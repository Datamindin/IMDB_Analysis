from bs4 import BeautifulSoup
import requests,openpyxl

excel = openpyxl.Workbook()
print(excel.sheetnames)
sheet = excel.active
sheet.title = 'Top Rated Movies'
print(excel.sheetnames)
sheet.append(['Movie Title', 'Release year', 'Runtime','IMDB rating','No of Votes'])


url = "https://www.imdb.com/chart/top/"
try:
#     source  = requests.get(url)
    
    HEADERS = {'User-Agent': 'Mozilla/5.0 (iPad; CPU OS 12_2 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/15E148'}

    source = requests.get(url, headers=HEADERS)
    source.raise_for_status()
    soup = BeautifulSoup(source.text, 'html.parser')

    div_text=soup.find("ul",{"ipc-metadata-list ipc-metadata-list--dividers-between sc-3f13560f-0 sTTRj compact-list-view ipc-metadata-list--base"}).find_all('li')
#     print(len(div_text))
    
    for i in div_text:
        
        name = i.find('div', class_= "ipc-title ipc-title--base ipc-title--title ipc-title-link-no-icon ipc-title--on-textPrimary sc-b85248f1-7 lhgKeb cli-title").a.text
        year = i.find('span', class_ = 'sc-b85248f1-6 bnDqKN cli-title-metadata-item')

        duration = i.find('span', class_ = 'sc-b85248f1-6 bnDqKN cli-title-metadata-item').get_text()
        infos=i.find("div",class_ = {"sc-b85248f1-5 kZGNjY cli-title-metadata"}).find_all('span')    
        
        rating = i.find("span",class_ = {"ipc-rating-star ipc-rating-star--base ipc-rating-star--imdb ratingGroup--imdb-rating"}).text.split("\xa0(")[0]
        voting = i.find("span",class_ = {"ipc-rating-star ipc-rating-star--base ipc-rating-star--imdb ratingGroup--imdb-rating"}).text.split("\xa0(")[1].replace(")", "")

        
#         print(infos)
        info_list = []
        for info in infos:
            val = info.text
            info_list.append(val)
#         print(info_list)   
        
        
        result = " ".join(info_list)
        year = info_list[0]
        runtime = info_list[1]
        
            
#         print(result)
#         print(rating)

        print(name, year, runtime, rating, voting)
        sheet.append([name, year, runtime, rating, voting])
        
except Exception as e:
    print(e)
    
excel.save("python_loaded_movies.csv")