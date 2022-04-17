from bs4 import BeautifulSoup
import requests, openpyxl


excel = openpyxl.Workbook()

sheet = excel.active

sheet.title = "Top Rated Movies"
sheet.append(['Movie Rank', 'Movie Name', 'Release Year', 'IMDB Rating'])


url = 'https://www.imdb.com/chart/top/'

try:
    res = requests.get(url)
    res.raise_for_status()

    soup = BeautifulSoup(res.text, 'html.parser')

    movies = soup.find('tbody', class_='lister-list').find_all('tr')

   
    # iterate over tr tag
    for movie in movies:
        title_column = movie.find('td', class_='titleColumn')
        rating_column = movie.find('td', class_='ratingColumn imdbRating')

        rank = title_column.get_text(strip=True).split('.')[0]
        name = title_column.a.text # we can use the tag name (.a) or find
        year = title_column.span.text

        rating =  rating_column.strong.text

        print(rank, name, year, rating)
        sheet.append([rank, name, year, rating])
        

except Exception as e:
    print(e)


excel.save('IMDB_Movie_Rating.xlsx')
