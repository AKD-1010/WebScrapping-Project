from bs4 import BeautifulSoup
import requests
import openpyxl

# creating an excel file
excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top Rated Movies'
sheet.append(['Movie Rank', 'Movie Name', 'Year of Release', 'IMDB Rating'])

# Every code is under the try block so that if any connection issue arises we can catch that
try:
    # source = html response object
    source = requests.get('https://www.imdb.com/chart/top/')
    source.raise_for_status()

    # Parsing the response object from web using the inbuilt html parser 
    # and beautifulsoup lib and it will return beautiful soup object to soup variable
    soup = BeautifulSoup(source.text, 'html.parser')
    
    # we pass the html tag name along with the class name
    # since 'class' is a reserved keyword in python, we use class_ to specify html class, here html classes are attributes to the tag
    # now everything under tbody (i.e. every info about top 250 movies) is in movies variable
    movies = soup.find('tbody', class_ = 'lister-list' ).find_all('tr')
    

    for movie in movies:

        # getting the movie title  text from the 'a' tag under the "td" tag of the titlecolumn class
        name = movie.find('td', class_ = 'titleColumn').a.text

        # since under "td" tag we have more than one text, we need to split the text by . and get 
        # the rank no. which is in the left
        rank = movie.find('td', class_ = 'titleColumn').get_text(strip = True).split('.')[0]
        
        year = movie.find('td', class_ = 'titleColumn').span.text.strip('()')

        rating = movie.find('td', class_ = 'ratingColumn imdbRating').strong.text

        sheet.append([rank, name, year, rating])



except Exception as e:
    print(e)


# Saving the excel file
excel.save('Top Rated Movies.xlsx')

print("Scrapping done")