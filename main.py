# import external modules
from bs4 import BeautifulSoup
import requests


try:
    # requesting the IMDB website
    url = "https://www.imdb.com/chart/top"

    source = requests.get(url)
    # throws an error in case of the URL has issues
    source.raise_for_status()

    # read and parse HTML file
    soup = BeautifulSoup(source.text, "html.parser")

    # reading movies table
    movies = soup.find("tbody", class_="lister-list").find_all("tr")

    for movie in movies:
        # read child elements
        titleColumnEl = movie.find("td", class_="titleColumn")
        ratingColumnEl = movie.find("td", class_="ratingColumn")

        # parse rank
        rank = titleColumnEl.get_text(strip=True).split(".")[0]

        # parse title
        title = titleColumnEl.a.text

        # parse year
        year = titleColumnEl.span.text.strip("()")

        # parse rating
        rating = ratingColumnEl.strong.text

        print(rank, title, year, rating)

except Exception as e:
    print(e)
