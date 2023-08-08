import requests, random, webbrowser
from bs4 import BeautifulSoup
url = 'https://www.nytimes.com/?auth=login-google1tap&login=google1tap'
r = requests.get(url)
r_html = r.text

soup = BeautifulSoup(r_html, 'html.parser')


#p = open("div_results", "w")
#p.write(str(soup.find_all('div')))
#class="css-xdandi"><h3'

t= open("NY_Times_HTML Code.txt", "w")
t.write(soup.prettify())

article_div_elements = soup.find_all("div", class_="css-xdandi")

Dict1 = {}

for element in article_div_elements:
    # Extract article title
    article_title = element.get_text()
    # Extract href from the parent <a> element
    article_link = element.find_previous("a")["href"]

    Dict1[article_title] = article_link

random_article_title = random.choice(list(Dict1.keys()))

#opens the webpage
webbrowser.open(Dict1[random_article_title])