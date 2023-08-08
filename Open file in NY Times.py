import requests, random, webbrowser
from bs4 import BeautifulSoup
url = 'https://www.nytimes.com/?auth=login-google1tap&login=google1tap'
r = requests.get(url)
r_html = r.text
#using requests we get the html code and make it into soup with html parser

soup = BeautifulSoup(r_html, 'html.parser')


#upon inspection of HTML code, title with links are embedded under this:
#class="css-xdandi"'

#writes a file with the HTML code of NY times version run formatted in a easier way to read with the prettify function
t= open("NY_Times_HTML Code.txt", "w")
t.write(soup.prettify())

#this is to find the titles which goes in and looks for the div class defined by the name mentioned in an earlier comment
article_div_elements = soup.find_all("div", class_="css-xdandi")

#create a dictionary which has relatively fast run-time to store the elements and randomize the process
Dict1 = {}

for element in article_div_elements:
    # Extract article title
    article_title = element.get_text()
    # Extract href from the parent <a> element
    article_link = element.find_previous("a")["href"]

    Dict1[article_title] = article_link

#function chooses random element from dictionary
random_article_title = random.choice(list(Dict1.keys()))

#opens the webpage
webbrowser.open(Dict1[random_article_title])