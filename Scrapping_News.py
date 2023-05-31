from textblob import TextBlob
from bs4 import BeautifulSoup
from urllib import request

class Analysis:
    link = 'https://news.google.com'
    def __init__(self, term):
        self.term = term
        self.sentiment = 0
        self.subjectivity = 0
        self.url = 'https://news.google.com/search?q={}&hl=en-US&gl=US&ceid=US%3Aen'.format(self.term)

    def fetch_url(self):
        opener = request.build_opener()
        opener.addheaders = [('User-agent', 'Mozilla/5.0 (X11; Linux x86_64; rv:10.0) Gecko/20100101 Firefox/10.0')]
        request.install_opener(opener)
        html_string = request.urlopen(self.url).read()
        return html_string.decode()

    def run(self):
        google_html = self.fetch_url()
        soup = BeautifulSoup(google_html, 'html.parser')
        topics = soup.select("h3 a")
        for headline in topics:
            #url = headline["href"]
            #fixed_url = self.link + url[1:]
            headline_results = headline.get_text()
            blob = TextBlob(headline_results)
            print(blob)
            self.subjectivity += blob.sentiment.subjectivity / len(topics)
            #print(blob.sentiment.subjectivity)
            self.sentiment += blob.sentiment.polarity / len(topics)
            #print(blob.sentiment.polarity)

searchword = input()
searchword_fixed = searchword.replace(" ", "%20")
word = Analysis(searchword_fixed)
word.run()
print("===========================================================================================================================")
print(searchword, 'Subjectivity:', word.subjectivity, 'Sentiment:', word.sentiment)