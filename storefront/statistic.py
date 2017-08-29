from bs4 import BeautifulSoup
soup = BeautifulSoup(xml, 'xml')

from bs4 import BeautifulStoneSoup
xml = "<doc><tag1>Contents 1<tag2>Contents 2<tag1>Contents 3"
soup = BeautifulStoneSoup(xml)
print(soup.prettify())