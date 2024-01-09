import requests
from bs4 import BeautifulSoup as BS
from googlesearch import search

"""
response = requests.get(
	url="https://he.wikipedia.org/wiki/%D7%AA%D7%97%D7%A0%D7%AA_%D7%94%D7%9B%D7%95%D7%97_%D7%93%D7%9C%D7%99%D7%94",
)
print(response.status_code)

soup = BS(response.content, "html.parser")
temp=soup.find_all('p')
kaki=1
"""

def get_wiki_link(param):
	query = f"{param} ויקי "
	search_res=search(query, tld="co.il",stop=3)
	for j in search_res:
		if 'he.wikipedia.org/wiki/' in j:
			return j
	return None


def get_data_from_wiki(wiki_link):
	response = requests.get(url=wiki_link)
	soup = BS(response.content, "html.parser")
	temp = soup.find_all('p')


if __name__=='__main__':
	wiki_link=get_wiki_link("דליה חברות אנרגיה")
	if wiki_link is not None:
		get_data_from_wiki(wiki_link)
	else:
		print ("No Wikipedia")
