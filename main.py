import requests
from bs4 import BeautifulSoup as BS
from googlesearch import search
from docx import Document

"""
response = requests.get(
	url="https://he.wikipedia.org/wiki/%D7%AA%D7%97%D7%A0%D7%AA_%D7%94%D7%9B%D7%95%D7%97_%D7%93%D7%9C%D7%99%D7%94",
)
print(response.status_code)

soup = BS(response.content, "html.parser")
temp=soup.find_all('p')
kaki=1
"""
document=Document()
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
	return temp


if __name__=='__main__':
	paragraph = document.add_paragraph()
	company_name='דליה חברות אנרגיה'
	wiki_headline='	1. כללי: ויקיפדיה ואתר אינטרנט'
	document.add_heading(company_name, 0)
	document.add_heading(wiki_headline, 1)

	wiki_link=get_wiki_link(company_name)
	if wiki_link is not None:
		wiki_text=get_data_from_wiki(wiki_link)[0]
	else:
		wiki_text='NO WIKIPEDIA LINK FOUND'
	document.add_paragraph(wiki_text)


	document.save('test.docx')



