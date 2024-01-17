import requests
from bs4 import BeautifulSoup as BS
from googlesearch import search
from docx import Document
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

document = Document()
chrome_driver_path = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'
profile_directory = 'C:\\Users\\40gil\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 1'
chrome_options = Options()
chrome_options.add_argument(f'--user-data-dir={chrome_driver_path}/user-data')
chrome_options.add_argument(f'--profile-directory={profile_directory}')


def get_link(company, wiki=False, maya=False):
    if wiki:
        query = f"{company} ויקי "
    elif maya:
        query = f"{company} מאיה "
    search_res = search(query, tld="co.il", stop=3)
    for j in search_res:
        if (wiki and 'he.wikipedia.org/wiki/' in j) or (maya and 'maya.tase.co.il' in j):
            return j.split('?')[0]
    return None


def get_data_from_site(link,wiki=False, maya=False):
    response = requests.get(url=link)
    soup = BS(response.content, "html.parser")
    if wiki:
        return soup.find_all('p')
    elif maya:
        driver = webdriver.Chrome(executable_path=chrome_driver_path, options=chrome_options)
        driver.get(link)
        driver.implicitly_wait(10)
        page_source = driver.page_source
        driver.quit()
        #temp= soup.find_all('div', {'class': 'tableCol col1'})
        kaki=1
    return None


def format_html_txt(html_string):
    soup = BS(html_string, 'html.parser')
    text_without_tags = soup.get_text()
    return text_without_tags


def add_wiki_sum(wiki_link, paragraphs=3):
    wiki_text = ''
    if wiki_link is not None:
        wiki_data = get_data_from_site(link=wiki_link,wiki=True)
        for i in range(0, paragraphs):
            wiki_text += format_html_txt(wiki_data[i].get_text())
            wiki_text += '\n'
    else:
        wiki_text = 'לא נמצא לינק לויקיפדיה'
    document.add_paragraph(wiki_text)


def add_maya_sum(maya_link):
    maya_text = ''
    if maya_link is not None:
        maya_data = get_data_from_site(link=maya_link,maya=True)
    else:
        maya_text = 'לא נמצא לינק למאיה'
    document.add_paragraph(maya_text)


def headline_and_wiki_txt(_company_name):
    wiki_headline = '    1. כללי: ויקיפדיה'
    maya_headline = 'בעלות: מאיה'
    document.add_heading(_company_name, 0)

    # ----------- WIKIPEDIA DATA -----------#
    document.add_heading(wiki_headline, 1)
    wiki_link = get_link(company=company_name, wiki=True)
    add_wiki_sum(wiki_link=wiki_link)
    # -------------- MAYA DATA --------------#
    document.add_heading(maya_headline, 1)
    maya_link = get_link(company=company_name, maya=True)
    add_maya_sum(maya_link=maya_link)


if __name__ == '__main__':
    paragraph = document.add_paragraph()
    company_name = 'דליה אנרגיה'
    headline_and_wiki_txt(_company_name=company_name)
    document.save(f'{company_name}.docx')
