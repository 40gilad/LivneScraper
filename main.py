import requests
from bs4 import BeautifulSoup as BS
from googlesearch import search
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
from dotenv import load_dotenv
import os
from selenium import webdriver


#region Globals

document = Document()
section = document.sections[0]
section.right_to_left = True
driver = webdriver.Chrome(executable_path=r"C:\Users\40gil\Desktop\Helpful\Scraping\chromedriver.exe")  # You may need to download the appropriate webdriver for your browser (e.g., ChromeDriver)


#endregion

# region Load .env


env_path = r'C:\Users\40gil\Desktop\Helpful\Scraping\LivneScraperPY\LivneScraper.env'
try:
    load_dotenv(dotenv_path=env_path)
    api_key = os.getenv('OPENAI_KEY')
    if api_key is None:
        raise ValueError("OpenAI API key not found in environment variables.")
except Exception as err:
    print(f"Error: {err}")
    # Handle the error accordingly, e.g., exit the script or provide a default API key.



# endregion



def add_paragraph(text):
    par= document.add_paragraph(text)
    run = par.runs[0]  # Assuming there is only one run in the paragraph

    # Set the font of the text
    font = run.font
    font.name = 'Calibri'  # Set the font name
    font.size = Pt(12)  # Set the font size
    par.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

def add_headline(text,size=1):
    hed= document.add_heading(text, size)
    hed.alignment =WD_PARAGRAPH_ALIGNMENT.RIGHT

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

def crawl_shareholders_table(table):
    shareholders_data = '\n'
    for t in table:
        shareholders_data += "\n" + t.get_text(separator=' ', strip=True)
    extracted_data = ask_gepeto(f"the following text represent a table.\
                extract for me the data in bullets for each row. before the dates there is the company name."
                                f"please make a newline before you write the company name: {shareholders_data}")
    return extracted_data

def get_data_from_site(link, wiki=False, maya=False):
    if wiki:
        response = requests.get(url=link)
        soup = BS(response.content, "html.parser")
        wiki_data= soup.find_all('p')
        wiki_text='\n'
        for i in range(0,5):
            wiki_text += format_html_txt(wiki_data[i].get_text())+'\n'
        return ask_gepeto(f"summerize this whole text for me in hebrew:{wiki_text}")
    elif maya:
        driver.get(link)
        page_source = driver.page_source
        soup = BS(page_source, 'html.parser')
        driver.quit()
        shareholders_section = soup.find('div', class_='listTable share-holders-grid')
        table=shareholders_section.find_all('div',class_='tableCol')
        return crawl_shareholders_table(table=table)
    return None


def ask_gepeto(prompt):
    # ChatGPT API endpoint
    endpoint = "https://api.openai.com/v1/chat/completions"


    # Headers containing the Authorization with your API key
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }

    # Data payload containing the prompt and other parameters
    data = {
        "model": "gpt-3.5-turbo",  # You can choose different models as per your requirement
        "messages": [
            {"role": "user", "content": prompt}
        ]
    }

    try:
        # Making a POST request to the API
        response = requests.post(endpoint, json=data, headers=headers)

        # Check if the request was successful
        if response.status_code == 200:
            return response.json()["choices"][0]["message"]["content"]
        else:
            return f"Error: {response.status_code} - {response.text}"
    except Exception as e:
        return f"Error: {str(e)}"

def format_html_txt(html_string):
    soup = BS(html_string, 'html.parser')
    text_without_tags = soup.get_text()
    return text_without_tags


def add_wiki_sum(wiki_link):
    if wiki_link is not None:
        wiki_data = get_data_from_site(link=wiki_link, wiki=True)
    else:
        wiki_text = 'לא נמצא לינק לויקיפדיה'
    add_paragraph(wiki_data)


def add_maya_sum(maya_link):
    maya_text = ''
    if maya_link is not None:
        maya_text = get_data_from_site(link=maya_link, maya=True)
    else:
        maya_text = 'לא נמצא לינק למאיה'
    add_paragraph(maya_text)


def headline_and_wiki_txt(_company_name):
    wiki_headline = '    1. כללי: ויקיפדיה'
    maya_headline = 'בעלות: מאיה'
    add_headline(text=_company_name,size=0)

    # ----------- WIKIPEDIA DATA -----------#
    add_headline(text=wiki_headline)
    wiki_link = get_link(company=company_name, wiki=True)
    add_wiki_sum(wiki_link=wiki_link)
    # -------------- MAYA DATA --------------#
    add_headline(text=maya_headline)
    maya_link = get_link(company=company_name, maya=True)
    add_maya_sum(maya_link=maya_link)


if __name__ == '__main__':
    paragraph = document.add_paragraph()
    company_name = 'האחים צברי'
    headline_and_wiki_txt(_company_name=company_name)
    document.save(f'{company_name}.docx')
