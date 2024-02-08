import requests
from bs4 import BeautifulSoup as BS
from googlesearch import search
from docx import Document
from selenium.webdriver.chrome.options import Options
from requests_html import HTMLSession
from dotenv import load_dotenv
import openai
import os
from selenium import webdriver


#region Globals

document = Document()
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


def get_data_from_site(link, wiki=False, maya=False):
    if wiki:
        response = requests.get(url=link)
        soup = BS(response.content, "html.parser")
        return soup.find_all('p')
    elif maya:
        driver.get(link)
        page_source = driver.page_source
        soup = BS(page_source, 'html.parser')
        return soup
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


def add_wiki_sum(wiki_link, paragraphs=3):
    wiki_text = ''
    if wiki_link is not None:
        wiki_data = get_data_from_site(link=wiki_link, wiki=True)
        for i in range(0, paragraphs):
            wiki_text += format_html_txt(wiki_data[i].get_text())
            wiki_text += '\n'
    else:
        wiki_text = 'לא נמצא לינק לויקיפדיה'
    document.add_paragraph(wiki_text)


def add_maya_sum(maya_link):
    maya_text = ''
    if maya_link is not None:
        maya_data = get_data_from_site(link=maya_link, maya=True)
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
    get_data_from_site(link='https://maya.tase.co.il/company/1840', maya=True)
    prompt = "check openai api from my python script"
    generated_response = ask_gepeto(prompt)
    print("Generated Response:")
    print(generated_response)


    paragraph = document.add_paragraph()
    company_name = 'דליה אנרגיה'
    headline_and_wiki_txt(_company_name=company_name)
    document.save(f'{company_name}.docx')
