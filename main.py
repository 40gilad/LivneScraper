import requests
from bs4 import BeautifulSoup as BS
from googlesearch import search
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
from dotenv import load_dotenv
import os
from selenium import webdriver

# region Globals

GPT3 = "gpt-3.5-turbo"
GPT4 = "gpt-4-0125-preview"
document = Document()
section = document.sections[0]
section.right_to_left = True
chrome_driver_path=r"C:\Users\40gil\Desktop\Helpful\Scraping\chromedriver.exe"

# endregion

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

# region Doc manipulation
def add_paragraph(text, style=None):
    par = document.add_paragraph(text, style=style)
    run = par.runs[0]  # Assuming there is only one run in the paragraph

    # Set the font of the text
    font = run.font
    font.name = 'Calibri'  # Set the font name
    font.size = Pt(12)  # Set the font size
    par.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT


def add_headline(text, size=1):
    global document
    hed = document.add_heading(text, size)
    hed.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT


# endregion

# region Scraping helpers
def format_html_txt(html_string):
    soup = BS(html_string, 'html.parser')
    text_without_tags = soup.get_text()
    return text_without_tags


def get_link(company, wiki=False, maya=False, bizportal=False,facebook=False,instagram=False,linkedin=False):
    if wiki:
        query = f"{company} ויקי "
    elif maya:
        query = f"{company} מאיה "
    elif bizportal:
        query = f"{company} ביזפורטל אגח "
    elif instagram:
        query = f"{company} instagram"
    elif facebook:
        query = f"{company} facebook"
    elif linkedin:
        query = f"{company} linkedin"

    search_res = search(query, tld="co.il", stop=3)
    for j in search_res:
        if (
                (wiki and 'he.wikipedia.org/wiki/' in j)
                or (maya and 'maya.tase.co.il' in j and 'details' in j)
                or (bizportal and 'bizportal.co.il' in j)
                or (facebook and 'facebook.com' in j)
                or (instagram and 'instagram.com' in j)
                or (linkedin and 'linkedin.com' in j)

        ):
            return j.split('?')[0]
    return None


def crawl_shareholders_table(table):
    shareholders_data = '\n'
    for t in table:
        shareholders_data += "\n" + t.get_text(separator=' ', strip=True)
    extracted_data = ask_gepeto(f"the following text represent a table.\
                extract for me the name of the company, and it's percentage in bullets for each row. before the dates there is the company name.: {shareholders_data}")
    return extracted_data


def crawl_reports(reports=None):
    if reports:
        for i in range(0, 10):
            txt ='- '+reports[i].find_all('span', class_='feedItemDateMobile ng-binding')[0].get_text(separator=' ',
                                                                                                  strip=True)
            if txt is not None:
                add_paragraph(txt, style='List Bullet')
            txt = reports[i].find_all('a', class_='messageContent')[0].get_text(separator=' ', strip=True) + '\n'

            if txt is not None:
                add_paragraph(txt)


def get_data_from_site(link, wiki=False, maya=False, bizportal=False, maya_reports=False):
    if wiki:
        response = requests.get(url=link)
        soup = BS(response.content, "html.parser")
        wiki_data = soup.find_all('p')
        wiki_text = '\n'
        for i in range(0, 5):
            wiki_text += format_html_txt(wiki_data[i].get_text()) + '\n'
        return ask_gepeto(f"summerize this whole text for me in *hebrew*:{wiki_text}")
    elif maya:
        driver = webdriver.Chrome(executable_path=chrome_driver_path)
        driver.get(link)
        page_source = driver.page_source
        soup = BS(page_source, 'html.parser')
        driver.quit()
        try:
            shareholders_section = soup.find('div', class_='listTable share-holders-grid')
            table = shareholders_section.find_all('div', class_='tableCol')
            return crawl_shareholders_table(table=table)
        except Exception as e:
            print(f"{link}, for maya\nexception: {e}")
            return 'משהו השתבש'
    elif maya_reports:
        driver = webdriver.Chrome(executable_path=chrome_driver_path)
        driver.get(link)
        page_source = driver.page_source
        soup = BS(page_source, 'html.parser')
        driver.quit()
        try:
            reports_section = soup.find('maya-reports')
            reports = reports_section.find_all('div', class_='feedItem ng-scope')
            return crawl_reports(reports=reports)
        except Exception as e:
            print(f"{link}, for maya reports\nexception: {e}")
            return 'משהו השתבש'
    elif bizportal:
        print("BIZPORTAL, GPT4 PROMPT:")
        print(f"{link} \n"
                          f"give me the bonds detailes for this company.\
                            i want the bond value, kind, it's Interest rate and other useful data\
                            give me the data in bullets and in *hebrew*"
                          f"Please write only the information, without additions of introduction or conclusion")
        return ask_gepeto(f"{link} \n"
                          f"give me the bonds detailes for this company.\
                            i want the bond value, kind, it's Interest rate and other useful data\
                            give me the data in bullets and in *hebrew*"
                          f"Please write only the information, without additions of introduction or conclusion", model=GPT4)
    return None


# endregion

# region GEPETO

def ask_gepeto(prompt, model=GPT3):
    # ChatGPT API endpoint
    endpoint = "https://api.openai.com/v1/chat/completions"

    # Headers containing the Authorization with your API key
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }

    # Data payload containing the prompt and other parameters
    data = {
        "model": model,  # You can choose different models as per your requirement
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


# endregion

# region add_<section>_sum
def add_wiki_sum(wiki_link=None):
    if wiki_link is not None:
        wiki_data = get_data_from_site(link=wiki_link, wiki=True)
    else:
        wiki_text = 'לא נמצא לינק לויקיפדיה'
    add_paragraph(wiki_data)


def add_maya_sum(maya_link=None):
    maya_text = ''
    if maya_link is not None:
        maya_text = get_data_from_site(link=maya_link, maya=True)
    else:
        maya_text = 'לא נמצא לינק למאיה'
    add_paragraph(maya_text)


def add_bizportal_sum(bizportal_link=None):
    bizportal_text = ''
    if bizportal_link is not None:
        bizportal_text = get_data_from_site(link=bizportal_link, bizportal=True)
    else:
        bizportal_text = 'לא נמצא לינק לביזפורטל'
    add_paragraph(bizportal_text)


def add_key_people(company_name=None):
    key_people_text = ''
    if company_name is not None:
        key_people_text = ask_gepeto(
            prompt=f"give me in bullets and in *HEBREW* the key people from the company {company_name}"
                   f"Please write only the information, without additions of introduction or conclusion",
            model=GPT4)
    else:
        key_people_text = "לא נמצאו אנשי מפתח"
    add_paragraph(key_people_text)


def add_last_reports(maya_link=None):
    maya_text = ''
    if maya_link is not None:
        maya_text = get_data_from_site(link=maya_link, maya_reports=True)
    else:
        maya_text = 'לא נמצאו דיווחים אחרונים במאיה'
        add_paragraph(maya_text)

def add_juice(company_name=None):
    juice_text =''
    if company_name is not None:
        juice_text = ask_gepeto(prompt=f"give me a short summary *IN HEBREW* about the company {company_name}"
                                       f"everything intersting and relevant for money investing"
                                       f"Please write only the information, without additions of introduction or conclusion",model=GPT4)
    else:
        juice_text = 'לא נמצא מידע רלוונטי בעיתונות ובאינטרנט'
    add_paragraph(juice_text)

def add_social(company_name=None):
    # ----------- FACEBOOK URL -----------#
    txt = 'פייסבוק- '
    add_paragraph(txt, style='List Bullet')
    link=get_link(company_name,facebook=True)
    if link is not None:
        add_paragraph(link)

    # ----------- INSTAGRAM URL -----------#
    txt = 'אינסטגרם- '
    add_paragraph(txt, style='List Bullet')
    link=get_link(company_name,instagram=True)
    if link is not None:
        add_paragraph(link)

    # ----------- LINKEDIN URL -----------#
    txt = 'לינקדאין- '
    add_paragraph(txt, style='List Bullet')
    link=get_link(company_name,linkedin=True)
    if link is not None:
        add_paragraph(link)

# endregion
def scrape_and_sum(_company_name):
    wiki_headline = '    1. כללי: ויקיפדיה'
    maya_headline = 'בעלות: מאיה'
    bizportal_headline = 'ני"ע: ביזפורטל'
    key_people_headline = 'אנשי מפתח: CHATGPT'
    imeidiate_reports = 'דווחים מידיים: מאיה'
    juice_headline= 'עיתונות: CHATGPT'
    social_headline= 'סושיאל:'
    add_headline(text=_company_name, size=0)

    # ----------- WIKIPEDIA DATA -----------#
    add_headline(text=wiki_headline)
    wiki_link = get_link(company=company_name, wiki=True)
    add_wiki_sum(wiki_link=wiki_link)
    # -------------- MAYA DATA --------------#
    add_headline(text=maya_headline)
    maya_link = get_link(company=company_name, maya=True)
    add_maya_sum(maya_link=maya_link)
    # -------------- BIZPORTAL DATA --------------#
    add_headline(text=bizportal_headline)
    bizportal_link = get_link(company=company_name, bizportal=True)
    add_bizportal_sum(bizportal_link=bizportal_link)
    # -------------- KEY PEOPLE DATA --------------#
    add_headline(text=key_people_headline)
    add_key_people(company_name=_company_name)
    # -------------- IMMIDIATE REPORTS --------------#
    add_headline(text=imeidiate_reports)
    link = f"{maya_link}?view=reports&q=%7B%22DateFrom%22:%222023-02-13T22:00:00.000Z%22,%22DateTo%22:%222024-02-13T22:00:00.000Z%22,%22Page%22:1,%22entity%22:%221840%22,%22events%22:%5B%5D,%22subevents%22:%5B%5D%7D"
    add_last_reports(maya_link=link)
    # -------------- JUICE DATA --------------#
    add_headline(text=juice_headline)
    add_juice(company_name=_company_name)
    # -------------- SOCIAL DATA --------------#
    add_headline(text=social_headline)
    add_social(company_name=_company_name)







if __name__ == '__main__':
    paragraph = document.add_paragraph()
    company_name = 'הפניקס'
    scrape_and_sum(_company_name=company_name)
    try:
        document.save(f'{company_name}.docx')
    except:
        document.save(f'{company_name}_second.docx')

