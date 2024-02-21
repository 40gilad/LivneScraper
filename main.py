import requests
from bs4 import BeautifulSoup as BS
from googlesearch import search
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
from dotenv import load_dotenv
import os
from selenium import webdriver
import datetime

from selenium.webdriver.support.wait import WebDriverWait

# region Globals

GPT3 = "gpt-3.5-turbo"
GPT4 = "gpt-4-0125-preview"
document = Document()
section = document.sections[0]
section.right_to_left = True
output_folder = None
bizportal_root = "https://www.bizportal.co.il"
bizportal_data_link = "https://www.bizportal.co.il/bonds/quote/bondsdata/"
bizportal_search_page = "https://www.bizportal.co.il/list/searchpapers?search="
chrome_driver_path = r"C:\Users\40gil\Desktop\Helpful\Scraping\chromedriver.exe"

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

def create_output_folder(path=None):
    output_folder = path
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)


def add_paragraph(text=None, style=None, subheadline=False):
    if text is None:
        raise ValueError("text is required")
        return
    par = document.add_paragraph(text, style=style)
    run = par.runs[0]  # Assuming there is only one run in the paragraph
    font = run.font
    font.name = 'Calibri'  # Set the font name
    if subheadline:
        font.size = Pt(18)  # Set the font size
    else:
        font.size = Pt(12)  # Set the font size

    par.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT


def add_headline(text, size=2):
    global document
    hed = document.add_heading(text, size)
    hed.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT


def insert_bizportal_data(root_bonds_url_data=None, driver=None):
    if root_bonds_url_data is None or driver is None:
        raise ValueError("root_bonds_url_data or driver are None")
        return
    sub_headline = f'{len(root_bonds_url_data)}: אגח '
    add_paragraph(text=f"{sub_headline}\n", subheadline=True)

    for d in root_bonds_url_data:

        # get root bond page
        driver.get(d['root_link'])
        page_source = driver.page_source
        soup = BS(page_source, 'html.parser')

        # add bonds sum as subheadline and the name of the bond
        add_paragraph(text=f" - {d['name']}  ", style='List Bullet', subheadline=True)

        # from root page
        spans = soup.find_all(name='span', class_='label')

        # every if inc counter, so if all ifs happens -> break
        iter_counter = 0

        for span in spans:

            if iter_counter == 2:  # already got into all ifs
                iter_counter = 0
                break
            curr_span = span.get_text()

            # todo NEED TO INSERT VALUE

            if curr_span == 'ענף':

                iter_counter += 1

                add_paragraph(text=f"{span.find_next('span').get_text()} ", style='List Bullet 2')
            elif curr_span == 'מח"מ:':

                iter_counter += 1

                val = span.find_next('span').get_text()
                text_to_insert = f"{curr_span} {val} "
                add_paragraph(text=text_to_insert, style='List Bullet 2')

        # get data bond page
        driver.get(d['data_link'])
        page_source = driver.page_source
        soup = BS(page_source, 'html.parser')

        # from data page
        spans = soup.find_all(name='td', class_='label')
        for span in spans:

            if iter_counter == 3:
                # unknwon ribit type and ribit val order of reading, so insrting at the end
                add_paragraph(text=f"{ribit_type} {ribit_val}", style='List Bullet 2')
                break

            curr_span = span.get_text()
            if curr_span == 'סוג ריבית':
                ribit_type = f" ריבית {span.find_next(name='td', class_='num').get_text()} "
                iter_counter += 1

            elif curr_span == 'שיעור ריבית':

                iter_counter += 1
                val = span.find_next(name='td', class_="num").get_text()
                ribit_val = f"% {val} "

            elif curr_span == 'דרוג מידרוג + אופק':

                iter_counter += 1
                text_to_insert = "דירוג מידרוג "
                text_to_insert += span.find_next(name='td', class_="num").get_text()
                add_paragraph(text=text_to_insert, style='List Bullet 2')


# endregion

# region Scraping helpers
def format_html_txt(html_string):
    soup = BS(html_string, 'html.parser')
    text_without_tags = soup.get_text()
    return text_without_tags


def get_link(company, wiki=False, maya=False, bizportal=False, facebook=False, instagram=False, linkedin=False):
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
                or (maya and 'maya.tase.co.il' in j)
                or (bizportal and 'bizportal.co.il' in j)
                or (facebook and 'facebook.com' in j)
                or (instagram and 'instagram.com' in j)
                or (linkedin and 'linkedin.com' in j)

        ):
            return j.split('?')[0]
    return None


def get_bonds(bond_name, driver):
    """

    :param bond_name:
    :param driver:
    :return: [{
                'root_link': root_link,
                'name': bond_name,
                'bond_number': bond_number
            },
            {
                .....
            }
            ]
    """
    bond_name = bond_name.replace(" ", "%20")
    link = f"{bizportal_search_page}{bond_name}"
    driver.get(link)
    page_source = driver.page_source
    soup = BS(page_source, 'html.parser')
    bonds_links = soup.find_all('a', class_='link')
    ret = []
    for link in bonds_links:
        root_link = f"{bizportal_root}{link.get('href')}"
        root_splitted = root_link.split("/")
        bond_number = root_splitted[len(root_splitted) - 1]
        ret.append(
            {
                'root_link': root_link,
                'name': link.get_text(),
                'bond_number': bond_number,
                'data_link': f"{bizportal_data_link}{bond_number}"
            }
        )
    return ret


def crawl_shareholders_table(table):
    shareholders_data = '\n'
    for t in table:
        shareholders_data += "\n" + t.get_text(separator=' ', strip=True)
    extracted_data = ask_gepeto(f"the following text represent a table.\
                                extract for me the name of the company, and it's percentage for each row."
                                f" before the dates there is the company name.: {shareholders_data}"
                                f"Please write only the information, without additions of introduction or conclusion.",
                                model=GPT4)
    return extracted_data


def crawl_reports(reports=None):
    if reports:
        for i in range(0, 10):
            txt = '- ' + reports[i].find_all('span', class_='feedItemDateMobile ng-binding')[0].get_text(separator=' ',
                                                                                                         strip=True)
            if txt is not None:
                add_paragraph(txt, style='List Bullet')
            txt = reports[i].find_all('a', class_='messageContent')[0].get_text(separator=' ', strip=True) + '\n'

            if txt is not None:
                add_paragraph(f"    {txt}")


def get_data_from_site(link=None, company_name=None, bond_name=None,
                       wiki=False, maya=False, key_people=False, bizportal=False, maya_reports=False):
    if wiki:
        paragraphs_limit=5
        response = requests.get(url=link)
        soup = BS(response.content, "html.parser")
        wiki_data = soup.find_all('p')
        wiki_text = '\n'
        if len(wiki_data) < paragraphs_limit:
            paragraphs_limit =len(wiki_data)
        for i in range(0, paragraphs_limit):
            wiki_text += format_html_txt(wiki_data[i].get_text()) + '\n'
        return ask_gepeto(f"summerize this whole text for me in *hebrew*:{wiki_text}")
    elif maya:
        driver = webdriver.Chrome(executable_path=chrome_driver_path)
        driver.get(link)
        page_source = driver.page_source
        soup = BS(page_source, 'html.parser')
        driver.quit()
        share_holders_table = soup.find_all(name="div", class_="listTable share-holders-grid")
        rows = share_holders_table[0].find_all(name="div", class_="ng-scope tableRow")

        ret_txt = ''
        for row in rows:
            comp_name_tag=row.find(name="div")
            precentage_tag=row.find(name="div", class_="tableCol col_6 ng-binding ng-scope")
            if comp_name_tag is not None and precentage_tag is not None:
                comp_name = comp_name_tag.get_text()
                precentage = precentage_tag.get_text()
                ret_txt += f"% {comp_name} - {precentage}\n"

        return ret_txt

    elif key_people:
        driver = webdriver.Chrome(executable_path=chrome_driver_path)
        driver.get(link)
        page_source = driver.page_source
        soup = BS(page_source, 'html.parser')
        driver.quit()
        key_people_table = soup.find_all(name="div", class_="mobileTableFrame")[2]
        rows = key_people_table.find_all(name="div", class_="ng-scope tableRow")

        ret_txt = ''
        for row in rows:
            person_name_tag = row.find(name="div")
            person_role = person_name_tag.find_next(name="div").get_text()
            if person_role != "דירקטור רגיל":
                person_name = person_name_tag.get_text()
                if "מאזן" in person_name:
                    return "לא נמצאו אנשי מפתח"
                ret_txt += f"{person_role} - {person_name}\n"
        return ret_txt





    elif maya_reports:
        driver = webdriver.Chrome(executable_path=chrome_driver_path)
        driver.get(link)
        page_source = driver.page_source
        soup = BS(page_source, 'html.parser')
        driver.quit()
        try:
            reports_section = soup.find('maya-reports')
            reports = reports_section.find_all('div', class_='feedItem ng-scope')
            crawl_reports(reports=reports)

        except Exception as e:
            print(f"{link}, for maya reports\nexception: {e}")
            return 'משהו השתבש'

    elif bizportal:
        driver = webdriver.Chrome(executable_path=chrome_driver_path)
        root_bonds_url_data = get_bonds(bond_name=bond_name, driver=driver)
        insert_bizportal_data(root_bonds_url_data=root_bonds_url_data, driver=driver)
        driver.quit()

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


def add_bizportal_sum(company_name=None, bond_name=None):
    bizportal_text = ''
    if company_name is not None:
        get_data_from_site(company_name=company_name, bond_name=bond_name, bizportal=True)


def add_key_people(maya_link=None):
    key_text = ''
    if maya_link is not None:
        key_text = get_data_from_site(link=maya_link, key_people=True)
    else:
        key_text = 'לא נמצא לינק למאיה'
    add_paragraph(key_text)


def add_last_reports(maya_link=None):
    maya_text = ''
    if maya_link is not None:
        maya_text = get_data_from_site(link=maya_link, maya_reports=True)
    else:
        maya_text = 'לא נמצאו דיווחים אחרונים במאיה'
        add_paragraph(maya_text)


def add_juice(company_name=None):
    juice_text = ''
    if company_name is not None:
        juice_text = ask_gepeto(prompt=f"give me a short summary *IN HEBREW* about the company {company_name}"
                                       f"everything intersting and relevant for money investing"
                                       f"Please write only the information, without additions of introduction or conclusion",
                                model=GPT4)
    else:
        juice_text = 'לא נמצא מידע רלוונטי בעיתונות ובאינטרנט'
    add_paragraph(juice_text)


def add_social(company_name=None):
    # ----------- FACEBOOK URL -----------#
    txt = '-פייסבוק '
    add_paragraph(txt, style='List Bullet')
    link = get_link(company_name, facebook=True)
    if link is not None:
        add_paragraph(f"    {link}")

    # ----------- INSTAGRAM URL -----------#
    txt = '-אינסטגרם '
    add_paragraph(txt, style='List Bullet')
    link = get_link(company_name, instagram=True)
    if link is not None:
        add_paragraph(f"    {link}")

    # ----------- LINKEDIN URL -----------#
    txt = '-לינקדאין '
    add_paragraph(txt, style='List Bullet')
    link = get_link(company_name, linkedin=True)
    if link is not None:
        add_paragraph(f"    {link}")


# endregion
def scrape_and_sum(_company_name=None, bond_name=None):
    if _company_name is None or bond_name is None:
        raise ValueError('_Company_name or bond_name are None. both should have valid value')

    wiki_headline = '    כללי: ויקיפדיה'
    maya_headline = 'בעלות: מאיה'
    bizportal_headline = 'ני"ע: ביזפורטל'
    key_people_headline = 'אנשי מפתח : מאיה '
    imeidiate_reports = 'דיווחים מידיים: מאיה'
    juice_headline = 'עיתונות (מומלץ לאמת את המידע) : CHATGPT'
    social_headline = ':סושיאל'
    add_headline(text=_company_name, size=0)

    # ----------- WIKIPEDIA DATA -----------#
    add_headline(text=wiki_headline)
    wiki_link = get_link(company=_company_name, wiki=True)
    add_wiki_sum(wiki_link=wiki_link)
    # -------------- MAYA DATA --------------#
    add_headline(text=maya_headline)
    maya_link = get_link(company=_company_name, maya=True)
    add_maya_sum(maya_link=maya_link)
    # -------------- BIZPORTAL DATA --------------#
    add_headline(text=bizportal_headline)
    bizportal_link = get_link(company=_company_name, bizportal=True)
    add_bizportal_sum(company_name=_company_name, bond_name=bond_name)
    # -------------- KEY PEOPLE DATA --------------#
    add_headline(text=key_people_headline)
    add_key_people(maya_link=maya_link)
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


def start(_company_name=None, bond_name=None, to_save_path=None, chrome_driver=None):
    global output_folder, chrome_driver_path

    output_folder = to_save_path
    chrome_driver_path = chrome_driver
    scrape_and_sum(_company_name=_company_name, bond_name=bond_name)
    if output_folder is None:
        try:
            document.save(f'{_company_name}.docx')
        except:
            document.save(f'{_company_name}_{datetime.datetime.now().strftime("%d.%m_%H.%M")}.docx')
    else:
        try:
            document.save(f'{output_folder}\\{_company_name}.docx')
        except:
            document.save(f'{output_folder}\\{_company_name}_{datetime.datetime.now().strftime("%d.%m_%H.%M")}.docx')


if __name__ == '__main__':
    start()
    """
    paragraph = document.add_paragraph()
    company_name = 'אלביט מערכות'
    scrape_and_sum(_company_name=company_name, bond_name='אלביט מערכות אגח')

    if output_folder is None:
        try:
            document.save(f'{company_name}.docx')
        except:
            document.save(f'{company_name}_{datetime.datetime.now().strftime("%d.%m_%H.%M")}.docx')
    else:
        try:
            document.save(f'{output_folder}\\{company_name}.docx')
        except:
            document.save(f'{output_folder}\\{company_name}_{datetime.datetime.now().strftime("%d.%m_%H.%M")}.docx')
            """
