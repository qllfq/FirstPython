import math
from xlutils.copy import copy
import requests
import re
# import pandas as pd
import xlrd as xlrd
import xlwt as xlwt
from requests.exceptions import RequestException
# from bs4 import BeautifulSoup
# from selenium import webdriver
# from pyquery import PyQuery as pq
# from selenium.common.exceptions import TimeoutException
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC, expected_conditions
# from selenium.webdriver.chrome.options import Options

# chrome_options = Options()
# chrome_options.add_argument('--no-sandbox')  # 解决DevToolsActivePort文件不存在的报错
# chrome_options.add_argument('--disable-gpu')  # 谷歌文档提到需要加上这个属性来规避bug
# chrome_options.add_argument('--hide-scrollbars')  # 隐藏滚动条, 应对一些特殊页面
# chrome_options.add_argument('blink-settings=imagesEnabled=false')  # 不加载图片, 提升速度
# chrome_options.add_argument('--headless')
# browser = webdriver.Chrome()
# wait = WebDriverWait(browser,30)
#
#
# def get_list():
#     url = 'https://ieeexplore.ieee.org/search/searchresult.jsp?queryText=IEEE%20Intelligent%20Systems&highlight=true&returnFacets=ALL&returnType=SEARCH&matchPubs=true&ranges=2010_2014_Year_Year&pageNumber={}'
#     pages = get_page_count()
#     print(pages)
#     page = 748
#     try:
#         while page <= pages:
#             print(url.format(page))
#             browser.get(url.format(page))
#             locator_link_list = (By.XPATH, '//*[@id="xplMainContent"]/div[2]/div[2]/xpl-results-list/div/xpl-results-item/div[1]/div[1]/div[2]/h2/a')
#             WebDriverWait(browser, 60, 0.5).until(expected_conditions.presence_of_element_located(locator_link_list))
#             link_list = browser.find_elements_by_xpath(
#                 '//*[@id="xplMainContent"]/div[2]/div[2]/xpl-results-list/div/xpl-results-item/div[1]/div[1]/div[2]/h2/a')
#             for link in link_list:
#                 ele_num = link.get_attribute('href')
#                 print(ele_num)
#                 herf_txt = open('IEEE Intelligent Systems.txt','a',encoding='utf-8')
#                 herf_txt.write(ele_num+'\n')
#                 herf_txt.close()
#                 #search(ele_num)
#             page += 1
#     except TimeoutException:
#         browser.close()
#         return get_list()
#
#
# def get_page_count():
#     try:
#         browser.get('https://ieeexplore.ieee.org/search/searchresult.jsp?queryText=IEEE%20Intelligent%20Systems&highlight=true&returnFacets=ALL&returnType=SEARCH&matchPubs=true&ranges=2010_2014_Year')
#         n = 0
#         locator = (By.XPATH, '//*[@id="xplMainContent"]/div[1]/div[2]/xpl-search-dashboard/section/div/div[1]/span')
#         WebDriverWait(browser, 30, 0.5).until(expected_conditions.presence_of_element_located(locator))
#         if browser.find_element_by_xpath(
#                 '//*[@id="xplMainContent"]/div[1]/div[2]/xpl-search-dashboard/section/div/div[1]/span'):
#             if browser.find_element_by_xpath(
#                     '//*[@id="xplMainContent"]/div[1]/div[2]/xpl-search-dashboard/section/div/div[1]/span').text != 'No results found':
#                 number = browser.find_element_by_xpath(
#                     '//*[@id="xplMainContent"]/div[1]/div[2]/xpl-search-dashboard/section/div/div[1]/span[1]/span[2]').text
#                 number = number.replace(",", '')
#                 n = int(number) / 25
#                 n = math.ceil(n)
#                 print(n)
#                 return n
#     except TimeoutException:
#         return get_page_count()
headers = {
    'User-agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.132 Safari/537.36',
}

#authors有时候最后一个多余，institution正则表达式有问题，keywords,abstract不全
def search(row,url):
    # try:
    #     browser.get(url)
    #     #WebDriverWait.until(browser,30)
    #     html = browser.page_source
    #     #print(html)
    #     soup = BeautifulSoup(html,'lxml')
    #     title = None
    #     abstract = None
    #     authors = None
    #     first_author = None
    #     PublishIn = None
    #     volume_list = None
    #     volume = None
    #     issue = None
    #     Date = None
    #     year = None
    #     month = None
    #     publisher = None
    #     keywords = None
    #     doi = None
    #     author_card = None
    #     institution = None
    #     if soup.find('h1', class_='document-title'):
    #         title = soup.find('h1', class_='document-title').text.strip()
    #         print(title)
    #     if soup.find('div', class_='abstract-text row'):
    #         abstract = soup.find('div', class_='abstract-text row').text.strip()
    #         abstract = re.sub(r'\n|\r|\t', '', abstract)
    #         abstract = re.sub('Abstract:', '', abstract)
    #         print(abstract)
    #     if soup.find('div', class_='authors-info-container overflow-ellipsis'):
    #         authors = soup.find('div', class_='authors-info-container overflow-ellipsis').text.strip()
    #         authors = re.sub(r'\n|\r|\t', '', authors)
    #         print(authors)
    #         list_author = authors.split(';')
    #         first_author = list_author[0]
    #         print(first_author)
    #         #u-pb-1 stats-document-abstract-publishedIn
    #     if soup.find('a', class_='u-pb-1 stats-document-abstract-publishedIn'):
    #         PublishIn = soup.find('div', class_='u-pb-1 stats-document-abstract-publishedIn').text.strip()
    #         PublishIn = re.sub(r'\n|\r|\t', '', PublishIn)
    #         PublishIn = re.sub('Published in: ', '', PublishIn)
    #         volume_list = re.findall(r"\b\d+\b", PublishIn)
    #         print(volume_list)
    #         volume = volume_list[0]
    #         issue = volume_list[1]
    #         print(''+PublishIn)
    #     if soup.find('div', class_='u-pb-1 doc-abstract-pubdate'):
    #         Date = soup.find('div', class_='u-pb-1 doc-abstract-pubdate').text.strip()
    #         Date = re.sub(r'\n|\r|\t', '', Date)
    #         Date = re.sub('Date of Publication:', '', Date)
    #         print(''+Date)
    #         month = ''.join(re.findall(r'[A-Za-z]', Date))
    #         print(month)
    #         Data_list = Date.split(' ')
    #         year = Data_list[len(Data_list)-1]
    #         print(year)
    #     if soup.find('div', class_='u-pb-1 doc-abstract-publisher publisher-info-container black-tooltip'):
    #         publisher = soup.find('div', class_='u-pb-1 doc-abstract-publisher publisher-info-container black-tooltip').text.strip()
    #         publisher = re.sub(r'\n|\r|\t', '', publisher)
    #         publisher = re.sub('Publisher:', '', publisher)
    #         print(publisher)
    #         #doc-keywords-list-item
    #     if soup.find('li', class_='doc-keywords-list-item'):
    #         keywords = soup.find('li', class_='doc-keywords-list-item').text
    #         keywords = re.sub(r'\n|\r|\t', '', keywords)
    #         keywords = keywords.replace('Keywords', 'Keywords:')
    #         keywords = re.sub('IEEE Keywords:', '', keywords)
    #         print(keywords)
    #     if soup.find('div', class_='u-pb-1 stats-document-abstract-doi'):
    #         doi = soup.find('div', class_='u-pb-1 stats-document-abstract-doi').text.strip()
    #         doi = re.sub(r'\n|\r|\t', '', doi)
    #         doi = re.sub('DOI:', '', doi)
    #         print(doi)
    #     if title != None:
    #         work_book = xlrd.open_workbook("IEEE.xls")
    #         sheet = work_book.sheet_by_name('TransactionsCommunications')
    #         rows = sheet.nrows
    #         new_workbook = copy(work_book)
    #         new_worksheet = new_workbook.get_sheet('TransactionsCommunications')
    #         new_worksheet.write(rows, 0, title)
    #         new_worksheet.write(rows, 1, authors)
    #         #Institutions_1
    #         new_worksheet.write(rows, 2, first_author)
    #         new_worksheet.write(rows, 3, PublishIn)
    #         new_worksheet.write(rows, 4, volume)
    #         new_worksheet.write(rows, 5, issue)
    #         new_worksheet.write(rows, 6, Date)
    #         new_worksheet.write(rows, 7, year)
    #         new_worksheet.write(rows, 8, month)
    #         new_worksheet.write(rows, 9, publisher)
    #         new_worksheet.write(rows, 10, keywords)
    #         new_worksheet.write(rows, 11, abstract)
    #         new_worksheet.write(rows, 12, doi)
    #         new_workbook.save("IEEE.xls")
    # except TimeoutException:
    #     return search(row,url)
    response = requests.get(url=url, headers=headers)
    # print(response)
    # response.json().get('doi')#
    html = response.text
    #print(html)
    pattern = 'xplGlobal.document.metadata.*'
    # # res = re.match(pattern,html)
    # # print(res.group())
    # html = response.content.decode('utf-8')
    # print(html)
    res = re.search('xplGlobal.document.metadata.*', html)
    title = None
    authors = ''
    first_author = None
    institution = None
    doi = None
    date = None
    year = None
    month = None
    published_in = None
    volume = None
    issue = None
    publisher = None
    keywords = None
    abstract = None
    published_date = None
    published_year = None
    published_month = None
    if res:
        data = res.group()
        print(data)
        print('-' * 100)
        doi_str = re.search(r'"doi":"[\w\s.-/]+[^"]', data)
        if doi_str:
            doi = re.sub(r'doi":"', '', doi_str.group())
            print('doi: ' + doi)
        title_str = re.search(r'displayDocTitle":"[\w][^"]+', data)
        if title_str:
            print(title_str.group())
            title = re.sub(r'displayDocTitle":"', '', title_str.group())
            print('title: ' + title)
        authors_list = re.findall(r'"name":".+,"affiliation"', data)
        if len(authors_list) > 0:
            authors_str = re.findall(r'"name":"[\w\s.-]+[^",]', authors_list[0])
            print(authors_str)
            for name in authors_str:
                name1 = re.sub(r'"name":"', '', name)
                authors = authors + name1 + ';'
            print(authors)
            first_author = re.sub(r'"name":"', '', authors_str[0])
            print(first_author)
        institution_str = re.search(r'"affiliation":."[\w][^"]+',data)
        if institution_str:
            institution_str1 = re.search(r'[A-Z][^"]+', institution_str.group())
            if institution_str1:
                institution = institution_str1.group()
                print('institution: ' + institution)
        online_date = re.search(r'"journalDisplayDateOfPublication":"\d.{0,30}\d|"onlineDate":"\d.{0,30}\d|"conferenceDate":"\d.{0,30}\d|"dateOfInsertion":"\d[\w\s]{0,30}', data)
        #online_date = re.search(r'journalDisplayDateOfPublication":"\d[\w\s]{0,30}\d',data)
        if online_date:
            print(online_date)
            date = re.sub(r'"journalDisplayDateOfPublication":"|"onlineDate":"|"conferenceDate":"|"dateOfInsertion":"', '', online_date.group())
            #date_group = re.search(r'\d[\w\s]+[-]?[.]?\d',online_date.group())
            #date = re.sub(r'"journalDisplayDateOfPublication":"','',online_date.group())
            # if date_group:
            #     date = date_group.group()
            print('date: ' + date)
            year_str = re.search(r'20\d\d',date)
            if year_str:
                year = year_str.group()
                print('year: ' + year)
            month_str = re.search(r'[A-Z][\w.]+',date)
            if month_str:
                month = month_str.group()
                print('month: ' + month)
        publication_title = re.search(r'"publicationTitle":"[A-Z][\w\s]+[a-z]', data)
        if publication_title:
            published_in = re.sub(r'"publicationTitle":"', '', publication_title.group())
            print('published_in: ' + published_in)
        volume_str = re.search(r'"volume":"\d{0,3}', data)
        if volume_str:
            volume = re.sub(r'"volume":"', '', volume_str.group())
            print('volume: ' + volume)
        issue_str = re.search(r'"issue":"\d{0,3}', data)
        if issue_str:
            issue = re.sub(r'"issue":"', '', issue_str.group())
            print('issue: ' + issue)
        published_date_str = re.search(r'"publicationDate":"[\w\s.,-/]+',data)
        if published_date_str:
            published_date = re.sub(r'"publicationDate":"','',published_date_str.group())
            print("published_date:" + published_date)
            published_month_str = re.search(r'[A-Za-z.-]+',published_date)
            if published_month_str:
                published_month = published_month_str.group()
                print("published_month:" + published_month)
            published_year = re.search(r'20\d\d',published_date).group()
            print("published_year:"+published_year)

        publisher_str = re.search(r'"publisher":"[A-Z][\w\s]+', data)
        if publisher_str:
            publisher = re.sub(r'"publisher":"', '', publisher_str.group())
            print('publisher: ' + publisher)
        keywords_str = re.search(r'"kwd":."[\w]+[^]]+', data)
        if keywords_str:
            # keywords = keywords_str.group()
            print(keywords_str.group())
            #keywords1 = re.search(r'[A-Z][^]]+', keywords_str.group())
            # print(keywords1.group())
            keywords = re.sub(r'"kwd":."', '', keywords_str.group())
            keywords = re.sub(r'"','',keywords)
            print('keywords: ' + keywords)
        abstract_str = re.search(r'"abstract":"[A-Z][^"]*', data)
        if abstract_str:
            abstract = re.sub(r'"abstract":"', '', abstract_str.group())
            print('abstract: ' + abstract)
        print('*' * 100)
        work_book = xlrd.open_workbook("IEEE.xls")
        sheet = work_book.sheet_by_name('TransactionsCommunications')
        rows = sheet.nrows
        new_workbook = copy(work_book)
        new_worksheet = new_workbook.get_sheet('TransactionsCommunications')
        new_worksheet.write(rows, 0, title)
        new_worksheet.write(rows, 1, authors)
        new_worksheet.write(rows, 2, first_author)
        new_worksheet.write(rows, 3, institution)
        new_worksheet.write(rows, 4, published_in)
        new_worksheet.write(rows, 5, volume)
        new_worksheet.write(rows, 6, issue)
        new_worksheet.write(rows, 7, published_date)
        new_worksheet.write(rows, 8, published_year)
        new_worksheet.write(rows, 9, published_month)
        new_worksheet.write(rows, 10, date)
        new_worksheet.write(rows, 11, year)
        new_worksheet.write(rows, 12, month)
        new_worksheet.write(rows, 13, publisher)
        new_worksheet.write(rows, 14, keywords)
        new_worksheet.write(rows, 15, abstract)
        new_worksheet.write(rows, 16, doi)
        new_workbook.save("IEEE.xls")
    else:
        print('no')


def get_one_page(url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            return response.text
        return None
    except RequestException:
        return None


def save_to_excel():
    urls = []
    row = 0
    for url in open('1.txt'):
        if url.strip('\n') != '':
            urls.append(url.strip('\n') + 'keywords')
            search(row,urls[row])
            row += 1


def main():
    save_to_excel()
    #get_list()
    # scExcel = pd.read_excel('new.xls')
    # scExcel.sort_values(by='Year',ascending=True)
    # scExcel.to_excel('hello.xls')


if __name__ == '__main__':
    main()

