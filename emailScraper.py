import re
import requests
import requests.exceptions
from urllib.parse import urlsplit
from collections import deque
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook
import xlwt
import xlrd
from xlutils.copy import copy

#load excel file
rb = xlrd.open_workbook('emailsCompany.xls')

result_sheet = rb.sheet_by_index(0)

wb = copy(rb) 
sheet = wb.get_sheet(0) 

companyLinks = []
companyNames = []

excel_data_df = pd.read_excel('linksCompany.xlsx', sheet_name=0)

print(excel_data_df.columns.ravel())

companyNames = excel_data_df['CompanyName'].tolist()
companyLinks = excel_data_df['CompanyLink'].tolist()
my_timeout = 3
counter = 0

  

# starting url. replace google with your own url.
starting_url = 'https://www.firstinspires.org/'

# a queue of urls to be crawled
unprocessed_urls = deque([starting_url])

# set of already crawled urls for email
processed_urls = set()

# a set of fetched emails
emails = set()

# process urls one by one from unprocessed_url queue until queue is empty
def findEmail():
    for i in range(10):
        try:
            # move next url from the queue to the set of processed urls
            url = unprocessed_urls.popleft()
            processed_urls.add(url)

            # extract base url to resolve relative links
            parts = urlsplit(url)
            base_url = "{0.scheme}://{0.netloc}".format(parts)
            path = url[:url.rfind('/')+1] if '/' in parts.path else url

            # get url's content
            print("Crawling URL %s" % url)
            try:
                response = requests.get(url, timeout=my_timeout)
            except (requests.exceptions.Timeout, requests.exceptions.ConnectionError, requests.exceptions.MissingSchema, requests.exceptions.ConnectionError):
                # ignore pages with errors and continue with next url
                continue

            # extract all email addresses and add them into the resulting set
            # You may edit the regular expression as per your requirement
            new_emails = set(re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.com+", response.text, re.I))
            emails.update(new_emails)
            print(emails)
            # create a beutiful soup for the html document
            soup = BeautifulSoup(response.text, 'lxml')

            # Once this document is parsed and processed, now find and process all the anchors i.e. linked urls in this document
            for anchor in soup.find_all("a"):
                # extract link url from the anchor
                link = anchor.attrs["href"] if "href" in anchor.attrs else ''
                # resolve relative links (starting with /)
                if link.startswith('/'):
                    link = base_url + link
                elif not link.startswith('http'):
                    link = path + link
                
                # add the new url to the queue if it was not in unprocessed list nor in processed list yet
                if not link in unprocessed_urls and not link in processed_urls:
                    unprocessed_urls.append(link)
        except:
            print(emails)
            break

for link in companyLinks:
    starting_url = str(link)
    if not link.startswith("http://"):
                    starting_url = "http://" + link
    unprocessed_urls = deque([starting_url])
    unprocessed_urls = deque([starting_url])
    companyName = companyNames[counter]
    # set of already crawled urls for email
    processed_urls = set()
    # set of already crawled urls for email
    findEmail()

    for email in emails:
        # row number = 0 , column number = 1
        sheet.write(counter+1,0, companyName)
        sheet.write(counter+1,1, email)
        wb.save('my_workbook.xls')
        
    
    
    counter += 1
    emails = set()

