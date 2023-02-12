from bs4 import BeautifulSoup
import requests
import random
from random_user_agent.user_agent import UserAgent
from random_user_agent.params import SoftwareName, OperatingSystem
import xlsxwriter


# Initializes variables
linkPart = ""
masterList = []
col = 0
pgNum = 1
linksList= []


# Creates random user agent
software_names = [SoftwareName.CHROME.value]
operating_systems = [OperatingSystem.WINDOWS.value, OperatingSystem.LINUX.value] 
user_agent_rotator = UserAgent(software_names=software_names, operating_systems=operating_systems, limit=100)
randomUserAgent = user_agent_rotator.get_random_user_agent()


# Changes my headers
headers = {
    'User-Agent': randomUserAgent,
    'Content-Type': 'text/html',
}

# Creates link from input
print("ENTER ITEM:")
item = input().split(" ")
print("ENTER # OF AMAZON PAGES:")
page = int(input())
for count, i in enumerate(item):
     linkPart += i
     if count != len(item)-1:
          linkPart += "+"
for pgNum in range(1,page+1):
    link = "https://www.amazon.ca/s?k="+linkPart+"&page="+str(pgNum)
    linksList.append(link)


def scrapePage(link):
    # Error detection
    if(requests.get(link, headers = headers).status_code != 200):
            print("ERROR")


    # Gets html file from link
    html_text = requests.get(link, headers = headers).text
    soup = BeautifulSoup(html_text, "lxml")


    # Searches for product information
    productInfo = soup.find_all('div',{'data-component-type': 's-search-result'})
    for product in productInfo:
        if product.find("span", class_="a-size-base-plus a-color-base a-text-normal") is not None: 
            name = product.find("span", class_="a-size-base-plus a-color-base a-text-normal").text
        if product.find("span", class_="a-offscreen") is not None:
            price = product.find("span", class_="a-offscreen").text.replace("$", "")
        ratingSub = product.find("div", class_="s-card-container s-overflow-hidden aok-relative puis-expand-height puis-include-content-margin puis s-latency-cf-section s-card-border")
        if ratingSub.find("span", class_="a-icon-alt") is not None: 
            rating = ratingSub.find("span", class_="a-icon-alt").text[:3]
        else:
            rating = "N/A"
        if ratingSub.find("span", class_="a-size-base s-underline-text") is not None:
            reviews = ratingSub.find("span", class_="a-size-base s-underline-text").text.replace(",", "")
        else:
            reviews = "N/A"
        if reviews[0] == "(":
            reviews = reviews[1:-1]
        temp = [name, price, rating, reviews]
        masterList.append(temp)
    
for currentLink in linksList:
    scrapePage(currentLink)

# Creates xlsx file
workbook = xlsxwriter.Workbook('results.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write_row(0, 0, ["Product Name", "Price (USD)", "Rating (out of 5)", "Number of Reviews"])
for row, data in enumerate(masterList):
    worksheet.write_row(row+1, col, data)
workbook.close()
