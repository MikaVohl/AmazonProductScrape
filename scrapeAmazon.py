from bs4 import BeautifulSoup
import requests
import random
from random_user_agent.user_agent import UserAgent
from random_user_agent.params import SoftwareName, OperatingSystem
import xlsxwriter
import tkinter as tk
from tkinter import filedialog

# Initializes variables
masterList = []

# Creates random user agent
def randomAgent():
    software_names = [SoftwareName.CHROME.value]
    operating_systems = [OperatingSystem.WINDOWS.value, OperatingSystem.LINUX.value] 
    user_agent_rotator = UserAgent(software_names=software_names, operating_systems=operating_systems, limit=100)
    return user_agent_rotator.get_random_user_agent()

# Changes my headers
headers = {
    'User-Agent': randomAgent(),
    'Content-Type': 'text/html',
}

def scrapePage(link):
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
def save_text_to_file():
    # GUI prompts and information
    item = text_entry.get().split(" ") # input prompt
    page = int(pages_entry.get()) # input number of pages
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
    text_entry.delete(0, tk.END) # emptys the prompt
    pages_entry.delete(0, tk.END) # emptys the prompt

    # initialize variables
    linkPart = ""
    pgNum = 1
    linksList = []

    # create link
    for count, i in enumerate(item):
        linkPart += i
        if count != len(item)-1:
            linkPart += "+"

    # create link array
    for pgNum in range(1,page+1):
        link = "https://www.amazon.ca/s?k="+linkPart+"&page="+str(pgNum)
        linksList.append(link)

    # iterate through links and scrape
    for currentLink in linksList:
        scrapePage(currentLink)

    # Creates xlsx file
    col = 0
    workbook = xlsxwriter.Workbook(file_path)
    worksheet = workbook.add_worksheet()
    worksheet.write_row(0, 0, ["Product Name", "Price (USD)", "Rating (out of 5)", "Number of Reviews"])
    for row, data in enumerate(masterList):
        worksheet.write_row(row+1, col, data)
    workbook.close()

# Create GUI
root = tk.Tk()
root.title("Find Amazon Product Information")

# Create / style input boxes
text_label = tk.Label(root, text="Enter Product:", font=("Lato", 14))
text_label.pack(pady=20)
text_entry = tk.Entry(root, width=50, font=("Helvetica", 14)) # creates input box
text_entry.pack(pady=20) # styles / makes input box visible
pages_label = tk.Label(root, text="Enter Number of Pages to Search:", font=("Lato", 14))
pages_label.pack(pady=20)
pages_entry = tk.Entry(root, width=50, font=("Helvetica", 14)) # creates input box
pages_entry.pack(pady=20) # styles / makes input box visible

# Create / style save button
save_button = tk.Button(root, text="Save Text to File", font=("Helvetica", 14), command=save_text_to_file)
save_button.pack(pady=20)

root.state('zoomed')

# Run the GUI
root.mainloop()
