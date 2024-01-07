import csv
import os
import time

from openpyxl import Workbook, load_workbook
import requests
from bs4 import BeautifulSoup
import re


def GetMonitorURLs(startPage, endPage):
    page = startPage
    monitorURLs = []
    while page <= endPage:
        productsPerPage = 0
        url = f'https://setec.mk/index.php?route=product/category&path=10019_10031&limit=100&page={page}'
        response = requests.get(url)
        if response.status_code != 200:
            continue
        soup = BeautifulSoup(response.text, "html.parser")
        monitorAnchorTags = soup.select("div.product > div.right > div.name > a")
        for tag in monitorAnchorTags:
            monitorURLs.append(tag.get('href'))
            productsPerPage += 1
        if productsPerPage <= 99:
            break
        page += 1
    return monitorURLs


def SaveMonitorToCSV(monitor, path):
    with open(path, 'a', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=['Name', 'Price(МКД)', 'In stock', 'Panel size(inches)', 'Horizontal resolution', 'Vertical resolution','Refresh rate(Hz)','URL'])
        writer.writerow(monitor)




def GetMonitorData(url: str) -> dict[str, str]:
    monitor: dict[str, str] = {}
    response = requests.get(url)
    if response.status_code != 200:
        return {}
    soup = BeautifulSoup(response.text, "html.parser")

    # Extracting relevant information
    panel_size = ""
    horizontal_res = ""
    vertical_res = ""
    refresh_rate = ""

    # Regular expressions for matching patterns
    size_pattern = re.compile(r'(\d+(\.\d+)?)\s?(inch|"|\'|nch|”)', re.IGNORECASE)
    resolution_pattern = re.compile(r'\b\d+\s?[xX]\s?\d+\b', re.IGNORECASE)
    refresh_rate_pattern = re.compile(r'\b\d+\s?Hz\b', re.IGNORECASE)

    div_content = soup.find('div', {'id': 'tab-description'}).get_text()

    # Find matches for each pattern
    size_matches = size_pattern.findall(div_content)
    resolution_matches = resolution_pattern.findall(div_content)
    refresh_rate_matches = refresh_rate_pattern.findall(div_content)

    # Extracting specific information from matches
    if size_matches:
        panel_size = size_matches[0][0] + " inch"

    if resolution_matches:
        if "X" in resolution_matches[0]:
            parts = resolution_matches[0].split('X')
        else:
            parts = resolution_matches[0].split('x')
        horizontal_res = parts[0].strip()
        vertical_res = parts[1].strip()

    if refresh_rate_matches:
        refresh_rate = refresh_rate_matches[0].replace("Hz", "").replace("hz","").strip()

    monitor["Name"] = soup.find('h1', {'id': 'title-page'}).text.strip()
    monitor["Price(МКД)"] = soup.select_one("div.price > span").text.replace("Ден.", "").replace(",", "").strip().replace("'","")
    imgSource = soup.find('div', class_='description').find('img').get('src')
    if imgSource == "image/no.png":
        monitor["In stock"] = "0"
    else:
        monitor["In stock"] = "1"
    monitor["Panel size(inches)"] = panel_size.replace("inch","").strip().replace("'","")
    if horizontal_res == "" and vertical_res == "":
        vertical_res = "0"
        horizontal_res = "0"
    if horizontal_res == "920" and vertical_res == "1":
        horizontal_res = "1920"
        vertical_res = "1080"
    if horizontal_res == "560" and vertical_res == "1":
        horizontal_res = "2560"
        vertical_res = "1440"
    if horizontal_res == "4" and vertical_res == "392":
        horizontal_res = "1920"
        vertical_res = "1080"
    if horizontal_res == "840" and vertical_res == "2":
        horizontal_res = "3840"
        vertical_res = "2160"
    if horizontal_res == "344" and vertical_res == "392":
        horizontal_res = "2560"
        vertical_res = "1440"
    if horizontal_res == "306" and vertical_res == "392":
        horizontal_res = "3840"
        vertical_res = "2160"
    if horizontal_res == "888" and vertical_res == "336":
        horizontal_res = "1920"
        vertical_res = "1080"
    if horizontal_res == "112" and vertical_res == "392":
        horizontal_res = "3840"
        vertical_res = "2160"
    horizontal_res.replace("'","")
    vertical_res.replace("'","")
    refresh_rate.replace("'","")
    panel_size.replace("'","")
    monitor["Horizontal resolution"] = horizontal_res
    monitor["Vertical resolution"] = vertical_res
    if refresh_rate == "":
        refresh_rate = "60"
    monitor["Refresh rate(Hz)"] = refresh_rate
    monitor["URL"] = url
    return monitor


def main():
    start = time.time()
    print("Retrieving monitor urls from https://setec.mk/index.php?route=product/category&path=10019_10031")
    monitorURLs = GetMonitorURLs(1, 6)
    print("Getting monitor data...")
    monitors = []
    count = 1
    for url in monitorURLs:
        monitor = GetMonitorData(url)
        if monitor == {}:
            continue
        monitors.append(monitor)
        if count % 20 == 0:
            print(f"Done {count} out of {len(monitorURLs)}")
        count += 1
    print("Finished scraping")
    print("Saving data...")
    csv_file_path = os.getcwd() + "\\data.csv"
    if not os.path.isfile(csv_file_path):
        with open(csv_file_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=['Name', 'Price(МКД)', 'In stock', 'Panel size(inches)', 'Horizontal resolution', 'Vertical resolution','Refresh rate(Hz)','URL'])
            writer.writeheader()
    print("Saving data to CSV " + csv_file_path)
    for monitor in monitors:
        SaveMonitorToCSV(monitor, csv_file_path)
    print(f"Finished in {(time.time() - start):.3f} seconds")


if __name__ == '__main__':
    main()
