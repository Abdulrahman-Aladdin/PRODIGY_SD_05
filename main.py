import requests
from bs4 import BeautifulSoup as Bs
import pickle
import re
import json
import xlsxwriter


'''
I used this code on different stages
first -> used `get_data` method to get data from the site, it was not easy since I had to search for the request that
fetches the data I want, and find the required headers for it. then I stored the response data in a pickle file using
`write_data` method.

second -> used `load_data` that reads the saved data, and then I processed it to extract information I want, then save
it to an excel sheet.
'''


def write_data(data):
    with open('data.pickle', 'wb') as file:
        pickle.dump(data, file)


def load_data():
    with open('data.pickle', 'rb') as file:
        return pickle.load(file)


def get_data(src_link):

    # You must replace `SECRET DATA` with the appropriate data on your machine
    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'en-US,en;q=0.5',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Cookie': 'SECRET DATA',
        'Host': 'www.ubuy.com.eg',
        'Pragma': 'no-cache',
        'Referer': 'https://www.ubuy.com.eg/en/category/electronics/laptops/all-in-ones-13896603011',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1',
        'TE': 'trailers',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'SECRET DATA'
    }
    return Bs(requests.get(src_link, headers=headers).content, 'lxml')


def write_data_to_excel(headers, data, file_path, file_name):
    workbook = xlsxwriter.Workbook(file_path + file_name)
    worksheet = workbook.add_worksheet()

    header_format = workbook.add_format({'bold': True})
    worksheet.write_row(0, 0, headers, header_format)

    for i in range(1, len(data) + 1):
        worksheet.write_row(i, 0, data[i - 1])
    workbook.close()


link = ('https://www.ubuy.com.eg/en/ubcommon/mongo/search/products?ubuy=es1&docType=offRHF&q=&node_id=13896603011&'
        'page=1&brand=&ufulfilled=&price_range=&sort_by=&s_id=11&lang=&dc=&search_type=category&skus=&store=us')

# Use the next two lines to fetch the data from the web-site
# soup = get_data(link)
# write_data(soup)

soup = load_data()
wanted_text = soup.find_all('script')[1].text
match = re.search(r'impressionProducts\s*=\s*(.*?);', wanted_text)
products_list = []
if match:
    products_list_txt = match.group(1)
    products_list = json.loads(products_list_txt)

products_data = []
for product in products_list:
    products_data.append([product['name'], 'EGP ' + str(product['price']), product['category']])

write_data_to_excel(['name', 'price', 'store'], products_data, '', 'extracted_products.xlsx')
print('DONE!')
