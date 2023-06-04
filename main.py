import PySimpleGUI as sg
from bs4 import BeautifulSoup
import requests
import time
import matplotlib.pyplot as plt
import pandas


html_text = requests.get(
    'https://www.etsy.com/search/jewelry/bracelets?ref=cat_hobby_1193_2&explicit=1&q=Yoga').text
soup = BeautifulSoup(html_text, 'lxml')

mydict = {}
matrix = []


sg.theme('DarkAmber')

layout = [
    [sg.Text("Excel name"), sg.Input(key='_IN_')],
    [sg.B("Retrieve data", key='Retrieve data')],
    [sg.B("Create graph", key='Create graph')],
    [sg.B("Display the matrix", key='Display the matrix')],
    [sg.B("Save to excel file", key='FolderBrowse')],
]

window = sg.Window('Viewer of an album of images',
                   layout, element_justification='c')

while True:
    event, values = window.read()

    print(event, values)

    if event in (None, '_exit_'):
        break

    if event == "Retrieve data":

        time.sleep(0.2)
        products = soup.find_all(
            'li', class_='wt-list-unstyled wt-grid__item-xs-6 wt-grid__item-md-4 wt-grid__item-lg-3 wt-order-xs-0 wt-order-md-0 wt-order-lg-0 wt-show-xs wt-show-md wt-show-lg')

        for product in products:
            names = product.find(
                'h3', class_="wt-text-caption v2-listing-card__title wt-text-truncate")
            prices = product.find('span', class_='currency-value')

            if names is not None and prices is not None:
                name = names.text.strip()
                price = prices.text.strip()

            mydict[name.title()] = float(price)
            time.sleep(0.2)

        matrix = [[key, values]
                  for key, values in mydict.items()]

    if event == "Create graph":

        if bool(mydict):
            names = list(mydict.keys())
            prices = list(mydict.values())
            plt.barh(range(len(mydict)), prices,
                     tick_label=names)
            plt.show()
        else:
            sg.popup_error("Please retrieve the data first.")

    if event == "Display the matrix":
        try:
            sg.popup_scrolled(
                'Matrix of products and their prices', str(matrix))
        except UnboundLocalError:
            sg.popup_error('Please retrieve the data first.')

    if event == "FolderBrowse":

        if values['_IN_'] and bool(mydict):
            folder = sg.popup_get_folder(
                'Select a folder to save the Excel file')
            filename = values['_IN_']

            df = pandas.DataFrame(mydict, index=['Price'])
            df = df.transpose()
            writer = pandas.ExcelWriter(
                folder + '/' + filename + '.xlsx', engine='xlsxwriter')
            worksheet = writer.book.add_worksheet('Sheet1')
            worksheet.write('A1', 'Product Name')

            df.to_excel(writer, index=True)

            writer._save()
        elif not values['_IN_']:
            sg.popup_error('Please insert a name for the excel file.')
        else:
            sg.popup_error('Please retrieve the data first.')

window.close()
