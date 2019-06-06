import xlrd, pandas as pd, re, csv
from os import walk
from xlutils.copy import copy


dir = walk('//itimpf01/EXCHANGE/AndreaAmoretti/Packing Lists/All')

# =========================================================================================== #
# Reading dictionaries in python from itemsSpace and fileNameSpace
file = open('//itimpf01/EXCHANGE/AndreaAmoretti/Dizionario.csv', mode='r')
reader = csv.reader(file, delimiter =';')
dictionary = {}
for row in reader:
    key = row[0]
    key = key.replace(' ', '')
    dictionary[key] = row[1]
file.close()

file2 = open('//itimpf01/EXCHANGE/AndreaAmoretti/DizionarioFileNames.csv', mode='r')
reader2 = csv.reader(file2, delimiter =';')
dictFileNames = {}
for row2 in reader2:
    key2 = row2[0]
    key2 = key2.replace(' ', '')
    dictFileNames[key2] = row2[1]
file2.close()

# =========================================================================================== #
# Standardizing Packing Lists wordings
for (dirPath, dirs, files) in dir:
    for file in files:
        if file.endswith(".xls"):
            path = dirPath + '/' + file
            wb = xlrd.open_workbook(path, formatting_info=True)
            new_wb = copy(wb)
            new_sheet = new_wb.get_sheet(0)
            first_sheet = wb.sheet_by_index(0)
            rows = first_sheet.nrows
            for row in range(5, rows):
                value = first_sheet.cell(row, 2).value
                if '+' in value:
                    item = value.split('+')
                    quantities = []
                    items = []

                    for i in item:

                        prod = re.sub(r'\d+', '', i)
                        prod = prod.replace(' ', '')
                        items.append(prod)

                        qty = re.search(r'\d+', str(i))
                        if qty is None:
                            quantities.append('0')
                        else:
                            qty = qty.group(0)
                            quantities.append(qty)

                else:
                    quantities = []
                    items = []
                    prod = re.sub(r'\d+', '', value)
                    prod = prod.replace(' ', '')
                    items.append(prod)

                    qty = re.search(r'\d+', str(value))
                    if qty is None:
                        quantities.append('')
                    else:
                        qty = qty.group(0)
                        quantities.append(qty)

                items = [dictionary.get(x, x) for x in items]
                new_items = [str(a) + ' ' + b for a, b in zip(quantities, items)]
                new_items = ' + '.join(new_items)
                new_items = new_items.replace('+ 0', ' +')
                new_sheet.write(row, 2, new_items)
                new_sheet.write(row, 3, '')
                new_sheet.write(row, 4, '')
                new_sheet.write(row, 5, '')

            new_sheet.write(1, 3, 'ORD')
            new_sheet.write(2, 3, 'CLIENTE')

            file = file.lower()
            pQty = re.search(r'\d+', str(file))
            product = file[pQty.span()[1]:]
            product = product.strip()
            item2 = product.split(' ')
            items2 = [dictFileNames.get(x, x) for x in item2]
            new_items2 = pQty.group() + ' ' + ' '.join(items2).strip() + '.xls'

            new_path = 'C:/Users/itimpalfe/OneDrive - munters.com/OldStuff/Computer/WIP/Test/' + str(new_items2)
            new_wb.save(new_path)

# =========================================================================================== #
# Create Item Space
itemSpace = pd.DataFrame()
for (dirPath, dirs, files) in dir:
    for file in files:
        if file.endswith(".xls"):
            path = dirPath + '/' + file
            wb = xlrd.open_workbook(path)
            first_sheet = wb.sheet_by_index(0)
            rows = first_sheet.nrows
            for row in range(5, rows):
                value = first_sheet.cell(row, 2).value
                value = re.sub(r'\d+', '', value)
                value = value.strip()
                if '+' in value:
                    item = value.split('+')
                    itemSpace = itemSpace.append(item, ignore_index=True)
                else:
                    itemSpace = itemSpace.append([value], ignore_index=True)
            print("Working on: " + file)

itemSpace.to_csv('C:/Users/itimpalfe/OneDrive - munters.com/SideProjects/PackingList/Input/Dizionario2.csv', sep=';',
                 index_label=False)


# =========================================================================================== #
# Create FileName Space
fileNameSpace = pd.DataFrame()
for (dirPath, dirs, files) in dir:
    for file in files:
        if file.endswith(".xls"):
            qty = re.search(r'\d+', str(file))
            product = file[qty.span()[1]:]
            item = product.split(' ')
            fileNameSpace = fileNameSpace.append(item, ignore_index=True)
            print('Working on: ' + file)
fileNameSpace.to_csv('//Itimpf01/exchange/AndreaAmoretti/DizionarioFileNames.csv', sep=';', index_label=False)

# =========================================================================================== #
# Second pass to adjust filename
dir = walk('C:/Users/itimpalfe/OneDrive - munters.com/SideProjects/PackingList/PackingLists_StdFilename')

for (dirPath, dirs, files) in dir:
    for file in files:
        if file.endswith(".xls"):
            print('Working on: ' + file)
            path = dirPath + '/' + file
            wb = xlrd.open_workbook(path, formatting_info=True)
            new_wb = copy(wb)
            first_sheet = wb.sheet_by_index(0)
            rows = first_sheet.nrows
            items = []
            for row in range(5, rows):
                print(row)
                value = first_sheet.cell(row, 2).value
                items.append(value)

        items = ''.join(items)
        items = items.replace(' ', '')
        items = items.lower()

        filename = file.replace('.xls', '')

        if 'motori' not in items:
            filename = filename + ' NoMot'

        if 'kitprotplastica' or 'kitplastica' in items:
            filename = filename + ' Pl'

        if 'retipiramidali' in items:
            filename = filename + ' Pm'


        path = 'C:/Users/itimpalfe/OneDrive - munters.com/OldStuff/Computer/WIP/Test/' + filename +\
               ' Smo' + '.xls'
        new_wb.save(path)















































