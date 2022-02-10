from dependency import *

data = r'D:\DataAnalytics\database\is.xlsx'

dfs = pd.read_excel(data, sheet_name='Sheet1')
a = dfs.columns.tolist()
col_name = []
for i in a:
    if 'Consolidated Year' in i:
        i = i.replace('Consolidated Year: 2018 Quarter: 1 Unit: VND', '')
        i = i.split(' ')
        del i[0]
        listToStr = '_'.join([str(elem).lower() for elem in i])[:-1]
        if listToStr[-1] == '_':
            listToStr = listToStr[:-1]
        if ',' in listToStr:
            listToStr = listToStr.replace(',', '_&')
        if '-' in listToStr:
            listToStr =  listToStr.replace('-', '_')
        listToStr = f'[{listToStr}] bigint,'
        # i = i.split(' ')
        col_name.append(listToStr)



