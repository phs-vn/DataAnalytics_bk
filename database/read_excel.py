from request import *

path_bank_bs = r'D:\DataAnalytics\database\fs_bank\bs_2018q2.xlsm'
txt_bs = r'D:\DataAnalytics\database\txt_col_name\fs_bank\fs_bank_bs.txt'


########################################################################################################################
########################################################################################################################


def read_excel(excel_path: str, col_name_path: str) -> pd.DataFrame():
    df = pd.read_excel(
        excel_path,
        sheet_name='ExportData',
        skiprows=7,
    )
    df = df.iloc[1:, 1:]
    # loại những dòng nào mà giá trị ở cả 3 cột đều là NaN
    df = df.dropna(how='all', subset=['Ticker', 'Name', 'Exchange'])
    df = df.fillna(0)
    if len((excel_path.split('\\')[-1]).split('_')[0]) == 2:
        year = int(excel_path.split('\\')[-1][3:7])
        quarter = int(excel_path.split('\\')[-1][8:9])
    else:
        year = int(excel_path.split('\\')[-1][4:8])
        quarter = int(excel_path.split('\\')[-1][9:10])
    old_name = []
    new_name = []
    # process and set name for column in df
    for col in df.columns:
        str_del = f'Consolidated Year: {year} Quarter: {quarter} Unit: VND'
        col = col.split(str_del)[0]
        if len(col.split()) < 2:
            col = col.lower()
            old_name.append(col)
        else:
            col = ''.join(col.split()[1:])
            col = re.sub("[^a-zA-Z]", '', col).lower()
            if col in old_name:
                col = col + '_'
                old_name.append(col)
            else:
                old_name.append(col)
    df.columns = old_name

    # read file txt has new col name from SQL script
    f = open(col_name_path, 'r')
    lst_tmp = f.read().split(',')
    for col in lst_tmp:
        col = re.sub("[\n]", '', col).split()[0]
        col = col.replace('[', '').replace(']', '')
        new_name.append(col)

    col_name_dict = {key:val for key,val in zip(old_name,new_name)}
    df.columns = df.columns.to_series().map(col_name_dict)

    df.insert(0, column='consolidated_year', value=year)
    df.insert(1, column='quarter', value=quarter)

    return df
