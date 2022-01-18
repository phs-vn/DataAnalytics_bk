from request_phs.stock import *


def run():
    dictionary = pd.read_csv(
        r'D:\DataAnalytics\News Filter Project\Vietnamese vocab\financial dictionary\viet_dict.txt',
        delimiter='\t',
        encoding='utf-8'
    )
    dictionary.columns = ['vie_dict']
    vie_dict = dictionary['vie_dict'].tolist()
    vie_dict = [' {0} '.format(elem) for elem in vie_dict]

    data = pd.read_pickle(r"D:\DataAnalytics\News Filter Project\output_data\vietstock_data"
                          r"\vietstock_kinh-te_data.pickle")
    content = data['Nội dung'].tolist()
    save_dict = {}
    idx = 0

    while True:
        for i in vie_dict:
            count = content[idx].count(i)
            if count > 0:
                save_dict.update({i: count})
        idx += 1
        print(idx)
        if idx > len(content) - 1:
            break

    data_items = save_dict.items()
    data_list = list(data_items)

    df = pd.DataFrame(data_list, columns=['word', 'times'])
    df.to_pickle("./News Filter Project/count_word_output/vietstock_kinh-te_count.pickle")
