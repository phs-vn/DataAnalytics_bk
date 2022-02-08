from request import *


def run():
    start = time.time()
    read_vietstock_1 = pd.read_pickle(r"D:\DataAnalytics\news_analysis\output_data\vietstock_data"
                                      r"\vietstock_bat-dong-san_data.pickle")
    read_vietstock_2 = pd.read_pickle(r"D:\DataAnalytics\news_analysis\output_data\vietstock_data"
                                      r"\vietstock_chung-khoan_data.pickle")
    read_vietstock_3 = pd.read_pickle(r"D:\DataAnalytics\news_analysis\output_data\vietstock_data"
                                      r"\vietstock_doanh-nghiep_data.pickle")
    read_vietstock_4 = pd.read_pickle(r"D:\DataAnalytics\news_analysis\output_data\vietstock_data"
                                      r"\vietstock_kinh-te_data.pickle")
    read_vietstock_5 = pd.read_pickle(r"D:\DataAnalytics\news_analysis\output_data\vietstock_data"
                                      r"\vietstock_tai-chinh_data.pickle")

    vietstock_news = pd.concat(
        [read_vietstock_1, read_vietstock_2, read_vietstock_3, read_vietstock_4, read_vietstock_5],
        ignore_index=True)

    ###################################################
    ###################################################
    ###################################################

    file_name = f'news.xlsx'
    writer = pd.ExcelWriter(
        fr"D:\DataAnalytics\news_analysis\output_data\{file_name}",
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    )
    workbook = writer.book

    time_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'num_format': 'dd/mm/yy hh:mm'
        }
    )
    text_left_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_size': 10,
            'font_name': 'Times New Roman'
        }
    )
    text_right_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman'
        }
    )
    headers_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
        }
    )

    header = [
        'STT',
        'Thời gian',
        'Loại tin tức',
        'Tiêu đề',
        'Link',
        'Nội dung'
    ]

    vietstock_sheet = workbook.add_worksheet('vietstock')
    # Set Column Width and Row Height
    vietstock_sheet.set_column('A:A', 4)
    vietstock_sheet.set_column('B:B', 14)
    vietstock_sheet.set_column('C:C', 26.57)
    vietstock_sheet.set_column('D:D', 25)
    vietstock_sheet.set_column('E:E', 25)
    vietstock_sheet.set_column('F:F', 55)

    vietstock_sheet.write_row('A1', header, headers_format)
    vietstock_sheet.write_column(
        'A2',
        np.arange(vietstock_news.shape[0]) + 1,
        text_right_format
    )
    vietstock_sheet.write_column(
        'B2',
        vietstock_news['Thời gian'],
        time_format
    )
    vietstock_sheet.write_column(
        'C2',
        vietstock_news['Loại tin tức'],
        text_left_format
    )
    vietstock_sheet.write_column(
        'D2',
        vietstock_news['Tiêu đề'],
        text_left_format
    )
    vietstock_sheet.write_column(
        'E2',
        vietstock_news['Link'],
        text_left_format
    )
    vietstock_sheet.write_column(
        'F2',
        vietstock_news['Nội dung'],
        text_left_format
    )

    # ====================================================================================
    read_tnck_1 = pd.read_pickle(r"D:\DataAnalytics\news_analysis\output_data\tinnhanhchungkhoan_2"
                                 r"\tinnhanhCK_chung-khoan_data_2.pickle")
    read_tnck_2 = pd.read_pickle(r"D:\DataAnalytics\news_analysis\output_data\tinnhanhchungkhoan_2"
                                 r"\tinnhanhCK_doanh-nghiep_data_2.pickle")
    tnck_news = pd.concat([read_tnck_1, read_tnck_2], ignore_index=True)

    tnck_sheet = workbook.add_worksheet('tinnhanhck')
    header_2 = [
        'STT',
        'Link',
        'Tiêu đề',
        'Nội dung'
    ]
    # Set Column Width and Row Height
    tnck_sheet.set_column('A:A', 4)
    tnck_sheet.set_column('B:B', 25)
    tnck_sheet.set_column('C:C', 25)
    tnck_sheet.set_column('D:D', 55)

    tnck_sheet.write_row('A1', header_2, headers_format)
    tnck_sheet.write_column(
        'A2',
        np.arange(tnck_news.shape[0]) + 1,
        text_right_format
    )
    tnck_sheet.write_column(
        'B2',
        tnck_news['Link'],
        text_left_format
    )
    tnck_sheet.write_column(
        'C2',
        tnck_news['Tiêu đề'],
        text_left_format
    )
    tnck_sheet.write_column(
        'D2',
        tnck_news['Nội dung'],
        text_left_format
    )

    # ====================================================================================
    read_cafef_1 = pd.read_pickle(r"D:\DataAnalytics\news_analysis\output_data\cafef_data"
                                  r"\cafef_doanh-nghiep_data.pickle")
    read_cafef_2 = pd.read_pickle(r"D:\DataAnalytics\news_analysis\output_data\cafef_data"
                                  r"\cafef_thi-truong-chung-khoan_data.pickle")
    read_cafef_3 = pd.read_pickle(r"D:\DataAnalytics\news_analysis\output_data\cafef_data"
                                  r"\cafef_bat-dong-san_data.pickle")
    read_cafef_4 = pd.read_pickle(r"D:\DataAnalytics\news_analysis\output_data\cafef_data"
                                  r"\cafef_tai-chinh-ngan-hang_data.pickle")
    cafef_news = pd.concat([read_cafef_1, read_cafef_2, read_cafef_3, read_cafef_4], ignore_index=True)

    cafef_sheet = workbook.add_worksheet('cafef')
    header_3 = [
        'STT',
        'Link',
        'Tiêu đề',
        'Mô tả',
        'Nội dung'
    ]
    # Set Column Width and Row Height
    cafef_sheet.set_column('A:A', 4)
    cafef_sheet.set_column('B:B', 25)
    cafef_sheet.set_column('C:C', 28)
    cafef_sheet.set_column('D:D', 30)
    cafef_sheet.set_column('E:E', 55)

    cafef_sheet.write_row('A1', header_3, headers_format)
    cafef_sheet.write_column(
        'A2',
        np.arange(cafef_news.shape[0]) + 1,
        text_right_format
    )
    cafef_sheet.write_column(
        'B2',
        cafef_news['News link'],
        text_left_format
    )
    cafef_sheet.write_column(
        'C2',
        cafef_news['Tiêu đề'],
        text_left_format
    )
    cafef_sheet.write_column(
        'D2',
        cafef_news['Mô tả'],
        text_left_format
    )
    cafef_sheet.write_column(
        'E2',
        cafef_news['Nội dung'],
        text_left_format
    )

    # ====================================================================================
    read_ndh_1 = pd.read_pickle(r"D:\DataAnalytics\news_analysis\output_data\ndh"
                                r"\ndh_chung-khoan_data.pickle")
    read_ndh_2 = pd.read_pickle(r"D:\DataAnalytics\news_analysis\output_data\ndh"
                                r"\ndh_doanh-nghiep_data.pickle")
    read_ndh_3 = pd.read_pickle(r"D:\DataAnalytics\news_analysis\output_data\ndh"
                                r"\ndh_tai-chinh_data.pickle")
    read_ndh_4 = pd.read_pickle(r"D:\DataAnalytics\news_analysis\output_data\ndh"
                                r"\ndh_bat-dong-san_data.pickle")
    ndh_news = pd.concat([read_ndh_1, read_ndh_2, read_ndh_3, read_ndh_4], ignore_index=True)

    ndh_sheet = workbook.add_worksheet('ndh')
    header_4 = [
        'STT',
        'Link',
        'Loại tin tức',
        'Tiêu đề',
        'Nội dung'
    ]
    # Set Column Width and Row Height
    ndh_sheet.set_column('A:A', 4)
    ndh_sheet.set_column('B:B', 28)
    ndh_sheet.set_column('C:C', 25)
    ndh_sheet.set_column('D:D', 30)
    ndh_sheet.set_column('E:E', 55)

    ndh_sheet.write_row('A1', header_4, headers_format)
    ndh_sheet.write_column(
        'A2',
        np.arange(ndh_news.shape[0]) + 1,
        text_right_format
    )
    ndh_sheet.write_column(
        'B2',
        ndh_news['Link'],
        text_left_format
    )
    ndh_sheet.write_column(
        'C2',
        ndh_news['Loại tin tức'],
        text_left_format
    )
    ndh_sheet.write_column(
        'D2',
        ndh_news['Tiêu đề'],
        text_left_format
    )
    ndh_sheet.write_column(
        'E2',
        ndh_news['Nội dung'],
        text_left_format
    )

    # ====================================================================================

    writer.close()

    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py', '')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start, 1)}s')
