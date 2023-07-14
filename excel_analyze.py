import os, pandas, win32com.client, openpyxl, xlrd

def checking_excel_encrypted(full_filepath):
    if full_filepath.find('.xlsx') != -1:
        # xlsx 확장자 파일에 대해 암호화 확인
        try:
            xl = openpyxl.load_workbook(full_filepath, read_only=True)
            xl.close()
            return False
        except Exception as e:
            return True
    else:
        try:
            wb = xlrd.open_workbook(full_filepath, on_demand=True)
            wb.release_resources()

            return False
        except Exception as e:
            return True

def uf_excel_reader(full_filepath, bank_name):
    '''
    매개변수로 받은 엑셀에 존재하는 이름, 금액을 추출하여
    결측치를 제거하고, 병합하여 하나의 데이터 프레임으로 리턴
    '''
    if checking_excel_encrypted(full_filepath) == False:
        df_temp = pandas.read_excel(full_filepath, header=None)
        data_start_row = 0
        money_col = 0
        name_col = 0
        if bank_name == '농협':
            data_start_row = 8
            money_col = 4
            name_col = 7
        elif bank_name == '신한':
            data_start_row = 7
            money_col = 4
            name_col = 5
        r_money = (df_temp[data_start_row:][money_col]) # 금액
        r_name = (df_temp[data_start_row:][name_col]) # 이름
        test = pandas.concat([r_name, r_money], axis=1)
        test.reset_index(inplace=True, drop=True)
        test.rename(columns={name_col:'이름', money_col:'금액'}, inplace=True)
        test = test.dropna(axis=0) # 결측치 제거
        return test
    else:
        xlApp = win32com.client.Dispatch("Excel.Application")
        xlApp.Visible = False
        excel_password = '961006'

        try:
            book = xlApp.Workbooks.Open(full_filepath, False, True, None, excel_password)
            ws = book.ActiveSheet
            for i in range(12, 38):
                print(ws.Cells(i, 4).Value, ws.Cells(i, 7).Value)
            xlApp.Quit()
            return 0
        except Exception as e:
            if str(e).find('암호가 잘못되었습니다.') != -1:
                print('비밀번호가 틀렸습니다.')
            else:
                print('uf_excel_reader 오류 : ', e)
                
def uf_excel_writer(df_data):
    if os.path.isfile('result_data.xlsx'):
        writer = pandas.ExcelWriter('result_data.xlsx', mode='a', engine='openpyxl', if_sheet_exists='overlay')
        max_row = writer.sheets['sheet1'].max_row
        
        df_data.to_excel(writer, sheet_name='sheet1', startcol=0, startrow=max_row, index=False, encoding='utf-8', header=None)
        writer.save()
        writer.close()  
    else:
        df_data.to_excel(excel_writer='result_data.xlsx', sheet_name='sheet1', index=False, encoding='utf-8')

    # if type(df_data) == int:
    #     print('엥')
    # try:
    #     writer = pandas.ExcelWriter('result_data.xlsx', mode = 'a', if_sheet_exists='overlay', engine='openpyxl')
    #     last_row = writer.sheets['sheet'].max_row
    #     if last_row == 1:
    #         type(df_data)
    #         df_data.to_excel(writer, sheet_name = 'sheet', index = False, encoding = 'utf-8', startrow = 0, startcol = 0)
    #     else:
    #         type(df_data)
    #         df_data.to_excel(writer, sheet_name = 'sheet', startrow = int(writer.sheets['sheet'].max_row), startcol = 0, index = False, header = None)
    # except Exception as e:
    #     if str(e).find('No such file') != -1:
    #         print(f'엑셀이 없어 생성합니다.')
    #         df_data.to_excel(excel_writer = 'result_data.xlsx', sheet_name = 'sheet', index = False)
    #     else:
    #         print(f'uf_excel_writer 오류 : {e}')
            
    writer.close()
    
    
def excel_analysis(dirpath):
    data_list = os.listdir(dirpath)
    bank_name = '농협' # 파일명으로 은행 가를수 있는가?

    for file_list in data_list:
        if file_list.find('농협') != -1:
            bank_name = '농협'
        elif file_list.find('신한') != -1:
            bank_name = '신한'
        elif file_list.find('카카오') != -1:
            bank_name = '카카오'
        else:
            print('사용 불가능한 엑셀입니다.')
            bank_name = ''
            continue
        dataframe_result = uf_excel_reader(dirpath + file_list, bank_name)
        uf_excel_writer(dataframe_result)

temp = 'F:/VSC_Project/funeral_project/data/'
excel_analysis(temp)
