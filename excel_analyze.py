import os, pandas, win32com.client, openpyxl, xlrd, sqlite3

def is_encrypted_excel(full_filepath):
    if full_filepath.find('.xlsx') != -1:
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
    if is_encrypted_excel(full_filepath) == False:
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
        r_money = (df_temp[data_start_row:][money_col])
        r_name = (df_temp[data_start_row:][name_col])
        df_data = pandas.concat([r_name, r_money], axis=1)
        df_data.reset_index(inplace=True, drop=True)
        df_data.rename(columns={name_col:'이름', money_col:'금액'}, inplace=True)
        df_data = df_data.dropna(axis=0)
        return df_data
    else:
        xlApp = win32com.client.Dispatch("Excel.Application")
        xlApp.Visible = False
        excel_password = '961006'

        try:
            book = xlApp.Workbooks.Open(full_filepath, False, True, None, excel_password)
            ws = book.ActiveSheet
            temp_dict = {}
            temp_namelist = []
            temp_moneylist = []
            for i in range(12, 38):
                if ws.Cells(i, 7).Value in temp_dict.keys():
                    temp_namelist.append(ws.Cells(i, 7).Value + str(i))
                    temp_moneylist.append(ws.Cells(i, 4).Value)
                else:
                    temp_namelist.append(ws.Cells(i, 7).Value)
                    temp_moneylist.append(ws.Cells(i, 4).Value)
            temp_dict['이름'] = temp_namelist
            temp_dict['금액'] = temp_moneylist
            df_data = pandas.DataFrame(data = temp_dict, index=None)
            xlApp.Quit()
            return df_data
        except Exception as e:
            if str(e).find('암호가 잘못되었습니다.') != -1:
                print('Invalid password.')
            else:
                print('uf_excel_reader error : ', e)
                
def uf_excel_writer(df_data):
    if os.path.isfile('result_data.xlsx'):
        writer = pandas.ExcelWriter(path = 'result_data.xlsx', mode='a', engine='openpyxl', if_sheet_exists='overlay')
        max_row = writer.sheets['Sheet1'].max_row
        df_data.to_excel(writer, startcol=0, startrow=max_row, index=False, header=None)
        writer.close()
        db_connect(df_data)
    else:
        df_data.to_excel(excel_writer='result_data.xlsx', index=False)
        db_connect(df_data)

def uf_data_excel_writer(name, money):
    temp_data = {'이름' : name, '금액' : money}
    df_data = pandas.DataFrame(temp_data)
    uf_excel_writer(df_data)

def excel_analysis(dirpath):
    data_list = os.listdir(dirpath)
    bank_name = '농협'

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

def file_path():
    script_file_path = os.path.abspath(__file__)
    script_dir_path = os.path.dirname(script_file_path)
    script_dir_path = str(script_dir_path).replace('\\', '/')
    return script_dir_path

def user_input_validate(str_name, str_money):
    if str_name == '' or str_money == '':
        return False, '이름 또는 금액이 빈칸입니다.'
    
    if str_money.isdecimal() == False:
        return False, '금액 란에는 숫자만 입력하세요.'

    if str_name.isalpha() == False:
        return False, '이름 란에는 한글 또는 영문만 입력하세요.'
    
    return True, ''

def db_connect(df_data):
    conn = sqlite3.connect('funeral_db.db')
    cur = conn.cursor()
    cur.execute('CREATE TABLE IF NOT EXSIST money_table(id PRIMARY KEY autoincrement, s_name TEXT, s_moeny TEXT)')
    sql = 'INSERT INTO money_table(s_name, s_money) VALUES (?, ?)'
    for idx, values in df_data.iterrows():
        df_name = values['이름']
        df_money = values['금액']
        cur.execute(sql, (df_name, df_money))
