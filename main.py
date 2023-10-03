import pandas as pd
import os
import requests
from bs4 import BeautifulSoup
import json
import ssl
import urllib.request
import urllib.parse


class EMT_clawer():
    def __init__(self):
        self.n = 0
        self.data = {}

    def load_data_from_excel(self, excel_path):
        excel_path = excel_path.replace("\\", "\\\\")
        if os.path.isfile(excel_path):
            data = pd.read_excel(excel_path)
            results = []
            for _, row in data.iterrows():
                name, id, birthday = row['姓名'], row['身分證字號'], row['出生年月日']
                birthday_date = birthday.date()
                result = self.get_emtinfo(id, birthday_date, name)

                results.append(result)
            print('共有' + str(self.n) + '筆資料查無資訊')
            return results
        else:
            print("Excel 文件不存在:", excel_path)

    def get_emtinfo(self, IDNB, birthday, FName):
        global new_df
        import requests
        url = 'https://ems.mohw.gov.tw/EMTHome/TrainingHistory?_=1695758666461'
        response = requests.get(url)
        cookies = response.cookies
        # 檢查是否成功獲取網頁
        if response.status_code == 200:
            # 使用Beautiful Soup解析HTML
            soup = BeautifulSoup(response.text, 'html.parser')

            # 找到具有指定name屬性的input元素
            input_element = soup.find('input', {'name': '__RequestVerificationToken'})

            # 檢查元素是否存在
            if input_element:
                # 獲取value屬性的值
                value = input_element['value']
            else:
                print("找不到指定的元素")
        else:
            print("無法獲取網頁")

        # 獲取cookies
        cookie_all = ''
        if cookies:
            for cookie in cookies:
                cookie_linked = cookie.name + '=' + cookie.value + ';'
                cookie_all += cookie_linked

        url = 'https://ems.mohw.gov.tw/EMTHome/QueryEMTTrainInfo'

        headers = {
            'Cookie': cookie_all,
        }
        data = {
            'ID': IDNB,
            'BIRTHDAY': birthday,
            '__RequestVerificationToken': value,
        }

        data = urllib.parse.urlencode(data).encode('utf-8')
        requests = urllib.request.Request(url=url, data=data, headers=headers)
        ssl_context = ssl._create_unverified_context()
        response = urllib.request.urlopen(requests, context=ssl_context)
        content = response.read().decode('utf-8')
        data = json.loads(content)
        self.data = data
        self.get_result(IDNB, birthday, FName)
        self.get_history()

    def get_result(self,IDNB, birthday, FName):
        data = self.data
        if not os.path.isfile('result.xlsx'):
            df = pd.DataFrame(columns=[
                "查詢結果",
                "姓名",
                "身分證字號",
                "生日",
                "證書有效期間(起)",
                "證書有效期間(結)",
                "複訓累計時數",
                "說明",
                "服務狀態",
                "服務單位",
                "資格",
                "系統有效期間(起)",
                "系統有效期間(結)",
            ])
            # 將資料保存到Excel檔
            df.to_excel('result.xlsx', index=False)

        existing_df = pd.read_excel('result.xlsx')

        # 創建一個新的 DataFrame，包含新資料
        try:
            new_data = [[
                data['Success'],
                data['EMT_NAME'],
                data['EMT_NNIID'],
                data['EMT_BIRTHDAY'],
                data['EMT_NNI_SDATE'],
                data['EMT_NNI_EDATE'],
                data['EMT_NNI_SUM_HOUR'],
                data['EMT_NNI_SUM_HOUR_TXT'],
                data['EMT_SERVICE_TYPE'],
                data['EMT_SERVICE_DEPT'],
                data['EMT_TYPE'],
                data['NEW_NNI_SDATE'],
                data['NEW_NNI_EDATE']
            ]]
            new_df = pd.DataFrame(new_data, columns=existing_df.columns)
            # 將新資料插入到現有資料框的第二行
            updated_df = pd.concat([existing_df.iloc[:1], new_df, existing_df.iloc[1:]], ignore_index=True)
            # 寫回 Excel 文件
            updated_df.to_excel('result.xlsx', index=False)


        except KeyError:
            self.n += 1
            print("("+FName+")"+"("+IDNB+")"+"("+str(birthday)+")" + '查無資料')

    def get_history(self):
        data = self.data
        history = []

        if 'EMT_TRAIN_HISTORY' in data:
            history = data.pop('EMT_TRAIN_HISTORY')

        if not os.path.isfile('history.xlsx'):
            df = pd.DataFrame(columns=[
                "身分證字號",
                "受訓城市",
                "受訓單位名稱",
                "受訓日期",
                "受訓時數",
                "初訓/複訓 類型",
                "結果",
                "單位電話"
            ])
            # 將資料保存到Excel檔
            df.to_excel('history.xlsx', index=False)

        existing_df = pd.read_excel('history.xlsx')
        new_data = []
        for emt_history in history:
            try:
                new_data.append([
                    emt_history['R_NNI_ID'],
                    emt_history['R_CH_CITY'],
                    emt_history['R_NCH_DEPTNAME'],
                    emt_history['R_CH_DATE'],
                    emt_history['R_CB_SUM'],
                    emt_history['R_CH_CLSLEVEL'],
                    emt_history['R_NCP_STSEXTIME'],
                    emt_history['R_TEL_TXT']
                ])
                new_df = pd.DataFrame(new_data, columns=existing_df.columns)
                # 將新資料插入到現有資料框的第二行
                updated_df = pd.concat([existing_df.iloc[:1], new_df, existing_df.iloc[1:]], ignore_index=True)
                # 寫回 Excel 文件
                updated_df.to_excel('history.xlsx', index=False)

            except:
                pass

    def comby(self):
        # 讀取第一個Excel文件
        file1 = 'history.xlsx'
        df1 = pd.read_excel(file1)

        # 讀取第二個Excel檔
        file2 = 'result.xlsx'
        df2 = pd.read_excel(file2)

        # 創建一個Excel writer物件，指定輸出檔案名
        output_file = 'emt.xlsx'
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            # 將第一個DataFrame寫入新的Excel檔的第一頁
            df1.to_excel(writer, sheet_name='受訓紀錄', index=False)

            # 將第二個DataFrame寫入新的Excel檔的第二頁
            df2.to_excel(writer, sheet_name='證照效期', index=False)

        print('合併完成，輸出檔案名為:', output_file)

    def input_excel_exp(self):
        if not os.path.isfile('年籍冊.xlsx'):
            df = pd.DataFrame(columns=[
                "姓名",
                "身分證字號",
                "出生年月日",
            ])
            # 將資料保存到Excel檔
            df.to_excel('年籍冊.xlsx', index=False)
        else:
            print('已有【年籍冊.xlsx】檔案請確認')


if __name__ == '__main__':
    excel_path = r'年籍冊.xlsx'
    emt_crawler = EMT_clawer()
    while True:
        print("\n選擇操作:")
        print("0. 產生輸入檔案")
        print("1. 從Excel文件中加載數據並獲取EMT信息")
        print("2. 獲取EMT歷史記錄")
        print("3. 將檔案合併為'emt.xlsx'")
        print("4. 退出")

        choice = input("請輸入選擇 (0/1/2/3/4): ")

        if choice == '0':
            emt_crawler.input_excel_exp()
        elif choice == '1':
            emt_data = emt_crawler.load_data_from_excel(excel_path)
        elif choice == '2':
            emt_crawler.get_history()
            print("EMT歷史記錄已獲取並保存到Excel文件【history.xlsx】&【result.xlsx】")
        elif choice == '3':
            emt_crawler.comby()
            print("已將【history.xlsx】&【result.xlsx】檔案合併至【emt.xlsx】")
        elif choice == '4':
            print("退出程序")
            break
        else:
            print("無效的選擇，請重新輸入。")
