import os


class read_list:

    def __init__(self, date, id_NB, path):
        self.date = date
        self.id_NB = id_NB
        self.path = path

    def path_ads(self):
        path = input("請輸入清單檔案路徑:")
        r_path = r(path)
        while True:
            if not os.path.exists(r_path):
                print("請重新輸入路徑:")
            else:
                False
        return path


if __name__ == '__main__':
    read_list.path_ads(0)