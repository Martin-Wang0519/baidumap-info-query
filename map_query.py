import os

from time import time

from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from openpyxl import load_workbook, Workbook

from window_ui import Ui_MainWindow
from config import settings

import json
from datetime import datetime
from urllib import request
from urllib.parse import quote
import sys
import time


class Window(Ui_MainWindow, QMainWindow):
    def __init__(self, parent):
        super(Window, self).__init__(parent)
        self.setupUi(self)
        # self.setFixedSize(self.width(), self.height())

        self.url_amap = settings.get('url_amap')
        self.english_header = settings.get('english_header')
        self.query = None
        self.region = None

        self.page_size = 20
        self.records = []
        self.current_province = None
        self.current_city = None
        self.provinceComboBox.addItem('全国')
        self.provinceComboBox.addItems(settings.get('china_administrative_divisions').keys())

        self.excel_save_folder = None

    def selectSavePathButton_clicked(self):
        self.excel_save_folder = QFileDialog.getExistingDirectory(self, "选择输出文件夹", "./")

        if self.excel_save_folder is not None and self.excel_save_folder != "":
            self.textBrowser.append("选择文件夹成功" + "  " + self.excel_save_folder)
        else:
            self.textBrowser.append("选择文件夹失败！ 请重新选择文件夹")

    def queryButton_clicked(self):
        self.records.clear()
        if self.current_province == '全国':
            for province in settings.get('china_administrative_divisions').keys():
                if province in settings.get('municipality'):
                    self.getPOIdata(self.query, province)
                else:
                    for city in settings.get('china_administrative_divisions').get(province):
                        self.getPOIdata(self.query, city)

        elif self.current_city == '全省':
            for city in settings.get('china_administrative_divisions').get(self.current_province):
                self.getPOIdata(self.query, city)

        elif self.current_province in settings.get('municipality'):
            self.getPOIdata(self.query, self.current_province)
        else:
            self.getPOIdata(self.query, self.current_city)

    def keyWords_textChanged(self):
        self.query = self.keyWordsEdit.text()

    def provinceComboBox_currentIndexChanged(self):
        self.current_province = self.provinceComboBox.currentText()
        self.cityComboBox.clear()
        cities = settings.get('china_administrative_divisions').get(self.current_province)
        if cities is not None:
            self.cityComboBox.addItem('全省')
            self.cityComboBox.addItems(settings.get('china_administrative_divisions').get(self.current_province))

    def cityComboBox_currentIndexChanged(self):
        self.current_city = self.cityComboBox.currentText()

    def get_data(self, query, page_num, city_name):
        time.sleep(0.5)
        self.textBrowser.append('解析页码： ' + str(page_num) + ' ... ...')
        self.textBrowser.repaint()
        url = self.url_amap.format(query, page_num, city_name)
        # 中文编码
        url = quote(url, safe='/:?&=')
        with request.urlopen(url) as f:
            html = f.read()
        results = json.loads(html)['results']
        for result in results:
            record = []
            for key in self.english_header:
                if key in result.keys():
                    record.append(result[key])
                else:
                    record.append('')
            self.records.append(record)

    def get_total_record(self):
        url = self.url_amap.format(self.query, 1, self.region)
        url = quote(url, safe='/:?&=')
        with request.urlopen(url) as f:
            html = f.read()
        return json.loads(html)['total']

    def getPOIdata(self, query, city_name):
        self.textBrowser.append('正在查询{}的{}信息'.format(city_name, query))
        self.textBrowser.repaint()

        total_record = self.get_total_record()
        if total_record % self.page_size != 0:
            page_num = int(total_record / self.page_size) + 2
        else:
            page_num = int(total_record / self.page_size) + 1
        for each_page in range(1, page_num):
            self.get_data(query, each_page, city_name)

        self.textBrowser.append('查询完毕！')
        self.textBrowser.append('正在写入excel文件...')
        self.write_excel()

    def write_excel(self):
        work_book = Workbook()
        sheet = work_book.active
        sheet.append(settings.get('header'))

        for record in self.records:
            sheet.append(record)

        current_date = datetime.now().strftime('%Y-%m-%d')

        if self.current_province == '全国':
            excel_name = '全国{}查询{}.xlsx'.format(self.query, current_date)
        elif self.current_city == '全省' or self.current_province in settings.get('municipality'):
            excel_name = '{}{}查询{}.xlsx'.format(self.current_province, self.query, current_date)
        else:
            excel_name = '{}{}{}查询{}.xlsx'.format(self.current_province, self.current_city, self.query, current_date)

        excel_path = os.path.join(self.excel_save_folder, excel_name)

        work_book.save(excel_path)
        work_book.close()

        self.textBrowser.append('excel写入完毕')

    def closeEvent(self, event):
        event.accept()
        sys.exit(0)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    mw = Window(None)
    mw.show()
    sys.exit(app.exec_())
