import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import time

output_path = '本科提前批B段-美术统考.xlsx'
workbook = Workbook()
worksheet = workbook.active
header = ['序号', '年度', '院校代号', '院校名称', '专业代号', '专业名称', '批次', '科类', '最低分', '平均分', '最低位次', '志愿号']
worksheet.append(header)
for page in range(1, 200):
    try:
        url = 'https://zy.hebeea.edu.cn:7001/hebgkzyfz/zyfz/web/lnwc?page=' + str(page) + '&csrfToken=63b8f15e-e5a1-4556-9fc4-bdf2d944312e1687250136336&queryTime=1687251004484&lsnf=&lspcdm=1&lskldm=4&lsyxmc=&lszymc=&lskswc=&lsjswc='
        print(page)
        response = requests.post(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        table = soup.find('table')
        rows = table.find_all('tr')
        data = []
        for row in rows:
            cells = row.find_all('td')
            row_data = []
            for cell in cells:
                row_data.append(cell.text.strip())
            if row_data:
                data.append(row_data)
        for row in data:
            worksheet.append(row)
    except Exception as e:
        print(f"当前页{page}发生了其他异常：{str(e)}")
        time.sleep(2)
workbook.save(output_path)
