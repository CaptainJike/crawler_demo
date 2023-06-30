import requests
import pandas as pd
import re
import json
import xlsxwriter

try:
    url = 'https://www.aisnsw.edu.au/finding-a-nsw-independent-school'
    response = requests.get(url)

    # 提取schools数组数据
    regex = r'var schools = (\[.*?\]);'
    match = re.search(regex, response.text)
    if match:
        schools = json.loads(match.group(1))

        # 构建数据列表
        data = []
        for school in schools:
            if school.get("secondary") and school["secondary"] == True:
                metroRegion = school.get('metroRegion', '')
                country = school.get('country', '')
                filtered_entry = {
                    'Source': 'aisnsw',
                    'Id': school['id'],
                    'Name': school['name'],
                    'Address1': school['location'],
                    'Address2': school['name'],
                    'Address3': school['name'],
                    'City': metroRegion,
                    'State': '',
                    'Mailing Country': country,
                    'Country': country,
                    'Postcode': '',
                    'Sub-Region': '',
                    'Region': '',
                    'Website': school['web'],
                    'General contact email address': school['email'],
                    'Lowest Age': '',
                    'Highest Age': '',
                    'Total Enrolment': '',
                    'Lowest Tuition Fee': '',
                    'Lowest Tuition Fee USD': '',
                    'Highest Tuition Fee': '',
                    'Highest Tuition Fee USD': '',
                    'Lowest Boarding Fee': '',
                    'Lowest Boarding Fee USD': '',
                    'Highest Boarding Fee': '',
                    'Highest Boarding Fee USD': '',
                    'Fee Currency': '',
                    'Mid-Market Fee': '',
                    'Premium Fee': '',
                    'Not for Profit': '',
                    'Orientations': '',
                    'Examinations': '',
                    'Groups': ''
                }
                data.append(filtered_entry)

        # 创建DataFrame并写入Excel文件
        df = pd.DataFrame(data)
        with pd.ExcelWriter('学校.xlsx', engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)
            worksheet = writer.sheets['Sheet1']
            worksheet.freeze_panes(1, 0)

    else:
        print("未找到schools数组数据")

except requests.exceptions.RequestException as e:
    print("请求发生异常：", str(e))

except Exception as e:
    print(f"发生了其他异常：{str(e)}")