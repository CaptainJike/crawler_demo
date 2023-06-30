import requests
import pandas as pd
import re
import json
import xlsxwriter
from bs4 import BeautifulSoup
import traceback


def send_post_request(webUrl):
    response = requests.get(webUrl)
    return response


try:
    page = 0
    # 每页调用数量
    count = 50
    while page != -1:
        url = 'https://www.isc.co.uk/Umbraco/Api/FindSchoolApi/FindSchoolListResults?skip={}&take=50'
        headers = {
            "Content-Type": "application/json"
        }
        data = {"locationLatitude": "", "locationLongitude": "", "distanceInMiles": 0, "residencyTypes": [],
                "genderGroup": "", "ageRange": 3, "religiousAffiliation": "", "financialAssistances": [],
                "examinations": [], "specialNeeds": "false", "scholarshipsAndBurseries": "false",
                "latitudeSW": 49.39021469482424, "longitudeSW": -18.0605503125, "latitudeNE": 60.57129810250982,
                "longitudeNE": 12.964840312500021, "contactCountyID": 0, "contactCountryID": 0, "londonBoroughID": 0,
                "filterByBounds": "true", "savedBounds": "true", "zoom": 5,
                "center": {"lat": 55.373470708722714, "lng": -2.547854999999999}}
        response = requests.post(url.format(page, count), data=json.dumps(data), headers=headers)
        schools = response.json()
        print('页码：{}，返回结果：{}'.format(page, schools))
        if not schools:
            page = -1
        else:
            excel_data = []
            for school in schools:
                address_parts = school.get('FullAddress', '').replace('\r\n', '').split(', ')
                address_parts1 = address_parts[1] if len(address_parts) > 1 else ""
                address_parts2 = address_parts[2] if len(address_parts) > 2 else ""
                address_parts3 = address_parts[3] if len(address_parts) > 3 else ""
                webUrl = 'https://www.isc.co.uk' + school['Url']
                soup = BeautifulSoup(send_post_request(webUrl).content, 'html.parser')
                script_tag = soup.find('script', type='application/ld+json')
                json_data = script_tag.string.strip()
                data1 = json.loads(json_data)
                filtered_entry = {
                    'Source': 'ISC Portal',
                    'Id': school.get('Id', ''),
                    'Name': school.get('Name', ''),
                    'Address1': school.get('Address1', ''),
                    'Address2': school.get('Address2', ''),
                    'Address3': school.get('Address3', ''),
                    'City': address_parts1,
                    'State': '',
                    'Mailing Country': address_parts2,
                    'Country': address_parts2,
                    'Postcode': address_parts3,
                    'Sub-Region': '',
                    'Region': '',
                    'Website': data1.get('url', ''),
                    'General contact email address': data1.get('email', ''),
                    'Lowest Age': '',
                    'Highest Age': '',
                    'Total Enrolment': school.get('PupilCount', ''),
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
                excel_data.append(filtered_entry)
            page += 50
        # 创建DataFrame并写入Excel文件
        df = pd.DataFrame(excel_data)
        with pd.ExcelWriter('学校1.xlsx', engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)
            worksheet = writer.sheets['Sheet1']
            worksheet.freeze_panes(1, 0)
except Exception as e:
    traceback.print_exc()
    print(f"发生了其他异常：{str(e)}")