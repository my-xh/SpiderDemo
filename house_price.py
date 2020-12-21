import requests
import bs4
import openpyxl
import re
import os, sys

def open_url(url):
    headers = {
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36',
    }
    res = requests.get(url, headers=headers)
    return res

def get_data(res):
    soup = bs4.BeautifulSoup(res.text, 'html.parser')
    content = soup.find('div', id="Cnt-Main-Article-QQ")
    targets = content.find_all('p', style="TEXT-INDENT: 2em")
    targets = iter(targets)
    data = []
    for each in targets:
        if each.text.isnumeric():
            data.append([
                re.search('\[(.+)\]', next(targets).text).group(1),
                re.search('(\d.*)', next(targets).text).group(1),
                re.search('(\d.*)', next(targets).text).group(1),
                re.search('(\d.*)', next(targets).text).group(1),
            ])
    return data

def save_excel(data):
    wb = openpyxl.Workbook()
    wb.guess_types = True
    ws = wb.active
    ws.append(['城市', '平均房价', '平均工资', '房价工资比'])
    for each in data:
        ws.append(each)
    wb.save('2017年全国城市房价工资比排行榜.xlsx')

def main():
    url = 'https://news.house.qq.com/a/20170702/003985.htm'
    res = open_url(url)
    data = get_data(res)
    save_excel(data)

if __name__ == '__main__':
    main()
