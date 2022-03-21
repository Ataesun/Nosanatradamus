from flask import Flask, render_template, request, redirect
from splinter import Browser
import random, codecs, wget, os
from time import sleep
from bs4 import BeautifulSoup
import pandas as pd
from twocaptcha import TwoCaptcha

app = Flask(__name__)

# Данные для входа
login = 'ivanivanovic.ivanovich'
password = 'ivan123456654321'
request_word = ''

# Обход капчи
solver = TwoCaptcha('e0589f6f38e9d06b089aaa8867542129')

count_circle, global_array, itog_array, counter_itog, kp_array, kp_final = 0, [], [], 2, [0.001, 0.001, 0.005, 0.01, 0.02, 0.03, 0.04, 0.06, 0.07, 0.08, 0.09, 0.1], []
kp_counter = 2
df = pd.DataFrame()
global_array = {}

# Словарь с регионами
region_dict = {
    1: '1',  # Москва и область
    2: '10174',  # Санкт-Петербург и Ленинградская область
    3: '187',  # Украина
    4: '225%2C166%2C169'  # Россиия, СНГ и Грузия
}

region_name = {
    '1': 'Москва и область',
    '10174': 'Санкт-Петербург и Ленинградская область',
    '187': 'Украина',
    '225%2C166%2C169': 'Россиия, СНГ и Грузия'
}

param_dic = {
    1: "B",
    2: "C",
    3: "D",
    4: "E",
    5: "F",
    6: "G",
    7: "H",
    8: "I",
    9: "J",
    10: "K",
    11: "L",
    12: "M",
    13: "N",
    14: "O",
    15: "P",
    16: "Q",
    17: "R",
    18: "S",
    19: "T",
    20: "U",
    21: "V",
    22: "W",
    23: "X",
    24: "Y",
    25: "Z"
}

choose = region_dict[1]


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/post', methods=['POST'])
def post():
    global request_word, choose
    request_word = request.form['key_words'].split(',')
    choose = region_dict[int(request.form['drop'])]
    parse_wordstat(login, password, request_word, choose)
    return redirect('/static/results/result.xlsx')

@app.route('/mistake')
def mistake():
    return render_template('mistake.html')


def parse_wordstat(login, password, request_word, region):
    if ('./static/results/result.xlsx') == True:
        os.remove('./static/results/result.xlsx')
    try:
        for j in range(len(request_word)):
            browser = Browser()
            url = "https://wordstat.yandex.ru/#!/history?regions=" + str(region) + "&words=" + str(request_word[j])
            browser.visit(url)
            browser.click_link_by_href(
                'https://passport.yandex.ru/passport?mode=auth&msg=&retpath=https%3A%2F%2Fwordstat.yandex.ru%2F')
            browser.find_by_id('b-domik_popup-username').fill(login)
            sleep(random.randint(5, 10) / 10)
            browser.find_by_id('b-domik_popup-password').fill(password)
            sleep(random.randint(5, 10) / 10)
            button = browser.find_by_css('input[class="b-form-button__input"]')[2]
            button.click()
            sleep(random.randint(5, 10) / 10)
            pointer = 0
            while browser.find_by_css('img[class="b-popupa__image"]')['src'] != pointer:
                sleep(3)
                pointer = browser.find_by_css('img[class="b-popupa__image"]')['src']
                captcha = browser.find_by_css('img[class="b-popupa__image"]')
                name = os.environ.get("USERNAME")
                url_pack = 'C:/Users/' + str(name) + '/PycharmProjects/pythonProject/test.jpg'
                wget.download(captcha['src'], url_pack)
                result = solver.normal(url_pack)
                sleep(4)
                res = result['code']
                browser.find_by_css('input[class="b-form-input__input"]')[1].fill(res)
                os.remove(url_pack)
                sleep(2)
                btn = browser.find_by_css('input[class="b-form-button__input"]')[1]
                btn.click()
                sleep(6)
            odd, even, data_odd, data_even, final_array_date, final_array_count, final_array_relative = 0, 0, [], [], [], [], []
            screenshot_path = browser.html_snapshot()
            lst = list(screenshot_path)
            for i in range(37):
                lst.pop(0)
            file_name = ''.join(lst)
            path_to_file = 'C:/Users/WINDOW~1/AppData/Local/Temp/' + file_name
            f = codecs.open(path_to_file, 'r', 'utf8')
            html = BeautifulSoup(f.read(), 'lxml')
            table = html.find('div', class_="b-history__table-layout b-history__table-not-first").find('table',
                                                                                                       class_='b-history__table').find(
                'tbody', class_='b-history__table-body')
            odd = table.find_all('tr', class_='odd')
            for ch in odd:
                data_odd.append(ch.text)
            even = table.find_all('tr', class_='even')
            for ch in even:
                data_even.append(ch.text)
            for i in range(len(data_odd)):
                data_odd[i] = (data_odd[i][:23] + ',' + data_odd[i][23:])
                data_odd[i] = data_odd[i].split(',')
            for i in range(len(data_even)):
                data_even[i] = (data_even[i][:23] + ',' + data_even[i][23:])
                data_even[i] = data_even[i].split(',')
            for i in range(len(data_even)):
                final_array_date.append(data_odd[i][0])
                final_array_date.append(data_even[i][0])
                final_array_count.append(int(data_odd[i][1]))
                final_array_count.append(int(data_even[i][1]))
            global_array[request_word[j]] = final_array_count
            df['Месяц'] = final_array_date
            for i in range(len(global_array)):
                df[request_word[j]] = global_array[request_word[j]]

            browser.quit()
        global counter_itog, kp_counter
        array_counter = 0

        # Создание итого
        if len(request_word) > 1:
            for r in range(12):
                itog_array.append(r'=SUM(B{}:{}{})'.format(counter_itog, param_dic[len(request_word)], counter_itog))
                counter_itog += 1
        else:
            for r in range(12):
                itog_array.append(r'=SUM(B{})'.format(counter_itog))
                counter_itog += 1
        df['Итого'] = itog_array

        for i in range(12):
            kp_final.append((r'={}{}*{}').format(param_dic[len(request_word)+1], kp_counter, kp_array[array_counter]))
            kp_counter += 1
            array_counter +=1

        df['Данные по КП'] = kp_final
        writer = pd.ExcelWriter('./static/results/result.xlsx', engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Лист1', index=False)

        writer.sheets['Лист1'].set_column('A:A', 30)
        writer.sheets['Лист1'].set_column('B:B', 20)
        writer.sheets['Лист1'].set_column('C:C', 20)
        writer.sheets['Лист1'].set_column('D:D', 20)
        writer.sheets['Лист1'].set_column('E:E', 20)
        writer.sheets['Лист1'].set_column('F:F', 20)
        writer.sheets['Лист1'].set_column('G:G', 20)
        writer.sheets['Лист1'].set_column('H:H', 20)
        writer.sheets['Лист1'].set_column('I:I', 20)
        writer.sheets['Лист1'].set_column('J:J', 20)
        writer.sheets['Лист1'].set_column('K:K', 20)
        writer.sheets['Лист1'].set_column('L:L', 20)
        writer.sheets['Лист1'].set_column('M:M', 20)

        writer.save()
    except TwoCaptcha.api.ApiException :
        redirect('/mistake')


if __name__ == '__main__':
    app.run()

