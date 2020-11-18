import os
import sys
import csv
import tabula
import requests
from time import sleep
from bs4 import BeautifulSoup as Be
from anticaptchaofficial.recaptchav2proxyless import *


def solve_capture():
    solver = recaptchaV2Proxyless()
    solver.set_verbose(1)
    solver.set_key("67cf868b894c969c392310b070613ed9")
    solver.set_website_url("https://www.qkr.gov.al")
    solver.set_website_key("6LfcaioUAAAAAKe5mOhor1VqquuNuw_NyIbzdNUW")

    g_response = solver.solve_and_return_solution()
    if g_response != 0:
        return g_response
    else:
        print("task finished with error " + solver.error_code)


def search_key(Id):
    url = "http://www.qkr.gov.al/kerko/kerko-ne-regjistrin-tregtar/kerko-per-subjekt/"

    payload = {
        '__RequestVerificationToken': 'wgv0wCUbikBDY2bqIxOMALOM3qQLwxrH8dyGo3y84OZEOvuFu65gYRH3Dz8vroO1XTQ7IHqKsMTLx99_sd8rdSapCq1ALivWHfMhGQhjouM1',
        'Nipt': Id,
        'ufprt': '8354D8AD9CD23844D1F1C4FF15F16EB62C41B8DF181434A9E438CAF92CFB31C69E05A056C00BD3A3A3BE83182694363679724396E83D8ABAB63A08F9F687F228AFF5B57ACF8B915E1952A5AC46715EAB88F4E475B5D8AB6A059A8D7E0E5F7DE9FBC169ABE7FB11F028BF142C5C0AA2E43D2A9B05FA2EBCA53B8DF61182E07379BFC7BAC568930613D9568E4E07AB70D69244B5B766530DD3C32E3FB3189AEE3C'}

    headers = {
        'Cookie': 'ASP.NET_SessionId=4rz0fete2tyybhzvsv51d1hu; __RequestVerificationToken=Fb567JV4xM5XsxRHLZVHakArmnjwuLdlmKazC_-niEAlZ2ZDiLT2Rd8oD6iqkW4bljglZJYZXTBUr4T-s-KR5wbRTH_EUqkkyZ_Ka9oEuU81',
        'user-agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.121 Safari/537.36"
    }

    try:
        response = requests.request("POST", url, headers=headers, data=payload)
        if response.status_code == requests.codes.ok:
            return response.text
        else:
            return None
    except ConnectionError:
        sleep(3)
        return search_key(Id)
    except Exception as e:
        print(e)
        return None


def parse(content):
    return Be(content, 'html5lib')


def write(lines):
    with open(file='result.csv', encoding='utf-8', mode='a', newline='') as csv_file:
        writer = csv.writer(csv_file, delimiter=',')
        writer.writerows(lines)


def get_indexes(rows):
    indexes = []
    for index, row in enumerate(rows):
        value = -1
        finalIndex = -1
        finalIndexes = []
        row_head = row[0].replace('.', '').strip()
        try:
            value = int(row_head[:1])
            finalIndex = index
            try:
                value = int(row_head[:2])
                finalIndex = index
                mainNum = row_head.strip()[:2]
                i = 0
                while True:
                    i = i + 1
                    subNum = mainNum + str(i)
                    if subNum in row_head:
                        finalIndexes.append(index + i)
                        continue
                    break
                try:
                    value = int(row_head[:3])
                    finalIndex = index
                except:
                    pass
            except:
                pass
        except:
            pass
        if index > 0 and value == 1:
            break
        if finalIndex != -1 and finalIndex not in indexes:
            indexes.append(finalIndex)
        for i in finalIndexes:
            if i not in indexes:
                indexes.append(i)
    indexes.append(len(rows))
    return indexes


def get_head_num(head):
    value = -1
    row_head = head.replace('.', '').strip()
    try:
        value = int(row_head[:1])
        try:
            value = int(row_head[:2])
            mainNum = row_head.strip()[:2]
            i = 0
            while True:
                i = i + 1
                subNum = mainNum + str(i)
                if subNum in row_head:
                    finalIndexes.append(index + i)
                    continue
                break
            try:
                value = int(row_head[:3])
            except:
                pass
        except:
            pass
    except:
        pass
    return value


def check_first_column_is_full_number(column):
    try:
        int(column.replace('.', '').strip())
        return True
    except:
        return False


def rearrange(row):
    if check_first_column_is_full_number(row[0]):
        return row[1:]
    else:
        return row


def convert_to_csv(input_path):
    output = "output.csv"
    tabula.convert_into(input_path, output, output_format="csv", pages='all')
    with open(file=output, mode='r') as file:
        rows = list(csv.reader(file))
    # os.remove(output)
    # os.remove(input_path)
    indexes = get_indexes(rows)
    records = []
    head_content = ''
    row_content = ''
    sub_dup_num = 0
    save_head_num = 0
    save_head = ''
    for key, index in enumerate(indexes):
        start_index = index
        end_index = indexes[key+1]
        head_num = get_head_num(rows[start_index][0])
        for index_range in range(start_index, end_index):
            if index_range == start_index and head_num == -1:
                sub_dup_num += 1
                real_head_num = str(save_head_num) + '.' + str(sub_dup_num)
                forward_head_num = str(save_head_num) + '.' + str(sub_dup_num + 1)
                if len(save_head.split(real_head_num)) > 1:
                    head_content = save_head.split(real_head_num)[1].split(forward_head_num)[0]
                for r in rows[index_range]:
                    row_content += r
            else:
                sub_dup_num = 0
                row = rearrange(rows[index_range])
                head_content += row[0]
                for r in row[1:]:
                    row_content += r
        save_head_num = head_num
        save_head = head_content
        record = [head_content, row_content]
        head_content = ''
        row_content = ''
        records.append(record)
    print(records)
    header = []
    write_row = []
    for record in records:
        header.append(record[0])
        write_row.append(record[1])
    write(lines=[header])
    write(lines=[write_row])


def write_pdf(content):
    pdf_path = 'result.pdf'
    with open(file=pdf_path, mode='wb') as file:
        file.write(content)
    convert_to_csv(pdf_path)


def download(code):
    url = "http://www.qkr.gov.al/umbraco/Surface/Documents/GenerateSimpleExtract"

    recaptcha_response = solve_capture()

    payload = {'SubjectDefCode': code,
               'g-recaptcha-response': recaptcha_response
               }

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.121 Safari/537.36',
        'Referer': 'http://www.qkr.gov.al/kerko/kerko-ne-regjistrin-tregtar/kerko-per-subjekt/',
        'Cookie': 'ASP.NET_SessionId=pxn25n4s1g2kzzdk1quko4ay; __RequestVerificationToken=06W4uFrgBbf2QndfrFWyyz19aSkYaoS0NumYit5ksMxT3OX30M6rJsMjC7lnhIgc9aXS53zoJNvSyAHAXvk5lxBnnZhO6OD7X7B0L225oC41; _ga=GA1.3.1010961615.1605535770; _gid=GA1.3.1222653889.1605535770; _gat=1'
    }

    response = requests.request("POST", url, headers=headers, data=payload)

    write_pdf(response.content)


def read_txt():
    with open(file='20201111-DPT-TaxpayerList.csv', encoding='utf-8', mode='r') as file:
        rows = list(csv.reader(file))
    Ids = []
    for row in rows:
        Ids.append(row[0].split(';')[0].strip())
    return Ids


def main():
    Ids = sys.argv[1:]
    Ids = read_txt()
    for Id in Ids:
        print(Id)
        response = search_key(Id)
        if response is not None:
            soup = parse(response)
            elements = soup.select(".result-element")
            for element in elements:
                subjectDefCode = element.find(id="subjectDefCode")['value']
                download(code=subjectDefCode)
                exit()


if __name__ == '__main__':
    convert_to_csv("EkstraktIThjeshte-11_18_2020.pdf")
    # main()
