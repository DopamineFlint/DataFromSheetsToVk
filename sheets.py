import vk_api
import time
import pandas
import dataframe_image as dfi
import re
import requests
import gspread
import base64

from google.auth.exceptions import TransportError
from gspread import CellNotFound
from oauth2client.service_account import ServiceAccountCredentials

rucaptcha_key = ""

counter = 0


def main():
    start()


def start():
    worksheet_list = open_sheet()

    while True:
        sheet_name = input("Введите точное название листа с которого вы хотите начать делать проход: ")

        correct_name = False
        index = 0
        i = 0

        for sheets in worksheet_list:
            if sheets.title == sheet_name:
                index = i
                correct_name = True
                break
            i += 1
        if correct_name:
            break
        else:
            print("Введёный вами лист не существует")

    while True:
        how_many_sheets = input("Введите сколько листов программе нужно пройти: ")
        if how_many_sheets.isalpha():
            print("Вы ввели символ, а не цифру")
        else:
            break
    google_sheets(index, worksheet_list, how_many_sheets)


def open_sheet():
    scope = [
        "https://spreadsheets.google.com/feeds",
        'https://www.googleapis.com/auth/spreadsheets.readonly',
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
    client = gspread.authorize(creds)
    while True:
        name = input("Введите точное название документа таблицы: ")
        try:
            sheet = client.open(name)
            return sheet.worksheets()
        except gspread.SpreadsheetNotFound:
            print("Неправильное название таблицы. Введите точное название таблицы")
        except TransportError:
            print("Проблема с соединением. Возможно у вас отключен интернет.")


def login_pass_check():
    while True:
        login = input("Введите логин: ")
        password = input("Введите пароль: ")
        try:
            vk_session = vk_api.VkApi(
                login,
                password,
                auth_handler=auth_handler,
                captcha_handler=captcha_handler
            )
            vk_session.auth()
            return vk_session
        except vk_api.BadPassword:
            print("Неправильный логин и(или) пароль. Введите правильные логин и пароль")
        except requests.ConnectionError:
            print("Проблема с соединением. Возможно у вас отключен интернет.")


def captcha_handler(captcha):
    print("Решается каптча...")
    key = uncaptcha(captcha.get_url())
    return captcha.try_again(key)


def auth_handler():
    key = input('Введите код подтверждения: ')
    remember_device = True

    return key, remember_device


def uncaptcha(url):
    imager = requests.get(url)
    r = requests.post(
        'http://rucaptcha.com/in.php',
        data={'method': 'base64', 'key': rucaptcha_key, 'body': base64.b64encode(imager.content)}
    )
    if r.text[:3] != 'OK|':
        print("11")
        print('captcha failed')
        return -1
    print(r.text)
    capid = r.text[3:]
    print(capid)
    time.sleep(5)
    capanswer = requests.post(
        'http://rucaptcha.com/res.php',
        data={'key': rucaptcha_key, 'id': capid, 'action': 'get'}
    ).text
    print(capanswer)
    if capanswer[:3] != 'OK|':
        print("22")
        print('captcha failed')
        return -1
    print("Капча решена!")
    return capanswer[3:]


def pause_time():
    global counter
    if counter == 9:
        print("Делаю паузу на 60 сек")
        time.sleep(61)
        counter = 0
        print("Пауза закончилась")
    else:
        counter += 1


def google_sheets(index, worksheet_list, how_many_sheets):
    hms = int(how_many_sheets) + index

    vk_session = login_pass_check()
    vk = vk_session.get_api()

    upload = vk_api.VkUpload(vk_session)

    for sheetItem in worksheet_list[index:hms]:
        time.sleep(2)
        pause_time()

        print("Проверяю " + sheetItem.title)
        link = sheetItem.cell(16, 2).value

        if link is not None:
            if "topic" not in link:
                print("Ссылка не найдена или написана не правильно. Пропуск.")

                pause_time()
            else:
                time.sleep(2)
                pause_time()
                current_link_cell = re.findall(r'\d+', link)
                print(current_link_cell)
                week = 1
                set_finish = False
                is_cell_none = False
                while True:
                    try:
                        if not is_cell_none:

                            if not set_finish:
                                current_week = sheetItem.find("Неделя " + str(week))
                            else:
                                current_week = sheetItem.find("Финиш")

                            if current_week:
                                current_week_row = current_week.row
                                current_week_col = current_week.col
                                if sheetItem.cell(current_week_row, current_week_col + 5).value == "POSTED":
                                    print("ОПУБЛИКОВАНО")
                                else:
                                    print("Не опубликовано")
                                    if not set_finish:
                                        cell_coords = sheetItem.find("Неделя " + str(week))
                                    else:
                                        cell_coords = sheetItem.find("Финиш")

                                    cell_row = cell_coords.row
                                    cell_col = cell_coords.col
                                    data = {'Калории': [], 'Белки': [], 'Жиры': [], 'Углеводы': []}

                                    for item in range(7):
                                        time.sleep(2)
                                        cell_row += 1
                                        if sheetItem.cell(cell_row, cell_col + 1).value is not None:
                                            data['Калории'].append(sheetItem.cell(cell_row, cell_col + 1).value)
                                        else:
                                            print("Одна из ячеек пуста. Пропускаю лист")
                                            is_cell_none = True
                                            break

                                    cell_row = cell_coords.row
                                    for item in range(7):
                                        time.sleep(2)
                                        cell_row += 1
                                        if sheetItem.cell(cell_row, cell_col + 2).value is not None:
                                            data['Белки'].append(sheetItem.cell(cell_row, cell_col + 2).value)
                                        else:
                                            print("Одна из ячеек пуста. Пропускаю лист")
                                            is_cell_none = True
                                            break

                                    if not is_cell_none:
                                        cell_row = cell_coords.row
                                        for item in range(7):
                                            time.sleep(2)
                                            cell_row += 1
                                            if sheetItem.cell(cell_row, cell_col + 3).value is not None:
                                                data['Жиры'].append(sheetItem.cell(cell_row, cell_col + 3).value)
                                            else:
                                                print("Одна из ячеек пуста. Пропускаю лист")
                                                is_cell_none = True
                                                break

                                    if not is_cell_none:
                                        cell_row = cell_coords.row
                                        for item in range(7):
                                            time.sleep(2)
                                            cell_row += 1
                                            if sheetItem.cell(cell_row, cell_col + 4).value is not None:
                                                data['Углеводы'].append(sheetItem.cell(cell_row, cell_col + 4).value)
                                            else:
                                                print("Одна из ячеек пуста. Пропускаю лист")
                                                is_cell_none = True
                                                break

                                    if not is_cell_none:
                                        df = pandas.DataFrame(
                                            data,
                                            columns=['Калории', 'Белки', 'Жиры', 'Углеводы'],
                                            index=['пон', 'вт', 'ср', 'четв', 'пятн', 'суб', 'воскр']
                                        )
                                        dfi.export(df, "mytable1.png")
                                        album = vk.photos.getAlbums(
                                            group_id=current_link_cell[0],
                                        )
                                        album_id = 0
                                        has_album = False
                                        for album_current_name in album['items']:
                                            if album_current_name['title'] == 'bot':
                                                album_id = album_current_name['id']
                                                has_album = True
                                                break

                                        if not has_album:
                                            album_json = vk.photos.createAlbum(
                                                title="bot",
                                                group_id=current_link_cell[0],
                                                upload_by_admins_only=1,
                                                comments_disabled=1
                                            )
                                            album_id = album_json['id']

                                        photo = upload.photo(
                                            'mytable1.png',
                                            album_id=album_id,
                                            group_id=current_link_cell[0]
                                        )
                                        current_photo_id = photo[0]['id']

                                        if not set_finish:
                                            print(vk.board.createComment(
                                                group_id=current_link_cell[0],
                                                topic_id=current_link_cell[1],
                                                message="Неделя " + str(week),
                                                from_group=1,
                                                attachments="photo-" + str(current_link_cell[0]) + "_" + str(current_photo_id)
                                                )
                                            )
                                        else:
                                            print(vk.board.createComment(
                                                group_id=current_link_cell[0],
                                                topic_id=current_link_cell[1],
                                                message="Финиш",
                                                from_group=1,
                                                attachments="photo-" + str(current_link_cell[0]) + "_" + str(
                                                    current_photo_id)
                                            )
                                            )

                                        sheetItem.update_cell(cell_coords.row, cell_col + 5, "POSTED")
                                        if not set_finish:
                                            print(
                                                "Неделя " +
                                                str(week) +
                                                " для " +
                                                str(link) +
                                                " загружена. Проверяю след недели или листы..."
                                            )
                                        else:
                                            print(
                                                "Финиш таблица " +
                                                "для " +
                                                str(link) +
                                                " загружена. Проверяю след недели или листы..."
                                            )
                                            time.sleep(2)
                                        pause_time()
                            if set_finish or is_cell_none:
                                break
                            week += 1
                    except CellNotFound:
                        if not set_finish:
                            set_finish = True
                        else:
                            break
        else:
            print("Ссылка пуста. Пропускаю лист.")
    print("Задача выполнена!")


if __name__ == "__main__":
    main()
