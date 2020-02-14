from vk_api.bot_longpoll import VkBotLongPoll, VkBotEventType
import vk_api
import datetime
import os
import urllib
import wget
import xlrd

vk = vk_api.VkApi(token="904da0e00c6997209af932a9a93b92fbda3226dbe61deb9756393ee818fee6c30b92405740d0101314f6c")

vk._auth_token()

vk.get_api()

longpoll = VkBotLongPoll(vk, 191998639)

while True:
    for event in longpoll.listen():
        if event.type == VkBotEventType.MESSAGE_NEW:
            if event.object.peer_id != event.object.from_id & event.object.from_id != 496208206:
                print(str(event.object.text.lower()) + ": " + str(event.object.from_id))
                if event.object.text.lower() == "!1с1":

                    date = datetime.datetime.now()

                    day = date.day
                    dayPlus1 = date.day + 1
                    month = date.month
                    year = str(date.year)

                    dayNikita = ""

                    # Скачивание .xls файла для дальнейшего чтения
                    if month < 10:
                        month = "0" + str(month)

                    if dayPlus1 < 10:
                        day = "0" + str(dayPlus1)

                    link = "http://www.urtk-mephi.ru/Raspisan/" + str(day) + "." + str(month) + ".20" + ".xls"
                    linkPlus1 = "http://www.urtk-mephi.ru/Raspisan/" + str(dayPlus1) + "." + str(month) + ".20" + ".xls"
                    try:
                        file = wget.download(linkPlus1, "day_temp.xls")

                        availabilityFile = os.path.isfile("day.xls")
                        if availabilityFile == True:
                            os.remove("day.xls")

                        os.rename(file, "day.xls")
                        dayNikita = dayPlus1
                    except urllib.error.HTTPError:
                        # print("\nРасписание на следующий день ещё не выложили, показано расписание на сегодня:\n")
                        file = wget.download(link, "day_temp.xls")

                        availabilityFile = os.path.isfile("day.xls")
                        if availabilityFile == True:
                            os.remove("day.xls")

                        os.rename(file, "day.xls")
                        dayNikita = day
                    # -------------------------------

                    # Проверка на чётность недели
                    today = datetime.datetime.today()
                    w = int(today.strftime("%U"))

                    if (w % 2 == 0):
                        ParityOfTheWeek = True
                    else:
                        ParityOfTheWeek = False

                    # -------------------------------

                    # Чтение .xls файла
                    try:
                        xls = xlrd.open_workbook("day.xls")
                        sheet = xls.sheet_by_index(0)
                    except FileNotFoundError:
                        print("Расписание на следующий день ещё не выложили, показано стандартное расписание:")

                    group1 = sheet.cell(2, 0).value
                    group2 = sheet.cell(2, 3).value
                    group3 = sheet.cell(2, 6).value
                    group4 = sheet.cell(2, 9).value

                    if group1 == 'Группа 1С1':
                        couple1 = "1. " + sheet.cell(3, 1).value
                        couple2 = "2. " + sheet.cell(4, 1).value
                        couple3 = "3. " + sheet.cell(5, 1).value
                        couple4 = "4. " + sheet.cell(6, 1).value
                    elif group2 == 'Группа 1С1':
                        couple1 = "1. " + sheet.cell(3, 4).value
                        couple2 = "2. " + sheet.cell(4, 4).value
                        couple3 = "3. " + sheet.cell(5, 4).value
                        couple4 = "4. " + sheet.cell(6, 4).value
                    elif group3 == 'Группа 1С1':
                        couple1 = "1. " + sheet.cell(3, 7).value
                        couple2 = "2. " + sheet.cell(4, 7).value
                        couple3 = "3. " + sheet.cell(5, 7).value
                        couple4 = "4. " + sheet.cell(6, 7).value
                    elif group4 == 'Группа 1С1':
                        couple1 = "1. " + sheet.cell(3, 10).value
                        couple2 = "2. " + sheet.cell(4, 10).value
                        couple3 = "3. " + sheet.cell(5, 10).value
                        couple4 = "4. " + sheet.cell(6, 10).value
                    else:
                        import datetime

                        d = datetime.date(int(year), int(month), int(day))
                        dayOfWeek = d.isoweekday()

                        '''
                       1 - Понедельник
                       2 - Вторник
                       3 - Среда
                       4 - Четверг
                       5 - Пятница
                       6 - Суббота
                       7 - Воскресенье
                       '''

                        if ParityOfTheWeek == True:
                            if dayOfWeek == 7:
                                couple1 = "1. Физика"
                                couple2 = "2. Математика"
                                couple3 = "3. Астрономия"
                                couple4 = ""
                            elif dayOfWeek == 1:
                                couple1 = "1. Физкультура"
                                couple2 = "2. Физика"
                                couple3 = "3. Обшествознание"
                                couple4 = "4. Ин.яз"
                            elif dayOfWeek == 2:
                                couple1 = "1. Математика"
                                couple2 = "2. История"
                                couple3 = "3. Химия"
                                couple4 = ""
                            elif dayOfWeek == 3:
                                couple1 = "1. Математика"
                                couple2 = "2. Русский язык"
                                couple3 = "3. Физ-ра"
                                couple4 = ""
                            elif dayOfWeek == 4:
                                couple1 = "1. Литература"
                                couple2 = "2. Математика"
                                couple3 = "3. Информатика"
                                couple4 = ""
                            elif dayOfWeek == 5:
                                couple1 = "1. История "
                                couple2 = "2. Обж"
                                couple3 = ""
                                couple4 = ""
                            elif dayOfWeek == 6:
                                couple1 = "1. Физика"
                                couple2 = "2. Математика"
                                couple3 = "3. Астрономия"
                                couple4 = ""
                        elif ParityOfTheWeek == False:
                            if dayOfWeek == 7:
                                couple1 = "1. Физика"
                                couple2 = "2. Математика"
                                couple3 = "3. Астрономия"
                                couple4 = ""
                            elif dayOfWeek == 1:
                                couple1 = "1. Русский язык"
                                couple2 = "2. Математика"
                                couple3 = "3. Информатика"
                                couple4 = "4. Обществознание"
                            elif dayOfWeek == 2:
                                couple1 = "1. Физкультура"
                                couple2 = "2. ВВС"
                                couple3 = "3. Физика"
                                couple4 = ""
                            elif dayOfWeek == 3:
                                couple1 = "1. История"
                                couple2 = "2. Математика"
                                couple3 = "3. Английский язык"
                                couple4 = ""
                            elif dayOfWeek == 4:
                                couple1 = "1. Математика"
                                couple2 = "2. Астрономия"
                                couple3 = "3. Физкультура"
                                couple4 = ""
                            elif dayOfWeek == 5:
                                couple1 = "1. История"
                                couple2 = "2. Литература"
                                couple3 = ""
                                couple4 = ""
                            elif dayOfWeek == 6:
                                couple1 = "1. Обществознание"
                                couple2 = "2. История"
                                couple3 = "3. Английский язык"
                                couple4 = ""
                    vk.method("messages.send", {"peer_id": event.object.peer_id,
                                                "message": "Расписание 1С1 на " + str(dayNikita) + "." + str(
                                                    month) + "." + str(
                                                    year) + "\n" + "\n\n" + couple1 + "\n" + couple2 + "\n" + couple3 + "\n" + couple4,
                                                "random_id": 0})
            elif event.object.from_id == 496208206:
                if event.object.text.lower() == "!1":
                    vk.method("messages.send", {"peer_id": event.object.peer_id,
                                                "message": "Витя, иди в пизду, тебе даже собака подчинятся не будет",
                                                "random_id": 0})



