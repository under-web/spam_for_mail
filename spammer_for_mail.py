import smtplib  # Импортируем библиотеку по работе с SMTP
import os  # Функции для работы с операционной системой, не зависящие от используемой операционной системы
import time

# Добавляем необходимые подклассы - MIME-типы
import mimetypes  # Импорт класса для обработки неизвестных MIME-типов, базирующихся на расширении файла
from email import encoders  # Импортируем энкодер
from email.mime.base import MIMEBase  # Общий тип
from email.mime.text import MIMEText  # Текст/HTML
from email.mime.image import MIMEImage  # Изображения
from email.mime.audio import MIMEAudio  # Аудио
from email.mime.multipart import MIMEMultipart  # Многокомпонентный объект


def send_email(addr_to, msg_subj, msg_text, files):
    addr_from = "magic-magnet.ru@mail.ru"  # Отправитель
    password = "emufef56"  # Пароль

    msg = MIMEMultipart()  # Создаем сообщение
    msg['From'] = addr_from  # Адресат
    msg['To'] = addr_to  # Получатель
    msg['Subject'] = msg_subj  # Тема сообщения

    body = msg_text  # Текст сообщения
    msg.attach(MIMEText(body, 'plain'))  # Добавляем в сообщение текст

    process_attachement(msg, files)

    # ======== Этот блок настраивается для каждого почтового провайдера отдельно
    # ===============================================
    server = smtplib.SMTP('smtp.mail.ru', port=25)  # Создаем объект SMTP
    server.starttls()  # Начинаем шифрованный обмен по TLS
    # server.set_debuglevel(True)  # Включаем режим отладки, если не нужен - можно закомментировать
    try:
        print("Соединяюсь")
        server.login(addr_from, password)  # Получаем доступ
        print("Успешно".center(40, '#'))
    except Exception as e:
        print('Не получилось залогиниться')
        print(e)
    try:
        server.send_message(msg)  # Отправляем сообщение
        print('Отправлено\n\n')
    except Exception as e:
        print('Не отправили')
        print(e)
    server.quit()  # Выходим
    # ==========================================================================================================================


def process_attachement(msg, files):  # Функция по обработке списка, добавляемых к сообщению файлов
    for f in files:
        if os.path.isfile(f):  # Если файл существует
            attach_file(msg, f)  # Добавляем файл к сообщению
        elif os.path.exists(f):  # Если путь не файл и существует, значит - папка
            dir = os.listdir(f)  # Получаем список файлов в папке
            # print(dir)
            for file in dir:  # Перебираем все файлы и...
                attach_file(msg, f + "/" + file)  # ...добавляем каждый файл к сообщению


def attach_file(msg, filepath):  # Функция по добавлению конкретного файла к сообщению
    filename = os.path.basename(filepath)  # Получаем только имя файла
    ctype, encoding = mimetypes.guess_type(filepath)  # Определяем тип файла на основе его расширения
    if ctype is None or encoding is not None:  # Если тип файла не определяется
        ctype = 'application/octet-stream'  # Будем использовать общий тип
    maintype, subtype = ctype.split('/', 1)  # Получаем тип и подтип
    if maintype == 'text':  # Если текстовый файл
        with open(filepath) as fp:  # Открываем файл для чтения
            file = MIMEText(fp.read(), _subtype=subtype)  # Используем тип MIMEText
            fp.close()  # После использования файл обязательно нужно закрыть
    elif maintype == 'image':  # Если изображение
        with open(filepath, 'rb') as fp:
            file = MIMEImage(fp.read(), _subtype=subtype)
            fp.close()
    elif maintype == 'audio':  # Если аудио
        with open(filepath, 'rb') as fp:
            file = MIMEAudio(fp.read(), _subtype=subtype)
            fp.close()
    else:  # Неизвестный тип файла
        with open(filepath, 'rb') as fp:
            file = MIMEBase(maintype, subtype)  # Используем общий MIME-тип
            file.set_payload(fp.read())  # Добавляем содержимое общего типа (полезную нагрузку)
            fp.close()
            encoders.encode_base64(file)  # Содержимое должно кодироваться как Base64
    file.add_header('Content-Disposition', 'attachment', filename=filename)  # Добавляем заголовки
    msg.attach(file)  # Присоединяем файл к сообщению



def main():
    files = [r'C:\Users\AppDev\Desktop\Рассылка']  # папка с файлами

    telo_mail = ("\n"
              "Мы обращаемся к тому, кто знаком с туристическим бизнесом не понаслышке. Кто давно занимается торговлей сувенирной продукции. Нечем удивить клиента в прикассовой зоне? Продаёте один и тот же ассортимент из года в год? Ваши устаревшие магниты пылятся годами? Вы вынуждены распродавать их за копейки? \n"
              "Выход есть! Мы предлагаем уникальный продукт, аналогов которому в мире нет – ЖИВОЙ МАГНИТ!!! Повысьте свои продажи. Удивите своих клиентов. Станьте первым и уникальным среди конкурентов!\n"
              "\nНаши контакты: +7 952 407-10-11 WhatsApp "
                 "    По всем вопросам 233373g@gmail.com или 89524071011@mail.ru"
              "\n")
    with open('1.txt', 'r', encoding='utf-8', errors='ignore') as lf:
        mails = lf.readlines()
        lfg = 765
        for i in mails:
            print('\nОтправляю письмо №{} на {}'.format(str(lfg), i))
            try:
                send_email(i, "Уникальные подарочные магниты !", telo_mail, files)
            except Exception as e:
                print(e)
                continue

            lfg += 1
            time.sleep(1)


if __name__ == '__main__':
    main()