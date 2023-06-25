import win32com.client
import os
from datetime import datetime, timedelta
import schedule
import time

# Создаем экземпляр приложения Outlook
outlook = win32com.client.Dispatch('outlook.application')

# Получаем доступ к объекту MAPI
mapi = outlook.GetNamespace("MAPI")

# Выводим названия учетных записей Outlook
for account in mapi.Accounts:
    print(account.DeliveryStore.DisplayName)  # Название учетной записи Outlook

# Получаем доступ к папке "Входящие" (Inbox)
inbox = mapi.GetDefaultFolder(6)  # Папка "Входящие"

# Получаем доступ к папке внутри "Входящие"
inbox = inbox.Folders['VN']  # Папка внутри "Входящие" (замените "your folder" на имя нужной папки)

# Функция для обработки писем
def process_emails():
    # Получаем все сообщения в выбранной папке
    messages = inbox.Items 

    # Устанавливаем временной диапазон для ограничения поиска писем (последние 24 часа)
    received_dt = datetime.now() - timedelta(days=1)  # Сюда вводим значения дней по умолчанию 24 часа (1 день)
    received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')

    # Задаем отправителя и тему письма для фильтрации
    email_sender = 'Василий Наумов'  # Отправитель письма
    email_subject = 'FW: Отчет об активациях Nurtelecom'  # Тема письма

    # Применяем ограничения к набору сообщений
    messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")

    # Задаем путь для сохранения вложений в текущую директорию
    outputDir = r'D:\Aidar\Python\outlook_info'  # если хотим в текущую директорию, то оставляем outputDir = os.getcwd()

    try:
        for message in list(messages):
            # Проверяем соответствие отправителя, темы письма и даты получения
            if email_subject == message.subject and (message.SenderEmailAddress == email_sender or message.sender.Name == email_sender):
                try:
                    # Сохраняем все вложения письма
                    s = message.sender
                    for attachment in message.Attachments:
                        attachment.SaveASFile(os.path.join(outputDir, attachment.FileName))
                        print(f"Вложение {attachment.FileName} от {s} сохранено")
                except Exception as e:
                    print("Ошибка при сохранении вложения: " + str(e))
    except Exception as e:
        print("Ошибка при обработке писем: " + str(e))

# Запускаем обработку писем каждый день в определенное время
schedule.every().day.at("16:10").do(process_emails)  # Здесь установите желаемое время запуска

# Бесконечный цикл для выполнения заданий по расписанию
while True:
    schedule.run_pending()
    time.sleep(1)
