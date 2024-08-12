# Project name: Automatic_email_sending
## Version: 1.2
## Date: 12.08.2024
Исходный код: https://github.com/gleb-pan/GenReport

Программа не требует установки Excel на сервер. Данная конфигурация предназначена для выгрузки из БД
системы GE CIMPLICITY HMI\SCADA.

Этот проект нацелен на ежедневную отправку выгрузки архивов событий из базы данных MS SQL.
Главный принцип работы приложения: SCADA система ежедневно запускает приложение automatic_email_sending,
которое производит выгрузку данных за предыдущие сутки из базы данных, формирует .xlsx файл и отправляет
его указанным получателям. Все настройки приложения производятся через файл .\_internal\config.ini.
А таже все действия программы и обшибки программы записываются в файл .\_internal\app_log.ini.

Настройки системы скада произведены в EventManager. Ежедневно в 8:00 запускается скрипт
Script Engine\Scripts\send_emails.bcl, запускающий automatic_email_sending через командную строку.

Пример скрипта (VBscript):

    Sub Main()
    Set objShell = CreateObject("Shell.Application")
    objShell.ShellExecute "C:\Users\Public\Documents\Automatic_email_sending_v1.2\Automatic_email_sending_v1.2.exe", "", "C:\Users\Public\Documents\Automatic_email_sending_v1.2", "runas", 0
    End Sub

Первый запуск:
    Убедитесь что в папке \_internal\ имеется файл config.ini и задайте в нем все необходимые настройки.
    
    Если этот файл отствует, вам необходимо его создать и затем заполнить.
    Шаблон файла конфигурации (config.ini), копировать без знаков "=":

    ========================================================================
    [Settings]
    smtp_server = smtp-mail.outlook.com
    smtp_port = 587
    
    [Credentials]
    username = test@email.com
    password = password
    
    [Other]
    to = recipient1@email.com,recipient2@email.com,recipient3@email.com
    
    [db_conn]
    driver = ODBC Driver 17 for SQL Server
    server = PC-NAME\DATABASE_NAME
    database = DATABASE_NAME
    username = sa
    password = password
    QUERY = SELECT * FROM table;
    ========================================================================

Все созданные .xlsx файлы сохраняются в директорию .\internal\Data\.