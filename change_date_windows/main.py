import datetime
import win32api

def change_system_datetime(new_datetime):
    try:
        new_date = new_datetime.strftime('%m-%d-%Y')
        win32api.SetSystemTime(new_datetime.year, new_datetime.month, new_datetime.weekday(),
                               new_datetime.day, new_datetime.hour, new_datetime.minute,
                               new_datetime.second, 0)
        print("Date and time modified successfully.")
    except Exception as e:
        print("An error occurred while modifying the date and time:", str(e))


new_datetime = datetime.datetime(2022, 6, 11, 12, 30, 0)
change_system_datetime(new_datetime)
