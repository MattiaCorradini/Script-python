import datetime
import subprocess

def change_system_datetime(new_datetime):
    try:
        date_command = ['date', '-s', new_datetime.strftime('%Y-%m-%d')]
        subprocess.run(date_command, check=True)
        print("Date modified successfully.")
    except Exception as e:
        print("An error occurred while modifying the date:", str(e))

    try:
        time_command = ['date', '-s', new_datetime.strftime('%H:%M:%S')]
        subprocess.run(time_command, check=True)
        print("Time modified successfully.")
    except Exception as e:
        print("An error occurred while modifying the time:", str(e))


new_datetime = datetime.datetime(2022, 6, 23, 14, 30, 0)
change_system_datetime(new_datetime)
