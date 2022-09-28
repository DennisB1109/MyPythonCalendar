from datetime import *
from datetime import timedelta
from distutils.filelist import glob_to_re
from uuid import uuid1                                          # To assign each event a unique ID
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import sys
from tkinter import *


print("LETS GET STARTED\n")

def get_curr_day():
    """_summary_

    Returns:
        _type_: _description_
    """
    convert_to_german = {"Monday" : "Montag", "Tuesday" : "Dienstag", "Wednesday" : "Mittwoch", "Thursday" : "Donnerstag", "Friday": "Freitag", "Saturday" : "Samstag", "Sunday" : "Sonntag"}
    weekday = datetime.datetime.now().strftime("%A")
    print("Heute ist der %ste, das ist ein" % (datetime.datetime.now().day), (convert_to_german.get(weekday)))
    return weekday

def get_curr_month():
    """_summary_

    Returns:
        _type_: _description_
    """
    month_name = {1 : "Januar", 2: "Februar", 3: "März", 4: "April", 5: "Mai", 6: "Juni", 7: "Juli", 8: "August", 9: "September", 10: "Oktober", 11: "November", 12: "Dezember"}
    number_of_month = datetime.datetime.now().month
    print("Wir haben", month_name.get(number_of_month), "(%s)" % (number_of_month))
    return number_of_month
    

def get_curr_year():
    """_summary_

    Returns:
        _type_: _description_
    """
    print("Wir haben das Jahr:", datetime.datetime.now().year)
    return datetime.datetime.now().year

def get_date():
    """_summary_

    Returns:
        _type_: _description_
    """
    print("Heutiges datum: %s-%s-%s" % (datetime.now().day, datetime.now().month, datetime.now().year))
    return f'{datetime.now().day}.{datetime.now().month}.{datetime.now().year}'

def add_event():
    """_summary_

    Returns:
        _type_: _description_
    """
    wb = load_workbook('my_Events.xlsx')
    ws = wb.active
    print("Create an Appointment/Event\n")
    uid = str(uuid1().int>>64)
    day = input("Day: ")
    month = input("Month: ")
    year = input("Year: ")
    if len(year) == 2:
        temp = str(datetime.now().year)
        century = temp[0] + temp[1] + year
        temp = ''.join(century)
        year = temp
    description = input("Description: ")
    reminder = int(input("How many days before, do you want to get notified?: "))
    ws.append([uid, day, month, year, description, reminder])
    wb.save('my_Events.xlsx')
    added_event = f"Your event for the {day}.{month}.{year} was successfully added to your Calendar"
    print(added_event)
    return added_event

def del_event():
    """_summary_

    Returns:
        _type_: _description_
    """
    wb = load_workbook('my_Events.xlsx')
    ws = wb.active
    print("Which event should be deleted?\n")
    uid = input("UniqueID: ")
    temp_list = []
    was_event_deleted = False
    for row in range(2,100):
        temp_list.clear()
        temp_list.append(ws[get_column_letter(1) + str(row)].value)
        if temp_list[0] == uid:
            deleted_day = ws[get_column_letter(2) + str(row)].value
            deleted_month =ws[get_column_letter(3) + str(row)].value
            deleted_year =ws[get_column_letter(4) + str(row)].value
            print(f"Your Event for the {deleted_day}.{deleted_month}.{deleted_year} with the ID {uid} was successfully deleted")
            ws.delete_rows(row)
            wb.save('my_Events.xlsx')
            was_event_deleted = True
            break
    if was_event_deleted is False:
        print(f"There is no event with the ID {uid}")
    return uid

def show_all_events():
    """_summary_

    Returns:
        _type_: _description_
    """
    wb = load_workbook('my_Events.xlsx')
    ws = wb.active    
    temp_list = []
    for row in range(3,100):
        check_empty = ws[get_column_letter(1) + str(row)].value
        if check_empty is None:
            break
        temp_list.clear()
        for col in range(1,6):
            temp_list.append(ws[get_column_letter(col) + str(row)].value)
        print(temp_list)
    return temp_list

def check_events():
    """_summary_

    Returns:
        _type_: _description_
    """
    wb = load_workbook('my_Events.xlsx')
    ws = wb.active
    upcoming_events = False
    for row in range(3, 100):
        save_reminder = ws[get_column_letter(6) + str(row)].value
        if save_reminder is None:
            break
        day = timedelta(days = save_reminder)
        get_event_day = ws[get_column_letter(2) + str(row)].value
        get_event_month = ws[get_column_letter(3) + str(row)].value
        get_event_year = ws[get_column_letter(4) + str(row)].value[2] + ws[get_column_letter(4) + str(row)].value[3]
        date_event_string = f'{get_event_day}/{get_event_month}/{get_event_year} 00:00:01'
        date_time_obj = datetime.strptime(date_event_string, '%d/%m/%y %H:%M:%S')
        
        remind_on_this_date = str(date_time_obj - day)

        remind_on_this_date = remind_on_this_date[:-9].strip()
        temp_datetime_now = str(datetime.now())[:-16].strip()
        
        if temp_datetime_now == remind_on_this_date:
            print(f"{ws[get_column_letter(5) + str(row)].value} in {save_reminder} days")
            upcoming_events = True
    return upcoming_events

def test_gui():
    """_summary_
    """
    root = Tk()

    display_date = get_date()
    my_label = Label(root, text=('Date', display_date))
    my_label.place(relx=0.0, rely=0.0, anchor='nw')

    invalid_label = Label(root)
    my_label_added_event = Label(root)

    def do_an_entry():
        valid_day = True
        day_value = entry_day.get()
        valid_month = True
        month_value = entry_month.get()
        valid_year = True
        year_value = entry_year.get()
        description_value = entry_description.get()
        
        for character in day_value:
            try:
                converted_character = int(character)
                #print(type(converted_character))
                if converted_character not in range(0, 10):
                    print("Invalid input")
                    valid_day = False
                if len(day_value) > 2:
                    print("Invalid input")
                    valid_day = False
            except ValueError:
                print("Invalid input")
                valid_day = False
                pass
        
        for character in month_value:
            try:
                converted_character = int(character)
                if converted_character not in range(0, 10):
                    print("Invalid input month_value")
                    valid_month = False
                if len(month_value) > 2:
                    print("Invalid input")
                    valid_month = False
            except ValueError:
                print("Invalid input")
                valid_month = False
                pass
        
        for character in year_value:
            try:
                converted_character = int(character)
                if converted_character not in range(0, 10):
                    print("Invalid input year_value")
                    valid_year = False
                if len(year_value) > 4 or len(year_value) < 2 or len(year_value) == 3:
                    print("Invalid input")
                    valid_year = False
            except ValueError:
                print("Invalid input")
                valid_year = False
                pass

        if valid_day is True and valid_month is True and valid_year is True:
            success_label = f'Your event "{description_value}" for the {day_value}.{month_value}.{year_value} was successfully added'
            nonlocal my_label_added_event
            my_label_added_event.destroy()
            nonlocal invalid_label
            invalid_label.destroy()
            my_label_added_event = Label(root, text=success_label)
            my_label_added_event.place(x=5, y=200)
        else:
            invalid_label.destroy()
            my_label_added_event.destroy()
            invalid_label = Label(root, text='Invalid Input')
            invalid_label.place(x=5, y=200)

    entry_day = Entry(root)
    entry_day.place(x=5, y=60)
    entry_month = Entry(root)
    entry_month.place(x=5, y=82)
    entry_year = Entry(root)
    entry_year.place(x=5, y=104)
    entry_description = Entry(root)
    entry_description.place(height=40 ,x=5, y=126)

    submit_event = Button(root, text="Add Event", padx=8, pady=3, command=do_an_entry)
    submit_event.place(x=5, y=168)

    root.title('Calendar')
    root.iconbitmap('Apple_Calendar_Icon.png')
    root.geometry("500x400")
    root.mainloop()

test_gui()

# def main():
#     function = sys.argv[1]
#     if function == "get_curr_day":
#         get_curr_day()
#     if function == "get_curr_month":
#         get_curr_month()
#     if function == "get_curr_year":
#         get_curr_year()
#     if function == "get_date":
#         get_date()
#     if function == "add_event":
#         add_event()
#     if function == "del_event":
#         del_event()
#     if function == "show_all_events":
#         show_all_events()        

# if (__name__ == "__main__"):
#      main()
