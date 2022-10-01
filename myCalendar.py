from datetime import *
from datetime import timedelta
from distutils.command import check
from distutils.filelist import glob_to_re
from msilib import type_binary
from turtle import bgcolor, color
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

def add_event(day: str, month: str, year: str, description: str, reminder: str):
    """_summary_

    Returns:
        _type_: _description_
    """
    wb = load_workbook('my_Events.xlsx')
    ws = wb.active
    print("Create an Appointment/Event\n")
    uid = str(uuid1().int>>64)
    #day = input("Day: ")
    #month = input("Month: ")
    #year = input("Year: ")
    if len(year) == 2:
        temp = str(datetime.now().year)
        century = temp[0] + temp[1] + year
        temp = ''.join(century)
        year = temp
    #description = input("Description: ")
    #reminder = int(input("How many days before, do you want to get notified?: "))
    ws.append([uid, day, month, year, description, reminder])
    wb.save('my_Events.xlsx')
    added_event = f"Your event for the {day}.{month}.{year} was successfully added to your Calendar"
    print(added_event)
    return added_event

def del_event(uid: str):
    """_summary_

    Returns:
        _type_: _description_
    """
    wb = load_workbook('my_Events.xlsx')
    ws = wb.active
    print("Which event should be deleted?\n")
    #uid = input("UniqueID: ")
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
        day = timedelta(days = int(save_reminder))
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

    appointment_label = Label(root, text=('Create an Appointment/Event'))
    appointment_label.place(x=38, y=35)

    delete_label = Label(root, text='Delete an Appointment/Event')
    delete_label.place(x=329, y=35)

    invalid_label = Label(root)
    my_label_added_event = Label(root)

    invalid_delete_label = Label(root)
    my_label_deleted_event = Label(root)

    def do_an_entry():
        valid_day = True
        day_value = entry_day.get()
        valid_month = True
        month_value = entry_month.get()
        valid_year = True
        year_value = entry_year.get()
        description_value = entry_description.get()
        valid_reminder = True
        reminder_value = entry_reminder.get()

        # After the Add Event button is pressed, the current entries get deleted
        entry_day.delete(0, "end")
        entry_month.delete(0, "end")
        entry_year.delete(0, "end")
        entry_description.delete(0, "end")
        entry_reminder.delete(0, "end")
        
        if day_value == "":
            print("Invalid input")
            valid_day = False
        for character in day_value:
            print("Reached loop")
            try:
                converted_character = int(character)
                if converted_character not in range(0, 10) or int(day_value) > 31:
                    print("Invalid input")
                    valid_day = False
                if len(day_value) > 2:
                    print("Length Error")
                    print("Invalid input")
                    valid_day = False
            except ValueError:
                print("Value Error")
                print("Invalid input")
                valid_day = False
                pass
        
        if month_value == "":
            print("Invalid input")
            valid_month = False
        for character in month_value:
            try:
                converted_character = int(character)
                if converted_character not in range(0, 10) or int(month_value) > 12:
                    print("Invalid input month_value")
                    valid_month = False
                if len(month_value) > 2 or len(month_value) == 0 or month_value is None:
                    print("Invalid input")
                    valid_month = False
            except ValueError:
                print("Invalid input")
                valid_month = False
                pass
        
        if year_value == "":
            print("Invalid input")
            valid_year = False
        for character in year_value:
            try:
                converted_character = int(character)
                if converted_character not in range(0, 10):
                    print("Invalid input year_value")
                    valid_year = False
                if len(year_value) > 4 or len(year_value) < 2 or len(year_value) == 3 or len(year_value) == 0 or not year_value:
                    print("Invalid input")
                    valid_year = False
            except ValueError:
                print("Invalid input")
                valid_year = False
                pass
        print(f"Type of empty is: {type(year_value)} and value is: {year_value}")
        
        if reminder_value == "":
            reminder_value = "0"
        for character in reminder_value:
            try:
                converted_character = int(character)
                if converted_character not in range(0, 10):
                    print("Invalid input year_value")
                    valid_reminder = False
            except ValueError:
                print("Invalid input")
                valid_reminder = False
                pass

        if valid_day is True and valid_month is True and valid_year is True and valid_reminder is True:
            success_label = f'Your event "{description_value}" for the {day_value}.{month_value}.{year_value} was successfully added\nYou will get notified {reminder_value} days before'
            nonlocal my_label_added_event
            my_label_added_event.destroy()
            nonlocal invalid_label
            invalid_label.destroy()
            my_label_added_event = Label(root, text=success_label)
            my_label_added_event.place(x=5, y=225)
            add_event(day_value, month_value, year_value, description_value, reminder_value)
        else:
            invalid_label.destroy()
            my_label_added_event.destroy()
            invalid_label = Label(root, text='Invalid Input')
            invalid_label.place(x=5, y=225)

    def delete_an_entry():
        valid_id = True
        id_value = entry_id.get()

        # After the Delete Event button is pressed, the current entries get deleted
        entry_id.delete(0, "end")

        for character in id_value:
            try:
                converted_character = int(character)
                #print(type(converted_character))
                if converted_character not in range(0, 10):
                    print("Invalid input")
                    valid_id = False
            except ValueError:
                print("Invalid input")
                valid_day = False
                pass

        if valid_id is True:
            wb = load_workbook('my_Events.xlsx')
            ws = wb.active
            temp_list = []
            for row in range(3,100):
                check_id = ws[get_column_letter(1) + str(row)].value
                if check_id is None:
                    print(f"Event with the id {id_value} could not be found")
                    break
                if check_id == id_value:
                    temp_list.append(ws[get_column_letter(2) + str(row)].value)
                    temp_list.append(ws[get_column_letter(3) + str(row)].value)
                    temp_list.append(ws[get_column_letter(4) + str(row)].value)
                    temp_list.append(ws[get_column_letter(5) + str(row)].value)
                    break
            if temp_list == []:
                nonlocal invalid_delete_label
                invalid_delete_label.destroy()
                nonlocal my_label_deleted_event
                my_label_deleted_event.destroy()
                invalid_delete_label = Label(root, text=f"Event with the id\n{id_value}\ncould not be found")
                invalid_delete_label.place(x=390, y=120)
            else:
                invalid_delete_label.destroy()
                my_label_deleted_event.destroy()
                my_label_deleted_event = Label(root, text=f"Your event {temp_list[3]}\nfor the {temp_list[0]}.{temp_list[1]}.{temp_list[2]}\nwas successfully deleted")
                my_label_deleted_event.place(x=350, y=120)
                del_event(id_value)
            

    def event_text(e):
        entry_day.delete(0, "end")
        entry_month.delete(0, "end")
        entry_year.delete(0, "end")
        entry_description.delete(0, "end")
        entry_reminder.delete(0, "end")
    
    def delete_text(e):
        entry_id.delete(0, "end")

    # Form to create an Event
    label_day = Label(root, text="Day")
    label_day.place(x=5, y=60)
    entry_day = Entry(root)
    entry_day.insert(0,"11")                                # Preview Text
    entry_day.place(x=75, y=60)
    label_month = Label(root, text="Month")
    label_month.place(x=5, y=82)
    entry_month = Entry(root)
    entry_month.insert(0, "09")                             # Preview Text
    entry_month.place(x=75, y=82)
    label_year = Label(root, text="Year")
    label_year.place(x=5, y=104)
    entry_year = Entry(root)
    entry_year.insert(0, "2064")                            # Preview Text
    entry_year.place(x=75, y=104)
    label_description = Label(root, text="Description")
    label_description.place(x=5, y=126)
    entry_description = Entry(root)
    entry_description.insert(0, "Dennis 63th Bday")         # Preview Text
    entry_description.place(height=40 ,x=75, y=126)
    label_reminder = Label(root, text="Reminder")
    label_reminder.place(x=5, y=168)
    entry_reminder = Entry(root)
    entry_reminder.insert(0, "5")                           # Preview Text
    entry_reminder.place(width=40, x=75, y= 168)

    submit_event = Button(root, text="Add Event", padx=8, pady=3, command=do_an_entry)
    submit_event.place(x=75, y=193)

    entry_day.bind("<FocusIn>", event_text)

    # Form to delete an Event
    label_id = Label(root, text="Event ID")
    label_id.place(x=320, y=60)
    entry_id = Entry(root)
    entry_id.insert(0,"3067211771744094029")                                # Preview Text
    entry_id.place(x=370, y=60)

    submit_delete = Button(root, text="Delete Event", padx=8, pady=3, command=delete_an_entry)
    submit_delete.place(x=400, y=85)

    entry_id.bind("<FocusIn>", delete_text)

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
