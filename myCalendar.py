import datetime
from uuid import uuid1                                          # To assign each event a unique ID
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

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
    print("Heutiges datum: %s-%s-%s" % (datetime.datetime.now().day, datetime.datetime.now().month, datetime.datetime.now().year))
    return datetime.datetime.now().day, datetime.datetime.now().month, datetime.datetime.now().year

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
        temp = str(datetime.datetime.now().year)
        century = temp[0] + temp[1] + year
        temp = ''.join(century)
        year = temp
    description = input("Description: ")
    ws.append([uid, day, month, year, description])
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
    for row in range(2,100):
        temp_list.clear()
        temp_list.append(ws[get_column_letter(1) + str(row)].value)
        if temp_list[0] == uid:
            wb.save('my_Events.xlsx')
            deleted_day = ws[get_column_letter(2) + str(row)].value
            deleted_month =ws[get_column_letter(3) + str(row)].value
            deleted_year =ws[get_column_letter(4) + str(row)].value
            print(f"Your Event for the {deleted_day}.{deleted_month}.{deleted_year} with the ID {uid} was successfully deleted")
            ws.delete_rows(row)
            break
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
