import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

print("LETS GET STARTED\n")

def get_CurrDay():
    """_summary_

    Returns:
        _type_: _description_
    """
    convert_to_german = {"Monday" : "Montag", "Tuesday" : "Dienstag", "Wednesday" : "Mittwoch", "Thursday" : "Donnerstag", "Friday": "Freitag", "Saturday" : "Samstag", "Sunday" : "Sonntag"}
    weekday = datetime.datetime.now().strftime("%A")
    print("Heute ist der %ste, das ist ein" % (datetime.datetime.now().day), (convert_to_german.get(weekday)))
    return weekday

def get_CurrMonth():
    """_summary_

    Returns:
        _type_: _description_
    """
    month_name = {1 : "Januar", 2: "Februar", 3: "März", 4: "April", 5: "Mai", 6: "Juni", 7: "Juli", 8: "August", 9: "September", 10: "Oktober", 11: "November", 12: "Dezember"}
    number_of_month = datetime.datetime.now().month
    print("Wir haben", month_name.get(number_of_month), "(%s)" % (number_of_month))
    return number_of_month
    

def get_CurrYear():
    """_summary_

    Returns:
        _type_: _description_
    """
    print("Wir haben das Jahr:", datetime.datetime.now().year)
    return datetime.datetime.now().year

def get_Date():
    """_summary_

    Returns:
        _type_: _description_
    """
    print("Heutiges datum: %s-%s-%s" % (datetime.datetime.now().day, datetime.datetime.now().month, datetime.datetime.now().year))
    return datetime.datetime.now().day, datetime.datetime.now().month, datetime.datetime.now().year

def add_Event(): 
    wb = load_workbook('my_Events.xlsx')
    ws = wb.active
    print("Create an Appointment/Event\n")
    day = input("Day: ")
    month = input("Month: ")
    year = input("Year: ")
    if len(year) == 2:
        temp = str(datetime.datetime.now().year)
        century = temp[0] + temp[1] + year
        temp = ''.join(century)
        year = temp
    description = input("Description: ")
    ws.append([day, month, year, description])
    wb.save('my_Events.xlsx')
    added_event = f"Your event for the {day}.{month}.{year} was successfully added to your Calendar"
    return added_event

def del_event():
    wb = load_workbook('my_Events.xlsx')
    ws = wb.active
    print("Which event should be deleted?\n")
    day = input("Day: ")
    month = input("Month: ")
    year = input("Year: ")
    temp_list = []
    for row in range(2,10):
        print(temp_list)
        temp_list.clear()
        for col in range (1,5):
            print(ws[get_column_letter(col) + str(row)].value)
            temp_list.append(ws[get_column_letter(col) + str(row)].value)
        if temp_list[0] == day and temp_list[0] == month and temp_list[0] == year:
            ws.delete_rows(row)
            wb.save('my_Events.xlsx')

del_event()