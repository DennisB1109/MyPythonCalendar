import datetime

print("LETS GET STARTED\n")

def getCurrDay():
    convert_to_german = {"Monday" : "Montag", "Tuesday" : "Dienstag", "Wednesday" : "Mittwoch", "Thursday" : "Donnerstag", "Friday": "Freitag", "Saturday" : "Samstag", "Sunday" : "Sonntag"}
    weekday = datetime.datetime.now().strftime("%A")
    print("Heute ist der %ste, das ist ein" % (datetime.datetime.now().day), (convert_to_german.get(weekday)))

def getCurrMonth():
    month_name = {1 : "Januar", 2: "Februar", 3: "März", 4: "April", 5: "Mai", 6: "Juni", 7: "Juli", 8: "August", 9: "September", 10: "Oktober", 11: "November", 12: "Dezember"}
    number_of_month = datetime.datetime.now().month
    print("Wir haben", month_name.get(number_of_month), "(%s)" % (number_of_month))
    

def getCurrYear():
    print("Wir haben das Jahr:", datetime.datetime.now().year)

def getDate():
    print("Heutiges datum: %s-%s-%s" % (datetime.datetime.now().day, datetime.datetime.now().month, datetime.datetime.now().year))

getCurrDay()
getCurrMonth()
getCurrYear()
getDate()