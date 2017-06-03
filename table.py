import sys
import datetime
import os
from docx import Document


def main():

    start, end, days, doc_name = schedule_info()

    # Make date list

    dates = date_list(start, end, days)

    print('dates are ')
    print(dates)

    make_word(dates, doc_name)

    os.system( 'open ' + doc_name)



def date_list(start, end, days):

    # Make list of Strings of American format dates.
    # Assumes that days is a list of the numerical weekdays in sorted order.
    # e.g for MW days = [0, 2]

    # date.weekday()  Monday = 0, Tues = 1, Weds = 3 ....

    # TODO (does it?) assumes start and end are on the day of week as days in days string.  Should fix.

    include_weekday = len(days) > 1

    delta = datetime.timedelta(days=1)

    counter_date = start

    date_list = []

    while counter_date <= end:

        print('loop ', counter_date)

        # is this date one for the schedule?
        if counter_date.weekday() in days:
            date_list.append(date_format(counter_date, include_weekday))

        # Add 1 to counter_date
        counter_date += delta


    return date_list


def make_word(dates, filename):

    document = Document()

    table = document.add_table(rows=1, cols = 4)

    # TODO can create border for table?

    # Create header
    # This is really cool
    # https://python-docx.readthedocs.io/en/latest/

    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Session'
    hdr_cells[1].text = 'Date'
    hdr_cells[2].text = 'Topics'
    hdr_cells[3].text = 'Assignments'

    counter = 1

    for d in dates:

        row_cells = table.add_row().cells
        row_cells[0].text = str(counter)
        row_cells[1].text = d

        counter += 1

    document.save(filename)




def date_format(d, include_weekday):
    # Return string for schedule. e.g. "Mon May 4 2017"
    # or "May 4 2017"

    if (include_weekday):
        return d.strftime('%a, %b %d %Y')
    else :
        return d.strftime('%b %d %Y')


def schedule_info():

    args = sys.argv
    if len(args) is 5:

        start = args[1]
        end = args[2]
        weekly = args[3]
        doc_name = args[4]

    else :

        print('Need more data.')
        start = input('Enter start date as DD-MM-YYYY: ')
        end = input('Enter end date as DD-MM-YYYY: ')
        weekly = input('Enter days of week to generate e.g. M or TW or MTH:')
        doc_name = input('Enter document name')


    # Validate. TODO replace with loop.

    start_date = ensure_date(start)
    end_date = ensure_date(end)
    days = ensure_days(weekly)
    docname = ensure_valid_filename(doc_name)

    print('Your args are ', start_date, end_date, days, docname)

    return start_date, end_date, days, doc_name


def ensure_valid_filename(name):

    #does it ends with docx (or append docx)

    if name.endswith('doc'):
        return name + 'x'

    if name.endswith('docx'):
        return name

    return name + 'docx'


    # TODO check if valid filename, does file exist?



def ensure_days(weekly):

    weekly = weekly.upper()

    map

    valid = 'MTWTF'

    for d in weekly:
        if d not in valid:
            return None

    weekdays = weekly.replace('M', '0')
    weekdays = weekdays.replace('T', '1')
    weekdays = weekdays.replace('W', '2')
    weekdays = weekdays.replace('T', '3')
    weekdays = weekdays.replace('F', '4')


    weekdays_int = [ int(day) for day in weekdays ]
    weekdays_int.sort()   # That makes sorting easier :)


    print(weekdays_int)

    return weekdays_int


def ensure_date(date_string):

    try:
        dmy = date_string.split('-')
        assert len(dmy) is 3
        dmy = [ int(item) for item in dmy ]

        if dmy[2] < 99:
            dmy[2] += 2000  # This is how millenium bugs happen

        date = datetime.date(dmy[2], dmy[1], dmy[0])
        return date
    except:
        return None


if __name__ == '__main__':
    main()
