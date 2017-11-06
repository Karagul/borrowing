import xlwings as xw
import datetime
import os
import sqlite3
import pandas as pd

def create_connection(db_file):
    try:
        conn = sqlite3.connect(db_file, detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES)
        return conn
    except Error as e:
        print(e)

    return None

def create_borrowing(conn, client):
    sql = ''' INSERT INTO borrowing(title, body, rate, start_date, end_date)
              VALUES(?,?,?,?,?)'''
    cur = conn.cursor()
    cur.execute(sql, client)
    return cur.lastrowid

def insert_a_borrowing():
    wb = xw.Book.caller()
    title = str(wb.sheets['management'].range("A4").value)
    body = wb.sheets['management'].range("B4").value
    rate = wb.sheets['management'].range("C4").value
    try:
        start_date = wb.sheets['management'].range("D4").value.strftime('%Y-%m-%d')
        end_date = wb.sheets['management'].range("E4").value.strftime('%Y-%m-%d')
    except AttributeError as e:
        wb.sheets['management'].range("A18").color = (240, 100, 77)
        wb.sheets['management'].range("A18").value = str(datetime.datetime.now()) + ': ' + str(e)
        return None
    database = os.path.join(os.path.dirname(wb.fullname), 'borrowing.db')
    # create a database connection
    conn = create_connection(database)
    try:
        with conn:
        # create a new project
            borrowing = (title, body, rate, start_date, end_date);
            borrowing_id = create_borrowing(conn, borrowing)
            wb.sheets['management'].range("A18").color = (146, 208, 80)
            wb.sheets['management'].range("A18").value = str(datetime.datetime.now()) + ": Создан договор займа № " + str(title)

    except sqlite3.IntegrityError as e:
        wb.sheets['management'].range("A18").color = (240, 100, 77)
        wb.sheets['management'].range("A18").value = str(datetime.datetime.now()) + ': ' + str(e)

def create_payment(conn, client):
    sql = ''' INSERT INTO payment(date, type, amount, borrowing)
              VALUES(?,?,?,?)'''
    cur = conn.cursor()
    cur.execute(sql, client)
    return cur.lastrowid

def insert_a_payment():
    wb = xw.Book.caller()
    borrowing = wb.api.ActiveSheet.OLEObjects("ComboBox1").Object.Value
    amount = wb.sheets['management'].range("D9").value
    type = ''

    try:
        date = wb.sheets['management'].range("B9").value.strftime('%Y-%m-%d')
    except AttributeError as e:
        wb.sheets['management'].range("A18").color = (240, 100, 77)
        wb.sheets['management'].range("A18").value = str(datetime.datetime.now()) + ': ' + str(e)
        return None

    try:
        if wb.api.ActiveSheet.OLEObjects("OptionButton1").Object.Value == True:
            type = '1'
        elif wb.api.ActiveSheet.OLEObjects("OptionButton2").Object.Value == True:
            type = '2'
    except:
        return None

    database = os.path.join(os.path.dirname(wb.fullname), 'borrowing.db')
    # create a database connection
    conn = create_connection(database)
    try:
        with conn:
        # create a new project
            payment = (date, type, amount, borrowing);
            payment_id = create_payment(conn, payment)
            wb.sheets['management'].range("A18").color = (146, 208, 80)
            wb.sheets['management'].range("A18").value = str(datetime.datetime.now()) + ": Создан платеж в размере " + str(amount)

    except sqlite3.IntegrityError as e:
        wb.sheets['management'].range("A18").color = (240, 100, 77)
        wb.sheets['management'].range("A18").value = str(datetime.datetime.now()) + ': ' + str(e)

def create_sup_agreement(conn, client):
    sql = ''' INSERT INTO sup_agreement(title, borrowing, rate, date)
              VALUES(?,?,?,?)'''
    cur = conn.cursor()
    cur.execute(sql, client)
    return cur.lastrowid

def insert_a_sup_agreement():
    wb = xw.Book.caller()
    title = wb.sheets['management'].range("B14").value
    borrowing = wb.api.ActiveSheet.OLEObjects("ComboBox2").Object.Value
    rate = wb.sheets['management'].range("D14").value

    try:
        date = wb.sheets['management'].range("E14").value.strftime('%Y-%m-%d')
    except AttributeError as e:
        wb.sheets['management'].range("E14").color = (240, 100, 77)
        wb.sheets['management'].range("E14").value = str(datetime.datetime.now()) + ': ' + str(e)
        return None

    database = os.path.join(os.path.dirname(wb.fullname), 'borrowing.db')
    # create a database connection
    conn = create_connection(database)
    try:
        with conn:
        # create a new project
            sup_agreement = (title, borrowing, rate, date);
            sup_agreement_id = create_sup_agreement(conn, sup_agreement)
            wb.sheets['management'].range("A18").color = (146, 208, 80)
            wb.sheets['management'].range("A18").value = str(datetime.datetime.now()) + ": Создано доп. соглашение " + str(title)

    except sqlite3.IntegrityError as e:
        wb.sheets['management'].range("A18").color = (240, 100, 77)
        wb.sheets['management'].range("A18").value = str(datetime.datetime.now()) + ': ' + str(e)

def combobox(command, combo_box_name, source_cell):
    wb = xw.Book.caller()
    source = wb.sheets['source']

    # Place the database next to the Excel file
    database = os.path.join(os.path.dirname(wb.fullname), 'borrowing.db')

    # Database connection and creation of cursor
    con = sqlite3.connect(database, detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES)
    cursor = con.cursor()

    # Database Query
    cursor.execute(command)

    #
    source.range(source_cell).expand().clear_contents()
    source.range(source_cell).value = cursor.fetchall()

    #
    combo = combo_box_name
    wb.api.ActiveSheet.OLEObjects(combo).Object.ListFillRange = \
        'Source!{}'.format(str(source.range(source_cell).expand().address))
    wb.api.ActiveSheet.OLEObjects(combo).Object.BoundColumn = 1
    wb.api.ActiveSheet.OLEObjects(combo).Object.ColumnCount = 2
    wb.api.ActiveSheet.OLEObjects(combo).Object.ColumnWidths = 0

    # Close cursor and connection
    cursor.close()
    con.close()

def up_to_date_report():
    wb = xw.Book.caller()
    database = os.path.join(os.path.dirname(wb.fullname), 'borrowing.db')
    conn = create_connection(database, )
    cursor = conn.cursor()
    sql = \
    '''SELECT
        start_date AS 'Дата начала', end_date AS 'Дата окончания',
        title AS '№ договора займа', body AS 'Тело', rate AS 'Ставка', payment.date AS 'Дата платежа',
        payment.amount AS 'Размер платежа', payment.type AS 'Тип платежа' FROM borrowing
       LEFT JOIN payment ON payment.borrowing=borrowing.id'''

    # date = wb.sheets['up_to_date'].range("B3").value.strftime('%Y-%m-%d')
    query = cursor.execute(sql, )
    cols = [column[0] for column in query.description]
    data = pd.DataFrame(query.fetchall(), columns=cols)
    wb.sheets['up_to_date'].range('A6').options(index=False).value = data
