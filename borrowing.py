import xlwings as xw
import datetime
import os
import sqlite3
import pandas as pd


def create_connection(db_file):
    try:
        conn = \
            sqlite3.connect(db_file, detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES)
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
        start_date = wb.sheets['management'].\
            range("D4").value.strftime('%Y-%m-%d')
        end_date = wb.sheets['management'].\
            range("E4").value.strftime('%Y-%m-%d')
    except AttributeError as e:
        wb.sheets['management'].range("A18").color = (240, 100, 77)
        wb.sheets['management'].range("A18").value = \
            str(datetime.datetime.now()) + ': ' + str(e)
        return None
    database = os.path.join(os.path.dirname(wb.fullname), 'borrowing.db')
    # create a database connection
    conn = create_connection(database)
    try:
        with conn:
            # create a new borrowing
            borrowing = (title, body, rate, start_date, end_date)
            borrowing_id = create_borrowing(conn, borrowing)
            wb.sheets['management'].range("A18").color = (146, 208, 80)
            wb.sheets['management'].range("A18").value = \
                str(datetime.datetime.now()) + \
                ": Создан договор займа № " + str(title)

    except sqlite3.IntegrityError as e:
        wb.sheets['management'].range("A18").color = (240, 100, 77)
        wb.sheets['management'].range("A18").value = \
            str(datetime.datetime.now()) + ': ' + str(e)


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
        wb.sheets['management'].range("A18").value = \
            str(datetime.datetime.now()) + ': ' + str(e)
        return None

    try:
        if wb.api.ActiveSheet.OLEObjects("OptionButton1").\
                Object.Value is True:
            type = '1'
        elif wb.api.ActiveSheet.OLEObjects("OptionButton2").\
                Object.Value is True:
            type = '2'
    except:
        return None

    database = os.path.join(os.path.dirname(wb.fullname), 'borrowing.db')
    # create a database connection
    conn = create_connection(database)
    try:
        with conn:
            # create a new payment
            payment = (date, type, amount, borrowing)
            payment_id = create_payment(conn, payment)
            wb.sheets['management'].range("A18").color = (146, 208, 80)
            wb.sheets['management'].range("A18").value = \
                str(datetime.datetime.now()) + ": Создан платеж в размере " + str(amount)

    except sqlite3.IntegrityError as e:
        wb.sheets['management'].range("A18").color = (240, 100, 77)
        wb.sheets['management'].range("A18").value = \
            str(datetime.datetime.now()) + ': ' + str(e)


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
        wb.sheets['management'].range("E14").value = \
            str(datetime.datetime.now()) + ': ' + str(e)
        return None

    database = os.path.join(os.path.dirname(wb.fullname), 'borrowing.db')
    # create a database connection
    conn = create_connection(database)
    try:
        with conn:
        # create a new sup_agreement
            sup_agreement = (title, borrowing, rate, date);
            sup_agreement_id = create_sup_agreement(conn, sup_agreement)
            wb.sheets['management'].range("A18").color = (146, 208, 80)
            wb.sheets['management'].range("A18").value = str(datetime.datetime.now()) + ": Создано доп. соглашение " + str(title)

    except sqlite3.IntegrityError as e:
        wb.sheets['management'].range("A18").color = (240, 100, 77)
        wb.sheets['management'].range("A18").value = \
            str(datetime.datetime.now()) + ': ' + str(e)


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
        '''
        SELECT * FROM corr_br_sp
        WHERE (start_date < ? OR start_date IS NULL)
        AND (s_date < ? OR s_date IS NULL)
        '''
    date = wb.sheets['up_to_date'].range("B3").value
    query = cursor.execute(sql, [date, date])
    cols = [column[0] for column in query.description]
    data = pd.DataFrame(query.fetchall(), columns=cols)
    try:
        grouped = \
            data.set_index(['title', 'agr_title'], drop=True)
        grouped['valid_from'] = grouped['start_date']
        grouped['valid_until'] = grouped['end_date']
        grouped['valid_rate'] = grouped['rate']
        grouped['valid_days'] = \
            (grouped['valid_until'] - grouped['valid_from']).dt.days
        wb.sheets['up_to_date'].range('A9').options(index=True).value = \
            valid_rate(grouped)
    except AttributeError:
        pass

def join_py_on_sp(date):
    wb = xw.Book.caller()
    database = os.path.join(os.path.dirname(wb.fullname), 'borrowing.db')
    conn = create_connection(database, )
    cursor = conn.cursor()
    sql = \
        '''SELECT * FROM corr_br_py
        WHERE (date < ?)'''
    query = cursor.execute(sql, [date])
    cols = [column[0] for column in query.description]
    payments = pd.DataFrame(query.fetchall(), columns=cols)
    payments = \
        payments.set_index(['title', 'p_id'])
    if payments is None:
        return None
    else:
        return payments


def valid_rate(data):
    wb = xw.Book.caller()
    group = None
    groups = []
    indexes = []
    date = wb.sheets['up_to_date'].range("B3").value.strftime('%Y-%m-%d')
    date = datetime.datetime.date(datetime.datetime.strptime(date, '%Y-%m-%d'))
    payments = join_py_on_sp(date)
    if payments is None:
        return None
    new_data = pd.DataFrame()
    new_data['percents'] = 0

    for item in data.index:
        indexes.append(item[0])
    indexes = list(set(indexes))

    for item in indexes:
        group = data.xs(item, level='title')
        if len(group) >= 2:
            for i in range(0, len(group)):
                group['valid_rate'][i] = group['agr_rate'][i]
                group['valid_from'][i] = group['s_date'][i]
                if i < len(group)-1:
                    group['valid_until'][i] = group['s_date'][i+1]
                    if date >= group['valid_from'][i] \
                            and date <= group['valid_until'][i]:
                        group['valid_until'] = date
                else:
                    group['valid_until'][i] = group['prlng_until'][i]
                    if date >= group['valid_from'][i] \
                            and date <= group['valid_until'][i]:
                        group['valid_until'] = date
                    group['valid_days'] = \
                        (group['valid_until'] - group['valid_from']).dt.days
        else:
            if date >= group['valid_from'][0] \
                    and date <= group['valid_until'][0]:
                group['valid_until'] = date
        groups.append(group)

    new_data = pd.concat(groups).sort_values(['borrowing'])
    new_data.reset_index(drop=True, inplace=True)
    new_data['percents'] = \
        ((new_data['valid_until'] -
            new_data['valid_from']).dt.days / 365) * \
        new_data['body'] * new_data['valid_rate']

    for i, row in payments.iterrows():
            for j, g in new_data.iterrows():
                if row['borrowing'] == g['borrowing'] and g['valid_from'] < row['date'] <= g['valid_until']:
                    if row['type'] == '1':
                        new_data['percents'][j] -= row['amount']
                    elif row['type'] == '2':
                        new_data['body'][j] -= row['amount']
                        try:
                            if new_data['borrowing'][j+1] == row['borrowing']:
                                new_data['body'][j+1] -= row['amount']
                        except KeyError as e:
                            pass

    new_data['percents'] += \
        ((new_data['valid_until'] -
            new_data['valid_from']).dt.days / 365) * \
        new_data['body'] * new_data['valid_rate']
