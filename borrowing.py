import xlwings as xw
import datetime
import os
import sqlite3
import pandas as pd


def create_connection(db_file):
    """Создает полключение к sqlite3 БД"""
    try:
        conn = \
            sqlite3.connect(db_file,
                            detect_types=sqlite3.PARSE_DECLTYPES |
                            sqlite3.PARSE_COLNAMES)
        return conn
    except Error as e:
        print(e)

    return None


def create_borrowing(conn, client):
    """Создает в sqlite3 БД новый договор займа"""
    sql = ''' INSERT INTO borrowing(title, body, rate, start_date, end_date)
              VALUES(?,?,?,?,?)'''
    cur = conn.cursor()
    cur.execute(sql, client)
    return cur.lastrowid


def insert_a_borrowing():
    """Определяет параметры договора займа для последующей вставки в БД"""
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
    conn = create_connection(database)
    try:
        with conn:
            borrowing = (title, body, rate, start_date, end_date)
            borrowing_id = create_borrowing(conn, borrowing)
            sup_agreement = (start_date,
                             '__' + title, borrowing_id, rate, end_date)
            sup_agreement_id = create_sup_agreement(conn, sup_agreement)
            wb.sheets['management'].range("A18").color = (146, 208, 80)
            wb.sheets['management'].range("A18").value = \
                str(datetime.datetime.now()) + \
                ": Создан договор займа № " + str(title)

    except sqlite3.IntegrityError as e:
        wb.sheets['management'].range("A18").color = (240, 100, 77)
        wb.sheets['management'].range("A18").value = \
            str(datetime.datetime.now()) + ': ' + str(e)


def create_payment(conn, client):
    """Создает в sqlite3 БД новый платеж"""
    sql = ''' INSERT INTO payment(date, type, amount, borrowing)
              VALUES(?,?,?,?)'''
    cur = conn.cursor()
    cur.execute(sql, client)
    return cur.lastrowid


def insert_a_payment():
    """Определяет параметры платежа для последующей вставки в БД"""
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
    conn = create_connection(database)
    try:
        with conn:
            payment = (date, type, amount, borrowing)
            payment_id = create_payment(conn, payment)
            wb.sheets['management'].range("A18").color = (146, 208, 80)
            wb.sheets['management'].range("A18").value = \
                str(datetime.datetime.now()) + \
                ": Создан платеж в размере " + str(amount) + \
                " к договору займа " + borrowing

    except sqlite3.IntegrityError as e:
        wb.sheets['management'].range("A18").color = (240, 100, 77)
        wb.sheets['management'].range("A18").value = \
            str(datetime.datetime.now()) + ': ' + str(e)


def create_sup_agreement(conn, client):

    sql = '''INSERT INTO sup_agreement(date, title, borrowing, rate, prlng_until)
             VALUES(?,?,?,?,?)'''
    cur = conn.cursor()
    cur.execute(sql, client)
    return cur.lastrowid


def insert_a_sup_agreement():
    """Создает в sqlite3 БД новое дополнительное соглашение к договору займа"""
    wb = xw.Book.caller()
    title = wb.sheets['management'].range("B14").value
    borrowing = wb.api.ActiveSheet.OLEObjects("ComboBox2").Object.Value
    rate = wb.sheets['management'].range("D14").value

    try:
        prlng_until = wb.sheets['management'].range("E14").value.strftime('%Y-%m-%d')
        date = wb.sheets['management'].range("A14").value.strftime('%Y-%m-%d')
    except AttributeError as e:
        wb.sheets['management'].range("E14").color = (240, 100, 77)
        wb.sheets['management'].range("E14").value = \
            str(datetime.datetime.now()) + ': ' + str(e)
        return None

    database = os.path.join(os.path.dirname(wb.fullname), 'borrowing.db')
    conn = create_connection(database)
    try:
        with conn:
            sup_agreement = (date, title, borrowing, rate, prlng_until)
            sup_agreement_id = create_sup_agreement(conn, sup_agreement)
            wb.sheets['management'].range("A18").color = (146, 208, 80)
            wb.sheets['management'].range("A18").value = \
                str(datetime.datetime.now()) + \
                ": Создано доп. соглашение " + str(title) + \
                " к договору займа " + borrowing

    except sqlite3.IntegrityError as e:
        wb.sheets['management'].range("A18").color = (240, 100, 77)
        wb.sheets['management'].range("A18").value = \
            str(datetime.datetime.now()) + ': ' + str(e)


def combobox(command, combo_box_name, source_cell):
    """Создает новый выпадающий список. Например, с
    названиями договоров займа"""
    wb = xw.Book.caller()
    source = wb.sheets['source']
    database = os.path.join(os.path.dirname(wb.fullname), 'borrowing.db')
    con = sqlite3.connect(database,
                          detect_types=sqlite3.PARSE_DECLTYPES |
                          sqlite3.PARSE_COLNAMES)
    cursor = con.cursor()
    cursor.execute(command)

    # Очищает данные в целевом диапазоне
    # затем вставляет в диапазон данные из sql-запроса
    source.range(source_cell).expand().clear_contents()
    source.range(source_cell).value = cursor.fetchall()

    # Создает выпадающий список
    combo = combo_box_name
    wb.api.ActiveSheet.OLEObjects(combo).Object.ListFillRange = \
        'Source!{}'.format(str(source.range(source_cell).expand().address))
    wb.api.ActiveSheet.OLEObjects(combo).Object.BoundColumn = 1
    wb.api.ActiveSheet.OLEObjects(combo).Object.ColumnCount = 2
    wb.api.ActiveSheet.OLEObjects(combo).Object.ColumnWidths = 0

    # Закрывает соединение с БД
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
        WHERE (s_date <= ?)
        '''
    start_date = wb.sheets['up_to_date'].range("B3").value
    end_date = wb.sheets['up_to_date'].range("C3").value
    query = cursor.execute(sql, [end_date])
    cols = [column[0] for column in query.description]
    data = pd.DataFrame(query.fetchall(), columns=cols)
    try:
        grouped = \
            data.set_index(['title', 'agr_title'], drop=True)
        wb.sheets['up_to_date'].range('A10:O1000').clear_contents()
        wb.sheets['up_to_date'].range('A10').options(index=False).value = \
            valid_rate(grouped)
        wb.sheets['up_to_date'].range('B6').options(index=False).value = \
            valid_rate(grouped)['percents'].sum()
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


def overlap(A_start, A_end, B_start, B_end):
    latest_start = max(A_start, B_start)
    earliest_end = min(A_end, B_end)
    return (latest_start, earliest_end)

def valid_rate(data):
    wb = xw.Book.caller()
    group = None
    groups = []
    indexes = []
    start_date = wb.sheets['up_to_date'].range("B3").value.strftime('%Y-%m-%d')
    start_date = datetime.datetime.date(datetime.datetime.strptime(start_date, '%Y-%m-%d'))
    end_date = wb.sheets['up_to_date'].range("C3").value.strftime('%Y-%m-%d')
    end_date = datetime.datetime.date(datetime.datetime.strptime(end_date, '%Y-%m-%d'))
    data['valid_from'] = None
    data['valid_until'] = None
    data['valid_rate'] = data['agr_rate']
    payments = join_py_on_sp(end_date)
    new_data = pd.DataFrame()
    new_data['percents'] = 0

    for item in data.index:
        indexes.append(item[0])
    indexes = list(set(indexes))

    for item in indexes:
        group = data.xs(item, level='title')
        for i in range((len(group)-1), -1, -1):
            if i == (len(group)-1):
                group['valid_from'][i], group['valid_until'][i] = \
                    overlap(group['s_date'][i], group['prlng_until'][i],
                            start_date, end_date)
            elif i == 0:
                group['valid_from'][i] = group['s_date'][i]
                if len(group) == 1:
                    group['valid_until'][i] = group['s_date'][i+1]
                    group['valid_from'][i], group['valid_until'][i] = \
                        overlap(group['valid_from'][i], group['valid_until'][i],
                                start_date, end_date)
                else:
                    group['valid_until'][i] = group['s_date'][i+1] - datetime.timedelta(days=1)
                    group['valid_from'][i], group['valid_until'][i] = \
                        overlap(group['valid_from'][i], group['valid_until'][i],
                                start_date, end_date)
            elif 0 < i < (len(group)-1):
                group['valid_from'][i] = group['s_date'][i]
                group['valid_until'][i] = group['s_date'][i+1] - datetime.timedelta(days=1)
                group['valid_from'][i], group['valid_until'][i] = \
                    overlap(group['valid_from'][i], group['valid_until'][i],
                            start_date, end_date)

        groups.append(group)

    new_data = pd.concat(groups).sort_values(['borrowing'])
    new_data.reset_index(drop=False, inplace=True)
    new_data['valid_days'] = \
        (new_data['valid_until'] - new_data['valid_from']).dt.days
    new_data['percents'] = \
        ((new_data['valid_until'] -
            new_data['valid_from']).dt.days / 365) * \
        new_data['body'] * new_data['valid_rate']

    for i, row in payments.iterrows():
            for j, g in new_data.iterrows():
                if row['borrowing'] == g['borrowing'] \
                        and g['valid_from'] < row['date'] <= g['valid_until']:
                    if row['type'] == '1':
                        new_data['percents'][j] -= row['amount']
                    elif row['type'] == '2':
                        new_data['body'][(j+1):].loc[new_data['borrowing'] == g['borrowing']] -= row['amount']

    new_data['percents'] += ((new_data['valid_until'] -
            new_data['valid_from']).dt.days / 365) * \
        new_data['body'] * new_data['valid_rate']
    new_data = new_data.loc[(new_data['valid_from'] >= start_date) & (new_data['valid_until'] <= end_date) & (new_data['valid_days'] > 0)]

    return new_data
