from datetime import date

import pyodbc


def write_data(Users, dateaptis):
    try:
        con_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D:\db_aptis\aptisdb.accdb;'
        conn = pyodbc.connect(con_string)
        print("MS Accessga muvafaqqiyatli bog'landi.")

        cur = conn.cursor()
        cur.execute('SELECT * FROM users')
        users = cur.fetchall()
        kun, oy, yil = dateaptis.split('.')
        now = date.today()
        time = date(int(yil), int(oy), int(kun))

        for user in Users:
            for row in users:
                if user[0] == row[1] and user[1] == row[2] and user[2] == int(row[6]) and user[3] == int(row[7]) and \
                        user[4] == int(row[8]):
                    cur.execute(
                        'UPDATE users SET Grammar = ?, Listening = ?, Reading = ?, Speaking = ?, Writing = ?, CEFR = ? '
                        'WHERE  Familiya = ? and  Ism = ? and Kun = ? and Oy = ? and Yil = ?',
                        (user[5], user[6], user[7], user[8], user[9], user[10], user[0], user[1], user[2], user[3],
                         user[4])
                    )
                    conn.commit()
                    print("Ma'lumot yozildi.")
                    break
                elif user[0] == row[1] and user[1] == row[2] and row[14] is None and now > time:
                    cur.execute(
                        'UPDATE users SET Kun = ?, Oy = ?, Yil = ?,  Grammar = ?, Listening = ?, Reading = ?, '
                        'Speaking = ?, Writing = ?, CEFR = ? '
                        'WHERE  Familiya = ? and  Ism = ?',
                        (user[2], user[3], user[4], user[5], user[6], user[7], user[8], user[9], user[10], user[0],
                         user[1])
                    )
                    conn.commit()
                    print("Ma'lumot yozildi.")
                    break
                else:
                    continue

    except pyodbc.Error as e:
        print("Error in Connection", e)


def write_newdata(Users):
    try:
        con_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D:\db_aptis\aptisdb.accdb;'
        conn = pyodbc.connect(con_string)
        print("MS Accessga muvafaqqiyatli bog'landi.")

        # Users_tuple = []
        # for user in Users:
        #     Users_tuple.append(tuple(user))
        # Users = tuple(Users_tuple)

        cur = conn.cursor()

        for user in Users:
            sql = "INSERT INTO users (Familiya, Ism, Kun, Oy, Yil, Grammar, Listening, Reading, Speaking, Writing, CEFR) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

            cur.execute(
                sql,
                (user[0], user[1], user[2], user[3], user[4], user[5], user[6], user[7], user[8], user[9], user[10])
            )
            conn.commit()

        print(f"{Users[0][2]}/{Users[0][3]}/{Users[0][4]} dagi userlar ma'lumotlari yozildi.")

    except pyodbc.Error as e:
        print("Error in Connection", e)
