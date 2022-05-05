from functions import *
from db_access.db_connect import *
from readEx import load_skills

if __name__ == '__main__':

    date = input("Sanani kiriting: ")
    print("1-split pdf\n"
          "2-Kandedatlarni ismini pdf faylni nomiga yozish\n"
          "3-Xatlarni yozish\n"
          "4-MS Access bazaga yozish\n"
          "0-Dasturdan chiqish\n"
          )

    print("Loading...")
    ish = None
    Users = None

    while ish != 0:

        ish = int(input("Ish="))

        if ish == 1:
            file = f'main/{date}/aptis.pdf'
            split_pdf(file, date)

        elif ish == 2:
            file = f'main/{date}/sheets.xlsx'
            file_id = f'main/{date}/data_id.xlsx'
            names = get_all_names(file)
            load_skills, can_ref = load_skills(file)
            names_id = load_id(file_id)
            # rename_file(names, date)
            rename_file(names, date, names_id, load_skills, can_ref)

        elif ish == 3:
            file = f'main/{date}/results.xlsx'
            C, B, Users = load_data(file, date)
            df = pd.DataFrame(Users, columns=["Familiya", "Ism", "Kun", "Oy", "Yil", "Grammar", "Listening", "Reading",
                                              "Speaking", "Writing", "CEFR"])
            if not os.path.isfile("users/users.xlsx"):
                df.to_excel(f"users/users.xlsx", sheet_name=f"aptis_{date}")
            else:
                with pd.ExcelWriter('users/users.xlsx',
                                    mode='a') as writer:
                    df.to_excel(writer, sheet_name=f"aptis_{date}")
            word = readtxt('base_word/base.docx')
            if len(C) != 0:
                writetxt(C, 'base_word/base.docx')
                files = sorted(os.listdir(f'xatlar/{date}/'))
                combine_all_docx(f'xatlar/{date}/{files[0]}', [x for x in files if x != files[0]])


        elif ish == 4:
            # write_data(Users, date)
            write_newdata(Users)

        else:
            if ish == 0:
                break
