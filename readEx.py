import pandas as pd
import json


def load_skills(file):
    xls = pd.ExcelFile(file)
    sheets = xls.sheet_names
    n = len(sheets)
    results = []
    can_ref = []

    for i in range(0, n, 2):
        data = pd.read_excel(xls, sheet_name=sheets[i])
        df = pd.DataFrame(data)
        df.fillna(0, inplace=True)

        t = []
        candidate_r_n = df.iloc[5, 37]
        fio = df.iloc[5, 3]

        t.append(candidate_r_n)
        t.append(fio)

        can_ref.append(t)

        listening = df.iloc[25, 18]
        reading = df.iloc[27, 18]
        speaking = df.iloc[28, 18]
        writing = df.iloc[29, 18]
        total = df.iloc[30, 18]
        grammar = df.iloc[31, 18]

        person = []

        skills = {
            "Listening": listening,
            "Reading": reading,
            "Speaking": speaking,
            "Writing": writing,
            "Final Scale Score": total,
            "Grammar and Vocabulary": grammar
        }

        for name, ball in skills.items():
            r = {"name": name, "ball": f"{ball}"}
            person.append(r)
        results.append(person)
    return results, can_ref
