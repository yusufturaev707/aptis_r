import pandas as pd
import os
import docx
from docxcompose.composer import Composer
from docx import Document as Document_compose
import PyPDF2

sana = ''


def split_pdf(file, date):
    path = f"results_pdf/{date}"
    if not os.path.isdir(path):
        os.mkdir(path=path)
    pdf = PyPDF2.PdfFileReader(open(file, "rb"))
    n = pdf.numPages
    print(f"nn={n}")
    k = 0
    for i in range(n):
        if i % 2 == 0:
            newpdf = PyPDF2.PdfFileWriter()
            newpdf.addPage(pdf.getPage(i))
            newpdf.addPage(pdf.getPage(1))
            k += 1
            with open(f"results_pdf/{date}/page-{k}.pdf", "wb") as f:
                newpdf.write(f)


def get_all_names(file):
    xls = pd.ExcelFile(file)
    sheets = xls.sheet_names
    n = len(sheets)
    names = list()
    for i in range(0, n, 2):
        data = pd.read_excel(xls, sheet_name=sheets[i])
        df = pd.DataFrame(data)
        df.fillna(0, inplace=True)
        names.append(df.iloc[5, 3])
    return names


def rename_file(names, date):
    for i in range(len(names)):
        os.rename(f'results_pdf/{date}/page-{i + 1}.pdf', f'results_pdf/{date}/{names[i]}.pdf')


def get_month(date):
    switcher = {
        1: "yanvar",
        2: "fevral",
        3: "mart",
        4: "aprel",
        5: "may",
        6: "iyun",
        7: "iyul",
        8: "avgust",
        9: "sentabr",
        10: "oktabr",
        11: "noyabr",
        12: "dekabr",
    }
    day = date.split('.')[0]
    month = date.split('.')[1]
    return f"{day}-{switcher.get(int(month), 'Noaniq')}"


def merge_fio(firstname, lastname):
    fio = f"{firstname} {lastname}"
    return fio


def load_data(file):
    data = pd.read_excel(file)
    df = pd.DataFrame(data, columns=['Unnamed: 3', 'Unnamed: 4', 'Unnamed: 10', 'Unnamed: 17'])
    n = df.shape[0]
    print(f'n={n}')
    m = df.shape[1]
    print(f'm={m}')

    date = df.iloc[6, 3]
    kun, oy, yil = date.split('/')
    global sana
    sana = f"{kun}.{oy}.{yil}"

    A = []
    for i in range(5, n, 6):
        a = []
        if df.iloc[i, 2] == 'C':
            a.append(merge_fio(df.iloc[i, 0], df.iloc[i, 1]))
            a.append(df.iloc[i, 2])
            A.append(a)

    df = pd.DataFrame(A, columns=["FIO", "CEFR"])
    df.to_excel(f'cefrs_C/{sana}.xlsx')

    return A


def readtxt(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)


def writetxt(data, filename):
    doc = docx.Document(filename)
    for row in data:
        doc.paragraphs[15].runs[3].text = get_month(sana)
        doc.paragraphs[15].runs[13].text = f"{row[0]}"
        path = f"xatlar/{sana}"
        if not os.path.isdir(path):
            os.mkdir(path=path)
        doc.save(f"{path}/{row[0]}.docx")


def combine_all_docx(filename_master, files_list):
    number_of_sections = len(files_list)
    master = Document_compose(filename_master)
    composer = Composer(master)
    for i in range(0, number_of_sections):
        doc_temp = Document_compose(f'xatlar/{sana}/{files_list[i]}')
        composer.append(doc_temp)
    composer.save(f"xatlar/{sana}/{sana}.docx")
