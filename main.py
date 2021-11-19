from functions import *

if __name__ == '__main__':

    date = input("Sanani kiriting: ")

    print("1-Kandedatlarni ismini pdf faylni nomiga yozish\n2-Xatlarni yozish\n")
    ish = int(input("Ish="))

    if ish == 1:
        file = f'main/{date}/sheets.xlsx'
        names = get_all_names(file)
        print(names)
        rename_file(names, date)

    elif ish == 2:
        file = f'main/{date}/results.xlsx'
        Data = load_data(file)
        word = readtxt('base.docx')
        writetxt(Data, 'base.docx')

        files = sorted(os.listdir(f'xatlar/{date}/'))
        combine_all_docx(f'xatlar/{date}/{files[0]}', [x for x in files if x != files[0]])

    else:
        print("Xato son kiritdingiz !")

        from PyPDF2 import PdfFileWriter, PdfFileReader

        inputpdf = PdfFileReader(open("document.pdf", "rb"))

        for i in range(inputpdf.numPages):
            output = PdfFileWriter()
            output.addPage(inputpdf.getPage(i))
            with open("document-page%s.pdf" % i, "wb") as outputStream:
                output.write(outputStream)
