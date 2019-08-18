print('pdf reviewer')
import os
from PyPDF2.pdf import PdfFileReader

class color:
   PURPLE = '\033[95m'
   CYAN = '\033[96m'
   DARKCYAN = '\033[36m'
   BLUE = '\033[94m'
   GREEN = '\033[92m'
   YELLOW = '\033[93m'
   RED = '\033[91m'
   BOLD = '\033[1m'
   UNDERLINE = '\033[4m'
   END = '\033[0m'

print(color.BOLD + 'Hello World !' + color.END)


root_folder = 'C:/Users/perso/Google Drive/Book-PDF/'
paths = ['']


def find_term(doc, *args):
    for page_num in range(doc.getNumPages()):
        page = doc.getPage(page_num)
        text = page.extractText()
        for term in args:
            # pos = p.text.lower().find(term)
            pos = text.find(term)
            if pos > -1:
                print(f'Page #{page_num} [{term}] {text[pos-20:pos]}{color.YELLOW}{text[pos:pos+len(term)]}{color.END}{text[pos+len(term):pos + 40]}')


def find_style(doc, *args):
    for p in doc.paragraphs:
        found = False
        for r in p.runs:
            for term in args:
                if r.text.lower().find(term) > -1:
                    print(f'[{term}] ({r.style.name}) {p.text[:70]}')
                    found = True


def show_toc(doc, *argv):
    for p in doc.paragraphs:
        if p.style.name.find('Heading 1') > -1:
            print(p.text)


def run(func, *argv):
    for p in paths:
        folder = os.path.join(root_folder, p)
        print(f'---searching in {folder}')
        files = os.listdir(folder)
        for f in files:
            (name, ext) = os.path.splitext(f)
            if ext == '.pdf':
                print(f'---Analyzing [{os.path.join(p, f)}]')
                try:
                    reader = PdfFileReader(open(os.path.join(folder, f), "rb"))
                except:
                    print(f'error analyizng file {f}')
                    continue
                func(reader, *argv)


run(find_term, "ensor class")
# run(find_style, 'the case for o')
# run(show_toc, 0)