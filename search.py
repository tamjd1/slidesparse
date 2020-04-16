from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
import os
from nltk.tokenize import sent_tokenize
from pptx import Presentation


def __find(text, keyword, filename=None, page=0):
    sentences = sent_tokenize(text)
    for s in sentences:
        if keyword.lower() in s.lower():
            print()
            print("- {} --- {}".format(page, filename))
            print()
            for _s in s.split("\n"):
                print("\t {}".format(_s))


def find_in_pptx(path, keyword):  # not tested
    presentation = Presentation(path)
    for i, slide in enumerate(presentation.slides):
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                __find(shape.text, keyword, os.path.basename(path), i)


def find_in_pdf(path, keyword):
    resource_manager = PDFResourceManager()
    s = StringIO()
    la_params = LAParams()
    device = TextConverter(resource_manager, s, laparams=la_params)
    f = open(path, 'rb')
    interpreter = PDFPageInterpreter(resource_manager, device)

    for page in PDFPage.get_pages(f):
        interpreter.process_page(page)

    text = s.getvalue()
    __find(text, keyword, os.path.basename(path))

    f.close()
    device.close()
    s.close()


def main(keyword):
    slides_dir = "/Users/T/Desktop/slides/"  # todo : configure your directory here
    ld = os.listdir(slides_dir)
    for l in ld:
        if l.endswith(".pdf"):
            find_in_pdf(slides_dir + l, keyword)
        elif l.endswith(".ppt") or l.endswith(".pptx"):
            find_in_pptx(slides_dir + l, keyword)


if __name__ == '__main__':
    import sys
    args = sys.argv
    __keyword = args[1]
    print("Searching for {} ...".format(__keyword))
    main(__keyword)
