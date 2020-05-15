"""
Author: Zhang Chengsheng, @2020.05.14
python3 script
从PDF中提取图片
"""

import os,sys
from pdfminer.pdfparser import PDFParser,PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal,LAParams,LTFigure,LTRect
import fitz
from PIL import Image


def parse(pdf_file,pngDir):
    fp = open(pdf_file,'rb')
    parser = PDFParser(fp)
    doc = PDFDocument()
    parser.set_document(doc)
    doc.set_parser(parser)
    doc.initialize()

    rsrcmgr = PDFResourceManager()
    laparams = LAParams()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)

    for idx1,page in enumerate(doc.get_pages()):
        print('正在处理{}第{}页...'.format(pdf_file, idx1 + 1))
        interpreter.process_page(page)
        zoom = 5
        pngre = pdf2png(pdf_file, zoom, idx1, pngDir)
        layout = device.get_result()
        drawDict = {}
        for x in layout:
            if (isinstance(x, LTTextBoxHorizontal)):
                if '.fsa' in x.get_text():
                    name = x.get_text().strip().strip('.fsa')
                    drawDict[name] = list(x.bbox)
        for idx,name in enumerate(sorted(drawDict,key=lambda x:drawDict[x][1])):
            heightMulti = 19.22
            heightMinus = 3.00
            topAdd = 0.00
            leftBorderMult = 0.565
            rightBorderMult = 50.37
            #print(name,drawDict[name])
            baseScale = drawDict[name][3]-drawDict[name][1]
            if not idx:
                drawDict[name][2] = drawDict[name][0]+baseScale*rightBorderMult
                drawDict[name][1] = drawDict[name][1]-(baseScale*heightMulti)+heightMinus
                drawDict[name][0] -= baseScale*leftBorderMult
                heightTemp = drawDict[name][3]
                drawDict[name][3] += topAdd
            else:
                drawDict[name][2] = drawDict[name][0]+baseScale*rightBorderMult
                drawDict[name][1] = heightTemp + heightMinus
                drawDict[name][0] -= baseScale*leftBorderMult
                heightTemp = drawDict[name][3]
                drawDict[name][3] += topAdd
            outputPath = os.path.join(pngDir,'{}.png'.format(name))
            locPNG(pngre,drawDict[name],zoom,outputPath)
        os.remove(pngre)


def pdf2png(pdf,zoom,page,pic_path):
    rotate = int(0)
    doc = fitz.open(pdf)
    d = doc[page]
    trans = fitz.Matrix(zoom, zoom).preRotate(rotate)
    pm = d.getPixmap(matrix=trans, alpha=False)
    path = os.path.join(pic_path, '{}_{}.png'.format(pdf.lower().strip('.pdf'),page+1))
    pm.writePNG(path)
    return path


def locPNG(png,position,zoom,cropped_path):
    img = Image.open(png)
    loc = [i*zoom for i in position]
    cropped_img = img.crop((loc[0],img.size[1]-loc[3],loc[2],img.size[1]-loc[1]))
    cropped_img.save(cropped_path)


def run(basedir,pngDir):
    if not os.path.exists(pngDir):
        os.makedirs(pngDir)
    pdfs = [i for i in os.listdir(basedir) if i.endswith('pdf') or i.endswith('PDF')]
    for pdf in pdfs:
        parse(pdf,pngDir)
