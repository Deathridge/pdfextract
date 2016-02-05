import StringIO as StringIO

import xlwt

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LAParams

laparams = LAParams()
laparams.word_margin = float(1.0)
laparams.char_margin = float(2.0)
#laparams.line_margin = float(0.55)
#laparams.boxes_flow = float(0.7)
#laparams.detect_vertical = True
#laparams.all_texts = True
caching = True
fp = open('C:\Users\daniel.betteridge\Downloads\Aeopi 9 Sep to 23 Sept page set up.pdf', 'rb')
#outfp = open('C:\Users\daniel.betteridge\Documents\pdfextract\Aeopi.csv', 'wb')
rsrc = PDFResourceManager()
restr = StringIO.StringIO()
device =TextConverter(rsrc, restr,laparams=laparams) #replace restr with outfp for file output

interpreter = PDFPageInterpreter(rsrc, device)
book = xlwt.Workbook(encoding="utf-8")

for pageNumber,page in enumerate(PDFPage.get_pages(fp, [1200], password=None, caching=caching, check_extractable=True)):
    if (pageNumber+1)%3 == 0:        
        numcolumns = 8
    else:
        numcolumns = 15

    print pageNumber+1
    interpreter.process_page(page)
    rstr = restr.getvalue()
    
    cleanup = str(rstr)
    
    cleanuplist = cleanup.split("\n\n")
    #print cleanuplist
    newlist = []
    headings = []
    notemptylist = []

    for each in cleanuplist:
        if each != ' ':
            notemptylist.append(each)
    
    if numcolumns !=8:
        for each in notemptylist[numcolumns-1:len(notemptylist)]:
            newlist.append(each.split())
        for each in notemptylist[0:numcolumns-1]:
            new = each.replace("\n"," ")
            #print new
            if new != " ":
                headings.append(new)

        if ((pageNumber)-1)%3==0:
            headings.append(" ")
            headings[numcolumns-3] += " " + headings[numcolumns-2]
            headings[numcolumns-2] =  ""
            for each in newlist[0]:
                headings[numcolumns-2] += " " + each   
            newlist = newlist[1::]
            
    else:
        for each in cleanuplist[12::]:
            newlist.append(each.split())
        headings.append(cleanuplist[0])
        headings.append(cleanuplist[2])
        headings.append(cleanuplist[4])
        headings.append(cleanuplist[5] + ' ' + cleanuplist[6])
        headings.append(cleanuplist[7] + ' ' + cleanuplist[8])
        headings.append(cleanuplist[9] + ' ' + cleanuplist[10])
        headings.append(cleanuplist[11])
    
    completeList = newlist
    #print completeList[0]
    
    sheet1 = book.add_sheet(str(pageNumber+1))
    i=0
    for n in headings:
        i = i + 1
        if str(n) != ' ':
            sheet1.write(0, i, n)
    
    
    for r in range(0,numcolumns):
        
        for c in range(0, len(completeList[r])):
            sheet1.write(c+1, r, completeList[r][c])
    book.save("Aeopi.xls")
    restr.truncate(0)
    
    
    
    
        
        
    
fp.close()
#outfp.close()




