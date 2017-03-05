#coding:utf-8
#python3.5
"""
使用python3.5编写，需额外安装BeautifulSoup及使用lxml解析html（默认的html.parser解析的不好用）。
必须在windows下的python环境执行，否则没有win32com的支持，无法生成word文档，建议直接安装anaconda_python。
直接执行
python jg2word56.py 可将同目录下的index.html导出为word
python jg2word56.py -f index.html index1.html index2.html ...可将多个html合并后导出为word
极光5、6通用版本
"""

import time
from bs4 import BeautifulSoup
import re
from win32com.client import Dispatch
import win32com.client
from io import open
import argparse


class WordWrap:
    """Wrapper aroud Word 8 documents to make them easy to build.
    Has variables for the Applications, Document and Selection;
    most methods add things at the end of the document"""
    def __init__(self, templatefile=None):
        self.wordApp = Dispatch('Word.Application')
        if templatefile == None:
            self.wordDoc = self.wordApp.Documents.Add()
        else:
            self.wordDoc = self.wordApp.Documents.Add(Template=templatefile)

        #set up the selection
        self.wordDoc.Range(0,0).Select()
        self.wordSel = self.wordApp.Selection

    def show(self):
        # convenience when debugging
        self.wordApp.Visible = 1

    def getStyleList(self):
        self.styles = []
        # returns a dictionary of the styles in a document
        stylecount = self.wordDoc.Styles.Count
        for i in range(1, stylecount + 1):
            styleObject = self.wordDoc.Styles(i)
            self.styles.append(styleObject.NameLocal)

    def saveAs(self, filename):
        self.wordDoc.SaveAs(filename)

    def printout(self):
        self.wordDoc.PrintOut()

    def selectEnd(self):
        # ensures insertion point is at the end of the document
        self.wordSel.Collapse(0)
        # 0 is the constant wdCollapseEnd; don't weant to depend
        # on makepy support.

    def addText(self, text):
        self.wordSel.InsertAfter(text)
        self.selectEnd()


    def addStyledPara(self, text, stylename):
        if text[-1] != '\n':
            text = text + '\n'

        self.wordSel.InsertAfter(text)
        self.wordSel.Style = stylename
        self.selectEnd()

    def addTable(self, table, styleid=None):
        # Takes a 'list of lists' of data.
        # first we format the text.  You might want to preformat
        # numbers with the right decimal places etc. first.
        textlines = []
        for row in table:
            textrow = map(str, row)   #convert to strings
            textline = string.join(textrow, '\t')
            textlines.append(textline)
        text = string.join(textlines, '\n')

        # add the text, which remains selected
        self.wordSel.InsertAfter(text)

        #convert to a table
        wordTable = self.wordSel.ConvertToTable(Separator='\t')
        #and format
        if styleid:
            wordTable.AutoFormat(Format=styleid)

        #all table style left
        wordTable.Style = u'网格型'
        wordTable.AutoFitBehavior(1)  #wdAutoFitContent
        wordTable.AutoFitBehavior(2)  #wdAutoFitWindow

        wordTable.Rows(1).Shading.BackgroundPatternColor = '-738132071'

        self.selectEnd()
    def addTable2(self,table,styleid=None):
        rows = len(table) +1
        newtable = self.wordSel.Tables.Add(self.wordSel.Range,rows,5)
        select_index = self.wordDoc.Range()
        newtable.Cell(1,1).Range.InsertAfter(u'序号')
        newtable.Cell(1,2).Range.InsertAfter(u'漏洞')
        newtable.Cell(1,3).Range.InsertAfter(u'严重程度')
        newtable.Cell(1,4).Range.InsertAfter(u'涉及IP')
        newtable.Cell(1,5).Range.InsertAfter(u'解决方案')
        for i in range(len(table)):
            newtable.Cell(i+2,1).Range.InsertAfter(str(i+1))
            newtable.Cell(i+2,1).Range.ParagraphFormat.Alignment = 1
            x = 2
            for j in table[i]:
                newtable.Cell(i+2,x).Range.InsertAfter(j)
                if j in (table[i][1],table[i][2]):
                    newtable.Cell(i+2,x).Range.ParagraphFormat.Alignment = 1
                else:
                    newtable.Cell(i+2,x).Range.ParagraphFormat.Alignment = 0
                x = x+1

        #newtable.AutoFitBehavior(1)    #按内容自动适应表格
        newtable.Style = u'网格型'
        #newtable.AutoFitBehavior(2)    #按窗口自动适应表格

        newtable.Rows(1).Shading.BackgroundPatternColor = '-738132071'
        newtable.Rows(1).Range.ParagraphFormat.Alignment = 1
#        newtable.Rows(1).Range.VerticalAligment = 1
        newtable.Range.Cells.VerticalAlignment = 1   #所有的垂直对齐都是居中对齐
        #newtable.AutoFitBehavior(3)  #  自定义宽度
        newtable.Columns(1).PreferredWidth = 1*28.57
        newtable.Columns(2).PreferredWidth = 3.6*28.57
        newtable.Columns(3).PreferredWidth = 1.2*28.57
        newtable.Columns(4).PreferredWidth = 3*28.57
        newtable.Columns(5).PreferredWidth = 6*28.57

#        newtable.Cells(1).Range.ParagraphFormat.Alignment = 1
        #newtable.Columns(1).Range.Cells.VerticalAlignment = 1   #目前首列无法居中
        select_index.Select()
        #self.wordDoc.Range().Select()  目前的方案也是直接到文档底部
        self.selectEnd()


    def  addInlineExcelChart(self, filename, caption='', height=216, width=432):
        # adds a chart inline within the text, caption below.


        # add an InlineShape to the InlineShapes collection
        #- could appear anywhere
        shape = self.wordDoc.InlineShapes.AddOLEObject(
            ClassType='Excel.Chart',
            FileName=filename
            )
        # set height and width in points
        shape.Height = height
        shape.Width = width

        # put it where we want
        shape.Range.Cut()

        self.wordSel.InsertAfter('chart will replace this')
        self.wordSel.Range.Paste()  # goes in selection
        self.addStyledPara(caption, 'Normal')

    def getFontList(self):
        # returns a dictionary of the styles in a document
        self.fonts = []
        fontcount = self.wordDoc.Fonts.Count
        for i in range(1, fontcount + 1):
            fontObject = self.wordDoc.Fonts(i)
            self.fonts.append(fontObject.NameLocal)


def gendoc(result):
    w = WordWrap()
    w.show()
    w.addStyledPara(u'漏洞扫描报告', u'标题')
    w.addStyledPara(u'XX系统漏洞扫描报告', u'标题 1')
    w.addStyledPara(u'XX时间扫描', u'正文')
    w.addTable2(result)


def htmlread6(soup):
    #soup = BeautifulSoup(file,"html.parser")
    htmlresult = []
    highvul=[]
    midvul=[]
    num = len(soup.find(id="vuln_distribution").find("tbody").contents)
    result = soup.find(id="vuln_distribution").find("tbody").children
    next(result)
    print(u"共{0}个漏洞".format(int((num-2)/3)))
    # content[0]='/n',content[1]=name,content[2]='/n',content[3]=content,content[4]=nextname,content[-1]='/n'
    for i in range(int((num-2)/3)):
        a = next(result)
        onevul=[]
        #print(a)
        if a.find(align="absmiddle")['src'] == "reportfiles/images/vuln_high.gif" or a.find(align="absmiddle")['src'] == "reportfiles/images/vuln_middle.gif":
            #onevul.append(a.find("span").string)
            onevul = []
            if a.find(align="absmiddle")['src'] == "reportfiles/images/vuln_high.gif":
                onevul.append(a.find("span").string)
                onevul.append(u"高")
                next(result)
                b = next(result).find("table")
                c = b.contents[0].text   #受影响ip所在行
                d = re.compile(r'(?<![\.\d])(?:\d{1,3}\.){3}\d{1,3}(?![\.\d])').findall(c)
                #onevul.append(','.join(d))
                onevul.append(d)
                e = b.contents[4].find('td').text.replace('NSFOCUS','').replace('\n\n','\n').replace('\n        \n','\n').strip()
                onevul.append(e)   #加入解决办法
                highvul.append(onevul)
            else:
                onevul.append(a.find("span").string)
                onevul.append(u"中")
                next(result)
                b = next(result).find("table")
                c = b.contents[0].text   #受影响ip所在行
                d = re.compile(r'(?<![\.\d])(?:\d{1,3}\.){3}\d{1,3}(?![\.\d])').findall(c)
                #onevul.append(','.join(d))
                onevul.append(d)
                e = b.contents[4].find('td').text.replace('NSFOCUS','').replace('\n\n','\n').replace('\n        \n','\n').strip()
                onevul.append(e)   #加入解决办法
                midvul.append(onevul)
        else:
            break
    htmlresult.append(highvul)
    htmlresult.append(midvul)
        #print(htmlresult)
    return htmlresult


def htmlread5(soup):
    #soup = BeautifulSoup(file,"html.parser")
    htmlresult = []
    highvul=[]
    midvul=[]
    num = len(soup.find(id="vulnDistribution").find("tbody").contents)
    # content[0]='/n',content[1]=name,content[2]='/n',content[3]=content,content[4]='/n',content[5]=nextname,
    #content[-1]='/n',content[-2]=sum,content[-3]='/n',,
    result = soup.find(id="vulnDistribution").find("tbody").children
    #next(result)
    print(u"共{0}个漏洞".format(int((num-3)/4)))
    for i in range(int((num-3)/4)):
        next(result)
        a = next(result)
        if a.find(class_='vul-vh') or a.find(class_='vul-vm'):
            onevul = []
            #highvul.append(a.find("a").string)
            if a.find(class_='vul-vh'):
                onevul.append(a.find("a").string)
                onevul.append(u"高")
                next(result)
                b = next(result).find("table")
                c = b.contents[1].text #受影响ip所在行
                d = re.compile(r'(?<![\.\d])(?:\d{1,3}\.){3}\d{1,3}(?![\.\d])').findall(c)
                #onevul.append(','.join(d))
                onevul.append(d)
                e = b.contents[5].text.replace('NSFOCUS','').replace('\n\n','\n').replace(u'解决办法','').replace('\n        ','\n').replace('\n   ','\n').strip()
                onevul.append(e)   #加入解决办法
                highvul.append(onevul)
            else:
                onevul.append(a.find("a").string)
                onevul.append(u"中")
                next(result)
                b = next(result).find("table")
                c = b.contents[1].text #受影响ip所在行
                d = re.compile(r'(?<![\.\d])(?:\d{1,3}\.){3}\d{1,3}(?![\.\d])').findall(c)
                #onevul.append(','.join(d))
                onevul.append(d)
                e = b.contents[5].text.replace('NSFOCUS','').replace('\n\n','\n').replace(u'解决办法','').replace('\n        ','\n').replace('\n   ','\n').strip()
                onevul.append(e)   #加入解决办法
                midvul.append(onevul)
        else:
            break
    htmlresult.append(highvul)
    htmlresult.append(midvul)
    return htmlresult


def judge56read(filename):
    file = open(filename,'r',encoding='utf-8')
    soup = BeautifulSoup(file,"lxml")
    if u'系统版本' in soup.find('table').text:
        print(u"使用极光5")
        htmlresult = htmlread5(soup)
    else:
        print(u"使用极光6")
        htmlresult = htmlread6(soup)
    file.close()
    return htmlresult


def hebing(result1,result2):
    longs1 = len(result1)
    longs2 = len(result2)
    tmp=[]
    for i in result1:
        tmp.append(i[0])

    for i in result2:
        for j in range(longs1):
            if i[0] == result1[j][0]:
                result1[j][2] = list(set(i[2]+result1[j][2]))

    for j in result2:
        if j[0] not in tmp:
            result1.append(j)
    return result1


def two2one(htmlresult1,htmlresult2):
    #print(htmlresult2[1])
    temp = []
    highresult = hebing(htmlresult1[0],htmlresult2[0])
    midresult = hebing(htmlresult1[1],htmlresult2[1])
    #print(midresult)
    #highresult.extend(midresult)
    temp.append(highresult)
    temp.append(midresult)

    return temp


def zhengli(htmlresults):
    longs = len(htmlresults)
    a = htmlresults[0]
    print(len(a[0])+len(a[1]))
    for i in range(longs-1):
        a = two2one(a,htmlresults[i+1])
        print(len(htmlresults[i+1][0])+len(htmlresults[i+1][1]))
        print(len(a))
    #print(len(a))
    return a

def chuliip(a):
    htmlresult=[]
    for i in range(len(a)):
        b=[]
        b.append(a[i][0])
        b.append(a[i][1])
        b.append(",".join(a[i][2]))
        b.append(a[i][3])
        htmlresult.append(b)
    return htmlresult


def main():
    print(u"开始时间：" + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) )
    parser = argparse.ArgumentParser(description=u"%(prog)s 整理极光报告")
    parser.add_argument("-f", dest="filenames", default=["index.html"], nargs='*',
                        help=u"极光报告，可多个，用空格隔开。 '-f index.html index1.html index2.html'")
    argv = parser.parse_args()
    print(u"共输入{0}个文件".format(len(argv.filenames)))

    htmlresults = []
    for i in argv.filenames:
        htmlresults.append(judge56read(i))
    #print(htmlresults)
    if len(argv.filenames)<2:
        a = htmlresults[0][0]+htmlresults[0][1]
    else:
        b = zhengli(htmlresults)
        a = b[0]+b[1]
    #print(a)
    htmlresult = chuliip(a)

    #file = open("index.html",'r',encoding='utf-8')
    #htmlresult = htmlread5(file)
    #file.close()
    print(u"开始生成报告：" + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) )
    gendoc(htmlresult)
    #file1 = open("1.log",'w')
    #file1.write(str(htmlresult))
    print(u"结束时间：" + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) )


if __name__ == '__main__':
    main()
