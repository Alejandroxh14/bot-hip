import win32com.client as win32
import re

class Words:
    def __init__(self, filePath, inmobiliaria, visible):
        self.filePath = filePath
        self.inmobiliaria = inmobiliaria
        self.wordApp = win32.gencache.EnsureDispatch("Word.Application")
        self.wordApp.Visible = visible
        self.document = self.wordApp.Documents.Open(filePath)
        self.paragraphs = self.document.Paragraphs
        print(self.document.Words.Count)

    def fix(self):
        print("parrafos: ", self.paragraphs.Count)
        self.testing2(self.paragraphs)
        for paragraph in self.paragraphs:
            self.toCapital(paragraph)
            #self.changeYO(paragraph)
            #self.changeExclamationMarkByL(paragraph)
            #self.changeNumberSimbol(paragraph)
            #self.changeSolSimbol(paragraph)
            #self.changeYByVromanNumber(paragraph)
            #self.changeFlat(paragraph)
            #self.removeAccents(paragraph)
            self.removeBold(paragraph)
            self.formatLetter(paragraph)

    def toCapital(self, paragraph):
        paragraph.Range.Font.AllCaps = True
        return paragraph

    def switchExclamationMarkByL(self, paragraph):
        """txt = paragraph.Range.Text
        subTxt = re.sub("!", "L", txt)
        #print(subTxt)
        paragraph.Range.Text = subTxt
        return paragraph
        """
        rangeParagraph = paragraph.Range
        for index in range(rangeParagraph.End):
            if self.document.Range(index, index+1).Text == "!":
                self.document.Range(index, index+1).Text = "L"

    def changeYO(self, paragraph):
        paragraph.Range.Find.Execute(FindText="Y/0", ReplaceWith="Y/O", Replace=2)

    def changeExclamationMarkByL(self, paragraph):
        paragraph.Range.Find.Execute(FindText="!", ReplaceWith="L", Replace=2)

    def changeNumberSimbol(self, paragraph):
        paragraph.Range.Find.Execute(FindText="N2", ReplaceWith="N°", Replace=2)
        paragraph.Range.Find.Execute(FindText="NG ", ReplaceWith="N°", Replace=2)
        paragraph.Range.Find.Execute(FindText="N2 ", ReplaceWith="N°", Replace=2)

    def changeSolSimbol(self, paragraph):
        paragraph.Range.Find.Execute(FindText="5/", ReplaceWith="S/", Replace=2)

    def change4ByL(self, paragraph):
        paragraph.Range.Find.Execute(FindText="4A", ReplaceWith="LA", Replace=2)

    def changeYByVromanNumber(self, paragraph):
        paragraph.Range.Find.Execute(FindText="(Y)", ReplaceWith="(V)", Replace=2)

    def changeFlat(self, paragraph):
        paragraph.Range.Find.Execute(FindText="FIAT", ReplaceWith="FLAT", Replace=2)

    def removeAccents(self, paragraph):
        paragraph.Range.Find.Execute(FindText="Á", ReplaceWith="A", Replace=2)
        paragraph.Range.Find.Execute(FindText="É", ReplaceWith="E", Replace=2)
        paragraph.Range.Find.Execute(FindText="Í", ReplaceWith="I", Replace=2)
        paragraph.Range.Find.Execute(FindText="Ó", ReplaceWith="O", Replace=2)
        paragraph.Range.Find.Execute(FindText="Ú", ReplaceWith="U", Replace=2)

    def removeBold(self, paragraph):
        paragraph.Range.Font.Bold = False

    def formatLetter(self, paragraph):
        paragraph.Range.Font.Name = "Anonymous"
        paragraph.Range.Font.Size = 8
    

    def testing0(self, paragraph):
        palabra = 10
        print(self.document.Words.Item(palabra).Text == "EXTENDER ")
        for index in range(self.document.Words.Item(palabra).Start, self.document.Words.Item(palabra).End):
            print("letra: ", self.document.Range(index, index+1))
        print(len(self.document.Words.Item(palabra).Text), len("EXTENDER "))
        print(self.document.Words.Item(palabra).Start)
        print(self.document.Words.Item(palabra).End)
        count = self.document.Words.Count
        mil=1000
        print(count, range(count))
        for num in range(count):
            if self.document.Words.Item(num+1).Text == "EXTENDER ":
                print("cambiar")
            if num == mil:
                print(num, " mil palabras")
                mil = mil+1000

            #if self.document.Words.Item(num+1) == "Y/0":
            #    print("cambiar")

    def testing1(self):
        paragraph = self.paragraphs(1)
        print(paragraph)
        rangeParagraph = paragraph.Range
        rangeDocument = self.document.Range
        #listFind = rangeParagraph.Find.Execute(FindText="IDENTIFICADO", ReplaceWith="HOLA", Replace=2)
        #print(listFind)
        for paraf in self.paragraphs:
            paraf.Range.Find.Execute(FindText="REGIST", ReplaceWith="HOLA", Replace=2)
    
    def testing2(self, paragraph):
        print(self.wordApp.ActiveDocument.Name)
        #self.wordApp.ActiveDocument.Paragraphs(1).Range.Bold = True
        self.paragraphs(1).Range.Font.Bold = False
        #self.document.ActiveWindow.Selection.Font.Bold