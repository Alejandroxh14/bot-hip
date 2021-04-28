import win32com.client as win32
from win32com.client import constants
from unidecode import unidecode
import os

class FixWords:
    def __init__(self, doc_path, visible):
        """ visible : bool , doc_path : string """
        self.wordApp = win32.gencache.EnsureDispatch('Word.Application')
        self.wordApp.Visible = visible
        self.document = self.wordApp.Documents.Open(doc_path)
        self.paragraphs = self.document.Paragraphs

    def close_word_app(self):
        self.wordApp.Quit(win32.constants.wdDoNotSaveChanges)

    def get_range_between_clauses(self, first_clause, second_clause):
        """ first_clause : list of strings, second_clause : list of strings """
        self.primera, self.segunda = first_clause, second_clause
        self.start, self.end, self.off, self.found = 0, 0, True, False

        for paragraph in self.paragraphs:
            for element in self.primera:
                if unidecode(element.upper()) in unidecode(paragraph.Range.Text) and self.off:
                    self.start = paragraph.Range.Start
                    self.off = False
            for ele in self.segunda:
                if unidecode(ele.upper()) in unidecode(paragraph.Range.Text) and not self.off and not self.found:
                    self.end = paragraph.Range.Start
                    self.found = True
                break 
  
        return self.start, self.end

    def show_paragraphs(self, start, end):
        try:
            if start != 0 and end != 0:
                for par in self.doc.Range(start,end).Paragraphs:
                    print(par.Range.Text)
                    print("------------------")
        except Exception:
            print("No encontro match pero hubo error")

    def doc_to_capital(self, paragraph):
        """
            Convierte el documento a letras mayusculas
            Salida: {
                "name": "string",
                "datetime": "string",
                "data": "word-document"
            }
        """
        pass

    def remove_accents(self, paragraph):
        """
            Remueve los acentos del documento
            Salida: {
                "name": "string",
                "datetime": "string",
                "data": "word-document"
            }
        """
        pass

    def numero(self, paragraph):
        """
            N2, N9 o NG a simbolo de numero No
        """
        pass

    def soles(self, paragraph):
        """
            Convertir 5/ a simbolo de sol S/
        """
        pass

    def format_LA(self, paragraph):
        """
            Convertir 4A a LA
        """
        pass

    def roman_V(self, paragraph):
        """
            numeracion romana (Y) a (V)
        """
        pass


    def arroba_email(self, paragraph): ##############
        """
            Reconocer arroba en correos electronicos
        """
        pass

    def and_or(self, paragraph):
        """
            Cambiar Y/0 por Y/O
        """
        pass

    def remove_admiration(self, paragraph):
        """
            El signo de admiracion con una L
        """
        pass


    def agregar_error(self):
        pass