class Document():
    def __init__(self, documentData):
        self.documentData = documentData 
        self.name = self.documentData.name
        self.docType = self.documentData.type
        self.minuta = self.documentData.minuta
        self.clausulas = self.documentData.clausulas
        self.docBanco = self.documentData.docBanco
        self.images = self.documentData.images
        self.tables = self.documentData.tables
        self.comparecientes = self.documentData.comparecientes
        self.banco = self.documentData.banco
        self.inmobiliaria = self.documentData.inmobiliaria

    def to_dict(self):
        self.document_dict = {
            "name" : self.name
        }
        pass

class Comparecientes():
    pass

class Banco():
    pass

class inmobiliaria():
    pass