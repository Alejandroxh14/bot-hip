class FixDocumentUseCase:
    def __init__(self, wordsInterface):
        self.wordsInterface = wordsInterface
        #self.inmobiliaria = inmobiliaria
        pass
    
    #def execute(self, name, date, minuta, clausula, bankDoc, signers, inmob, bank, tables, images):
    def execute(self):
        try:
            fixWords = self.wordsInterface.fix()
            return fixWords
        except Exception as exc:
            print(exc)

class editSignersUseCase():
    def __init_(self):
        pass

    def execute(self, document_signers, document_inmob, document_bank):
        try:
            print("editSigners")
        except Exception as exc:
            print(exc)

class createContractUseCase():
    def __init__(self):
        pass

    def execute(self, document, contractFormat, contracRules):
        try:
            print("create contract")
        except Exception as exc:
            print(exc)

class openDocumentUseCase():
    def __init__(self):
        pass
    
    def execute(self, document_name):
        try:
            print("open document")
        except Exception as exc:
            print(exc)