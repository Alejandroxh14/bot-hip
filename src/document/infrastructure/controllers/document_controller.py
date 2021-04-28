import sys
sys.path.append("C:\\Users\\Alejandro Herrera\\Documents\\Desarrollo2\\bloomcker\\20210222BotHip\\bot-hip")

from src.document.infrastructure.interfaces.wordsInterface import Words
from src.document.application.usecases.processDocumentUseCases import FixDocumentUseCase

from src.document.infrastructure.interfaces.fixWords_interface import FixWords
import os
# import repositories
# import ui
myDir = "C:\\Users\\Alejandro Herrera\\Documents\\Desarrollo2\\bloomcker\\20210222BotHip\\alcanfores\\alcanfores-459680"

def fixDController(): 
    words = Words(myDir + "\\459680-1.rtf", "alcanfores", True)
    fixDocument = FixDocumentUseCase(words)
    responseFixDocument = fixDocument.execute()

def fixDocumentController():
    #fixtool = FixWords(os.getcwd()+"\\459680.rtf", True)
    fixtool = FixWords(myDir + "\\457968.rtf", True)
    start, end = fixtool.get_range_between_clauses(["primera"], ["segunda"])
    fixtool.show_paragraphs(start, end)
    input("Presiona Enter:")
    fixtool.close_word_app()
    #name
    #date
    #minuta
    #clausula
    #bankDoc
    #signers
    #inmob
    #bank
    #tables
    #images
    # documentUseCase(---)
    pass

def editSignersController():
    #signers
    #inmob
    #bank
    # editSignersUseCase
    pass

def createContracController():
    #contractFormat
    #contractRule
    #document
    # createContractUseCase()
    pass

def openDocumentController():
    #document DB
    # openDocumentUseCase()
    pass

if __name__ =="__main__":
    #funcio() doc
    fixDController()