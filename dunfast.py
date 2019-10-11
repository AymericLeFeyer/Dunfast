# coding=utf-8

from tkinter import *

import openpyxl
from openpyxl import *

import creationTableau
import interface

# Debut du programme

NumSemaine = 41

# Création du fichier principal

wb2 = Workbook()
fn2 = "Antilles.xlsx"
ws2 = wb2.active
ws2.title = "Antilles"

# Création du nouvau fichier temporaire
wb = Workbook()
fn = "AntillesTemp.xlsx"
ws = wb.active
ws.title = "Antilles"

# Style du document principal

ws.column_dimensions['E'].width = 20
ws.column_dimensions['F'].width = 50

# Creation et ouverture de la fenetre
screen = Tk()
interface = interface.Interface(screen, ws)
interface.mainloop()
interface.destroy()
screen.destroy()

creationTableau.createNewTab(ws, ws2)

wb.save(filename=fn)
wb2.save(filename=fn2)
