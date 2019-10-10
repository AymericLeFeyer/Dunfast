# coding=utf-8

from tkinter import *

from openpyxl import *

import containers
import interface

# Debut du programme


NumSemaine = 41


# Création du vrai fichier principal

wb2 = Workbook()
fn2 = "Antilles2.xlsx"
ws2 = wb2.active
ws2.title = "Antilles"

# Création du nouvau fichier principal de la semaine
wb = Workbook()
fn = "Antilles.xlsx"
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

wb.save(filename=fn)
wb2.save(filename=fn2)
