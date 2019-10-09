# coding=utf-8

from tkinter import *

from openpyxl import *

import interface
import containers

# Debut du programme

C = containers.Containers()

NumSemaine = 41
Lots = []
# createFolders(NumSemaine)

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
interface = interface.Interface(screen, ws, Lots, C)
interface.mainloop()
interface.destroy()
screen.destroy()

wb.save(filename=fn)
wb2.save(filename=fn2)
