# coding=utf-8
import unicodedata
from openpyxl import *
from openpyxl.styles import PatternFill
import os
import numpy as np
import pyexcel as p
from tkinter import *


def start():
    tryTo(DYNAMAN)
    tryTo(FRET)
    tryTo(CROSS_ST_DIRECT)
    tryTo(CENTRE_EMPOTAGE)
    tryTo(LOTS_BLOQUES)
    tryTo(FENES)
    tryTo(SCAFRUIT)
    COMMENTS()


# Appels des fonctions
def tryTo(f):
    try:
        f()
    except Exception as exception:
        print("Il faut insérer le fichier dans le dossier " + str(f) + ". Erreur " + str(exception))
    else:
        print("Fichier " + str(f) + " accepté")


# Création des dossiers pour la semaine S
def createFolders(s):
    print("Création des dossiers pour la semaine ...")
    semaine = "Semaine " + str(s)
    os.mkdir(semaine)
    os.mkdir(semaine + "/FRET")
    os.mkdir(semaine + "/CROSS_ST_DIRECT")
    os.mkdir(semaine + "/FENES")
    os.mkdir(semaine + "/CENTRE_EMPOTAGE")
    os.mkdir(semaine + "/LOTS_BLOQUES")
    os.mkdir(semaine + "/SCAFRUIT")
    os.mkdir(semaine + "/COMMENTAIRES")
    os.mkdir(semaine + "/AUTRES_CONFIRMATIONS")
    os.mkdir(semaine + "/FICHIERS_DYNAMAN")
    CREATE_SCAFRUIT()
    CREATE_COMMENTS()
    print("Dossiers créés, vous pouvez y introduire les différents documents")


def remove_accents(input_str):
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    only_ascii = nfkd_form.encode('ASCII', 'ignore')
    return only_ascii


def DYNAMAN():
    a = os.listdir(r"Semaine " + str(NumSemaine) + "/FICHIERS_DYNAMAN")
    dyna1 = load_workbook(filename="Semaine " + str(NumSemaine) + "/FICHIERS_DYNAMAN/" + str(a[0]))
    dyna2 = load_workbook(filename="Semaine " + str(NumSemaine) + "/FICHIERS_DYNAMAN/" + str(a[1]))
    feuilleDyna1 = dyna1.active
    feuilleDyna2 = dyna2.active

    ws['D2'] = feuilleDyna2['E2'].value
    ws['E2'] = feuilleDyna2['A2'].value
    ws['C2'] = "SQ"

    j = 1
    fruidor = 0

    for index in range(2, feuilleDyna2.max_row + 1):
        if ws['D' + str(j)].value != feuilleDyna2['E' + str(index)].value:
            j += 1
            ws['D' + str(j)] = int(feuilleDyna2['F' + str(index)].value.replace("_x003", "").replace("_", ""))
            ws['E' + str(j)] = feuilleDyna2['E' + str(index)].value
            ws['C' + str(j)] = "SQ"
            Lots.append(ws['D' + str(j)].value)
            # C.fruidor.append(int(feuilleDyna1['F' + str(index)].value))
            fruidor += 1

    for index in range(2, feuilleDyna1.max_row + 1):
        if ws['D' + str(j)].value != feuilleDyna1['E' + str(index)].value:
            j += 1
            ws['D' + str(j)] = int(feuilleDyna1['F' + str(index)].value.replace("_x003", "").replace("_", ""))
            ws['E' + str(j)] = feuilleDyna1['E' + str(index)].value
            ws['C' + str(j)] = "SQ"
            Lots.append(ws['D' + str(j)].value)
            # C.fruidor.append(int(feuilleDyna1['F' + str(index)].value))
            fruidor += 1

    print("Il y a un total de " + str(fruidor) + " containers")

    return True


# Récupération des informations du fichier Fret
def FRET():
    a = os.listdir(r"Semaine " + str(NumSemaine) + "/FRET")
    fret = load_workbook(filename="Semaine " + str(NumSemaine) + "/FRET/" + str(a[0]))
    feuilleFret = fret.active

    # Importations des informations du fichier Fret vers le fichier principal

    j = 2
    francit = 2

    for index in range(2, feuilleFret.max_row):
        if remove_accents(str(feuilleFret['F' + str(index)].value)) == b'reserve francite':
            C.francite.append(int(feuilleFret['E' + str(index)].value))
            for j in range(2, ws.max_row):
                if ws['D' + str(j)].value in C.francite:
                    if ws['B' + str(j)].value:
                        pass
                    else:
                        ws['B' + str(j)].value = "F"
                        francit += 1

    print("Il y a " + str(francit) + " francite")

    return True


# Récupération des informations dans le fichier CrossDock, ST, Direct
def CROSS_ST_DIRECT():
    cross_st_direct = 1
    b = os.listdir(r"Semaine " + str(NumSemaine) + "/CROSS_ST_DIRECT")
    cross = load_workbook(filename="Semaine " + str(NumSemaine) + "/CROSS_ST_DIRECT/" + str(b[0]))
    feuilleCross = cross.active

    # Récupération des 3 catégories

    st = []
    direct = []
    cross = []

    for index in range(6, 100):
        if feuilleCross['A' + str(index)].value is not None:
            st.append(feuilleCross['A' + str(index)].value)
        if feuilleCross['C' + str(index)].value is not None:
            direct.append(feuilleCross['C' + str(index)].value)
        if feuilleCross['F' + str(index)].value is not None:
            cross.append(feuilleCross['F' + str(index)].value)

    # Importations des informations du fichier Cross vers le fichier principal

    for index in range(len(st)):
        if st[index] not in Lots:
            ws['D' + str(ws.max_row + 1)] = st[index]
            ws['F' + str(ws.max_row)].value = ""
            Lots.append(st[index])
    for index in range(len(direct)):
        if direct[index] not in Lots:
            ws['D' + str(ws.max_row + 1)] = direct[index]
            ws['F' + str(ws.max_row)].value = ""
            Lots.append(direct[index])
    for index in range(len(cross)):
        if cross[index] not in Lots:
            ws['D' + str(ws.max_row + 1)] = cross[index]
            ws['F' + str(ws.max_row)].value = ""
            Lots.append(cross[index])

    for index in range(2, ws.max_row + 1):
        if st:
            if ws['D' + str(index)].value in st:
                if ws['F' + str(index)].value:
                    ws['F' + str(index)].value += "Sans tri"
                else:
                    ws['F' + str(index)].value = "Sans tri"
                ws['C' + str(index)].value = "SQ"
        if direct:
            if ws['D' + str(index)].value in direct:
                if ws['F' + str(index)].value:
                    ws['F' + str(index)].value += "Direct"
                else:
                    ws['F' + str(index)].value = "Direct"
                ws['C' + str(index)].value = "SQ"
        if cross:
            if ws['D' + str(index)].value in cross:
                if ws['F' + str(index)].value:
                    ws['F' + str(index)].value += "CrossDock"
                else:
                    ws['F' + str(index)].value = "CrossDock"
                ws['C' + str(index)].value = "SQ"

    return True


# Récupération des informations pour le centre d'empotage
def CENTRE_EMPOTAGE():
    c = os.listdir(r"Semaine " + str(NumSemaine) + "/CENTRE_EMPOTAGE")
    zoneC = load_workbook(filename="Semaine " + str(NumSemaine) + "/CENTRE_EMPOTAGE/" + str(c[0]))
    feuilleZoneC = zoneC.active

    totalC = 0

    # Importations des informations du fichier Centre d'Empotage vers le fichier principal

    dataZoneC = [[cell.value for cell in row] for row in feuilleZoneC.rows]
    betterDataZoneC = np.ravel(dataZoneC)

    for index in range(2, ws.max_row):
        if ws['D' + str(index)].value in betterDataZoneC:
            if ws['A' + str(index)].value:
                pass
            else:
                ws['A' + str(index)].value = "C"
                totalC += 1

    print("Il y a " + str(totalC) + " zone c")

    return True


# Récupération des informations pour les lots bloqués
def LOTS_BLOQUES():
    d = os.listdir(r"Semaine " + str(NumSemaine) + "/LOTS_BLOQUES")
    bloquer = load_workbook(filename="Semaine " + str(NumSemaine) + "/LOTS_BLOQUES/" + str(d[0]))
    feuilleBloquer = bloquer.active

    # Importations des informations du fichier Lots bloqués vers le fichier principal

    dataBloquer = []
    for i in range(2, feuilleBloquer.max_row):
        dataBloquer.append(feuilleBloquer['A' + str(i)].value)

    for i in range(2, ws.max_row):
        if ws['D' + str(i)].value in dataBloquer:
            ws['E' + str(i)].fill = greyFill

    return True


# Récupérations des informations pour les contremarques spécifiques (fenes)
def FENES():
    e = os.listdir(r"Semaine " + str(NumSemaine) + "/FENES")
    p.save_book_as(file_name="Semaine " + str(NumSemaine) + "/FENES/" + str(e[0]),
                   dest_file_name="Semaine " + str(NumSemaine) + "/FENES/" + "true.xlsx")
    spe = load_workbook(filename="Semaine " + str(NumSemaine) + "/FENES/true.xlsx")
    feuilleFenes = spe.active

    # Importations des informations du fichier Fenes vers le fichier principal

    lotsFenes = []
    letter = 'B'
    for i in range(2, feuilleFenes.max_column):
        lotsFenes.append(feuilleFenes[letter + str(4)].value)
        lotsFenes.append(feuilleFenes[letter + str(19)].value)
        letter = chr(ord(letter) + 1)

    for i in range(2, ws.max_row):
        if ws['D' + str(i)].value in lotsFenes:
            if ws['F' + str(i)].value is None:
                ws['F' + str(i)].value = "Contremarque spé "
            else:
                ws['F' + str(i)].value += " + Contremarque spé "

    return True


# Création du fichier Scafruit, à modifier par Dunfresh
def CREATE_SCAFRUIT():
    sf = Workbook()
    ss = sf.active
    ss.title = "Scafruit"
    ss.column_dimensions['A'].width = 30
    ss.column_dimensions['B'].width = 30
    ss.column_dimensions['C'].width = 30
    ss.column_dimensions['D'].width = 30
    ss.column_dimensions['F'].width = 30
    ss.column_dimensions['G'].width = 30
    ss.column_dimensions['H'].width = 30

    ss['A1'].value = "Numéro du lot"
    ss['B1'].value = "Quantité"
    ss['C1'].value = "Catégorie"
    ss['D1'].value = "Par combien ?"
    ss['E1'].value = "+"
    ss['F1'].value = "Quantité"
    ss['G1'].value = "Catégorie ?"
    ss['H1'].value = "Par combien ?"

    sf.save('Semaine ' + str(NumSemaine) + '/SCAFRUIT/Scafruit.xlsx')


# Récupération des informations pour le Scafruit
def SCAFRUIT():
    f = os.listdir(r"Semaine " + str(NumSemaine) + "/SCAFRUIT")
    scaf = load_workbook(filename="Semaine " + str(NumSemaine) + "/SCAFRUIT/" + str(f[0]))
    feuilleScafruit = scaf.active
    Lot = []
    Qte = []
    Cat = []
    How = []
    ind = 0

    # Récupération des informations du fichier Scafruit.xlsx

    for i in range(2, feuilleScafruit.max_row):
        Lot.append(feuilleScafruit['A' + str(i)].value)
        Qte.append(feuilleScafruit['B' + str(i)].value)
        Cat.append(feuilleScafruit['C' + str(i)].value)
        How.append(str(feuilleScafruit['D' + str(i)].value))

        if feuilleScafruit['F' + str(i)].value:
            How[len(How) - 1] += " + " + str(feuilleScafruit['F' + str(i)].value) + " " + str(
                feuilleScafruit['G' + str(i)].value) + " en " + str(feuilleScafruit['H' + str(i)].value)

    for j in range(2, ws.max_row):
        if ws['D' + str(j)].value in Lot:

            ind = Lot.index(ws['D' + str(j)].value)
            C.scafruit.append(Lot[ind])

            if ws['F' + str(j)].value is None:
                ws['F' + str(j)].value = "SCAFRUIT " + str(Qte[ind]) + " " + str(Cat[ind]) + " en " + str(How[ind])
            else:
                ws['F' + str(j)].value += "+ SCAFRUIT " + str(Qte[ind]) + " " + str(Cat[ind]) + " en " + str(How[ind])


# Création du fichier Commentaires, à modifier par Dunfresh
def CREATE_COMMENTS():
    pr = Workbook()
    ps = pr.active
    ps.title = "Commentaires"
    ps.column_dimensions['A'].width = 30
    ps.column_dimensions['B'].width = 50
    ps.column_dimensions['C'].width = 15
    ps.column_dimensions['D'].width = 15
    ps.column_dimensions['E'].width = 15
    ps.column_dimensions['F'].width = 15
    ps.column_dimensions['G'].width = 15

    ps['A1'].value = "Numéro du lot"
    ps['B1'].value = "Commentaire"
    ps['C1'].value = "Appel SQ"
    ps['D1'].value = "Nexy"
    ps['E1'].value = "Polybag orange"
    ps['F1'].value = "Polybag complet"
    ps['G1'].value = "Prio"

    for i in range(2, 200):
        ps['A' + str(i)].value = int(i - 1)

    pr.save('Semaine ' + str(NumSemaine) + '/COMMENTAIRES/Commentaires.xlsx')


# Récupération des informations pour les Commentaires
def COMMENTS():
    g = os.listdir(r"Semaine " + str(NumSemaine) + "/COMMENTAIRES")
    com = load_workbook(filename="Semaine " + str(NumSemaine) + "/COMMENTAIRES/" + str(g[0]))
    feuilleComments = com.active

    for i in range(2, feuilleComments.max_row):
        a = feuilleComments['A' + str(i)].value

        b = feuilleComments['B' + str(i)].value
        if b:
            addComment(a, b)

        b = feuilleComments['C' + str(i)].value
        if b:
            addComment(a, "Appel SQ ")

        b = feuilleComments['D' + str(i)].value
        if b:
            addComment(a, "Nexy ")

        b = feuilleComments['E' + str(i)].value
        if b:
            addComment(a, "Poly orange ")

        b = feuilleComments['F' + str(i)].value
        if b:
            addComment(a, "Poly complet ")

        b = feuilleComments['G' + str(i)].value
        if b:
            addComment(a, "Prio ")


# Ajouter les commentaires associes
def addComment(n, c):
    if n in Lots:
        if ws['F' + str(Lots.index(n))].value:
            ws['F' + str(Lots.index(n))].value += " + " + str(c)
        else:
            ws['F' + str(Lots.index(n))].value = str(c)


# Colonnes Francite, Zone A et Scafruit
def CREATE_TRUE_FILE():
    C.francite = list(set(C.francite))
    C.francite.sort()

    C.zoneA = list(set(C.zoneA))
    C.zoneA.sort()

    C.scafruit = list(set(C.scafruit))
    C.scafruit.sort()
    ind = 1
    for j in range(2, ws.max_row):
        for n in C.francite:
            if ws['D' + str(j)].value == n:
                ind += 1
                ws2['A' + str(ind)].value = ws['A' + str(j)].value
                ws2['B' + str(ind)].value = ws['B' + str(j)].value
                ws2['C' + str(ind)].value = ws['C' + str(j)].value
                ws2['D' + str(ind)].value = ws['D' + str(j)].value
                ws2['E' + str(ind)].value = ws['E' + str(j)].value
                ws2['F' + str(ind)].value = ws['F' + str(j)].value

    ind = 1
    for j in range(2, ws.max_row):
        for n in C.zoneA:
            if ws['D' + str(j)].value == n:
                ind += 1
                ws2['H' + str(ind)].value = ws['A' + str(j)].value
                ws2['I' + str(ind)].value = ws['B' + str(j)].value
                ws2['J' + str(ind)].value = ws['C' + str(j)].value
                ws2['K' + str(ind)].value = ws['D' + str(j)].value
                ws2['L' + str(ind)].value = ws['E' + str(j)].value
                ws2['M' + str(ind)].value = ws['F' + str(j)].value

    ind = 1
    for j in range(2, ws.max_row):
        for n in C.scafruit:
            if ws['D' + str(j)].value == n:
                ind += 1
                ws2['O' + str(ind)].value = ws['A' + str(j)].value
                ws2['P' + str(ind)].value = ws['B' + str(j)].value
                ws2['Q' + str(ind)].value = ws['C' + str(j)].value
                ws2['R' + str(ind)].value = ws['D' + str(j)].value
                ws2['S' + str(ind)].value = ws['E' + str(j)].value
                ws2['T' + str(ind)].value = ws['F' + str(j)].value


# Classe interface
class Interface(Frame):
    """Notre fenêtre principale.
    Tous les widgets sont stockés comme attributs de cette fenêtre."""

    def __init__(self, fenetre, **kwargs):
        Frame.__init__(self, fenetre, width=768, height=576, **kwargs)
        self.pack(fill=BOTH)
        self.nb_clic = 0

        # Création de nos widgets
        self.message = Label(self, text="Vous n'avez pas cliqué sur le bouton.")
        self.message.pack()

        self.bouton_quitter = Button(self, text="Quitter", command=self.fermer)
        self.bouton_quitter.pack(side="left")

        self.bouton_cliquer = Button(self, text="Cliquez ici", fg="red",
                                     command=start)
        self.bouton_cliquer.pack(side="right")

    def fermer(self):
        self.quit()
        self.destroy()


# Classe Containers
class Containers:
    def __init__(self):
        self.scafruit = []
        self.zoneA = []
        self.zoneC = []
        self.francite = []
        self.total = []


# Debut du programme

C = Containers()

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
greyFill = PatternFill(start_color='969696',
                       end_color='969696',
                       fill_type='solid')

# Creation et ouverture de la fenetre
screen = Tk()
interface = Interface(screen)
interface.mainloop()
interface.destroy()
screen.destroy()

wb.save(filename=fn)
wb2.save(filename=fn2)