import datetime

import creationTableau
from functions import *
import containers

C = containers.Containers()

date = datetime.datetime.now()


# Classe interface
class Interface(Frame):
    """Notre fenêtre principale.
    Tous les widgets sont stockés comme attributs de cette fenêtre."""

    def __init__(self, fenetre, ws, wb2, ws2, **kwargs):
        credits = "LE FEYER Aymeric | Dunfast v1.7 | 27/02/2020"
        Frame.__init__(self, fenetre, width=768, height=576, **kwargs)
        self.pack(fill=BOTH)

        self.num_semaine = IntVar()
        self.num_semaine.set(datetime.datetime(date.year, date.month, date.day).isocalendar()[1])

        self.ws = ws
        self.ws2 = ws2
        self.wb2 = wb2

        self.total = []

        # Création de nos frames
        self.frame = LabelFrame(self, borderwidth=2, relief=GROOVE, text="Dunfast")
        self.frame2 = LabelFrame(self, borderwidth=2, relief=GROOVE, text="Containers")
        self.frame3 = LabelFrame(self, borderwidth=2, relief=GROOVE, text="Infos")

        self.message_semaine = Label(self.frame, text="Semaine : ", width=20)
        self.nb_containers = Label(self.frame2, text="Total : " + str(len(C.total)), width=20)
        self.nb_francite = Label(self.frame2, text="Francite : " + str(len(C.francite)), width=20)
        self.nb_scafruit = Label(self.frame2, text="Francite : " + str(len(C.francite)), width=20)
        self.nb_zone_c = Label(self.frame2, text="Francite : " + str(len(C.francite)), width=20)
        self.nb_zone_a = Label(self.frame2, text="Francite : " + str(len(C.francite)), width=20)
        self.nb_st = Label(self.frame2, text="Francite : " + str(len(C.francite)), width=20)
        self.nb_cross = Label(self.frame2, text="Francite : " + str(len(C.francite)), width=20)
        self.nb_direct = Label(self.frame2, text="Francite : " + str(len(C.francite)), width=20)
        self.nb_appel_sq = Label(self.frame2, text="Francite : " + str(len(C.francite)), width=20)
        self.nb_nexy = Label(self.frame2, text="Francite : " + str(len(C.francite)), width=20)
        self.nb_polybag = Label(self.frame2, text="Francite : " + str(len(C.francite)), width=20)
        self.nb_prio = Label(self.frame2, text="Francite : " + str(len(C.francite)), width=20)

        self.bouton_quitter = Button(self.frame, text="Quitter", command=self.fermer)
        self.creditsLabel = Label(self.frame, text=credits)

        self.bouton_cliquer = Button(self.frame, text="Creer les dossiers", command=self.creerLesDossiers)

        self.ligne_semaine = Spinbox(self.frame, textvariable=self.num_semaine, width=10, from_=1, to=57, increment=1)

        self.button_launch = Button(self.frame, text="Creer le tableau", command=self.commencer)

        self.frame3.pack()

        self.frame.pack(side=LEFT)
        self.frame2.pack(side=RIGHT)

        self.message_semaine.pack(padx=30)
        self.ligne_semaine.pack(padx=30, pady=10)
        self.bouton_cliquer.pack(padx=30, pady=10)
        self.button_launch.pack(padx=30, pady=10)
        self.bouton_quitter.pack(side="right", pady=30, padx=30)
        self.creditsLabel.pack(side="right", padx=30, pady=30)

    def fermer(self):
        self.quit()
        self.destroy()

    def creerLesDossiers(self):
        try:
            createFolders(self.num_semaine.get(), self.ws, C)
        except Exception as e:
            Label(self.frame3, text="Ce dossier existe deja, supprime le puis retente. Erreur : " + str(e)).pack(
                padx=30)
        else:
            Label(self.frame3, text="Dossiers créés, vous pouvez y introduire les différents documents").pack(padx=30)

    def commencer(self):
        try:
            self.total = start(self.num_semaine.get(), self.ws, C, self)
            self.updateCompteurs()
            creationTableau.createNewTab(self.ws, self.ws2, self.num_semaine.get())
            saveFile(self.num_semaine.get(), self.wb2)

        except Exception as e:
            Label(self.frame3, text="Erreur : " + str(e)).pack(padx=30)
        else:
            Label(self.frame3,
                  text="Le tableau est généré dans [Semaine " + str(self.num_semaine.get()) + "/Antilles " + str(
                      self.num_semaine.get()) + ".xlsx]").pack(padx=30)

    def updateCompteurs(self):
        self.nb_containers = Label(self.frame2, text="Total : " + str(self.total[0]) + " containers", width=20)
        self.nb_francite = Label(self.frame2, text="Francite : " + str(self.total[1]) + " containers", width=20)
        self.nb_scafruit = Label(self.frame2, text="Scafruit : " + str(self.total[6]) + " containers", width=20)
        self.nb_zone_c = Label(self.frame2, text="Zone C : " + str(self.total[3]) + " containers", width=20)
        self.nb_zone_a = Label(self.frame2,
                               text="Zone A : " + str(self.total[0] - self.total[1] - self.total[6]) + " containers",
                               width=20)
        self.nb_st = Label(self.frame2, text="Sans Tri : " + str(self.total[2][0]) + " containers", width=20)
        self.nb_cross = Label(self.frame2, text="Crossdock : " + str(self.total[2][1]) + " containers", width=20)
        self.nb_direct = Label(self.frame2, text="Direct : " + str(self.total[2][2]) + " containers", width=20)
        self.nb_appel_sq = Label(self.frame2, text="Appel SQ : " + str(self.total[7][0]) + " containers", width=20)
        self.nb_nexy = Label(self.frame2, text="Nexy : " + str(self.total[7][1]) + " containers", width=20)
        self.nb_polybag = Label(self.frame2, text="Polybag : " + str(self.total[8]) + " containers", width=20)
        self.nb_prio = Label(self.frame2, text="Prio : " + str(self.total[7][3]) + " containers", width=20)

        self.nb_containers.pack(padx=30)
        self.nb_francite.pack(padx=30)
        self.nb_scafruit.pack(padx=30)
        self.nb_zone_c.pack(padx=30)
        self.nb_zone_a.pack(padx=30)
        self.nb_st.pack(padx=30)
        self.nb_cross.pack(padx=30)
        self.nb_direct.pack(padx=30)
        self.nb_appel_sq.pack(padx=30)
        self.nb_nexy.pack(padx=30)
        self.nb_polybag.pack(padx=30)
        self.nb_prio.pack(padx=30)


def printException(self, s):
    Label(self.frame3, text=s).pack(padx=30)
