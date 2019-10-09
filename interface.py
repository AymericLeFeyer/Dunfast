from functions import *


# Classe interface
class Interface(Frame):
    """Notre fenêtre principale.
    Tous les widgets sont stockés comme attributs de cette fenêtre."""

    def __init__(self, fenetre, ws, Lots, C, **kwargs):
        Frame.__init__(self, fenetre, width=768, height=576, **kwargs)
        self.pack(fill=BOTH)
        self.num_semaine = IntVar()

        self.ws = ws
        self.Lots = Lots
        self.C = C

        # Création de nos widgets
        self.message = Label(self, text="Dunfast")
        self.message.pack()

        self.bouton_quitter = Button(self, text="Quitter", command=self.fermer)
        self.bouton_quitter.pack(side="left")

        self.bouton_cliquer = Button(self, text="Creer les dossiers", command=self.creerLesDossiers)
        self.bouton_cliquer.pack(side="right")

        self.ligne_semaine = Entry(self, textvariable=self.num_semaine, width=10)
        self.ligne_semaine.pack()

        self.button_launch = Button(self, text="Creer le tableau", command=self.commencer)
        self.button_launch.pack()

    def fermer(self):
        self.quit()
        self.destroy()

    def creerLesDossiers(self):
        createFolders(self.num_semaine.get(), self.ws, self.Lots)

    def commencer(self):
        start(self.num_semaine.get(), self.ws, self.Lots, self.C)

