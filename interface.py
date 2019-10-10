from functions import *
import containers

C = containers.Containers()


# Classe interface
class Interface(Frame):
    """Notre fenêtre principale.
    Tous les widgets sont stockés comme attributs de cette fenêtre."""

    def __init__(self, fenetre, ws, Lots, **kwargs):
        Frame.__init__(self, fenetre, width=768, height=576, **kwargs)
        self.pack(fill=BOTH)
        self.num_semaine = IntVar()

        self.ws = ws
        self.Lots = Lots

        self.total = []

        # Création de nos widgets
        self.frame = LabelFrame(self, borderwidth=2, relief=GROOVE, text="Dunfast")
        self.frame2 = LabelFrame(self, borderwidth=2, relief=GROOVE, text="Containers")

        self.message_semaine = Label(self.frame, text="Semaine : ", width=20)
        self.nb_containers = Label(self.frame2, text="Total : " + str(len(C.total)), width=20)
        self.nb_francite = Label(self.frame2, text="Francite : " + str(len(C.francite)), width=20)

        self.bouton_quitter = Button(self.frame, text="Quitter", command=self.fermer)

        self.bouton_cliquer = Button(self.frame, text="Creer les dossiers", command=self.creerLesDossiers)

        self.ligne_semaine = Spinbox(self.frame, textvariable=self.num_semaine, width=10, from_=0, to=55, increment=1)

        self.button_launch = Button(self.frame, text="Creer le tableau", command=self.commencer)

        self.frame.pack()
        self.frame2.pack()

        self.message_semaine.pack(padx=30)
        self.ligne_semaine.pack(padx=30, pady=10)
        self.bouton_cliquer.pack(padx=30, pady=10)
        self.button_launch.pack(padx=30, pady=10)
        self.bouton_quitter.pack(side="right", pady=30, padx=30)

    def fermer(self):
        self.quit()
        self.destroy()

    def creerLesDossiers(self):
        createFolders(self.num_semaine.get(), self.ws, self.Lots, C)

    def commencer(self):
        self.total = start(self.num_semaine.get(), self.ws, self.Lots, C)
        self.updateCompteurs()

    def updateCompteurs(self):

        self.nb_containers = Label(self.frame2, text="Total : " + str(self.total[0]) + " containers", width=20)
        self.nb_francite = Label(self.frame2, text="Francite : " + str(self.total[1]) + " containers", width=20)

        self.nb_containers.pack(padx=30)
        self.nb_francite.pack(padx=30)

