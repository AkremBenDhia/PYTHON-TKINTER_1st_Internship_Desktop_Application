import tkinter as tk
from tkinter import filedialog, Text
from tkinter import *
from PIL import ImageTk,Image
from tkinter import ttk
import os

import os.path


import numpy as np
import io
import time

from tkinter.messagebox import showinfo




import pandas as pd


root = tk.Tk()
top= Toplevel()
root.title("RTUTOP_Creator_Application")
top.title("Adresses Dupliqées")


apps=[]
def addAppdatalist():
    for widget in frame.winfo_children():
        widget.destroy()
    excel_filepath1 = filedialog.askopenfilename(initialdir="/", title="Choisir Fichier Data_list",
    filetypes=(("xlsx","*.xlsx"),("all files","*.*")))
    showinfo(
        title='Fichier Data_List est bien choisir',
        message=excel_filepath1
    )
    showinfo(
        title='Choisir le fichier RTUTOP',
        message="Entrer le fichier RTUTOP Maintenant"
    )


    apps.append(excel_filepath1)
    print(excel_filepath1)
    for app in apps:

        label =tk.Label(frame, text=app, bg="white")
        label.pack()

    dfliste = pd.read_excel(excel_filepath1, header=None)




    for widget in frame.winfo_children():
        widget.destroy()
    excel_filepath2 = filedialog.askopenfilename(initialdir="/", title="Choisir_Fichier_RTUTOP",
                                                 filetypes=(("xlsx", "*.xlsx"), ("all files", "*.*")))
    showinfo(
        title='Fichier RTUTOP est bien choisir',
        message=excel_filepath2
    )
    showinfo(
        title='Trés Bien',
        message="Ne Fermer pas l'application svp,Votre fichier est en cours de création..."
    )

    apps.append(excel_filepath2)
    print(excel_filepath2)
    for app in apps:
        label = tk.Label(frame, text=app, bg="white")
        label.pack()

    dfrtutop = pd.read_excel(excel_filepath2, header=None)


    list = []
    list2 = []
    for k in range(len(dfliste)):
        list.append(dfliste.loc[k, 7])
    s = 0
    for l in range(len(dfliste)):
        s = 0
        for m in range(len(list)):
            if list[m] == dfliste.loc[l, 7]:
                s = s + 1
        if s > 1:
            x = ("Adresse dupliqué dans le fichier Data_list", dfliste.loc[l, 7], "ligne", l + 1)
            list2.append(x)


    for compt in list2:
        label = Label(top, text=compt, bg="white")
        label.pack()

    dfrtutop = pd.ExcelFile(excel_filepath2)

    with pd.ExcelWriter('output.xlsx') as writer:

        for i in range(len(dfliste)):

            cle = dfliste.loc[i, 3] + dfliste.loc[i, 4] + dfliste.loc[i, 5]

            for sheet in dfrtutop.sheet_names:

                df = pd.read_excel(excel_filepath2, sheet_name=sheet, header=None)

                for j in range(len(df)):

                    cle2 = str(df.loc[j, 6]) + str(df.loc[j, 7]) + str(df.loc[j, 8])

                    if cle == cle2:

                        column = dfliste.loc[i, 7]
                        df.loc[j, 3] = column
                        df.to_excel(writer, sheet_name=sheet)
                    else:

                        df.to_excel(writer, sheet_name=sheet)




canvas=tk.Canvas(root, height=300, width=300, bg="gray")
canvas.pack(fill = "both", expand = True)


canvas.create_text( 100, 50, text = "Bienvenue à RTUTOP_Creator")


frame= tk.Canvas(root, bg="gray")
frame.place(relwidth=0.8, relheight=0.4, relx=0.1, rely=0.3)

frame.create_text( 100, 120, text = "Chemins des fichiers apparaîtront ici")

bouton= tk.Button(root, text="Choisir votre fichier Data_list", padx=10, pady=5, fg="black", bg="gray", command=addAppdatalist)
bouton.pack()



def output():
    for widget in frame.winfo_children():
        widget.destroy()
    outputfile = filedialog.askopenfilename(initialdir="/", title="Choisir Fichier Data_list",
    filetypes=(("xlsx","*.xlsx"),("all files","*.*")))

bouton3= tk.Button(root, text="OUTPUT", padx=10, pady=5, fg="black", bg="gray", command=output)
bouton3.pack()






root.mainloop()

