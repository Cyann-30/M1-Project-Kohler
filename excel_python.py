# -*- coding: utf-8 -*-
"""
Created on Tue Mar  7 17:20:02 2023

@author: cyann
"""

import pandas as pd
 

def moteur(excel_file_motor_path, column_name_motor):
    """
    Fonction qui créer un fichier MOTEUR_CSV.csv à partir du fichier CAT_MOTEUR.xlsx
    Le nouveau fichier CSV contiendra des données selon des entêtes spécifiques du fichier initial
    Prend en paramètre le chemin du fichier Excel et les entêtes chhoisies
    """
    #Lecture du fichier
    dfm = pd.read_excel(excel_file_motor_path)

    #Séléction des colonnes que l'on souhaite traiter 
    df = dfm[column_name_motor]
    
    #Boucle de tri : seulement les élements ACTIF
    i=0
    while i<=(len(df)-1):
        if df["ID_STATUT\nStatus"][i] != "ACTIF":
            df = df.drop(labels=i,axis=0)
        i+=1
    
    #Ajout de la colonne type 
    df['type'] = 'moteur'
    #Conversion des données au format CSV 
    csv_data = df.to_csv(header=False, index=False)
    
    #Enregistrement des données dans le fichier CSV MOTEUR_CSV.csv
    with open("C:/Users/Cyann/Documents/ISEN/M1/Projet M1/SOURCES_KOHLER/CSV/MOTEUR_CSV.csv", "w") as f:
        f.write(csv_data)

    
def alternateur(excel_file_alternator_path):
    """
    Fonction qui créer un fichier ALTERNATEUR_CSV.csv à partir du fichier CSV ALTERNATEUR.csv
    Le nouveau fichier CSV contiendra des données selon des entêtes spécifiques du fichier initial
    Prend en paramètre le chemin du fichier initial
    """
    # ouvrir le fichier CSV d'origine
    with open(excel_file_alternator_path, newline='') as csvfile_origine:
        reader = pd.read_csv(csvfile_origine, delimiter=';')
    
        # récupérer les colonnes que l'on souhaite
        colonne1 = []
        colonne2 = []
        colonne3 = []
        colonneff2 = []

        colonne1.append(reader["ID_ALT Alternator type"])
        colonne2.append(reader["ID_STATUT Status"])
        colonne3.append(reader["ALT_GN_POIDS_MO Net Weight of the alternator with single bearing config (kg)"])
    
        colonnef1 = list(set(colonne1[0]))
        colonnef2 = list(set(colonne2[0]))
        for i in range(len(colonnef1)):
            colonneff2.append(colonnef2[0])
        colonnef3 = list(set(colonne3[0]))
      
        type_element = "alternateur"  
        consommation = 0
        colonnef4 = []  
        colonnef5=[]
    
        for i in range(len(colonnef1)):
            colonnef4.append(type_element)
            colonnef5.append(consommation)
    
        df = pd.DataFrame({'ID': colonnef1, 'Status': colonneff2, 'Poids': colonnef3, 'Conso': colonnef5, 'Type': colonnef4})
    
        #Conversion des données au format CSV 
        csv_data = df.to_csv(header=False, index=False)
    
        #Enregistrement des données dans le fichier CSV ALTERNATEUR_CSV.csv
        with open("C:/Users/Cyann/Documents/ISEN/M1/Projet M1/SOURCES_KOHLER/CSV/ALTERNATEUR_CSV.csv","w") as f:
            f.write(csv_data)

    
def radiateur(excel_file_cooling_path, column_name_cooling):
    """
    Fonction qui créer un fichier RADIATEUR_CSV.csv à partir du fichier REF_COOLING.xlsx
    Le nouveau fichier CSV contiendra des données selon des entêtes spécifiques du fichier initial
    Prend en paramètre le chemin du fichier Excel et les entêtes chhoisies
    """
    #Lecture du fichier
    dfr = pd.read_excel(excel_file_cooling_path)

    #Séléction des colonnes que l'on souhaite traiter 
    df = dfr[column_name_cooling]
    
    #Boucle de tri : seulement les élements ACTIF
    i=0
    while i<=(len(dfr)-1):
        #if df["ID_STATUT\nStatus"][i] != "ACTIF" or df["ID_REFROID\nType of Cooling"][i] != "RADIA" or df["ENC_TYPE\nEnclosure"][i] != "BASE":
        if df["ID_STATUT\nStatus"][i] != "ACTIF":
            df = df.drop(labels=i,axis=0)
        i+=1

    #ajout des colonnes conso et type 
    df['conso'] = 0
    df['type'] = 'radiateur'
    df = df.drop_duplicates("REF_REFROID\nCooling part number")
    #df = df.drop("ENC_TYPE\nEnclosure", axis=1)

    #Conversion des données au format CSV 
    csv_data = df.to_csv(header=False, index=False)

    #Enregistrement des données dans le fichier CSV RADIATEUR_CSV.csv
    with open("C:/Users/Cyann/Documents/ISEN/M1/Projet M1/SOURCES_KOHLER/CSV/RADIATEUR_CSV.csv","w") as f:
        f.write(csv_data)


def capot(excel_file_capot_path):
    
    #Colonnes que l'on souhaite traiter 
    column_name_radiator = ["IDENT\nItem","ID_STATUT\nStatus","ENC_TYPE\nEnclosure","GEN_ORD\nOrder Number","ENC_PDSBRT_IN\nGross weight (kg)"]

    #Lecture du fichier
    dfg = pd.read_excel(excel_file_capot_path)

    #Séléction des colonnes que l'on souhaite traiter 
    df = dfg[column_name_radiator]
    df.rename(columns={'IDENT\nItem': 'ID', 'ID_STATUT\nStatus': 'Status','ENC_PDSBRT_IN\nGross weight (kg)': 'pds','ENC_TYPE\nEnclosure' : 'capot',"GEN_ORD\nOrder Number":'nb'}, inplace=True)
    i=0
    while i<=(len(dfg)-1):
        if df["Status"][i] != "ACTIF":
            df = df.drop(labels=i,axis=0)
        i+=1
        
    df_filtre = df[~df['capot'].str.contains('DW')]
    df_filtre = df_filtre.reset_index(drop=True)
    liste_nom_capot = []
    liste_poids_capot = []
    j=0
    while j<len(df_filtre):
        
        if df_filtre['nb'][j] == 2.0 and df_filtre['nb'][j+1] == 1.0:
            liste_nom_capot.append(df_filtre['capot'][j])
            liste_poids_capot.append(df_filtre['pds'][j]-df_filtre['pds'][j+1])

        if df_filtre['nb'][j] == 2.0 and df_filtre['nb'][j+1] != 1.0:
            if df_filtre['nb'][j+2] == 1.0:
               liste_nom_capot.append(df_filtre['capot'][j])
               liste_poids_capot.append(df_filtre['pds'][j]-df_filtre['pds'][j+2])
        j+=1
        
    dataframe = pd.DataFrame({'id': liste_nom_capot, 'poids': liste_poids_capot})

    status = []
    element = 'ACTIF' 
    for i in range(len(dataframe)):
        status.append(element)
        
    conso = []
    cons = 0 
    for i in range(len(dataframe)):
        conso.append(cons)
        
    typ = []
    ty = 'capot'
    for i in range(len(dataframe)):
        typ.append(ty)
    
    dataframe.insert(loc=1, column='Status', value=status)
    dataframe = dataframe.assign(conso=conso)
    dataframe = dataframe.assign(type=typ)
    dataframe = dataframe.drop_duplicates('id')
    
    #Conversion des données au format CSV 
    csv_data = dataframe.to_csv(header=False, index=False)

    #Enregistrement des données CSV dans un fichier
    with open("C:/Users/Cyann/Documents/ISEN/M1/Projet M1/SOURCES_KOHLER/CSV/CAPOT_CSV.csv","w") as f:
        f.write(csv_data)


def groupe(excel_file_encombrement_path, column_name_encombrement):
    """
    Fonction qui créer un fichier GROUPE_CSV.csv à partir du fichier ENCOMBREMENT.xlsx
    Le nouveau fichier CSV contiendra des données selon des entêtes spécifiques du fichier initial
    Prend en paramètre le chemin du fichier Excel et les entêtes chhoisies
    """
    #Lecture du fichier
    dfm = pd.read_excel(excel_file_encombrement_path)

    #Séléction des colonnes que l'on souhaite traiter 
    df = dfm[column_name_encombrement]
    
    #Boucle de tri : seulement les élements ACTIF
    i=0
    while i<=(len(df)-1):
        if df["ID_STATUT\nStatus"][i] != "ACTIF":
            df = df.drop(labels=i,axis=0)
        i+=1
    
    #Conversion des données au format CSV 
    csv_data = df.to_csv(header=False, index=False)
    
    #Enregistrement des données dans le fichier CSV GROUPE_CSV.csv
    with open("C:/Users/Cyann/Documents/ISEN/M1/Projet M1/SOURCES_KOHLER/CSV/GROUPE_CSV.csv", "w") as f:
        f.write(csv_data)


def groupe_element(excel_file_group_path, excel_file_cooling_path):
    """
    Fonction qui créer un fichier GROUPE_ELEM_CSV.csv à partir du fichier CAT_GROUPE.xlsx
    Le nouveau fichier CSV contiendra des données selon des entêtes spécifiques du fichier initial
    Prend en paramètre le chemin du fichier Excel
    """
    #Colonnes que l'on souhaite traiter 
    column_name_groupe_moteur = ["IDENT\nItem",'ID_MOT\nEngine type']
    colomn_name_groupe_alternateur = ["IDENT\nItem","ID_ALT\nAlternator type"]
    colomn_name_groupe_radiateur = ["IDENT\nItem", "REF_REFROID\nCooling part number"]
    colomn_name_groupe_pupitre = ["IDENT\nItem"]

    #Séléction des colonnes que l'on souhaite traiter
    dfg = pd.read_excel(excel_file_group_path)
    df1 = dfg[column_name_groupe_moteur]
    df1.rename(columns={'IDENT\nItem': 'ID', 'ID_MOT\nEngine type': 'element'}, inplace=True)
    
    #Pour l'altérnateur
    dfa = pd.read_excel(excel_file_group_path)
    df2 = dfa[colomn_name_groupe_alternateur]
    df2.rename(columns={'IDENT\nItem': 'ID', 'ID_ALT\nAlternator type': 'element'}, inplace=True)
    
    #Pour le radiateur
    dfr = pd.read_excel(excel_file_group_path)
    dfc = pd.read_excel(excel_file_cooling_path)
    
    # Suppression des lignes où la colonne "ID_STATUT\nStatus" a une valeur différente de 'ACTIF'
    dfc = dfc[dfc['ID_STATUT\nStatus'] == 'ACTIF']
    
    # Fusionner les deux dataframes en utilisant les colonnes "DESC_GEN\nDescription" et "IDENT_GE\nGenset model" comme colonnes communes
    merged_df = pd.merge(dfr, dfc, left_on='DESC_GEN\nDescription', right_on='IDENT_GE\nGenset model')
           
    df3 = merged_df[colomn_name_groupe_radiateur]
    df3.rename(columns={'IDENT\nItem': 'ID', "REF_REFROID\nCooling part number": 'element'}, inplace=True)
    df3 = df3.drop_duplicates()
    
    #Pour le pupitre
    dfr = pd.read_excel(excel_file_group_path)
    df4 = dfr[colomn_name_groupe_pupitre]
    df4.rename(columns={'IDENT\nItem': 'ID'}, inplace=True)
    pupitre = []
    element = 'PUPITRE'
    for i in range(len(df4)):
        pupitre.append(element)
    df4 = df4.assign(element=pupitre)
    
    #Pour le chassis
    dfr = pd.read_excel(excel_file_group_path)
    df5 = dfr[colomn_name_groupe_pupitre]
    df5.rename(columns={'IDENT\nItem': 'ID'}, inplace=True)
    chassis = []
    element = 'CHASSIS'
    for i in range(len(df5)):
        chassis.append(element)
    df5 = df5.assign(element=chassis)
    
    nouvelle_dataframe =  pd.concat([df1, df2, df3, df4, df5])
    
    # inversion des données dans chaque ligne
    nouvelle_dataframe = nouvelle_dataframe.apply(lambda x: x[::-1], axis=1)
    
    #Conversion des données au format CSV 
    csv_data = nouvelle_dataframe.to_csv(header=False, index=False)

    #Enregistrement des données dans le fichier CSV GROUPE_ELEM_CSV.csv
    with open("C:/Users/Cyann/Documents/ISEN/M1/Projet M1/SOURCES_KOHLER/CSV/GROUPE_ELEM_CSV.csv","w") as f:
        f.write(csv_data)


def element_matiere(CSV_file_motor_path, CSV_file_alternator_path, CSV_file_radiator_path, CSV_file_capot_path):
    
    #Lecture du fichier CSV
    data_frame_motor = pd.read_csv(CSV_file_motor_path, header=None)
    data_frame_alt = pd.read_csv(CSV_file_alternator_path, header=None)
    data_frame_radiator = pd.read_csv(CSV_file_radiator_path, header=None)
    data_frame_capot = pd.read_csv(CSV_file_capot_path, header=None)
    
    lst_df = [data_frame_motor, data_frame_alt, data_frame_radiator, data_frame_capot]
    mes_variables = {}
    matieres = ['acier','aluminium','cuivre','plastique']
    
    for i, element in enumerate(lst_df):
        nom_variable_1 = f"donnees_colonnes_{i}"
        mes_variables[nom_variable_1] = element.iloc[:, 0].tolist()
        nom_variable_2 = f"lst_{i}"
        mes_variables[nom_variable_2] = [x for x in mes_variables[nom_variable_1] for _ in range(4)]
        nom_variable_3 = f"lst_matieres_{i}"
        mes_variables[nom_variable_3] = []
        
        #Ajout des 4 matieres pour chaque element
        for j in range(len(mes_variables[nom_variable_1])):
            mes_variables[nom_variable_3] += matieres
        
        #Création d'une dataframe
        nom_variable_4 = f"elem_mat_{i}"
        mes_variables[nom_variable_4] = {'elements' : nom_variable_2, 'matieres' : nom_variable_3}
        nom_variable_4 = f"frame_elem_mat_{i}"
        mes_variables[nom_variable_4] = pd.DataFrame({'elements': mes_variables[nom_variable_2], 'matieres': mes_variables[nom_variable_3]})
    
    dataframe_elem_matiere = pd.concat([mes_variables['frame_elem_mat_0'], mes_variables['frame_elem_mat_1'], mes_variables['frame_elem_mat_2'], mes_variables['frame_elem_mat_3']])


    #Conversion des données au format CSV 
    csv_data = dataframe_elem_matiere.to_csv(header=False, index=False)

    #Enregistrement des données CSV dans un fichier
    with open("C:/Users/Cyann/Documents/ISEN/M1/Projet M1/SOURCES_KOHLER/CSV/ELEMENTS_MATIERES.csv","w") as f:
        f.write(csv_data)

        
def type_element():
    """
    Fonction qui créer un fichier CSV TYPE.csv
    Le fichier CSV contiendra les 6 éléments qui composent un GE
    """
    #création d'un DataFrame avec une seule colonne de données
    df = pd.DataFrame({'Colonne': ['alternateur', 'capot', 'chassis', 'moteur', 'pupitre', 'radiateur']})

    #écriture du DataFrame dans un fichier CSV
    df.to_csv("C:/Users/Cyann/Documents/ISEN/M1/Projet M1/SOURCES_KOHLER/CSV/TYPE.csv", header=False, index=False)
    

def main():
    """
    Fonction principale qui fait appel aux fonctions précédentes
    """
        
    #Appel des fonctions
    moteur(excel_file_motor_path, column_name_motor)
    alternateur(excel_file_alternator_path)
    radiateur(excel_file_cooling_path, column_name_cooling)
    groupe(excel_file_encombrement_path, column_name_encombrement)
    groupe_element(excel_file_group_path, excel_file_cooling_path)
    capot(excel_file_encombrement_path)
    
    #chemin des fichiers CSV
    CSV_file_motor_path = "C:/Users/Cyann/Documents/ISEN/M1/Projet M1/SOURCES_KOHLER/CSV/MOTEUR_CSV.csv"
    CSV_file_alternator_path = "C:/Users/Cyann/Documents/ISEN/M1/Projet M1/SOURCES_KOHLER/CSV/ALTERNATEUR_CSV.csv"
    CSV_file_radiator_path = "C:/Users/Cyann/Documents/ISEN/M1/Projet M1/SOURCES_KOHLER/CSV/RADIATEUR_CSV.csv"
    CSV_file_capot_path = "C:/Users/Cyann/Documents/ISEN/M1/Projet M1/SOURCES_KOHLER/CSV/CAPOT_CSV.csv"
    
    #Appel des fonctions
    element_matiere(CSV_file_motor_path, CSV_file_alternator_path, CSV_file_radiator_path, CSV_file_capot_path)
    type_element()


# Programe principal

#Chemins des fichiers à traiter 
excel_file_motor_path = "C:/Users/Cyann/Documents/ISEN/M1/Projet M1/SOURCES_KOHLER/CAT_MOTEUR.xlsx"
excel_file_alternator_path = "C:/Users/Cyann/Documents/ISEN/M1/Projet M1/SOURCES_KOHLER/ALTERNATEUR.csv"
excel_file_group_path = "C:/Users/Cyann/Documents/ISEN/M1/Projet M1/SOURCES_KOHLER/CAT_GROUPE.xlsx"
excel_file_cooling_path = "C:/Users/Cyann/Documents/ISEN/M1/Projet M1/SOURCES_KOHLER/REF_COOLING.xlsx"
excel_file_encombrement_path = "C:/Users/Cyann/Documents/ISEN/M1/Projet M1/SOURCES_KOHLER/TD_ENCOMBREMENT.xlsx"

#Colonnes que l'on souhaite traiter 
column_name_motor = ["IDENT\nItem","ID_STATUT\nStatus","MOT_GN_46_IN\nWet weight (kg)","MOT_H5_07_IN\nFuel consumption @ ESP Max Power (l/h)"]
#column_name_alternator = ["DESC_ALT Kohler Alternator description","ID_STATUT Status","ALT_GN_POIDS_MO Net Weight of the alternator with single bearing config (kg)"]
column_name_group = ["IDENT\nItem","ID_STATUT\nStatus","ID_MOT\nEngine type","ID_ALT\nAlternator type", "ID_CAPOT\nCanopy", "ID_ENC01\nDimensions 01", "ID_ENC02\nDimensions 02"]
column_name_cooling = ["REF_REFROID\nCooling part number","ID_STATUT\nStatus", "COOL_CAR_42_IN\nRadiator Weight (kg)"] #"ID_MOT\nEngine type", "ENC_TYPE\nEnclosure",
column_name_encombrement = ["IDENT\nItem","ID_STATUT\nStatus", "ENC_PDSBRT_IN\nGross weight (kg)"] #"ENC_TYPE\nEnclosure",

if __name__ == '__main__':
    main()

        
