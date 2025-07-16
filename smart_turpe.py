# -*- coding: utf-8 -*-
"""
Created on Thu Oct  3 17:08:16 2024

@author: Hidekela
"""
import pandas as pd
import fitz
import re
from datetime import datetime, date
from pathlib import Path
import os

# Même si inutilisé dans le code ci-dessous, l'import openpyxl.cell._writer est nécessaire 
# au bon foncitonnement du fichier .exe créé à partir de ce programme
import openpyxl.cell._writer


def formater_concat_date(date_):
    """ Permet transformer un format date de jj.mm.aaaa à m/aaaa afin de créer
    une chaine de caratère que concatène plusieurs info
    
    :type date_: str
    :param date_: date de début sous la forme jj.mm.aaaa
    ---- RETURNS
    :type resultat: str
    :param resultat: date de début sous la forme m/aaaa
    """
      
    # Convertir la chaîne en objet datetime
    date_obj = datetime.strptime(date_, "%d/%m/%Y")
    
    # Formater la date pour obtenir "m/aaaa"
    resultat = date_obj.strftime("%m/%Y").lstrip("0")

    return resultat

def format_date(date_string):
    """ Permet transformer un format date de par exemple 13 août 2001 à 12/08/2001 afin de 
    formater les dates d'écriture dans le fichier Excel.
    
    :type date_string: str
    :param date_string: date d'écriture sous la forme jj mois aaaa
    ---- RETURNS
    :type date_: str
    :param date_: date d'écriture sous la forme jj/mm/aaaa
    """
    # Dictionnaire pour mapper les noms de mois en français à leurs numéros
    mois_francais = {
        'janvier': '01', 'février': '02', 'mars': '03', 'avril': '04',
        'mai': '05', 'juin': '06', 'juillet': '07', 'août': '08',
        'septembre': '09', 'octobre': '10', 'novembre': '11', 'décembre': '12'
    }

    # Séparation des éléments de la date
    jour, mois, annee = date_string.split()

    # Conversion du jour en format à deux chiffres
    jour = jour.zfill(2)

    # Obtention du numéro du mois
    mois_num = mois_francais[mois.lower()]
    
    # Formatage de la date
    date_ = f"{jour}/{mois_num}/{annee}"
    
    return date_



def create_df():
    """ Permet de créer le pandas dataframe avec les bonnes colonnes que 
    l'on va remplir pour ensuite transformer en fichier Excel une fois complété.
    
    ---- RETURNS
    :type df: pandas dataframe
    :param df: dataframe comportant les info des pdf
    """
    
    # Définition des noms de colonnes
    columns = [
        'CARDI', 'Mapping', 'Société et/ou', 'Etablissement', 'Type de compte', 'Journal', 'Type de pièce', 'Date écriture',
        'Code compte', 'Code tiers (valeur par défaut)', 'Profil de TVA', 'N° pièce', 'Libellé pièce (nom du client)',
        'Libellé écriture', "Date d'échéance", 'Sens', 'Montant EUR', 'Montant DEVISE(par défaut)', 'DEVISE (par défaut)',
        'Mode reg (par défaut)', 'PLAN 1', 'PLAN 2', 'Axe analytique', 'Champs supplémentaire ESM',
        'Champs supplémentaire ESM_', 'Date Début', 'Date Fin'
    ]
    
    # Création du DataFrame vide
    df = pd.DataFrame(columns=columns)
    return df

def extract_cardi(doc):
    """ Permet d'aller chercher le numéro de CARDI dans la facture.
    
    :type doc: Document object
    :param doc: fichier PDF ouvert
    ---- RETURNS
    :type cardi: str
    :param cardi: numéro de CARDI présent dans la facture
    """
    cardi = None

    for page in doc:
        text = page.get_text()
        
        # Recherche de "N° de contrat:" suivi de l'information
        match = re.search(r"Nº contrat\s*:?\s*(.*)", text, re.IGNORECASE)
        if match:
            cardi = match.group(1).strip()
            break  # On arrête dès qu'on trouve l'information
        else:
            print("Numéro de contrat non trouvé")
    return cardi

def extract_date_ecriture(doc):
    """ Permet d'aller chercher la date d'écriture dans la facture.
    
    :type doc: Document object
    :param doc: fichier PDF ouvert
    ---- RETURNS
    :type ecriture: str
    :param ecriture: date d'écriture présente dans la facture
    """
    # On extrait le texte de la 1e page
    page = doc[0]
    text = page.get_text()
    # Expression régulière pour trouver la date
    date_pattern = r"(\d{1,2}\s+(?:janvier|février|mars|avril|mai|juin|juillet|août|septembre|octobre|novembre|décembre)\s+\d{4})"
    
    # Recherche de la date dans le texte
    match = re.search(date_pattern, text)
    
    if match:
        ecriture = match.group(1)
    else:
        print("Date d'écriture non trouvée")
    return ecriture

def extract_nom_client(doc):
    """ Permet d'aller chercher le nom du client dans la facture.
    
    :type doc: Document object
    :param doc: fichier PDF ouvert
    ---- RETURNS
    :type nom_client: str
    :param nom_client: nom du client présent dans la facture
    """
    nom_client = None

    for page in doc:
        text = page.get_text()
        
        # Recherche de "Facture N°" suivi de l'information
        match = re.search(r"Facture N°\s*:?\s*(.*)", text, re.IGNORECASE)
        if match:
            full_match = match.group(1).strip()
            # Extraire uniquement le numéro de facture
            numero_facture = re.search(r"(\d+)", full_match)
            if numero_facture:
                nom_client = numero_facture.group(1)
            break  # On arrête dès qu'on trouve l'information
        else:
            print("Nom client non trouvé")
    return nom_client


def extract_echeance(doc):
    """ Permet d'aller chercher la date d'échéance dans la facture.
    
    :type doc: Document object
    :param doc: fichier PDF ouvert
    ---- RETURNS
    :type echeance: str
    :param echeance: date d'échéance présente dans la facture
    """
    echeance = None

    for page in doc:
        text = page.get_text()
        # Recherche de la date au format JJ/MM/AAAA
        match = re.search(r'le (\d{2}/\d{2}/\d{4})', text)
        if match:
            echeance = match.group(1)
            break  # On arrête dès qu'on trouve la date
        else:
            print("Échéance non trouvée")
    return echeance

def extract_montant(doc):
    """ Permet d'aller chercher le montant en EUR dans la facture.
    
    :type doc: Document object
    :param doc: fichier PDF ouvert
    ---- RETURNS
    :type montant: str
    :param montant: montant présent dans la facture
    """
    montant = None
    pattern = r"Sous-Total Accès au réseau H\.T\.\s+20,00\s+%\s+(\d{1,3}(?:\s?\d{3})*,\d{2})\s+€"



    
    for page in doc:
        text = page.get_text()
        # print(text)
        match = re.search(pattern, text)
        if match:
            montant_str = match.group(1)
            montant = montant_str.replace(',', '.')

            break  # On arrête dès qu'on trouve le montant
        else:
            pattern = r"Sous-Total Accès au réseau H\.T\.\s+20,00\s+%\s+(-\s{2})?(\d+,\d+)\s+€"
            match = re.search(pattern, text)
            if match:
                signe = '-' if match.group(1) else ''
                montant_str = match.group(2)
                montant = signe + montant_str.replace(',', '.')
                break  # On arrête dès qu'on trouve le montant
            else:
                print("Montant non trouvé")
    montant = montant.replace(" ", "")
    return montant

def extract_date_debut(doc):
    """ Permet d'aller chercher la date de début dans la facture.
    
    :type doc: Document object
    :param doc: fichier PDF ouvert
    ---- RETURNS
    :type debut: str
    :param debut: date de début présente dans la facture
    """
    debut = None

    for page in doc:
        text = page.get_text()
        # Recherche de la date au format JJ/MM/AAAA
        match = re.search(r'pour la période du (\d{2}\.\d{2}\.\d{4})', text)
        if match:
            debut = match.group(1)
            break  # On arrête dès qu'on trouve la date
    
    #Modification du format
    date_object = datetime.strptime(debut, "%d.%m.%Y")
    date_format = date_object.strftime("%d/%m/%Y")

    return date_format

def extract_date_fin(doc):
    """ Permet d'aller chercher la date de fin dans la facture.
    
    :type doc: Document object
    :param doc: fichier PDF ouvert
    ---- RETURNS
    :type fin: str
    :param fin: date de fin présente dans la facture
    """
    fin = None

    for page in doc:
        text = page.get_text()
        # Recherche de la date au format JJ/MM/AAAA
        match = re.search(r'au (\d{2}\.\d{2}\.\d{4})', text)
        if match:
            fin = match.group(1)
            break  # On arrête dès qu'on trouve la date
    
    #Modification du format
    date_object = datetime.strptime(fin, "%d.%m.%Y")
    date_format = date_object.strftime("%d/%m/%Y")
    
    return date_format


def extract_CRD(doc):
    """ Permet d'aller chercher le code CRD dans la facture.
    
    :type doc: Document object
    :param doc: fichier PDF ouvert
    ---- RETURNS
    :type crd: str
    :param crd: Code CRD présent dans la facture
    """
    crd = None

    for page in doc:
        text = page.get_text()
        # Recherche de la date au format JJ/MM/AAAA
        match = re.search(r'Mandat SEPA n°\s*:?\s*(.*)', text)
        if match:
            crd = match.group(1)
            break  # On arrête dès qu'on trouve la date
    return crd

def get_info_to_fill():
    """ Permet de récupérer toutes les info des factures grâces aux fonctions précédentes
    """
    
    # Recherche des éléments
    cardi = str(extract_cardi(doc))
    ecriture = format_date(extract_date_ecriture(doc))
    nom_client = extract_nom_client(doc)
    echeance = extract_echeance(doc)
    montant = extract_montant(doc)
    date_debut = extract_date_debut(doc)
    date_fin = extract_date_fin(doc)
    crd = extract_CRD(doc)

    
    return cardi, ecriture, nom_client, echeance, montant, date_debut, date_fin, crd

def df_to_excel(savingxlsx_path, df_trie): 
    """ Fonction qui crée le fichier excel en utilisant le dataframe
    
    :type savingxlsx: string
    :param doc: path d'enregistrement du fichier excel
    
    :type df_trie: pandas dataframe
    :param df_trie: df trié à partir duquel on crée le fichier excel
    ---- RETURNS TRUE
    """
    with pd.ExcelWriter(savingxlsx_path, engine='xlsxwriter') as writer:
        df_trie.to_excel(writer, sheet_name='Recap_factures_ENEDIS', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Recap_factures_ENEDIS']
        
        # Format pour les en-têtes
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        
        # Format particulier pour le nom client
        custom_format = workbook.add_format({'num_format': '0'})
        
        # Format pour la colonne "Société et/ou" (texte)
        text_format = workbook.add_format({'num_format': '@'})
        
        # Appliquer le format aux en-têtes
        for col_num, value in enumerate(df_trie.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Ajuster la largeur des colonnes et appliquer les formats spécifiques
        for i, col in enumerate(df_trie.columns):
            column_len = max(df_trie[col].astype(str).apply(len).max(), len(col)) 
            worksheet.set_column(i, i, column_len + 2)
            
            # Appliquer le format personnalisé à la colonne "Libellé pièce (nom du client)"
            if col == 'Libellé pièce (nom du client)':  
                worksheet.set_column(i, i, column_len + 2, custom_format)
            
            # Appliquer le format texte à la colonne "Société et/ou"
            elif col == 'Société et/ou':
                worksheet.set_column(i, i, column_len + 2, text_format)
    return True


if __name__ == "__main__":
    
    # Obtenir le répertoire de travail actuel
    chemin_courant = os.getcwd()

    
    #Ouverture du fichier excel 'Gestion SPV' de l'exploitation
    # excel_pcard_path = "I:/ACTIFS EN EXPLOITATION/0- Exploitation/0 - Gestion centrale PV/Gestion SPV.xlsx"
    excel_pcard_path = chemin_courant + "/Gestion SPV.xlsx"
    
    #Envoie un message d'erreur si le fichier Gestion SPV.xlsx n'est pas retrouvé (notamment si le nom n'est pas le bon)
    try:
        excel_pcard = pd.read_excel(excel_pcard_path, sheet_name='PCARD.I')
    except Exception:
        print("Le fichier Gestion SPV.xlsx est introuvable, vérifiez le nom de votre fichier\n")

    excel_pcard['N° CARD I'] = excel_pcard['N° CARD I'].astype(str)


    fichier_facture = chemin_courant
    chemin_dossier = Path(fichier_facture)
    # Création du pandas dataframe qu'on va remplir avec les info puis transformer en excel
    df = create_df()
    i = 0 #num facture dans dossier
   
    for pdf_file in chemin_dossier.glob("*.pdf"):
        with fitz.open(pdf_file) as doc:  # Ouvrir le fichier PDF            
            print(f"Traitement de {pdf_file.name}...")
            cardi, ecriture, nom_client, echeance, montant, date_debut, date_fin, crd = get_info_to_fill() 
            try: 
                montant = float(montant)
            except Exception as e:
                print(f"Le montant de la facture {pdf_file.name} a été mal lu: {str(e)}\n")

                
            
            # Si une facture est en double dans le dossier, ne pas la traiter 2 fois.
            if nom_client in df['Libellé pièce (nom du client)'].values:
                print(f"Le libellé pièce {nom_client} existe déjà: PDF en double. Passage au fichier suivant.")
                continue  # Passer au fichier PDF suivant
            
            # Si le montant de la facture est nul
            if montant == 0:
                print(f"La facture {nom_client} n'a pas été prise en compte car le montant à régler est nul. Passage au fichier suivant.")
                continue 
            
            #Pour N° pièce
            i+=1
            # Ajout des éléments dans le df
            # Code tiers
            df.loc[0+(4*i), 'Code tiers (valeur par défaut)'] = 'ERDF'
            
            #Lorsque toute la colonne vaut la même valeur
            for a in range(4):
                # CARDI
                df.loc[a+(4*i), 'CARDI'] = cardi    
                
                # Mapping
                try:
                    df.loc[a+(4*i), 'Mapping'] = (excel_pcard.loc[excel_pcard['N° CARD I'] == cardi, 'Centrale'].values[0]).upper()
                except IndexError:
                    print(f'Pas de correspondance trouvée pour le cardi {cardi} dans le fichier 202405 - Gestion SPV.xlsx\n')
                except Exception as e:
                    print(f"Une erreur inattendue s'est produite pour la facture {pdf_file.name} lors du remplissage de la colonne 'Mapping': {str(e)}\n")
                    
                #Société et/ou
                try:
                    df.loc[a+(4*i), 'Société et/ou'] = (str(excel_pcard.loc[excel_pcard['N° CARD I'] == cardi, 'Code SPV'].values[0])).upper()
                except Exception as e:
                    print(f"Une erreur inattendue s'est produite pour la facture {pdf_file.name} lors du remplissage de la colonne 'Société et/ou': {str(e)}\n")
                
                #Etablissement
                try:
                    df.loc[a+(4*i), 'Etablissement'] = "SIEGE-" + df.loc[a+4*i, "Société et/ou"]
                except Exception as e:
                    print(f"Une erreur inattendue s'est produite pour la facture {pdf_file.name} lors du remplissage de la colonne 'Etablissement': {str(e)}\n")
                
                # Journal
                df.loc[a+(4*i), 'Journal'] = "ACH"
                
                # Type de pièce
                df.loc[a+(4*i), 'Type de pièce'] = "FF"
                
                # Date d'écriture
                df.loc[a+(4*i), 'Date écriture'] = ecriture
                
                # N° pièce
                df.loc[a+(4*i), 'N° pièce'] = i
                
                # Libellé pièce
                df.loc[a+(4*i), 'Libellé pièce (nom du client)'] = nom_client
                
                # Libellé écriture
                try:
                    if montant < 0:
                        df.loc[a+(4*i), "Libellé écriture"] = df.loc[0+(4*i), 'Code tiers (valeur par défaut)'] + "-" + formater_concat_date(date_debut) + "-" + df.loc[0+4*i, "Mapping"]
                        
                    else:
                        df.loc[a+(4*i), "Libellé écriture"] = df.loc[0+(4*i), 'Code tiers (valeur par défaut)'] + "-" + formater_concat_date(date_debut) + "-" + df.loc[0+4*i, "Mapping"] + "-" + crd
                except Exception as e:
                    print(f"Une erreur inattendue s'est produite pour la facture {pdf_file.name} lors du remplissage de la colonne 'Libellé écriture': {str(e)}\n")
              
            
            try:
                # Type de compte    
                df.loc[0+(4*i), 'Type de compte'] = "X"
                df.loc[1+(4*i), 'Type de compte'] = "G"
                df.loc[2+(4*i), 'Type de compte'] = "A"
                df.loc[3+(4*i), 'Type de compte'] = "G"
                
                # Code compte
                df.loc[0+(4*i), 'Code compte'] = 40110000
                df.loc[1+(4*i), 'Code compte'] = 60410000
                df.loc[2+(4*i), 'Code compte'] = 60410000
                df.loc[3+(4*i), 'Code compte'] = 44561200
                
                
                # Profil de TVA
                df.loc[1+(4*i), 'Profil de TVA'] = 'TVA Déd. 20% (débits)'
                
                # Date d'échéance
                df.loc[0+(4*i), "Date d'échéance"] = echeance
                
                # Sens
                df.loc[0+(4*i), "Sens"] = "C"
                df.loc[1+(4*i), "Sens"] = "D"
                df.loc[2+(4*i), "Sens"] = "D"
                df.loc[3+(4*i), "Sens"] = "D"
                
                
                #PLAN 1
                df.loc[2+(4*i), "PLAN 1"] = df.loc[0+(4*i), "Mapping"]
                
                #PLAN 2
                df.loc[2+(4*i), "PLAN 2"] = "EBITDA_OPEX_REC"
                
                # Date début
                df.loc[1+(4*i), 'Date Début'] = date_debut
                
                #  Date fin
                df.loc[1+(4*i), 'Date Fin'] = date_fin
            except Exception as e:
                print(f"Une erreur inattendue s'est produite pour la facture {pdf_file.name}: {str(e)}\n")
      
        
            try:
                # Montant EUR
                df.loc[2+(4*i), 'Montant EUR'] = montant
                df.loc[3+(4*i), 'Montant EUR'] = round(float(df.loc[2+(4*i), 'Montant EUR'])*0.2, 2)
                df.loc[1+(4*i), 'Montant EUR'] = df.loc[2+(4*i), 'Montant EUR'] 
                df.loc[0+(4*i), 'Montant EUR'] = float(df.loc[1+(4*i), 'Montant EUR']) + float(df.loc[3+(4*i), 'Montant EUR'])
            except Exception as e:
                print(f"Une erreur inattendue s'est produite pour la facture {pdf_file.name} lors de la lecture du montant: {str(e)}\n")
      
        
      
    #On trie le dataframe pour que les lignes du fichiers excel soient triées 
    #par ordre croissant de la société puis du numéro de pièce
    df_trie = df.sort_values(['Société et/ou', 'N° pièce'], ascending=[True, True])
    df_trie['Société et/ou'] = df_trie['Société et/ou'].astype('string')
    
    #Nom fichiers avec date à laquelle il est généré

    current_date = date.today()
    formatted_date = current_date.strftime("%d-%m-%Y")
    name = "Facture ENEDIS - Traitement du " + str(formatted_date)
        
    #Création du fichier excel
    savingxlsx_path = chemin_courant + "\\" + name + ".xlsx"
    df_to_excel(savingxlsx_path, df_trie)
    
    #Enregistrement du fichier au format csv délimité par un point-virgule              
    savingcsv_path = chemin_courant + "\\" + name + ".csv"
    
    #Sélectionner seulement à partir de la colonne Société et/ou
    df_csv = df_trie.loc[:, 'Société et/ou':]
    #On change les points en virgule pour que ce soit bien reconnu par le logiciel d'intégration 
    df_csv['Montant EUR'] = df_csv['Montant EUR'].astype(str).str.replace('.', ',')
    
    # Pour qu'Excel ne mettre pas l'écriture de 100000157963 en écriture scientifique lors de l'importation du csv
    df_csv['Libellé pièce (nom du client)'] = '="' + df_csv['Libellé pièce (nom du client)'].astype(str) + '"'

    # Enregistrement du DataFrame au format CSV
    df_csv.to_csv(savingcsv_path, encoding='utf-8-sig', sep=';', date_format='%Y/%m/%d', index=False)


 


