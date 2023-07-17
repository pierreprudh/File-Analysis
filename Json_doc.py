#!/usr/bin/env python
# coding: utf-8

# <div style="background-image: linear-gradient(to right, #4b4cff , #00d4ff); text-align: center; padding: 50px;">
#     <h1 style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 24px; color: white; text-shadow: 2px 2px #4b4cff;">
#         JSON to Documentation
#     </h1>
# </div>

# <h2 style="color: #3366cc; font-family: Arial, sans-serif; font-size: 24px; font-weight: bold; text-transform: uppercase; text-align: center; border-bottom: 2px solid #3366cc; padding-bottom: 5px;">Bibliothèque</h2>

# In[1]:


import pandas as pd
import numpy as np 
import os

import random

import json 

from datetime import date
from datetime import datetime

from scipy import stats

from fpdf import FPDF
import json


# <h2 style="color: #3366cc; font-family: Arial, sans-serif; font-size: 24px; font-weight: bold; text-transform: uppercase; text-align: center; border-bottom: 2px solid #3366cc; padding-bottom: 5px;">doc_json</h2>

# In[2]:


def doc_json(file, sep = ',', ID = '') : 
    """
    Convertit un fichier CSV ou Excel en un format JSON contenant des informations statistiques sur les variables.

    Args:
        file (str): Chemin du fichier à convertir.
        sep (str, optional): Délimiteur utilisé pour les fichiers CSV. Par défaut, ','.
        ID (str or list, optional): Colonne(s) utilisée(s) comme identifiant(s). Par défaut, ''.

    Returns:
        dict: Un dictionnaire JSON contenant les informations statistiques sur les variables du fichier.

    Raises:
        TypeError: Si le fichier n'est pas au format CSV ou Excel.

    """
    #--------------------------------Ouverture du fichier----------------------------------------------------------------------
    if file.endswith('.csv'):
        data = pd.read_csv(file, sep = sep)
    elif file.endswith('.xlsx') or file.endswith('.xls'):
        data = pd.read_excel(file)
    else:
        return 'Erreur : Le fichier doit être au format CSV ou Excel.'
    # Récupération du fichier
    #--------------------------------------------------------------------------------------------------------------------------
    
    
    #--------------------------------Nom et valeurs principales du fichier-----------------------------------------------------
    file_name = file.split('/')[-1]
    current_date = date.today().strftime('%d-%m-%Y')
    # Nom et date d'ouverture du fichier
    
    types = data.dtypes
    columns = data.columns
    missing_values = data.isna().sum()
    len_data = len(data)
    missing_values_total = data.isna().sum().sum()
    # type des colonnes, colonnes et valeurs manquantes
    #--------------------------------------------------------------------------------------------------------------------------
    
    
    #--------------------------------Mise en ébauche des différentes colonnes--------------------------------------------------
    variables = {}
    for column, data_type in zip(columns, types):
        variables[column] = {
            'type': str(data_type),
            'missing_values': int(missing_values[column])
        }
    # valeurs classiques (type et nombre de valeurs manquantes)
    #--------------------------------------------------------------------------------------------------------------------------
    
    
    #--------------------------------Colonnes de types 'float/int'-------------------------------------------------------------
    float_int_columns = types[(types == 'float64') | (types == 'int64')].index.tolist()
    # On récupère les colonnes de types int et float

    for column in float_int_columns:
        column_data = data[column]

        if len(data[column] < 3000) : 
            sample_size = len(data[column])
        else : 
            sample_size = 3000
        # Taille du sample pour le test statistique
        
        sample = random.sample(column_data.tolist(), sample_size)
        shapiro_stat, shapiro_pvalue = stats.shapiro(sample)
        # Test de shapiro
        
        q1 = np.percentile(column_data, 25)
        q3 = np.percentile(column_data, 75)
        iqr = q3 - q1
        lower_bound = q1 - 1.5 * iqr
        upper_bound = q3 + 1.5 * iqr
        outliers = column_data[(column_data < lower_bound) | (column_data > upper_bound)]
        # Calcul des valeurs aberrantes
        
        if missing_values_total == 0 : 
            msg_val_per_tot = 0
        else : 
            msg_val_per_tot = (int(missing_values[column]/missing_values_total))*100
        # On fait attention car si la valeur du nombre de valeurs manquantes total est null 
        
        variables[column] = {
            'type': str(types[column]),
            'statistics': {
                'mean': np.mean(column_data),
                'std': np.std(column_data),
                'min': np.min(column_data),
                'max': np.max(column_data),
                'q1' : q1, 
                'q3' : q3,
                'shapiro_statistic': shapiro_stat,
                'shapiro_pvalue': shapiro_pvalue,
                'outliers': len(outliers.tolist())
            },
            'missing values': int(missing_values[column]),  
            'missing values percent column':(int(missing_values[column])/len_data) *100,
            'missing values percent total': msg_val_per_tot
        }
        # On remplit le dictionnaire avec les valeurs 
    #--------------------------------------------------------------------------------------------------------------------------
    
    
    #--------------------------------Colonne de type date----------------------------------------------------------------------
    date_columns = [col for col, dtype in types.items() if dtype == 'datetime64[ns]']
    # On récupère les colonnes de type date
    
    for column in date_columns:
        column_data = data[column]
        date_formats = set() 

        for date_value in column_data:
            date_str = date_value.strftime('%d-%m-%Y') 
            # Conversion en chaine de caractère
            
            try:
                datetime.strptime(date_str, '%d-%m-%Y')
                date_formats.add('European')
            except ValueError:
                pass
            # Format européen

            try:
                datetime.strptime(date_str, '%m-%d-%Y')
                date_formats.add('American')
            except ValueError:
                pass
            # Format américain 
            
            if 'European' in date_formats:
                selected_format = 'European'
            elif 'American' in date_formats:
                selected_format = 'American'
            else:
                selected_format = 'Unknown'
            # Récupération du format

        variables[column] = {
            'type': str(types[column]),
            'extremum': {
                'earliest_date': np.min(column_data).strftime('%d-%m-%Y'),
                'latest_date': np.max(column_data).strftime('%d-%m-%Y')
            },
            'missing_values': int(missing_values[column]),
            'date_format': selected_format
        }
    # On remplit avec les valeurs obtenues
    #--------------------------------------------------------------------------------------------------------------------------
    
    
    #--------------------------------Colonne de type 'année'-------------------------------------------------------------------
    year_columns = [col for col, dtype in types.items() if dtype == 'int64' and (col.lower() == 'annee' or col.lower() == 'année' or col.lower() == 'year')] 
    # On cherche si on a des colonnes de 'type' année
    
    for column in year_columns:
        column_data = data[column]
        variables[column] = {
            'type': str(types[column]),
            'extremum': {
                'earliest_year': np.min(column_data),
                'latest_year': np.max(column_data)
            },
            'missing_values': int(missing_values[column])
        }
    # On remplit avec les valeurs extrêmes
    #--------------------------------------------------------------------------------------------------------------------------
    
    
    #--------------------------------Colonne de type 'object'------------------------------------------------------------------
    object_columns = date_columns = [col for col, dtype in types.items() if dtype == 'object']
    # On récupère les colonnes de type object 
    
    for column in object_columns: 
        column_data = data[column]
        
        unique = column_data.unique()
        value_counts = column_data.value_counts(normalize=True)
        top_values = value_counts[value_counts > 0.1]
        top10 = top_values.index.tolist()
        top10_percentages = (top_values * 100).round(2).tolist()
        # On calcule les pourcentages des valeurs les plus présentes dans chaque colonne de type 'object' en affichant que celle de plus de 10%
        
        variables[column] = {
            'type': str(types[column]),
            'missing_values': int(missing_values[column]), 
            '>10% appearance': dict(zip(top10, top10_percentages))
        }
    #--------------------------------------------------------------------------------------------------------------------------
    
    
    #--------------------------------Colonne de type ID------------------------------------------------------------------------
    try:
    # On teste si ID n'est pas vide
        if ID:
            id_colonne = data[ID]

            unique_values = len(id_colonne.value_counts())
            highest_app = max(id_colonne.value_counts())
            nbr_hgh_app = sum(id_colonne.value_counts() == highest_app)
            # valeur unique et nombre d'apparitions

            unique = id_colonne.unique()
            value_counts = id_colonne.value_counts(normalize=True)
            top_values = value_counts[value_counts > 0.1]
            top10 = top_values.index.tolist()
            top10_percentages = (top_values * 100).round(2).tolist()
            # Valeur avec le plus de présence et son pourcentage

            variables[ID] = {
                'type': str(types[ID]),
                'missing_values': int(missing_values[ID]),
                'unique value': unique_values,
                'highest appearance': highest_app,
                'nbr of highest app': nbr_hgh_app,
                '>10% appearance': dict(zip(top10, top10_percentages))
                }
    except KeyError:
        pass
    #--------------------------------------------------------------------------------------------------------------------------
    
    
    #--------------------------------Mise en forme des résultats---------------------------------------------------------------
    json_resultat = {
        'file_properties': {
            'file_name': file_name,
            'date': current_date,
            'Nombre de lignes' : len_data,
            'Missing values' : missing_values_total
        },
        'variables': variables
    }
    #--------------------------------------------------------------------------------------------------------------------------

    return json_resultat


# In[3]:


doc_json('./data/fr-en-ips_colleges.csv',';','UAI')


# <h2 style="color: #3366cc; font-family: Arial, sans-serif; font-size: 24px; font-weight: bold; text-transform: uppercase; text-align: center; border-bottom: 2px solid #3366cc; padding-bottom: 5px;">Mise en page</h2>

# In[4]:


def format_json(json_data):
    """
    Formate le contenu JSON pour une meilleure lisibilité.

    Arguments :
    - json_data : dict : Les données JSON à formater.

    Retourne :
    - str : Le contenu JSON formaté.
    """
    indent = 4 
    sorted_json = json.dumps(json_data, indent=indent, sort_keys=True, ensure_ascii=False, default=convert_to_builtin_type)
    formatted_json = ""

    level = 0
    for char in sorted_json:
        if char == '{':
            formatted_json += char + "\n" + " " * (indent * (level + 1))
            level += 1
        elif char == '}':
            formatted_json += "\n" + " " * (indent * (level - 1)) + char
            level -= 1
        elif char == ',':
            formatted_json += char + "\n" + " " * (indent * level)
        else:
            formatted_json += char

    return formatted_json

def convert_to_builtin_type(obj):
    """
    Convertit les types d'objets personnalisés en types JSON sérialisables.

    Arguments :
    - obj : object : L'objet à convertir.

    Retourne :
    - object : L'objet converti.

    Raises :
    - TypeError : Si l'objet n'est pas sérialisable en JSON.
    """
    if isinstance(obj, np.int64):
        return int(obj)
    raise TypeError("Object of type '{}' is not JSON serializable".format(type(obj).__name__))

result = doc_json('./data/fr-en-ips_colleges.csv', ';')
formatted_json = format_json(result)
#print(formatted_json)


# In[5]:


formatted_json = format_json(doc_json('./data/Extraction EGIDE - Nombre de formations dispensées - Année scolaire 2020-2021.xlsx'))
#print(formatted_json)


# In[6]:


formatted_json = format_json(doc_json('M:/str-dne-sn-3/Ingenierie/PNE/Tableau/DonneesEGIDE/Agents_EGIDES.csv'))
#print(formatted_json)


# <h2 style="color: #3366cc; font-family: Arial, sans-serif; font-size: 24px; font-weight: bold; text-transform: uppercase; text-align: center; border-bottom: 2px solid #3366cc; padding-bottom: 5px;">Mise en page PDF</h2>

# In[7]:


class PDF(FPDF):
    """
    Classe PDF étendue pour la création de fichiers PDF personnalisés.

    Cette classe hérite de la classe FPDF du module fpdf.

    Méthodes spéciales :
    - header() : Définit le contenu de l'en-tête du document PDF.
    - footer() : Définit le contenu du pied de page du document PDF.

    Méthodes personnalisées :
    - chapter_title(title) : Ajoute un titre de chapitre au document PDF.
    - chapter_body(content) : Ajoute le contenu d'un chapitre au document PDF.
    """
    
    def header(self):
        """
        Définit le contenu de l'en-tête du document PDF.
        """
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, 'Documentation de CSV/XLSX', align='C')
        self.ln(15)

    def footer(self):
        """
        Définit le contenu du pied de page du document PDF.
        """
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

    def chapter_title(self, title):
        """
        Ajoute un titre de chapitre au document PDF.

        Arguments :
        - title : str : Le titre du chapitre.
        """
        self.set_font('Arial', 'B', 20)
        self.cell(0, 10, title, ln=True)
        self.ln(5)

    def chapter_body(self, content):
        """
        Ajoute le contenu d'un chapitre au document PDF.

        Arguments :
        - content : str : Le contenu du chapitre.
        """
        self.set_font('Arial', '', 12)
        self.multi_cell(0, 10, content)
        self.ln(10)


# In[8]:


def create_pdf_from_json(json_data, output_file, file_name):
    """
    Crée un fichier PDF à partir des données JSON fournies.

    Arguments :
    - json_data : str : Les données JSON au format texte.
    - output_file : str : Le nom du fichier PDF de sortie.
    - file_name : str : Le nom du fichier d'origine.

    Cette fonction crée un fichier PDF à partir des données JSON fournies.
    
    Returns:
        .pdf: un pdf contenant les informations du json.

    """
    data = json_data    
    pdf = PDF()
    pdf.add_page()
    # Initialisation
    
    pdf.set_font('Arial', 'B', 24)  
    pdf.cell(0, 10, file_name, ln=True, align="C", border=1, fill=False)
    # Titre
    
    pdf.ln(10)
    # Saut de ligne
    
    pdf.chapter_title("File_properties")
    pdf.chapter_body(str(data['file_properties']))
    # Titre de chapitre

    pdf.ln(10)  
    # Saut de ligne
    
    pdf.chapter_title("Variables")
    # Titre de chapitre
    
    for key, value in data['variables'].items():
        pdf.chapter_body(f"{key}: {value}")
        pdf.ln(5)  # Saut de ligne entre les valeurs
    # Ajouter chaque paire clé-valeur avec un saut de ligne
    
    pdf.output(output_file)


# In[9]:


def main() : 
    file = './data/test.xlsx'
    file_name = file.split('/')[-1]
    result = doc_json(file, ';', 'UAI')
    output_file = f"./documentation/{file_name}.pdf"
    create_pdf_from_json(result, output_file, file_name)

main()

