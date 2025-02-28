#!/usr/bin/env python
# coding: utf-8

# # Statistical Analysis

# In[1]:


from pandas import (
    read_excel,
    DataFrame, 
    ExcelWriter
)
from pathlib import Path
import numpy as np
import scipy.stats as stats
from IPython.core.interactiveshell import InteractiveShell
from openpyxl.styles import Border, Side


# Enable automatic display of all expressions in the cell
InteractiveShell.ast_node_interactivity = "all"


# In[2]:


# config
project_root = Path.cwd()
project_root


# In[3]:


df_path = project_root / "analysededonnees.xlsx"
df_path


# In[4]:


df = read_excel(df_path)

df.columns


# ## First Treatment
# 
# ---
# 
# [x] remove non participant <br/>
# [x] fill NAN 

# In[5]:


df['Q0-Acceptez-vous de participer à cette enquête?'].value_counts()


# In[6]:


remove_non_participant = df[df["Q0-Acceptez-vous de participer à cette enquête?"]!="Non, je ne souhaite pas participer"]
remove_non_participant


# In[7]:


remove_np_fillna = remove_non_participant.fillna("n/a")
remove_np_fillna


# ## univariate qUALITATIVE

# In[8]:


class QuantitativeData:
    def __init__(self, data):
        # Convertir "n/a" en NaN (Not a Number) pour les ignorer dans les calculs
        self.data = [np.nan if str(x).lower() in ["n/a", "na", ""] else x for x in data]

    def mean(self):
        """Calcule la moyenne en ignorant les NaN."""
        return np.nanmean(self.data)

    def median(self):
        """Calcule la médiane en ignorant les NaN."""
        return np.nanmedian(self.data)

    def mode(self):
        """Calcule le mode en ignorant les NaN."""
        try:
            return stats.mode([x for x in self.data if not np.isnan(x)]).mode
        except:
            return "Aucun mode trouvé"

    def std_dev(self):
        """Calcule l'écart type en ignorant les NaN."""
        return np.nanstd(self.data)

    def range(self):
        """Calcule l'étendue en ignorant les NaN."""
        cleaned_data = [x for x in self.data if not np.isnan(x)]
        return np.max(cleaned_data) - np.min(cleaned_data)

    def variance(self):
        """Calcule la variance en ignorant les NaN."""
        return np.nanvar(self.data)


# In[9]:


age = remove_np_fillna["Q2-Veuillez indiquer votre Age."].to_list()
personne_vivant_meme_toit = remove_np_fillna["Q5-Combien de personnes vivent-ils sous le même toit?"].to_list()
montant_total_credit = remove_np_fillna["Q10-Montant total du dernier crédit contracté"].to_list()
temps_ecoule_app_credit = remove_np_fillna["Q11-Temps écoulé entre la demande et l\'approbation du crédit."].to_list()


# In[10]:


age_analysis = QuantitativeData(age)
personne_vivant_meme_toit_analysis = QuantitativeData(personne_vivant_meme_toit)
montant_total_credit_analysis = QuantitativeData(montant_total_credit)
temps_ecoule_app_credit_analysis = QuantitativeData(temps_ecoule_app_credit)


# In[11]:


resultats = {
    "age": {
        "Moyenne": [age_analysis.mean()],
        "Médiane": [age_analysis.median()],
        "Mode": [age_analysis.mode()],
        "Écart type": [age_analysis.std_dev()],
        "Étendue": [age_analysis.range()],
        "Variance": [age_analysis.variance()]
    },
    "personne_vivant_meme_toit": {
        "Moyenne": [personne_vivant_meme_toit_analysis.mean()],
        "Médiane": [personne_vivant_meme_toit_analysis.median()],
        "Mode": [personne_vivant_meme_toit_analysis.mode()],
        "Écart type": [personne_vivant_meme_toit_analysis.std_dev()],
        "Étendue": [personne_vivant_meme_toit_analysis.range()],
        "Variance": [personne_vivant_meme_toit_analysis.variance()]
    },
    "montant_total_credit": {
        "Moyenne": [montant_total_credit_analysis.mean()],
        "Médiane": [montant_total_credit_analysis.median()],
        "Mode": [montant_total_credit_analysis.mode()],
        "Écart type": [montant_total_credit_analysis.std_dev()],
        "Étendue": [montant_total_credit_analysis.range()],
        "Variance": [montant_total_credit_analysis.variance()]
    },
    "temps_ecoule_app_credit": {
        "Moyenne": [temps_ecoule_app_credit_analysis.mean()],
        "Médiane": [temps_ecoule_app_credit_analysis.median()],
        "Mode": [temps_ecoule_app_credit_analysis.mode()],
        "Écart type": [temps_ecoule_app_credit_analysis.std_dev()],
        "Étendue": [temps_ecoule_app_credit_analysis.range()],
        "Variance": [temps_ecoule_app_credit_analysis.variance()]
    }
}


# In[12]:


age_univariate_results = DataFrame(resultats["age"])
personne_vivant_meme_toit_univariate_results = DataFrame(resultats["personne_vivant_meme_toit"])
montant_total_credit_univariate_results = DataFrame(resultats["montant_total_credit"])
temps_ecoule_app_credit_univariate_results = DataFrame(resultats["temps_ecoule_app_credit"])


# In[13]:


# Création du fichier Excel avec plusieurs feuilles
with ExcelWriter("Analyse univariee des variables quantitatives.xlsx", engine='openpyxl') as writer:
    # Écrire chaque DataFrame dans une feuille distincte avec des noms raccourcis
    age_univariate_results.to_excel(writer, sheet_name="age", index=False)
    personne_vivant_meme_toit_univariate_results.to_excel(writer, sheet_name="personnes_meme_toit", index=False)
    montant_total_credit_univariate_results.to_excel(writer, sheet_name="montant_credit", index=False)
    temps_ecoule_app_credit_univariate_results.to_excel(writer, sheet_name="temps_approbation", index=False)

    # Ajouter des bordures aux tableaux
    workbook = writer.book
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Appliquer les bordures à chaque feuille
    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
            for cell in row:
                cell.border = thin_border


# ## Turn data quality variable with formula
# 
# ---
# 
# (
#     Q2,
#     Q5,
#     q10,
#     q11
# )

# In[ ]:





# In[ ]:




