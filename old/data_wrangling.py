#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd


# In[2]:


# Charger les données depuis le fichier Excel
data = pd.read_excel('./data_set.xlsx',sheet_name="data")


# In[3]:


# Déterminer les limites des classes
n_classes = 6

min_val = data['Chi_Aff_Men_Ac'].min()
max_val = data['Chi_Aff_Men_Ac'].max()

min_val_pma = data['Px_Min_Av'].min()
max_val_pma = data['Px_Min_Av'].max()

min_val_pmac = data['Px_Min_Ac'].min()
max_val_pmac = data['Px_Min_Ac'].max()

min_val_pmav = data['Px_Max_Av'].min()
max_val_pmav = data['Px_Max_Av'].max()

min_val_pmax = data['Px_Max_Ac'].min()
max_val_pmax = data['Px_Max_Ac'].max()

min_val_men = data['Chi_Aff_Men'].min()
max_val_men = data['Chi_Aff_Men'].max()

min_val_emp_av = data['Nb_Emp_Av'].min()
max_val_emp_av = data['Nb_Emp_Av'].max()

min_val_emp_ac = data['Nb_Emp_Ac'].min()
max_val_emp_ac = data['Nb_Emp_Ac'].max()


class_width_Chi_Aff_Men_Ac = (max_val - min_val) / n_classes

class_width_Chi_Aff_Men = (max_val_men - min_val_men) / n_classes

class_width_Px_Min_Av = (max_val_pma - min_val_pma) / n_classes

class_width_Px_Min_Ac = (max_val_pmac - min_val_pmac) / n_classes

class_width_Px_Max_Av = (max_val_pmav - min_val_pmav) / n_classes

class_width_Px_Max_Ac = (max_val_pmax - min_val_pmax) / n_classes

class_width_Nb_Emp_Ac = (max_val_emp_ac - min_val_emp_ac) / n_classes

class_width_Nb_Emp_Av = (max_val_emp_av - min_val_emp_av) / n_classes


# In[4]:


# Définir les intervalles
bins_Chi_Aff_Men_Ac = [min_val + i * class_width_Chi_Aff_Men_Ac for i in range(n_classes + 1)]
bins_Chi_Aff_Men = [min_val_men + i * class_width_Chi_Aff_Men for i in range(n_classes + 1)]
bins_Px_Min_Av = [min_val_pma + i * class_width_Px_Min_Av for i in range(n_classes + 1)]
bins_Px_Min_Ac = [min_val_pmac + i * class_width_Px_Min_Ac for i in range(n_classes + 1)]
bins_Px_Max_Av = [min_val_pmav + i * class_width_Px_Max_Av for i in range(n_classes + 1)]
bins_Px_Max_Ac = [min_val_pmax + i * class_width_Px_Max_Ac for i in range(n_classes + 1)]
bins_Nb_Emp_Av = [min_val_emp_av + i * class_width_Nb_Emp_Av for i in range(n_classes + 1)]
bins_Nb_Emp_Ac = [min_val_emp_ac + i * class_width_Nb_Emp_Ac for i in range(n_classes + 1)]


# In[5]:


# Générer les labels au format souhaité
labels_Chi_Aff_Men_Ac = [f"[{int(bins_Chi_Aff_Men_Ac[i])},{int(bins_Chi_Aff_Men_Ac[i+1])}[" for i in range(len(bins_Chi_Aff_Men_Ac) - 1)]
labels_Chi_Aff_Men_Ac[-1] = f"[{int(bins_Chi_Aff_Men_Ac[-2])},{int(bins_Chi_Aff_Men_Ac[-1])}]"  # Dernière classe fermée à droite

labels_Chi_Aff_Men = [f"[{int(bins_Chi_Aff_Men[i])},{int(bins_Chi_Aff_Men[i+1])}[" for i in range(len(bins_Chi_Aff_Men) - 1)]
labels_Chi_Aff_Men[-1] = f"[{int(bins_Chi_Aff_Men[-2])},{int(bins_Chi_Aff_Men[-1])}]"  # Dernière classe fermée à droite

labels_Px_Min_Av = [f"[{int(bins_Px_Min_Av[i])},{int(bins_Px_Min_Av[i+1])}[" for i in range(len(bins_Px_Min_Av) - 1)]
labels_Px_Min_Av[-1] = f"[{int(bins_Px_Min_Av[-2])},{int(bins_Px_Min_Av[-1])}]"  # Dernière classe fermée à droite

labels_Px_Min_Ac = [f"[{int(bins_Px_Min_Ac[i])},{int(bins_Px_Min_Ac[i+1])}[" for i in range(len(bins_Px_Min_Ac) - 1)]
labels_Px_Min_Ac[-1] = f"[{int(bins_Px_Min_Ac[-2])},{int(bins_Px_Min_Ac[-1])}]"  # Dernière classe fermée à droite 

labels_Px_Max_Av = [f"[{int(bins_Px_Max_Av[i])},{int(bins_Px_Max_Av[i+1])}[" for i in range(len(bins_Px_Max_Av) - 1)]
labels_Px_Max_Av[-1] = f"[{int(bins_Px_Max_Av[-2])},{int(bins_Px_Max_Av[-1])}]"  # Dernière classe fermée à droite

labels_Px_Max_Ac = [f"[{int(bins_Px_Max_Ac[i])},{int(bins_Px_Max_Ac[i+1])}[" for i in range(len(bins_Px_Max_Ac) - 1)]
labels_Px_Max_Ac[-1] = f"[{int(bins_Px_Max_Ac[-2])},{int(bins_Px_Max_Ac[-1])}]"  # Dernière classe fermée à droite

labels_Nb_Emp_Av = [f"[{int(bins_Nb_Emp_Av[i])},{int(bins_Nb_Emp_Av[i+1])}[" for i in range(len(bins_Nb_Emp_Av) - 1)]
labels_Nb_Emp_Av[-1] = f"[{int(bins_Nb_Emp_Av[-2])},{int(bins_Nb_Emp_Av[-1])}]"  # Dernière classe fermée à droite

labels_Nb_Emp_Ac = [f"[{int(bins_Nb_Emp_Ac[i])},{int(bins_Nb_Emp_Ac[i+1])}[" for i in range(len(bins_Nb_Emp_Ac) - 1)]
labels_Nb_Emp_Ac[-1] = f"[{int(bins_Nb_Emp_Ac[-2])},{int(bins_Nb_Emp_Ac[-1])}]"  # Dernière classe fermée à droite


# In[7]:


data['Classe_chiffre_aff_men_ac'] = pd.cut(data['Chi_Aff_Men_Ac'], bins=bins_Chi_Aff_Men_Ac, labels=labels_Chi_Aff_Men_Ac, include_lowest=True)
data['Classe_Chi_Aff_Men'] = pd.cut(data['Chi_Aff_Men'], bins=bins_Chi_Aff_Men, labels=labels_Chi_Aff_Men, include_lowest=True)
data['Classe_Px_Min_Av'] = pd.cut(data['Px_Min_Av'], bins=bins_Px_Min_Av, labels=labels_Px_Min_Av, include_lowest=True)
data['Classe_Px_Min_Ac'] = pd.cut(data['Px_Min_Ac'], bins=bins_Px_Min_Ac, labels=labels_Px_Min_Ac, include_lowest=True)
data['Classe_Px_Max_Av'] = pd.cut(data['Px_Max_Av'], bins=bins_Px_Max_Av, labels=labels_Px_Max_Av, include_lowest=True)
data['Classe_Px_Max_Ac'] = pd.cut(data['Px_Max_Ac'], bins=bins_Px_Max_Ac, labels=labels_Px_Max_Ac, include_lowest=True)
data['Classe_Nb_Emp_Av'] = pd.cut(data['Nb_Emp_Av'], bins=bins_Nb_Emp_Av, labels=labels_Nb_Emp_Av, include_lowest=True)
data['Classe_Nb_Emp_Ac'] = pd.cut(data['Nb_Emp_Ac'], bins=bins_Nb_Emp_Ac, labels=labels_Nb_Emp_Ac, include_lowest=True)


data


# In[8]:


data.to_excel("./EI.xlsx",index=False)

