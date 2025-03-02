#!/usr/bin/env python
# coding: utf-8

from pandas import (
    read_excel,
    DataFrame, 
    ExcelWriter,
    to_numeric,
    Series,
    cut
)
from pathlib import Path
import numpy as np
import scipy.stats as stats
from IPython.core.interactiveshell import InteractiveShell
from openpyxl.styles import Border, Side
import math

# Enable automatic display of all expressions in the cell
InteractiveShell.ast_node_interactivity = "all"

# Config
project_root = Path.cwd()
df_path = project_root / "analysededonnees.xlsx"
df = read_excel(df_path)

# Remove non-participants and fill NaN
remove_non_participant = df[df["Q0-Acceptez-vous de participer à cette enquête?"] != "Non, je ne souhaite pas participer"]
remove_np_fillna = remove_non_participant.fillna("n/a")

# Variables to analyze
variables = {
    "Q2-Veuillez indiquer votre Age.": "Age",
    "Q5-Combien de personnes vivent-ils sous le même toit?": "Personnes_toit",
    "Q10-Montant total du dernier crédit contracté": "Montant_credit",
    "Q11-Temps écoulé entre la demande et l'approbation du crédit.": "Temps_approbation"
}

# Function to generate bins using Sturges' formula
def generate_sturges_bins(data):
    """
    Generate bins using Sturges' formula.
    
    Parameters:
    ----------
    data : list or pandas.Series
        The data to bin
        
    Returns:
    -------
    list
        Bin edges
    """
    # Calculate number of bins using Sturges' formula
    n = len(data)
    if n == 0:
        return []
    
    k = math.ceil(1 + 3.322 * math.log10(n))
    
    # Calculate bin edges
    min_val = min(data)
    max_val = max(data)
    bin_width = (max_val - min_val) / k
    bins = [min_val + i * bin_width for i in range(k + 1)]
    
    return bins

# Function to generate labels for bins
def generate_bin_labels(bins):
    """
    Generate labels for bins in the format [a, b[.
    
    Parameters:
    ----------
    bins : list
        Bin edges
        
    Returns:
    -------
    list
        Bin labels
    """
    labels = [f"[{bins[i]:.2f},{bins[i+1]:.2f}[" for i in range(len(bins) - 1)]
    labels[-1] = f"[{bins[-2]:.2f},{bins[-1]:.2f}]"  # Close the last interval
    return labels

# Transform variables using Sturges' formula
def transform_variable(data, column_name):
    """
    Transform a variable into classes using Sturges' formula.
    
    Parameters:
    ----------
    data : pandas.DataFrame
        The dataframe containing the data
    column_name : str
        The name of the column to transform
        
    Returns:
    -------
    pandas.Series
        Transformed series with classes
    """
    # Extract data and convert to numeric, forcing non-numeric to NaN
    numeric_values = to_numeric(data[column_name], errors='coerce')
    
    # Remove NaN values (including 'n/a')
    cleaned_data = numeric_values.dropna().tolist()
    
    # Skip transformation if no valid numeric values
    if len(cleaned_data) == 0:
        return Series(["n/a"] * len(data), index=data.index)
    
    # Generate bins and labels
    bins = generate_sturges_bins(cleaned_data)
    labels = generate_bin_labels(bins)
    
    # Transform data into classes
    transformed = cut(numeric_values, bins=bins, labels=labels, include_lowest=True)
    
    # Add 'n/a' as a valid category
    transformed = transformed.cat.add_categories("n/a")
    
    # Restore 'n/a' values
    transformed[numeric_values.isna()] = "n/a"
    
    return transformed

# Apply transformation to all variables
transformed_data = remove_np_fillna.copy()
for col, name in variables.items():
    transformed_data[f"{name}_classe"] = transform_variable(transformed_data, col)

# Save results to Excel
with ExcelWriter("Analyse_univariee_avec_classes.xlsx", engine='openpyxl') as writer:
    transformed_data.to_excel(writer, sheet_name="Donnees_Transformees", index=False)
    
    # Add bin information for each variable
    for col, name in variables.items():
        numeric_values = to_numeric(transformed_data[col], errors='coerce')
        cleaned_data = numeric_values.dropna().tolist()
        
        if len(cleaned_data) > 0:
            bins = generate_sturges_bins(cleaned_data)
            labels = generate_bin_labels(bins)
            
            bin_info = DataFrame({
                "Bin Start": bins[:-1],
                "Bin End": bins[1:],
                "Label": labels
            })
            
            bin_info.to_excel(writer, sheet_name=f"{name}_Bins", index=False)

print("Transformation terminée. Les résultats ont été sauvegardés dans 'Analyse_univariee_avec_classes.xlsx'.")
