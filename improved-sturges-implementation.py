#!/usr/bin/env python
# coding: utf-8

class SturgesTransformer:
    """
    A class to transform variables using Sturges' formula.
    
    The Sturges formula determines the optimal number of bins:
    k = 1 + 3.322 * log10(n)
    
    where:
    - k is the number of bins
    - n is the sample size
    """
    
    def __init__(self):
        """Initialize the transformer."""
        import numpy as np
        self.bin_edges = None
        self.bin_midpoints = None
        self.n_bins = None
        self.frequencies = None
    
    def _calculate_bins(self, data):
        """
        Calculate the number of bins and bin edges using Sturges' formula.
        
        Parameters:
        ----------
        data : array-like
            Numeric data to calculate bins for
            
        Returns:
        -------
        tuple
            (n_bins, bin_edges, bin_midpoints, frequencies)
        """
        import numpy as np
        import math
        
        # Calculate number of bins using Sturges' formula
        n = len(data)
        n_bins = math.ceil(1 + 3.322 * math.log10(n))
        
        # Create histogram bins
        hist, bin_edges = np.histogram(data, bins=n_bins)
        
        # Calculate bin midpoints
        bin_midpoints = (bin_edges[1:] + bin_edges[:-1]) / 2
        
        return n_bins, bin_edges, bin_midpoints, hist
    
    def transform(self, series):
        """
        Transform a series using Sturges' method.
        
        Parameters:
        ----------
        series : pandas.Series
            The series to transform
            
        Returns:
        -------
        pandas.Series
            Transformed series with n/a values preserved
        """
        import pandas as pd
        import numpy as np
        
        # Create a copy to avoid modifying the original
        result = series.copy()
        
        # Identify 'n/a' values to preserve them
        na_mask = series.astype(str).isin(['n/a', 'na', ''])
        
        # Convert to numeric, forcing non-numeric to NaN (but keep track of original n/a values)
        numeric_values = pd.to_numeric(series, errors='coerce')
        
        # Get valid values for transformation (not NaN)
        valid_mask = ~numeric_values.isna()
        valid_values = numeric_values[valid_mask].values
        
        # Skip transformation if no valid values
        if len(valid_values) == 0:
            return series
        
        # Calculate bins
        self.n_bins, self.bin_edges, self.bin_midpoints, self.frequencies = self._calculate_bins(valid_values)
        
        # Transform valid values to bin midpoints
        transformed_values = np.zeros_like(valid_values, dtype=float)
        
        for i, value in enumerate(valid_values):
            # Find the bin this value belongs to
            bin_index = np.digitize(value, self.bin_edges) - 1
            
            # Handle edge case where value equals the maximum
            if bin_index >= len(self.bin_midpoints):
                bin_index = len(self.bin_midpoints) - 1
            
            # Assign the bin midpoint
            transformed_values[i] = self.bin_midpoints[bin_index]
        
        # Create a new series with transformed values
        result = pd.Series(index=series.index, dtype=object)  # Use object dtype to store both numbers and strings
        
        # Assign transformed values
        result[valid_mask] = transformed_values
        
        # Restore original 'n/a' values
        result[na_mask] = 'n/a'
        
        return result
    
    def get_bin_info(self):
        """
        Get information about the bins.
        
        Returns:
        -------
        dict
            Dictionary with bin information
        """
        if self.bin_edges is None:
            return None
        
        return {
            'n_bins': self.n_bins,
            'bin_edges': self.bin_edges,
            'bin_midpoints': self.bin_midpoints,
            'frequencies': self.frequencies
        }


class DataTransformer:
    """
    Class to handle the transformation of multiple variables in a dataset.
    """
    
    def __init__(self, df):
        """
        Initialize with a dataframe.
        
        Parameters:
        ----------
        df : pandas.DataFrame
            The dataframe to transform
        """
        self.df = df
        self.transformers = {}
        self.results = {}
    
    def transform_variable(self, column_name, output_column_name=None):
        """
        Transform a single variable using Sturges method.
        
        Parameters:
        ----------
        column_name : str
            The name of the column to transform
        output_column_name : str, optional
            The name for the transformed column. If None, will use column_name + "_sturges"
            
        Returns:
        -------
        pandas.Series
            Transformed series
        """
        # Create a new transformer for this column
        transformer = SturgesTransformer()
        
        # Apply transformation
        transformed = transformer.transform(self.df[column_name])
        
        # Store the transformer for later reference
        self.transformers[column_name] = transformer
        
        # Determine output column name
        if output_column_name is None:
            output_column_name = f"{column_name}_sturges"
            
        # Store the result
        self.results[column_name] = {
            'transformed': transformed, 
            'output_name': output_column_name
        }
        
        return transformed
    
    def transform_all(self, columns_map):
        """
        Transform multiple variables with specified output names.
        
        Parameters:
        ----------
        columns_map : dict
            Dictionary mapping original column names to output column names
            
        Returns:
        -------
        pandas.DataFrame
            DataFrame with all columns from original df and added transformed columns
        """
        import pandas as pd
        
        # Create a copy of the original dataframe
        result_df = self.df.copy()
        
        # Transform each variable and add to the result
        for original_col, output_col in columns_map.items():
            transformed = self.transform_variable(original_col, output_col)
            result_df[output_col] = transformed
        
        return result_df
    
    def get_bin_info(self, column_name):
        """
        Get bin information for a transformed variable.
        
        Parameters:
        ----------
        column_name : str
            The name of the column
            
        Returns:
        -------
        dict or None
            Dictionary with bin information or None if not transformed
        """
        if column_name in self.transformers:
            return self.transformers[column_name].get_bin_info()
        return None
    
    def save_to_excel(self, filename="sturges_transformed_variables.xlsx"):
        """
        Save transformed data and bin information to Excel.
        
        Parameters:
        ----------
        filename : str
            The name of the Excel file to create
        """
        import pandas as pd
        
        if not self.results:
            print("No transformations to save.")
            return
        
        # Create a DataFrame with original data and transformations
        result_df = self.df.copy()
        
        # Add transformed columns
        for col, info in self.results.items():
            result_df[info['output_name']] = info['transformed']
        
        # Write to Excel with multiple sheets
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            # Main sheet with all data
            result_df.to_excel(writer, sheet_name="Complete_Dataset", index=False)
            
            # Create sheets with bin information
            for col in self.results.keys():
                bin_info = self.get_bin_info(col)
                if bin_info:
                    # Create a simplified name for the sheet
                    short_name = col.split('-')[0] if '-' in col else col
                    
                    bin_df = pd.DataFrame({
                        'Bin Start': bin_info['bin_edges'][:-1],
                        'Bin End': bin_info['bin_edges'][1:],
                        'Bin Midpoint': bin_info['bin_midpoints'],
                        'Frequency': bin_info['frequencies']
                    })
                    
                    bin_df.to_excel(writer, sheet_name=f"{short_name}_Bins", index=False)
            
            # Create a summary sheet with information about transformations
            summary_data = []
            for col, info in self.results.items():
                bin_info = self.get_bin_info(col)
                if bin_info:
                    summary_data.append({
                        'Original Column': col,
                        'Transformed Column': info['output_name'],
                        'Number of Bins': bin_info['n_bins']
                    })
            
            if summary_data:
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name="Transformation_Summary", index=False)
        
        print(f"Transformations saved to {filename}")


def main():
    """
    Main function to transform variables in the dataset using remove_np_fillna.
    """
    import pandas as pd
    from pathlib import Path
    
    # Load data
    project_root = Path.cwd()
    df_path = project_root / "analysededonnees.xlsx"
    df = pd.read_excel(df_path)
    
    # Remove non-participants and fill NAs as in the original code
    remove_non_participant = df[df["Q0-Acceptez-vous de participer à cette enquête?"]!="Non, je ne souhaite pas participer"]
    remove_np_fillna = remove_non_participant.fillna("n/a")
    
    # Define columns to transform with simplified output names
    columns_to_transform = {
        "Q2-Veuillez indiquer votre Age.": "Age_cont",
        "Q5-Combien de personnes vivent-ils sous le même toit?": "Personnes_toit_cont",
        "Q10-Montant total du dernier crédit contracté": "Montant_credit_cont",
        "Q11-Temps écoulé entre la demande et l'approbation du crédit.": "Temps_approbation_cont"
    }
    
    # Create transformer using remove_np_fillna
    transformer = DataTransformer(remove_np_fillna)
    
    # Transform all variables
    result_df = transformer.transform_all(columns_to_transform)
    
    # Save to Excel
    transformer.save_to_excel("variables_continues_sturges.xlsx")
    
    return transformer, result_df


if __name__ == "__main__":
    main()
