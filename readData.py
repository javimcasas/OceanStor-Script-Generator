import pandas as pd

def readFile(nombre_archivo):
    try:
        # Load the Excel file into a DataFrame
        df = pd.read_excel(nombre_archivo)
        
        # Convert the DataFrame into a list of lists, each row becomes a list
        data = df.values.tolist()  # Convert DataFrame rows into lists of values
        return data
    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        return None
