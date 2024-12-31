import pandas as pd
from codicefiscale import codicefiscale

# Nome del file Excel da elaborare
input_file = 'tesserati.xlsx'  # Sostituisci con il nome del tuo file
output_file = 'tesserati_con_codici_fiscali.xlsx'

# Funzione per calcolare il codice fiscale
def calcola_cf(row):
    try:
        return codicefiscale.encode({
            'name': row['Nome'],
            'surname': row['Cognome'],
            'gender': row['Genere'],
            'birthdate': row['Data di nascita'],
            'birthplace': row['Luogo di nascita']
        })
    except Exception as e:
        print(f"Errore per {row['Nome']} {row['Cognome']}: {e}")
        return None

# Caricamento del file Excel
try:
    df = pd.read_excel(input_file)
except FileNotFoundError:
    print(f"Errore: il file '{input_file}' non è stato trovato.")
    exit()

# Verifica delle colonne necessarie
required_columns = ['Cognome', 'Nome', 'Data di nascita', 'Luogo di nascita', 'Genere']
if not all(col in df.columns for col in required_columns):
    print(f"Errore: il file deve contenere le colonne: {', '.join(required_columns)}")
    exit()

# Calcolo dei codici fiscali
df['Codice Fiscale'] = df.apply(calcola_cf, axis=1)

# Salvataggio del nuovo file Excel
df.to_excel(output_file, index=False)

print(f"File generato con successo: {output_file}")
