import pandas as pd
from pathlib import Path

# Dein Projektpfad
BASE_PATH = Path(r"C:\Users\janni\OneDrive\Desktop\MathesData\Daten und Zahlen")

# Hier sollten deine Excel-Dateien liegen
DATA_PATH = BASE_PATH

files = [
    "Alle_Auftraege_2020-2022.xlsx",
    "Alle_Auftraege_2022-2024.xlsx",
    "Alle_Auftraege_2022-2026.xlsx"
]

dfs = []

for file in files:
    path = BASE_PATH / file
    print(f"Lade Datei: {file}")
    
    df = pd.read_excel(path)
    
    # WICHTIG: hier setzen
    df["source_file"] = file
    
    dfs.append(df)


# Zusammenführen
combined_df = pd.concat(dfs, ignore_index=True)

print("\nGesamtdaten:")
print(combined_df.shape)

# Erste Vorschau
print("\nSpalten:")
print(combined_df.columns)

print("\nBeispieldaten:")
print(combined_df.head())


print("\n--- Grundanalyse ---")

print("\nAnzahl eindeutiger Aufträge:")
print(combined_df["Auftrag"].nunique())

print("\nDurchschnittliche Positionen pro Auftrag:")
print(len(combined_df) / combined_df["Auftrag"].nunique())

print("\nFehlende Werte:")
print(combined_df.isnull().sum())

print("\nDatentypen:")
print(combined_df.dtypes)


#Überschneidungen Prüfen

print("\n--- Überschneidung prüfen ---")

# nach Source File gruppieren
for file in files:
    df_temp = combined_df[combined_df["source_file"] == file]
    print(f"\n{file}")
    print("Min Datum:", df_temp["Rechnungsdatum"].min())
    print("Max Datum:", df_temp["Rechnungsdatum"].max())
    print("Anzahl Aufträge:", df_temp["Auftrag"].nunique())

##Felder verstehen
#Umsatz sinnig?
combined_df["umsatz_test"] = combined_df["Menge"] * combined_df["Nettopreis"]

print("\nUmsatz Test:")
print(combined_df[["Menge", "Nettopreis", "umsatz_test"]].head(10))


#gutschriften verstehen
print("\n--- Gutschriften prüfen ---")

# Filter: Auftragsart enthält "Gutschrift"
gutschriften = combined_df[combined_df["Auftragsart"].str.contains("Gutschrift", na=False)]

print("Anzahl Gutschriften:", len(gutschriften))

print("\nBeispiele:")
print(gutschriften[["Auftragsart", "Menge", "Nettopreis"]].head(10))

#rabatt verstehen
print("\n--- Rabatt prüfen ---")

print("Einzigartige Rabattwerte:")
print(combined_df["Rabatt"].value_counts().head(20))

#skonto verstehen
print("\n--- Skonto prüfen ---")

print("Skonto Werte:")
print(combined_df["Skonto"].value_counts().head(20))

print("\nSkontono Werte:")
print(combined_df["Skontono"].value_counts())

#sonderrabatt prüfen
print("\n--- Sonderrabatt prüfen ---")

print(combined_df["Sonderrabatt1"].value_counts().head(20))


##umsatz berechnen
combined_df["umsatz"] = (
    combined_df["Menge"]
    * combined_df["Nettopreis"]
    * (1 - combined_df["Rabatt"])
    * (1 - combined_df["Sonderrabatt1"])
    * (1 - combined_df["Skonto"])
)

print("\nUmsatz mit Skonto:")
print(combined_df[[
    "Menge",
    "Nettopreis",
    "Rabatt",
    "Sonderrabatt1",
    "Skonto",
    "umsatz"
]].head(10))


# echte werte sehen
print("\n--- Nur echte Verkäufe ---")

echte_verkaeufe = combined_df[
    (combined_df["Nettopreis"] > 0) &
    (combined_df["Rabatt"] < 1)
]

print(echte_verkaeufe[[
    "Menge",
    "Nettopreis",
    "Rabatt",
    "Sonderrabatt1",
    "Skonto",
    "umsatz"
]].head(10))


#checke 0 Felder
print("\n--- 0€ Positionen prüfen ---")

zero_sales = combined_df[combined_df["umsatz"] == 0]

print("Anzahl 0€ Positionen:", len(zero_sales))
print("Anteil:", len(zero_sales) / len(combined_df))

#null werte für umsatz herausfiltern
combined_df["is_sale"] = combined_df["umsatz"] != 0


# berechnung Gesamtumsatz 
print("\n--- Gesamtumsatz ---")

gesamt_umsatz = combined_df["umsatz"].sum()

print("Gesamtumsatz:", round(gesamt_umsatz, 2))

#pro jahr
print("\n--- Umsatz pro Jahr ---")

combined_df["Jahr"] = combined_df["Rechnungsdatum"].dt.year

umsatz_jahr = combined_df.groupby("Jahr")["umsatz"].sum()

print(umsatz_jahr)

####länder codes
print(combined_df["LLand"].value_counts(dropna=False).head(20))
#länder immer gleich schreiben:
land_mapping = {
    "D": "Deutschland",
    "DE": "Deutschland",
    "DEUTSCHLAND": "Deutschland",
    "Deutschland": "Deutschland",
    
    "SPANIEN": "Spanien",
    "Spanien": "Spanien",
    
    "BELGIEN": "Belgien",
    "Belgien": "Belgien",
    
    "NIEDERLANDE": "Niederlande",
    "Niederlande": "Niederlande",
    
    "SCHWEIZ": "Schweiz",
    "Schweiz": "Schweiz",
    
    "LUXEMBURG": "Luxemburg",
    "Luxemburg": "Luxemburg",
    
    "Grossbritannien": "Großbritannien",
    
    "Österreich": "Österreich",
    
    "Portugal": "Portugal",
    
    "Frankreich": "Frankreich",
    
    "I": "Italien"
}

combined_df["LLand_clean"] = combined_df["LLand"].map(land_mapping)

print(combined_df["LLand_clean"].value_counts(dropna=False).head(20))

## plzs verstehen
print("\n--- PLZ + Land Analyse ---")

# Nur Zeilen mit vorhandenem Land
df_known = combined_df[combined_df["LLand_clean"].notna()]

# PLZ als String
df_known["PLZ_str"] = df_known["LPostleitzahl"].astype(str)

# Länge der PLZ
df_known["PLZ_len"] = df_known["PLZ_str"].str.len()

# Übersicht: PLZ-Länge vs Land
print("\nPLZ Länge vs Land:")
print(df_known.groupby(["PLZ_len", "LLand_clean"]).size().sort_values(ascending=False).head(20))

##bsp anschauen pro Land
print("\n--- Beispiel PLZ pro Land ---")

for land in ["Deutschland", "Belgien", "Schweiz", "Niederlande", "Spanien"]:
    print(f"\n{land}:")
    
    sample = combined_df[
        combined_df["LLand_clean"] == land
    ]["LPostleitzahl"].dropna().astype(str).unique()[:10]
    
    print(sample)

##modifizieren
# nl
mask_missing = combined_df["LLand_clean"].isna()

plz = combined_df["LPostleitzahl"].astype(str)

# NL: enthält Buchstaben
combined_df.loc[
    mask_missing & plz.str.contains("[A-Za-z]", regex=True),
    "LLand_clean"
] = "Niederlande"

print("\nNach NL-Ergänzung:")
print(combined_df["LLand_clean"].value_counts(dropna=False).head(20))

#de
mask_missing = combined_df["LLand_clean"].isna()
plz = combined_df["LPostleitzahl"].astype(str)

# typische deutsche PLZ-Bereiche (aus deinen Daten)
combined_df.loc[
    mask_missing & plz.str.match(r"^(50|51|52|40|52|52|52)\d{3}$"),
    "LLand_clean"
] = "Deutschland"

#be
mask_missing = combined_df["LLand_clean"].isna()
plz = combined_df["LPostleitzahl"].astype(str)

# Belgien: typische 4-stellige PLZ aus deinen Daten
combined_df.loc[
    mask_missing & plz.str.match(r"^(1|2|3|4|5)\d{3}$"),
    "LLand_clean"
] = "Belgien"

#überprüfung be
print("\n--- Verdächtige Belgien prüfen ---")

belgien = combined_df[combined_df["LLand_clean"] == "Belgien"]

print(belgien["LPostleitzahl"].astype(str).value_counts().head(20))

#einige wieder unebekannt machen

# bekannte belgische Präfixe (aus deinen Daten)
belgien_prefix = ["47", "48"]

mask_belgien = (
    combined_df["LLand_clean"] == "Belgien"
)

plz = combined_df["LPostleitzahl"].astype(str)

# alles behalten, was wirklich belgisch aussieht
combined_df.loc[
    mask_belgien & ~plz.str.startswith(tuple(belgien_prefix)),
    "LLand_clean"
] = None

#es
mask_missing = combined_df["LLand_clean"].isna()
plz = combined_df["LPostleitzahl"].astype(str)

combined_df.loc[
    mask_missing & plz.str.match(r"^(07|08)\d{3}$"),
    "LLand_clean"
] = "Spanien"

print(combined_df["LLand_clean"].value_counts(dropna=False).head(20))

combined_df["Land"] = combined_df["LLand_clean"].fillna("Unbekannt")

print(combined_df["Land"].value_counts().head(20))


#### als neue csv speichern
output_path = BASE_PATH / "auftraege_final.csv"

combined_df.to_csv(output_path, index=False)

print("Datei gespeichert unter:", output_path)