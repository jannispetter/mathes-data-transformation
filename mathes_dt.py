import pandas as pd
from pathlib import Path

# --------------------------------------------------
# PFAD
# --------------------------------------------------
BASE_PATH = Path(r"C:\Users\janni\OneDrive\Desktop\MathesData\Daten und Zahlen\Daten")

# --------------------------------------------------
# DATEIEN LADEN
# --------------------------------------------------
files = sorted(BASE_PATH.glob("*.xlsx"))

print(f"\nGefundene Dateien: {len(files)}")

dfs = []

for file in files:
    print(f"Lade Datei: {file.name}")
    
    df = pd.read_excel(file)
    
    df["source_file"] = file.name
    dfs.append(df)

# --------------------------------------------------
# ZUSAMMENFÜHREN
# --------------------------------------------------
combined_df = pd.concat(dfs, ignore_index=True)

print("\nGesamtdaten vor Dedup:", combined_df.shape)

# --------------------------------------------------
# DATUM
# --------------------------------------------------
combined_df["Rechnungsdatum"] = pd.to_datetime(
    combined_df["Rechnungsdatum"], errors="coerce"
)

# --------------------------------------------------
# NUMERISCHE SPALTEN FIXEN
# --------------------------------------------------
numeric_cols = ["Menge", "Nettopreis", "Rabatt", "Sonderrabatt1", "Skonto"]

for col in numeric_cols:
    combined_df[col] = pd.to_numeric(combined_df[col], errors="coerce")

# NaN entfernen → wichtig für Power BI
combined_df[numeric_cols] = combined_df[numeric_cols].fillna(0)

# --------------------------------------------------
# DUPLIKATE ENTFERNEN
# --------------------------------------------------
combined_df = combined_df.sort_values("source_file")

combined_df = combined_df.drop_duplicates(
    subset=["Auftrag", "Position"],
    keep="last"
)

print("Gesamtdaten nach Dedup:", combined_df.shape)

# --------------------------------------------------
# UMSATZ
# --------------------------------------------------
combined_df["umsatz"] = (
    combined_df["Menge"]
    * combined_df["Nettopreis"]
    * (1 - combined_df["Rabatt"])
    * (1 - combined_df["Sonderrabatt1"])
    * (1 - combined_df["Skonto"])
)

combined_df["umsatz"] = combined_df["umsatz"].fillna(0)

# --------------------------------------------------
# JAHR
# --------------------------------------------------
combined_df["Jahr"] = combined_df["Rechnungsdatum"].dt.year

# --------------------------------------------------
# LAND MAPPING (ROBUST)
# --------------------------------------------------
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

# sauberer String
combined_df["LLand_clean"] = (
    combined_df["LLand"]
    .astype(str)
    .str.strip()
    .map(land_mapping)
)

# --------------------------------------------------
# LAND LOGIK (WICHTIG!)
# --------------------------------------------------
combined_df["Land"] = combined_df["LLand_clean"]

# Fallback: PLZ Logik
mask_missing = combined_df["Land"].isna()
plz = combined_df["LPostleitzahl"].astype(str)

# Niederlande (Buchstaben)
combined_df.loc[
    mask_missing & plz.str.contains("[A-Za-z]", regex=True),
    "Land"
] = "Niederlande"

# Deutschland
combined_df.loc[
    mask_missing & plz.str.match(r"^(50|51|52|40)\d{3}$"),
    "Land"
] = "Deutschland"

# Belgien
combined_df.loc[
    mask_missing & plz.str.match(r"^(1|2|3|4|5)\d{3}$"),
    "Land"
] = "Belgien"

# Spanien
combined_df.loc[
    mask_missing & plz.str.match(r"^(07|08)\d{3}$"),
    "Land"
] = "Spanien"

# Final fallback
combined_df["Land"] = combined_df["Land"].fillna("Unbekannt")

print("\n--- Länderverteilung ---")
print(combined_df["Land"].value_counts().head(10))

# --------------------------------------------------
# OUTPUT
# --------------------------------------------------
output_path = BASE_PATH / "auftraege_final.csv"

combined_df.to_csv(output_path, index=False)

print("\nDatei gespeichert unter:", output_path)