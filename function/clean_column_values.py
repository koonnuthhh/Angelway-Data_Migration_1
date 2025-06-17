import pandas as pd

def clean_column_names(df: pd.DataFrame, column: str) -> pd.DataFrame:
    # Hardcoded correction mapping
    clean_column_values = {
    "MITSUBISHI": ["Mitsubihi", "Mitsubshi", "Misubishi", "Mitsubashi", "Mitzubishi"],
    "MG": ["Mg", "mg", "MGG", "MG.", "M-G"],
    "TOYOTA": ["Toyata", "Totota", "Toyoya", "Toyoya", "T0yota"],
    "ISUZU": ["Iszuzu", "Isuzu", "Isuzu", "Iszsu", "Izuzu"],
    "DFSK": ["DFK", "DFSKK", "D-FSK", "DSFK", "DFFSK"],
    "CHEVROLET": ["Chevorlet", "Chevy", "Chevrolat", "Chev", "Cheverlot"],
    "FORD": ["Forrd", "Fod", "Foed", "F0rd", "Frd"],
    "HONDA": ["Hondai", "Hando", "Hondo", "Hondar", "Honnda", "ฮอนด้า"],
    "D-MAX": ["Dmax", "D Max", "D_MAZ", "DMAX", "D-Maax"],
    "HINO": ["Hinoo", "Hiino", "Hinno", "HINO.", "Hino-"],
    "GPX": ["GPXX", "GP-PX", "G-PX", "GPPX", "GXP"],
    "NISSAN": ["Nisan", "Nissin", "Nissam", "Nissaan", "Nissn"],
    "KUBOTA": ["Kubotaa", "Cubota", "Kobota", "Kuboto", "Kuboata"],
    "MAZDA": ["Mazdaa", "Masda", "Mazd", "Mazta", "Mazdah"],
    "SUBARU": ["Suberu", "Subaro", "SubaRu", "Subauru", "Sbaru"],
    "YAMAHA": ["Yamha", "Yamama", "Yamhaa", "Yamah", "Yamhaa", "YAMAYA"],
    "VESPA": ["Vesbaa", "Vesspa", "Vspa", "Veespa", "Vesap"],
}


    # Create a reverse lookup for faster correction
    reverse_lookup = {}
    for correct, wrongs in clean_column_values.items():
        for w in wrongs:
            reverse_lookup[w.lower()] = correct  # Case insensitive

    # Apply corrections
    def correct_value(value):
        if isinstance(value, str) and value.lower() in reverse_lookup:
            return reverse_lookup[value.lower()]
        return value

    df[column] = df[column].apply(correct_value)

    return df
