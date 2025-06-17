import pandas as pd

def clean_column_names(df: pd.DataFrame, column: str, clean_column_values) -> pd.DataFrame:


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
