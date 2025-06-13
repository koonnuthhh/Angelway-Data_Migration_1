import pandas as pd

def replace_text_in_column(df, column_name, old_word, new_word, case_sensitive=True):
    """
    Replaces occurrences of 'old_word' with 'new_word' in a specific DataFrame column.

    Parameters:
        df (pd.DataFrame): The input DataFrame.
        column_name (str): The name of the column to perform the replacement in.
        old_word (str): The word or substring to find.
        new_word (str): The word or substring to replace with.
        case_sensitive (bool): If True, the replacement is case-sensitive.
                               If False, it's case-insensitive. Default is True.

    Returns:
        pd.DataFrame: A new DataFrame with the replacements made.
                      The original DataFrame is not modified.
    """
    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' not found in the DataFrame.")
    
    # Use .str.replace() for string replacement
    # regex=False for literal string replacement, not regex patterns
    # case=case_sensitive for case sensitivity control
    df[column_name] = df[column_name].astype(str).str.replace(
        old_word, new_word, case=case_sensitive, regex=False
    )
    print(f"\nDataFrame after replacement ('{old_word}' -> '{new_word}') in column '{column_name}':")
    return df