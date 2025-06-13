import pandas as pd

import pandas as pd

def split_value_by_keyword(df, source_col, steps, leftover_col='leftover'):
    """
    Extracts keywords step-by-step from a source column in a DataFrame,
    creates new columns for matches, and assigns the leftover after all
    extractions to a separate leftover column, while keeping the original
    source column unchanged.

    Parameters:
        df: pandas.DataFrame
        source_col: str - column containing input text (unchanged)
        steps: list of tuples (keyword_list, match_column)
        leftover_col: str - column to store leftover text after all extractions

    Returns:
        pandas.DataFrame - modified DataFrame with extracted keyword columns
                          and leftover column, original source_col untouched
    """

    def extract_keywords(text, keyword_set):
        text = str(text)
        matches = []
        i = 0
        while i < len(text):
            found = False
            for word in sorted(keyword_set, key=len, reverse=True):
                if text[i:i+len(word)] == word:
                    matches.append(word)
                    i += len(word)
                    found = True
                    break
            if not found:
                i += 1
        matched_str = ' '.join(matches)
        leftover = text
        for m in matches:
            leftover = leftover.replace(m, '', 1)
        return matched_str, leftover

    # Initialize a temporary column with the original text
    # Convert to string to handle potential non-string types gracefully
    temp_leftover = df[source_col].astype(str).copy()

    # --- Preprocessing Step ---
    # 1. Remove '_P' and '_H'
    temp_leftover = temp_leftover.str.replace('_P', '', regex=False)
    temp_leftover = temp_leftover.str.replace('_H', '', regex=False)


    for keywords, match_col in steps:
        keyword_set = set(keywords)
        results = temp_leftover.apply(lambda x: pd.Series(extract_keywords(x, keyword_set)))
        df[match_col] = results[0]  # matched keywords
        temp_leftover = results[1]  # update leftover for next step

    # After all steps, assign leftover to leftover_col
    df[leftover_col] = temp_leftover

    return df


def copy_column_and_delete_old(df, old_col, new_col):
    """
    Copies data from one column (old_col) to another column (new_col) in a DataFrame,
    then deletes the old column.

    Parameters:
        df: pandas.DataFrame
        old_col: str - name of the column to copy from and delete
        new_col: str - name of the column to copy data to

    Returns:
        pandas.DataFrame with updated columns
    """
    if old_col not in df.columns:
        raise KeyError(f"Old column '{old_col}' not found in DataFrame.")

    # Copy the old column data to the new column
    df[new_col] = df[old_col]

    # Drop the old column
    df = df.drop(columns=[old_col])

    return df