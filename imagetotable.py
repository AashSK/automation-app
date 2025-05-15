# %%
import numpy as np
import pandas as pd
from sklearn.cluster import KMeans

def convert_ocrdata_to_table(df):
    # Datenbereinigung
    df = df.copy()
    df = df.dropna(subset=['text'])
    df['text'] = df['text'].astype(str).str.strip()
    df = df[df['text'] != ''] 
    
    # Text ab '|' teilen und in separate Zeilen erweitern, um Gruppierungsprobleme zu vermeiden
    df = df.assign(text=df['text'].str.split('|')).explode('text').reset_index(drop=True)
    df['text'] = df['text'].str.strip()
    df = df[df['text'] != '']  # Entfernen von Leerzeichen
    
    # Sortieren des DataFrames
    df_sorted = df.sort_values(['block_num', 'par_num', 'line_num', 'word_num', 'top'])
    
    # Gruppierung von Wörten und Sätzen in Zeilen
    def group_words_in_row(group):
        words = group.sort_values('left')
        word_groups = []
        current_words = []
        current_left = None
        last_right = None
        
        # Textzeilen durchlaufen
        for _, word in words.iterrows():
            current_text = word['text']
            current_position = word['left']
            
            # Prüfen Ob Text als Erstes vorkommnt
            if current_left is None:
                current_words = [current_text]
                current_left = current_position
                last_right = current_position + word['width']
                continue
            
            # Abstand zum letzten Wort berechnen
            gap = current_position - last_right
            
            has_continuation_marker = (
                current_words and 
                (current_words[-1].endswith(',') or 
                 current_words[-1].endswith('-') or
                 current_words[-1].endswith(':')))
            
            # Lückengröße ermitteln um die in Gruppen zuzuordnen
            gap_threshold = 50 if has_continuation_marker else 30
            
            if gap > gap_threshold:
                # Zuordnung der einzelnen Zeilen in Gruppen
                if current_words:
                    word_groups.append({
                        'text': ' '.join(current_words),
                        'left': current_left
                    })
                current_words = [current_text]
                current_left = current_position
            else:
                current_words.append(current_text)
            
            last_right = current_position + word['width']
        
        # Letzte Gruppe
        if current_words:
            word_groups.append({
                'text': ' '.join(current_words),
                'left': current_left
            })
            #print(word_groups)
        
        return word_groups
    
    grouped = df_sorted.groupby(['block_num', 'par_num', 'line_num'])
    row_data = []
    for (block, par, line), group in grouped:
        text_groups = group_words_in_row(group)
        min_top = group['top'].min()
        for text_group in text_groups:
            row_data.append({
                'block_num': block,
                'par_num': par,
                'line_num': line,
                'text': text_group['text'],
                'left': text_group['left'],
                'top': min_top
            })
    
    result_df = pd.DataFrame(row_data)
    if result_df.empty:
        return pd.DataFrame()
    
    # Verwendung von K-Means Algorithmus um die texte in Spalten zu verteilen
    num_columns = min(4, len(result_df['left'].unique())) 
    kmeans = KMeans(n_clusters=num_columns, random_state=42, n_init=15)
    result_df['column'] = kmeans.fit_predict(result_df[['left']])
    
    # Zuordnung von Clusterbezeichnungen zu Spaltennummern auf der Grundlage der sortierten links Positionen
    column_order = sorted(result_df.groupby('column')['left'].mean().items(), key=lambda x: x[1])
    column_map = {col[0]: idx for idx, col in enumerate(column_order)}
    result_df['column'] = result_df['column'].map(column_map)
    
    # Anordnen von Texten in strukturierten Zeilen
    unique_tops = sorted(result_df['top'].unique())
    threshold = 10
    row_groups = []
    current_group = [unique_tops[0]]
    
    for top in unique_tops[1:]:
        if top - current_group[-1] <= threshold:
            current_group.append(top)
        else:
            row_groups.append(current_group)
            current_group = [top]
    if current_group:
        row_groups.append(current_group)
    
    top_to_row = {}
    for row_idx, group in enumerate(row_groups):
        for top in group:
            top_to_row[top] = row_idx
    
    num_rows = len(row_groups)
    final_table = pd.DataFrame(np.empty((num_rows, num_columns), dtype='object'))
    
    for _, row in result_df.iterrows():
        row_idx = top_to_row[row['top']]
        col_idx = row['column']
        final_table.at[row_idx, col_idx] = row['text']
    
    return final_table.dropna(how='all', axis=0).reset_index(drop=True)


