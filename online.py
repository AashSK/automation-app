# %%
import csv
import os
import chardet
import pandas as pd
from pandas import DataFrame
from IPython.display import display, Markdown

def detect_encoding(file_path):
    with open(file_path, 'rb') as file:
        raw_data = file.read()
        result = chardet.detect(raw_data)
        encoding = result['encoding']
        euc = 'EUC'
        if euc.lower() in encoding.lower():
            encoding = 'UTF-8'
        confidence = result['confidence']
    return encoding, confidence

def detect_delimiter(file_path,encoding):
    with open(file_path, 'r', newline='', encoding=encoding) as file:
        # # Lesen von Stichprobezeilen, um das Trennzeichen zu erraten
        sample = file.read(4096)
        #print(sample)
        delimiter = csv.Sniffer().sniff(sample).delimiter
        return delimiter

def contains_keywords(df, keywords):
    keyword_match_count = sum(keyword == col for col in df.columns for keyword in keywords)
    return keyword_match_count >= len(keywords) / 2

def clean_zoom_data(df:DataFrame):
    index_of_name = df[df.apply(lambda row: row.astype(str).str.contains(r'Name \(Originalname\)|Benutzer-E-Mail:|Beitrittszeit|Beendigungszeit|Gast|Dauer \(Minuten\)', na=False).any(), axis=1)].index[0]
    #print(index_of_name)
    df.columns = df.iloc[index_of_name-1]                                                                 # Zeile Name (Originalname) als neue Tabellenzeile gesetzt
    df = df.iloc[index_of_name:].reset_index(drop=True)                                                   # Zeilen Oberhalb Name (Originalname) werden entfernt
    df = df[df['Gast'].str.lower() != 'nein']                                                            # Gastgeber Zeile wird entfernt
    df['Name (Originalname)'] = df['Name (Originalname)'].replace(regex=r'^\d+\s*-\s*|\(.*\)', value='')  # Namen werden bereinigt 1 - Vorname Nachname (Vorname Nachname) -> Vorname Nachname
    df = df.drop_duplicates(subset='Name (Originalname)', keep='last')                                    # Duplikate werden entfernt letzte wird einbehalten
    display(Markdown('#### --- **Zoom Data cleaned** ---')) 
    display(df)
    return df

def clean_webex_data(df:DataFrame):
    df = df[df['Rolle'] != 'host']
    df = df.drop_duplicates(subset='Anzeigename', keep='last')
    display(Markdown('#### --- **Webex Data cleaned** ---')) 
    display(df)
    return df

def filter_data(df:DataFrame, column_name):
    df = df.filter(items=column_name)
    df = df.astype(str)  
    df = df.drop_duplicates()
    return df

def clean_meeting_data(path):
    df = None
    keywords ={
    "zoom_keywords": ["Meeting-ID","Thema","Startzeit","Endzeit","Teilnehmer","Benutzer-E-Mail:","Dauer (Minuten)"],
    "webex_keywords": ["Meeting-Name","Meeting-Startzeit","Meeting-Endzeit","Anzeigename","Vorname","Nachname","Rolle","Beitrittszeit","Verbindungstyp","Sitzungsname"]
    }

    if os.path.splitext(path)[1].lower() == '.csv':
        encoding,_ = detect_encoding(path)
        delimiter = detect_delimiter(path,encoding)
        df = pd.read_csv(path,delimiter=delimiter,encoding=encoding)
    if os.path.splitext(path)[1].lower() == '.xlsx':
        df = pd.read_excel(path)

    is_zoom_data = contains_keywords(df, keywords["zoom_keywords"])
    is_webex_data = contains_keywords(df, keywords["webex_keywords"])
    dfc=df.copy()
    dfc = dfc.dropna(how='all') 

    if is_zoom_data:
        display(Markdown('#### --- **Zoom Data** ---'))
        display(dfc)
        dfc = clean_zoom_data(dfc)
        column_to_keep = ['Name (Originalname)']
        dfc = filter_data(dfc,column_to_keep)
        display(dfc)
    elif is_webex_data:
        display(Markdown('#### --- **WebEx Data** ---'))   
        display(dfc)                                                                      #alle Zeilen mit null werte werden entfernt
        dfc = clean_webex_data(dfc)
        column_to_keep = ['Anzeigename']
        dfc = filter_data(dfc,column_to_keep)
        display(dfc)
    return dfc