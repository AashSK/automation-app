# %%
import os
import re
import cv2
import numpy as np
import pandas as pd
from matplotlib import pyplot as plt
import pytesseract
from pandas import DataFrame

def display_image(img,title="image"):
    dpi = 80
    height, width  = img.shape[:2]
    figsize = (width / float(dpi))*0.6, (height / float(dpi))*0.6
    fig = plt.figure(figsize=figsize)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.axis('off')
    ax.imshow(img, cmap='gray')
    ax.set_title(title, fontsize=18, color='red', pad=20)
    plt.show()

def ocr_roi(img):
    ocr_result = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT,lang="deu",config="--psm 6")
    text = " ".join(word for word, conf in zip(ocr_result['text'], ocr_result['conf']) if float(conf) > -1)
    conf_values = [float(c) for c in ocr_result['conf'] if float(c) > -1]
    conf = sum(conf_values) / len(conf_values) if conf_values else 0.0 
    return text, conf

def improve_ocr_result(df,img,path):
    for index, row in df.iterrows():
        if row['conf'] > -1 and row['conf'] < 50:
            x, y, w, h = row['left']-2, row['top']-2, row['width']+4, row['height']+4
            roi = img[y:y+h, x:x+w] 
            os.makedirs(path, exist_ok=True)
            cv2.imwrite(f'{path}/roi_{x}.png',roi)
            new_text, new_conf = ocr_roi(roi)
            if new_conf > row['conf'] and new_text:
                df.at[index, 'text'] = new_text
                df.at[index, 'conf'] = new_conf 
    return df

def draw_contours(img, df):
    img = img.copy()

    boxes = df[['left', 'top', 'width', 'height']].values

    rectangles = np.array([
        [[x, y], [x + w, y], [x + w, y + h], [x, y + h], [x, y]]
        for x, y, w, h in boxes
    ], dtype=np.int32)

    cv2.polylines(img, rectangles, isClosed=True, color=(0, 255, 0), thickness=2)

    return img

def recheck_output_table(df:DataFrame):
    matching_indices_al = df[df.apply(lambda row: row.astype(str).str.contains('Nr.|Name|Vorname|Druckschrift!|Hochschule|Unterschrift|Matr.-Nr. oder FH-zugehörigkeit', na=False).any(), axis=1)].index[0]
    dfc =  df.iloc[matching_indices_al:].reset_index(drop=True)
    for row in dfc.itertuples():
        i = row.Index
        if i == 0:
            continue
        if not row[2]:
            if not row[1]:
                continue
            else:
               if len(row[1]) > 3:
                    row_in_converted = df.iloc[i+matching_indices_al,:]
                    row_in_converted.iloc[0] = f"{i}."
                    row_in_converted.iloc[1] = re.sub(r'[^a-zA-Z\s]', '', row[1]) 
                    df.iloc[i + matching_indices_al, :] = row_in_converted
