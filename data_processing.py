from multiprocessing import Value
import pandas as pd
import numpy as np

def corregir_dtype(df, columna, dtype):
    '''Función que corrige los datatypes de una serie de Pandas
    '''
    if dtype == int:
        serie_corregida = pd.to_numeric(df[columna], errors='coerce', downcast="integer")
    elif dtype == np.datetime64:
        serie_corregida = pd.to_datetime(df[columna], dayfirst=True, errors='coerce')
    elif dtype == str:
        serie_corregida = df[columna].str.title()
        serie_corregida = serie_corregida.str.strip()
    else:
        raise Exception(f"El tipo '{dtype}' es desconocido")
    return serie_corregida

def get_price(row, price_data, summary_df, threshold=1000):
    '''Función que obtiene el precio de un dado material.
    '''
    material = row.MATERIAL
    try:
        monthly_total = summary_df[summary_df.CUIT == row["CUIT"]]["KG"].iloc[0]
        bonus = monthly_total >= threshold
    except Exception:
        bonus = False
    if bonus and any(price_data.MATERIAL == material+"_Bono"):
        material += "_Bono"
    
    price = price_data.loc[price_data.MATERIAL == material]["PRECIO POR KG"]
    try:
        price = int(price.iloc[0])
    except Exception:
        print(f"Hay un material desconocido:{material}")
        price = 0

    return price

def calculate_monthly_total(df, month, year):
    month_series = pd.DatetimeIndex(df["FECHA"]).month
    year_series = pd.DatetimeIndex(df["FECHA"]).year
    df_filter = df[(month_series==month) & (year_series==year)]
    monthly_total = df_filter[["CUIT", "KG"]].groupby("CUIT").sum().reset_index()
    return monthly_total
