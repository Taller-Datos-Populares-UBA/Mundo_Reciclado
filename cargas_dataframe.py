import pandas as pd
import numpy as np
from data_processing import corregir_dtype


class MTEDataFrame:
    FILES_TO_LOAD = None
    _instance = None

    print("RONILOG cargando la clase")

    COLUMNAS = ['FECHA', 'ORIGEN', 'NRO LEGAJO', 'APELLIDO', 'NOMBRE', 'APODO',
                'DNI', 'CUIT', 'FRECUENCIA DE PAGO', 'MODALIDAD DE PAGO',
                'MATERIAL', 'KG', 'OBSERVACIONES', 'PRECIO x KG', 'KG VALORIZADO']
    
    def __init__(self):
        raise Exception("Cannot instanciate a singleton")

    @classmethod
    def get_instance(cls):
        if cls._instance is None:
            cls._instance = cls._create_instance()
        return cls._instance.copy()

    @classmethod
    def _create_instance(cls):
        print("files to load ")
        print(cls.FILES_TO_LOAD)
        dtypes = {'CUIT': int,
                  'KG': int,
                  'FECHA': np.datetime64,
                  'MATERIAL': str}
        
        dfs_per_date = [cls._read_cargas_excel_sheet(filename) for filename in cls.FILES_TO_LOAD]
        df = pd.concat(dfs_per_date, ignore_index=True)

        # Renombrar columnas mal escritas en el excel original
        df.rename({"Unnamed: 2": "NRO LEGAJO", " ": "FECHA", "MEZCLA": "MATERIAL",
                "Unnamed: 1": "ORIGEN"}, axis=1, inplace=True)
        # Hay que quitar los guiones para transformar al CUIT en int 
        df["CUIT"] = df["CUIT"].replace("-", "")
        for columna, dtype in dtypes.items():
            df[columna] = corregir_dtype(df, columna, dtype)
        #df = df.fillna('No especificado')
        df = df[cls.COLUMNAS[:-2]]
        #df.rename(lambda x: str.lower(x),axis=1, inplace=True)
        return df

    @classmethod
    def _read_cargas_excel_sheet(cls, filename, sheet_name="INGRESO DE MATERIAL"):
        '''Read the excel sheet with the data and normalize the data
        '''

        df = pd.read_excel(filename, sheet_name=sheet_name, header=3)
        return df