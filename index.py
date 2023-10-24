from dash import Dash, dcc, html, dash_table, callback, Output, Input, State
import pandas as pd
from datetime import date, timedelta


columns = ['FECHA', 'ORIGEN', 'NRO LEGAJO', 'APELLIDO', 'NOMBRE', 'APODO',
           'DNI', 'CUIT', 'FRECUENCIA DE PAGO', 'MODALIDAD DE PAGO',
           'MATERIAL', 'KG', 'OBSERVACIONES', 'PRECIO x KG', 'KG VALORIZADO']


def read_excel_sheet(filename, sheet_name):
    '''Read the excel sheet with the data and normalize the data
    '''

    df = pd.read_excel(filename, sheet_name=sheet_name, header=3)
    df.rename({"Unnamed: 2": "NRO LEGAJO", " ": "FECHA", "MEZCLA": "MATERIAL",
               "Unnamed: 1": "ORIGEN"}, axis=1, inplace=True)
    dtypes = {"NRO LEGAJO": "Int64", "MATERIAL": str}
    df.astype(dtypes)
    df["KG"] = pd.to_numeric(df.KG, errors='coerce', downcast="integer")
    df["CUIT"] = pd.to_numeric(df.CUIT.replace("-", ""), errors='coerce',
                               downcast="integer")
    df["FECHA"] = pd.to_datetime(df["FECHA"], dayfirst=True, errors='coerce')
    df["MATERIAL"] = df["MATERIAL"].str.title()
    df = df[columns[:-2]]  # ["NRO LEGAJO", "FECHA", "MATERIAL", "PESO"]]
    return df


df = read_excel_sheet("Copia de BASE OPERACIÓN - AVELLANEDA.xlsx",
                      "INGRESO DE MATERIAL")
price_df = pd.read_excel("PLANILLA BASE para armar SOLICITUD PAGO (2).xlsx",
                         sheet_name="BASE PRECIOS")
price_df.dropna(inplace=True)
price_df["MATERIAL"] = price_df["MATERIAL"].str.title()

# Initialize app
app = Dash(__name__)


app.layout = html.Div(children=[
    html.Div(children="Interactive App"),
    dash_table.DataTable(id="summary-table", data=[], page_size=10),

    html.Div(id="total-table"),
    dash_table.DataTable(id='price-table',
                         columns=[{"id": "MATERIAL", "name": "MATERIAL"},
                                  {"id": "PRECIO POR KG",
                                   "name": "PRECIO POR KG"}],
                         data=price_df.to_dict("records"), persistence=True,
                         persisted_props=["data"], editable=True),
    # Add dropdown
    dcc.DatePickerRange(
        id='date-picker',
        min_date_allowed=date(2015, 8, 5),
        max_date_allowed=date.today(),
        start_date=date.today()-timedelta(weeks=4),
        end_date=date.today()
    ),
    html.Button('Guardar', id='save-button', n_clicks=0),  # Add a save button
    # Add a Div that displays the status
    html.Div(id="save-status", children=["Nothing saved yet"])
])


@callback(
    [
        Output(component_id='summary-table', component_property='data'),
        Output(component_id='total-table', component_property='children')
    ],
    [
        Input(component_id='date-picker', component_property='start_date'),
        Input(component_id='date-picker', component_property='end_date'),
        Input(component_id='price-table', component_property='data'),
        Input(component_id='price-table', component_property='columns')
    ]
)
def update_table(start_date, end_date, rows, columns):
    try:
        df_filter = df.loc[df.FECHA >= start_date].copy()
        df_filter = df_filter.loc[df.FECHA <= end_date].copy()
    except Exception:
        print("An error ocurred while filtering the dataframe")
    df_filter["FECHA"] = pd.DatetimeIndex(df_filter.FECHA).date
    current_prices = pd.DataFrame(rows, columns=[c['name'] for c in columns])
    df_filter["PRECIO POR KG"] = df_filter.apply(get_price, axis=1,
                                                 args=[current_prices])
    df_filter["KG VALORIZADO"] = df_filter["PRECIO POR KG"]*df_filter["KG"]
    df_group = df_filter[["CUIT",
                          "KG VALORIZADO"]].groupby("CUIT").sum().reset_index()
    total_table = dash_table.DataTable(data=df_group.to_dict("records"),
                                       page_size=10),
    df_filter = df_filter.to_dict("records")
    return df_filter, total_table


def get_price(row, data_precios):
    material = row.MATERIAL
    df_precios = data_precios
    price = df_precios.loc[df_precios.MATERIAL == material]["PRECIO POR KG"]
    try:
        price = int(price.iloc[0])
    except Exception:
        print(f"Hay un material desconocido:{material}")
        price = 0

    return price


# Callback to save table data to Excel
@app.callback(
    Output("save-status", "children"),
    State("summary-table", "data"),
    Input("save-button", "n_clicks")
)
def save_table_data_to_excel(table_data, n_clicks):
    df_filter = pd.DataFrame(table_data)
    sheet_name = "INGRESO DE MATERIAL VALORIZADO "
    if n_clicks == 0:
        return "Nothing saved yet"
    # Open the file with an Excel writer object
    try:
        file = "PLANILLA BASE para armar SOLICITUD PAGO (2).xlsx"
        with pd.ExcelWriter(file, engine='openpyxl', mode="a",
                            if_sheet_exists="replace") as excel_writer:

            # Add the Inland data to a new sheet
            df_filter.to_excel(excel_writer, sheet_name=sheet_name,
                               index=False)
    except Exception as error:
        print(error)
        save_status = "Error while saving the data"
    else:
        save_status = "Data successfully saved"
    return save_status


if __name__ == "__main__":
    app.run_server(debug=True)
