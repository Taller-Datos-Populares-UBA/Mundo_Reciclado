from dash import Dash, dcc, html, dash_table, callback, Output, Input, State
import pandas as pd
from datetime import date, timedelta, datetime
from data_processing import get_price, calculate_monthly_total
from cargas_dataframe import MTEDataFrame
from utils.utils import estilo_tabla

import dash_bootstrap_components as dbc

MTEDataFrame.FILES_TO_LOAD = ["Copia de BASE OPERACIÓN - AVELLANEDA.xlsx"]
df = MTEDataFrame.get_instance()

price_df = pd.read_excel("PLANILLA BASE para armar SOLICITUD PAGO (2).xlsx",
                         sheet_name="BASE PRECIOS")
price_df.dropna(inplace=True)
price_df["MATERIAL"] = price_df["MATERIAL"].str.title()
price_df["MATERIAL"] = price_df["MATERIAL"].str.strip()


# meta_tags are required for the app layout to be mobile responsive
external_scripts = ["https://cdn.plot.ly/plotly-locale-es-latest.js"]  # Agregamos español como lenguaje
# Initialize app
app = Dash(__name__, suppress_callback_exceptions=True,
                external_stylesheets=['https://codepen.io/chriddyp/pen/bWLwgP.css', dbc.themes.LITERA],
                external_scripts=external_scripts,
                meta_tags=[{'name': 'viewport',
                            'content': 'width=device-width, initial-scale=1.0'}]
                )


app.layout = html.Div([
        html.Div(id="div-izquierda-panel", className="botonera", children=[
            html.H6(
                "Filtros",
                className="title-botonera"
            ),
            html.Div(children=[
                html.Label("Elegí el rango de fechas", className="labels"),
                html.Div(children=[
                    dcc.DatePickerRange(
                        id='date-picker',
                        min_date_allowed=date(2015, 8, 5),
                        max_date_allowed=date.today(),
                        start_date=date.today()-timedelta(weeks=4),
                        end_date=date.today()
                    )
                ], className="dates-container"
                ),
            ]),
            html.Div(children=[
                html.Label("Elegí el periodo para calcular el bono:", className="labels"),
                html.Div(children=[
                    dcc.DatePickerSingle(
                        id="bonus-button",
                        min_date_allowed=date(2015, 8, 5),
                        max_date_allowed=date.today(),
                        date=date.today(),
                        display_format="MM/YYYY"
                    )
                ], className="dates-container"
                ),
            ]),
            html.Div(children=[
                html.Button('Guardar', id='save-button', n_clicks=0, className="mr-1 mt-1 btn btn-primary search-button"),  # Add a save button
                
            ], className="button-container"
            ),
            # Add a Div that displays the status
            html.Div(id="save-status", children=["Nothing saved yet"], className="button-container"),
        ]),
    html.Div(id="div-derecha", className="output", children=[
    
        dash_table.DataTable(id="summary-table", data=[], page_size=10, **estilo_tabla),
        dash_table.DataTable(id="total-table", data=[],page_size=10, **estilo_tabla),
        dash_table.DataTable(id='price-table',
                            columns=[{"id": "MATERIAL", "name": "MATERIAL"},
                                    {"id": "PRECIO POR KG",
                                    "name": "PRECIO POR KG"}],
                            data=price_df.to_dict("records"), persistence=True,
                            persisted_props=["data"], editable=True,
                            **estilo_tabla),
    ])
])


@callback(
    [
        Output(component_id='summary-table', component_property='data'),
        Output(component_id='total-table', component_property='data')
    ],
    [
        Input(component_id='date-picker', component_property='start_date'),
        Input(component_id='date-picker', component_property='end_date'),
        Input(component_id='price-table', component_property='data'),
        Input(component_id='price-table', component_property='columns'),
        Input(component_id='bonus-button', component_property='date')
    ]
)
def update_table(start_date, end_date, rows, columns, bonus_date):
    try:
        df_filter = df.loc[df.FECHA >= start_date].copy()
        df_filter = df_filter.loc[df.FECHA <= end_date].copy()
    except Exception:
        print("An error ocurred while filtering the dataframe")
    df_filter["FECHA"] = pd.DatetimeIndex(df_filter.FECHA).date
    current_prices = pd.DataFrame(rows, columns=[c['name'] for c in columns])
    bonus_date = pd.to_datetime(bonus_date)
    bonus_month = bonus_date.month
    bonus_year = bonus_date.year
    monthly_total = calculate_monthly_total(df, bonus_month, bonus_year)
    df_filter["PRECIO POR KG"] = df_filter.apply(get_price, axis=1,
                                                 args=[current_prices, monthly_total])
    df_filter["KG VALORIZADO"] = df_filter["PRECIO POR KG"]*df_filter["KG"]
    df_group = df_filter[["CUIT",
                          "KG VALORIZADO"]].groupby("CUIT").sum().reset_index()
    df_total = df_group.to_dict("records")
    df_filter = df_filter.to_dict("records")
    return df_filter, df_total




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
        return "Nada guardado aún"
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
        save_status = "Error al guardar los datos"
    else:
        save_status = "Datos guardados exitosamente"
    return save_status


if __name__ == "__main__":
    app.run_server(debug=True)
