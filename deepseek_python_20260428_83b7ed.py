import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import dash
from dash import dcc, html, Input, Output
import dash_bootstrap_components as dbc
from datetime import datetime, timedelta
import numpy as np

# ===================== CONFIGURACIÓN INICIAL =====================
# Cargar los datos
df = pd.read_excel('Tiempos.xlsx')

# Limpieza de nombres de columnas
df.columns = df.columns.str.strip()

# Filtrar solo filas con inicio de sesión válido
df = df.dropna(subset=['Inicio de Sesión'])

# Convertir fechas
df['Fecha'] = pd.to_datetime(df['Fecha'])
df['Inicio de Sesión'] = pd.to_datetime(df['Inicio de Sesión'])
df['Cerrado de Sesión'] = pd.to_datetime(df['Cerrado de Sesión'])

# ===================== DEFINIR REGLAS DE TIEMPOS =====================
AGENTES_ESPECIALES = [
    "PAMELA VALERIA  ARELLANO MARADIAGA",
    "JESUS ANGEL ZAVALA SANCHEZ",
    "JATZIRI NOHEMI SANCHEZ ARELLANO",
    "JAIME PACHECO GONZALEZ",
    "STEFAN ARROYO OCHOA",
    "WENDY VERONICA BAEZ MONROY",
    "PAULA BENITEZ RODRIGUEZ",
    "EVELYN TAMARA LARIOS RUBIO"
]

REGLAS = {
    'normal': {
        'Comida': 60,
        'Baño': 15,
        'Retro': 10
    },
    'especial': {
        'Comida': 20,  # Break en lugar de comida
        'Baño': 10,
        'Retro': 10
    }
}

# Mapear columnas de tiempo
COLUMNAS_TIEMPO = {
    'PAUSECOMIDAMinutes': 'Comida',
    'TIMEEXCEEDEDPAUSECOMIDAMinutes': 'Comida_Exceso',
    'PAUSEBAÑOMinutes': 'Baño',
    'TIMEEXCEEDEDPAUSEBAÑOMinutes': 'Baño_Exceso',
    'PAUSERETROMinutes': 'Retro',
    'TIMEEXCEEDEDPAUSERETROMinutes': 'Retro_Exceso',
    'PAUSERETROCALIDADMinutes': 'Retro_Calidad',
    'TIMEEXCEEDEDPAUSERETROCALIDADMinutes': 'Retro_Calidad_Exceso',
    'PAUSECAPACITACIONMinutes': 'Capacitacion',
    'TIMEEXCEEDEDPAUSECAPACITACIONMinutes': 'Capacitacion_Exceso',
    'PAUSECOTIZACIONMinutes': 'Cotizacion',
    'TIMEEXCEEDEDPAUSECOTIZACIONMinutes': 'Cotizacion_Exceso'
}

# Crear columnas de tiempo real usado
for col, nombre in COLUMNAS_TIEMPO.items():
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

# Calcular si excede o tiene tiempo de más
def calcular_estado_tiempo(row, tipo_pausa, limite):
    tiempo_usado = row.get(f'PAUSE{tipo_pausa}Minutes', 0)
    exceso_col = f'TIMEEXCEEDEDPAUSE{tipo_pausa}Minutes'
    exceso = row.get(exceso_col, 0)
    
    if tiempo_usado > limite:
        return f"⚠️ Excede por {tiempo_usado - limite:.1f} min", tiempo_usado - limite
    elif tiempo_usado < limite and tiempo_usado > 0:
        return f"✅ Tiene {limite - tiempo_usado:.1f} min disponibles", -(limite - tiempo_usado)
    else:
        return "✅ Dentro del límite", 0

# Aplicar reglas
df['Tipo_Agente'] = df['Nombre'].apply(lambda x: 'especial' if x in AGENTES_ESPECIALES else 'normal')

# Calcular para Comida, Baño, Retro
df['Comida_Estado'] = df.apply(lambda row: calcular_estado_tiempo(row, 'COMIDA', REGLAS[row['Tipo_Agente']]['Comida'])[0], axis=1)
df['Comida_Diferencia'] = df.apply(lambda row: calcular_estado_tiempo(row, 'COMIDA', REGLAS[row['Tipo_Agente']]['Comida'])[1], axis=1)

df['Baño_Estado'] = df.apply(lambda row: calcular_estado_tiempo(row, 'BAÑO', REGLAS[row['Tipo_Agente']]['Baño'])[0], axis=1)
df['Baño_Diferencia'] = df.apply(lambda row: calcular_estado_tiempo(row, 'BAÑO', REGLAS[row['Tipo_Agente']]['Baño'])[1], axis=1)

df['Retro_Estado'] = df.apply(lambda row: calcular_estado_tiempo(row, 'RETRO', REGLAS[row['Tipo_Agente']]['Retro'])[0], axis=1)
df['Retro_Diferencia'] = df.apply(lambda row: calcular_estado_tiempo(row, 'RETRO', REGLAS[row['Tipo_Agente']]['Retro'])[1], axis=1)

# ===================== PREPARAR DATOS PARA EL DASHBOARD =====================
# Agregar columnas de tiempo
df['Semana'] = df['Fecha'].dt.isocalendar().week
df['Mes'] = df['Fecha'].dt.strftime('%Y-%m')
df['Dia'] = df['Fecha'].dt.date

# Calcular tiempo total en diferentes estados
df['Tiempo_Disponible'] = pd.to_numeric(df.get('DISPONIBLE (Minutos)', 0), errors='coerce').fillna(0)
df['Tiempo_EnLlamada'] = pd.to_numeric(df.get('EN LLAMADA (Minutos)', 0), errors='coerce').fillna(0)
df['Tiempo_DespuesLlamada'] = pd.to_numeric(df.get('DESPUES DE LLAMADA (Minutos)', 0), errors='coerce').fillna(0)
df['Tiempo_Pausa'] = pd.to_numeric(df.get('PAUSA (Minutos)', 0), errors='coerce').fillna(0)

# ===================== CREAR DASHBOARD =====================
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

app.layout = dbc.Container([
    dbc.Row([
        dbc.Col(html.H1("Dashboard de Tiempos de Agentes", className="text-center my-4"), width=12)
    ]),
    
    dbc.Row([
        dbc.Col([
            html.Label("Seleccionar Periodo:"),
            dcc.RadioItems(
                id='periodo-selector',
                options=[
                    {'label': 'Día', 'value': 'Dia'},
                    {'label': 'Semana', 'value': 'Semana'},
                    {'label': 'Mes', 'value': 'Mes'}
                ],
                value='Dia',
                inline=True
            )
        ], width=6),
        dbc.Col([
            html.Label("Seleccionar Fecha/Semana/Mes:"),
            dcc.Dropdown(id='periodo-dropdown', value=None)
        ], width=6)
    ], className="mb-4"),
    
    dbc.Row([
        dbc.Col(dcc.Graph(id='tiempos-agentes'), width=12)
    ]),
    
    dbc.Row([
        dbc.Col(dcc.Graph(id='cumplimiento-pausas'), width=12)
    ]),
    
    dbc.Row([
        dbc.Col(dcc.Graph(id='excesos-tiempo'), width=12)
    ]),
    
    dbc.Row([
        dbc.Col(html.Div(id='tabla-resumen', className="mt-4"), width=12)
    ])
], fluid=True)

# ===================== CALLBACKS =====================
@app.callback(
    [Output('periodo-dropdown', 'options'),
     Output('periodo-dropdown', 'value')],
    [Input('periodo-selector', 'value')]
)
def update_dropdown_options(periodo):
    if periodo == 'Dia':
        opciones = [{'label': str(d), 'value': d} for d in sorted(df['Dia'].unique())]
    elif periodo == 'Semana':
        opciones = [{'label': f"Semana {semana}", 'value': semana} for semana in sorted(df['Semana'].unique())]
    else:  # Mes
        opciones = [{'label': mes, 'value': mes} for mes in sorted(df['Mes'].unique())]
    
    valor_defecto = opciones[0]['value'] if opciones else None
    return opciones, valor_defecto

@app.callback(
    [Output('tiempos-agentes', 'figure'),
     Output('cumplimiento-pausas', 'figure'),
     Output('excesos-tiempo', 'figure'),
     Output('tabla-resumen', 'children')],
    [Input('periodo-selector', 'value'),
     Input('periodo-dropdown', 'value')]
)
def update_dashboard(periodo, valor):
    if valor is None:
        return {}, {}, {}, ""
    
    # Filtrar datos según periodo
    if periodo == 'Dia':
        df_filtrado = df[df['Dia'] == valor]
        titulo_periodo = f"Día: {valor}"
    elif periodo == 'Semana':
        df_filtrado = df[df['Semana'] == valor]
        titulo_periodo = f"Semana {valor}"
    else:
        df_filtrado = df[df['Mes'] == valor]
        titulo_periodo = f"Mes: {valor}"
    
    # Agrupar por agente
    df_agente = df_filtrado.groupby('Nombre').agg({
        'Tiempo_Disponible': 'sum',
        'Tiempo_EnLlamada': 'sum',
        'Tiempo_DespuesLlamada': 'sum',
        'Tiempo_Pausa': 'sum',
        'Comida_Diferencia': 'sum',
        'Baño_Diferencia': 'sum',
        'Retro_Diferencia': 'sum',
        'Llamadas Manejadas': 'sum'
    }).reset_index()
    
    # Gráfico 1: Tiempos por agente
    fig1 = make_subplots(rows=1, cols=1)
    for col in ['Tiempo_Disponible', 'Tiempo_EnLlamada', 'Tiempo_DespuesLlamada', 'Tiempo_Pausa']:
        fig1.add_trace(go.Bar(
            x=df_agente['Nombre'], y=df_agente[col], name=col.replace('Tiempo_', '')
        ))
    fig1.update_layout(
        title=f"Distribución de Tiempos por Agente - {titulo_periodo}",
        barmode='group', xaxis_tickangle=-45, height=500
    )
    
    # Gráfico 2: Cumplimiento de pausas (diferencias)
    df_melt = df_agente.melt(id_vars=['Nombre'], value_vars=['Comida_Diferencia', 'Baño_Diferencia', 'Retro_Diferencia'],
                             var_name='Tipo_Pausa', value_name='Diferencia')
    df_melt['Estado'] = df_melt['Diferencia'].apply(lambda x: 'Excede' if x > 0 else ('Faltante' if x < 0 else 'Cumple'))
    
    fig2 = px.bar(df_melt, x='Nombre', y='Diferencia', color='Tipo_Pausa',
                  title=f"Cumplimiento de Tiempos de Pausa - {titulo_periodo}",
                  barmode='group', labels={'Diferencia': 'Diferencia (minutos)', 'Nombre': 'Agente'})
    fig2.update_layout(xaxis_tickangle=-45, height=500)
    fig2.add_hline(y=0, line_dash="dash", line_color="red")
    
    # Gráfico 3: Excesos de tiempo (solo valores positivos)
    excesos = df_agente[df_agente['Comida_Diferencia'] > 0][['Nombre', 'Comida_Diferencia']].rename(columns={'Comida_Diferencia': 'Exceso Comida'})
    excesos_bano = df_agente[df_agente['Baño_Diferencia'] > 0][['Nombre', 'Baño_Diferencia']].rename(columns={'Baño_Diferencia': 'Exceso Baño'})
    excesos_retro = df_agente[df_agente['Retro_Diferencia'] > 0][['Nombre', 'Retro_Diferencia']].rename(columns={'Retro_Diferencia': 'Exceso Retro'})
    
    if not excesos.empty or not excesos_bano.empty or not excesos_retro.empty:
        # Combinar excesos
        df_excesos = pd.DataFrame()
        if not excesos.empty:
            df_excesos = excesos
        if not excesos_bano.empty:
            if df_excesos.empty:
                df_excesos = excesos_bano
            else:
                df_excesos = df_excesos.merge(excesos_bano, on='Nombre', how='outer')
        if not excesos_retro.empty:
            if df_excesos.empty:
                df_excesos = excesos_retro
            else:
                df_excesos = df_excesos.merge(excesos_retro, on='Nombre', how='outer')
        
        df_excesos = df_excesos.fillna(0)
        df_excesos_melt = df_excesos.melt(id_vars=['Nombre'], var_name='Tipo', value_name='Minutos_Exceso')
        df_excesos_melt = df_excesos_melt[df_excesos_melt['Minutos_Exceso'] > 0]
        
        fig3 = px.bar(df_excesos_melt, x='Nombre', y='Minutos_Exceso', color='Tipo',
                      title=f"Excesos de Tiempo por Agente - {titulo_periodo}",
                      labels={'Minutos_Exceso': 'Minutos en Exceso', 'Nombre': 'Agente'})
        fig3.update_layout(xaxis_tickangle=-45, height=500)
    else:
        fig3 = go.Figure()
        fig3.add_annotation(text="No hay excesos de tiempo en este periodo", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False)
        fig3.update_layout(title=f"Excesos de Tiempo - {titulo_periodo}", height=500)
    
    # Tabla resumen
    tabla = dbc.Table([
        html.Thead(html.Tr([html.Th("Agente"), html.Th("Llamadas"), html.Th("Tiempo Disp (min)"), 
                            html.Th("Tiempo Llamada (min)"), html.Th("Estado Comida"), 
                            html.Th("Estado Baño"), html.Th("Estado Retro")])),
        html.Tbody([
            html.Tr([
                html.Td(row['Nombre']),
                html.Td(row['Llamadas Manejadas']),
                html.Td(f"{row['Tiempo_Disponible']:.1f}"),
                html.Td(f"{row['Tiempo_EnLlamada']:.1f}"),
                html.Td(df_filtrado[df_filtrado['Nombre'] == row['Nombre']]['Comida_Estado'].iloc[0] if not df_filtrado[df_filtrado['Nombre'] == row['Nombre']].empty else "N/A"),
                html.Td(df_filtrado[df_filtrado['Nombre'] == row['Nombre']]['Baño_Estado'].iloc[0] if not df_filtrado[df_filtrado['Nombre'] == row['Nombre']].empty else "N/A"),
                html.Td(df_filtrado[df_filtrado['Nombre'] == row['Nombre']]['Retro_Estado'].iloc[0] if not df_filtrado[df_filtrado['Nombre'] == row['Nombre']].empty else "N/A")
            ]) for _, row in df_agente.iterrows()
        ])
    ], striped=True, bordered=True, hover=True, responsive=True)
    
    return fig1, fig2, fig3, tabla

if __name__ == '__main__':
    app.run_server(debug=True, port=8050)