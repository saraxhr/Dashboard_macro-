"""RETO MACRO · Guatemala | Streamlit, Python 3.10 o superior.

Instalación (una vez en tu entorno):
    python -m pip install 'streamlit>=1.50,<2' 'pandas>=2.2,<4' 'plotly>=5.24,<7' openpyxl xlrd
Ejecución:
    python -m streamlit run dashboard_macro.py

Coloca los diez archivos originales junto a este script, en data/ o en upload/.
También puedes cargarlos desde la barra lateral. No se consulta ninguna API.
Se admiten los Excel originales y CSV con las mismas columnas de la hoja de datos.
Las explicaciones y escenarios son orientaciones conceptuales, no pronósticos.
"""
from __future__ import annotations

from io import BytesIO
from pathlib import Path
from datetime import date
from html import escape
import re

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


FILES = {
    1: 'Variable1PIB_Guatemala_Unificado_Sin_Colores.xlsx',
    2: 'Variable2Inflacion_Guatemala_Unificada_Sin_Colores.xlsx',
    3: 'Variable3tasadeintereslider.xls',
    4: 'Variable4Tasas_Activa_Pasiva_Spread_Guatemala.xlsx',
    5: 'Variable5Base_Monetaria_M1_M2_Guatemala_Limpia.xlsx',
    6: 'Variable6CreditosPrivado.xlsx',
    7: 'Variable7TipodeCambio.csv',
    8: 'Variable8Remesas_Familiares_Guatemala_Base_Limpia.xlsx',
    9: 'Variable9Reservas_Internacionales_Guatemala_Base_Limpia.xlsx',
    10: 'Variable10Petroleo_Combustibles_Guatemala_Base_Limpia.xlsx',
}
COLORS = ['#12b8a6', '#5778e8', '#edaa43', '#de6d91', '#8e73d3']

# No se inventan enlaces de descarga. Solo se incorporan enlaces presentes
# en los archivos, con posibilidad de completar los pendientes desde la interfaz.
DEFAULT_URLS = {2: 'https://www.ine.gob.gt/',
                10: 'https://www.eia.gov/dnav/pet/hist/LeafHandler.ashx?n=PET&s=RWTC&f=M'}


def file_number(name):
    m = re.match(r'Variable(10|[1-9])(?:\D|$)', Path(name).name, re.I)
    return int(m.group(1)) if m else None


@st.cache_data(show_spinner=False)
def read_file(name: str, content: bytes):
    """Lee la hoja principal y conserva los diccionarios para inspección."""
    if name.lower().endswith('.csv'):
        text = None
        for enc in ('utf-8-sig', 'cp1252'):
            try:
                text = content.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            raise ValueError('No se pudo identificar la codificación del CSV.')
        lines = text.splitlines()
        skip = next((i for i, s in enumerate(lines)
                     if 'Fecha' in s and 'Compra' in s and 'Venta' in s), 0)
        df = pd.read_csv(BytesIO(content), encoding=enc, skiprows=skip, sep=None, engine='python')
        return {'Datos': df}
    book = pd.ExcelFile(BytesIO(content))
    return {s: pd.read_excel(book, sheet_name=s) for s in book.sheet_names}


def primary(book):
    return book.get('Datos', next(iter(book.values()))).copy()


def clean_series(df, dates, values):
    out = pd.DataFrame({'fecha': pd.to_datetime(dates, errors='coerce'),
                        'valor': pd.to_numeric(values, errors='coerce')})
    out = out.replace([np.inf, -np.inf], np.nan).dropna().sort_values('fecha')
    if out.fecha.duplicated().any():
        raise ValueError('Fechas duplicadas: revise el filtro de indicador antes de graficar.')
    return out.reset_index(drop=True)


def annual_dates(year):
    return pd.to_datetime(year.astype('Int64').astype(str) + '-01-01', errors='coerce')


def yoy_monthly(series):
    """Une contra el mismo mes del año anterior; nunca desplaza 12 filas a ciegas."""
    s = series.copy()
    s['mes'] = s.fecha.dt.to_period('M')
    lookup = s.set_index('mes').valor
    previous = (s.mes - 12).map(lookup)
    s['valor'] = (s.valor / previous.where(previous.ne(0)) - 1) * 100
    return s[['fecha', 'valor']].dropna().reset_index(drop=True)


def build_catalog(books, credit_rebuild=False):
    """Cada serie tiene seis elementos de ficha y trazabilidad al archivo recibido."""
    result, issues = {}, []

    def add(key, n, title, group, unit, freq, df, dates, values, definition,
            reading, relations, decision, source, note='', method='Valor del archivo, sin estimación.'):
        series = clean_series(df, dates, values)
        result[key] = dict(key=key, n=n, title=title, group=group, unit=unit,
                           freq=freq, data=series, definition=definition, reading=reading,
                           relations=relations, decision=decision, source=source,
                           note=note, method=method)

    for n, book in books.items():
        d = primary(book)
        try:
            if n == 1:
                a = d[(d.Codigo == 'PIB') & (d.Enfoque == 'Produccion') & (d.Medicion == 'Real')]
                add('pib', n, 'Crecimiento del PIB real', 'Actividad', '%', 'Anual', a,
                    annual_dates(a.Anio), a.Crecimiento_interanual_pct,
                    'Variación anual del volumen de producción de bienes y servicios.',
                    'Un crecimiento mayor sugiere expansión de actividad; uno negativo, contracción. No existe un umbral universal para las ventas de cada empresa.',
                    'Mayor actividad suele acompañarse de más consumo y crédito; el efecto depende del sector.',
                    'Escalonar inventario y metas de venta según la demanda observada; evitar expandirse solo por el PIB agregado.',
                    'Banco de Guatemala, según el archivo.',
                    'El archivo identifica 2023–2025 como preliminares. Se utiliza un solo enfoque del PIB.',
                    'Crecimiento_interanual_pct publicado en la base; PIB real, enfoque Producción.')
            elif n == 2:
                for key, col, title, unit, definition, reading, decision in [
                    ('inflacion', 'Inflacion_interanual_pct', 'Inflación interanual', '%',
                     'Cambio del nivel general de precios frente al mismo mes del año anterior.',
                     'Comparar con 3–5 %, rango de la consigna. Por debajo indica menor inflación, no necesariamente caída de precios; por encima, mayor presión de precios.',
                     'Revisar precios y tamaños de presentación según costos propios y poder adquisitivo.'),
                    ('ipc', 'IPC', 'Índice de precios al consumidor', 'puntos de índice',
                     'Índice que resume el nivel de precios de una canasta de consumo.',
                     'Un IPC mayor significa una canasta más cara respecto a la base del índice. El nivel del IPC no es una tasa de inflación.',
                     'Comparar la evolución de precios de la categoría con la canasta general antes de ajustar el catálogo.')]:
                    a = d[(d.grupo == 'Inflacion_general') & (d.indicador == col)]
                    add(key, n, title, 'Precios', unit, 'Mensual', a, a.periodo, a.valor,
                        definition, reading, 'La inflación puede responder a costos energéticos, tipo de cambio y demanda; el IPC permite calcular variaciones de precios.',
                        decision, 'INE: IPC octubre de 2026, según el archivo.',
                        'Solo 13 meses disponibles (octubre 2025–octubre 2026); no se completa el histórico faltante.')
            elif n == 3:
                add('lider', n, 'Tasa de interés líder', 'Financiamiento', '%', 'Mensual', d,
                    pd.to_datetime(d.Fechas, format='%d-%m-%Y', errors='coerce'), d.TasaInteresLider,
                    'Tasa de referencia para la política monetaria.',
                    'Una subida señala una postura monetaria más restrictiva; una baja, más expansiva. Evaluar su trayectoria, sin asumir un umbral fijo de normalidad.',
                    'Puede influir en tasas bancarias, demanda de crédito e inflación con rezagos y transmisión incompleta.',
                    'Revisar el costo de promociones financiadas y priorizar ofertas con menor necesidad de capital de trabajo.',
                    'Institución no consignada dentro del archivo; la consigna identifica a Banguat como referencia.')
            elif n == 4:
                specs = [
                    ('activa','tasa_activa_pct','Tasa activa','%', 'Costo promedio ponderado de préstamos bancarios en moneda nacional; excluye tarjetas.',
                     'Si aumenta, el financiamiento se encarece; si baja, puede facilitar compras a crédito.',
                     'Cotizar el costo real de vender en cuotas y ajustar plazos según margen y mora.'),
                    ('pasiva','tasa_pasiva_pct','Tasa pasiva','%', 'Remuneración promedio ponderada de depósitos en moneda nacional; excluye depósitos a la vista.',
                     'Una subida mejora el incentivo a ahorrar; una baja reduce el rendimiento del depósito.',
                     'Comparar el retorno de mantener caja en depósitos con el de invertir en campañas.'),
                    ('spread','spread_pp','Margen bancario (spread)','pp', 'Diferencia entre la tasa activa y la tasa pasiva.',
                     'Un margen mayor implica mayor diferencia entre costo de préstamo y rendimiento de depósito. No representa por sí solo la utilidad del banco.',
                     'Negociar alternativas de financiamiento al evaluar el costo de campañas o inventario.')]
                for key,col,title,unit,definition,reading,decision in specs:
                    add(key,n,title,'Financiamiento',unit,'Mensual',d,d.fecha,d[col],definition,reading,
                        'Tasa activa y pasiva responden a liquidez, riesgo y política monetaria; spread = activa − pasiva.',
                        decision,'Archivos de tasas bancaria activa/pasiva; institución por verificar en la fuente original.',
                        method='Spread conservado del archivo; comprobación activa − pasiva en el panel de calidad.' if key=='spread' else 'Valor del archivo, sin estimación.')
                if not np.allclose(d.tasa_activa_pct - d.tasa_pasiva_pct, d.spread_pp, atol=0.011, equal_nan=True):
                    issues.append('El spread contiene diferencias frente a activa menos pasiva.')
            elif n == 5:
                for key,ind,currency,title,definition in [
                    ('base','Base_monetaria','Nacional','Base monetaria','Agregado de dinero primario reportado como base monetaria en el archivo.'),
                    ('m1','M1','Nacional','M1 · moneda nacional','Medio circulante en moneda nacional; dinero disponible para transacciones.'),
                    ('m2','M2','Nacional','M2 · moneda nacional','Medios de pago en moneda nacional, agregado más amplio que M1.'),
                    ('m2total','M2_total','Total','Medios de pago totales','Columna de medios de pago totales de la fuente, conservada como agregado independiente.')]:
                    a=d[(d.indicador==ind)&(d.moneda==currency)]
                    freq='Mensual' if key=='base' else 'Anual'
                    dates=a.fecha_referencia if key=='base' else annual_dates(a.anio)
                    add(key,n,title,'Liquidez','millones de Q',freq,a,dates,a.valor_millones_quetzales,
                        definition,'Una expansión indica más liquidez nominal; no equivale automáticamente a mayor poder adquisitivo ni a inflación. Comparar con precios y actividad.',
                        'La liquidez se relaciona con depósitos y crédito. Su efecto sobre precios depende también de la demanda de dinero y la producción.',
                        'Contrastar liquidez con ventas y cobros antes de ampliar inventarios o promociones.',
                        'Banco de Guatemala, según los archivos originales descritos en el diccionario.',
                        'Base mensual; M1/M2 anuales. No se suman estos agregados entre sí ni se interpolan frecuencias.')
            elif n == 6:
                months={'ene':1,'feb':2,'mar':3,'abr':4,'may':5,'jun':6,'jul':7,'ago':8,'sep':9,'oct':10,'nov':11,'dic':12}
                a=d[d.Mes.notna()].copy()
                a['mes_num']=a.Mes.astype(str).str.strip().str[:3].str.lower().map(months)
                # El año aparece al cierre de los primeros bloques. Hay además años
                # incompatibles con los cierres de diciembre en bloques posteriores.
                block=(a.mes_num==1).cumsum()
                end_year=a.Anio.where(a.mes_num==12).groupby(block).transform('max')
                candidate=end_year.fillna(a.Anio)
                conflict=a.Anio.notna() & candidate.notna() & a.Anio.ne(candidate)
                missing=a.Anio.isna().sum()
                issues.append(f'Crédito: {missing} años vacíos y {int(conflict.sum())} años incompatibles con el cierre de diciembre. Revisar el archivo original.')
                if credit_rebuild:
                    years=candidate
                    note='Cronología RECONSTRUIDA por bloques enero–diciembre usando el año de diciembre; último bloque con año explícito. Pendiente de validar contra la fuente. No son datos oficialmente verificados.'
                else:
                    # Excluir bloques incompletos o contradictorios, sin corregir años
                    # silenciosamente. La opción explícita permite explorar la hipótesis.
                    valid_block=(a.Anio.notna() & ~conflict).groupby(block).transform('all')
                    years=a.Anio.where(valid_block)
                    note='Se excluyen bloques con años vacíos o contradictorios. No hay crecimiento anual comparable en modo estricto; la opción lateral permite explorar una reconstrucción identificada.'
                dates=pd.to_datetime(dict(year=years,month=a.mes_num,day=1),errors='coerce')
                saldo=clean_series(a,dates,a.Total)
                growth=yoy_monthly(saldo)
                add('credito',n,'Crecimiento anual del crédito','Financiamiento','%','Mensual',growth,growth.fecha,growth.valor,
                    'Variación del total de préstamos por destino en moneda nacional frente al mismo mes del año anterior; cobertura del archivo, no todo el crédito multimoneda.',
                    'Una subida indica mayor saldo de crédito frente al año anterior; una baja indica desaceleración. El saldo no mide desembolsos nuevos.',
                    'Tasas más altas pueden frenar crédito y compras financiadas; también intervienen riesgo y demanda.',
                    'Evaluar ventas a crédito junto con mora y costo de fondeo antes de ampliar plazos.',
                    'Institución y unidad del saldo no consignadas en la hoja recibida.',note,
                    '(Saldo del mes / saldo del mismo mes del año anterior − 1) × 100. Sin interpolación.')
            elif n == 7:
                for key,col in [('cambio','TCR 1/'),('cambio_compra','Compra'),('cambio_venta','Venta')]:
                    title='Tipo de cambio · referencia' if key=='cambio' else 'Tipo de cambio · '+col.lower()
                    add(key,n,title,'Sector externo','Q por US$','Diaria',d,
                        pd.to_datetime(d.Fecha,format='%d/%m/%Y',errors='coerce'),d[col],
                        f'Quetzales por dólar de la serie de {col.lower()} del archivo.',
                        'Una baja significa apreciación del quetzal; una subida, depreciación. Referencia, compra y venta son series distintas y no se empalman.',
                        'La apreciación puede abaratar importaciones y reducir quetzales por remesas; la depreciación opera en sentido contrario, con transmisión variable.',
                        'Usar referencia para el contexto macroeconómico y solicitar una cotización bancaria vigente antes de comprar o vender divisas.',
                        'Banco de Guatemala, indicado en el pie del CSV.',
                        'El CSV cambia de compra/venta a TCR el 02/12/2006. No se unen las tres columnas: compra y venta terminan en 2006; TCR llega a noviembre de 2026.')
            elif n == 8:
                a=d[d.tipo_registro=='Mensual']
                add('remesas',n,'Remesas familiares','Sector externo','millones de US$','Mensual',a,a.fecha_referencia,a.monto_millones_usd,
                    'Ingreso mensual de divisas por remesas familiares.',
                    'Una subida puede sostener consumo de hogares receptores. Comparar el mismo mes de años distintos para evitar confundir estacionalidad con tendencia.',
                    'Aumentan oferta de divisas y recursos para consumo; su valor en quetzales también depende del tipo de cambio.',
                    'Programar campañas en zonas receptoras y validar el calendario con ventas reales.',
                    'Banco de Guatemala, según el diccionario del archivo.',
                    'Solo registros mensuales; los totales anuales y el acumulado parcial 2026 quedan fuera para evitar doble conteo.')
            elif n == 9:
                add('reservas',n,'Reservas internacionales netas','Sector externo','millones de US$','Mensual',d,d.fecha,d.reservas_millones_usd,
                    'Saldo de reservas monetarias internacionales netas reportado por el archivo.',
                    'Un aumento amplía el saldo de activos de respaldo externo. El nivel aislado no basta para evaluar suficiencia: harían falta importaciones y obligaciones externas.',
                    'Se relacionan con flujos de divisas y operaciones del banco central; no determinan por sí solas el tipo de cambio.',
                    'Incorporar el contexto externo en compras importadas sin tratar las reservas como garantía de un tipo de cambio fijo.',
                    'Banco de Guatemala, según el diccionario.', 'No se suman saldos mensuales para calcular un total anual.')
            elif n == 10:
                specs=[('wti','petroleo_wti_cushing','Petróleo WTI','USD por barril','Mensual'),
                       ('superior','gasolina_superior','Gasolina superior','Q por galón','Diaria'),
                       ('regular','gasolina_regular','Gasolina regular','Q por galón','Diaria'),
                       ('diesel','diesel','Diésel','Q por galón','Diaria')]
                for key,ind,title,unit,freq in specs:
                    a=d[d.indicador==ind]
                    source=' / '.join(a.fuente_institucional.dropna().unique().astype(str))
                    note='WTI internacional mensual.' if key=='wti' else 'Precio promedio al consumidor final en Ciudad Capital; NO identificado como precio en pista. Solo hasta octubre de 2024.'
                    add(key,n,title,'Energía',unit,freq,a,a.fecha,a.valor,
                        'Precio internacional spot FOB WTI en Cushing.' if key=='wti' else 'Precio promedio al consumidor final en Ciudad Capital, según la base.',
                        'Un aumento presiona costos de transporte y producción; una caída puede aliviarlos. No existe un umbral universal: evaluar el cambio y el costo propio.',
                        'Petróleo, combustibles e inflación se conectan mediante costos; impuestos, márgenes y tipo de cambio modifican la transmisión local.',
                        'Recalcular rutas, costos de entrega y márgenes antes de modificar precios o prometer envío gratuito.',source,note)
        except (KeyError,ValueError,TypeError,AttributeError) as exc:
            issues.append(f'{FILES[n]}: estructura no compatible o error de validación: {exc}')
    return result,issues


def period_text(ts,freq):
    if pd.isna(ts): return 'Sin fecha'
    return ts.strftime('%Y') if freq=='Anual' else ts.strftime('%d/%m/%Y') if freq=='Diaria' else ts.strftime('%m/%Y')


def num(value, digits=2):
    return 'No disponible' if pd.isna(value) else f'{value:,.{digits}f}'


def previous_comparison(item):
    d=item['data']
    if len(d)<2: return None,None
    last=d.iloc[-1]
    if item['freq']=='Anual':
        ref=d[d.fecha.dt.year==last.fecha.year-1]
    elif item['freq']=='Mensual':
        ref=d[d.fecha.dt.to_period('M')==last.fecha.to_period('M')-12]
    else:
        # Última cotización disponible en o antes del aniversario, hasta 7 días atrás.
        target=last.fecha-pd.DateOffset(years=1)
        ref=d[(d.fecha<=target)&(d.fecha>=target-pd.Timedelta(days=7))]
    if ref.empty:return None,None
    p=ref.iloc[-1]
    if item['unit'] in ('%','pp'):return last.valor-p.valor,'pp frente a un año antes'
    if p.valor==0:return None,None
    return (last.valor/p.valor-1)*100,'% frente a un año antes'


def signal(item):
    d=item['data']
    if d.empty:return '⚪ Dato no disponible'
    v=d.iloc[-1].valor
    if item['key']=='inflacion':
        return '🟢 Dentro de 3–5 %' if 3<=v<=5 else '🟠 Bajo la meta' if v<3 else '🔴 Sobre la meta'
    change,label=previous_comparison(item)
    if change is None:return '⚪ Sin comparación anual'
    return '↑ Sube interanual' if change>0 else '↓ Baja interanual' if change<0 else '→ Sin cambio interanual'


def chart(items,title,start_year=None,end_year=None,indexed=False):
    fig=go.Figure()
    has_data=False
    for i,item in enumerate(items):
        d=item['data'].copy()
        if start_year is not None:d=d[d.fecha.dt.year.between(start_year,end_year)]
        if d.empty:continue
        y=d.valor.copy()
        if indexed:
            if y.iloc[0]==0:continue
            y=y/y.iloc[0]*100
        has_data=True
        fig.add_trace(go.Scatter(x=d.fecha,y=y,name=item['title'],mode='lines+markers' if len(d)<70 else 'lines',
                                line=dict(color=COLORS[i%len(COLORS)],width=2.5),connectgaps=False,
                                hovertemplate='%{x|%d/%m/%Y}<br>%{y:,.2f}<extra>%{fullData.name}</extra>'))
    if len(items)==1 and items[0]['key']=='inflacion' and not indexed:
        fig.add_hrect(y0=3,y1=5,fillcolor='#12b8a6',opacity=.10,line_width=0)
        fig.add_hline(y=4,line_dash='dot',line_color='#12b8a6',annotation_text='Meta 4 % · consigna')
    fig.update_layout(title=dict(text=title,font=dict(size=17)),height=350,
                      paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',
                      font=dict(color='#59657c'),margin=dict(l=10,r=20,t=55,b=15),
                      xaxis_title='Año de referencia' if items and all(i['freq']=='Anual' for i in items) else 'Fecha / período de referencia',
                      yaxis_title='Índice (primera observación visible = 100)' if indexed else items[0]['unit'] if items else '',
                      hovermode='x unified',legend=dict(orientation='h',y=-.25),
                      xaxis=dict(showgrid=False),yaxis=dict(gridcolor='rgba(130,150,170,.15)'))
    return fig,has_data


def plot(items,title,key,years=None,indexed=False):
    fig,ok=chart(items,title,*(years or (None,None)),indexed=indexed)
    if ok:st.plotly_chart(fig,width='stretch',key=key,config={'displaylogo':False})
    else:st.info('No hay observaciones para este intervalo. Amplía el período o revisa la calidad del archivo.')


SCENARIOS={
    '01 · Sube la tasa líder': [
        ('Tasa líder','↑','Impulso de política monetaria restrictiva.'),
        ('Tasas activa y pasiva','Posible ↑','Transmisión con rezagos; depende de liquidez y competencia.'),
        ('Crédito y demanda','Posible ↓ en crecimiento','Financiar compras e inventario puede costar más.'),
        ('Inflación','Posible ↓ posterior','Menor presión de demanda, sin efecto inmediato garantizado.'),
        ('Tipo de cambio Q/US$','Posible ↓','Mayor atractivo de activos en Q; otros flujos pueden contrarrestarlo.')],
    '02 · Aumenta el petróleo': [
        ('WTI','↑','Choque internacional de costos.'),
        ('Combustibles locales','Posible ↑','Impuestos, márgenes y tipo de cambio alteran el traslado.'),
        ('Inflación y logística','Presión ↑','Aumentan costos directos e indirectos.'),
        ('Poder adquisitivo real','Presión ↓','El presupuesto alcanza para menos si el ingreso no se ajusta.'),
        ('Demanda de dólares','Posible ↑','Importadores necesitan más divisas para una misma cantidad.')],
    '03 · Se aprecia el quetzal': [
        ('Tipo de cambio Q/US$','↓','Se necesitan menos quetzales por dólar.'),
        ('Costo importado en Q','Posible ↓','Mismo precio en dólares y demás costos constantes.'),
        ('Remesas convertidas a Q','↓ para iguales US$','Menos ingreso local por cada dólar recibido.'),
        ('Ingresos de exportación en Q','Presión ↓','Para iguales ventas en dólares.'),
        ('Inflación importada','Posible ↓','Traslado parcial y con rezagos a los precios locales.')],
}
ACTIONS={
    '01 · Sube la tasa líder': 'Negociar fondeo y plazos; calcular el margen de promociones a cuotas; segmentar descuentos y cuidar el cobro antes de ampliar crédito a clientes.',
    '02 · Aumenta el petróleo': 'Actualizar costo por entrega, consolidar rutas, revisar inventario de insumos y ajustar precios por categoría según margen y sensibilidad del cliente.',
    '03 · Se aprecia el quetzal': 'Recotizar compras importadas y evaluar trasladar parte del ahorro a precios o promociones; revisar márgenes si la empresa exporta o vende a hogares receptores de remesas.',
}


def main():
    st.set_page_config(page_title='Pulso Macro · Guatemala',page_icon='◈',layout='wide')
    st.markdown('''<style>
    .stApp{background:linear-gradient(135deg,#f7f9fc 0%,#edf3f8 100%);color:#182b49}
    [data-testid="stSidebar"]{background:#e6edf5}
    .block-container{max-width:1440px;padding-top:2rem;padding-bottom:2rem}
    h1,h2,h3{color:#182b49!important;letter-spacing:-.025em}
    [data-testid="stMetric"]{background:#fff;border:1px solid #dce5ef;border-radius:16px;padding:16px}
    [data-testid="stMetricLabel"]{color:#53637a!important}
    [data-testid="stMetricValue"]{color:#182b49!important;font-size:1.85rem}
    .hero{background:linear-gradient(115deg,#142947,#204b66);padding:28px 32px;border-radius:22px;margin-bottom:22px;color:#fff}
    .hero h1{color:#fff!important;font-size:2.5rem;margin:0}
    .hero p{color:#cfdee9;margin-bottom:0}
    .eyebrow{color:#72e1d2;font-size:.78rem;letter-spacing:.18em;font-weight:700}
    .stTabs [data-baseweb="tab-list"]{gap:16px}
    .stTabs [data-baseweb="tab"]{font-weight:650}
    [data-testid="stAlert"]{border-radius:12px}
    </style>''',unsafe_allow_html=True)
    st.markdown('<div class="hero"><div class="eyebrow">GUATEMALA / INTELIGENCIA PARA DECIDIR</div><h1>Pulso Macro</h1><p>De las variables económicas a las decisiones de mercadeo.</p></div>',unsafe_allow_html=True)

    with st.sidebar:
        st.markdown('### ◈ Centro de control')
        root=st.text_input('Carpeta de datos',value=str(Path(__file__).resolve().parent))
        uploads=st.file_uploader('Cargar archivos originales',type=['csv','xls','xlsx'],accept_multiple_files=True)
        st.caption('Nombres Variable1…Variable10. La carga manual reemplaza el archivo local de la misma variable.')
        rebuild=st.checkbox('Explorar reconstrucción de fechas del crédito',value=False,
                            help='El archivo tiene años vacíos y contradictorios. Reconstruye bloques con el año de diciembre. Esta opción exige validación de la cronología original.')
    blobs={}
    for folder in [Path(root),Path(root)/'data',Path(root)/'upload']:
        if folder.is_dir():
            for p in sorted(folder.iterdir()):
                n=file_number(p.name)
                if n and p.suffix.lower() in ('.csv','.xls','.xlsx') and n not in blobs:
                    blobs[n]=(p.name,p.read_bytes())
    for upload in uploads or []:
        n=file_number(upload.name)
        if n:blobs[n]=(upload.name,upload.getvalue())
    if not blobs:
        st.info('Coloca tus diez archivos junto al script o cárgalos desde la barra lateral para comenzar.')
        with st.expander('Archivos esperados',expanded=True):st.write(list(FILES.values()))
        st.stop()
    books,read_errors={},[]
    for n,(name,content) in sorted(blobs.items()):
        try:books[n]=read_file(name,content)
        except Exception as exc:read_errors.append(f'{name}: {exc}')
    catalog,issues=build_catalog(books,rebuild)
    issues=read_errors+issues
    if not catalog:st.error('No se pudo construir ninguna serie. Revisa nombres y columnas.');st.stop()
    missing=sorted(set(FILES)-set(books))
    if missing:st.warning('Faltan archivos de las variables: '+', '.join(map(str,missing)))
    nonempty=[s for s in catalog.values() if not s['data'].empty]
    if not nonempty:st.error('Los archivos no contienen observaciones utilizables.');st.stop()
    lo=min(s['data'].fecha.dt.year.min() for s in nonempty)
    hi=max(s['data'].fecha.dt.year.max() for s in nonempty)
    with st.sidebar:
        years=st.slider('Intervalo de los gráficos',int(lo),int(hi),(max(int(lo),int(hi)-5),int(hi))) if lo<hi else (int(lo),int(hi))
        st.caption('El resumen y las fichas muestran el último dato de cada archivo. El intervalo afecta los gráficos, no la fecha del último dato.')
        st.divider()
        st.caption(f'{len(books)}/10 archivos · {len(catalog)} series de consulta')
        with st.expander('Completar trazabilidad'):
            st.caption('Opcional: registra enlaces oficiales y fechas de descarga reales. No se infieren de la fecha de carga.')
            source_n=st.selectbox('Archivo',list(books),format_func=lambda n:FILES[n],key='source_n')
            st.text_input('Institución confirmada',key=f'institution_{source_n}')
            st.text_input('Enlace a la fuente original',value=DEFAULT_URLS.get(source_n,''),key=f'url_{source_n}')
            st.text_input('Fecha real de descarga (AAAA-MM-DD)',key=f'download_{source_n}')
    st.caption(f'Cobertura: diez familias obligatorias y {len(catalog)} series de consulta. Las subseries no equivalen necesariamente a veinte variables distintas del curso.')
    panel,fichas=st.tabs(['01 · Panel de decisión','02 · Fichas y datos'])
    with panel:
        st.subheader('El pulso en cuatro cifras')
        for col,key in zip(st.columns(4),['pib','inflacion','lider','cambio']):
            with col:
                if key in catalog:
                    s=catalog[key];d=s['data']
                    if not d.empty:
                        change,label=previous_comparison(s)
                        st.metric(s['title'],num(d.iloc[-1].valor)+' '+s['unit'],
                                  None if change is None else f'{change:+.2f} {label}',delta_color='off')
                        st.caption(period_text(d.iloc[-1].fecha,s['freq']))
        st.markdown('#### Panorama de variables')
        rows=[]
        for s in catalog.values():
            d=s['data'];last=d.iloc[-1] if not d.empty else None
            inst=st.session_state.get(f'institution_{s["n"]}','').strip() or s['source']
            rows.append({'Variable':s['title'],'Último dato':num(last.valor) if last is not None else 'No disponible',
                         'Unidad':s['unit'],'Período':period_text(last.fecha,s['freq']) if last is not None else 'Sin fecha comparable',
                         'Señal':signal(s),'Fuente declarada':inst,
                         'Cobertura':s['freq'],'Observación':s['note']})
        summary=pd.DataFrame(rows)
        st.dataframe(summary,hide_index=True,width="stretch")
        st.caption('Solo inflación tiene rango normativo (3–5 %, consigna). Las flechas describen cambios interanuales, sin calificarlos como buenos o malos. Cada serie conserva su propio período.')
        st.download_button('↓ Descargar resumen CSV',summary.to_csv(index=False).encode('utf-8-sig'),'resumen_macro.csv','text/csv')
        c1,c2=st.columns(2)
        with c1:
            if 'inflacion' in catalog:plot([catalog['inflacion']],'Inflación frente al rango de la consigna','overview_inflation',years)
        with c2:
            rates=[catalog[k] for k in ['lider','activa','pasiva'] if k in catalog]
            if rates:plot(rates,'El costo del dinero','overview_rates',years)
        st.markdown('#### Tres escenarios para actuar')
        scenario=st.selectbox('¿Qué cambia en el entorno?',list(SCENARIOS))
        st.dataframe(pd.DataFrame(SCENARIOS[scenario],columns=['Variable','Dirección esperada','Mecanismo']),hide_index=True,width="stretch")
        st.success('**Decisión de mercadeo:** '+ACTIONS[scenario])
        st.caption('Escenarios cualitativos resueltos: no son eventos confirmados ni estimaciones econométricas. Las demás condiciones pueden modificar los efectos.')
        with st.expander('Laboratorio: efecto directo sobre un presupuesto'):
            st.caption('Cálculo condicional con supuestos ingresados por el usuario; no crea datos históricos ni predice variables macroeconómicas.')
            a,b=st.columns(2)
            with a:amount=st.number_input('Factura de importación (US$)',min_value=0.,value=10000.,step=100.)
            fx=catalog.get('cambio',{}).get('data',pd.DataFrame())
            if not fx.empty:
                baseline=float(fx.iloc[-1].valor)
                with b:appreciation=st.slider('Apreciación hipotética del quetzal (%)',0.,20.,5.,.5)
                hypothetical=baseline*(1-appreciation/100)
                a.metric('Costo al último tipo de referencia',f'Q {amount*baseline:,.2f}')
                b.metric('Costo con el supuesto',f'Q {amount*hypothetical:,.2f}',f'Ahorro Q {amount*(baseline-hypothetical):,.2f}',delta_color='off')
                st.caption(f'Base: {baseline:.5f} Q/US$, {period_text(fx.iloc[-1].fecha,"Diaria")}. Referencia macroeconómica, no cotización efectiva de venta bancaria. Fórmula: US$ × Q/US$. No incluye comisiones, fletes, impuestos ni cambios del precio en dólares.')
    with fichas:
        st.subheader('Una variable, una decisión')
        a,b=st.columns([1,3])
        groups=['Todas']+list(dict.fromkeys(s['group'] for s in catalog.values()))
        with a:group=st.selectbox('Familia',groups)
        options=[k for k,s in catalog.items() if group=='Todas' or s['group']==group]
        with b:key=st.selectbox('Variable para consultar',options,format_func=lambda k:catalog[k]['title'])
        s=catalog[key];d=s['data']
        st.markdown('#### 1 · Nombre y definición')
        st.write(f'**{s["title"]}** · {s["definition"]}')
        st.caption(f'Unidad: {s["unit"]} · Frecuencia: {s["freq"]}.')
        st.markdown('#### 2 · Último dato y fuente')
        if d.empty:st.warning('Dato no disponible con fechas comparables en el archivo. Revisa las observaciones de calidad.')
        else:
            a,b,c=st.columns([1,1,2])
            a.metric('Último dato en el archivo',num(d.iloc[-1].valor)+' '+s['unit'])
            b.metric('Período',period_text(d.iloc[-1].fecha,s['freq']))
            c.info(signal(s))
        inst=st.session_state.get(f'institution_{s["n"]}','').strip() or s['source']
        st.write('**Fuente declarada:** '+inst)
        st.caption('Archivo recibido: '+blobs[s['n']][0])
        # El enlace EIA corresponde solo a WTI, nunca a precios locales MEM.
        url=st.session_state.get(f'url_{s["n"]}',DEFAULT_URLS.get(s['n'],''))
        if s['n']==10 and key!='wti' and url==DEFAULT_URLS[10]:url=''
        if url.startswith(('https://','http://')):st.markdown(f'[Abrir enlace registrado]({url})')
        else:st.caption('Enlace oficial específico: no disponible en el archivo; pendiente de completar.')
        downloaded=st.session_state.get(f'download_{s["n"]}','').strip()
        st.caption('Fecha de descarga original: '+(downloaded or 'no consta en los archivos.'))
        st.caption('Actualidad: último dato de la base recibida; no se ha contrastado con publicaciones posteriores.')
        if s['note']:st.info(s['note'])
        st.markdown('#### 3 · Serie histórica')
        if not d.empty:
            st.caption(f'Cobertura disponible: {period_text(d.iloc[0].fecha,s["freq"])} → {period_text(d.iloc[-1].fecha,s["freq"])} · {len(d):,} observaciones.')
            if (d.iloc[-1].fecha-d.iloc[0].fecha).days<365*5:st.caption('El archivo aporta menos de cinco años; se muestra únicamente lo disponible.')
        full=st.checkbox('Mostrar toda la historia disponible',value=False,key='full_history')
        plot([s],s['title']+' · '+s['unit'],'detail_chart',None if full else years)
        st.markdown('#### 4 · Cómo leerlo');st.write(s['reading'])
        st.markdown('#### 5 · Relaciones');st.write(s['relations'])
        st.markdown('#### 6 · Decisión de mercadeo');st.success(s['decision'])
        st.caption('Método: '+s['method'])
        with st.expander('Comparar trayectorias (índice base 100)'):
            selected=st.multiselect('Hasta cuatro series',list(catalog),default=[key],format_func=lambda k:catalog[k]['title'],max_selections=4)
            selected_items=[catalog[k] for k in selected]
            st.caption('Cada serie inicia en 100 en su primera observación visible. Fechas iniciales y frecuencias pueden diferir; el índice no hace equivalentes las unidades ni demuestra causalidad.')
            if selected_items:plot(selected_items,'Evolución relativa, sin interpolación','comparison',years,indexed=True)
        with st.expander('Ver y descargar observaciones de la ficha'):
            export=d.copy()
            export['unidad']=s['unit'];export['variable']=s['title'];export['archivo']=blobs[s['n']][0]
            st.dataframe(export,hide_index=True,width="stretch")
            st.download_button('↓ Descargar serie CSV',export.to_csv(index=False).encode('utf-8-sig'),key+'.csv','text/csv')
        with st.expander('Explorar archivos originales y diccionarios'):
            raw_n=st.selectbox('Archivo recibido',list(books),format_func=lambda n:blobs[n][0])
            sheet=st.selectbox('Hoja',list(books[raw_n]))
            raw=books[raw_n][sheet]
            st.caption('Incluye todas las filas y columnas originales; los filtros de gráficos no modifican esta vista.')
            st.dataframe(raw,hide_index=True,width="stretch")
            st.download_button('↓ Exportar hoja original a CSV',raw.to_csv(index=False).encode('utf-8-sig'),f'variable_{raw_n}_{sheet}.csv','text/csv')



if __name__=='__main__':
    main()
