import pandas as pd
from datetime import date
import sys
import os
from dotenv import load_dotenv
from sqlalchemy import create_engine
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

sys.path.append(r"C:\Users\Admin\OneDrive\Documentos\mis_funciones")
from etl_funciones import(
    cargar_excel,
    validar_calidad,
    convertir_numero_texto,
    limpiar_dataframe,
    cargar_postgresql,
    identificar_inexistentes,
    ejecutar_query,
    exportar_excel,
    enviar_email
)

#CONECTAR POSTGRESQL
load_dotenv()

cargar_postgresql(
    database=os.getenv("DB_NAME"),
    user=os.getenv("DB_USER"),
    password=os.getenv("DB_PASSWORD"),
)

db_url = (
    f"postgresql://{os.getenv('DB_USER')}:"
    f"{os.getenv('DB_PASSWORD')}@"
    f"{os.getenv('DB_HOST')}:"
    f"{os.getenv('DB_PORT')}/"
    f"{os.getenv('DB_NAME')}"
)

engine = create_engine(db_url)

#CARGAR A EXCEL
df_clientes = cargar_excel("farmacia_clientes.xlsx")
df_productos = cargar_excel("farmacia_productos.xlsx")
df_vendedores = cargar_excel("farmacia_vendedores.xlsx")
df_ventas = cargar_excel("farmacia_ventas.xlsx")

print("carga completada")

#VALIDAR CALIDAD 
validar_productos = validar_calidad(df_productos, "Precio_unitario")
validar_clientes = validar_calidad(df_clientes) 
validar_ventas = validar_calidad(df_ventas, "Cantidad")
validar_ventas_precio = validar_calidad(df_ventas, "Precio_unitario")
validar_vendedores = validar_calidad(df_vendedores)

print(type(df_ventas))

print("\nNULOS")
print(validar_clientes["nulos"])

print("\nDUPLICADOS")
print(validar_clientes["duplicados"])

print("\nNULOS")
print(validar_vendedores["nulos"])

print("\nDUPLICADOS")
print(validar_vendedores["duplicados"])

print("\nNULOS")
print(validar_productos["nulos"])

print("\nDUPLICADOS")
print(validar_productos["duplicados"])

print("VALORES INVALIDOS (Precio_unitario)")
print(validar_productos["valores_invalidos"])

print("\nNULOS")
print(validar_ventas["nulos"])

print("\nDUPLICADOS")
print(validar_ventas["duplicados"])

print("VALORES INVALIDOS (Cantidad)")
print(validar_ventas["valores_invalidos"])

print("VALORES INVALIDOS (Precio_unitario)")
print(validar_ventas_precio["valores_invalidos"])

#EXPLORAR DATOS (antes de limpiar): Busco valores raros que validar calidad no detecto
print(df_ventas["Estado"].unique())#¿hay el mismo estado escrito de varias formas?
print(df_ventas[["Cantidad", "Precio_unitario","Total"]].describe()) #¿min y max tienen sentido para una farmacia?
print(df_ventas["Id_venta"].duplicated().sum()) #¿se repite la clave aunque la fila no sea idéntica?
pd.set_option("display.max_columns", None)
print(df_ventas[df_ventas["Id_venta"].duplicated(keep=False)].sort_values("Id_venta"))# Ver las filas completas con Id_venta repetido (keep=False muestra todas las copias)
print(df_ventas.duplicated().sum())
print(df_clientes["Tipo_cliente"].unique()) #¿hay categorías escritas distinto?


# Ver las filas de ventas con Cantidad nula, para decidir si se borran o se reconstruyen
print(df_ventas[df_ventas["Cantidad"].isna()])


# --- Nulos en ventas: Cantidad ---
# Hay 3 ventas sin Cantidad pero con Total y Precio_unitario válidos.
# En vez de borrarlas (se perdería ingreso en los reportes), se reconstruye:
# Cantidad = Total / Precio_unitario (válido porque Descuento = 0 y precio > 0)
mascara = df_ventas["Cantidad"].isna()
df_ventas.loc[mascara, "Cantidad"] = df_ventas.loc[mascara, "Total"] / df_ventas.loc[mascara, "Precio_unitario"]

# Verificación: deben salir 4.0, 5.0, 6.0 y 0 nulos restantes
print(df_ventas.loc[mascara, "Cantidad"])
print(df_ventas["Cantidad"].isna().sum())

#Limpiar nulos 
df_clientes["Ciudad"] = df_clientes["Ciudad"].fillna("Sin ciudad")
print(df_clientes["Ciudad"].isna().sum())

df_productos["Categoria"] = df_productos["Categoria"].fillna("Sin categoria")
print(df_productos["Categoria"].isna().sum())

#Convertir numero a texto (telefono )
df_clientes = convertir_numero_texto(df_clientes, ["Telefono"])
print(df_clientes["Telefono"].isna().sum())
print(df_clientes["Telefono"].unique())

#Rellenar telefono con sin telefono
df_clientes["Telefono"] = df_clientes["Telefono"].fillna("Sin telefono")
print(df_clientes["Telefono"].isna().sum())
print(df_clientes["Telefono"].unique())

#quitqr duplicaodos de ventas 
print(len(df_ventas))
df_ventas = df_ventas.drop_duplicates() 
print(len(df_ventas))
print(df_ventas["Id_venta"].duplicated().sum())

#limpiar dataframe 
df_clientes = limpiar_dataframe(
    df_clientes,
    columnas_texto=["Nombre_cliente", "Tipo_cliente", "Ciudad", "Telefono"]
)
df_productos = limpiar_dataframe(
    df_productos,
    columnas_texto=["Nombre_producto", "Categoria"]
)
df_vendedores = limpiar_dataframe(
    df_vendedores,
    columnas_texto=["Nombre_vendedor", "Zona"]
)
df_ventas = limpiar_dataframe(
    df_ventas,
    columnas_texto=["Estado"],
    columnas_fecha=["Fecha"],
    formato_latino=False
)

print(df_ventas[df_ventas["Id_venta"] == "F00101"])
print(df_ventas["Fecha"].min(), df_ventas["Fecha"].max())

print(df_clientes["Id_cliente"].duplicated().sum())
print(df_productos["Id_producto"].duplicated().sum())

print(df_ventas[(df_ventas["Cantidad"] <= 0) | (df_ventas["Cantidad"] > 100)])
print(df_ventas[df_ventas["Precio_unitario"] == 0])

# Corrección de 3 ventas con datos erróneos; el Total es coherente, así que se despeja la columna mala
df_ventas.loc[df_ventas["Id_venta"] == "F01201", "Cantidad"] = 3      # era -4; 35.22 / 11.74 = 3
df_ventas.loc[df_ventas["Id_venta"] == "F00301", "Cantidad"] = 5      # era 900; 338.87 / (71.34 * 0.95) = 5
df_ventas.loc[df_ventas["Id_venta"] == "F01801", "Precio_unitario"] = 72.17  # era 0; 505.19 / 7 = 72.17

print(df_ventas[(df_ventas["Cantidad"] <= 0) | (df_ventas["Cantidad"] > 100) | (df_ventas["Precio_unitario"] == 0)])

print(df_productos[df_productos["Id_producto"] == "P017"])

#pasar a entero 
df_ventas["Cantidad"] = df_ventas["Cantidad"].astype(int)
print(df_ventas["Cantidad"].dtypes)

#cargar a postgresql()
df_clientes.to_sql("clientes", engine, index=False, if_exists="replace", )
df_productos.to_sql("productos", engine, index=False, if_exists="replace", )
df_vendedores.to_sql("vendedores", engine, index=False, if_exists="replace", )
df_ventas.to_sql("ventas", engine, index=False, if_exists="replace", )

print("tabla cargada a postgresql")

#Identificar inexistentes
clientes_inexistentes = identificar_inexistentes(df_ventas, df_clientes, "Id_cliente")
print(clientes_inexistentes)

productos_inexistentes = identificar_inexistentes(df_ventas, df_productos, "Id_producto")
print(productos_inexistentes)

vendedores_inexistentes = identificar_inexistentes(df_ventas, df_vendedores, "Id_vendedor")
print(vendedores_inexistentes)

clientes_inexistentes = clientes_inexistentes.assign(Motivo="Cliente inexistente")

productos_inexistentes = productos_inexistentes.assign(Motivo="Producto inexistente")

vendedores_inexistentes = vendedores_inexistentes.assign(Motivo="Vendedor inexistente")


inconsistencia = pd.concat([clientes_inexistentes, productos_inexistentes, vendedores_inexistentes])
print(inconsistencia)

#quitar precio_unitario en ventas 
df_productos = df_productos.drop(columns=["Precio_unitario"])
print(df_productos.columns.tolist())

#unir tablas merge 
df_paso1 = df_ventas.merge(df_productos, on="Id_producto", how="inner")
print(len(df_paso1))

df_combinado = df_paso1.merge(df_clientes, on="Id_cliente", how="inner")
print(len(df_combinado))

df_final = df_combinado.merge(df_vendedores, on="Id_vendedor", how="inner")
print(len(df_final))

#Cargar a postgresql
df_final.to_sql("venta_drogueria", engine, index=False, if_exists="replace")

print(df_final.columns.tolist())

#ejecutar query
#QUERY 1 ventan por categoria 
df_categoria = pd.read_sql("""
    SELECT "Categoria",
        COUNT(*) AS "Cantidad_ventas",
        SUM("Cantidad") AS "Unidades_vendidas",
        ROUND(SUM("Total")::numeric, 2) AS "Ingresos_totales"
    FROM venta_drogueria
    WHERE "Estado" = 'Pagada'
    GROUP BY "Categoria"
    ORDER by "Ingresos_totales" DESC
""",engine)

print("\nventas por categoria")
print(df_categoria)

#QUERY 2 top 5 productos 
df_productos_top5 = pd.read_sql("""
    SELECT "Nombre_producto", "Categoria",
        SUM("Cantidad") AS "Unidades_vendidas",
        ROUND(SUM("Total")::numeric, 2) AS "Ingresos_totales"
    FROM venta_drogueria
    WHERE "Estado" = 'Pagada'
    GROUP BY "Nombre_producto", "Categoria"
    ORDER BY "Unidades_vendidas" DESC
    LIMIT 5
""",engine)

print("\ntop 5 productos")
print(df_productos_top5)

#QUERY 3 Cliente que mas compro
df_top_clientes = pd.read_sql("""
    SELECT "Nombre_cliente",
        COUNT(*) AS "Cantidad_compras",
        ROUND(SUM("Total")::numeric, 2) AS "Totales_comprados"
    FROM venta_drogueria
    WHERE "Estado" = 'Pagada'
    GROUP BY "Nombre_cliente"
    ORDER BY "Totales_comprados" DESC
    limit 10
""",engine)

print("\ncliente que mas compro")
print(df_top_clientes)

#QUERY 4 Ciudad con mas ventas 
df_ciudad = pd.read_sql("""
    SELECT "Ciudad",
        COUNT(*) AS "Cantidad_ventas",
        ROUND(SUM("Total")::numeric, 2) AS "Ingresos_totales"
    FROM venta_drogueria
    WHERE "Estado" = 'Pagada'
    GROUP BY "Ciudad"
    ORDER BY "Ingresos_totales" DESC
""",engine)

print("\nciudad con mas ventas")
print(df_ciudad)

#QUERY 5 ventas por vendedor 
df_vendedor = pd.read_sql("""
    SELECT "Nombre_vendedor",
        COUNT(*) AS "Cantidad_ventas",
        ROUND(SUM("Total")::numeric, 2) AS "Ingresos_totales"
    FROM venta_drogueria
    WHERE "Estado" = 'Pagada'
    GROUP BY "Nombre_vendedor"
    ORDER BY "Ingresos_totales" DESC
""",engine)

print("\nventas por vendedor")
print(df_vendedor)

#QUERY 6 venta mes a mes 
df_crecimiento_mes = pd.read_sql("""
    SELECT
        TO_CHAR("Fecha" :: date, 'YYYY-MM') AS "Mes",
        ROUND(SUM("Total")::numeric, 2) AS "Ingresos_totales",
        LAG(ROUND(SUM("Total")::numeric, 2)) OVER (
            ORDER BY TO_CHAR("Fecha" :: date, 'YYYY-MM')) AS "Mes_anterior",
        ROUND(SUM("Total")::numeric, 2) - LAG(ROUND(SUM("Total")::numeric, 2)) OVER (
            ORDER BY TO_CHAR("Fecha" :: date, 'YYYY-MM')) AS "Diferencia"
    FROM venta_drogueria
    WHERE "Estado" = 'Pagada'
    GROUP BY TO_CHAR("Fecha" :: date, 'YYYY-MM')
    ORDER BY TO_CHAR("Fecha" :: date, 'YYYY-MM')
""",engine)

print("\nventa mes a mes")
print(df_crecimiento_mes)

#QUERY 7 Ventas por tipo de clientes 
df_tipo_cliente = pd.read_sql("""
    SELECT "Tipo_cliente",
        COUNT(*) AS "Cantidad_compras",
        ROUND(SUM("Total")::numeric, 2 ) AS "Totales_comprados"
    FROM venta_drogueria
    WHERE "Estado" = 'Pagada'
    GROUP BY "Tipo_cliente"
    ORDER BY "Totales_comprados" DESC
""", engine)

print("\nTipo cliente")
print(df_tipo_cliente)

#QUERY 8 Comicion por vendedor
df_comision = pd.read_sql("""
    SELECT "Nombre_vendedor", 
        ROUND(SUM("Total")::numeric, 2) AS "Ingresos_totales",
        "Comision_pct",
        ROUND((SUM("Total") * "Comision_pct") ::numeric, 2) AS "Comision"
    FROM venta_drogueria
    WHERE "Estado" = 'Pagada'
    GROUP BY "Nombre_vendedor", "Comision_pct"
    ORDER BY "Comision" DESC
""", engine)

print("\nComision por vendedor")
print(df_comision)

#Exportar excel
df_final["Fecha"] =df_final["Fecha"].dt.date
inconsistencia["Fecha"] = inconsistencia["Fecha"].dt.date
hoy = date.today().strftime("%Y%m%d")

exportar_excel(f"reporte_venta_drogueria_{hoy}.xlsx", {
    "datos_completos": df_final,
    "top_5_productos": df_productos_top5,
    "cliente_que_mas_compro": df_top_clientes,
    "ciudad_con_mas_ventas": df_ciudad,
    "ventas_por_vendedor": df_vendedor,
    "venta_mes_a_mes": df_crecimiento_mes,
    "ventas_por_tipo_clientes": df_tipo_cliente,
    "comision_por_vendedor": df_comision,
    "inconsistencias": inconsistencia
})

#Formato profesional 
wb = load_workbook(f"reporte_venta_drogueria_{hoy}.xlsx")

azul_oscuro = "1F4E79"
azul_medio  = "2E75B6"
azul_claro  = "D6E4F0"
blanco      = "FFFFFF"

borde_fino = Border(
    left=Side(style='thin', color='CCCCCC'),
    right=Side(style='thin', color='CCCCCC'),
    top=Side(style='thin', color='CCCCCC'),
    bottom=Side(style='thin', color='CCCCCC')
)

for nombre_hoja in wb.sheetnames:
    ws = wb[nombre_hoja]

    #Encabezados fila 1
    for cell in ws["1"]:
        cell.font      = Font(name="Arial", size=9, bold=True, color=blanco)
        cell.fill      = PatternFill("solid", fgColor=azul_oscuro)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = borde_fino
    ws.row_dimensions[1].height = 35

    #Filas datos
    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        color = azul_claro if i % 2 == 0 else blanco
        for j, cell in enumerate(row, start=1):
            cell.fill   = PatternFill("solid", fgColor=color)
            cell.border = borde_fino
            cell.font   = Font(name="Arial", size=9)
            if j == 1:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.font = Font(name="Arial", size=9, bold=True)
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if isinstance(cell.value, int):
                    cell.number_format = '#,##0'
                
                elif isinstance(cell.value, float):
                    cell.number_format = '#,##0.00'
        ws.row_dimensions[i].height = 18
    
    #Ancho automatico 
    for col in ws.iter_cols():
        max_len = 0
        col_letter = None
        for cell in col:
            if hasattr(cell, 'column_letter'):
                col_letter = cell.column_letter
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
        if col_letter:
            ws.column_dimensions[col_letter].width = min(max_len + 2, 15)

     # Ajustar ancho según encabezado
    for cell in ws[1]:
        if cell.value:
            col_letter = cell.column_letter
            ancho_encabezado = len(str(cell.value)) + 2
            ancho_actual = ws.column_dimensions[col_letter].width
            ws.column_dimensions[col_letter].width = max(ancho_actual, ancho_encabezado)

wb.save(f"reporte_venta_drogueria_{hoy}.xlsx")
print("Formato aplicado correctamente")

#Enviar email
enviar_email(
    destinatario=os.getenv("EMAIL_USER"),
    asunto=f"Reporte_venta_drogueria_{hoy}",
    mensaje="Adjunto el reporte de ventas actualizado",
    archivo=f"reporte_venta_drogueria_{hoy}.xlsx",
    user=os.getenv("EMAIL_USER"),
    password=os.getenv("EMAIL_PASSWORD")
)
































