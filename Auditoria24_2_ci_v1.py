import os
import math
import pandas as pd
import numpy as np
from docx import Document
from openpyxl import load_workbook
from openpyxl.drawing.image import Image


file_path = '/content/PIAM_UNICAUCA_24_2.xlsx'
output_pathXlsx = '/content/AuditoriaPiam20242CiV.xlsx'

def agregar_mensaje(doc, mensaje):
    print(mensaje)
    doc.add_paragraph(mensaje)

def cargar_archivos_y_dataframes(file_path):
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"{file_path} no encontrado.")
    print(f"Archivo {file_path} encontrado.")
    try:
        dic_insumos = pd.read_excel(file_path, sheet_name=['24_2_VAL_211024','PIAM20242_AJ2V','ICTEX_ANGELA', 'ICTEX_CULTURA', 'SQ_24_2_221024','SQPG_24_2_22102024'], engine='openpyxl')
        for df in dic_insumos.values():
            df.columns = df.columns.str.strip()
        return dic_insumos['24_2_VAL_211024'], dic_insumos['PIAM20242_AJ2V'], dic_insumos['ICTEX_ANGELA'], dic_insumos['ICTEX_CULTURA'], dic_insumos['SQ_24_2_221024'], dic_insumos['SQPG_24_2_22102024']
    except Exception as e:
        raise Exception(f"Error al cargar los DataFrames: {e}")


def depuradorIcetex(icetexAng, icetexCul):
    filtroIcetexAng = icetexAng[['TERCERO', 'ID FACTURA', 'NUMERO', 'VALOR DEL GIRO ICETEX','VALOR PAGO FACTURA APLICADO', 'SALDO A FAVOR', 'MERITO UNICAUCA']]
    filtroIcetexCul = icetexCul[['Documento', 'Codigo', 'Sublínea Crédito', 'Relación de Giro', 'Total a Girar']]
    filtroIcetexCul = filtroIcetexCul.rename(columns={'Documento': 'TERCERO'})
    insumoIcetex = pd.merge(filtroIcetexAng, filtroIcetexCul, on='TERCERO', how='inner')
    insumoIcetex['Observacion'] = insumoIcetex.apply(
        lambda row: 'Valor giro distinto' if row['VALOR DEL GIRO ICETEX'] != row['Total a Girar'] else '',
        axis=1
    )
    return insumoIcetex

def depuradorFacturacion(Fact20242, FactPol20242):
    Fact20242['Documento'] = pd.to_numeric(Fact20242['Documento'].astype(str).str.strip(), errors='coerce')
    FactPol20242['Documento'] = pd.to_numeric(FactPol20242['Documento'].astype(str).str.strip(), errors='coerce')
    filtroFactPol = FactPol20242[['Documento', 'Cuenta bancaria', 'Aplica gratuidad']]
    insumoFacturacion = pd.merge(Fact20242, filtroFactPol, on='Documento', how='left')
    insumoFacturacion['Observacion'] = insumoFacturacion.apply(
        lambda row: 'Registro solo en Fact20242' if pd.isna(row['Aplica gratuidad']) else '',
        axis=1
    )
    soloFactPol = pd.merge(Fact20242[['Documento']], filtroFactPol, on='Documento', how='right', indicator=True)
    soloFactPol = soloFactPol[soloFactPol['_merge'] == 'right_only'].drop(columns='_merge')
    soloFactPol['Observacion'] = 'Registro solo en FactPol20242'
    facturacionFinal = pd.concat([insumoFacturacion, soloFactPol], ignore_index=True)
    return facturacionFinal

# Extraccion
piam20242civ, piam20242ci, icetexAng, icetexCul, Fact20242, FactPol20242 = cargar_archivos_y_dataframes(file_path)

# Manipulación
try:
    insumoIcetex = depuradorIcetex(icetexAng, icetexCul)
    insumoFacturacion = depuradorFacturacion(Fact20242, FactPol20242)
    print("Los DataFrames han sido procesados correctamente.")
except KeyError as e:
    print(f"Error de clave: {e}")
    print("Revisa que todas las columnas necesarias estén presentes en los DataFrames.")

# Carga
with pd.ExcelWriter(output_pathXlsx, engine='xlsxwriter') as writer:

    if insumoIcetex is not None:
      insumoIcetex.to_excel(writer, sheet_name='Icetex', index=False)
      print('Se ha generado la plantilla de caracterizacion')
    if insumoFacturacion is not None:
      insumoFacturacion.to_excel(writer, sheet_name='Facturacion', index=False)
      print('Se ha generado la plantilla de facturación')

print("Los resultados han sido guardados en el documento y archivo Excel.")
