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
    FactPol20242 = FactPol20242.rename(columns={
        'Id Factura':'Id  factura',
        'Identificacion tercero':'Tercero',
        'Nombre Tercero':'Nombre del Tercero',
        'Valor':'Valor Factura',
        'Valor ajuste':'Valor Ajuste',
        'Pago':'Valor Pagado',
        'Valor anulado':'Valor Anulado',
        'Documento integración':'Id Integracion',
        'Estado':'Estado Actual',
        'Periodo academico':'Periodico Academico',
        'Tipo financiación':'Tipo de Financiacion'
    })
    Fact20242['Documento'] = pd.to_numeric(Fact20242['Documento'].astype(str).str.strip(), errors='coerce')
    FactPol20242['Documento'] = pd.to_numeric(FactPol20242['Documento'].astype(str).str.strip(), errors='coerce')
    insumoFacturacion = pd.merge(Fact20242, FactPol20242, on='Documento', how='left')
    insumoFacturacion['Observacion'] = insumoFacturacion.apply(
        lambda row: 'Registro solo en Fact20242' if pd.isna(row['Aplica gratuidad']) else '',
        axis=1
    )
    soloFactPol = pd.merge(Fact20242[['Documento']], FactPol20242, on='Documento', how='right', indicator=True)
    soloFactPol = soloFactPol[soloFactPol['_merge'] == 'right_only'].drop(columns='_merge')
    soloFactPol['Observacion'] = 'Registro solo en FactPol20242'
    facturacionFinal = pd.concat([insumoFacturacion, soloFactPol], ignore_index=True)
    for col in Fact20242.columns:
      if col in facturacionFinal.columns:
        col_x = f"{col}_x"
        col_y = f"{col}_y"
        if col_x in facturacionFinal.columns and col_y in facturacionFinal.columns:
            facturacionFinal[col] = (
                facturacionFinal[col_x].combine_first(facturacionFinal[col_y])
                if not facturacionFinal[col_x].isna().all() or not facturacionFinal[col_y].isna().all()
                else np.nan
            )
            facturacionFinal.drop(columns=[col_x, col_y], inplace=True)
    return facturacionFinal

def depuradorPiam(piam20242civ, piam20242ci):
    filtroPiam20242ci = piam20242ci[['CODIGO','ID-SNIES','RECIBO','DERECHOS_MATRICULA','BIBLIOTECA_DEPORTES',
                          'LABORATORIOS','RECURSOS_COMPUTACIONALES','SEGURO_ESTUDIANTIL',
                          'VRES_COMPLEMENTARIOS','RESIDENCIAS','REPETICIONES','VOTO',
                          'CONVENIO_DESCENTRALIZACION','BECA','MATRICULA_HONOR','MEDIA_MATRICULA_HONOR',
                          'TRABAJO_GRADO','DOS_PROGRAMAS','DESCUENTO_HERMANO','ESTIMULO_EMP_DTE_PLANTA',
                          'ESTIMULO_CONYUGE','EXEN_HIJOS_CONYUGE_CATEDRA','HIJOS_TRABAJADORES_OFICIALES',
                          'ACTIVIDAES_LUDICAS_DEPOR','DESCUENTOS','SERVICIOS_RELIQUIDACION',
                          'DESCUENTO_LEY_1171','PROGRAMA','TELEFONO','CELULAR','EMAILINSTITUCIONAL',
                          'BRUTA','BRUTAORD','NETAORD','MERITO','MTRNETA','NETAAPL']]
    filtroPiam20242ci = filtroPiam20242ci.rename(columns={'CODIGO': 'codigo'})
    insumoPiam20242 = pd.merge(piam20242civ, filtroPiam20242ci, on='codigo', how='inner')
    return  insumoPiam20242

def integradorPiam(insumoPiam20242,insumoFacturacion,insumoIcetex):
    filtroInsumoFacturacion = insumoFacturacion[['Documento','Id  factura','Aplica gratuidad','Estado Actual',
                                                 'Valor Factura','Valor Pagado','Saldo','Cuenta bancaria','Tipo de Financiacion']]
    filtroInsumoIcetex = insumoIcetex[['ID FACTURA','Relación de Giro','VALOR DEL GIRO ICETEX','VALOR PAGO FACTURA APLICADO',
                                        'SALDO A FAVOR','MERITO UNICAUCA','Sublínea Crédito']]
    filtroInsumoFacturacion = filtroInsumoFacturacion.rename(columns={'Documento': 'RECIBO'})
    InsumoPiamF20242 = pd.merge(insumoPiam20242, filtroInsumoFacturacion, on='RECIBO', how='left')
    filtroInsumoIcetex = filtroInsumoIcetex.rename(columns={'ID FACTURA': 'Id  factura'})
    InsumoPiamFC20242 = pd.merge(InsumoPiamF20242, filtroInsumoIcetex, on='Id  factura', how='left')
    return InsumoPiamFC20242

# Extraccion
piam20242civ, piam20242ci, icetexAng, icetexCul, Fact20242, FactPol20242 = cargar_archivos_y_dataframes(file_path)

# Manipulación
try:
    insumoIcetex = depuradorIcetex(icetexAng, icetexCul)
    insumoFacturacion = depuradorFacturacion(Fact20242, FactPol20242)
    insumoPiam20242 = depuradorPiam(piam20242civ, piam20242ci)
    InsumoPiamFC20242 = integradorPiam(insumoPiam20242,insumoFacturacion,insumoIcetex) 
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
    if insumoPiam20242 is not None:
      insumoPiam20242.to_excel(writer, sheet_name='Piam20242', index=False)
      print('Se ha generado la plantilla del Piam 2024-2')
    if InsumoPiamFC20242 is not None:
      InsumoPiamFC20242.to_excel(writer, sheet_name='PIAMFC20242', index=False)
      print('Se ha generado la plantilla del PIAMF 2024-2')

print("Los resultados han sido guardados en el documento y archivo Excel.")
