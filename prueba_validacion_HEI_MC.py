## comentario sin relevancia
import os
import glob
import argparse
import unicodedata
import openpyxl
import pandas as pd
import pyodbc
import sqlalchemy as sa
from sqlalchemy import text
import numpy as np
from numpy import nan
from datetime import datetime
import uuid
from tqdm import tqdm
import re
import logging
import subprocess

logging.info("define arguments")

def getargs():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input",
                        default="/home/scristanchob/CO-JEP/migraciones/TANURIWA/new_work_flow/homl_validate/output/caso_02.parquet")
    parser.add_argument("--dict_juptr",
                        default="/datos/grai/gestion_de_la_informacion/parametros-jupiter/Dicionario_Migracion_MC_28052026.xlsx")
    parser.add_argument("--output",
                        default="/home/scristanchob/CO-JEP/migraciones/TANURIWA/new_work_flow/homl_validate/output/reporte_validacion_HEI.xlsx")
    return parser.parse_args()

args = getargs()

archivo_config = args.dict_juptr

archivo_datos = args.input

result_path = args.output

logging.warning("create function")
def vd(df):
    csv_memoria = df.write_csv().encode('utf-8')

    subprocess.run(["vd", "-f", "csv", "-"], input=csv_memoria)

logging.warning("import data")
# MACROCASOS
C_MACROCASOS = pd.read_excel(archivo_config, sheet_name="C_MACROCASOS",
                              usecols=["ID_MACROCASOS", "MACROCASOS"])

# CONDUCTA PENAL
C_CONDUCTA_PENAL = pd.read_excel(archivo_config, sheet_name="C_CONDUCTA_PENAL",
                              usecols=["ID_CONDUCTA_PENAL", "CONDUCTA_PENAL"])

# MODUS OPERANDI
C_MODUS_OPERANDI = pd.read_excel(archivo_config, sheet_name="C_MODUS_OPERANDI",
                              usecols=["ID_MODUS_OPERANDI", "MODUS_OPERANDI"])
# TIPO FAMILIA
C_TIPO_FAMILIA = pd.read_excel(archivo_config, sheet_name="C_TIPO_FAMILIA",
                              usecols=["ID_TIPO_FAMILIA", "TIPO_FAMILIA"])

# PARTIDO POLITICO
C_PARTIDO_POLITICO = pd.read_excel(archivo_config, sheet_name="C_TIPO_FAMILIA",
                              usecols=["ID_TIPO_FAMILIA", "TIPO_FAMILIA"])

# DEPARTAMENTO
L_DEPARTAMENTO = pd.read_excel(archivo_config, sheet_name="L_DEPARTAMENTO",
                              usecols=["ID_DEPARTAMENTO", "CODIGO_DANE", "DEPARTAMENTO"],
                               dtype={"ID_DEPARTAMENTO": str, "CODIGO_DANE": str})


L_CIUDAD = pd.read_excel(archivo_config, sheet_name="L_CIUDAD",
                              usecols=["ID_CIUDAD", "ID_DEPARTAMENTO", "CODIGO_DANE",
                                      "CIUDAD"],
                            dtype={"CODIGO_DANE": str})

# CENTRO POBLADO
L_CENTRO_POBLADO = pd.read_excel(archivo_config, sheet_name="L_CENTRO_POBLADO",
                              usecols=["ID_CENTRO_POBLADO", "ID_CIUDAD",
                                      "CODIGO_DANE", "CENTRO_POBLADO"])

# TIPO DE HECHO
C_TIPO_HECHO = pd.read_excel(archivo_config, sheet_name="C_TIPO_HECHO",
                              usecols=["ID_TIPO_HECHO", "TIPO_HECHO"])

# CATEGORIA DE CONDUCTA
C_CATEGORIA_CONDUCTA = pd.read_excel(archivo_config, sheet_name="C_CATEGORIA_CONDUCTA",
                              usecols=["ID_CATEGORIA_CONDUCTA", "CATEGORIA_CONDUCTA"])

# TIPO SUJETO
C_TIPO_SUJETO = pd.read_excel(archivo_config, sheet_name="C_TIPO_SUJETO",
                              usecols=["ID_TIPO_SUJETO", "TIPO_SUJETO"])

# PATRON
C_PATRON = pd.read_excel(archivo_config, sheet_name="C_PATRON",
                              usecols=["ID_PATRON", "PATRON"])

# SEXO
C_SEXO = pd.read_excel(archivo_config, sheet_name="C_SEXO",
                              usecols=["ID_SEXO", "SEXO"])

# TIPO DOCUMENTO
C_TIPO_DOCUMENTO = pd.read_excel(archivo_config, sheet_name="C_TIPO_DOCUMENTO",
                              usecols=["ID_TIPO_DOCUMENTO", "TIPO_DOCUMENTO"])

# PERTENENCIA ETNICA
C_PERTENENCIA_ETNICA = pd.read_excel(archivo_config, sheet_name="C_PERTENENCIA_ETNICA",
                              usecols=["ID_PERTENENCIA_ETNICA", "PERTENENCIA_ETNICA"])

# PUEBO O COMUNIDAD ETNICA
C_PUEBLO_O_COMUNIDAD_ETNICA = pd.read_excel(archivo_config, sheet_name="C_PUEBLO_O_COMUNIDAD_ETNICA",
                              usecols=["ID_PUEBLO_O_COMUNIDAD_ETNICA", "PUEBLO_O_COMUNIDAD_ETNICA"])

# TERRITORIO COLECTIVO
C_TERRITORIO_COLECTIVO = pd.read_excel(archivo_config, sheet_name="C_TERRITORIO_COLECTIVO",
                              usecols=["ID_TERRITORIO_COLECTIVO", "TERRITORIO_COLECTIVO"])

# RANGO DE EDAD
C_RANGO_EDAD = pd.read_excel(archivo_config, sheet_name="C_RANGO_EDAD",
                              usecols=["ID_RANGO_EDAD", "RANGO_EDAD"])

# IDENTIDAD DE GENERO
C_IDENTIDAD_DE_GENERO = pd.read_excel(archivo_config, sheet_name="C_IDENTIDAD_DE_GENERO",
                              usecols=["ID_IDENTIDAD_DE_GENERO", "IDENTIDAD_DE_GENERO"])

# ORIENTACION SEXUAL
C_ORIENTACION_SEXUAL = pd.read_excel(archivo_config, sheet_name="C_ORIENTACION_SEXUAL",
                              usecols=["ID_ORIENTACION_SEXUAL", "ORIENTACION_SEXUAL"])

# TOPONIMIA
C_TOPONIMIA = pd.read_excel(archivo_config, sheet_name="C_TOPONIMIA",
                              usecols=["ID_TOPONIMIA", "TOPONIMIA"])

# ACTOR
C_ACTOR = pd.read_excel(archivo_config, sheet_name="C_ACTOR",
                              usecols=["ID_ACTOR", "ACTOR"])

#GRUPO
C_GRUPO = pd.read_excel(archivo_config, sheet_name="C_GRUPO",
                              usecols=["ID_ACTOR", "ID_GRUPO", "GRUPO"])

#ESTRUCTURA NIVEL 1
C_ESTRUCTURA_NIVEL_1 = pd.read_excel(archivo_config, sheet_name="C_ESTRUCTURA_NIVEL_1",
                              usecols=["ID_GRUPO", "ID_ESTRUCTURA_NIVEL_1",
                                      "ESTRUCTURA_NIVEL_1"])

#ESTRUCTURA NIVEL 2
C_ESTRUCTURA_NIVEL_2 = pd.read_excel(archivo_config, sheet_name="C_ESTRUCTURA_NIVEL_2",
                              usecols=["ID_ESTRUCTURA_NIVEL_1", "ID_ESTRUCTURA_NIVEL_2",
                                      "ESTRUCTURA_NIVEL_2"])

#ESTRUCTURA NIVEL 3
C_ESTRUCTURA_NIVEL_3 = pd.read_excel(archivo_config, sheet_name="C_ESTRUCTURA_NIVEL_3",
                              usecols=["ID_ESTRUCTURA_NIVEL_2", "ID_ESTRUCTURA_NIVEL_3",
                                      "ESTRUCTURA_NIVEL_3"])

#ESTRUCTURA NIVEL 4
C_ESTRUCTURA_NIVEL_4 = pd.read_excel(archivo_config, sheet_name="C_ESTRUCTURA_NIVEL_4",
                              usecols=["ID_ESTRUCTURA_NIVEL_3", "ID_ESTRUCTURA_NIVEL_4",
                                      "ESTRUCTURA_NIVEL_4"])

#C_OCUPACION_NIVEL_1
C_OCUPACION_NIVEL_1 = pd.read_excel(archivo_config, sheet_name="C_OCUPACION_NIVEL_1",
                              usecols=["ID_OCUPACION_NIVEL_1", "OCUPACION_NIVEL_1"])

#C_OCUPACION_NIVEL_2
C_OCUPACION_NIVEL_2 = pd.read_excel(archivo_config, sheet_name="C_OCUPACION_NIVEL_2",
                              usecols=["ID_OCUPACION_NIVEL_2", "OCUPACION_NIVEL_2"])

#C_ROL_SOCIAL
C_ROL_SOCIAL = pd.read_excel(archivo_config, sheet_name="C_ROL_SOCIAL",
                              usecols=["ID_ROL_SOCIAL", "ROL_SOCIAL"])

#C_TIPO_SOLICITUD_ACREDITACION
C_TIPO_SOLICITUD_ACREDITACION = pd.read_excel(archivo_config, sheet_name="C_TIPO_SOLICITUD_ACREDITACION",
                              usecols=["ID_TIPO_SOLICITUD_ACREDITACION", "TIPO_SOLICITUD_ACREDITACION"])

#C_RAZON_DECISION
C_RAZON_DECISION = pd.read_excel(archivo_config, sheet_name="C_RAZON_DECISION",
                              usecols=["ID_RAZON_DECISION", "RAZON_DECISION"])

#C_CLASE_ACREDITACION
C_CLASE_ACREDITACION = pd.read_excel(archivo_config, sheet_name="C_CLASE_ACREDITACION",
                              usecols=["ID_CLASE_ACREDITACION", "CLASE_ACREDITACION"])

#C_RELACION_CON_LA_VICTIMA
C_RELACION_CON_LA_VICTIMA = pd.read_excel(archivo_config, sheet_name="C_RELACION_CON_LA_VICTIMA",
                              usecols=["ID_RELACION_CON_LA_VICTIMA", "RELACION_CON_LA_VICTIMA"])

#C_DECISION_ACREDITACION
C_DECISION_ACREDITACION = pd.read_excel(archivo_config, sheet_name="C_DECISION_ACREDITACION",
                              usecols=["ID_DECISION_ACREDITACION", "DECISION_ACREDITACION"])

#C_TIPO_AFECTACION
C_TIPO_AFECTACION = pd.read_excel(archivo_config, sheet_name="C_TIPO_AFECTACION",
                              usecols=["ID_TIPO_AFECTACION", "TIPO_AFECTACION"])

#C_IMPACTOS
C_IMPACTOS = pd.read_excel(archivo_config, sheet_name="C_IMPACTOS",
                              usecols=["ID_IMPACTOS", "IMPACTOS"])

#C_IMPACTOS_ESPECIFICOS
C_IMPACTOS_ESPECIFICOS = pd.read_excel(archivo_config, sheet_name="C_IMPACTOS_ESPECIFICOS",
                              usecols=["ID_IMPACTOS_ESPECIFICOS", "IMPACTOS_ESPECIFICOS"])

#C_LINEA_BASE_REPARACION
C_LINEA_BASE_REPARACION = pd.read_excel(archivo_config, sheet_name="C_LINEA_BASE_REPARACION",
                              usecols=["ID_LINEA_BASE_REPARACION", "LINEA_BASE_REPARACION"])

#C_DEMANDA_ESPECIFICA
C_DEMANDA_ESPECIFICA = pd.read_excel(archivo_config, sheet_name="C_DEMANDA_ESPECIFICA",
                              usecols=["ID_DEMANDA_ESPECIFICA", "DEMANDA_ESPECIFICA"])

#C_DEMANDA_ESPECIFICA
C_DEMANDA_ESPECIFICA = pd.read_excel(archivo_config, sheet_name="C_DEMANDA_ESPECIFICA",
                              usecols=["ID_DEMANDA_ESPECIFICA", "DEMANDA_ESPECIFICA"])

#C_ESTADO_SUJETO
C_ESTADO_SUJETO = pd.read_excel(archivo_config, sheet_name="C_ESTADO_SUJETO",
                              usecols=["ID_ESTADO_SUJETO", "ESTADO_SUJETO"])

#C_ESCOLARIDAD
C_ESCOLARIDAD = pd.read_excel(archivo_config, sheet_name="C_ESCOLARIDAD",
                              usecols=["ID_ESCOLARIDAD", "ESCOLARIDAD"])

# M_FUENTE
M_FUENTE = pd.read_excel(archivo_config, sheet_name="M_FUENTE")

# C_LUGAR
df_lugar =  pd.read_excel(archivo_config, sheet_name="C_LUGAR")
C_LUGAR = df_lugar.loc[df_lugar['ACTIVO'] == 1, ['ID_LUGAR', 'LUGAR']]

df_actor = pd.read_excel(archivo_config, sheet_name='C_ACTOR')
df_grupo = pd.read_excel(archivo_config, sheet_name='C_GRUPO')
df_n1    = pd.read_excel(archivo_config, sheet_name='C_ESTRUCTURA_NIVEL_1')
df_n2    = pd.read_excel(archivo_config, sheet_name='C_ESTRUCTURA_NIVEL_2')
df_n3    = pd.read_excel(archivo_config, sheet_name='C_ESTRUCTURA_NIVEL_3')
df_n4    = pd.read_excel(archivo_config, sheet_name='C_ESTRUCTURA_NIVEL_4')

df_n1 = df_n1.loc[df_n1['ACTIVO'] == 1, ['ID_ESTRUCTURA_NIVEL_1', 'ID_GRUPO',
                                                 "ESTRUCTURA_NIVEL_1"]]

df_n2 = df_n2.loc[df_n2['ACTIVO'] == 1, ['ID_ESTRUCTURA_NIVEL_2', 'ID_ESTRUCTURA_NIVEL_1',
                                                 "ESTRUCTURA_NIVEL_2"]]

df_n3 = df_n3.loc[df_n3['ACTIVO'] == 1, ['ID_ESTRUCTURA_NIVEL_3', 'ID_ESTRUCTURA_NIVEL_2',
                                                 "ESTRUCTURA_NIVEL_3"]]

df_n4 = df_n4.loc[df_n4['ACTIVO'] == 1, ['ID_ESTRUCTURA_NIVEL_4', 'ID_ESTRUCTURA_NIVEL_3',
                                                 "ESTRUCTURA_NIVEL_4"]]

C_GRUPO_ARMADO = (
    df_actor
    .merge(df_grupo, on='ID_ACTOR', how='left')
    .merge(df_n1, on='ID_GRUPO', how='left')
    .merge(df_n2, on='ID_ESTRUCTURA_NIVEL_1', how='left')
    .merge(df_n3, on='ID_ESTRUCTURA_NIVEL_2', how='left')
    .merge(df_n4, on='ID_ESTRUCTURA_NIVEL_3', how='left')
)

# 3. Creación de las tablas finales (Vistas)
C_GRUPO_ACTOR_GRUPO = C_GRUPO_ARMADO[["ID_ACTOR", "ACTOR", "ID_GRUPO", "GRUPO"]].drop_duplicates()
C_GRUPO_GRUPO_N1    = C_GRUPO_ARMADO[["ID_GRUPO", "GRUPO", "ID_ESTRUCTURA_NIVEL_1", "ESTRUCTURA_NIVEL_1"]].drop_duplicates()
C_GRUPO_N1_N2       = C_GRUPO_ARMADO[["ID_ESTRUCTURA_NIVEL_1", "ESTRUCTURA_NIVEL_1", "ID_ESTRUCTURA_NIVEL_2", "ESTRUCTURA_NIVEL_2"]].drop_duplicates()
C_GRUPO_N2_N3       = C_GRUPO_ARMADO[["ID_ESTRUCTURA_NIVEL_2", "ESTRUCTURA_NIVEL_2", "ID_ESTRUCTURA_NIVEL_3", "ESTRUCTURA_NIVEL_3"]].drop_duplicates()
C_GRUPO_N3_N4       = C_GRUPO_ARMADO[["ID_ESTRUCTURA_NIVEL_3", "ESTRUCTURA_NIVEL_3", "ID_ESTRUCTURA_NIVEL_4", "ESTRUCTURA_NIVEL_4"]].drop_duplicates()

ocupacn_1 = pd.read_excel(archivo_config, sheet_name='C_OCUPACION_NIVEL_1')

ocupacn_1 = ocupacn_1.loc[ocupacn_1['ACTIVO'] == 1, ['ID_OCUPACION_NIVEL_1', 'CODIGO',
                                                 "OCUPACION_NIVEL_1"]]

ocupacn_2 = pd.read_excel(archivo_config, sheet_name='C_OCUPACION_NIVEL_2')

ocupacn_2 = ocupacn_2.loc[ocupacn_2['ACTIVO'] == 1, ['ID_OCUPACION_NIVEL_2', 'ID_OCUPACION_NIVEL_1',
                                                 "OCUPACION_NIVEL_2"]]

C_OCUPACION = (
    ocupacn_1
    .merge(ocupacn_2, on='ID_OCUPACION_NIVEL_1', how='inner')
)

thcp_1 = pd.read_excel(archivo_config, sheet_name='C_TIPO_HECHO')

thcp_1 = thcp_1.loc[thcp_1['ACTIVO'] == 1, ['ID_TIPO_HECHO', 'TIPO_HECHO']]

thcp_2 = pd.read_excel(archivo_config, sheet_name='C_CONDUCTA_PENAL')

C_TIPO_HECHO_CONDUCTA_PENAL = (
    thcp_1
    .merge(thcp_2, on='ID_TIPO_HECHO', how='left')
)

divi_1 = pd.read_excel(archivo_config, sheet_name='L_DEPARTAMENTO',
                            dtype={"CODIGO_DANE": str})

divi_1 = divi_1.loc[divi_1['ACTIVO'] == 1, ['ID_DEPARTAMENTO', 'CODIGO_DANE',
                                           'DEPARTAMENTO']]

divi_1["DEPARTAMENTO_CODIGO_DANE"] = divi_1["CODIGO_DANE"]

divi_2 = pd.read_excel(archivo_config, sheet_name='L_CIUDAD',
                            dtype={"CODIGO_DANE": str})

divi_2 = divi_2.loc[divi_2['ACTIVO'] == 1, ['ID_CIUDAD', 'ID_DEPARTAMENTO',
                                           'CODIGO_DANE', 'CIUDAD']]

divi_2["CIUDAD_CODIGO_DANE"] = divi_2["CODIGO_DANE"]

divi_3 = pd.read_excel(archivo_config, sheet_name='L_CENTRO_POBLADO')

divi_3 = divi_3.loc[divi_3['ACTIVO'] == 1, ['ID_CENTRO_POBLADO', 'TIPO_MUNICIPIO', 'ID_CIUDAD',
                                           'DISTRITO', 'AREA_METROPOLITANA', "CODIGO_DANE", "CENTRO_POBLADO"]]

divi_3["CENTRO_POBLADO_CODIGO_DANE"] = divi_3["CODIGO_DANE"]

C_DIVIPOLA = (
    divi_1
    .merge(divi_2, on='ID_DEPARTAMENTO', how='left')
    .merge(divi_3, on='ID_CIUDAD', how='left')
)

C_DEPTO_CIUDAD          = C_DIVIPOLA[["DEPARTAMENTO_CODIGO_DANE", "DEPARTAMENTO", "CIUDAD_CODIGO_DANE", "CIUDAD"]].drop_duplicates()
C_CIUDAD_CENTRO_POBLADO = C_DIVIPOLA[["CIUDAD_CODIGO_DANE", "CIUDAD", "CENTRO_POBLADO_CODIGO_DANE", "CENTRO_POBLADO"]].drop_duplicates()

ge = pd.read_excel(archivo_config, sheet_name='C_PERTENENCIA_ETNICA')

ge = ge.loc[ge['ACTIVO'] == 1, ['ID_PERTENENCIA_ETNICA', 'PERTENENCIA_ETNICA']]

pce = pd.read_excel(archivo_config, sheet_name='C_PUEBLO_O_COMUNIDAD_ETNICA')

pce = pce.loc[pce['ACTIVO'] == 1, ['ID_PUEBLO_O_COMUNIDAD_ETNICA', 'ID_PERTENENCIA_ETNICA',
                                  "PUEBLO_O_COMUNIDAD_ETNICA"]]

tc = pd.read_excel(archivo_config, sheet_name='C_TERRITORIO_COLECTIVO')

tc = tc.loc[tc['ACTIVO'] == 1, ['ID_TERRITORIO_COLECTIVO', 'ID_PUEBLO_O_COMUNIDAD_ETNICA',
                                  "TERRITORIO_COLECTIVO"]]

C_GRUPO_ETNICO = (
    ge
    .merge(pce, on='ID_PERTENENCIA_ETNICA', how='left')
    .merge(tc, on='ID_PUEBLO_O_COMUNIDAD_ETNICA', how='left')
)

C_PERTENECIA_COMUNIDAD_ETNICA    = C_GRUPO_ETNICO[["ID_PERTENENCIA_ETNICA", "PERTENENCIA_ETNICA", "ID_PUEBLO_O_COMUNIDAD_ETNICA", "PUEBLO_O_COMUNIDAD_ETNICA"]].drop_duplicates()
C_COMUNIDAD_TERRITORIO_COLECTIVO = C_GRUPO_ETNICO[["ID_PUEBLO_O_COMUNIDAD_ETNICA", "PUEBLO_O_COMUNIDAD_ETNICA", "ID_TERRITORIO_COLECTIVO", "TERRITORIO_COLECTIVO"]].drop_duplicates()


# ---------------------------
# CARGUE DE CATALOGOS EN LISTAS
# ---------------------------

# ---------------- CONFIGURACIÓN DE CATALOGOS SENCILLOS
catalogos = {
    "C_MACROCASOS"                  : C_MACROCASOS,
    "C_CONDUCTA_PENAL"              : C_CONDUCTA_PENAL,
    "C_MODUS_OPERANDI"              : C_MODUS_OPERANDI,
    "C_TIPO_FAMILIA"                : C_TIPO_FAMILIA,
    "C_PARTIDO_POLITICO"            : C_PARTIDO_POLITICO,
    "L_DEPARTAMENTO"                : L_DEPARTAMENTO,
    "L_CIUDAD"                      : L_CIUDAD,
    "L_CENTRO_POBLADO"              : L_CENTRO_POBLADO,
    "C_TIPO_HECHO"                  : C_TIPO_HECHO,
    "C_CATEGORIA_CONDUCTA"          : C_CATEGORIA_CONDUCTA,
    "C_TIPO_SUJETO"                 : C_TIPO_SUJETO,
    "C_PATRON"                      : C_PATRON,
    "C_SEXO"                        : C_SEXO,
    "C_TIPO_DOCUMENTO"              : C_TIPO_DOCUMENTO,
    "C_PERTENENCIA_ETNICA"          : C_PERTENENCIA_ETNICA,
    "C_PUEBLO_O_COMUNIDAD_ETNICA"   : C_PUEBLO_O_COMUNIDAD_ETNICA,
    "C_TERRITORIO_COLECTIVO"        : C_TERRITORIO_COLECTIVO,
    "C_RANGO_EDAD"                  : C_RANGO_EDAD,
    "C_IDENTIDAD_DE_GENERO"         : C_IDENTIDAD_DE_GENERO,
    "C_ORIENTACION_SEXUAL"          : C_ORIENTACION_SEXUAL,
    "C_TOPONIMIA"                   : C_TOPONIMIA,
    "C_ACTOR"                       : C_ACTOR,
    "C_GRUPO"                       : C_GRUPO,
    "C_ESTRUCTURA_NIVEL_1"          : C_ESTRUCTURA_NIVEL_1,
    "C_ESTRUCTURA_NIVEL_2"          : C_ESTRUCTURA_NIVEL_2,
    "C_ESTRUCTURA_NIVEL_3"          : C_ESTRUCTURA_NIVEL_3,
    "C_ESTRUCTURA_NIVEL_4"          : C_ESTRUCTURA_NIVEL_4,
    "C_OCUPACION_NIVEL_1"           : C_OCUPACION_NIVEL_1,
    "C_OCUPACION_NIVEL_2"           : C_OCUPACION_NIVEL_2,
    "C_ROL_SOCIAL"                  : C_ROL_SOCIAL,
    "C_TIPO_SOLICITUD_ACREDITACION" : C_TIPO_SOLICITUD_ACREDITACION,
    "C_RAZON_DECISION"              : C_RAZON_DECISION,
    "C_CLASE_ACREDITACION"          : C_CLASE_ACREDITACION,
    "C_RELACION_CON_LA_VICTIMA"     : C_RELACION_CON_LA_VICTIMA,
    "C_DECISION_ACREDITACION"       : C_DECISION_ACREDITACION,
    "C_TIPO_AFECTACION"             : C_TIPO_AFECTACION,
    "C_IMPACTOS"                    : C_IMPACTOS,
    "C_IMPACTOS_ESPECIFICOS"        : C_IMPACTOS_ESPECIFICOS,
    "C_LINEA_BASE_REPARACION"       : C_LINEA_BASE_REPARACION,
    "C_DEMANDA_ESPECIFICA"          : C_DEMANDA_ESPECIFICA,
    "C_ESTADO_SUJETO"               : C_ESTADO_SUJETO,
    "C_ESCOLARIDAD"                 : C_ESCOLARIDAD,
    "M_FUENTE"                      : M_FUENTE,
    "C_LUGAR"                       : C_LUGAR
}

# ---------------- CONFIGURACIÓN DE CATALOGOS DEPENDIENTES
dependencias = {
    # ----------------------------------------------------------------- HECHOS
    "H_CONDUCTA_PENAL": {
        "campo_padre"   : "H_TIPO_HECHO",
        "catalogo"      : C_TIPO_HECHO_CONDUCTA_PENAL,
        "col_padre"     : "TIPO_HECHO",
        "col_hijo"      : "CONDUCTA_PENAL"
    },
    "H_CIUDAD": {
        "campo_padre"   : "H_DEPARTAMENTO",
        "catalogo"      : C_DEPTO_CIUDAD,
        "col_padre"     : "DEPARTAMENTO_CODIGO_DANE",
        "col_hijo"      : "CIUDAD_CODIGO_DANE"
    },
    "H_CENTRO_POBLADO": {
        "campo_padre"   : "H_CIUDAD",
        "catalogo"      : C_CIUDAD_CENTRO_POBLADO,
        "col_padre"     : "CIUDAD_CODIGO_DANE",
        "col_hijo"      : "CENTRO_POBLADO_CODIGO_DANE"
    },
    "H_COMUNIDAD_ETNICA": {
        "campo_padre"   : "H_PERTENENCIA_ETNICA",
        "catalogo"      : C_PERTENECIA_COMUNIDAD_ETNICA,
        "col_padre"     : "PERTENENCIA_ETNICA",
        "col_hijo"      : "PUEBLO_O_COMUNIDAD_ETNICA"
    },
    "H_NOMBRE_TERRITORIO_COLECTIVO": {
        "campo_padre"   : "H_COMUNIDAD_ETNICA",
        "catalogo"      : C_COMUNIDAD_TERRITORIO_COLECTIVO,
        "col_padre"     : "PUEBLO_O_COMUNIDAD_ETNICA",
        "col_hijo"      : "TERRITORIO_COLECTIVO"
    },
    # ----------------------------------------------------------------- ENTIDADES
    "E_CIUDAD_NACIMIENTO": {
        "campo_padre"   : "E_DEPARTAMENTO_NACIMIENTO",
        "catalogo"      : C_DEPTO_CIUDAD,
        "col_padre"     : "DEPARTAMENTO_CODIGO_DANE",
        "col_hijo"      : "CIUDAD_CODIGO_DANE"
    },
    "E_CENTRO_POBLADO_NACIMIENTO": {
        "campo_padre"   : "E_CIUDAD_NACIMIENTO",
        "catalogo"      : C_CIUDAD_CENTRO_POBLADO,
        "col_padre"     : "CIUDAD_CODIGO_DANE",
        "col_hijo"      : "CENTRO_POBLADO_CODIGO_DANE"
    },
    "E_COMUNIDAD_ETNICA": {
        "campo_padre"   : "E_PERTENENCIA_ETNICA",
        "catalogo"      : C_PERTENECIA_COMUNIDAD_ETNICA,
        "col_padre"     : "PERTENENCIA_ETNICA",
        "col_hijo"      : "PUEBLO_O_COMUNIDAD_ETNICA"
    },
    "E_NOMBRE_TERRITORIO_COLECTIVO": {
        "campo_padre"   : "E_COMUNIDAD_ETNICA",
        "catalogo"      : C_COMUNIDAD_TERRITORIO_COLECTIVO,
        "col_padre"     : "PUEBLO_O_COMUNIDAD_ETNICA",
        "col_hijo"      : "TERRITORIO_COLECTIVO"
    },    
    "E_OCUPACION_NIVEL_2": {
        "campo_padre"   : "E_OCUPACION_NIVEL_1",
        "catalogo"      : C_OCUPACION,
        "col_padre"     : "OCUPACION_NIVEL_1",
        "col_hijo"      : "OCUPACION_NIVEL_2"
    },
    "E_CIUDAD_RESIDENCIA": {
        "campo_padre"   : "E_DEPARTAMENTO_RESIDENCIA",
        "catalogo"      : C_DEPTO_CIUDAD,
        "col_padre"     : "DEPARTAMENTO_CODIGO_DANE",
        "col_hijo"      : "CIUDAD_CODIGO_DANE"
    },
    "E_CENTRO_POBLADO_RESIDENCIA": {
        "campo_padre"   : "E_CIUDAD_RESIDENCIA",
        "catalogo"      : C_CIUDAD_CENTRO_POBLADO,
        "col_padre"     : "CIUDAD_CODIGO_DANE",
        "col_hijo"      : "CENTRO_POBLADO_CODIGO_DANE"
    },
    "E_CIUDAD_NOTIFICACION": {
        "campo_padre"   : "E_DEPARTAMENTO_NOTIFICACION",
        "catalogo"      : C_DEPTO_CIUDAD,
        "col_padre"     : "DEPARTAMENTO_CODIGO_DANE",
        "col_hijo"      : "CIUDAD_CODIGO_DANE"
    },
    "E_CENTRO_POBLADO_NOTIFICACION": {
        "campo_padre"   : "E_CIUDAD_NOTIFICACION",
        "catalogo"      : C_CIUDAD_CENTRO_POBLADO,
        "col_padre"     : "CIUDAD_CODIGO_DANE",
        "col_hijo"      : "CENTRO_POBLADO_CODIGO_DANE"
    },
    "E_GRUPO": {
        "campo_padre"   : "E_ACTOR",
        "catalogo"      : C_GRUPO_ACTOR_GRUPO,
        "col_padre"     : "ACTOR",
        "col_hijo"      : "GRUPO"
    },
    "E_ESTRUCTURA_NIVEL_1": {
        "campo_padre"   : "E_GRUPO",
        "catalogo"      : C_GRUPO_GRUPO_N1,
        "col_padre"     : "GRUPO",
        "col_hijo"      : "ESTRUCTURA_NIVEL_1"
    },
    "E_ESTRUCTURA_NIVEL_2": {
        "campo_padre"   : "E_ESTRUCTURA_NIVEL_1",
        "catalogo"      : C_GRUPO_N1_N2,
        "col_padre"     : "ESTRUCTURA_NIVEL_1",
        "col_hijo"      : "ESTRUCTURA_NIVEL_2"
    },
    "E_ESTRUCTURA_NIVEL_3": {
        "campo_padre"   : "E_ESTRUCTURA_NIVEL_2",
        "catalogo"      : C_GRUPO_N2_N3,
        "col_padre"     : "ESTRUCTURA_NIVEL_2",
        "col_hijo"      : "ESTRUCTURA_NIVEL_3"
    },
    "E_ESTRUCTURA_NIVEL_4": {
        "campo_padre"   : "E_ESTRUCTURA_NIVEL_3",
        "catalogo"      : C_GRUPO_N3_N4,
        "col_padre"     : "ESTRUCTURA_NIVEL_3",
        "col_hijo"      : "ESTRUCTURA_NIVEL_4"
    },
    "E_CIUDAD_COMUNIDAD_MUNICIPAL": {
        "campo_padre"   : "E_DEPARTAMENTO_COMUNIDAD_MUNICIPAL",
        "catalogo"      : C_DEPTO_CIUDAD,
        "col_padre"     : "DEPARTAMENTO_CODIGO_DANE",
        "col_hijo"      : "CIUDAD_CODIGO_DANE"
    },
    "E_CENTRO_POBLADO_COMUNIDAD_MUNICIPAL": {
        "campo_padre"   : "E_CIUDAD_COMUNIDAD_MUNICIPAL",
        "catalogo"      : C_CIUDAD_CENTRO_POBLADO,
        "col_padre"     : "CIUDAD_CODIGO_DANE",
        "col_hijo"      : "CENTRO_POBLADO_CODIGO_DANE"
    },    
    # ----------------------------------------------------------------- INFORMACION AL MOMENTO DEL HECHO
    "I_OCUPACION_NIVEL_2": {
        "campo_padre"   : "I_OCUPACION_NIVEL_1",
        "catalogo"      : C_OCUPACION,
        "col_padre"     : "OCUPACION_NIVEL_1",
        "col_hijo"      : "OCUPACION_NIVEL_2"
    },
    "I_GRUPO": {
        "campo_padre"   : "I_ACTOR",
        "catalogo"      : C_GRUPO_ACTOR_GRUPO,
        "col_padre"     : "ACTOR",
        "col_hijo"      : "GRUPO"
    },
    "I_ESTRUCTURA_NIVEL_1": {
        "campo_padre"   : "I_GRUPO",
        "catalogo"      : C_GRUPO_GRUPO_N1,
        "col_padre"     : "GRUPO",
        "col_hijo"      : "ESTRUCTURA_NIVEL_1"
    },
    "I_ESTRUCTURA_NIVEL_2": {
        "campo_padre"   : "I_ESTRUCTURA_NIVEL_1",
        "catalogo"      : C_GRUPO_N1_N2,
        "col_padre"     : "ESTRUCTURA_NIVEL_1",
        "col_hijo"      : "ESTRUCTURA_NIVEL_2"
    },
    "I_ESTRUCTURA_NIVEL_3": {
        "campo_padre"   : "I_ESTRUCTURA_NIVEL_2",
        "catalogo"      : C_GRUPO_N2_N3,
        "col_padre"     : "ESTRUCTURA_NIVEL_2",
        "col_hijo"      : "ESTRUCTURA_NIVEL_3"
    },
    "I_ESTRUCTURA_NIVEL_4": {
        "campo_padre"   : "I_ESTRUCTURA_NIVEL_3",
        "catalogo"      : C_GRUPO_N3_N4,
        "col_padre"     : "ESTRUCTURA_NIVEL_3",
        "col_hijo"      : "ESTRUCTURA_NIVEL_4"
    },
    "I_CIUDAD": {
        "campo_padre"   : "I_DEPARTAMENTO",
        "catalogo"      : C_DEPTO_CIUDAD,
        "col_padre"     : "DEPARTAMENTO_CODIGO_DANE",
        "col_hijo"      : "CIUDAD_CODIGO_DANE"
    },
    "I_CENTRO_POBLADO": {
        "campo_padre"   : "I_CIUDAD",
        "catalogo"      : C_CIUDAD_CENTRO_POBLADO,
        "col_padre"     : "CIUDAD_CODIGO_DANE",
        "col_hijo"      : "CENTRO_POBLADO_CODIGO_DANE"
    },
    "I_CIUDAD_UBICACION": {
        "campo_padre"   : "I_DEPARTAMENTO_UBICACION",
        "catalogo"      : C_DEPTO_CIUDAD,
        "col_padre"     : "DEPARTAMENTO_CODIGO_DANE",
        "col_hijo"      : "CIUDAD_CODIGO_DANE"
    },
    "I_CENTRO_POBLADO_UBICACION": {
        "campo_padre"   : "I_CIUDAD_UBICACION",
        "catalogo"      : C_CIUDAD_CENTRO_POBLADO,
        "col_padre"     : "CIUDAD_CODIGO_DANE",
        "col_hijo"      : "CENTRO_POBLADO_CODIGO_DANE"
    },    
}

logging.warning("create functions")

def to_varchar(x):
    if pd.isna(x):
        return None
    try:
        return str(int(float(x)))
    except:
        return str(x).strip()

def normalizar_texto(texto):
    if pd.isna(texto):
        return None

    texto = str(texto).strip()

    # quitar comillas normales y tipográficas
    comillas = ['"', "'", "“", "”", "‘", "’", "`", "´"]
    for c in comillas:
        texto = texto.replace(c, "")

    texto = texto.upper()

    # quitar tildes
    texto = ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )

    # quitar dobles espacios
    texto = " ".join(texto.split())

    return texto

def obtener_columna_alfanumerica(df_catalogo):
    for col in df_catalogo.columns:
        if not col.upper().startswith("ID_"):
            return col
    return None

def split_valores(valor):
    if pd.isna(valor):
        return []
    return [v.strip() for v in str(valor).split("|") if v.strip()]

reglas_condicionales = [
    {
        "si_columna"        : "H_INCLUYE_INFORMACION_ETNICA", 
        "si_valor"          : "SI",
        "entonces_columna"  : "H_PERTENENCIA_ETNICA",
        "tipo"              : "REQUERIDO"
    },
    {
        "si_columna"        : "E_INCLUYE_INFORMACION_DE_VINCULACION", 
        "si_valor"          : "SI",
        "entonces_columna"  : "E_ACTOR",
        "tipo"              : "REQUERIDO"
    },
    {
        "si_columna"        : "I_CONOCE_INFORMACION_PERSONA_NATURAL", 
        "si_valor"          : "SI",
        "entonces_columna"  : "I_ACTOR",
        "tipo"              : "REQUERIDO"
    },
]

reglas_i_id = [
    {
        "col_hecho"     : "H_ID",
        "col_entidad"   : "E_ID",
        "col_resultado" : "I_ID"        
    },
    # {
    #     "col_resultado": "H_ID",
    #     "col_hecho": "ID_HECHO",
    #     "col_entidad": "ID_PERSONA"
    # }
]

campos_pipe       = {"H_EVENTOS", "H_CATEGORIA_CONDUCTA", "H_PATRONES", "H_TITULO_FUENTE", "E_DOCUMENTO(S)", "E_TITULO_FUENTE", "I_TIPO_SUJETO","I_DISCAPACIDAD", "I_TITULO_FUENTE"}
campos_lista      = {"E_LISTADOS_MINDEFENSA", "E_CORREOS_ELECTRONICOS", "E_EMAILS_NOTIFICACION"}
campos_lista_num  = {"E_ACTAS_SUSCRITAS", "E_RESOLUCIONES_OACP"}
campos_titulo     = {"H_TITULO_FUENTE", "E_TITULO_FUENTE", "I_TITULO_FUENTE"}

def validar_campos_existentes(config, df, errores):
    campos_config = set(config["CAMPO"])
    campos_datos = set(df.columns)

    faltantes = campos_config - campos_datos

    for campo in faltantes:
        errores.append([ None, campo, "CAMPO_NO_EXISTE", "Campo no existe en archivo de datos"])
        
def validar_rango_fechas(df, errores):
    if "H_FECHA_INICIAL" in df.columns and "H_FECHA_FINAL" in df.columns:
        fecha_ini = pd.to_datetime(df["H_FECHA_INICIAL"], format="%d/%m/%Y", errors="coerce")
        fecha_fin = pd.to_datetime(df["H_FECHA_FINAL"], format="%d/%m/%Y", errors="coerce")
        mask = (fecha_ini > fecha_fin) & fecha_ini.notna() & fecha_fin.notna()

        for idx in df[mask].index:
            errores.append([ df.loc[idx,"_fila_excel"], "H_FECHA_INICIAL", "RANGO_FECHA", "H_FECHA_INICIAL mayor que H_FECHA_FINAL"])

def validar_obligatorio(df, serie, campo, obligatorio, errores):
    if obligatorio:
        mask = serie.isna() | (serie.astype(str).str.strip() == "")
        for idx in df[mask].index:
            errores.append([ df.loc[idx,"_fila_excel"], campo, "OBLIGATORIO", "Valor vacío"])

def validar_longitud(df, serie, campo, longitud, errores):

    # campos a omitir
    campos_omitidos = { "H_CODIGO_INDEXACION_FUENTE", "E_CODIGO_INDEXACION_FUENTE", "I_CODIGO_INDEXACION_FUENTE",
                        "E_TITULO_FUENTE", "H_TITULO_FUENTE", "I_TITULO_FUENTE" }

    if campo in campos_omitidos:
        return
    
    if pd.isna(longitud):
        return

    # función para normalizar valores numéricos
    def normalizar_valor(v):
        if pd.isna(v):
            return ""
        
        if isinstance(v, float) and v.is_integer(): # si es número entero representado como float
            return str(int(v))
        return str(v).strip()

    serie_norm = serie.apply(normalizar_valor)
    mask_con_info = serie_norm != "" # solo valores con información
    mask = mask_con_info & (serie_norm.str.len() > int(longitud)) # validar longitud

    for idx in df[mask].index:
        errores.append([ df.loc[idx,"_fila_excel"], campo, "LONGITUD", f"Supera longitud max: {int(longitud)}"])

def validar_fecha(df, serie, campo, tipo, errores):
    if tipo == "date":
        fechas = pd.to_datetime(serie, format="%d/%m/%Y", errors="coerce")
        mask = fechas.isna() & serie.notna()

        for idx in df[mask].index:
            errores.append([ df.loc[idx,"_fila_excel"], campo, "FECHA", "Formato debe ser DD/MM/AAAA" ])

def validar_si_no(df, serie, campo, observacion, errores):

    if any(x in observacion for x in [
        "SI / NO",
        "SI / NO / SE DESCONOCE",
        "SI / NO / SIN INFORMACION"
    ]):

        permitidos = {"SI","NO"}
        if "SE DESCONOCE" in observacion:
            permitidos.add("SE DESCONOCE")

        if "SIN INFORMACION" in observacion:
            permitidos.add("SIN INFORMACION")

        for idx, valor in serie.items():
            if pd.isna(valor):
                continue

            for v in split_valores(valor):
                if normalizar_texto(v) not in permitidos:
                    errores.append([ df.loc[idx,"_fila_excel"], campo, "VALOR", f"Solo permite {' / '.join(sorted(permitidos))}"])

def validar_tipo_entidad(df, serie, campo, observacion, errores):
    if "PERSONA NATURAL / PERSONA JURIDICA / COMUNIDAD ETNICA / GRUPO ARMADO" in observacion:
        permitidos = {"PERSONA NATURAL", "PERSONA JURIDICA", "COMUNIDAD ETNICA", "GRUPO ARMADO"}
        for idx, valor in serie.items():
            if pd.isna(valor):
                continue

            for v in split_valores(valor):
                if normalizar_texto(v) not in permitidos:
                    errores.append([ df.loc[idx,"_fila_excel"], campo, "VALOR", f"Solo permite {' / '.join(sorted(permitidos))}"])
                    
def validar_catalogo(df, serie, campo, catalogo, catalogos, admite_varios, errores):
    
    if str(catalogo).strip().upper() == "M_FUENTE": # Excluir catálogo M_FUENTE
        return
    
    if pd.notna(catalogo) and catalogo in catalogos:
        df_cat = catalogos[catalogo].copy()

        col_valor = obtener_columna_alfanumerica(df_cat)

        df_cat["_norm_"] = df_cat[col_valor].apply(normalizar_texto)
        valores_validos = set(df_cat["_norm_"].dropna())

        for idx, valor in serie.items():
            fila_excel = df.loc[idx,"_fila_excel"]
            if pd.isna(valor):
                continue

            valores = split_valores(valor) if admite_varios else [valor]

            for v in valores:
                if normalizar_texto(v) not in valores_validos:
                    errores.append([ fila_excel, campo, "CATALOGO", f"No existe '{v}' en {catalogo}"])

def validar_catalogo_dependiente(df, serie, campo, dependencias, errores):
    if campo not in dependencias:
        return

    dep = dependencias[campo]
    campo_padre = dep["campo_padre"]
    df_dep = dep["catalogo"].copy()
    col_padre = dep["col_padre"]
    col_hijo = dep["col_hijo"]

    df_dep["_padre"] = df_dep[col_padre].apply(normalizar_texto)
    df_dep["_hijo"] = df_dep[col_hijo].apply(normalizar_texto)

    for idx, valor in serie.items():
        fila_excel = df.loc[idx, "_fila_excel"]

        if pd.isna(valor) or pd.isna(df.loc[idx, campo_padre]):
            continue

        padre_valor = normalizar_texto(df.loc[idx, campo_padre])

        for v in split_valores(valor):
            v_norm = normalizar_texto(v)

            existe = ((df_dep["_padre"] == padre_valor) & (df_dep["_hijo"] == v_norm)).any()

            if not existe:
                errores.append([ fila_excel, campo, "CATALOGO DEPENDIENTE", f"'{v}' no tiene relación con el catálogo padre '{df.loc[idx, campo_padre]}'"])
                
def validar_condicionales(df, reglas_condicionales, errores):

    for regla in reglas_condicionales:
        col_si       = regla["si_columna"]
        val_si       = regla["si_valor"]
        col_entonces = regla["entonces_columna"]

        # ---------------------------------------------------------
        # Determinar el ID que controla la validación
        # ---------------------------------------------------------
        if col_entonces.startswith("H_"):
            id_col = "H_ID"

        elif col_entonces.startswith("E_"):
            id_col = "E_ID"

        elif col_entonces.startswith("I_"):
            id_col = "I_ID"

        else:
            id_col = None

        for idx, row in df.iterrows():

            if id_col is not None: # Si el bloque no existe, omitir la validación
                valor_id = str(row.get(id_col, "")).strip()

                if valor_id == "0":
                    continue

            fila_excel = row["_fila_excel"]

            if normalizar_texto(row.get(col_si)) == normalizar_texto(val_si):

                valor = row.get(col_entonces)
                if pd.isna(valor) or str(valor).strip() == "":
                    errores.append([ fila_excel, col_entonces, "CONDICIONAL", f"Debe contener información porque {col_si} = '{val_si}'"])


def validar_documentos(df, serie, campo, catalogos, errores):

    if campo != "E_DOCUMENTO(S)":
        return

    if "C_TIPO_DOCUMENTO" not in catalogos:
        return
    # ---------------------------------------------------------
    # Catálogo tipos válidos
    df_cat = catalogos["C_TIPO_DOCUMENTO"].copy()
    col_valor = obtener_columna_alfanumerica(df_cat)

    df_cat["_norm_"] = df_cat[col_valor].apply(normalizar_texto)
    tipos_validos = set(df_cat["_norm_"].dropna())

    # ---------------------------------------------------------
    # Formato general NUMERO:TIPO
    patron_general = re.compile(r"^\s*[\d\-]+\s*:\s*.+\s*$")

    # ---------------------------------------------------------
    # Tipos especiales con reglas estrictas
    tipos_especiales = { "cedula de ciudadania", "tarjeta de identidad", "nuip", "nip" }

    for idx, valor in serie.items():
        if pd.isna(valor):
            continue

        for v in split_valores(valor):
            v = str(v).strip()

            # -------------------------------------------------
            # Validar estructura NUMERO:TIPO
            if not patron_general.match(v):
                errores.append([ df.loc[idx, "_fila_excel"], campo, "FORMATO", "Formato debe ser NUMERO:TIPO"])
                continue

            numero, tipo = v.split(":", 1)
            numero = numero.strip()
            tipo = tipo.strip()
            tipo_norm = normalizar_texto(tipo)

            # -------------------------------------------------
            # Validar existencia en catálogo
            if tipo_norm not in tipos_validos:
                errores.append([ df.loc[idx, "_fila_excel"], campo, "CATALOGO", f"Tipo documento '{tipo}' no existe en C_TIPO_DOCUMENTO"])
                continue

            # -------------------------------------------------
            # Validaciones especiales
            if tipo_norm in tipos_especiales:

                # CEDULA / NUIP / NIP => 4 a 10 dígitos
                if tipo_norm in {"cedula de ciudadania", "nuip", "nip"}:
                    if not re.fullmatch(r"\d{4,10}", numero):
                        errores.append([ df.loc[idx, "_fila_excel"], campo, "FORMATO", f"Número inválido para {tipo}. Debe tener 4 a 10 dígitos" ])

                # TARJETA IDENTIDAD
                elif tipo_norm == "tarjeta de identidad":

                    valido = ( re.fullmatch(r"\d{4,12}", numero) or re.fullmatch(r"\d{6}-\d{4,6}", numero))

                    if not valido:
                        errores.append([ df.loc[idx, "_fila_excel"], campo, "FORMATO", f"Número inválido para {tipo}. Debe ser 4-12 dígitos o 123456-1234"])


def validar_admite_varios(df, serie, campo, observacion: None, errores):

    for idx, valor in serie.items():
        
        if pd.isna(valor): # ignorar NaN
            continue

        valor_str = str(valor).strip()

        if valor_str == "": # ignorar vacíos
            continue

        fila_excel = df.loc[idx,"_fila_excel"]

        # -----------------------------
        # VALIDACION SEPARADOS POR |
        # -----------------------------
        if campo in campos_pipe:

            # si tiene varios valores debe contener |
            if "|" not in valor_str and len(split_valores(valor_str)) > 1:
                errores.append([ fila_excel, campo, "FORMATO", "Valores deben estar separados por |"])

        # -----------------------------
        # VALIDACION FORMATO LISTA JSON  ["VAL1","VAL2","VAL3"]
        # -----------------------------
        elif campo in campos_lista:

            patron_lista = re.compile( r'^\[\s*"[^"]+"\s*(,\s*"[^"]+"\s*)*\]$')
            if not patron_lista.match(valor_str):
                errores.append([ fila_excel, campo, "FORMATO", 'Formato debe ser ["VAL1","VAL2","VAL3"]'])

        # -----------------------------
        # VALIDACION FORMATO LISTA JSON NUMERICA ["12","321","0025"]
        # -----------------------------
        elif campo in campos_lista_num:

            patron_lista_num = re.compile( r'^\[\s*"\d+"\s*(,\s*"\d+"\s*)*\]$')
            if not patron_lista_num.match(valor_str):
                errores.append([ fila_excel, campo, "FORMATO", 'Formato debe ser ["12","321","0025"] solo con valores numéricos'])

def validar_titulos(df, serie, campo, errores):
    if campo not in campos_titulo:
        return

    for idx, valor in serie.items():

        if pd.isna(valor):
            continue

        valor_str = str(valor).strip()
        if valor_str == "":
            continue

        fila_excel = df.loc[idx,"_fila_excel"]

        titulos = split_valores(valor_str) # separar por |

        for t in titulos:
            titulo = str(t).strip()
            if titulo == "":
                continue

            if len(titulo) > 255:
                errores.append([ fila_excel, campo, "LONGITUD_TITULO", "Cada título separado por | no debe superar 255 caracteres"])
                
def validar_titulo_catalogo(df, serie, campo, M_FUENTE, errores):
    if campo not in campos_titulo:
        return

    # preparar catálogo
    df_cat = M_FUENTE.copy()
    df_cat["_titulo"] = df_cat["TITULO"].apply(normalizar_texto)
    titulos_validos = set(df_cat["_titulo"].dropna())

    for idx, valor in serie.items():
        if pd.isna(valor):
            continue

        valor_str = str(valor).strip()
        if valor_str == "":
            continue

        fila_excel = df.loc[idx, "_fila_excel"]
        
        for titulo in split_valores(valor_str): # admite varios separados por |
            titulo_norm = normalizar_texto(titulo)
            if titulo_norm not in titulos_validos:
                errores.append([ fila_excel, campo, "CATALOGO_TITULO", f"Título '{titulo}' no existe en las FUENTES"])

def validar_i_id(df, reglas, errores):

    for regla in reglas:
        col_resultado = regla["col_resultado"]
        col_hecho     = regla["col_hecho"]
        col_entidad   = regla["col_entidad"]

        columnas = {col_resultado, col_hecho, col_entidad}
        if not columnas.issubset(df.columns):
            continue

        for idx, row in df.iterrows():
            fila_excel  = row["_fila_excel"]
            val_hecho   = row[col_hecho]
            val_entidad = row[col_entidad]
            resultado   = row[col_resultado]

            if pd.isna(val_hecho) or pd.isna(val_entidad):
                continue

            esperado = f"{str(val_hecho).strip()}-{str(val_entidad).strip()}"

            # ---------------- resultado vacío
            if pd.isna(resultado) or str(resultado).strip() == "":
                errores.append([ fila_excel, col_resultado, "CONCATENACION", f"Debe ser {col_hecho}-{col_entidad}"])
                continue

            resultado_str = str(resultado).strip()

            # ---------------- validar concatenación completa
            if resultado_str != esperado:
                errores.append([ fila_excel, col_resultado, "CONCATENACION", f"Debe ser {col_hecho}-{col_entidad}"])

            # ---------------- validar partes internas SIEMPRE
            if "-" in resultado_str:
                parte_hecho, parte_entidad = resultado_str.split("-", 1)

                if parte_hecho != str(val_hecho).strip():
                    errores.append([ fila_excel, col_hecho, "CONCATENACION", f"No coincide con {col_resultado}"])

                if parte_entidad != str(val_entidad).strip():
                    errores.append([ fila_excel, col_entidad, "CONCATENACION", f"No coincide con {col_resultado}"])
                    
def validar_duplicidad_por_id_prefijo(df, errores, id_col="E_ID", prefijo="E_"):

    if id_col not in df.columns:
        return

    columnas = [
        c for c in df.columns
        if c.startswith(prefijo) and c != id_col and c != "_fila_excel"]

    if not columnas:
        return

    df_tmp = df[[id_col, "_fila_excel"] + columnas].copy()

    # ------------------------------------------------------------------
    # Para H_ID no se valida cuando el identificador es 0
    # ------------------------------------------------------------------
    if id_col == "H_ID":
        df_tmp = df_tmp[ df_tmp[id_col].fillna("").astype(str).str.strip() != "0"]

        if df_tmp.empty:
            return
        
    # normalizar
    for col in columnas:
        df_tmp[col] = (
            df_tmp[col]
            .fillna("")
            .astype(str)
            .str.strip()
            .apply(normalizar_texto)
        )

    # firma única
    df_tmp["_firma"] = df_tmp[columnas].agg("|".join, axis=1)

    # ids con más de una versión
    resumen = (
        df_tmp.groupby(id_col)["_firma"]
        .nunique()
        .reset_index(name="firmas")
    )

    inconsistentes = resumen[resumen["firmas"] > 1][id_col]

    # nombre amigable según prefijo
    nombres = {
        "E_": "ENTIDAD",
        "H_": "HECHO",
        "I_": "INFORMACION"
    }

    nombre_tipo = nombres.get(prefijo, prefijo.replace("_", ""))

    for eid in inconsistentes:

        filas = ( df_tmp[df_tmp[id_col] == eid]["_fila_excel"].astype(str).tolist() )

        errores.append([ ",".join(filas), id_col, f"INCONSISTENCIA_{nombre_tipo}", f"{nombre_tipo} con {id_col} '{eid}' presenta diferencias entre registros"])

def obtener_df_bloque(df, campo, mask_h, mask_e, mask_i):

    if campo.startswith("H_"):
        return df[mask_h]

    if campo.startswith("E_"):
        return df[mask_e]

    if campo.startswith("I_"):
        return df[mask_i]

    return df

config  = pd.read_excel(archivo_config, sheet_name="HECHO_ENTIDAD_INF")

df = pd.read_parquet(archivo_datos)

columnas_a_convertir = [
    "H_CIUDAD", "H_DEPARTAMENTO", "H_CENTRO_POBLADO", "E_DEPARTAMENTO_NACIMIENTO",
    "E_CIUDAD_NACIMIENTO", "E_CENTRO_POBLADO_NACIMIENTO", "E_DEPARTAMENTO_RESIDENCIA",
    "E_CIUDAD_RESIDENCIA", "E_CENTRO_POBLADO_RESIDENCIA", "E_DEPARTAMENTO_NOTIFICACION",
    "E_CIUDAD_NOTIFICACION", "E_CENTRO_POBLADO_NOTIFICACION", "E_DEPARTAMENTO_COMUNIDAD_MUNICIPAL",
    "E_CIUDAD_COMUNIDAD_MUNICIPAL", "E_CENTRO_POBLADO_COMUNIDAD_MUNICIPAL",
    "I_CENTRO_POBLADO", "I_DEPARTAMENTO_UBICACION", "I_CIUDAD_UBICACION",
    "I_CENTRO_POBLADO_UBICACION"
]

df[columnas_a_convertir] = df[columnas_a_convertir].astype("string")

df["_fila_excel"] = df.index + 2 # guardar numero fila excel (sumamos 2 por encabezado Excel)
errores = []

# --------------------------- --------------------------- --------------------------- 
#                                   VALIDACIONES
# --------------------------- --------------------------- ---------------------------
logging.warning("validate")

validar_campos_existentes(config, df, errores)
validar_rango_fechas(df, errores)

# --------------------------- --------------------------- ---------------------------
#                         DATAFRAMES POR BLOQUE (ID != 0)
# --------------------------- --------------------------- ---------------------------

mask_h = df["H_ID"].fillna("").astype(str).str.strip() != "0"
mask_e = df["E_ID"].fillna("").astype(str).str.strip() != "0"
mask_i = df["I_ID"].fillna("").astype(str).str.strip() != "0"

df_h = df.loc[mask_h].copy()
df_e = df.loc[mask_e].copy()
df_i = df.loc[mask_i].copy()

for _, row in config.iterrows():
    campo       = row["CAMPO"]
    tipo        = str(row["TIPO DE CAMPO"]).lower()
    longitud    = row["LONGITUD"]
    obligatorio = str(row["OBLIGATORIO"]).strip().upper() == "S"
    catalogo    = row["CATALOGO"]
    observacion = str(row.get("OBSERVACION","")).upper()

    if campo not in df.columns:
        continue

    # Seleccionar el DataFrame correspondiente
    if campo.startswith("H_"):
        df_val = df_h

    elif campo.startswith("E_"):
        df_val = df_e

    elif campo.startswith("I_"):
        df_val = df_i

    else:
        df_val = df

    if df_val.empty:
        continue

    serie = df_val[campo]
    admite_varios = "ADMITE VARIOS" in observacion

    validar_obligatorio(df_val, serie, campo, obligatorio, errores)
    validar_longitud(df_val, serie, campo, longitud, errores)
    validar_titulos(df_val, serie, campo, errores)
    validar_titulo_catalogo(df_val, serie, campo, M_FUENTE, errores)
    validar_fecha(df_val, serie, campo, tipo, errores)
    validar_documentos(df_val, serie, campo, catalogos, errores)
    validar_si_no(df_val, serie, campo, observacion, errores)
    validar_admite_varios(df_val, serie, campo, observacion, errores)
    validar_tipo_entidad(df_val, serie, campo, observacion, errores)

    validar_catalogo( df_val, serie, campo, catalogo, catalogos, admite_varios, errores)
    validar_catalogo_dependiente( df_val, serie, campo, dependencias, errores)

# --------------------------- --------------------------- ---------------------------
#                                   VALIDACIONES GLOBALES
# --------------------------- --------------------------- ---------------------------

validar_i_id(df_i, reglas_i_id, errores)
validar_condicionales(df, reglas_condicionales, errores)

# Duplicidad por bloque
if not df_e.empty:
    validar_duplicidad_por_id_prefijo( df_e, errores, id_col="E_ID", prefijo="E_")

if not df_h.empty:
    validar_duplicidad_por_id_prefijo( df_h, errores, id_col="H_ID", prefijo="H_")

if not df_i.empty:
    validar_duplicidad_por_id_prefijo( df_i, errores, id_col="I_ID", prefijo="I_")

# --------------------------- --------------------------- --------------------------- 
#                                    RESULTADOS
# --------------------------- --------------------------- --------------------------- 
df_errores = pd.DataFrame( errores, columns=["fila_excel","campo","tipo_error","descripcion"])
# df_errores['fila_excel'] = df_errores['fila_excel'].astype('Int64')  # Convertir a entero, permitiendo NA
df_errores["fila_excel"] = df_errores["fila_excel"].astype("string")

# --------------------------- --------------------------- --------------------------- 
#                                   RESUMEN DE ERRORES
# --------------------------- --------------------------- --------------------------- 
resumen = (
    df_errores
    .groupby(["campo","tipo_error"])
    .agg(
        cantidad=("fila_excel","count"),
        filas=("fila_excel", lambda x: ",".join(map(str, sorted(set(x)))))
    )
    .reset_index()
)
                      
logging.warning("export")

with pd.ExcelWriter(result_path, engine="openpyxl") as writer:
    resumen.to_excel(writer, sheet_name="RESUMEN", index=False)
    df_errores.to_excel(writer, sheet_name="DETALLE", index=False)

