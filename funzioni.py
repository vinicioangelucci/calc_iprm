import math
import pandas as pd


def calcola_velocita_monte(row,k):

    q_monte=(row.Qimp_m3_h*((1-0.002*row.Pmin_bar)/(row.Pmin_bar+1)))/3600
    sez_monte = ((math.pi * row.Dnmonte_mm ** 2) / 4) / 1000000
    velocita_monte = round(q_monte / sez_monte * k, 2)

    descrizione=""
    if velocita_monte<30:
        descrizione="VERIFICATO"
    else:
        descrizione="NON VERIFICATO"
    return velocita_monte,descrizione

def calcola_dp_monte(row,k1):
    Dp = round(1000 * (
    row.Pmin_bar+1 - math.sqrt(((row.Pmin_bar+1) ** 2) - 25.24 * row.Lmonte_m * (row.Qimp_m3_h ** 1.82) * ((row.Dnmonte_mm*k1)) ** (-4.82))), 2)
    descrizione = ""
    if Dp < 100:
        descrizione = 'VERIFICATO'
    else:
        descrizione = 'NON VERIFICATO'
    return Dp,descrizione


def calcola_dp_linea(row,k1):
    Dp = round(1000 * (
        ((row.Pmin_bar+1)+(row.Pvalle_bar+1))/2 - math.sqrt(((((row.Pmin_bar+1)+(row.Pvalle_bar+1))/2) ** 2) - 25.24 * row.lunghezza_linea * (row.Qimp_m3_h ** 1.82) * ((row.DNlinea_rid_mm *k1))** (-4.82))), 2)
    descrizione = ""
    if Dp < 100:
        descrizione = 'VERIFICATO'
    else:
        descrizione = 'NON VERIFICATO'
    return Dp, descrizione




def calcola_velocita_uscita(row,k):
    q_valle = (row.Qimp_m3_h * ((1 - 0.002 * (row.Pvalle_bar )) / (row.Pvalle_bar+1))) / 3600
    sez_valle = ((row.Dnvalle_mm ** 2 * math.pi) / 4) / 1000000
    velocita_valle = round(q_valle / sez_valle * k, 2)
    descrizione = ""
    if velocita_valle < 25:
        descrizione = "VERIFICATO"
    else:
        descrizione = "NON VERIFICATO"
    return velocita_valle, descrizione



def calcola_portata_max_filtro(row):
    q_massima_filtro = round(row.A_filtro_m2 * 0.33 *row.N_cartucce* (row.Pmin_bar+1)  * 3600, 2)
    descrizione=""
    if q_massima_filtro>(row.Qimp_m3_h/(row.N_linee-1)):
        descrizione='VERIFICATO'
    else:
        descrizione='NON VERIFICATO'

    return q_massima_filtro,descrizione

def calcola_velocita_linea_filtro(row,k):
    sez_filtro = ((row.Dnfiltro_mm ** 2 * math.pi) / 4) / 1000000
    q_monte = ((row.Qimp_m3_h/(row.N_linee-1))* ((1 - 0.002 * row.Pmax_bar) / (row.Pmax_bar+1))) / 3600

    velocita_filtro = round(q_monte/ sez_filtro * k, 2)
    descrizione = ""
    if velocita_filtro < 30:
        descrizione = "VERIFICATO"
    else:
        descrizione = "NON VERIFICATO"
    return velocita_filtro, descrizione

def calcola_velocita_linea_riduttore(row,k):
    sez_linea_rid = ((row.DNlinea_rid_mm ** 2 * math.pi) / 4) / 1000000
    q_monte = ((row.Qimp_m3_h/(row.N_linee))* ((1 - 0.002 * row.Pmin_bar) / (row.Pmin_bar+1))) / 3600

    velocita_linea_rid = round(q_monte/ sez_linea_rid * k, 2)
    descrizione = ""
    if velocita_linea_rid < 30:
        descrizione = "VERIFICATO"
    else:
        descrizione = "NON VERIFICATO"
    return velocita_linea_rid, descrizione




def calcolo_c(row):

    if row.N_linee==2:
        c = round(0.000216 * (row.Qimp_m3_h*(row.N_linee-1) * row.salto_entalpico),1)
        descrizione=''

        if c < (row.Potenzialità_1_kw+row.Potenzialità_2_kw)*(row.N_linee)/row.N_linee:
            descrizione = 'VERIFICATO'
        else:
            descrizione = 'NON VERIFICATO'
    else:
        c = round(0.000216 * (row.Qlin_m3_h * row.salto_entalpico),1)
        descrizione=''

        if c*(row.N_linee-1)/row.N_linee < (row.Potenzialità_1_kw+row.Potenzialità_2_kw+row.Potenzialità_3_kw):
            descrizione = 'VERIFICATO'
        else:
            descrizione = 'NON VERIFICATO'


    return c,descrizione



def calcolo_capacita_riduttore(row):

    portata_max_rid = round(1.17*0.55 * row.cg * (row.Pmin_bar+1) * math.sin(
        ((3417 / row.c1) * math.sqrt((row.Pmin_bar - row.Pvalle_bar) / (row.Pmin_bar+1)) * math.pi / 180)), 2)

    #portata_max_rid=0.55*row.cg*(row.Pmin_bar+1)

    descrizione = ""
    if portata_max_rid > (row.Qimp_m3_h / (row.N_linee - 1)):
        descrizione = 'VERIFICATO'
    else:
        descrizione = 'NON VERIFICATO'

    return portata_max_rid,descrizione

def calcolo_capacita_riduttore1(row):

    #portata_max_rid = round(0.55 * row.cg * (row.Pmin_bar+1) * math.sin(
        #((3417 / row.c1) * math.sqrt((row.Pmin_bar - row.Pvalle_bar) / (row.Pmin_bar+1)) * math.pi / 180)), 2)

    portata_max_rid=round((13.57/math.sqrt((0.6*278)))*row.cg*(row.Pmin_bar/2),2)

    descrizione = ""
    if portata_max_rid > (row.Qimp_m3_h / (row.N_linee - 1)):
        descrizione = 'VERIFICATO'
    else:
        descrizione = 'NON VERIFICATO'

    return portata_max_rid,descrizione

def calcolo_velocita_riduttore(row,k2):

    q_valle = (row.Qimp_m3_h * ((1 - 0.002 * (row.Pmin_bar)) / (row.Pmin_bar+1))) / 3600
    q_linea=q_valle/(row.N_linee-1)

    sez_riduttore = ((row.Dnrid_mm ** 2 * math.pi) / 4) / 1000000
    velocita_riduttore = round(q_linea / sez_riduttore * k2, 2)
    descrizione = ""
    if velocita_riduttore < 130:
        descrizione = "VERIFICATO"
    else:
        descrizione = "NON VERIFICATO"

    return velocita_riduttore,descrizione

def calcolo_velocita_riduttore_valle(row,k):
    q_valle = (row.Qimp_m3_h * ((1 - 0.002 * row.Pvalle_bar ) / (row.Pvalle_bar + 1))) / 3600
    q_linea = q_valle / (row.N_linee - 1)
    sez_riduttore_valle = ((row.Dntub_rid_val ** 2 * math.pi) / 4) / 1000000
    velocita_riduttore_valle = round(q_linea / sez_riduttore_valle * k, 2)
    if velocita_riduttore_valle < 25:
        descrizione = "VERIFICATO"
    else:
        descrizione = "NON VERIFICATO"

    return velocita_riduttore_valle,descrizione



def calcolo_diamtero_teorico(row):
    diametro_teorico = round(row.Dnvalle_mm * 1.00 / 10, 2)
    if diametro_teorico <= (row.Dnvalv_scarico_mm):
        descrizione = "VERIFICATO"
    else:
        descrizione = "NON VERIFICATO"


    return diametro_teorico,descrizione

def calcolo_portata_massima_volumetrica(row,k):

    q_max_contatore = round(1.025 * row.Qero_m3_h / (row.Pmisura_bar+1),1)

    sez_monte = ((math.pi * row.Dnmis_2_mm ** 2) / 4) / 1000000
    v_max_valle = 25
    q = v_max_valle * sez_monte * k
    q_max_mis_t = round((q / ((1 - 0.002 * row.Pmisura_bar) / (row.Pmisura_bar + 1))) * 3600, 1)
    descrizione = ""
    if q_max_mis_t  > q_max_contatore:
        descrizione = 'VERIFICATO'
    else:
        descrizione = 'NON VERIFICATO'

    return q_max_mis_t,descrizione


def calcolo_portata_massima_ventirumetria(row):
    q_max_venturi= round(1.3 * row.Qimp_m3_h,1)
    q_max_imp = round(1.025 * row.Qero_m3_h / (row.Pmisura_bar + 1), 1)

    descrizione = ""
    if q_max_venturi > q_max_imp:
        descrizione = 'VERIFICATO'
    else:
        descrizione = 'NON VERIFICATO'
    return q_max_venturi, descrizione


def calcolo_portata_bypass(row):
    sez_bypass = ((row.BYPASS_DN** 2 * math.pi) / 4) / 1000000
    q_massima_bypass = round((30 * sez_bypass * (row.Pmisura_bar+1) / (1 - 0.002 * (row.Pmisura_bar) )) * 3600, 1)
    q_max_contatore = round(1.025 * row.Qero_m3_h / (row.Pmisura_bar + 1), 1)
    descrizione = ""
    if q_massima_bypass > q_max_contatore:
        descrizione = 'VERIFICATO'
    else:
        descrizione = 'NON VERIFICATO'
    return q_massima_bypass,descrizione


def calcolo_velocita_bypass(row,k):
    q_valle = round((row.Qimp_m3_h * ((1 - 0.002 * (row.Pmisura_bar )) / (row.Pmisura_bar + 1))) / 3600,1)
    sez_bypass = ((row.BYPASS_DN ** 2 * math.pi) / 4) / 1000000
    velocita_bypass = round(q_valle / sez_bypass * k, 2)
    descrizione = ""
    if velocita_bypass < 30:
        descrizione = "VERIFICATO"
    else:
        descrizione = "NON VERIFICATO"

    return velocita_bypass,descrizione

def calcolo_potenzialita(row):
    m = round(0.00024 * row.salto_entalpico * row.Qimp_m3_h, 2)
    descrizione = ''

    if m < (row.Potenzialità_caldaia_1_kw + row.Potenzialità_caldaia_2_kw + row.Potenzialità_caldaia_3_kw):
        descrizione = 'VERIFICATO'
    else:
        descrizione = 'NON VERIFICATO'

    return m, descrizione



def calcolo_q_valle(row,k):
    q_valle = round((row.Qimp_m3_h * ((1 - 0.002 * row.Pvalle_bar ) / (row.Pvalle_bar ))) / 3600,1)

    return q_valle

def calcola_q_max_monte(row,k):
    sez_monte = ((math.pi * row.Dnmonte_mm ** 2) / 4) / 1000000
    v_max_monte=30
    q=v_max_monte*sez_monte*k
    q_max_monte_t=round((q/((1-0.002*row.Pmax_bar)/(row.Pmax_bar+1)))*3600,1)

    descrizione = ""
    if q_max_monte_t > (row.Qimp_m3_h / (row.N_linee - 1)):
        descrizione = 'VERIFICATO'
    else:
        descrizione = 'NON VERIFICATO'

    return q_max_monte_t, descrizione


def calcola_q_max_valle(row,k):
    sez_monte = ((math.pi * row.Dnvalle_mm** 2) / 4) / 1000000
    v_max_valle=25
    q=v_max_valle*sez_monte*k
    q_max_valle_t=round((q/((1-0.002*row.Pvalle_bar)/(row.Pvalle_bar+1)))*3600,1)

    descrizione = ""
    if q_max_valle_t > (row.Qimp_m3_h / (row.N_linee - 1)):
        descrizione = 'VERIFICATO'
    else:
        descrizione = 'NON VERIFICATO'

    return q_max_valle_t, descrizione



def calcola_q_max_misura_vol(row,k):
    sez_monte = ((math.pi * row.Dnmis_2_mm** 2) / 4) / 1000000
    v_max_valle=25
    q=v_max_valle*sez_monte*k
    q_max_mis_t=round((q/((1-0.002*row.Pmisura_bar)/(row.Pmisura_bar+1)))*3600,1)
    return q_max_mis_t

def calcola_portata_contatore(row):
    qc=round(row.classe_q*(row.Pmisura_bar+1),1)
    q_max_contatore = round(1.025 * row.Qero_m3_h , 1)
    descrizione=""
    if qc > (q_max_contatore):
        descrizione = 'VERIFICATO'
    else:
        descrizione = 'NON VERIFICATO'

    return descrizione




def calcola_q_max_scambiatore_t(row,k):
    sez_monte = ((math.pi * row.Dnmonte_mm ** 2) / 4) / 1000000
    v_max_monte = 30
    q = v_max_monte * sez_monte * k
    q_max_scambiatore_t = round(((q * ((1 - 0.002 * row.Pmax_bar) / (row.Pmax_bar + 1))) * 3600)*(row.N_linee-1),1)

    descrizione = ""
    if q_max_scambiatore_t > (row.Qimp_m3_h / (row.N_linee - 1)):
        descrizione = 'VERIFICATO'
    else:
        descrizione = 'NON VERIFICATO'

    return q_max_scambiatore_t, descrizione











