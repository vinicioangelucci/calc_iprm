import math
import pandas as pd
import funzioni
import xlsxwriter

dati = pd.read_excel("C:\\dimensionamenti\input\\dati_ingr1_nuovi2.xlsx")
dati = dati.rename(columns=lambda x: x.replace(' ', '_'))
dati = dati.rename(columns=lambda x: x.replace('/', '_'))
dati = dati.rename(columns=lambda x: x.replace('-', '_'))
dati = dati.fillna(10)
dati = dati.set_index('codice_remi')
# dati.style.set_properties(**{'text-align': 'left'},in_place=True)

area_filtrante = {'GC01': 0.066, 'GC02': 0.12, 'GC10': 0.12, 'GC20': 0.22, 'GC0,5': 0.06, 'G1': 0.125, 'G1,5': 0.23,
                  'G2': 0.47, 'G2,5': 0.725, 'G3': 0.95, 'G4': 1.45, 'G5': 2.3, 'G6': 4.2}

classi_contatore = {'G65':100,'G160':250,'G250': 1000, 'G400': 650, 'G650': 1000, 'G1000': 1600, 'G1600': 2500, 'G2500': 4000, 'G4000': 6500,
                    'G6500': 10000}


for i in dati.index:

    if dati.tipo_misura.loc[i] == 'volumetrico':

        k = 0.977
        k1=0.977
        k2=0.977
        start = 10

        dati['velocità_monte'] = dati.apply(funzioni.calcola_velocita_monte, args=[k], axis=1).apply(lambda x: x[0])
        dati['velocità_monte_ver']=dati.apply(funzioni.calcola_velocita_monte, args=[k], axis=1).apply(lambda x: x[1])

        dati['DP_monte'] = dati.apply(funzioni.calcola_dp_monte,args=[k1],axis=1).apply(lambda x: x[0])
        dati['DP_monte_ver'] = dati.apply(funzioni.calcola_dp_monte,args=[k1], axis=1).apply(lambda x: x[1])

        dati['DP_valle'] = dati.apply(funzioni.calcola_dp_linea,args=[k1], axis=1).apply(lambda x: x[0])
        dati['DP_valle_ver'] = dati.apply(funzioni.calcola_dp_linea,args=[k1], axis=1).apply(lambda x: x[1])


        dati['velocita_uscita'] = dati.apply(funzioni.calcola_velocita_uscita, args=[k], axis=1).apply(lambda x: x[0])
        dati['velocita_uscita_ver'] = dati.apply(funzioni.calcola_velocita_uscita, args=[k], axis=1).apply(lambda x: x[1])


        dati['portata_max_filtro'] = dati.apply(funzioni.calcola_portata_max_filtro, axis=1).apply(lambda x: x[0])
        dati['portata_max_filtro_ver'] = dati.apply(funzioni.calcola_portata_max_filtro, axis=1).apply(lambda x: x[1])


        dati['velcità_linea_filtro'] = dati.apply(funzioni.calcola_velocita_linea_filtro, args=[k], axis=1).apply(lambda x: x[0])
        dati['velcità_linea_filtro_ver'] = dati.apply(funzioni.calcola_velocita_linea_filtro, args=[k], axis=1).apply(lambda x: x[1])


        dati['pot1'] = dati.apply(funzioni.calcolo_c, axis=1).apply(lambda x: x[0])
        dati['pot2'] = dati.apply(funzioni.calcolo_c, axis=1).apply(lambda x: x[0])
        dati['pot3'] = dati.apply(funzioni.calcolo_c, axis=1).apply(lambda x: x[0])
        dati['pot1_ver'] = dati.apply(funzioni.calcolo_c, axis=1).apply(lambda x: x[1])
        dati['pot2_ver'] = dati.apply(funzioni.calcolo_c, axis=1).apply(lambda x: x[1])
        dati['pot3_ver'] = dati.apply(funzioni.calcolo_c, axis=1).apply(lambda x: x[1])

        dati['V_serp_scambiatore'] = dati.apply(funzioni.calcola_velocita_linea_filtro, args=[k], axis=1).apply(lambda x: x[0])
        dati['V_serp_scambiatore_ver'] = dati.apply(funzioni.calcola_velocita_linea_filtro, args=[k], axis=1).apply(
            lambda x: x[1])

        dati['Q_max_rid'] = dati.apply(funzioni.calcolo_capacita_riduttore, axis=1).apply(lambda x: x[0])
        dati['Q_max_rid_ver'] = dati.apply(funzioni.calcolo_capacita_riduttore, axis=1).apply(lambda x: x[1])


        dati['v_riduttore'] = dati.apply(funzioni.calcolo_velocita_riduttore, args=[k], axis=1).apply(lambda x: x[0])
        dati['v_riduttore_ver'] = dati.apply(funzioni.calcolo_velocita_riduttore, args=[k], axis=1).apply(lambda x: x[1])

        dati['v_linea_riduttore'] = dati.apply(funzioni.calcola_velocita_linea_riduttore, args=[k], axis=1).apply(lambda x: x[0])
        dati['v_linea_riduttore_ver'] = dati.apply(funzioni.calcola_velocita_linea_riduttore, args=[k], axis=1).apply(
            lambda x: x[1])

        dati['v_tub_rid_valle'] = dati.apply(funzioni.calcolo_velocita_riduttore_valle, args=[k], axis=1).apply(lambda x: x[0])
        dati['v_tub_rid_valle_ver'] = dati.apply(funzioni.calcolo_velocita_riduttore_valle, args=[k], axis=1).apply(
            lambda x: x[1])

        dati['diametro_scarico'] = dati.apply(funzioni.calcolo_diamtero_teorico, axis=1).apply(lambda x: x[0])
        dati['diametro_scarico_ver'] = dati.apply(funzioni.calcolo_diamtero_teorico, axis=1).apply(lambda x: x[1])

        # dati['Q_max_misura'] = dati.apply(funzioni.calcolo_portata_massima_volumetrica, axis=1)
        dati['Q_max_misura'] = dati.apply(funzioni.calcolo_portata_massima_volumetrica, args=[k2], axis=1).apply(lambda x: x[0])
        dati['Q_max_misura_ver'] = dati.apply(funzioni.calcolo_portata_massima_volumetrica, args=[k2], axis=1).apply(lambda x: x[1])


        dati['classe_q']=dati.classe_2.map(classi_contatore)

        dati['Q_max_contatore_ver'] = dati.apply(funzioni.calcola_portata_contatore, axis=1)




        dati['Q_max_venturi'] = dati.apply(funzioni.calcolo_portata_massima_ventirumetria, axis=1)

        dati['v_bypass'] = dati.apply(funzioni.calcolo_velocita_bypass, args=[k], axis=1).apply(lambda x: x[0])
        dati['v_bypass_ver'] = dati.apply(funzioni.calcolo_velocita_bypass, args=[k], axis=1).apply(lambda x: x[1])

        dati['q_va'] = dati.apply(funzioni.calcolo_q_valle, args=[k], axis=1)

        dati['q_max_bypass'] = dati.apply(funzioni.calcolo_portata_bypass, axis=1).apply(lambda x: x[0])
        dati['q_max_bypass_ver'] = dati.apply(funzioni.calcolo_portata_bypass, axis=1).apply(lambda x: x[1])


        dati['m'] = dati.apply(funzioni.calcolo_potenzialita, axis=1).apply(lambda x: x[0])
        dati['m_ver'] = dati.apply(funzioni.calcolo_potenzialita, axis=1).apply(lambda x: x[1])

        dati['riferimento_REMI'] = dati.apply(lambda row: 'REMI ' + row.descrizione, axis=1)

        dati['Q_max_teorica_monte'] = dati.apply(funzioni.calcola_q_max_monte, args=[k2], axis=1).apply(lambda x: x[0])
        dati['Q_max_teorica_monte_ver'] = dati.apply(funzioni.calcola_q_max_monte, args=[k2], axis=1).apply(lambda x: x[1])

        dati['Q_max_teorica_valle'] = dati.apply(funzioni.calcola_q_max_valle, args=[k2], axis=1).apply(lambda x: x[0])
        dati['Q_max_teorica_valle_ver'] = dati.apply(funzioni.calcola_q_max_valle, args=[k2], axis=1).apply(lambda x: x[1])

        dati['Q_max_teorica_scambiatore'] = dati.apply(funzioni.calcola_q_max_scambiatore_t, args=[k2], axis=1).apply(lambda x: x[0])
        dati['Q_max_teorica_scambiatore_ver'] = dati.apply(funzioni.calcola_q_max_scambiatore_t, args=[k2], axis=1).apply(
            lambda x: x[1])




        df_dati_impianti = {}
        #dati della testata
        dati_testata = {}
        dati_impianto = {}
        dati_impiantov = {}
        dati_verifica = {}


        dati_portate_teoriche_max = {}

        dati_portata = {}
        dati_portata_v={}
        dati_potenzialita = {}
        dati_capacita = {}
        dati_scarico = {}
        dati_misura = {}
        dati_misura_contatore={}
        dati_misura1 = {}
        dati_pote = {}

        dati_impianto['codice REMI'] = i

        dati_testata['riferimento_remi'] = 'REMI ' + str(i) + ' ' + dati.descrizione.loc[i]
        dati_testata['impianto'] = dati.codice_impianto.loc[i]
        dati_testata['comune'] = dati.comune.loc[i]
        df_dati_impianti['rif_remi']='REMI ' + ' ' + dati.descrizione.loc[i]

        dati_impianto['Q_imp'] = dati.Qimp_m3_h.loc[i].astype(str) + ' ' * 2 + "m3/h"
        dati_impianto['Q_ero'] = dati.Qero_m3_h.loc[i].astype(str) + ' ' * 2 + "m3/h"
        dati_impianto['Q_linea_filtro'] = dati.Qlin_m3_h.loc[i].astype(str) + ' ' * 2 + "m3/h"
        dati_impianto['Q_linea_scambiatore'] = dati.Qlin_m3_h.loc[i].astype(str) + ' ' * 2 + "m3/h"
        dati_impianto['Q_linea_riduttore'] = dati.Qlin_m3_h.loc[i].astype(str) + ' ' * 2 + "m3/h"
        dati_impianto['Coeff sicurezza'] = "10%"
        dati_impianto['Pressione minima'] = dati.Pmin_bar.loc[i].astype(str) + ' ' * 2 + "bar"
        dati_impianto['Pressione minima assoluta'] = dati.Pmin_bar.loc[i] + 1
        dati_impianto['Pressione minima assoluta'] = dati_impianto['Pressione minima assoluta'].astype(
            str) + ' ' * 2 + "bar"
        dati_impianto['Pressione massima'] = dati.Pmax_bar.loc[i].astype(str) + ' ' * 2 + "bar"
        dati_impianto['Pressione massima assoluta'] = dati.Pmax_bar.loc[i] + 1
        dati_impianto['Pressione massima assoluta'] = dati_impianto['Pressione massima assoluta'].astype(
            str) + ' ' * 2 + "bar"
        dati_impianto['Pressione preriscaldo '] = dati.Pmax_bar.loc[i].astype(str) + ' ' * 2 + "bar"
        dati_impianto['Pressione preriscaldo assoluta'] = dati.Pmax_bar.loc[i] + 1
        dati_impianto['Pressione preriscaldo assoluta'] = dati_impianto['Pressione preriscaldo assoluta'].astype(
            str) + ' ' * 2 + "bar"
        dati_impianto['Pressione misura'] = dati.Pmisura_bar.loc[i].astype(str) + ' ' * 2 + "bar"
        dati_impianto['Pressione misura assoluta'] = dati.Pmisura_bar.loc[i] + 1
        dati_impianto['Pressione misura assoluta'] = dati_impianto['Pressione misura assoluta'].astype(
            str) + ' ' * 2 + "bar"
        dati_impianto['Numero linee filtro'] = dati.N_linee.loc[i]
        dati_impianto['Numero linee scambiatore'] = dati.N_linee.loc[i]
        dati_impianto['Numero linee preriscaldo'] = dati.N_linee.loc[i]
        dati_impianto['Altezza livello mare'] = dati.altezzas_lm.loc[i].astype(str) + ' ' * 2 + "m"
        dati_impianto['Lunghezza monte'] = dati.Lmonte_m.loc[i].astype(str) + ' ' * 2 + "m"
        dati_impianto['DNmonte'] = dati.Dnmonte_mm.loc[i].astype(str) + ' ' * 2 + "mm"
        dati_impianto['DNvalle riduzione'] = dati.DNlinea_rid_mm.loc[i].astype(int).astype(str) + ' ' * 2 + "mm"
        dati_impianto['DNvalle'] = dati.Dnvalle_mm.loc[i].astype(str) + ' ' * 2 + "mm"
        dati_impianto['PN/ANSI sino 1a valvola valle regolatori'] = dati.PN_D_U_ANSI_monte.loc[i].astype(
            str) + ' ' * 2 + "ANSI"
        dati_impianto['PN/ANSI dopo 1a valvola valle regolatori'] = dati.PN_D_U_ANSI_monte.loc[i].astype(
            str) + ' ' * 2 + "ANSI"
        dati_impianto['FILTRO'] = ''
        dati_impianto['Tipo filtro'] = dati.tipo_filtro.loc[i]
        dati_impianto['N° cartucce'] = dati.N_cartucce.loc[i]
        dati_impianto['DNfiltro'] = dati.Dnfiltro_mm.loc[i].astype(str) + ' ' * 2 + "mm"
        dati_impianto['Pressione bollo filtro '] = dati.Pbollo_filtro_bar.loc[i].astype(str) + ' ' * 2 + "bar"
        dati_impianto['A'] = dati.A_filtro_m2.loc[i].astype(str) + ' ' * 2 + "m2"
        dati_impianto['SCAMBIATORI'] = ''
        dati_impianto['DNscambiatore'] = dati.DNscambiatore.loc[i].astype(str) + ' ' * 2 + "mm"
        dati_impianto['Potenzialita1'] = dati.Potenzialità_1_kw.loc[i].astype(str) + ' ' * 2 + "KW"
        dati_impianto['Potenzialita2'] = dati.Potenzialità_2_kw.loc[i].astype(str) + ' ' * 2 + "KW"
        dati_impianto['Potenzialita3'] = dati.Potenzialità_3_kw.loc[i].astype(str) + ' ' * 2 + "KW"
        dati_impianto['Pressione bollo filtro '] = dati.Pbollo_scamb.loc[i].astype(str) + ' ' * 2 + "bar"
        dati_impianto['RIDUTTORE'] = ''
        dati_impianto['Marca_r'] = dati.marca_riduttore.loc[i]

        dati_impianto['Tipo_r'] = dati.Tipo_riduttore.loc[i]
        dati_impianto['DNrid'] = dati.Dnrid_mm.loc[i].astype(str) + ' ' * 2 + "mm"
        dati_impianto['C1'] = dati.c1.loc[i]
        dati_impianto['Cg'] = dati.cg.loc[i]
        dati_impianto['DNtub rid valle'] = dati.Dntub_rid_val.loc[i].astype(str) + ' ' * 2 + "mm"
        dati_impianto['SCARICO'] = ''
        dati_impianto['DNvalle riduzione'] = dati.Dntub_rid_val.loc[i].astype(str) + ' ' * 2 + "mm"
        #dati_impianto['Marca'] = dati.marca_scarico.loc[i]
        #dati_impianto['Tipo'] = dati.Tipo_scarico.loc[i]
        dati_impianto['Dv_teo'] = dati.Dnvalv_scarico_mm.loc[i].astype(str) + ' ' * 2 + "mm"

        dati_impianto['A scarico'] = dati.A_scarico_mm2.loc[i].astype(str) + ' ' * 2 + "mm2"
        dati_impianto['K scarico'] = dati.k_scarico.loc[i]
        dati_impianto['MISURA'] = ''
        dati_impianto['TIPO MISURA'] = dati.tipo_misura.loc[i]
        # print(dati_impianto['TIPO MISURA'])

        # if dati_impianto['TIPO MISURA'] == 'volumetrico':


        dati_impianto['DNmisura'] = dati.Dnmis_2_mm.loc[i].astype(str) + ' ' * 2 + "mm"
        dati_impianto['Tipo contatore1'] = dati.tipo_contatore_1.loc[i]
        dati_impianto['Classe'] = dati.classe_1.loc[i]
        dati_impianto['Tipo contatore2'] = dati.tipo_contatore_2.loc[i]
        dati_impianto['Classe'] = dati.classe_2.loc[i]

        # dati_impianto['DP bassa'] = dati.DP_bassa_mbar.astype(str) + ' ' * 2 + "mbar"
        dati_impianto['DN BY-PASS'] = dati.BYPASS_DN.loc[i].astype(str) + ' ' * 2 + "mm"

        dati_impianto['IMPIANTO TERMICO'] = ''
        dati_impianto['Potenzialita caldaia1'] = dati.Potenzialità_caldaia_1_kw.loc[i].astype(str) + ' ' * 2 + "KW"
        dati_impianto['Potenzialita caldaia2'] = dati.Potenzialità_caldaia_2_kw.loc[i].astype(str) + ' ' * 2 + "KW"
        dati_impianto['Potenzialita caldaia3'] = dati.Potenzialità_caldaia_3_kw.loc[i].astype(str) + ' ' * 2 + "KW"


        dati_verifica['Q_max_monte'] = [dati.Q_max_teorica_monte.loc[i].astype(str) + ' ' * 2 + "m3/h",dati.Q_max_teorica_monte_ver.loc[i]]
        dati_verifica['Q_max_valle'] = [dati.Q_max_teorica_valle.loc[i].astype(str) + ' ' * 2 + "m3/h",dati.Q_max_teorica_valle_ver.loc[i]]
        dati_verifica['Velocita_monte'] = [dati.velocità_monte.loc[i].astype(str) + ' ' * 2 + "m/s",dati.velocità_monte_ver.loc[i]]
        dati_verifica['DP_monte'] = [dati.DP_monte.loc[i].astype(str) + ' ' * 2 + "mbar",dati.DP_monte_ver.loc[i]]
        dati_verifica['DP_linea'] = [dati.DP_valle.loc[i].astype(str) + ' ' * 2 + "mbar",dati.DP_valle_ver.loc[i]]
        dati_verifica['velocità_uscita'] = [dati.velocita_uscita.loc[i].astype(str) + ' ' * 2 + "m/s",dati.velocita_uscita_ver.loc[i]]

        dati_portata['Q max filtro'] = [dati.portata_max_filtro.loc[i].astype(str) + ' ' * 2 + "m3/h",dati.portata_max_filtro_ver.loc[i]]
        dati_portata['V linea filtro'] = [dati.velcità_linea_filtro.loc[i].astype(str) + ' ' * 2 + "m/s",dati.velcità_linea_filtro_ver.loc[i]]

        for c in range(1, (dati.N_linee.loc[i]) + 1):

            dati_potenzialita[f'C{c}'] = [dati[f'pot{c}'].loc[i].astype(str) + ' ' * 2 + "KW",dati[f'pot{c}_ver'].loc[i]]

        #dati_potenzialita['Velocita serp,scam'] = [dati.V_serp_scambiatore.loc[i].astype(str) + ' ' * 2 + "m/s",dati.V_serp_scambiatore_ver.loc[i]]
        dati_potenzialita['Velocita scam'] = [dati.V_serp_scambiatore.loc[i].astype(str) + ' ' * 2 + "m/s",
                                                   dati.V_serp_scambiatore_ver.loc[i]]


        dati_capacita['Q_max_riduttore'] = [dati.Q_max_rid.loc[i].astype(str) + ' ' * 2 + "m3/h",dati.Q_max_rid_ver.loc[i]]
        dati_capacita['v_riduttore'] = [dati.v_riduttore.loc[i].astype(str) + ' ' * 2 + "m/s",dati.v_riduttore_ver.loc[i]]
        dati_capacita['v_linea_riduttore'] = [dati.v_linea_riduttore.loc[i].astype(str) + ' ' * 2 + "m/s",dati.v_linea_riduttore_ver.loc[i]]
        dati_capacita['v_linea_riduttore_valle'] = [dati.v_tub_rid_valle.loc[i].astype(str) + ' ' * 2 + "m/s",dati.v_tub_rid_valle_ver.loc[i]]

        dati_scarico['diametro_teorico'] = [dati.diametro_scarico.loc[i].astype(str) + ' ' * 2 + "mm",dati.diametro_scarico_ver.loc[i]]

        dati_misura['Q_max_misura'] = [dati.Q_max_misura.loc[i].astype(str) + ' ' * 2 + "m3/h",dati.Q_max_misura_ver.loc[i]]
        dati_misura['V_bypass'] = [dati.v_bypass.loc[i].astype(str) + ' ' * 2 + "m/s",dati.v_bypass_ver.loc[i]]
        dati_misura['Q_max_bypass'] = [dati.q_max_bypass.loc[i].astype(str) + ' ' * 2 + "m3/h",dati.q_max_bypass_ver.loc[i]]
        #dati_misura_contatore['Q_max_contatore'] = [dati.Q_max_contatore.loc[i].astype(str) + ' ' * 2 + "m3/h",
                                      # dati.Q_max_contatore_ver.loc[i]]

        dati_pote['M'] = [dati.m.loc[i].astype(str) + ' ' * 2 + "KW",dati.m_ver.loc[i]]



        serie_dati_impianto = pd.Series(dati_impianto, name='DATI IMPIANTO')
        print(serie_dati_impianto)

        serie_dati_verifica = pd.Series(dati_verifica, name='VELOCITA')
        serie_dati_portata = pd.Series(dati_portata, name='PORTATA')
        serie_dati_potenzialita = pd.Series(dati_potenzialita, name='POTENZIALITA')
        serie_dati_capacita = pd.Series(dati_capacita, name='capacita')
        serie_dati_teorico = pd.Series(dati_scarico, name='teorico')
        serie_dati_misura = pd.Series(dati_misura, name='teorico1')
        serie_dati_pote = pd.Series(dati_pote, name='potenzialit')





        writer = pd.ExcelWriter(
            'C:\\dimensionamenti\\risultati\\{}_nuovi.xlsx'.format(dati_impianto['codice REMI'],
                                                                      engine='xlsxwriter'))

        serie_dati_impianto.to_excel(writer, sheet_name='Sheet1', startrow=start + 1, engine='xlsxwriter')
        # serie_dati_verifica.to_excel(writer, sheet_name='Sheet1', startrow=10,startcol=3)

        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        cell_formatH1 = workbook.add_format({'bold': False, 'align': 'right', 'bottom': True, 'top': True})
        cell_format = workbook.add_format({'bold': False, 'align': 'left', 'bottom': False, 'top': False})
        cell_format1 = workbook.add_format({'bold': False, 'align': 'right'})
        cell_format2 = workbook.add_format({'right': True, 'align': 'right'})
        cell_format3 = workbook.add_format({'bottom': True, 'align': 'right'})
        cell_format4 = workbook.add_format({'top': True, 'align': 'right'})
        cell_format5 = workbook.add_format({'right': True, 'align': 'right'})
        cell_format6 = workbook.add_format({'left': True, 'align': 'right'})

        cell_format_verifica1 = workbook.add_format({'bold': True, 'align': 'left'})
        cell_format_verifica2 = workbook.add_format({'bold': False, 'align': 'right'})
        cell_format_verifica3 = workbook.add_format({'bold': False, 'align': 'right', 'font_color': '#008000'})
        cell_format_verifica4 = workbook.add_format({'bold': False, 'align': 'right', 'font_color': '#DC143C'})
        header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'fg_color': '#1e90ff', 'border': 1})

        # formattazaione HEADER1
        '''
        worksheet.conditional_format(1 , 0, 6 , 0,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_formatH1})
        '''

        # formattazaione colonna 0
        worksheet.conditional_format(3 + start, 0, 59 + start, 0,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format})

        # formattazione colonna1
        worksheet.conditional_format(2 + start, 1, 59 + start, 1,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format2})
        # linee orizzontali superiori sulla destra
        worksheet.conditional_format(26 + start, 2, 26 + start, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format3})

        worksheet.conditional_format(32 + start, 2, 32 + start, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format3})
        worksheet.conditional_format(37 + start, 2, 37 + start, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format3})
        worksheet.conditional_format(44 + start, 2, 44 + start, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format3})
        worksheet.conditional_format(48 + start, 2, 48 + start, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format3})
        worksheet.conditional_format(55 + start, 2, 55 + start, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format3})
        # formattazione ultima riga tabella
        worksheet.conditional_format(60 + start, 0, 60 + start, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format4})

        # formattazione bordo esterno
        worksheet.conditional_format(1 + start, 5, 59 + start, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format5})
        worksheet.conditional_format(1 + start, 5, 59 + start, 6,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format6})

        #formattazione verificato
        worksheet.conditional_format(13 + start, 5, 59 + start, 5,
                                     {'type': 'text', 'criteria': 'begins with', 'value': 'VERI',
                                      'format': cell_format_verifica3})
        worksheet.conditional_format(13 + start, 5, 59 + start, 5,
                                     {'type': 'text', 'criteria': 'begins with', 'value': 'NON',
                                      'format': cell_format_verifica4})


        worksheet.write(1 + start, 1, 'DATI IMPIANTO', header_format)
        worksheet.merge_range(1 + start, 2, 1 + start, 5, 'VERIFICA', header_format)

        worksheet.write(2 + start, 0, 'codice REMI', header_format)

        worksheet.merge_range(27 + start, 0, 27 + start, 1, 'FILTRO', header_format)
        # worksheet.merge_range(28+start,1,28+start,2, 'FILTRO', header_format)

        worksheet.merge_range(33 + start, 0, 33 + start, 1, 'SCAMBIATORI', header_format)

        worksheet.merge_range(38 + start, 0, 38 + start, 1, 'RIDUTTORE', header_format)
        worksheet.merge_range(45 + start, 0, 45 + start, 1, 'SCARICO', header_format)

        worksheet.merge_range(49 + start, 0, 49 + start, 1, 'MISURA', header_format)

        worksheet.merge_range(56 + start, 0, 56 + start, 1, 'IMPIANTO TERMICO', header_format)

        worksheet.set_column('A:A', 37, cell_format)

        worksheet.set_column('B:B', 25, cell_format1)
        worksheet.set_column('D:D', 22)
        worksheet.set_column('E:E', 15)
        worksheet.set_column('F:F', 18)

        worksheet.write(12 + start, 3, 'PORTATA E VELOCITA', cell_format_verifica1)
        #worksheet.write(12 + start, 3, 'VELOCITA', cell_format_verifica1)
        worksheet.write(27 + start, 3, 'PORTATA', cell_format_verifica1)
        worksheet.write(33 + start, 3, 'POTENZIALITA', cell_format_verifica1)
        worksheet.write(38 + start, 3, 'CAPACITA', cell_format_verifica1)
        worksheet.write(45 + start, 3, 'DIAMETRO TEORICO', cell_format_verifica1)
        worksheet.write(49 + start, 3, 'MISURA', cell_format_verifica1)
        worksheet.write(53 + start, 3, 'taglia misuratori', cell_format_verifica1)
        worksheet.write(53 + start, 5, dati.Q_max_contatore_ver.loc[i], cell_format_verifica3)
        worksheet.write(56 + start, 3, 'POTENZIALITA', cell_format_verifica1)

        for i, value in enumerate(serie_dati_verifica):
            worksheet.write(13 + start + i, 4, serie_dati_verifica[i][0], cell_format_verifica2)
            worksheet.write(13 + start + i, 3, serie_dati_verifica.keys()[i], cell_format_verifica1)
            worksheet.write(13 + start + i, 5, serie_dati_verifica[i][1], cell_format_verifica3)

        for i, value1 in enumerate(serie_dati_portata):
            worksheet.write(28 + start + i, 4, serie_dati_portata[i][0], cell_format_verifica2)
            worksheet.write(28 + start + i, 3, serie_dati_portata.keys()[i], cell_format_verifica1)
            worksheet.write(28 + start + i, 5, serie_dati_portata[i][1])

        for i, value2 in enumerate(serie_dati_potenzialita):
            worksheet.write(34 + start + i, 4, serie_dati_potenzialita[i][0], cell_format_verifica2)
            worksheet.write(34 + start + i, 3, serie_dati_potenzialita.keys()[i], cell_format_verifica1)
            worksheet.write(34 + start + i, 5, serie_dati_potenzialita[i][1], cell_format_verifica3)

        for i, value3 in enumerate(serie_dati_capacita):
            worksheet.write(39 + start + i, 4, serie_dati_capacita[i][0], cell_format_verifica2)
            worksheet.write(39 + start + i, 3, serie_dati_capacita.keys()[i], cell_format_verifica1)
            worksheet.write(39 + start + i, 5,serie_dati_capacita[i][1], cell_format_verifica3)

        for i, value4 in enumerate(serie_dati_teorico):
            worksheet.write(46 + start + i, 4, serie_dati_teorico[i][0], cell_format_verifica2)
            worksheet.write(46 + start + i, 3, serie_dati_teorico.keys()[i], cell_format_verifica1)
            worksheet.write(46 + start + i, 5, serie_dati_teorico[i][1], cell_format_verifica3)

        for i, value5 in enumerate(serie_dati_misura):
            worksheet.write(50 + start + i, 4, serie_dati_misura[i][0], cell_format_verifica2)
            worksheet.write(50 + start + i, 3, serie_dati_misura.keys()[i], cell_format_verifica1)
            worksheet.write(50 + start + i, 5, serie_dati_misura[i][1], cell_format_verifica3)

        for i, value6 in enumerate(serie_dati_pote):
            worksheet.write(57 + start + i, 4, serie_dati_pote[i][0], cell_format_verifica2)
            worksheet.write(57 + start + i, 3, serie_dati_pote.keys()[i], cell_format_verifica1)
            worksheet.write(57 + start + i, 5, serie_dati_pote[i][1], cell_format_verifica3)

        # worksheet.merge_range('C2:C60',None, cell_format2)


        chart = workbook.add_chart({'type': 'column'})

        chart1 = workbook.add_chart({'type': 'column'})

        chart.set_x_axis({
            'major_gridlines': {
                'visible': True,
                'line': {'width': 1.25, 'dash_type': 'dash'}
            },
        })

        #scrittura primo grafico
        worksheet.write(62 + start, 1, 'name')
        worksheet.write(62 + start, 2, 'velocità impianto')
        worksheet.write(62 + start, 3, 'velocità max')

        worksheet.write(63 + start, 3, float(dati_verifica['Velocita_monte'][0].split(' ')[0]))
        worksheet.write(64 + start, 3, float(dati_portata['V linea filtro'][0].split(' ')[0]))
        worksheet.write(65 + start, 3, float(dati_capacita['v_linea_riduttore'][0].split(' ')[0]))
        worksheet.write(66 + start, 3, float(dati_verifica['velocità_uscita'][0].split(' ')[0]))
        worksheet.write(67 + start, 3, float(dati_misura['V_bypass'][0].split(' ')[0]))
        worksheet.write(68 + start, 3, float(dati_verifica['velocità_uscita'][0].split(' ')[0]))

        worksheet.write(63 + start, 2, 30)
        worksheet.write(64 + start, 2, 30)
        worksheet.write(65 + start, 2, 25)
        worksheet.write(66 + start, 2, 25)
        worksheet.write(67 + start, 2, 30)
        worksheet.write(68 + start, 2, 25)

        worksheet.write(63 + start, 1, 'Ingresso')
        worksheet.write(64 + start, 1, 'Linea')
        worksheet.write(65 + start, 1, 'Riduttori')
        worksheet.write(66 + start, 1, 'Misura')
        worksheet.write(67 + start, 1, 'Bypass')
        worksheet.write(68 + start, 1, 'Uscita')\


        # chart.add_series({'name':'=Sheet1!$C$68','categories':'=Sheet1!$B$69:$B$74','values': '=Sheet1!$D$69:$D$74','line': {'color': 'blue'}})
        chart.add_series({'name': ['Sheet1', 62 + start, 2], 'categories': ['Sheet1', 63 + start, 1, 68 + start, 1],
                          'values': ['Sheet1', 63 + start, 3, 68 + start, 3],
                          'line': {'color': 'blue'}})

        # chart.add_series({'name':['Sheet1',67,3],'categories':'=Sheet1!$B$69:$B$74','values': '=Sheet1!$C$69:$C74', 'line': {'color': 'red'}})

        chart.add_series({'name': ['Sheet1', 62 + start, 3], 'categories': ['Sheet1', 63 + start, 1, 68 + start, 1],
                          'values': ['Sheet1', 63 + start, 2, 68 + start, 2],
                          'line': {'color': 'red'}})

        chart.set_title({'name': 'Impianto {}'.format(df_dati_impianti['rif_remi'])})

        chart.set_y_axis({'name': 'Velocita (m/s)'})

        chart.set_plotarea({
            'layout': {
                'x': 0.13,
                'y': 0.2,
                'width': 0.73,
                'height': 0.6,
            }
        })

        chart.set_legend({'num_font': {'name': 'Arial', 'size': 3}, 'position': 'bottom'})
        worksheet.insert_chart(62 + start, 1, chart)

        # scrittura secondo grafico
        worksheet.write(82 + start, 1, 'name')
        worksheet.write(82 + start, 2, 'portata impianto')
        worksheet.write(82 + start, 3, 'porata max')

        worksheet.write(83 + start, 3, float( dati_verifica['Q_max_monte'][0].split(' ')[0]))
        worksheet.write(84 + start, 3, float(dati_portata['Q max filtro'][0].split(' ')[0]))
        worksheet.write(85 + start, 3, float( dati_capacita['Q_max_riduttore'][0].split(' ')[0]))
        worksheet.write(86 + start, 3, float(dati_capacita['Q_max_riduttore'][0].split(' ')[0]))
        worksheet.write(87 + start, 3, float(dati_misura['Q_max_misura'][0].split(' ')[0]))
        worksheet.write(88 + start, 3, float(dati_misura['Q_max_bypass'][0].split(' ')[0]))
        worksheet.write(89 + start, 3, float(dati_verifica['Q_max_valle'][0].split(' ')[0]))

        worksheet.write(83 + start, 2, float(dati_impianto['Q_imp'].split(' ')[0]))
        worksheet.write(84 + start, 2, float(dati_impianto['Q_imp'].split(' ')[0]))
        worksheet.write(85 + start, 2, float(dati_impianto['Q_imp'].split(' ')[0]))
        worksheet.write(86 + start, 2, float(dati_impianto['Q_imp'].split(' ')[0]))
        worksheet.write(87 + start, 2, float(dati_impianto['Q_imp'].split(' ')[0]))
        worksheet.write(88 + start, 2, float(dati_impianto['Q_imp'].split(' ')[0]))
        worksheet.write(89 + start, 2, float(dati_impianto['Q_imp'].split(' ')[0]))

        worksheet.write(83 + start, 1, 'Ingresso')
        worksheet.write(84 + start, 1, 'Filtri')
        worksheet.write(85 + start, 1, 'Scamb')
        worksheet.write(86 + start, 1, 'Riduttori')
        worksheet.write(87 + start, 1, 'Misura')
        worksheet.write(88 + start, 1, 'Bypass')
        worksheet.write(89 + start, 1, 'Uscita') \
 \
 \
            # chart.add_series({'name':'=Sheet1!$C$68','categories':'=Sheet1!$B$69:$B$74','values': '=Sheet1!$D$69:$D$74','line': {'color': 'blue'}})

        chart1.set_x_axis({
            'major_gridlines': {
                'visible': True,
                'line': {'width': 1.25, 'dash_type': 'dash'}
            },
        })

        chart1.add_series({'name': ['Sheet1', 82 + start, 2], 'categories': ['Sheet1', 83 + start, 1, 89 + start, 1],
                          'values': ['Sheet1', 83 + start, 3, 89 + start, 3],
                          'line': {'color': 'blue'}})

        # chart.add_series({'name':['Sheet1',67,3],'categories':'=Sheet1!$B$69:$B$74','values': '=Sheet1!$C$69:$C74', 'line': {'color': 'red'}})

        chart1.add_series({'name': ['Sheet1', 82 + start, 3], 'categories': ['Sheet1', 83 + start, 1, 89 + start, 1],
                          'values': ['Sheet1', 83 + start, 2, 89 + start, 2],
                          'line': {'color': 'red'}})

        chart1.set_title({'name': 'Impianto {}'.format(df_dati_impianti['rif_remi'])})

        chart1.set_y_axis({'name': 'Portata (m3/s)'})

        chart1.set_plotarea({
            'layout': {
                'x': 0.13,
                'y': 0.2,
                'width': 0.73,
                'height': 0.6,
            }
        })

        chart1.set_legend({'num_font': {'name': 'Arial', 'size': 1}, 'position': 'bottom'})
        worksheet.insert_chart(82 + start, 1, chart1)


        #intestazione

        worksheet.write(2, 1, 'Riferimento Re.Mi')
        worksheet.write(3, 1, 'Codice Re.Mi')
        worksheet.write(4, 1, 'Comune')
        worksheet.write(5, 1, 'Operazione')
        worksheet.write(6, 1, 'Data')
        worksheet.write(7, 1, 'Versione')

        cell_format_header1 = workbook.add_format({'top': True, 'left': True})
        cell_format_header2 = workbook.add_format({'top': True, })
        cell_format_header3 = workbook.add_format({'top': True, 'right': True})
        cell_format_header4 = workbook.add_format({'left': True})
        cell_format_header5 = workbook.add_format({'bottom': True, 'left': True})
        cell_format_header6 = workbook.add_format({'bottom': True})
        cell_format_header7 = workbook.add_format({'bottom': True, 'right': True})
        cell_format_header8 = workbook.add_format({'right': True})

        worksheet.conditional_format(2, 0, 2, 0,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header1})
        worksheet.conditional_format(2, 1, 2, 1,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header3})

        worksheet.conditional_format(2, 2, 2, 4,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header2})
        worksheet.conditional_format(2, 2, 2, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header3})
        worksheet.conditional_format(3, 0, 6, 0,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header4})
        worksheet.conditional_format(7, 0, 7, 0,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header5})
        worksheet.conditional_format(7, 1, 7, 1,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header7})
        worksheet.conditional_format(3, 1, 6, 1,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header8})
        worksheet.conditional_format(3, 5, 6, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header8})
        worksheet.conditional_format(7, 2, 7, 2,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header5})
        worksheet.conditional_format(7, 3, 7, 4,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header6})
        worksheet.conditional_format(7, 5, 7, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header7})

        worksheet.conditional_format(12, 0, 71, 0,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header8})

        worksheet.write(2, 3, dati_testata['riferimento_remi'])
        worksheet.write(3, 3, dati_testata['impianto'])
        worksheet.write(4, 3, dati_testata['comune'])
        worksheet.write(5, 3, 'Verifica stato futuro')
        worksheet.write(6, 3, '15/10/2019')
        worksheet.write(7, 3, '0')

        # worksheet.write('A4', 'Insert a scaled image:')

        cell_format_image = workbook.add_format({'bold': False, 'align': 'center', 'top': True})
        worksheet.insert_image('A4', 'C:\\dimensionamenti\\input\\Logo.jpg',
                               {'x_scale': 0.25, 'y_scale': 0.25, 'x_offset': 15, 'y_offset': 15})

        writer.save()

    else:

        k = 0.977
        k1 = 0.977
        k2 = 0.977
        start = 10

        dati['velocità_monte'] = dati.apply(funzioni.calcola_velocita_monte, args=[k], axis=1).apply(lambda x: x[0])
        dati['velocità_monte_ver'] = dati.apply(funzioni.calcola_velocita_monte, args=[k], axis=1).apply(lambda x: x[1])

        dati['DP_monte'] = dati.apply(funzioni.calcola_dp_monte, args=[k1],axis=1).apply(lambda x: x[0])
        dati['DP_monte_ver'] = dati.apply(funzioni.calcola_dp_monte,args=[k1], axis=1).apply(lambda x: x[1])

        dati['DP_valle'] = dati.apply(funzioni.calcola_dp_linea, args=[k1],axis=1).apply(lambda x: x[0])
        dati['DP_valle_ver'] = dati.apply(funzioni.calcola_dp_linea,args=[k1], axis=1).apply(lambda x: x[1])

        dati['velocita_uscita'] = dati.apply(funzioni.calcola_velocita_uscita, args=[k], axis=1).apply(lambda x: x[0])
        dati['velocita_uscita_ver'] = dati.apply(funzioni.calcola_velocita_uscita, args=[k], axis=1).apply(
            lambda x: x[1])

        dati['portata_max_filtro'] = dati.apply(funzioni.calcola_portata_max_filtro, axis=1).apply(lambda x: x[0])
        dati['portata_max_filtro_ver'] = dati.apply(funzioni.calcola_portata_max_filtro, axis=1).apply(lambda x: x[1])

        dati['velcità_linea_filtro'] = dati.apply(funzioni.calcola_velocita_linea_filtro, args=[k2], axis=1).apply(
            lambda x: x[0])
        dati['velcità_linea_filtro_ver'] = dati.apply(funzioni.calcola_velocita_linea_filtro, args=[k2], axis=1).apply(
            lambda x: x[1])

        dati['pot1'] = dati.apply(funzioni.calcolo_c, axis=1).apply(lambda x: x[0])
        dati['pot2'] = dati.apply(funzioni.calcolo_c, axis=1).apply(lambda x: x[0])
        dati['pot3'] = dati.apply(funzioni.calcolo_c, axis=1).apply(lambda x: x[0])
        dati['pot1_ver'] = dati.apply(funzioni.calcolo_c, axis=1).apply(lambda x: x[1])
        dati['pot2_ver'] = dati.apply(funzioni.calcolo_c, axis=1).apply(lambda x: x[1])
        dati['pot3_ver'] = dati.apply(funzioni.calcolo_c, axis=1).apply(lambda x: x[1])

        dati['V_serp_scambiatore'] = dati.apply(funzioni.calcola_velocita_linea_filtro, args=[k], axis=1).apply(
            lambda x: x[0])
        dati['V_serp_scambiatore_ver'] = dati.apply(funzioni.calcola_velocita_linea_filtro, args=[k], axis=1).apply(
            lambda x: x[1])

        dati['Q_max_rid'] = dati.apply(funzioni.calcolo_capacita_riduttore, axis=1).apply(lambda x: x[0])
        dati['Q_max_rid_ver'] = dati.apply(funzioni.calcolo_capacita_riduttore, axis=1).apply(lambda x: x[1])

        dati['v_riduttore'] = dati.apply(funzioni.calcolo_velocita_riduttore, args=[k], axis=1).apply(lambda x: x[0])
        dati['v_riduttore_ver'] = dati.apply(funzioni.calcolo_velocita_riduttore, args=[k], axis=1).apply(
            lambda x: x[1])

        dati['v_linea_riduttore'] = dati.apply(funzioni.calcola_velocita_linea_riduttore, args=[k], axis=1).apply(
            lambda x: x[0])
        dati['v_linea_riduttore_ver'] = dati.apply(funzioni.calcola_velocita_linea_riduttore, args=[k], axis=1).apply(
            lambda x: x[1])

        dati['v_tub_rid_valle'] = dati.apply(funzioni.calcolo_velocita_riduttore_valle, args=[k], axis=1).apply(
            lambda x: x[0])
        dati['v_tub_rid_valle_ver'] = dati.apply(funzioni.calcolo_velocita_riduttore_valle, args=[k], axis=1).apply(
            lambda x: x[1])

        dati['diametro_scarico'] = dati.apply(funzioni.calcolo_diamtero_teorico, axis=1).apply(lambda x: x[0])
        dati['diametro_scarico_ver'] = dati.apply(funzioni.calcolo_diamtero_teorico, axis=1).apply(lambda x: x[1])

        # dati['Q_max_misura'] = dati.apply(funzioni.calcolo_portata_massima_volumetrica, axis=1)
        dati['Q_max_misura'] = dati.apply(funzioni.calcolo_portata_massima_volumetrica, args=[k2], axis=1).apply(
            lambda x: x[0])
        dati['Q_max_misura_ver'] = dati.apply(funzioni.calcolo_portata_massima_volumetrica, args=[k2], axis=1).apply(
            lambda x: x[1])

        dati['Q_max_venturi'] = dati.apply(funzioni.calcolo_portata_massima_ventirumetria, axis=1).apply(
            lambda x: x[0])
        dati['Q_max_venturi_ver'] = dati.apply(funzioni.calcolo_portata_massima_ventirumetria, axis=1).apply(
            lambda x: x[1])

        dati['v_bypass'] = dati.apply(funzioni.calcolo_velocita_bypass, args=[k], axis=1).apply(lambda x: x[0])
        dati['v_bypass_ver'] = dati.apply(funzioni.calcolo_velocita_bypass, args=[k], axis=1).apply(lambda x: x[1])

        dati['q_va'] = dati.apply(funzioni.calcolo_q_valle, args=[k], axis=1)

        dati['q_max_bypass'] = dati.apply(funzioni.calcolo_portata_bypass, axis=1).apply(lambda x: x[0])
        dati['q_max_bypass_ver'] = dati.apply(funzioni.calcolo_portata_bypass, axis=1).apply(lambda x: x[1])

        dati['m'] = dati.apply(funzioni.calcolo_potenzialita, axis=1).apply(lambda x: x[0])
        dati['m_ver'] = dati.apply(funzioni.calcolo_potenzialita, axis=1).apply(lambda x: x[1])

        dati['riferimento_REMI'] = dati.apply(lambda row: 'REMI ' + row.descrizione, axis=1)

        dati['Q_max_teorica_monte'] = dati.apply(funzioni.calcola_q_max_monte, args=[k2], axis=1).apply(lambda x: x[0])
        dati['Q_max_teorica_monte_ver'] = dati.apply(funzioni.calcola_q_max_monte, args=[k2], axis=1).apply(
            lambda x: x[1])

        dati['Q_max_teorica_valle'] = dati.apply(funzioni.calcola_q_max_valle, args=[k2], axis=1).apply(lambda x: x[0])
        dati['Q_max_teorica_valle_ver'] = dati.apply(funzioni.calcola_q_max_valle, args=[k2], axis=1).apply(
            lambda x: x[1])

        dati['Q_max_teorica_scambiatore'] = dati.apply(funzioni.calcola_q_max_scambiatore_t, args=[k], axis=1).apply(
            lambda x: x[0])
        dati['Q_max_teorica_scambiatore_ver'] = dati.apply(funzioni.calcola_q_max_scambiatore_t, args=[k],
                                                           axis=1).apply(
            lambda x: x[1])

        df_dati_impianti = {}
        dati_testata = {}
        dati_impianto = {}
        dati_impiantov = {}
        dati_verifica = {}
        dati_portate_teoriche_max = {}

        dati_portata = {}
        dati_potenzialita = {}
        dati_capacita = {}
        dati_scarico = {}
        dati_misura = {}
        dati_misura1 = {}
        dati_pote = {}

        dati_impiantov['codice REMI'] = i

        dati_testata['riferimento_remi'] = 'REMI ' + str(i) + ' ' + dati.descrizione.loc[i]
        dati_testata['impianto'] = dati.codice_impianto.loc[i]
        dati_testata['comune'] = dati.comune.loc[i]
        df_dati_impianti['rif_remi'] = 'REMI ' + ' ' + dati.descrizione.loc[i]

        dati_impiantov['Q_imp'] = dati.Qimp_m3_h.loc[i].astype(str) + ' ' * 2 + "m3/h"
        dati_impiantov['Q_ero'] = dati.Qero_m3_h.loc[i].astype(str) + ' ' * 2 + "m3/h"
        dati_impiantov['Q_linea_filtro'] = dati.Qlin_m3_h.loc[i].astype(str) + ' ' * 2 + "m3/h"
        dati_impiantov['Q_linea_scambiatore'] = dati.Qlin_m3_h.loc[i].astype(str) + ' ' * 2 + "m3/h"
        dati_impiantov['Q_linea_riduttore'] = dati.Qlin_m3_h.loc[i].astype(str) + ' ' * 2 + "m3/h"
        dati_impiantov['Coeff sicurezza'] = "10%"
        dati_impiantov['Pressione minima'] = dati.Pmin_bar.loc[i].astype(str) + ' ' * 2 + "bar"
        dati_impiantov['Pressione minima assoluta'] = dati.Pmin_bar.loc[i] + 1
        dati_impiantov['Pressione minima assoluta'] = dati_impiantov['Pressione minima assoluta'].astype(
            str) + ' ' * 2 + "bar"
        dati_impiantov['Pressione massima'] = dati.Pmax_bar.loc[i].astype(str) + ' ' * 2 + "bar"
        dati_impiantov['Pressione massima assoluta'] = dati.Pmax_bar.loc[i] + 1
        dati_impiantov['Pressione massima assoluta'] = dati_impiantov['Pressione massima assoluta'].astype(
            str) + ' ' * 2 + "bar"
        dati_impiantov['Pressione preriscaldo '] = dati.Pmax_bar.loc[i].astype(str) + ' ' * 2 + "bar"
        dati_impiantov['Pressione preriscaldo assoluta'] = dati.Pmax_bar.loc[i] + 1
        dati_impiantov['Pressione preriscaldo assoluta'] = dati_impiantov['Pressione preriscaldo assoluta'].astype(
            str) + ' ' * 2 + "bar"
        dati_impiantov['Pressione misura'] = dati.Pmisura_bar.loc[i].astype(str) + ' ' * 2 + "bar"
        dati_impiantov['Pressione misura assoluta'] = dati.Pmisura_bar.loc[i] + 1
        dati_impiantov['Pressione misura assoluta'] = dati_impiantov['Pressione misura assoluta'].astype(
            str) + ' ' * 2 + "bar"
        dati_impiantov['Numero linee filtro'] = dati.N_linee.loc[i]
        dati_impiantov['Numero linee scambiatore'] = dati.N_linee.loc[i]
        dati_impiantov['Numero linee preriscaldo'] = dati.N_linee.loc[i]
        dati_impiantov['Altezza livello mare'] = dati.altezzas_lm.loc[i].astype(str) + ' ' * 2 + "m"
        dati_impiantov['Lunghezza monte'] = dati.Lmonte_m.loc[i].astype(str) + ' ' * 2 + "m"
        dati_impiantov['DNmonte'] = dati.Dnmonte_mm.loc[i].astype(str) + ' ' * 2 + "mm"
        dati_impiantov['DNvalle riduzione'] = dati.DNlinea_rid_mm.loc[i].astype(int).astype(str) + ' ' * 2 + "mm"
        dati_impiantov['DNvalle'] = dati.Dnvalle_mm.loc[i].astype(str) + ' ' * 2 + "mm"
        dati_impiantov['PN/ANSI sino 1a valvola valle regolatori'] = dati.PN_D_U_ANSI_monte.loc[i].astype(
            str) + ' ' * 2 + "ANSI"
        dati_impiantov['PN/ANSI dopo 1a valvola valle regolatori'] = dati.PN_D_U_ANSI_monte.loc[i].astype(
            str) + ' ' * 2 + "ANSI"
        dati_impiantov['FILTRO'] = ''
        dati_impiantov['Tipo filtro'] = dati.tipo_filtro.loc[i]
        dati_impiantov['N° cartucce'] = dati.N_cartucce.loc[i]
        dati_impiantov['DNfiltro'] = dati.Dnfiltro_mm.loc[i].astype(str) + ' ' * 2 + "mm"
        dati_impiantov['Pressione bollo filtro '] = dati.Pbollo_filtro_bar.loc[i].astype(str) + ' ' * 2 + "bar"
        dati_impiantov['A'] = dati.A_filtro_m2.loc[i].astype(str) + ' ' * 2 + "m2"
        dati_impiantov['SCAMBIATORI'] = ''
        dati_impiantov['DNscambiatore'] = dati.DNscambiatore.loc[i].astype(str) + ' ' * 2 + "mm"
        dati_impiantov['Potenzialita1'] = dati.Potenzialità_1_kw.loc[i].astype(str) + ' ' * 2 + "KW"
        dati_impiantov['Potenzialita2'] = dati.Potenzialità_2_kw.loc[i].astype(str) + ' ' * 2 + "KW"
        dati_impiantov['Potenzialita3'] = dati.Potenzialità_3_kw.loc[i].astype(str) + ' ' * 2 + "KW"
        dati_impiantov['Pressione bollo filtro '] = dati.Pbollo_scamb.loc[i].astype(str) + ' ' * 2 + "bar"
        dati_impiantov['RIDUTTORE'] = ''
        dati_impiantov['Marca_r'] = dati.marca_riduttore.loc[i]
        dati_impiantov['Tipo_r'] = dati.Tipo_riduttore.loc[i]
        dati_impiantov['DNrid'] = dati.Dnrid_mm.loc[i].astype(str) + ' ' * 2 + "mm"
        dati_impiantov['C1'] = dati.c1.loc[i]
        dati_impiantov['Cg'] = dati.cg.loc[i]
        dati_impiantov['DNtub rid valle'] = dati.Dntub_rid_val.loc[i].astype(str) + ' ' * 2 + "mm"
        dati_impiantov['SCARICO'] = ''
        dati_impiantov['DNvalle riduzione'] = dati.Dntub_rid_val.loc[i].astype(str) + ' ' * 2 + "mm"
        #dati_impiantov['Marca'] = dati.marca_scarico.loc[i]
        #dati_impiantov['Tipo'] = dati.Tipo_scarico.loc[i]
        dati_impiantov['Dv_teo'] = dati.Dnvalv_scarico_mm.loc[i].astype(str) + ' ' * 2 + "mm"
        dati_impiantov['A scarico'] = dati.A_scarico_mm2.loc[i].astype(str) + ' ' * 2 + "mm2"
        dati_impiantov['K scarico'] = dati.k_scarico.loc[i]
        dati_impiantov['MISURA1'] = ''
        dati_impiantov['TIPO MISURA'] = dati.tipo_misura.loc[i]
        # print(dati_impianto['TIPO MISURA'])


        dati_impiantov['DN misura venturimetrica'] = dati.Dnmis_vent_mm.loc[i].astype(str) + ' ' * 2 + "mm"
        dati_impiantov['DN orifizio 1'] = dati.Ø_orifizio_1.loc[i].astype(str) + ' ' * 2 + "mm"
        dati_impiantov['DN orifizio 2'] = dati.Ø_orifizio_2.loc[i].astype(str) + ' ' * 2 + "mm"
        dati_impiantov['DN orifizio 3'] = dati.Ø_orifizio_3.loc[i].astype(str) + ' ' * 2 + "mm"
        dati_impiantov['DP alta'] = dati.DP_alta_mbar.loc[i].astype(str) + ' ' * 2 + "mbar"
        # dati_impianto['DP bassa'] = dati.DP_bassa_mbar.astype(str) + ' ' * 2 + "mbar"
        dati_impiantov['DN BY-PASS'] = dati.BYPASS_DN.loc[i].astype(str) + ' ' * 2 + "m2"

        # dati_impiantov['IMPIANTO TERMICO1'] = ''
        dati_impiantov['Potenzialita caldaia1'] = dati.Potenzialità_caldaia_1_kw.loc[i].astype(str) + ' ' * 2 + "KW"
        dati_impiantov['Potenzialita caldaia2'] = dati.Potenzialità_caldaia_2_kw.loc[i].astype(str) + ' ' * 2 + "KW"
        dati_impiantov['Potenzialita caldaia3'] = dati.Potenzialità_caldaia_3_kw.loc[i].astype(str) + ' ' * 2 + "KW"

        dati_verifica['Q_max_monte'] = [dati.Q_max_teorica_monte.loc[i].astype(str) + ' ' * 2 + "m3/h",
                                        dati.Q_max_teorica_monte_ver.loc[i]]
        dati_verifica['Q_max_valle'] = [dati.Q_max_teorica_valle.loc[i].astype(str) + ' ' * 2 + "m3/h",
                                        dati.Q_max_teorica_valle_ver.loc[i]]
        dati_verifica['Velocita_monte'] = [dati.velocità_monte.loc[i].astype(str) + ' ' * 2 + "m/s",
                                           dati.velocità_monte_ver.loc[i]]
        dati_verifica['DP_monte'] = [dati.DP_monte.loc[i].astype(str) + ' ' * 2 + "mbar", dati.DP_monte_ver.loc[i]]
        dati_verifica['DP_valle'] = [dati.DP_valle.loc[i].astype(str) + ' ' * 2 + "mbar", dati.DP_valle_ver.loc[i]]
        dati_verifica['velocità_uscita'] = [dati.velocita_uscita.loc[i].astype(str) + ' ' * 2 + "m/s",
                                            dati.velocita_uscita_ver.loc[i]]

        dati_portata['Q max filtro'] = [dati.portata_max_filtro.loc[i].astype(str) + ' ' * 2 + "m3/h",
                                        dati.portata_max_filtro_ver.loc[i]]
        dati_portata['V linea filtro'] = [dati.velcità_linea_filtro.loc[i].astype(str) + ' ' * 2 + "m/s",
                                          dati.velcità_linea_filtro_ver.loc[i]]

        for c in range(1, (dati.N_linee.loc[i]) + 1):

            dati_potenzialita[f'C{c}'] = [dati[f'pot{c}'].loc[i].astype(str) + ' ' * 2 + "KW",dati[f'pot{c}_ver'].loc[i]]

        dati_potenzialita['Velocita serp,scam'] = [dati.V_serp_scambiatore.loc[i].astype(str) + ' ' * 2 + "m/s",
                                                   dati.V_serp_scambiatore_ver.loc[i]]

        dati_capacita['Q_max_riduttore'] = [dati.Q_max_rid.loc[i].astype(str) + ' ' * 2 + "m3/h",
                                            dati.Q_max_rid_ver.loc[i]]
        dati_capacita['v_riduttore'] = [dati.v_riduttore.loc[i].astype(str) + ' ' * 2 + "m/s",
                                        dati.v_riduttore_ver.loc[i]]
        dati_capacita['v_linea_riduttore'] = [dati.v_linea_riduttore.loc[i].astype(str) + ' ' * 2 + "m/s",
                                              dati.v_linea_riduttore_ver.loc[i]]
        dati_capacita['v_linea_riduttore_valle'] = [dati.v_tub_rid_valle.loc[i].astype(str) + ' ' * 2 + "m/s",
                                                    dati.v_tub_rid_valle_ver.loc[i]]

        dati_scarico['diametro_teorico'] = [dati.diametro_scarico.loc[i].astype(str) + ' ' * 2 + "mm",
                                            dati.diametro_scarico_ver.loc[i]]

        dati_misura1['Q_max_venturi'] = [dati.Q_max_venturi.loc[i].astype(str) + ' ' * 2 + "m3/h",dati.Q_max_venturi_ver.loc[i]]
        dati_misura1['V_bypass'] = [dati.v_bypass.loc[i].astype(str) + ' ' * 2 + "m/s",dati.v_bypass_ver.loc[i]]
        dati_misura1['Q_max_bypass'] = [dati.q_max_bypass.loc[i].astype(str) + ' ' * 2 + "m3/h",dati.q_max_bypass_ver.loc[i]]

        dati_pote['M'] = [dati.m.loc[i].astype(str) + ' ' * 2 + "KW", dati.m_ver.loc[i]]

        # workbook = xlsxwriter.Workbook('C:\\vinicio\\dimensionamenti\\risultati\\{}.xlsx'.format(dati_impianto['codice REMI']))
        # worksheet = workbook.add_worksheet('dati_impianto')

        # df_dati_impianto=pd.DataFrame(dati_impianto)
        # df_dati_impianto=df_dati_impianto.set_index('codice REMI').stack()
        # workbook.close()


        serie_dati_impiantov = pd.Series(dati_impiantov, name='DATI IMPIANTO')

        serie_dati_verifica = pd.Series(dati_verifica, name='VELOCITA')
        serie_dati_portata = pd.Series(dati_portata, name='PORTATA')
        serie_dati_potenzialita = pd.Series(dati_potenzialita, name='POTENZIALITA')
        serie_dati_capacita = pd.Series(dati_capacita, name='capacita')
        serie_dati_teorico = pd.Series(dati_scarico, name='teorico')
        serie_dati_misura1 = pd.Series(dati_misura1, name='teorico1')

        serie_dati_pote = pd.Series(dati_pote, name='potenzialit')

        writer = pd.ExcelWriter(
            'C:\\dimensionamenti\\risultati\\{}_nuovi.xlsx'.format(dati_impiantov['codice REMI'],
                                                                      engine='xlsxwriter'))

        serie_dati_impiantov.to_excel(writer, sheet_name='Sheet1', startrow=start + 1, engine='xlsxwriter')
        # serie_dati_verifica.to_excel(writer, sheet_name='Sheet1', startrow=10,startcol=3)

        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        cell_format = workbook.add_format({'bold': False, 'align': 'left', 'bottom': False, 'top': False})
        cell_format1 = workbook.add_format({'bold': False, 'align': 'right'})
        cell_format2 = workbook.add_format({'right': True, 'align': 'right'})
        cell_format3 = workbook.add_format({'bottom': True, 'align': 'right'})
        cell_format4 = workbook.add_format({'top': True, 'align': 'right'})
        cell_format5 = workbook.add_format({'right': True, 'align': 'right'})
        cell_format6 = workbook.add_format({'left': True, 'align': 'right'})

        cell_format_verifica1 = workbook.add_format({'bold': True, 'align': 'left'})
        cell_format_verifica2 = workbook.add_format({'bold': False, 'align': 'right'})
        cell_format_verifica3 = workbook.add_format({'bold': False, 'align': 'right', 'font_color': '#008000'})
        cell_format_verifica4 = workbook.add_format({'bold': False, 'align': 'right', 'font_color': '#DC143C'})

        header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'fg_color': '#1e90ff', 'border': 1})

        # formattazaione colonna 0
        worksheet.conditional_format(3 + start, 0, 59 + start, 0,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format})

        # formattazione colonna1
        worksheet.conditional_format(2 + start, 1, 59 + start, 1,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format2})
        # linee orizzontali superiori sulla destra
        worksheet.conditional_format(26 + start, 2, 26 + start, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format3})

        worksheet.conditional_format(32 + start, 2, 32 + start, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format3})
        worksheet.conditional_format(37 + start, 2, 37 + start, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format3})
        worksheet.conditional_format(44 + start, 2, 44 + start, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format3})
        worksheet.conditional_format(48 + start, 2, 48 + start, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format3})
        worksheet.conditional_format(55 + start, 2, 55 + start, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format3})
        # formattazione ultima riga tabella
        worksheet.conditional_format(60 + start, 0, 60 + start, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format4})

        # formattazione bordo esterno
        worksheet.conditional_format(1 + start, 5, 59 + start, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format5})
        worksheet.conditional_format(1 + start, 5, 59 + start, 6,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format6})

        worksheet.conditional_format(13 + start, 5, 59 + start, 5,
                                     {'type': 'text', 'criteria': 'begins with', 'value': 'VERI',
                                      'format': cell_format_verifica3})
        worksheet.conditional_format(13 + start, 5, 59 + start, 5,
                                     {'type': 'text', 'criteria': 'begins with', 'value': 'NON',
                                      'format': cell_format_verifica4})



        worksheet.write(1 + start, 1, 'DATI IMPIANTO', header_format)
        worksheet.merge_range(1 + start, 2, 1 + start, 5, 'VERIFICA', header_format)

        worksheet.write(2 + start, 0, 'codice REMI', header_format)

        worksheet.merge_range(27 + start, 0, 27 + start, 1, 'FILTRO', header_format)
        # worksheet.merge_range(28+start,1,28+start,2, 'FILTRO', header_format)

        worksheet.merge_range(33 + start, 0, 33 + start, 1, 'SCAMBIATORI', header_format)

        worksheet.merge_range(38 + start, 0, 38 + start, 1, 'RIDUTTORE', header_format)
        worksheet.merge_range(45 + start, 0, 45 + start, 1, 'SCARICO', header_format)

        worksheet.merge_range(49 + start, 0, 49 + start, 1, 'MISURA', header_format)

        worksheet.merge_range(56 + start, 0, 56 + start, 1, 'IMPIANTO TERMICO', header_format)

        worksheet.set_column('A:A', 37, cell_format)

        worksheet.set_column('B:B', 25, cell_format1)
        worksheet.set_column('D:D', 22)
        worksheet.set_column('E:E', 15)
        worksheet.set_column('F:F', 20)

        worksheet.write(12 + start, 3, 'VELOCITA', cell_format_verifica1)
        worksheet.write(27 + start, 3, 'PORTATA', cell_format_verifica1)
        worksheet.write(33 + start, 3, 'POTENZIALITA', cell_format_verifica1)
        worksheet.write(38 + start, 3, 'CAPACITA', cell_format_verifica1)
        worksheet.write(45 + start, 3, 'DIAMETRO TEORICO', cell_format_verifica1)
        worksheet.write(49 + start, 3, 'MISURA', cell_format_verifica1)
        worksheet.write(53 + start, 3, 'tronco misura', cell_format_verifica1)
        worksheet.write(53 + start, 5, 'VERIFICATO', cell_format_verifica3)
        worksheet.write(56 + start, 3, 'POTENZIALITA', cell_format_verifica1)

        for i, value in enumerate(serie_dati_verifica):
            worksheet.write(13 + start + i, 4, serie_dati_verifica[i][0], cell_format_verifica2)
            worksheet.write(13 + start + i, 3, serie_dati_verifica.keys()[i], cell_format_verifica1)
            worksheet.write(13 + start + i, 5, serie_dati_verifica[i][1], cell_format_verifica3)

        for i, value1 in enumerate(serie_dati_portata):
            worksheet.write(28 + start + i, 4, serie_dati_portata[i][0], cell_format_verifica2)
            worksheet.write(28 + start + i, 3, serie_dati_portata.keys()[i], cell_format_verifica1)
            worksheet.write(28 + start + i, 5, serie_dati_portata[i][1], cell_format_verifica3)

        for i, value2 in enumerate(serie_dati_potenzialita):
            worksheet.write(34 + start + i, 4, serie_dati_potenzialita[i][0], cell_format_verifica2)
            worksheet.write(34 + start + i, 3, serie_dati_potenzialita.keys()[i], cell_format_verifica1)
            worksheet.write(34 + start + i, 5, serie_dati_potenzialita[i][1], cell_format_verifica3)

        for i, value3 in enumerate(serie_dati_capacita):
            worksheet.write(39 + start + i, 4, serie_dati_capacita[i][0], cell_format_verifica2)
            worksheet.write(39 + start + i, 3, serie_dati_capacita.keys()[i], cell_format_verifica1)
            worksheet.write(39 + start + i, 5, serie_dati_capacita[i][1], cell_format_verifica3)

        for i, value4 in enumerate(serie_dati_teorico):
            worksheet.write(46 + start + i, 4, serie_dati_teorico[i][0], cell_format_verifica2)
            worksheet.write(46 + start + i, 3, serie_dati_teorico.keys()[i], cell_format_verifica1)
            worksheet.write(46 + start + i, 5, serie_dati_teorico[i][1], cell_format_verifica3)

        for i, value5 in enumerate(serie_dati_misura1):
            worksheet.write(50 + start + i, 4, serie_dati_misura1[i][0], cell_format_verifica2)
            worksheet.write(50 + start + i, 3, serie_dati_misura1.keys()[i], cell_format_verifica1)
            worksheet.write(50 + start + i, 5, serie_dati_misura1[i][1], cell_format_verifica3)

        for i, value6 in enumerate(serie_dati_pote):
            worksheet.write(57 + start + i, 4, serie_dati_pote[i][0], cell_format_verifica2)
            worksheet.write(57 + start + i, 3, serie_dati_pote.keys()[i], cell_format_verifica1)
            worksheet.write(57 + start + i, 5, serie_dati_pote[i][1], cell_format_verifica3)

        # worksheet.merge_range('C2:C60',None, cell_format2)


        chart = workbook.add_chart({'type': 'column'})

        chart1 = workbook.add_chart({'type': 'column'})

        chart.set_x_axis({
            'major_gridlines': {
                'visible': True,
                'line': {'width': 1.25, 'dash_type': 'dash'}
            },
        })

        worksheet.write(62 + start, 1, 'name')
        worksheet.write(62 + start, 2, 'velocità impianto')
        worksheet.write(62 + start, 3, 'velocità max')

        worksheet.write(63 + start, 3, float(dati_verifica['Velocita_monte'][0].split(' ')[0]))
        worksheet.write(64 + start, 3, float(dati_portata['V linea filtro'][0].split(' ')[0]))
        worksheet.write(65 + start, 3, float(dati_capacita['v_linea_riduttore'][0].split(' ')[0]))
        worksheet.write(66 + start, 3, float(dati_verifica['velocità_uscita'][0].split(' ')[0]))
        worksheet.write(67 + start, 3, float(dati_misura1['V_bypass'][0].split(' ')[0]))
        worksheet.write(68 + start, 3, float(dati_verifica['velocità_uscita'][0].split(' ')[0]))

        worksheet.write(63 + start, 2, 30)
        worksheet.write(64 + start, 2, 30)
        worksheet.write(65 + start, 2, 25)
        worksheet.write(66 + start, 2, 25)
        worksheet.write(67 + start, 2, 30)
        worksheet.write(68 + start, 2, 25)

        worksheet.write(63 + start, 1, 'Ingresso')
        worksheet.write(64 + start, 1, 'Linea')
        worksheet.write(65 + start, 1, 'Riduttori')
        worksheet.write(66 + start, 1, 'Misura')
        worksheet.write(67 + start, 1, 'Bypass')
        worksheet.write(68 + start, 1, 'Uscita')

        worksheet.write(2, 1, 'Riferimento Re.Mi')
        worksheet.write(3, 1, 'Codice Re.Mi')
        worksheet.write(4, 1, 'Comune')
        worksheet.write(5, 1, 'Operazione')
        worksheet.write(6, 1, 'Data')
        worksheet.write(7, 1, 'Versione')



        cell_format_header1 = workbook.add_format({'top': True, 'left': True})
        cell_format_header2 = workbook.add_format({'top': True, })
        cell_format_header3 = workbook.add_format({'top': True, 'right': True})
        cell_format_header4 = workbook.add_format({'left': True})
        cell_format_header5 = workbook.add_format({'bottom': True, 'left': True})
        cell_format_header6 = workbook.add_format({'bottom': True})
        cell_format_header7 = workbook.add_format({'bottom': True, 'right': True})
        cell_format_header8 = workbook.add_format({'right': True})

        worksheet.conditional_format(2, 0, 2, 0,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header1})
        worksheet.conditional_format(2, 1, 2, 1,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header3})

        worksheet.conditional_format(2, 2, 2, 4,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header2})
        worksheet.conditional_format(2, 2, 2, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header3})
        worksheet.conditional_format(3, 0, 6, 0,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header4})
        worksheet.conditional_format(7, 0, 7, 0,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header5})
        worksheet.conditional_format(7, 1, 7, 1,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header7})
        worksheet.conditional_format(3, 1, 6, 1,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header8})
        worksheet.conditional_format(3, 5, 6, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header8})
        worksheet.conditional_format(7, 2, 7, 2,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header5})
        worksheet.conditional_format(7, 3, 7, 4,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header6})
        worksheet.conditional_format(7, 5, 7, 5,
                                     {'type': 'text', 'criteria': 'not containing', 'value': 'impossibile',
                                      'format': cell_format_header7})



        worksheet.write(2, 3, dati_testata['riferimento_remi'])
        worksheet.write(3, 3, dati_testata['impianto'])
        worksheet.write(4, 3, dati_testata['comune'])
        worksheet.write(5, 3, 'Verifica stato futuro')
        worksheet.write(6, 3, '15/10/2019')
        worksheet.write(7, 3, '0')


        # worksheet.write('A4', 'Insert a scaled image:')

        cell_format_image = workbook.add_format({'bold': False, 'align': 'center', 'top': True})
        worksheet.insert_image('A4', 'C:\\vinicio\\dimensionamenti\\input\\Logo.jpg',
                               {'x_scale': 0.25, 'y_scale': 0.25, 'x_offset': 15, 'y_offset': 15})
        # chart.add_series({'name':'=Sheet1!$C$68','categories':'=Sheet1!$B$69:$B$74','values': '=Sheet1!$D$69:$D$74','line': {'color': 'blue'}})
        chart.add_series({'name': ['Sheet1', 62 + start, 2], 'categories': ['Sheet1', 63 + start, 1, 68 + start, 1],
                          'values': ['Sheet1', 63 + start, 3, 68 + start, 3],
                          'line': {'color': 'blue'}})

        # chart.add_series({'name':['Sheet1',67,3],'categories':'=Sheet1!$B$69:$B$74','values': '=Sheet1!$C$69:$C74', 'line': {'color': 'red'}})

        chart.add_series({'name': ['Sheet1', 62 + start, 3], 'categories': ['Sheet1', 63 + start, 1, 68 + start, 1],
                          'values': ['Sheet1', 63 + start, 2, 68 + start, 2],
                          'line': {'color': 'red'}})

        chart.set_title({'name': 'Impianto {}'.format(df_dati_impianti['rif_remi'])})

        chart.set_y_axis({'name': 'Velocita (m/s)'})

        chart.set_plotarea({
            'layout': {
                'x': 0.13,
                'y': 0.2,
                'width': 0.73,
                'height': 0.6,
            }
        })
        """
        chart.set_legend({
            'layout': {
                'x': 1,
                'y': 0,
                'width': 0.18,
                'height': 0.25,

            }
        })
        """
        # chart.set_legend({'position': 'bottom'})
        chart.set_legend({'num_font': {'name': 'Arial', 'size': 3}, 'position': 'bottom'})
        worksheet.insert_chart(62 + start, 1, chart)

        # scrittura secondo grafico
        worksheet.write(82 + start, 1, 'name')
        worksheet.write(82 + start, 2, 'portata impianto')
        worksheet.write(82 + start, 3, 'porata max')

        worksheet.write(83 + start, 3, float(dati_verifica['Q_max_monte'][0].split(' ')[0]))
        worksheet.write(84 + start, 3, float(dati_portata['Q max filtro'][0].split(' ')[0]))
        worksheet.write(85 + start, 3, float(dati_capacita['Q_max_riduttore'][0].split(' ')[0]))
        worksheet.write(86 + start, 3, float(dati_capacita['Q_max_riduttore'][0].split(' ')[0]))
        worksheet.write(87 + start, 3, float(dati_misura1['Q_max_venturi'][0].split(' ')[0]))
        worksheet.write(88 + start, 3, float(dati_misura1['Q_max_bypass'][0].split(' ')[0]))
        worksheet.write(89 + start, 3, float(dati_verifica['Q_max_valle'][0].split(' ')[0]))

        worksheet.write(83 + start, 2, float(dati_impiantov['Q_imp'].split(' ')[0]))
        worksheet.write(84 + start, 2, float(dati_impiantov['Q_imp'].split(' ')[0]))
        worksheet.write(85 + start, 2, float(dati_impiantov['Q_imp'].split(' ')[0]))
        worksheet.write(86 + start, 2, float(dati_impiantov['Q_imp'].split(' ')[0]))
        worksheet.write(87 + start, 2, float(dati_impiantov['Q_imp'].split(' ')[0]))
        worksheet.write(88 + start, 2, float(dati_impiantov['Q_imp'].split(' ')[0]))
        worksheet.write(89 + start, 2, float(dati_impiantov['Q_imp'].split(' ')[0]))

        worksheet.write(83 + start, 1, 'Ingresso')
        worksheet.write(84 + start, 1, 'Filtri')
        worksheet.write(85 + start, 1, 'Scamb')
        worksheet.write(86 + start, 1, 'Riduttori')
        worksheet.write(87 + start, 1, 'Misura')
        worksheet.write(88 + start, 1, 'Bypass')
        worksheet.write(89 + start, 1, 'Uscita') \
 \
 \
            # chart.add_series({'name':'=Sheet1!$C$68','categories':'=Sheet1!$B$69:$B$74','values': '=Sheet1!$D$69:$D$74','line': {'color': 'blue'}})

        chart1.set_x_axis({
            'major_gridlines': {
                'visible': True,
                'line': {'width': 1.25, 'dash_type': 'dash'}
            },
        })

        chart1.add_series({'name': ['Sheet1', 82 + start, 2], 'categories': ['Sheet1', 83 + start, 1, 89 + start, 1],
                           'values': ['Sheet1', 83 + start, 3, 89 + start, 3],
                           'line': {'color': 'blue'}})

        # chart.add_series({'name':['Sheet1',67,3],'categories':'=Sheet1!$B$69:$B$74','values': '=Sheet1!$C$69:$C74', 'line': {'color': 'red'}})

        chart1.add_series({'name': ['Sheet1', 82 + start, 3], 'categories': ['Sheet1', 83 + start, 1, 89 + start, 1],
                           'values': ['Sheet1', 83 + start, 2, 89 + start, 2],
                           'line': {'color': 'red'}})

        chart1.set_title({'name': 'Impianto {}'.format(df_dati_impianti['rif_remi'])})

        chart1.set_y_axis({'name': 'Q(m3/s)'})

        chart1.set_plotarea({
            'layout': {
                'x': 0.13,
                'y': 0.2,
                'width': 0.73,
                'height': 0.6,
            }
        })

        chart1.set_legend({'num_font': {'name': 'Arial', 'size': 1}, 'position': 'bottom'})
        worksheet.insert_chart(82 + start, 1, chart1)

        writer.save()

