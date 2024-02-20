import os
import zipfile
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
import shutil
import pandas as pd
import numpy as np
import requests
from bs4 import BeautifulSoup
from soup2dict import convert
from datetime import datetime
import re
import time
import joblib
import calendar



def get_chromedriver(chromedriver_path=None):
    
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
    
    return driver

# def get_chromedriver(chromedriver_path=None, use_proxy=False, user_agent=None,
#                     PROXY_HOST=None, PROXY_PORT=None, PROXY_USER=None, PROXY_PASS=None, download_folder=None,
#                     return_options_only=False):

#     manifest_json = """
#     {
#         "version": "1.0.0",
#         "manifest_version": 2,
#         "name": "Chrome Proxy",
#         "permissions": [
#             "proxy",
#             "tabs",
#             "unlimitedStorage",
#             "storage",
#             "<all_urls>",
#             "webRequest",
#             "webRequestBlocking"
#         ],
#         "background": {
#             "scripts": ["background.js"]
#         },
#         "minimum_chrome_version":"22.0.0"
#     }
#     """

#     background_js = """
#     var config = {
#             mode: "fixed_servers",
#             rules: {
#             singleProxy: {
#                 scheme: "http",
#                 host: "%s",
#                 port: parseInt(%s)
#             },
#             bypassList: ["localhost"]
#             }
#         };

#     chrome.proxy.settings.set({value: config, scope: "regular"}, function() {});

#     function callbackFn(details) {
#         return {
#             authCredentials: {
#                 username: "%s",
#                 password: "%s"
#             }
#         };
#     }

#     chrome.webRequest.onAuthRequired.addListener(
#                 callbackFn,
#                 {urls: ["<all_urls>"]},
#                 ['blocking']
#     );
#     """ % (PROXY_HOST, PROXY_PORT, PROXY_USER, PROXY_PASS)

#     chrome_options = webdriver.ChromeOptions()
#     # allow multiple download
#     prefs_experim = {'profile.default_content_setting_values.automatic_downloads': 1}
    
#     if use_proxy:
#         pluginfile = 'proxy_auth_plugin.zip'

#         with zipfile.ZipFile(pluginfile, 'w') as zp:
#             zp.writestr("manifest.json", manifest_json)
#             zp.writestr("background.js", background_js)
#         chrome_options.add_extension(pluginfile)
#     if user_agent:
#         chrome_options.add_argument('--user-agent=%s' % user_agent)
#     if download_folder:
#         prefs_experim["download.default_directory"] = download_folder

#     chrome_options.add_experimental_option("prefs", prefs_experim)

#     if return_options_only:
#         return chrome_options
#     else:
#         driver = webdriver.Chrome(
#             service=Service(chromedriver_path), # executable_path=chromedriver_path,
#             options=chrome_options)
        
#         return driver

    
def get_seduta(CHROMEDRIVER_PATH, GRADUATION_DATE, ESSE3_URL, ROOT_URL, TRIENNALI, config):
    
    if len(GRADUATION_DATE) == 1:
        GRADUATION_DATE=GRADUATION_DATE[0]    # when TRIENNALI=False

    driver=get_chromedriver(chromedriver_path=CHROMEDRIVER_PATH)
    driver.get(ESSE3_URL)
    driver.maximize_window()
    time.sleep(2)

    wait=WebDriverWait(driver, 10)

    # close cookies
    try:
        cookie=wait.until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/div/div[1]/div/div[2]/button[2]')))
        cookie.click()
    except:
        print('\n## Error for cookies')
        raise
    # login button
    try:
        login=wait.until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[2]/div[2]/a[4]")))
        login.click()
    except:
        print('\n## Error for login button')
        raise
    # enter credentials
    try:
        cf=wait.until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div/div[2]/div[2]/div[1]/div/div/div[2]/form/div[1]/input")))
        cf.send_keys(config.Codice_Fiscale)
        pwd=wait.until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div/div[2]/div[2]/div[1]/div/div/div[2]/form/div[2]/input")))
        pwd.send_keys(config.Password)
        time.sleep(1)
        driver.execute_script(f"window.scrollTo(0, {600});")
        time.sleep(1)
        login=wait.until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='auth-internal']/div/div/div[2]/form/div[3]/button")))
        login.click()
    except:
        print('\n## Error for credentials')
        raise
    # select Docente
    try:
        time.sleep(2)
        docente=wait.until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(),'Accedi come DOCENTE')]")))
        docente.click()
    except:
        print('\n## Error for Docente')
        raise
    # open seduta
    try:
        menu=wait.until(
            EC.presence_of_element_located((By.XPATH, "/html/body/header/div/div/a[2]")))
        menu.click()
        time.sleep(1)
        commissioni=wait.until(
                EC.presence_of_element_located((By.XPATH, "/html/body/header/div/nav/div/div[2]/ul/li[4]")))
        commissioni.click()
        time.sleep(1)
        titolo=wait.until(
                EC.presence_of_element_located((By.XPATH, "/html/body/header/div/nav/div/div[2]/ul/li[4]/ul/li[1]")))
        titolo.click()
        time.sleep(1)
    #     seduta=wait.until(
    #             EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table/tbody/tr/td[4]/a/img")))
    #     seduta.click()
        seduta=wait.until(
                        EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table")))
        seduta=convert(BeautifulSoup(seduta.get_attribute('outerHTML'), 'html.parser'))
        link_list=[]
        if TRIENNALI == False:
            for sed in seduta['table'][0]['tbody']:
                if GRADUATION_DATE in sed['#text']:
                    link_list.append(ROOT_URL + seduta['table'][0]['tbody'][0]['tr'][0]['td'][3]['a'][0]['@href'])
    #                 driver.get(ROOT_URL + link)
        else:
            for sed in seduta['table'][0]['tbody'][0]['tr']:
                if any(x in sed['#text'] for x in GRADUATION_DATE):
                    for td in sed['td']:
                        if 'a' in td.keys():
                            link_list.append(ROOT_URL + td['a'][0]['@href'])
    #                         driver.get(ROOT_URL + link)

    except:
        print('\n## Error for seduta')
        raise
        
    return driver, link_list


def get_studenti(driver, link_list, MAPPING_CORSO):

    df_studenti=pd.DataFrame()
    for link_seduta in link_list:

        driver.execute_script(f"window.open('{link_seduta}')")
        driver.switch_to.window(driver.window_handles[-1])
        wait=WebDriverWait(driver, 10)

        # get info seduta
        try:
            time.sleep(1)
            info_seduta=wait.until(
                    EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/div[1]")))
            info_seduta=convert(BeautifulSoup(info_seduta.get_attribute('outerHTML'), 'html.parser'))
            data_ora=info_seduta['div'][0]['dl'][0]['dd'][0]['navigablestring'][0].encode('ascii', 'replace').decode().replace('?','')
            luogo=info_seduta['div'][0]['dl'][0]['dd'][1]['navigablestring'][0].encode('ascii', 'replace').decode().replace('?','')
        except:
            print('\n## Error for info seduta')
            raise
            
        # get commissione
        try:
            time.sleep(1)
            commissione=wait.until(
                    EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table[1]")))
            commissione=convert(BeautifulSoup(commissione.get_attribute('outerHTML'), 'html.parser'))
            comm_list=commissione['table'][0]['tbody'][0]['tr']
            df_commissione=pd.DataFrame()
            for i, com in enumerate(comm_list):
                df_commissione=pd.concat([df_commissione,
                                          pd.DataFrame({'Docente': com['td'][0]['#text'], 'Ruolo': com['td'][1]['#text']}, index=[i+1])])
        except:
            print('\n## Error for commissione')
            raise
            
        print(f'\n\n########################### {data_ora}  ({luogo}) ###########################')
        print(f'\n- Found {len(comm_list)} in commissione: ')
        display(df_commissione)
        file_name='.\\Sedute\\Seduta Laurea '+'_'.join(data_ora.split(' ore ')[0].split('/')[::-1] + [data_ora.split(' ore ')[1].replace(':', '.')])+'_commissione.csv'
        df_commissione.to_csv(file_name, index=False, sep=';')
        print('\nFile saved in', file_name)

        # get student list
        if "footable-page-link" not in driver.page_source:     # single page
            try:
            #     driver.execute_script(f"window.scrollTo(0, {600});")   # scroll down to let the table appear
                time.sleep(1)
                studenti=wait.until(
                        EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table[2]")))
                studenti=convert(BeautifulSoup(studenti.get_attribute('outerHTML'), 'html.parser'))
                stud_list=studenti['table'][0]['tbody'][0]['tr']
                print(f'\n- Found {len(stud_list)} students')
            except:
                print('\n\n\n## Error for lista studenti pagina singola')
                raise
        else:                                                  # multiple pages
            try:
                expected_rows = int(driver.find_element(By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table[2]/tfoot/tr/td/div/span").get_attribute("textContent").split(' di ')[-1])
                num_pages = int(driver.find_element(By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table[2]/tfoot/tr/td/div/span").get_attribute("textContent").split(',')[0].split('di')[-1])

                print(f'\n- Found {num_pages} pages')

                stud_list = []
                for i in range(num_pages):

                    time.sleep(1)
                    studenti=wait.until(
                            EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table[2]")))
                    studenti=convert(BeautifulSoup(studenti.get_attribute('outerHTML'), 'html.parser'))
                    stud_list.extend(studenti['table'][0]['tbody'][0]['tr'])

                    if i < (num_pages-1):
                        driver.find_element(By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table[2]/tfoot/tr/td/div/ul/li[5]/a").click()
                        time.sleep(2)

                print(f'\n- Found {len(stud_list)} students')
                if len(stud_list) != expected_rows:
                    print(f'\n\n\n## Error for lista studenti. Trovati {len(stud_list)}, attesi {expected_rows}')
            except:
                print('\n\n\n## Error for lista studenti pagine multiple')
                raise

        df_studenti_t=pd.DataFrame()
        start_tab_handle=len(driver.window_handles)
        driver.switch_to.window(driver.window_handles[start_tab_handle-1])
        for i, st in enumerate(stud_list):

            print(f"Reading student {i+1}/{len(stud_list)}", end = "\r")

            link=st['td'][10]['a'][0]['@href']
            driver.execute_script(f"window.open('{link}')")
            new_tab= driver.window_handles[i+start_tab_handle]
            driver.switch_to.window(new_tab)

            # get anno iscrizione
            iscrizione=wait.until(
                        EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table[1]")))
            iscrizione=convert(BeautifulSoup(iscrizione.get_attribute('outerHTML'), 'html.parser'))
            anno_iscrizione=[]
            for el in iscrizione['table'][0]['tbody'][0]['tr']:
                anno_iscrizione.append(el['#text'])

            # get punti tesi
            media=wait.until(
                        EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/div[3]")))
            media=convert(BeautifulSoup(media.get_attribute('outerHTML'), 'html.parser'))
            crediti=media['div'][0]['dl'][0]['dd'][3]['#text'].replace('\xa0', '').replace('\u200b', '')
            crediti_tesi=media['div'][0]['dl'][0]['dd'][4]['#text'].replace('\xa0', '').replace('\u200b', '')
            media_pesata=media['div'][0]['dl'][0]['dd'][5]['#text'].replace('\xa0', '').replace('\u200b', '')

            # get voto proposto
            voto=wait.until(
                        EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/div[4]")))
            voto=convert(BeautifulSoup(voto.get_attribute('outerHTML'), 'html.parser'))
            df_voto=pd.DataFrame(index=[i])
            for ind, el in enumerate(voto['div'][0]['dl'][0]['dt']):
                el_name=el['#text']
                el_val=voto['div'][0]['dl'][0]['dd'][ind]['navigablestring']
                el_val='\n'.join(el_val).replace('\xa0', '').replace('\u200b', '')
                if el_name == 'Voto proposto':
                    el_val=int(el_val) 
                df_voto[el_name]=[el_val]

            add_row=pd.concat([pd.DataFrame({'Nome': st['td'][0]['#text'],
                                             'Matricola': st['td'][1]['#text'],
                                             'Nascita': st['td'][2]['#text'],
                                             'Corso': st['td'][3]['#text'],
                                             'Titolo tesi': st['td'][4]['#text'],
                                             'Relatore': st['td'][5]['#text'],
                                             'Anno iscrizione': '\n'.join(anno_iscrizione),
                                             'Crediti': crediti,
                                             'Crediti per tesi': crediti_tesi,
                                             'Media pesata': media_pesata.replace(' - ', '\n'),
                                             'link': link}, index=[i]),
                               df_voto], axis=1)

            df_studenti_t=pd.concat([df_studenti_t, add_row])
            df_studenti_t['Punti tesi']=''
            df_studenti_t['Lode']=''
            df_studenti_t['Encomio']=''
            df_studenti_t['Voto Finale']=0
            move_col = df_studenti_t.pop('link')
            df_studenti_t.insert(df_studenti_t.shape[1], 'link', move_col)
        df_studenti_t.insert(0, 'Luogo', luogo)
        df_studenti_t.insert(0, 'Data', data_ora)
        df_studenti_t['link_seduta']=link_seduta
        df_studenti_t['file_commissione']=file_name
        df_studenti=pd.concat([df_studenti, df_studenti_t])
    df_studenti=df_studenti.merge(MAPPING_CORSO, on='Corso', how='left')
    display(df_studenti.head(5))
    display(df_studenti.groupby(['Data', 'Corso_lab']).size().to_frame())
    
    joblib.dump(df_studenti, 'df_studenti.pkl', compress=('lzma', 3))
    
    return df_studenti


def export_triennali(df_studenti, QUERY_TRIENNALI, GRADUATION_DATE, TEMPESTIVITA_YEAR):


    df_check=pd.read_excel(os.path.join('Query', QUERY_TRIENNALI))
    df_check['MATRICOLA']=df_check['MATRICOLA'].astype(str)
    df_studenti=df_studenti.merge(df_check[['NOME', 'COGNOME', 'MATRICOLA', 'PUNTI_TOTALI', 'AA_IMM_SU']], left_on='Matricola', right_on='MATRICOLA', how='left')
    if df_studenti.isna().sum().sum() > 0:
        print('#### NAs in df_studenti')
        raise

    # check voto proposto
    df_studenti['check voto']=df_studenti['Voto proposto']==round(df_studenti['PUNTI_TOTALI'])
    cc=df_studenti[df_studenti['check voto'] == False]
    if len(cc) > 0:
        print(f'#### found {len(cc)} rows with mismatch in Voto Proposto')
        display(cc)
    # check anno immatricolazione
    df_studenti['check anno iscrizione']=df_studenti['Anno iscrizione'].apply(lambda x: int(x.split('\n')[0].split(' ')[1][:4]))==df_studenti['AA_IMM_SU']
    cc=df_studenti[df_studenti['check anno iscrizione'] == False]
    if len(cc) > 0:
        print(f'#### found {len(cc)} rows with mismatch in Anno iscrizione')
        display(cc)
    # add tempestività
    df_studenti['Tempestività']=np.where(df_studenti['AA_IMM_SU'] >= int(TEMPESTIVITA_YEAR[:4]), 'Yes', 'No')
    display(df_studenti.groupby('Tempestività').size().to_frame())
    main_name='Seduta Laurea '+'_'.join(GRADUATION_DATE[0].split('/')[::-1])+','+','.join([x[:2] for x in GRADUATION_DATE[1:]])
    df_studenti.to_csv(os.path.join('Sedute', main_name+'_studenti.csv'), index=False, sep=';')
    print('\nStudenti list saved in', os.path.join('Sedute', main_name+'_studenti.csv'))
    
    # create template
    file_name=os.path.join('Sedute', main_name+'.xlsx')
    month_num=int(GRADUATION_DATE[0].split('/')[1])
    dates_num=[str(int(x.split('/')[0])) for x in GRADUATION_DATE]
    year=GRADUATION_DATE[0].split('/')[2]

    df_email=pd.DataFrame()
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    for sed in df_studenti['Data'].unique():
        df_studenti_sed=df_studenti[df_studenti['Data']==sed].copy()

        date_full=datetime.strptime(sed.split(' ore')[0], '%d/%m/%Y')
        day_week=calendar.day_name[date_full.weekday()]
        sheet_name=day_week[:3].capitalize()+' '+sed.split('/')[0]+' '+sed.split('ore ')[1].replace(':', '.')
        df_commissione=pd.read_csv(df_studenti_sed['file_commissione'].iloc[0], sep=';')
        pres=df_commissione[df_commissione['Ruolo']=='Presidente']['Docente'].values[0]
        procl=df_commissione[df_commissione['Ruolo']=='Membro Effettivo']['Docente'].values[0]
        segr=df_commissione[df_commissione['Ruolo']=='Segretario']['Docente'].values[0]
        df_email=pd.concat([df_email, pd.DataFrame({'Docente': [pres, procl, segr], 'Ruolo': ['Presidente', 'Proclamatore', 'Segretario']})])

        pd.DataFrame().to_excel(writer, sheet_name=sheet_name, startcol=1, startrow=18, header=True, index=False)  # initialize empty sheet

        workbook  = writer.book
        worksheet = writer.sheets[sheet_name]
        worksheet.write(1, 0, 'DIPARTIMENTO DI SCIENZE ECONOMICHE E AZIENDALI',
                        workbook.add_format({'bold': True, 'italic': True, 'font_size': 14, 'font_name': 'Arial'}))
        worksheet.write(3, 0, 'SEDUTA DI LAUREA '+'-'.join(dates_num)+' '+calendar.month_name[month_num].upper()+' '+year,
                       workbook.add_format({'bold': True, 'color': '#08a129', 'font_size': 14, 'font_name': 'Arial'}))
        worksheet.write(5, 0, 'CORSI DI LAUREA TRIENNALI',
                       workbook.add_format({'bold': True, 'color': '#092273', 'font_size': 14, 'font_name': 'Arial'}))
        worksheet.write(7, 0, 'Presidente '+pres+' Segretario '+segr,
                       workbook.add_format({'bold': True, 'color': '#476ce6', 'font_size': 14, 'font_name': 'Arial'}))
        worksheet.write(9, 0, 'Proclamatore '+procl,
                       workbook.add_format({'bold': True, 'color': '#476ce6', 'font_size': 14, 'font_name': 'Arial'}))
        worksheet.write(9, 6, day_week.upper().replace('Ì', "I'")+' '+str(date_full.day)+' - '+df_studenti_sed['Luogo'].iloc[0].upper() ,
                       workbook.add_format({'bold': True, 'color': '#bd1136', 'font_size': 14, 'font_name': 'Arial'}))

        worksheet.set_column('B:D', 30)
        worksheet.set_column('E:E', 11)
        worksheet.set_column('F:I', 20)
        worksheet.set_column('J:J', 10)
        worksheet.set_column('K:L', 13)
        worksheet.set_column('M:M', 12)

        start_row=11
        for sed_corso in df_studenti_sed['Corso_lab'].unique():
            df_studenti_corso=df_studenti_sed[df_studenti_sed['Corso_lab']==sed_corso].copy()
            worksheet.write(start_row, 0, sed_corso,
                           workbook.add_format({'bold': True, 'color': '#082ca1', 'font_size': 14, 'font_name': 'Arial'}))
            worksheet.write(start_row, 6, 'PROCLAMAZIONE ORE '+sed.split('ore ')[1],
                           workbook.add_format({'bold': True, 'color': 'red', 'font_size': 14, 'font_name': 'Arial'}))
            worksheet.write(start_row+2, 0, 'Numero candidati: '+str(len(df_studenti_corso)),
                           workbook.add_format({'bold': True, 'color': '#092273', 'font_size': 14, 'font_name': 'Arial'}))

            df=df_studenti_corso[['NOME', 'COGNOME', 'Relatore']].copy().rename(columns={'NOME': 'Nome', 'COGNOME': 'Cognome'})
            df['Punteggio iniziale']=df_studenti_corso['Voto proposto']
            df['Punti per attività automatico (3 punti PRIMA del 17/18)']=np.where(df_studenti_corso['AA_IMM_SU'] <= 2017, 3, 0)
            df['Punti tesi (da 0 a 2 punti PRIMA del 17/18)']=np.where(df_studenti_corso['AA_IMM_SU'] <= 2017, 2, 0)
            df['Punti tesi (da 0 a 5 punti immatricolati DAL 17/18)']=np.where(df_studenti_corso['AA_IMM_SU'] > 2017, 5, 0)
            df['Punti tempestività (laurea entro 3 anni solari da anno 1^ immatricolazione 2 punti)']=np.where(df_studenti_corso['Tempestività'] == 'Yes', 2, 0)
            df['Totale']=df.sum(numeric_only = True, axis=1)
            df['Lode automatica (totale >= 112)']=np.where(df['Totale'] >= 112, 'Sì', 'No')
        #     df['Lode su richiesta relatore (totale=111)']=np.where(df['Totale'] == 111, 'Sì', 'No')
            df['Voto di Laurea']=''
            df=df.replace(0, None)
            df.to_excel(writer, sheet_name=sheet_name, startcol=1, startrow=start_row+5, header=True, index=False, na_rep='')
            for j, col in enumerate(df.columns):
                frm1 = workbook.add_format({'bold': True, 'italic': True, 'text_wrap' : True, 'align': 'top', 'font_name': 'Book Antiqua', 'font_size': 12})
                frm1.set_font_color('#c76408')
                if j > 2 and j < 11:
                    frm1.set_bg_color('#faef91')
                frm1.set_border()
                worksheet.write(start_row+5, j+1, col, frm1)
            worksheet.set_row(start_row+5, 90)
            cell_format = workbook.add_format()
            cell_format.set_border()
            for col_num, col in enumerate(df.columns.values):
                for row_num, value in enumerate(df[col].values):
                    worksheet.write(start_row+5+row_num+1, 1 + col_num, value, cell_format)
            start_row+=len(df)+11
            df_t=df[['Relatore']].rename(columns={'Relatore': 'Docente'})
            df_t['Ruolo']='Relatore'
            df_email=pd.concat([df_email, df_t])

    writer.close()
    df_email=df_email.drop_duplicates().sort_values(by=['Ruolo', 'Docente'])
    df_email.to_csv(os.path.join('Sedute', main_name+'_email.csv'), index=False, sep=';')
    
    print('\nFile saved in', file_name)
    print('\nEmail list saved in', os.path.join('Sedute', main_name+'_email.csv'))