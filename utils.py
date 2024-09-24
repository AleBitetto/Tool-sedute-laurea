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
from urllib.parse import urljoin
import re
import time
import joblib
import calendar
from termcolor import colored
import config



def get_chromedriver(chromedriver_path=None):
    
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--lang=it")
    
    chrome_install = ChromeDriverManager().install()
    folder = os.path.dirname(chrome_install)
    chromedriver_path = os.path.join(folder, "chromedriver.exe")
    driver = webdriver.Chrome(service=ChromeService(chromedriver_path),
                             options=chrome_options)
    
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


def login_and_profilo(driver, login_as='DOCENTE'):

    wait=WebDriverWait(driver, 10)
    
    # enter credentials
    try:
        cf=wait.until(
            EC.presence_of_element_located((By.ID, "username")))
#             EC.presence_of_element_located((By.XPATH, "/html/body/div/div[2]/div[2]/div[1]/div/div/div[2]/form/div[1]/input")))
        cf.send_keys(config.Codice_Fiscale)
        pwd=wait.until(
            EC.presence_of_element_located((By.ID, "password")))
#             EC.presence_of_element_located((By.XPATH, "/html/body/div/div[2]/div[2]/div[1]/div/div/div[2]/form/div[2]/input")))
        pwd.send_keys(config.Password)
        time.sleep(1)
        driver.execute_script(f"window.scrollTo(0, {600});")
        time.sleep(1)
        login=wait.until(            
            EC.presence_of_element_located((By.XPATH, "//button[text()='Accedi']")))
#             EC.presence_of_element_located((By.XPATH, "//*[@id='auth-internal']/div/div/div[2]/form/div[3]/button")))
        login.click()
    except:
        print(colored('\n## Error for credentials', 'black', 'on_yellow'))
        raise
    # close cookies
    time.sleep(3)
    try:
        cookie=wait.until(            
            EC.presence_of_element_located((By.XPATH, "//button[text()='Accetta tutti']")))
#             EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/div/div[1]/div/div[2]/button[2]')))
        cookie.click()
    except:
        print(colored('\n## Error for cookies', 'black', 'on_yellow'))
    # select Docente
    try:
        time.sleep(2)
        docente=wait.until(
            EC.presence_of_element_located((By.XPATH, f"//a[contains(text(),'Accedi come {login_as}')]")))
        docente.click()
    except:
        print(colored('\n## Error for Docente', 'black', 'on_yellow'))
        
        
def get_seduta(GRADUATION_DATE, ESSE3_URL, ROOT_URL, TRIENNALI):
    
    if len(GRADUATION_DATE) == 1:
        GRADUATION_DATE=GRADUATION_DATE[0]    # when TRIENNALI=False

    driver=get_chromedriver()
    driver.get(ESSE3_URL)
    driver.maximize_window()
    time.sleep(2)

    wait=WebDriverWait(driver, 10)

#     # login button
#     try:
#         login=wait.until(
#             EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[2]/div[2]/a[4]")))
#         login.click()
#     except:
#         print('\n## Error for login button')
#         raise
    # login, cookies and profilo Docente
    login_and_profilo(driver, login_as='DOCENTE')
    # open seduta
    try:
        menu=wait.until(
            EC.presence_of_element_located((By.ID, "hamburger")))
#             EC.presence_of_element_located((By.XPATH, "/html/body/header/div/div/a[2]")))
        menu.click()
        time.sleep(1)
        commissioni=wait.until(
            EC.presence_of_element_located((By.ID, "menu_link-navbox_docenti_Commissioni")))        
#             EC.presence_of_element_located((By.XPATH, "/html/body/header/div/nav/div/div[2]/ul/li[4]")))
        commissioni.click()
        time.sleep(1)
        titolo=wait.until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(),'Conseguimento titolo')]")))
#             EC.presence_of_element_located((By.XPATH, "/html/body/header/div/nav/div/div[2]/ul/li[4]/ul/li[1]")))
        titolo.click()
        time.sleep(1)
#         seduta=wait.until(
#                 EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table/tbody/tr/td[4]/a/img")))
#         seduta.click()
        seduta=wait.until(
                        EC.presence_of_element_located((By.ID, "seduteAperte")))  
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
        print(colored('\n## Error for seduta', 'black', 'on_yellow'))
        raise
        
    return driver, link_list


def get_studenti(driver, link_list, MAPPING_CORSO, GRADUATION_DATE):

    all_drivers = []
    df_studenti=pd.DataFrame()
    for link_seduta in link_list:

        # open new window
        driver = get_chromedriver()            
        driver.maximize_window()
        driver.get(link_seduta)
        login_and_profilo(driver, login_as='DOCENTE')
        driver.get(link_seduta)
#             driver.execute_script(f"window.open('{link_seduta}')")
#             driver.switch_to.window(driver.window_handles[-1])
        all_drivers.append(driver)    # so to keep all drivers open
        
        wait=WebDriverWait(driver, 10)

        # get info seduta
        try:
            time.sleep(1)
            info_seduta=wait.until(
                EC.presence_of_element_located((By.ID, "seduta")))
#                     EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/div[1]")))
            info_seduta=convert(BeautifulSoup(info_seduta.get_attribute('outerHTML'), 'html.parser'))
            data_ora=info_seduta['div'][0]['dl'][0]['dd'][0]['navigablestring'][0].encode('ascii', 'replace').decode().replace('?','')
            luogo=info_seduta['div'][0]['dl'][0]['dd'][1]['navigablestring'][0].encode('ascii', 'replace').decode().replace('?','')
        except:
            print(colored('\n## Error for info seduta', 'black', 'on_yellow'))
            raise
            
        # get commissione
        try:
            time.sleep(1)
            commissione=wait.until(
                EC.presence_of_element_located((By.ID, "commissione")))
#                     EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table[1]")))
            commissione=convert(BeautifulSoup(commissione.get_attribute('outerHTML'), 'html.parser'))
            comm_list=commissione['table'][0]['tbody'][0]['tr']
            df_commissione=pd.DataFrame()
            for i, com in enumerate(comm_list):
                df_commissione=pd.concat([df_commissione,
                                          pd.DataFrame({'Docente': com['td'][0]['#text'], 'Ruolo': com['td'][1]['#text']}, index=[i+1])])
        except:
            print(colored('\n## Error for commissione', 'black', 'on_yellow'))
            raise
            
        print(f'\n\n########################### {data_ora}  ({luogo}) ###########################')
        print(f'\n- Found {len(comm_list)} in commissione: ')
        display(df_commissione)
        file_name='.\\Sedute\\Log\\Seduta Laurea '+'_'.join(data_ora.split(' ore ')[0].split('/')[::-1] + [data_ora.split(' ore ')[1].replace(':', '.')])+'_commissione.csv'
        df_commissione['Presente']=''
        df_commissione.to_csv(file_name, index=False, sep=';')
        print('\nFile saved in', file_name)

        # get student list
        if "footable-page-link" not in driver.page_source:     # single page
            try:
            #     driver.execute_script(f"window.scrollTo(0, {600});")   # scroll down to let the table appear
                time.sleep(1)
                studenti=wait.until(
                    EC.presence_of_element_located((By.ID, "elencoLaureandi")))
#                         EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table[2]")))
                studenti=convert(BeautifulSoup(studenti.get_attribute('outerHTML'), 'html.parser'))
                stud_list=studenti['table'][0]['tbody'][0]['tr']
                print(f'\n- Found {len(stud_list)} students')
            except:
                print(colored('\n\n\n## Error for lista studenti pagina singola', 'black', 'on_yellow'))
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
                        EC.presence_of_element_located((By.ID, "elencoLaureandi")))
#                             EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table[2]")))
                    studenti=convert(BeautifulSoup(studenti.get_attribute('outerHTML'), 'html.parser'))
                    stud_list.extend(studenti['table'][0]['tbody'][0]['tr'])

                    if i < (num_pages-1):
                        time.sleep(1)
                        driver.find_element(By.CSS_SELECTOR, 'li[data-page="next"] a').click()    # next page
                        time.sleep(2)

                print(f'\n- Found {len(stud_list)} students')
                if len(stud_list) != expected_rows:
                    print(colored(f'\n\n\n## Error for lista studenti. Trovati {len(stud_list)}, attesi {expected_rows}', 'black', 'on_yellow'))
            except:
                print(colored('\n\n\n## Error for lista studenti pagine multiple', 'black', 'on_yellow'))
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
                EC.presence_of_element_located((By.ID, "grad-dettLau-iscrizioni")))
#                         EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table[1]")))
            iscrizione=convert(BeautifulSoup(iscrizione.get_attribute('outerHTML'), 'html.parser'))
            anno_iscrizione=[]
            for el in iscrizione['table'][0]['tbody'][0]['tr']:
                anno_iscrizione.append(el['#text'])

            # get punti tesi
            media=wait.until(
                EC.presence_of_element_located((By.ID, "grad-dettLau-tesi")))
#                         EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/div[3]")))
            media=convert(BeautifulSoup(media.get_attribute('outerHTML'), 'html.parser'))
            crediti=media['div'][0]['dl'][0]['dd'][3]['#text'].replace('\xa0', '').replace('\u200b', '')
            crediti_tesi=media['div'][0]['dl'][0]['dd'][4]['#text'].replace('\xa0', '').replace('\u200b', '')
            media_pesata=media['div'][0]['dl'][0]['dd'][5]['#text'].replace('\xa0', '').replace('\u200b', '')

            # get voto proposto
            voto=wait.until(
                EC.presence_of_element_located((By.ID, "grad-dettLau-boxVerbalizzazione")))
#                         EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/div[4]")))
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
            time.sleep(0.5)
        df_studenti_t.insert(0, 'Luogo', luogo)
        df_studenti_t.insert(0, 'Data', data_ora)
        df_studenti_t['link_seduta']=link_seduta
        df_studenti_t['file_commissione']=file_name
        df_studenti=pd.concat([df_studenti, df_studenti_t])
    df_studenti=df_studenti.merge(MAPPING_CORSO, on='Corso', how='left')
    print(f'\n\nTotal students: {len(df_studenti)}')
    display(df_studenti.head(5))
    display(df_studenti.groupby(['Data', 'Corso_lab']).size().to_frame())
   
    file_name_pkl=os.path.join('Checkpoints', 'Seduta Laurea '+'_'.join(GRADUATION_DATE[0].split('/')[::-1])+','+','.join([x[:2] for x in GRADUATION_DATE[1:]])+' - df_studenti.pkl')
    joblib.dump(df_studenti, file_name_pkl, compress=('lzma', 3))
    print('\nLog studenti saved in', file_name_pkl)
                               
    
    return df_studenti, all_drivers


def export_triennali(df_studenti, QUERY_TRIENNALI, GRADUATION_DATE, TEMPESTIVITA_YEAR, final_version):

    QUERY_REQUIRED_COLUMNS = ['NOME', 'COGNOME', 'MATRICOLA', 'PUNTI_TOTALI', 'AA_IMM_SU']
    
    main_name='Seduta Laurea '+'_'.join(GRADUATION_DATE[0].split('/')[::-1])+','+','.join([x[:2] for x in GRADUATION_DATE[1:]])

    df_check=pd.read_excel(os.path.join('Query', QUERY_TRIENNALI))
    if not all([x in df_check.columns for x in QUERY_REQUIRED_COLUMNS]):
        print(colored(f'\n## Missing columns in "{QUERY_TRIENNALI}":', 'black', 'on_yellow'),
             '\n   -', '\n   - '.join([x for x in QUERY_REQUIRED_COLUMNS if x not in df_check.columns]))
        raise
    df_check['MATRICOLA']=df_check['MATRICOLA'].astype(str)
    df_studenti=df_studenti.merge(df_check[['NOME', 'COGNOME', 'MATRICOLA', 'PUNTI_TOTALI', 'AA_IMM_SU']], left_on='Matricola', right_on='MATRICOLA', how='left')

    if not final_version:    
        if df_studenti.isna().sum().sum() > 0:
            print(colored('#### NAs in df_studenti', 'black', 'on_yellow'))
            raise

        # check voto proposto
        df_studenti['check voto']=df_studenti['Voto proposto']==df_studenti['PUNTI_TOTALI'].apply(lambda x: min(round(x), 110))
        cc=df_studenti[df_studenti['check voto'] == False]
        if len(cc) > 0:
            print(colored(f'#### found {len(cc)} rows with mismatch between "PUNTI_TOTALI" and "Voto Proposto"', 'black', 'on_yellow'))
            display(cc)
        # check anno immatricolazione
        df_studenti['check anno iscrizione']=df_studenti['Anno iscrizione'].apply(lambda x: int(x.split('\n')[0].split(' ')[1][:4]))==df_studenti['AA_IMM_SU']
        cc=df_studenti[df_studenti['check anno iscrizione'] == False]
        if len(cc) > 0:
            print(colored(f'#### found {len(cc)} rows with mismatch between "AA_IMM_SU" and "Anno iscrizione"', 'black', 'on_yellow'),
                 '   ->  Keeping information from Esse3')
            for count, row in cc.reset_index(drop=True).iterrows():
                print(count+1, '-', row['Nome'],' (Matricola', row['Matricola']+') - Corso:',row['Corso'])
                print('   AA_IMM_SU:', row['AA_IMM_SU'])
                print('   Da esse3:', '\n      '+row['Anno iscrizione'].replace('\n', '\n      '), end='\n\n')
        df_studenti['AA_IMM_SU_ESSE3']=df_studenti['Anno iscrizione'].apply(lambda x: int(x.split('\n')[0].split(' ')[1][:4]))
        check_range = [2000, 2030]
        if df_studenti['AA_IMM_SU_ESSE3'].unique().max() > check_range[1] or df_studenti['AA_IMM_SU_ESSE3'].unique().min() < check_range[0]:
            print(colored(f'#### Check "AA_IMM:SU_ESSE33", values outside range {check_range[0]}-{check_range[1]}', 'black', 'on_yellow'))
            display(df_studenti[(df_studenti['AA_IMM_SU_ESSE3'] < check_range[0]) | (df_studenti['AA_IMM_SU_ESSE3'] > check_range[1])]['AA_IMM_SU_ESSE3'].value_counts().to_frame())
            raise        
        # add tempestività
        df_studenti['Tempestività']=np.where(df_studenti['AA_IMM_SU_ESSE3'] >= int(TEMPESTIVITA_YEAR[:4]), 'Yes', 'No')
        display(df_studenti.groupby('Tempestività').size().to_frame())
        df_studenti.to_csv(os.path.join('Sedute', 'Log', main_name+'_studenti.csv'), index=False, sep=';')
        print('\nStudenti list saved in', os.path.join('Sedute', 'Log', main_name+'_studenti.csv'))

    # create template
    file_name=os.path.join('Sedute', main_name+'.xlsx')
    month_num=int(GRADUATION_DATE[0].split('/')[1])
    dates_num=[str(int(x.split('/')[0])) for x in GRADUATION_DATE]
    year=GRADUATION_DATE[0].split('/')[2]

    df_email=pd.DataFrame()
    if not final_version:
        writer=pd.ExcelWriter(file_name, engine='xlsxwriter')
    for sed in df_studenti['Data'].unique():
        df_studenti_sed=df_studenti.copy()[df_studenti['Data']==sed]

        date_full=datetime.strptime(sed.split(' ore')[0], '%d/%m/%Y')
        day_week=calendar.day_name[date_full.weekday()]
        sheet_name=day_week[:3].capitalize()+' '+sed.split('/')[0]+' '+sed.split('ore ')[1].replace(':', '.')
        df_commissione=pd.read_csv(df_studenti_sed['file_commissione'].iloc[0], sep=';')
        pres=df_commissione[df_commissione['Ruolo']=='Presidente']['Docente'].values[0]
        procl=df_commissione[df_commissione['Ruolo']=='Membro Effettivo']['Docente'].values[0]
        segr=df_commissione[df_commissione['Ruolo']=='Segretario']['Docente'].values[0]
        df_email=pd.concat([df_email, pd.DataFrame({'Docente': [pres, procl, segr], 'Ruolo': ['Presidente', 'Proclamatore', 'Segretario']})])

        if final_version:
            file_name_final=os.path.join('Sedute', main_name+' '+procl.split(' ')[0]+'.xlsx')
            writer=pd.ExcelWriter(file_name_final, engine='xlsxwriter')
            df_read=pd.read_excel(os.path.join('Sedute', main_name+'.xlsx'), sheet_name=sheet_name)

        pd.DataFrame().to_excel(writer, sheet_name=sheet_name, startcol=1, startrow=1, header=True, index=False)  # initialize empty sheet
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
        worksheet.write(9, 4 if final_version else 6, day_week.upper().replace('Ì', "I'")+' '+str(date_full.day)+' - '+df_studenti_sed['Luogo'].iloc[0].upper() ,
                       workbook.add_format({'bold': True, 'color': '#bd1136', 'font_size': 14, 'font_name': 'Arial'}))

        worksheet.set_column('B:D', 30)
        worksheet.set_column('E:E', 11)
        if not final_version:
            worksheet.set_column('F:I', 20)
            worksheet.set_column('J:J', 10)
            worksheet.set_column('K:L', 13)
            worksheet.set_column('M:M', 12)

        start_row=11
        for sed_corso in df_studenti_sed['Corso_lab'].unique():
            df_studenti_corso=df_studenti_sed.copy()[df_studenti_sed['Corso_lab']==sed_corso]
            worksheet.write(start_row, 0, sed_corso,
                           workbook.add_format({'bold': True, 'color': '#082ca1', 'font_size': 14, 'font_name': 'Arial'}))
            worksheet.write(start_row, 4 if final_version else 6, 'PROCLAMAZIONE ORE '+sed.split('ore ')[1],
                           workbook.add_format({'bold': True, 'color': 'red', 'font_size': 14, 'font_name': 'Arial'}))
            worksheet.write(start_row+2, 0, 'Numero candidati: '+str(len(df_studenti_corso)),
                           workbook.add_format({'bold': True, 'color': '#092273', 'font_size': 14, 'font_name': 'Arial'}))

            df=df_studenti_corso[['NOME', 'COGNOME', 'Relatore']].copy().rename(columns={'NOME': 'Nome', 'COGNOME': 'Cognome'})
            if not final_version:
                df['Punteggio iniziale']=df_studenti_corso['Voto proposto']
                df['Punti per attività automatico (3 punti PRIMA del 17/18)']=np.where(df_studenti_corso['AA_IMM_SU_ESSE3'] <= 2017, 3, 0)
                df['Punti tesi (da 0 a 2 punti PRIMA del 17/18)']=np.where(df_studenti_corso['AA_IMM_SU_ESSE3'] <= 2017, 2, 0)
                df['Punti tesi (da 0 a 5 punti immatricolati DAL 17/18)']=np.where(df_studenti_corso['AA_IMM_SU_ESSE3'] > 2017, 5, 0)
                df['Punti tempestività (laurea entro 3 anni solari da anno 1^ immatricolazione 2 punti)']=np.where(df_studenti_corso['Tempestività'] == 'Yes', 2, 0)
                df['Totale']=df.sum(numeric_only = True, axis=1)
                df['Lode automatica (totale >= 112)']=''#np.where(df['Totale'] >= 112, 'Sì', 'No')
                df['Lode su richiesta relatore (totale=111)']=''#np.where(df['Totale'] == 111, 'Sì', 'No')
                df['Voto di Laurea']=''
                df=df.replace(0, None)
            else:
                df['Voto di Laurea']=''
                df=df.reset_index(drop=True)
                # find columns
                for i, _ in enumerate(df_read.columns):
                    if 'Voto di Laurea' in df_read.iloc[:, i].values:
                        voto_col=i
                    if 'Nome' in df_read.iloc[:, i].values:
                        nome_col=i
                    if 'Cognome' in df_read.iloc[:, i].values:
                        cognome_col=i
                # get voto finale
                for i, row in df.iterrows():
                    i_loc=np.where((df_read.iloc[:, nome_col].values == row['Nome']) & (df_read.iloc[:, cognome_col].values == row['Cognome']))[0][0]
                    voto=df_read.iloc[i_loc, voto_col]
                    df.iloc[i]['Voto di Laurea']=str(voto)

            df.to_excel(writer, sheet_name=sheet_name, startcol=1, startrow=start_row+5, header=True, index=False, na_rep='')
    #             df_log=pc.concat([df_log, pd.DataFrame({'sed': sed, 'start_row': start_row+5})])
            # format header
            for j, col in enumerate(df.columns):
                frm1 = workbook.add_format({'bold': True, 'italic': True, 'text_wrap' : True, 'align': 'top', 'font_name': 'Book Antiqua', 'font_size': 12})
                frm1.set_font_color('#c76408')
                if j > 2 and j < 11:
                    frm1.set_bg_color('#faef91')
                frm1.set_border()
                worksheet.write(start_row+5, j+1, col, frm1)
            if not final_version:
                worksheet.set_row(start_row+5, 90)
            else:
                worksheet.set_row(start_row+5, 40)
            # add cell border
            border_format = workbook.add_format()
            border_format.set_border()
            border_yellow_format = workbook.add_format({'bg_color': 'yellow', 'border': 1})
            for col_num, col in enumerate(df.columns.values):
                for row_num, value in enumerate(df[col].values):
                    # add highlight for Totale = 111
                    if col == 'Totale':
                        if value == 111:
                            worksheet.write(start_row+5+row_num+1, 1 + col_num, value, border_yellow_format)
                    else:    
                        worksheet.write(start_row+5+row_num+1, 1 + col_num, value, border_format)
            # add formula for Totale and Lode     https://xlsxwriter.readthedocs.io/working_with_formulas.html
            row_111 = np.where(df['Totale'] == 111)[0].tolist()
            if not final_version:
                for ind in range(len(df)):
                    row_num=str(start_row+5+ind+2)
                    if ind in row_111:
                        worksheet.write_formula('J'+row_num, '=SUM(E'+row_num+':I'+row_num+')', border_yellow_format)
                    else:
                        worksheet.write_formula('J'+row_num, '=SUM(E'+row_num+':I'+row_num+')', border_format)
                    worksheet.write_formula('K'+row_num, '=IF(J'+row_num+'>=112,"sì","no")', border_format)     # USE comma not semicolon
                    worksheet.write_formula('M'+row_num, '=IF(OR(K'+row_num+'="sì",L'+row_num+'="sì"),"110 e lode",J'+row_num+')', border_format)
            start_row+=len(df)+11
            df_t=df[['Relatore']].rename(columns={'Relatore': 'Docente'})
            df_t['Ruolo']='Relatore'
            df_email=pd.concat([df_email, df_t])
        if final_version:
            writer.close()
            print('File saved in', file_name_final)

    if final_version:
        print('\n\n- Available sessions:\n')
        print('\n'.join(sorted(df_studenti['Data'].unique())))
    else:
        writer.close()
        df_email_list=pd.concat([pd.read_csv(os.path.join('Checkpoints', 'lista_email_docenti.csv'), sep=';'),
                                 pd.read_csv(os.path.join('Checkpoints', 'lista_email_docenti_aggiuntivi_MANUAL.csv'), sep=';')])
        df_email=df_email.drop_duplicates().sort_values(by=['Ruolo', 'Docente'])
        df_email=df_email.merge(df_email_list[['Docente', 'Email']], on='Docente', how='left')
        df_email.to_csv(os.path.join('Sedute', main_name+'_email.csv'), index=False, sep=';')
        print('\nFile saved in', file_name)
        print('\nEmail list saved in', os.path.join('Sedute', main_name+'_email.csv'))


def fill_voti(driver, df_studenti_sed, df_commissione_presenze_sed, ROOT_URL):

    wait=WebDriverWait(driver, 10)
    driver.switch_to.window(driver.window_handles[0])

    for i, row in df_studenti_sed.iterrows():

        print(f"  -Uploading student {i+1}/{len(df_studenti_sed)}", end = "\r")

        link=ROOT_URL+row['link']
        driver.execute_script(f"window.open('{link}')")
        new_tab= driver.window_handles[i+1]
        driver.switch_to.window(new_tab)

        # input punti tesi
        try:
            punti_tesi=wait.until(
                            EC.presence_of_element_located((By.ID, 'grad-dettLau-puntiTesi')))
            punti_tesi.click()
            punti_tesi.clear()
            punti_tesi.send_keys(row['Punti tesi'])
            wait.until(EC.presence_of_element_located((By.ID, 'grad-dettLau-annotazioni'))).click()
        except:
            print(colored(f'\n## Error for punti tesi {row["Nome"]} ({i+1})', 'black', 'on_yellow'))

        # tick Lode and Encomio
        if row['Lode'] != '':
            try:
                lode=wait.until(
                        EC.presence_of_element_located((By.ID, 'grad-dettLau-lode1')))
                lode.click()
            except:
                print(colored(f'\n## Error for Lode {row["Nome"]} ({i+1})', 'black', 'on_yellow'))
        if row['Encomio'] != '':
            try:
                encomio=wait.until(
                        EC.presence_of_element_located((By.ID, 'grad-dettLau-encomio1')))
                encomio.click()
            except:
                print(colored(f'\n## Error for Encomio {row["Nome"]} ({i+1})', 'black', 'on_yellow'))

        # tick Commissione
        try:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            names_list=df_commissione_presenze_sed['Docente'].values

            # check if multiple pages in Commissione checkbox - and check presenti
            ss=BeautifulSoup(driver.page_source, 'html.parser').prettify()
            ind=ss.find('Visualizzati')
            s1=ss[ind:(ind+40)].split('\n')[0]
            tot_elem=s1.split(' di ')[1]
            stop=False
            tot_checked=0
            while True:
                # find names and check
                table_checkbox=driver.find_elements(By.CSS_SELECTOR, "table[id^='gradDettLauCommissione'] td:nth-of-type(2)")
                for i, item in enumerate(table_checkbox):
                    tt=convert(BeautifulSoup(item.get_attribute('outerHTML'), 'html.parser'))
                    for name_present in names_list:
                        if name_present in tt['td'][0]['#text']:
                            prev_sibl=item.find_element(By.XPATH, "preceding-sibling::*[1]")
                            t1=convert(BeautifulSoup(prev_sibl.get_attribute('outerHTML'), 'html.parser'))
                            if t1['td'][0]['@class'] == ['td_cbox']:
                                prev_sibl.click()
                                tot_checked+=1
                            else:
                                print(f'Missing checkbox for {name_present} for {row["NOME"]} {row["COGNOME"]}')
                # check next page
                ss=BeautifulSoup(driver.page_source, 'html.parser').prettify()
                ind=ss.find('Visualizzati')
                s1=ss[ind:(ind+40)].split('\n')[0]
                current_range_max=s1.replace('Visualizzati ', '').replace(' di '+tot_elem, '').split(' - ')[1]
                if current_range_max == tot_elem:    # no more pages
                    break
                else:
                    button_path='/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/form/div[3]/div/div[5]/div/table/tbody/tr/td[2]/table/tbody/tr/td[6]'
                    if driver.find_elements(By.XPATH, button_path):
                        driver.find_element(By.XPATH, button_path).click()
                        time.sleep(1)

            if len(names_list) != tot_checked:
                print(colored(f'\n## Error for checkbox Commissione {row["Nome"]} ({i+1})', 'black', 'on_yellow'))
        except:
            print(colored(f'\n## Error for checkbox Commissione {row["Nome"]} ({i+1})', 'black', 'on_yellow'))

        # save
        try:
            save=wait.until(
                        EC.presence_of_element_located((By.ID, 'grad-dettLau-btnSubmit')))
            save.click()
        except:
            print(colored(f'\n## Error for Save {row["Nome"]} ({i+1})', 'black', 'on_yellow'))

        # save and exit
        try:
            save_exit=wait.until(
                        EC.presence_of_element_located((By.ID, 'grad-dettLau-btnSubmitExit')))
            save_exit.click()
        except:
            print(colored(f'\n## Error for Save and Exit {row["Nome"]} ({i+1})', 'black', 'on_yellow'))

    return driver


def check_voti(driver, df_studenti_sed):

    wait=WebDriverWait(driver, 10)
    driver.switch_to.window(driver.window_handles[-1])
    driver.refresh()
    time.sleep(3)

    if "footable-page-link" not in driver.page_source:     # single page
        final_grade=wait.until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table[2]")))
        final_grade=convert(BeautifulSoup(final_grade.get_attribute('outerHTML'), 'html.parser'))
        elem_list=final_grade['table'][0]['tbody'][0]['tr']
    else:                                                  # multiple pages
        expected_rows = int(driver.find_element(By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table[2]/tfoot/tr/td/div/span").get_attribute("textContent").split(' di ')[-1])
        num_pages = int(driver.find_element(By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table[2]/tfoot/tr/td/div/span").get_attribute("textContent").split(',')[0].split('di')[-1])

        print(f'\n- Found {num_pages} pages')

        elem_list=[]
        for i in range(num_pages):

            time.sleep(1)
    #         final_grade=wait.until(
    #             EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table[2]")))
            final_grade=driver.find_element(By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table[2]")
            final_grade=convert(BeautifulSoup(final_grade.get_attribute('outerHTML'), 'html.parser'))
            elem_list.extend(final_grade['table'][0]['tbody'][0]['tr'])

            if i < (num_pages-1):
                driver.find_element(By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table[2]/tfoot/tr/td/div/ul/li[5]/a").click()
                time.sleep(2)

        print(f'- Found {len(elem_list)} students')
        if len(elem_list) != expected_rows:
            print(colored(f'\n\n\n## Error for lista studenti. Trovati {len(elem_list)}, attesi {expected_rows}', 'black', 'on_yellow'))

    for i, el in enumerate(elem_list):
        grade=el['td'][8]['#text']
        lode = True if 'L' in grade else False
        grade_evaluated=int(grade.replace('L', ''))
        matricola=int(el['td'][1]['#text'])
        name=el['td'][0]['#text']
        grade_expected=min(df_studenti_sed[df_studenti_sed['Matricola'] == matricola]['Punteggio finale'].values[0], 110)
        lode_expected = True if df_studenti_sed[df_studenti_sed['Matricola'] == matricola]['Lode'].values[0] != '' else False
        if grade_evaluated != grade_expected or lode != lode_expected:
            print(f'\n## Expected Votazione Finale mismatch: {name} ({i+1})')

    print('\nVotazione Finale: Check Done')
    
    
def upload_voti(GRADUATION_DATE, ESSE3_URL, ROOT_URL, TRIENNALI, upload_single_session=[]):

    main_name='Seduta Laurea '+'_'.join(GRADUATION_DATE[0].split('/')[::-1])+','+','.join([x[:2] for x in GRADUATION_DATE[1:]])
    df_studenti=pd.read_csv(os.path.join('Sedute', 'Log', main_name+'_studenti.csv'), sep=';')

    if len(upload_single_session) > 0:
        df_studenti=df_studenti[df_studenti['Data'].isin(upload_single_session)]

    # read presenze commissione
    df_commissione_presenze=pd.DataFrame()
    for path_commissione in df_studenti['file_commissione'].unique():
        df_comm_sed=pd.read_csv(path_commissione, sep=';').fillna('')
        sed=df_studenti[df_studenti['file_commissione'] == path_commissione]['Data'].values[0]
        if sum(df_comm_sed['Presente'] != '') == 0:
            print('Nobody present at seduta: ', sed)
            raise
        df_comm_sed=df_comm_sed[df_comm_sed['Presente'] != '']
        print(f'Presence at {sed}: {len(df_comm_sed)}')
        display(df_comm_sed)
        df_comm_sed['Data']=sed
        df_commissione_presenze=pd.concat([df_commissione_presenze, df_comm_sed])

    # read and match final votes
    df_t=pd.DataFrame()
    df_t['NOME']=''
    df_t['COGNOME']=''
    df_t['Punteggio iniziale']=0
    df_t['Punteggio finale']=0
    df_t['Punti tesi']=0
    df_t['Lode']=''
    for sh in pd.ExcelFile(os.path.join('Sedute', main_name+'.xlsx')).sheet_names:
        df_read=pd.read_excel(os.path.join('Sedute', main_name+'.xlsx'), sheet_name=sh)
        for i, _ in enumerate(df_read.columns):
            if 'Voto di Laurea' in df_read.iloc[:, i].values:
                voto_col=i
            if 'Nome' in df_read.iloc[:, i].values:
                nome_col=i
            if 'Cognome' in df_read.iloc[:, i].values:
                cognome_col=i
            if 'Punteggio iniziale' in df_read.iloc[:, i].values:
                iniziale_col=i
            if 'Totale' in df_read.iloc[:, i].values:
                finale_col=i
        for i, row in df_studenti.iterrows():
            i_loc=np.where((df_read.iloc[:, nome_col].values == row['NOME']) & (df_read.iloc[:, cognome_col].values == row['COGNOME']))
            if len(i_loc[0]) > 0:
                punteggio_iniziale=df_read.iloc[i_loc[0][0], iniziale_col]
                punteggio_finale=df_read.iloc[i_loc[0][0], finale_col]
                punti_tesi=punteggio_finale-punteggio_iniziale
                voto_finale=df_read.iloc[i_loc[0][0], voto_col]

                if "lode" in str(voto_finale):
                    voto=110
                    lode='sì'
                else:
                    voto=int(voto_finale)
                    lode=''

                add_row=pd.DataFrame({'NOME': row['NOME'],
                                      'COGNOME': row['COGNOME'],
                                      'Punteggio iniziale': punteggio_iniziale,
                                      'Punti tesi': punti_tesi,
                                      'Lode': lode,
                                      'Encomio': '',
                                      'Punteggio finale': punteggio_finale}, index=[0])
                df_t=pd.concat([df_t, add_row])
    col_drop=[x for x in df_studenti.columns if x in df_t.columns and x not in ['NOME', 'COGNOME']]
    if len(col_drop) > 0:
        df_studenti=df_studenti.drop(columns=col_drop)
    df_studenti=df_studenti.merge(df_t, on=['NOME', 'COGNOME'], how='left')
    if df_studenti[df_t.columns].isna().sum().sum() > 0:
        print(colored('\n## Error for df_studenti when matching the final votes', 'black', 'on_yellow'))
        raise


    # fill voti
    driver_dict={}
    avail_sed=df_studenti['Data'].unique().tolist()
    for sed in avail_sed:

        print('\n\n#####################################')
        print('\nFilling voti for: ', sed, '\n')
        df_studenti_sed=df_studenti[df_studenti['Data'] == sed].copy()
        df_commissione_presenze_sed=df_commissione_presenze[df_commissione_presenze['Data'] == row['Data']]

        # set up page
        driver_dict[sed], _ = get_seduta(GRADUATION_DATE, ESSE3_URL, ROOT_URL, TRIENNALI)

        # fill final grade
        driver_dict[sed] = fill_voti(driver_dict[sed], df_studenti_sed, df_commissione_presenze_sed, ROOT_URL)

        # check final grade mismatch
        check_voti(driver_dict[sed], df_studenti_sed)


def get_emails(ROOT_URL):
    
    # parse html
    file = open("lista_docenti.txt", "r")
    html = '\n'.join(file.readlines())
    file.close()
    tt=convert(BeautifulSoup(html, 'html.parser'))

    df_email=pd.DataFrame()
    elem_list=tt['html'][0]['body'][0]['div'][1]['div'][0]['div'][1]['div'][0]['div'][0]['div'][0]['div'][0]['div'][0]['div'][0]['div'][1]['div']
    for i, el in enumerate(elem_list):

        print(f"Downloading email {i+1}/{len(elem_list)}", end = "\r")

        name=el['#text']
        # get email
        url=urljoin(ROOT_URL, el['div'][0]['a'][0]['@href'])
        page=requests.get(url)
        soup=convert(BeautifulSoup(page.content, 'html.parser'))
        email=soup['html'][0]['body'][0]['div'][1]['div'][0]['div'][2]['div'][0]['div'][0]['div'][0]['ul'][0]['li'][1]['#text']
        add_row=pd.DataFrame({'Docente': name.split(',')[0].strip().upper(),
                              'Ruolo': ', '.join([x.strip() for x in name.split(',')[1:]]),
                              'Email': email,
                             'url': url}, index=[i])
        df_email=pd.concat([df_email, add_row])
    
    df_email.to_csv(os.path.join('Checkpoints', 'lista_email_docenti.csv'), index=False, sep=';')
    print('\n- File saved in', os.path.join('Checkpoints', 'lista_email_docenti.csv'))
    display(df_email.head(5))
    
    return df_email