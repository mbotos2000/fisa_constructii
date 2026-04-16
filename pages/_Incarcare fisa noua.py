from __future__ import print_function
from io import BytesIO
from datetime import *
import streamlit as st
import pandas as pd
from pandas import *
from docx2python import docx2python
import os
import base64
import time
import ftplib
from mailmerge import MailMerge
from difflib import get_close_matches
import pickle
import string
from auth_simple import require_login
name, user = require_login("🔐 Incarcare credentiale")
st.title("Fisa disciplinei")
st.success(f"Bine ai venit, {name}!")
@st.dialog("info")
def info():
 st.write("Aplicația este pusă la dispoziția dumneavoastră pentru a elimina confuziile și neconcordanțele generate de modificările și noile reglementări ARACIS. Formularul _Fișei disciplinei_ a suferit schimbări de structură, iar această aplicație automatizează procesul de actualizare. Toate informațiile introduse de dumneavoastră sunt transpuse automat în noul șablon oficial.")
 st.write("Datele referitoare la disciplină (denumire, număr de credite, tip de examinare, număr de ore, codul disciplinei etc.) sunt preluate direct din planurile de învățământ. Au fost introduse două capitole noi: Competențe și Rezultatele învățării. Pentru fiecare specializare, formulările oficiale se regăsesc în prezentarea planului de învățământ și sunt afișate de aplicație în etapa de completare. Acestea trebuie adaptate pentru fiecare disciplină în secțiunile corespunzătoare ale fișei.")
 st.write("Pentru fișele încărcate în anii anteriori în baza de date, aceste capitole conțin propuneri destinate titularului de curs. Disciplinele aflate la prima completare pot fi încărcate în format .docx, iar aplicația va prelua automat cât mai multe dintre informațiile existente. De asemenea, este posibilă completarea manuală prin câmpurile afișate. Propunerile generate pot fi acceptate, preluate sau editate.")
 st.write("Datele privind aprobarea în departament și consiliu vor fi actualizate automat atunci când devin disponibile. La final, fișa completată poate fi descărcată și vizualizată în formatul oficial aprobat.")
if st.button("Informatii aplicatie"):
	info()

def preprocess(text):
    return text.strip().lower().translate(str.maketrans('', '', string.punctuation))

def find_closest_match_index(word, word_list, cutoff=0.6):
    word = preprocess(word)
    word_list = [preprocess(w) for w in word_list]
    
    closest_matches = get_close_matches(word, word_list, n=1, cutoff=cutoff)
    if closest_matches:
        return word_list.index(closest_matches[0])
    return 0
	
def clean_value(value):
    if pd.isna(value):  # Replaces NaN or None with an empty string
        return ''
    elif isinstance(value, bool):  # Convert boolean values to strings
        return str(value)
    elif isinstance(value, (int, float, str)):  # Keep numbers and strings as they are
        return value
    else:
        return 'Unknown object'  # Handle unrecognized objects by converting them to a string
def fix_encoding(text):
    if isinstance(text, str):
        try:
            return text.encode('latin1').decode('utf-8')  # Fix incorrectly decoded text
        except UnicodeEncodeError:
            return text  # Return text unchanged if no encoding issues
    return text  # If it's not a string, return as is
def get_binary_file_downloader_html(bin_file, file_label='File'):
    with open(bin_file, 'rb') as f:
        data = f.read()
    bin_str = base64.b64encode(data).decode()
    href = f'<a href="data:application/octet-stream;base64,{bin_str}" download="{os.path.basename(bin_file)}">Download {file_label}</a>'
    return href

def strip_last(x):
	return x.strip()

def my_function(x):
  return list(dict.fromkeys(x))

def schimba_1_1(new):
    st.session_state['M_1_1'] = str(new)

def schimba_1_2(new):
    st.session_state['M_1_2'] = str(new)

def schimba_1_3(new):
    st.session_state['M_1_3'] = str(new)

def schimba_1_4(new):
    st.session_state['M_1_4'] = str(new)

def schimba_1_5(new):
    st.session_state['M_1_5'] = str(new)

def schimba_1_6(new):
    st.session_state['M_1_6'] = str(new)

def schimba_1_7(new):
    st.session_state['M_1_7'] = str(new)

def schimba_1_8(new):
    st.session_state['M_1_8'] = str(new)

def schimba_2_1(new):
    st.session_state['M_2_1'] = str(new)

def schimba_2_2(new):
    st.session_state['M_2_2'] = str(new)

def schimba_2_3(new):
    st.session_state['M_2_3'] = str(new)

def schimba_2_2_1(new):
    st.session_state['M_2_2_1'] = str(new)

def schimba_2_3_1(new):
    st.session_state['M_2_3_1'] = str(new)

def schimba_2_4(new):
    st.session_state['M_2_4'] = str(new)

def schimba_2_5(new):
    st.session_state['M_2_5'] = str(new)

def schimba_2_6(new):
    st.session_state['M_2_6'] = str(new)

def schimba_2_7_1(new):
    st.session_state['M_2_7_1'] = str(new)

def schimba_2_7_2(new):
    st.session_state['M_2_7_2'] = str(new)    

def schimba_M_3_1(new):
    st.session_state['M_3_1'] = str(new)

def schimba_M_3_2(new):
    st.session_state['M_3_2'] = str(new)
    st.session_state['test_curs']=True

def schimba_3_3_s(new):
    st.session_state['M_3_3_s'] = str(new) 
    st.session_state['test_aplicatie']=True

def schimba_3_3_l(new):
    st.session_state['M_3_3_l'] = str(new) 
    st.session_state['test_aplicatie']=True

def schimba_3_3_p(new):
    st.session_state['M_3_3_p'] = str(new) 
    st.session_state['test_aplicatie']=True

def schimba_M_3_4(new):
    st.session_state['M_3_4'] = str(new) 

def schimba_M_3_5(new):
    st.session_state['M_3_5'] = str(new)
    st.session_state['test_curs']=True

def schimba_3_6_s(new):
    st.session_state['M_3_6_s'] = str(new) 
    st.session_state['test_aplicatie']=True

def schimba_3_6_l(new):
    st.session_state['M_3_6_l'] = str(new) 
    st.session_state['test_aplicatie']=True

def schimba_3_6_p(new):
    st.session_state['M_3_6_p'] = str(new) 
    st.session_state['test_aplicatie']=True

def schimba_M_3_7_a(new):
    st.session_state['M_3_7_a'] = str(new)

def schimba_M_3_7_b(new):
    st.session_state['M_3_7_b'] = str(new)

def schimba_M_3_7_c(new):
    st.session_state['M_3_7_c'] = str(new) 

def schimba_M_3_7_d(new):
    st.session_state['M_3_7_d'] = str(new) 

def schimba_M_3_7_e(new):
    st.session_state['M_3_7_e'] = str(new)

def schimba_M_3_7_f(new):
    st.session_state['M_3_7_f'] = str(new)

def schimba_M_3_8(new):
    st.session_state['M_3_8'] = str(new)

def schimba_M_3_9(new):
    st.session_state['M_3_9'] = str(new)

def schimba_M_3_11(new):
    st.session_state['M_3_11'] = str(new)

def schimba_M_4_1(new):
    st.session_state['M_4_1'] = str(new)

def schimba_M_4_2(new):
    st.session_state['M_4_2'] = str(new)

def schimba_M_5_1(new):
    st.session_state['M_5_1'] = str(new)

def schimba_M_5_2(new):
    st.session_state['M_5_2'] = str(new)

def schimba_M_6_cp(new):
    st.session_state['M_6_cp'] = str(new)

def schimba_M_6_ct(new):
    st.session_state['M_6_ct'] = str(new)

def schimba_M_7_1(new):
    st.session_state['M_7_1'] = str(new)

def schimba_M_7_2(new):
    st.session_state['M_7_2'] = str(new)

def schimba_M_8_1_1(new):
    st.session_state['M_8_1_1'] = str(new)

def schimba_M_8_1_2(new):
    st.session_state['M_8_1_2'] = str(new)

def schimba_M_8_1_3(new):
    st.session_state['M_8_1_3'] = str(new)

def schimba_M_8_1_4(new):
    st.session_state['M_8_1_4'] = str(new)

def schimba_M_8_1_5(new):
    st.session_state['M_8_1_5'] = str(new)

def schimba_M_8_1_6(new):
    st.session_state['M_8_1_6'] = str(new)

def schimba_M_8_1_7(new):
    st.session_state['M_8_1_7'] = str(new)

def schimba_M_8_1_8(new):
    st.session_state['M_8_1_8'] = str(new)

def schimba_M_8_1_9(new):
    st.session_state['M_8_1_9'] = str(new)

def schimba_M_8_1_10(new):
    st.session_state['M_8_1_10'] = str(new)

def schimba_M_8_1_11(new):
    st.session_state['M_8_1_11'] = str(new)

def schimba_M_8_1_12(new):
    st.session_state['M_8_1_12'] = str(new)

def schimba_M_8_1_13(new):
    st.session_state['M_8_1_13'] = str(new)

def schimba_M_8_1_14(new):
    st.session_state['M_8_1_14'] = str(new)

def schimba_M_8_1_mp(new):
    st.session_state['M_8_1_mp'] = str(new)

def schimba_8_1_o(new):
    st.session_state['8_1_o'] = str(new)

def schimba_Biblio_c(new):
    st.session_state['Biblio_c'] = str(new)

def schimba_M_8_2_1(new):
    st.session_state['M_8_2_1'] = str(new)

def schimba_M_8_2_2(new):
    st.session_state['M_8_2_2'] = str(new)

def schimba_M_8_2_3(new):
    st.session_state['M_8_2_3'] = str(new)

def schimba_M_8_2_4(new):
    st.session_state['M_8_2_4'] = str(new)

def schimba_M_8_2_5(new):
    st.session_state['M_8_2_5'] = str(new)

def schimba_M_8_2_6(new):
    st.session_state['M_8_2_6'] = str(new)

def schimba_M_8_2_7(new):
    st.session_state['M_8_2_7'] = str(new)

def schimba_M_8_2_8(new):
    st.session_state['M_8_2_8'] = str(new)

def schimba_M_8_2_9(new):
    st.session_state['M_8_2_9'] = str(new)

def schimba_M_8_2_10(new):
    st.session_state['M_8_2_10'] = str(new)

def schimba_M_8_2_11(new):
    st.session_state['M_8_2_11'] = str(new)

def schimba_M_8_2_12(new):
    st.session_state['M_8_2_12'] = str(new)

def schimba_M_8_2_13(new):
    st.session_state['M_8_2_13'] = str(new)

def schimba_M_8_2_14(new):
    st.session_state['M_8_2_14'] = str(new)

def schimba_8_1_mp(new):
    st.session_state['8_1_mp'] = str(new)

def schimba_8_1_o(new):
    st.session_state['8_1_o'] = str(new)

def schimba_Biblio_a(new):
    st.session_state['Biblio_a'] = str(new)

def schimba_9(new):
    st.session_state['9'] = str(new)

def schimba_10_1_c(new):
    st.session_state['10_1_c'] = str(new)

def schimba_10_2_c(new):
    st.session_state['10_2_c'] = str(new)

def schimba_10_3_c(new):
    st.session_state['10_3_c'] = str(new)

def schimba_10_1_a(new):
    st.session_state['10_1_a'] = str(new)

def schimba_10_2_a(new):
    st.session_state['10_2_a'] = str(new)

def schimba_10_3_a(new):
    st.session_state['10_3_a'] = str(new)

def schimba_10_6(new):
    st.session_state['10_6'] = str(new)

def schimba_da_cu(new):
    st.session_state['da_cu'] = str(new)

def schimba_data_dep(new):
    st.session_state['data_dep'] = str(new)

def schimba_data_fac(new):
    st.session_state['data_fac'] = str(new)

def schimba_decan(new):
    st.session_state['decan'] = str(new)

def schimba_dir_dep(new):
    st.session_state['dir_dep'] = str(new)

def schimba_tip(new):
    st.session_state['tip'] = str(new)

def schimba_mail(new):
    st.session_state['mail'] = str(new)

def schimba_dep(new):
    st.session_state['dep'] = str(new)

def schimba_fac(new):
    st.session_state['fac'] = str(new)

def buton_t():
    st.session_state['ut'] = 1

def parcurge(aa,bb,cc,dd):
    output=''
    if len(doc_result.body[aa][bb][cc])!=0:
            for i in range (len(doc_result.body[aa][bb][cc])):
                if  len(doc_result.body[aa][bb][cc][0])==1:
                    break
                    output=doc_result.body[aa][bb][cc][0]
                else:    
                    output+=doc_result.body[aa][bb][cc][i]+'\n'
    return output[:len(output)-1]
st.set_page_config(page_title="Fisa disciplinei",layout="wide", initial_sidebar_state="auto")

if 'FormSubmitter:Fisa disciplinei-Treceti la alegerea specializarii' not in st.session_state:
    st.session_state["FormSubmitter:Fisa disciplinei-Treceti la alegerea specializarii"]=''
if 'M_1_7' not in st.session_state:
    st.session_state['M_1_7']='IF – învăţământ cu frecvenţă'
if 'M_1_1' not in st.session_state:
    st.session_state['M_1_1']='Universitatea Tehnica din Cluj Napoca'
keys_dash=['M_3_1' ,'M_3_2' ,'M_3_3_s','M_3_3_l','M_3_3_p','M_3_4','M_3_5','M_3_6_s','M_3_6_l','M_3_6_p']
for key in keys_dash:
    st.session_state.setdefault(key, '-')
if 'M_3_7_a' not in st.session_state:
    st.session_state['M_3_7_a']=0.0

if 'M_3_7_b' not in st.session_state:
    st.session_state['M_3_7_b']=0.0
 
if 'M_3_7_c' not in st.session_state:
    st.session_state['M_3_7_c']=0.0

if 'M_3_7_d' not in st.session_state:
    st.session_state['M_3_7_d']=0.0
 
if 'M_3_7_e' not in st.session_state:
    st.session_state['M_3_7_e']=0.0
 
if 'M_3_7_f' not in st.session_state:
    st.session_state['M_3_7_f']=0.0
if 'decan' not in st.session_state:
    st.session_state['decan']='prof.dr.ing Daniela Manea'

if 'fac' not in st.session_state:
    st.session_state['fac']='Constructii'

if 'test_curs' not in st.session_state:
    st.session_state['test_curs']=False
if 'test_aplicatie' not in st.session_state:
    st.session_state['test_aplicatie']=False
if 'ut' not in st.session_state:
    st.session_state['ut']=False

keys_none=['cap2','cap3','cap4','resetare' ,'file','M_8_1_o1','M_8_1_mp1','M_8_1_o','M_8_1_mp']
for key in keys_none:
    st.session_state.setdefault(key, None)
keys_space=['M_1_2','M_1_3','M_1_4','M_1_5','M_1_6','M_1_8','M_2_1','M_2_2','M_2_3','M_2_2_1','M_2_3_1','M_2_4','M_2_5','M_2_6','M_2_7_1','M_2_7_2',
	    'M_3_8','M_3_9','M_3_11','M_4_1','M_4_2','M_5_1','M_5_2', 'M_6_cp','M_6_ct','M_7_1','M_7_2','M_8_1_1','M_8_1_2','M_8_1_3','M_8_1_4','M_8_1_5',
	    'M_8_1_6', 'M_8_1_7','M_8_1_8', 'M_8_1_9','M_8_1_10','M_8_1_11','M_8_1_12','M_8_1_13','M_8_1_14','M_8_1_mp','8_1_o','Biblio_c','M_8_2_1',
	    'M_8_2_2','M_8_2_3','M_8_2_4','M_8_2_5','M_8_2_6','M_8_2_7','M_8_2_8','M_8_2_9','M_8_2_10','M_8_2_11','M_8_2_12','M_8_2_13','M_8_2_14',
	    '8_1_mp','8_1_o','Biblio_a','9','10_1_c','10_2_c','10_3_c','10_1_a','10_2_a','10_3_a','10_6','da_cu','data_dep','data_fac','dir_dep','tip','mail','dep']
for key in keys_space:
    st.session_state.setdefault(key, "")
ver={
    'examen':'E',
    'verificare':'V',
    'colocviu':'C'
}
domeniu={'Amenajari si constructii hidrotehnice - (ACH)':'Inginerie civila'
             ,'Cai Ferate, Drumuri si Poduri-(CFDP)':'Inginerie civila'
             ,'Constructii civile, industriale si agricole (CCIA-eng)':'Inginerie civila'
             ,'Constructii civile, industriale si agricole (CCIA-Baia_Mare)':'Inginerie civila'
             ,'Constructii civile, industriale si agricole (CCIA)':'Inginerie civila'
             ,'Inginerie Civila  - (CCIA,CFDP,ACH,IUDR)':'Inginerie civila'
             ,'Inginerie urbana si dezvoltare regionala (IUDR)':'Inginerie civila'
             ,'Masuratori terestre si cadastru (MTC)': 'Inginerie geodezica'
             ,'Inginerie si Management in Constructii (IMC)':'Inginerie si management'
             ,'Cladiri verzi (CV)':'Inginerie civila'
            ,'Constructii durabile din beton (CDB)':'Inginerie civila'
            ,'Ingineria infrastructurii transporturilor (IIT)':'Inginerie civila'
            ,'Ingineria tehnologiilor speciale in constructii (ITSC)':'Inginerie civila'
            ,'Inginerie geotehnica (IG)':'Inginerie civila'
            ,'Inginerie structurala (IS)':'Inginerie civila'
            ,'Proiectarea avansata a structurilor din lemn si metal (PASLM - Baia Mare)':'Inginerie civila'
            ,'Managementul proiectelor si evaluarea proprietatii (MPEP)':'Inginerie si management'
            ,'Inteligenta artificiala in ing.civila si manag.constructiilor (AICIV)':'Inginerie civila'
	    ,'Inteligenta artificiala in ing.civila si manag.constructiilor (AICIV-e)':'Inginerie civila'}
pres={'Amenajari si constructii hidrotehnice - (ACH)':'ACH'
             ,'Cai Ferate, Drumuri si Poduri-(CFDP)':'CFDP'
             ,'Constructii civile, industriale si agricole (CCIA-eng)':'CCIA-eng'
             ,'Constructii civile, industriale si agricole (CCIA-Baia_Mare)':'CCIA-Baia_Mare'
             ,'Constructii civile, industriale si agricole (CCIA)':'CCIA'
             ,'Inginerie Civila  - (CCIA,CFDP,ACH,IUDR)':'Inginerie civila'
             ,'Inginerie urbana si dezvoltare regionala (IUDR)':'IUDR'
             ,'Masuratori terestre si cadastru (MTC)': 'MTC'
             ,'Inginerie si Management in Constructii (IMC)':'IMC'
             ,'Cladiri verzi (CV)':'CV'
            ,'Constructii durabile din beton (CDB)':'CDB'
            ,'Ingineria infrastructurii transporturilor (IIT)':'IIT'
            ,'Ingineria tehnologiilor speciale in constructii (ITSC)':'ITSC'
            ,'Inginerie geotehnica (IG)':'IG'
            ,'Inginerie structurala (IS)':'IS'
            ,'Proiectarea avansata a structurilor din lemn si metal (PASLM - Baia Mare)':'PASLM'
            ,'Managementul proiectelor si evaluarea proprietatii (MPEP)':'MPEP'
            ,'Inteligenta artificiala in ing.civila si manag.constructiilor (AICIV)':'AICIV'
            ,'Inteligenta artificiala in ing.civila si manag.constructiilor (AICIV-e)': 'AICIV-e'
     }
specializari= {
  '':[],
  'Licenta':['',
             'Amenajari si constructii hidrotehnice - (ACH)'
             ,'Cai Ferate, Drumuri si Poduri-(CFDP)'
             ,'Constructii civile, industriale si agricole (CCIA-eng)'
             ,'Constructii civile, industriale si agricole (CCIA-Baia_Mare)'
             ,'Constructii civile, industriale si agricole (CCIA)'
             ,'Inginerie Civila  - (CCIA,CFDP,ACH,IUDR)'
             ,'Inginerie urbana si dezvoltare regionala (IUDR)'
             ,'Masuratori terestre si cadastru (MTC)'
             ,'Inginerie si Management in Constructii (IMC)'],
  'Master':['',
            'Cladiri verzi (CV)'
            ,'Constructii durabile din beton (CDB)'
            ,'Ingineria infrastructurii transporturilor (IIT)'
            ,'Ingineria tehnologiilor speciale in constructii (ITSC)'
            ,'Inginerie geotehnica (IG)'
            ,'Inginerie structurala (IS)'
            ,'Proiectarea avansata a structurilor din lemn si metal (PASLM - Baia Mare)'
            ,'Managementul proiectelor si evaluarea proprietatii (MPEP)'
	    ,'Inteligenta artificiala in ing.civila si manag.constructiilor (AICIV)'
	    ,'Inteligenta artificiala in ing.civila si manag.constructiilor (AICIV-e)'
            ]
 }
directori = {
  'Mecanica constructiilor':'conf.dr.ing. Anca-Gabriela POPA',
  'Constructii civile si management':'conf.dr.ing. Claudiu ACIU',
  'Structuri':'conf.dr.ing. Attila Puskas',
  'Masuratori terestre':'conf.dr.ing. Sanda NAS',
  'Cai ferate, drumuri si poduri':'conf.dr.ing. Mihai Liviu DRAGOMIR',
  'Matematica':'prof. dr. Vasile-Dorian Popa',
  'Fizica':'prof. dr. Petru Pascuta',
  'Limbi straine':'conf.dr. Ruxanda Literat'}
decan= {
  'Constructii':'prof.dr.ing Daniela Lucia Manea',
  'Mecanica':'',
  'Arhitectura':''
  }
departamentele= {
  'Mecanica constructiilor':'Structural Mechanics',
  'Constructii civile si management':'Buildings and Management',
  'Structuri':'Structures',
  'Masuratori terestre':'Land Measurements and Cadastre',
  'Cai ferate, drumuri si poduri':'Railways, Roads and Bridges ',
  'Matematica':'Mathematics',
  'Fizica':'Physics',
  'Limbi straine':'Languages'}
@st.cache_resource 
def load_ftp_file():
    # Establish FTP connection
    #ftp_server = ftplib.FTP("users.utcluj.ro", st.secrets['u'], st.secrets['p'])
    ftp_server = ftplib.FTP_TLS("users.utcluj.ro")
    ftp_server.login(user=st.secrets['u'], passwd=st.secrets['p'])
    ftp_server.prot_p()
    ftp_server.encoding = "utf-8"  # Force UTF-8 encoding
    ftp_server.cwd('./public_html')

    # Download CSV files
    csv_data = {}
    for filename in ["lista_cd.csv", "planinv_2025.csv","lista_cd.csv"]:
        with BytesIO() as file_data:
            ftp_server.retrbinary(f"RETR {filename}", file_data.write)
            file_data.seek(0)  # Reset file pointer to the start
            csv_data[filename] = pd.read_csv(file_data, encoding="ISO-8859-1")

    # Download DOCX templates
    docx_files = {}
    for filename in [
        "fisa_template_Mail_.docx", 
        "fisa_template_Mail_eng.docx", 
        "fisa_template_Mail_curs_.docx",
        "fisa_template_Mail_curs_eng.docx",
        "fisa_template_Mail_aplicatie_.docx",
        "fisa_template_Mail_aplicatie_eng.docx"
    ]:
        file_data = BytesIO()
        ftp_server.retrbinary(f"RETR {filename}", file_data.write)
        file_data.seek(0)  # Reset file pointer to the start
        docx_files[filename] = file_data

    # Close FTP connection
    ftp_server.quit()

    # Return downloaded files
    return (
        csv_data["lista_cd.csv"], 
        csv_data["planinv_2025.csv"], 
        docx_files["fisa_template_Mail_.docx"], 
        docx_files["fisa_template_Mail_eng.docx"], 
        docx_files["fisa_template_Mail_curs_.docx"], 
        docx_files["fisa_template_Mail_curs_eng.docx"], 
        docx_files["fisa_template_Mail_aplicatie_.docx"], 
        docx_files["fisa_template_Mail_aplicatie_eng.docx"],
        csv_data["lista_cd.csv"]
    )
# Use a session state flag to control cache invalidation
data,data1,_,_,_,_,_,_,data2=load_ftp_file()
if "refresh_data" not in st.session_state:
    st.session_state.refresh_data = False

if st.button("🔄 Refresh FTP Data (apasa doar daca nu s-a actualizat baza de date!!!)"):
    st.session_state.refresh_data = True
if st.session_state.refresh_data:
    load_ftp_file.clear()  # Clear the cache
    st.session_state.refresh_data = False          


st.session_state['file'] = st.file_uploader("Incarca o fisa a disciplinei in format *docx")
if not(st.session_state['ut']):
    if st.button('Nu am fisa de incarcat'):
        st.session_state['ut']=True
if st.session_state['file']!=None or st.session_state['ut']:
  if st.session_state['file']!=None: 
    doc_result = docx2python(st.session_state['file'])
   
  nume_tit = data['nume'].tolist()
 
  st.title("Fisa disciplinei - 2025/2026")
  st.write('{:%d-%b-%Y}'.format(date.today()))

  with st.form('Fisa disciplinei pentru Facultatea de constructii - 2025/2026'):
    st.header('Capitolul 1: Date despre program')
    #st.write(data1['ore'])
    try:
        dfr=doc_result.body[3][0][1][0]
    except:
        pass
    #st.write(dfr)
    #lista_f=doc_result.body[1][1][1]+['Constructii', 'Mecanica', 'Arhitectura']
    #lista_f=['Constructii', 'Mecanica', 'Arhitectura']
    #Slista_f=my_function(lista_f)
    add_select = 'Constructii'
    st.session_state['M_1_2']= 'Constructii'

    lista_ci=['Licenta', 'Master']
    lista_ci=my_function(lista_ci)
    add_selectbox_C = st.selectbox(
            'Ciclul de studii?',
            (lista_ci),key='M_1_5'
        )
        
    #lista_d=doc_result.body[1][2][1]+['Mecanica constructiilor', 'Constructii civile si management', 'Structuri', 'Masuratori terestre', 'Cai ferate, drumuri si poduri']
    lista_d=['Mecanica constructiilor', 'Constructii civile si management', 'Structuri', 'Masuratori terestre', 'Cai ferate, drumuri si poduri','Matematica','Fizica','Limbi straine']
    lista_dd=lista_d
    lista_d=my_function(lista_d)
    try:
     add_selectbox_DP = st.selectbox('Departamentul?',lista_d,index=find_closest_match_index(doc_result.body[1][2][1][0], lista_d))
    except:
     add_selectbox_DP = st.selectbox('Departamentul?',lista_d)
	    
    #add_selectbox_dom = st.selectbox(
    #        'Domeniul de studii?',
    #        ('Inginerie civila', 'Inginerie si management', 'Inginerie geodezica'),key='M_1_4'
    #    )
    #lista_ci=doc_result.body[1][4][1]+['Licenta', 'Master']

    submitted = st.form_submit_button("Treceti la alegerea specializarii")
    if submitted:
      st.session_state['M_1_3']=add_selectbox_DP
      #st.write()
      schimba_1_1("Universitatea Tehnica din Cluj Napoca")
      ##st.write(departamentele[add_selectbox_SP])
      #schimba_1_2(add_selectbox)
      #schimba_1_3(add_selectbox_DP)
      #schimba_1_4(add_selectbox_dom)
      #schimba_1_5(add_selectbox_SP)
      #schimba_1_6(add_selectbox_C)
      #schimba_dep(add_selectbox_DP)
      schimba_1_7('IF – învăţământ cu frecvenţă')
      st.session_state['cap2']='1'
     
  if st.session_state['cap2']!=None:
    with st.form('Alege specializarea:'):
        try:
         add_selectbox_SP = st.selectbox(
          'Programul de studii?',my_function(specializari[st.session_state['M_1_5']]),key='M_1_6',index=find_closest_match_index(doc_result.body[1][5][1][0], specializari[st.session_state['M_1_5']]))
        except:
         add_selectbox_SP = st.selectbox('Programul de studii?',my_function(specializari[st.session_state['M_1_5']]),key='M_1_6')
        submitted = st.form_submit_button("Treceti la capitolul 2-3")
        if submitted:
            st.session_state['M_1_4']=domeniu[add_selectbox_SP]
            st.session_state['cap2']='2'
  if st.session_state['cap2']=='2':
    with st.form('capitolul 2 si 3'):        
        st.header('Capitolul 2 si 3:Date despre disciplină si Timp total estimat')
        
        data1['nume_disciplina'] = data1['nume_disciplina'].apply(strip_last)
        data1['specializare'] = data1['specializare'].apply(strip_last)
        #st.write(st.session_state['M_1_6'])
        nume_di = data1['nume_disciplina'].loc[(data1['specializare']==st.session_state['M_1_6'])].drop_duplicates().tolist()
        data1['ore_s']=14*data1['numarore'].astype(int)
        
        try:
         add_selectbox_D = st.selectbox('Disciplina?',
                nume_di,key='M_2_1',index=find_closest_match_index(doc_result.body[3][0][3][0], nume_di),
                help='Toate datele asociate disciplinei vor fi inserate in fisa disciplinei automat din planul de invatamant')
        except:
         add_selectbox_D = st.selectbox(
                'Disciplina?',
                nume_di,key='M_2_1',
                help='Toate datele asociate disciplinei vor fi inserate in fisa disciplinei automat din planul de invatamant')
	#nume_tit1=doc_result.body[3][2][1]+nume_tit
        nume_tit1=nume_tit
        #nume_tit=doc_result.body[3][1][1]+nume_tit
        add_selectbox_TC = st.multiselect('Titulari curs?',
					  nume_tit, 
					  #placeholder="De exemplu"+find_closest_match(doc_result.body[3][2][1], nume_tit),
					  help='Pot fi selectati mai multi titulari de curs.')
        add_selectbox_TA = st.multiselect('Titular aplicatii?',
					  nume_tit,  
					  #placeholder="De exemplu"+find_closest_match(doc_result.body[3][1][1], nume_tit)
					 )
        submitted = st.form_submit_button("Treci la subcapitolul 3.7")
        if submitted:
                
                data1['nume_disciplina'] = data1['nume_disciplina'].apply(strip_last)

                st.write("A fost selectata disciplina "+st.session_state['M_2_1'])
                s=''
                ss=''
                for d in add_selectbox_TC:
                    s+=d.title()+'-'+str(data['mail'].loc[(data['nume']==d)].values[0])+'\n'
                    ss+=d.title()+'\n'
                s=s[:len(s)-1]
                ss=ss[:len(ss)-1]
                st.session_state['M_2_2']=s
                st.session_state['M_2_2_1']=ss
                s=''
                ss=''
                for d in add_selectbox_TA:
                    s+=d.title()+'-'+str(data['mail'].loc[(data['nume']==d)].values[0])+'\n'
                    ss+=d.title()+'\n'
                s=s[:len(s)-1]
                ss=ss[:len(ss)-1]
                st.session_state['M_2_3']=s
                st.session_state['M_2_3_1']=ss
                st.session_state['M_1_8']=str(data1['nrcrt'].loc[(data1['specializare']==st.session_state['M_1_6']) & (data1['nume_disciplina']==st.session_state['M_2_1'])].values[0])

                #st.write('Numele disciplinei este:',st.session_state['M_2_1'])
                #st.write(st.session_state['M_2_2'])
                #st.write('Codul disciplinei este:',st.session_state['M_1_8'])

                #schimba_2_3(add_selectbox_TA)
                #st.write(st.session_state['M_2_3_1'])
                
		
                #st.session_state['M_2_4']=str(data1['an'].loc[(data1['specializare']==st.session_state['M_1_6']) & (data1['nume_disciplina']==st.session_state['M_2_1'])].values[0])
                #st.write("Anul in care e studiata disciplina aleasa: ",st.session_state['M_2_4'])
                st.session_state['M_2_5']=str(data1['semestru'].loc[(data1['specializare']==st.session_state['M_1_6']) & (data1['nume_disciplina']==st.session_state['M_2_1'])].values[0])
                indices = data1.loc[(data1['specializare'] == st.session_state['M_1_6']) & (data1['nume_disciplina'] == st.session_state['M_2_1']),'semestru'].index
                
                st.session_state['M_2_4']=str(data1.iloc[indices.tolist()[0], 0])
		
		#st.write("Semestrul in care e studiata disciplina aleasa: ",st.session_state['M_2_5'])
                st.session_state['M_2_6']=ver[str(data1['examin'].loc[(data1['specializare']==st.session_state['M_1_6']) & (data1['nume_disciplina']==st.session_state['M_2_1'])].values[0]).strip()]
                #st.write("Tipul de examinare pentru disciplina aleasa: ",st.session_state['M_2_6'])

                st.session_state['M_2_7_1']=str(data1['numecat'].loc[(data1['specializare']==st.session_state['M_1_6']) & (data1['nume_disciplina']==st.session_state['M_2_1'])].values[0])
                #st.write(st.session_state['M_2_7_1'])
                st.session_state['M_2_7_2']=str(data1['obligativ'].loc[(data1['specializare']==st.session_state['M_1_6']) & (data1['nume_disciplina']==st.session_state['M_2_1'])].values[0])
                #st.write(st.session_state['M_2_7_2'])
                #data1['curs'].loc[(data1['nume_disciplina']==add_selectbox_D)&(data1['curs']!='curs'),0]
                #st.write(add_selectbox_D)
                #st.write('Lungimea string-ului este:',len(str(data1['curs'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0])))
                try:
                    st.session_state['tip']=str(data1['curs'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0]).lower()
                except:
                    pass
                schimba_mail('')

                try:
                    st.session_state['M_3_1']=str(int(data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']=='CURS      ')].values[0]+data1['numarore'].loc[(data1['nume_disciplina']==add_selectbox_D)&(data1['curs']!='CURS      ')].values[0]))
                    st.session_state['M_3_4']=str(int(14*(data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']=='CURS      ')].values[0]+data1['numarore'].loc[(data1['nume_disciplina']==add_selectbox_D)&(data1['curs']!='CURS      ')].values[0])))
			
                    st.session_state['test_curs']=True
                    st.session_state['test_aplicatie']=True
                except:
                    try: 
                       st.session_state['M_3_1']=str(int(data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0]))
                       st.session_state['M_3_4']=str(int(14*(data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0])))

                    except:
                        try:
                            #schimba_M_3_1(data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']=='CURS      ')].values[0])
                            #schimba_M_3_4(14*(data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']=='CURS      ')].values[0]))
                            st.session_state['M_3_1']=str(int(data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']=='CURS      ')].values[0]))
                            st.session_state['M_3_4']=str(int(14*(data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']=='CURS      ')].values[0])))

                            st.session_state['test_curs']=True
                        except:
                            pass
                try:
                    st.session_state['M_3_2']=str(int(data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']=='CURS      ')].values[0]))
                    st.session_state['M_3_5']=str(int(14*data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']=='CURS      ')].values[0]))

                except:
                    pass
                try:
                    if data1['curs'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0]=='SEMINAR   ':
                        #schimba_3_3_s(data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0])
                        #schimba_3_6_s(14*data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0])
                        st.session_state['M_3_3_s']=str(int(data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0]))
                        st.session_state['M_3_6_s']=str(int(14*data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0]))

                        # st.write('Am gasit seminar')
                        st.session_state['test_aplicatie']=True
                except:
                    pass
                try:
                    if data1['curs'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0]=='LABORATOR ':
                        #schimba_3_3_l(data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0])
                        #schimba_3_6_l(14*data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0])
                        st.session_state['M_3_3_l']=str(int(data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0]))
                        st.session_state['M_3_6_l']=str(int(14*data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0]))
                         #st.write('Am gasit lucrari')
                        st.session_state['test_aplicatie']=True
                except:
                    pass
                try:
                    if data1['curs'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0]=='PROIECT   ':
                        #schimba_3_3_p(data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0])
                        #schimba_3_6_p(14*data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0])
                        st.session_state['M_3_3_p']=str(data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0])
                        st.session_state['M_3_6_p']=str(14*data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0])
                        st.session_state['test_aplicatie']=True
                except:
                    pass
                try:
                    st.session_state['M_3_8']=str(data1['orestud'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])].values[0])
                except:
                    pass
                try:
                    st.session_state['M_3_9']=str(25*data1['credite'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])].values[0])
                except:
                    pass
                try:
                    st.session_state['M_3_11']=str(data1['credite'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])].values[0])
                except:
                    pass
                st.session_state['cap3']='1'
  if st.session_state['cap3']!=None:
    st.write('Distribuția fondului de timp (ore pe semestru)')
    #st.session_state['M_3_8']=str(data1['orestud'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1']) & (data1['curs']=='CURS      ')].values[0])
    tosi=data1['orestud'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1']) ].values[0]

    #st.write('Total ore studiu individual ', tosi)

    slide_37a=0
    slide_37b=0
    slide_37c=0
    slide_37d=0
    slide_37e=0
    slide_37f=0
    st.write('Distribuția fondului de timp:')
    try:
     slide_37a=st.slider(
      '(a) Studiul după manual, suport de curs, bibliografie şi notițe',
      min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
      value=int(doc_result.body[5][3][-1][0]),
      help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
    except:
     slide_37a=st.slider(
      '(a) Studiul după manual, suport de curs, bibliografie şi notițe',
      min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
      
      help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')	    
    try:  
     slide_37b=st.slider(
      '(b) Documentare suplimentară în bibliotecă, pe platforme electronice de specialitate şi pe teren',
      min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
      value=int(doc_result.body[5][4][-1][0]),
      help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
    except:  
     slide_37b=st.slider(
      '(b) Documentare suplimentară în bibliotecă, pe platforme electronice de specialitate şi pe teren',
      min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
      
      help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
    try:
     slide_37c=st.slider(
      'c) Pregătire seminarii / laboratoare, teme, referate, portofolii şi eseuri',
      min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
      value=int(doc_result.body[5][5][-1][0]),
      help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
    except:
     slide_37c=st.slider(
      'c) Pregătire seminarii / laboratoare, teme, referate, portofolii şi eseuri',
      min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
      help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
    try:
     slide_37d=st.slider(
      '(d) Tutoriat',
      min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
      value=int(doc_result.body[5][6][-1][0]),
      help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
    except:
     slide_37d=st.slider(
      '(d) Tutoriat',
      min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
      help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
    sd=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f))
    try:
     slide_37e=st.slider(
      'e) Examinări',
      min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
      value=int(doc_result.body[5][7][-1][0]),
      help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
    except:
     slide_37e=st.slider(
      'e) Examinări',
      min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
      
      help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
    sd=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f))
    if not(sd<=0):
        slide_37f=st.slider(
          '(f) Alte activități:',
           max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
      
          value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
          help='Completati de la a spre f. Suma orelor de studiu individual este cea din planurile de invatamant')
    else:
            st.write('(f) Alte activități: 0 ore')
            slide_37f=0
            slide_37e+=-sd
    a=st.button('Treci la capitolul 4')
    if a:
      st.write('Capitolul 4')
      schimba_M_3_7_a(slide_37a)
      schimba_M_3_7_b(slide_37b)
      schimba_M_3_7_c(slide_37c)
      schimba_M_3_7_d(slide_37d)
      schimba_M_3_7_e(slide_37e)
      schimba_M_3_7_f(slide_37f)
      
      st.session_state['cap4']='1'
	    
  if st.session_state['cap4']!=None:
    with st.form('capitolul 4'):
      try:
        c41=''
        #st.write(doc_result.body[11][0][1])
        for i in range (len(doc_result.body[7][0][1])):
            if len(doc_result.body[7][0][1][0])==1:
                c41=doc_result.body[7][0][1][0]
                break
                
            c41+=doc_result.body[7][0][1][i]+'\n'
        #c41=parcurge(7,0,1,0)
        #c41=doc_result.body[7][0][1][0]
        c41=c41[:len(c41)-1]
      except:
        c41=''      
      text_4_1=st.text_area('4.1 Preconditii din curriculum',value=c41,key='M_4_1',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
      try:
        #c42=doc_result.body[7][1][1][0]
        c42=parcurge(7,1,1,0)
      except:
        c42=''      
      text_4_2=st.text_area('4.2 Preconditii de competente',value=c42,key='M_4_2',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
      try:
        #c51=doc_result.body[9][0][1][0]
        c51=parcurge(9,0,1,0)
      except:
        c51=''       
      text_5_1=st.text_area('5.1 Conditii de desfasurare a cursului',value=c51,key='M_5_1',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
      try:
        c52=parcurge(9,1,1,0)
        #c52=doc_result.body[9][1][1][0]
      except:
        c52=''      
      text_5_2=st.text_area('5.2 Conditii de desfasurare a aplicatiilor',value=c52,key='M_5_2',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
      try:
        text_c=''
        #st.write(doc_result.body[11][0][1])
        for i in range (len(doc_result.body[11][0][1])):
          text_c+=doc_result.body[11][0][1][i]+'\n'
        text_c=text_c[:len(text_c)-1]
      except:
        text_c=''
      try:
        text_p=''
        #st.write(doc_result.body[11][0][1])
        for i in range (len(doc_result.body[11][1][1])):
          text_p+=doc_result.body[11][1][1][i]+'\n'
        text_p=text_p[:len(text_p)-1]
      except:
        text_p=''
      text_6_1=st.text_area('6.a Competente profesionale acumulate',value=text_c,key='M_6_cp',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
      
      text_6_2=st.text_area('6.b Competente transversale',value=text_p,key='M_6_ct',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
      try:
        c71=parcurge(13,0,1,0)
      except:
        c71=''      
      text_7_1=st.text_area('7.1 Obiectivul general al disciplinei',value=c71,key='M_7_1',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
      try:
        c73=parcurge(13,1,1,0)
      except:
        c73=''
      text_7_2=st.text_area('7.2 Obiectivele specifice',value=c73,key='M_7_2',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
      #st.write(doc_result.body[15])
      if st.session_state['test_curs']:
          try:
            tc1=st.text_area('Curs 1',value=parcurge(15,1,0,0),key='M_8_1_1',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          except:
            tc1=st.text_area('Curs 1',value='',key='M_8_1_1',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            tc2=st.text_area('Curs 2',value=parcurge(15,2,0,0),key='M_8_1_2',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          except:
            tc2=st.text_area('Curs 2',value='',key='M_8_1_2',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            c3=parcurge(15,3,0,0)
          except:
            c3=''
          st.text_area('Curs 3',value=c3,key='M_8_1_3',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          
          try:
            c4=parcurge(15,4,0,0)
          except:
            c4=''
          st.text_area('Curs 4',value=c4,key='M_8_1_4',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            c5=parcurge(15,5,0,0)
          except:
            c5=''
          st.text_area('Curs 5',value=c5,key='M_8_1_5',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            c6=parcurge(15,6,0,0)
          except:
            c6=''
          st.text_area('Curs 6',value=c6,key='M_8_1_6',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            c7=parcurge(15,7,0,0)
          except:
            c7=''
          st.text_area('Curs 7',value=c7,key='M_8_1_7',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            c8=parcurge(15,8,0,0)
          except:
            c8=''
          st.text_area('Curs 8',value=c8,key='M_8_1_8',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            c9=parcurge(15,9,0,0)
          except:
            c9=''
          st.text_area('Curs 9',value=c9,key='M_8_1_9',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            c10=parcurge(15,10,0,0)
          except:
            c10=''
          st.text_area('Curs 10',value=c10,key='M_8_1_10',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            c11=parcurge(15,11,0,0)
          except:
            c11=''
          st.text_area('Curs 11',value=c11,key='M_8_1_11',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            c12=parcurge(15,12,0,0)
          except:
            c12=''
          st.text_area('Curs 12',value=c12,key='M_8_1_12',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            c13=parcurge(15,13,0,0)
          except:
            c13=''
          st.text_area('Curs 13',value=c13,key='M_8_1_13',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            c14=parcurge(15,14,0,0)
          except:
            c14=''
          st.text_area('Curs 14',value=c14,key='M_8_1_14',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            a111=doc_result.body[15][1][1][0]
            #st.write(a111)
            st.text_area('Metode de predare-Curs',value=a111,key='M_8_1_mp',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          except:
            a111=''
            st.text_area('Metode de predare-Curs',value=a111,key='M_8_1_mp',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            b111=doc_result.body[15][1][2][0]
            #parcurge(15,1,2,0)
          except:
            b111=''
          st.text_area('Observatii-Curs',value=b111,key='M_8_1_o',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          text_bc=''
          try:
              for i in range(len(doc_result.body[15][15][0])):
                text_bc+=doc_result.body[15][15][0][i]+'\n'
              text_bc=text_bc[:len(text_bc)-1]
          except:
              pass
          st.text_area('Biliografie-Curs',value=text_bc,key='Biblio_c',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
      if st.session_state['test_aplicatie']:
          try:
            if st.session_state['M_3_2'].isdigit():
              p1=parcurge(15,17,0,0)
            else:
              p1=parcurge(15,1,0,0)          
          except:
            p1=''
          st.text_area('Aplicatia 1',value=p1,key='M_8_2_1',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            if st.session_state['M_3_2'].isdigit():
              p2=parcurge(15,18,0,0)
            else:
              p2=parcurge(15,2,0,0)  
          except:
            p2=''
          st.text_area('Aplicatia 2',value=p2,key='M_8_2_2',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            if st.session_state['M_3_2'].isdigit():
              p3=parcurge(15,19,0,0)
            else:
              p3=parcurge(15,3,0,0) 
          except:
            p3=''
          st.text_area('Aplicatia 3',value=p3,key='M_8_2_3',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            if st.session_state['M_3_2'].isdigit():
              p4=parcurge(15,20,0,0)
            else:
              p4=parcurge(15,4,0,0) 
          except:
            p4=''
          st.text_area('Aplicatia 4',value=p4,key='M_8_2_4',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            if st.session_state['M_3_2'].isdigit():
              p5=parcurge(15,21,0,0)
            else:
              p5=parcurge(15,5,0,0) 
          except:
            p5=''
          st.text_area('Aplicatia 5',value=p5,key='M_8_2_5',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            if st.session_state['M_3_2'].isdigit():
              p6=parcurge(15,22,0,0)
            else:
              p6=parcurge(15,6,0,0) 
          except:
            p6=''
          st.text_area('Aplicatia 6',value=p6,key='M_8_2_6',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            if st.session_state['M_3_2'].isdigit():
              p7=parcurge(15,23,0,0)
            else:
              p7=parcurge(15,7,0,0) 
          except:
            p7=''
          st.text_area('Aplicatia 7',value=p7,key='M_8_2_7',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            if st.session_state['M_3_2'].isdigit():
              p8=parcurge(15,24,0,0)
            else:
              p8=parcurge(15,8,0,0) 
          except:
            p8=''
          st.text_area('Aplicatia 8',value=p8,key='M_8_2_8',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            if st.session_state['M_3_2'].isdigit():
              p9=parcurge(15,25,0,0)
            else:
              p9=parcurge(15,9,0,0) 
          except:
            p9=''
          st.text_area('Aplicatia 9',value=p9,key='M_8_2_9',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            if st.session_state['M_3_2'].isdigit():
              p10=parcurge(15,26,0,0)
            else:
              p10=parcurge(15,10,0,0) 
          except:
            p10=''
          st.text_area('Aplicatia 10',value=p10,key='M_8_2_10',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            if st.session_state['M_3_2'].isdigit():
              p11=parcurge(15,27,0,0)
            else:
              p11=parcurge(15,11,0,0) 
          except:
            p11=''
          st.text_area('Aplicatia 11',value=p11,key='M_8_2_11',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            if st.session_state['M_3_2'].isdigit():
              p12=parcurge(15,28,0,0)
            else:
              p12=parcurge(15,12,0,0) 
          except:
            p12=''
          st.text_area('Aplicatia 12',value=p12,key='M_8_2_12',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            if st.session_state['M_3_2'].isdigit():
              p13=parcurge(15,29,0,0)
            else:
              p13=parcurge(15,13,0,0) 
          except:
            p13=''
          st.text_area('Aplicatia 13',value=p13,key='M_8_2_13',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            if st.session_state['M_3_2'].isdigit():
              p14=parcurge(15,30,0,0)
            else:
              p14=parcurge(15,14,0,0) 
          except:
            p14=''
          st.text_area('Aplicatia 14',value=p14,key='M_8_2_14',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            aa1=parcurge(15,1,1,0)
          except:
            aa1=''  
          st.text_area('Metode de predare-aplicatii',value=aa1, key='M_8_1_mp1',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          try:
            bb1=parcurge(15,1,2,0)
          except:
            bb1=''
          try:
            text_ba=parcurge(15,31,0,0)
          except:
            text_ba=''
            
          st.text_area('Observatii-aplicatii',value=bb1, key='M_8_1_o1',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
          st.text_area('Bibliografie-Aplicatii',value=text_ba,key='Biblio_a',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")

      try:
        c9=parcurge(17,0,0,0)
      except:
        c9=''
      text_9=st.text_area('9. Coroborarea conținuturilor disciplinei cu așteptările reprezentanţilor comunităţii epistemice, a.p. s.a.m.d. din domeniul aferent programului',value=c9,key='M_9')
      try:
        c101=parcurge(19,1,1,0)
      except:
        c101=''
      text_10_1=st.text_area('10.1 Criterii de evaluare curs',value=c101,key='M_10_1_c',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
      try:
        c102=parcurge(19,1,2,0)
      except:
        c102=''
      text_10_2=st.text_area('10.2 Metode de evaluare curs',value=c102,key='M_10_2_c',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
      try:
        c1021=parcurge(19,2,1,0)
      except:
        c1021=''
      text_10_2_1=st.text_area('10.1.1 Criterii de evaluare aplicatii',value=c1021,key='M_10_1_a',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
      try:
        c1022=parcurge(19,2,2,0)
      except:
        c1022=''      
      text_10_2_2=st.text_area('10.2 Metode de evaluare aplicatii',value=c1022,key='M_10_2_a',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
      try:
        c103=parcurge(19,1,3,0)
      except:
        c103=''      
      text_10_3=st.text_area('10.3 Ponderea notei la curs in nota finala',value=c103,key='M_10_3_c',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
      try:
        c1032=parcurge(19,2,3,0)
      except:
        c1032=''      
      text_10_3_2=st.text_area('10.3.1 Ponderea notei la aplicatii in nota finala',value=c1032,key='M_10_3_a',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
      try:
        cst=parcurge(19,4,0,0)
      except:
        cst=''      
      text_st=st.text_area('10.6 Standard minim de performanţă',value=cst,key='M_10_6',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
      #d_com=st.date_input("Data completarii",date.today())
      d_com=date.today()
      d_dep=st.date_input("Data avizari in departament",date.today())
      d_fac=st.date_input("Data avizari in consiliul facultatii",date(2025, 6, 25))
      #d_dep='04.09.2022'
      #d_fac='21.09.2022'
      submitted= st.form_submit_button("finalizeaza")
      if submitted:
        schimba_decan(decan[add_select])
        st.session_state['dir_dep']=directori[st.session_state['M_1_3']]
        st.session_state['data_dep']=str(d_dep)
        schimba_fac(add_select)
        st.session_state['data_fac']=str(d_fac)
        st.session_state['d_com']=str(d_com)
        del st.session_state["FormSubmitter:capitolul 2 si 3-Treci la subcapitolul 3.7"]
        del st.session_state["FormSubmitter:Alege specializarea:-Treceti la capitolul 2-3"]
        del st.session_state["resetare"]
        del st.session_state["FormSubmitter:Fisa disciplinei-Treceti la alegerea specializarii"]
        del st.session_state["FormSubmitter:capitolul 4-finalizeaza"]

        #st.write(st.session_state)
        if (st.session_state['test_curs'])&(st.session_state['test_aplicatie']):
            if st.session_state['M_1_6']!='Constructii civile, industriale si agricole (CCIA-eng)' or st.session_state['M_1_6']!='Inteligenta artificiala in ing.civila si manag.constructiilor (AICIV-e) ':
                
                _,_,template,_,_,_,_,_,_=load_ftp_file()
            else:
                _,_,_,template,_,_,_,_,_=load_ftp_file()
        
        if (st.session_state['test_curs'])&(not(st.session_state['test_aplicatie'])):
            if st.session_state['M_1_6']!='Constructii civile, industriale si agricole (CCIA-eng)' or st.session_state['M_1_6']!='Inteligenta artificiala in ing.civila si manag.constructiilor (AICIV-e) ':
               
                _,_,_,_,template,_,_,_,_=load_ftp_file()
                st.session_state['M_1_3']=departamentele[st.session_state['M_1_3']]
            else:
                _,_,_,_,_,template,_,_,_=load_ftp_file()
                st.session_state['M_1_3']=departamentele[st.session_state['M_1_3']]
        if (not(st.session_state['test_curs']))&(st.session_state['test_aplicatie']):
            if st.session_state['M_1_6']!='Constructii civile, industriale si agricole (CCIA-eng)':
                _,_,_,_,_,_,template,_,_=load_ftp_file()
            else:
                _,_,_,_,_,_,_,template,_=load_ftp_file()
                st.session_state['M_1_3']=departamentele[st.session_state['M_1_3']]


    
        document = MailMerge(template)
        #st.write(document.get_merge_fields())
        document.merge(da_cu=st.session_state['d_com'])
        keys_to_merge=['denumirefisa','dataintocmire','M_1_1','M_1_2','M_1_3','M_1_4','M_1_5','M_1_6','M_1_8',
		       'M_2_1','M_2_2','M_2_3','M_2_3_1','M_2_4','M_2_5','M_2_6','M_2_2_1','M_2_7_1','M_2_7_2',
		       'M_3_1','M_3_2','M_3_3_l','M_3_3_s','M_3_3_p','M_3_4','M_3_5','M_3_6_s','M_3_6_l','M_3_6_p','M_3_7_a','M_3_7_b','M_3_7_c','M_3_7_d','M_3_7_e','M_3_7_f',
		       'M_3_8','M_3_9','M_3_11',
		       'M_4_1','M_4_2',
		       'M_5_1','M_5_2',
		       'M_6_cp','M_6_ct',
		       'M_7_1','M_7_2',
                       'M_8_1_1','M_8_1_2','M_8_1_3','M_8_1_4','M_8_1_5','M_8_1_6','M_8_1_7','M_8_1_8','M_8_1_9','M_8_1_11','M_8_1_12','M_8_1_13','M_8_1_14','Biblio_c',
                       'M_8_2_1','M_8_2_2','M_8_2_3','M_8_2_4','M_8_2_5','M_8_2_6','M_8_2_7','M_8_2_8','M_8_2_9','M_8_2_10','M_8_2_12','M_8_2_13','M_8_2_14','Biblio_a',
		       'M_9',
		       'M_10_1_a','M_10_1_c','M_10_2_c','M_10_3_a','M_10_3_c','M_10_6','M_10_2_a','M_8_1_10',
		       'M_8_1_o1','M_8_1_mp','M_8_1_mp1','M_8_1_o',
		       'dep','da_cu','data_fac','data_dep','tip','dir_dep','decan','fac']
        data_ftp=pd.DataFrame(columns=keys_to_merge)
	
        for key in keys_to_merge:
          if key in st.session_state:
            document.merge(**{key: st.session_state[key]})
        file_name=st.session_state['M_1_8']+'_FD_an'+st.session_state['M_2_4']+'_s'+st.session_state['M_2_5']+'_'+pres[st.session_state['M_1_6']]+'_'+st.session_state['M_2_1']+'_26-27.docx'
        file_name_bak=st.session_state['M_1_8']+'_FD_an'+st.session_state['M_2_4']+'_s'+st.session_state['M_2_5']+'_'+pres[st.session_state['M_1_6']]+'_'+st.session_state['M_2_1']+'_26-27_'+str(int(time.time()))+'.pkl'		  
        remote_filename=st.session_state['M_1_8']+'_FD_an'+st.session_state['M_2_4']+'_s'+st.session_state['M_2_5']+'_'+pres[st.session_state['M_1_6']]+'_'+st.session_state['M_2_1']+'_26-27.pkl'
        remote_filename_csv=st.session_state['M_1_8']+'_FD_an'+st.session_state['M_2_4']+'_s'+st.session_state['M_2_5']+'_'+pres[st.session_state['M_1_6']]+'_'+st.session_state['M_2_1']+'_26-27.csv'
        current_datetime = datetime.now()    
        document.write(file_name)
	
        st.markdown(get_binary_file_downloader_html(file_name, 'Word document'), unsafe_allow_html=True)
        st.session_state['denumirefisa']=file_name
        st.session_state['dataintocmire']=str(current_datetime)
        #os.startfile(file_name)
        def fix_encoding(text):
            return text.encode('latin1').decode('utf-8')
        df = data2
        required_keys = data2.columns  # Use `data2` column names as expected keys
        for key in keys_to_merge:
          if key in st.session_state:
           data_ftp[key]=st.session_state[key]
	# Define the new row based on session state
	#Add the new row to `df` using pd.concat
        new_row_df = pd.DataFrame([{key: st.session_state.get(key, '') for key in st.session_state.keys()}])
        new_row_df = new_row_df.fillna('')  # Replace with appropriate default values if needed
        for col in data_ftp.columns:
          if data_ftp[col].dtype == 'object':  # Convert object columns to strings
            data_ftp[col] = data_ftp[col].astype(str)
          elif data_ftp[col].dtype.name == 'category':  # Convert categories to strings
            data_ftp[col] = data_ftp[col].astype(str)
        #st.dataframe(new_row_df)
        dict_from_df = new_row_df.to_dict(orient='list')
        #df = pd.concat([data2, new_row_df], ignore_index=True)
        csv_buffer = BytesIO()
        data_ftp.to_csv(csv_buffer, index=False)
        csv_buffer.seek(0)  # Reset buffer pointer to the beginning
        pickle_buffer = BytesIO()
        pickle_buffer_bak = BytesIO()
        #!!!!!!!!!!!!
        pickle.dump({key: str(st.session_state.get(key, '')) for key in st.session_state.keys()}, pickle_buffer)
        pickle_buffer.seek(0) 
        pickle.dump({key: str(st.session_state.get(key, '')) for key in st.session_state.keys()}, pickle_buffer_back)
        pickle_buffer_back.seek(0) 

        ftp_server1 = ftplib.FTP_TLS("users.utcluj.ro")
        ftp_server1.login(user=st.secrets['u'], passwd=st.secrets['p'])
        ftp_server1.prot_p()
        ftp_server1.encoding = "utf-8"
        ftp_server1.cwd('./public_html/Fise/2026')
        ftp_server1.storbinary(f'STOR {remote_filename}', pickle_buffer)  
		  # Send the file
	#ftp_server1.storbinary(f'STOR {file_name}', docx_buffer)
        ftp_server1.cwd('..')
        ftp_server1.cwd('./public_html/Fise/2026_bak')
        ftp_server1.storbinary(f'STOR {remote_filename}', pickle_buffer_bak)     
        ftp_server1.quit()
        docx_buff=BytesIO()
        document.write(docx_buff)
        docx_buff.seek(0)
        ftp_server1 = ftplib.FTP_TLS("users.utcluj.ro")
        ftp_server1.login(user=st.secrets['u'], passwd=st.secrets['p'])
        ftp_server1.prot_p()
        ftp_server1.encoding = "utf-8"
        ftp_server1.cwd('./public_html/Fise/2026')
        ftp_server1.storbinary(f'STOR {file_name}', docx_buff)
        ftp_server1.quit()

        #sub1=st.form_submit_button("Incarca alt fisier docx")
        #if sub1:
        # st.write("Acceseaza linkul de mai jos pentru a incarca o alta fisa in format docx")
        # redirect_url1 = "https://fisaconstructiiutcn.streamlit.app/"
        # st.markdown(f"[Incaraca alt fisier docx]({redirect_url1})")
       #  st.stop()  # Only if needed, or use JS for redirect
       # sub2=st.form_submit_button("Cauta fisa noua in baza de date")
       # if sub2:
        st.write("Acceseaza linkul de mai jos pentru a cauta din nou in baza de date o fisa")
        redirect_url2 = "https://fisaconstructii-utcn.streamlit.app/"
        st.markdown(f"[Cauta in baza de date]({redirect_url2})")
        st.stop()  # Only if needed, or use JS for redirect# Convert the updated DataFrame to CSV format
        #data_baza = df.to_csv(index=False)
        #def clear_cache():
        #  st.cache_data.clear()  # Clear @st.cache_data cache
        #  st.cache_resource.clear()  # Clear @st.cache_resource cache
        #def clear_resource(file):
         
        #  file_buffer = BytesIO()
         # file.to_csv(file_buffer, index=False)  # Save DataFrame as CSV to BytesIO
          
        #  file_buffer.seek(0)  # Reset the buffer's position to the start
        #  ftp_server = ftplib.FTP("users.utcluj.ro", st.secrets['u'], st.secrets['p'])
         # ftp_server.encoding = "utf-8"
         # ftp_server.cwd('./public_html')
          #ftp_server.delete('baza.csv')
         # ftp_server.storbinary('STOR baza.csv', file_buffer)  # Send the file
         # ftp_server.quit()
          #st.cache_data.clear()
          #st.cache_resource.clear()
        # Button to clear cache
        #clear_cache_button = st.form_submit_button("Incarca alta fisa")
        #clear_resource_button = st.form_submit_button("Scrie datele in baza")
        #if clear_cache_button:
        #    clear_cache()
        #   st.success("Cache cleared!")
        #if clear_resource_button:
         #  clear_resource(df)
         #  st.success("Datele au fost scrise!")
