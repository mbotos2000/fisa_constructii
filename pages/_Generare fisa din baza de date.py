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
#from docx import Document
name, user = require_login("🔐 Incarcare credentiale")
st.title("Fisa disciplinei")
st.success(f"Bine ai venit, {name}!")
st.session_state['NAME']=name
@st.dialog("info")
def info():
 st.write("Aplicația este pusă la dispoziția dumneavoastră pentru a elimina confuziile și neconcordanțele generate de modificările și noile reglementări ARACIS. Formularul _Fișei disciplinei_ a suferit schimbări de structură, iar această aplicație automatizează procesul de actualizare. Toate informațiile introduse de dumneavoastră sunt transpuse automat în noul șablon oficial.")
 st.write("Datele referitoare la disciplină (denumire, număr de credite, tip de examinare, număr de ore, codul disciplinei etc.) sunt preluate direct din planurile de învățământ. Au fost introduse două capitole noi: Competențe și Rezultatele învățării. Pentru fiecare specializare, formulările oficiale se regăsesc în prezentarea planului de învățământ și sunt afișate de aplicație în etapa de completare. Acestea trebuie adaptate pentru fiecare disciplină în secțiunile corespunzătoare ale fișei.")
 st.write("Pentru fișele încărcate în anii anteriori în baza de date, aceste capitole conțin propuneri destinate titularului de curs. Disciplinele aflate la prima completare pot fi încărcate în format .docx, iar aplicația va prelua automat cât mai multe dintre informațiile existente. De asemenea, este posibilă completarea manuală prin câmpurile afișate. Propunerile generate pot fi acceptate, preluate sau editate.")
 st.write("Datele privind aprobarea în departament și consiliu vor fi actualizate automat atunci când devin disponibile. La final, fișa completată poate fi descărcată și vizualizată în formatul oficial aprobat.")
if st.button("Apasa aici daca doresti sa stii mai multe despre aplicatie"):
	info()

def remaining(exclude=()):
    used = sum(
        st.session_state[f"s37{k}"]
        for k in ["a", "b", "c", "d", "e"]
        if k not in exclude
    )
    return max(tosi - used, 0)

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
def schimba_NAME(new):
    st.session_state['NAME'] = str(new)
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

def schimba_M_r_c(new):
    st.session_state['M_r_c'] = str(new)

def schimba_M_r_a(new):
    st.session_state['M_r_a'] = str(new)

def schimba_M_r_r(new):
    st.session_state['M_r_r'] = str(new)
	
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
if "refresh_data" not in st.session_state:
    st.session_state.refresh_data = False
if 'FormSubmitter:Fisa disciplinei-Treceti la alegerea specializarii' not in st.session_state:
    st.session_state["FormSubmitter:Fisa disciplinei-Treceti la alegerea specializarii"]=''
if 'M_1_7' not in st.session_state:
    st.session_state['M_1_7']='IF – învăţământ cu frecvenţă'
if 'M_1_1' not in st.session_state:
    st.session_state['M_1_1']='Universitatea Tehnica din Cluj Napoca'
keys_dash=['M_3_1' ,'M_3_2' ,'M_3_3_s','M_3_3_l','M_3_3_p','M_3_4','M_3_5','M_3_6_s','M_3_6_l','M_3_6_p']
for key in keys_dash:
    st.session_state.setdefault(key, '-')
if 'NAME' not in st.session_state:
    st.session_state['NAME']="noname"
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
c_trans=['da dovada de initiativa','adopta modalitati de promovare a biodiversitatii si a bunastarii animalelor','ofera consiliere altora','accepta critici si orientari','aplica cunostinte de filozofie, etica si religie','adopta modalitati de reducere a poluarii','respecta angajamente','se adapteaza la cerinte fizice','îsi asuma responsabilitatea','da dovada de hotarâre','efectueaza calcule','gestioneaza frustrarea','calculeaza probabilitati','gestioneaza resurse financiare si materiale',
'demonstreaza angajament','gândeste rapid','evalueaza impactul comportamentului individual asupra mediului','conduce controlul calitatii','da dovada de auto-reflectie','da dovada de curiozitate',
'gestioneaza evolutia personala','abordeaza provocarile în mod pozitiv','gândeste analitic','se adapteaza la schimbare','adopta modalitati de reducere a impactului negativ al consumului',
'mentine bunastarea psihologica','protejeaza sanatatea celorlalti','asigura orientarea catre client','modereaza discutii','aplica cunostinte stiintifice, tehnologice si ingineresti','prelucreaza informatii spatiale','respecta diversitatea valorilor si a normelor culturale','face fata incertitudinii','se adreseaza unui public','lucreaza în echipe','arata încredere','instruieste pe ceilalti','opereaza echipamente hardware digitale','organizeaza informatii, obiecte si resurse','interpreteaza informatii matematice','demonstreaza loialitate','îi conduce pe altii','arata empatie','negociaza compromisuri','face fata stresului','lucreaza eficient','participa activ la viata civica','gândeste critic','îi implica pe ceilalti în comportamente favorabile mediului','este atent la detalii','utilizeaza în cunostinta de cauza sistemul de sanatate',
'promoveaza principiile democratiei si ale statului de drept','se exprima într-un mod creativ','solutioneaza conflicte','apreciaza diversitatea expresiilor culturale si artistice','da dovada de dorinta de învatare',
'utilizeaza cu precizie echipamente, instrumente sau echipamente tehnologice','îsi mentine concentrarea pentru perioade lungi de timp','identifica probleme','îsi mentine aptitudinile fizice',
'solutioneaza probleme','deleaga responsabilitati','utilizeaza software de comunicare si colaborare','construieste spirit de echipa','improvizeaza','respecta reglementarile','deplaseaza obiecte','efectueaza cautari pe internet','administreaza identitatea digitala','demonstreaza credibilitate','demonstreaza spirit antreprenorial','raporteaza fapte','aplica competente de baza în materie de programare','creeaza retele','da dovada de competenta interculturala','lucreaza independent','gândeste în mod inovator','demonstreaza constientizarea riscurilor pentru sanatate','gândeste în mod creativ','gândeste holistic','ia decizii','creeaza continut digital','gestioneaza timpul','aplica standarde de igiena','evalueaza în mod critic informatiile si sursele acestora','gestioneaza bolile cronice','aplica cunostinte în domeniul stiintelor sociale si umaniste','promoveaza idei, produse sau servicii','îi motiveaza pe altii','respecta obligatiile de confidentialitate','planifica','reactioneaza la schimbari fizice sau pericole',
'aplica masuri de securitate digitala','îsi exercita drepturile si responsabilitatile','memoreaza informatii','are o minte deschisa']
c_spec=['abordeaza problemele în mod critic','ajusteaza proiectele produselor','analizeaza nevoile comunitatii','analizeaza traseele potentiale la proiectele de conducte',
'aplica competente de calcul numeric','aplica competente de comunicare în domeniul tehnic','aplica standarde de sanatate si siguranta','asigura conformitatea cu legislatia de mediu','asigura conformitatea cu legislatia în materie de securitate','asigura conformitatea infrastructurii de conducte',
'asigura conservarea resurselor de apa','asigura managementul de proiect','asigura managementul proceselor de licitatie','atenueaza impactul de mediu al proiectelor de conducte','comunica cu echipele de constructori','conduce o echipa în cadrul activitatilor legate de managementul apei','da instructiuni personalului','defineste cerinte tehnice','depisteaza defectiuni ale infrastructurii de conducte','deseneaza schite','detine competente informatice','efectueaza analiza apei',
'efectueaza analiza riscurilor','elaboreaza previziuni statistice','elaboreaza strategii de aparare împotriva inundatiilor si a fenomenelor periculoase','elaboreaza studii tehnice în domeniul feroviar','elaboreaza studiul de fezabilitate','evalueaza impactul de mediu','examineaza constrângerile de constructie în proiectarea arhitecturala','examineaza principii tehnice','executa calcule matematice analitice',
'foloseste instrumentele de masura','gestioneaza bugete','gestioneaza personalul geotehnic','gestioneaza proiecte de inginerie','gestioneaza proiectele de constructie a cailor ferate',
'ia în calcul impactul proprietatilor materialelor asupra transportului prin conducte','identifica riscul de inundatii','integreaza cerintele în materie de constructie în proiectarea arhitecturala','integreaza masuri în proiecte arhitecturale','întelege terminologia financiar-bancara','interpreteaza date stiintifice pentru a evalua calitatea apei','întocmeste rapoarte de lucru','lucreaza la proiectarea conductelor','mentine relatiile cu clientii','mentine relatiile cu furnizorii','monitorizeaza santierul','ofera consiliere în domeniul constructiilor','ofera consiliere în legatura cu proiectele de irigatii','ofera consiliere pentru materiale de constructie','ofera consiliere privind destinatia terenurilor','planifica investigatii geotehnice pe teren',
'proiecteaza conducte cu diferite învelisuri','proiecteaza harti personalizate','proiecteaza lucrari de drenaj','proiecteaza sisteme de energie electrica','proiecteaza sisteme de transport','promoveaza constientizarea problemelor legate de mediu','promoveaza proiectarea inovatoare a infrastructurii',
'promoveaza utilizarea transportului durabil','realizeaza studii privind traseele conductelor','redacteaza rapoarte tehnice','relationeaza cu experti în domeniul feroviar','respecta codul deontologic al serviciilor de transport','respecta reglementarile juridice','satisface cerinte tehnice','sintetizeaza informatii',
'studiaza viitoarele amplasamente ale conductelor','supravegheaza proiecte de constructii','supravegheaza siguranta mediului de lucru','utilizeaza diferite canale de comunicare',
'utilizeaza software CAD','utilizeaza software de desen tehnic']
keys_space=['M_1_2','M_1_3','M_1_4','M_1_5','M_1_6','M_1_8','M_2_1','M_2_2','M_2_3','M_2_2_1','M_2_3_1','M_2_4','M_2_5','M_2_6','M_2_7_1','M_2_7_2',
	    'M_3_8','M_3_9','M_3_11','M_4_1','M_4_2','M_5_1','M_5_2', 'M_6_cp','M_6_ct','M_r_c','M_r_a','M_r_r','M_7_1','M_7_2','M_8_1_1','M_8_1_2','M_8_1_3','M_8_1_4','M_8_1_5',
	    'M_8_1_6', 'M_8_1_7','M_8_1_8', 'M_8_1_9','M_8_1_10','M_8_1_11','M_8_1_12','M_8_1_13','M_8_1_14','M_8_1_mp','8_1_o','Biblio_c','M_8_2_1',
	    'M_8_2_2','M_8_2_3','M_8_2_4','M_8_2_5','M_8_2_6','M_8_2_7','M_8_2_8','M_8_2_9','M_8_2_10','M_8_2_11','M_8_2_12','M_8_2_13','M_8_2_14',
	    '8_1_mp','8_1_o','Biblio_a','9','10_1_c','10_2_c','10_3_c','10_1_a','10_2_a','10_3_a','10_6','da_cu','data_dep','data_fac','dir_dep','tip','mail','dep']
for key in keys_space:
    st.session_state.setdefault(key, '')
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
	    ,'Inteligenta artificiala in ing.civila si manag.constructiilor (AICIV-e) ':'Inginerie civila'}
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
            ,'Managementul proiectelor si evaluarea proprietatii (MPEP)',
	    'Inteligenta artificiala in ing.civila si manag.constructiilor (AICIV)'
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
    for filename in ["lista_cd.csv", "planinv_2025_v2.csv"]:
        with BytesIO() as file_data:
            ftp_server.retrbinary(f"RETR {filename}", file_data.write)
            file_data.seek(0)  # Reset file pointer to the start
            csv_data[filename] = pd.read_csv(file_data, encoding="ISO-8859-1")

    # Download DOCX templates
    docx_files = {}
    for filename in [
        "fisa_template_Mail_R.docx", 
        "fisa_template_Mail_eng.docx", 
        "fisa_template_Mail_curs_R.docx",
        "fisa_template_Mail_curs_eng.docx",
        "fisa_template_Mail_aplicatie_R.docx",
        "fisa_template_Mail_aplicatie_eng.docx"
    ]:
        file_data = BytesIO()
        ftp_server.retrbinary(f"RETR {filename}", file_data.write)
        file_data.seek(0)  # Reset file pointer to the start
        docx_files[filename] = file_data

    # Close FTP connection
    ftp_server.cwd('..')
    ftp_server.cwd('./public_html/Fise/2025_r')

    # Get list of .pkl files
    files = ftp_server.nlst()
    pkl_files = [f for f in files if f.endswith('.pkl')]
    pkl_1=pkl_files
    ftp_server.cwd('..')
    ftp_server.cwd('./2024')

    # Get list of .pkl files
    files = ftp_server.nlst()
    pkl_files = pkl_files + [f for f in files if f.endswith('.pkl') and f not in pkl_1]
    ftp_server.cwd('..')
    ftp_server.cwd('./2025')
    files = ftp_server.nlst()
    pkl_files = pkl_files + [f for f in files if f.endswith('.pkl') and f not in pkl_1]
    ftp_server.quit()

    # Return downloaded files
    return (
        csv_data["lista_cd.csv"], 
        csv_data["planinv_2025_v2.csv"], 
        docx_files["fisa_template_Mail_R.docx"], 
        docx_files["fisa_template_Mail_eng.docx"], 
        docx_files["fisa_template_Mail_curs_R.docx"], 
        docx_files["fisa_template_Mail_curs_eng.docx"], 
        docx_files["fisa_template_Mail_aplicatie_R.docx"], 
        docx_files["fisa_template_Mail_aplicatie_eng.docx"],
        csv_data["lista_cd.csv"],pkl_files
    )


def load_ftp_pdf_file(presc):
    # Establish FTP connection
    ftp_server = ftplib.FTP_TLS("users.utcluj.ro")
    ftp_server.login(user=st.secrets['u'], passwd=st.secrets['p'])
    ftp_server.prot_p()
    
    ftp_server.encoding = "utf-8"
    ftp_server.cwd('./public_html')

    file_data = BytesIO()
    ftp_server.retrbinary(f"RETR prezentare_{presc}.pdf", file_data.write)
    file_data.seek(0)  # Reset pointer so Streamlit reads correctly
    
    return file_data  # return stream, NOT file_data.read()


def load_pkl_from_ftp(file_path):
   
        #ftp = ftplib.FTP("users.utcluj.ro", st.secrets['u'], st.secrets['p'])
 ftp = ftplib.FTP_TLS("users.utcluj.ro")
 ftp.login(user=st.secrets['u'], passwd=st.secrets['p'])
 ftp.prot_p() 
 ftp.encoding = "utf-8"  # Force UTF-8 encoding
 buffer = BytesIO()
 ftp.retrbinary(f"RETR {file_path}", buffer.write)
 buffer.seek(0)
 data = pickle.load(buffer)
 if not isinstance(data, dict):
  data = {"data": data}
 return data  
if st.button("🔄 Refresh FTP Data (apasa doar daca nu s-a actualizat baza de date!!!)"):
    st.session_state.refresh_data = True
if st.session_state.refresh_data:
    load_ftp_file.clear()  # Clear the cache
    st.session_state.refresh_data = False
data,data1,_,_,_,_,_,_,data2,Lista_fisiere=load_ftp_file()

data1['nume_disciplina'] = data1['nume_disciplina'].apply(strip_last)
data1['specializare'] = data1['specializare'].apply(strip_last)
lista_ci=['Licenta', 'Master']
lista_ci=my_function(lista_ci)
st.header("Aplicatie generare fisa disciplinei folosind variante anterioare salvate in baza de date a Facultatii de constructii", divider="gray")
if st.button("Apasa aici doar daca nu este afisata nici o optiune de continuare"):
	st.session_state['ut']=False
if not(st.session_state['ut']):
    add_selectbox_C = st.selectbox(
                'Ciclul de studii?',
                my_function((lista_ci)),key='M_1_5'
            )
    try:
     add_selectbox_SP =  st.selectbox(
                'Programul de studii?',
                specializari[add_selectbox_C],
                key='M_1_6')
    except:
        pass
    try:
        st.session_state['M_1_4']=domeniu[add_selectbox_SP]
        nume_di = data1['nume_disciplina'].loc[(data1['specializare']==st.session_state['M_1_6'])].drop_duplicates().tolist()
        add_selectbox_D = st.selectbox(
                    'Disciplina?',
                    ['']+nume_di,key='M_2_1',
                    help='Toate datele asociate disciplinei vor fi inserate in fisa disciplinei automat din planul de invatamant curent')
    except:
        pass
    
    try:
        filtered = [s for s in Lista_fisiere if add_selectbox_D in s and pres[add_selectbox_SP] in s]
        st.write("Au fost gasite "+str(len(filtered))+" variante ale fisei introduse anterior")
        if len(filtered) != 0:
            
            add_selectbox_C = st.selectbox(
                    'Alege varianta de fisa?',
                    my_function(filtered), key='file'
                )
        if len(filtered) == 0:
                st.write("Nu am gasit nici o varianta afisei introdusa anterior!")
                st.write("Acceseaza linkul de mai jos pentru a incarca o fisa in format docx")
                redirect_url = "https://fisaconstructii-utcn.streamlit.app/Incarcare_fisa_noua"
                st.markdown(f"[Click here to continue]({redirect_url})")
                st.experimental_rerun()  # Only if needed, or use JS for redirect
             
    except:
        pass
#st.session_state['file'] = st.file_uploader("Incarca o fisa a disciplinei in format *docx")

    
    if st.session_state['file']!=None or st.session_state['ut']:
        if st.session_state['file']!=None:
          try:
           data_fis = load_pkl_from_ftp('./public_html/Fise/20'+st.session_state['file'][-9:-7]+'/'+st.session_state['file'])      
          except:
           try:
            data_fis = load_pkl_from_ftp('./public_html/Fise/2025_r/'+st.session_state['file'])   
           except Exception as e:
            st.error(f"Error loading file: {e}")

        
          nume_tit = data['nume'].tolist()
         
          st.title("Fisa disciplinei")
          
          st.write('{:%d-%b-%Y}'.format(date.today()))
          nume_tit1=nume_tit
        
          add_selectbox_TC = st.multiselect('Titulari curs?',
                                                  nume_tit, 
                                                  #placeholder="De exemplu"+find_closest_match(doc_result.body[3][2][1], nume_tit),
                                                  help='Pot fi selectati mai multi titulari de curs.')
          add_selectbox_TA = st.multiselect('Titular aplicatii?',
                                                  nume_tit,  
                                                  #placeholder="De exemplu"+find_closest_match(doc_result.body[3][1][1], nume_tit)
                                                 )
          
          if st.button("Treci la subcapitolul 3.7"):   
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
                        st.write(s)
                        st.write(ss)
                        st.session_state['M_2_3']=s
                        st.session_state['M_2_3_1']=ss
                        st.session_state['M_1_8']=str(data1['nrcrt'].loc[(data1['specializare']==st.session_state['M_1_6']) & (data1['nume_disciplina']==st.session_state['M_2_1'])].values[0])
                        st.session_state['M_2_5']=str(data1['semestru'].loc[(data1['specializare']==st.session_state['M_1_6']) & (data1['nume_disciplina']==st.session_state['M_2_1'])].values[0])
                        indices = data1.loc[(data1['specializare'] == st.session_state['M_1_6']) & (data1['nume_disciplina'] == st.session_state['M_2_1']),'semestru'].index
                        st.session_state['M_2_4']=str(data1.iloc[indices.tolist()[0], 0])
                        st.session_state['M_2_6']=ver[str(data1['examin'].loc[(data1['specializare']==st.session_state['M_1_6']) & (data1['nume_disciplina']==st.session_state['M_2_1'])].values[0]).strip()]
                        st.session_state['M_2_7_1']=str(data1['numecat'].loc[(data1['specializare']==st.session_state['M_1_6']) & (data1['nume_disciplina']==st.session_state['M_2_1'])].values[0])
                        st.session_state['M_2_7_2']=str(data1['obligativ'].loc[(data1['specializare']==st.session_state['M_1_6']) & (data1['nume_disciplina']==st.session_state['M_2_1'])].values[0])
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
                                st.session_state['M_3_3_s']=str(int(data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0]))
                                st.session_state['M_3_6_s']=str(int(14*data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0]))
                                st.session_state['test_aplicatie']=True
                        except:
                            pass
                        try:
                            if data1['curs'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0]=='LABORATOR ':
                                st.session_state['M_3_3_l']=str(int(data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0]))
                                st.session_state['M_3_6_l']=str(int(14*data1['numarore'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0]))
                                st.session_state['test_aplicatie']=True
                        except:
                            pass
                        try:
                            if data1['curs'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1'])&(data1['curs']!='CURS      ')].values[0]=='PROIECT   ':

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
          try:
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
                slide_37a=st.slider('(a) Studiul după manual, suport de curs, bibliografie şi notițe',
                 min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
                 value=int(doc_result.body[5][3][-1][0]),
                 help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
            except:
                 slide_37a=st.slider(
                 '(a) Studiul după manual, suport de curs, bibliografie şi notițe',
                 min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
                 help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')	    
            try:  
                 slide_37b=st.slider('(b) Documentare suplimentară în bibliotecă, pe platforme electronice de specialitate şi pe teren',
                  min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
                  value=int(doc_result.body[5][4][-1][0]),
                help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
            except:  
                 slide_37b=st.slider('(b) Documentare suplimentară în bibliotecă, pe platforme electronice de specialitate şi pe teren',
                 min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
                 help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
            try:
                 slide_37c=st.slider( 'c) Pregătire seminarii / laboratoare, teme, referate, portofolii şi eseuri',
                 min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
                 value=int(doc_result.body[5][5][-1][0]),
                 help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
            except:
                 slide_37c=st.slider('c) Pregătire seminarii / laboratoare, teme, referate, portofolii şi eseuri',
                 min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
                 help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
            try:
                 slide_37d=st.slider( '(d) Tutoriat',
                 min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
                  value=int(doc_result.body[5][6][-1][0]),
                help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
            except:
                 slide_37d=st.slider( '(d) Tutoriat',
                 min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
                 help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
            sd=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f))
            try:
                 slide_37e=st.slider( 'e) Examinări',
                 min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
                 value=int(doc_result.body[5][7][-1][0]),
                 help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
            except:
                 slide_37e=st.slider( 'e) Examinări',
                 min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
                 help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
            sd=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f))
            if not(sd<=0):
                   slide_37f=st.slider('(f) Alte activități:',
                   max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
                   value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
                   help='Completati de la a spre f. Suma orelor de studiu individual este cea din planurile de invatamant')
            else:
                    st.write('(f) Alte activități: 0 ore')
                    slide_37f=0
                    slide_37e+=-sd
           except:
                pass
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
				  
                         
               st.text_area('4.1 Preconditii din curriculum',value=data_fis['M_4_1'],key='M_4_1',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
               st.text_area('4.2 Preconditii de competente',value=data_fis['M_4_2'],key='M_4_2',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
               st.text_area('5.1 Conditii de desfasurare a cursului',value=data_fis['M_5_1'],key='M_5_1',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")        
               st.text_area('5.2 Conditii de desfasurare a aplicatiilor',value=data_fis['M_5_2'],key='M_5_2',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
               try:
                with st.expander("Arată prezentarea planului de invatamant pentru a completa urmatoarea sectiune in concordanta cu principiile definite in misiunea specializarii pentru care scriu aceasta fisa"):
                  pdf_stream=(load_ftp_pdf_file(pres[st.session_state['M_1_6']]))
                  st.pdf(pdf_stream)
               except:
                  pass
               add_selectbox_cp = st.multiselect('6.a Competente profesionale acumulate',c_spec,help='Pot fi alese oricate competente.')
			                           
               ss=''
               for d in add_selectbox_cp:
                            ss+=d.title()+'\n'
                        
               ss=ss[:len(ss)-1]      
               st.session_state['M_6_cp']=ss
               #st.text_area('6.a Competente profesionale acumulate',value=data_fis['M_6_cp'],key='M_6_cp',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")              
               add_selectbox_ct = st.multiselect('6.b Competente transversale',c_trans,help='Pot fi alese oricate competente.')
               ss=''
               for d in add_selectbox_ct:
                            ss+=d.title()+'\n'
                        
               ss=ss[:len(ss)-1]      
               st.session_state['M_6_ct']=ss
               
				   
			   #st.text_area('6.b Competente transversale',value=data_fis['M_6_ct'],key='M_6_ct',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")            
               try:
                st.text_area('7.1 Rezultatele asteptate a invatarii - Cunostiinte',value=data_fis['M_r_c'],key='M_r_c',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
               except:  
                st.text_area('7.1 Rezultatele asteptate a invatarii - Cunostiinte',key='M_r_c',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
               try:			  
                st.text_area('7.2 Rezultatele asteptate a invatarii - Abilitati',value=data_fis['M_r_a'],key='M_r_a',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
               except:
                st.text_area('7.2 Rezultatele asteptate a invatarii - Abilitati',key='M_r_a',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
               try:
                st.text_area('7.3 Rezultatele asteptate a invatarii - Responsabilitatea si autonomie',value=data_fis['M_r_r'],key='M_r_r',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
               except:
                st.text_area('7.3 Rezultatele asteptate a invatarii - Responsabilitatea si autonomie',key='M_r_r',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
               st.text_area('8.1 Obiectivul general al disciplinei',value=data_fis['M_7_1'],key='M_7_1',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
               st.text_area('8.2 Obiectivele specifice',value=data_fis['M_7_2'],key='M_7_2',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
            
               if st.session_state['test_curs']:         
                  st.text_area('Curs 1',value=data_fis['M_8_1_1'],key='M_8_1_1',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")                 
                  st.text_area('Curs 2',value=data_fis['M_8_1_2'],key='M_8_1_2',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")                  
                  st.text_area('Curs 3',value=data_fis['M_8_1_3'],key='M_8_1_3',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
                  st.text_area('Curs 4',value=data_fis['M_8_1_4'],key='M_8_1_4',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
                  st.text_area('Curs 5',value=data_fis['M_8_1_5'],key='M_8_1_5',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!") 
                  st.text_area('Curs 6',value=data_fis['M_8_1_6'],key='M_8_1_6',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
                  st.text_area('Curs 7',value=data_fis['M_8_1_7'],key='M_8_1_7',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")              
                  st.text_area('Curs 8',value=data_fis['M_8_1_8'],key='M_8_1_8',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")                  
                  st.text_area('Curs 9',value=data_fis['M_8_1_9'],key='M_8_1_9',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")                 
                  st.text_area('Curs 10',value=data_fis['M_8_1_10'],key='M_8_1_10',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")                 
                  st.text_area('Curs 11',value=data_fis['M_8_1_11'],key='M_8_1_11',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")                  
                  st.text_area('Curs 12',value=data_fis['M_8_1_12'],key='M_8_1_12',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")                 
                  st.text_area('Curs 13',value=data_fis['M_8_1_13'],key='M_8_1_13',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
                  st.text_area('Curs 14',value=data_fis['M_8_1_14'],key='M_8_1_14',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
                  st.text_area('Metode de predare-Curs',value=data_fis['M_8_1_mp'],key='M_8_1_mp',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
                  st.text_area('Observatii-Curs',value=data_fis['M_8_1_o'],key='M_8_1_o',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")                  
                  st.text_area('Biliografie-Curs',value=data_fis['Biblio_c'],key='Biblio_c',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
               if st.session_state['test_aplicatie']:
                  st.text_area('Aplicatia 1',data_fis['M_8_2_1'],key='M_8_2_1',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
                  st.text_area('Aplicatia 2',data_fis['M_8_2_2'],key='M_8_2_2',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
                  st.text_area('Aplicatia 3',data_fis['M_8_2_3'],key='M_8_2_3',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
                  st.text_area('Aplicatia 4',data_fis['M_8_2_4'],key='M_8_2_4',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
                  st.text_area('Aplicatia 5',data_fis['M_8_2_5'],key='M_8_2_5',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
                  st.text_area('Aplicatia 6',data_fis['M_8_2_6'],key='M_8_2_6',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
                  st.text_area('Aplicatia 7',data_fis['M_8_2_7'],key='M_8_2_7',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
                  st.text_area('Aplicatia 8',data_fis['M_8_2_8'],key='M_8_2_8',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
                  st.text_area('Aplicatia 9',data_fis['M_8_2_9'],key='M_8_2_9',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
                  st.text_area('Aplicatia 10',data_fis['M_8_2_10'],key='M_8_2_10',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
                  st.text_area('Aplicatia 11',data_fis['M_8_2_11'],key='M_8_2_11',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
                  st.text_area('Aplicatia 12',data_fis['M_8_2_12'],key='M_8_2_12',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")         
                  st.text_area('Aplicatia 13',data_fis['M_8_2_13'],key='M_8_2_13',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
                  st.text_area('Aplicatia 14',data_fis['M_8_2_14'],key='M_8_2_14',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")         
                  st.text_area('Metode de predare-aplicatii', data_fis['M_8_1_mp1'],key='M_8_1_mp1',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")                  
                  st.text_area('Observatii-aplicatii',data_fis['M_8_1_o1'],key='M_8_1_o1',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
                  st.text_area('Bibliografie-Aplicatii',data_fis['Biblio_a'],key='Biblio_a',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")

               st.text_area('10. Coroborarea conținuturilor disciplinei cu așteptările reprezentanţilor comunităţii epistemice, a.p. s.a.m.d. din domeniul aferent programului',data_fis['M_9'],key='M_9')
               st.text_area('11.1 Criterii de evaluare curs',data_fis['M_10_1_c'],key='M_10_1_c',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
               st.text_area('11.2 Metode de evaluare curs',data_fis['M_10_2_c'],key='M_10_2_c',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
               st.text_area('11.1.1 Criterii de evaluare aplicatii',data_fis['M_10_1_a'],key='M_10_1_a',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
               st.text_area('11.2 Metode de evaluare aplicatii',data_fis['M_10_2_a'],key='M_10_2_a',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
               st.text_area('11.3 Ponderea notei la curs in nota finala',data_fis['M_10_3_c'],key='M_10_3_c',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
               st.text_area('11.3.1 Ponderea notei la aplicatii in nota finala',data_fis['M_10_3_a'],key='M_10_3_a',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
               st.text_area('11.6 Standard minim de performanţă',data_fis['M_10_6'],key='M_10_6',placeholder="Completati manual. Aplicatia nu a reusit sa identifice text in fisa incarcata!")
              #d_com=st.date_input("Data completarii",date.today())

               d_com=date.today()
               d_dep=st.date_input("Data avizari in departament",date.today())
               d_fac=st.date_input("Data avizari in consiliul facultatii",date.today())
               submitted= st.form_submit_button("finalizeaza")
               if submitted:
                #schimba_decan(decan[add_select])
                st.session_state['dir_dep']=directori[data_fis['M_1_3']]
                st.session_state['data_dep']=str(d_dep)
                #schimba_fac(add_select)
                st.session_state['data_fac']=str(d_fac)
                st.session_state['d_com']=str(d_com)

                del st.session_state["resetare"]
                del st.session_state["FormSubmitter:Fisa disciplinei-Treceti la alegerea specializarii"]
                del st.session_state["FormSubmitter:capitolul 4-finalizeaza"]

                if bool(st.session_state['test_curs'])& bool(st.session_state['test_aplicatie']):
                    if st.session_state['M_1_6']!='Constructii civile, industriale si agricole (CCIA-eng)' or st.session_state['M_1_6']!='Inteligenta artificiala in ing.civila si manag.constructiilor (AICIV-e) ':
                                                _,_,template,_,_,_,_,_,_,_=load_ftp_file()
                    else:
                        _,_,_,template,_,_,_,_,_,_=load_ftp_file()
                
                if bool(st.session_state['test_curs'])& (not(bool(st.session_state['test_aplicatie']))):
                    if st.session_state['M_1_6']!='Constructii civile, industriale si agricole (CCIA-eng)' or st.session_state['M_1_6']!='Inteligenta artificiala in ing.civila si manag.constructiilor (AICIV-e) ':
                                               _,_,_,_,template,_,_,_,_,_=load_ftp_file()
                                               st.session_state['M_1_3']=departamentele[data_fis['M_1_3']]
                    else:
                        _,_,_,_,_,template,_,_,_,_=load_ftp_file()
                        st.session_state['M_1_3']=departamentele[data_fis['M_1_3']]
                if (not(bool(st.session_state['test_curs'])))& bool(st.session_state['test_aplicatie']):
                    if st.session_state['M_1_6']!='Constructii civile, industriale si agricole (CCIA-eng)':
                        _,_,_,_,_,_,template,_,_,_=load_ftp_file()
                    else:
                        _,_,_,_,_,_,_,template,_,_=load_ftp_file()
                        st.session_state['M_1_3']=departamentele[data_fis['M_1_3']]
                st.session_state['M_1_3']=data_fis['M_1_3']
                st.session_state['M_1_2']=data_fis['M_1_2']
                document = MailMerge(template)
                #st.write(document.get_merge_fields())
                document.merge(da_cu=st.session_state['d_com'])
                keys_to_merge=['denumirefisa','dataintocmire','M_1_1','M_1_2','M_1_3','M_1_4','M_1_5','M_1_6','M_1_8',
                               'M_2_1','M_2_2','M_2_3','M_2_3_1','M_2_4','M_2_5','M_2_6','M_2_2_1','M_2_7_1','M_2_7_2',
                               'M_3_1','M_3_2','M_3_3_l','M_3_3_s','M_3_3_p','M_3_4','M_3_5','M_3_6_s','M_3_6_l','M_3_6_p','M_3_7_a','M_3_7_b','M_3_7_c','M_3_7_d','M_3_7_e','M_3_7_f',
                               'M_3_8','M_3_9','M_3_11','M_4_1','M_4_2','M_5_1','M_5_2','M_6_cp','M_6_ct','M_r_c','M_r_a','M_r_r','M_7_1','M_7_2',
                               'M_8_1_1','M_8_1_2','M_8_1_3','M_8_1_4','M_8_1_5','M_8_1_6','M_8_1_7','M_8_1_8','M_8_1_9','M_8_1_11','M_8_1_12','M_8_1_13','M_8_1_14','Biblio_c',
                               'M_8_2_1','M_8_2_2','M_8_2_3','M_8_2_4','M_8_2_5','M_8_2_6','M_8_2_7','M_8_2_8','M_8_2_9','M_8_2_10','M_8_2_12','M_8_2_13','M_8_2_14','Biblio_a','M_9',
                               'M_10_1_a','M_10_1_c','M_10_2_c','M_10_3_a','M_10_3_c','M_10_6','M_10_2_a','M_8_1_10',
                               'M_8_1_o1','M_8_1_mp','M_8_1_mp1','M_8_1_o','dep','da_cu','data_fac','data_dep','tip','dir_dep','decan','fac','NAME']
                data_ftp=pd.DataFrame(columns=keys_to_merge)
                
                for key in keys_to_merge:
                  if key in st.session_state:
                    document.merge(**{key: st.session_state[key]})
                file_name=st.session_state['M_1_8']+'_FD_an'+st.session_state['M_2_4']+'_s'+st.session_state['M_2_5']+'_'+pres[st.session_state['M_1_6']]+'_'+st.session_state['M_2_1']+'_26-27.docx'
                file_name_bak=st.session_state['M_1_8']+'_FD_an'+st.session_state['M_2_4']+'_s'+st.session_state['M_2_5']+'_'+pres[st.session_state['M_1_6']]+'_'+st.session_state['M_2_1']+'_26-27_'+str(int(time.time()))+'.pkl'
                remote_filename=st.session_state['M_1_8']+'_FD_an'+st.session_state['M_2_4']+'_s'+st.session_state['M_2_5']+'_'+pres[st.session_state['M_1_6']]+'_'+st.session_state['M_2_1']+'_26-27.pkl'
                remote_filename_csv=st.session_state['M_1_8']+'_FD_an'+st.session_state['M_2_4']+'_s'+st.session_state['M_2_5']+'_'+pres[st.session_state['M_1_6']]+'_'+st.session_state['M_2_1']+'_24-25.csv'
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
                session_data = {key: str(st.session_state.get(key, '')) for key in st.session_state}
                pickle_buffer = BytesIO()
                pickle.dump(session_data, pickle_buffer)
                pickle_buffer.seek(0)
                pickle_buffer_bak = BytesIO(pickle_buffer.getvalue())
                docx_buff = BytesIO()
                document.write(docx_buff)
                docx_buff.seek(0)
                with ftplib.FTP_TLS("users.utcluj.ro") as ftp:
                 ftp.login(user=st.secrets['u'], passwd=st.secrets['p'])
                 ftp.prot_p()
                 ftp.encoding = "utf-8"
                 ftp.cwd("/public_html/Fise/2026")
                 ftp.storbinary(f"STOR {remote_filename}", pickle_buffer)
                 ftp.cwd("/public_html/Fise/2026_bak")
                 ftp.storbinary(f"STOR {file_name_bak}", pickle_buffer_bak)
                 ftp.cwd("/public_html/Fise/2026")
                 ftp.storbinary(f"STOR {file_name}", docx_buff)
