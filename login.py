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
import hashlib
import time


name, user = require_login("🔐 App Login")
st.title("Fisa disciplinei")
st.success(f"Bine ai venit, {name}!")
@st.dialog("info")
def info():
 st.write("Aplicația este pusă la dispoziția dumneavoastră pentru a elimina confuziile și neconcordanțele generate de modificările și noile reglementări ARACIS. Formularul _Fișei disciplinei_ a suferit schimbări de structură, iar această aplicație automatizează procesul de actualizare. Toate informațiile introduse de dumneavoastră sunt transpuse automat în noul șablon oficial.")
 st.write("Datele referitoare la disciplină (denumire, număr de credite, tip de examinare, număr de ore, codul disciplinei etc.) sunt preluate direct din planurile de învățământ. Au fost introduse două capitole noi: Competențe și Rezultatele învățării. Pentru fiecare specializare, formulările oficiale se regăsesc în prezentarea planului de învățământ și sunt afișate de aplicație în etapa de completare. Acestea trebuie adaptate pentru fiecare disciplină în secțiunile corespunzătoare ale fișei.")
 st.write("Pentru fișele încărcate în anii anteriori în baza de date, aceste capitole conțin propuneri destinate titularului de curs. Disciplinele aflate la prima completare pot fi încărcate în format .docx, iar aplicația va prelua automat cât mai multe dintre informațiile existente. De asemenea, este posibilă completarea manuală prin câmpurile afișate. Propunerile generate pot fi acceptate, preluate sau editate.")
 st.write("Datele privind aprobarea în departament și consiliu vor fi actualizate automat atunci când devin disponibile. La final, fișa completată poate fi descărcată și vizualizată în formatul oficial aprobat.")

info()
