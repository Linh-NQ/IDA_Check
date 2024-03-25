#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import shutil
import tempfile
import xlwings as xw
import numpy as np
import glob
import re
from datetime import datetime
import os
import regex

from tkinter import *
from tkinter import ttk
from tkinter import filedialog


# In[2]:


# Allgemeine Funktionen
def return_error_rows_as_string(error_row):
    error_str = ''
    for e in error_row:
        error_str = error_str + str(e) + ', '
    error_str = error_str[:-2]    
    # Split the string by commas and convert each element to an integer
    numbers_list = [int(num.strip()) for num in error_str.split(',')]
    # Sort the list in ascending order
    numbers_list.sort()
    # Initialize variables
    result = []
    start = numbers_list[0]
    end = numbers_list[0]
    # Iterate over the list and find consecutive numbers
    for i in range(1, len(numbers_list)):
        if numbers_list[i] == end + 1:
            end = numbers_list[i]
        else:
            # Add the range to the result
            if start == end:
                result.append(str(start))
            else:
                result.append(f'{start}-{end}')
            start = numbers_list[i]
            end = numbers_list[i]
    # Add the last range to the result
    if start == end:
        result.append(str(start))
    else:
        result.append(f'{start}-{end}')
    # Join the ranges with commas
    error_str = ', '.join(result)    
    return(error_str)



def zellen_bunt_malen(error_str, column_name, ws, color):
    # get column index from column name
    column_index = vorlage.columns.get_loc(column_name)
    column_letter = xw.utils.col_name(column_index + 1)
    # cell coordinates
    error_row_list = error_str.split(',')
    for error_row in error_row_list:
        if '-' in error_row:
            first_row = error_row.split('-')[0].strip()
            last_row = error_row.split('-')[1].strip()
            cell_coord = column_letter + first_row + ':' + column_letter + last_row
        else:
            cell_coord = column_letter + error_row.strip()
        cell = ws.range(cell_coord)
        cell.color = color
        
        

def open_dm_file(name, path, sheetname):
    """ Öffnet eine der Dateien des DMs.
        Dafür werden der Name der Datei, Pfad und sheet name als Argumente übergeben.
    """
    
    files = glob.glob(path + '\*.xlsx*')
    for i in range(len(files)):
        if name in files[i]:
            path = files[i]
    try:
        excel = pd.read_excel(path, sheet_name = sheetname)
    # falls Datei geöffnet ist:
    except:
        if '~$' in path:
            path = path.replace('~$', '')
        # Generate a temporary directory to store the copied file
        temp_dir = tempfile.mkdtemp()        
        # Generate a temporary file name with a .xlsx extension in the temporary directory
        temp_file_path = tempfile.mktemp(suffix='.xlsx', dir=temp_dir)        
        # Copy the open Excel file to the temporary location
        shutil.copy2(path, temp_file_path)
        excel = pd.read_excel(temp_file_path, sheet_name = sheetname)    
        shutil.rmtree(temp_dir)
    return(excel)

def open_closed_or_opened_file(path):
    """ Mit der Funktion wird zuerst versucht, die Excel-Datei mit pandas zu öffnen.
        Falls die Datei bereits von einem User geöffnet ist, wird eine Kopie erstellt.
    """
    date_format = '%d.%m.%Y'
    try:
        df = pd.read_excel(path, decimal=',')
    except:
        if '~$' in path:
            path = path.replace('~$', '')
        temp_dir = tempfile.mkdtemp()        
        temp_file_path = tempfile.mktemp(suffix='.xlsx', dir=temp_dir)
        shutil.copy2(path, temp_file_path)
        try:
            df = pd.read_excel(temp_file_path, decimal=',')  # Die temporäre Datei wird hier geöffnet
        finally:
            shutil.rmtree(temp_dir)
            
    for col in df.columns:
        # Check if the column contains datetime-like values
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            for i, value in enumerate(df[col]):
                # Check if the value is a datetime object with time included
                if hasattr(value, 'hour'):
                    # Convert datetime object to date (without time)
                    df.at[i, col] = value.date()
                    if str(df.at[i, col]) != 'NaT':
                        try:
                            date_object = value.strftime('%d.%m.%Y')
                            df.at[i, col] = date_object
                            
                        except:
                            pass
                
    return df


# In[21]:


# Check-Funktionen

# Check 1: Feldcode-Name
cols_richtig = {}
def check_feldcode(cols, felder, felder_origin, ws, error_count, error_count_total):
    global felder_f, cols_richtig
    
    cols_richtig = {}
    felder_liste = list(felder_origin['Feldcode'])
    felder_liste = [feld.strip() for feld in felder_liste]
    felder_liste_lower = [entry.lower() for entry in felder_liste]
    
    for col in cols:
        if col not in felder_liste:
            # Suche nach richtiger Schreibweise
            for i in range(len(felder_liste)):
                if col.lower() in felder_liste_lower:
                    index_feld = felder_liste_lower.index(col.lower())
                    col_richtig = felder_liste[index_feld]
                    cols_richtig[col_richtig] = col
                    # felder_f mit falsch geschriebenem Spaltennamen ergänzen
                    felder_f = pd.concat([felder_f, felder[felder['Feldcode']==col_richtig]], ignore_index=True)
                    text_feldcode = Label(frame, text = "Feldcode '{}' ist falsch, richtig wäre '{}'".format(col, col_richtig),
                                            bg = '#eeeee4', font=('Ink free',11))
                    text_feldcode.pack()
                    zellen_bunt_malen('1', col, ws, (220, 20, 60))
                    error_count_total[0] += 1
                    error_count[0] += 1
                    break
                else:
                    if len(col) < 15:
                        match = regex.search(r'({}){}'.format(col.lower(), '{e<3}'), felder_liste_lower[i])
                    else:
                        match = regex.search(r'({}){}'.format(col.lower(), '{e<6}'), felder_liste_lower[i])
                    if match:
                        if ((len(felder_liste[i]) <= len(col) + 2) & (len(felder_liste[i]) >= len(col) - 2)):
                            col_richtig = felder_liste[i]
                            cols_richtig[col_richtig] = col
                            # felder_f mit falsch geschriebenem Spaltennamen ergänzen
                            felder_f = pd.concat([felder_f, felder[felder['Feldcode']==col_richtig]], ignore_index=True)
                            text_feldcode = Label(frame, text = "Feldcode '{}' ist falsch, richtig wäre '{}'".format(col, col_richtig),
                                                    bg = '#eeeee4', font=('Ink free',11))
                            text_feldcode.pack()
                            zellen_bunt_malen('1', col, ws, (220, 20, 60))
                            error_count_total[0] += 1
                            error_count[0] += 1
                            break
                        else:
                            # Leerzeichen vor oder nach dem Wort
                            if len(col) != len(col.strip()):
                                text_feldcode = Label(frame, text = "Feldcode '{}' ist falsch (Leerzeichen zu viel)".format(col),
                                                        bg = '#eeeee4', font=('Ink free',11))
                                text_feldcode.pack()
                                zellen_bunt_malen('1', col, ws, (220, 20, 60))
                                error_count_total[0] += 1
                                error_count[0] += 1
                                break                                
                            else:
                                text_feldcode = Label(frame, text = "Feldcode '{}' ist falsch bzw. nicht in IDA".format(col),
                                                        bg = '#eeeee4', font=('Ink free',11))
                                text_feldcode.pack()
                                zellen_bunt_malen('1', col, ws, (220, 20, 60))
                                error_count_total[0] += 1
                                error_count[0] += 1
                                break
                    # Richtige Schreibweise nicht gefunden
                    else:
                        if i == len(felder_liste)-1:
                            text_feldcode = Label(frame, text = "Feldcode '{}' ist falsch bzw. nicht in IDA".format(col),
                                                    bg = '#eeeee4', font=('Ink free',11))
                            text_feldcode.pack()
                            zellen_bunt_malen('1', col, ws, (220, 20, 60))
                            error_count_total[0] += 1
                            error_count[0] += 1
    

# Check 2: Master ID
def check_masterid(vorlage, ws, error_count, error_count_total):
    master_id_pattern = r'I\d{9}\b'
    error_master_id_row = []
    if 'master_id' in cols_richtig:
        master_id = cols_richtig['master_id']
    else:
        master_id = 'master_id'
    for i in range(len(vorlage[master_id])):
        match = re.search(master_id_pattern, str(vorlage[master_id][i]))
        if match:
            continue
        else:
            if str(vorlage[master_id][i]) == 'nan':
                continue
            else:
                error_master_id_row.append(i+2)
    if error_master_id_row != []:
        error_str = return_error_rows_as_string(error_master_id_row)
        text_master = Label(frame, text = "Fehler in Spalte 'master_id' Zeile(n) {}".format(error_str),
                            bg = '#eeeee4', font=('Ink free',11))
        text_master.pack()
        # Zelle in Excel-Datei rot färben
        zellen_bunt_malen(error_str, master_id, ws, (220, 20, 60))
        error_count_total[0] += 1
        error_count[0] += 1
        

# Check 3: Sample ID
def check_sampleid(cols, cols_l, vorlage, ws, error_count, error_count_total):
    sample_id_pattern = r'I\d{12}\b'
    if ('sample_id' in cols_richtig):
        sample_id = cols_richtig['sample_id']
    else:
        sample_id = 'sample_id'
    error_sample_id_row = []
    for i in range(len(vorlage[sample_id])):
        match = re.search(sample_id_pattern, str(vorlage[sample_id][i]))
        if match:
            continue
        else:
            if str(vorlage[sample_id][i]) == 'nan':
                continue
            else:
                error_sample_id_row.append(i+2)
    if error_sample_id_row != []:
        error_str = return_error_rows_as_string(error_sample_id_row)
        text_sample = Label(frame, text = "Fehler in Spalte 'sample_id' Zeile(n) {}".format(error_str),
                                bg = '#eeeee4', font=('Ink free',11))
        text_sample.pack()
        # Zelle in Excel-Datei rot färben
        zellen_bunt_malen(error_str, sample_id, ws, (220, 20, 60))
        error_count_total[0] += 1
        error_count[0] += 1        

# Check 4: Pat ID
def check_patid(vorlage, ws, error_count, error_count_total):
    error_pat_id_row = []
    if ('pat_id' in cols_richtig):
        pat_id = cols_richtig['pat_id']
    else:
        pat_id = 'pat_id'
    # Zuerst wird geprüft, ob die gesamte Spalte im Format np.int64 ist
    if vorlage[pat_id].dtype != np.int64:
        if vorlage[pat_id].dtype == 'object':
            for i in range(len(vorlage[pat_id])):
                if not str(vorlage[pat_id][i]).isdigit():
                    error_pat_id_row.append(i+2)
    if error_pat_id_row != []:
        error_str = return_error_rows_as_string(error_pat_id_row)
        text_pat = Label(frame, text = "(Format-)Fehler in Spalte 'pat_id' Zeile(n) {}".format(error_str),
                        bg = '#eeeee4', font=('Ink free',11))
        text_pat.pack()
        # Zelle in Excel-Datei rot färben
        zellen_bunt_malen(error_str, pat_id, ws, (220, 20, 60))
        error_count_total[0] += 1
        error_count[0] += 1
        
# Check 5: Sample und Master ID
def check_sample_master(vorlage, ws, error_count, error_count_total):
    if ('sample_id' in cols_richtig):
        sample_id = cols_richtig['sample_id']
    else:
        sample_id = 'sample_id'
        
    if ('master_id' in cols_richtig):
        master_id = cols_richtig['master_id']
    else:
        master_id = 'master_id'
        
    error_sample_master = []
    for i in range(len(vorlage[sample_id])):
        if str(vorlage[sample_id][i]).split('I')[-1][:-3] != str(vorlage[master_id][i]).split('I')[-1]:
            error_sample_master.append(i+2)
    if error_sample_master != []:
        error_str = return_error_rows_as_string(error_sample_master)
        text_pat = Label(frame, text = "'sample_id' stimmt nicht mit 'master_id' überein in Zeile(n) {}".format(error_str),
                        bg = '#eeeee4', font=('Ink free',11))
        text_pat.pack()
        # Zelle in Excel-Datei rot färben
        zellen_bunt_malen(error_str, sample_id, ws, (220, 20, 60))
        zellen_bunt_malen(error_str, master_id, ws, (220, 20, 60))
        error_count_total[0] += 1
        error_count[0] += 1
        
# Check 6: monovalent lists
def check_monovalent(cols, felder_f, ws, error_count, error_count_total):
    ana_fields = ['Ana_parameter', 'Ana_parameter_value', 'Ana_parameter_date']
    error_ana_block = {}
    error_ana_field1 = {}
    error_ana_field2 = {}
    
    for i in range(len(felder_f)):
        field_error_row = []

        if 'Monovalent list' in felder_f['Feldart'][i]:
            check_col = felder_f['Feldcode'][i]
            if check_col in cols_richtig:
                check_col = cols_richtig[check_col]
            # Ausnahme bei 'Ana_parameter_avail_xy'
            if 'Ana_parameter_avail' in check_col:
                possible_choices = ['01', '02', '1', '2', '1.0', '2.0']
                nr = check_col.split('_')[-1]
                check_entry = list(vorlage[check_col])
                for j in range(len(check_entry)):                      
                    if str(check_entry[j])=='nan':
                        field_error_row.append(j+2)
                    else:
                        try:
                            if str(int(check_entry[j])) not in possible_choices:
                                ana_error_row.append(j+2)
                        except:
                            ana_error_row.append(j+2)
                    # Check bei 1
                    if '1' in str(check_entry[j]):
                        for field in ana_fields:
                            if str(vorlage[field+str(nr)][j]) == 'nan':
                                if field+str(nr) not in error_ana_field1:
                                    error_ana_field1[field+str(nr)] = [j+2]
                                else:
                                    if j+2 not in error_ana_field1[field+str(nr)]:
                                        error_ana_field1[field+str(nr)].append(j+2)

                        # Check, ob in den Blöcken davor eine 2 steht
                        if check_col != 'Ana_parameter_avail_1':
                            akt_nr = int(check_col.split('_')[-1])
                            list_nr = list(range(1, int(akt_nr)))
                            for l_nr in list_nr:
                                ana_col = 'Ana_parameter_avail_' + str(l_nr)
                                if ana_col in cols_richtig:
                                    ana_col = cols_richtig[ana_col]
                                if (str(vorlage[ana_col][j]) == '2') | (str(vorlage[ana_col][j]) == '2.0'):
                                    if ana_col not in error_ana_block:
                                        error_ana_block[ana_col] = [j+2]
                                    else:
                                        if j+2 not in error_ana_block[ana_col]:
                                            error_ana_block[ana_col].append(j+2)

                    # Check bei 2
                    if '2' in str(check_entry[j]):
                        ana_check_start = cols.index(check_col) +1
                        ana_check_stop = ana_check_start
                        for l in range(ana_check_start, len(cols)):
                            if 'Ana_parameter_avail' not in cols[l]:
                                ana_check_stop += 1
                            else:
                                break
                        for l in range(ana_check_start, ana_check_stop):
                            field = cols[l]
                            if str(vorlage[field][j]) != 'nan':
                                if field not in error_ana_field2:
                                    error_ana_field2[field] = [j+2]
                                else:
                                    if j+2 not in error_ana_field2[field]:
                                        error_ana_field2[field].append(j+2)
                            
            else:           
                # falls Auswahlmöglichkeiten in einer weiteren Excel-Datei hinterlegt sind,
                # wird der Pfad angegeben
                possible_choices = felder_f['Auswahlmöglichkeiten bei Listen'][i]
                if 'O:' in possible_choices:
                    if '\n' in possible_choices:
                        possible_choices = possible_choices.replace('\n', '')
                    path = possible_choices.split('siehe ')[-1]
                    if '-->' in path:
                        filename = path.split('-->')[1].strip()
                        path = path.split('-->')[0].strip()
                    else:
                        filename = path.split('\\')[-1].split('_')[0]
                        path = '\\'.join(path.split('\\')[:-1])
                    files = glob.glob(path+'\*.xlsx')
                    for file in files:
                        if (filename in file) | (' '.join(filename.split('_')) in file):
                            path = file                
                    df_choices = open_closed_or_opened_file(path)
                    # erste Spalte ablesen und als Liste abspeichern
                    possible_choices = df_choices.iloc[:,0].tolist()
                else:
                    if '\n' in possible_choices:
                        possible_choices = possible_choices.split('\n')
                    elif ',' in possible_choices:
                        possible_choices = possible_choices.split(',')
                # muss ggf. in List umgewandelt werden
                if not isinstance(possible_choices, list):
                    possible_choices = [possible_choices]    
                for j in range(len(possible_choices)):
                    c = possible_choices[j].split('=')[0]
                    try:
                        c = int(c)
                        possible_choices[j] = possible_choices[j].split('=')[0]
                    except:
                        try:
                            possible_choices[j] = possible_choices[j].split('=')[1]
                        except:
                            possible_choices[j] = '00'

                possible_choices = [str(int(c)) for c in possible_choices]
                # Ablesen des Eintrags in der zu importierenden Datei (vorlage)
                check_entry = list(vorlage[check_col])
                for j in range(len(check_entry)):                      
                    if str(check_entry[j])=='nan':
                        continue
                    else:
                        try:
                            if str(int(check_entry[j])) not in possible_choices:
                                field_error_row.append(j+2)
                            elif ('.' in str(check_entry[j])) & (str(check_entry[j]).split('.')[-1]!='0'):
                                field_error_row.append(j+2)
                        except:
                            field_error_row.append(j+2)
            # Falls Fehler in eins der Felder, dann erscheint Fehlermeldung
            if field_error_row != []:
                error_str = return_error_rows_as_string(field_error_row)
                text_field = Label(frame, text = "Fehler in Spalte '{}' Zeile(n) {}".format(check_col, error_str),
                                    bg = '#eeeee4', font=('Ink free',11))
                text_field.pack()
                # Zelle in Excel-Datei rot färben
                zellen_bunt_malen(error_str, check_col, ws, (220, 20, 60))
                error_count_total[0] += 1
                error_count[0] += 1
                                    
                    
    if error_ana_field1 != {}:
        for para in list(error_ana_field1.keys()):
            error_str = return_error_rows_as_string(error_ana_field1[para])
            text_field = Label(frame, text = "Feld muss ausgefüllt sein in Spalte '{}' Zeile(n) {}".format(para, error_str),
                                bg = '#eeeee4', font=('Ink free',11))
            text_field.pack()
            # Zelle in Excel-Datei rot färben
            zellen_bunt_malen(error_str, para, ws, (220, 20, 60))
            error_count_total[0] += 1
            error_count[0] += 1
            
    if error_ana_field2 != {}:
        for para in list(error_ana_field2.keys()):
            error_str = return_error_rows_as_string(error_ana_field2[para])
            text_field = Label(frame, text = "Feld muss leer sein sein in Spalte '{}' Zeile(n) {}".format(para, error_str),
                                bg = '#eeeee4', font=('Ink free',11))
            text_field.pack()
            # Zelle in Excel-Datei rot färben
            zellen_bunt_malen(error_str, para, ws, (220, 20, 60))
            error_count_total[0] += 1
            error_count[0] += 1
                                        
    if error_ana_block != {}:
        for ana_para in list(error_ana_block.keys()):
            error_str = return_error_rows_as_string(error_ana_block[ana_para])
            text_field = Label(frame, text = "Fehler in Anordnung der Blöcke Spalte '{}' Zeile(n) {}".format(ana_para, error_str),
                                bg = '#eeeee4', font=('Ink free',11))
            text_field.pack()
            # Zelle in Excel-Datei rot färben
            zellen_bunt_malen(error_str, ana_para, ws, (220, 20, 60))
            error_count_total[0] += 1
            error_count[0] += 1
                    
            
# Check 7: Datum-Spalte
def check_datum(felder_f, vorlage, ws, error_count, error_count_total):
    for i in range(len(felder_f)):
        date_error_row = []
        if 'Date' in felder_f['Feldart'][i]:
            date_col = felder_f['Feldcode'][i]
            if date_col in cols_richtig:
                date_col = cols_richtig[date_col]
            for j in range(len(vorlage)):
                if not pd.isnull(vorlage[date_col][j]):
                    try:
                        date_string = str(vorlage[date_col][j]).split(' ')[0]
                        try:
                            date_object = datetime.strptime(date_string, "%d.%m.%Y").date()
                        except:
                            try:
                                date_object = datetime.strptime(date_string, "%Y-%m").date()
                            except:
                                try:
                                    date_object = datetime.strptime(date_string, "%Y-%m-%d").date()
                                except:
                                    date_error_row.append(j+2)
                    except:
                        date_error_row.append(j+2)
        if date_error_row != []:
            error_str = return_error_rows_as_string(date_error_row)
            text_date = Label(frame, text = "Fehler in Datumspalte '{}' Zeile(n) {}".format(date_col, error_str),
                              bg = '#eeeee4', font=('Ink free',11))
            text_date.pack()
            # Zelle in Excel-Datei rot färben
            zellen_bunt_malen(error_str, date_col, ws, (220, 20, 60))
            error_count_total[0] += 1
            error_count[0] += 1
            
            
# Check 8: One line text
def check_one_line_text(felder_f, vorlage, ws, error_count, error_count_total):
    for i in range(len(felder_f)):
        long_text_row = []
        if 'One line text' in felder_f['Feldart'][i]:
            check_col = felder_f['Feldcode'][i]
            if check_col in cols_richtig:
                check_col = cols_richtig[check_col]
            for j in range(len(vorlage)):
                text = vorlage[check_col][j]
                if len(str(text)) > 120:
                    long_text_row.append(j+2)
        if long_text_row != []:
            error_str = return_error_rows_as_string(long_text_row)
            text_long = Label(frame, text = "Text wurde abgeschnitten in Spalte '{}' Zeile(n) {}".format(check_col, error_str),
                              bg = '#eeeee4', font=('Ink free',11))
            text_long.pack()
            # Zelle in Excel-Datei gelb färben
            zellen_bunt_malen(error_str, check_col, ws, (238, 232, 170))
            error_count_total[0] += 1
            error_count[0] += 1
            
            
# Check 9: Katalog-Felder
def check_katalog(felder_f, ws, error_count, error_count_total):
    for i in range(len(felder_f)):
        field_error_row = []
        if 'Katalog' in felder_f['Feldart'][i]:
            check_col = felder_f['Feldcode'][i]
            if check_col in cols_richtig:
                check_col = cols_richtig[check_col]
            possible_choices = felder_f['Auswahlmöglichkeiten bei Listen'][i]
            if '\n' in possible_choices:
                possible_choices = possible_choices.replace('\n', '')
            if 'O:' in possible_choices:
                path = possible_choices.split('siehe ')[-1]
                if '-->' in path:
                    filename = path.split('-->')[1].strip()
                    path = path.split('-->')[0].strip()
                else:
                    filename = path.split('\\')[-1].split('_')[0]
                    path = '\\'.join(path.split('\\')[:-1])
                files = glob.glob(path+'\*.xlsx')
                for file in files:
                    if (filename in file) | (' '.join(filename.split('_')) in file):
                        path = file                
            df_choices = open_closed_or_opened_file(path)
            # erste Spalte ablesen und als Liste abspeichern
            possible_choices = df_choices.iloc[:,0].tolist()
            possible_choices = [str(int(c)) for c in possible_choices if str(c) != 'nan']
            # Ablesen des Eintrags in der zu importierenden Datei (vorlage)
            check_entry = list(vorlage[check_col])
            for j in range(len(check_entry)):
                if str(check_entry[j])=='nan':
                    continue
                else:
                    try:
                        if str(int(check_entry[j])) not in possible_choices:
                            field_error_row.append(j+2)
                        elif ('.' in str(check_entry[j])) & (str(check_entry[j]).split('.')[-1]!='0'):
                            field_error_row.append(j+2)
                    except:
                        field_error_row.append(j+2)
        # Falls Fehler in eins der Felder, dann erscheint Fehlermeldung
        if field_error_row != []:
            error_str = return_error_rows_as_string(field_error_row)
            text_kat = Label(frame, text = "Fehler in Spalte '{}' Zeile(n) {}".format(check_col, error_str),
                            bg = '#eeeee4', font=('Ink free',11))
            text_kat.pack()
            # Zelle in Excel-Datei rot färben
            zellen_bunt_malen(error_str, check_col, ws, (220, 20, 60))
            error_count_total[0] += 1
            error_count[0] += 1
                                        
            
# Check 10: Pflichtfeld
def check_pflichtfeld(cols, felder_f, ws, error_count, error_count_total):
    
    # Funktion für Spalte Audit_Trail
    def is_valid_format(input_string):
        pattern = re.compile(r'^(\d{2}\.\d{2}\.\d{4})/([a-zA-Z0-9]{3}):')
        match = pattern.match(input_string)

        if match:
            date_str = match.group(1).split('/')[0]
            try:
                # Attempt to parse the extracted date string as a date
                datetime.strptime(date_str, '%d.%m.%Y')
                return True
            except ValueError:
                return False
        else:
            return False
        
    # falsch geschriebene Spalten
    fillcap = 'fillcapacity'
    if 'fillcapacity' in cols_richtig:
        fillcap = cols_richtig['fillcapacity']
    sample_approval = 'sampleApproval'
    if 'sampleApproval' in cols_richtig:
        sample_approval = cols_richtig['sampleApproval'] 
                    
    
    # Ausnahmefelder definieren
    ausnahme_felder = ['diagnosis_ICD', 'diagnosis_text', 'x', 'y', 'cbh', 'comments', 'lockReason']
    # alle Felder mit 'discharge' im Namen sind ebenfalls Ausnahmefelder
    discharge_felder = ['discharge_quantity', 'discharge_reason', 'discharge_project', 
                        'discharge_recipient', 'discharge_date', 'discharge_cost_unit']
    # discharge_felder mit falsch geschriebenen Feldern ergänzen
    for feld in discharge_felder:
        if feld in cols_richtig:
            discharge_felder.append(cols_richtig[feld])
    
    for i in range(len(felder_f)):
        if ('discharge' in felder_f['Feldcode'][i]) & (felder_f['Feldcode'][i] not in discharge_felder):
            ausnahme_felder.append(felder_f['Feldcode'][i])
    
    for i in range(len(felder_f)):
        pflichtfeld_row = []
        kein_pflichtfeld_row = []
        pflichtfeld_period =[]
        pflichtfeld_row_diag = []
        pflichtfeld_row_dis = []
        pflichtfeld_cbh = []
        pflichtfeld_lock = []
        soll_row = []
        pflichtfeld_container = []
        audit_semikolon = []
        audit_fail = []
        audit_komma = []
        # Pflichtfeld muss immer gefüllt sein
        if ('leer' not in felder_f['Pflichtfeld'][i]) & ((felder_f['Pflichtfeld'][i] == 'Pflichtfeld') | ('Pflichtfeld?' in felder_f['Pflichtfeld'][i])):
            # Ausnahmen bei diagnosis_ICD und diagnosis_text:
            # sind beide als Pflichtfelder angegeben, aber nur eins von beiden muss ausgefüllt sein
            if (('diagnosis_ICD' in cols) & ('diagnosis_text' in cols)) & ('diagnosis_ICD' == felder_f['Feldcode'][i]):
                for j in range(len(vorlage)):
                    if (str(vorlage['diagnosis_ICD'][j]) == 'nan') & (str(vorlage['diagnosis_text'][j]) == 'nan'):
                        pflichtfeld_row_diag.append(j+2)
                if pflichtfeld_row_diag != []:
                    error_str = return_error_rows_as_string(pflichtfeld_row_diag)
                    text_pflicht = Label(frame, text = "diagnosis_ICD und diagnosis_text sind beide leer Zeile(n) {}".format(error_str),
                                            bg = '#eeeee4', font=('Ink free',11))
                    text_pflicht.pack()
                    # Zelle in Excel-Datei orange färben
                    zellen_bunt_malen(error_str, 'diagnosis_ICD', ws, (255, 127, 80))
                    zellen_bunt_malen(error_str, 'diagnosis_text', ws, (255, 127, 80))
                    error_count_total[0] += 1
                    error_count[0] += 1
            
            check_col = felder_f['Feldcode'][i]
            if check_col in cols_richtig:
                check_col = cols_richtig[check_col]
            
            # Ausnahme bei x und y
            if (check_col.lower() == 'x') | (check_col.lower() == 'y'):
                continue
                
            # Ausnahme bei discharge-Felder
            if check_col in discharge_felder:
                for j in range(len(vorlage)):
                    if ((str(vorlage[fillcap][j]) == '0') | (str(vorlage[fillcap][j]) == '0.0')) & (str(vorlage[check_col][j]) == 'nan'):
                        pflichtfeld_row_dis.append(j+2)
                if pflichtfeld_row_dis != []:
                    error_str = return_error_rows_as_string(pflichtfeld_row_dis)
                    text_pflicht = Label(frame, text = "Pflichtfeld ist leer in Spalte '{}' Zeile(n) {}".format(check_col, error_str),
                                            bg = '#eeeee4', font=('Ink free',11))
                    text_pflicht.pack()
                    # Zelle in Excel-Datei orange färben
                    zellen_bunt_malen(error_str, check_col, ws, (255, 127, 80))
                    error_count_total[0] += 1
                    error_count[0] += 1
                    
            # Ausnahme bei cbh
            try:
                check_col_richtig = cols_richtig['cbh']
            except:
                check_col_richtig = ''
            if (check_col == 'cbh') | (check_col == check_col_richtig):                                                           
                for j in range(len(vorlage)):
                    try:
                        check_cap = float(vorlage[fillcap][j])
                    except:
                        check_cap = vorlage[fillcap][j]
                    try:
                        check_sample = float(vorlage[sample_approval][j])
                    except:
                        check_sample = vorlage[sample_approval][j]
                    try:
                        check_cbh = float(vorlage[check_col][j])
                    except:
                        check_cbh = vorlage[check_col][j]
                    
                    if (check_cap == 0.0) | (check_sample == 2.0):
                        if str(vorlage[check_col][j]) != 'nan':
                            pflichtfeld_cbh.append(j+2)
                    elif check_sample == 1.0:
                        if check_cbh != 1.0:
                            pflichtfeld_cbh.append(j+2)
                if pflichtfeld_cbh != []:
                    error_str = return_error_rows_as_string(pflichtfeld_cbh)
                    text_pflicht = Label(frame, text = "Feldcode ist falsch in Spalte '{}' Zeile(n) {}".format(check_col, error_str),
                                            bg = '#eeeee4', font=('Ink free',11))
                    text_pflicht.pack()
                    # Zelle in Excel-Datei orange färben
                    zellen_bunt_malen(error_str, check_col, ws, (255, 127, 80))
                    error_count_total[0] += 1
                    error_count[0] += 1
            
            # Ausnahme bei lockReason
            try:
                check_col_richtig = cols_richtig['lockReason']
            except:
                check_col_richtig = ''
            if (check_col == 'lockReason') | (check_col == check_col_richtig):
                for j in range(len(vorlage)):
                    if str(vorlage[check_col][j]) == 'nan':
                        if (str(vorlage[sample_approval][j]) == '2') | (str(vorlage[sample_approval][j]) == '2.0') | (vorlage[sample_approval][j] == '02'):
                            pflichtfeld_lock.append(j+2)
                if pflichtfeld_lock != []:
                    error_str = return_error_rows_as_string(pflichtfeld_lock)
                    text_pflicht = Label(frame, text = "Pflichtfeld ist leer in Spalte '{}' Zeile(n) {}".format(check_col, error_str),
                                            bg = '#eeeee4', font=('Ink free',11))
                    text_pflicht.pack()
                    # Zelle in Excel-Datei orange färben
                    zellen_bunt_malen(error_str, check_col, ws, (255, 127, 80))
                    error_count_total[0] += 1
                    error_count[0] += 1
                    
            # Check bei container
            try:
                check_col_richtig = cols_richtig['container']
            except:
                check_col_richtig = ''
            if (check_col == 'container') | (check_col == check_col_richtig):
                for j in range(len(vorlage)): 
                    try:
                        check_container = float(vorlage[check_col][j])
                    except:
                        check_container = vorlage[check_col][j]
                    try:
                        check_cap = float(vorlage[fillcap][j])
                    except:
                        check_cap = vorlage[fillcap][j]
                    if (check_container == 99.0) & ((check_cap != 0.0) & (check_cap != '0,0')):
                        pflichtfeld_container.append(j+2)
                    if ((check_cap == 0.0) & (check_cap == '0,0')) & (check_container != 99.0):
                        pflichtfeld_container.append(j+2)
                        
                if pflichtfeld_container != []:
                    error_str = return_error_rows_as_string(pflichtfeld_container)
                    text_pflicht = Label(frame, text = "Lager 'container' oder Volumen 'fillcapacity' falsch Zeile(n) {}".format(error_str),
                                            bg = '#eeeee4', font=('Ink free',11))
                    text_pflicht.pack()
                    # Zelle in Excel-Datei orange färben
                    zellen_bunt_malen(error_str, check_col, ws, (220, 20, 60))
                    zellen_bunt_malen(error_str, fillcap, ws, (220, 20, 60))
                    error_count_total[0] += 1
                    error_count[0] += 1                    
                
                                
            if (check_col not in ausnahme_felder) & (check_col not in discharge_felder):
                # weiterer Check für 'Audit_Trail'
                try:
                    check_col_richtig = cols_richtig['Audit_Trail']
                except:
                    check_col_richtig = ''
                if (check_col == 'Audit_Trail') | (check_col == check_col_richtig):
                    for j in range(len(vorlage)):
                        if ';' in str(vorlage[check_col][j]):
                            audit_semikolon.append(j+2)
                        if not is_valid_format(vorlage[check_col][j][:15]):
                            audit_fail.append(j+2)
                        if ',' in str(vorlage[check_col][j]):
                            index_komma = vorlage[check_col][j].index(',')
                            if vorlage[check_col][j][index_komma + 1] != ' ':
                                audit_komma.append(j+2)                  
                            
                if audit_semikolon != []:
                    error_str = return_error_rows_as_string(audit_semikolon)
                    text_pflicht = Label(frame, text = "Format-Fehler in Spalte '{}' Zeile(n) {} (Semikolon)".format(check_col, error_str),
                                            bg = '#eeeee4', font=('Ink free',11))
                    text_pflicht.pack()
                    # Zelle in Excel-Datei rot färben
                    zellen_bunt_malen(error_str, check_col, ws, (220, 20, 60))
                    error_count_total[0] += 1
                    error_count[0] += 1
                    
                if audit_fail != []:
                    error_str = return_error_rows_as_string(audit_fail)
                    text_pflicht = Label(frame, text = "Format-Fehler in Spalte '{}' Zeile(n) {} (Datum/Kürzel)".format(check_col, error_str),
                                            bg = '#eeeee4', font=('Ink free',11))
                    text_pflicht.pack()
                    # Zelle in Excel-Datei rot färben
                    zellen_bunt_malen(error_str, check_col, ws, (220, 20, 60))
                    error_count_total[0] += 1
                    error_count[0] += 1
                    
                if audit_komma != []:
                    error_str = return_error_rows_as_string(audit_komma)
                    text_pflicht = Label(frame, text = "Format-Fehler in Spalte '{}' Zeile(n) {} (Komma)".format(check_col, error_str),
                                            bg = '#eeeee4', font=('Ink free',11))
                    text_pflicht.pack()
                    # Zelle in Excel-Datei rot färben
                    zellen_bunt_malen(error_str, check_col, ws, (220, 20, 60))
                    error_count_total[0] += 1
                    error_count[0] += 1   
                    
                # Standardvorgang:
                for j in range(len(vorlage)):
                    if str(vorlage[check_col][j]) == 'nan':
                        pflichtfeld_row.append(j+2)                        
                if pflichtfeld_row != []:
                    error_str = return_error_rows_as_string(pflichtfeld_row)
                    text_pflicht = Label(frame, text = "Pflichtfeld ist leer in Spalte '{}' Zeile(n) {}".format(check_col, error_str),
                                            bg = '#eeeee4', font=('Ink free',11))
                    text_pflicht.pack()
                    # Zelle in Excel-Datei orange färben
                    zellen_bunt_malen(error_str, check_col, ws, (255, 127, 80))
                    error_count_total[0] += 1
                    error_count[0] += 1
        
        # Feld, das gefüllt sein sollte, aber kein Pflichtfeld ist
        elif ('soll immer gefüllt sein, aber kein Pflichtfeld' in felder_f['Pflichtfeld'][i]) & ('abhängig' not in felder_f['Pflichtfeld'][i]):
            check_col = felder_f['Feldcode'][i]
            if check_col in cols_richtig:
                check_col = cols_richtig[check_col]

            # weiterer Check bei Period_of_blood_draw
            if check_col == 'Period_of_blood_draw':
                for j in range(len(vorlage)):
                    if str(vorlage[check_col][j]) == 'nan':
                        kein_pflichtfeld_row.append(j+2)
                    elif bool(re.compile(r'^\d{4}-\d{2}$').match(vorlage[check_col][j])):
                        try:
                            date_object = datetime.strptime(vorlage[check_col][j], "%Y-%m").date()
                        except:
                            pflichtfeld_period.append(j+2)
                    elif (vorlage[check_col][j] != 'not specified'):
                        pflichtfeld_period.append(j+2)
                        
                if kein_pflichtfeld_row != []:
                    error_str = return_error_rows_as_string(kein_pflichtfeld_row)
                    text_pflicht = Label(frame, text = "Feld ist leer in Spalte '{}' Zeile(n) {} (kein Pflichtfeld)".format(check_col, error_str),
                                            bg = '#eeeee4', font=('Ink free',11))
                    text_pflicht.pack()
                    # Zelle in Excel-Datei grau färben
                    zellen_bunt_malen(error_str, check_col, ws, (168, 168, 168))
                    
                if pflichtfeld_period != []:
                    error_str = return_error_rows_as_string(pflichtfeld_period)
                    text_pflicht = Label(frame, text = "Fehler in Spalte '{}' Zeile(n) {}".format(check_col, error_str),
                                            bg = '#eeeee4', font=('Ink free',11))
                    text_pflicht.pack()
                    # Zelle in Excel-Datei rot färben
                    zellen_bunt_malen(error_str, check_col, ws, (220, 20, 60))                    
                
            else:
                for j in range(len(vorlage)):
                    if str(vorlage[check_col][j]) == 'nan':
                        kein_pflichtfeld_row.append(j+2)
                if kein_pflichtfeld_row != []:
                    error_str = return_error_rows_as_string(kein_pflichtfeld_row)
                    text_pflicht = Label(frame, text = "Feld ist leer in Spalte '{}' Zeile(n) {} (kein Pflichtfeld)".format(check_col, error_str),
                                            bg = '#eeeee4', font=('Ink free',11))
                    text_pflicht.pack()
                    # Zelle in Excel-Datei grau färben
                    zellen_bunt_malen(error_str, check_col, ws, (168, 168, 168))
            
        # Felder, die abhängig von anderen Feldern gefüllt sein sollen
        elif ('soll abhängig von' in felder_f['Pflichtfeld'][i]) & ('aber kein Pflichtfeld' in felder_f['Pflichtfeld'][i]):
            # Remarks_on_diagnosis kann auch leer sein
            if felder_f['Feldcode'][i] != 'Remarks_on_diagnosis':
                dependent_field = felder_f['Pflichtfeld'][i].split('soll abhängig von')[-1].split('gefüllt')[0].strip()
                # Name von dependent_field bearbeiten, weil der manchmal anders geschrieben wird
                if ' ' in dependent_field:
                    dep_field_words = dependent_field.split(' ')
                    for col in vorlage.columns:
                        if all(string in col for string in dep_field_words):
                            dependent_field = col                        
                check_col = felder_f['Feldcode'][i]
                if check_col in cols_richtig:
                    check_col = cols_richtig[check_col]
                # Klein- und Großschreibung von dependent field könnte anders sein
                cols_klein = [col.lower() for col in cols]
                for j in range(len(vorlage)):
                    # prüfen, ob dependent field existiert
                    if dependent_field in cols:
                        if (str(vorlage[check_col][j]) == 'nan') & (str(vorlage[dependent_field][j]) != 'nan'):
                            pflichtfeld_row.append(j+2)
                    elif dependent_field in cols_klein:
                        col_index = cols_klein.index(dependent_field)
                        if (str(vorlage[check_col][j]) == 'nan') & (str(vorlage[cols[col_index]][j]) != 'nan'):
                            pflichtfeld_row.append(j+2)
                if pflichtfeld_row != []:
                    error_str = return_error_rows_as_string(pflichtfeld_row)
                    text_pflicht = Label(frame, text = "Feld ist leer in Spalte '{}' Zeile(n) {}".format(check_col, error_str),
                                            bg = '#eeeee4', font=('Ink free',11))
                    text_pflicht.pack()
                    # Zelle in Excel-Datei orange färben
                    zellen_bunt_malen(error_str, check_col, ws, (255, 127, 80))
                    error_count_total[0] += 1
                    error_count[0] += 1


reiter = {
    'Patient': 'Patients',
    'Diagnosis': 'Diagnosis',
    'Therapie': 'Therapy',
    'Parameter': 'Laboratory', 
    'Medication': 'Medication',
    'Audit': 'Audit Trail',
    'Specs': 'Specs'
}
reiter_namen = list(reiter.keys())

        
# Check 11: alle Pflichtfelder vorhanden?
def check_pflichtfeld_fehlend(file_path, cols, felder, ws, error_count, error_count_total):
    # Reiter identifizieren (aus Dateiname)
    fehlende_cols = []
    
    # alle Pflichtfelder zu den entsprechenden Reitern ablesen
    felder_pflicht = felder[felder['Pflichtfeld'].isin(['Pflichtfeld', 'Pflichtfeld?', 'leeres Pflichtfeld', 'leeres Pflichtfeld?'])]
    reiter_namen = list(reiter.keys())
    for i in range(len(reiter_namen)):
        if reiter_namen[i] in file_path:
            reiter_name = reiter_namen[i]
            felder_reiter = felder_pflicht[felder_pflicht['Reiter']==reiter[reiter_name]]
            reiter_cols = list(felder_reiter['Feldcode'])
            
            for col in reiter_cols:
                if (col not in cols) & (col not in cols_richtig):
                    fehlende_cols.append(col)
            break
                    
    # Ausnahme bei Reservierung
    if 'Reservierung' in file_path:
        reiter_cols = ['reservedCapacity', 'isReserved', 'reservationDate', 'reservedUntil', 'reservedFor']
        for col in reiter_cols:
            if (col not in cols) & (col not in cols_richtig):
                fehlende_cols.append(col)
                
    # Ausnahme bei 'Sample & Aliquot'
    elif 'Sample & Aliquot' in file_path:
        reiter_cols = ['master_id', 'sample_id', 'comments', 'amountOfAliquotes']
        for col in reiter_cols:
            if (col not in cols) & (col not in cols_richtig):
                fehlende_cols.append(col)
                
    if fehlende_cols != []:
        fehlende_cols_str = ', '.join(fehlende_cols)
        text_pflicht = Label(frame, text = "Spalte(n) '{}' fehlt/fehlen".format(fehlende_cols_str),
                                bg = '#eeeee4', font=('Ink free',11))
        text_pflicht.pack()
        error_count_total[0] += 1
        error_count[0] += 1
        
        
# Check 12: 2 Nachkommastellen und weitere Formatierungen
# nur bei csv-Dateien kann auf Nachkommastellen geprüft werden
def check_komma_stellen(cols, felder_f, ws, error_count, error_count_total):
    komma_felder = ['Measured_value', 'Reference_cut_off']
    besondere_zeichen = ['<', '>']
    
    def check_nachkommastellen(entry):
        if ',' in entry:
            if '-' not in entry:
                check_entry = entry.split(',')[-1]
                if len(check_entry) != 2:
                    if j+2 not in komma_fehler:
                        komma_fehler.append(j+2)
            else:
                check_number = entry
                if '(' in check_number:
                    check_number = check_number.split('(')[0].strip()
                erste_zahl = check_number.split('-')[0].strip()
                zweite_zahl = check_number.split('-')[-1].strip()
                try:
                    if (len(erste_zahl.split(',')[-1]) != 2) | (len(zweite_zahl.split(',')[-1]) != 2):
                        if j+2 not in komma_fehler:
                            komma_fehler.append(j+2)
                except:
                    if j+2 not in komma_fehler:
                        komma_fehler.append(j+2)
                        
    def check_doppelpunkt(entry):
        if ':' in entry:
            erste_zahl = entry.split(':')[0]
            zweite_zahl = entry.split(':')[-1]
            if erste_zahl[-1] == ' ':
                if j+2 not in komma_fehler:
                    komma_fehler.append(j+2)
            if zweite_zahl[0] != ' ':
                if j+2 not in komma_fehler:
                    komma_fehler.append(j+2)
            if zweite_zahl[1] == ' ':
                if j+2 not in komma_fehler:
                    komma_fehler.append(j+2)
                    
    def check_bindestrich(entry):
        if '-' in entry:
            erste_zahl = entry.split('-')[0]
            zweite_zahl = entry.split('-')[-1]
            if erste_zahl[-1] != ' ':
                if j+2 not in komma_fehler:
                    komma_fehler.append(j+2)
            if erste_zahl[-2] == ' ':
                if j+2 not in komma_fehler:
                    komma_fehler.append(j+2)
            if zweite_zahl[0] != ' ':
                if j+2 not in komma_fehler:
                    komma_fehler.append(j+2)
            if zweite_zahl[1] == ' ':
                if j+2 not in komma_fehler:
                    komma_fehler.append(j+2)
                                    
    
    for col in komma_felder:
        if col in cols_richtig:
            komma_felder.append(cols_richtig[col])
    
    
    for check_col in cols:
        komma_fehler = []
        if check_col in komma_felder:
            for j in range(len(vorlage)):
                # Überprüfen auf 2 Nachkommastellen
                # bei csv-Dateien
                if csv_flag:
                    # Überprüfen bei Einträgen mit Komma als Dezimalzeichen
                    check_nachkommastellen(str(vorlage[check_col][j]))
                    
                    # Punkt als Dezimalzeichen ist falsch
                    if (':' not in str(vorlage[check_col][j])) & ((',' not in str(vorlage[check_col][j])) | ('.' in str(vorlage[check_col][j]))):
                        if ('positive' not in vorlage[check_col][j]) & ('negative' not in vorlage[check_col][j]) & ('not specified' not in vorlage[check_col][j]):
                            if j+2 not in komma_fehler:
                                komma_fehler.append(j+2)
                    
                    # Überprüfen, ob nach < oder > Leerzeichen folgt
                    for zeichen in besondere_zeichen:
                        if zeichen in str(vorlage[check_col][j]):
                            if str(vorlage[check_col][j]).split(zeichen)[-1][0] != ' ':
                                if j+2 not in komma_fehler:
                                    komma_fehler.append(j+2)
                                    
                    # Doppelpunkt-Formatierung
                    check_doppelpunkt(str(vorlage[check_col][j]))
                    
                    # Bindestrich-Formatierung
                    check_bindestrich(str(vorlage[check_col][j]))
                        
                # Überprüfung bei Excel-Tabellen
                # Aufgrund des Einlesens mittels pandas gehen alle 0-Nachkommastellen verloren
                # Hier kommt nur eine allgemeine Fehlermeldung 
                else:
                    # in der Output-Dateu werden die Daten richtig angezeigt
                    try:
                        value = float(vorlage[check_col][j])
                        value = "{:.2f}".format(value).replace('.', ',')
                        vorlage[check_col][j] = value
                    except:
                        pass                    
                                
            if komma_fehler != []:
                error_str = return_error_rows_as_string(komma_fehler)
                text_komma_fehler = Label(frame, text = "Formatierungsfehler in Spalte '{}' Zeile(n) {}".format(check_col, error_str),
                                        bg = '#eeeee4', font=('Ink free',11))
                text_komma_fehler.pack()
                # Zelle in Excel-Datei rot färben
                zellen_bunt_malen(error_str, check_col, ws, (220, 20, 60))
                error_count_total[0] += 1
                error_count[0] += 1
        
                


# In[22]:


def go_dodo():
    """ Überprüfen aller Excel-Dateien im ausgewählten Ordner
    """
    
    # Ladebalken
    progress_bar = ttk.Progressbar(root, orient="horizontal", length=250, mode="determinate")
    progress_bar.grid(row=4, column=1, columnspan=5)    
    
    global error_title, root_m, vorlage, canvas, frame, error_title2, felder_f, csv_flag, para_check_text
    felder_f = pd.DataFrame()
    try:
        error_title.grid_forget()
    except:
        pass
    try:
        error_title2.grid_forget()
    except:
        pass
    try:
        root_m.grid_forget()
    except:
        pass
    try:
        para_check_text.grid_forget()
    except:
        pass

    
    error_count_total = [0] # weil Listen im Gegensatz zu Integers mutable sind
    
    
    folder_path = filedialog.askdirectory(title = "Ordner auswählen")
    file_list = os.listdir(folder_path)
    file_list = [file for file in file_list if '$' not in file]
    
    progress_bar.start()
    
    ### Interface ###    
    root_m = Frame(root, width=900, height=200, bg = '#eeeee4', highlightbackground='#869287',
                   highlightthickness=2)
    root_m.grid(row=7, column=1, columnspan=100)
    # Create a Canvas to hold the Frame and Scrollbars
    canvas = Canvas(root_m, width=800, bg = '#eeeee4')
    canvas.pack(side="left", fill="both", expand=True)

    # Create vertical scrollbar and associate it with the Canvas
    v_scrollbar = Scrollbar(root_m, command=canvas.yview)
    v_scrollbar.pack(side="right", fill="y")
    canvas.configure(yscrollcommand=v_scrollbar.set)

    # Create horizontal scrollbar and associate it with the Canvas
    h_scrollbar = Scrollbar(root_m, command=canvas.xview, orient="horizontal")
    h_scrollbar.pack(side="bottom", fill="x")
    canvas.configure(xscrollcommand=h_scrollbar.set)

    # Create a Frame to hold the label widgets
    frame = Frame(canvas, bg='#eeeee4')
    canvas.create_window((0, 0), window=frame, anchor="nw")

    # Einlesen der Excel-Tabelle 'Aufbau und Felder'
    felder_origin = open_dm_file('IDA_Aufbau und Felder', 'O:\Datenmanagement\IDA_in.vent Datenbank', 'Felderbezeichnung')
    felder_origin.columns = felder_origin.iloc[6, :]
    felder_origin = felder_origin.drop(felder_origin.index[:7]).reset_index().drop('index', axis=1)
    felder_origin['Feldcode'] = felder_origin['Feldcode'].map(lambda x: x.strip())
    
    # Info, dass Parameter-Tabellen selbständig überprüft werden müssen
    para_flag = False
    for file in file_list:
        if 'Parameter' in file:
            para_flag = True
    if para_flag:
        para_check_text = Label(root, text = 'Achtung: Überprüfe in der Parameter-Tabelle nochmal \n die Nachkommastellen der Messwerte und des Referenzbereichs',
                               bg = '#eeeee4', font=('Ink free',11,'bold'), fg = '#ff5349')
        para_check_text.grid(row=6, column=4, columnspan=30)

    
    csv_flag = False
    
    error_notifications = []
    
    for i in range(len(file_list)):
        
        error_count = [0]
        if 'csv' in file_list[i]:
            vorlage = pd.read_csv(folder_path+'/'+file_list[i], encoding='latin_1', delimiter=';')
            file_name = file_list[i].split('.csv')[0]
            csv_flag = True
        elif 'xlsx' in file_list[i]:
            vorlage = open_closed_or_opened_file(folder_path+'/'+file_list[i])
            file_name = file_list[i].split('.xlsx')[0]
        cols = list(vorlage.columns)

        
        vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
        wb = xw.Book('{}_Check.xlsx'.format(file_name))
        ws = wb.sheets['Sheet1']
        
        filename_text = Label(frame, text = "{}".format(file_name),
                                   bg = '#eeeee4', font=('Ink free',11,'bold'))
        filename_text.pack()

        felder = felder_origin.copy()
        
        # Filtern nach Pflichtfeld:
        # Zeilen rausnehmen, wenn Feld 'nie genutzt' oder 'derzeit nicht genutzt' wird
        rows = felder.loc[felder['Pflichtfeld'].isin(['nie genutzt', 'derzeit nicht genutzt'])]
        rows = rows.index.tolist()
        felder = felder.drop(rows)
        felder = felder.reset_index().drop('index', axis=1)
        
        # Spalten der hochgeladenen Datei mit den Feldern abgleichen und ein Subset erstellen
        index_felder = []
        for col in cols:
            for i in range(len(felder)):
                if col == felder['Feldcode'][i]:
                    index_felder.append(i)
        felder_f = felder.iloc[index_felder].reset_index().drop('index', axis=1)

        # Feldcode-Namen überprüfen
        check_feldcode(cols, felder, felder_origin, ws, error_count, error_count_total)        
        
        # Format von pat_id, master_id und sample_id überprüfen       
        # Spalte master_id
        if ('master_id' in vorlage.columns) | ('master_id' in cols_richtig):
            check_masterid(vorlage, ws, error_count, error_count_total)

        # Spalte sample_id
        cols = list(vorlage.columns)
        cols_l = [col.lower() for col in cols]
        if ('sample_id' in vorlage.columns) | ('sample_id' in cols_richtig):
            check_sampleid(cols, cols_l, vorlage, ws, error_count, error_count_total)

        # Spalte pat_id
        if ('pat_id' in vorlage.columns) | ('pat_id' in cols_richtig):
            check_patid(vorlage, ws, error_count, error_count_total)

        # Check, ob master_id und sample_id identisch sind
        if (('sample_id' in vorlage.columns) | ('sample_id' in cols_richtig)) & (('master_id' in vorlage.columns) | ('master_id' in cols_richtig)):
            check_sample_master(vorlage, ws, error_count, error_count_total)
            
        # Überprüfen der Auswahlmöglichkeiten, wenn Feldart monovalent list ist
        check_monovalent(cols, felder_f, ws, error_count, error_count_total)
        
        # Datum-Spalte überprüfen (dd.mm.yyyy)
        check_datum(felder_f, vorlage, ws, error_count, error_count_total)

        # 'One line text'-Felder überprüfen, sodass es eine Warnung gibt,
        # wenn Text ab dem 120. Zeichen abgeschnitten wird
        check_one_line_text(felder_f, vorlage, ws, error_count, error_count_total)

        # Katalog-Felder überprüfen
        check_katalog(felder_f, ws, error_count, error_count_total)

        # Pflichtfelder füllen
        check_pflichtfeld(cols, felder_f, ws, error_count, error_count_total)
        
        # Überprüfen, ob Pflichtfelder vorhanden sind
        check_pflichtfeld_fehlend(file_name, cols, felder, ws, error_count, error_count_total)
        
        # Überprüfen, ob Formatierungen bei Zahlen stimmen
        check_komma_stellen(cols, felder_f, ws, error_count, error_count_total)
                    
        error_file = Label(frame, text = 'Anzahl Errors für {}: '.format(file_name) + str(error_count[0]), bg = '#eeeee4', font=('Ink free',9,'bold'))
        error_file.pack()
        error_notifications.append(error_file)
        error_notifications.append(filename_text)
        
        ##### User Interface #####
        error_title = Label(root, text = '\nError-Meldungen: ' + str(error_count_total[0]), bg = '#eeeee4', font=('Ink free',11,'bold'))
        error_title.grid(row=5, column=7, columnspan=15)
        ##########################

        wb.save()
        wb.close()

        step = (100/(len(file_list)+2))
        progress_bar.step(step)
        root.update()

        
    # Feedback, wenn Dateien alle keine Fehler haben
    if (error_count_total == [0]) & (para_flag == False):
        for message in error_notifications:
            message.pack_forget()
            
        error_file = Label(frame, text = 'Alles tip top :)', bg = '#eeeee4', font=('Ink free',15,'bold'))
        error_file.pack()    
        
        
    # Calculate the desired frame width based on the canvas width
    frame_width = canvas.winfo_reqwidth()  # Use canvas width
    # Set the width of the scrollable_frame and prevent resizing
    frame.grid_propagate(False)  # Prevent resizing
    frame.config(width=frame_width)  # Set the width
    # Configure Scrollbar to control scrolling
    frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))
        
    progress_bar.stop()
    progress_bar.grid_forget()
    
        
def go_dodo_solo():
    """ Überprüfen einer einzelnen Excel-Datei
    """
    
    # Ladebalken
    progress_bar_solo = ttk.Progressbar(root, orient="horizontal", length=250, mode="determinate")
    progress_bar_solo.grid(row=4, column=1, columnspan=5)
    
    global error_title, root_m, vorlage, canvas, frame, error_title2, felder_f, csv_flag, para_check_text
    felder_f = pd.DataFrame()
    try:
        error_title.grid_forget()
    except:
        pass
    try:
        error_title2.grid_forget()
    except:
        pass
    try:
        root_m.grid_forget()
    except:
        pass
    try:
        para_check_text.grid_forget()
    except:
        pass

        
    error_count = 0
    
    
    vorlage_path = filedialog.askopenfilename(title = "Datei auswählen")
    
    progress_bar_solo.start(10)
    
    # Info, dass Parameter-Tabellen selbständig überprüft werden müssen
    if 'Parameter' in vorlage_path:
        para_check_text = Label(root, text = 'Achtung: Überprüfe in der Parameter-Tabelle nochmal \n die Nachkommastellen der Messwerte und des Referenzbereichs',
                               bg = '#eeeee4', font=('Ink free',11,'bold'), fg = '#ff5349')
        para_check_text.grid(row=6, column=4, columnspan=30)
    
    csv_flag = False
    if vorlage_path not in ['', ' ']:
        if 'csv' in vorlage_path:
            vorlage = pd.read_csv(vorlage_path, encoding='latin_1', delimiter=';')
            csv_flag = True
        elif 'xlsx' in vorlage_path:
            vorlage = open_closed_or_opened_file(vorlage_path)
        cols = list(vorlage.columns)
        # neue Datei xxx_Check.xlsx erstellen, um Farbe von fehlerhaften Zellen zu ändern
        file_name = vorlage_path.split('/')[-1].split('.csv')[0]        

        progress_bar_solo.step(5)
        root.update()

            
        vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
        wb = xw.Book('{}_Check.xlsx'.format(file_name))
        ws = wb.sheets['Sheet1']

        progress_bar_solo.step(20)
        root.update()
        ### Interface ###    
        root_m = Frame(root, width=900, height=200, bg = '#eeeee4', highlightbackground='#869287',
                       highlightthickness=2)
        root_m.grid(row=7, column=1, columnspan=100)

        # Create a Canvas to hold the Frame and Scrollbars
        canvas = Canvas(root_m, width=800, bg = '#eeeee4')
        canvas.pack(side="left", fill="both", expand=True)

        # Create vertical scrollbar and associate it with the Canvas
        v_scrollbar = Scrollbar(root_m, command=canvas.yview)
        v_scrollbar.pack(side="right", fill="y")
        canvas.configure(yscrollcommand=v_scrollbar.set)

        # Create horizontal scrollbar and associate it with the Canvas
        h_scrollbar = Scrollbar(root_m, command=canvas.xview, orient="horizontal")
        h_scrollbar.pack(side="bottom", fill="x")
        canvas.configure(xscrollcommand=h_scrollbar.set)

        # Create a Frame to hold the label widgets
        frame = Frame(canvas, bg='#eeeee4')
        canvas.create_window((0, 0), window=frame, anchor="nw")

        # Einlesen der Excel-Tabelle 'Aufbau und Felder'
        felder_origin = open_dm_file('IDA_Aufbau und Felder', 'O:\Datenmanagement\IDA_in.vent Datenbank', 'Felderbezeichnung')
        felder_origin.columns = felder_origin.iloc[6, :]
        felder_origin = felder_origin.drop(felder_origin.index[:7]).reset_index().drop('index', axis=1)
        felder_origin['Feldcode'] = felder_origin['Feldcode'].map(lambda x: x.strip())
        
        felder = felder_origin.copy()
        # Filtern nach Pflichtfeld:
        # Zeilen rausnehmen, wenn Feld 'nie genutzt' oder 'derzeit nicht genutzt' wird
        rows = felder.loc[felder['Pflichtfeld'].isin(['nie genutzt', 'derzeit nicht genutzt'])]
        rows = rows.index.tolist()
        felder = felder.drop(rows)
        felder = felder.reset_index().drop('index', axis=1)

        # Spalten der hochgeladenen Datei mit den Feldern abgleichen und ein Subset erstellen
        index_felder = []
        for col in cols:
            for i in range(len(felder)):
                if col == felder['Feldcode'][i]:
                    index_felder.append(i)
        felder_f = felder.iloc[index_felder].reset_index().drop('index', axis=1)

        error_count = [0]
        error_count_total = [0]    
        
        # Feldcode-Namen überprüfen
        check_feldcode(cols, felder, felder_origin, ws, error_count, error_count_total)
        
        # Format von pat_id, master_id und sample_id überprüfen       
        # Spalte master_id
        if ('master_id' in vorlage.columns) | ('master_id' in cols_richtig):
            check_masterid(vorlage, ws, error_count, error_count_total)

        # Spalte sample_id
        cols = list(vorlage.columns)
        cols_l = [col.lower() for col in cols]
        if ('sample_id' in vorlage.columns) | ('sample_id' in cols_richtig):
            check_sampleid(cols, cols_l, vorlage, ws, error_count, error_count_total)

        # Spalte pat_id
        if ('pat_id' in vorlage.columns) | ('pat_id' in cols_richtig):
            check_patid(vorlage, ws, error_count, error_count_total)

        # Check, ob master_id und sample_id identisch sind
        if (('sample_id' in vorlage.columns) | ('sample_id' in cols_richtig)) & (('master_id' in vorlage.columns) | ('master_id' in cols_richtig)):
            check_sample_master(vorlage, ws, error_count, error_count_total)

        # Überprüfen der Auswahlmöglichkeiten, wenn Feldart monovalent list ist
        check_monovalent(cols, felder_f, ws, error_count, error_count_total)
        progress_bar_solo.step(5)
        root.update()

        # Datum-Spalte überprüfen (dd.mm.yyyy)
        check_datum(felder_f, vorlage, ws, error_count, error_count_total)
        progress_bar_solo.step(5)
        root.update()

        # 'One line text'-Felder überprüfen, sodass es eine Warnung gibt,
        # wenn Text ab dem 120. Zeichen abgeschnitten wird
        check_one_line_text(felder_f, vorlage, ws, error_count, error_count_total)
        progress_bar_solo.step(15)
        root.update()

        # Katalog-Felder überprüfen
        check_katalog(felder_f, ws, error_count, error_count_total)
        progress_bar_solo.step(5)
        root.update()

        # Pflichtfelder füllen
        check_pflichtfeld(cols, felder_f, ws, error_count, error_count_total)
        progress_bar_solo.step(5)
        root.update()
        
        # Überprüfen, ob Pflichtfelder vorhanden sind
        check_pflichtfeld_fehlend(vorlage_path, cols, felder, ws, error_count, error_count_total)
        progress_bar_solo.step(15)
        root.update()
        
        # Überprüfen, ob Formatierungen bei Zahlen stimmen
        check_komma_stellen(cols, felder_f, ws, error_count, error_count_total)
        progress_bar_solo.step(5)
        root.update()  
        
        # Falls keine Fehler vorhanden sind, Nachrich im GUI
        if (error_count_total == [0]) & ('Parameter' not in vorlage_path):
            null_fehler = Label(frame, text = "Alles tip top :)",
                                    bg = '#eeeee4', font=('Ink free',15))
            null_fehler.pack()

        ##### User Interface #####
        error_title2 = Label(root, text = '\nError-Meldungen: ' + str(error_count[0]), bg = '#eeeee4', font=('Ink free',11,'bold'))
        error_title2.grid(row=5, column=7, columnspan=15)
        ##########################

        wb.save()

        # Calculate the desired frame width based on the canvas width
        frame_width = canvas.winfo_reqwidth()  # Use canvas width
        # Set the width of the scrollable_frame and prevent resizing
        frame.grid_propagate(False)  # Prevent resizing
        frame.config(width=frame_width)  # Set the width
        # Configure Scrollbar to control scrolling
        frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))
        
    progress_bar_solo.stop()
    progress_bar_solo.grid_forget()


# In[23]:


# User Interface
# Basis
root = Tk()
root.title('Import-Check')
root.geometry('1000x550')
root.iconbitmap('O:\Forschung & Entwicklung\Allgemein\Vorlagen\Abbildungen\Dodo\dodo_icon.ico')
root.config(bg='#eeeee4')

# alle Buttons auf Startseite
hi = Label(root, text = '\nÜberprüfen der Dateien vor dem Import in IDA                        \n',
           bg = '#eeeee4', font=('Ink free',15,'bold'))
hi.grid(row=1, column=1, columnspan=30)

browse_text = Label(root, text = 'Bitte wähle den Ordner aus  ', bg = '#eeeee4', font=('Ink free',12))
browse_text.grid(row=2, column=1, columnspan=5)
browse_button = Button(root, text='Browse', font=('Ink free',10,'bold'), bg='#869287',
                       command = go_dodo)
browse_button.grid(row=2, column=6)

browse_solo_text = Label(root, text = 'Oder wähle eine einzelne Datei aus ', bg = '#eeeee4', font=('Ink free',12))
browse_solo_text.grid(row=3, column=1, columnspan=5)
browse_solo_button = Button(root, text='Browse', font=('Ink free',10,'bold'), bg='#869287',
                       command = go_dodo_solo)
browse_solo_button.grid(row=3, column=6)

# Logo
from PIL import ImageTk, Image
frame_logo = Frame(root, width=1, height=1)
frame_logo.grid(row=1, column=0)

img = ImageTk.PhotoImage(Image.open("O:\Forschung & Entwicklung\Allgemein\Vorlagen\Abbildungen\Dodo\dodo-dancing_ohne Hintergrund_ohne Schatten.png").resize((70,70)), master = root)
label = Label(frame_logo, image = img, bg = '#eeeee4')
label.pack()

root.mainloop()


# In[ ]:




