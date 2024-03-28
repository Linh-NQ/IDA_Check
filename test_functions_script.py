import xlwings as xw
import pandas as pd
import glob
import tempfile
import os
import pytest
from IDA_Check import frame
from IDA_Check import open_closed_or_opened_file, open_dm_file, return_error_rows_as_string, zellen_bunt_malen, check_feldcode, check_masterid, check_sampleid, check_patid, check_sample_master, check_monovalent, check_datum, check_one_line_text, check_katalog, check_pflichtfeld, check_pflichtfeld_fehlend, check_komma_stellen


# Unit Test für return_error_rows_as_string
def test_return_error_rows_as_string():
    assert return_error_rows_as_string([1]) == '1'
    assert return_error_rows_as_string([1, 3]) == '1, 3'
    assert return_error_rows_as_string([1, 2, 3]) == '1-3'
    assert return_error_rows_as_string([1, 2, 3, 5, 6, 7]) == '1-3, 5-7'
    assert return_error_rows_as_string([1, 2, 3, 5, 6, 7, 9]) == '1-3, 5-7, 9'
    assert return_error_rows_as_string([1, 2, 3, 5, 7, 8]) == '1-3, 5, 7-8'
    assert return_error_rows_as_string([101, 103, 104, 105, 110]) == '101, 103-105, 110'
    

# Helper function to check cell color
def check_cell_color(ws, column_name, row, expected_color):
    global file_name
    column_index = ord(column_name) - ord('A') + 1  # Convert column letter to index
    actual_color = ws.cells(row, column_index).color
    return actual_color == expected_color, f"Cell color mismatch at {column_name}{row}. Expected: {expected_color}, Actual: {actual_color}"

# Unit Test für zellen_bunt_malen
def test_zellen_bunt_malen():
    global file_name
    vorlage_path = 'excel_unit_tests/test1.xlsx'
    vorlage = open_closed_or_opened_file(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']

    # Rot
    zellen_bunt_malen('1', 'master_id', vorlage, ws, (220, 20, 60))
    assert check_cell_color(ws, 'A', 1, (220, 20, 60))

    # Grau
    zellen_bunt_malen('2, 4', 'pat_id', vorlage, ws, (168, 168, 168))
    assert check_cell_color(ws, 'B', 2, (168, 168, 168))
    assert check_cell_color(ws, 'B', 4, (168, 168, 168))

    # Gelb
    zellen_bunt_malen('1-3, 5, 7-8', 'Audit_Trail', vorlage, ws, (238, 232, 170))
    assert check_cell_color(ws, 'C', 1, (238, 232, 170))
    assert check_cell_color(ws, 'C', 2, (238, 232, 170))
    assert check_cell_color(ws, 'C', 3, (238, 232, 170))
    assert check_cell_color(ws, 'C', 5, (238, 232, 170))
    assert check_cell_color(ws, 'C', 7, (238, 232, 170))
    assert check_cell_color(ws, 'C', 8, (238, 232, 170))

    wb.close()
    

# Unit Test für check_feldcode
def test_check_feldcode():
    global felder_f
    vorlage_path = r'excel_unit_tests/test_feldcode.xlsx'
    vorlage = open_closed_or_opened_file(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)
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

    error_count = [0]
    error_count_total = [0]

    check_feldcode(vorlage, cols, felder, felder_origin, ws, error_count, error_count_total)
    assert check_cell_color(ws, 'K', 1, (220, 20, 60))
    assert check_cell_color(ws, 'M', 1, (220, 20, 60))
    assert check_cell_color(ws, 'N', 1, (220, 20, 60))
    assert check_cell_color(ws, 'O', 1, (220, 20, 60))

    wb.close()



# Unit Test für check_masterid
def test_check_masterid():
    vorlage_path = r'excel_unit_tests/test_masterid.xlsx'
    vorlage = open_closed_or_opened_file(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)
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

    error_count = [0]
    error_count_total = [0]

    check_masterid(vorlage, ws, error_count, error_count_total)
    assert check_cell_color(ws, 'A', 3, (220, 20, 60))
    assert check_cell_color(ws, 'A', 4, (220, 20, 60))
    assert check_cell_color(ws, 'A', 5, (220, 20, 60))
    assert check_cell_color(ws, 'A', 6, (220, 20, 60))
    assert check_cell_color(ws, 'A', 7, (220, 20, 60))
    assert check_cell_color(ws, 'A', 8, (220, 20, 60))
    assert check_cell_color(ws, 'A', 12, (220, 20, 60))
    assert check_cell_color(ws, 'A', 13, (220, 20, 60))
    assert check_cell_color(ws, 'A', 15, (220, 20, 60))

    wb.close()



# Unit Test für check_sampleid
def test_check_sampleid():
    vorlage_path = r'excel_unit_tests/test_sampleid.xlsx'
    vorlage = open_closed_or_opened_file(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)
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

    error_count = [0]
    error_count_total = [0]

    check_sampleid(cols, vorlage, ws, error_count, error_count_total)
    assert check_cell_color(ws, 'A', 2, (220, 20, 60))
    assert check_cell_color(ws, 'A', 4, (220, 20, 60))
    assert check_cell_color(ws, 'A', 5, (220, 20, 60))
    assert check_cell_color(ws, 'A', 6, (220, 20, 60))
    assert check_cell_color(ws, 'A', 7, (220, 20, 60))
    assert check_cell_color(ws, 'A', 10, (220, 20, 60))
    assert check_cell_color(ws, 'A', 11, (220, 20, 60))
    assert check_cell_color(ws, 'A', 12, (220, 20, 60))

    wb.close()



# Unit Test für check_patid
def test_check_patid():
    vorlage_path = r'excel_unit_tests/test_patid.xlsx'
    vorlage = open_closed_or_opened_file(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)
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

    error_count = [0]
    error_count_total = [0]

    check_patid(vorlage, ws, error_count, error_count_total)
    assert check_cell_color(ws, 'A', 3, (220, 20, 60))
    assert check_cell_color(ws, 'A', 5, (220, 20, 60))
    assert check_cell_color(ws, 'A', 6, (220, 20, 60))
    assert check_cell_color(ws, 'A', 7, (220, 20, 60))

    wb.close()



# Unit Test für check_sample_master
def test_check_sample_master():
    vorlage_path = r'excel_unit_tests/test_sample_master.xlsx'
    vorlage = open_closed_or_opened_file(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)
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

    error_count = [0]
    error_count_total = [0]

    check_sample_master(vorlage, ws, error_count, error_count_total)
    assert check_cell_color(ws, 'A', 3, (220, 20, 60))
    assert check_cell_color(ws, 'A', 8, (220, 20, 60))
    assert check_cell_color(ws, 'A', 11, (220, 20, 60))
    assert check_cell_color(ws, 'B', 4, (220, 20, 60))
    assert check_cell_color(ws, 'B', 5, (220, 20, 60))
    assert check_cell_color(ws, 'B', 7, (220, 20, 60))
    assert check_cell_color(ws, 'B', 9, (220, 20, 60))

    wb.close()



# Unit Test für check_movovalent
def test_check_monovalent():
    vorlage_path = r'excel_unit_tests/test_monovalent.xlsx'
    vorlage = open_closed_or_opened_file(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)
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

    index_felder = []
    for col in cols:
        for i in range(len(felder)):
            if col == felder['Feldcode'][i]:
                index_felder.append(i)

    felder_f = felder.iloc[index_felder].reset_index().drop('index', axis=1)                

    error_count = [0]
    error_count_total = [0]

    check_monovalent(vorlage, cols, felder_f, ws, error_count, error_count_total)
    assert check_cell_color(ws, 'A', 1, (220, 20, 60))
    assert check_cell_color(ws, 'A', 7, (220, 20, 60))
    assert check_cell_color(ws, 'B', 3, (220, 20, 60))
    assert check_cell_color(ws, 'B', 5, (220, 20, 60))
    assert check_cell_color(ws, 'C', 9, (220, 20, 60))
    assert check_cell_color(ws, 'D', 4, (220, 20, 60))

    wb.close()



# Unit Test für check_datum
def test_check_datum():
    vorlage_path = r'excel_unit_tests/test_datum.xlsx'
    vorlage = open_closed_or_opened_file(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)
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

    index_felder = []
    for col in cols:
        for i in range(len(felder)):
            if col == felder['Feldcode'][i]:
                index_felder.append(i)

    felder_f = felder.iloc[index_felder].reset_index().drop('index', axis=1)                

    error_count = [0]
    error_count_total = [0]

    check_datum(felder_f, vorlage, ws, error_count, error_count_total)
    assert check_cell_color(ws, 'A', 3, (220, 20, 60))
    assert check_cell_color(ws, 'B', 4, (220, 20, 60))
    assert check_cell_color(ws, 'B', 5, (220, 20, 60))
    assert check_cell_color(ws, 'C', 8, (220, 20, 60))
    assert check_cell_color(ws, 'C', 9, (220, 20, 60))
    assert check_cell_color(ws, 'C', 12, (220, 20, 60))
    assert check_cell_color(ws, 'C', 13, (220, 20, 60))

    wb.close()



# Unit Test für one_line_test
def test_check_datum():
    vorlage_path = r'excel_unit_tests/test_one_line.xlsx'
    vorlage = open_closed_or_opened_file(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)
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

    index_felder = []
    for col in cols:
        for i in range(len(felder)):
            if col == felder['Feldcode'][i]:
                index_felder.append(i)

    felder_f = felder.iloc[index_felder].reset_index().drop('index', axis=1)                

    error_count = [0]
    error_count_total = [0]

    check_one_line_text(felder_f, vorlage, ws, error_count, error_count_total)
    assert check_cell_color(ws, 'A', 2, (238, 232, 170))
    assert check_cell_color(ws, 'B', 7, (238, 232, 170))
    assert check_cell_color(ws, 'B', 8, (238, 232, 170))

    wb.close()



# Unit Test für check_katalog
def test_check_katalog():
    vorlage_path = r'excel_unit_tests/test_katalog.xlsx'
    vorlage = open_closed_or_opened_file(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)
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

    index_felder = []
    for col in cols:
        for i in range(len(felder)):
            if col == felder['Feldcode'][i]:
                index_felder.append(i)

    felder_f = felder.iloc[index_felder].reset_index().drop('index', axis=1)                

    error_count = [0]
    error_count_total = [0]

    check_katalog(vorlage, felder_f, ws, error_count, error_count_total)
    assert check_cell_color(ws, 'A', 7, (220, 20, 60))
    assert check_cell_color(ws, 'A', 8, (220, 20, 60))
    assert check_cell_color(ws, 'A', 11, (220, 20, 60))
    assert check_cell_color(ws, 'A', 12, (220, 20, 60))
    assert check_cell_color(ws, 'A', 14, (220, 20, 60))

    wb.close()



# Unit Test für check_pflichtfeld
def test_check_katalog():
    vorlage_path = r'excel_unit_tests/test_pflichtfeld.xlsx'
    vorlage = open_closed_or_opened_file(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)
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

    index_felder = []
    for col in cols:
        for i in range(len(felder)):
            if col == felder['Feldcode'][i]:
                index_felder.append(i)

    felder_f = felder.iloc[index_felder].reset_index().drop('index', axis=1)                

    error_count = [0]
    error_count_total = [0]

    check_pflichtfeld(vorlage, cols, felder_f, ws, error_count, error_count_total)
    assert check_cell_color(ws, 'A', 3, (220, 20, 60))
    assert check_cell_color(ws, 'B', 7, (220, 20, 60))
    assert check_cell_color(ws, 'B', 8, (220, 20, 60))
    assert check_cell_color(ws, 'E', 6, (220, 20, 60))

    wb.close()



# Unit Test für check_pflichtfeld_fehlend
def test_check_katalog():
    vorlage_path = r'excel_unit_tests/test_pflichtfeld_fehlend_Diagnosis.xlsx'
    vorlage = open_closed_or_opened_file(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)
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

    index_felder = []
    for col in cols:
        for i in range(len(felder)):
            if col == felder['Feldcode'][i]:
                index_felder.append(i)

    felder_f = felder.iloc[index_felder].reset_index().drop('index', axis=1)                

    error_count = [0]
    error_count_total = [0]

    check_pflichtfeld_fehlend(vorlage, file_name, cols, felder, ws, error_count, error_count_total)
    assert error_count == [1]

    wb.close()


    vorlage_path = r'excel_unit_tests/test_pflichtfeld_fehlend_Reservierung.xlsx'
    vorlage = open_closed_or_opened_file(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)
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

    index_felder = []
    for col in cols:
        for i in range(len(felder)):
            if col == felder['Feldcode'][i]:
                index_felder.append(i)

    felder_f = felder.iloc[index_felder].reset_index().drop('index', axis=1)                

    error_count = [0]
    error_count_total = [0]

    check_pflichtfeld_fehlend(vorlage, file_name, cols, felder, ws, error_count, error_count_total)
    assert error_count == [1]

    wb.close()



    vorlage_path = r'excel_unit_tests/test_pflichtfeld_fehlend_Sample & Aliquot.xlsx'
    vorlage = open_closed_or_opened_file(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)
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

    index_felder = []
    for col in cols:
        for i in range(len(felder)):
            if col == felder['Feldcode'][i]:
                index_felder.append(i)

    felder_f = felder.iloc[index_felder].reset_index().drop('index', axis=1)                

    error_count = [0]
    error_count_total = [0]

    check_pflichtfeld_fehlend(vorlage, file_name, cols, felder, ws, error_count, error_count_total)
    assert error_count == [2]

    wb.close()



# Unit Test für check_komma_stellen
def test_check_komma_stellen():
    vorlage_path = r'excel_unit_tests/test_pflichtfeld_fehlend_Diagnosis.xlsx'
    vorlage = open_closed_or_opened_file(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.csv')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)
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

    index_felder = []
    for col in cols:
        for i in range(len(felder)):
            if col == felder['Feldcode'][i]:
                index_felder.append(i)

    felder_f = felder.iloc[index_felder].reset_index().drop('index', axis=1)                

    error_count = [0]
    error_count_total = [0]

    assert check_cell_color(ws, 'A', 2, (220, 20, 60))
    assert check_cell_color(ws, 'A', 3, (220, 20, 60))
    assert check_cell_color(ws, 'A', 4, (220, 20, 60))
    assert check_cell_color(ws, 'A', 5, (220, 20, 60))
    assert check_cell_color(ws, 'A', 6, (220, 20, 60))
    assert check_cell_color(ws, 'A', 7, (220, 20, 60))
    assert check_cell_color(ws, 'A', 8, (220, 20, 60))
    assert check_cell_color(ws, 'A', 9, (220, 20, 60))
    assert check_cell_color(ws, 'A', 11, (220, 20, 60))
    assert check_cell_color(ws, 'A', 12, (220, 20, 60))
    assert check_cell_color(ws, 'B', 2, (220, 20, 60))
    assert check_cell_color(ws, 'B', 3, (220, 20, 60))
    assert check_cell_color(ws, 'B', 5, (220, 20, 60))
    assert check_cell_color(ws, 'B', 7, (220, 20, 60))
    assert check_cell_color(ws, 'B', 9, (220, 20, 60))
    assert check_cell_color(ws, 'B', 11, (220, 20, 60))
    assert check_cell_color(ws, 'B', 12, (220, 20, 60))

    wb.close()