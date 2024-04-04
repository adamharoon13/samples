# -*- coding: utf-8 -*-

"""
Created on Wed Mar 23 18:12:49 2022

@author: marce
"""
import os
from datetime import time
from datetime import date
import numpy as np 
import pandas as pd 
import xlrd as xd
import Rechnungen

#Excel laden (Betriebspunkte):
#Öffnen der Hauptdatei zur Suche der Dateinamen --> relevanten Infos
df = pd.read_excel('Alle Betriebspunkte Gemischkondensation.xlsx', header = 20 )
Lcol = list(df['Messdaten Name'])
Dat = list(df['Datum'])
SDate = list(df['Uhrzeit Beginn'])
EDate = list(df['UhrzeitEnde'])
Lcol = list(df['Messdaten Name'])



#Diese Liste von Spaltennamen wurde für die einfache Navigation innerhalb des Dataframes erstellt
cols = ["Datum/Zeit" , "Pumpenleistung [%]", "FU-Spannung [V]", "Frequenz [Hz]", "Spannung Liquiphant [V]",
                          "Massenstrom [g/s]", "Gefoerderte Masse [g]", "Spannung Nadelventil [V]", "Spannung Nadelventil [%]", 
                          "Leistung Heizdraht 1 [%]",  "Leistung Heizdraht 2 [%]", "Druck vor MS [bar]", "Druckverlust vor VK [mbar]", 
                          "Druckverlust nach MS [mbar]", "T0 [°C]", "T1 [°C]", "T2 [°C]", "T6 [°C]", "T7 [°C]", "T8 [°C]", "T11 [°C]", 
                          "T12 [°C]", "T13 [°C]", "T14 [°C]", "T_0 [°C]", "T_1 [°C]", "T_2 [°C]", "T_6 [°C]", "T_7 [°C]", "T_8 [°C]", 
                          "T_11 [°C]", "T_12 [°C]", "T_13 [°C]", "T_14 [°C]", "Prozesszeit [s]", "I_0 [mA]", "I_1[mA]", "I_2 [mA]",
                          "T_pc [°C]", "T_huelse_oben [°C]", "T_huelse_unten [°C]", "T5 [°C]", "T_5 [°C]", "T9 [°C]", "T_9 [°C]", 
                          "T10 [°C]", "T_10 [°C]", "R0 [°C]", "R_0 [Ohm]", "R1 [°C]", "R_1 [Ohm] "]

#Liste der Namen der Spalten, die für die Navigation und das Hinzufügen von Mittelwert und Standardabweichung verwendet werden sollen
meanfilecols = {'Datum' : [],'Uhrzeit Beginn' : [],'Uhrzeit Ende': [], 'Spektren' : [], 'Amplitude-Auswertung [-]': [],'Dampfeintrittszusammensetzung [-]' : [], 'Dampfmassenstrom' : [],'BP':[] , 'Pumpenleistung [%]' : [], 'FU-Spannung [V]' : [], 'Frequenz [Hz]' : [], 'Spannung Liquiphant [V]' : [],
                          "Massenstrom [g/s]" : [], "Gefoerderte Masse [g]" : [], 'Spannung Nadelventil [V]' : [], 'Spannung Nadelventil [%]' : [], 
                          "Leistung Heizdraht 1 [%]" : [],  "Leistung Heizdraht 2 [%]" : [], "Druck vor MS [bar]" : [], "Druckverlust vor VK [mbar]" : [], 
                          "Druckverlust nach MS [mbar]" : [], "T0 [°C]" : [], "T1 [°C]" : [], "T2 [°C]" : [], "T6 [°C]" : [], "T7 [°C]" : [], "T8 [°C]" : [], "T11 [°C]" : [], 
                          "T12 [°C]" : [], "T13 [°C]" : [], "T14 [°C]" : [], "T_0 [°C]" : [], "T_1 [°C]" : [], "T_2 [°C]" : [], "T_6 [°C]" : [], "T_7 [°C]" : [], "T_8 [°C]" : [], 
                          "T_11 [°C]" : [], "T_12 [°C]" : [], "T_13 [°C]" : [], "T_14 [°C]" : [], "Prozesszeit [s]" : [], "I_0 [mA]" : [], "I_1[mA]" : [], "I_2 [mA]" : [],
                          "T_pc [°C]" : [], "T_huelse_oben [°C]" : [], "T_huelse_unten [°C]" : [], "T5 [°C]" : [], "T_5 [°C]" : [], "T9 [°C]" : [], "T_9 [°C]" : [], 
                          "T10 [°C]" : [], "T_10 [°C]" : [], "R0 [°C]" : [], "R_0 [Ohm]" : [], "R1 [°C]" : [], "R_1 [Ohm] " : []}

test = ["Pumpenleistung [%]", "FU-Spannung [V]", "Frequenz [Hz]", "Spannung Liquiphant [V]",
                          "Massenstrom [g/s]", "Gefoerderte Masse [g]", "Spannung Nadelventil [V]", "Spannung Nadelventil [%]", 
                          "Leistung Heizdraht 1 [%]",  "Leistung Heizdraht 2 [%]", "Druck vor MS [bar]", "Druckverlust vor VK [mbar]", 
                          "Druckverlust nach MS [mbar]", "T0 [°C]", "T1 [°C]", "T2 [°C]", "T6 [°C]", "T7 [°C]", "T8 [°C]", "T11 [°C]", 
                          "T12 [°C]", "T13 [°C]", "T14 [°C]", "T_0 [°C]", "T_1 [°C]", "T_2 [°C]", "T_6 [°C]", "T_7 [°C]", "T_8 [°C]", 
                          "T_11 [°C]", "T_12 [°C]", "T_13 [°C]", "T_14 [°C]", "Prozesszeit [s]", "I_0 [mA]", "I_1[mA]", "I_2 [mA]",
                          "T_pc [°C]", "T_huelse_oben [°C]", "T_huelse_unten [°C]", "T5 [°C]", "T_5 [°C]", "T9 [°C]", "T_9 [°C]", 
                          "T10 [°C]", "T_10 [°C]", "R0 [°C]", "R_0 [Ohm]", "R1 [°C]", "R_1 [Ohm] "]

#Liste der Spaltennamen, die für die Erstellung neuer Dataframes zu verwenden sind, in denen die extrahierten Daten gespeichert werden
data = {'Datum/Zeit' : [] , 'Pumpenleistung [%]' : [], 'FU-Spannung [V]' : [], 'Frequenz [Hz]' : [], 'Spannung Liquiphant [V]' : [],
                          "Massenstrom [g/s]" : [], "Gefoerderte Masse [g]" : [], 'Spannung Nadelventil [V]' : [], 'Spannung Nadelventil [%]' : [], 
                          "Leistung Heizdraht 1 [%]" : [],  "Leistung Heizdraht 2 [%]" : [], "Druck vor MS [bar]" : [], "Druckverlust vor VK [mbar]" : [], 
                          "Druckverlust nach MS [mbar]" : [], "T0 [°C]" : [], "T1 [°C]" : [], "T2 [°C]" : [], "T6 [°C]" : [], "T7 [°C]" : [], "T8 [°C]" : [], "T11 [°C]" : [], 
                          "T12 [°C]" : [], "T13 [°C]" : [], "T14 [°C]" : [], "T_0 [°C]" : [], "T_1 [°C]" : [], "T_2 [°C]" : [], "T_6 [°C]" : [], "T_7 [°C]" : [], "T_8 [°C]" : [], 
                          "T_11 [°C]" : [], "T_12 [°C]" : [], "T_13 [°C]" : [], "T_14 [°C]" : [], "Prozesszeit [s]" : [], "I_0 [mA]" : [], "I_1[mA]" : [], "I_2 [mA]" : [],
                          "T_pc [°C]" : [], "T_huelse_oben [°C]" : [], "T_huelse_unten [°C]" : [], "T5 [°C]" : [], "T_5 [°C]" : [], "T9 [°C]" : [], "T_9 [°C]" : [], 
                          "T10 [°C]" : [], "T_10 [°C]" : [], "R0 [°C]" : [], "R_0 [Ohm]" : [], "R1 [°C]" : [], "R_1 [Ohm] " : []}
              
rechnunDat = { 'Datum' : [], 'BP' : [], 'Zusammensetzung am Eintritt' : [], 'Massenstrom Messstrecke' : [], 'Eintrittstemperatur in °C' : [], 'Austrittstemperatur in °C' : [], 'Druck vor Messstrecke in bar' : [],
              'T_dampf' : [], 'T_gas' : [], 'T_l' : [], 'Amplitude' : [], 'Temperatur FTIR' : [], 'Differenz Temperatur TF in K' : [], 'Mittlere Temperatur TF in °C' : [], 'Wärmekapazität TF in J/(kgK)' : [],
              'Dichte TF in kg/m3' : [], 'Wärmeleitfähigkeit TF in W/(mK)' : [], 'Dynamische Viskosität in Pas' : [], 'Prandtl-Zahl' : [] , 'Geschwindigkeit TF' : [], 'Re-Zahl TF' : [], 'Reibungsbeiwert TF' : [],
              'Nu-Zahl TF' : [], 'alpha_innen  in W/m2K' : [], 'Wärmestom in W' : [], 'Wärmestromdichte in W/m2' : [], 'Temperatur Taulinie in °C' : [], 'Temperatur Siedelinie in °C' : [], 'Temperaturgleit in K' : [],
              'Logarithmische Temperaturdifferenz in K' : [], 'Overall heat transfer coefficient in W/m2K' : [], 'alpha_kondensation in W/m2K' : [], 'Zusammensetzung Kondensat' : []
              }
#Laden unserer Datei in einen Datenrahmen und Setzen einiger Zähler und Flags zur späteren Verwendung  
filedf = pd.DataFrame(data)
sheetcount = 0
writerFlag = False

#Erstellen des Dataframes für die Mittelwerte/standardabweichungen und die Namensgebeung der Dateien 
meandf = pd.DataFrame(meanfilecols)
stddf = pd.DataFrame(meanfilecols)
rechnunDf = pd.DataFrame(rechnunDat)
meanfile = 'All means.xlsx'
stdfile = 'All Std_Dev.xlsx'

#Zuordnung bzw. Beschreibung verschiedener 'Variablen' bzw. relevanter Spaltennamen
beginn = time()
ende = time()
datum = date.fromisoformat('2019-12-04')
Spektren = ''
AmpSpek = None
Eintritt = None
Dampfmass = None
MassenDampf = None
rechnunList = []


rechnun_row = {'Datum' : '-', 'BP' : ' ', 'Zusammensetzung am Eintritt' : 'y_ein', 'Massenstrom Messstrecke' : 'dot_m_ms', 'Eintrittstemperatur in °C' : 'T_tf_ein', 'Austrittstemperatur in °C' : 'T_tf_aus', 'Druck vor Messstrecke in bar' : 'p_ms',
                                  'T_dampf' : 'T_dampf', 'T_gas' : 'T_g', 'T_l' : 'T_l', 'Amplitude' : 'a', 'Temperatur FTIR' : 'T_ftir', 'Differenz Temperatur TF in K' : 'dT_tf', 'Mittlere Temperatur TF in °C' : 'T_tf', 'Wärmekapazität TF in J/(kgK)' : 'C',
                                  'Dichte TF in kg/m3' : 'D', 'Wärmeleitfähigkeit TF in W/(mK)' : 'L', 'Dynamische Viskosität in Pas' : 'V', 'Prandtl-Zahl' : 'PR' , 'Geschwindigkeit TF' : 'w_tf', 'Re-Zahl TF' : 'Re_tf', 'Reibungsbeiwert TF' : 'zeta_tf',
                                  'Nu-Zahl TF' : 'Nu_tf', 'alpha_innen  in W/m2K' : 'alpha_innen', 'Wärmestom in W' : 'dot_Q', 'Wärmestromdichte in W/m2' : 'dot_q', 'Temperatur Taulinie in °C' : 'T-dew', 'Temperatur Siedelinie in °C' : 'T_bub', 'Temperaturgleit in K' : 'T_dew - T_bub',
                                  'Logarithmische Temperaturdifferenz in K' : 'dT_log', 'Overall heat transfer coefficient in W/m2K' : 'k', 'alpha_kondensation in W/m2K' : 'alpha_kondensation', 'Zusammensetzung Kondensat' : 'x'}
#The symbols for the indivdual variables/columns are added to the dataframe in this code              
rechnunDf = rechnunDf.append(rechnun_row,ignore_index = True)

#ex.func(12)

#In der folgenden Zeile werden Eingaben des Benutzers zu den Betriebspunkten entgegengenommen
choice = input('Enter 1 to start at a custom datapoint and 2 to cover all datapoints.')

#Wird '1' eingegeben: Das Programm startet an einem vom Benutzer angegebenen Punkt --> Zeilen der Hauptdatei.
if choice == '1':
    datapoint = input('Enter row of the datapoint you wish to start at: ')
    datapoint = int(datapoint)
    
    val = datapoint - 22 
    #'22'um die ersten nicht relevanten Zeilen der Hauptdatei zu überspringen 
    #print('This is current ind: ',val) #--> hier kann angegenen werden, welche Zeile bearbeitet wird --> val = Zeilennummer -22
    
    #Die folgende Schleife navigiert durch die Hauptdatei, Zeile für Zeile, extrahiert die Dateinamen und extrahiert die relevanten Daten aus den genannten Dateien und speichert sie in einem neuen Datenrahmen
    #Dies wird für jeden Index wiederholt, bis die Datei endet 
    for ind in range(val,len(df.index)):
      #print('df.index length: ',len(df.index))
      find = False
      fileflag = False
      count = 0
      
      #Extrahieren der Namen der einzelnen Dateien als Liste
      file = df['Messdaten Name'][ind]
      
      #die Datei enthält einige leere Zeilen, die das Programm zum Absturz bringen können
      #Die folgende Bedingung prüft, ob der Index einen Dateinamen hat oder nicht und fährt nur fort, wenn die Zeile nicht leer ist
      if type(file) != float:
          
          #print(df.columns)
          
          #Erstellung verschiedener Versionen des Dateinamens zur späteren Verwendung --> csv, excel
          file_name = file
          raw_excel = file + '_raw.xlsx'
          file_excel = file + '.xlsx'
          rawname = file + '_raw.txt'
          file += '.txt' 
          
          
          #Hier wird geprüft, ob die Datei in unserem Verzeichnis existiert. Wenn ja, wird sie verarbeitet, wenn nicht, geht es weiter zum nächsten Index mit dem nächsten Dateinamen
          if os.path.isfile(file):
              
              #Erstellen eines Writer-Objekt, das am Ende verwendet wird, um die extrahierten Daten an die erstellte Exceldatei anzuhängen.
              #die Bedingung unten ist auf false gesetzt, so dass ein Writer-Objekt nur erstellt wird, wenn ein neuer Dateiname in den Indizes erkannt wird, 
              #um nicht für jeden Index einen neuen Writer erstellen zu müssen
              if writerFlag == False:
                  writer = pd.ExcelWriter(file_excel, engine='xlsxwriter')
                  writerFlag = True
              
              #dieser Dataframe wird verwendet, der alle extrahierten Daten enthält 
              newdf = pd.DataFrame(data)
              
              #Öffnen der Datei, die im Verzeichnis gefunden wurde --> printen der CSV Datei (Name) in Konsole
              f1 = open(file,'r')
              print(file)
              #Extrahieren aller Zeilen aus der Datei in Form einer Liste
              lines = f1.readlines()
            
              sheet = 'BP'
              
              #Indizierung durch alle Zeilen der Datei und Extraktion der Zeile beim Durchlaufen
              for i in lines:
                  
                  #die folgenden Bedingungen gelten für die ersten sechs Zeilen der Daten, 
                  #da sie keine nützlichen Daten enthalten
                  if not i:
                      continue
                  if count < 6:
                      
                      count+=1
                      continue
                      
                  #Trennzeichen der Zeilen definieren und in  Form einer Liste zurückgeben
                  #Nun vereinfachter Vergleich der Zeiten durchführbar 
                  Sline = i.split(';')
                  
                  #Erstellen einer neuen Zeile, die an den neuen Datenrahmen angehängt wird
                  #folgende Befehle werden benutzt, um jeweils die Daten einer Zeile zu speichern
                  #Zeitvergleich: wenn sie zwischen der "Startzeit" und der "Endzeit" liegt, wird sie 
                  #in Form einer Zeile hinzugefügt
                  new_row = {'Datum/Zeit' : Sline[0] , 'Pumpenleistung [%]' : Sline[1], 'FU-Spannung [V]' : Sline[2], 'Frequenz [Hz]' : Sline[3], 'Spannung Liquiphant [V]' : Sline[4],
                          "Massenstrom [g/s]" :Sline[5], "Gefoerderte Masse [g]" : Sline[6], 'Spannung Nadelventil [V]' : Sline[7], 'Spannung Nadelventil [%]' : Sline[8], 
                          "Leistung Heizdraht 1 [%]" : Sline[9],  "Leistung Heizdraht 2 [%]" : Sline[10], "Druck vor MS [bar]" : Sline[11], "Druckverlust vor VK [mbar]" : Sline[12], 
                          "Druckverlust nach MS [mbar]" : Sline[13], "T0 [°C]" : Sline[14], "T1 [°C]" : Sline[15], "T2 [°C]" : Sline[16], "T6 [°C]" : Sline[17], "T7 [°C]" : Sline[18], "T8 [°C]" : Sline[19], "T11 [°C]" : Sline[20], 
                          "T12 [°C]" : Sline[21], "T13 [°C]" : Sline[22], "T14 [°C]" : Sline[23], "T_0 [°C]" : Sline[24], "T_1 [°C]" : Sline[25], "T_2 [°C]" : Sline[26], "T_6 [°C]" : Sline[27], "T_7 [°C]" : Sline[28], "T_8 [°C]" : Sline[29], 
                          "T_11 [°C]" : Sline[30], "T_12 [°C]" : Sline[31], "T_13 [°C]" : Sline[32], "T_14 [°C]" : Sline[33], "Prozesszeit [s]" : Sline[34], "I_0 [mA]" : Sline[35], "I_1[mA]" : Sline[36], "I_2 [mA]" : Sline[37],
                          "T_pc [°C]" : Sline[38], "T_huelse_oben [°C]" : Sline[39], "T_huelse_unten [°C]" : Sline[40], "T5 [°C]" : Sline[41], "T_5 [°C]" : Sline[42], "T9 [°C]" : Sline[43], "T_9 [°C]" : Sline[44], 
                          "T10 [°C]" : Sline[45], "T_10 [°C]" : Sline[46], "R0 [°C]" : Sline[47], "R_0 [Ohm]" : Sline[48], "R1 [°C]" : Sline[49], "R_1 [Ohm] " : Sline[50]}
                  
                  tim = time.fromisoformat(Sline[0])
                  #print(tim)
                  if tim < df['Uhrzeit Beginn'][ind]:
                      continue
                  if tim > df['Uhrzeit Beginn'][ind] and tim < df['UhrzeitEnde'][ind]:
                      newdf = newdf.append(new_row,ignore_index=True)
                      beginn = df['Uhrzeit Beginn'][ind]
                      ende = df['UhrzeitEnde'][ind]
                      datum = df['Datum'][ind].date()
                      Spektren = df['Bemerkungen'][ind]
                      AmpSpek = df['Amplitude-Auswertung [-]'][ind]
                      Eintritt = df['Dampfeintrittszusammensetzung [-]'][ind] #This is the value that will be used for the value of y_ein in the rechnungen function
                      Dampfmass = df['Dampfmassenstrom [g/s]'][ind]
                      MassenDampf = df['Massenstrom Dampf'] [ind]
                      
                      find = True
                      
                  count+=1
                  
            #Die folgenden Zeilen werden ausgeführt, wenn eine Datei geöffnet wird, 
            #aber die erforderlichen Messungen in der Datei nicht vorhanden sind. --> Fehler bei nicht vorhandenen Messungen
              if find == False:
                  print('File Error: Time does not exist')
                  continue
                
              filedf = filedf.append(newdf)  
              
              #Die folgenden 8 Zeilen durchsuchen den gesamten Datenrahmen, Spalte für Spalte 
              #Ziel ist es ihr Format von String auf Float ändern
              #Da die Datums-/Zeitspalte keine Formatänderung erfordert, wird sie übersprungen
              for colName in cols:
                  if colName == 'Datum/Zeit':
                      continue
                  for i in newdf.index:
                      x = newdf[colName][i]
                      x = str(x).replace(',','.')
                      newdf.loc[i,colName] = x
                
                #Fehlerbehebung in der Konsole über '.loc'
                
                  newdf[colName] = newdf[colName].astype(float)
                  
              #df.loc['mean'] = df.mean(numeric_only=True)
              
              
              #Berechnung von Mittelwert und Standardabweichung für jede einzelne Spalte
              averages = [newdf[key].describe()['mean'] for key in test]
              stds = [newdf[key].describe()['std'] for key in test]
              indexes = newdf.index.tolist()
              indexes.append('mean')
              indexes.append('std_dev')
              newdf.reindex(indexes)
             
              #Hinzufügen der Zeilen für Mittelwert und Standardabweichung am unteren Rand des DataFrame
              i = 0
              for key in newdf:
                  if key == 'Datum/Zeit':
                      continue
                  newdf.at['mean', key] = averages[i]
                  newdf.at['std_dev', key] = stds[i]
                  i += 1
              
              #die folgenden Zeilen konvertieren den gesamten Datenrahmen in eine Excel-Tabelle --> printen der ganzen BP
              sheetcount += 1
              sheet += str(sheetcount) 
              newdf.to_excel(writer, sheet_name=sheet)
              print('Currently working on sheet: ',sheet,'\n') #--> sheet printet den gesamten BP
              #print(newdf, sheet) um alles zu printen
              
              mean_row = {'Datum' : datum ,'Uhrzeit Beginn' : beginn,'Uhrzeit Ende' : ende, 'Spektren' : Spektren, 'Amplitude-Auswertung [-]': AmpSpek,'Dampfeintrittszusammensetzung [-]' : Eintritt, 'Dampfmassenstrom' : Dampfmass,'BP' : sheet,'Pumpenleistung [%]' : averages[0], 'FU-Spannung [V]' : averages[1], 'Frequenz [Hz]' : averages[2], 'Spannung Liquiphant [V]' : averages[3],
                          "Massenstrom [g/s]" : averages[4], "Gefoerderte Masse [g]" : averages[5], 'Spannung Nadelventil [V]' : averages[6], 'Spannung Nadelventil [%]' : averages[7], 
                          "Leistung Heizdraht 1 [%]" : averages[8],  "Leistung Heizdraht 2 [%]" : averages[9], "Druck vor MS [bar]" : averages[10], "Druckverlust vor VK [mbar]" : averages[11], 
                          "Druckverlust nach MS [mbar]" : averages[12], "T0 [°C]" : averages[13], "T1 [°C]" : averages[14], "T2 [°C]" : averages[15], "T6 [°C]" : averages[16], "T7 [°C]" : averages[17], "T8 [°C]" : averages[18], "T11 [°C]" : averages[19], 
                          "T12 [°C]" : averages[20], "T13 [°C]" : averages[21], "T14 [°C]" : averages[22], "T_0 [°C]" : averages[23], "T_1 [°C]" : averages[24], "T_2 [°C]" : averages[25], "T_6 [°C]" : averages[26], "T_7 [°C]" : averages[27], "T_8 [°C]" : averages[28], 
                          "T_11 [°C]" : averages[29], "T_12 [°C]" : averages[30], "T_13 [°C]" : averages[31], "T_14 [°C]" : averages[32], "Prozesszeit [s]" : averages[33], "I_0 [mA]" : averages[34], "I_1[mA]" : averages[35], "I_2 [mA]" : averages[36],
                          "T_pc [°C]" : averages[37], "T_huelse_oben [°C]" : averages[38], "T_huelse_unten [°C]" : averages[39], "T5 [°C]" : averages[40], "T_5 [°C]" : averages[41], "T9 [°C]" : averages[42], "T_9 [°C]" : averages[43], 
                          "T10 [°C]" : averages[44], "T_10 [°C]" : averages[45], "R0 [°C]" : averages[46], "R_0 [Ohm]" : averages[47], "R1 [°C]" : averages[48], "R_1 [Ohm] " : averages[49]}
              
              std_row = {'Datum' : datum ,'Uhrzeit Beginn' : beginn,'Uhrzeit Ende' : ende, 'Spektren' : Spektren, 'Amplitude-Auswertung [-]': AmpSpek,'Dampfeintrittszusammensetzung [-]' : Eintritt, 'Dampfmassenstrom' : Dampfmass,'BP' : sheet,'Pumpenleistung [%]' : stds[0], 'FU-Spannung [V]' : stds[1], 'Frequenz [Hz]' : stds[2], 'Spannung Liquiphant [V]' : stds[3],
                          "Massenstrom [g/s]" : stds[4], "Gefoerderte Masse [g]" : stds[5], 'Spannung Nadelventil [V]' : stds[6], 'Spannung Nadelventil [%]' : stds[7], 
                          "Leistung Heizdraht 1 [%]" : stds[8],  "Leistung Heizdraht 2 [%]" : stds[9], "Druck vor MS [bar]" : stds[10], "Druckverlust vor VK [mbar]" : stds[11], 
                          "Druckverlust nach MS [mbar]" : stds[12], "T0 [°C]" : stds[13], "T1 [°C]" : stds[14], "T2 [°C]" : stds[15], "T6 [°C]" : stds[16], "T7 [°C]" : stds[17], "T8 [°C]" : stds[18], "T11 [°C]" : stds[19], 
                          "T12 [°C]" : stds[20], "T13 [°C]" : stds[21], "T14 [°C]" : stds[22], "T_0 [°C]" : stds[23], "T_1 [°C]" : stds[24], "T_2 [°C]" : stds[25], "T_6 [°C]" : stds[26], "T_7 [°C]" : stds[27], "T_8 [°C]" : stds[28], 
                          "T_11 [°C]" : stds[29], "T_12 [°C]" : stds[30], "T_13 [°C]" : stds[31], "T_14 [°C]" : stds[32], "Prozesszeit [s]" : stds[33], "I_0 [mA]" : stds[34], "I_1[mA]" : stds[35], "I_2 [mA]" : stds[36],
                          "T_pc [°C]" : stds[37], "T_huelse_oben [°C]" : stds[38], "T_huelse_unten [°C]" : stds[39], "T5 [°C]" : stds[40], "T_5 [°C]" : stds[41], "T9 [°C]" : stds[42], "T_9 [°C]" : stds[43], 
                          "T10 [°C]" : stds[44], "T_10 [°C]" : stds[45], "R0 [°C]" : stds[46], "R_0 [Ohm]" : stds[47], "R1 [°C]" : stds[48], "R_1 [Ohm] " : stds[49]}
              
              
              rechnunList = Rechnungen.rechnungen( Eintritt, Dampfmass, averages[27], averages[18], averages[10], averages[19], averages[16], averages[13], AmpSpek, averages[44])
                  
              #Below code obtains the values from the above function in a list and then that list and some additional values are used to fill the dataframe
              #for the rechnungen file
              rechnun_row = {'Datum' : datum, 'BP' : sheet, 'Zusammensetzung am Eintritt' : Eintritt, 'Massenstrom Messstrecke' : Dampfmass, 'Eintrittstemperatur in °C' : averages[27], 'Austrittstemperatur in °C' : averages[18], 'Druck vor Messstrecke in bar' : averages[10],
                                  'T_dampf' : averages[19], 'T_gas' : averages[16], 'T_l' : averages[13], 'Amplitude' : AmpSpek, 'Temperatur FTIR' : averages[44], 'Differenz Temperatur TF in K' : rechnunList[0], 'Mittlere Temperatur TF in °C' : rechnunList[1], 'Wärmekapazität TF in J/(kgK)' : rechnunList[2],
                                  'Dichte TF in kg/m3' : rechnunList[3], 'Wärmeleitfähigkeit TF in W/(mK)' : rechnunList[4], 'Dynamische Viskosität in Pas' : rechnunList[5], 'Prandtl-Zahl' : rechnunList[6] , 'Geschwindigkeit TF' : rechnunList[7], 'Re-Zahl TF' : rechnunList[8], 'Reibungsbeiwert TF' : rechnunList[9],
                                  'Nu-Zahl TF' : rechnunList[10], 'alpha_innen  in W/m2K' : rechnunList[11], 'Wärmestom in W' : rechnunList[12], 'Wärmestromdichte in W/m2' : rechnunList[13], 'Temperatur Taulinie in °C' : rechnunList[14], 'Temperatur Siedelinie in °C' : rechnunList[15], 'Temperaturgleit in K' : rechnunList[16],
                                  'Logarithmische Temperaturdifferenz in K' : rechnunList[17], 'Overall heat transfer coefficient in W/m2K' : rechnunList[18], 'alpha_kondensation in W/m2K' : rechnunList[19], 'Zusammensetzung Kondensat' : rechnunList[20]}
                  
              meandf = meandf.append(mean_row,ignore_index=True)
              stddf = stddf.append(std_row,ignore_index = True)
              rechnunDf = rechnunDf.append(rechnun_row,ignore_index = True)
              
              print(meandf) #um die berechneten mittelwerte zu printen 
              print(rechnunDf)
              f1.close()
              
              #Dieser Teil prüft den Dateinamen in der nächsten Zeile. Wenn die nächste Zeile denselben Dateinamen hat,
              #fährt der Code mit dem nächsten Index fort
              #Wenn dies nicht der Fall ist und sich der Dateiname ändert, bedeutet dies, dass alle relevanten Daten 
              #extrahiert wurden und die Excel-Datei nun ordnungsgemäß gespeichert werden kann
              if ind == df.shape[0] - 1:
                      writer.save()
                      sheetcount = 0
                      writerFlag = False
                      filedf.to_csv(rawname,index = False)
                      filedf.to_excel(raw_excel)
                      filedf = pd.DataFrame(data)
                      #print('here')
                      break
              if df['Messdaten Name'][ind + 1] !=  file_name:
                  writer.save()
                  sheetcount = 0
                  writerFlag = False
                  filedf.to_csv(rawname,index = False)
                  filedf.to_excel(raw_excel)
                  filedf = pd.DataFrame(data)
                  
                  
#sollte 2 eingegeben werden, wird alles von "oben nach unten" durchgespielt (siehe oben, Aufbau grob der gleiche)
elif choice == '2':
    for ind in df.index:
          find = False
            
          
          count = 0
          #Extrahieren der "CSV-Namen" der einzelnen Dateien als Liste
          file = df['Messdaten Name'][ind]
          
          #die Datei enthält einige leere Zeilen, die das Programm zum Absturz bringen können
          #deswegen prüft die folgende Bedingung, ob der Index einen "CSV-Namen" hat oder nicht und fährt nur fort, wenn die Zeile nicht leer ist
          if type(file) != float:
              
              #Erstellung verschiedener Versionen des Dateinamens zur späteren Verwendung
              file_name = file
              raw_excel = file + '_raw.xlsx'
              file_excel = file + '.xlsx'
              rawname = file + '_raw.txt'
              file += '.txt' 
              
              #Hier wird geprüft, ob die Datei zum zugehörgen "CSV-Namen" in unserem Verzeichnis existiert
              #Wenn ja, wird sie verarbeitet, wenn nicht, wird der nächste Index mit dem nächsten Dateinamen aufgerufen
              if os.path.isfile(file):
                  
                 #Erstellen eines "Writers", der am Ende verwendet wird, um die extrahierten Daten an die erstellte Excel-Datei anzuhängen
                 #Die unten stehende Bedingung wird auf "false" gesetzt, damit ein "Writer" nur erstellt wird, wenn ein neuer Dateiname 
                 #in den Indizes erkannt wird, da wir nicht für jeden Index einen neuen Writer erstellen müssen
                  if writerFlag == False:
                      writer = pd.ExcelWriter(file_excel, engine='xlsxwriter')
                      writerFlag = True
                  
                  #der folgende Datenrahmen wird verwendet, der alle extrahierten Daten enthält (Data s.o.)
                  newdf = pd.DataFrame(data)
                  
                  #Öffnen der Datei, die im Verzeichnis gefunden wurde und ausgeben der Datei in der Konsole
                  f1 = open(file,'r')
                  print(file)
                 
                  #Extrahieren aller Zeilen aus der Datei in Form einer Liste
                  lines = f1.readlines()
                  
                  sheet = 'BP'
                  #Indizierung durch alle Zeilen der Datei und Extraktion der Zeile beim Durchlaufen
                  for i in lines:
                      
                      #die folgenden Bedingungen gelten für die ersten sechs Zeilen der CSV Daten, da diese keine Daten enthalten nützliche Daten
                      if not i:
                          continue
                      if count < 6:
                          
                          count+=1
                          continue
                          
                      #Trennzeichen der Zeilen definieren und in  Form einer Liste zurückgeben
                      #Nun vereinfachter Vergleich der Zeiten durchführbar 
                      Sline = i.split(';')
                      
                     #Erstellen einer neuen Zeile, die an den neuen Datenrahmen angehängt wird
                     #folgende Befehle werden benutzt, um jeweils die Daten einer Zeile zu speichern
                     #Zeitvergleich: wenn sie zwischen der "Startzeit" und der "Endzeit" liegt, wird sie 
                     #in Form einer Zeile hinzugefügt
                      new_row = {'Datum/Zeit' : Sline[0] , 'Pumpenleistung [%]' : Sline[1], 'FU-Spannung [V]' : Sline[2], 'Frequenz [Hz]' : Sline[3], 'Spannung Liquiphant [V]' : Sline[4],
                              "Massenstrom [g/s]" :Sline[5], "Gefoerderte Masse [g]" : Sline[6], 'Spannung Nadelventil [V]' : Sline[7], 'Spannung Nadelventil [%]' : Sline[8], 
                              "Leistung Heizdraht 1 [%]" : Sline[9],  "Leistung Heizdraht 2 [%]" : Sline[10], "Druck vor MS [bar]" : Sline[11], "Druckverlust vor VK [mbar]" : Sline[12], 
                              "Druckverlust nach MS [mbar]" : Sline[13], "T0 [°C]" : Sline[14], "T1 [°C]" : Sline[15], "T2 [°C]" : Sline[16], "T6 [°C]" : Sline[17], "T7 [°C]" : Sline[18], "T8 [°C]" : Sline[19], "T11 [°C]" : Sline[20], 
                              "T12 [°C]" : Sline[21], "T13 [°C]" : Sline[22], "T14 [°C]" : Sline[23], "T_0 [°C]" : Sline[24], "T_1 [°C]" : Sline[25], "T_2 [°C]" : Sline[26], "T_6 [°C]" : Sline[27], "T_7 [°C]" : Sline[28], "T_8 [°C]" : Sline[29], 
                              "T_11 [°C]" : Sline[30], "T_12 [°C]" : Sline[31], "T_13 [°C]" : Sline[32], "T_14 [°C]" : Sline[33], "Prozesszeit [s]" : Sline[34], "I_0 [mA]" : Sline[35], "I_1[mA]" : Sline[36], "I_2 [mA]" : Sline[37],
                              "T_pc [°C]" : Sline[38], "T_huelse_oben [°C]" : Sline[39], "T_huelse_unten [°C]" : Sline[40], "T5 [°C]" : Sline[41], "T_5 [°C]" : Sline[42], "T9 [°C]" : Sline[43], "T_9 [°C]" : Sline[44], 
                              "T10 [°C]" : Sline[45], "T_10 [°C]" : Sline[46], "R0 [°C]" : Sline[47], "R_0 [Ohm]" : Sline[48], "R1 [°C]" : Sline[49], "R_1 [Ohm] " : Sline[50]}
                      
                      tim = time.fromisoformat(Sline[0])
                      #print(tim)
                      if tim < df['Uhrzeit Beginn'][ind]:
                          continue
                      if tim > df['Uhrzeit Beginn'][ind] and tim < df['UhrzeitEnde'][ind]:
                          newdf = newdf.append(new_row,ignore_index=True)
                          beginn = df['Uhrzeit Beginn'][ind]
                          ende = df['UhrzeitEnde'][ind]
                          datum = df['Datum'][ind].date()
                          Spektren = df['Bemerkungen'][ind]
                          AmpSpek = df['Amplitude-Auswertung [-]'][ind]
                          Eintritt = df['Dampfeintrittszusammensetzung [-]'][ind]#This is the value that will be used for the value of y_ein in the rechnungen function
                          Dampfmass = df['Dampfmassenstrom [g/s]'][ind]
                          MassenDampf = df['Massenstrom Dampf'] [ind]
                          find = True
                          
                    #auch hier die Fehlermeldung, sollten die Zeiten der CSV-Datei und die der Hauptdatei nicht übereinstimmen
                      count+=1
                  if find == False:
                      print('File Error: Time does not exist, program continues with next datapoint')
                      continue
                    
                  filedf = filedf.append(newdf)  
                  
                  #Die folgenden 8 Zeilen analysieren den gesamten Datenrahmen, Spalte für Spalte, und 
                  #ändern ihr Format von String zu Float. Da die Datums-/Zeitspalte keine Formatänderung erfordert,
                  #wird sie übersprungen.
                  for colName in cols:
                      if colName == 'Datum/Zeit':
                          continue
                      for i in newdf.index:
                          x = newdf[colName][i]
                          x = str(x).replace(',','.')
                          newdf.loc[i,colName] = x
                          
                      #print(newdf.columns)
                      newdf[colName] = newdf[colName].astype(float)
                      
                  #df.loc['mean'] = df.mean(numeric_only=True)
                  
                  
                  #Berechnung von Mittelwert und Standardabweichung für jede einzelne Spalte
                  averages = [newdf[key].describe()['mean'] for key in test]
                  stds = [newdf[key].describe()['std'] for key in test]
                  indexes = newdf.index.tolist()
                  indexes.append('mean')
                  indexes.append('std_dev')
                  newdf.reindex(indexes)
                  
                  #Hinzufügen der Zeilen für den Mittelwert und die Standardabweichung am unteren Rand des DataFrame
                  i = 0
                  for key in newdf:
                      if key == 'Datum/Zeit':
                          continue
                      newdf.at['mean', key] = averages[i]
                      newdf.at['std_dev', key] = stds[i]
                      i += 1
                  
                  #die folgenden Zeilen konvertieren den gesamten Datenrahmen in eine Excel-Tabelle.
                  sheetcount += 1
                  sheet += str(sheetcount) 
                  newdf.to_excel(writer, sheet_name=sheet)
                  print('Currently working on sheet: ',sheet,'\n') #--> sheet printet den gesamten BP
                  #print(newdf, sheet) um alles zu printen
                  
                  mean_row = {'Datum' : datum ,'Uhrzeit Beginn' : beginn,'Uhrzeit Ende' : ende, 'Spektren' : Spektren, 'Amplitude-Auswertung [-]': AmpSpek,'Dampfeintrittszusammensetzung [-]' : Eintritt, 'Dampfmassenstrom' : Dampfmass,'BP' : sheet,'Pumpenleistung [%]' : averages[0], 'FU-Spannung [V]' : averages[1], 'Frequenz [Hz]' : averages[2], 'Spannung Liquiphant [V]' : averages[3],
                              "Massenstrom [g/s]" : averages[4], "Gefoerderte Masse [g]" : averages[5], 'Spannung Nadelventil [V]' : averages[6], 'Spannung Nadelventil [%]' : averages[7], 
                              "Leistung Heizdraht 1 [%]" : averages[8],  "Leistung Heizdraht 2 [%]" : averages[9], "Druck vor MS [bar]" : averages[10], "Druckverlust vor VK [mbar]" : averages[11], 
                              "Druckverlust nach MS [mbar]" : averages[12],  "T0 [°C]" : averages[13], "T1 [°C]" : averages[14], "T2 [°C]" : averages[15], "T6 [°C]" : averages[16], "T7 [°C]" : averages[17], "T8 [°C]" : averages[18], "T11 [°C]" : averages[19], 
                              "T12 [°C]" : averages[20], "T13 [°C]" : averages[21], "T14 [°C]" : averages[22], "T_0 [°C]" : averages[23], "T_1 [°C]" : averages[24], "T_2 [°C]" : averages[25], "T_6 [°C]" : averages[26], "T_7 [°C]" : averages[27], "T_8 [°C]" : averages[28], 
                              "T_11 [°C]" : averages[29], "T_12 [°C]" : averages[30], "T_13 [°C]" : averages[31], "T_14 [°C]" : averages[32], "Prozesszeit [s]" : averages[33], "I_0 [mA]" : averages[34], "I_1[mA]" : averages[35], "I_2 [mA]" : averages[36],
                              "T_pc [°C]" : averages[37], "T_huelse_oben [°C]" : averages[38], "T_huelse_unten [°C]" : averages[39], "T5 [°C]" : averages[40], "T_5 [°C]" : averages[41], "T9 [°C]" : averages[42], "T_9 [°C]" : averages[43], 
                              "T10 [°C]" : averages[44], "T_10 [°C]" : averages[45], "R0 [°C]" : averages[46], "R_0 [Ohm]" : averages[47], "R1 [°C]" : averages[48], "R_1 [Ohm] " : averages[49]}
                  
                  std_row = {'Datum' : datum ,'Uhrzeit Beginn' : beginn,'Uhrzeit Ende' : ende, 'Spektren' : Spektren, 'Amplitude-Auswertung [-]': AmpSpek,'Dampfeintrittszusammensetzung [-]' : Eintritt, 'Dampfmassenstrom' : Dampfmass,'BP' : sheet,'Pumpenleistung [%]' : stds[0], 'FU-Spannung [V]' : stds[1], 'Frequenz [Hz]' : stds[2], 'Spannung Liquiphant [V]' : stds[3],
                              "Massenstrom [g/s]" : stds[4], "Gefoerderte Masse [g]" : stds[5], 'Spannung Nadelventil [V]' : stds[6], 'Spannung Nadelventil [%]' : stds[7], 
                              "Leistung Heizdraht 1 [%]" : stds[8],  "Leistung Heizdraht 2 [%]" : stds[9], "Druck vor MS [bar]" : stds[10], "Druckverlust vor VK [mbar]" : stds[11], 
                              "Druckverlust nach MS [mbar]" : stds[12], "T0 [°C]" : stds[13], "T1 [°C]" : stds[14], "T2 [°C]" : stds[15], "T6 [°C]" : stds[16], "T7 [°C]" : stds[17], "T8 [°C]" : stds[18], "T11 [°C]" : stds[19], 
                              "T12 [°C]" : stds[20], "T13 [°C]" : stds[21], "T14 [°C]" : stds[22], "T_0 [°C]" : stds[23], "T_1 [°C]" : stds[24], "T_2 [°C]" : stds[25], "T_6 [°C]" : stds[26], "T_7 [°C]" : stds[27], "T_8 [°C]" : stds[28], 
                              "T_11 [°C]" : stds[29], "T_12 [°C]" : stds[30], "T_13 [°C]" : stds[31], "T_14 [°C]" : stds[32], "Prozesszeit [s]" : stds[33], "I_0 [mA]" : stds[34], "I_1[mA]" : stds[35], "I_2 [mA]" : stds[36],
                              "T_pc [°C]" : stds[37], "T_huelse_oben [°C]" : stds[38], "T_huelse_unten [°C]" : stds[39], "T5 [°C]" : stds[40], "T_5 [°C]" : stds[41], "T9 [°C]" : stds[42], "T_9 [°C]" : stds[43], 
                              "T10 [°C]" : stds[44], "T_10 [°C]" : stds[45], "R0 [°C]" : stds[46], "R_0 [Ohm]" : stds[47], "R1 [°C]" : stds[48], "R_1 [Ohm] " : stds[49]}
              
                  
                  
                  #print( Eintritt, 0.28, averages[27], averages[18], averages[10], averages[19], 80.14053417, 39.75105742, AmpSpek, averages[44])
                  rechnunList = Rechnungen.rechnungen( Eintritt, Dampfmass, averages[27], averages[18], averages[10], averages[19], averages[16], averages[13], AmpSpek, averages[44])
                  #same as the description given at line 282
                  rechnun_row = {'Datum' : datum, 'BP' : sheet, 'Zusammensetzung am Eintritt' : Eintritt, 'Massenstrom Messstrecke' : Dampfmass, 'Eintrittstemperatur in °C' : averages[27], 'Austrittstemperatur in °C' : averages[18], 'Druck vor Messstrecke in bar' : averages[10],
                                  'T_dampf' : averages[19], 'T_gas' : averages[16], 'T_l' : averages[13], 'Amplitude' : AmpSpek, 'Temperatur FTIR' : averages[44], 'Differenz Temperatur TF in K' : rechnunList[0], 'Mittlere Temperatur TF in °C' : rechnunList[1], 'Wärmekapazität TF in J/(kgK)' : rechnunList[2],
                                  'Dichte TF in kg/m3' : rechnunList[3], 'Wärmeleitfähigkeit TF in W/(mK)' : rechnunList[4], 'Dynamische Viskosität in Pas' : rechnunList[5], 'Prandtl-Zahl' : rechnunList[6] , 'Geschwindigkeit TF' : rechnunList[7], 'Re-Zahl TF' : rechnunList[8], 'Reibungsbeiwert TF' : rechnunList[9],
                                  'Nu-Zahl TF' : rechnunList[10], 'alpha_innen  in W/m2K' : rechnunList[11], 'Wärmestom in W' : rechnunList[12], 'Wärmestromdichte in W/m2' : rechnunList[13], 'Temperatur Taulinie in °C' : rechnunList[14], 'Temperatur Siedelinie in °C' : rechnunList[15], 'Temperaturgleit in K' : rechnunList[16],
                                  'Logarithmische Temperaturdifferenz in K' : rechnunList[17], 'Overall heat transfer coefficient in W/m2K' : rechnunList[18], 'alpha_kondensation in W/m2K' : rechnunList[19], 'Zusammensetzung Kondensat' : rechnunList[20]}
                  #print(rechnunList)
                  meandf = meandf.append(mean_row,ignore_index=True)
                  stddf = stddf.append(std_row,ignore_index = True)
                  rechnunDf = rechnunDf.append(rechnun_row,ignore_index = True)
                  
                  print(meandf)
                  print(rechnunDf)
                  f1.close()
                  
                  #Dieser Teil prüft den Dateinamen in der nächsten Zeile. Wenn die nächste Zeile denselben Dateinamen hat,
                  #fährt der Code mit dem nächsten Index fort
                  #Wenn dies nicht der Fall ist und sich der Dateiname ändert, bedeutet dies, dass alle relevanten Daten 
                  #extrahiert wurden und die Excel-Datei nun ordnungsgemäß gespeichert werden kann
                  if ind == df.shape[0] - 1:
                      writer.save()
                      sheetcount = 0
                      writerFlag = False
                      filedf.to_csv(rawname,index = False)
                      filedf.to_excel(raw_excel)
                      filedf = pd.DataFrame(data)
                      #print('here')
                      break
                  if df['Messdaten Name'][ind + 1] !=  file_name:
                      writer.save()
                      sheetcount = 0
                      writerFlag = False
                      filedf.to_csv(rawname,index = False)
                      filedf.to_excel(raw_excel)
                      filedf = pd.DataFrame(data)

#Speichern der Mittelwert- und Std.Devdatei als Excel (in Custom und nicht Custom)


if choice == '2':
    meandf.to_excel('All Means.xlsx')
    stddf.to_excel('All Std_Dev.xlsx')
    rechnunDf.to_excel('Rechnungen_All.xlsx')
else:
    meandf.to_excel('All Means_Custom.xlsx')
    stddf.to_excel('All Std_Dev_Custom.xlsx')
    rechnunDf.to_excel('Rechnungen_Custom.xlsx')

#HINWEIS: Der Code benötigt aufgrund seiner Vorverarbeitung Zeit, um vollständig ausgeführt zu werden. Es dauert mindestens 2-3 Minuten pro Datei.
# Unterbrechen Sie das Programm in dieser Zeit nicht.