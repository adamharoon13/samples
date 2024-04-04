# -*- coding: utf-8 -*-

#Serdar-Marcel Yaldiz, Matrikelnr: 03251450




#Laden verschiedener Bibliotheken
from matplotlib import pyplot as plt 
import pandas as pd




#Die folgende Funktion  verwendet zwei getrennte Schleifen, um die beiden Graphen unabhängig voneinander zu erzeugen
#Das Vorgehen ist das Selbe, weswegen die zweite Schleife bzr grob kommentiert ist
def Display_visual(fileName):

    #Lesen des Dateinamens
    #Die erste Zeile des Datensatzes entfernen, da sie keine numerischen Daten enthält
    df = pd.read_excel(fileName )
    df = df.loc[1:]
    
    
    #eine Liste mit Markierungen und eine weitere mit Farben erstellen 
    markers = [  "o" , "v" , "^" , "<", ">"]
    colors = ['r','g','b','c','m', 'y', 'k']
    m = 0
    c = 0
    
    
    #Festsetzen der für die Sortierung der Messpunkte relevanten Größen y_ein und dot_m
    yein = df['Zusammensetzung am Eintritt']                                          
    yein = set(yein)                            #y_ein gibt die Farbe des Punktes an  
    dot_m = df['Massenstrom Messstrecke']                                               
    dot_m = set(dot_m)                          #dot_m gibt die Form des Punktes an 
    
    

    #Die folgenden Zeilen legen die Dimensionen und Details des Plots fest (bspw.: Schriftart und -größe)    
    plt.style.use('bmh') 
    font = {'family' : 'normal',
            'weight' : 'normal', 
            'size'   : 18}
    plt.rc('font', **font)
    
    #Die folgende Zeile bestimmt die Abmessungen der ersten Figur
    #Plot 1: alpha_kon über dT_log
    fig, ax = plt.subplots(figsize=(10, 6))
    plt.grid(True)                                                                      #Rasterwert auf 'true' setzen, damit auch ein Raster angezeigt wird
    ax.set_xlabel("$\mathrm{d}T_{\mathrm{log}}$ in $\mathrm{K}$")                       #Name der X-Achse ['dT_log in K']
    ax.set_ylabel("$\u03B1_{\mathrm{kon}}$ in $\mathrm{W}/(\mathrm{m^{2}}\mathrm{K})$") #Name der Y-Achse ['alpha_kon in W/(m^2 *K)']
    ax.set_xlim([0, 18])                                                                #Max,Min Werte x-Achse
    ax.set_ylim([0,10000])                                                              #Max,Min Werte y-Achse
    ax.set_facecolor('white')                                                           #Hintergrund des Plots
    
    f1 = plt.figure(1) #Funktion, die zwischen den beiden erzeugten Figuren unterscheidet 
    
    
    x_coordinate = [] # eine Liste, die die jeweiligen Koordinaten der x-Achse für die Darstellung enthält
    y_coordinate = [] # eine Liste, die die jeweiligen Koordinaten der y-Achse für die Darstellung enthält
    
    
    #Schleife 1 zur Erzeugung des ersten Figur bzw. Plots:
    for i in df.index:
        
        if i == 1:                                                                                      #Diese Bedingung wird so gesetzt, dass keine Abhängigkeiten zwischen den farb- oder formverändernden Variablen geprüft werden
            x_coordinate.append(df.loc[i,'Logarithmische Temperaturdifferenz in K'])
            y_coordinate.append(df.loc[i,'alpha_kondensation in W/m2K'])
            #print(y_coordinate)
        if i > 1 and df.loc[i,'Massenstrom Messstrecke'] == df.loc[i - 1,'Massenstrom Messstrecke']:
            x_coordinate.append(df.loc[i,'Logarithmische Temperaturdifferenz in K'])
            y_coordinate.append(df.loc[i,'alpha_kondensation in W/m2K'])
            
        if i > 1 and df.loc[i,'Massenstrom Messstrecke'] != df.loc[i - 1,'Massenstrom Messstrecke']:
            mi = markers[m]
            ci = colors[c]
            if df.loc[i,'Zusammensetzung am Eintritt'] != df.loc[i - 1,'Zusammensetzung am Eintritt']:   #Diese Bedingung prüft, ob das Symbol der gezeichneten Punkte geändert werden soll oder nicht
                a = df.loc[i-1,'Massenstrom Messstrecke']                                                #dot_m_ms: Massenstrom der Messtrecke
                b = df.loc[i-1,'Zusammensetzung am Eintritt']                                            #y_ein: Zusammensetzung am Eintritt
                lab =  str(b) + ',' + str(a)                                                             #Diese Variable steuert was in der Legende angezeigt wird: 
                #'y_ein= '   "$\.\mathrm{m}_{\mathrm{ms}}$= "                                            #um y_ein und dot_m in Legende mit aufzunehmen 
                
                x_coordinate.sort()                                                                      #Diese Funktion sortiert die dT_log-Werte in aufsteigender Reihenfolge
                                
                #Die folgende Funktion wird verwendet, um alle gesammelten Punkte darzustellen
                #'linestyle = "" ' gibt an wie die eingezeichneten Punkte verbunden werden 
                ax.plot(x_coordinate,y_coordinate,color=ci,linestyle = "",marker=mi, markersize=5, label = lab) 
               
                #Die erfassten Koordinaten werden dann gelöscht, damit neuere Punkte gespeichert werden können
                x_coordinate.clear()
                y_coordinate.clear()
            
                x_coordinate.append(df.loc[i,'Logarithmische Temperaturdifferenz in K'])
                y_coordinate.append(df.loc[i,'alpha_kondensation in W/m2K'])
                m += 1      #Diese Variable dient als Zähler für die Liste, die alle Markierungen enthält
                c += 1      #Diese Variable dient als Zähler für die Liste, die alle Farben enthält
                m = 0
                
            elif df.loc[i,'Zusammensetzung am Eintritt'] == df.loc[i - 1,'Zusammensetzung am Eintritt']:
                a = df.loc[i-1,'Massenstrom Messstrecke']
                b = df.loc[i-1,'Zusammensetzung am Eintritt']
                lab = str(b) + ',' + str(a)
                #print('Coordinate Values: ',x_coordinate,y_coordinate)
                ax.plot(x_coordinate,y_coordinate,linestyle = "",marker=mi, color=ci,label = lab) 
                x_coordinate.clear()
                y_coordinate.clear()
                x_coordinate.append(df.loc[i,'Logarithmische Temperaturdifferenz in K'])
                y_coordinate.append(df.loc[i,'alpha_kondensation in W/m2K'])
                m += 1
        if i > 1 and i == df.shape[0]: #This condition ensure that the last file's datapoints are also plotted as the normal ones depend on change in dot_m_ms and y_ein values and the last datapoint doesn't
            mi = markers[m]
            ci = colors[c]
            a = df.loc[i-1,'Massenstrom Messstrecke']
            b = df.loc[i-1,'Zusammensetzung am Eintritt']
            lab = str(b) + ',' + str(a)
            #print('Coordinate Values: ',x_coordinate,y_coordinate)
            ax.plot(x_coordinate,y_coordinate,linestyle = "",marker=mi, color=ci,label = lab) 
            x_coordinate.clear()
            y_coordinate.clear()
            x_coordinate.append(df.loc[i,'Logarithmische Temperaturdifferenz in K'])
            y_coordinate.append(df.loc[i,'alpha_kondensation in W/m2K'])
            m += 1
    #'legend' legt die Position der Legende in dem Graphen fest
    #Funktion 'savefig' konventiert den Graphen in ein png-Bild --> Namensgebung des ersten Graphens 
    legend = plt.legend(loc='center left', bbox_to_anchor=(1, 0.5))
    plt.savefig('plot1.png', bbox_inches = 'tight')
    plt.show()
    
    
    
    
    x_coordinate = []
    y_coordinate = []
    m = 0
    c = 0
    
    
    
    
    #Ab hier beginnt der Code für die zweite Abbildung und verwendet genau dieselben Funktionen
    #Plot 2: Q_dot_tf über dT_log
    f2 = plt.figure(2)
    
    plt.style.use('bmh') 
    font = {'family' : 'normal',
            'weight' : 'normal', 
            'size'   : 18}
    plt.rc('font', **font)
    
    fig, ax = plt.subplots(figsize=(10, 6))
    
    plt.grid(True)
    ax.set_xlabel("$\mathrm{d}T_{\mathrm{log}}$ in $\mathrm{K}$")   #dT_log -> X-Achse
    ax.set_ylabel("$\.\mathrm{Q}_{\mathrm{tf}}$ in $\mathrm{W}$")   #Q_dot_tf -> y-Achse
    ax.set_xlim([0, 18])
    ax.set_ylim([0,800])
    ax.set_facecolor('white')
    
    for i in df.index:
        if i == 1:
            x_coordinate.append(df.loc[i,'Logarithmische Temperaturdifferenz in K'])
            y_coordinate.append(df.loc[i,'Wärmestom in W'])
        if i > 1 and df.loc[i,'Massenstrom Messstrecke'] == df.loc[i - 1,'Massenstrom Messstrecke']:
            x_coordinate.append(df.loc[i,'Logarithmische Temperaturdifferenz in K'])
            y_coordinate.append(df.loc[i,'Wärmestom in W'])
        if i > 1 and df.loc[i,'Massenstrom Messstrecke'] != df.loc[i - 1,'Massenstrom Messstrecke']:
            mi = markers[m]
            ci = colors[c]
            if df.loc[i,'Zusammensetzung am Eintritt'] != df.loc[i - 1,'Zusammensetzung am Eintritt']:
                a = df.loc[i-1,'Massenstrom Messstrecke']
                b = df.loc[i-1,'Zusammensetzung am Eintritt']
                lab = str(b) + ',' + str(a)

                x_coordinate.sort()
                
                ax.plot(x_coordinate,y_coordinate,color=ci,linestyle = "",marker=mi, markersize=5, label = lab) 
                
                x_coordinate.clear()
                y_coordinate.clear()
                
                x_coordinate.append(df.loc[i,'Logarithmische Temperaturdifferenz in K'])
                y_coordinate.append(df.loc[i,'Wärmestom in W'])
                m += 1
                c += 1
                m = 0
                
            elif df.loc[i,'Zusammensetzung am Eintritt'] == df.loc[i - 1,'Zusammensetzung am Eintritt']:
                a = df.loc[i-1,'Massenstrom Messstrecke']
                b = df.loc[i-1,'Zusammensetzung am Eintritt']
                lab = str(b) + ',' + str(a)
                ax.plot(x_coordinate,y_coordinate,linestyle = "",marker=mi, color=ci,label = lab) 
                x_coordinate.clear()
                y_coordinate.clear()
                x_coordinate.append(df.loc[i,'Logarithmische Temperaturdifferenz in K'])
                y_coordinate.append(df.loc[i,'Wärmestom in W'])
                m += 1
            
        if i > 1 and i == df.shape[0]:
            mi = markers[m]
            ci = colors[c]
            a = df.loc[i-1,'Massenstrom Messstrecke']
            b = df.loc[i-1,'Zusammensetzung am Eintritt']
            lab = str(b) + ',' + str(a)
            #print('Coordinate Values: ',x_coordinate,y_coordinate)
            ax.plot(x_coordinate,y_coordinate,linestyle = "",marker=mi, color=ci,label = lab) 
            x_coordinate.clear()
            y_coordinate.clear()
            x_coordinate.append(df.loc[i,'Logarithmische Temperaturdifferenz in K'])
            y_coordinate.append(df.loc[i,'alpha_kondensation in W/m2K'])
            m += 1
    #Funktion 'savefig' konventiert den Graphen in ein png-Bild --> Namensgebung des zweiten Graphens         
    legend = plt.legend(loc='center left', bbox_to_anchor=(1, 0.5))
    plt.savefig('plot2.png', bbox_inches = 'tight')
    plt.show()
    
    
    

    
    
    