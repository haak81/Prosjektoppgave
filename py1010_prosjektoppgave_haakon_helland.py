# # Prosjektoppgaven PY1010
# ## Haakon Helland
# ### 2025 04 

##############################################################################

#Del a

#Leser inn filen 'support_uke_24.xlsx'

import pandas as pd

data_support_u24 = pd.read_excel('support_uke_24.xlsx', dtype = {'Ukedag':str, 'Klokkeslett': str, 'Varighet': str })

#Array med ukedag henvendelsen fant sted
u_dag = data_support_u24['Ukedag'].values

# Array med klokkeslett for henvendelsen
kl_slett = data_support_u24['Klokkeslett'].values 

# Array med samtalens varighet
varighet = data_support_u24['Varighet'].values 

# Array score på tilfredshet fra kunden
score = data_support_u24['Tilfredshet'].values 

##############################################################################

#Del b

import numpy as np
import matplotlib.pyplot as plt

#Gjør om til array streng
u_dag = u_dag.astype(str) 

#Summerer antall henvendelser per dag
ant_henv_man = sum(np.char.count(u_dag,'Mandag'))

ant_henv_tir = sum(np.char.count(u_dag,'Tirsdag'))

ant_henv_ons = sum(np.char.count(u_dag,'Onsdag'))

ant_henv_tors = sum(np.char.count(u_dag,'Torsdag'))

ant_henv_fre = sum(np.char.count(u_dag,'Fredag'))

#liste med antall henvendelser per dag
ant_henv_dag_uke24 = [ant_henv_man, ant_henv_tir, ant_henv_ons, ant_henv_tors, ant_henv_fre]

#Lager stolpediagram med antall henvendelser per dag

#Lagel dag
labels_dag = ['Mandag', 'Tirsdag', 'Onsdag', 'Torsdag', 'Fredag']

#Farge på søylene
colors = ['brown', 'blue', 'yellow', 'red', 'orange']

#selve diagrammet
plt.bar(labels_dag, ant_henv_dag_uke24, color=colors)

#Tittel dag
plt.xlabel("Dag")

#Tittel ant. henvendelser
plt.ylabel("Antall henvendelser")

#Tittel diagram
plt.title("Antall telefonhenvendelser per dag uke 24")

##############################################################################

#Del c
#import time
#from datetime import datetime
import re 

#Gjør om til array streng
varighet = varighet.astype(str) 

#Tom liste som skal fylles med antall minutter for hver samtale
varighet_min = []

#Legger til minutter fra varighet, går gjennom listen og henter ut minuttene

for n in range(0, len(varighet)):
               temp = int(re.search(':(.+?):', varighet[n]).group(1))
               varighet_min.append(temp)
        
#Gjør om liste til array        
varighet_min = np.array(varighet_min) 

#Tom liste som skal fylles med sekunder for hver samtale
varighet_sek = [] 

#Legger til sekunder fra varighet, går gjennom listen og henter ut sekundene

for n in range(0, len(varighet)):
               temp = int(re.search(r'..$', varighet[n]).group(0))
               varighet_sek.append(temp)

#Gjør om liste til array        
varighet_sek = np.array(varighet_sek) 

#Finner maks antall minutter

#Lager en funskjon som finner maks antall minutter i array minutter 
def finn_max_minutt(varighet_min, n):
    
    max_min = varighet_min[0]
    
    for i in range(1, n):
        if varighet_min[i] > max_min:
            max_min = varighet_min[i]
    return max_min
    
n = len(varighet_min)
maks_varighet_min = finn_max_minutt(varighet_min, n)

#Lager en funksjon som finner maks antall sekunder til maks antall minutter

def finn_max_sekunder(varighet_sek, n):
    
    max_sek = varighet_sek[0]
    index = np.where(varighet_min == maks_varighet_min)[0]
    
    #Hvis det er kun en registreritn med maks minutt, så hentes sekund herfra
    if len(index) == 1:
        max_sek = maks_varighet_min[index[0]]
    #Hvis det er flere registreringer med maks minutt, så velges den med maks sekund
    else:
    
        for i in index:
            if varighet_sek[i] > max_sek:
                max_sek = varighet_sek[i]
    
    return max_sek

n = len(varighet_sek)
maks_varighet_sek = finn_max_sekunder(varighet_sek, n)
    

#Funksjon som finner minimum antall minutter 
def finn_min_minutt(varighet_min, n):
    
    min_min = varighet_min[0]
    
    for i in range(1, n):
        if varighet_min[i] < min_min:
            min_min = varighet_min[i]
    return min_min
    
n = len(varighet_min)
min_varighet_min = finn_min_minutt(varighet_min, n)

#Funksjon som finner minimum antall sekunder til minimum antall minutter

def finn_min_sekunder(varighet_sek, n):
    
    min_sek = varighet_sek[0]
    index = np.where(varighet_min == min_varighet_min)[0]
    
    #Hvis det er bare en registrerting med minimum minutt, så hentes sekund herfra
    if len(index) == 1:
        min_sek = varighet_sek[index[0]]
    #Hvis det er flere registreringer med minimum minutt, så velges den med minimum sekund    
    else:  
   
        for i in index:
            if varighet_sek[i] < min_sek:
                min_sek = varighet_sek[i]
    
    return min_sek

n = len(varighet_sek)
min_varighet_sek = finn_min_sekunder(varighet_sek, n)

print('Korteste telefonsamtale i uke 24 var',min_varighet_min,'minutt og',min_varighet_sek,'sekunder' )

print('Lengste telefonsamtale i uke 24 var',maks_varighet_min,'minutt og',maks_varighet_sek,'sekunder' )

###############################################################################

#Del d
#Skal regne ut gjennomsnittlig samtaletid

#Gjør om minutt til sekunder
varighet_min_tilsek = np.multiply(varighet_min,60)

#Legger sammen for å få totalt antall sekudner
varighet_totalt_sekunder = np.add(varighet_min_tilsek, varighet_sek)

#Finner gjennomsnittlig samtaletid i sekunder
varighet_gjennomsnitt_sekund = np.mean(varighet_totalt_sekunder)

#Gjør om gjennomsnittlig samtaletid i sekunder til minutter (float)
varighet_gjennomsnitt_minutt = np.divide(varighet_gjennomsnitt_sekund,60)

#lager en funksjon som tar inn float (minutter) og gjør om til minutt og sekund

def float_til_min_sek(fl):
    
    fl_int = int(fl)
    fl_dec = (fl - fl_int)
    minutt = fl_int
    sekund = int(fl_dec*60)
    
    return minutt, sekund

minutt,sekund = float_til_min_sek(varighet_gjennomsnitt_minutt)

print('Gjennomsnittlig samtaletid i uke 24 var', minutt, 'minutt og', sekund, 'sekund')
    
###############################################################################

#Del e
#Skal finne antall henvendelser per tidsrom
#Må ta ut time fra klokkeslett

#Gjør om til array streng
kl_slett = kl_slett.astype(str) 

#henter ut time tt
kl_slett_tt = [] #Tom liste

#Legger til timer fra klokkeslett
#Tar de to første tallene i strengen- gjør det om til integer
for n in range(0, len(kl_slett)):
               temp = int(re.search(r'^\d{2}', kl_slett[n]).group(0))
               kl_slett_tt.append(temp)

#Gjør om liste til array        
kl_slett_tt = np.array(kl_slett_tt) 

#Teller opp til de forskjellige gruppene

kl_08_10 = np.count_nonzero( (kl_slett_tt >= 8) & (kl_slett_tt < 10) ) 

kl_10_12 = np.count_nonzero( (kl_slett_tt >= 10) & (kl_slett_tt < 12) )

kl_12_14 = np.count_nonzero( (kl_slett_tt >= 12) & (kl_slett_tt < 14) )

kl_14_16 = np.count_nonzero( (kl_slett_tt >= 14) & (kl_slett_tt < 16) )

#Lager liste med de forskjellige tidsgruppene
kl_slett_grupper = [kl_08_10, kl_10_12, kl_12_14, kl_14_16]

#Labels for tidsgruppene
kl_slett_labels = ['kl 08-10', 'kl 10-12', 'kl 12-14', 'kl 14-16']

#Farger i kakediagrammet
kl_slett_colors = ['orange', 'blue', 'yellow', 'red',]

#den største kaken skal stå litt ut
skille_ut = [0.03, 0, 0, 0]

#Selve diagrammet
plt.pie(kl_slett_grupper, labels=kl_slett_labels, colors=kl_slett_colors, startangle = 90, explode = skille_ut)
#Tittel
plt.title('Antall henvendelser per klokkeslett i uke 24')

###############################################################################

#Del f

#Teller opp antall henvendelser i forskjellige score-grupper

score_1_6 = np.count_nonzero( (score >= 1) & (score < 7) ) 

score_7_8 = np.count_nonzero( (score >= 7) & (score < 9) ) 

score_9_10 = np.count_nonzero( (score >= 9) & (score <= 10) ) 

#Regner ut % for score-gruppene
import math as m

#totalt antall som har gitt score
totalt_gitt_score = [score_1_6, score_7_8, score_9_10]

#Lager en funksjon som regner ut NPS
#Input er antall i de forskjellige scoregruppene
def score_NPS(s_1_6, s_7_8, s_9_10):
    
    #totalt antall
    s_tot = s_1_6 + s_7_8 + s_9_10
    
    p_1_6 = ( (s_1_6/s_tot)*100 )
    p_7_8 = ( (s_7_8/s_tot)*100 )
    p_9_10 = ( (s_9_10/s_tot)*100 )
    
    NPS = (p_9_10 - p_1_6)
    
    return NPS

NPS = score_NPS(score_1_6, score_7_8, score_9_10)

print('Supportavdelingens NPS i uke 24 var', round(NPS,1),'%') 
    




   



