import pandas as pd
from datetime import timedelta
from openpyxl.worksheet.table import Table, TableStyleInfo


#Uus fail
failinimi = "andmed.xlsx"

#Loeme andmed lähtefailist
f = pd.read_excel("andmed2.xlsx", sheet_name="Sheet1")
f2 = pd.read_excel("andmed2.xlsx", sheet_name="Sheet2")

#Teeme Pandas exceli, millesse kirjutada
writer = pd.ExcelWriter(failinimi, engine="openpyxl")


#Paneme kuupäevad õigesse formaati
f["Kuupäev"] = pd.to_datetime(f["Kuupäev"], dayfirst = True, format="%d-%m-%Y")
f['Kuupäev'] = f["Kuupäev"].dt.date

#Sorteerime kuupäeva järgi
f = f.sort_values("Kuupäev")


#Format as table funktsioon
def tabel1(lehenimi, vahemik, kujundus, tabelinimi):
    tab = Table(displayName=tabelinimi, ref=vahemik)

    style = TableStyleInfo(
        name=kujundus,
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False)

    tab.tableStyleInfo = style
    writer.sheets[lehenimi].add_table(tab)


#Kokkuvõtte arvutamine
df1 = pd.DataFrame({
    "Makseviis" : ["Sularahas","Pangakontolt","Kokku"],
    "Väljaminekud" : [pd.DataFrame.sum(f[(f["Makseviis"]=="s") & (f["Summa"]<0)]["Summa"]),
                      pd.DataFrame.sum(f[(f["Makseviis"]=="k") & (f["Summa"]<0)]["Summa"]),
                      pd.DataFrame.sum(f[(f["Makseviis"]=="s") & (f["Summa"]<0)]["Summa"])+ pd.DataFrame.sum(f[(f["Makseviis"]=="k") & (f["Summa"]<0)]["Summa"])],
    "Sissetulekud" : [pd.DataFrame.sum(f[(f["Makseviis"]=="s") & (f["Summa"]>0)]["Summa"]),
                      pd.DataFrame.sum(f[(f["Makseviis"]=="k") & (f["Summa"]>0)]["Summa"]),
                      pd.DataFrame.sum(f[(f["Makseviis"]=="s") & (f["Summa"]>0)]["Summa"])+ pd.DataFrame.sum(f[(f["Makseviis"]=="k") & (f["Summa"]>0)]["Summa"])]
    })

df2 = pd.DataFrame({
    "Vabad vahendid": [f2.at[0,"Summa"]+df1.at[0,df1.columns[1]]+df1.at[0,df1.columns[2]],
                       f2.at[1,"Summa"]+df1.at[1,df1.columns[1]]+df1.at[1,df1.columns[2]],
                       f2.at[2,"Summa"]+df1.at[2,df1.columns[1]]+df1.at[2,df1.columns[2]]]
     })

#Lisame vabad vahendid ülevaatesse
df_vabad = pd.DataFrame({
    "Makseviis": ["Sularahas","Pangakontol","Kokku"],
    "Vabad vahendid": [f2.at[0,"Summa"]+df1.at[0,df1.columns[1]]+df1.at[0,df1.columns[2]],
                       f2.at[1,"Summa"]+df1.at[1,df1.columns[1]]+df1.at[1,df1.columns[2]],
                       f2.at[2,"Summa"]+df1.at[2,df1.columns[1]]+df1.at[2,df1.columns[2]]]
     })
df_vabad.to_excel(writer, "Ülevaade", startcol=6, index=False)


#Loome töötlemiseks uue df, indeksiks kuupäev
df=f.set_index('Kuupäev')

#Andmed aastate kaupa 
summa = []
aastad = []
for i in range(2021, 2015, -1):
    
    algus = pd.to_datetime(str(i)).date()
    lõpp = pd.to_datetime(str(i+1)).date()
    
    if len(df.loc[algus:lõpp]) > 0:
        aastad.append(str(i))

#Andmed kuude kaupa
nihe = 0
e = 5
for aasta in aastad:
    df=f.set_index('Kuupäev')
    kuud1 = ["Jaanuar", "Veebruar", "Märts", "Aprill", "Mai", "Juuni", "Juuli", "August", "September", "Oktoober", "November", "Detsember"]
    kuud = []
    kuude_summa = []
    kuude_sisse = []
    kuude_välja = []
    
    for i in range(1,13):
        if i == 12:
            algus = pd.to_datetime(str(aasta)+"-"+str(i)).date()
            lõpp = pd.to_datetime(str(int(aasta)+1)+"-1-1").date()-timedelta(days=1)
        else:
            algus = pd.to_datetime(str(aasta)+"-"+str(i)).date()
            lõpp = pd.to_datetime(str(aasta)+"-"+str(i+1)).date()-timedelta(days=1)
        dfa = df.loc[algus:lõpp]
        if len(dfa) > 0:
            kuude_summa.append(pd.DataFrame.sum(dfa["Summa"]))
            kuude_sisse.append(pd.DataFrame.sum(dfa[(dfa["Summa"]>0)]["Summa"]))
            kuude_välja.append(pd.DataFrame.sum(dfa[(dfa["Summa"]<0)]["Summa"]))
            kuud.append(kuud1[i-1])
    
    kuude_summa.append(sum(kuude_summa))
    kuude_sisse.append(sum(kuude_sisse))
    kuude_välja.append(sum(kuude_välja))
    kuud.append("Kokku")
    
    df_aastad = pd.DataFrame({
        aasta : kuud,
        "Väljaminekud" : kuude_välja,
        "Sissetulekud" : kuude_sisse,
        "Kokku" : kuude_summa
        })
    df_aastad.to_excel(writer, "Ülevaade", startrow=nihe, index=False)
    
    #tabeli kujundus
    tabel1('Ülevaade', 'A'+str(nihe+1)+':D'+str(nihe+len(kuude_summa)+1), "TableStyleLight15", 'tabel'+str(e) )
    e += 1
    nihe += 2+len(df_aastad)


#Paneme kokkuvõtte excelisse
f.to_excel(writer, "Andmed", index=False)
df1.to_excel(writer, "Andmed", startcol=7, index=False)
df2.to_excel(writer, "Andmed", startcol=10, index=False)


#Teeme sheeti, kuhu kopeeritud sularahatehingute info
makseviis = f[f["Makseviis"]=="s"]
makseviis.to_excel(writer, "Sularaha", index=False)
df_sularaha = pd.DataFrame({
    "Väljaminekud" : df1.at[0,df1.columns[1]],
    "Sissetulekud" : df1.at[0,df1.columns[2]],
    "Kokku" : df1.at[0,df1.columns[1]]+df1.at[0,df1.columns[2]],
    "Vabad vahendid" : [df2.at[0,df2.columns[0]]]
    })
df_sularaha.to_excel(writer, "Sularaha", startcol=7, index=False)

#Teeme tabelid
w = str(len(makseviis)+1)
tabel1('Sularaha', 'A1:F'+w, "TableStyleLight20", 'tabel100')
tabel1('Sularaha', 'H1:K2', "TableStyleLight20", 'tabel101')

#Teeme sheeti, kuhu kopeeritud pangatehingute info
f[f["Makseviis"]=="k"].to_excel(writer, "Konto", index=False)
df_konto = pd.DataFrame({
    "Väljaminekud" : df1.at[1,df1.columns[1]],
    "Sissetulekud" : df1.at[1,df1.columns[2]],
    "Kokku" : df1.at[1,df1.columns[1]]+df1.at[1,df1.columns[2]],
    "Vabad vahendid" : [df2.at[1,df2.columns[0]]]
    })
df_konto.to_excel(writer, "Konto", startcol=7, index=False)

#Teeme tabelid
w = str(len(f[f["Makseviis"]=="k"])+1)
tabel1('Konto', 'A1:F'+w, "TableStyleLight19", 'tabel102')
tabel1('Konto', 'H1:K2', "TableStyleLight19", 'tabel103')


#Teeme listi siltidest
sildid = []
for silt in f["Silt"]:
    if silt not in sildid:
        sildid.append(silt)

#Teeme iga sildi jaoks eraldi sheeti, kus selle info ja kokku summa ning lisasin sheeti koos kogu summaga esimesele sheetile
kulutulu = []
for sheet in sildid:
    with writer as writer:
        tabel = f[f["Silt"]==sheet]
        tabel.to_excel(writer, sheet, index=False)
        sum1 = pd.DataFrame({
            "Kokku":[pd.DataFrame.sum((f[f["Silt"]==sheet]["Summa"]))]
            })
        kulutulu.append(pd.DataFrame.sum((f[f["Silt"]==sheet]["Summa"])))
        sum1.to_excel(writer, sheet, index=False, startcol=(tabel.count(axis="columns").max()))
        tabel1(sheet, 'A1:G'+str(len(tabel)+1), "TableStyleLight15", 'tabel'+str(e))
        e += 1

#Sildid tabelina esimesele sheetile
df3 = pd.DataFrame({
    'Sildid':sildid, 'Kulu/tulu':kulutulu})

df3.to_excel(writer, 'Andmed', index=False, startcol=12)


#Kujundame lehe 'Andmed' tabelid
w = str(len(f)+1)
tabel1('Andmed', 'A1:F'+w, "TableStyleLight15", 'tabel1')
tabel1('Andmed', 'H1:K4', "TableStyleLight16", 'tabel2')
w = str(len(df3)+1)
tabel1('Andmed', 'M1:N'+w, "TableStyleLight16", 'tabel3')

#Kujundmae lehe 'Ülevaade' tabelid
tabel1('Ülevaade', 'G1:H4', "TableStyleLight18", 'tabel4')


#Muudame lehtedel veergude laiuseid
sheets = ['Andmed','Sularaha','Konto']

for sh in sheets:
    writer.sheets[sh].column_dimensions['H'].width = 13
    writer.sheets[sh].column_dimensions['I'].width = 13
    writer.sheets[sh].column_dimensions['J'].width = 13
    writer.sheets[sh].column_dimensions['K'].width = 15

sheets.extend(sildid)

for el in sheets:
    writer.sheets[el].column_dimensions['A'].width = 12
    writer.sheets[el].column_dimensions['B'].width = 15
    writer.sheets[el].column_dimensions['C'].width = 15
    writer.sheets[el].column_dimensions['D'].width = 20

writer.sheets["Ülevaade"].column_dimensions['A'].width = 15
writer.sheets["Ülevaade"].column_dimensions['B'].width = 13
writer.sheets["Ülevaade"].column_dimensions['C'].width = 13
writer.sheets["Ülevaade"].column_dimensions['G'].width = 13
writer.sheets["Ülevaade"].column_dimensions['H'].width = 15
    
writer.sheets['Andmed'].column_dimensions['M'].width = 15

writer.save()
