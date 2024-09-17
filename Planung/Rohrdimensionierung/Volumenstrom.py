
#%%
import numpy as np
from IPython.display import display, Math
import pandas as pd
import matplotlib.pyplot as plt

filename = '/Users/benjaminbartels/VS Code/Python/lib/stylesheets/20240402_ci-style.py'
exec(compile(open(filename, "rb").read(), filename, 'exec'))
# %%
#Formel Volumenstrom 
# 
print("Formel für den Volumenstrom:")
display(Math(r"$V=\frac{\dot Q} {\rho \cdot c \cdot \Delta\vartheta}$"))
print("Q: Wärmestrom (Heizlast) in W")
print("p: Dichte c: Wärmeleitkoeffizient (p*c Wasser = 0.87)")
print("delta t: Temperaturspreizung in K")
# %%
#V = Q/c*p*dt
Q = 0
#C*p
x = 0.86
dt = 15
# %%
def getVolumenstrom(heizlast):
    #in m3/h
    hl = heizlast * 1000
    V = hl/dt * x

    return V/1000
def getStroemungsgeschwindigkeit(Volumenstrom: pd.Series, durchmesser: float):
    p = np.pi/4
    d = np.power((durchmesser/1000),2)
    S = (Volumenstrom/3600)/(p*d)
    S.columns = [durchmesser]
    return S
def getSpezDruckverlust(Stroemungsgeschwindigkeit: pd.Series, durchmesser: float):
    
    lbd = 0.0005
    rho = 996.511
    stroe = np.power(Stroemungsgeschwindigkeit,2) 
    d = durchmesser/1000
    D = (lbd * rho * stroe)/(d*2)
    return D


def getKurzBelastung(Stroemungsgeschwindigkeit):
    # wir wollen die Gesamtstunfden an kurzbelastung und die längste zusammenhängende kette und villeicht das maximum
    r = pd.DataFrame()

    r["Maximale Belastung"] = Stroemungsgeschwindigkeit.max()

    b = Stroemungsgeschwindigkeit > 1.5
    r["Stunden über Belastung"] = b.sum()
    x = b.iloc[:, 0].ne(b.iloc[:, 0].shift()).cumsum()

    z = b.iloc[:, 0].groupby(x).transform('size') * np.where(b.iloc[:, 0], 1, -1)
    t = z.max()
    if(t!=-8760):
        r["Maximaler Zeitraum über Belastung"] = z.max()
    else:
        r["Maximaler Zeitraum über Belastung"] = 0
        
    r.insert(0, "D_i", Stroemungsgeschwindigkeit.columns[0])
    return r.reset_index(drop=True)


def filterGrenzwerte(Variante_Stro,upper,lower):
    data = Variante_Stro 
    x = data.filter(like='Strö', axis=1)[data.filter(like='Strö', axis=1) < upper]
    x = x.fillna(9999)
    x = x.filter(like='Strö', axis=1)[x.filter(like='Strö', axis=1) >= lower]
    x = x.fillna("Passt!")
    x = x.replace(9999,"Boom!")
    
    data.update(x)
    return data
# %%
print(getVolumenstrom(5000))
# %%

# %%
data = pd.read_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Import/Heizlast.xlsx", sheet_name=1, usecols=[2], skiprows=16)


print(data)
# %%
Vol = getVolumenstrom(data)
print(Vol)
# %%

# %%

# %%
Str50 = getStroemungsgeschwindigkeit(Vol, 51.4)
Str65 = getStroemungsgeschwindigkeit(Vol, 61.4)
Str40 = getStroemungsgeschwindigkeit(Vol, 40.8)
#plt.plot(Str)
# %%
#Dr = getSpezDruckverlust(Str, 51.4)
#plt.plot(Dr)
# %%
fig, (ax,ax1,ax2) = plt.subplots(3,1, figsize = (10,18))

ax.plot(Vol, label = "Volumenstrom")

ax.axhline(4.5, c = color5, linewidth = 2.5, label = "DN40")
ax.axhline(3, c = color4, linewidth = 2.5, label = "DN32")
ax.axhline(1.8, c = color3, linewidth = 2.5, label = "DN26")
ax.set_xlabel("Stunde des Jahres", size = 18)
ax.set_ylabel("Volumenstrom in l/s", size = 22)
ax.tick_params(axis='both', which='both', labelsize=15,width = 3,length = 5)
ax.legend()

ax1.plot(Str40, label = "Strömungsgeschwindigkeit (DN40)")
ax1.plot(Str50, label = "Strömungsgeschwindigkeit (DN50)")
ax1.plot(Str65, label = "Strömungsgeschwindigkeit (DN65)")

#ax1.axhline(1.6, c = color3, linewidth = 2.5, label = "DN65")
#ax1.axhline(1.4, c = color2, linewidth = 2.5, label = "DN50")
#ax1.axhline(1.2, c = color1, linewidth = 2.5, label = "DN40")
ax1.set_xlabel("Stunde des Jahres", size = 18)
ax1.set_ylabel("Strömungsgeschwindigkeit in m/s", size = 22)
ax1.tick_params(axis='both', which='both', labelsize=15,width = 3,length = 5)
ax1.legend()

#ax2.plot(Dr, label = "Volumenstrom")

ax2.axhline(250, c = color1, linewidth = 2.5, label = "DN40")
#ax2.axhline(3, c = color4, linewidth = 2.5, label = "DN32")
#ax2.axhline(1.8, c = color3, linewidth = 2.5, label = "DN26")
ax2.set_xlabel("Stunde des Jahres", size = 18)
ax2.set_ylabel("spez. Druckverlust in Pa/m", size = 22)
ax2.tick_params(axis='both', which='both', labelsize=15,width = 3,length = 5)
ax2.legend()
# %%
fig, ax = plt.subplots(1, figsize = (10,5))
ax.plot(Str50, label = "Strömungsgeschwindigkeit (DN50)")
ax.axhline(1.4, c = color4, linewidth = 2.5, label = "DN50")
ax.set_xlabel("Stunde des Jahres", size = 18)
ax.set_ylabel("Strömungsgeschwindigkeit in m/s", size = 22)
ax.tick_params(axis='both', which='both', labelsize=15,width = 3,length = 5)
ax.legend()
# %%
Str = getStroemungsgeschwindigkeit(Vol, 51.4)
plt.plot(Str)
# %%
dr = getSpezDruckverlust(Str, 51.4)
plt.plot(dr)
# %%
#getStroemungsgeschwindigkeit()
# %%

# %%
'''print(Str40 > 1.5)
s = Str40 > 1.5

x = s['Strömungsgeschwindigkeit (m/s)'].ne(s['Strömungsgeschwindigkeit (m/s)'].shift()).cumsum()
print(x)
s['Count'] = s.groupby(x)['Strömungsgeschwindigkeit (m/s)'].transform('size') * np.where(s['Strömungsgeschwindigkeit (m/s)'], 1, -1)
print(s["Count"].max())
getKurzBelastung(Str40)'''
# %%

# %%
getKurzBelastung(Str40)

# %% [M]

# %%
#Volumenströme für Variante Wagen:
#Low
#A
A_l = 20+61+28+11+13+24+13
A_l_Vol = getVolumenstrom(A_l)
d1 = {"Strang":"A" , "Szenario":"Low", "Volumenstrom":A_l_Vol }

#B
B_l = 28+61 #Low_COW+6
B_l_Vol = getVolumenstrom(B_l)
d2 = {"Strang":"B" , "Szenario":"Low", "Volumenstrom":B_l_Vol }

#C
C_l = 11+13+20+24+13 #Low_HH+2+3+4+5
C_l_Vol = getVolumenstrom(C_l)
d3 = {"Strang":"C" , "Szenario":"Low", "Volumenstrom":C_l_Vol }


#High
#A
A_h = 31+90+45+17+20+38+20
A_h_Vol = getVolumenstrom(A_h)
d4 = {"Strang":"A" , "Szenario":"High", "Volumenstrom":A_h_Vol}

#B
B_h = 45+90 #Hgh_COW+6
B_h_Vol = getVolumenstrom(B_h)
d5 = {"Strang":"B" , "Szenario":"High", "Volumenstrom":B_h_Vol }

#C
C_h = 17+20+31+38+20 #High_HH+2+3+4+5
C_h_Vol = getVolumenstrom(C_h)
d6 = {"Strang":"C" , "Szenario":"High", "Volumenstrom":C_h_Vol }

Wagen_Vol = pd.DataFrame.from_dict((d1,d2,d3,d4,d5,d6))



print(Wagen_Vol)
# %%
#Volumenströme für Variante Schlange:
#Low
#A
A_l = 20+61+28+11+13+24+13
A_l_Vol = getVolumenstrom(A_l)
Sd1 = {"Strang":"A" , "Szenario":"Low", "Volumenstrom":A_l_Vol }

#B
SB_l = 20+11+13+24+13 #Low_HH+2+3+4+5
SB_l_Vol = getVolumenstrom(SB_l)
Sd2 = {"Strang":"B" , "Szenario":"Low", "Volumenstrom":SB_l_Vol }


#High
#A
SA_h = 31+90+45+17+20+38+20
SA_h_Vol = getVolumenstrom(SA_h)
Sd4 = {"Strang":"A" , "Szenario":"High", "Volumenstrom":SA_h_Vol}

#B
SB_h = 31+17+20+38+20 #High_HH+2+3+4+5
SB_h_Vol = getVolumenstrom(SB_h)
Sd5 = {"Strang":"B" , "Szenario":"High", "Volumenstrom":SB_h_Vol }

Schlange_Vol = pd.DataFrame.from_dict((Sd1,Sd2,Sd4,Sd5))



print(Schlange_Vol)
# %%
#Volumenströme für Variante Wagen - NPRO:
#Low
#A
data = pd.read_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Import/Low_Alles.xlsx",sheet_name=1, usecols=[2],skiprows=(12),nrows=1)
A_l = data.values[0][0]
A_l_Vol = getVolumenstrom(A_l)
d1 = {"Strang":"A" , "Szenario":"Low", "Volumenstrom":A_l_Vol }

#B
data = pd.read_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Import/Low_COW+6.xlsx",sheet_name=1, usecols=[2],skiprows=(12),nrows=1)
B_l = data.values[0][0] #Low_COW+6
B_l_Vol = getVolumenstrom(B_l)
d2 = {"Strang":"B" , "Szenario":"Low", "Volumenstrom":B_l_Vol }

#C
data = pd.read_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Import/Low_HH+2+3+4+5.xlsx",sheet_name=1, usecols=[2],skiprows=(12),nrows=1)
C_l = data.values[0][0] #Low_HH+2+3+4+5
C_l_Vol = getVolumenstrom(C_l)
d3 = {"Strang":"C" , "Szenario":"Low", "Volumenstrom":C_l_Vol }


#High
#A
data = pd.read_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Import/High_Alles.xlsx",sheet_name=1, usecols=[2],skiprows=(12),nrows=1)
A_h = data.values[0][0]
A_h_Vol = getVolumenstrom(A_h)
d4 = {"Strang":"A" , "Szenario":"High", "Volumenstrom":A_h_Vol}

#B
data = pd.read_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Import/High_COW+6.xlsx",sheet_name=1, usecols=[2],skiprows=(12),nrows=1)
B_h = data.values[0][0] #Hgh_COW+6
B_h_Vol = getVolumenstrom(B_h)
d5 = {"Strang":"B" , "Szenario":"High", "Volumenstrom":B_h_Vol }

#C
data = pd.read_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Import/High_HH+2+3+4+5.xlsx",sheet_name=1, usecols=[2],skiprows=(12),nrows=1)
C_h = data.values[0][0] #High_HH+2+3+4+5
C_h_Vol = getVolumenstrom(C_h)
d6 = {"Strang":"C" , "Szenario":"High", "Volumenstrom":C_h_Vol }

Wagen_Vol = pd.DataFrame.from_dict((d1,d2,d3,d4,d5,d6))



print(Wagen_Vol)

# %%
#Volumenströme für Variante Schlange - NPRO:
#Low
#A
data = pd.read_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Import/Low_Alles.xlsx",sheet_name=1, usecols=[2],skiprows=(12),nrows=1)
A_l = data.values[0][0]
A_l_Vol = getVolumenstrom(A_l)
Sd1 = {"Strang":"A" , "Szenario":"Low", "Volumenstrom":A_l_Vol }

#B
data = pd.read_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Import/Low_HH+2+3+4+5.xlsx",sheet_name=1, usecols=[2],skiprows=(12),nrows=1)
SB_l = data.values[0][0] #Low_HH+2+3+4+5
SB_l_Vol = getVolumenstrom(SB_l)
Sd2 = {"Strang":"B" , "Szenario":"Low", "Volumenstrom":SB_l_Vol }


#High
#A
data = pd.read_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Import/High_Alles.xlsx",sheet_name=1, usecols=[2],skiprows=(12),nrows=1)
SA_h = data.values[0][0]
SA_h_Vol = getVolumenstrom(SA_h)
Sd4 = {"Strang":"A" , "Szenario":"High", "Volumenstrom":SA_h_Vol}

#B
data = pd.read_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Import/High_HH+2+3+4+5.xlsx",sheet_name=1, usecols=[2],skiprows=(12),nrows=1)
SB_h = data.values[0][0] #High_HH+2+3+4+5
SB_h_Vol = getVolumenstrom(SB_h)
Sd5 = {"Strang":"B" , "Szenario":"High", "Volumenstrom":SB_h_Vol }

Schlange_Vol = pd.DataFrame.from_dict((Sd1,Sd2,Sd4,Sd5))



print(Schlange_Vol)
# %%
Wagen_Stro = Wagen_Vol
Schlange_Stro = Schlange_Vol
Wagen_Stro["Max Strö - DN20"] = getStroemungsgeschwindigkeit(Wagen_Stro["Volumenstrom"],20.4)
Schlange_Stro["Max Strö - DN20"] = getStroemungsgeschwindigkeit(Schlange_Stro["Volumenstrom"],20.4)

Wagen_Stro["Max Strö - DN25"] = getStroemungsgeschwindigkeit(Wagen_Stro["Volumenstrom"],26.2)
Schlange_Stro["Max Strö - DN25"] = getStroemungsgeschwindigkeit(Schlange_Stro["Volumenstrom"],26.2)

Wagen_Stro["Max Strö - DN32"] = getStroemungsgeschwindigkeit(Wagen_Stro["Volumenstrom"],32.6)
Schlange_Stro["Max Strö - DN32"] = getStroemungsgeschwindigkeit(Schlange_Stro["Volumenstrom"],32.6)

Wagen_Stro["Max Strö - DN40"] = getStroemungsgeschwindigkeit(Wagen_Stro["Volumenstrom"],40.8)
Schlange_Stro["Max Strö - DN40"] = getStroemungsgeschwindigkeit(Schlange_Stro["Volumenstrom"],40.8)

Wagen_Stro["Max Strö - DN50"] = getStroemungsgeschwindigkeit(Wagen_Stro["Volumenstrom"],51.4)
Schlange_Stro["Max Strö - DN50"] = getStroemungsgeschwindigkeit(Schlange_Stro["Volumenstrom"],51.4)

Wagen_Stro["Max Strö - DN65"] = getStroemungsgeschwindigkeit(Wagen_Stro["Volumenstrom"],61.4)
Schlange_Stro["Max Strö - DN65"] = getStroemungsgeschwindigkeit(Schlange_Stro["Volumenstrom"],61.4)

Wagen_Stro["Max Strö - DN80"] = getStroemungsgeschwindigkeit(Wagen_Stro["Volumenstrom"],73.6)
Schlange_Stro["Max Strö - DN80"] = getStroemungsgeschwindigkeit(Schlange_Stro["Volumenstrom"],73.6)

Wagen_Stro["Max Strö - DN100"] = getStroemungsgeschwindigkeit(Wagen_Stro["Volumenstrom"],81)
Schlange_Stro["Max Strö - DN100"] = getStroemungsgeschwindigkeit(Schlange_Stro["Volumenstrom"],81)

Wagen_Stro.applymap(lambda x: round(x, 2) if isinstance(x, (float, int)) else x)
Schlange_Stro.applymap(lambda x: round(x, 2) if isinstance(x, (float, int)) else x)


#Wagen_Stro.to_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Export/Wagen_Stro.xlsx")
#Schlange_Stro.to_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Export/Schlange_Stro.xlsx")

#NPROOOROOROROROR
Wagen_Stro.to_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Export/Wagen_Npro_Stro.xlsx")
Schlange_Stro.to_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Export/Schlange_Npro_Stro.xlsx")

#print(Wagen_Stro[Wagen_Stro.filter(like="Strömung")])

#print(Wagen_Stro.filter(like='Ström', axis=1).mask())

# %%
#Stroemungsgescvhwindigkeiten filtern
Fil_Wagen_Stro = filterGrenzwerte(Wagen_Stro,2,1.5)
Fil_Schlange_Stro = filterGrenzwerte(Schlange_Stro,2,1.5)

#Fil_Wagen_Stro.to_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Export/Filtered_Wagen_Stro.xlsx")
#Fil_Schlange_Stro.to_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Export/Filtered_Schlange_Stro.xlsx")

#NRPRPRPRPRPRPPRPRpoooooooooo
Fil_Wagen_Stro.to_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Export/Filtered_Wagen_Npro_Stro.xlsx")
Fil_Schlange_Stro.to_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Export/Filtered_Schlange_Npro_Stro.xlsx")


# %%
print(Fil_Wagen_Stro)
print(Fil_Wagen_Stro.applymap(lambda x: round(x, 2) if isinstance(x, (float, int)) else x))

# %%
Low_B_Wagen =pd.read_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Import/Low_COW+6.xlsx", sheet_name="Gesamtbedarf", usecols=[2], skiprows=16)
Low_C_Wagen =pd.read_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Import/Low_HH+2+3+4+5.xlsx", sheet_name="Gesamtbedarf", usecols=[2], skiprows=16)
High_B_Wagen = pd.read_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Import/High_COW+6.xlsx", sheet_name="Gesamtbedarf", usecols=[2], skiprows=16)
High_C_Wagen = pd.read_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Import/High_HH+2+3+4+5.xlsx", sheet_name="Gesamtbedarf", usecols=[2], skiprows=16)

Low_B_Schlange = pd.read_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Import/Low_HH+2+3+4+5.xlsx", sheet_name="Gesamtbedarf", usecols=[2], skiprows=16)
High_B_Schlange= pd.read_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Import/High_HH+2+3+4+5.xlsx", sheet_name="Gesamtbedarf", usecols=[2], skiprows=16)

# %%
V_Low_B_Wagen = getVolumenstrom(Low_B_Wagen)
V_Low_C_Wagen = getVolumenstrom(Low_C_Wagen)
V_High_B_Wagen = getVolumenstrom(High_B_Wagen)
V_High_C_Wagen = getVolumenstrom(High_C_Wagen)

V_Low_B_Schlange = getVolumenstrom(Low_B_Schlange)
V_High_B_Schlange = getVolumenstrom(High_B_Schlange)
# %%
S_Low_B_Wagen = getStroemungsgeschwindigkeit(V_Low_B_Wagen,32.6)
S_Low_C_Wagen = getStroemungsgeschwindigkeit(V_Low_C_Wagen,32.6)
S_High_B_Wagen = getStroemungsgeschwindigkeit(V_High_B_Wagen,40.8)
S_High_C_Wagen = getStroemungsgeschwindigkeit(V_High_C_Wagen, 40.8)

S_Low_B_Schlange = getStroemungsgeschwindigkeit(V_Low_B_Schlange, 32.6)
S_High_B_Schlange = getStroemungsgeschwindigkeit(V_High_B_Schlange, 40.8)
# %%
kb1 = getKurzBelastung(S_Low_B_Wagen)
kb1.insert(0,"Strang","B")
kb1.insert(0,"Variante", "Wagen")
kb1.insert(0,"Szenario","Low")

kb2 = getKurzBelastung(S_Low_C_Wagen)
kb2.insert(0,"Strang","C")
kb2.insert(0,"Variante", "Wagen")
kb2.insert(0,"Szenario","Low")

kb3 = getKurzBelastung(S_High_B_Wagen)
kb3.insert(0,"Strang","B")
kb3.insert(0,"Variante", "Wagen")
kb3.insert(0,"Szenario","High")

kb4 = getKurzBelastung(S_High_C_Wagen)
kb4.insert(0,"Strang","C")
kb4.insert(0,"Variante", "Wagen")
kb4.insert(0,"Szenario","High")


kb5 = getKurzBelastung(S_Low_B_Schlange)
kb5.insert(0,"Strang","B")
kb5.insert(0,"Variante", "Schlange")
kb5.insert(0,"Szenario","Low")

kb6 = getKurzBelastung(S_High_B_Schlange)
kb6.insert(0,"Strang","B")
kb6.insert(0,"Variante", "Schlange")
kb6.insert(0,"Szenario","High")

# %%
result = pd.concat((kb1,kb2,kb3,kb4,kb5,kb6)).reset_index()
print(result)
# %%
Low_A = pd.read_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Import/Low_Alles.xlsx",sheet_name=1, usecols=[2],skiprows=(16))
High_A = pd.read_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Import/High_Alles.xlsx",sheet_name=1, usecols=[2],skiprows=(16))

V_Low_A = getVolumenstrom(Low_A)
V_High_A = getVolumenstrom(High_A)

S_Low_A = getStroemungsgeschwindigkeit(V_Low_A, 40.8)
S_High_A = getStroemungsgeschwindigkeit(V_High_A, 51.4)


# %%
b1 = getKurzBelastung(S_Low_A)
b1.insert(0,"Strang","A")
b1.insert(0,"Variante", "Wagen/Schlange")
b1.insert(0,"Szenario","Low")

b2 = getKurzBelastung(S_High_A)
b2.insert(0,"Strang","A")
b2.insert(0,"Variante", "Wagen/Schlange")
b2.insert(0,"Szenario","High")

result = pd.concat((b1,b2)).reset_index()
print(result)

result.to_excel("/Users/benjaminbartels/VS Code/Python/Hofenergiesystem/Planung/Rohrdimensionierung/Export/kurzbelastung.xlsx")