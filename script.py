#import modules
import pandas as pd
import datetime
from datetime import timedelta
import requests
import json
from requests.auth import HTTPBasicAuth
import tkinter as tk
from tkinter import ttk

my_font = ("Helvetica", 12)

def popupmsg(msg):
	popup = tk.Tk()
	popup.wm_title("Got it turttle!")
	label = ttk.Label(popup, text=msg, font=my_font)
	label.pack(side="top", fill="x", padx=10, pady=20)
	#B1 = ttk.Button(popup, text="Okay", command = popup.destroy)
	#B1 = pack()
	popup.mainloop()

#display options
pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 20)
pd.set_option('display.width', 8000)

#Today's day value in order to compare with worksheet
now = pd.Timestamp(datetime.date.today())

#set file path
file_path = 'G:\My Drive\Formulario Abastecimiento\Abastecimiento.xlsm'
file_path2 = 'G:\My Drive\Planos\Criterios.xlsx' 


def readFiles(f1):  #That is if you can 
	try:
		#load Abastecimiento.xlsm
		f = pd.read_excel(f1, sheet_name = 'BD_Apr', dtype=str) #importing workbook
	except:
		return print('Tha fuck is happening with your paths???')
	return f

df = readFiles(file_path)
df2 = pd.read_excel(file_path2, sheet_name = 'Bodega', dtype=str) #Calling criterios to perform combine action
df3 = pd.read_excel(file_path2, sheet_name = 'Bodega', dtype=str) #Calling criterios to perform combine action
#df = df.parse() #Selecting sheet
#############################################################################################
df = df.drop(['User_Apr', 'D_Item_Apr', 'Factor_Um_Item_Apr', 'Peso_Um_Item_Apr', 'Usr_LE', 'Fecha_VbOk', 'If_ok', 'TB_LE', 'Orig_Bodega_Apr', 'Dest_Bodega_Apr'], axis=1) #Droping columns
df = df.drop(df.columns[df.columns.str.contains('Unnamed',case = False)],axis = 1) #drop columns with no title
df['Date_Apr'] = df['Date_Apr'].astype('datetime64[ns]') 

df = df.loc[df['Date_Apr'] == now] #replace with now #selecting last register if applies
df = df.loc[df['Estado'] == 'OK'] 
df = df.drop(['Estado'], axis=1)
df = df.set_index('Id_Apr') #Asigning new column to index
df = df.reset_index() #self explanatory
round(pd.to_numeric(df['Qty_kg_Apr']),0)
df['Qty_Um_Apr'] = round(pd.to_numeric(df['Qty_Um_Apr']),0)
df = df.loc[df['Qty_Um_Apr'] > 0]
df = df.set_index('Id_Apr') #Asigning new column to inex
df = df.reset_index() #self explanatory
#df = df.loc[:0,:]
df = df.loc[df['Cod_Bodega_O_Apr'] != df['Cod_Bodega_D_Apr']]
df = df.set_index('Id_Apr') #Asigning new column to inex
df = df.reset_index() #self explanatory
df.rename(columns={'Cod_Bodega_O_Apr': 'id_Bodega', 'Cod_Bodega_D_Apr': 'id_Bodega2'}, inplace=True) #Renaming columns
df3.rename(columns={'id_Bodega':'id_Bodega2'}, inplace=True) #preparation to combine id_Bodega
df2 = df2.drop(['descBodega', 'estado'], axis=1) #Vlookup(combination) for id_COpera
df3 = df3.drop(['descBodega', 'estado'], axis=1)
df = df.merge(df2, on='id_Bodega')
df = df.merge(df3, on='id_Bodega2')
df = df.loc[df['id_Bodega2'] != '00092']
df = df.loc[df['id_Bodega2'] != '00095']
df = df.loc[df['id_Bodega2'] != '00-91']
df = df.loc[df['id_Bodega'] != '00-91'] #sometimes
df['ctrl1'] = df['id_Bodega'] + df['id_Bodega2'] #Control column
df['Date_ent'] = df['Date_Apr'] + timedelta(days=6)

df = df.sort_values('ctrl1')

l = list(df['ctrl1']) # extracting a list to pass to find_unique function


def find_unique(lists): # function to find unique values 
	unique = []
	i=0
	for regis  in lists:
		if regis != lists[abs(i-1)]:
			unique.append('1')
		else:
			unique.append('0')
		i=i+1
	unique[0] = '1'
	return unique

ctrl2 = find_unique(l) #creating list and 
df.insert(loc=11, column='ctrl2', value=ctrl2)

l2 = list(df['ctrl2']) # extracting a list to pass to find_unique function


def find_consec(lists2): #function to find consecutive values
	consec = []
	#consec[0] = '1'
	i=0
	for regis in lists2:
		if regis == '0':
			consec.append(str(i))
		else:
			consec.append('1')
			i=1
		i=i+1
	return consec

ctrl3 = find_consec(l2)

df.insert(loc=12, column='ctrl3', value=ctrl3)

l3 = list(df['ctrl2']) # extracting a list to pass to find_unique function


def find_doc(lists3): #function to find doc values
	doc = []
	#consec[0] = '1'
	i=0
	for regis in lists3:
		if regis == '1':
			doc.append(str(i+1))
			i=i+1
		else:
			doc.append(str(i))
	return doc

ctrl4 = find_doc(l3)

df.insert(loc=11, column='ctrl4', value=ctrl4)
df['Date_Apr'] = df['Date_Apr'].astype(str)
df['Date_ent'] = df['Date_ent'].astype(str)

l4 = list(df['Date_Apr']) #Extraction a list to remove the hyphens
l5 = list(df['Date_ent'])


def kill_hyphens(lists4):
	no_hyphens = []
	for d in lists4:
		no_hyphens.append(d.replace('-',''))
	return no_hyphens

df['Date_ent'] = kill_hyphens(l5)
df['Date_Apr'] = kill_hyphens(l4)

docs = df.loc[df['ctrl2'] == '1']
docs = docs.drop(['ctrl2', 'id_COpera_x', 'Um_Item_Apr', 'Qty_kg_Apr', 'Item_Apr', 'ctrl3'], axis=1)
docs= docs[['id_COpera_y', 'ctrl4', 'Date_Apr', 'Date_ent', 'id_Bodega', 'id_Bodega2']]
docs = docs.set_index('id_COpera_y')
docs = docs.reset_index()

df = df.drop(['Id_Apr', 'ctrl1', 'Date_Apr', 'ctrl2', 'id_Bodega2', 'Qty_kg_Apr'], axis=1)
df = df[['id_COpera_y', 'ctrl4', 'ctrl3', 'Item_Apr', 'id_Bodega', 'Um_Item_Apr', 'Qty_Um_Apr', 'Date_ent', 'id_COpera_x']]
df = df.set_index('id_COpera_y')
df = df.reset_index()

nee = {"id_COpera_y": "copera", "ctrl4":"numDoc", "Date_Apr":"doc_FechaDocumento", "Date_ent":"doc_FechaEntrega", "id_Bodega":"doc_BodegaSalida", "id_Bodega2":"doc_BodegaEntrada"}
nee2 = {"id_COpera_y": "copera", "ctrl4":"numDoc", "ctrl3":"mov_Registro", "Item_Apr":"mov_Item", "id_Bodega":"mov_Bodega", "Um_Item_Apr":"mov_UMedida", "Qty_Um_Apr":"mov_Cantidad", "Date_ent":"mov_FechaEntrega", "id_COpera_x":"mov_Copera"}

docs.rename(columns=nee, inplace=True) #Renaming columns
df.rename(columns=nee2, inplace=True) #Renaming columns

df = df.set_index('copera')
df = df.reset_index()

docs = docs.set_index('copera')
docs = docs.reset_index()

'''
date_to_name = str(now)
date_to_name = date_to_name[0:10]
file_name = 'REQUISICIONES_PLANEACION_' + date_to_name + '.xlsx'

write_dfs = pd.ExcelWriter(file_name, engine = 'xlsxwriter')
docs.to_excel(write_dfs, sheet_name = 'Documentos')
df.to_excel(write_dfs, sheet_name = 'Movimientos')
write_dfs.save()
write_dfs.close() #This is in case i need it as emergency solution!'''

emptyDIC1 = {}
emptyDIC2 = {}
fullDIC = {}
listOUTTER =[]
listINNER = []

listOUTTER.append(emptyDIC1)
listOUTTER.append(fullDIC)
listOUTTER.append(emptyDIC2)	

data = ''
s=0
j=0
rows_docs = len(docs.index)
rows_df = len(df.index)

for k in range(rows_docs):
	id_doc = int(docs.iat[s,1])
	fullDIC.update( {"copera": docs.iat[s,0], "numDoc": docs.iat[s,1], "doc_FechaDocumento": docs.iat[s,2], "doc_FechaEntrega": docs.iat[s,3], "doc_BodegaSalida": docs.iat[s,4], "doc_BodegaEntrada": docs.iat[s,5]} )
	j=0
	for i in range(rows_df): #built an identifier that corelates rows_df with rows_docs wich is ctrl 2 or numdoc that i always do.
		id_mov = int(df.iat[j,1])
		if id_doc == id_mov:
			listINNER.append({"copera": (df.iat[j,0]), "numDoc": df.iat[j,1], "mov_Registro": (df.iat[j,2]), "mov_Item": (df.iat[j,3]), "mov_Bodega": (df.iat[j,4]), "mov_UMedida": (df.iat[j,5]), "mov_Cantidad": (df.iat[j,6]), "mov_FechaEntrega": (df.iat[j,7])})
		j=j+1
	# Adding a new key value pair n values
	fullDIC.update( {'mov_' : listINNER} )
	data = json.dumps(str(listOUTTER))
	data = data.replace('\'', '\"')
	data = data[1:-1]
	#print(data)
	listINNER = []
	s=s+1
	
	My_url = 'http://olympus.web.lan/Olympus/api/siesa/importar/rstOly_sob_aba_sie_pedInterno'
	My_auth = HTTPBasicAuth('dpto_user', 'pass2019')

	r = requests.post( url = My_url, data = data, auth = My_auth)
	#print(r.text)

#print('Success!')
popupmsg("Hurry up and beat me dowm, i don't like what im supposed to handle.")
