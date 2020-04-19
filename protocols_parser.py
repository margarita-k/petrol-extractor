import docx
import csv
from docx import Document
import pandas as pd
import myfunction
import cProfile
import pstats

p = cProfile.Profile()
p.enable()

df = pd.read_csv("properties.csv")
files_list = myfunction.get_names("D:\Margo\python\protocols")
prop_dict = {
	"Густина кг/м3": "густина",
	"ОЧД":"дослідним",
	"ОЧМ":"моторним",
	"Фракція_70":"70\s?оС",
	"Фракція_100":"100\s?оС",
	"Фракція_150":"150\s?оС",
	"10%":"10\s?%",
	"50%":"50\s?%",
	"90%":"90\s?%",
	"Кінець кипіння":"кіне?ц",
	"Залишок після випаровування":"залиш",
	"ДНП":"насиченої",
	"Індукційний період":"ст(абільн|ійк)",
	"Фактичні смоли":"смол",
	"Сірка":"сірки",
	"Олефінові вуглеводні":"олефін",
	"Ароматичні вуглеводні":"аромат",
	"Бензол":"бензол",
	"Кисневмісні: метанол":"метанол",
	"Кисневмісні: етанол":"біо\)?етанол",
	"Кисневмісні: ізопропіловий спирт":"ізопропіл",
	"Кисневмісні: ізобутиловий спирт":"ізобутил",
	"Кисневмісні: третбутиловий":"третбутил",
	"Кисневмісні: етери":"е(те|фі)р"
}

i = 0
#k = 0

for item in files_list:	
	try:
		doc = Document(item)
	except docx.opc.exceptions.PackageNotFoundError:
		print("PackageNotFoundError - ", item)
		continue

	mydoc = myfunction.get_format(doc)

	if (mydoc.id == -1):
		print("ID not found -", item)
	if (mydoc.results_col == -1):
		print("Результати not found -", item)
		continue

	df.at[i,"ID"] = mydoc.id
	df.at[i,"doc"] = item
	df.at[i,"Тип документу"] = mydoc.type
	df.at[i,"Зразок"] = mydoc.sample
	df.at[i,"Виробник"] = mydoc.producer
	df.at[i,"Замовник"] = mydoc.client

	df = myfunction.find_properties(df,prop_dict,mydoc,i)                                                                    
	'''for column in df.columns[6:]:
		df.at[i,column] = myfunction.find_property(mylist[k], mydoc)
		k = k + 1
	'''
	i = i + 1
	# k = 0
	print(item, " COMPLETED")	
p.disable()
p_stats = pstats.Stats(p)
p_stats.sort_stats('cumtime')
p_stats.print_stats()


df.to_csv("protocol_data.csv")



