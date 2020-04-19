import docx
import csv
from docx import Document
import pandas as pd
import myfunction
import cProfile
import pstats

p = cProfile.Profile()
p.enable()

df = pd.read_csv("dict_properties.csv")
files_list = myfunction.get_names("D:\Margo\python\protocols")
mylist = ["густина","дослідним","моторним", "70\s?оС", "100\s?оС", "150\s?оС", "10\s?%", "50\s?%", "90\s?%", "кіне?ц", "залиш", "насиченої", "ст(абільн|ійк)", "смол", "сірки", "олефін", "аромат", "бензол", "метанол", "біоетанол", "ізопропіл", "ізобутил", "третбутил","е(те|фі)р"]

i = 0
k = 0

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
                                                                          
	for column in df.columns[6:]:
		df.at[i,column] = myfunction.find_property(mylist[k], mydoc)
		k = k + 1
	
	i = i + 1
	k = 0
	print(item, " COMPLETED")	
p.disable()
p_stats = pstats.Stats(p)
p_stats.sort_stats('cumtime')
p_stats.print_stats()


df.to_csv("protocol_data.csv")



