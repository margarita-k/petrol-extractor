import os
	
path = "D:\Margo\python\protocols\\"
print(path)
i = 0
for root, dirs, files in os.walk(os.path.abspath(path)):
	for file in files:
		os.replace(path + file, path + "Protocol_" + str(i) + ".doc")
		i = i+1