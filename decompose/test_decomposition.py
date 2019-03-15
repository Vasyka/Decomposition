from decompose import *
'''
# Создаем новый экземпляр класса, загружаем данные и готовим их к декомпозиции
dec = Decomposition()
path_and_sheetnames = {"./data/задание 1 для студентов.xlsx": [0, 1]}
dec.load_WIOD2013_merged_data(**path_and_sheetnames)
dec.prepare_data(column_order='WIOD')

# Запускаем методы декомпозиции
dec.decomposition_Baranov_2016()
dec.decomposition_Baranov_2018()
dec.decomposition_Magacho_2018()
'''

# World WIOD2013
dec = Decomposition()
dec.load_WIOD2013_world()

'''
#Для WIOD16 года
# Создаем новый экземпляр класса, загружаем данные и готовим их к декомпозиции
dec = Decomposition()
path_and_sheetnames = {"./data/RUS_niot_nov16 (done) 27 ДЕКАБРЯ .xlsx": ['2003 в ценах 2011', '2010 в ценах 2011']}
dec.load_WIOD2016_merged_data(**path_and_sheetnames)
dec.prepare_data(column_order='WIOD')

# Запускаем методы декомпозиции
dec.decomposition_Baranov_2016()
dec.decomposition_Baranov_2018()
dec.decomposition_Magacho_2018()
dec.decomposition_Chenery_extended()



#Для WIOD16 года
# Создаем новый экземпляр класса, загружаем данные и готовим их к декомпозиции
dec = Decomposition()
path_and_sheetnames = {"./data/RUS_niot_nov16 (done) 27 ДЕКАБРЯ .xlsx": ['2011 в рублях', '2014 в ценах 2011']}
dec.load_WIOD2016_merged_data(**path_and_sheetnames)
dec.prepare_data(column_order='WIOD')

# Запускаем методы декомпозиции
dec.decomposition_Baranov_2016()
dec.decomposition_Baranov_2018()
dec.decomposition_Magacho_2018()
dec.decomposition_Chenery_extended()

# На новых данных:
dec1 = Decomposition()
path_and_sheetnames1 = {"./data/симметричные таблицы WIOD 2003 и 2010 в ценах 2008 25 декабря.xlsx": [0, 1]}
dec1.load_WIOD2013_merged_data(**path_and_sheetnames1)
dec1.prepare_data(column_order='WIOD')

# Запускаем методы декомпозиции
dec1.decomposition_Baranov_2016()
dec1.decomposition_Baranov_2018()
dec1.decomposition_Magacho_2018()
dec1.decomposition_Chenery_extended()



# Создаем новый экземпляр класса, загружаем данные и готовим их к декомпозиции
dec = Decomposition()
path_and_sheetnames = {"./data/all2011.xlsx": ['SD calculated def', 'SM calculated def'],
                       "./data/all2014 (проверочная).xlsx": ['SD calculated def', 'SM calculated def']}
dec.load_Rosstat_separated_data(**path_and_sheetnames)
dec.prepare_data(column_order='Rosstat')
# Запускаем методы декомпозиции
dec.decomposition_Baranov_2016()
dec.decomposition_Baranov_2018()
dec.decomposition_Magacho_2018()
dec.decomposition_Chenery_extended()



# Создаем новый экземпляр класса, загружаем данные и готовим их к декомпозиции
dec = Decomposition()
path_and_sheetnames = {"./data/all2014 (проверочная).xlsx": ['SD calculated def', 'SM calculated def'],
                       "./data/all2015 (проверочная).xlsx": ['SD calculated def', 'SM calculated def']}
dec.load_Rosstat_separated_data(**path_and_sheetnames)
dec.prepare_data(column_order='Rosstat')

# Запускаем методы декомпозиции
dec.decomposition_Baranov_2016()
dec.decomposition_Baranov_2018()
dec.decomposition_Magacho_2018()'''
#dec.decomposition_Chenery_extended()