from decompose import *

# Создаем новый экземпляр класса, загружаем данные и готовим их к декомпозиции
dec = Decomposition()
path_and_sheetnames = {"./data/задание 1 для студентов.xlsx": [0, 1]}
dec.load_WIOD2013_merged_data(**path_and_sheetnames)
dec.prepare_data(column_order='eng')

# Запускаем методы декомпозиции
dec.decomposition_Baranov_2016()
dec.decomposition_Baranov_2018()
dec.decomposition_Magacho_2018()





# Создаем новый экземпляр класса, загружаем данные и готовим их к декомпозиции
dec = Decomposition()
path_and_sheetnames = {"./data/all2011.xlsx": ['SD calculated def', 'SM calculated def'],
                       "./data/all2014 (проверочная).xlsx": ['SD calculated def', 'SM calculated def']}
dec.load_Rosstat_separated_data(**path_and_sheetnames)
dec.prepare_data(column_order='rus')
# Запускаем методы декомпозиции
dec.decomposition_Baranov_2016()
dec.decomposition_Baranov_2018()
dec.decomposition_Magacho_2018()




# Создаем новый экземпляр класса, загружаем данные и готовим их к декомпозиции
dec = Decomposition()
path_and_sheetnames = {"./data/all2014 (проверочная).xlsx": ['SD calculated def', 'SM calculated def'],
                       "./data/all2015 (проверочная).xlsx": ['SD calculated def', 'SM calculated def']}
dec.load_Rosstat_separated_data(**path_and_sheetnames)
dec.prepare_data(column_order='rus')

# Запускаем методы декомпозиции
dec.decomposition_Baranov_2016()
dec.decomposition_Baranov_2018()
dec.decomposition_Magacho_2018()
