import os
import time
#import openpyxl
import numpy as np
import pandas as pd
from collections import Counter
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from matplotlib.colors import Normalize
from matplotlib.cm import ScalarMappable
from scipy.stats import pearsonr, ttest_1samp
<<<<<<< HEAD

INPUT_FILE = 'Data_Excel.xlsx'
PREPARED_DATA_FILE = 'Data_Excel_Bins.xlsx'


# Функція підготовки excel-таблиці до роботи
def data_preparation():
    values_to_replace = {
        "5 = Total agreement": 5,
        "DezTotal agreement=1": 1,
        "5=Excelente": 5,
        "F. Weak =1": 1,
        "F. Weak =1, 2": 1,
        "Totally disagree=1": 1,
        "5 = Total Agreement": 5,
        "5 = Totally agree": 5,
        "3, 4": 3,
        "Good": 4,
        "Neutre": 3,
        "Very Good": 5,
        "Weak": 2,
        "Very Weak": 1,
        "I can't answer": 3,
        "Agree": 4,
        "Neutru": 3,
        "Total agreement": 5,
        "DezAgree": 2,
        "Total DezAgree": 1
        # "I can't answer":3
    }

    workbook = openpyxl.load_workbook(INPUT_FILE)

    for worksheet in workbook:
        for column in worksheet.iter_cols():
            for cell in column:
                # Перевірка, чи містить клітинка текст, який потрібно замінити
                if cell.value in values_to_replace:
                    cell.value = values_to_replace[cell.value]
                # Якщо в непорожньому стовпці є порожня клітинка, вписати в неї "I can't answer"
                elif cell.value is None and any(cell1.value is not None for cell1 in column):
                    # cell.value = "I can't answer"
                    cell.value = 3
    workbook.save(PREPARED_DATA_FILE)
    workbook.close()


# Функція створення гістограм по всіх стовпцях кожного аркуша таблиці
# Не запускати без попереднього запуску функції "data_preparation()"
def histogram_creation():
    workbook = pd.ExcelFile(PREPARED_DATA_FILE)

    output_directory = 'histograms'
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    for sheet_name in workbook.sheet_names:
        df = workbook.parse(sheet_name)
        for column_name in df.columns:
            column_name = str(column_name)
            if not df[column_name].dropna().empty:
                plt.figure(figsize=(8, 6))
                try:
                    plt.hist(df[column_name], bins=10, color='skyblue', edgecolor='black')
                except:
                    print(column_name, df[column_name])
                    continue
                plt.title(column_name)
                plt.xlabel('Value')
                plt.ylabel('Frequency')
                plt.grid(True)

                # Формування шляху для збереження гістограми із заміною "проблемних символів"
                clean_sheet_name = str(sheet_name).replace(' ', '_').replace('.', '').replace('/', '_').replace('"', '')
                clean_column_name = str(column_name).replace(' ', '_').replace('.', '').replace('/', '_').replace('"', '')
                histogram_filename = f'{output_directory}/{clean_sheet_name}_{clean_column_name}_histogram.png'

                plt.savefig(histogram_filename)
                plt.show()
                # Затримка між запитами, щоб уникнути блокування
                time.sleep(0.2)

# Функція отримання масивів даних із стовпців excel-таблиці
# Не запускати без попереднього запуску функції "data_preparation()"
def get_arrays_of_data():
    workbook = load_workbook(PREPARED_DATA_FILE)

    # Створення словника для збереження масивів з кожного аркуша
    all_arrays = {}

    # Створення словника для збереження кількості стовпців до того, як зустрінеться перший порожній, з кожного аркуша (це використовуватиметься у побудові графіків)
    all_num_of_cols_before_empty = {}

    # Створення масиву для збереження назв стопців, щоб уникнути повторів
    array_of_column_names = []

    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        arrays = {}
        num_of_cols_before_empty = 0

        # Знаходження к-ті стовпців до першого порожнього
        for column in sheet.iter_cols(values_only=True):
            if any(cell is not None for cell in column):
                num_of_cols_before_empty += 1
            else:
                break

        all_num_of_cols_before_empty[sheet_name] = num_of_cols_before_empty

        for column in sheet.iter_cols(values_only=True):
            # Перевірка, чи стовпець не є порожнім
            if any(cell is not None for cell in column):

                column_name = column[0]
                """
                # Перевірка, чи не опрацьований вже поточний стовпець
                if column_name not in array_of_column_names:
                    array_of_column_names.append(column_name)
                    arrays[column_name] = [cell for cell in column if cell is not None and cell != column_name]
                """
                arrays[column_name] = [cell for cell in column if cell is not None and cell != column_name]

        all_arrays[sheet_name] = arrays
    """
    for sheet_name, arrays in all_arrays.items():
        print(f"Аркуш: {sheet_name}")
        for column_name, array in arrays.items():
            print(f"\tНазва стовпця: {column_name}")
            print("\tМасив:", array)

    #print(all_num_of_cols_before_empty)
    """
    result = {
        "all_arrays": all_arrays,
        "all_num_of_cols_before_empty": all_num_of_cols_before_empty
    }

    return result


# Функція побудови графіків, використовуючи масиви даних, здобуті за допомогою функції "get_arrays_of_data()"
# Не запускати без попереднього запуску функції "data_preparation()"
def construction_of_graphs():

    graphs_directory = 'graphs'
    if not os.path.exists(graphs_directory):
        os.makedirs(graphs_directory)

    letters_of_columns = {
        "PROACTIVE_ORIENTATION_[Technologies_used_are_the_latest]": "A",
        "PROACTIVE_ORIENTATION_[We_anticipate_the_potential_of_new_technologies_practices]": "B",
        "PROACTIVE_ORIENTATION_[We_systematically_try_to_acquire_and_implement_new_technologies]": "C",
        "PROACTIVE_ORIENTATION_[The_research_and_development_department_is_a_leader_in_the_field]": "D",
        "PERFORMANȚELE_FINANCIARE_[Profitul_brut]": "A",
        "PERFORMANȚELE_FINANCIARE(1)_[Profitul_brut]": "B",
        "FINANCIAL_PERFORMANCE_[Gross_Profit]": "C",
        "FINANCIAL_PERFORMANCE_[Return_on_assets]": "D",
        "FINANCIAL_PERFORMANCE_[Sales]": "E",
        "FINANCIAL_PERFORMANCE_[Earnings_per_share]": "F",
        "FINANCIAL_PERFORMANCE(1)_[Earnings_per_share]": "G",
        "FINANCIAL_PERFORMANCE_[Rate_of_Profit]": "H",
        "INNOVATION_[Research_activity]": "A",
        "INNOVATION_[The_degree_of_product_novelty]": "B",
        "INNOVATION_[Using_the_latest_technologies]": "C",
        "INNOVATION_[Speed_of_development_of_new_products]": "D",
        "INNOVATION_[Share_of_new_products_within_the_range]": "E",
        "AgeGroup": 'A',
        "EcoGroup": "A",
        "Transport": 'B'
    }

    for sheet_name, arrays in get_arrays_of_data()["all_arrays"].items():

        # Побудова графіків для кожної пари стовпців, де х-координати - це значення зі стовпців справа, а у-координати - значення зі стовпців зліва кожного аркуша.
        counter1 = 0
        for left_column_name, left_array in arrays.items():
            # left_column_name.sorted()
            counter1 += 1
            counter2 = 0
            if counter1 <= get_arrays_of_data()["all_num_of_cols_before_empty"][sheet_name]:
                for right_column_name, right_array in arrays.items():
                    counter2 += 1
                    if counter2 > get_arrays_of_data()["all_num_of_cols_before_empty"][sheet_name]:

                        plt.figure(figsize=(15, 10))
                        point_counts = Counter(zip(right_array, left_array))

                        # Створення мапера кольорів на основі кількості повторень для кожної точки
                        max_count = max(point_counts.values())
                        norm = Normalize(vmin=0, vmax=max_count)
                        cmap = plt.get_cmap('viridis')
                        sm = ScalarMappable(norm=norm, cmap=cmap)

                        for (x, y), count in point_counts.items():
                            color = sm.to_rgba(count)
                            size = count * 65
                            plt.scatter(x, y, color=color, s=size, alpha=1)
                            plt.annotate(f'{count}', (x, y), fontsize=10, textcoords="offset points", xytext=(15, 15), ha='left')

                        #plt.scatter(right_array, left_array, color='red', alpha=1)
                        plt.title('Plot of dependancy')
                        plt.xlabel(right_column_name)
                        plt.ylabel(left_column_name)
                        plt.grid(True)

                        # Формування шляху для збереження графіка із заміною "проблемних символів"
                        clean_sheet_name = str(sheet_name).replace(' ', '_').replace('.', '').replace('/', '_').replace(
                            '"', '')
                        clean_left_column_name = str(left_column_name).replace(' ', '_').replace('.', '').replace('/',
                                                                                                        '_').replace(
                            '"', '')
                        clean_right_column_name = str(right_column_name).replace(' ', '_').replace('.', '').replace('/',
                                                                                                                  '_').replace(
                            '"', '')
                        print(f'{clean_sheet_name}_{letters_of_columns}{clean_left_column_name}_{clean_right_column_name}')
                        graph_filename = f'{graphs_directory}/{clean_sheet_name[0:2]}_{letters_of_columns[clean_left_column_name]}_{clean_right_column_name}_graph.png'

                        plt.savefig(graph_filename)
                        plt.show()
                        # Затримка між запитами, щоб уникнути блокування
                        time.sleep(0.2)

# Функція обчислення статистичної значущості по впливу параметрів
# Не запускати без попереднього запуску функції "data_preparation()"
def correlation_calculation():
    # Словник для заміни текстових значень з таблиці на відповідні числові
    numerical_equivalent = {
        "Total DezAgree": 1,
        "Very Weak": 1,
        "DezAgree": 2,
        "Weak": 2,
        "Neutru": 3,
        "Neutre": 3,
        "Agree": 4,
        "Good": 4,
        "Total agreement": 5,
        "Very Good": 5,
        "I can't answer": 3
    }

    letters_of_columns = {
        "PROACTIVE_ORIENTATION_[Technologies_used_are_the_latest]": "A",
        "PROACTIVE_ORIENTATION_[We_anticipate_the_potential_of_new_technologies_practices]": "B",
        "PROACTIVE_ORIENTATION_[We_systematically_try_to_acquire_and_implement_new_technologies]": "C",
        "PROACTIVE_ORIENTATION_[The_research_and_development_department_is_a_leader_in_the_field]": "D",
        "PERFORMANȚELE_FINANCIARE_[Profitul_brut]": "A",
        "PERFORMANȚELE_FINANCIARE(1)_[Profitul_brut]": "B",
        "FINANCIAL_PERFORMANCE_[Gross_Profit]": "C",
        "FINANCIAL_PERFORMANCE_[Return_on_assets]": "D",
        "FINANCIAL_PERFORMANCE_[Sales]": "E",
        "FINANCIAL_PERFORMANCE_[Earnings_per_share]": "F",
        "FINANCIAL_PERFORMANCE(1)_[Earnings_per_share]": "G",
        "FINANCIAL_PERFORMANCE_[Rate_of_Profit]": "H",
        "INNOVATION_[Research_activity]": "A",
        "INNOVATION_[The_degree_of_product_novelty]": "B",
        "INNOVATION_[Using_the_latest_technologies]": "C",
        "INNOVATION_[Speed_of_development_of_new_products]": "D",
        "INNOVATION_[Share_of_new_products_within_the_range]": "E",
        "AgeGroup": 'A',
        "EcoGroup": "A",
        "Transport": 'B'
    }

    with open('correlation_calculation_results.txt', 'w+', encoding='utf-8') as file:

        for sheet_name, arrays in get_arrays_of_data()["all_arrays"].items():
            counter1 = 0
            for left_column_name, left_array in arrays.items():
                counter1 += 1
                counter2 = 0
                if counter1 <= get_arrays_of_data()["all_num_of_cols_before_empty"][sheet_name]:
                    for right_column_name, right_array in arrays.items():
                        counter2 += 1
                        if counter2 > get_arrays_of_data()["all_num_of_cols_before_empty"][sheet_name]:
                            # Заміна текстових параметрів з лівих стовпців кожного аркуша на відповідні числові
                            numerical_left_array = []
                            for cell in left_array:
                                if cell in numerical_equivalent.keys():
                                    numerical_cell = numerical_equivalent[cell]
                                    numerical_left_array.append(numerical_cell)
                                else:
                                    numerical_left_array.append(cell)

                            # Формування заголовку для результатів із заміною "проблемних символів"
                            clean_sheet_name = str(sheet_name).replace(' ', '_').replace('.', '').replace('/',
                                                                                                          '_').replace(
                                '"', '')
                            clean_left_column_name = str(left_column_name).replace(' ', '_').replace('.',
                                                                                                     '').replace(
                                '/', '_').replace(
                                '"', '')
                            clean_right_column_name = str(right_column_name).replace(' ', '_').replace('.',
                                                                                                       '').replace(
                                '/',
                                '_').replace(
                                '"', '')
                            results_header = f'{clean_sheet_name[0:2]}_{letters_of_columns[clean_left_column_name]}_{clean_right_column_name}'

                            # Обчислення кореляції Пірсона для кожної пари стовпців зліва та справа кожного аркуша.
                            correlation_coefficient, p_value = pearsonr(right_array, numerical_left_array)
                            print(results_header, '\n')
                            print("Pearson corellation:", correlation_coefficient)
                            print("p-value:", p_value)
                            # Обчислення різниці в значенні
                            difference = np.array(right_array) - np.array(numerical_left_array)
                            # Проведення t-тесту
                            t_statistic, p_value_ttest = ttest_1samp(difference, 0)
                            print("t-statistics:", t_statistic)
                            print("p-value for t-test:", p_value_ttest)
                            print('\n')

                            file.write(f"{results_header}\n")
                            file.write(f"Pearson c.c: {correlation_coefficient}\n")
                            file.write(f"p-value: {p_value}\n")
                            file.write(f"t-stat for differences: {t_statistic}\n")
                            file.write(f"p-value for t-test: {p_value_ttest}\n\n")

# data_preparation()
histogram_creation()
construction_of_graphs()
correlation_calculation()

#
# FIRMS_DATA = "Chestionar_Viktor.xlsx"
# FIRMS_DATA2 = "Firms"
#
# firm_data ={
#     "name":"",
#     "field":"",
#     "employers":0,
#     "progress":True,
#     "perform":0
# }
#
# SHEET1 = 'Sheet1'
#
# firms = pd.read_excel(FIRMS_DATA, sheet_name=SHEET1, usecols="A", dtype=str).dropna(how='all')
# print(firms.values[1:],firms.index)
#
#
# avg_stuff = pd.read_excel(FIRMS_DATA, sheet_name=SHEET1, usecols="K", dtype=str)
# print(avg_stuff.values[:10])
#
# stuff = []
# stuffs = []
# for i,peoples in enumerate(avg_stuff.values):
=======
#
# INPUT_FILE = 'Data_Excel.xlsx'
# PREPARED_DATA_FILE = 'Data_Excel_Bins.xlsx'
#
#
# # Функція підготовки excel-таблиці до роботи
# def data_preparation():
#     values_to_replace = {
#         "5 = Total agreement": 5,
#         "DezTotal agreement=1": 1,
#         "5=Excelente": 5,
#         "F. Weak =1": 1,
#         "F. Weak =1, 2": 1,
#         "Totally disagree=1": 1,
#         "5 = Total Agreement": 5,
#         "5 = Totally agree": 5,
#         "3, 4": 3,
#         "Good": 4,
#         "Neutre": 3,
#         "Very Good": 5,
#         "Weak": 2,
#         "Very Weak": 1,
#         "I can't answer": 3,
#         "Agree": 4,
#         "Neutru": 3,
#         "Total agreement": 5,
#         "DezAgree": 2,
#         "Total DezAgree": 1
#         # "I can't answer":3
#     }
#
#     workbook = openpyxl.load_workbook(INPUT_FILE)
#
#     for worksheet in workbook:
#         for column in worksheet.iter_cols():
#             for cell in column:
#                 # Перевірка, чи містить клітинка текст, який потрібно замінити
#                 if cell.value in values_to_replace:
#                     cell.value = values_to_replace[cell.value]
#                 # Якщо в непорожньому стовпці є порожня клітинка, вписати в неї "I can't answer"
#                 elif cell.value is None and any(cell1.value is not None for cell1 in column):
#                     # cell.value = "I can't answer"
#                     cell.value = 3
#     workbook.save(PREPARED_DATA_FILE)
#     workbook.close()
#
#
# # Функція створення гістограм по всіх стовпцях кожного аркуша таблиці
# # Не запускати без попереднього запуску функції "data_preparation()"
# def histogram_creation():
#     workbook = pd.ExcelFile(PREPARED_DATA_FILE)
#
#     output_directory = 'histograms'
#     if not os.path.exists(output_directory):
#         os.makedirs(output_directory)
#
#     for sheet_name in workbook.sheet_names:
#         df = workbook.parse(sheet_name)
#         for column_name in df.columns:
#             column_name = str(column_name)
#             if not df[column_name].dropna().empty:
#                 plt.figure(figsize=(8, 6))
#                 try:
#                     plt.hist(df[column_name], bins=10, color='skyblue', edgecolor='black')
#                 except:
#                     print(column_name, df[column_name])
#                     continue
#                 plt.title(column_name)
#                 plt.xlabel('Value')
#                 plt.ylabel('Frequency')
#                 plt.grid(True)
#
#                 # Формування шляху для збереження гістограми із заміною "проблемних символів"
#                 clean_sheet_name = str(sheet_name).replace(' ', '_').replace('.', '').replace('/', '_').replace('"', '')
#                 clean_column_name = str(column_name).replace(' ', '_').replace('.', '').replace('/', '_').replace('"',                                                                                             '')
#                 histogram_filename = f'{output_directory}/{clean_sheet_name}_{clean_column_name}_histogram.png'
#
#                 plt.savefig(histogram_filename)
#                 plt.show()
#                 # Затримка між запитами, щоб уникнути блокування
#                 time.sleep(1)
#
# # Функція отримання масивів даних із стовпців excel-таблиці
# # Не запускати без попереднього запуску функції "data_preparation()"
# def get_arrays_of_data():
#     workbook = load_workbook(PREPARED_DATA_FILE)
#
#     # Створення словника для збереження масивів з кожного аркуша
#     all_arrays = {}
#
#     # Створення словника для збереження кількості стовпців до того, як зустрінеться перший порожній, з кожного аркуша (це використовуватиметься у побудові графіків)
#     all_num_of_cols_before_empty = {}
#
#     # Створення масиву для збереження назв стопців, щоб уникнути повторів
#     array_of_column_names = []
#
#     for sheet_name in workbook.sheetnames:
#         sheet = workbook[sheet_name]
#         arrays = {}
#         num_of_cols_before_empty = 0
#
#         # Знаходження к-ті стовпців до першого порожнього
#         for column in sheet.iter_cols(values_only=True):
#             if any(cell is not None for cell in column):
#                 num_of_cols_before_empty += 1
#             else:
#                 break
#
#         all_num_of_cols_before_empty[sheet_name] = num_of_cols_before_empty
#
#         for column in sheet.iter_cols(values_only=True):
#             # Перевірка, чи стовпець не є порожнім
#             if any(cell is not None for cell in column):
#
#                 column_name = column[0]
#                 """
#                 # Перевірка, чи не опрацьований вже поточний стовпець
#                 if column_name not in array_of_column_names:
#                     array_of_column_names.append(column_name)
#                     arrays[column_name] = [cell for cell in column if cell is not None and cell != column_name]
#                 """
#                 arrays[column_name] = [cell for cell in column if cell is not None and cell != column_name]
#
#         all_arrays[sheet_name] = arrays
#     """
#     for sheet_name, arrays in all_arrays.items():
#         print(f"Аркуш: {sheet_name}")
#         for column_name, array in arrays.items():
#             print(f"\tНазва стовпця: {column_name}")
#             print("\tМасив:", array)
#
#     #print(all_num_of_cols_before_empty)
#     """
#     result = {
#         "all_arrays": all_arrays,
#         "all_num_of_cols_before_empty": all_num_of_cols_before_empty
#     }
#
#     return result
#
#
# # Функція побудови графіків, використовуючи масиви даних, здобуті за допомогою функції "get_arrays_of_data()"
# # Не запускати без попереднього запуску функції "data_preparation()"
# def construction_of_graphs():
#
#     graphs_directory = 'graphs'
#     if not os.path.exists(graphs_directory):
#         os.makedirs(graphs_directory)
#
#     letters_of_columns = {
#         "PROACTIVE_ORIENTATION_[Technologies_used_are_the_latest]": "A",
#         "PROACTIVE_ORIENTATION_[We_anticipate_the_potential_of_new_technologies_practices]": "B",
#         "PROACTIVE_ORIENTATION_[We_systematically_try_to_acquire_and_implement_new_technologies]": "C",
#         "PROACTIVE_ORIENTATION_[The_research_and_development_department_is_a_leader_in_the_field]": "D",
#         "PERFORMANȚELE_FINANCIARE_[Profitul_brut]": "A",
#         "PERFORMANȚELE_FINANCIARE(1)_[Profitul_brut]": "B",
#         "FINANCIAL_PERFORMANCE_[Gross_Profit]": "C",
#         "FINANCIAL_PERFORMANCE_[Return_on_assets]": "D",
#         "FINANCIAL_PERFORMANCE_[Sales]": "E",
#         "FINANCIAL_PERFORMANCE_[Earnings_per_share]": "F",
#         "FINANCIAL_PERFORMANCE(1)_[Earnings_per_share]": "G",
#         "FINANCIAL_PERFORMANCE_[Rate_of_Profit]": "H",
#         "INNOVATION_[Research_activity]": "A",
#         "INNOVATION_[The_degree_of_product_novelty]": "B",
#         "INNOVATION_[Using_the_latest_technologies]": "C",
#         "INNOVATION_[Speed_of_development_of_new_products]": "D",
#         "INNOVATION_[Share_of_new_products_within_the_range]": "E",
#     }
#
#     for sheet_name, arrays in get_arrays_of_data()["all_arrays"].items():
#
#         # Побудова графіків для кожної пари стовпців, де х-координати - це значення зі стовпців справа, а у-координати - значення зі стовпців зліва кожного аркуша.
#         counter1 = 0
#         for left_column_name, left_array in arrays.items():
#             # left_column_name.sorted()
#             counter1 += 1
#             counter2 = 0
#             if counter1 <= get_arrays_of_data()["all_num_of_cols_before_empty"][sheet_name]:
#                 for right_column_name, right_array in arrays.items():
#                     counter2 += 1
#                     if counter2 > get_arrays_of_data()["all_num_of_cols_before_empty"][sheet_name]:
#
#                         plt.figure(figsize=(15, 10))
#                         point_counts = Counter(zip(right_array, left_array))
#
#                         # Створення мапера кольорів на основі кількості повторень для кожної точки
#                         max_count = max(point_counts.values())
#                         norm = Normalize(vmin=0, vmax=max_count)
#                         cmap = plt.get_cmap('viridis')
#                         sm = ScalarMappable(norm=norm, cmap=cmap)
#
#                         for (x, y), count in point_counts.items():
#                             color = sm.to_rgba(count)
#                             size = count * 65
#                             plt.scatter(x, y, color=color, s=size, alpha=1)
#                             plt.annotate(f'{count}', (x, y), fontsize=10, textcoords="offset points", xytext=(15, 15), ha='left')
#
#                         #plt.scatter(right_array, left_array, color='red', alpha=1)
#                         plt.title('Plot of dependancy')
#                         plt.xlabel(right_column_name)
#                         plt.ylabel(left_column_name)
#                         plt.grid(True)
#
#                         # Формування шляху для збереження графіка із заміною "проблемних символів"
#                         clean_sheet_name = str(sheet_name).replace(' ', '_').replace('.', '').replace('/', '_').replace(
#                             '"', '')
#                         clean_left_column_name = str(left_column_name).replace(' ', '_').replace('.', '').replace('/',
#                                                                                                         '_').replace(
#                             '"', '')
#                         clean_right_column_name = str(right_column_name).replace(' ', '_').replace('.', '').replace('/',
#                                                                                                                   '_').replace(
#                             '"', '')
#
#                         graph_filename = f'{graphs_directory}/{clean_sheet_name[0:2]}_{letters_of_columns[clean_left_column_name]}_{clean_right_column_name}_graph.png'
#
#                         plt.savefig(graph_filename)
#                         plt.show()
#                         # Затримка між запитами, щоб уникнути блокування
#                         time.sleep(1)
#
# # Функція обчислення статистичної значущості по впливу параметрів
# # Не запускати без попереднього запуску функції "data_preparation()"
# def correlation_calculation():
#     # Словник для заміни текстових значень з таблиці на відповідні числові
#     numerical_equivalent = {
#         "Total DezAgree": 1,
#         "Very Weak": 1,
#         "DezAgree": 2,
#         "Weak": 2,
#         "Neutru": 3,
#         "Neutre": 3,
#         "Agree": 4,
#         "Good": 4,
#         "Total agreement": 5,
#         "Very Good": 5,
#         "I can't answer": 3
#     }
#
#     letters_of_columns = {
#         "PROACTIVE_ORIENTATION_[Technologies_used_are_the_latest]": "A",
#         "PROACTIVE_ORIENTATION_[We_anticipate_the_potential_of_new_technologies_practices]": "B",
#         "PROACTIVE_ORIENTATION_[We_systematically_try_to_acquire_and_implement_new_technologies]": "C",
#         "PROACTIVE_ORIENTATION_[The_research_and_development_department_is_a_leader_in_the_field]": "D",
#         "PERFORMANȚELE_FINANCIARE_[Profitul_brut]": "A",
#         "PERFORMANȚELE_FINANCIARE(1)_[Profitul_brut]": "B",
#         "FINANCIAL_PERFORMANCE_[Gross_Profit]": "C",
#         "FINANCIAL_PERFORMANCE_[Return_on_assets]": "D",
#         "FINANCIAL_PERFORMANCE_[Sales]": "E",
#         "FINANCIAL_PERFORMANCE_[Earnings_per_share]": "F",
#         "FINANCIAL_PERFORMANCE(1)_[Earnings_per_share]": "G",
#         "FINANCIAL_PERFORMANCE_[Rate_of_Profit]": "H",
#         "INNOVATION_[Research_activity]": "A",
#         "INNOVATION_[The_degree_of_product_novelty]": "B",
#         "INNOVATION_[Using_the_latest_technologies]": "C",
#         "INNOVATION_[Speed_of_development_of_new_products]": "D",
#         "INNOVATION_[Share_of_new_products_within_the_range]": "E",
#     }
#
#     with open('correlation_calculation_results.txt', 'w+', encoding='utf-8') as file:
#
#         for sheet_name, arrays in get_arrays_of_data()["all_arrays"].items():
#             counter1 = 0
#             for left_column_name, left_array in arrays.items():
#                 counter1 += 1
#                 counter2 = 0
#                 if counter1 <= get_arrays_of_data()["all_num_of_cols_before_empty"][sheet_name]:
#                     for right_column_name, right_array in arrays.items():
#                         counter2 += 1
#                         if counter2 > get_arrays_of_data()["all_num_of_cols_before_empty"][sheet_name]:
#                             # Заміна текстових параметрів з лівих стовпців кожного аркуша на відповідні числові
#                             numerical_left_array = []
#                             for cell in left_array:
#                                 if cell in numerical_equivalent.keys():
#                                     numerical_cell = numerical_equivalent[cell]
#                                     numerical_left_array.append(numerical_cell)
#                                 else:
#                                     numerical_left_array.append(cell)
#
#                             # Формування заголовку для результатів із заміною "проблемних символів"
#                             clean_sheet_name = str(sheet_name).replace(' ', '_').replace('.', '').replace('/',
#                                                                                                           '_').replace(
#                                 '"', '')
#                             clean_left_column_name = str(left_column_name).replace(' ', '_').replace('.',
#                                                                                                      '').replace(
#                                 '/', '_').replace(
#                                 '"', '')
#                             clean_right_column_name = str(right_column_name).replace(' ', '_').replace('.',
#                                                                                                        '').replace(
#                                 '/',
#                                 '_').replace(
#                                 '"', '')
#                             results_header = f'{clean_sheet_name[0:2]}_{letters_of_columns[clean_left_column_name]}_{clean_right_column_name}'
#
#                             # Обчислення кореляції Пірсона для кожної пари стовпців зліва та справа кожного аркуша.
#                             correlation_coefficient, p_value = pearsonr(right_array, numerical_left_array)
#                             print(results_header, '\n')
#                             print("Pearson corellation:", correlation_coefficient)
#                             print("p-value:", p_value)
#                             # Обчислення різниці в значенні
#                             difference = np.array(right_array) - np.array(numerical_left_array)
#                             # Проведення t-тесту
#                             t_statistic, p_value_ttest = ttest_1samp(difference, 0)
#                             print("t-statistics:", t_statistic)
#                             print("p-value for t-test:", p_value_ttest)
#                             print('\n')
#
#                             file.write(f"{results_header}\n")
#                             file.write(f"Pearson c.c: {correlation_coefficient}\n")
#                             file.write(f"p-value: {p_value}\n")
#                             file.write(f"t-stat for differences: {t_statistic}\n")
#                             file.write(f"p-value for t-test: {p_value_ttest}\n\n")

# data_preparation()
# histogram_creation()
# construction_of_graphs()
# correlation_calculation()


FIRMS_DATA = "Chestionar_Viktor.xlsx"
FIRMS_DATA2 = "Firms"

firm_data ={
    "name":"",
    "field":"",
    "employers":0,
    "progress":True,
    "perform":0
}

SHEET1 = 'Sheet1'

firms = pd.read_excel(FIRMS_DATA, sheet_name=SHEET1, usecols="A", dtype=str).dropna(how='all')
print(firms.values[1:],firms.index)
input()

avg_stuff = pd.read_excel(FIRMS_DATA, sheet_name=SHEET1, usecols="K", dtype=str)
print(avg_stuff.values[:10])

stuff = []
stuffs = []
for i,peoples in enumerate(avg_stuff.values):
    print(peoples)
    if not pd.isnull(peoples):
        print("P", i, peoples, int(peoples))
        stuff.append(int(peoples[0]))
    elif i+1 in firms.index:
        print("ok")
        stuffs.append(stuff)
        stuff = []
    # input()
stuffs.append(stuff)
print(stuffs[:5])

L =[]
for i in range(len(firms)):
    print(i,firms.values[i],stuffs[i])
    L.append((str(firms.values[i][0]),stuffs[i][0]))

print(L[:10])

#
# avg_stuff = pd.read_excel(FIRMS_DATA, sheet_name=SHEET1, usecols="K", dtype=str)
# print(avg_stuff.values[:10])
# profit = []
# p = []
# for i,pr in enumerate(avg_stuff.values):
>>>>>>> 1dee9f0a6249c4130f481f9e19bb8243a182efb9
#     print(peoples)
#     if not pd.isnull(peoples):
#         print("P", i, peoples, int(peoples))
#         stuff.append(int(peoples[0]))
#     elif i+1 in firms.index:
#         print("ok")
#         stuffs.append(stuff)
#         stuff = []
#     # input()
# stuffs.append(stuff)
# print(stuffs[:5])
<<<<<<< HEAD
#
# L =[]
# for i in range(len(firms)):
#     print(i,firms.values[i],stuffs[i])
#     L.append((str(firms.values[i][0]),stuffs[i][0]))
#
# print(L[:10])
#
# TRANS = "Chestionar_Transforemed.xlsx"
#
# years = pd.read_excel(TRANS, sheet_name="FirmAnalysis", usecols="X:AE", dtype=str).dropna(how='all')
# print(years.values[:10])
#
# years.to_excel("output.xlsx",
#              sheet_name='Sheet_name_1')
#
# input()
#
# #
# # avg_stuff = pd.read_excel(FIRMS_DATA, sheet_name=SHEET1, usecols="K", dtype=str)
# # print(avg_stuff.values[:10])
# # profit = []
# # p = []
# # for i,pr in enumerate(avg_stuff.values):
# #     print(peoples)
# #     if not pd.isnull(peoples):
# #         print("P", i, peoples, int(peoples))
# #         stuff.append(int(peoples[0]))
# #     elif i+1 in firms.index:
# #         print("ok")
# #         stuffs.append(stuff)
# #         stuff = []
# #     # input()
# # stuffs.append(stuff)
# # print(stuffs[:5])
#
# # Read the Excel file
# # df = pd.read_excel(file_path)
#
# df = pd.read_excel(FIRMS_DATA, sheet_name=SHEET1, dtype=object)
#
#
# # Find the indices of the empty rows which separate the firms
# empty_indices = df[df.isnull().all(axis=1)].index
# print(empty_indices)
#
# # Add an index at the end to handle the last segment
# empty_indices = firms.index[1:] #  empty_indices.append(pd.Index([len(df)]))
#
#
# print(empty_indices.values)
#
# # Initialize a list to store results
# firm_results = []
#
# # Initialize start index
# start_idx = 1
#
# # Iterate over the indices to separate the data for each firm
# for end_idx in empty_indices:
#     # Slice the DataFrame to get the firm's data
#     firm_data = df.iloc[start_idx+1:end_idx-1].dropna(how='all')
#     firm_data = firm_data.replace(u'\xa0', u'', regex=True).astype(float)
#     print(firm_data[:5].values)
#
#
#     if not firm_data.empty:
#         # Calculate the required metrics
#         print(firm_data['Turnover'])
#
#         t1 = firm_data['Turnover'].iloc[-1] # float("".join(firm_data['Turnover'].iloc[-1].split()))
#         t2 = firm_data['Turnover'].iloc[-1] # float("".join(firm_data['Turnover'].iloc[0].split()))
#         print("T", t1, " and", t2, float(t1))
#         turnover_growth = (t1 / t2) # ** (1 / len(firm_data)) - 1) * 100
#         print("grows", turnover_growth)
#
#         print(firm_data['Profit Net'])
#         print(firm_data['Profit Net'].values)
#         print(firm_data['Turnover'].values)
#         avg_profit_margin = (firm_data['Profit Net'].values / firm_data['Turnover'].values).mean() * 100
#         print(avg_profit_margin)
#         liabilities_to_assets = (
#                     firm_data['Liailities'] / (firm_data['Fixed assets'] + firm_data['Circulant Assets'])).mean()
#         print(liabilities_to_assets)
#         fixed_assets_growth = (firm_data['Fixed assets'].iloc[-1] / firm_data['Fixed assets'].iloc[0])  #** ( 1 / (len(firm_data)) - 1) * 100
#
#
#         current_ratio = (firm_data['Circulant Assets'] / firm_data['Liailities']).mean()
#         capital_reserves_growth = ((firm_data['Capitals and reserves'].iloc[-1] /
#                                     firm_data['Capitals and reserves'].iloc[0])) # ** (1 / (len(firm_data) - 1)) - 1) * 100
#
#         # Determine scores for each metric
#         turnover_score = min(max((turnover_growth // 5) + 1, 1), 5)
#         profit_margin_score = min(max((avg_profit_margin // 5) + 1, 1), 5)
#         liabilities_score = min(max(((1 - liabilities_to_assets) * 10), 1), 5)
#         fixed_assets_score = min(max((fixed_assets_growth // 2) + 1, 1), 5)
#         current_ratio_score = min(max((current_ratio * 2.5), 1), 5)
#         capital_reserves_score = min(max((capital_reserves_growth // 3) + 1, 1), 5)
#
#         # Analyze employees trend (here assuming a simple average trend as positive growth)
#         employee_trend = (firm_data['The average number of employees'].iloc[-1] -
#                           firm_data['The average number of employees'].iloc[0]) / len(firm_data)
#         employee_score = 5 if employee_trend > 0 else 3 if employee_trend == 0 else 1
#
#         # Calculate final score as average of all scores
#         final_score = round((turnover_score + profit_margin_score + liabilities_score + fixed_assets_score + current_ratio_score + capital_reserves_score + employee_score) / 7,0)
#
#         # Store the firm's final score
#         firm_results.append(final_score)
#
#     # Update start index for next firm
#     start_idx = end_idx + 1
#
# print(firm_results)
=======

# Read the Excel file
# df = pd.read_excel(file_path)

df = pd.read_excel(FIRMS_DATA, sheet_name=SHEET1, dtype=object)


# Find the indices of the empty rows which separate the firms
empty_indices = df[df.isnull().all(axis=1)].index
print(empty_indices)

# Add an index at the end to handle the last segment
empty_indices = firms.index[1:] #  empty_indices.append(pd.Index([len(df)]))


print(empty_indices.values)

# Initialize a list to store results
firm_results = []

# Initialize start index
start_idx = 1

# Iterate over the indices to separate the data for each firm
for end_idx in empty_indices:
    # Slice the DataFrame to get the firm's data
    firm_data = df.iloc[start_idx+1:end_idx-1].dropna(how='all')
    firm_data = firm_data.replace(u'\xa0', u'', regex=True).astype(float)
    print(firm_data[:5].values)


    if not firm_data.empty:
        # Calculate the required metrics
        print(firm_data['Turnover'])

        t1 = firm_data['Turnover'].iloc[-1] # float("".join(firm_data['Turnover'].iloc[-1].split()))
        t2 = firm_data['Turnover'].iloc[-1] # float("".join(firm_data['Turnover'].iloc[0].split()))
        print("T", t1, " and", t2, float(t1))
        turnover_growth = (t1 / t2) # ** (1 / len(firm_data)) - 1) * 100
        print("grows", turnover_growth)

        print(firm_data['Profit Net'])
        print(firm_data['Profit Net'].values)
        print(firm_data['Turnover'].values)
        avg_profit_margin = (firm_data['Profit Net'].values / firm_data['Turnover'].values).mean() * 100
        print(avg_profit_margin)
        liabilities_to_assets = (
                    firm_data['Liailities'] / (firm_data['Fixed assets'] + firm_data['Circulant Assets'])).mean()
        print(liabilities_to_assets)
        fixed_assets_growth = (firm_data['Fixed assets'].iloc[-1] / firm_data['Fixed assets'].iloc[0])  #** ( 1 / (len(firm_data)) - 1) * 100


        current_ratio = (firm_data['Circulant Assets'] / firm_data['Liailities']).mean()
        capital_reserves_growth = ((firm_data['Capitals and reserves'].iloc[-1] /
                                    firm_data['Capitals and reserves'].iloc[0])) # ** (1 / (len(firm_data) - 1)) - 1) * 100

        # Determine scores for each metric
        turnover_score = min(max((turnover_growth // 5) + 1, 1), 5)
        profit_margin_score = min(max((avg_profit_margin // 5) + 1, 1), 5)
        liabilities_score = min(max(((1 - liabilities_to_assets) * 10), 1), 5)
        fixed_assets_score = min(max((fixed_assets_growth // 2) + 1, 1), 5)
        current_ratio_score = min(max((current_ratio * 2.5), 1), 5)
        capital_reserves_score = min(max((capital_reserves_growth // 3) + 1, 1), 5)

        # Analyze employees trend (here assuming a simple average trend as positive growth)
        employee_trend = (firm_data['The average number of employees'].iloc[-1] -
                          firm_data['The average number of employees'].iloc[0]) / len(firm_data)
        employee_score = 5 if employee_trend > 0 else 3 if employee_trend == 0 else 1

        # Calculate final score as average of all scores
        final_score = round((turnover_score + profit_margin_score + liabilities_score + fixed_assets_score + current_ratio_score + capital_reserves_score + employee_score) / 7,0)

        # Store the firm's final score
        firm_results.append(final_score)

    # Update start index for next firm
    start_idx = end_idx + 1

print(firm_results)
>>>>>>> 1dee9f0a6249c4130f481f9e19bb8243a182efb9
