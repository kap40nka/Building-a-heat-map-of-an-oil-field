import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns
from scipy.spatial import cKDTree
import xlwings as xw

fig, axs = plt.subplots(1, 2, figsize=(8,4))
grid = plt.grid(True)
sns.set_theme()

# Извлечение данных
extraction_values = input("УКАЖИТЕ ТОЧНЫЙ ПУТЬ К ДАННЫМ СЕЙСМИЧЕСКОЙ СЕТКИ: ")
well_coord = input("УКАЖИТЕ ТОЧНЫЙ ПУТЬ К КООРДИНАТАМ СКВАЖИН: ")
data1 = pd.read_table(extraction_values, sep=' ', header=None, names=['x', 'y', 'z'])
data2 = pd.read_excel(well_coord,  header=None, names=['x', 'y', 'well'])
data2 = data2.drop(index=0)
data2_copy = data2.copy()
data2_copy = data2_copy.drop(columns='well')

# создаем графики
sns.heatmap(data1.pivot_table(index='y', columns='x', values='z'), ax=axs[1])
sns.scatterplot(x=data2['x'], y=data2['y'], color='black', ax=axs[0],hue= data2['well'])

#для каждой скважины находим координаты ближайшего узла и создаем data3, содержащую координаты этих узлов
tree = cKDTree(data1[['x', 'y']])
distances, indices = tree.query(data2_copy, k=1)
data3 = pd.DataFrame({'well': data2.reset_index(drop=True)['well'],
                      'x': np.zeros(len(data2)),
                      'y': np.zeros(len(data2)),
                      'z': np.zeros(len(data2))})

for i in range(len(data3)): 
    data3.loc[i, 'z'] =data1.loc[indices[i], 'z'] 
    data3.loc[i, 'x'] =data1.loc[indices[i], 'x'] 
    data3.loc[i, 'y'] =data1.loc[indices[i], 'y']
print('Координаты ближайших узлов для каждой скважины: ')
print(data3)

#настройка отображения
plt.gca().invert_yaxis()
axs[0].set_facecolor('none')
axs[0].set_position([0.1, 0.1, 0.7, 0.8])
axs[1].set_position([0.1, 0.1, 0.7, 0.8])
axs[0].set_zorder(2)
axs[1].set_zorder(1)
axs[1].set_xlabel('X')
axs[1].set_ylabel('Y')
axs[0].set_axis_off()
plt.tight_layout()
plt.title('Тепловая карта')
plt.show()



# вывод данных в отдельный лист документа well_coord.xlsx
sheet_df_mapping = {"Координаты узлов": data3}

with xw.App(visible=False) as app:
    wb = app.books.open(well_coord)


    current_sheets = [sheet.name for sheet in wb.sheets]


    for sheet_name in sheet_df_mapping.keys():
        if sheet_name in current_sheets:
            wb.sheets(sheet_name).range("A1").value = sheet_df_mapping.get(sheet_name)
        else:
            new_sheet = wb.sheets.add(after=wb.sheets.count)
            new_sheet.range("A1").value = sheet_df_mapping.get(sheet_name)
            new_sheet.name = sheet_name
    wb.save()
    wb.close()