import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

#read file
file = pd.read_excel("Отчет.xlsx",
                     sheet_name="Брак", usecols="E,N")

#select individual elements
columns = []
for i in file["Дефект"]:
    if not i in columns:
        columns.append(i)

rows = []
for i in file["Станок"]:
    if not i in rows:
        rows.append(i)

#process file
results = {}
for i in rows:
    results[i] = [0] * len(columns)

for machine, problem in zip(file["Станок"], file["Дефект"]):
    results[machine][columns.index(problem)] += 1

#create dataframe
data = pd.DataFrame(results, index=columns).transpose()

#sort by rows
data_sum = data.sum(axis=1).sort_values(ascending=False)
data = data.reindex(index=data_sum.index)

#sort by columns
data_sum = data.sum(axis=0).sort_values(ascending=False)
data = data.reindex(columns=data_sum.index)

#select only 10 rows from dataframe
SIZE = 10
data = data.iloc[ : SIZE, :]

#get rid of useless info
to_pop = []
zeros = pd.Series([0] * SIZE, index=data.index)
for name, column in data.iteritems():
    if column.equals(zeros):
        to_pop.append(name)

data = data.drop(to_pop, axis=1)

#create colormap
cmap = plt.get_cmap('rainbow')(np.linspace(0, 1, len(data.columns)))

#build barh
data_cum = data.cumsum(axis=1)
starts = data_cum - data

for problem, color in zip(data.columns, cmap):
    plt.barh(data.index, data[problem], left=starts[problem], color=color)

    for name, value in zip(data.index, data[problem]):
        if value != 0:
            plt.text(starts.loc[name, problem] + (value / 2),
                     name, str(value), color='lightgrey',
                     ha='center', va='center')

for name, value in data.sum(axis=1).iteritems():
    plt.text(value + 0.2, name, str(value),
             ha='left', va='center')

#set limit so text doesn't go outside the figure
plt.xlim(0, data.sum(axis=1).max() + 10 - (data.sum(axis=1).max() % 10))

#build legend and title
plt.legend(data.columns)
plt.title(f"Топ {SIZE} дефектных станков")

#show
plt.show()
