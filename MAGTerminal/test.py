date = "2022-01-24"

name_sech_dict = {
    "10": "КС Донское (ток)",
    "19": "КС Донское (статика)",
    "2": "КС Воронежское-2 на Север (ток)",
    "18": "КС Воронежское-2 на Север (статика)",
}

date_dict = {
    "25": f'{date} 00:00:00',
    "26": f'{date} 01:00:00',
    "27": f'{date} 02:00:00',
    "28": f'{date} 03:00:00',
    "29": f'{date} 04:00:00',
    "30": f'{date} 05:00:00',
    "31": f'{date} 06:00:00',
    "32": f'{date} 07:00:00',
    "33": f'{date} 08:00:00',
    "34": f'{date} 09:00:00',
    "35": f'{date} 10:00:00',
    "36": f'{date} 11:00:00',
    "37": f'{date} 12:00:00',
    "38": f'{date} 13:00:00',
    "39": f'{date} 14:00:00',
    "40": f'{date} 15:00:00',
    "41": f'{date} 16:00:00',
    "42": f'{date} 17:00:00',
    "43": f'{date} 18:00:00',
    "44": f'{date} 19:00:00',
    "45": f'{date} 20:00:00',
    "46": f'{date} 21:00:00',
    "47": f'{date} 22:00:00',
    "48": f'{date} 23:00:00',
}

list_ = []
file_BARS_MDP = 'ogrsech.csv'
with open(file_BARS_MDP) as file:
    file_line = file.readlines()
    for i, j in enumerate(file_line):
        name_sech = j.split(";")[0].replace(j.split(";")[0], name_sech_dict[j.split(";")[0]])
        t = (j.split(";")[1].replace(j.split(";")[1], date_dict[j.split(";")[1]]))
        list_.append([name_sech, t, j.split(";")[2]])
        print(j.split(";")[0].replace(j.split(";")[0], name_sech_dict[j.split(";")[0]]))
print(list_)

for i, j in enumerate(list_):
    print(i, j)
