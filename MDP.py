import pandas as pd
import win32com.client
rastr = win32com.client.Dispatch('Astra.Rastr')
# Импорт своих модулей
import UT

# Указываем путь к шаблону растра на компьютере
path_regime = 'C:/Users/Fanit/Documents/RastrWin3/SHABLON/режим.rg2'
path_sech = 'C:/Users/Fanit/Documents/RastrWin3/SHABLON/сечения.sch'

# Считываем заданный файл, в котором перечисленны элементы сечения
sech = pd.read_json('flowgate.json')

# Формирование траектории утяжеления
vector_UT = pd.read_csv('vector.csv')

# Нерегулярные колебания
P_nk = 30

# Расчет МДП по Обеспечению 20% запаса статической апериодической устойчивости в КС в нормальной схеме.
P_mdp1 = round(abs(UT.utyazhelenie(vector_UT, path_regime, path_sech, sech)) * (1 - 0.2) - P_nk, 2)

# Расчет МДП по Обеспечению 15% коэффициента запаса статической устойчивости по напряжению в узлах нагрузки в нормальной схеме.
P_mdp2 = round(abs(UT.utyazhelenie_U(vector_UT, path_regime, 1.15, 0)) - P_nk, 2)

# Расчет МДП по Обеспечению 8% запаса статической апериодической устойчивости в КС в ПАВ.
# Заданные возмущения
faults = pd.read_json('faults.json')
# Считаем переток, соответствующий 8% запасу
P_mdp3 = UT.PAV(faults, path_regime, vector_UT, path_sech, sech)

# Расчет МДП по обеспечению 10% запаса по U в ПАВ
P_mdp4 = UT.PAV_U(faults, path_regime, vector_UT, 1.1)

# Расчет МДП по ДДТН
P_mdp5_1 = round(abs(UT.utyazhelenie_I(vector_UT, path_regime, 'zag_i', 0)) - P_nk, 2)

# Расчет МДП по АДТН
P_mdp5_2 = UT.PAV_I(faults, path_regime, vector_UT, 'zag_i_av')

all_result = pd.DataFrame(columns = ['Критерий определения перетока', 'Максимальный допустимый переток, МВт'])
mdp1 = {'Критерий определения перетока':'Обеспечение 20% коэффициента запаса статической устойчивости в нормальной схеме', 
    'Максимальный допустимый переток, МВт':P_mdp1}
mdp2 = {'Критерий определения перетока':'Обеспечение 15% коэффициента запаса по напряжению в узлах нагрузки в нормальной схеме', 
    'Максимальный допустимый переток, МВт':P_mdp2}
mdp3 = {'Критерий определения перетока':'Обеспечение 8% коэффициента запаса статической  устойчивости в послеаварийных режимах' 
    'после нормативных возмущений', 'Максимальный допустимый переток, МВт':P_mdp3}
mdp4 = {'Критерий определения перетока':'Обеспечение 10% коэффициента запаса по напряжению в узлах нагрузки в послеаварийных режимах'
    'после нормативных возмущений', 'Максимальный допустимый переток, МВт':P_mdp4}
mdp5_1 = {'Критерий определения перетока':'Обеспечение ДДТН линий электропередачи и электросетевого оборудования в нормальной схеме', 
    'Максимальный допустимый переток, МВт':P_mdp5_1}
mdp5_2 = {'Критерий определения перетока':'Обеспечение АДТН линий электропередачи и электросетевого оборудования в послеаварийных'
    'режимах после нормативных возмущений', 'Максимальный допустимый переток, МВт':P_mdp5_2}
all_result = all_result.append(mdp1, ignore_index = True)
all_result = all_result.append(mdp2, ignore_index = True)
all_result = all_result.append(mdp3, ignore_index = True)
all_result = all_result.append(mdp4, ignore_index = True)
all_result = all_result.append(mdp5_1, ignore_index = True)
all_result = all_result.append(mdp5_2, ignore_index = True)
print(all_result)