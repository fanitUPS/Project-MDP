import conculations_powerflow as cp
import pandas as pd
import win32com.client
rastr = win32com.client.Dispatch('Astra.Rastr')

# Указываем путь к шаблону растра на компьютере
path_regime = 'C:/Users/Fanit/Documents/RastrWin3/SHABLON/режим.rg2'
path_sech = 'C:/Users/Fanit/Documents/RastrWin3/SHABLON/сечения.sch'

# Считываем заданный файл, в котором перечисленны элементы сечения
sech = pd.read_json('flowgate.json', ).T

# Формирование траектории утяжеления
vector_ut = pd.read_csv('vector.csv')

# Нерегулярные колебания
p_nk = 30

# Загрузка файлов в Rastr
rastr.Load(1, 'regime.rg2', path_regime)
rastr.rgm('p')
# Находим индексы узлов
rastr_node = rastr.Tables('node').Size
indexes = {}
for node in vector_ut['node']:
    for index in range(rastr_node):
        if node == rastr.Tables('node').Cols('ny').Z(index):
            indexes[node] = index

# Расчет МДП по Обеспечению 20% запаса статической апериодической
# устойчивости в КС в нормальной схеме.
p_mdp1 = round(abs(cp.utyazhelenie(vector_ut, path_regime,
               path_sech, sech, 0, indexes)) * (1 - 0.2) - p_nk, 2)

# Расчет МДП по Обеспечению 15% коэффициента запаса статической
# устойчивости по напряжению в узлах нагрузки в нормальной схеме.
p_mdp2 = round(abs(cp.utyazhelenie_u(
               vector_ut, path_regime, 1.15, 0, indexes)) - p_nk, 2)

# Расчет МДП по Обеспечению 8% запаса статической апериодической
# устойчивости в КС в ПАВ.
# Заданные возмущения
faults = pd.read_json('faults.json').T
# Считаем переток, соответствующий 8% запасу
p_mdp3 = cp.alert_state(faults, path_regime,
                        vector_ut, path_sech, sech, indexes)

# Расчет МДП по обеспечению 10% запаса по U в ПАВ
p_mdp4 = cp.voltage_alert_state(faults, path_regime, vector_ut, 1.1, indexes)

# Расчет МДП по ДДТН
p_mdp5_1 = round(abs(cp.utyazhelenie_i(
    vector_ut, path_regime, 'zag_i', 0, indexes)) - p_nk, 2)

# Расчет МДП по АДТН
p_mdp5_2 = cp.current_alert_state(faults, path_regime,
                                  vector_ut, 'zag_i_av', indexes)

result = {
    'Критерий определения перетока': [
        'Обеспечение 20% коэффициента запаса статической устойчивости \
        в нормальной схеме',
        'Обеспечение 15% коэффициента запаса по напряжению в узлах нагрузки \
        в нормальной схеме',
        'Обеспечение 8% коэффициента запаса статической  устойчивости в \
        послеаварийных режимах',
        'Обеспечение 10% коэффициента запаса по напряжению в узлах нагрузки в\
        послеаварийных режимах',
        'Обеспечение ДДТН линий электропередачи и электросетевого оборудования в \
        нормальной схеме',
        'Обеспечение АДТН линий электропередачи и электросетевого оборудования в\
        послеаварийных'],
    'Максимальный допустимый переток, МВт': [
        p_mdp1,
        p_mdp2,
        p_mdp3,
        p_mdp4,
        p_mdp5_1,
        p_mdp5_2]}
all_result = pd.DataFrame.from_dict(result)
print(all_result)
