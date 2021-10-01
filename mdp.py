import ut
import pandas as pd
import win32com.client
rastr = win32com.client.Dispatch('Astra.Rastr')
# Импорт своих модулей

# Указываем путь к шаблону растра на компьютере
path_regime = 'C:/Users/Fanit/Documents/RastrWin3/SHABLON/режим.rg2'
path_sech = 'C:/Users/Fanit/Documents/RastrWin3/SHABLON/сечения.sch'

# Считываем заданный файл, в котором перечисленны элементы сечения
sech = pd.read_json('flowgate.json')

# Формирование траектории утяжеления
vector_ut = pd.read_csv('vector.csv')

# Создаем словарь
result = {}
# Нерегулярные колебания
p_nk = 30

# Расчет МДП по Обеспечению 20% запаса статической апериодической
# устойчивости в КС в нормальной схеме.
p_mdp1 = round(abs(ut.utyazhelenie(vector_ut, path_regime,
               path_sech, sech)) * (1 - 0.2) - p_nk, 2)

# Расчет МДП по Обеспечению 15% коэффициента запаса статической
# устойчивости по напряжению в узлах нагрузки в нормальной схеме.
p_mdp2 = round(abs(ut.utyazhelenie_u(
    vector_ut, path_regime, 1.15, 0)) - p_nk, 2)

# Расчет МДП по Обеспечению 8% запаса статической апериодической
# устойчивости в КС в ПАВ.
# Заданные возмущения
faults = pd.read_json('faults.json')
# Считаем переток, соответствующий 8% запасу
p_mdp3 = ut.pav(faults, path_regime, vector_ut, path_sech, sech)

# Расчет МДП по обеспечению 10% запаса по U в ПАВ
p_mdp4 = ut.pav_u(faults, path_regime, vector_ut, 1.1)

# Расчет МДП по ДДТН
p_mdp5_1 = round(abs(ut.utyazhelenie_i(
    vector_ut, path_regime, 'zag_i', 0)) - p_nk, 2)

# Расчет МДП по АДТН
p_mdp5_2 = ut.pav_i(faults, path_regime, vector_ut, 'zag_i_av')

result = {
    'Критерий определения перетока': [
        'Обеспечение 20% коэффициента запаса статической устойчивости в нормальной схеме',
        'Обеспечение 15% коэффициента запаса по напряжению в узлах нагрузки в нормальной схеме',
        'Обеспечение 8% коэффициента запаса статической  устойчивости в послеаварийных режимах',
        'Обеспечение 10% коэффициента запаса по напряжению в узлах нагрузки в послеаварийных режимах',
        'Обеспечение ДДТН линий электропередачи и электросетевого оборудования в нормальной схеме',
        'Обеспечение АДТН линий электропередачи и электросетевого оборудования в послеаварийных'],
    'Максимальный допустимый переток, МВт': [
        p_mdp1,
        p_mdp2,
        p_mdp3,
        p_mdp4,
        p_mdp5_1,
        p_mdp5_2]}
x = pd.DataFrame.from_dict(result)
print(x)
