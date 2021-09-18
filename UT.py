import pandas as pd
import win32com.client
rastr = win32com.client.Dispatch('Astra.Rastr')


# Функция прирощения по узлам
def prirost_uzl(vector):
    # Выделение генераторных узлов
    vector_g = vector[vector['variable'] == 'pg']
    vector_g = vector_g.reset_index()
    vector_g = vector_g.drop('index', 1)
    vector_g = vector_g.drop('variable', 1)
    count_g = vector_g.shape[0]
    # Выделение нагрузочных узлов
    vector_n = vector[vector['variable'] == 'pn']
    vector_n = vector_n.reset_index()
    vector_n = vector_n.drop('index', 1)
    vector_n = vector_n.drop('variable', 1)
    count_n = vector_n.shape[0]
    i = 0
    # Изменение генерации в узле, условия необходимы для расчета реактивной мощности
    while i < count_g:
        if vector_g['tg'][i] == 0:
            # к текущей генерации узла прибавляем значение из траектории
            rastr.Tables('node').Cols('pg').SetZ(vector_g['node'][i] - 1, rastr.Tables('node').Cols('pg').Z(vector_g['node'][i] - 1)
                + vector_g['value'][i])
        elif vector_g['tg'][i] != 0 and rastr.Tables('node').Cols('qg').Z(vector_g['node'][i] - 1) != 0:
            # считаем тангенс
            old_tg1 = (rastr.Tables('node').Cols('pg').Z(vector_g['node'][i] - 1) / rastr.Tables('node').Cols('qg').Z(vector_g['node'][i] - 1))
            rastr.Tables('node').Cols('pg').SetZ(vector_g['node'][i] - 1, rastr.Tables('node').Cols('pg').Z(vector_g['node'][i] - 1)
                + vector_g['value'][i])
            y = rastr.Tables('node').Cols('pg').Z(vector_g['node'][i] - 1)
            rastr.Tables('node').Cols('qg').SetZ(vector_g['node'][i] - 1, (y / old_tg1))
        elif vector_g['tg'][i] != 0 and rastr.Tables('node').Cols('qg').Z(vector_g['node'][i] - 1) == 0:
            rastr.Tables('node').Cols('pg').SetZ(vector_g['node'][i] - 1, rastr.Tables('node').Cols('pg').Z(vector_g['node'][i] - 1)
                + vector_g['value'][i])
        i = i + 1
    j = 0
    # Изменение потребления в узлах
    while j < count_n:
        if vector_n['tg'][j] == 0:
            rastr.Tables('node').Cols('pn').SetZ(vector_n['node'][j] - 1, rastr.Tables('node').Cols('pn').Z(vector_n['node'][j] - 1)
                + vector_n['value'][j])
        elif vector_n['tg'][j] != 0 and rastr.Tables('node').Cols('qn').Z(vector_n['node'][j] - 1) != 0:
            old_tg = (rastr.Tables('node').Cols('pn').Z(vector_n['node'][j] - 1) / rastr.Tables('node').Cols('qn').Z(vector_n['node'][j] - 1))
            rastr.Tables('node').Cols('pn').SetZ(vector_n['node'][j] - 1, rastr.Tables('node').Cols('pn').Z(vector_n['node'][j] - 1)
                + vector_n['value'][j])
            x = rastr.Tables('node').Cols('pn').Z(vector_n['node'][j] - 1)
            rastr.Tables('node').Cols('qn').SetZ(vector_n['node'][j] - 1, (x / old_tg))
        elif vector_n['tg'][j] != 0 and rastr.Tables('node').Cols('qn').Z(vector_n['node'][j] - 1) == 0:
            rastr.Tables('node').Cols('pn').SetZ(vector_n['node'][j] - 1, rastr.Tables('node').Cols('pn').Z(vector_n['node'][j] - 1)
                + vector_n['value'][j])
        j = j + 1


# Функция утяжеления, расчитывающая предел по статической устойчивости
def utyazhelenie(vector, path_regime, path_sech, sech):
    
    # Загрузка файлов в Rastr
    rastr.Load(1, 'regime.rg2', path_regime)
    result = rastr.rgm('p')

    # Добавление нового сечения
    # Создаем файл сечения
    rastr.Save('regime.sch', path_sech)

    # Загружаем созданный файл в Растр
    rastr.Load(1, 'regime.sch', path_sech)
    rastr.Tables('sechen').AddRow()
    # Создаем сечение с названием 333
    rastr.Tables('sechen').Cols('ns').SetZ(0, 333)
    
    # Вносим в сечение заданные ЛЭП
    i = 0
    for label, contents in sech.items():
        rastr.Tables('grline').AddRow()
        rastr.Tables('grline').Cols('ns').SetZ(i, 333)
        rastr.Tables('grline').Cols('ip').SetZ(i, contents[0])
        rastr.Tables('grline').Cols('iq').SetZ(i, contents[1])
        i = i + 1

    # Утяжеление
    result = rastr.rgm('p')
    # В данном случае используется изменение генерации и потребления в таблице Узлы
    while result == 0:
        prirost_uzl(vector)
        # Расчет УР
        result = rastr.rgm('p')
    P_predel = round(rastr.Tables('sechen').Cols('psech').Z(0), 2)
    return P_predel


# Функция утяжеления, расчитывающая устойчивость нагрузки по напряжению
# Аналогичная функия, но также учитываются напряжения в узле, функции необходимо передать траекторию утяжеления
# и указать по какому критерию выполнять расчет, 10% или 15%
def utyazhelenie_U(vector, path_regime, koeff, off):
    if off == 0:
        # Загрузка файлов в Rastr
        rastr.Load(1, 'regime.rg2', path_regime)
        result = rastr.rgm('p')
    # Нахождение Минимальных напряжений
    uzl = 0
    count_uzl = rastr.Tables('node').Size
    Uzl_min = []
    while uzl < count_uzl:
        U = rastr.Tables('node').Cols('vras').Z(uzl)
        U_kr = rastr.Tables('node').Cols('uhom').Z(uzl) * 0.7
        U_1_10 = U_kr * koeff
        if U < U_1_10:
            Uzl_min.append(rastr.Tables('node').Cols('ny').Z(uzl))
        uzl = uzl + 1
    if len(Uzl_min) != 0:
        print('Недопустимые напряжения в узлах')
        raise SystemExit     
    # Утяжеление

    result = rastr.rgm('p')

    while result == 0 and len(Uzl_min) == 0:
        prirost_uzl(vector)
        result = rastr.rgm('p')
        # Проверка узлов на отклонение напряжения
        uzl = 0
        count_uzl = rastr.Tables('node').Size
        Uzl_min = []
        while uzl < count_uzl:
            U = rastr.Tables('node').Cols('vras').Z(uzl)
            U_kr = rastr.Tables('node').Cols('uhom').Z(uzl) * 0.7
            U_1_10 = U_kr * koeff
            if U < U_1_10:
                Uzl_min.append(rastr.Tables('node').Cols('ny').Z(uzl))
            uzl = uzl + 1
    P_predelU = round(rastr.Tables('sechen').Cols('psech').Z(0), 2)
    return P_predelU


# Функция утяжеления по току
def utyazhelenie_I(vector, path_regime, I, off):
    tok_max = []
    if off == 0:
        # Загрузка файлов в Rastr
        rastr.Load(1, 'regime.rg2', path_regime)
        result = rastr.rgm('p')
    result = rastr.rgm('p')
    while result == 0 and len(tok_max) == 0:
        prirost_uzl(vector)
        result = rastr.rgm('p')
        # Проверка ветвей на токовую нагрузку
        count_vetv = rastr.Tables('vetv').Size
        uzl_i = 0
        while uzl_i < count_vetv:
            I_v = rastr.Tables('vetv').Cols(I).Z(uzl_i)
            if I_v * 1000 >= 100:
                tok_max.append(rastr.Tables('vetv').Cols(I).Z(uzl_i))
            uzl_i = uzl_i + 1
    P_predelI = round(rastr.Tables('sechen').Cols('psech').Z(0), 2)
    return P_predelI


def outage(path_regime, faults, z):
    # Загрузка файлов в Rastr
    rastr.Load(1, 'regime.rg2', path_regime)
    result = rastr.rgm('p')
    vetv = 0
    count_vetv = rastr.Tables('vetv').Size

    # Перебор всех ветвей, если отключаемая ветвь совпадает с перебираемой, меняем ее состояние на отключенное
    while vetv < count_vetv:
        ip = rastr.Tables('vetv').Cols('ip').Z(vetv)
        iq = rastr.Tables('vetv').Cols('iq').Z(vetv)
        np = rastr.Tables('vetv').Cols('np').Z(vetv)
        if ip == faults['ip'][z] and iq == faults['iq'][z] and np == faults['np'][z]:
            rastr.Tables('vetv').Cols('sta').SetZ(vetv, 1)
            off_vetv = vetv
            result = rastr.rgm('p')
        vetv = vetv + 1
    result = rastr.rgm('p')
    return off_vetv


def PAV(faults, path_regime, vector, path_sech, sech):
    # Заданные возмущения
    faults = faults.T
    faults_shape = faults.shape
    count_faults = faults_shape[0]
    # Утяжеление
    z = 0
    mdp_PAV = pd.DataFrame(columns = ['MDP'])
 
    # Цикл, делающий перебор и расчет режимов для каждого возмущения в заданном списке
    while z < count_faults:
        off_vetv = outage(path_regime, faults, z)
        # Находим переток соответствующий 8% запасу
        P_PAV = abs(utyazhelenie_light(vector, path_regime, path_sech, sech) * 0.92)
        P = 0
        
        # Заново загружаем режим
        rastr.Load(1, 'regime.rg2', path_regime)
        result = rastr.rgm('p')
        # отключаем ЛЭП снова
        rastr.Tables('vetv').Cols('sta').SetZ(off_vetv, 1)
        result = rastr.rgm('p')

        # Выставляем в сечении переток, соответствующий 8% запасу
        while  P < P_PAV:
            prirost_uzl(vector)
            result = rastr.rgm('p')
            P = abs(round(rastr.Tables('sechen').Cols('psech').Z(0), 2))
        
        # Включаем отключенную ветвь
        result = rastr.rgm('p')
        rastr.Tables('vetv').Cols('sta').SetZ(off_vetv, 0)
        result = rastr.rgm('p')
        # Расчитываем МДП и записываем в датафрейм
        mdp_PAV.loc[z] = [abs(round(rastr.Tables('sechen').Cols('psech').Z(0), 2)) - 30]
        z = z + 1
            
    # Находим наименьший переток
    P_mdp3 = mdp_PAV['MDP'].min()
    return P_mdp3


# Функция утяжеления не загружаящая шаблон, необходим для расчета 3 критерия
def utyazhelenie_light (vector, path_regime, path_sech, sech): 
        
    # Добавление нового сечения
    # Создаем файл сечения
    rastr.Save('regime.sch', path_sech)

    # Загружаем созданный файл в Растр
    rastr.Load(1, 'regime.sch', path_sech)
    rastr.Tables('sechen').AddRow()
    # Создаем сечение с названием 333
    rastr.Tables('sechen').Cols('ns').SetZ(0, 333)
    
    # Вносим в сечение заданные ЛЭП
    i = 0
    for label, contents in sech.items():
        rastr.Tables('grline').AddRow() 
        rastr.Tables('grline').Cols('ns').SetZ(i, 333) 
        rastr.Tables('grline').Cols('ip').SetZ(i, contents[0])
        rastr.Tables('grline').Cols('iq').SetZ(i, contents[1])
        i = i + 1
    
    # Утяжеление
    result = rastr.rgm('p')
    # В данном случае используется изменение генерации и потребления в таблице Узлы
    while result == 0:
        prirost_uzl(vector)
        # Расчет УР
        result = rastr.rgm('p')
    P_predel = round(rastr.Tables('sechen').Cols('psech').Z(0), 2)  
    return P_predel

# Функция считающая ПАВ по напряжению
def PAV_U (faults, path_regime, vector, koeff):
    # Заданные возмущения
    faults = faults.T
    faults_shape = faults.shape
    count_faults = faults_shape[0]
    # Утяжеление
    z = 0
    mdp_PAVU = pd.DataFrame(columns = ['MDP'])
    off = 1
    # Цикл, делающий перебор и расчет режимов для каждого возмущения в заданном списке
    while z < count_faults:
        off_vetv = outage(path_regime, faults, z)
        # Находим 10% запас по U
        P_PAVU = abs(utyazhelenie_U(vector, path_regime, koeff, off))
                  
        # Включаем отключенную ветвь
        result = rastr.rgm('p')
        rastr.Tables('vetv').Cols('sta').SetZ(off_vetv, 0)
        result = rastr.rgm('p')
        # Расчитываем МДП и записываем в датафрейм
        mdp_PAVU.loc[z] = [abs(round(rastr.Tables('sechen').Cols('psech').Z(0), 2)) - 30]
        z = z + 1
            
    # Находим наименьший переток
    P_mdp4 = mdp_PAVU['MDP'].min()
    return P_mdp4


# Функция считающая ПАВ по напряжению
def PAV_I (faults, path_regime, vector, I):
    # Заданные возмущения
    faults = faults.T
    faults_shape = faults.shape
    count_faults = faults_shape[0]
    # Утяжеление
    z = 0
    mdp_PAVI = pd.DataFrame(columns = ['MDP'])
    
    # Цикл, делающий перебор и расчет режимов для каждого возмущения в заданном списке
    while z < count_faults:
        off = 1
        off_vetv = outage(path_regime, faults, z)
        # Находим предел по АДТН
        P_PAVI = abs(utyazhelenie_I(vector, path_regime, I, off))
                  
        # Включаем отключенную ветвь
        result = rastr.rgm('p')
        rastr.Tables('vetv').Cols('sta').SetZ(off_vetv, 0)
        result = rastr.rgm('p')
        # Расчитываем МДП и записываем в датафрейм
        mdp_PAVI.loc[z] = [abs(round(rastr.Tables('sechen').Cols('psech').Z(0), 2)) - 30]
        z = z + 1
    #Находим наименьший переток
    P_mdp5_2 = mdp_PAVI['MDP'].min()
    return P_mdp5_2