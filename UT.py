import pandas as pd
import win32com.client
rastr = win32com.client.Dispatch('Astra.Rastr')


# Функция прирощения по узлам
def prirost_uzl1(vector):
    """Function changes generation and load in nodes

    :param vector: include changing nodes
    :type: pandas dataframe
    :return changed regime
    """

    # Выделение генераторных узлов
    count_vector = vector.shape[0]

    i = 0

    pg = rastr.Tables('node').Cols('pg')
    pn = rastr.Tables('node').Cols('pn')
    qg = rastr.Tables('node').Cols('qg')
    qn = rastr.Tables('node').Cols('qn')
    # Изменение генерации в узле, условия необходимы для расчета реактивной
    # мощности
    while i < count_vector:
        if vector['variable'][i] == 'pg':
            if vector['tg'][i] == 0:
                # к текущей генерации узла прибавляем значение из траектории
                pg.SetZ(
                    vector['node'][i] -
                    1,
                    pg.Z(
                        vector['node'][i] -
                        1) +
                    vector['value'][i])
            elif vector['tg'][i] != 0 and qg.Z(vector['node'][i] - 1) != 0:
                # считаем тангенс
                old_tg1 = (
                    pg.Z(
                        vector['node'][i] -
                        1) /
                    qg.Z(
                        vector['node'][i] -
                        1))
                pg.SetZ(
                    vector['node'][i] -
                    1,
                    pg.Z(
                        vector['node'][i] -
                        1) +
                    vector['value'][i])
                y = pg.Z(vector['node'][i] - 1)
                qg.SetZ(vector['node'][i] - 1, (y / old_tg1))
            elif vector['tg'][i] != 0 and qg.Z(vector['node'][i] - 1) == 0:
                pg.SetZ(
                    vector['node'][i] -
                    1,
                    pg.Z(
                        vector['node'][i] -
                        1) +
                    vector['value'][i])
        if vector['variable'][i] == 'pn':
            if vector['tg'][i] == 0:
                # к текущей генерации узла прибавляем значение из траектории
                pn.SetZ(
                    vector['node'][i] -
                    1,
                    pn.Z(
                        vector['node'][i] -
                        1) +
                    vector['value'][i])
            elif vector['tg'][i] != 0 and qn.Z(vector['node'][i] - 1) != 0:
                # считаем тангенс
                old_tg = (
                    pn.Z(
                        vector['node'][i] -
                        1) /
                    qn.Z(
                        vector['node'][i] -
                        1))
                pn.SetZ(
                    vector['node'][i] -
                    1,
                    pn.Z(
                        vector['node'][i] -
                        1) +
                    vector['value'][i])
                y = pn.Z(vector['node'][i] - 1)
                qn.SetZ(vector['node'][i] - 1, (y / old_tg))
            elif vector['tg'][i] != 0 and qn.Z(vector['node'][i] - 1) == 0:
                pn.SetZ(
                    vector['node'][i] -
                    1,
                    pn.Z(
                        vector['node'][i] -
                        1) +
                    vector['value'][i])
        i = i + 1


# Функция утяжеления, расчитывающая предел по статической устойчивости
def utyazhelenie(vector, path_regime, path_sech, sech):
    """Function changes regime until reaches limit of steady state stability

    :param vector: include changing nodes (type: pandas dataframe)
    :param path_regime: path to shablon .rg2
    :param path_sech: path to shablon .sch
    :param sech: include flowgate (type: pandas dataframe)
    :return limit power flow
    """
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
    for index, contents in sech.items():
        rastr.Tables('grline').AddRow()
        rastr.Tables('grline').Cols('ns').SetZ(i, 333)
        rastr.Tables('grline').Cols('ip').SetZ(i, contents[0])
        rastr.Tables('grline').Cols('iq').SetZ(i, contents[1])
        i = i + 1
    # Утяжеление
    result = rastr.rgm('p')
    # В данном случае используется изменение генерации и потребления в таблице
    # Узлы
    while result == 0:
        prirost_uzl1(vector)
        # Расчет УР
        result = rastr.rgm('p')
    p_predel = round(rastr.Tables('sechen').Cols('psech').Z(0), 2)
    return p_predel


# Функция утяжеления, расчитывающая устойчивость нагрузки по напряжению
# Аналогичная функия, но также учитываются напряжения в узле, функции необходимо передать траекторию утяжеления
# и указать по какому критерию выполнять расчет, 10% или 15%
def utyazhelenie_u(vector, path_regime, koeff, off):
    """Function changes regime until reaches limit of steady state stability by voltage in nodes

    :param vector: include changing nodes (type: pandas dataframe)
    :param path_regime: path to shablon .rg2
    :param koeff: type float, shows margin by voltage in nodes
    :param off: type int, 0 if normal mode, 1 if alert state
    :return limit power flow
    """
    if off == 0:
        # Загрузка файлов в Rastr
        rastr.Load(1, 'regime.rg2', path_regime)
        result = rastr.rgm('p')
    # Нахождение Минимальных напряжений
    uzl = 0
    count_uzl = rastr.Tables('node').Size
    uzl_min = []
    while uzl < count_uzl:
        u = rastr.Tables('node').Cols('vras').Z(uzl)
        u_kr = rastr.Tables('node').Cols('uhom').Z(uzl) * 0.7
        u_1_10 = u_kr * koeff
        if u < u_1_10:
            uzl_min.append(rastr.Tables('node').Cols('ny').Z(uzl))
        uzl = uzl + 1
    if len(uzl_min) != 0:
        print('Недопустимые напряжения в узлах')
        raise SystemExit
    # Утяжеление

    result = rastr.rgm('p')

    while result == 0 and len(uzl_min) == 0:
        prirost_uzl1(vector)
        result = rastr.rgm('p')
        # Проверка узлов на отклонение напряжения
        uzl = 0
        count_uzl = rastr.Tables('node').Size
        uzl_min = []
        while uzl < count_uzl:
            U = rastr.Tables('node').Cols('vras').Z(uzl)
            U_kr = rastr.Tables('node').Cols('uhom').Z(uzl) * 0.7
            U_1_10 = U_kr * koeff
            if U < U_1_10:
                uzl_min.append(rastr.Tables('node').Cols('ny').Z(uzl))
            uzl = uzl + 1
    p_predel_u = round(rastr.Tables('sechen').Cols('psech').Z(0), 2)
    return p_predel_u


# Функция утяжеления по току
def utyazhelenie_i(vector, path_regime, I, off):
    """Function changes regime until reaches thermal limits of line

    :param vector: include changing nodes (type: pandas dataframe)
    :param path_regime: path to shablon .rg2
    :param I: type string, choose control parameter, normal or alert current
    :param off: type int, 0 if normal mode, 1 if alert state
    :return limit power flow
    """

    tok_max = []
    if off == 0:
        # Загрузка файлов в Rastr
        rastr.Load(1, 'regime.rg2', path_regime)
        result = rastr.rgm('p')
    result = rastr.rgm('p')
    while result == 0 and len(tok_max) == 0:
        prirost_uzl1(vector)
        result = rastr.rgm('p')
        # Проверка ветвей на токовую нагрузку
        count_vetv = rastr.Tables('vetv').Size
        uzl_i = 0
        while uzl_i < count_vetv:
            I_v = rastr.Tables('vetv').Cols(I).Z(uzl_i)
            if I_v * 1000 >= 100:
                tok_max.append(rastr.Tables('vetv').Cols(I).Z(uzl_i))
            uzl_i = uzl_i + 1
    p_predel_i = round(rastr.Tables('sechen').Cols('psech').Z(0), 2)
    return p_predel_i


def outage(path_regime, faults, z):
    """Function iterated given faults

    :param path_regime: path to shablon .rg2
    :param faults: type pandas dataframe, choose control parameter, normal or alert current
    :param z: type int, number of faults
    :return limit power flow
    """

    # Загрузка файлов в Rastr
    rastr.Load(1, 'regime.rg2', path_regime)
    result = rastr.rgm('p')
    vetv = 0
    count_vetv = rastr.Tables('vetv').Size

    # Перебор всех ветвей, если отключаемая ветвь совпадает с перебираемой,
    # меняем ее состояние на отключенное
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


def pav(faults, path_regime, vector, path_sech, sech):
    """Function changes regime until reaches limit of steady state stability in alert state

    :param faults: type pandas dataframe, choose control parameter, normal or alert current
    :param path_regime: path to shablon .rg2
    :param vector: include changing nodes (type: pandas dataframe)
    :param path_sech: path to shablon .sch
    :param sech: include flowgate (type: pandas dataframe)
    :return limit power flow
    """
    # Заданные возмущения
    faults = faults.T
    faults_shape = faults.shape
    count_faults = faults_shape[0]
    # Утяжеление
    z = 0
    mdp_pav = pd.DataFrame(columns=['MDP'])

    # Цикл, делающий перебор и расчет режимов для каждого возмущения в
    # заданном списке
    while z < count_faults:
        off_vetv = outage(path_regime, faults, z)
        # Находим переток соответствующий 8% запасу
        p_pav = abs(
            utyazhelenie_light(
                vector,
                path_regime,
                path_sech,
                sech) * 0.92)
        p = 0

        # Заново загружаем режим
        rastr.Load(1, 'regime.rg2', path_regime)
        result = rastr.rgm('p')
        # отключаем ЛЭП снова
        rastr.Tables('vetv').Cols('sta').SetZ(off_vetv, 1)
        result = rastr.rgm('p')

        # Выставляем в сечении переток, соответствующий 8% запасу
        while p < p_pav:
            prirost_uzl1(vector)
            result = rastr.rgm('p')
            p = abs(round(rastr.Tables('sechen').Cols('psech').Z(0), 2))

        # Включаем отключенную ветвь
        result = rastr.rgm('p')
        rastr.Tables('vetv').Cols('sta').SetZ(off_vetv, 0)
        result = rastr.rgm('p')
        # Расчитываем МДП и записываем в датафрейм
        mdp_pav.loc[z] = [
            abs(round(rastr.Tables('sechen').Cols('psech').Z(0), 2)) - 30]
        z = z + 1

    # Находим наименьший переток
    P_mdp3 = mdp_pav['MDP'].min()
    return P_mdp3


# Функция утяжеления не загружаящая шаблон, необходим для расчета 3 критерия
def utyazhelenie_light(vector, path_regime, path_sech, sech):
    """Function changes regime until reaches limit of steady state stability

    :param vector: include changing nodes (type: pandas dataframe)
    :param path_regime: path to shablon .rg2
    :param path_sech: path to shablon .sch
    :param sech: include flowgate (type: pandas dataframe)
    :return limit power flow
    """
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
    # В данном случае используется изменение генерации и потребления в таблице
    # Узлы
    while result == 0:
        prirost_uzl1(vector)
        # Расчет УР
        result = rastr.rgm('p')
    p_predel = round(rastr.Tables('sechen').Cols('psech').Z(0), 2)
    return p_predel

# Функция считающая ПАВ по напряжению


def pav_u(faults, path_regime, vector, koeff):
    """Function changes regime until reaches limit of steady state stability by voltage in nodes in alert state

    :param faults: type pandas dataframe, choose control parameter, normal or alert current
    :param vector: include changing nodes (type: pandas dataframe)
    :param path_regime: path to shablon .rg2
    :param koeff: type float, shows margin by voltage in nodes
    :return limit power flow
    """
    # Заданные возмущения
    faults = faults.T
    faults_shape = faults.shape
    count_faults = faults_shape[0]
    # Утяжеление
    z = 0
    mdp_pav_u = pd.DataFrame(columns=['MDP'])
    off = 1
    # Цикл, делающий перебор и расчет режимов для каждого возмущения в
    # заданном списке
    while z < count_faults:
        off_vetv = outage(path_regime, faults, z)
        # Находим 10% запас по U
        p_pav_u = abs(utyazhelenie_u(vector, path_regime, koeff, off))

        # Включаем отключенную ветвь
        result = rastr.rgm('p')
        rastr.Tables('vetv').Cols('sta').SetZ(off_vetv, 0)
        result = rastr.rgm('p')
        # Расчитываем МДП и записываем в датафрейм
        mdp_pav_u.loc[z] = [
            abs(round(rastr.Tables('sechen').Cols('psech').Z(0), 2)) - 30]
        z = z + 1

    # Находим наименьший переток
    P_mdp4 = mdp_pav_u['MDP'].min()
    return P_mdp4


# Функция считающая ПАВ по напряжению
def pav_i(faults, path_regime, vector, I):
    """Function changes regime until reaches thermal limits of line in alert state

    :param faults: type pandas dataframe, choose control parameter, normal or alert current
    :param vector: include changing nodes (type: pandas dataframe)
    :param path_regime: path to shablon .rg2
    :param I: type string, choose control parameter, normal or alert current
    :return limit power flow
    """
    # Заданные возмущения
    faults = faults.T
    faults_shape = faults.shape
    count_faults = faults_shape[0]
    # Утяжеление
    z = 0
    mdp_pav_i = pd.DataFrame(columns=['MDP'])

    # Цикл, делающий перебор и расчет режимов для каждого возмущения в
    # заданном списке
    while z < count_faults:
        off = 1
        off_vetv = outage(path_regime, faults, z)
        # Находим предел по АДТН
        p_pav_i = abs(utyazhelenie_i(vector, path_regime, I, off))

        # Включаем отключенную ветвь
        result = rastr.rgm('p')
        rastr.Tables('vetv').Cols('sta').SetZ(off_vetv, 0)
        result = rastr.rgm('p')
        # Расчитываем МДП и записываем в датафрейм
        mdp_pav_i.loc[z] = [
            abs(round(rastr.Tables('sechen').Cols('psech').Z(0), 2)) - 30]
        z = z + 1
    # Находим наименьший переток
    P_mdp5_2 = mdp_pav_i['MDP'].min()
    return P_mdp5_2
