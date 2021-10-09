import pandas as pd
import win32com.client
rastr = win32com.client.Dispatch('Astra.Rastr')


# Функция прирощения по узлам
def prirost_uzl(vector, indexes):
    """Function changes generation and load in nodes

    Args:
        vector(pandas dataframe): include changing nodes.
        indexes(dict): indexes of nodes.
    Returns:
        changed regime.
    """

    # Выделение генераторных узлов
    count_vector = vector.shape[0]

    rastr_cols = {
                 'pg': rastr.Tables('node').Cols('pg'),
                 'pn': rastr.Tables('node').Cols('pn'),
                 'qg': rastr.Tables('node').Cols('qg'),
                 'qn': rastr.Tables('node').Cols('qn')
    }

    # Изменение генерации в узле, условия необходимы для расчета реактивной
    # мощности
    for i in range(count_vector):
        if vector['variable'][i] == 'pg':
            if vector['tg'][i] == 0:
                # к текущей генерации узла прибавляем значение из траектории
                # g_in_n = Generation_in_node
                g_in_n = rastr_cols['pg'].Z(indexes[vector['node'][i]])
                change = vector['value'][i]
                rastr_cols['pg'].SetZ(indexes[vector['node'][i]],
                                      g_in_n + change)

            elif (vector['tg'][i] != 0 and
                  rastr_cols['qg'].Z(indexes[vector['node'][i]]) != 0):
                # считаем тангенс
                old_tg_gen = (rastr_cols['pg'].Z(indexes[vector['node'][i]]) /
                              rastr_cols['qg'].Z(indexes[vector['node'][i]]))

                g_in_n = rastr_cols['pg'].Z(indexes[vector['node'][i]])
                change = vector['value'][i]
                rastr_cols['pg'].SetZ(indexes[vector['node'][i]],
                                      g_in_n + change)

                actual_gen = rastr_cols['pg'].Z(indexes[vector['node'][i]])
                rastr_cols['qg'].SetZ(indexes[vector['node'][i]],
                                      (actual_gen / old_tg_gen))

            elif (vector['tg'][i] != 0 and
                  rastr_cols['qg'].Z(indexes[vector['node'][i]]) == 0):

                g_in_n = rastr_cols['pg'].Z(indexes[vector['node'][i]])
                change = vector['value'][i]
                rastr_cols['pg'].SetZ(indexes[vector['node'][i]],
                                      g_in_n + change)

        if vector['variable'][i] == 'pn':
            if vector['tg'][i] == 0:
                # к текущей генерации узла прибавляем значение из траектории
                # l_in_n == load_in_node
                l_in_n = rastr_cols['pn'].Z(indexes[vector['node'][i]])
                change = vector['value'][i]
                rastr_cols['pn'].SetZ(indexes[vector['node'][i]],
                                      l_in_n + change)

            elif (vector['tg'][i] != 0 and
                  rastr_cols['qn'].Z(indexes[vector['node'][i]]) != 0):
                # считаем тангенс
                old_tg_load = (rastr_cols['pn'].Z(indexes[vector['node'][i]]) /
                               rastr_cols['qn'].Z(indexes[vector['node'][i]]))

                l_in_n = rastr_cols['pn'].Z(indexes[vector['node'][i]])
                change = vector['value'][i]
                rastr_cols['pn'].SetZ(indexes[vector['node'][i]],
                                      l_in_n + change)

                actual_load = rastr_cols['pn'].Z(indexes[vector['node'][i]])
                rastr_cols['qn'].SetZ(indexes[vector['node'][i]],
                                      (actual_load / old_tg_load))

            elif (vector['tg'][i] != 0 and
                  rastr_cols['qn'].Z(indexes[vector['node'][i]]) == 0):
                l_in_n = rastr_cols['pn'].Z(indexes[vector['node'][i]])
                change = vector['value'][i]
                rastr_cols['pn'].SetZ(indexes[vector['node'][i]],
                                      l_in_n + change)


# Функция утяжеления, расчитывающая предел по статической устойчивости
def utyazhelenie(vector, path_regime, path_sech, sech, indexes):
    """Function changes regime until reaches limit of steady state stability

    Args:
        vector (pandas dataframe): include changing nodes.
        path_regime (str): path to shablon .rg2.
        path_sech (str): path to shablon .sch.
        sech (pandas dataframe): include flowgate.
        indexes(dict): indexes of nodes.
    Return:
        float: limit power flow.
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
    for null, contents in sech.items():
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
        prirost_uzl(vector, indexes)
        # Расчет УР
        result = rastr.rgm('p')
    p_predel = round(rastr.Tables('sechen').Cols('psech').Z(0), 2)
    return p_predel


# Функция утяжеления, расчитывающая устойчивость нагрузки по напряжению
# Аналогичная функия, но также учитываются напряжения в узле,
# функции необходимо передать траекторию утяжеления
# и указать по какому критерию выполнять расчет, 10% или 15%
def utyazhelenie_u(vector, path_regime, koeff, off, indexes):
    """Function changes regime until reaches limit of steady state stability
    by voltage in nodes

    Args:
        vector (pandas dataframe): include changing nodes.
        path_regime (str): path to shablon .rg2.
        koeff (float): shows margin by voltage in nodes.
        off (int): 0 if normal mode, 1 if alert state.
        indexes(dict): indexes of nodes.
    Return:
        float: limit power flow.
    """
    if off == 0:
        # Загрузка файлов в Rastr
        rastr.Load(1, 'regime.rg2', path_regime)
        result = rastr.rgm('p')
    # Нахождение Минимальных напряжений
    count_node = rastr.Tables('node').Size
    node_min_voltage = []
    for node in range(count_node):
        actual_voltage = rastr.Tables('node').Cols('vras').Z(node)
        critical_voltage = rastr.Tables('node').Cols('uhom').Z(node) * 0.7
        voltage_margin = critical_voltage * koeff
        if actual_voltage < voltage_margin:
            node_min_voltage.append(rastr.Tables('node').Cols('ny').Z(node))
    if len(node_min_voltage) != 0:
        print('Недопустимые напряжения в узлах')
        raise SystemExit
    # Утяжеление

    result = rastr.rgm('p')

    while result == 0 and len(node_min_voltage) == 0:
        prirost_uzl(vector, indexes)
        result = rastr.rgm('p')
        # Проверка узлов на отклонение напряжения
        count_node = rastr.Tables('node').Size
        node_min_voltage = []
        for node in range(count_node):
            actual_voltage = rastr.Tables('node').Cols('vras').Z(node)
            critical_voltage = rastr.Tables('node').Cols('uhom').Z(node) * 0.7
            voltage_margin = critical_voltage * koeff
            ny = rastr.Tables('node').Cols('ny')
            if actual_voltage < voltage_margin:
                node_min_voltage.append(ny.Z(node))
    p_predel_u = round(rastr.Tables('sechen').Cols('psech').Z(0), 2)
    return p_predel_u


# Функция утяжеления по току
def utyazhelenie_i(vector, path_regime, current_control, off, indexes):
    """Function changes regime until reaches thermal limits of line

    Args:
        vector (pandas dataframe): include changing nodes.
        path_regime (str): path to shablon .rg2.
        current_control (str): choose control parameter,
        zag_i or zag_i_av.
        off (int): 0 if normal mode, 1 if alert state.
        indexes(dict): indexes of nodes.
    Return:
        float: limit power flow.
    """

    tok_max = []
    if off == 0:
        # Загрузка файлов в Rastr
        rastr.Load(1, 'regime.rg2', path_regime)
        result = rastr.rgm('p')
    result = rastr.rgm('p')
    while result == 0 and len(tok_max) == 0:
        prirost_uzl(vector, indexes)
        result = rastr.rgm('p')
        # Проверка ветвей на токовую нагрузку
        branch = rastr.Tables('vetv')
        count_vetv = branch.Size
        for actual_branch in range(count_vetv):
            actual_current = branch.Cols(current_control).Z(actual_branch)
            if actual_current * 1000 >= 100:
                tok_max.append(branch.Cols(current_control).Z(actual_branch))
    p_predel_i = round(rastr.Tables('sechen').Cols('psech').Z(0), 2)
    return p_predel_i


def outage(path_regime, faults, z):
    """Function iterated given faults

    Args:
        path_regime (str): path to shablon .rg2.
        faults (pandas dataframe): line outage.
        z (int): number of faults.
        indexes(dict): indexes of nodes.
    Return:
        float: limit power flow.
    """

    # Загрузка файлов в Rastr
    rastr.Load(1, 'regime.rg2', path_regime)
    rastr.rgm('p')
    count_vetv = rastr.Tables('vetv').Size

    # Перебор всех ветвей, если отключаемая ветвь совпадает с перебираемой,
    # меняем ее состояние на отключенное
    for vetv in range(count_vetv):
        ip = rastr.Tables('vetv').Cols('ip').Z(vetv)
        iq = rastr.Tables('vetv').Cols('iq').Z(vetv)
        np = rastr.Tables('vetv').Cols('np').Z(vetv)
        if (ip == faults['ip'][z] and iq == faults['iq'][z] and
                np == faults['np'][z]):
            rastr.Tables('vetv').Cols('sta').SetZ(vetv, 1)
            off_vetv = vetv
            rastr.rgm('p')
    rastr.rgm('p')
    return off_vetv


def alert_state(faults, path_regime, vector, path_sech, sech, indexes):
    """Function changes regime until reaches limit of steady
    state stability in alert state

    Args:
        faults (pandas dataframe): line outage.
        path_regime (str): path to shablon .rg2.
        vector (pandas dataframe): include changing nodes.
        path_sech (str): path to shablon .sch.
        sech (pandas dataframe): include flowgate.
        indexes(dict): indexes of nodes.
    Return:
        float: limit power flow.
    """
    # Заданные возмущения
    faults = faults.T
    faults_shape = faults.shape
    count_faults = faults_shape[0]
    # Утяжеление
    mdp_alert_state = pd.DataFrame(columns=['MDP'])

    # Цикл, делающий перебор и расчет режимов для каждого возмущения в
    # заданном списке
    for num_faults in range(count_faults):
        off_vetv = outage(path_regime, faults, num_faults)
        # Находим переток соответствующий 8% запасу
        power_alert_state = abs(
            utyazhelenie_light(
                vector,
                path_sech,
                sech,
                indexes) * 0.92)
        actual_power = 0

        # Заново загружаем режим
        rastr.Load(1, 'regime.rg2', path_regime)
        rastr.rgm('p')
        # отключаем ЛЭП снова
        rastr.Tables('vetv').Cols('sta').SetZ(off_vetv, 1)
        rastr.rgm('p')

        # Выставляем в сечении переток, соответствующий 8% запасу
        power_flowgate = rastr.Tables('sechen').Cols('psech')
        while actual_power < power_alert_state:
            prirost_uzl(vector, indexes)
            rastr.rgm('p')
            actual_power = abs(round(power_flowgate.Z(0), 2))

        # Включаем отключенную ветвь
        rastr.rgm('p')
        rastr.Tables('vetv').Cols('sta').SetZ(off_vetv, 0)
        rastr.rgm('p')
        # Расчитываем МДП и записываем в датафрейм
        mdp_alert_state.loc[num_faults] = [
            abs(round(rastr.Tables('sechen').Cols('psech').Z(0), 2)) - 30]

    # Находим наименьший переток
    power_mdp3 = mdp_alert_state['MDP'].min()
    return power_mdp3


# Функция утяжеления не загружаящая шаблон, необходим для расчета 3 критерия
def utyazhelenie_light(vector, path_sech, sech, indexes):
    """Function changes regime until reaches limit of steady state stability

    Args:
        vector (pandas dataframe): include changing nodes.
        path_regime (str): path to shablon .rg2.
        path_sech (str): path to shablon .sch.
        sech (pandas dataframe): include flowgate.
        indexes(dict): indexes of nodes.
    Return:
        float: limit power flow.
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
        prirost_uzl(vector, indexes)
        # Расчет УР
        result = rastr.rgm('p')
    p_predel = round(rastr.Tables('sechen').Cols('psech').Z(0), 2)
    return p_predel

# Функция считающая ПАВ по напряжению


def voltage_alert_state(faults, path_regime, vector, koeff, indexes):
    """Function changes regime until reaches limit of steady state stability
    by voltage in nodes in alert state

    Args:
        faults (pandas dataframe): line outage.
        vector (pandas dataframe): include changing nodes.
        path_regime (str): path to shablon .rg2.
        koeff (float): shows margin by voltage in nodes.
        indexes(dict): indexes of nodes.
    Return:
        float: limit power flow.
    """
    # Заданные возмущения
    faults = faults.T
    faults_shape = faults.shape
    count_faults = faults_shape[0]
    # Утяжеление
    mdp_voltage_alert_state = pd.DataFrame(columns=['MDP'])
    off = 1
    # Цикл, делающий перебор и расчет режимов для каждого возмущения в
    # заданном списке
    for num_faults in range(count_faults):
        off_vetv = outage(path_regime, faults, num_faults)
        # Находим 10% запас по U
        utyazhelenie_u(vector, path_regime, koeff, off, indexes)

        # Включаем отключенную ветвь
        rastr.rgm('p')
        rastr.Tables('vetv').Cols('sta').SetZ(off_vetv, 0)
        rastr.rgm('p')
        # Расчитываем МДП и записываем в датафрейм
        mdp_voltage_alert_state.loc[num_faults] = [
            abs(round(rastr.Tables('sechen').Cols('psech').Z(0), 2)) - 30]

    # Находим наименьший переток
    p_voltage_alert_state = mdp_voltage_alert_state['MDP'].min()
    return p_voltage_alert_state


# Функция считающая ПАВ по напряжению
def current_alert_state(faults, path_regime, vector, current_control, indexes):
    """Function changes regime until reaches thermal limits of line in alert state

    Args:
        faults (pandas dataframe): line outage.
        vector (pandas dataframe): include changing nodes.
        path_regime (str): path to shablon .rg2.
        current_control (str): choose control parameter,
        normal or alert current.
        indexes(dict): indexes of nodes.
    Return:
        float: limit power flow.
    """
    # Заданные возмущения
    faults = faults.T
    faults_shape = faults.shape
    count_faults = faults_shape[0]
    # Утяжеление
    mdp_current_alert_state = pd.DataFrame(columns=['MDP'])

    # Цикл, делающий перебор и расчет режимов для каждого возмущения в
    # заданном списке
    for z in range(count_faults):
        off = 1
        off_vetv = outage(path_regime, faults, z)
        # Находим предел по АДТН
        (utyazhelenie_i(vector, path_regime, current_control, off, indexes))

        # Включаем отключенную ветвь
        rastr.rgm('p')
        rastr.Tables('vetv').Cols('sta').SetZ(off_vetv, 0)
        rastr.rgm('p')
        # Расчитываем МДП и записываем в датафрейм
        mdp_current_alert_state.loc[z] = [
            abs(round(rastr.Tables('sechen').Cols('psech').Z(0), 2)) - 30]
    # Находим наименьший переток
    p_current_alert_state = mdp_current_alert_state['MDP'].min()
    return p_current_alert_state
