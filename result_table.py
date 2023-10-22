from time import sleep

import pandas as pd


def write_to_table_result(score, step):
    sleep(5)
    df_result = pd.read_excel('./data/Result.xlsx', sheet_name='Result',
                              dtype={'Команда': str, 'Лига': str,
                                     'Баллы 1': float, 'Рейтинг 1': int,
                                     'Баллы 2': float, 'Рейтинг 2': int,
                                     'Баллы': float, 'Рейтинг': int})
    for command in score:
        df_result.loc[df_result['Команда'] == command, f'Баллы {step}'] = score[command][0]
        t = df_result.loc[df_result['Команда'] == command, 'Баллы']
        df_result.loc[df_result['Команда'] == command, 'Баллы'] = t + score[command][0]
        df_result.loc[df_result['Команда'] == command, f'Рейтинг {step}'] = score[command][1]
        t = df_result.loc[df_result['Команда'] == command, 'Рейтинг']
        df_result.loc[df_result['Команда'] == command, 'Рейтинг'] = t + score[command][1]
    df_result.sort_values(['Лига', 'Рейтинг'], ascending=False, inplace=True)
    df_result.to_csv('./data/Result.csv', index=False, encoding="ANSI")


def write_to_table_individual(score, step):
    sleep(5)
    df_ind = pd.read_excel('./data/Individual_Results.xlsx', sheet_name='Individual_Results',
                           dtype={'ФИО': str, 'Команда': str, 'Школа': str,
                                  'Бой 1 Д': float, 'Бой 1 О': float, 'Бой 1 Р': float,
                                  'Бой 2 Д': float, 'Бой 2 О': float, 'Бой 2 Р': float,
                                  'Бой 3 Д': float, 'Бой 3 О': float, 'Бой 3 Р': float,
                                  'Общ Д': float, 'Общ О': float, 'Общ Р': float,
                                  'Общ': float})
    for command in score:
        for i in range(2, len(score[command])):
            df_ind.loc[df_ind['ФИО'] == score[command][i][0], 'Команда'] = command
            df_ind.loc[df_ind['ФИО'] == score[command][i][0],
                       f'Бой {step} {score[command][i][1]}'] = score[command][i][2]
            df_ind.loc[df_ind['ФИО'] == score[command][i][0], f'Общ {score[command][i][1]}'] = (
                    df_ind.loc[df_ind['ФИО'] == score[command][i][0], f'Бой 1 {score[command][i][1]}'] +
                    df_ind.loc[df_ind['ФИО'] == score[command][i][0], f'Бой 2 {score[command][i][1]}'] +
                    df_ind.loc[df_ind['ФИО'] == score[command][i][0], f'Бой 3 {score[command][i][1]}'])
            df_ind.loc[df_ind['ФИО'] == score[command][i][0], f'Общ'] = (
                    df_ind.loc[df_ind['ФИО'] == score[command][i][0], f'Общ Д'] +
                    df_ind.loc[df_ind['ФИО'] == score[command][i][0], f'Общ О'] +
                    df_ind.loc[df_ind['ФИО'] == score[command][i][0], f'Общ Р'])
        df_ind.sort_values('Общ', ascending=False, inplace=True)
        df_ind.to_csv('./data/Individual_Results.csv', index=False, encoding="ANSI")


def write_to_table_tasks(score, step):
    sleep(5)
    df_pt = pd.read_excel('./data/Played_tasks.xlsx', sheet_name='Played_tasks',
                          dtype={'Команда': str, 'Школа': str, 'Отказы': str,
                                 'Бой 1 Д': int, 'Бой 1 О': int,
                                 'Бой 2 Д': int, 'Бой 2 О': int,
                                 'Бой 3 Д': int, 'Бой 3 О': int,
                                 'Финал Д': int, 'Финал О': int})
    for act in score:
        if act[-1] == 'Д':
            df_pt.loc[df_pt['Команда'] == score[act][0], f'Бой {step} Д'] = score[act][1]
            if score[act][2] != '':
                if score[act][2].startswith(', '):
                    score[act][2] = score[act][2][2:]
                    df_pt.loc[df_pt['Команда'] == score[act][0], f'Отказы'] = score[act][2] + ';'
                else:
                    df_pt.loc[df_pt['Команда'] == score[act][0], f'Отказы'] = score[act][2] + ';'
        elif act[-1] == 'О':
            df_pt.loc[df_pt['Команда'] == score[act][0], f'Бой {step} О'] = score[act][1]

    df_pt.sort_values('Команда', ascending=True, inplace=True)
    df_pt.to_csv('./data/Played_tasks.csv', index=False, encoding="ANSI")
