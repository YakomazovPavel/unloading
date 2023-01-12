from connections import connectiondef
import pandas as pd
import numpy as np


def main():
    io = connectiondef('ИО')

    df_io = pd.DataFrame(io[1:], columns=io[0]).fillna('')[[
        'Тег сигнала', 'Позиция', 'Тип сигнала']].copy()
    df_io = df_io[(df_io['Позиция'] != '')]
    # .sort_values(['Позиция'], ascending=[True])
    df_io.reset_index(drop=True, inplace=True)
    df_io.reset_index(drop=False, inplace=True)

    df_io['Жила1/2'] = 1

    df_io_2 = df_io.copy()
    df_io_2['Жила1/2'] = 2

    df = pd.concat([df_io, df_io_2]) \
        .sort_values(['index', 'Жила1/2'], ascending=[True, True]) \
        .reset_index(drop=True)
    df['Номер кабеля'] = np.where(df['Тип сигнала'] != 'Namur', 'C-' + df['Позиция'].astype('str'), '')
    df['Жила кабеля'] = np.where(df['Тип сигнала'] != 'Namur',
                                 df['Номер кабеля'].astype('str') + '-' + df['Жила1/2'].astype('str'), '')
    df = df[['Тег сигнала', 'Жила1/2', 'Жила кабеля', 'Номер кабеля']].copy()
    df.to_excel("output.xlsx")
    # (df_io['Тип сигнала'] != 'Namur') &

    print('')

    # df_pp = df_pp[df_pp['Параметр'] == 'Расход']
    # df_env = pd.DataFrame(env[1:], columns=env[0]).fillna('')
    # df_pressure = pd.DataFrame(pressure[1:], columns=pressure[0]).fillna('')
    #
    # df = pd.merge(left=df_pp,
    #               right=df_pressure,
    #               how='left',
    #               on='Позиция',
    #               indicator=True,
    #               suffixes=('', '_y')).drop(['_merge'], axis=1).reset_index(drop=True).fillna('')
    # df.drop(df.filter(regex='_y$').columns, axis=1, inplace=True)


if __name__ == '__main__':
    main()
