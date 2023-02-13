from data_engine import getTemptureForOL, getPressureForOL, getFlowForOL, getLevelForOL
from docClass import unload


def main():

    print('Выгрузка начата!')
    unload('Температура', *getTemptureForOL())
    print('Температура - ГОТОВО')
    unload('Давление', *getPressureForOL())
    print('Давление - ГОТОВО')
    unload('Расход', *getFlowForOL())
    print('Расход - ГОТОВО')
    unload('Уровень', *getLevelForOL())
    print('Уровень - ГОТОВО')
    print('Выгрузка завершена!')


if __name__ == '__main__':
    main()