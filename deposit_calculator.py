import PySimpleGUI as sg
import numpy as np
import xlsxwriter

from tabulate import tabulate
from ctypes import windll
windll.shcore.SetProcessDpiAwareness(1)  # убирает размытость!!!

numbers = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']


def fill_xml(arr_rows_xml):
    with xlsxwriter.Workbook('deposit_calc_option.xlsx') as workbook:
        sheet = workbook.add_worksheet()
        for row_num, data in enumerate(arr_rows_xml):
            sheet.write_row(row_num, 0, data)


def calc_deposit(s, month, rate):
    mprofit_perc = s * (rate / 1200)  # ((rate / 100) / 365) * 30

    # month_sums = [s+mprofit_perc]
    # for i in range(1, month):
    #     month_sums.append(month_sums[-1] + mprofit_perc)
    month_perc_sums = [mprofit_perc]
    for _ in range(1, month):
        month_perc_sums.append(month_perc_sums[-1] + mprofit_perc)

    total_profit_perc = mprofit_perc * month
    total_sum = s + total_profit_perc
    return round(total_sum,2), round(total_profit_perc, 2), round(mprofit_perc, 2), \
           np.round(np.array(month_perc_sums), 2)


def calc_deposit_capit(s, month, rate):
    r = (rate / 1200)
    mprofits = [s * r]
    month_sums = [s * (1+r)]
    s_now = s + mprofits[-1]
    for _ in range(1, month):
        now_profit_perc = s_now * r
        mprofits.append(now_profit_perc)
        month_sums.append(month_sums[-1] + now_profit_perc)
        s_now += now_profit_perc

    total_sum = month_sums[-1]
    return round(total_sum,2), round(sum(mprofits), 2), \
           np.round(np.array(mprofits), 2), np.round(np.array(month_sums), 2)


if __name__ == '__main__':
    sg.theme('Reddit')
    font_window = 'Arial 12 bold'
    font_input = 'Arial 12'
    layout = [
        [sg.Checkbox('Капитализация', font=font_window, default=False, key='-cap-')],
        [sg.Text('Сумма (в руб.)', font=font_window),
         sg.InputText(size=(13, 2), font=font_input, key='-sum-')],
        [sg.Text('Срок (в мес.)', font=font_window),
         sg.InputText(size=(7, 2), font=font_input, key='-time-'), ],
        [sg.Text('Процентная ставка ', font=font_window),
         sg.InputText(size=(7, 2), font=font_input, key='-rate-'),
         sg.Text('% годовых', font=font_window)],
        [sg.Button('Рассчитать', button_color='PaleGreen4', font=font_window)],
        [sg.Text('\n\n')],
        [sg.Text('Общая итоговая сумма ... руб.', font=font_input, key='-total-')],
        [sg.Text('Итоговая сумма процентов ... руб.', font=font_input, key='-totalperc-')],
        [sg.Text('Помесячное состояние счета: ', font=font_input, key='-msum-'),
         sg.Button('Показать таблицу', button_color='PaleGreen4',
                   font=font_input, key='-show_btn-', visible=False)],
        [sg.Text('Все данные сохраняются в файл deposit_calc_option.xlsx в директории с файлом программы', font='Arial 8')]
    ]
    window = sg.Window('Simple deposit calculator ©KVA', layout)

    arr_rows_xml = []
    data_for_table = None
    while True:
        event, values = window.read()

        if event in (None, 'Exit'):
            break

        if event == 'Рассчитать':
            sum_ = values['-sum-'].replace(' ', '')
            time_ = values['-time-'].replace(' ', '')
            rate_ = values['-rate-'].replace(' ', '')

            if sum_ == '' or time_ == '' or rate_ == '':
                sg.PopupOK('Пожалуйста, заполните все поля.', title='Ошибка')
                continue

            if not all([sym in numbers or sym == '.' or sym == ',' for sym in sum_]):
                sg.PopupOK('При заполнении поля "Сумма" можно использовать '
                           'только цифры и десятичные разделители.', title='Ошибка')
                continue
            if not all([sym in numbers for sym in time_]):
                sg.PopupOK('При заполнении поля "Срок" можно использовать '
                           'только цифры.', title='Ошибка')
                continue
            if not all([sym in numbers or sym == '.' or sym == ',' for sym in rate_]):
                sg.PopupOK('При заполнении поля "Процентная ставка" можно использовать '
                           'только цифры и десятичные разделители.', title='Ошибка')
                continue

            sum_ = sum_.replace(',', '.')
            if sum_[0] == '.' or sum_[-1] == '.':
                sg.PopupOK('Поле "Сумма" заполнено некорректно.', title='Ошибка')
                continue
            rate_ = rate_.replace(',', '.')
            if rate_[0] == '.' or rate_[-1] == '.':
                sg.PopupOK('Поле "Процентная ставка" заполнено некорректно.')
                continue

            window['-show_btn-'].update(visible=True)

            if values['-cap-']:
                total, total_perc, month_percs, month_sums = \
                    calc_deposit_capit(float(sum_), int(time_), float(rate_))
                window['-total-'].update(f'Общая итоговая сумма {total} руб.', background_color='gray85')
                window['-totalperc-'].update(f'Итоговая сумма процентов {total_perc} руб.', background_color='gray85')
                window['-msum-'].update(f'Помесячное состояние счета: ', background_color='gray85')

                data_for_table = \
                    [[i+1 , month_percs[i], month_sums[i]] for i in range(int(time_))]
                data_for_table.insert(0, ['Месяц |', 'Сумма начисленных процентов |', 'Текущая сумма на счете |'])

                list_for_save = ['С капитализацией', float(sum_), int(time_), f'{float(rate_)}%',
                                 '', total, 'Помесячное состояние счета:']
                list_for_save.extend(month_sums)
                arr_rows_xml.append(list_for_save)

                try:
                    fill_xml(arr_rows_xml)
                except xlsxwriter.exceptions.FileCreateError as e:
                    sg.PopupOK('Данные текущих вычислений не были сохранены.\n'
                               'Пожалуйста, закройте файл deposit_calc_option.xlsx '
                               'и нажмите на кнопку "Рассчитать" еще раз.', title='Ошибка')
                    arr_rows_xml.pop()
            else:
                total, total_perc, month_perc, month_perc_sums = \
                    calc_deposit(float(sum_), int(time_), float(rate_))
                window['-total-'].update(f'Общая итоговая сумма {total} руб.', background_color='gray85')
                window['-totalperc-'].update(f'Итоговая сумма процентов {total_perc} руб.', background_color='gray85')
                window['-msum-'].update(f'Каждый месяц выплачивается {month_perc} руб. : ', background_color='gray85')

                data_for_table = \
                    [[i+1, float(sum_), month_perc_sums[i]] for i in range(int(time_))]
                data_for_table.insert(0, ['Месяц |', 'Сумма на счете |', 'Полученные проценты |'])

                list_for_save = ['Без капитализации', float(sum_), int(time_), f'{float(rate_)}%',
                                 '', total, 'Помесячно выплачиваемые проценты:', month_perc]
                arr_rows_xml.append(list_for_save)

                try:
                    fill_xml(arr_rows_xml)
                except xlsxwriter.exceptions.FileCreateError as e:
                    sg.PopupOK('Данные текущих вычислений не были сохранены.\n'
                               'Пожалуйста, закройте файл deposit_calc_option.xlsx '
                               'и нажмите на кнопку "Рассчитать" еще раз.', title='Ошибка')
                    arr_rows_xml.pop()

        if event == '-show_btn-':
            table_str = tabulate(data_for_table, headers="firstrow",
                                 numalign="right", floatfmt=".2f")
            table_str = table_str.replace('-', '')
            sg.PopupScrolled(table_str, title='График состояния счета помесячно',
                             background_color='white', button_color='PaleGreen4',
                             font=font_input, size=(70,30), non_blocking=True, modal=False)
