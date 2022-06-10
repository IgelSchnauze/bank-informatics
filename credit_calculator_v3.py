import PySimpleGUI as sg
import numpy as np
import xlsxwriter

from tabulate import tabulate
from ctypes import windll
windll.shcore.SetProcessDpiAwareness(1)  # убирает размытость!!!

numbers = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']


def fill_xml(arr_rows_xml):
    with xlsxwriter.Workbook('credit_calc_option.xlsx') as workbook:
        sheet = workbook.add_worksheet()
        for row_num, data in enumerate(arr_rows_xml):
            sheet.write_row(row_num, 0, data)


def calc_differnt_payment(s, month, rate):
    s_rest = s
    mpay_no_perc = s / month
    arr_mpays_real = []
    arr_mpays_perc = []
    while month != 0:
        arr_mpays_real.append(mpay_no_perc)
        arr_mpays_perc.append(s_rest * rate / 1200)
        s_rest -= mpay_no_perc
        month -= 1

    arr_mpays_perc = np.around(np.array(arr_mpays_perc), decimals=2)
    arr_mpays_real = np.around(np.array(arr_mpays_real), decimals=2)
    arr_mpays = np.around(arr_mpays_real + arr_mpays_perc, decimals=2)
    return arr_mpays_real, arr_mpays_perc, arr_mpays, round(sum(arr_mpays) - s, 2)


def calc_annuit_payment(s, month, rate):
    month_rate = rate / 1200
    ak = (month_rate * (1 + month_rate) ** month) \
         / (((1 + month_rate) ** month) - 1)
    mpay = s * ak
    return round(mpay, 2), round((mpay * month) - s, 2)


if __name__ == '__main__':
    sg.theme('Reddit')
    font_window = 'Arial 12 bold'
    font_input = 'Arial 12'
    layout = [
        [sg.Text('Вид платежа', font=font_window),
         sg.Radio('Дифференцированный', "Pay", font=font_window, default=True, key='-payd-'),
         sg.Radio('Аннуитетный', "Pay", font=font_window, key='-paya-')],
        [sg.Text('Сумма (в руб.)', font=font_window),
         sg.InputText(size=(13, 2), font=font_input, key='-sum-')],
        [sg.Text('Срок (в мес.)', font=font_window),
         sg.InputText(size=(7, 2), font=font_input, key='-time-'), ],
        [sg.Text('Процентная ставка ', font=font_window),
         sg.InputText(size=(7, 2), font=font_input, key='-rate-'),
         sg.Text('% годовых', font=font_window)],
        [sg.Button('Рассчитать', button_color='PaleGreen4', font=font_window)],
        [sg.Text('\n\n')],
        [sg.Text('Общая переплата ... руб.', font=font_input, key='-total-')],
        [sg.Text('Ежемесячная выплата ... руб.', font=font_input, key='-mpay-'),
         sg.Button('Показать таблицу', button_color='PaleGreen4',
                   font=font_input, key='-show_btn-', visible=False)],
        [sg.Text('Все данные сохраняются в файл credit_calc_option.xlsx в директории с файлом программы', font='Arial 8')]
    ]
    window = sg.Window('Simple credit calculator ©KVA', layout)

    arr_rows_xml = []
    diff_for_popup = None
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

            if values['-paya-']:
                window['-show_btn-'].update(visible=False)
                monthpay, total = calc_annuit_payment(float(sum_), int(time_), float(rate_))
                window['-total-'].update(f'Общая переплата {total} руб.', background_color='gray85')
                window['-mpay-'].update(f'Ежемесячная выплата {monthpay} руб.', background_color='gray85')

                arr_rows_xml.append(['Аннуитетный', float(sum_), int(time_), f'{float(rate_)}%',
                                     '', total, 'Ежемесячный платеж:', monthpay])
                try:
                    fill_xml(arr_rows_xml)
                except xlsxwriter.exceptions.FileCreateError as e:
                    sg.PopupOK('Данные текущих вычислений не были сохранены.\n'
                               'Пожалуйста, закройте файл credit_calc_option.xlsx '
                               'и нажмите на кнопку "Рассчитать" еще раз.', title='Ошибка')
                    arr_rows_xml.pop()

            if values['-payd-']:
                window['-show_btn-'].update(visible=True)
                monthpay_real_arr, monthpay_perc_arr, monthpay_arr, total = \
                    calc_differnt_payment(float(sum_), int(time_), float(rate_))
                window['-total-'].update(f'Общая переплата {total} руб.', background_color='gray85')
                window['-mpay-'].update(f'Помесячный график платежей: ', background_color='gray85')

                diff_for_popup = \
                    [[i , monthpay_real_arr[i], monthpay_perc_arr[i], monthpay_arr[i]] for i in range(int(time_))]

                list_for_save = ['Дифференцированный', float(sum_), int(time_), f'{float(rate_)}%',
                                 '', total, 'Помесячные платежи:']
                list_for_save.extend(monthpay_arr)
                arr_rows_xml.append(list_for_save)

                try:
                    fill_xml(arr_rows_xml)
                except xlsxwriter.exceptions.FileCreateError as e:
                    sg.PopupOK('Данные текущих вычислений не были сохранены.\n'
                               'Пожалуйста, закройте файл credit_calc_option.xlsx '
                               'и нажмите на кнопку "Рассчитать" еще раз.', title='Ошибка')
                    arr_rows_xml.pop()

        if event == '-show_btn-':
            column_names = ["Месяц |", "Погашение основного долга |", "Погашение процентов |", "Общая сумма платежа |"]
            diff_for_popup.insert(0, column_names)
            table_str = tabulate(diff_for_popup, headers="firstrow",
                                 numalign="right", floatfmt=".2f")
            table_str = table_str.replace('-', '')
            sg.PopupScrolled(table_str, title='Помесячный график платежей',
                             background_color='white', button_color='PaleGreen4',
                             font=font_input, size=(80,30), non_blocking=True, modal=False)
            diff_for_popup.pop(0)
            # sg.Print(table_str, background_color='white', no_button=True, size=(90,30))
