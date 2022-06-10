import os

import PySimpleGUI as sg
import matplotlib.pyplot as plt
import xlsxwriter

# from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

numbers = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']


def set_scale(scale):
    root = sg.tk.Tk()
    root.tk.call('tk', 'scaling', scale)
    root.destroy()


'''
def draw_table(month, monthpay_real_arr, monthpay_perc_arr, monthpay_arr):
    cell_matr = np.array((monthpay_real_arr, monthpay_perc_arr, monthpay_arr)).T
    cell_text = []
    for row in cell_matr:
        cell_text.append([f'{x}' for x in row])
    column_names = ["Погашение основного долга", "Погашение процентов", "Общая сумма платежа"]
    row_names = [f'{i+1} месяц' for i in range(month)]
    colors = plt.cm.Greens(np.linspace(0, 0.5, len(row_names)))

    plt.figure(linewidth=2, tight_layout={'pad':1})  # figsize=(9,5)
    the_table = plt.table(cellText=cell_text,
                          rowLabels=row_names,
                          rowColours=colors,
                          colLabels=column_names,
                          loc='center')
    the_table.scale(1, 1.2)
    ax = plt.gca()
    ax.get_xaxis().set_visible(False)
    ax.get_yaxis().set_visible(False)
    plt.box(on=None)
    plt.gcf()
    plt.savefig('credit_calc_table.png', dpi=170)
'''


def draw_graph(month, monthpay_arr):
    plt.figure(figsize=(9, 5))
    plt.grid()  # axis = 'y'
    plt.xlabel('Месяц', fontsize=11)
    plt.ylabel('Сумма выплаты (в руб.)', fontsize=11)
    plt.plot(monthpay_arr, color='green')
    plt.scatter([i for i in range(month)], monthpay_arr, color='green')
    for i in range(month):
        plt.annotate(f'{monthpay_arr[i]}', xy=(i, monthpay_arr[i]), size=11)
    plt.savefig('credit_calc_graph')


def calc_differnt_payment(s, month, rate):
    arr_mpays = []
    s_rest = s
    mpay_no_perc = s / month
    while month != 0:
        arr_mpays.append(round(mpay_no_perc + (s_rest * rate / 1200), 2))
        s_rest -= mpay_no_perc
        month -= 1
    return arr_mpays, round(sum(arr_mpays) - s, 2)


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
         sg.Button('Показать график', button_color='PaleGreen4',
                   font=font_input, key='-show_btn-', visible=False)],
        [sg.Text('Все данные сохраняются в файл credit_calc_option.xlsx в директории с файлом программы', font='Arial 8')]
    ]
    window = sg.Window('Simple credit calculator ©KVA', layout)

    fig = plt.figure()
    set_scale(fig.dpi / 60)  # Set DPI of PySimpleGUI/tkinter to be same as it of Matplotlib/Qt5

    arr_rows_xml = []
    while True:
        event, values = window.read()

        if event in (None, 'Exit'):
            with xlsxwriter.Workbook('credit_calc_option.xlsx') as workbook:
                sheet = workbook.add_worksheet()
                for row_num, data in enumerate(arr_rows_xml):
                    sheet.write_row(row_num, 0, data)

            if os.path.isfile('credit_calc_graph.png'):
                os.remove('credit_calc_graph.png')
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

            if values['-payd-']:
                window['-show_btn-'].update(visible=True)
                monthpay_arr, total = calc_differnt_payment(float(sum_), int(time_), float(rate_))
                window['-total-'].update(f'Общая переплата {total} руб.', background_color='gray85') # DarkSeaGreen1
                window['-mpay-'].update(f'Помесячный график платежей: ', background_color='gray85')

                draw_graph(int(time_), monthpay_arr)

                list_for_save = ['Дифференцированный', float(sum_), int(time_), f'{float(rate_)}%',
                                 '', total, 'Помесячные платежи:']
                list_for_save.extend(monthpay_arr)
                arr_rows_xml.append(list_for_save)

                # figure_canvas_agg = FigureCanvasTkAgg(fig, window['-graph-'].TKCanvas)
                # figure_canvas_agg.draw()
                # figure_canvas_agg.get_tk_widget().pack(side='top', fill='both', expand=1)

        if event == '-show_btn-':
            sg.PopupNoButtons(title='Помесячный график платежей', image='credit_calc_graph.png')
