from pathlib import Path

import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme('Reddit')

current_dir = Path(__file__).parent if '__file__' in locals() else Path.cwd()
EXCEL_FILE = "Data_Entry.xlsx"
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Name', size=(15,1)), sg.InputText(key='Name')],
    [sg.Text('MiddleName', size=(15,1)), sg.InputText(key='MiddleName')],
    [sg.Text('LastName', size=(15,1)), sg.InputText(key='LastName')],
    [[sg.Text('Date-start', size=(10,1)), sg.Spin([i for i in range(0,32)], initial_value=0, key='date'), sg.Text('Month', size=(5,1)), sg.Spin([i for i in range(0,13)], initial_value=0, key='month'), sg.Text('Year', size=(5,1)), sg.Spin([i for i in range(2022,2031)], initial_value=0, key='year')]],
    [[sg.Text('Date-end', size=(10, 1)), sg.Spin([i for i in range(0, 32)], initial_value=0, key='date'),sg.Text('Month', size=(5, 1)), sg.Spin([i for i in range(0, 13)], initial_value=0, key='month'),sg.Text('Year', size=(5, 1)), sg.Spin([i for i in range(2022, 2031)], initial_value=0, key='yearend')]],
    [sg.Text('Branch area and pincode', size=(15,1)), sg.InputText(key='City')],
    [sg.Text('mobile no.', size=(15,1)), sg.InputText(key='mobile')],
    #[sg.Text('Favorite Colour', size=(15,1)), sg.Combo(['Green', 'Blue', 'Red'], key='Favorite Colour')],
    [sg.Text('gov. leave taken', size=(15,1)),sg.Checkbox('yes', key='govyes'),sg.Checkbox('no', key='govno'), sg.Text('if "yes"', size=(5,1)), sg.Spin([i for i in range(0,16)], initial_value=0, key='govnos')],
    [sg.Text('restricted leaves taken', size=(17, 1)), sg.Checkbox('yes', key='resyes'), sg.Checkbox('no', key='resno'),sg.Text('if "yes"', size=(5, 1)), sg.Spin([i for i in range(0, 16)], initial_value=0, key='restrictno')],
    [sg.Text('Reason for leave:', size=(15,1)), sg.InputText(key='reason')],
    [sg.Text('Address', size=(15,1)), sg.InputText(key='address')],
    [sg.Text('permission to leave headquarters', size=(17, 1)), sg.Checkbox('yes', key='hyes'), sg.Checkbox('no', key='hno')],



    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Simple data entry form', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        new_record = pd.DataFrame(values, index=[0])
        df = pd.concat([df, new_record], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()
window.close()