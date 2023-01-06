import pandas as pd
import PySimpleGUI as sg

import PySimpleGUI as sg

sg.theme('DarkAmber')   # Add a touch of color
# All the stuff inside your window.
EXCEL_FILE = 'BMR_data.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Name', size=(15,1)), sg.InputText(key='Name')],
    [sg.Text('Lenght (in cm) ', size=(15,1)), sg.Spin([i for i in range(120,240)],
                                                       initial_value=0, key='Lenght')],
    [sg.Text('Weight (in Kg) ', size=(15,1)), sg.Spin([i for i in range(35,300)],
                                                       initial_value=0, key='Weight')],
    [sg.Text('Age ', size=(15,1)), sg.Spin([i for i in range(18,100)],
                                                       initial_value=0, key='Age')],
    [sg.Text('Sex', size=(15,1)), sg.Combo(['Male', 'Female'], key='Sex')],
    [sg.Text('Body Shape', size=(15,1)), sg.Combo(['Normal', 'Muscle', 'Round'], key='Shape')],
    [sg.Text('Physical activity level ', size=(15,1)), sg.Combo(['Sedentary', 'Moderate', 'Active', 'Professional'], key='Pal')],

    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

def clear_input():
    for key in values:
        window[key]('')
    return None
# Create the Window
window = sg.Window('BMI and DEE calculator', layout)
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit': # if user closes window or clicks Exit
        break
    if event == 'Clear':
        clear_input()

    if event == 'Submit':
        new_record = pd.DataFrame(values, index=[0])
        new_record['BMI'] = new_record['Weight']/((new_record['Lenght']/100)*(new_record['Lenght']/100))
        new_record.loc[new_record['Sex'] == "Male", 'BMR' ] = [66.5 + (new_record['Weight'] * 13.75) + (new_record['Lenght'] * 5.003) - (
                            6.755 * new_record['Age'])]
        new_record.loc[new_record['Sex'] == "Female", 'BMR'] = [655.1 + (new_record['Weight'] * 9.563) + (new_record['Lenght'] * 1.850) - (
                           4.676 * new_record['Age']) for x in new_record['Sex']]
        new_record.loc[new_record['Shape'] == "Muscle", 'BMR'] = int(new_record['BMR'] * 1.15)
        new_record.loc[new_record['Shape'] == "Round", 'BMR'] = int(new_record['BMR'] * 0.85)
        new_record.loc[new_record['Shape'] == "Normal", 'BMR'] = int(new_record['BMR'])
        new_record.loc[new_record['Pal'] == "Sedentary", 'DEE'] = int(new_record['BMR'] * 1.55)
        new_record.loc[new_record['Pal'] == "Moderate", 'DEE'] = int(new_record['BMR'] * 1.85)
        new_record.loc[new_record['Pal'] == "Active", 'DEE'] = int(new_record['BMR'] * 2.10)
        new_record.loc[new_record['Pal'] == "Professional", 'DEE'] = int(new_record['BMR'] * 2.50)

        df = pd.concat([df, new_record], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        df = pd.read_excel(EXCEL_FILE)
        last_value = df['DEE'].iat[-1]
        last_value1 = df['BMR'].iat[-1]
        sg.popup('Basal Metabolic Rate ' + str(last_value1) + "  Kcal \n"
            'Daily Energy Expenditure ' + str(last_value) + "  Kcal")

window.close()