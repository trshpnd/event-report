import pandas as pd
import tkinter as tk
import numpy as np
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

col_event_index = 1
col_timestamp_index = 2
col_code_index = 0
col_eventname_index = 1
error_msg = 'ERRO: Código inexistente na base de dados.'

file_dir = filedialog.askopenfilename()
code_dir = 'EventCodes.txt'
df = pd.read_csv(file_dir)
df_codes = pd.read_csv(code_dir)

col_event_label = df.columns[col_event_index]
col_timestamp_label = df.columns[col_timestamp_index]
col_code_label = df_codes.columns[col_code_index]
col_eventname_label = df_codes.columns[col_eventname_index]

event = df[col_event_label].values
reset_array = np.empty([len(event)], '<U6')
reset_array[:] = ' '

for i in range(len(event)):
    event[i] = int(event[i], 16)
    if event[i] > 32768:
        event[i] -= 32768
        reset_array[i] = 'RESET'

timestamp = df[col_timestamp_label].values
codes = df_codes[col_code_label].values
eventnames = df_codes[col_eventname_label].values

for i in range(len(event)):
    code_exists = 0
    for j in range(len(codes)):
        if event[i] == codes[j]:
            event[i] = eventnames[j]
            code_exists = 1
            break
    if code_exists == 0:
        event[i] = error_msg

df_output = pd.DataFrame({  'Timestamp': timestamp,
                            'Event' : event,
                            'Command' : reset_array})
print(df_output)


## Configuração do arquivo .xlsx
writer = pd.ExcelWriter("output.xlsx", engine='xlsxwriter')

df_output.to_excel(writer, sheet_name = 'Sheet1', index = False)

workbook = writer.book
worksheet = writer.sheets['Sheet1']

max_row, max_col = df_output.shape

worksheet.set_column(0, max_col - 1, 14)
worksheet.add_table(0, 0, max_row, max_col - 1, {'columns':[{'header': 'Timestamp'},{'header': 'Event'},{'header': 'Command'}]})

writer.save()
#MessageBox = ctypes.windll.user32.MessageBoxW
#MessageBox(None, 'Processo finalizado.\nArquivos .csv e .xlsx salvos em:\n\n'+os.getcwd(), 'Sucesso', 0)
#print(event)