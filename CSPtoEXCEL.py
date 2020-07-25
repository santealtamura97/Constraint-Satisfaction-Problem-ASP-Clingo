import subprocess
import pandas as pd


def result_to_Excel(list_solution, file_name):
    writer = pd.ExcelWriter("{}.xlsx".format(file_name), engine="xlsxwriter")
    df = pd.DataFrame(index=range(len(list_solution)), columns=['Settimana', 'Giorno', 'Inizio', 'Fine', 'Docente', 'Corso'])
    for ind, s in enumerate(list_solution):
        row = s[s.index('(') + 1:-1].split(',')
        df.loc[ind] = row
    for settimana in df['Settimana'].unique():
        df.Giorno = df.Giorno.astype("category")
        df.Giorno.cat.set_categories(['lunedi', 'martedi', 'mercoledi', 'giovedi', 'venerdi', 'sabato'], inplace=True)
        df.Inizio = df.Inizio.astype("category")
        df.Inizio.cat.set_categories(['9', '10', '11', '12', '14', '15', '16', '17'], inplace=True)
        df[df['Settimana'] == settimana].sort_values(by=['Giorno', 'Inizio']).to_excel(writer, sheet_name=settimana, index=False)
    writer.save()


num_solution = input("Numero Soluzioni: ")
asp_calendar_path = "./Orario-Lezioni.cl"
runthis = "/Applications/clingo-5.4.0-macos-x86_64/clingo --verbose=0 --warn=no-global-variable {} {}".format(asp_calendar_path, num_solution)
osstdout = subprocess.Popen(runthis, shell=True, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, close_fds=True)
solution = osstdout.communicate()[0].strip()
return_code = osstdout.returncode
solution = solution.decode("utf-8").replace("\nSATISFIABLE", "")

all_solutions = solution.split("\n")
print(all_solutions)
for i, s in enumerate(all_solutions):
    list_solution = s.split(" ")
    result_to_Excel(list_solution, "Orario (Soluzione: {})".format(i+1))




