# Copies "PW_V_Ang/@CASENAME_GICHarmScenario.csv" files into Scenarios.xlsx with columns specifically for use in batch in GICHarm. 

from pathlib import Path
import pandas as pd
import os

def write_scenarios(file_ps: list[Path], out_file_p: Path, out_folder_p: Path, extension: str):
    # file_ps: list of CSV file results to take in.
    # extension: indicates the part of the filename to exclude from the sheet name. 
    # Outputs a spreadsheet at out_file_p with a tab for each input CSV. 

    # Delete scenarios file if it already exists.
    if os.path.exists(out_file_p):
        os.remove(out_file_p)

    headers = ['BusNum1','BusNum2','BusNum3','Ckt','EffectiveGIC','GICBus1','GICBus2']
    with pd.ExcelWriter(str(out_file_p), engine='xlsxwriter') as writer:
        for file_p in file_ps:
            df = pd.read_csv(file_p, skiprows=2, names=headers)
            sheet_name = file_p.name.replace(extension, '')
            
            # Prepare scenarios for GUI.
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            # Prepare scenarios for CMD-mode.
            df.to_csv(out_folder_p / (sheet_name + '.csv'), index=False)
    return

if __name__=='__main__':
    cur_dir = Path(__file__).parent

    out_file_p = cur_dir / 'Scenarios.xlsx'
    out_folder_p = cur_dir / 'GICHarmScenarios'

    result_p = cur_dir / 'PW_V_Ang'
    extension = '_GICHarmScenario.csv'
    file_ps = result_p.rglob('*' + extension)

    write_scenarios(file_ps, out_file_p, out_folder_p, extension)

    exit
