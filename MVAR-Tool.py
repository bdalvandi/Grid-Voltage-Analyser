#MVAR tool developped by Behdad Dalvandi, IGMC
import os
import glob
import pandas as pd
import numpy as np
from tkinter import filedialog, Tk

udb = pd.read_excel("./UNITS_DATABASE.xls")
udb_units = set(udb['Unit'])
root = os.path.dirname(__file__)
xls_filenames = glob.glob(os.path.join(root, "XLS_RESULTS/*.xls"))

print("Loading MVAR data from state estimation report files ...")
_dfs = []
for xls_file in xls_filenames:
    print('reading ', xls_file)
    _d = pd.read_excel(xls_file, header=None, usecols=[0,4], skiprows=[0,1], names=['Unit', 'Mvar'], dtype={'Mvar':np.float64})
    _d['isu'] = _d.Unit.str.match('^\w+\s+[GNSH]\d\d?')
    _d = _d.query('isu == True')
    _d = _d.drop('isu',1)
    _d['File'] = xls_file
    _dfs.append(_d)
_df = pd.concat(_dfs)
_df = _df.merge(udb,on='Unit',how='left')
_df_lead = _df.query('Mvar<-1')
_df_lag = _df.query('Mvar>1')
_pv_lead = _df_lead.pivot_table('Mvar', ['Code','Group'], 'File', margins=True, margins_name='Mean MVAR')
_pv_lag = _df_lag.pivot_table('Mvar', ['Code','Group'], 'File', margins=True, margins_name='Mean MVAR')

win = Tk()
final_file = filedialog.asksaveasfilename(filetypes=(('Excel files','*.xlsx'),('All files','*.*')), defaultextension=".xlsx")
win.destroy()

print('Generating result file ...')
xlw = pd.ExcelWriter(final_file)
_pv_lead.to_excel(xlw, 'Lead')
_pv_lag.to_excel(xlw, 'Lag')
xlw.save()

print("Job finished.")
