import pandas as pd
from datetime import datetime
import easygui
from easygui import *
import os





title = ""
ok_btn_txt = "Continue"
def load_invoice():
    message = "Pleas Load the Invoice CSV"
    msgbox(message, title, ok_btn_txt)
    inv = pd.DataFrame(pd.read_csv(easygui.fileopenbox()))
    return inv
def load_Sidur():
    message = "Pleas Load the Sidur CSV"
    msgbox(message, title, ok_btn_txt)
    sidur = pd.DataFrame(pd.read_csv(easygui.fileopenbox()))
    return sidur

def save_result(res):
    message = "Please Save the Results"
    msgbox(message, title, ok_btn_txt)
    path = easygui.diropenbox() # easygui is used to get a path
    name = str("Harigot."+str(datetime.now().strftime("%d.%m.%Y.%H.%M.%S"))+".xlsx") # the name of the file
    path = os.path.join(path, name)  # this is the full pathname
    res.to_excel(path, 'Harigot') # exporting the dataframe to excel


def start_dev_calc(inv_T, sidur_T):
    if inv_T >= sidur_T:
        return 0
    else:
        dlt = sidur_T - inv_T
        if (dlt.seconds / 60) <= 15:
            return 0
        else:
            return (dlt.seconds / 60 - 15)


def end_dev_calc(inv_T, sidur_T):
    if inv_T >= sidur_T:
        return (inv_T - sidur_T).seconds / 60
    else:
        return 0


def dev_calc(inv_start, inv_end, sidur_start, sidur_end):
    return (start_dev_calc(inv_start, sidur_start) + end_dev_calc(inv_end, sidur_end))

def check(inv,sidur):
    Date = 'תאריך'
    Employee = 'עובד'
    Invstar='התחלה - חשבונית'
    Invend="סיום - חשבונית"
    SidStart= 'התחלה - סידור'
    Sidend='סיום - סידור'
    Hariga = 'חריגה - דקות'
    Notes = 'הערות'
    diviations = pd.DataFrame(columns=[Date,Employee,Invstar,Invend,SidStart,Sidend,Hariga,Notes])

    for index, row_inv in inv.iterrows():
        count = 0
        res = 0
        dt = str(row_inv['date'])
        for index, row_sid in sidur.iterrows():
            if (row_inv['date']) == (row_sid['date']):
                if str(row_inv['employee']) == str(row_sid['employee']):
                    count = count + 1
                    dt = str(row_inv['date'])
                    inv_start = datetime.strptime(dt, '%d/%m/%Y')
                    inv_start = inv_start.replace(hour=datetime.strptime(row_inv['start'], '%H:%M').hour,
                                                  minute=(datetime.strptime(row_inv['start'],
                                                                            '%H:%M').minute))  # adding hour& minuts
                    inv_end = datetime.strptime(dt, '%d/%m/%Y')
                    inv_end = inv_end.replace(hour=datetime.strptime(row_inv['end'], '%H:%M').hour,
                                              minute=(datetime.strptime(row_inv['end'],
                                                                        '%H:%M').minute))  # adding hour& minuts
                    sid_start = datetime.strptime(dt, '%d/%m/%Y')
                    sid_start = sid_start.replace(hour=datetime.strptime(row_sid['start'], '%H:%M').hour,
                                                  minute=(datetime.strptime(row_sid['start'],
                                                                            '%H:%M').minute))  # adding hour& minuts

                    sid_end = datetime.strptime(dt, '%d/%m/%Y')
                    sid_end = sid_end.replace(hour=datetime.strptime(row_sid['end'], '%H:%M').hour,
                                              minute=(
                                                  datetime.strptime(row_sid['end'], '%H:%M').minute))  # adding hour& minuts

                    res = dev_calc(inv_start, inv_end, sid_start, sid_end)
        if count == 2:
            diviations = diviations.append({Date: str(dt),
                                            Employee: row_inv['employee'],
                                            Invstar: str(inv_start.strftime("%H:%M")),
                                            Invend: str(inv_end.strftime("%H:%M")),
                                            SidStart: str(sid_start.strftime("%H:%M")),
                                            Sidend: str(sid_end.strftime("%H:%M")),
                                            Hariga: str(int(res)),
                                            Notes: "משמרת כפולה"},ignore_index=True)
            print(str(dt) + "." + str(row_inv['employee']) + "." + str(inv_start.strftime("%H:%M")) + "." + str(inv_end.strftime("%H:%M")) +
                  "." + str(sid_start.strftime("%H:%M")) + "." + str(sid_end.strftime("%H:%M")) + "." + str(int(res))+" MULTIPLE-SHIFTS")
            continue
        if count == 0:
            print(str(dt) + "."+str(row_inv['employee']) + "." + str(row_inv['start']) + "." +
                  str(row_inv['end'])+" None"+" None"+" NOT-IN-SIDUR")
            diviations = diviations.append({Date: str(dt),
                                            Employee: row_inv['employee'],
                                            Invstar: str(row_inv['start']),
                                            Invend: str(row_inv['end']),
                                            SidStart: 'חסר',
                                            Sidend: 'חסר',
                                            Hariga: "",
                                            Notes: "לא בסידור"},ignore_index=True)
            continue
        if res == 0:
            continue
        else:
            print(str(dt) + "." + str(row_inv['employee']) + "." + str(inv_start.strftime("%H:%M")) + "." + str(inv_end.strftime("%H:%M")) +
                  "." + str(sid_start.strftime("%H:%M")) + "." + str(sid_end.strftime("%H:%M")) + "." + str(int(res)))
            diviations = diviations.append({Date: str(dt),
                                            Employee: row_inv['employee'],
                                            Invstar: str(inv_start.strftime("%H:%M")),
                                            Invend: str(inv_end.strftime("%H:%M")),
                                            SidStart: str(sid_start.strftime("%H:%M")),
                                            Sidend: str(sid_end.strftime("%H:%M")),
                                            Hariga: str(int(res))},ignore_index=True)
    save_result(diviations)

if __name__ == '__main__':
    inv=load_invoice()
    sidur=load_Sidur()
    check(inv,sidur)