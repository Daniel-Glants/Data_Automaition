#   region  -------------------------------   Imports   -----------------------------------
import win32com.client as client
import os, os.path
import glob
import time
from datetime import datetime
#   endregion

# region -------------------------------  Variables & Objects   --------------------------------

Contacts_Test = {

    'LP1': [
        Mail1; Mail2;Mail3],

    'LP2':[
        Mail1; Mail2;Mail3],

    'LP3': [
        Mail1; Mail2;Mail3],

    'LP24': [
        Mail1; Mail2;Mail3]
}
shared_dir = r'---'  # shared directory
local_dir = r'---'  # local temp directory
CHECK_FOLDER = os.path.isdir(local_dir)
if not CHECK_FOLDER:
    os.makedirs(local_dir)
    print("created folder : ", local_dir)
os.chdir(shared_dir)  # navigate to the shared directory folder
FileList = glob.glob('*.xlsx')  # Iterate through files in the folder & create a list of .xlsx
n = 0  # counter
new_emails = {}  # new emails dictionary
# endregion

#  region  ------------------------------   Functions   ----------------------------------
'''   ---------------------------   Initialize Outlook App Instance   -----------------------------    '''


def init_outlook(action):
    if action == 'search':
        outlook = client.Dispatch("Outlook.Application").GetnameSpace('MAPI')  # Create an OutLook Instance
        return outlook
    elif action == 'create':
        outlook = client.Dispatch("Outlook.Application")
        return outlook


'''   ---------------------------   Initialize Excel App Instance   -----------------------------    '''


def init_excel():
    xl_inst = client.Dispatch("Excel.Application")  # Create an OutLook Instance
    xl_inst.Visible = False
    return xl_inst


'''    --------------------------   Search For New Daily Mails   ------------------------------     '''


def search_for_email(folder):
    global n
    global new_emails
    if folder.Folders.Count > 0:
        for sub_f in folder.Folders:  # Iterate through emails recursively.
            search_for_email(sub_f)
    for item in folder.Items:
        if item.Subject == 'Subject':  # Check if Email's Subject is relevant.
            attachments = item.Attachments  # get list of attachments.
            for file in attachments:  # Iterate through mail's attachments.
                if file.FileName[-4:] == 'xlsx':
                    if file.FileName not in FileList:  # Check if the newly received file was already downloaded.
                        n += 1
                        key = f'{n}'
                        new_emails[str(key)] = str(file.FileName)
                        download_file(file)  # Download File.


'''     ---------------------------------   Download Emails   -------------------------------------     '''


def download_file(attachment):
    attachment.SaveAsFile(os.path.join(shared_dir, str(attachment.FileName)))  # save to shared directory
    attachment.SaveAsFile(os.path.join(local_dir, str(attachment.FileName)))  # save to local temp directory


'''     -------------------   Split files to separate worksheets   ---------------------------------      '''


def split_workbook(file_path, excel_instance):
    temp_path = file_path[:-5]
    check_path = os.path.isdir(temp_path)  # Check if there is a temp  folder for the file
    if not check_path:  # If there isn't then create one
        os.makedirs(temp_path)
        print("created folder : ", temp_path)
    wb = excel_instance.Workbooks.open(file_path)
    for ws in wb.worksheets:
        if ws.name not in ['New', 'Rates']:
            ws.copy
            time.sleep(30)
            excel_instance.ActiveWorkbook.SaveAs(os.path.join(temp_path, '{0} - {1}.xlsx'.format(wb.name[:-5], ws.name)),
                                                 FileFormat = 51)
            time.sleep(30)
            excel_instance.ActiveWorkbook.Close(True)
            time.sleep(30)
    excel_instance.ActiveWorkbook.Close(False)
    excel_instance.Quit()
    return temp_path


'''   ---------------------------   Create New Mail   -----------------------------    '''


def new_msg(outlook_instance):
    message = outlook_instance.CreateItem(0)  # Create message Object
    message.Display()  # Display Message
    return message  # Return Message Object


'''   ---------------------------   Identify Relevant LP   -----------------------------    '''


def get_recip(lp_name):
    if lp_name == 'LP1':
        return ''.join(Contacts_Test.get('LP1'))
    if lp_name == 'LP2':
        return ''.join(Contacts_Test.get('LP2'))
    if lp_name == 'LP3':
        return ''.join(Contacts_Test.get('LP3'))
    if (lp_name == 'LP4') or (lp_name == 'LP4'):
        return ''.join(Contacts_Test.get('LP4 & LP4'))


def identify(file_name):
    lp = file_name[file_name.rindex('-')+2:-5]  # get LP NAme
    recipients = get_recip(lp)
    return recipients


'''   ---------------------------   Get Date From File Path   -----------------------------    '''


def get_date(file_name):
    # x = file_name.rindex("Simplex.xlsx")  # index of 'Simplex,xlsx'
    dt_raw = file_name[:8]  # snip out the date text
    date = datetime.strptime(dt_raw, '%d%m%Y').strftime('%d/%m/%Y')  # convert to date format
    return date


'''   ---------------------------   Generate Emails body   -----------------------------    '''


def gen_body(file_name):
    date = get_date(file_name)  # get date from path
    html_body = ''' 
   body
    '''.format(date)  # attach body text with signature
    return html_body


'''   ---------------------------   Generate Emails Subject   -----------------------------    '''


def gen_subject(file_name):
    date = get_date(file_name)
    text = 'Daily Report for {0}'.format(date)
    return text


'''  -----------------------------   Send Relevant Worksheet to LP   ---------------------------   '''


def send_files(files_dir, outlook_instance):
    os.chdir(files_dir)  # navigate to the split directory folder
    batch_list = glob.glob('*.xlsx')  # list all mails for sending
    for item in batch_list:
        if item.split()[-1] not in ['LP1.xlsx', 'LP2.xlsx']:
            mail = new_msg(outlook_instance)
            mail.Attachments.Add(os.path.join(files_dir, item))
            mail.Subject = gen_subject(item)
            mail.HTMLbody = gen_body(item)
            mail.To = identify(item)
    mail = new_msg(outlook_instance)
    for item in batch_list:
        if item.split()[-1] in ['LP1.xlsx', 'LP2.xlsx']:
            mail.Attachments.Add(os.path.join(files_dir, item))
    mail.Subject = gen_subject(batch_list[0])
    mail.HTMLbody = gen_body(batch_list[0])
    mail.To = identify(batch_list[0])


'''   ---------------------------   Process Attachment   -----------------------------    '''


def process_file(file_path):
    ol_inst = init_outlook('create')
    xl_inst = init_excel()
    xl_inst.Visible = False
    split_dir = split_workbook(file_path, xl_inst)
    send_files(split_dir, ol_inst)


# TODO   ---------------------------------   Clear Local directory   ---------------------------    '''

# def clear_dir():


# endregion


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    simplex_email = init_outlook('search')  # Initiate Outlook App Instance
    for Folder in simplex_email.Folders.Item(1).Folders:  # scan each folder for Daily ECP mails.
        search_for_email(Folder)  # Download new Emails

    print(f'{len(new_emails)} New Emails Where Found :')
    for key in new_emails:  # Print out new Emails detected
        print(new_emails[str(key)])
    print('Phase 1 Completed, new daily emails collected and ready for inspection')

    new_emails_list = glob.glob(os.path.join(local_dir, '*.xlsx'))  # list all new downloaded emails
    for file in new_emails_list:
        process_file(file)
    print('Files sent successfully')




