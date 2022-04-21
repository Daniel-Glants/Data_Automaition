#   region  -------------------------------   Imports   -----------------------------------
import win32com.client as client
import os, os.path
import glob
import time
from datetime import datetime
#   endregion

# region -------------------------------  Variables & Objects   --------------------------------

Contacts_Test = {

    'Paybis': [
        'Arturs Markevics <a.mark@paybis.com>; Elvisa Brahja <elvisa.brahja@paybis.com>; AB_finance <AB_finance@simplex.com>'],

    'Debex & Elastum': [
        'Gabriele Lipniute <gabriele@elastum.io>; Ina Ruskiene <ina@elastum.io>; Goda Bumbliauskaite <goda@elastum.io>; AB_finance <AB_finance@simplex.com>'],

    'Hodleris': [
        'Aiste  Hfinance <aiste@hfinance.co>; vytas@hodlfinance.io; gintautas@hodlfinance.io; Matas Ramanauskas <matas@hfinance.co>; AB_finance <AB_finance@simplex.com>'],

    'eToroX': [
        'adiga@etoro.com; Shay Meir <shayme@etoro.com>; Ravit Lotan <ravitlo@etoro.com>; Emilios Georgiou <emiliosge@etoro.com>; AB_finance <AB_finance@simplex.com>']
}
shared_dir = r'C:\Users\DanielGlants\OneDrive - Nuvei Technologies\Finance - Simplex\Reconciliations\Weekly Python Script\all 2019 files'  # shared directory
local_dir = r'C:\Users\DanielGlants\OneDrive - Nuvei Technologies\Daily - ECP'  # local temp directory
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
        if item.Subject == 'ECP Daily Payment':  # Check if Email's Subject is relevant.
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
    if lp_name == 'eToroX':
        return ''.join(Contacts_Test.get('eToroX'))
    if lp_name == 'Paybis':
        return ''.join(Contacts_Test.get('Paybis'))
    if lp_name == 'Hodleris':
        return ''.join(Contacts_Test.get('Hodleris'))
    if (lp_name == 'UAB Debex') or (lp_name == 'Elastum'):
        return ''.join(Contacts_Test.get('Debex & Elastum'))


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
    <p>Hello,<br /><br /><br />Please see attached ECP&nbsp;daily&nbsp;report for&nbsp;<strong><u>{0}</u>.</strong></p>
    <p>&nbsp;</p>
    <p>Sincerely,</p>
    <p>--</p>
    <p><span style="color: #41ac48;">Daniel Glants</span></p>
    <p>Financial Data Analyst</p>
    <p>Office: <span style="text-decoration: underline;">+972-3-7510192</span></p>
    <p>Email:<a href="mailto:daniel.glants@simplex.com">daniel.glants@simplex.com</a></p>
    <p><img src="https://pbs.twimg.com/profile_images/1437419441864814595/85g3O4A-_400x400.jpg" alt="" width="220" height="223" /></p>
    <p><a href="http://www.simplex.com">www.simplex.com</a></p><p><a href="https://www.facebook.com/Simplex.cc"><img src="https://www.nicepng.com/png/full/448-4482584_fb-icon-facebook-icon.png" alt="" width="40" height="40" /></a>&nbsp; &nbsp;<a href="https://twitter.com/SimplexCC"><img src="https://www.pngitem.com/pimgs/m/33-338808_logo-twitter-circle-png-transparent-image-transparent-twitter.png" alt="" width="40" height="40" /></a>&nbsp; &nbsp;<a href="https://www.linkedin.com/company/simplexcc?challengeId=AQH9cgE6Q9F4mwAAAX_awfnG30Y1vPp3R9Y-MgALZhbfEJzYucGdri8o48dSqHHaPbtnzRv94TKKp1c5-XhNQNIwH1l0GVk-KQ&amp;submissionId=da64f238-ba27-e116-317c-f52697f8d164"><img src="data:image/png;base64,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" alt="" width="40" height="40" /></a>&nbsp;&nbsp;<a href="https://www.instagram.com/simplex_life/"><img src="https://thumbs.dreamstime.com/b/web-180643611.jpg" alt="" width="40" height="40" /></a>&nbsp;&nbsp;<a href="https://www.glassdoor.com/Overview/Working-at-Simplex-Israel-EI_IE1366394.11,25.htm"><img src="https://simg.nicepng.com/png/small/128-1281856_glassdoor-social-icon-jimbere-fund.png" alt="" width="40" height="40" /></a></p>
    <p></p>
    '''.format(date)  # attach body text with signature
    return html_body


'''   ---------------------------   Generate Emails Subject   -----------------------------    '''


def gen_subject(file_name):
    date = get_date(file_name)
    text = 'ECP Daily Report for {0}'.format(date)
    return text


'''  -----------------------------   Send Relevant Worksheet to LP   ---------------------------   '''


def send_files(files_dir, outlook_instance):
    os.chdir(files_dir)  # navigate to the split directory folder
    batch_list = glob.glob('*.xlsx')  # list all mails for sending
    for item in batch_list:
        if item.split()[-1] not in ['Elastum.xlsx', 'Debex.xlsx']:
            mail = new_msg(outlook_instance)
            mail.Attachments.Add(os.path.join(files_dir, item))
            mail.Subject = gen_subject(item)
            mail.HTMLbody = gen_body(item)
            mail.To = identify(item)
    mail = new_msg(outlook_instance)
    for item in batch_list:
        if item.split()[-1] in ['Elastum.xlsx', 'Debex.xlsx']:
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




