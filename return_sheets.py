# Author: The Sang Nguyen
# Written for the University of Goettingen.
# For other universities, please adapt the script.

'''
This script automates the delivery of corrected homework by sending them by 
e-mail via SMTP.

--- SET UP ---:
Step 1:
    Save this script in a directory together with 'Punkteliste.csv' from StudIP.
    Then create a folder with name 'SheetXX' containing the corrected homework 
    with names of the form
        - 'XX_surname_corrected.pdf' (single submissions), or,
        - 'XX_surname_surname2_corrected.pdf' (group submissions), etc,
    where XX is the homework number.
Step 2 - in method 'main()':
    Adjust login data 'UNI_USER', 'UNI_MAIL' and 'UNI_PASS'.
    (It is recommended to use environment variables instead of clear text.)
Step 3 - in method 'send_mail()':
    Adjust e-mail text.
'''

# built-in modules
import os
import sys
import smtplib
from email.message import EmailMessage

# third-party modules
import pandas as pd


def get_hw_number():
    '''
    Get homework number from user.
    '''
    i = input('++++ Please enter the homework number: ')
    return '0' + i if len(i) == 1 else i


def get_path(hw_number):
    '''
    Get path to the folder containing the corrected homework.
    '''
    path = os.getcwd()
    if sys.platform.startswith('win32'):
        path += r'\Sheet'
    elif sys.platform.startswith('linux'):
        path += '/Sheet'
    # TODO: check if script works on macOS
    elif sys.platform.startswith('darwin'):
        path += '/Sheet'
    else:
        raise OSError('this script is not supported by your OS')
    return path + hw_number


def get_corr_hw(path, head, tail, filetypes):
    '''
    Get list of files of the corrected homework and list of students' names.
    '''
    files = []
    names = []
    for file in os.listdir(path):
        for filetype in filetypes:
            ending = tail + filetype
            if file[:len(head)] == head and file[-len(ending):] == ending:
                files.append(file)
                names.append(file[len(head):-len(ending)])
    return files, names


def send_mail(
    UNI_USER, UNI_MAIL, UNI_PASS,
    receiver_email, firstname, surname, hw_number, filename
):
    '''
    Send e-mail to student via SMTP.
    '''
    msg = EmailMessage()
    msg['From'] = UNI_MAIL
    msg['To'] = receiver_email
    msg['Subject'] = 'Sheet ' + hw_number + ' Correction'

    # NOTE: adjust e-mail text
    msg.set_content(
        f'Hallo {firstname},\n\n'
        + 'anbei findest Du deinen korrigierten Zettel.\n\n'
        + 'Viele Grüße,\n'
        + 'Sang'
    )

    with open(filename, 'rb') as f:
        file_data = f.read()
        file_name = f.name

    msg.add_attachment(
        file_data,
        maintype='application',  # for PDFs
        subtype='octet-stream',  # for PDFs
        filename=file_name,
    )

    # NOTE: adjust SMTP server according to your university
    with smtplib.SMTP('email.stud.uni-goettingen.de', 587) as smtp:
        smtp.ehlo()  # identify ourselves to smtp client
        smtp.starttls()  # secure our email with tls encryption
        smtp.ehlo()  # re-identify ourselves as an encrypted connection
        smtp.login(UNI_USER, UNI_PASS)
        smtp.send_message(msg)

    print(
        f'|    +--- E-Mail to \'{surname}, {firstname}\' was sent succesfully.'
    )


def main():
    UNI_USER = os.environ.get('UNI_USER')  # 'ug-student\firstname.surname'
    UNI_MAIL = os.environ.get('UNI_MAIL')  # eCampus e-mail address
    UNI_PASS = os.environ.get('UNI_PASS')  # password

    # NOTE: adjust column names according to your 'Punkteliste.csv'
    df = pd.read_csv('Punkteliste.csv', sep=';')
    username_col = 'Stud.IP Benutzername'
    surname_col = 'Nachname'
    firstname_col = 'Vorname'


    # FOR TESTING PURPOSES:
    # -------------------------------------------------
    df1 = pd.DataFrame(
        data={
            username_col: 'thesang.nguyen',
            surname_col: 'Nguyen',
            firstname_col: 'The Sang',
        },
        index=[0],
    )
    df2 = pd.DataFrame(
        data={
            username_col: 'thesang.nguyen',
            surname_col: 'Nguyen',
            firstname_col: 'Doppelgänger',
        },
        index=[0],
    )
    df3 = pd.DataFrame(
        data={
            username_col: 'thesang.nguyen',
            surname_col: 'Nguyen2',
            firstname_col: 'Doppelgänger2',
        },
        index=[0],
    )
    frames = [df, df1, df2, df3]
    df = pd.concat(frames, ignore_index=True)
    # -------------------------------------------------

    # duplicate surnames
    duplicate_names = list(df[df.duplicated(surname_col)][surname_col])

    hw_number = get_hw_number()
    path = get_path(hw_number)
    os.chdir(path)  # change into directory of given path

    # -------------------------------------------------
    # NOTE: adjust filetypes and head/tail of filenames
    head = hw_number + '_'
    tail = '_' + 'corrected' + '.'
    filetypes = ['zip', 'pdf', 'ipynb']
    # -------------------------------------------------

    corr_hw, names_in_filenames = get_corr_hw(path, head, tail, filetypes)
    print(f'+--- There are {len(corr_hw)} corrected homework files.')

    unknown = []  # names not found in database during following search
    for idx, names in enumerate(names_in_filenames):
        # split 'names' into names of group members
        names = names.split(sep='_')
        print(f'+--- Handling file number {idx+1} from:', *names)
        for name in names:
            # get dataframe row(s) of student(s) with given name
            dd = df.loc[df[surname_col] == name]
            if name in duplicate_names:  # if students have the same surname
                # list of firstnames with same surname
                options = list(dd[firstname_col])

                # let user choose which student to send e-mail to
                prompt = (
                    f'++++ There are {len(options)} {name}s.\n' +
                    f'|    Do you mean {options[0]} [0]'
                )
                for k in range(1, len(options)):
                    prompt += f' or {options[k]} [{k}]'
                prompt += '? '
                opt = int(input(prompt))

                # unlike surname_col the dataframe index is unique
                dd = dd.loc[df[firstname_col] == options[opt]]

            try:
                # get index of dataframe row of student with given name
                df_idx = dd.index[0]
            except IndexError:  # if student not found in database
                unknown.append(name)
                print(
                    '+--- UNICODE ERROR:\n' +
                    f'|    Name \"{name}\" not found in database ' 
                    + '(possibly due to weird Umlaut-formats)\n' +
                    '|    +-- possible SOLUTIONS: retype filename on PC or '
                    + 'change name in database'
                )
                continue
            else:
                dd = df.iloc[df_idx]  # wanted dataframe row

                # NOTE: adjust to your needs
                # 'corr_hw' and 'names_in_files' share same index of same file
                send_mail(
                    UNI_USER, UNI_MAIL, UNI_PASS,
                    dd[username_col] + '@stud.uni-goettingen.de',
                    dd[firstname_col], dd[surname_col], hw_number, corr_hw[idx],
                )

    if unknown:
        print('+--- ATTENTION: following names were NOT found:\n', unknown)


if __name__ == '__main__':
    main()
