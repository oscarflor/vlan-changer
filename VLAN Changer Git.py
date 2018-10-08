#!/usr/bin/python3.6
from __future__ import print_function
from googleapiclient.discovery import build
from googleapiclient import errors
from httplib2 import Http
from oauth2client import file, client, tools
from pytz import timezone
from datetime import datetime
from re import match
from netmiko import ConnectHandler, ssh_exception
from email.mime.text import MIMEText
import base64
import time
import threading
import queue

# If modifying these scopes, delete the file token.json.
SCOPES_SHEETS = 'https://www.googleapis.com/auth/spreadsheets'
SCOPES_GMAIL = 'https://www.googleapis.com/auth/gmail.send'

# The ID and range of the VLAN Change Request Form results spreadsheet.
SPREADSHEET_ID = 'yourSheetID'
RANGE_NAME = 'A2:I'

# List of people to receive notifications for failed attempts
EMAILS = ['netadminsemails@gmail.com']

# Maximum number of connections allowed at one time
MAX_THREADS = 5

# Job queue for new submissions (not a constant)
q = queue.Queue()

# Sleep time in seconds
seconds = 10

# TODO: Edit with appropriate VLAN IDs
VLAN_DICT = {
    'Management': 1,
    'Food Services': 2,
    'UPS': 3,
    'Copiers': 4,
    'Printers': 5,
    'Cameras': 6,
    'Air Con': 7,
    'DATA': 8,
    'Administrator': 9,
    'vBrick': 10,
    'Servers': 11,
    'Wireless': 12,
}


def main():
    """
    Reads through submitted responses making changes on new entries.
    :return: None
    """
    while True:
        service_sheets = get_google_service('sheets', 'v4', 'token_sheets.json', SCOPES_SHEETS)
        result = service_sheets.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID,
                                                            range=RANGE_NAME).execute()
        values = result.get('values', [])
        new_data_found = False

        if values:
            current_row = 2
            global q
            q = queue.Queue()
            threads = []
            for data in values:
                try:
                    status = data[6]
                    if status == 'Authentication Failure' \
                            or status == 'Connection Timeout':
                        new_data_found = True
                        q.put((data, current_row))
                except IndexError:
                    new_data_found = True
                    q.put((data, current_row))
                current_row += 1

            if new_data_found:
                num_workers = q.qsize() if q.qsize() < MAX_THREADS else MAX_THREADS
                print ('num workers: ', num_workers)
                for i in range(num_workers):
                    thread_name = 'thread {}'.format(i)
                    t = threading.Thread(target=worker, name=thread_name, daemon=True)
                    t.start()
                    threads.append(t)
                print('workers created')
                print('queue size: ', q.qsize())
                q.join()
                for i in range(num_workers):
                    q.put(None)
                for t in threads:
                    t.join()
                print('workers killed')
                print('queue size: ', q.qsize())

        if not new_data_found:
            print('sleeping')
            time.sleep(seconds)


def worker():
    """
    Creates a worker/thread. Changes vlans until no jobs are left in the queue.
    :return: None
    """
    service_gmail = get_google_service('gmail', 'v1', 'token_gmail.json', SCOPES_GMAIL)
    service_sheets = get_google_service('sheets', 'v4', 'token_sheets.json', SCOPES_SHEETS)
    while True:
        data = q.get()
        if data is None:
            print('Worker freed')
            break
        print('Worker assigned new job')
        device_type = get_system_type(ip=data[0][3])
        current_row = data[1]
        change_vlan(data[0], device_type, current_row, service_sheets, service_gmail)
        q.task_done()
        print('worker finished job')


def get_google_service(type, version, token, scope):
    """
    Builds and returns a google service.

    :param type: String - Type of service to build, either 'gmail' or 'sheets'
    :param version: String - Service version e.g. 'v1' 'v4'
    :param token: String - Name of token file to create
    :param scope: String - Scope of the service i.e. permissions
    :return: googleapiclient.discovery.Resource
    """
    store = file.Storage(token)
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', scope)
        creds = tools.run_flow(flow, store)
    return build(type, version, http=creds.authorize(Http()))


def create_message(form_data, status, to=', '.join(EMAILS), subject='VLAN change fail'):
    """
    Creates and returns a message.

    :param form_data: List - Submitted responses
    :param status: String - Status to include in the message body
    :param to: String - List of recipients
    :param subject: String - Subject of the email
    :return: Dict - Message
    """
    message_text = str(form_data) + '\n\nStatus: ' + status
    message = MIMEText(message_text)
    message['to'] = to if len(EMAILS) > 1 else EMAILS[0]
    message['from'] = ''
    message['subject'] = subject
    raw = base64.urlsafe_b64encode(message.as_bytes())
    raw = raw.decode()
    message = {'raw': raw}
    return message


def send_email(service_gmail, message, user_id='me'):
    """
    Sends an email

    :param service_gmail: Authorized Gmail API service instance.
    :param message: Message to be sent
    :param user_id: User's email address. The special value "me"
                    can be used to indicate the authenticated user.
    :return: None
    """
    try:
        message = (service_gmail.users().messages().send(userId=user_id, body=message)
                   .execute())
        return message
    except errors.HttpError as error:
        print('An error occurred: %s' % error)


def update_sheet(range_to_edit, status, service_sheets, desired_vlan, current_vlan="\' \'"):
    """
    Updates the status, change timestamp and vlan change details on the sheet

    :param range_to_edit: String - Range of cells to be updated e.g. 'G2:I2'
    :param status: String - New status e.g. 'Successful'
    :param service_sheets: googleapiclient.discovery.Resource
    :param desired_vlan: String - Desired VLAN according to submitted response
    :param current_vlan: String - Current VLAN 'name, id'
    :return: None
    """
    change_info = current_vlan + ' >> ' + desired_vlan + ', ' + str(VLAN_DICT.get(desired_vlan))
    new_status = {
        'range': range_to_edit,
        'majorDimension': 'ROWS',
        'values': [[status, get_time(), change_info]]
    }
    request = service_sheets.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID, range=range_to_edit,
        valueInputOption='USER_ENTERED', body=new_status)
    request.execute()


def get_time():
    """
    Returns a String representation of the current time

    :return: String - Month/Day/Year HH:MM:SS
    """
    tz = timezone('America/Chicago')
    current_date, current_time = str(datetime.now(tz)).split('.', 1)[0].split(' ')
    year, month, day = current_date.split('-')
    month = month.lstrip('0')
    day = day.lstrip('0')
    formatted_date = '%s/%s/%s %s' % (month, day, year, current_time)
    return formatted_date


def get_system_type(ip):
    """
    FOR FUTURE USE.
    If more than one OS is in use, ideally this function would read
    from a file that lists IPs and OS types. In this case only cisco_ios
    is in use.
    :param ip: String - IP address of the device
    :return: String - OS of the device
    """
    return 'cisco_ios'


def change_vlan(form_data, device_type, current_row, service_sheets, service_gmail):
    """

    :param form_data: List of data of the current row in the sheet
    :param device_type: OS of the device e.g. cisco_ios
    :param current_row: row in the google sheet
    :param service_sheets: googleapiclient.discovery.Resource
    :param service_gmail: googleapiclient.discovery.Resource
    :return: None
    """
    range_to_edit = 'G' + str(current_row) + ':I' + str(current_row)
    desired_vlan = form_data[1]
    status = 'Connecting...'
    update_sheet(range_to_edit, status, service_sheets, desired_vlan)

    if device_type == 'cisco_ios':
        cisco_ios_change(form_data, range_to_edit, service_sheets, service_gmail)
    elif device_type == 'some HP':
        print()
        # TODO: some HP function, Extreme etc


def cisco_ios_change(form_data, range_to_edit, service_sheets, service_gmail):
    """
    Attempts an SSH connection to the switch.
    Exits if connection fails to be established.
    Changes VLAN as needed, only saving if change is validated successfully.

    :param form_data: List of data of the current row in the sheet
    :param range_to_edit: String - Range of cells to be updated e.g. "G2:I2"
    :param service_sheets: googleapiclient.discovery.Resource
    :param service_gmail: googleapiclient.discovery.Resource
    :return: None
    """
    ip_address = form_data[3]
    port = str(form_data[4]).replace(' ', '')
    desired_vlan = form_data[1]
    if port.upper()[0] == 'G':
        port = port[:2] + port[15:]
    else:
        port = port[:2] + port[12:]

    # adjust user/pass accordingly
    cisco_ios_switch = {
        'device_type': 'cisco_ios',
        'ip': ip_address,
        'username': 'oscar',
        'password': 'cisco',
        'port': 22,         # optional, defaults to 22
        'secret': '',       # optional, defaults to ''
        'verbose': False    # optional, default to False
    }

    net_connect, status = get_connection(cisco_ios_switch)
    update_sheet(range_to_edit, status, service_sheets, desired_vlan)

    # Stops if a connection fails to be established
    if status != 'Connection Established':
        return send_email(service_gmail, message=create_message(form_data, status))

    output = net_connect.send_command('show vlan br')
    current_vlan = get_current_cisco_vlan(output, port)
    current_vlan = current_vlan[0] + ', ' + current_vlan[1]
    status = 'Attempting change...'

    update_sheet(range_to_edit, status, service_sheets, desired_vlan, current_vlan)

    # Stops if port is not access or port is not found (Should probably never happen)
    if current_vlan == 'None, -1':
        if is_trunk(net_connect.send_command('show int trunk'), port):
            status = 'Port is trunk'
        else:
            status = 'Port not found'
        update_sheet(range_to_edit, status, service_sheets, desired_vlan, current_vlan)
        return send_email(service_gmail, message=create_message(form_data, status))

    interface_cmd = 'interface ' + port
    vlan_cmd = 'switchport access vlan ' + str(VLAN_DICT.get(desired_vlan))
    net_connect.send_config_set([interface_cmd, vlan_cmd])

    output = net_connect.send_command('show vlan br')
    if validate_cisco(output, port, desired_vlan):
        net_connect.save_config()
        status = 'Successful'
        send_email(service_gmail, message=create_message(form_data, status, to=form_data[5], subject='VLAN Changed'))
    else:
        status = 'Failed'
        send_email(service_gmail, message=create_message(form_data, status))

    net_connect.cleanup()
    net_connect.disconnect()

    update_sheet(range_to_edit, status, service_sheets, desired_vlan, current_vlan)


def get_connection(device):
    """
    Returns a Connection to the device and a status message

    :param device: Dict - holds device information
    :return: Connection, String - Netmiko Connection if possible, Status
    """
    net_connect = ''
    try:
        net_connect = ConnectHandler(**device)
        status = 'Connection Established'
    except ssh_exception.NetMikoAuthenticationException:
        status = 'Authentication Failure'
    except ssh_exception.NetMikoTimeoutException:
        status = 'Connection Timeout'
    return net_connect, status


def get_current_cisco_vlan(output, port):
    """
    Returns the current VLAN assignment for a given interface.

    Reads the output of "show vlan brief" line by line storing the last read
    vlan ID and name. When the interface is found, the current stored
    ID and name are returned. If no match is found, ('None', '0') is returned.

    :param output: String - Output of "show vlan brief"
    :param port: String - Interface to be modified (e.g. Gi0/1 Fa0/1)
    :return: String Tuple - (vlan NAME, vlan ID) e.g. (DATA, 8)
    """
    for lines in output.strip().splitlines():
        if match(r'\d{1,4}', lines[0:3].strip()):
            vlan_id = lines[0:3].strip()
            vlan_name = lines[5:37].strip()
        if port in lines:
            if vlan_name.upper() != 'VOICE':
                return vlan_name, vlan_id
    return 'None', '-1'


def validate_cisco(output, port, vlan):
    """
    Returns True if current VLAN matches desired VLAN, False otherwise

    :param output: String - Output of "show vlan brief"
    :param port: String - Interface to be modified (e.g. Gi0/1 Fa0/1)
    :param vlan: String - Desired VLAN according to submitted response
    :return: Bool
    """
    desired_vlan = str(VLAN_DICT.get(vlan))
    if get_current_cisco_vlan(output, port)[1] == desired_vlan:
        return True
    return False


def is_trunk(output, port):
    """
    Returns True port is trunk, False otherwise

    :param output: String - Output of 'show int trunk'
    :param port: String - The port e.g. Gi0/1
    :return: Bool
    """
    for lines in output.strip().splitlines():
        if port in lines:
            return True
    return False


if __name__ == '__main__':
    main()

