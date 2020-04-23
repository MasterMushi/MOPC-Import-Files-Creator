import configparser
import copy
import csv
import ctypes
import os
import sys


def get_settings_from_ini_file(cwd, name='settings.ini'):
    # Configuration ini-file reading
    ini_path = os.path.join(cwd, name)

    with open(ini_path) as f:
        cfg = configparser.ConfigParser()
        cfg.read_file(f)
        param_vex_path = cfg['settings']['param_vex_path']

    return param_vex_path


def get_dict_from_param_file(path):
    with open(path, 'r') as csv_file:
        for i in range(2):
            # Start reading from the 2nd line
            next(csv_file)

        params_dict = csv.DictReader(f=csv_file,
                                     delimiter=';',
                                     quoting=csv.QUOTE_ALL)

        short_param_dict = []
        for row in params_dict:
            if '_DI_' in row['Group']:
                ctyp = 'DI'
                spec = row['Spec'][2:]
            if '_DO_' in row['Group']:
                ctyp = 'DO'
                spec = row['Spec'][2:]
            if 'UVL_' in row['Group']:
                ctyp = 'UVL'
                spec = row['Spec'][2:]
            if 'EvalPar_' in row['Group']:
                ctyp = 'EvalPar'
                spec = str(int(row['Spec'][2:]) - 400000)
            short_param_dict.append({'Varname': row['\'Varname'],
                                     'Conn': row['Conn'],
                                     'Group': row['Group'],
                                     'Spec': spec,
                                     'CTyp': ctyp})

    return short_param_dict


def create_mopc_folders_file(cwd, params):
    # Create "Folders" part
    groups = []
    for row in params:
        if row['Group'] not in groups:
            groups.append(row['Group'])

    folders = []
    for group in groups:
        folders.append({'Item Location': r'\\Address Space\OPC_SERVER\Virtual_Box',
                        'Name': group,
                        'Simulate': 'No',
                        'Use Simple Template': 'No',
                        'Simple Template': '',
                        'Use Parameterized Template': 'No',
                        'Parameterized Template': '',
                        'Starting Address Base': '0'})

    path = os.path.join(cwd, 'import\\folders_import.csv')
    with open(path, 'w', newline='') as csv_file:
        writer = csv.DictWriter(f=csv_file,
                                fieldnames=folders[0].keys(),
                                delimiter=',',
                                quoting=csv.QUOTE_ALL)
        writer.writeheader()
        for folder in folders:
            writer.writerow(folder)


def create_mopc_data_item_file(cwd, params):
    # Create "Data Items" part
    template_data_item = {'Item Location':        '',  # set group
                          'Name':                 '',  # set name
                          'Simulate':             'No',
                          'Simulation Signal':    '',
                          'Location Type':        '',  # set loc_type
                          'Read Access':          '',  # set r_access
                          'Write Access':         '',  # set w_access
                          'Starting Address':     '',  # set address
                          'Bit Field':            'No',
                          'Bit Number':           0,
                          'Bit Count':            1,
                          'Modbus Type':          '',  # set bool or real
                          'Data Length':          10,
                          'Vector':               'No',
                          'Number of elements':   20,
                          'Manual':               'No',
                          'Manual Value':         '',
                          'Use conversion':       'No',  # set Yes if conversion
                          'Conversion':           'None (to/from float)',
                          'Message Prefix':       '',
                          'Generate Alarms':      'No',
                          'Limit Alarm Def.':     '',
                          'Digital Alarm Def.':   '',
                          'EU Units':             '',
                          'Description':          '',
                          'Close Label':          '',
                          'Open Label':           '',
                          'Default Display':      '',
                          'Foreground Color':     0,
                          'Background Color':     0,
                          'Blink':                'No',
                          'BMP File':             '',
                          'Sound File':           '',
                          'HTML File':            '',
                          'AVI File':             '',
                          'Monitor':              'Yes',
                          'Associated':           'No',
                          'Associated type':      1,
                          'Association':          '',
                          '&Device	Ctrl+D':      '',
                          'Master Tag':           'No',
                          'Code Number':          'No',
                          'Number of a Code':     0
                          }

    data_items = []
    for row in params:
        template_data_item['Item Location'] = '\\\\Address Space\\OPC_SERVER\\Virtual_Box\\' + row['Group']
        template_data_item['Name'] = row['Varname']

        if row['CTyp'] == 'DI' or row['CTyp'] == 'UVL':
            template_data_item['Location Type'] = 'Coil (bit, r/w)'
            template_data_item['Read Access'] = 'Yes'
            template_data_item['Write Access'] = 'No'
            template_data_item['Modbus Type'] = 'BOOL'
            template_data_item['Use conversion'] = 'No'
        if row['CTyp'] == 'DO':
            template_data_item['Location Type'] = 'Coil (bit, r/w)'
            template_data_item['Read Access'] = 'No'
            template_data_item['Write Access'] = 'Yes'
            template_data_item['Modbus Type'] = 'BOOL'
            template_data_item['Use conversion'] = 'No'
        if row['CTyp'] == 'EvalPar':
            template_data_item['Location Type'] = 'Holding Register (word, r/w)'
            template_data_item['Read Access'] = 'Yes'
            template_data_item['Write Access'] = 'No'
            template_data_item['Modbus Type'] = 'REAL'
            template_data_item['Use conversion'] = 'Yes'

        template_data_item['Starting Address'] = row['Spec']
        copy_template_data_item = copy.deepcopy(template_data_item)
        data_items.append(copy_template_data_item)

    path = os.path.join(cwd, 'import\\data_items_import.csv')
    with open(path, 'w', newline='') as csv_file:
        writer = csv.DictWriter(f=csv_file,
                                fieldnames=template_data_item.keys(),
                                delimiter=',',
                                quoting=csv.QUOTE_ALL)
        writer.writeheader()
        for data_item in data_items:
            writer.writerow(data_item)


if __name__ == '__main__':
    cwd = os.path.abspath(os.path.dirname(sys.argv[0]))

    param_file_path = get_settings_from_ini_file(cwd)

    params_dict = get_dict_from_param_file(param_file_path)

    create_mopc_folders_file(cwd, params_dict)
    create_mopc_data_item_file(cwd, params_dict)

    ctypes.windll.user32.MessageBoxW(0, 'Import files created',
                                     'MOPC Import Files Creator!', 1)
