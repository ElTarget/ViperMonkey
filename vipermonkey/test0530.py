import os
import re


def exec_cmd(cmd):
    with os.popen(cmd) as f:
        text = f.read()
    return text


def tes1t():
    dir_path = r'D:\gitPro\ViperMonkey\new_test1'
    office_files = []
    no_found_pattern = 'No VBA macros found.'
    # detect_pattern = 'Recorded Actions:\t(.*?)\tsuspicion\t(.*?)\n'
    called_patterns = '(?<=VBA Builtins Called: )(.*?)(?=\n)'
    no_found = {}
    vba_builtins_called_dict = {}
    for root, dirs, files in os.walk(dir_path):
        for file in files:
            filesize = os.path.getsize(os.path.join(root, file)) / 1024
            path = os.path.join(root, file)
            office_files.append(path)
    for path in office_files:
        cmd_str = r'python vmonkey.py {}'.format(path)
        data = exec_cmd(cmd_str)
        no = re.findall(no_found_pattern, data)
        if no:
            no_found.update({path: no})
        call = re.findall(called_patterns, data)
        if call:
            vba_builtins_called_dict.update({path: no})

        # items = re.findall(pattern, data)
        # detects = re.findall(detect_pattern, data)
        # statistics = {}
        # for pattern in scan_result_patterns:
        #     items = re.findall(pattern, data)
        print('1----------------------------------------------', path)
        print(data)
        print('2----------------------------------------------')
    print(len(no_found), len(vba_builtins_called_dict))


if __name__ == '__main__':
    # tes1t()
    from vmonkey import vmonkey_scan

    # fileinfo = {'path': r"D:\gitPro\ViperMonkey\new_test1\0f0d40304fc53e1990ed8b499803d14f91d0bb65daad57bfea11837a43083d30.xlsx"}
    # fileinfo = {'path': r"D:\gitPro\ViperMonkey\new_test1\00367d927216bdf6e8d8e9d825efce6f356f469e5358e604bd26d2e9396e27de.xlsm"}
    dir_path = r'D:\gitPro\ViperMonkey\new_test2'
    office_files = []
    no_found = {}
    vba_builtins_called_dict = {}
    for root, dirs, files in os.walk(dir_path):
        for file in files:
            filesize = os.path.getsize(os.path.join(root, file)) / 1024
            path = os.path.join(root, file)
            office_files.append(path)
    for path in office_files:
        fileinfo = {'path': path}
        print('开始',path)
        res = vmonkey_scan(fileinfo["path"])
        print(path,res)

