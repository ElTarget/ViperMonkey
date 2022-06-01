import os


def exec_cmd(cmd):
    with os.popen(cmd) as f:
        text = f.read()
    return text

def tes1t():
    dir_path = r'D:\download_malware\output_daily\2022-04-12'
    # dir_path = r'D:\download_malware\output_daily\test'
    office_files = []
    for root, dirs, files in os.walk(dir_path):
        for file in files:
            filesize = os.path.getsize(os.path.join(root, file)) / 1024
            path = os.path.join(root, file)
            office_files.append(os.path.join(root, file))
    for path in office_files:
        cmd_str = r'python vmonkey.py {}'.format(path)
        print('1----------------------------------------------', path)
        print(exec_cmd(cmd_str))
        print('2----------------------------------------------')

if __name__ == '__main__':
    # tes1t()
    # path="12529c7d367852ae7f0617a54c4b1b6952c1ecb05b1e156a63568f0f43746464.xlsx"
    # cmd_str = r'python vmonkey.py {}'.format(path)
    # print(exec_cmd(cmd_str))
    # os.remove('C:\\Users\\792293\\AppData\\Local\\Temp\\tmpz2fmvuz1')
