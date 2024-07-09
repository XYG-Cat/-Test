import os
import shutil
import subprocess
import win32api
import win32file
import ctypes
from time import sleep

temp_dir = 'C:\\TempInstaller\\'


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception as e:
        print(e)
        return False


def get_cdrom_drives():
    drives = win32api.GetLogicalDriveStrings()
    drives = drives.split('\000')[:-1]
    cdrom_drives = [drive for drive in drives if win32file.GetDriveType(drive) == win32file.DRIVE_CDROM]
    return cdrom_drives


def delete_path(path):
    try:
        if os.path.isfile(path):
            os.remove(path)
            print(f"文件已成功删除: {path}")
        elif os.path.isdir(path):
            shutil.rmtree(path)
            print(f"文件夹及其内容已成功删除: {path}")
        else:
            print(f"路径不是文件或文件夹: {path}")
    except FileNotFoundError:
        print(f"路径不存在: {path}")
    except PermissionError:
        print(f"没有权限删除: {path}")
    except Exception as e:
        print(f"删除时发生错误: {e}")


def copy_files_from_cd(source_dir, target_dir):
    try:
        # 遍历光盘文件并复制到临时文件夹
        for filename in os.listdir(source_dir):
            source_path = os.path.join(source_dir, filename)
            if os.path.isfile(source_path):
                shutil.copy(source_path, target_dir)
                print(f"复制文件 {filename} 到临时文件夹")
    except FileNotFoundError:
        print(f"未找到光盘 {source_dir}，请确保光盘已插入。")


def extract_archive(path):
    try:
        # 找到所有的分卷压缩包
        os.system('cls')
        print('JetBrains 安装程序 | =================================================================================\
======')
        print()
        print('步骤二: 展开文件...')
        part1 = os.path.join(path, 'JetBrainsSetupForWindowsAmd64.zip.001')

        # 使用 7z 工具解压缩到临时文件夹
        subprocess.run([f'{temp_dir}\\7-Zip\\7z', 'x', part1, f'-o{path}', '-y'], check=True)
        print("缩分卷压缩包展开完成。")
    except subprocess.CalledProcessError:
        print("解压缩过程中出现错误！")


def setup():
    global temp_dir
    os.system('cls')
    print('JetBrains 安装程序 | =================================================================================\
======')
    print()
    print('步骤三: 选择安装...')
    print()
    sleep(0.1)
    print('1. 仅安装IDEA 2024.1')
    sleep(0.1)
    print('2. 仅安装Clion 2024.1')
    sleep(0.1)
    print('3. 仅安装Pycharm 2024.1')
    sleep(0.1)
    print('4. 安装IDEA 2024.1 & Clion 2024.1')
    sleep(0.1)
    print('5. 安装IDEA 2024.1 & Pycharm 2024.1')
    sleep(0.1)
    print('6. 安装Clion 2024.1 & Pycharm 2024.1')
    sleep(0.1)
    print('7. 安装全部')
    sleep(0.1)
    print()

    while True:
        user_choice = input('请选择安装类型(输入数字):')

        if user_choice == '1':
            subprocess.run([f'{temp_dir}/JetBrainsSetupForWindowsAmd64/ideaIU-2024.1.4.exe'])
            break
        elif user_choice == '2':
            subprocess.run([f'{temp_dir}/JetBrainsSetupForWindowsAmd64/CLion-2024.1.4.exe'])
            break
        elif user_choice == '3':
            subprocess.run([f'{temp_dir}/JetBrainsSetupForWindowsAmd64/pycharm-professional-2024.1.4.exe'])
            break
        elif user_choice == '4':
            subprocess.run([f'{temp_dir}/JetBrainsSetupForWindowsAmd64/ideaIU-2024.1.4.exe'])
            subprocess.run([f'{temp_dir}/JetBrainsSetupForWindowsAmd64/CLion-2024.1.4.exe'])
            break
        elif user_choice == '5':
            subprocess.run([f'{temp_dir}/JetBrainsSetupForWindowsAmd64/ideaIU-2024.1.4.exe'])
            subprocess.run([f'{temp_dir}/JetBrainsSetupForWindowsAmd64/pycharm-professional-2024.1.4.exe'])
            break
        elif user_choice == '6':
            subprocess.run([f'{temp_dir}/JetBrainsSetupForWindowsAmd64/CLion-2024.1.4.exe'])
            subprocess.run([f'{temp_dir}/JetBrainsSetupForWindowsAmd64/pycharm-professional-2024.1.4.exe'])
            break
        elif user_choice == '7':
            subprocess.run([f'{temp_dir}/JetBrainsSetupForWindowsAmd64/ideaIU-2024.1.4.exe'])
            subprocess.run([f'{temp_dir}/JetBrainsSetupForWindowsAmd64/CLion-2024.1.4.exe'])
            subprocess.run([f'{temp_dir}/JetBrainsSetupForWindowsAmd64/pycharm-professional-2024.1.4.exe'])
            break
    os.system('cls')
    print('JetBrains 安装程序 | =================================================================================\
======')
    print()
    print('步骤四: 完成安装...')
    print()
    input('按Enter键以继续')
    print('正在删除临时文件, 请耐心等待...')
    print(f'以完成0%')
    for i in range(1, 6):
        delete_path(f'{temp_dir}\\JetBrainsSetupForWindowsAmd64.zip.00{i}')
        print(f'以完成{i * 20}%')
    os.system('cls')
    print('JetBrains 安装程序 | =================================================================================\
======')
    print()
    print()
    print('感谢您使用本安装向导，JetBrains激活工具在 C:\\TempInstaller\\JetBrainsSetupForWindowsAmd64 文件夹下，请自便。')
    input('按Enter键以退出本向导...')
    exit()


def copy_right():
    print()
    print('Copyright (c) 2024 POINT organization.All rights reserved.')
    print()
    print('本安装程序仅供学习参考使用, 不得将本程序用于商业用途, 请于下载 24 小时内删除')
    input('按Enter以继续安装...')


def main():
    global temp_dir

    print(
        'JetBrains 安装程序 | ========================================================================================')
    copy_right()
    print('安装程序正在初始化...')
    try:
        cdrom_drives = get_cdrom_drives()
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        shutil.copytree(f'{cdrom_drives[0]}\\7-Zip', f'{temp_dir}\\7-Zip')

    except Exception as e:
        print('初始化失败, 错误代码:{}'.format(e))
        print('请手动清除C:\\TempInstaller 文件夹, 然后重启')
        input('按Enter以退出...')
        exit()

    print('初始化成功, 现在可以弹出#0光盘')
    print(f'检测到光驱{cdrom_drives[0]}')
    input('请确定光驱位置后按Enter以开始安装...')

    # 提示用户插入每张光盘
    for i in range(1, 6):
        while True:
            os.system('cls')
            print('JetBrains 安装程序 | =================================================================================\
======')
            print()
            print('步骤一: 复制文件...')
            print(
                '|----------------------------------------------------------------------------------------------------|'
            )
            print('|', end='')
            print('#' * i * 20, end='')
            print('|')
            print(
                '|----------------------------------------------------------------------------------------------------|'
            )
            input(f"请插入第 {i} 张光盘后按 Enter 键继续...")
            if os.path.exists(f"{cdrom_drives[0]}JetBrainsSetupForWindowsAmd64.zip.00{i}"):
                break
            else:
                print('未找到光盘，请检测光盘序号或稍等片刻。。。')

        source_dir = cdrom_drives[0]
        print('正在复制文件中，请耐心等待...')

        # 复制光盘文件到临时文件夹
        copy_files_from_cd(source_dir, temp_dir)

    # 解压缩分卷压缩包
    extract_archive(temp_dir)
    setup()


if __name__ == "__main__":
    main()
