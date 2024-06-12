import os
import shutil
from pathlib import Path

src_folder = os.path.abspath('.')

"""
先安装whl
前期准备工作参考：https://www.cnblogs.com/NaughtyCat/p/python-pip-freeze-to-package-offline-packages.html
"""
Python_Freeze_dir = Path(src_folder, "Python_Freeze")
requirements_dir = Path(src_folder, 'requirements.txt')
os.system(f'pip install --no-index --find-links={Python_Freeze_dir} -r {requirements_dir} ')

"""
将WorkSpace复制到C盘
"""
print(src_folder)
try:
    source_path = os.path.abspath(Path(src_folder, 'WorkSpace'))
    target_path = os.path.abspath(r'C:\WorkSpace')
    if os.path.exists(target_path):
        # 如果目标路径存在原文件夹的话就删除
        shutil.rmtree(target_path)
    # 粘贴文件夹
    shutil.copytree(source_path, target_path)
    print("复制WorkSpace成功！")
except:
    print("复制WorkSpace时出现问题")

"""
移动bert文件
"""
try:
    target_path = Path('C:\\', "Users", os.getlogin(), ".cache", 'torch')
    source_path = Path(src_folder, "torch")
    if os.path.exists(target_path):
        shutil.rmtree(target_path)
    shutil.copytree(source_path, target_path)
    print("复制Torch成功！")
except:
    print("复制Torch时出现问题")

"""
创建保存审批文书和合同的文件夹
"""
try:
    if os.path.exists(r'C:\examine'):
        shutil.rmtree(r'C:\examine')
    os.mkdir(r'C:\examine')
    os.mkdir(r'C:\examine\files')
    os.mkdir(r'C:\examine\contract')
    print("复制Examine成功！")
except:
    print("复制Examine时出现问题")
"""
等待用户按回车后退出
"""
input('脚本执行完成。请确认无报错后，按<Enter>退出。')
