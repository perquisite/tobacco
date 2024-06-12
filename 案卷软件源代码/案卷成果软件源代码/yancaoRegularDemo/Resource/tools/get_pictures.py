import zipfile
import shutil
import os


def get_pictures_single(dir_path):
    if os.path.exists(dir_path) == False:
        return False
    image_path = os.path.join(dir_path, "picture")
    if os.path.exists(os.path.join(dir_path, "picture")):
        shutil.rmtree(os.path.join(dir_path, "picture"))
    for filepath, dirnames, filenames in os.walk(dir_path):
        for filename in filenames:
            # print(filename)
            # print(filepath)
            if "~$" in filename or ".txt" in filename or '.xls' in filename:
                continue
            doc = zipfile.ZipFile(os.path.join(filepath, filename))
            for info in doc.infolist():
                if info.filename.endswith((".png", ".jpeg", ".gif", ".bmp", ".jpg")):
                    # print(info.filename)
                    doc.extract(info.filename, os.path.join(image_path, filename.split('.')[0]))


def get_pictures_multi(dir_path):
    if os.path.exists(dir_path) == False:
        return False
    for filepath, dirnames, filenames in os.walk(dir_path):
        for dirname in dirnames:
            get_pictures_single(dir_path + "\\" + dirname)


if __name__ == '__main__':
    dir_path = r'C:\Users\Zero\Desktop\副本\\'
    get_pictures_single(dir_path)
    # get_pictures_single(dir_path)
