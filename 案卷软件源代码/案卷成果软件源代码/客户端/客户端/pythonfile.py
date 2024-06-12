#coding=utf-8
import shutil
import os
import sys

#windows 环境下
orginurl=sys.argv[1]
newurl=sys.argv[2]
print("源文件地址:",orginurl)
print("新文件地址:",newurl)
shutil.copyfile(orginurl,newurl)
