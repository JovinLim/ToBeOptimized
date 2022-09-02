import os
from os.path import join
from os import listdir, rmdir
from shutil import move, copyfile, copytree


folderDir = os.getcwd()
target = "D:\\Jovin stuff\\Unreal\\NotreDame\\NotreDameUE\\Saved\\MovieRenders\\Exhibition_Alleyway_02"
prefix = target + "\\Exhibition_Alleyway_"
suffix = ".png"


folderContents = listdir(target)


for i in range(len(folderContents)):
    if i < 100:
        os.rename(folderContents[i], target+prefix+"0"+str(i)+suffix)
    else:
        os.rename(folderContents[i], target+prefix+str(i)+suffix)