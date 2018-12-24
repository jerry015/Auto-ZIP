
# import zipfile
import os
import subprocess
import time

fname = 'Jingui'
lname = 'Zhang'
files_path = 'D:\\bb'
rarname = r"D:\test"


# # 新建压缩包，放文件进去,若压缩包已经存在，将覆盖。可选择用a模式，追加
# azip = zipfile.ZipFile(r'D:\bb\bb.zip', 'w')


# for current_path, subfolders, filesname in os.walk(r'D:\bb'):
#     print(current_path, subfolders, filesname)
#     #  filesname是一个列表，我们需要里面的每个文件名和当前路径组合
#     for file in filesname:
#         # 将当前路径与当前路径下的文件名组合，就是当前文件的绝对路径
#         azip.write(os.path.join(current_path, file))

# # 关闭资源
# azip.close()

# azip.setpassword(lname+'!'+fname)



for current_path, subfolders, filesname in os.walk(files_path):
    for file in filesname:
        print(os.path.join(current_path, file))
        cmd = [r'C:\Program Files\WinRAR\WinRAR.exe', 'a', '-hp'+lname+'!'+fname, rarname, os.path.join(current_path, file)]
        sp = subprocess.Popen(cmd, stderr=subprocess.STDOUT, stdout=subprocess.PIPE)
        time.sleep(5) 