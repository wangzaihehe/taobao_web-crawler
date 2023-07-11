import os
# 添加需要安装的扩展包名称进去
libs = {"pyautogui" , "selenium" , "python-docx" , "beautifulsoup4","requests","Pillow","imageio","numpy ","python-libmagic","opencv-python"}
if "pip" in libs:
    print("found pip")
else:
    os.system(" curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py" )
    os.system(" sudo python get-pip.py")
try:
    for lib in libs:
        os.system(" pip install " + lib)
        print("{}   Install successful".format(lib))
except:
    print("{}   failed install".format(lib))
