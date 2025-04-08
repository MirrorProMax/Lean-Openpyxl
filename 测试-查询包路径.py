import sys
from SubFunc import SubFunc

SubFunc.clearConsole()

# 列出当前python的包路径
pathList = sys.path
for i in pathList:
    print(i)

