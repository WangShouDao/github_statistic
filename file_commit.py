import os
from git.repo import Repo
import xlrd
import xlwt
from xlutils.copy import copy

# 计算所有人的 hello-world 版本库的文件路径
def target_path(path):
    pathList = []
    if os.path.exists(path):
        files = os.listdir(path)
        for file in files:          # 每个人所在的文件间
            subPath = os.path.join(path, file)  # 拼接成每个人所在文件夹的完整路径
            if os.path.isdir(subPath):
                subFile = os.listdir(subPath)[-1]       # hello-world的版本库
                targetPath = os.path.join(subPath, subFile)
                pathList.append(targetPath)
    return pathList

# 从远程仓库拉取更新
def pull_request(pathList):
    addList, deleteList = [], []
    for path in pathList:
        repo = Repo(path)
        # 检查版本库是否为空
        if repo.bare:
            return None
        # 获取远程默认版本库 origin
        remote = repo.remote()
        try:
            # 从远程版本库拉取分支
            remote.pull()
        except Exception as e:
            print(path)
            print(str(e))
        git = repo.git
        # 读取本地版本库的信息
        strList = git.log('--numstat').split()
        # 需要计算的文件
        file = 'Units.cpp'
        # 运行前需要设置时间
        add, delete = count_add_delete(strList, file, 'Jun',10, 'Jun', 16)
        addList.append(add)
        deleteList.append(delete)
    print('add:', addList)
    print('delete:', deleteList)
    write_excel(addList, deleteList)

# 从读取的list中计算出增加删除的代码量
def count_add_delete(strList, file, month1, day1, month2, day2):
    add, delete = 0, 0
    # 记录有多少次提交
    numList = []
    for i in range(len(strList)):
        if strList[i] == 'Author:':
            numList.append(i)
    # 计算总的提交删除代码量
    for i in range(len(numList)):
        if i!=len(numList)-1:
            # 如果是添加代码者是本人并且“Units.cpp”包含在commit中
            if 'Lake' not in strList[numList[i]:numList[i+1]]:
                if 'Dashark' not in strList[numList[i]:numList[i+1]]:
                    if file in strList[numList[i]:numList[i+1]]:
                        li = strList[numList[i]:numList[i+1]]
                        month = strList[numList[i] + li.index('Date:') + 2]
                        day = int(strList[numList[i] + li.index('Date:') + 3])
                        # 如果一周在同一个月中
                        if month1 == month2:
                            # 如果日期在这周内
                            if (month == month1) & (day >=day1) & (day<=day2):
                                # 如果有插入或删除纪录
                                add, delete = calc_add_delete(file, add, delete, li, strList, numList, i)
                        else:
                            count = count_day(month1)
                            # 如果日期在这段时间内
                            if ((month == month1) & (day>=day1) & (day<=count)) | ((month==month2) & (day>=1) & (day<= day2)):
                                add, delete = calc_add_delete(file, add, delete, li, strList, numList, i)
            else:
                continue
        elif i==len(numList)-1:
            if 'Lake' not  in strList[numList[i]:]:
                if 'Dashark' not in strList[numList[i]:]:
                    if file in strList[numList[i]:]:
                        li = strList[numList[i]:]
                        month = strList[numList[i] + li.index('Date:') + 2]
                        day = int(strList[numList[i] + li.index('Date:') + 3])
                        # 如果一周在同一个月中
                        if month1 == month2:
                            # 如果日期在这周内
                            if (month == month1) & (day >= day1) & (day <= day2):
                                add, delete = calc_add_delete(file, add, delete, li, strList, numList, i)
                        else:
                            count = count_day(month1)
                            # 如果日期在这段时间内
                            if (month == month1 & day>=day1 & day<=count) | (month==month2 & day>=1 & day<= day2):
                                add, delete = calc_add_delete(file, add, delete, li, strList, numList, i)
            else:
                continue
    return add, delete

# 如果时间段不在一个月中，计算前一月有多少天
def count_day(month):
    if month=='May':
        return 31
    else:
        return 30

# 如果有插入或删除纪录, 计算每次的插入或删除量
def calc_add_delete(file, add, delete, li, strList, numList, i):
    if file in li:
        if strList[numList[i] + li.index(file)-2].isdigit():
            add += int(strList[numList[i] + li.index(file)-2])
        if strList[numList[i] + li.index(file)-1].isdigit():
            delete += int(strList[numList[i] + li.index(file)-1])
    return add, delete

# 写入Excel
def write_excel(addList, deleteList):
    # excel 文件路径
    xlsFile= r'F:\statistic_github\github.xls'
    # 获取Excel 文件的book对象，实例化对象
    rb = xlrd.open_workbook(xlsFile, formatting_info=True)
    # 复制产生一个新的excel进行写入
    wb = copy(rb)
    # 通过sheet索引获得sheet对象
    sheet = wb.get_sheet(1)
    # 设置写入格式
    style = xlwt.XFStyle()
    # 边框
    borders = xlwt.Borders()
    borders.bottom = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.left = xlwt.Borders.THIN
    style.borders = borders
    # 位置
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER
    style.alignment = alignment
    k = 7
    # 写入excel
    for i in range(2, 2 + len(addList)):
        # i为行， k为列
        sheet.write(i, k, addList[i-2], style)
        sheet.write(i, k+1, deleteList[i-2], style)
        sheet.write(i, k+2, addList[i-2] - deleteList[i-2], style)
    # 写完后进行保存
    wb.save(xlsFile)

if __name__=='__main__':
    path = r'F:\statistic_github'
    pathList = target_path(path)
    pull_request(pathList)