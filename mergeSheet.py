import pandas as pd
import msvcrt
try:
    # filename="new2.xlsx"
    # sheetNameLs=['碳排放量-马克','能源消耗量-马克']
    # match_col_ls=[[1,2],[1,2]]
    # add_col_ls=[[3],[3]]
    sheetNameLs=[]
    match_col_ls=[]
    add_col_ls=[]
    filename=input("请输入文件路径: ")
    print("-------------")
    sheetnum=int(input("请输入要合并的sheet数: "))
    print("-------------")
    for i in range(sheetnum):
        sheetname=input("请输入第{}个sheet的名字: ".format(i+1))
        sheetNameLs.append(sheetname)
        colls=input("请输入<{}>sheet的匹配列(以逗号分隔): ".format(sheetname))
        tmp=[]
        for i in colls.split(','):
            tmp.append(int(i))
        match_col_ls.append(tmp)
        colls=input("请输入<{}>sheet添加的列(以逗号分隔): ".format(sheetname))
        tmp=[]
        for i in colls.split(','):
            tmp.append(int(i))
        add_col_ls.append(tmp)
        print("-------------")
    outputfilename=input("请输入输出的文件名(以.xlsx结尾): ")
    print("开始处理....")
    #加载表格
    df_ls=[]
    for sheet in sheetNameLs:
        df_ls.append(pd.read_excel(filename,sheet_name=sheet))
    #开始处理
    df_num=len(sheetNameLs)
    dict__all={}
    for index,df in enumerate(df_ls):
        match_cols=match_col_ls[index]
        add_cols=add_col_ls[index]
        col_label=df.columns
        #遍历sheet行
        for i in range(df.shape[0]):
            dict_p=dict__all
            for match_col_index in match_cols:
                match_name=df.loc[i][match_col_index]
                if match_name not in dict_p.keys():
                    dict_p[match_name]={}
                dict_p=dict_p[match_name]
            for add_col_index in add_cols:
                add_value=df.loc[i][add_col_index]
                add_label=col_label[add_col_index]
                dict_p[add_label]=add_value
    #写入表格
    data_ls=[]

    ##获取match 标题
    match_label=[]
    for i in match_col_ls[0]:
        match_label.append(df_ls[0].columns[i])
    ##获取add标题
    add_label_all=[]
    for index,add_labels in enumerate(add_col_ls):
        cols=df_ls[index].columns
        for i in add_labels:
            add_label_all.append(cols[i])

    match_num=len(match_label)
    def dfs(depth,dict,line):
        if depth == match_num:
            newline=line[:]
            for key in add_label_all:
                if key in dict.keys():
                    newline.append(dict[key])
                else:
                    newline.append('')
            data_ls.append(newline)
            return 
        else:
            for key in dict:
                line.append(key)
                dfs(depth+1,dict[key],line)
                line.pop()
    dfs(0,dict__all,[])
    wpd=pd.DataFrame(data_ls,columns=match_label+add_label_all)
    # print(wpd)
    wpd.to_excel(outputfilename,index=False)
    print("完成，文件已写入到{}".format(outputfilename))
except Exception as ex:
    print("有一些错误发生了:",ex)
print("请按任意键退出~")
ord(msvcrt.getch())