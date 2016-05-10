# -*- coding: utf-8 -*-
__author__ = 'CHENZH'
import xlrd
import xlwt
import xdrlib, sys
import os

style_tmp = 'font: name Times New Roman'
#GGSN_device_list = [u'GGSN7', u'GGSN8', u'GGSN9', u'GGSN10', u'GGSN11', u'GGSN4-1', u'GGSN12', u'GGSN13',u'GGSN14',u'GGSN15']
GGSN_device_list = [u'UGW7', u'UGW8', u'UGW9', u'UGW10', u'UGW11', u'UGW4-1', u'UGW12', u'UGW13',u'UGW14',u'UGW15']
SGSN_device_list = [u'SGSN5', u'SGSN6', u'SGSN7', u'SGSN8', u'SGSN9', u'SGSN11',u'SGSN14',u'SGSN15']
####################################
############读取现有xls#############
####################################
def read_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception,e:
        print str(e)


def save_excel(book):
    book.save(u'资源预警.xls')

####################################
##########打开一个新xls#############
####################################
def open_excel():
    try:
        book = xlwt.Workbook()
        #GGSN = book.add_sheet('GGSN')
        return book
    except Exception,e:
        print str(e)

#################################
###用于设备名称的提取，暂时没用##
#################################

def detect_device_name(table, device_col):
    col_number=table.col_values(device_col)
    tmp_list=[]
    for i in col_number:
        if(i not in tmp_list):
            tmp_list.append(i)
    return tmp_list

#################################
############Gn 口 峰值提取#######
#################################
def find_Gn_top(table,device_list,Gn_top_col,device_col):
    top_value_list=[]
    for dev in device_list:
        num_max=0
        for num in range(table.nrows):
            if dev==table.cell_value(num,device_col):
                if table.cell_value(num,Gn_top_col) > num_max:
                    num_max = table.cell_value(num,Gn_top_col)
                else:
                    continue
            else:
                continue
        num_max = round(float(num_max) * 8/1024,2)
        top_value_list.append(num_max)

    #print top_value_list
    return top_value_list


def find_Gi_top(table,device_list,Gi_top_col,device_col):
    top_value_list=[]
    for dev in device_list:
        num_max=0
        for num in range(table.nrows):
            if dev==table.cell_value(num,device_col):
                if (table.cell_value(num,Gi_top_col)+table.cell_value(num,Gi_top_col+1)) > num_max:
                    num_max = (table.cell_value(num,Gi_top_col)+table.cell_value(num,Gi_top_col+1))
                else:
                    continue
            else:
                continue
        num_max = round(float(num_max) * 8/1024/1024,2)
        top_value_list.append(num_max)

    #print top_value_list
    return top_value_list

def find_PDP_top(table,device_list,PDP_top_col,device_col):
    top_value_list = []
    for dev in device_list:
        num_max=0
        for num in range(table.nrows):
            if dev==table.cell_value(num,device_col):
                if table.cell_value(num,PDP_top_col) > num_max:
                    num_max = table.cell_value(num,PDP_top_col)
                else:
                    continue
            else:
                continue
        num_max = round(float(num_max)/10000,2)
        top_value_list.append(num_max)

    #print top_value_list
    return top_value_list

def find_GTP_top(table,device_list,GTP_top_col,device_col):
    top_value_list = []
    for dev in device_list:
        num_max=0
        for num in range(table.nrows):
            if dev==table.cell_value(num,device_col):
                if table.cell_value(num,GTP_top_col) > num_max:
                    num_max = table.cell_value(num,GTP_top_col)
                else:
                    continue
            else:
                continue
        num_max = round(float(num_max)/1024/1024*8,2)
        top_value_list.append(num_max)

    #print top_value_list
    return top_value_list


def find_Iu_top(table,device_list,Iu_top_col,device_col,GTP_top_col):
    top_value_list = []
    for dev in device_list:
        num_max=0
        for num in range(table.nrows):
            if dev==table.cell_value(num,device_col):
                tmp_Iu = table.cell_value(num,GTP_top_col)-table.cell_value(num,Iu_top_col)
                if tmp_Iu > num_max:
                    num_max = tmp_Iu
                else:
                    continue
            else:
                continue
        num_max = round(float(num_max)/1024/1024*8,2)
        top_value_list.append(num_max)

    #print top_value_list
    return top_value_list

def find_Gb_User_top (table,device_list,Gb_User_col,device_col):
    top_value_list = []
    for dev in device_list:
        num_max=0
        for num in range(table.nrows):
            if dev == table.cell_value(num, device_col):
                if table.cell_value(num, Gb_User_col) > num_max:
                    num_max = table.cell_value(num, Gb_User_col)
                else:
                    continue
            else:
                continue
        num_max = round(float(num_max)/10000, 2)
        top_value_list.append(num_max)

    #print top_value_list
    return top_value_list


#################################
###########整列函数添加##########
#################################
def add_excel_col(sheet,content,row,style,mod):
    for num in range(len(content)):
        sheet.write(num+mod,row,content[num],xlwt.easyxf(style))

#def add_excel_col(sheet,content,row,style):
    #for num in range(len(content)):
        #sheet.write(num+12,row,content[num],xlwt.easyxf(style))

#def add_excel_col(sheet,content,row,style):
    #for num in range(len(content)):
        #sheet.write(num+23,row,content[num],xlwt.easyxf(style))

#################################
############初始化excel##########
#################################
def init_excel(book,GGSN_device_list,SGSN_device_list,style_tmp):
    sheet = book.add_sheet(u'资源预警')
    add_excel_col(sheet,GGSN_device_list,0,style_tmp,1)
    add_excel_col(sheet,GGSN_device_list,0,style_tmp,12)
    add_excel_col(sheet,SGSN_device_list,0,style_tmp,23)
    add_excel_col(sheet,SGSN_device_list,0,style_tmp,33)
    add_excel_col(sheet,SGSN_device_list,0,style_tmp,43)
    sheet.write(0,1,u'Gn口物理带宽')
    sheet.write(0,2,u'Gn口实际流量')
    sheet.write(0,3,u'利用率')
    sheet.write(0,4,u'Gi口物理带宽')
    sheet.write(0,5,u'Gi口实际流量')
    sheet.write(0,6,u'利用率')
    sheet.write(0,7,u'S1-U物理带宽（G）')
    sheet.write(0,8,u'S1-U实际流量（G）')
    sheet.write(0,9,u'利用率')
    sheet.write(11,1,u'正式/临时licesne吞吐量')
    sheet.write(11,2,u'实际吞吐量（M）')
    sheet.write(11,3,u'利用率')
    sheet.write(11,4,u'正式/临时licesne激活用户数')
    sheet.write(11,5,u'实际激活用户数(万)')
    sheet.write(11,6,u'利用率')
    sheet.write(11,10,u'临时license过期时间')
    sheet.write(11,7,u'正式/临时licesne激活用户数（4G）')
    sheet.write(11,8,u'实际激活用户数(万)')
    sheet.write(11,9,u'利用率')
    sheet.write(12,10,u'不涉及')
    sheet.write(13,10,u'不涉及')
    sheet.write(14,10,u'不涉及')
    sheet.write(15,10,u'不涉及')
    sheet.write(16,10,u'不涉及')
    sheet.write(17,10,u'不涉及')
    sheet.write(18,10,u'不涉及')
    sheet.write(19,10,u'不涉及')
    sheet.write(20,10,u'不涉及')
    sheet.write(21,10,u'不涉及')
    sheet.write(22,1,u'Gn口带宽')
    sheet.write(22,2,u'Gn口实际流量（G）')
    sheet.write(22,3,u'利用率')
    sheet.write(22,4,u'Iu口带宽')
    sheet.write(22,5,u'Iu口实际流量')
    sheet.write(22,6,u'利用率')
    sheet.write(32,1,u'2G License容量')
    sheet.write(32,2,u'2G现网附着用户数(万)')
    sheet.write(32,3,u'2G用户license利用率')
    sheet.write(32,4,u'3G licsnse容量')
    sheet.write(32,5,u'3G现网附着用户数(万)')
    sheet.write(32,6,u'3G用户license利用率')
    sheet.write(32,7,u'4G licsnse容量(万)')
    sheet.write(32,8,u'4G现网附着用户数(万)')
    sheet.write(32,9,u'4G用户license利用率')
    sheet.write(42,1,u'2G License容量(万)')
    sheet.write(42,2,u'2G现网激活用户数(万)')
    sheet.write(42,3,u'利用率')
    sheet.write(42,4,u'3G licsnse容量(万)')
    sheet.write(42,5,u'3G现网激活用户数(万)')
    sheet.write(42,6,u'利用率')
    sheet.write(42,7,u'4G licsnse容量(万)')
    sheet.write(42,8,u'4G承载建立数(万)')
    sheet.write(42,9,u'利用率')

    save_excel(book)
    return book


def main():

    pathlist=os.listdir('.')
    for name in pathlist:
        if 'GGSN' in name:
            GGSN = name
        elif 'SGSN' in name:
            SGSN = name
        elif u'附着用户数'.encode('gbk') in name:
            attach_user_num_xls = name
        else:
            continue
    data_GGSN = read_excel(GGSN)
    data_SGSN = read_excel(SGSN)
    data_attach = read_excel(attach_user_num_xls)
    table_GGSN = data_GGSN.sheet_by_index(0)
    table_SGSN = data_SGSN.sheet_by_index(0)
    table_attach = data_attach.sheet_by_index(0)
    device_col = 2
    Gn_top_col = 4
    Gi_top_col = 21
    S1U_top_col = 23
    PDP_4G_top_col = 20
    PDP_top_col = 8
    GTP_top_col = 17
    Iu_top_col = 23
    Gb_User_Activate_col = 6
    Gb_User_Attach_col = 4
    Iu_User_Attach_col = 5
    Iu_User_Activate_col = 12
    S1_User_Attach_col = 6
    Bear_User_col = 25

   # print "%d" %table_GGSN.nrows
   # col_number=table_GGSN.col_values(2)
    #GGSN_device_list=detect_device_name(table_GGSN,device_col)
    #del GGSN_device_list[0:2]      #delet the first two useless data
    #GGSN_device_list.sort()


    Gn_flow = [u'6',u'4',u'6',u'20',u'20',u'20',u'20',u'20',u'20',u'20']
    Gi_flow = [10,4,10,10,4,10,10,10,10,10]
    S1U_flow=[u'4',u'/',u'4',u'4',u'/',u'4',u'4',u'4',u'10',u'10']
    License_in_out = [u'5437M',u'5437M',u'5437M',u'16800M',u'16800M',u'12084M',u'21580M',u'21580M',u'21000M',u'21000M']
    License_Activate = [u'66万',u'100万',u'66万',u'66万',u'70万',u'66万',u'66万',u'66万',u'66万',u'66万']
    License_Activate_4G = [u'45万',u'/',u'45万',u'45万',u'5万',u'45万',u'45万',u'45万',u'45万',u'45万']
    Gn_band = [u'20G', u'16G', u'18G', u'14G', u'16G', u'16G', u'14G', u'16G']
    Iu_band = [u'20G', u'16G', u'18G', u'14G', u'16G', u'16G', u'14G', u'16G']
    User_License_2G = [u'110万',u'115万',u'132万',u'90万',u'105万',u'90万',u'105万',u'90万']
    User_License_3G = [u'90万',u'90万',u'132万',u'90万',u'100万',u'90万',u'105万',u'90万']
    User_License_4G = [u'45万',u'45万',u'45万',u'45万',u'45万',u'45万',u'45万',u'45万']
    User_Activate_2G = [u'77万',u'80万',u'77万',u'63万',u'73万',u'63万',u'73万',u'63万']
    User_Activate_3G = [u'63万',u'63万',u'116万',u'63万',u'70万',u'63万',u'73万',u'63万']
    Bear_4G = [u'49万',u'49万',u'49万',u'49万',u'49万',u'49万',u'49万',u'49万']
    Gn_actual_flow = find_Gn_top(table_GGSN,GGSN_device_list,Gn_top_col,device_col)
    SGi_actual_flow = find_Gi_top(table_GGSN,GGSN_device_list,Gi_top_col,device_col)
    S1U_actual_flow = find_Gi_top(table_GGSN,GGSN_device_list,S1U_top_col,device_col)
    PDP_Activate = find_PDP_top(table_GGSN,GGSN_device_list,PDP_top_col,device_col)
    PDP_4G_Activate = find_PDP_top(table_GGSN,GGSN_device_list,PDP_4G_top_col,device_col)
    GTP = find_GTP_top(table_SGSN,SGSN_device_list,GTP_top_col,device_col)
    Iu = find_Iu_top(table_SGSN,SGSN_device_list,Iu_top_col,device_col,GTP_top_col)
    #Gb_User = find_Gb_User_top(table_SGSN,SGSN_device_list,Gb_User_col,device_col)
    Gb_User_Attach = find_Gb_User_top(table_attach,SGSN_device_list,Gb_User_Attach_col,device_col)
    Iu_User_Attach = find_Gb_User_top(table_attach,SGSN_device_list,Iu_User_Attach_col,device_col)
    S1_User_Attach = find_Gb_User_top(table_attach,SGSN_device_list,S1_User_Attach_col,device_col)
    Gb_User_Activate = find_Gb_User_top(table_SGSN,SGSN_device_list,Gb_User_Activate_col,device_col)
    Iu_User_Activate = find_Gb_User_top(table_SGSN,SGSN_device_list,Iu_User_Activate_col,device_col)
    Bear_User = find_Gb_User_top(table_SGSN,SGSN_device_list,Bear_User_col,device_col)
    Actual_lisence = [int(i*1024) for i in Gn_actual_flow]
    #print Iu_User_Activate
    Gn_flow_int = [6,4,6,20,20,20,20,20,20,20]
    Gi_flow_int = [10,4,10,10,4,10,10,10,10,10]
    S1U_flow_int = [4,0,4,4,0,4,4,4,10,10]
    License_in_out_int = [5437,5437,5437,16800,16800,12084,21580,21580,21000,21000]
    License_Activate_int = [66,100,66,66,70,66,66,66,66,66]
    License_Activate_int_4G = [45,0,45,45,5,45,45,45,45,45]
    Gn_band_int = [20,16,18,14,16,16,14,16]
    Iu_band_int = [20,16,18,14,16,16,14,16]
    User_License_2G_int= [110,115,132,90,105,90,105,90]
    User_License_3G_int = [90,90,132,90,100,90,105,90]
    User_License_4G_int = [45,45,45,45,45,45,45,45]
    User_Activate_2G_int = [77,80,77,63,73,63,73,63]
    User_Activate_3G_int = [63,63,116,63,70,63,73,63]
    Bear_4G_int = [49,49,49,49,49,49,49,49]
    Gn_flow_per = []
    Gi_flow_per = []
    S1U_flow_per = []
    Lisence_per = []
    PDP_per = []
    PDP_4G_per=[]
    GTP_per = []
    Iu_per = []
    User_License_per = []
    User_License_3G_per = []
    User_License_4G_per = []
    User_Activate_2G_per = []
    User_Activate_3G_per =[]
    Bear_User_per=[]
    Gi_SGi=[]
    for n in range(len(Gn_flow)):
        tmp_Gn = str(round(Gn_actual_flow[n]/Gn_flow_int[n],4)*100)+'%'
        Gn_flow_per.append(tmp_Gn)
        Gi_SGi.append(SGi_actual_flow[n]+Gn_actual_flow[n])
        tmp_Gi = str(round(Gi_SGi[n]/Gi_flow_int[n],4)*100)+'%'
        Gi_flow_per.append(tmp_Gi)
        if S1U_flow_int[n]!=0:
            tmp_S1U = str(round(S1U_actual_flow[n]/S1U_flow_int[n],4)*100)+'%'
            S1U_flow_per.append(tmp_S1U)
        else:
            S1U_flow_per.append('/')



    for n in range(len(License_in_out)):
        tmp_Lisence = str(round(float(Actual_lisence[n])/License_in_out_int[n],4)*100)+'%'
        Lisence_per.append(tmp_Lisence)
    for n in range(len(License_Activate_int)):
        tmp_PDP = str(round(float(PDP_Activate[n])/License_Activate_int[n],4)*100)+'%'
        if License_Activate_int_4G[n]!=0:
            tmp_PDP_4G = str(round(float(PDP_4G_Activate[n])/License_Activate_int_4G[n],4)*100)+'%'
        else:
            tmp_PDP_4G='/'
        #tmp = float(PDP_Activate[n])/License_Activate_int[n]
        #tmp = round (tmp,4)
        #round 函数少当最后一位是0 会少一位显示
        PDP_per.append(tmp_PDP)
        PDP_4G_per.append(tmp_PDP_4G)
    for n in range(len(Gn_band_int)):
        tmp_GTP = str(round(float(GTP[n])/Gn_band_int[n],4)*100)+'%'
        GTP_per.append(tmp_GTP)
    for n in range(len(Iu_band_int)):
        tmp_Iu = str(round(float(Iu[n])/Iu_band_int[n],4)*100)+'%'
        Iu_per.append(tmp_Iu)
    for n in range(len(User_License_2G_int)):
        tmp_Gb = str(round(float(Gb_User_Attach[n])/User_License_2G_int[n],4)*100)+'%'
        User_License_per.append(tmp_Gb)

    for n in range(len(User_License_3G_int)):
        tmp_Iu_attache = str(round(float(Iu_User_Attach[n])/User_License_3G_int[n],4)*100)+'%'
        User_License_3G_per.append(tmp_Iu_attache)

    for n in range(len(User_License_4G_int)):
        tmp_S1_attache = str(round(float(S1_User_Attach[n])/User_License_4G_int[n],4)*100)+'%'
        User_License_4G_per.append(tmp_S1_attache)

    for n in range(len(User_Activate_2G_int)):
        tmp_activate = str(round(float(Gb_User_Activate[n])/User_Activate_2G_int[n],4)*100)+'%'
        User_Activate_2G_per.append(tmp_activate)

    for n in range(len(User_Activate_3G_int)):
        tmp_activate_3G = str(round(float(Iu_User_Activate[n])/User_Activate_3G_int[n],4)*100)+'%'
        User_Activate_3G_per.append(tmp_activate_3G)

    for n in range(len(Bear_4G_int)):
        tmp_Bear_User = str(round(float(Bear_User[n])/Bear_4G_int[n],4)*100)+'%'
        Bear_User_per.append(tmp_Bear_User)

    #print GGSN_device_list
    #for i in GGSN_device_list:
        #print i.encode("utf-8")
    book = open_excel()
    book = init_excel(book,GGSN_device_list,SGSN_device_list,style_tmp)
    sheet=book.get_sheet(0)
    add_excel_col(sheet, Gn_flow,1,style_tmp,1)
    add_excel_col(sheet, Gn_actual_flow,2,style_tmp,1)
    add_excel_col(sheet, Gn_flow_per,3,style_tmp,1)
    add_excel_col(sheet, Gi_flow,4,style_tmp,1)
    add_excel_col(sheet, Gi_SGi,5,style_tmp,1)
    add_excel_col(sheet, Gi_flow_per,6,style_tmp,1)
    add_excel_col(sheet, S1U_flow,7,style_tmp,1)
    add_excel_col(sheet, S1U_actual_flow,8,style_tmp,1)
    add_excel_col(sheet, S1U_flow_per,9,style_tmp,1)
    add_excel_col(sheet, License_in_out,1,style_tmp,12)
    add_excel_col(sheet, Actual_lisence,2,style_tmp,12)
    add_excel_col(sheet, Lisence_per,3,style_tmp,12)
    add_excel_col(sheet, License_Activate,4,style_tmp,12)
    add_excel_col(sheet, PDP_Activate,5,style_tmp,12)
    add_excel_col(sheet, PDP_per,6,style_tmp,12)
    add_excel_col(sheet, License_Activate_4G,7,style_tmp,12)
    add_excel_col(sheet, PDP_4G_Activate,8,style_tmp,12)
    add_excel_col(sheet, PDP_4G_per,9,style_tmp,12)
    add_excel_col(sheet, Gn_band,1,style_tmp,23)
    add_excel_col(sheet, GTP,2,style_tmp,23)
    add_excel_col(sheet, GTP_per,3,style_tmp,23)
    add_excel_col(sheet, Iu_band,4,style_tmp,23)
    add_excel_col(sheet, Iu,5,style_tmp,23)
    add_excel_col(sheet, Iu_per,6,style_tmp,23)
    add_excel_col(sheet, User_License_2G,1,style_tmp,33)
    add_excel_col(sheet, Gb_User_Attach,2,style_tmp,33)
    add_excel_col(sheet, User_License_per,3,style_tmp,33)
    add_excel_col(sheet, User_License_3G,4,style_tmp,33)
    add_excel_col(sheet, Iu_User_Attach,5,style_tmp,33)
    add_excel_col(sheet, User_License_3G_per,6,style_tmp,33)
    add_excel_col(sheet, User_License_4G,7,style_tmp,33)
    add_excel_col(sheet, S1_User_Attach,8,style_tmp,33)
    add_excel_col(sheet, User_License_4G_per,9,style_tmp,33)
    add_excel_col(sheet, User_Activate_2G,1,style_tmp,43)
    add_excel_col(sheet, Gb_User_Activate,2,style_tmp,43)
    add_excel_col(sheet, User_Activate_2G_per,3,style_tmp,43)
    add_excel_col(sheet, User_Activate_3G,4,style_tmp,43)
    add_excel_col(sheet, Iu_User_Activate,5,style_tmp,43)
    add_excel_col(sheet, User_Activate_3G_per,6,style_tmp,43)
    add_excel_col(sheet, Bear_4G,7,style_tmp,43)
    add_excel_col(sheet, Bear_User,8,style_tmp,43)
    add_excel_col(sheet, Bear_User_per,9,style_tmp,43)
    save_excel(book)
    #print os.listdir('./')
    #print("finish")
if __name__=="__main__":
    main()
