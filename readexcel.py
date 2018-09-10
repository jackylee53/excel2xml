#!/usr/bin/python
# -*- coding: utf-8 -*-

import random
import time
import datetime
from lxml import etree
from xlrd import xldate_as_tuple
from souppost import send_soap

# reload(sys)
# sys.setdefaultencoding('utf-8')

excelFileUrl = "/Users/jacky/Downloads/20180626新导入内容/媒体资源总6-21 - 技术版.xlsx"

xmlFileDic = "/Users/jacky/Downloads/XML/Program/"

ftpUser = 'ftp://ustorage:usxwzx@61.130.250.36/'

from xlrd import open_workbook

"""打开Excel表"""
wb = open_workbook(excelFileUrl, encoding_override='utf-8')

"""获取第N张sheet表"""
sheet_0 = wb.sheets()[5]

"""总行数和总列数"""
total_rows = sheet_0.nrows
total_cols = sheet_0.ncols

# 注入的动作
Action = "REGIST"

"""根据ProgramID,返回固定长度的数字值"""
def returnFixLen(pid):
    id = pid.split('/')[1].split('@')[0]
    if len(pid) > 32:
        modid = len(pid) - 31
        resid = 765432 + int(id[0:-modid])
    else:
        resid = 123456 + int(id)
    return resid

"""将ProgramID由不确定长度处理为32位"""
def returnProgramID(pid):
    print("错误PID", pid)
    id = pid.split('/')[1].split('@')[0]
    pidheadstr = pid.split('/')[0]
    if pidheadstr == "Umai:PRO":
        headstr = "Umai:PRO/"
    elif pidheadstr == "Umai:PROG":
        headstr = "Umai:PROG/"
    pidendstr = pid.split('@')[1]
    endstr = '@' + pidendstr
    if len(pid) > 32:
        modid = len(pid) - 31
        resid = id[0:-modid]
    else:
        resid = id
    ProgramID = headstr + resid + endstr
    return ProgramID

"""根据ProgramID生成MovieID"""
def rerurnMoiveID(pid):
    pidheadstr = pid.split('/')[0]
    if pidheadstr == "Umai:PRO":
        headstr = "Umai:MOV/"
    elif pidheadstr == "Umai:PROG":
        headstr = "Umai:MOVI/"
    pidendstr = pid.split('@')[1]
    endstr = '@' + pidendstr
    MovieID = headstr + str(returnFixLen(pid)) + endstr
    print(MovieID)
    return MovieID

"""根据ProgramID生成PictureID"""
def rerurnPictureID(pid, fixid):
    pidheadstr = pid.split('/')[0]
    if pidheadstr == "Umai:PRO":
        headstr = "Umai:PIC/"
    elif pidheadstr == "Umai:PROG":
        headstr = "Umai:PICT/"
    pidendstr = pid.split('@')[1]
    endstr = '@' + pidendstr
    PictureID = headstr + str(fixid) + endstr
    return PictureID

"""返回特定资产类型"""
def returnCType(cell, ctype):
    if ctype == 2 and cell % 1 == 0:  # 如果是整形
        cell = int(cell)
    elif ctype == 3:
        # 转成datetime对象
        date = datetime(*xldate_as_tuple(cell, 0))
        cell = date.strftime('%Y/%d/%m %H:%M:%S')
    elif ctype == 4:
        cell = True if cell == 1 else False
    return cell


def addProperty(parent,
                name, value, add_cdate=True):
    property = etree.SubElement(parent, "Property")
    property.set('Name', name)
    if isinstance(value, int):
        value = str(value)
    if value:
        if add_cdate:
            cdata = etree.CDATA(str(value))
            property.text = cdata
        else:
            property.text = str(value)
    return property


def generatePics(Objects, pics, pid, isspic=False):
    pic_ids = []
    i = 1
    if len(pics) > 0:
        for pic in pics.splitlines():
            pic_element = etree.SubElement(Objects, "Object")
            obj_picfixid = returnFixLen(pid)
            if isspic:
                pic_id = rerurnPictureID(pid, obj_picfixid + 5 + i)
            elif len(pics) == 1:
                pic_id = rerurnPictureID(pid, obj_picfixid + 2)
            else:
                pic_id = rerurnPictureID(pid, obj_picfixid + i)
            pic_element.set('ElementType', 'Picture')
            pic_element.set('Action', Action)
            pic_element.set('ID', pic_id)
            pic_element.set('Code', pic_id)
            addProperty(pic_element, "FileURL", ftpUser + pic.strip())
            addProperty(pic_element, "Description", None, False)
            pic_ids.append(pic_id)
            i = i + 1
    return pic_ids


def main():
    for i in range(total_rows - 1):
        i = i + 1
        """Excel中序号ID"""
        seq_id = sheet_0.cell(i, 0).value
        print("seq_id", seq_id)
        """Type节目的分类标签，如“体育”，多个标签用空格或“;”区分"""
        Genre = sheet_0.cell(i, 1).value
        """Keywords"""
        Keywords = sheet_0.cell(i, 2).value
        """导演(Director)、主持人(Compere)均用该字断"""
        Director = sheet_0.cell(i, 4).value
        """资产名称"""
        Name = sheet_0.cell(i, 5).value
        """ProgramID,生成TID,通过调用方法生成MovieID和PictureID"""
        rdpid = random.randint(1, 100000000)
        PID = sheet_0.cell(i, 7).value
        print(PID)
        if len(PID) == 0:
            PID = 'Umai:PROG/' + '20180621' + str(162717250056 + rdpid) + '@BST'
        ProgramID = returnProgramID(PID)
        """FileURL"""
        PlayUrl = sheet_0.cell(i, 8).value
        """Duration"""
        Duration = int(sheet_0.cell(i, 9).value)
        """Description"""
        Description = sheet_0.cell(i, 11).value
        pics = sheet_0.cell(i, 16).value
        # spics = sheet_0.cell(i, 17).value
        """LicensingWindowStart,YYYYMMDDHH24MiSS"""
        LicensingWindowStart = '20161124011001'
        # 发行时间
        ReleaseYear = '2018'

        # Program属性
        """LicensingWindowEnd"""
        LicensingWindowEnd = '20191124011001'
        """Tags"""
        # Tags = Keywords

        # 0: 影片,1: 单集
        SeriesFlag = 0
        # 原产地
        OriginalCountry = "中国大陆"
        # 语言
        Language = "国语"
        # 拷贝保护标志
        Macrovision = 0
        # 列表定价
        PriceTaxIn = 0
        # 状态标志0:失效 1:生效
        Status = 0
        # 1: 视频类节目，2: 图文类节目
        SourceType = 1
        StorageType = 1
        DefinitionFlag = 0

        # Moive属性定义
        SourceDRMType = 0
        DestDRMType = 0
        AudioType = 0
        # 0: 4x3，1: 16x9(Wide)
        ScreenFormat = 1
        # 是否有字幕
        ClosedCaptioning = 1
        # 文件大小
        FileSize = 10000
        # 分辨率
        BitRateType = 6
        VideoType = 1
        VideoProfile = 4
        SystemLayer = 1

        # 图片属性定义：0: 缩略图,1: 海报,2: 剧照,3: 图标,4: 标题图,5: 广告图,6: 草图,7: 背景图,9: 频道图片,10: 频道黑白图片,11: 频道Logo,12: 频道名字图片,99: 其他
        PicType0 = 0
        PicType1 = 1
        PicType2 = 2
        PicType3 = 3
        PicType4 = 4
        PicType5 = 5
        PicType6 = 6
        PicType7 = 7
        PicType8 = 8
        PicType9 = 9
        PicOtherType = 99

        root = etree.Element("ADI", nsmap={'xsi': 'http://www.w3.org/2001/XMLSchema-instance'})
        root.set('BizDomain', '0')
        root.set('Priority', '5')

        Objects = etree.SubElement(root, "Objects")

        program = etree.SubElement(Objects, "Object")
        program.set('ElementType', 'Program')
        program.set('Action', Action)
        program.set('ID', ProgramID)
        program.set('Code', ProgramID)

        addProperty(program, "Name", Name)
        addProperty(program, "OriginalName", Name)
        addProperty(program, "SortName", Name)
        addProperty(program, "SearchName", Name)
        addProperty(program, "Genre", Genre)
        addProperty(program, "ActorDisplay", None, False)
        addProperty(program, "WriterDisplay", None, False)
        addProperty(program, "OriginalCountry", OriginalCountry)
        addProperty(program, "Language", Language)
        addProperty(program, "ReleaseYear", ReleaseYear)
        addProperty(program, "OrgAirDate", None, False)

        addProperty(program, "LicensingWindowStart", LicensingWindowStart, False)
        addProperty(program, "LicensingWindowEnd", LicensingWindowEnd, False)
        addProperty(program, "DisplayAsNew", 7)
        addProperty(program, "DisplayAsLastChance", None, False)

        addProperty(program, "Macrovision", Macrovision, False)
        addProperty(program, "Description", Description)
        addProperty(program, "PriceTaxIn", PriceTaxIn, False)
        addProperty(program, "Status", Status, False)
        addProperty(program, "SourceType", SourceType, False)

        addProperty(program, "SeriesFlag", SeriesFlag, False)
        addProperty(program, "Type", Genre, False)
        addProperty(program, "Keywords", Keywords)
        addProperty(program, "Director", Director)
        addProperty(program, "Compere", Director)
        addProperty(program, "Tags", Genre, False)
        addProperty(program, "OrderNumber", None, False)
        addProperty(program, "Duration", Duration, False)
        addProperty(program, "StorageType", StorageType, False)

        addProperty(program, "RMediaCode", ProgramID, False)
        addProperty(program, "DefinitionFlag", DefinitionFlag, False)

        movie = etree.SubElement(Objects, "Object")
        movie.set('ElementType', 'Movie')
        movie.set('Action', Action)
        movie_id = rerurnMoiveID(PID)
        movie.set('ID', movie_id)
        movie.set('Code', movie_id)

        addProperty(movie, "Type", 1, False)
        addProperty(movie, "FileURL", PlayUrl)
        addProperty(movie, "SourceDRMType", SourceDRMType, False)
        addProperty(movie, "DestDRMType", DestDRMType, False)
        addProperty(movie, "AudioType", AudioType, False)
        addProperty(movie, "ScreenFormat", ScreenFormat, False)
        addProperty(movie, "ClosedCaptioning", ClosedCaptioning, False)
        addProperty(movie, "OCSURL", None, False)
        addProperty(movie, "Duration", Duration, False)
        addProperty(movie, "FileSize", FileSize, False)
        addProperty(movie, "BitRateType", BitRateType, False)
        addProperty(movie, "VideoType", VideoType, False)
        addProperty(movie, "Resolution", None, False)
        addProperty(movie, "VideoProfile", VideoProfile, False)
        addProperty(movie, "SystemLayer", SystemLayer, False)
        addProperty(movie, "AudioFormat", None, False)
        addProperty(movie, "ServiceType", None, False)

        pic_ids = generatePics(Objects, pics, PID)

        mappings = etree.SubElement(root, "Mappings")
        movie_mapping = etree.SubElement(mappings, "Mapping")
        movie_mapping.set('ParentType', 'Program')
        movie_mapping.set('ParentID', ProgramID)
        movie_mapping.set('ParentCode', ProgramID)
        movie_mapping.set('ElementType', 'Movie')
        movie_mapping.set('ElementID', movie_id)
        movie_mapping.set('ElementCode', movie_id)
        movie_mapping.set('Action', Action)

        if len(pic_ids) > 0:
            i = 1
            for pic_id in pic_ids:
                pic_mapping = etree.SubElement(mappings, "Mapping")
                pic_mapping.set('ParentType', 'Picture')
                pic_mapping.set('ParentID', pic_id)
                pic_mapping.set('ParentCode', pic_id)
                pic_mapping.set('ElementType', 'Program')
                pic_mapping.set('ElementID', ProgramID)
                pic_mapping.set('ElementCode', ProgramID)
                pic_mapping.set('Action', Action)
                if len(pic_ids) == 1:
                    addProperty(pic_mapping, "Type", 2, False)
                else:
                    addProperty(pic_mapping, "Type", i, False)
                i = i + 1

        # spic_ids = generatePics(Objects, spics, PID, isspic=True)
        #
        # if len(spic_ids) > 0:
        #     for spic_id in spic_ids:
        #         pic_mapping = etree.SubElement(mappings, "Mapping")
        #         pic_mapping.set('ParentType', 'Picture')
        #         pic_mapping.set('ParentID', spic_id)
        #         pic_mapping.set('ParentCode', spic_id)
        #         pic_mapping.set('ElementType', 'Program')
        #         pic_mapping.set('ElementID', ProgramID)
        #         pic_mapping.set('ElementCode', ProgramID)
        #         pic_mapping.set('Action', Action)
        #         addProperty(pic_mapping, "Type", 0, False)

        ########### 将DOM对象doc写入文件
        tree = etree.ElementTree(root)
        file_name = str(seq_id) + ".xml"
        print(etree.tostring(root, pretty_print=True))
        writefile_name = xmlFileDic + file_name
        tree.write(writefile_name, pretty_print=True, xml_declaration=True, encoding='utf-8')

        # sequnceid = random.randint(1, 100000000)
        # resp = send_soap(str(sequnceid), file_name)
        # print(resp)
        # time.sleep(1)

if __name__ == "__main__":
    main()
