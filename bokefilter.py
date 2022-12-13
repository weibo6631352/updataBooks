
import xlrd
import xlwt
import sys
import os
import re
from dateutil.parser import parse  # 日期解析器
from xlutils.copy import copy

sourceWrokBookPath = "./source/新图书信息表 - 副本.xlsx"
templateWrokBookPath = "./template/1670639422947template.xls"
targetWrokBookPath = './王凤歌 12.13.xls'
deleteWrokBookPath = './王凤歌 12.13删除书籍.xls'
modifyWrokBookPath = './王凤歌 12.13缺信息或修改的书籍.xls'


sourceStartRow = 11106 # 正常index是 11111  因为前面几本是一类所以包含进来吧  王凤歌部分
sourceStopRow = 13982 # 正常index是 13888  因为后面几本是一类所以包含进来吧  王凤歌部分

#sourceStartRow = 1 # 全部
#sourceStopRow = 72415 # 全部
templateStartRow = 2


warnBookNameList = ['血腥', '仙', '佛', '鬼', '妖', '魔', '神', '斗罗大陆', '暴力', '恐怖',  '玄幻',  '宗教',  '基督',  '耶稣',  '观音',  '菩萨',  '张三丰',  '玄奘', 
 '圣僧',  '姜子牙',  '钟馗',  '济公',  '托钵 在西藏行走',  '灵魂切成的碎片'
,  '都市鸡人',  '哈利路亚',  '聊斋',  '幻影迷踪',  '轴心国做恶悍将', '缔造者计划', '王子13号店'
,  '0号标本',  '抓痕',  '伤藤',  '缔造者计划',  '潮汐', '在自由的边缘行走']
warnAuthorNameList = ['杨红樱',  '北猫',  '龙应台']

sourceWrokBook = xlrd.open_workbook(sourceWrokBookPath)
sourceSheet = sourceWrokBook.sheet_by_name('Sheet1')   

templateWrokBook = xlrd.open_workbook(templateWrokBookPath, formatting_info=True)

targetWrokBook = copy(templateWrokBook)  
targetSheet = targetWrokBook.get_sheet(0) 

deleteWrokBook = copy(templateWrokBook)  
deleteSheet = deleteWrokBook.get_sheet(0) 

modifyWrokBook = copy(templateWrokBook)  
modifySheet = modifyWrokBook.get_sheet(0) 

targetBookInfos = {}
deleteBookInfos = {}
modifyBookInfos = {}

def getBookInfos(sourceSheet):
    bookInfos = {}
    for srouceRow in range(sourceStartRow, sourceStopRow+1):
        bookInfo = {}
        bookInfo['ISBN'] = str(sourceSheet.cell_value(srouceRow, 0))
        bookInfo['书名'] = str(sourceSheet.cell_value(srouceRow, 1))
        bookInfo['著者'] = str(sourceSheet.cell_value(srouceRow, 2))
        bookInfo['出版社'] = str(sourceSheet.cell_value(srouceRow, 3))
        bookInfo['出版日期'] = str(sourceSheet.cell_value(srouceRow, 4))
        bookInfo['价格'] = str(sourceSheet.cell_value(srouceRow, 5))
        bookInfo['分类号'] = str(sourceSheet.cell_value(srouceRow, 6))
        bookInfo['条形码'] = str(sourceSheet.cell_value(srouceRow, 7))
        bookInfo['馆藏地'] = str(sourceSheet.cell_value(srouceRow, 8))
        bookInfo['所在层架'] = str(sourceSheet.cell_value(srouceRow, 9))
        bookInfo['备注'] = ''
        bookInfo['原始档案行数'] = srouceRow
        bookInfos[srouceRow] = bookInfo
    return bookInfos

def deleteByNameSame(bookInfos):
    retBookInfos = {}
    midBookInfos = {}
    for bookKey in bookInfos:
        bookInfo = bookInfos[bookKey]
        curBooknameAndCurConcern = bookInfo['书名'] +  bookInfo['出版社'] 
        if curBooknameAndCurConcern not in midBookInfos:
            midBookInfos[curBooknameAndCurConcern] = []
        midBookInfos[curBooknameAndCurConcern].append(bookInfo)

    for bookInfoListIndex in midBookInfos:
        bookInfoList = midBookInfos[bookInfoListIndex]
        bookCount = 1
        for bookInfo in bookInfoList:
            if(bookCount <= 5):
                retBookInfos[bookInfo['原始档案行数']] = bookInfo.copy()
            else:
                deleteBookInfos[bookInfo['原始档案行数']] = bookInfo.copy()
                deleteBookInfos[bookInfo['原始档案行数']] ['备注']+= '书名和出版社名字完全相同册数大于5,当前第' +str(bookCount) + '本'
            bookCount+=1
    return retBookInfos

def deleteByBookNameWarn(bookInfos):
    retBookInfos = {}
    for bookKey in bookInfos:
        bookInfo = bookInfos[bookKey]
        bookname = bookInfo['书名'] 

        nowarning = True
        warningName = ''
        for warningField in warnBookNameList:
            if warningField in bookname:
                nowarning = False
                warningName = warningField
                break;
        if nowarning:
            retBookInfos[bookKey] = bookInfo.copy()
        else:
            deleteBookInfos[bookKey] = bookInfo.copy()
            deleteBookInfos[bookKey]['备注'] += ('书名包含违规关键词:' + warningName)
    return retBookInfos

def deleteByAuthorNameWarn(bookInfos):
    retBookInfos = {}
    for bookKey in bookInfos:
        bookInfo = bookInfos[bookKey]
        authorname = bookInfo['著者'] 

        nowarning = True
        warningName = ''
        for warningField in warnAuthorNameList:
            if warningField in authorname:
                nowarning = False
                warningName = warningField
                break;
        if nowarning:
            retBookInfos[bookKey] = bookInfo.copy()
        else:
            deleteBookInfos[bookKey] = bookInfo.copy()
            deleteBookInfos[bookKey]['备注'] += ('著者包含违规关键词:' + warningName)
    return retBookInfos

def deleteByISBNWarn(bookInfos):
    retBookInfos = {}
    for bookKey in bookInfos:
        bookInfo = bookInfos[bookKey]
        isbnName = bookInfo['ISBN'] 
        if(len(isbnName) == len('9787206075698')):
            retBookInfos[bookKey] = bookInfo.copy()
        else:
            deleteBookInfos[bookKey] = bookInfo.copy()
            deleteBookInfos[bookKey]['备注'] += ('ISBN异常:' + isbnName)
    return retBookInfos

def modifyDate(bookInfos):
    for bookKey in bookInfos:
        bookInfo = bookInfos[bookKey]
        sourceDate = bookInfo['出版日期']
        pattern = re.compile('^[0-9]{4}-[0-9]{2}$')
        if pattern.match(sourceDate):   #无误
            continue

        modifyDateStr = sourceDate

        modifyDateStr=re.sub('\(.*?\)','',modifyDateStr)
        modifyDateStr=re.sub('\（.*?\）','',modifyDateStr)
        modifyDateStr=re.sub('\[.*?\]','',modifyDateStr)
        modifyDateStr=re.sub('\【.*?\】','',modifyDateStr)
        modifyDateStr=re.sub(',.*?重印','',modifyDateStr)
        modifyDateStr=re.sub(',.*?印\)','',modifyDateStr)
        modifyDateStr=re.sub('第.*?版','',modifyDateStr)

        modifyDateStr = modifyDateStr.replace('，' , '-')
        modifyDateStr = modifyDateStr.replace(',' , '-')
        modifyDateStr = modifyDateStr.replace('.' , '-')
        modifyDateStr = modifyDateStr.replace('。' , '-')
        modifyDateStr = modifyDateStr.replace('·' , '-')
        modifyDateStr = modifyDateStr.replace('年' , '-')
        modifyDateStr = modifyDateStr.replace('月' , '-')
        modifyDateStr = modifyDateStr.replace('日' , ' ')
        modifyDateStr = modifyDateStr.replace('s' , ' ')
        modifyDateStr = modifyDateStr.replace('）' , '')
        modifyDateStr = modifyDateStr.replace(')' , '')
        modifyDateStr = modifyDateStr.strip()
        modifyDateStr = modifyDateStr.strip('-')
        try:
            date = parse(modifyDateStr, fuzzy=True)
        except:
            date = None
        if date == None:
            modifyBookInfos[bookKey] = bookInfo.copy()
            modifyBookInfos[bookKey]['备注'] += ('    无效的日期格式:' + sourceDate)
            bookInfo['备注'] += ('    无效的日期格式:' + sourceDate)
        else:
            bookInfo['出版日期'] = parse(modifyDateStr).strftime('%Y-%m')
            modifyBookInfos[bookKey] = bookInfo.copy()
            modifyBookInfos[bookKey]['备注'] += ('    修正日期格式，源日期为:' + sourceDate)
            bookInfo['备注'] += ('    修正日期格式，源日期为:' + sourceDate)
    return bookInfos

def modifyAuthor(bookInfos):
    for bookKey in bookInfos:
        bookInfo = bookInfos[bookKey]
        sourceAuthor = bookInfo['著者']
        modifyAuthorStr = sourceAuthor.replace(',' , ' ')
        modifyAuthorStr = modifyAuthorStr.replace('，' , ' ')
        modifyAuthorStr = modifyAuthorStr.replace(':' , ' ')
        modifyAuthorStr = modifyAuthorStr.replace(';' , ' ')
        modifyAuthorStr = modifyAuthorStr.replace('；' , ' ')
        modifyAuthorStr = modifyAuthorStr.replace('《' , '')
        modifyAuthorStr = modifyAuthorStr.replace('》' , '')
        modifyAuthorStr = modifyAuthorStr.replace('主编' , '')
        modifyAuthorStr = modifyAuthorStr.replace('编著' , '')
        modifyAuthorStr = modifyAuthorStr.replace('编注' , '')
        modifyAuthorStr = modifyAuthorStr.replace('编辑' , '')
        modifyAuthorStr = modifyAuthorStr.replace('编写' , '')
        modifyAuthorStr = modifyAuthorStr.replace('整理' , '')
        modifyAuthorStr = modifyAuthorStr.replace('改写' , '')
        modifyAuthorStr = modifyAuthorStr.replace('著' , '')
        modifyAuthorStr = modifyAuthorStr.replace('等' , '')
        modifyAuthorStr = modifyAuthorStr.replace('编' , '')
        modifyAuthorStr = modifyAuthorStr.replace('注' , '')
        modifyAuthorStr=re.sub('\(.*?\)','',modifyAuthorStr)
        modifyAuthorStr=re.sub('\（.*?\）','',modifyAuthorStr)
        modifyAuthorStr=re.sub('\[.*?\]','',modifyAuthorStr)
        modifyAuthorStr=re.sub('\【.*?\】','',modifyAuthorStr)

        if modifyAuthorStr.find(' '):
            sqList = modifyAuthorStr.split(' ')
            newstr = ''
            for str in sqList:
                if '译' not in str:
                    if newstr != '' : newstr+=' '
                    newstr+=str;
            modifyAuthorStr = newstr

        modifyAuthorStr = modifyAuthorStr.strip()
        if sourceAuthor!=modifyAuthor:
            bookInfo['著者'] = modifyAuthorStr
            modifyBookInfos[bookKey] = bookInfo.copy()
            modifyBookInfos[bookKey]['备注'] += ('    修正著者格式，源著者为:' + sourceAuthor)
            bookInfo['备注'] += ('    修正著者格式，源著者为:' + sourceAuthor)
    return bookInfos

def modifyPrice(bookInfos):
    for bookKey in bookInfos:
        bookInfo = bookInfos[bookKey]
        sourcePrice = bookInfo['价格']

        pattern = re.compile('^[0-9]+\.[0-9]{2}$')
        if pattern.match(sourcePrice):   #无误
            continue

        modifyPriceStr=re.sub('\(.*?\)','',sourcePrice)
        modifyPriceStr=re.sub('\（.*?\）','',modifyPriceStr)
        modifyPriceStr=re.sub('\[.*?\]','',modifyPriceStr)
        modifyPriceStr=re.sub('\【.*?\】','',modifyPriceStr)

        modifyPriceStr = modifyPriceStr.strip()
        spList = re.findall(r"\d+\.?\d*",modifyPriceStr)
        if len(spList) <= 0 : 
            modifyPriceStr = '80.00'
        else:
            modifyPriceStr = spList[0]
        modifyPriceStr = '%.2f' % float(modifyPriceStr)
        if sourcePrice!=modifyPriceStr:
            bookInfo['价格'] = modifyPriceStr
            modifyBookInfos[bookKey] = bookInfo.copy()
            modifyBookInfos[bookKey]['备注'] += ('    修正价格格式，源价格为:' + sourcePrice)
            bookInfo['备注'] += ('    修正价格格式，源价格为:' + sourcePrice)
    return bookInfos

def saveWrokBook(wrokBook, savePath, sheet, bookInfos, showMore=False):
    if os.path.exists(savePath):
        os.remove(savePath)  # 删除文件
    newIndex = templateStartRow
    for bookKey in bookInfos:
        bookInfo = bookInfos[bookKey]
        sheet.write(newIndex,0,bookInfo['ISBN'])  
        sheet.write(newIndex,1,bookInfo['书名'])  
        sheet.write(newIndex,2,bookInfo['著者'])  
        sheet.write(newIndex,3,bookInfo['出版社'])  
        sheet.write(newIndex,4,bookInfo['出版日期'])  
        sheet.write(newIndex,5,bookInfo['价格'])  
        sheet.write(newIndex,6,bookInfo['分类号'])  
        code = bookInfo['条形码']
        if len(code) == 6: code = '0' + code
        sheet.write(newIndex,7,code)  
        sheet.write(newIndex,8,bookInfo['馆藏地'])  
        sheet.write(newIndex,9,bookInfo['所在层架'])  
        if(showMore == True):
            sheet.write(newIndex,10,bookInfo['备注'])  
            sheet.write(newIndex,11,bookInfo['原始档案行数'] + 1)  
        newIndex += 1

    if(showMore == True):
        sheet.write(1,10,'备注')  
        sheet.write(1,11,'原始档案行数')  
    wrokBook.save(savePath)

bookInfos = getBookInfos(sourceSheet)
bookInfos = deleteByNameSame(bookInfos)
bookInfos = deleteByBookNameWarn(bookInfos)
bookInfos = deleteByAuthorNameWarn(bookInfos)
bookInfos = deleteByISBNWarn(bookInfos)
bookInfos = modifyDate(bookInfos)
bookInfos = modifyAuthor(bookInfos)
bookInfos = modifyPrice(bookInfos)

saveWrokBook(targetWrokBook, targetWrokBookPath, targetSheet, bookInfos)
saveWrokBook(deleteWrokBook, deleteWrokBookPath, deleteSheet, deleteBookInfos, True)
saveWrokBook(modifyWrokBook, modifyWrokBookPath, modifySheet, modifyBookInfos, True)




