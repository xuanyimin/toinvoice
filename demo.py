# -*- coding: utf-8 -*-

import xlrd
from xlutils.copy import copy
from xml.dom.minidom import Document

invoice_top = 100000
base= {u'Gfmc':'select',u'Gfsh':'9133052100000000000',u'Gfyhzh':'select',u'Gfdzdh':'select',u'Bz':u'昊添财务 tel: 18969275032',
           u'Fhr':'',u'Skr':'',u'Spbmbbh':'19.0',u'Hsbz':'0',u'Sgbz':'0'}
doc = Document()
Kp = doc.createElement('Kp')
doc.appendChild(Kp)
Version = doc.createElement('Version')
Version.appendChild(doc.createTextNode('3.0'))
Kp.appendChild(Version)
Fpxx = doc.createElement('Fpxx')
Kp.appendChild(Fpxx)

def to_xml(djh,invoice,number):
    #处理xml发票张数
    Zsl = doc.createElement('Zsl')
    Zsl.appendChild(doc.createTextNode(str(number+1)))
    Fpxx.appendChild(Zsl)
    xml_company(djh, invoice)

def xml_company(djh,invoice):
    #处理XML每张发票的公司抬头
    for newcows,list in invoice.iteritems():
        Fpsj = doc.createElement('Fpsj')
        Fpxx.appendChild(Fpsj)
        Fp = doc.createElement('Fp')
        Fpsj.appendChild(Fp)
        base[u'Djh'] = djh + newcows
        for key, vule in base.iteritems():
            company = doc.createElement(key)
            if isinstance(vule, (str, unicode)):
                company.appendChild(doc.createTextNode(vule.replace(u'\xa0', ' ')))
            else:
                company.appendChild(doc.createTextNode(str(vule).replace(u'\xa0', ' ')))
            Fp.appendChild(company)
        xml_line(list,Fp)

def xml_line(list,Fp):
    # 处理发票的开票明细
    xh = 0
    line = len(list)
    Spxx = doc.createElement('Spxx')
    for data in range(0, line):
        in_xls_data = list[data]
        if in_xls_data.get(u'Spmc'):
            xh += 1
            in_xls_data[u'Xh'] = xh
            Fp.appendChild(Spxx)
            Sph = doc.createElement('Sph')
            Spxx.appendChild(Sph)
            for key, vule in in_xls_data.iteritems():
                mi = doc.createElement(key)
                if isinstance(vule, (str, unicode)):
                    mi.appendChild(doc.createTextNode(vule.replace(u'\xa0', ' ')))
                else:
                    mi.appendChild(doc.createTextNode(str(vule).replace(u'\xa0', ' ')))
                Sph.appendChild(mi)

def formxls(select):
    #处理EXCEL
    xls_data = xlrd.open_workbook('test.xls')
    #从工作薄名中取得第几张发票
    books = xls_data.sheet_names()[0]
    djh = int(books)
    table = xls_data.sheets()[0]
    # 取得行数
    ncows = table.nrows
    colnames = table.row_values(1)
    invoice ={}
    list = []
    newcows = 0
    amount = 0
    number = 0
    for rownum in range(2, ncows):
        row = table.row_values(rownum)
        if row:
            app = {}
            for i in range(len(colnames)):
                app[colnames[i]] = row[i]
            if select == '1':
                app[u'Dj'] = app[u'Dj'] / (1+app[u'Slv'])
                app[u'Je'] = app[u'Je'] / (1 + app[u'Slv'])
            amount += app.get(u'Je')
            if app.get(u'Spmc'):
                if amount > invoice_top:
                    amount = 0
                    invoice[number] = list
                    list = []
                    list.append(app)
                    number += 1
                    newcows = 0
                else:
                    list.append(app)
                newcows += 1

    invoice[number] = list
    to_xml(djh,invoice,number)
    print u'已重新生test.xml请到开票系统去导入'
    # 写xls：
    wb = copy(xls_data)
    idx = xls_data.sheet_names().index(books)
    wb.get_sheet(idx).name = str(int(djh) + number + 1)
    wb.save('test.xls')
    # 写XML
    with open('test.xml', 'w') as f:
        f.write(doc.toprettyxml(indent='\t', encoding='GBK'))

if __name__ == "__main__":
    print u'含税请打1，不含税请打0'
    select = raw_input(u'selct:')
    if select == '1' or select == '0':
        formxls(select)
    else:
        print u'请重新运行并正确选择'
