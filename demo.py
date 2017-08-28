# -*- coding: utf-8 -*-

import xlrd
from xml.dom.minidom import Document

def to_xml():
    doc = Document()
    Kp = doc.createElement('Kp')
    doc.appendChild(Kp)
    Version = doc.createElement('Version')
    Version.appendChild(doc.createTextNode('2.0'))
    Kp.appendChild(Version)
    Fpxx = doc.createElement('Fpxx')
    Kp.appendChild(Fpxx)
    Zsl = doc.createElement('Zsl')
    Zsl.appendChild(doc.createTextNode('1'))
    Fpxx.appendChild(Zsl)
    Fpsj = doc.createElement('Fpsj')
    Fpxx.appendChild(Fpsj)
    Fp = doc.createElement('Fp')
    Fpsj.appendChild(Fp)
    Djh = doc.createElement('Djh')
    Djh.appendChild(doc.createTextNode('1'))
    Fp.appendChild(Djh)
    Gfmc = doc.createElement('Gfmc')
    Gfmc.appendChild(doc.createTextNode(u'Select'))
    Fp.appendChild(Gfmc)
    Gfsh = doc.createElement('Gfsh')
    Gfsh.appendChild(doc.createTextNode(u'00000000000'))
    Fp.appendChild(Gfsh)
    Gfyhzh = doc.createElement('Gfyhzh')
    Gfyhzh.appendChild(doc.createTextNode(u'Select'))
    Fp.appendChild(Gfyhzh)
    Gfdzdh = doc.createElement('Gfdzdh')
    Gfdzdh.appendChild(doc.createTextNode(u'Select'))
    Fp.appendChild(Gfdzdh)
    Bz = doc.createElement('Bz')
    Bz.appendChild(doc.createTextNode(u'昊添财务 tel: 18969275032'))
    Fp.appendChild(Bz)
    Fhr = doc.createElement('Fhr')
    Fhr.appendChild(doc.createTextNode(''))
    Fp.appendChild(Fhr)
    Skr = doc.createElement('Skr')
    Skr.appendChild(doc.createTextNode(''))
    Fp.appendChild(Skr)
    Spbmbbh = doc.createElement('Spbmbbh')
    Spbmbbh.appendChild(doc.createTextNode('13.0'))
    Fp.appendChild(Spbmbbh)
    Hsbz = doc.createElement('Hsbz')
    Hsbz.appendChild(doc.createTextNode('0'))
    Fp.appendChild(Hsbz)

    #处理EXCEL
    xls_data = xlrd.open_workbook('test.xls')
    table = xls_data.sheets()[0]
    # 取得行数
    ncows = table.nrows
    ncols = table.ncols
    colnames = table.row_values(1)
    list = []
    newcows = 0
    for rownum in range(2, ncows):
        row = table.row_values(rownum)
        if row:
            app = {}
            for i in range(len(colnames)):
                app[colnames[i]] = row[i]
            if app.get(u'Spmc'):
                list.append(app)
                newcows += 1
    xh = 0
    for data in range(0, newcows):
        in_xls_data = list[data]
        if in_xls_data.get(u'Spmc'):
            xh += 1
            in_xls_data[u'Xh'] = xh
            Spxx = doc.createElement('Spxx')
            Fp.appendChild(Spxx)
            Sph = doc.createElement('Sph')
            Spxx.appendChild(Sph)
            for key,vule in in_xls_data.iteritems():
                mi = doc.createElement(key)
                if isinstance(vule, (str, unicode)):
                    mi.appendChild(doc.createTextNode(vule))
                else:
                    mi.appendChild(doc.createTextNode(str(vule)))
                Sph.appendChild(mi)

    with open('test.xml', 'w') as f:
        f.write(doc.toprettyxml(indent='\t', encoding='GBK'))


if __name__ == "__main__":
    to_xml()
