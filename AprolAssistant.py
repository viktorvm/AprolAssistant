#! /usr/bin/env python
# -*- coding: utf-8 -*-

# ----------------------------------------------------------------------------------------------------------------------
#                                      Объявление переменных и функций
# ----------------------------------------------------------------------------------------------------------------------
# путь к папке с файлами
path = 'd:\Hlam\\'
# ------Считываемые файлы---------
# имя исходного файла Excel
srcXLS = 'DataTable.xls'
# ------Создаваемые файлы---------
# имя создаваемого файла для импорта Connectors
connFName = 'connectors.imp'
# имя создаваемого файла для импорта Constants
constFName = 'constants.imp'
# имя проекта (используется при генерации Constant и Connector'ов)
# припривязке тегов к дисплею, парсится и вытаскивает имя проекта с файла дисплея
_pName = 'suzun_s4'

# открываем файл логов
_log = open(path + 'log.txt', 'w')


# ---> преобразует значение MA в INT к значению трехзначного STRING
#       INT используется в логике, для определения максимального числа MA
#       STRING в заголовке в файле импорта
#       т.е. bиз 15 INT в 015 STRING
def maToString(ma):
    if ma < 10:
        return '00' + str(ma)
    if 9 < ma < 100:
        return '0' + str(ma)
    if ma > 99:
        return str(ma)

# ---> генерирует одну запись для импорта Connector'ов
def generateConnectorEntry(pn, ps, type):
    nstr = '\n\P/' + pn + '\I\V/' + ps.encode('cp1251') + '\import  (' \
                                                          '\n   var_type = "CON" ,' \
                                                          '\n   is_output = True ,' \
                                                          '\n   iec_type = "' + type + '" ,' \
                                                          '\n   unit = "" ,' \
                                                          '\n   desc001 = "" ,' \
                                                          '\n   tb_structure = "R 3.4-01" ,' \
                                                          '\n)' \
                                                          '\n'
    return nstr

# ---> создает файл импорта Connector'ов
def createConnectors(pn, sheets):
    # открываем файл, в который будем писать
    f = open(path + connFName, 'w')
    f.write('// Attention: Generated Automaticaly with Python script\n')

    for sheet in sheets:
        tags = sheets[sheet]
        if sheet == 'AI' or sheet == 'DI':
            for ps in tags:
                f.write(generateConnectorEntry(pn, ps + '_ACK', 'BOOL'))

    f.write(generateConnectorEntry(pn, 'LOCK_GLOB', 'LSTRING'))
    f.write(generateConnectorEntry(pn, 'LOCK_MHRES', 'LSTRING'))
    f.write(generateConnectorEntry(pn, 'LOCK_PIDTUNE', 'LSTRING'))
    f.write(generateConnectorEntry(pn, 'LOCK_AITUNE', 'LSTRING'))

    f.close()
    return 'Файл "' + connFName + '" создан'

# ---> генерирует одну запись для импорта Constant переменных
def generateConstantEntry(pn, ps, ds, un):
    nstr = '\P/' + pn + '\PD/consts\V1.1.1/' + ps.encode(
        'cp1251') + '_PS  (' \
                    '\n   desc001 = "" ,' \
                    '\n   is_const ,' \
                    '\n   type = "LSTRING" ,' \
                    '\n   value = "' + ps.encode('cp1251') + '" ,\n)\n\n'
    nstr += '\P/' + pn + '\PD/consts\V1.1.1/' + ps.encode(
        'cp1251') + '_DS  (' \
                    '\n   desc001 = "" ,' \
                    '\n   is_const ,' \
                    '\n   type = "LSTRING" ,' \
                    '\n   value = "' + ds.encode('cp1251') + '" ,\n)\n\n'

    if un is not None:
        nstr += '\P/' + pn + '\PD/consts\V1.1.1/' + ps.encode(
            'cp1251') + '_UN  (' \
                        '\n   desc001 = "" ,' \
                        '\n   is_const ,' \
                        '\n   type = "LSTRING" ,' \
                        '\n   value = "' + un.encode('cp1251') + '" ,\n)\n\n'

    return nstr

# ---> генерирует одну запись для импорта Constant постоянных
def generateConstantEntryTypical(pn, n, dt, vl):
    nstr = '\P/' + pn + '\PD/consts\V1.1.1/' + n + '  (\n' \
                                                   '   desc001 = "" ,\n' \
                                                   '   is_const ,\n' \
                                                   '   type = "' + dt + '" ,\n' \
                                                                        '   value = "' + vl + '" ,\n' \
                                                                                              ')\n'
    return nstr

# ---> создает файл импорта Constant
def createConstants(pn, sheets):
    # формируем заголовок
    wstr = '// Attention: Generated Automaticaly with Python script\n\n'
    wstr += '\P/' + pn + '\PD/consts  (' \
                         '\n   active = \P/' + pn + '\PD/consts\V1.1.1 ,' \
                                                    '\n   active_time = 31.01.2017 12:19:14.891 ,' \
                                                    '\n   active_user = "" ,' \
                                                    '\n   class = "AT" ,' \
                                                    '\n   current = \P/' + pn + '\PD/consts\V1.1.1 ,' \
                                                                                '\n   otref = \OT\P\AT ,' \
                                                                                '\n   tb_structure = "R 3.4-0301" ,' \
                                                                                '\n)' \
                                                                                '\n\n'

    wstr += '\P/' + pn + '\PD/consts\V1.1.1  (' \
                         '\n   active_save = (Vt_OPAQUE)"active" ,' \
                         '\n   crelease = "R 4.0-11" ,' \
                         '\n   instance = "consts" ,' \
                         '\n   orig = "V0.0.6" ,' \
                         '\n   page_comm = "" ,' \
                         '\n)' \
                         '\n\n'

    wstr += '\P/' + pn + '\PD/consts\V1.1.1\DOC  ( )' \
                         '\n\P/' + pn + '\PD/consts\V1.1.1\DOC\Images  ( )' \
                                        '\n\P/' + pn + '\PD/consts\V1.1.1\DOC\Links  ( )' \
                                                       '\n\n'

    # пишем константы тегов
    for sheet in sheets:
        tags = sheets[sheet]
        for ps in tags:
            ds = tags[ps][0]
            if sheet == 'AI':
                un = tags[ps][1]
            else:
                un = None
            wstr += generateConstantEntry(pn, ps, ds, un)

    # генерируем числа INT
    for i in range(0, 11):
        wstr += generateConstantEntryTypical(pn, 'INT_' + str(i), 'INT', str(i))

    # генерируем числа UINT
    for i in range(0, 11):
        wstr += generateConstantEntryTypical(pn, 'UINT_' + str(i), 'UINT', str(i))

    # генерируем числа BOOL
    for i in range(0, 2):
        wstr += generateConstantEntryTypical(pn, 'BOOL_' + str(i), 'BOOL', str(i))

    # делаем выборку по типу форматов и дописываем эти константы
    fms = set()
    for t in sheets['AI']:
        fms.add(sheets['AI'][t][2])

    for fm in fms:
        wstr += generateConstantEntryTypical(pn, 'FM_' + fm.encode('cp1251'), 'LSTRING',
                                             '%' + fm.encode('cp1251').replace('x', '.') + 'f')

    # Пишем строку в файл
    f = open(path + constFName, 'w')
    f.write(wstr)
    f.close()
    return 'Файл "' + constFName + '" создан'

# ---> генерирует одну запись переменной фейсплейта
def generateFpTagEntry(header, io, name, type, dflt, ref):
    nstr = '\n\n//Не верно отработала функция generateFpTagEntry( header, io, name, type, ref)\n' \
           + header.encode('cp1251') + '\n' + name.encode('cp1251') + '\n\n'

    if ref is None:
        ref = '\\'
    if dflt is None:
        dflt = 'nil'
    else:
        dflt = '"' + dflt + '"'

    if io == 'I':
        nstr = '\n' + header + '\I/' + name + '  (\n' \
               '   dflt = ' + dflt + ',\n' \
               '   force_output = False,\n' \
               '   iec_type = "' + type + '",\n' \
               '   pin_iec_type = "' + type + '",\n' \
               '   right = \,\n' \
               '   used_ref = ' + ref + ' ,\n)\n'
    if io == 'O':
        nstr = '\n' + header + '\O/' + name + '  (\n' \
               '   force_output = False,\n' \
               '   iec_type = "' + type + '",\n' \
               '   pin_iec_type = "' + type + '",\n' \
               '   right = \,\n' \
               '   startup_value = nil ,\n' \
               '   used_ref = ' + ref + ' ,\n)\n'

    return nstr

# ---> генерирует запись с фейсплейтами mainFP и optFP для заданного AI
def generateAIFPs(header, pn, newma, ps, vals):
    ind = str(vals[3])
    fm = vals[2]
    arr = vals[4]

    fpText = ''
    ma = maToString(newma)
    maopt = maToString(newma + 1)

    fpText += '\n' + header + ma + '  (\n' \
                                   '   fbref = \L\L/UserDefined\B/AI_mainFP ,\n' \
                                   '   height = 7351 ,\n' \
                                   '   is_popup = True ,\n' \
                                   '   is_prop = True ,\n' \
                                   '   name = "' + ps + '_mainFP" ,\n' \
                                   '   width = 11001 ,\n)\n'

    tv = {'ACK': ['I', 'BOOL', ps + '_ACK'],
          'DS': ['I', 'LSTRING', 'consts_' + ps + '_DS'],
          'EMAN': ['I', 'BOOL', arr + '___' + ind + '___EMAN'],
          'FM': ['I', 'LSTRING', 'consts_FM_' + fm],
          'PS': ['I', 'LSTRING', 'consts_' + ps + '_PS'],
          'PV': ['I', 'REAL', arr + '___' + ind + '___PV'],
          'SH': ['I', 'REAL', arr + '___' + ind + '___SH'],
          'SL': ['I', 'REAL', arr + '___' + ind + '___SL'],
          'ST': ['I', 'UINT', arr + '___' + ind + '___STATUS'],
          'UN': ['I', 'LSTRING', 'consts_' + ps + '_UN'],
          'EH': ['O', 'BOOL', arr + '___' + ind + '___EH'],
          'EHH': ['O', 'BOOL', arr + '___' + ind + '___EHH'],
          'EL': ['O', 'BOOL', arr + '___' + ind + '___EL'],
          'ELL': ['O', 'BOOL', arr + '___' + ind + '___ELL'],
          'VH': ['O', 'REAL', arr + '___' + ind + '___VH'],
          'VHH': ['O', 'REAL', arr + '___' + ind + '___VHH'],
          'VL': ['O', 'REAL', arr + '___' + ind + '___VL'],
          'VLL': ['O', 'REAL', arr + '___' + ind + '___VLL'],
          'LOCK': ['I', 'LSTRING', 'LOCK_GLOB']
          }

    for t in tv:
        fpText += generateFpTagEntry(header + ma, tv[t][0], t, tv[t][1], None,
                                     '\P/' + pn + '\\I\\V/' + tv[t][2])

    fpText += '\n' + header + ma + '\O/FP  (\n' \
                                   '   force_output = False,\n' \
                                   '   iec_type = nil,\n' \
                                   '   macro = ' + header + maopt + ' ,\n' \
                                   '   pin_iec_type = "FACEPLATE",\n' \
                                   '   right = \,\n' \
                                   '   used_ref = \,\n)\n'

    fpText += '\n' + header + maopt + '  (\n' \
                                      '   fbref = \L\L/UserDefined\B/AI_optFP ,\n' \
                                      '   height = 5301 ,\n' \
                                      '   is_popup = True ,\n' \
                                      '   is_prop = True ,\n' \
                                      '    name = "' + ps + '_optFP",\n' \
                                      '   width = 8901 ,\n)\n'

    tv = {'BL': ['I', 'BOOL', arr + '___' + ind + '___BL'],
          'DBL': ['I', 'BOOL', arr + '___' + ind + '___DBL'],
          'CUR': ['I', 'REAL', arr + '___' + ind + '___CUR'],
          'DS': ['I', 'LSTRING', 'consts_' + ps + '_DS'],
          'FM': ['I', 'LSTRING', 'consts_FM_' + fm],
          'PV': ['I', 'REAL', arr + '___' + ind + '___PV'],
          'SH': ['O', 'REAL', arr + '___' + ind + '___SH'],
          'SL': ['O', 'REAL', arr + '___' + ind + '___SL'],
          'UN': ['I', 'LSTRING', 'consts_' + ps + '_UN'],
          'EMAN': ['O', 'BOOL', arr + '___' + ind + '___EMAN'],
          'VMAN': ['O', 'REAL', arr + '___' + ind + '___VMAN'],
          'LOCK_TUNE': ['I', 'LSTRING', 'LOCK_AITUNE']
          }
    for t in tv:
        fpText += generateFpTagEntry(header + maopt, tv[t][0], t, tv[t][1], None,
                                     '\P/' + pn + '\\I\\V/' + tv[t][2])

    return fpText

# ---> генерирует запись с фейсплейтами mainFP для заданного DI
def generateDIFPs(header, pn, newma, ps, vals):
    ind = str(vals[1])
    arr = vals[2]
    ct = str(vals[3])
    cf = str(vals[4])

    fpText = ''
    ma = maToString(newma)

    fpText += '\n' + header + ma + '  (\n' \
                                   '   fbref = \L\L/UserDefined\B/DI_FP ,\n' \
                                   '   height = 4481 ,\n' \
                                   '   is_popup = True ,\n' \
                                   '   is_prop = True ,\n' \
                                   '   name = "' + ps + '_mainFP" ,\n' \
                                   '   width = 7041 ,\n)\n'

    tv = {'DESCR': ['I', 'LSTRING', 'consts_' + ps + '_DS', None],
          'PS': ['I', 'LSTRING', 'consts_' + ps + '_PS', None],
          'ST': ['I', 'BOOL', arr + '___' + ind + '___res', None],
          'BL': ['I', 'BOOL', arr + '___' + ind + '___bl', None],
          'DBL': ['I', 'BOOL', arr + '___' + ind + '___dbl', None],
          'COL_TRUE': ['I', 'UINT', None, ct],
          'COL_FALSE': ['I', 'UINT', None, cf]
          }

    for t in tv:
        if tv[t][2] is None:
            fpText += generateFpTagEntry(header + ma, tv[t][0], t, tv[t][1], tv[t][3], None)
        else:
            fpText += generateFpTagEntry(header + ma, tv[t][0], t, tv[t][1], tv[t][3],
                                     '\P/' + pn + '\\I\\V/' + tv[t][2])

    return fpText

# ---> генерирует запись с фейсплейтами mainFP, MH_FP, PID_FP, RES_FP_confirm и , RP_FM_confirm для заданного насоса
def generatePumpFPs(header, pn, newma, ps, sheets):
    vals = sheets['PUMP'][ps]
    ind = str(vals[1])
    arr = vals[2]
    fr = vals[3]
    pvps = vals[4]
    svfm = ''
    if fr == '+':
        try:
            svfm = sheets['AI'][vals[4]][2]
        except:
            _log.write(str(
                datetime.now()) + ' - Позиция SV: "' + pvps + '" для насоса "' + ps + '" не найдена в записях AI')

    fpText = ''
    ma = maToString(newma)
    mamh = maToString(newma + 1)
    marp = maToString(newma + 2)
    mares = maToString(newma + 3)
    mapid = maToString(newma + 4)

    # генерируем теги фейсплейта mainFP
    fpText += '\n' + header + ma + '  (\n' \
                                   '   fbref = \L\L/UserDefined\B/PUMP_FP ,\n' \
                                   '   height = 11651 ,\n' \
                                   '   is_popup = True ,\n' \
                                   '   is_prop = True ,\n' \
                                   '   name = "' + ps + '_mainFP" ,\n' \
                                   '   width = 6401 ,\n)\n'
    tv = {'BLR': ['I', 'USINT', arr + '___' + ind + '___IL'],
          'DBL': ['I', 'BOOL', arr + '___' + ind + '___DBL'],
          'DS': ['I', 'LSTRING', 'consts_' + ps + '_DS'],
          'PS': ['I', 'LSTRING', 'consts_' + ps + '_PS'],
          'PV': ['I', 'REAL', arr + '___' + ind + '___Herz'],
          'RES': ['I', 'BOOL', arr + '___' + ind + '___Mode_1'],
          'RP': ['I', 'BOOL', arr + '___' + ind + '___Rem'],
          'ST': ['I', 'UINT', arr + '___' + ind + '___Status'],
          'AM': ['O', 'UINT', arr + '___' + ind + '___Mode'],
          'MV': ['O', 'REAL', arr + '___' + ind + '___MV'],
          'SV': ['O', 'REAL', arr + '___' + ind + '___SV'],
          'XOFF': ['O', 'BOOL', arr + '___' + ind + '___MStop'],
          'XON': ['O', 'BOOL', arr + '___' + ind + '___MStart'],
          'LOCK': ['I', 'LSTRING', 'LOCK_GLOB']
          }
    for t in tv:
        fpText += generateFpTagEntry(header + ma, tv[t][0], t, tv[t][1], None,
                                     '\P/' + pn + '\\I\\V/' + tv[t][2])

    # при наличии ПЧ
    if fr == '+':
        # привязываем теги для PID
        tv = {
            'SVFM': ['I', 'LSTRING', 'consts_FM_' + svfm],
            'SVUN': ['I', 'LSTRING', 'consts_' + pvps + '_UN']
        }
        for t in tv:
            fpText += generateFpTagEntry(header + ma, tv[t][0], t, tv[t][1], None,
                                         '\P/' + pn + '\\I\\V/' + tv[t][2])
    else:
        fpText += generateFpTagEntry(header + ma, 'I', 'noFR', 'BOOL', '1', None)

    # генерируем привязку к фейсплейту MH_FP
    fpText += '\n' + header + ma + '\O/MH_FP  (\n' \
                                   '   force_output = False,\n' \
                                   '   iec_type = nil,\n' \
                                   '   macro = ' + header + mamh + ' ,\n' \
                                                                   '   pin_iec_type = "FACEPLATE",\n' \
                                                                   '   right = \,\n' \
                                                                   '   used_ref = \,\n)\n'
    # генерируем привязку к фейсплейту RES_FP
    fpText += '\n' + header + ma + '\O/RES_FP  (\n' \
                                   '   force_output = False,\n' \
                                   '   iec_type = nil,\n' \
                                   '   macro = ' + header + mares + ' ,\n' \
                                                                    '   pin_iec_type = "FACEPLATE",\n' \
                                                                    '   right = \,\n' \
                                                                    '   used_ref = \,\n)\n'
    # генерируем привязку к фейсплейту RP_FP
    fpText += '\n' + header + ma + '\O/RP_FP  (\n' \
                                   '   force_output = False,\n' \
                                   '   iec_type = nil,\n' \
                                   '   macro = ' + header + marp + ' ,\n' \
                                                                   '   pin_iec_type = "FACEPLATE",\n' \
                                                                   '   right = \,\n' \
                                                                   '   used_ref = \,\n)\n'
    # при наличии ПЧ
    if fr == '+':
        # генерируем теги фейсплейта PID_FP и привязку к нему из mainFP
        fpText += '\n' + header + ma + '\O/PID_FP  (\n' \
                                       '   force_output = False,\n' \
                                       '   iec_type = nil,\n' \
                                       '   macro = ' + header + mapid + ' ,\n' \
                                                                        '   pin_iec_type = "FACEPLATE",\n' \
                                                                        '   right = \,\n' \
                                                                        '   used_ref = \,\n)\n'

        fpText += '\n' + header + mapid + '  (\n' \
                                          '   fbref = \L\L/UserDefined\B/PUMP_FP_PID ,\n' \
                                          '   height = 7751 ,\n' \
                                          '   is_popup = True ,\n' \
                                          '   is_prop = True ,\n' \
                                          '   name = "' + ps + '_PID_FP" ,\n' \
                                                               '   width = 14901 ,\n)\n'
        tv = {'PS': ['I', 'LSTRING', 'consts_' + ps + '_PS'],
              'PV': ['I', 'REAL', arr + '___' + ind + '___Herz'],
              'PVFM': ['I', 'LSTRING', 'consts_FM_' + svfm],
              'PVPS': ['I', 'LSTRING', 'consts_' + pvps + '_PS'],
              'PVUN': ['I', 'LSTRING', 'consts_' + pvps + '_UN'],
              'AM': ['O', 'UINT', arr + '___' + ind + '___Mode'],
              'MV': ['O', 'REAL', arr + '___' + ind + '___MV'],
              'SV': ['O', 'REAL', arr + '___' + ind + '___SV'],
              'P': ['O', 'REAL', arr + '___' + ind + '___P'],
              'I': ['O', 'REAL', arr + '___' + ind + '___I'],
              'D': ['O', 'REAL', arr + '___' + ind + '___D'],
              'LOCK_GL': ['I', 'LSTRING', 'LOCK_GLOB'],
              'LOCK_PID': ['I', 'LSTRING', 'LOCK_PIDTUNE']
              }
        for t in tv:
            fpText += generateFpTagEntry(header + mapid, tv[t][0], t, tv[t][1], None,
                                         '\P/' + pn + '\\I\\V/' + tv[t][2])

    # генерируем теги фейсплейта MH_FP
    fpText += '\n' + header + mamh + '  (\n' \
                                     '   fbref = \L\L/UserDefined\B/PUMP_FP_MH ,\n' \
                                     '   height = 7751 ,\n' \
                                     '   is_popup = True ,\n' \
                                     '   is_prop = True ,\n' \
                                     '   name = "' + ps + '_MH_FP" ,\n' \
                                                          '   width = 8051 ,\n)\n'
    tv = {'PS': ['I', 'LSTRING', 'consts_' + ps + '_PS'],
          'AT': ['I', 'REAL', arr + '___' + ind + '___MHA_Total'],
          'CD': ['I', 'REAL', arr + '___' + ind + '___MHA_Day'],
          'CM': ['I', 'REAL', arr + '___' + ind + '___MHA_Month'],
          'LD': ['I', 'REAL', arr + '___' + ind + '___MHA_Daye'],
          'LM': ['I', 'REAL', arr + '___' + ind + '___MHA_Monthe'],
          'RESET': ['O', 'BOOL', arr + '___' + ind + '___MotoRes'],
          'LOCK': ['I', 'LSTRING', 'LOCK_MHRES']
          }
    for t in tv:
        fpText += generateFpTagEntry(header + mamh, tv[t][0], t, tv[t][1], None,
                                     '\P/' + pn + '\\I\\V/' + tv[t][2])

    # генерируем теги фейсплейта RP_FP
    fpText += '\n' + header + marp + '  (\n' \
                                     '   fbref = \L\L/UserDefined\B/PUMP_FP_confirm ,\n' \
                                     '   height = 3521 ,\n' \
                                     '   is_popup = True ,\n' \
                                     '   is_prop = True ,\n' \
                                     '   name = "' + ps + '_RP_FP_confirm" ,\n' \
                                                          '   width = 6081 ,\n)\n'
    tv = {'VAL': ['O', 'BOOL', arr + '___' + ind + '___Rem']}
    fpText += generateFpTagEntry(header + marp, 'O', 'VAL', 'BOOL', None,
                                 '\P/' + pn + '\\I\\V/' + arr + '___' + ind + '___Rem')
    fpText += generateFpTagEntry(header + marp, 'I', 'LOCK', 'LSTRING', None,
                                 '\P/' + pn + '\\I\\V/' + 'LOCK_GLOB')
    fpText += generateFpTagEntry(header + marp, 'I', 'VAR', 'UINT', '0', None)


    # генерируем теги фейсплейта RES_FP
    fpText += '\n' + header + mares + '  (\n' \
                                      '   fbref = \L\L/UserDefined\B/PUMP_FP_confirm ,\n' \
                                      '   height = 3521 ,\n' \
                                      '   is_popup = True ,\n' \
                                      '   is_prop = True ,\n' \
                                      '   name = "' + ps + '_RES_FP_confirm" ,\n' \
                                                           '   width = 6081 ,\n)\n'
    fpText += generateFpTagEntry(header + mares, 'O', 'VAL', 'BOOL', None,
                                 '\P/' + pn + '\\I\\V/' + arr + '___' + ind + '___Mode_1')
    fpText += generateFpTagEntry(header + mares, 'I', 'LOCK', 'LSTRING', None,
                                 '\P/' + pn + '\\I\\V/' + 'LOCK_GLOB')
    fpText += generateFpTagEntry(header + mares, 'I', 'VAR', 'UINT', '1', None)

    return fpText

# ---> генерирует запись с фейсплейтами mainFP для заданной задвижки
def generateValve1FPs(header, pn, newma, ps, vals):
    ind = str(vals[1])
    arr = vals[2]

    fpText = ''
    ma = maToString(newma)

    # генерируем теги фейсплейта mainFP
    fpText += '\n' + header + ma + '  (\n' \
                                   '   fbref = \L\L/UserDefined\B/Valve1_FP ,\n' \
                                   '   height = 5051 ,\n' \
                                   '   is_popup = True ,\n' \
                                   '   is_prop = True ,\n' \
                                   '   name = "' + ps + '_mainFP" ,\n' \
                                   '   width = 7041 ,\n)\n'
    tv = {'BLR': ['I', 'USINT', arr + '___' + ind + '___IL'],
          'DBL': ['I', 'BOOL', arr + '___' + ind + '___DBL'],
          'DS': ['I', 'LSTRING', 'consts_' + ps + '_DS'],
          'PS': ['I', 'LSTRING', 'consts_' + ps + '_PS'],
          'ST': ['I', 'UINT', arr + '___' + ind + '___Status'],
          'AM': ['O', 'UINT', arr + '___' + ind + '___Mode'],
          'XOFF': ['O', 'BOOL', arr + '___' + ind + '___MStop'],
          'XON': ['O', 'BOOL', arr + '___' + ind + '___MStart']
          }
    for t in tv:
        fpText += generateFpTagEntry(header + ma, tv[t][0], t, tv[t][1], None,
                                     '\P/' + pn + '\\I\\V/' + tv[t][2])

    return fpText

# ---> делает привязки переменных graphic block'а AI_visu1
def atachAIGBTags(ssn, pName, ps, var, vals, maMax, header):
    # ACK - к connector'у [ps]_ACK
    # ST - к тегу ПЛК [arr]___[ind]___STATUS
    # PV, EMAN, SL, SH, ELL, EL, EH, EHH, VLL, VL, VH, VHH, CUR, VMAN, BL, DBL - к тегу ПЛК [arr]___[ind]___[var]
    # PS, UN, DS - к constant'е consts_[ps]_[var]
    # FM - к к constant'е consts_FM_[fm] в зависимости от формата в исходных данных

    # название массива
    arr = vals[4]
    # индекс в массиве
    ind = str(vals[3])
    # формат отображения
    fm = vals[2]

    if var == 'ACK':
        return ssn.replace('used_ref = \\ ,', 'used_ref = \\P/' + pName + '\\I\\V/' + ps + '_ACK ,')

    if var == 'ST':
        return ssn.replace('used_ref = \\ ,',
                           'used_ref = \\P/' + pName + '\\I\\V/' + arr + '___' + ind + '___STATUS ,')

    if var == 'PV' or var == 'EMAN' or var == 'SL' or var == 'SH' or var == 'ELL' or var == 'EL' \
            or var == 'EH' or var == 'EHH' or var == 'VLL' or var == 'VL' or var == 'VH' or var == 'VHH' \
            or var == 'CUR' or var == 'VMAN' or var == 'BL' or var == 'DBL':
        return ssn.replace('used_ref = \\ ,',
                           'used_ref = \\P/' + pName + '\\I\\V/' + arr + '___' + ind + '___' + var + ' ,')

    if var == 'PS' or var == 'UN' or var == 'DS':
        return ssn.replace('used_ref = \\ ,',
                           'used_ref = \\P/' + pName + '\\I\\V/consts_' + ps + '_' + var + ' ,')

    if var == 'FM':
        return ssn.replace('used_ref = \\ ,',
                           'used_ref = \\P/' + pName + '\\I\\V/consts_FM_' + fm + ' ,')

    if var == 'FP':
        # прибавляем к maMax единицу, это номер сущности mainFP
        maMax += 1
        newma = maToString(maMax)

        # делаем привязку к Faceplate'у
        # (замечено, что iec_type может быть либо [""] либо [nil], перебираем оба варианта)
        ssn = ssn.replace('iec_type = nil ,\n',
                          'iec_type = nil ,\n   macro = ' + header + '\MA/' + newma + ' ,\n')
        ssn = ssn.replace('iec_type = "" ,\n',
                          'iec_type = "" ,\n   macro = ' + header + '\MA/' + newma + ' ,\n')

        # генерируем текст Faceplate'а
        ssn += generateAIFPs(header + '\MA/', pName, maMax, ps, vals)

        return ssn

    return ssn

# ---> делает привязки переменных graphic block'а DI_visu1
def atachDIGBTags(ssn, pName, ps, var, vals, maMax, header):
    # ACK - к connector'у [ps]_ACK
    # ST, BL, DBL - к тегу ПЛК [arr]___[ind]___res, BL, DBL соответственно
    # PS, DS - к constant'е consts_[ps]_PS, DS
    # COL_FALSE, COL_TRUE значение dflt из таблицы параметров

    # название массива
    arr = vals[2]
    # индекс в массиве
    ind = str(vals[1])
    # цвет в сработке
    ct = str(vals[3])
    # цвет не в сработке
    cf = str(vals[4])

    if var == 'ACK':
        return ssn.replace('used_ref = \\ ,', 'used_ref = \\P/' + pName + '\\I\\V/' + ps + '_ACK ,')
    if var == 'ST':
        return ssn.replace('used_ref = \\ ,',
                           'used_ref = \\P/' + pName + '\\I\\V/' + arr + '___' + ind + '___res ,')
    if var == 'BL':
        return ssn.replace('used_ref = \\ ,',
                           'used_ref = \\P/' + pName + '\\I\\V/' + arr + '___' + ind + '___bl ,')
    if var == 'DBL':
        return ssn.replace('used_ref = \\ ,',
                           'used_ref = \\P/' + pName + '\\I\\V/' + arr + '___' + ind + '___dbl ,')
    if var == 'PS' or var == 'DS':
        return ssn.replace('used_ref = \\ ,',
                           'used_ref = \\P/' + pName + '\\I\\V/consts_' + ps + '_' + var + ' ,')
    if var == 'COL_TRUE':
        return ssn.replace('dflt = nil ,',
                           'dflt = "' + ct + '" ,')
    if var == 'COL_FALSE':
        return ssn.replace('dflt = nil ,',
                           'dflt = "' + cf + '" ,')

    if var == 'FP':
        # прибавляем к maMax единицу, это номер сущности mainFP
        maMax += 1
        newma = maToString(maMax)

        # делаем привязку к Faceplate'у
        # (замечено, что iec_type может быть либо [""] либо [nil], перебираем оба варианта)
        ssn = ssn.replace('iec_type = nil ,\n',
                          'iec_type = nil ,\n   macro = ' + header + '\MA/' + newma + ' ,\n')
        ssn = ssn.replace('iec_type = "" ,\n',
                          'iec_type = "" ,\n   macro = ' + header + '\MA/' + newma + ' ,\n')

        # генерируем текст Faceplate'а
        ssn += generateDIFPs(header + '\MA/', pName, maMax, ps, vals)

        return ssn

    return ssn

# ---> делает привязки переменных graphic block'а Valve1_visu
def atachValve1GBTags(ssn, pName, ps, var, vals, maMax, header):
    # ST, AM, BLR, DBL  - к тегу ПЛК [arr]___[ind]___Status, Mode, IL, DBL соответственно
    # PS, DS - к constant'е consts_[ps]_[var]

    # название массива
    arr = vals[2]
    # индекс в массиве
    ind = str(vals[1])

    if var == 'ST':
        return ssn.replace('used_ref = \\ ,',
                           'used_ref = \\P/' + pName + '\\I\\V/' + arr + '___' + ind + '___Status ,')
    if var == 'AM':
        return ssn.replace('used_ref = \\ ,',
                           'used_ref = \\P/' + pName + '\\I\\V/' + arr + '___' + ind + '___Mode ,')
    if var == 'BLR':
        return ssn.replace('used_ref = \\ ,',
                           'used_ref = \\P/' + pName + '\\I\\V/' + arr + '___' + ind + '___IL ,')
    if var == 'DBL':
        return ssn.replace('used_ref = \\ ,',
                           'used_ref = \\P/' + pName + '\\I\\V/' + arr + '___' + ind + '___DBL ,')
    if var == 'PS' or var == 'DS':
        return ssn.replace('used_ref = \\ ,',
                           'used_ref = \\P/' + pName + '\\I\\V/consts_' + ps + '_' + var + ' ,')

    if var == 'FP':
        # прибавляем к maMax единицу, это номер сущности mainFP
        maMax += 1
        newma = maToString(maMax)

        # делаем привязку к Faceplate'у
        # (замечено, что iec_type может быть либо [""] либо [nil], перебираем оба варианта)
        ssn = ssn.replace('iec_type = nil ,\n',
                          'iec_type = nil ,\n   macro = ' + header + '\MA/' + newma + ' ,\n')
        ssn = ssn.replace('iec_type = "" ,\n',
                          'iec_type = "" ,\n   macro = ' + header + '\MA/' + newma + ' ,\n')

        # генерируем текст Faceplate'а
        ssn += generateValve1FPs(header + '\MA/', pName, maMax, ps, vals)

        return ssn

    return ssn

# ---> делает привязки переменных graphic block'а Pump_visu
def atachPumpGBTags(ssn, pName, ps, var, sheets, maMax, header):
    # PS, DS - к constant'е consts_[ps]_[var]
    # DBL - к тегу ПЛК [arr]___[ind]___[var]
    # ST, XON, XOFF, BLR, RP, AM - к тегу ПЛК [arr]___[ind]___STATUS, MStart, MStop, IL, Rem, Mode
    # PV, RES, CD, CM - к тегу ПЛК [arr]___[ind]___Herz, Mode_1, MHA_Day, MHA_Month
    # LD, LM, AT, REST - к тегу ПЛК [arr]___[ind]___MHA_Daye, MHA_Monthe, MHA_Total, Motores
    # (при наличии ПЧ)
    # SVUN - к constant'е consts_[позиция регулируемой величины]_UN
    # SVFM - к constant'е consts_FM_[fm] в зависимости от формата в исходных данных рег. величины
    # MV, P, I, D - к тегу ПЛК [arr]___[ind]___[var]
    # SP - к тегу ПЛК [arr]___[ind]___SV

    vals = sheets['PUMP'][ps]
    # название массива
    arr = vals[2]
    # индекс в массиве
    ind = str(vals[1])
    # наличие ПЧ
    fr = False
    if vals[3] == '+':
        fr = True

    if var == 'PS' or var == 'DS':
        return ssn.replace('used_ref = \\ ,',
                           'used_ref = \\P/' + pName + '\\I\\V/consts_' + ps + '_' + var + ' ,')
    rep = {
        'DBL': 'DBL',
        'ST': 'Status',
        'BLR': 'IL',
        'AM': 'Mode',
        'RES': 'Mode_1',
    }

    for original in rep:
        if var == original:
            return ssn.replace('used_ref = \\ ,',
                               'used_ref = \\P/' + pName + '\\I\\V/' + arr + '___' + ind + '___' + rep[original] + ' ,')

    if var == 'FP':
        # прибавляем к maMax единицу, это номер сущности главного FP
        maMax += 1
        newma = maToString(maMax)

        # делаем привязку к Faceplate'у
        # (замечено, что iec_type может быть либо [""] либо [nil], перебираем оба варианта)
        ssn = ssn.replace('iec_type = nil ,\n',
                          'iec_type = nil ,\n   macro = ' + header + '\MA/' + newma + ' ,\n')
        ssn = ssn.replace('iec_type = "" ,\n',
                          'iec_type = "" ,\n   macro = ' + header + '\MA/' + newma + ' ,\n')

        # генерируем текст Faceplate'а
        ssn += generatePumpFPs(header + '\MA/', pName, maMax, ps, sheets)

        return ssn

    return ssn

# ---> обрабатывает файл графического дисплея и делает привязку тегов и фейсплейтов
def atachTagsInDisplay(sheets):
    # имя файла графического дисплея
    # dispFName = raw_input('Введите имя графического дисплея для привязки тегов: ')
    # if dispFName == '':
    #     return 'Имя файла не введено'
    dispFName = 'main_display.imp'

    try:
        f = open(path + dispFName, 'r')
        s = f.read()
        f.close()
    except IOError:
        _log.write(str(
            datetime.now()) + ' - Файл дисплея "' + dispFName + '" не найден')
        return 'Ошибка имени файла дисплея'

    # в данной переменной хранится номер сущности элемента на графике, он же MA
    # необходимо знать для автогенерации сущностей Faceplat'ов
    maMax = 0

    # выделяем имя проекта [pName], путь к дисплею [dPath], имя дисплея [dName], весию дисплея [ver]
    t = s[s.find('// Context:	') + 12:]
    e = t.split('/')
    pName = e[1][:e[1].find('\\')]
    dPath = ''
    dName = ''
    for ee in e[2:]:
        if ee.find('\n') != -1:
            dName = ee[:ee.find('\n')]
            break
        dPath = dPath + '/' + ee

    # и формируем типовой заголовок переменных
    header = '\\P/%s\\PD%s/%s\\' % (pName, dPath, dName)
    ver = s[s.find(header + 'V') + header.__len__():]
    ver = ver[:ver.find(' ')]

    # сразу выделяем концовку документа
    # добавляем по ходу работы новые элементы в конец и концовку зашифрованную потом вернем на место
    endText = s[s.find(header + '\\'[:-1] + ver + '\PYTHON  ('):]
    s = s.replace(endText, '')

    # извлекаем имена graphic block'ов
    # regx = re.compile(r'\\P/%s\\PD%s/%s\\V\d+\.\d+\.\d+\\MA/\w+\s+\(\n.*\n.*\n.*\n.*\n.*\n.*\n\)' %
    #                   (pName, dPath, dName))
    regx = re.compile(r'\\P/%s\\PD%s/%s\\V\d+\.\d+\.\d+\\MA/\w+\s+\([\s\S]*?\n\)' %
                      (pName, dPath, dName))
    gBlocks = regx.findall(s)

    # создаем словарь [МА graphic block'а]:[имя]
    gb = {}
    for ss in gBlocks:
        # print ss + '/n-------------/n'
        gb[ss[ss.find('MA/') + 3:ss.rfind('  (')]] = ss[ss.find('"') + 1:ss.rfind('"')]

    # определяем максимальное значение ma
    for ma in gb:
        if int(ma) > maMax:
            maMax = int(ma)

    # извлекаем переменные сущностей graphic block'ов с привязкой к их номеру MA
    # I - Input
    regex = re.compile(r'\\P/%s\\PD%s/%s\\V\d+\.\d+\.\d+\\MA/\w+\\I/\w+\s+\(\n.*\n.*\n.*\n.*\n.*\n.*\n\)' %
                       (pName, dPath, dName))
    gbVarsI = regex.findall(s)

    # O - Output
    regex = re.compile(r'\\P/%s\\PD%s/%s\\V\d+\.\d+\.\d+\\MA/\w+\\O/\w+\s+\(\n.*\n.*\n.*\n.*\n.*\n*\)' %
                       (pName, dPath, dName))
    gbVarsO = regex.findall(s)

    # обходим каждую позицию каждого листа
    for sheet in sheets:
        tags = sheets[sheet]
        for xlsps in tags:
            # обходим все записи файла с Input переменными
            # и обрабатываем переменные, относящиеся к graphic block'у c данной позицией
            for ss in gbVarsI:
                # извлекаем значение МА
                ma = ss[ss.find('MA/') + 3:ss.find('\\I/')]
                # извлекаем имя переменной
                var = ss[ss.find('\\I/') + 3:ss.find('  (')]
                # извлекаем тип переменной
                varType = ss[ss.find('pin_iec_type = "') + 16:ss.rfind('"')]
                # извлекаем позицию graphic block'а по его МА
                ps = gb[ma]

                # если позиция блока записи с позицией записи Excel не совпадает -> идем дальше
                if ps != xlsps:
                    continue

                # иначе извлекаем значения переменных и работаем
                vals = tags[ps]

                # изменяем тип привязанной переменной
                # (замечено, что iec_type может быть либо [""] либо [nil], перебираем оба варианта)
                ssn = ss.replace('iec_type = ""', 'iec_type = "' + varType + '"')
                ssn = ssn.replace('iec_type = nil', 'iec_type = "' + varType + '"')

                # привязываем переменную данного блока
                if sheet == 'AI':
                    ssn = atachAIGBTags(ss, pName, ps, var, vals, None, None)
                    # заменяем фрагмент текста, содержащий переменную, на фрагмент с привязкой
                    s = s.replace(ss, ssn.encode('utf8'))
                if sheet == 'DI':
                    ssn = atachDIGBTags(ss, pName, ps, var, vals, None, None)
                    # заменяем фрагмент текста, содержащий переменную, на фрагмент с привязкой
                    s = s.replace(ss, ssn.encode('utf8'))
                if sheet == 'Valve1':
                    ssn = atachValve1GBTags(ss, pName, ps, var, vals, None, None)
                    # заменяем фрагмент текста, содержащий переменную, на фрагмент с привязкой
                    s = s.replace(ss, ssn.encode('utf8'))
                if sheet == 'PUMP':
                    ssn = atachPumpGBTags(ss, pName, ps, var, sheets, None, None)
                    # заменяем фрагмент текста, содержащий переменную, на фрагмент с привязкой
                    s = s.replace(ss, ssn.encode('utf8'))

            # обходим все записи файла с Output переменными
            for ss in gbVarsO:
                # извлекаем значение МА
                ma = ss[ss.find('MA/') + 3:ss.find('\\O/')]
                # извлекаем имя переменной
                var = ss[ss.find('\\O/') + 3:ss.find('  (')]
                # извлекаем позицию graphic block'а по его МА
                ps = gb[ma]

                # если позиция блока записи с позицией записи Excel не совпадает -> идем дальше
                if ps != xlsps:
                    continue

                # иначе извлекаем значения переменных и работаем
                vals = tags[ps]

                # обрабатываем тег Faceplate'а
                if sheet == 'AI':
                    ssn = atachAIGBTags(ss, pName, ps, var, vals, maMax, header + ver)
                    maMax += 2
                    s = s.replace(ss, ssn.encode('utf8'))
                if sheet == 'DI':
                    ssn = atachDIGBTags(ss, pName, ps, var, vals, maMax, header + ver)
                    maMax += 1
                    s = s.replace(ss, ssn.encode('utf8'))
                if sheet == 'Valve1':
                    ssn = atachValve1GBTags(ss, pName, ps, var, vals, maMax, header + ver)
                    maMax += 1
                    s = s.replace(ss, ssn.encode('utf8'))
                if sheet == 'PUMP':
                    ssn = atachPumpGBTags(ss, pName, ps, var, sheets, maMax, header + ver)
                    fr = sheets['PUMP'][ps][3]
                    if fr == '+':
                        maMax += 5
                    else:
                        maMax += 4
                    s = s.replace(ss, ssn.encode('utf8'))

    # возвращаем зашифрованную концовку документа на место
    s += '\n' + endText

    # пишем новый текст в новый файл
    f = open(path + '!new_' + dispFName, 'w')
    f.write(s)
    f.close()
    return 'Файл "!new_' + dispFName + '" создан'

# ---> обрабатывает файл PDA и изменяет в необходимых переменных тип доступа
def changePDA():
    # имя файла PDA
    # pdaFName = raw_input('Введите имя файла импорта PDA: ')
    # if dispFName == '':
    #     return 'Имя файла не введено'
    pdaFName = 'PDA.imp'

    try:
        f = open(path + pdaFName, 'r')
        s = f.read()
        f.close()
    except IOError:
        _log.write(str(
            datetime.now()) + ' - Файл импорта PDA "' + pdaFName + '" не найден')
        return 'Ошибка имени файла PDA'

    # извлекаем имена graphic block'ов
    regx = re.compile(r'\\IMPORT\\.*\n.*\n.*\n.*\n.*\n\)')
    tags = regx.findall(s)

    iot = ['EMAN', 'SL', 'SH','ELL', 'EL', 'EH', 'EHH', 'VLL', 'VL', 'VH', 'VHH', 'VMAN', 'DBL', 'dbl', 'Rem', 'Mode',
           'MV', 'SV', 'P', 'I', 'D', 'Mode_1']
    ot = ['MStart', 'MStop', 'MotoRes']

    for t in tags:
        name = t[t.find('.') + 1: t.find('  (')]
        for io in iot:
            if name == io:
                # изменяем тип доступа к необходимым переменным
                tt = t.replace('mode = "INPUT" ,', 'mode = "IN_OUT" ,')
                # заменяем фрагмент текста, содержащий переменную, на фрагмент с привязкой
                s = s.replace(t, tt)
        for o in ot:
            if name == o:
                # изменяем тип доступа к необходимым переменным
                tt = t.replace('mode = "INPUT" ,', 'mode = "OUTPUT" ,')
                # заменяем фрагмент текста, содержащий переменную, на фрагмент с привязкой
                s = s.replace(t, tt)

    # пишем новый текст в новый файл
    f = open(path + '!new_' + pdaFName, 'w')
    f.write(s)
    f.close()
    return 'Файл "!new_' + pdaFName + '" создан'


# ----------------------------------------------------------------------------------------------------------------------
#                                             Основной код
# ----------------------------------------------------------------------------------------------------------------------
import xlrd
import re
from datetime import datetime

# path = raw_input('Введите путь к проекту (если файлы в другой папке): ')
# if path == '':
#     path = ''

# открываем файл Excel, с которого будем читать конфигурацию
rb = xlrd.open_workbook(path + srcXLS, formatting_info=True)

# Обходим все листы и формируем словарь вида {'имя листа' : {'имя тега' : [параметры тега]}}
sheets = {}
for sh in rb.sheet_names():
    sheet = rb.sheet_by_name(sh)
    # считываем все строки листа
    allRows = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]

    # формируем словарь тегов листа вида {'имя тега' : [параметры тега]}
    tags = {}
    if sh == u'AI':
        for row in allRows:
            if row[1] == u'Позиция':
                continue
            if row[1] == '':
                break
            # Для AI:
            # Excel  values  Описание
            # 1      -       Позиция
            # 2      0       Описание
            # 3      1       Единизы изм-я
            # 4      2       Формат отображения
            # 5      3       Индекс в массиве
            # 6      4       Имя массива
            values = [row[2], row[3], row[4], int(row[5]), row[6]]
            tags[row[1]] = values
    if sh == u'DI':
        for row in allRows:
            if row[1] == u'Позиция':
                continue
            if row[1] == '':
                break
            # Для Задвижек:
            # Excel  values  Описание
            # 1      -       Позиция
            # 2      0       Описание
            # 3      1       Индекс в массиве
            # 4      2       Имя массива
            # 5      3       Цвет в сработке
            # 6      4       Цвет не в сработке
            values = [row[2], int(row[3]), row[4], int(row[5]), int(row[6])]
            tags[row[1]] = values
    if sh == u'PUMP':
        for row in allRows:
            if row[1] == u'Позиция':
                continue
            if row[1] == '':
                break
            # Для Насосов:
            # Excel  values  Описание
            # 1      -       Позиция
            # 2      0       Описание
            # 3      1       Индекс в массиве
            # 4      2       Имя массива
            # 5      3       Наличие ПЧ
            # 6      4       Позиция регулируемой величины (при наличии ПЧ)
            values = [row[2], int(row[3]), row[4], row[5], row[6]]
            tags[row[1]] = values
    if sh == u'Valve1':
        for row in allRows:
            if row[1] == u'Позиция':
                continue
            if row[1] == '':
                break
            # Для Задвижек:
            # Excel  values  Описание
            # 1      -       Позиция
            # 2      0       Описание
            # 3      1       Индекс в массиве
            # 4      2       Имя массива
            values = [row[2], int(row[3]), row[4]]
            tags[row[1]] = values

    sheets[sh] = tags

# изменяем тип доступа к переменным в файле PDA
print changePDA()

# генерируем файл коннекторов
print createConnectors(_pName, sheets)

# генерируем файл констант
print createConstants(_pName, sheets)

# привязываем graphic block'и дисплея к тегам
print atachTagsInDisplay(sheets)

_log.write(str(datetime.now()) + ' - Job is done!')
_log.close()
print '\nJob is done!'

# ----------------------------------------------------------------------------------------------------------------------
#                                              Для справки
# ----------------------------------------------------------------------------------------------------------------------

# так можно пройти все ячейки всех строк
# for row in allRows:
#     for cell in row:
#         if cell == u'Позиция датчика':
#             cell2 = cell.encode('cp1251') + '\n'
#             print type(cell2), cell2

# print (row_values + '\n').decode('ascii')
# f.write(row)
