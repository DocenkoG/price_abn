# -*- coding: UTF-8 -*-
import os
import os.path
import logging
import logging.config
import sys
import configparser
import time
import shutil
import csv


def convert_csv2csv(cfg):
    inFfileName = cfg.get('basic', 'filename_in')
    outFfileNameRUR = cfg.get('basic', 'filename_out_RUR')
    outFfileNameEUR = cfg.get('basic', 'filename_out_EUR')
    outFfileNameUSD = cfg.get('basic', 'filename_out_USD')
    # inFile  = open( inFfileName,  'r', newline='', encoding='CP1251', errors='replace')
    inFile = open(inFfileName, 'r', encoding='UTF-8', errors='replace')
    outFileRUR = open(outFfileNameRUR, 'w', newline='')
    outFileEUR = open(outFfileNameEUR, 'w', newline='')
    outFileUSD = open(outFfileNameUSD, 'w', newline='')

    outFields = cfg.options('cols_out')
    csvReader = csv.DictReader(inFile, delimiter=';')
    csvWriterRUR = csv.DictWriter(outFileRUR, fieldnames=cfg.options('cols_out'))
    csvWriterEUR = csv.DictWriter(outFileEUR, fieldnames=cfg.options('cols_out'))
    csvWriterUSD = csv.DictWriter(outFileUSD, fieldnames=cfg.options('cols_out'))

    print(csvReader.fieldnames)

    csvWriterRUR.writeheader()
    csvWriterEUR.writeheader()
    csvWriterUSD.writeheader()
    recOut = {}
    for recIn in csvReader:
        for outColName in outFields:
            shablon = cfg.get('cols_out', outColName)
            for key in csvReader.fieldnames:
                if shablon.find(key) >= 0:
                    shablon = shablon.replace(key, recIn[key])
            if outColName in ('закупка', 'продажа'):
                if shablon.find('Звоните') >= 0:
                    shablon = '0.1'
            recOut[outColName] = shablon
        if recOut['валюта'] == 'RUR' :
            csvWriterRUR.writerow(recOut)
        elif recOut['валюта'] == 'EUR' :
            csvWriterEUR.writerow(recOut)
        elif recOut['валюта'] == 'USD' :
            csvWriterUSD.writerow(recOut)
        else :
            log.error('нераспознана валюта "%s" для товара "%s"', recOut['валюта'], recOut['код производителя'] )
    log.info('Обработано ' + str(csvReader.line_num) + 'строк.')
    inFile.close()
    outFileRUR.close()
    outFileEUR.close()
    outFileUSD.close()


def wait_for_window(self, timeout=2):
    time.sleep(round(timeout / 1000))
    wh_now = self.driver.window_handles
    wh_then = self.vars["window_handles"]
    if len(wh_now) > len(wh_then):
        return set(wh_now).difference(set(wh_then)).pop()


def download(cfg):
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.action_chains import ActionChains
    from selenium.webdriver.support import expected_conditions
    from selenium.webdriver.support.wait import WebDriverWait
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
    from selenium.webdriver.remote.remote_connection import LOGGER
    LOGGER.setLevel(logging.WARNING)

    retCode = False
    filename_new = cfg.get('download', 'filename_new')
    filename_old = cfg.get('download', 'filename_old')
    login = cfg.get('download', 'login')
    password = cfg.get('download', 'password')
    url_lk = cfg.get('download', 'url_lk')
    url_file = cfg.get('download', 'url_file')

    download_path = os.path.join(os.getcwd(), 'tmp')
    if not os.path.exists(download_path):
        os.mkdir(download_path)

    for fName in os.listdir(download_path):
        os.remove(os.path.join(download_path, fName))
    dir_befo_download = set(os.listdir(download_path))

    if os.path.exists('geckodriver.log'): os.remove('geckodriver.log')
    try:
        ffprofile = webdriver.FirefoxProfile()
        ffprofile.set_preference("browser.download.dir", download_path)
        ffprofile.set_preference("browser.download.folderList", 2)
        ffprofile.set_preference("browser.helperApps.neverAsk.saveToDisk",
                                 ",application/octet-stream" +
                                 ",application/vnd.ms-excel" +
                                 ",application/vnd.msexcel" +
                                 ",application/x-excel" +
                                 ",application/x-msexcel" +
                                 ",application/zip" +
                                 ",application/xls" +
                                 ",application/vnd.ms-excel" +
                                 ",application/vnd.ms-excel.addin.macroenabled.12" +
                                 ",application/vnd.ms-excel.sheet.macroenabled.12" +
                                 ",application/vnd.ms-excel.template.macroenabled.12" +
                                 ",application/vnd.ms-excelsheet.binary.macroenabled.12" +
                                 ",application/vnd.ms-fontobject" +
                                 ",application/vnd.ms-htmlhelp" +
                                 ",application/vnd.ms-ims" +
                                 ",application/vnd.ms-lrm" +
                                 ",application/vnd.ms-officetheme" +
                                 ",application/vnd.ms-pki.seccat" +
                                 ",application/vnd.ms-pki.stl" +
                                 ",application/vnd.ms-word.document.macroenabled.12" +
                                 ",application/vnd.ms-word.template.macroenabed.12" +
                                 ",application/vnd.ms-works" +
                                 ",application/vnd.ms-wpl" +
                                 ",application/vnd.ms-xpsdocument" +
                                 ",application/vnd.openofficeorg.extension" +
                                 ",application/vnd.openxmformats-officedocument.wordprocessingml.document" +
                                 ",application/vnd.openxmlformats-officedocument.presentationml.presentation" +
                                 ",application/vnd.openxmlformats-officedocument.presentationml.slide" +
                                 ",application/vnd.openxmlformats-officedocument.presentationml.slideshw" +
                                 ",application/vnd.openxmlformats-officedocument.presentationml.template" +
                                 ",application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" +
                                 ",application/vnd.openxmlformats-officedocument.spreadsheetml.template" +
                                 ",application/vnd.openxmlformats-officedocument.wordprocessingml.template" +
                                 ",application/x-ms-application" +
                                 ",application/x-ms-wmd" +
                                 ",application/x-ms-wmz" +
                                 ",application/x-ms-xbap" +
                                 ",application/x-msaccess" +
                                 ",application/x-msbinder" +
                                 ",application/x-mscardfile" +
                                 ",application/x-msclip" +
                                 ",application/x-msdownload" +
                                 ",application/x-msmediaview" +
                                 ",application/x-msmetafile" +
                                 ",application/x-mspublisher" +
                                 ",application/x-msschedule" +
                                 ",application/x-msterminal" +
                                 ",application/x-mswrite" +
                                 ",application/xml" +
                                 ",application/xml-dtd" +
                                 ",application/xop+xml" +
                                 ",application/xslt+xml" +
                                 ",application/xspf+xml" +
                                 ",application/xv+xml" +
                                 ",application/excel")
        if os.name == 'posix':
            driver = webdriver.Firefox(ffprofile,
                                       executable_path=r'geckodriver')  # , executable_path=r'/usr/local/bin/geckodriver')
        elif os.name == 'nt':
            driver = webdriver.Firefox(ffprofile)
        driver.implicitly_wait(33)
        driver.set_page_load_timeout(33)

        driver.get(url_lk)
        time.sleep(1)
        driver.find_element(By.ID, "login").click()
        driver.find_element(By.ID, "login").send_keys(login)
        time.sleep(1)
        driver.find_element(By.ID, "pass").click()
        time.sleep(1)
        driver.find_element(By.ID, "pass").send_keys(password)
        time.sleep(1)
        driver.find_element(By.NAME, "Submit").click()
        time.sleep(1)
        driver.find_element(By.XPATH, "//a[contains(text(),\'Прайс-лист\')]").click()
        time.sleep(1)
        driver.find_element(By.ID, "fileModeCsv").click()
        time.sleep(1)
        driver.find_element(By.ID, "exportBrands").click()
        time.sleep(1)
        driver.find_element(By.XPATH, "//input[@value=\'Сгенерировать\']").click()
        time.sleep(120)

    except Exception as e:
        log.debug('Exception: <' + str(e) + '>')

    driver.quit()
    dir_afte_download = set(os.listdir(download_path))
    new_files = list(dir_afte_download.difference(dir_befo_download))
    print(new_files)
    if len(new_files) == 0:
        log.error('Не удалось скачать файл прайса ')
        return False
    elif len(new_files) > 1:
        log.error('Скачалось несколько файлов. Надо разбираться ...')
        return False
    else:
        new_file = new_files[0]  # загружен ровно один файл.
        new_ext = os.path.splitext(new_file)[-1].lower()
        DnewFile = os.path.join(download_path, new_file)
        new_file_date = os.path.getmtime(DnewFile)
        log.info('Скачанный файл ' + new_file + ' имеет дату ' + time.strftime("%Y-%m-%d %H:%M:%S",
                                                                               time.localtime(new_file_date)))

        print(new_ext)
        if new_ext in ('.xls', '.xlsx', '.xlsb', '.xlsm', '.csv'):
            if os.path.exists(filename_new) and os.path.exists(filename_old):
                os.remove(filename_old)
                os.rename(filename_new, filename_old)
            if os.path.exists(filename_new):
                os.rename(filename_new, filename_old)
            shutil.copy2(DnewFile, filename_new)
            return True


def config_read(cfgFName):
    cfg = configparser.ConfigParser(inline_comment_prefixes=('#'))
    if os.path.exists('private.cfg'):
        cfg.read('private.cfg', encoding='utf-8')
    if os.path.exists(cfgFName):
        cfg.read(cfgFName, encoding='utf-8')
    else:
        log.debug('Нет файла конфигурации ' + cfgFName)
    return cfg


def is_file_fresh(fileName, qty_days):
    qty_seconds = qty_days * 24 * 60 * 60
    if os.path.exists(fileName):
        price_datetime = os.path.getmtime(fileName)
    else:
        log.error('Не найден файл  ' + fileName)
        return False

    if price_datetime + qty_seconds < time.time():
        file_age = round((time.time() - price_datetime) / 24 / 60 / 60)
        log.error(
            'Файл "' + fileName + '" устарел!  Допустимый период ' + str(qty_days) + ' дней, а ему ' + str(file_age))
        return False
    else:
        return True


def make_loger():
    global log
    logging.config.fileConfig('logging.cfg')
    log = logging.getLogger('logFile')


def main(dealerName):
    """ Обработка прайсов выполняется согласно файлов конфигурации.
    Для этого в текущей папке должны быть файлы конфигурации, описывающие
    свойства файла и правила обработки. По одному конфигу на каждый
    прайс или раздел прайса со своими правилами обработки
    """
    make_loger()
    log.info('~~~~~~~~~~~~~~~~~~  ' + dealerName)

    rc_download = False
    if  os.path.exists('getting.cfg'):
        cfg = config_read('getting.cfg')
        if cfg.has_section('download'):
            rc_download = download(cfg)

    for cfgFName in os.listdir("."):
        if cfgFName.startswith("cfg") and cfgFName.endswith(".cfg"):
            log.info('----------------------- Processing '+cfgFName )
            cfg = config_read(cfgFName)
            filename_in = cfg.get('basic','filename_in')
            if rc_download==True or is_file_fresh( filename_in, int(cfg.get('basic','срок годности'))):
                convert_csv2csv(cfg)


if __name__ == '__main__':
    myName = os.path.basename(os.path.splitext(sys.argv[0])[0])
    mydir = os.path.dirname(sys.argv[0])
    print(mydir, myName)
    main(myName)