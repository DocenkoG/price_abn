# -*- coding: UTF-8 -*-
import os
import os.path
import logging
import logging.config
import io
import sys
import configparser
import time
import abn_downloader
import abn_converter
import shutil

global log
global myname



def make_loger():
    global log
    logging.config.fileConfig('logging.cfg')
    log = logging.getLogger('logFile')
#    log.debug('test debug message from abn')
#    log.info( 'test info message from abn')
#    log.warn( 'test warn message from abn')
#    log.error('test error message from abn')
#    log.critical('test critical message')



def main( ):
    global  myname
    global  mydir
   
    make_loger()
    log.info('------------  '+myname +'  ------------')

    if  abn_downloader.download( myname ) :
        abn_converter.convert2csv( myname )
        shutil.copy2( myname + '.csv', 'c://AV_PROM/prices/' + myname +'/'+ myname + '.csv')


if __name__ == '__main__':
    global  myname
    global  mydir
    myname   = os.path.basename(os.path.splitext(sys.argv[0])[0])
    mydir    = os.path.dirname (sys.argv[0])
    if ('' != mydir) : os.chdir(mydir)
    main( )
 
#os.system(r'c:\prices\_scripts\remove_tmp_profiles.cmd')