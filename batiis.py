#!/usr/bin/python3
# -*- coding: utf-8 -*-
import traceback
import logging
import subprocess
import os
import csv
import re

from socket import gethostbyname, gethostname


#Thiết lập cài đặt ghi log
logging.basicConfig(
        handlers=[
            logging.FileHandler(
                'logpython.txt',
                'w',
                'utf-8',
                )],
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s',
            )

#Ghi log lỗi
def logloi():
    logging.error(traceback.format_exc())

#Khai báo các biến dùng chung
#Đường dẫn gốc
duongdangoc = os.getcwd()

#Đường dẫn file features
dsfeatures = os.path.join(duongdangoc, 'danhsachfeatures.txt')

#Đường dẫn chương trình appcmd.exe
duongdan_appcmd = os.path.join(
        os.path.normpath(r'C:\Windows\System32\inetsrv'),
        'appcmd.exe',
        )

def xuly_noidung(noidung=None, so_dong=None):
    """Xử lý nội dung để lấy tên features"""
    if noidung is None:
        logging.error('Không có nội dung tại %s' % (so_dong))
    if noidung is not None:
        logging.debug('Xử lý nội dung tại dòng %s' % (so_dong))
        try:
            noidung_chung = noidung[0]
            ten_features, tinh_trang = noidung_chung.split('|')
            ten_features = ten_features.strip()
            return ten_features
        except:
            logloi()

def caidatfeature(ten=None):
    """Cài đặt features"""
    if ten is not None:
        logging.debug('Bắt đầu cài đặt features - %s' % (ten))
        try:
            subprocess.call('dism /online /enable-feature /featurename:%s /all'
            % (ten), stdout=None)
            logging.debug('Cài đặt xong - %s' % (ten))
        except:
            logloi()

def chaylenh_cmd(duongdan_appcmd, lenh_cmd):
    lenh = ' '.join([duongdan_appcmd, lenh_cmd])
    logging.debug('Chạy lệnh %s' % (lenh))
    with subprocess.Popen(lenh, stdout=subprocess.PIPE) as ps:
        logging.debug(ps.stdout.read())

def caidat_features():
    try:
        #Tạo file danh sách features
        logging.debug('Lấy danh sách features')
        if os.path.exists(dsfeatures):
            logging.debug('Tìm thấy danh sách tại %s' % (dsfeatures))
        if not os.path.exists(dsfeatures):
            logging.debug('Tạo danh sách features tại %s' % (dsfeatures))
            try:
                with open(dsfeatures, 'w') as f:
                    subprocess.call(
                            'dism /online /get-features /format:table',
                            stdout=f,
                            )
            except:
                logloi()

        #Lấy tổng dòng trong file
        with open(dsfeatures) as file_ds:
            so_dong = sum(1 for row in file_ds)

        #Đọc dữ liệu các features cần cài
        logging.debug('Đọc dữ liệu features')
        with open(dsfeatures) as file_ds:
            doc_file = csv.reader(file_ds)
            ten_features = []
            mautimkiem = re.compile('IIS-')
            for noidung in doc_file:
                if doc_file.line_num >= 14 and\
                doc_file.line_num < (so_dong - 2):
                   ten = xuly_noidung(noidung, doc_file.line_num)
                   if mautimkiem.match(ten):
                       logging.debug('Tìm thấy features %s' % (ten))
                       ten_features.append(ten)

        #Xóa file danh sách features
        logging.debug('Xóa file danh sách features')
        try:
            os.remove(dsfeatures)
        except:
            logloi()

        #Cài đặt các features
        if len(ten_features) > 0:
            logging.debug('Cài đặt các features')
            for i in ten_features:
                caidatfeature(i)
        logging.debug('Cài đặt xong các features!')
    except:
        logloi()

def cauhinh_iis():
    try:
        #Kiểm tra chương trình appcmd.exe
        if not os.path.exists(duongdan_appcmd):
            print('Không thấy chương trình appcmd.exe')
            print('vui lòng thiết lập bằng tay!')
            input()
            quit()
        if os.path.exists(duongdan_appcmd):
            logging.debug("Bắt đầu cấu hình IIS")

            #Thiết lập Apppool .Net v4.5 Classic
            lenh_thietlap_apppool = 'set apppool ".NET v4.5 Classic" \
                                    -managedPipelineMode:Integrated \
                                    -enable32BitAppOnWin64:true \
                                    /autoStart:true'
            chaylenh_cmd(duongdan_appcmd, lenh_thietlap_apppool)

            #Thêm sites
            ip_may = gethostbyname(gethostname())
            duongdan_lanweb = os.path.normpath('C:\MasterLanTestWeb')
            if not os.path.exists(duongdan_lanweb):
                print('Chưa cài đặt phần mềm Master LanTest Website')
                input()
                quit()

            #Thêm site Lantest
            lenh_them_site = 'add site /name:"LanTest" \
                             /bindings:http://%s:80 \
                             /physicalPath:%s \
                             /serverAutoStart:true' % (ip_may, duongdan_lanweb)
            chaylenh_cmd(duongdan_appcmd, lenh_them_site)

            #Cài đặt application pools cho site
            lenh_caidat_site = 'set site /site.name:"LanTest" \
                    /[path=\'/\'].applicationPool:".NET v4.5 Classic"'
            chaylenh_cmd(duongdan_appcmd, lenh_caidat_site)

            #Chạy lại site
            lenh_tat_site = 'stop site /site.name:"LanTest"'
            lenh_chay_site = 'start site /site.name:"LanTest"'
            chaylenh_cmd(duongdan_appcmd, lenh_tat_site)
            chaylenh_cmd(duongdan_appcmd, lenh_chay_site)

            #Thêm quyền cho tài khoản IIS tại thư mục website và phần mềm
            duongdan_phanmem = os.path.normpath('C:\Program Files (x86)\Saosaigon')
            lenh_themquyen_website = '"%s" /grant IIS_IUSRS:(OI)(CI)F /T' % \
                                    (duongdan_lanweb)
            lenh_themquyen_phanmem = '"%s" /grant IIS_IUSRS:(OI)(CI)F /T' % \
                                    (duongdan_phanmem)
            chaylenh_cmd('icacls', lenh_themquyen_website)
            chaylenh_cmd('icacls', lenh_themquyen_phanmem)
    except:
        logloi()


if __name__ == '__main__':
    caidat_features()
    cauhinh_iis()
