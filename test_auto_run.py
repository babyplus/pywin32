import os
import win32api
import win32con
import time
import win32clipboard
import sys


def open_program(cmd):
    os.popen(cmd)
    time.sleep(5)


def press_key(x):
    dict_x = {
        'up': 38,
        'down': 40,
        'left': 37,
        'right': 39,
        'space': 32,
        'enter': 13,
        'tab': 9
    }
    dict_x.update({'F{}'.format(x): x+111 for x in range(1, 13)})
    dict_x.update({chr(x): x for x in range(65, 91)})
    dict_x.update({chr(x).lower(): x for x in range(65, 91)})
    win32api.keybd_event(dict_x.get(x), 0, 0, 0)
    win32api.keybd_event(dict_x.get(x), 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.5)


def ctrl_x(key_w):
    k = ord(key_w.upper())
    win32api.keybd_event(17, 0, 0, 0)
    win32api.keybd_event(k, 0, 0, 0)
    win32api.keybd_event(k, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(1)


def alt_x(key_w):
    win32api.keybd_event(18, 0, 0, 0)
    press_key(key_w)
    win32api.keybd_event(18, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(1)


def str_copy(text):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(text)
    win32clipboard.CloseClipboard()


if __name__ == '__main__':
    p = '''
    "G:/Users/babyplus/Documents/Spirent/TestCenter 4.63/Results/Spirent_port1port2_2017-01-10_17-09-54/2544-Tput_2017-01-10_17-11-18/2544-Tput-Summary-2_2017-01-10_17-11-18.db"
    '''
    p = sys.argv[1]
    my_path = 'C:/Users/Public/Desktop/Spirent TestCenter Results Reporter 4.63.lnk'
    file = sys.argv[2]
    file_name = file
    # 打开程序
    open_program(my_path)
    # 打开数据库
    ctrl_x('O')
    str_copy(p)
    ctrl_x('V')
    press_key('enter')
    # 保存表格
    ctrl_x('s')
    press_key('down')
    press_key('down')
    press_key('right')
    press_key('down')
    press_key('space')
    press_key('down')
    press_key('space')
    press_key('tab')
    press_key('enter')
    str_copy(file_name)
    ctrl_x('V')
    press_key('enter')
    # 关闭程序
    alt_x('F4')
