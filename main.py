from ctypes import *
import win32gui
import win32ui
import win32con
import win32api
import time
import winsound
from PIL import Image
from PIL import ImageFile
import imagehash
import datetime
import pytesseract
import cv2
import numpy as np
from matplotlib import pyplot as plt
# 注册DD DLL，64位python用64位，32位用32位，具体看DD说明文件。
# 测试用免安装版。
# 用哪个就调用哪个的dll文件。
dd_dll = windll.LoadLibrary('DD85590.64.dll')

# DD虚拟码，可以用DD内置函数转换。
vk = {'5': 205, 'c': 503, 'n': 506, 'z': 501, '3': 203, '1': 201, 'd': 403, '0': 210, 'l': 409, '8': 208, 'w': 302,
        'u': 307, '4': 204, 'e': 303, '[': 311, 'f': 404, 'y': 306, 'x': 502, 'g': 405, 'v': 504, 'r': 304, 'i': 308,
        'a': 401, 'm': 507, 'h': 406, '.': 509, ',': 508, ']': 312, '/': 510, '6': 206, '2': 202, 'b': 505, 'k': 408,
        '7': 207, 'q': 301, "'": 411, '\\': 313, 'j': 407, '`': 200, '9': 209, 'p': 310, 'o': 309, 't': 305, '-': 211,
        '=': 212, 's': 402, ';': 410}
# 需要组合shift的按键。
vk2 = {'"': "'", '#': '3', ')': '0', '^': '6', '?': '/', '>': '.', '<': ',', '+': '=', '*': '8', '&': '7', '{': '[', '_': '-',
        '|': '\\', '~': '`', ':': ';', '$': '4', '}': ']', '%': '5', '@': '2', '!': '1', '(': '9'}

#截屏
def catch_screen(savepath):
    # grab a handle to the main desktop window
    hdesktop = win32gui.GetDesktopWindow()
 
 
    # determine the size of all monitors in pixels
    # here to change your pixs
    width = 2560#//win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
    height = 1600#//win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
    left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
    top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
 
 
    # create a device context
    desktop_dc = win32gui.GetWindowDC(hdesktop)
    img_dc = win32ui.CreateDCFromHandle(desktop_dc)
 
    # create a memory based device context
    mem_dc = img_dc.CreateCompatibleDC()
 
    # create a bitmap object
    screenshot = win32ui.CreateBitmap()
    screenshot.CreateCompatibleBitmap(img_dc, width, height)
    mem_dc.SelectObject(screenshot)
 
 
    # copy the screen into our memory device context
    mem_dc.BitBlt((0, 0), (width, height), img_dc, (left, top),win32con.SRCCOPY)
 
 
    # save the bitmap to a file
    screenshot.SaveBitmapFile(mem_dc, savepath)
    # free our objects
    mem_dc.DeleteDC()
    win32gui.DeleteObject(screenshot.GetHandle())
    pass

#切割图片有用部分
def cut_pic(region_path,save_path,x,y,w,h):
    im = Image.open(region_path)
    # 图片的宽度和高度
    img_size = im.size
    #print("图片宽度和高度分别是{}".format(img_size))
    '''
    裁剪：传入一个元组作为参数
    元组里的元素分别是：（距离图片左边界距离x， 距离图片上边界距离y，距离图片左边界距离+裁剪框宽度x+w，距离图片上边界距离+裁剪框高度y+h）
    '''
    region = im.crop((x, y, x+w, y+h))
    region.save(save_path)
    pass

#比较两个部分是否相同
def compare_image_with_hash(image_file_name_1, image_file_name_2, max_dif=0):
        """
        max_dif: 允许最大hash差值, 越小越精确,最小为0
        推荐使用
        """
        ImageFile.LOAD_TRUNCATED_IMAGES = True
        hash_1 = None
        hash_2 = None
        with open(image_file_name_1, 'rb') as fp:
            hash_1 = imagehash.average_hash(Image.open(fp))
        with open(image_file_name_2, 'rb') as fp:
            hash_2 = imagehash.average_hash(Image.open(fp))
        dif = hash_1 - hash_2
        if dif < 0:
            dif = -dif
        if dif <= max_dif:
            return True
        else:
            return False
        
def compare_F_pic():
    img_savepath=r"C:\Users\busker\Desktop\兴趣\钓鱼\PIC\fishing_now.png"
    catch_screen(img_savepath)
    region_path=r"C:\Users\busker\Desktop\兴趣\钓鱼\PIC\fishing_now.png"
    save_path=r"C:\Users\busker\Desktop\兴趣\钓鱼\DEAL_PIC\fishing_cut.png"
    cut_pic(region_path,save_path,1497,980,35,35)
    pic_mode=r"C:\Users\busker\Desktop\兴趣\钓鱼\DEAL_PIC\fishing_mode.png"
    pic_now=r"C:\Users\busker\Desktop\兴趣\钓鱼\DEAL_PIC\fishing_cut.png"
    flag=compare_image_with_hash(pic_mode,pic_now, max_dif=6)
    return flag

#判断是否是普通鱼
def compare_29_equal():
    time.sleep(0.3)
    region=r"C:\Users\busker\Desktop\兴趣\钓鱼\PIC\fishing_now_2.png"
    save=r"C:\Users\busker\Desktop\兴趣\钓鱼\DEAL_PIC\time_cut.png"
    catch_screen(region)
    cut_pic(region,save,1228,1288,46,45)
    img = Image.open(save).convert('LA')
    img.save(save)
    text=pytesseract.image_to_string(Image.open(save),lang='eng')
    try:
        text=int(text)
        print(text)
    except Exception:
        print("未识别出")
        return False
    else:
        if text==29:
            return True
        else:
            print("不等于29")
            return False
    pass

def find_sigh():
    img = cv2.imread('final_f_time_cut.png',0)
    print(img.shape)
    img2 = img.copy()
    template = cv2.imread('final_model.png',0)
    w, h = template.shape[::-1]

    # All the 6 methods for comparison in a list
    img = img2.copy()
    method = eval('cv2.TM_CCOEFF')

        # Apply template Matching
    res = cv2.matchTemplate(img,template,method)
        #    print(res.shape)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)#找到最大值和最小值
        # If the method is TM_SQDIFF or TM_SQDIFF_NORMED, take minimum
    top_left = max_loc
    bottom_right = (top_left[0] + w, top_left[1] + h)
    print("x,y="+str(top_left)+","+str(bottom_right))
    mid_x=top_left[0]+7
    mid_y=top_left[1]+7
    if mid_y>8:
        return -1
    else:
        return mid_x
    
        
def final_f_time():
    region=r"C:\Users\busker\Documents\ocr\fishing_now_2.png"
    catch_screen(region)
    save=r"C:\Users\busker\Documents\ocr\final_f_time_cut.png"
    cut_pic(region,save,973,1198,390,53)
    position_x=find_sigh()
    if position_x<0:
        return 1
    else:
        return 0
    
def Fishing():
    a=404  #键码
    t1=0.2   #等待时间1
    t2=2.2  #等待时间2
    flag_29=False
    flag_final=0
    time.sleep(3)
    for item in range(0,50):
        time.sleep(3)
        dd_dll.DD_key(210,1)
        dd_dll.DD_key(210,2)#按0键抛线
        winsound.Beep(800,200)
        flag=False
        while flag==False:
            time.sleep(t1)
            flag=compare_F_pic()
            #print("未出现F键提示")
            pass
        dd_dll.DD_key(a,1)
        dd_dll.DD_key(a,2)
        winsound.Beep(1000,200)
        flag_29=compare_29_equal()
        if flag_29==True:
            dd_dll.DD_key(603,1)
            dd_dll.DD_key(603,2)
        else:
            flag_final=final_f_time()
            while flag_final==0:
                #time.sleep(0.1)
                flag_final=final_f_time()
            if flag_final==1:
                dd_dll.DD_key(a,1)
                dd_dll.DD_key(a,2)
                winsound.Beep(1000,200)
            time.sleep(3)
        pass
    pass  
