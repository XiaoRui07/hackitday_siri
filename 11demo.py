import pygame
import sys
from glob import glob
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import qrcode
import random
import threading
import time
import json
from http.server import SimpleHTTPRequestHandler, HTTPServer
import webbrowser
from pdf2image import convert_from_path
import os
from PIL import Image as PILImage, ImageTk, ImageFilter
from matplotlib import pyplot as plt
import subprocess
import pyautogui
import pickle
import google_auth_oauthlib.flow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from tkinter import ttk
# 初始化 Tkinter
root = tk.Tk()
root.title("高級 PPT 播放器")

# PPT 相關變數
slides = []
index = 0
running = False
vote_results = {}
chat_messages = []
start_time = time.time()
time_setting = 0
button_delete = True


detect_used = False

# 啟動 Pygame 視窗
pygame.init()

info = pygame.display.Info()
screen_width, screen_height = info.current_w, info.current_h
print(screen_width, screen_height)
screen = pygame.display.set_mode((screen_width, screen_height))
clock = pygame.time.Clock()
font = pygame.font.Font("C:/Windows/Fonts/msjh.ttc", 40)
timer_initiate = False
timer = font.render("0%", True, (255, 255, 255))

max_cols = screen_width // 150  # 計算每行最多 150px 寬的縮圖
max_rows = screen_height // 150  # 計算每列最多 150px 高的縮圖
thumbnail_width = screen_width // max_cols - 10  # 確保有間距
thumbnail_height = screen_height // max_rows - 10
thumbnail_margin = 10  # 間距
counter = 0
visual_start_time = time.time()

pygame.mixer.init()
alarm_sound = pygame.mixer.Sound("C:\\Users\\ray22\\Desktop\\works\\pptpro\\OIIAOIIA CAT but in 4K Not Actually.mp3")

# 設定 PDF 檔案與解析度
#c:\Users\ray22\Downloads\會考準備經驗分享.pdf"C:\Users\ray22\Downloads\資訊學科能力競賽市賽心得.pdf"
#"C:\\Users\\ray22\\Downloads\\2100_1.pptx.pdf""C:\Users\ray22\Downloads\資訊科技學習歷程.pdf"
PDF_FILE = "C:\\Users\\ray22\\Desktop\\works\\pptpro\\這是什麼，為何如此重要.pdf"
def get_pdf_page_size(pdf_path):
    """使用 pdfinfo 取得 PDF 頁面尺寸（點，1 點 = 1/72 英吋）"""
    try:
        result = subprocess.run(
            ["pdfinfo", pdf_path], capture_output=True, text=True, check=True
        )
        output_lines = result.stdout.split("\n")
        for line in output_lines:
            if line.startswith("Page size:"):
                parts = line.split(":")[1].strip().split()
                width = float(parts[0])  # 單位是 pt (點)
                height = float(parts[2])  # 單位是 pt (點)
                return width, height
    except Exception as e:
        print(f"取得 PDF 尺寸失敗: {e}")
        return None, None
    
page_width_pt, page_height_pt = get_pdf_page_size(PDF_FILE)
if page_width_pt and page_height_pt:
    # 計算 DPI，使 PDF 符合螢幕大小
    screen_ppi = 96  # 通常 Windows 螢幕為 96 PPI，Mac 可能為 110+ PPI
    dpi_x = screen_width / (page_width_pt / 72)
    dpi_y = screen_height / (page_height_pt / 72)
    print(min(dpi_x, dpi_y))
    dpii = int(min(dpi_x, dpi_y))  # 取較小值，避免超出螢幕
    print(f"自適應 DPI: {dpii}")

    images = convert_from_path(PDF_FILE, dpi=dpii+1, use_pdftocairo=True)
else:
    print("無法讀取 PDF 頁面大小，使用預設 DPI")
    images = convert_from_path(PDF_FILE, dpi=96, use_pdftocairo=True)
PAGE_INDEX = 0  # 初始顯示第 1 



# 假設圖片的原始比例 (從第一張圖片獲取)
original_width, original_height = images[0].size  # 取得圖片原始大小
aspect_ratio = original_width / original_height   # 計算寬高比例 W:H

# 設定間距
thumbnail_margin = 10  # 設定縮圖間距

# 找到最適合的行數和列數
best_cols = 1
best_rows = 1
max_thumbnail_width = 0
max_thumbnail_height = 0
open_trans = False


for cols in range(1, screen_width // 100):  # 最少 1 列，最多能放的列數
    # 計算對應的行數
    rows = len(images) // cols + (1 if len(images) % cols else 0)
    
    # 計算縮圖大小（確保不超出螢幕）
    temp_width = (screen_width - (cols + 1) * thumbnail_margin) // cols
    temp_height = int(temp_width / aspect_ratio)
    
    # 如果超過螢幕高度，縮小
    if rows * (temp_height + thumbnail_margin) > screen_height:
        temp_height = (screen_height - (rows + 1) * thumbnail_margin) // rows
        temp_width = int(temp_height * aspect_ratio)  # 重新計算寬度

    # 更新最佳數值（選擇能最大化填滿螢幕的組合）
    if temp_width > max_thumbnail_width and temp_height > max_thumbnail_height:
        best_cols = cols
        best_rows = rows
        max_thumbnail_width = temp_width
        max_thumbnail_height = temp_height

# 最終使用的縮圖大小
thumbnail_width = max_thumbnail_width
thumbnail_height = max_thumbnail_height

print(f"計算後的縮圖大小: {thumbnail_width}x{thumbnail_height}")
print(f"最佳排版: {best_cols} 列, {best_rows} 行")

#apisetting
def get_credentials():
    creds = None
    # 檢查 token 是否存在
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # 如果沒有有效的憑證，讓用戶登錄
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            try:
                flow = google_auth_oauthlib.flow.InstalledAppFlow.from_client_secrets_file()
                creds = flow.run_local_server(port=0)
            except Exception as e:
                print(f"json:{e}")
        # 保存憑證
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds
def get_sheet_data(spreadsheet_id, range_name):
    creds = get_credentials()
    try:
        service = build('sheets', 'v4', credentials=creds)
    except Exception as e:
        print(f"Error building the service: {e}")
        return None
    
    sheet = service.spreadsheets()
    try:
        # result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    except Exception as e:
        print(f"Error reading data from sheet: {e}")
        
    return result.get('values', [])



def pil_to_surface(pil_image, size=None):
    if size:
        pil_image = pil_image.resize(size)  # 縮小圖片
    if pil_image.mode != "RGB":
        pil_image = pil_image.convert("RGB")  # 轉換為 RGB 格式
    return pygame.image.fromstring(pil_image.tobytes(), pil_image.size, "RGB")

have_filter = False
def update_slide():
    """更新當前幻燈片"""
    global counter, timer_initiate, have_filter,surface,open_trans
    if images and 0 <= counter < len(images):
        try:
            screen.fill((255, 255, 255))  # 清空畫面
            if have_filter:
                new_img = images[counter].filter(ImageFilter.GaussianBlur(10)) 
            else:
                new_img = images[counter]
            surface = pil_to_surface(new_img)
            if open_trans:
                invisable()
            screen.blit(surface, (0, 0))  # 顯示當前頁面
            if timer_initiate:
                draw_timer()
            pygame.display.flip()
        except Exception as e:
            print(f"更新幻燈片時發生錯誤: {e}")
    else:
        print(f"錯誤：counter={counter} 超出範圍或 images 為空！")


def next_slide():

    movement = random.randint(0,2)
    if not movement:
        messagebox.showerror("錯誤", "免費試用已到期，請儲值3290解鎖高級版!")
    else:
        global counter,visual_start_time
        visual_start_time = time.time()
        if counter < len(images) - 1:
            counter += 1
            update_slide()

def prev_slide():

    movement = random.randint(0,2)
    if not movement:
        messagebox.showerror("錯誤", "免費試用已到期，請儲值3290解鎖高級版!")
    else:
        global counter,visual_start_time
        visual_start_time = time.time()
        if counter > 0:
            counter -= 1
            update_slide()

def draw_lottery():
    candidate_input = simpledialog.askstring("抽籤名單", "請輸入名字（用逗號分隔）")
    if not candidate_input:
        return
    choices = []
    candidates = candidate_input.split(",")
    if candidates:
        for i in range(int(candidates[0]), int(candidates[1])):
            choices.append(i) 
        winner = random.choice(choices)
        messagebox.showinfo("抽籤結果", f"恭喜 {winner} 被選中！")
    else:
        messagebox.showinfo("抽籤結果", "沒有候選人，無法抽籤！")

# ========== 可視化計時器 ==========
def start_timer():
    global timer_initiate, start_time, time_setting, timer_label, button_delete

    time_setting = simpledialog.askstring("時間設定", "輸入計時時間(秒)")
    
    if not time_setting or not time_setting.isdigit():  # 確保輸入為數字
        return
    
    time_setting = int(time_setting)
    start_time = time.time()
    timer_initiate = True
    button_delete = True
    # 如果 `timer_label` 尚未建立，則建立 Label
    if 'timer_label' not in globals():
        timer_label = tk.Label(root, text="剩餘時間: 0 秒", font=("Arial", 16))
        timer_label.pack()

def draw_timer():
    global timer_label, start_time, time_setting,root,button_delete
    
    elapsed_time = time.time() - start_time
    max_time = int(time_setting)
    progress = min(elapsed_time / max_time, 1)
    color = (0, 255, 0) if progress < 0.5 else (255, 255, 0) if progress < 0.8 else (255, 0, 0)
    timer_label.config(text=f"剩餘時間: {max(0, round(max_time - elapsed_time, 1))} 秒")
    pygame.draw.rect(screen, color, (0, screen_height - 20, int(screen_width * progress), 20))
    if max(0, round(max_time - elapsed_time, 1)) == 0 and button_delete:
        button_delete = False
        alarm_sound.play()
        button1 = tk.Button(root, text="關閉計時器", command=lambda:close_timer(button1))
        button1.pack()
    return elapsed_time
def close_timer(button):
    global button_delete
    alarm_sound.stop() 
    button.destroy()

#透明度
def invisable():
    global surface,visual_start_time
    current_time = time.time()
    minus = (current_time - visual_start_time)/1
    surface.set_alpha(max(255-10*minus,0))
# ==========快選頁面============
thumbnails = [pil_to_surface(img, (thumbnail_width, thumbnail_height)) for img in images]

# 計算顯示縮圖的位置
def get_thumbnail_position(index, cols):
    """計算縮圖位置"""
    row = index // cols
    col = index % cols
    x = col * (thumbnail_width + thumbnail_margin)
    y = row * (thumbnail_height + thumbnail_margin)
    return x, y


# 顯示縮圖並加上邊框
def display_thumbnails():
    """顯示所有縮圖並加上邊框"""
    screen.fill((0, 0, 0))  # 清空畫面，避免殘影
    cols = max(1, screen_width // (thumbnail_width + thumbnail_margin))  # 計算每行最大縮圖數
    for index, thumbnail in enumerate(thumbnails):
        x, y = get_thumbnail_position(index, cols)
        # 畫黑色邊框
        pygame.draw.rect(screen, (0, 0, 0), (x - 2, y - 2, thumbnail_width + 4, thumbnail_height + 4), 3)
        # 顯示縮圖
        screen.blit(thumbnail, (x, y))

    pygame.display.flip()  # 更新畫面

show_thumbnails = False
def selection():
    movement = random.randint(0,2)
    if not movement:
        ad_play()
    else:
        global counter, show_thumbnails
        show_thumbnails = True
        display_thumbnails()
        print(len(images))
        page_select =[f"頁面 {i}"for i in range(1,len(images)+1)]
        quick_window = tk.Toplevel(root)
        quick_window.title("快選頁面")
        quick_window.geometry("250x300")
        tk.Label(quick_window, text="請選擇頁面:", font=("Arial", 14)).pack(pady=10)
        listbox = tk.Listbox(quick_window, height=10)
        for page in page_select:
            listbox.insert(tk.END, page)
            listbox.pack(pady=10)

        select_button = tk.Button(quick_window, text="選擇", 
        command=lambda: selected_option(quick_window,listbox))
        select_button.pack(pady=5)

def selected_option(quick_window,listbox):
    global counter, timer_initiate, show_thumbnails
    page_index = listbox.curselection()
    print(page_index)
    if page_index:
        counter =page_index[0]
        quick_window.destroy()
    show_thumbnails = False      
    update_slide()

# ========== 放大鏡 ==========
points = []  
bigger= False
detection = False

def set_detection():
    choose_movement = random.randint(0,2)
    if not choose_movement:
        ad_play()
    else:
        global detection
        detection = True

def detect_frame():
    global images, detection, points, bigger,counter, screen
    detection = True
    zoomed_image = pil_to_surface(images[counter])
    new_x, new_y = 0, 0
    closeure = 0
    while 1:    
        for event in pygame.event.get():
            if event.type == pygame.MOUSEBUTTONDOWN and not bigger:
                points.append(event.pos)
                print(points)
                if len(points) == 2:
                    closeure = 1
                    # 取得兩點座標
                    x1, y1 = points[0]
                    x2, y2 = points[1]
                    min_x, max_x = min(x1, x2), max(x1, x2)
                    min_y, max_y = min(y1, y2), max(y1, y2)
                    width, height = max_x - min_x, max_y - min_y

                # 計算等比例放大倍率
                    scale = min(screen_width / width, screen_height / height)
                    new_width, new_height = int(width * scale), int(height * scale)

                # 計算置中座標
                    new_x = (screen_width - new_width) // 2
                    new_y = (screen_height - new_height) // 2
                    print(new_width, new_height)
                    pygame.display.flip()
                # 擷取選取區域並放大
                    selected_region = screen.subsurface(pygame.Rect(min_x, min_y, width, height))
                    zoomed_image = pygame.transform.scale(selected_region, (new_width, new_height))
                    bigger= True
                    points.clear() 
        if closeure:
            print(1)
            break
                    

    return zoomed_image,new_x,new_y
    
def display_zoom(zoom, new_x, new_y):
    global screen
    if zoom:  
        screen.blit(zoom, (new_x, new_y))
        pygame.display.flip()

#廣告==========================================
ad1 ="C:\\Users\\ray22\\Desktop\\works\\pptpro\\无标题视频——使用Clipchamp制作.mp4"
ad2 ="C:\\Users\\ray22\\Desktop\\works\\pptpro\\感冒用思思.mp4"
ad3 ="C:\\Users\\ray22\\Desktop\\works\\pptpro\\貓戰.mp4"
def ad_play():
    display_choose = random.randint(0,2)
    if display_choose == 0:
        os.startfile(ad1)
    elif display_choose == 1:
        os.startfile(ad2)
    else:
        os.startfile(ad3)   
# ========== 投票系統 ==========
def start_vote():
    move = random.randint(0,2)
    if not move:
        ad_play()
    else:
        url = "https://docs.google.com/forms/d/e/1FAIpQLSfj_CFn5s4DExnPwJOXNVop5SGY12CO5-U6pA0WVEn6_1LCKw/viewform"
        qr = qrcode.make(url)
        qr.save("vote_qr.png")

    # 顯示 QR Code
        qr_window = tk.Toplevel()
        qr_window.title("投票 QR Code")
        qr_img = PILImage.open("vote_qr.png")
        qr_img = ImageTk.PhotoImage(qr_img)
        qr_label = tk.Label(qr_window, image=qr_img)
        qr_label.image = qr_img
        qr_label.pack()

        tk.Button(qr_window, text="關閉", command=qr_window.destroy).pack()


#=======彈幕================================================================================================
damn = []  # 存放彈幕的列表
damn_lock = threading.Lock()  # 確保執行緒安全
displayed_answers = set()

def get_new_answers():
    global displayed_answers,damn
    all_values = damn
    new_answers = [row for row in all_values if row not in displayed_answers]
    displayed_answers.update(row for row in new_answers)  # 記錄已顯示過的答案
    return new_answers

outload = []  # 存放彈幕的列表
def chatting(text):
    """ 在畫面右側隨機位置新增彈幕，確保正確的字典結構 """
    global outload

    y_position = random.randint(0, screen_height//3)  # 避免太靠上下邊界
    with damn_lock:
        outload.append({"text": text, "x": screen_width, "y": y_position})  # 確保是字典格式

def draw_danmu():
    """ 繪製彈幕並向左移動 """
    global outload, damn_lock
    for danmu in outload[:]:
        if isinstance(danmu, dict) and "text" in danmu:
            text_surface = font.render(danmu["text"], True, (255, 255, 255))
            screen.blit(text_surface, (danmu["x"], danmu["y"]))
            danmu["x"] -= 8 # 彈幕向左移動
            pygame.display.flip()  # 更新畫面
    outload = [danmu for danmu in outload if danmu["x"] + text_surface.get_width() > 0]  # 移除已經移出螢幕的彈幕
    
                

    


def apply_effects():
    global counter, screen, screen_width, screen_height
    keys = pygame.key.get_pressed()
    if keys[pygame.K_e]:  
        pygame.draw.circle(screen, (255, 0, 0), (screen_width // 2, screen_height // 2), 100)  # 加上紅色圓形
        root.lower()
    if keys[pygame.K_r]:  
        screen.fill((random.randint(0, 255), random.randint(0, 255), random.randint(0, 255)))  # 隨機背景色
        root.lower()
    if keys[pygame.K_t]:  # 按 T 旋轉畫面
        root.lower()
        rotate_screen()
    if keys[pygame.K_SPACE]:  # 按 Y 重置畫面
        counter = random.randint(0, len(images) - 1)
        update_slide()
    if keys[pygame.K_a]:  # 按 A 開啟透明度
        for i in range(200):
            rotate_screen()
            toggle_cursor_drift()
            counter = random.randint(0, len(images) - 1)
            update_slide()

        messagebox.showerror("錯誤", "程序已崩潰，將關閉運行程式!")
        pygame.quit()
        root.quit()
#滑鼠飄移
cursor_drift = False  # 開關控制
drift_speed = 25  # 飄移速度（可以調整）
def toggle_cursor_drift():
    global cursor_drift
    cursor_drift = not cursor_drift
    print(f"游標飄移 {'啟動' if cursor_drift else '關閉'}")
# ========== 螢幕旋轉功能 ==========
rotation_angle = 0

def rotate_screen():
    global screen, rotation_angle

    # 更新旋轉角度（每次 +90°，最多 360°）
    rotation_angle = (rotation_angle + random.randint(0,360)) 

    # 取得當前畫面快照
    screen_copy = pygame.display.get_surface().copy()

    # 旋轉畫面
    rotated_screen = pygame.transform.rotate(screen_copy, rotation_angle)

    # 根據旋轉後的大小重新調整視窗
    new_rect = rotated_screen.get_rect(center=(screen_width // 2, screen_height // 2))

    # 重新繪製畫面
    screen.fill((0, 0, 0))  # 清空畫面
    screen.blit(rotated_screen, new_rect.topleft)
    pygame.display.flip()

damnuu = False
# ========== Pygame 主迴圈 ==========
def run_pygame():
    global running,detect_used,zoomed_i,new_a,new_b,detection,bigger,cursor_drift, have_filter,damnuu,data,damn,visual_start_time,open_trans
    last_fetch_time = time.time()
    running = True
    while running:
        keys = pygame.key.get_pressed()
        if keys[pygame.K_u]:
            root.wm_attributes("-topmost", True) 
            root.wm_attributes("-topmost", False) 
        if keys[pygame.K_ESCAPE]:
            pygame.quit()
            root.quit()
        if keys[pygame.K_f]:
            have_filter = True
        if keys[pygame.K_d]:
            visual_start_time = time.time()
            open_trans = True
            update_slide()
            

        #     detection = False
        #     detect_used =False
        #     bigger = False
        #     update_slide()
        if show_thumbnails:
            display_thumbnails()
        elif detection:
            if not detect_used:
                zoomed_i,new_a,new_b = detect_frame()
                detect_used = True 
            display_zoom(zoomed_i,new_a,new_b)
        else:
            update_slide()
        draw_danmu()
        # 用你的試算表 ID 和範圍來替換這些值

    # 每 5 秒檢查一次新答案
        if time.time() - last_fetch_time > 10:
            spreadsheet_id = '177pEuDlQmdqzxAsemXHhXPDg0Nhqtdrtubtoe6NfBcI'
            range_name = "彈幕留言"  # 可以自訂範圍
            data = get_sheet_data(spreadsheet_id, range_name)
            damn = [row[1] for row in data]
            new_answers = get_new_answers()
            for ans in new_answers:
                chatting(ans)  # 把新答案加入彈幕
            last_fetch_time = time.time()
            damnuu = True
        #if(damnuu):
        apply_effects()

        pygame.display.flip()
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                running = False
            if event.type == pygame.MOUSEBUTTONDOWN:
                if event.button == 3:
                    detection = False
                    detect_used =False
                    bigger = False
                    have_filter = False
                    visual_start_time = time.time()
                    open_trans = False
                    update_slide()
                elif event.button == 2:  # 滑鼠中鍵開關飄移
                    toggle_cursor_drift()
                
        if cursor_drift:
            # 取得當前滑鼠位置
            x, y = pygame.mouse.get_pos()
            # 隨機往四個方向之一小幅度移動
            x += random.choice([-drift_speed, 0, drift_speed])
            if x<0:
                x = 0   
            elif x>screen_width:
                x = screen_width
            y += random.choice([-drift_speed, 0, drift_speed])
            if y<0:
                y = 0   
            elif y>screen_height:
                y = screen_height
            pygame.mouse.set_pos(x, y)        
    pygame.quit()
    root.quit()

# ========== Tkinter 介面 ==========
# tk.Button(root, text="載入 PPT", command=convert_ppt_to_images).pack()
tk.Button(root, text="上一頁", command=prev_slide).pack()
tk.Button(root, text="下一頁", command=next_slide).pack()
tk.Button(root, text="投放彈幕", command=start_vote).pack()
tk.Button(root, text="抽籤", command=draw_lottery).pack()
tk.Button(root, text="計時器", command=start_timer).pack()
tk.Button(root, text="快選頁面", command=selection).pack()
tk.Button(root, text="放大鏡", command=lambda: set_detection()).pack()
# 啟動 Pygame 視窗的執行緒
threading.Thread(target=run_pygame, daemon=True).start()
root.mainloop()
