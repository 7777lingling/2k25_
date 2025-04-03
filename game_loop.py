import os
import time
import ctypes
import win32api
import win32con
import win32gui
import cv2
import numpy as np
from PIL import ImageGrab

PUL = ctypes.POINTER(ctypes.c_ulong)

class KeyBdInput(ctypes.Structure):
    _fields_ = [("wVk", ctypes.c_ushort), ("wScan", ctypes.c_ushort),
                ("dwFlags", ctypes.c_ulong), ("time", ctypes.c_ulong),
                ("dwExtraInfo", PUL)]

class HardwareInput(ctypes.Structure):
    _fields_ = [("uMsg", ctypes.c_ulong), ("wParamL", ctypes.c_short),
                ("wParamH", ctypes.c_ushort)]

class MouseInput(ctypes.Structure):
    _fields_ = [("dx", ctypes.c_long), ("dy", ctypes.c_long),
                ("mouseData", ctypes.c_ulong), ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong), ("dwExtraInfo", PUL)]

class Input_I(ctypes.Union):
    _fields_ = [("ki", KeyBdInput), ("mi", MouseInput), ("hi", HardwareInput)]

class Input(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong), ("ii", Input_I)]

class GameLoop:
    def __init__(self):
        self.is_running = False
        self.window_name = "NBA 2K25"
        self.game_hwnd = None
        self.KEYS = {
            "RIGHT": ord('D'), "LEFT": ord('A'), "SPACE": win32con.VK_SPACE,
            "E": ord('E'), "S": ord('S'), "W": ord('W'), "ESC": win32con.VK_ESCAPE
        }
        folder1 = "image1"  # 遊戲主要流程圖片
        folder2 = "image2"  # 遊戲進行時的動作圖片
        self.paths = {
            # 遊戲主要流程圖片
            "mycareer": os.path.join(folder1, "MyCAREER.png"),
            "myteam": os.path.join(folder1, "MyTEAM.png"),
            "new_content": os.path.join(folder1, "全新內容.png"),
            "daily_reward": os.path.join(folder1, "每日獎勵.png"),
            "domination": os.path.join(folder1, "稱霸賽主頁.png"),
            "domination_btn": os.path.join(folder1, "稱霸賽按鈕.png"),
            "select": os.path.join(folder1, "選擇.png"),
            "stars": os.path.join(folder1, "三星.png"),
            "match": os.path.join(folder1, "比賽.png"),  # 新增比賽圖片
            # 遊戲進行時的動作圖片
            "forward": os.path.join(folder2, "前進.png"),
            "pause": os.path.join(folder2, "暫停.png"),
            "continue": os.path.join(folder2, "繼續.png"),
        }
        self.search_count = 0  # 添加搜尋次數計數器

    def find_game_window(self):
        print(f"尋找遊戲視窗: {self.window_name}")
        self.game_hwnd = win32gui.FindWindow(None, self.window_name)
        if self.game_hwnd and win32gui.IsWindow(self.game_hwnd):
            print(f"找到遊戲視窗，句柄: {self.game_hwnd}")
            return True
        print("未找到遊戲視窗")
        return False

    def press_and_release(self, key):
        if not self.game_hwnd or not win32gui.IsWindow(self.game_hwnd):
            print("視窗無效，重新尋找...")
            if not self.find_game_window(): return
        if win32gui.GetWindowLong(self.game_hwnd, win32con.GWL_STYLE) & win32con.WS_MINIMIZE:
            print("視窗最小化，無法操作")
            return
        scan_code = win32api.MapVirtualKey(key, 0)
        win32gui.SetForegroundWindow(self.game_hwnd)
        time.sleep(0.1)
        win32api.keybd_event(key, scan_code, 0, 0)
        time.sleep(0.1)
        win32api.keybd_event(key, scan_code, win32con.KEYEVENTF_KEYUP, 0)

    def get_window_rect(self):
        try:
            return win32gui.GetWindowRect(self.game_hwnd) if self.game_hwnd else None
        except:
            return None

    def detect_image(self, path, threshold=0.7):
        if not os.path.exists(path):
            print(f"找不到圖片: {path}")
            return False
        rect = self.get_window_rect()
        if not rect:
            print("無法獲取視窗區域")
            return False
        screenshot = ImageGrab.grab(bbox=rect)
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        template = cv2.imread(path)
        if template is None:
            print(f"無法讀取圖片: {path}")
            return False
        result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, _ = cv2.minMaxLoc(result)
        print(f"圖片匹配度: {max_val:.3f} (閾值: {threshold})")
        return max_val > threshold

    def check_three_stars(self):
        print("檢查三星...")
        result = self.detect_image(self.paths["stars"], threshold=0.9)
        if result:
            print("找到三星！")
            return True
        print("未找到三星")
        return False

    def press_key_and_check_stars(self, key, times):
        for i in range(times):
            print(f"第 {i+1} 次按下按鍵")
            self.press_and_release(key)
            time.sleep(0.5)
            if self.check_three_stars():
                return True
        return False

    def handle_main_images(self):
        # 優先順序檢查主要圖片並執行對應操作
        if self.detect_image(self.paths["new_content"]):
            print("找到全新內容，按下E鍵")
            for _ in range(2):
                self.press_and_release(self.KEYS["E"])
                time.sleep(0.5)
            return True

        if self.detect_image(self.paths["domination_btn"]):
            print("找到稱霸賽按鈕，按下空白鍵")
            self.press_and_release(self.KEYS["SPACE"])
            self.in_domination = True
            time.sleep(1)
            if self.detect_image(self.paths["select"]):
                print("找到選擇按鈕，按下空白鍵")
                self.press_and_release(self.KEYS["SPACE"])
                time.sleep(1)
                self.handle_three_stars_search()
            return True

        if self.detect_image(self.paths["domination"], threshold=0.4):
            print("找到稱霸賽主頁，按下S鍵")
            for i in range(5):
                print(f"按下第 {i+1} 次S鍵")
                self.press_and_release(self.KEYS["S"])
                time.sleep(0.5)
            return True

        if self.detect_image(self.paths["mycareer"]):
            print("找到 MyCAREER，按下D鍵")
            self.press_and_release(self.KEYS["RIGHT"])
            time.sleep(1)
            return True

        if self.detect_image(self.paths["myteam"]):
            print("找到 MyTEAM，按下空白鍵")
            self.press_and_release(self.KEYS["SPACE"])
            self.in_myteam = True
            time.sleep(1)
            return True

        if self.detect_image(self.paths["daily_reward"]):
            print("找到每日獎勵，按下空白鍵")
            self.press_and_release(self.KEYS["SPACE"])
            time.sleep(3)
            print("再次按下空白鍵")
            self.press_and_release(self.KEYS["SPACE"])
            return True

        return False

    def handle_game_buttons(self):
        if self.detect_image(self.paths["forward"]):
            print("找到前進按鈕，按下空白鍵")
            self.press_and_release(self.KEYS["SPACE"])
            time.sleep(1)
            return True
        elif self.detect_image(self.paths["pause"]):
            print("找到暫停按鈕，按下空白鍵")
            self.press_and_release(self.KEYS["SPACE"])
            time.sleep(1)
            return True
        elif self.detect_image(self.paths["continue"]):
            print("找到繼續按鈕，按下空白鍵")
            self.press_and_release(self.KEYS["SPACE"])
            self.in_domination = False
            time.sleep(1)
            return True
        elif self.check_three_stars():
            print("找到三星按鈕，進入三星搜尋")
            self.handle_three_stars_search()  # 直接進入三星搜尋
            return True
        print("尋找圖片...")
        return False

    def handle_three_stars_search(self):
        while self.is_running:  # 加入循環確保持續搜尋
            self.search_count += 1
            print(f"第 {self.search_count} 次搜尋三星...")
            
            if self.search_count > 2:
                print("搜尋次數超過2次，按ESC退出")
                self.press_and_release(self.KEYS["ESC"])
                time.sleep(0.5)
                print("按D鍵往右移")
                self.press_and_release(self.KEYS["RIGHT"])
                time.sleep(0.5)
                print("按空白鍵確認")
                self.press_and_release(self.KEYS["SPACE"])
                self.search_count = 0
            
            print("開始查找三星...")
            print("按下A鍵")
            self.press_and_release(self.KEYS["LEFT"])
            print("按下W鍵")
            for i in range(4):
                print(f"第 {i+1} 次按W")
                self.press_and_release(self.KEYS["W"])
                time.sleep(0.5)
            
            if self.press_key_and_check_stars(self.KEYS["S"], 5):
                print("按S鍵找到三星，開始進入遊戲")
                self.search_count = 0
                self.trigger_game_start()
                continue  # 找到三星後繼續循環
            
            print("按S鍵未找到三星，按D鍵往右移")
            self.press_and_release(self.KEYS["RIGHT"])
            time.sleep(1)
            
            if self.check_three_stars():
                print("按D鍵後找到三星，開始進入遊戲")
                self.search_count = 0
                self.trigger_game_start()
                continue  # 找到三星後繼續循環
            
            time.sleep(0.5)  # 短暫延遲後繼續循環

    def trigger_game_start(self):
        print("開始進入遊戲流程...")
        # 第一次按空白鍵
        print("按下第一次空白鍵")
        self.press_and_release(self.KEYS["SPACE"])
        time.sleep(0.5)
        
        # 按兩次S鍵
        for i in range(2):
            print(f"按下第 {i+1} 次S鍵")
            self.press_and_release(self.KEYS["S"])
            time.sleep(0.5)
        
        # 再按一次空白鍵
        print("按下第二次空白鍵")
        self.press_and_release(self.KEYS["SPACE"])
        time.sleep(0.5)
        
        # 最後再按一次空白鍵
        print("按下第三次空白鍵")
        self.press_and_release(self.KEYS["SPACE"])

        # 進入遊戲循環
        print("進入遊戲循環...")
        while self.is_running:
            if self.handle_game_buttons():
                continue
            time.sleep(0.5)

    def start(self):
        if not self.find_game_window():
            print("找不到遊戲視窗，程式退出")
            return
        print("開始執行自動化程序...")
        self.is_running = True
        try:
            self.main_loop()
        except KeyboardInterrupt:
            print("\n使用者中止程式")
            self.is_running = False
        except Exception as e:
            print(f"發生錯誤: {str(e)}")
            self.is_running = False

    def stop(self):
        self.is_running = False

    def main_loop(self):
        print("進入主循環...")
        self.in_myteam = False
        self.in_domination = False

        while self.is_running:
            if not win32gui.IsWindow(self.game_hwnd):
                print("重新查找遊戲視窗...")
                if not self.find_game_window(): break

            # 檢查主要圖片
            if self.handle_main_images():
                continue

            if not self.in_myteam:
                print("尋找主要畫面...")
            time.sleep(0.5)

if __name__ == "__main__":
    game = GameLoop()
    game.start()
