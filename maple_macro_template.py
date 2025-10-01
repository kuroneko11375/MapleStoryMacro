#!/usr/bin/env python3
"""
楓之谷自動化腳本
作者：SchwarzeKatze_R
版本：1.0.1
"""

import warnings
warnings.filterwarnings("ignore", category=SyntaxWarning)  # 忽略語法警告

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import time
import threading
import keyboard
import win32gui
import win32con
import pydirectinput
import pyperclip
import os
import cv2
import numpy as np
from PIL import Image, ImageTk
import pyautogui
import sys
import ctypes

def is_admin():
    """檢查是否以管理員身分執行"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """以管理員身分重新啟動程序"""
    if is_admin():
        return True
    else:
        try:
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, " ".join(sys.argv), None, 1
            )
            return False
        except:
            messagebox.showerror("權限錯誤", "需要管理員權限才能正常運行，請以管理員身分執行此程序")
            return False

class MacroApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Maple Macro GUI")

        print("初始化 MacroApp...")
        
        # 初始化小地圖相關變數 (必須在最前面)
        self.minimap_var = tk.BooleanVar(value=False)  # 預設關閉小地圖
        
        # 初始化變數
        self.events = []
        self.recording = False
        self.playing = False
        self.current_loop = 0
        self.total_loops = 1
        self.hooked_hwnd = None
        self.scanning_dialog = None
        self.start_position = None  # 記錄開始位置

        # 位置偏離追蹤
        self.deviation_start_time = None
        self.last_deviation_check = 0
        self.deviation_threshold_time = 2.0
        self.is_currently_deviating = False

        # 暫停修正機制
        self.correction_pause_event = threading.Event()
        self.correction_pause_event.set()
        self.is_correcting = False

        # 技能連發設置 (完整字母表)
        self.skill_keys = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm',
                          'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z']
        self.last_skill_time = {}
        self.skill_repeat_interval = 0.05

        # 小地圖偵測設置
        self.minimap_enabled = True  # 預設啟用小地圖偵測
        self.minimap_region = None
        self.minimap_reference = None
        self.minimap_position_offset = (0, 0)
        self.minimap_path_points = []
        self.minimap_canvas = None
        self.current_minimap_pos = None
        self.replaying = False
        self.minimap_update_interval = 100  # 小地圖刷新頻率 (毫秒) - 預設100ms

        # 校正策略 / 基準腳本
        self.baseline_events = None  # 保存第一次播放時的原始腳本 (基準值)
        self.loop2_correction_threshold = 1.0  # 第二迴圈開始，嚴重偏移持續秒數門檻
        self.suppress_space_until_loop_end = False  # 校正後本迴圈抑制跳躍

        # 視窗與佈局
        self.root.geometry("560x530")
        
        # 強制置頂顯示
        self.root.attributes('-topmost', True)
        self.root.lift()
        self.root.focus_force()
        self.root.update()
        self.root.attributes('-topmost', False)
        
        self.main_frame = ttk.Frame(self.root, padding=5)
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        print("視窗佈局設定完成")

        # 建立控制元件
        self.create_widgets()

        print("控制元件建立完成")

        # 尋找遊戲視窗
        self.find_maple_window()

        # 自動啟用小地圖功能 (在UI元件創建完成後)
        self.auto_setup_minimap()

        # 延遲啟動位置更新，確保所有UI元件都已初始化
        self.root.after(1000, self.update_position)

        print("MacroApp 初始化完成")

    # ================== 自動小地圖設定 / 載入 ==================
    def auto_setup_minimap(self):
        """嘗試從設定檔載入小地圖區域 / 人物模板。若不存在則略過。"""
        try:
            config = self._load_minimap_config()
            if config:
                region = config.get('region')
                if region and len(region) == 4:
                    self.minimap_region = tuple(region)
                    # 安全地更新UI元件
                    if hasattr(self, 'minimap_status') and self.minimap_status.winfo_exists():
                        self.minimap_status.config(text=f"小地圖: 已載入 {region[2]}x{region[3]}")
                    print(f"🔁 已載入小地圖區域: {self.minimap_region}")
                # 人物模板暫不持久化（可日後擴充）
            else:
                print("ℹ️ 沒有找到小地圖設定檔，跳過自動載入")
        except Exception as e:
            print(f"❌ 自動載入小地圖設定失敗: {e}")

    def _load_minimap_config(self):
        path = os.path.join(os.path.dirname(__file__), 'minimap_config.json')
        if os.path.exists(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"讀取 minimap_config.json 失敗: {e}")
        return None

    def _save_minimap_config(self):
        try:
            data = {
                'region': list(self.minimap_region) if self.minimap_region else None,
            }
            path = os.path.join(os.path.dirname(__file__), 'minimap_config.json')
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print(f"💾 已保存小地圖設定: {path}")
        except Exception as e:
            print(f"❌ 保存小地圖設定失敗: {e}")

    def create_widgets(self):
        # 左側控制面板
        left_panel = ttk.LabelFrame(self.main_frame, text="控制面板", padding=5)
        left_panel.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        # 錄製控制
        record_frame = ttk.LabelFrame(left_panel, text="錄製控制", padding=5)
        record_frame.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        self.start_button = ttk.Button(record_frame, text="開始錄製", command=self.start_recording)
        self.start_button.grid(row=0, column=0, padx=2, pady=2)

        self.stop_button = ttk.Button(record_frame, text="停止錄製", command=self.stop_recording)
        self.stop_button.grid(row=0, column=1, padx=2, pady=2)
        self.stop_button['state'] = 'disabled'

        # 檔案操作
        file_frame = ttk.LabelFrame(left_panel, text="檔案操作", padding=5)
        file_frame.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

        ttk.Button(file_frame, text="保存腳本", command=self.save_macro).grid(row=0, column=0, padx=2, pady=2)
        ttk.Button(file_frame, text="載入腳本", command=self.load_macro).grid(row=0, column=1, padx=2, pady=2)
        ttk.Button(file_frame, text="清除暫存", command=self.clear_macro).grid(row=1, column=0, padx=2, pady=2)
        ttk.Button(file_frame, text="查看事件", command=self.debug_events).grid(row=1, column=1, padx=2, pady=2)
        ttk.Button(file_frame, text="測試連發", command=self.test_rapid_fire).grid(row=2, column=0, padx=2, pady=2)

        # 播放控制
        playback_frame = ttk.LabelFrame(left_panel, text="播放控制", padding=5)
        playback_frame.grid(row=2, column=0, padx=5, pady=5, sticky="ew")

        ttk.Label(playback_frame, text="迴圈次數:").grid(row=0, column=0, padx=2, pady=2)
        self.loop_count = ttk.Entry(playback_frame, width=5)
        self.loop_count.insert(0, "1")
        self.loop_count.grid(row=0, column=1, padx=2, pady=2)

        self.play_button = ttk.Button(playback_frame, text="開始播放", command=self.start_playback)
        self.play_button.grid(row=1, column=0, padx=2, pady=2)

        self.stop_play_button = ttk.Button(playback_frame, text="停止播放", command=self.stop_playback)
        self.stop_play_button.grid(row=1, column=1, padx=2, pady=2)

        # 右側狀態面板
        right_panel = ttk.LabelFrame(self.main_frame, text="狀態資訊", padding=5)
        right_panel.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

        # 視窗狀態
        window_frame = ttk.Frame(right_panel)
        window_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)

        self.window_status = ttk.Label(window_frame, text="視窗狀態: 尋找中...")
        self.window_status.pack(side=tk.LEFT)
        ttk.Button(window_frame, text="重新檢測", command=self.refresh_window).pack(side=tk.LEFT, padx=5)

        # 位置資訊
        self.position_label = ttk.Label(right_panel, text="角色位置: 未偵測")
        self.position_label.grid(row=1, column=0, sticky="w", padx=5, pady=5)

        # 錄製狀態
        self.recording_status = ttk.Label(right_panel, text="錄製狀態: 就緒 | 事件數: 0")
        self.recording_status.grid(row=2, column=0, sticky="w", padx=5, pady=5)

        # 播放狀態
        self.playback_status = ttk.Label(right_panel, text="播放狀態: 就緒")
        self.playback_status.grid(row=3, column=0, sticky="w", padx=5, pady=5)

        # 整合小地圖輔助控件與監控控件
        self.minimap_display_frame = ttk.LabelFrame(right_panel, text="小地圖輔助與監控", padding=5)
        self.minimap_display_frame.grid(row=4, column=0, padx=5, pady=5, sticky="ew")

        # 創建小地圖畫布
        self.minimap_canvas = tk.Canvas(self.minimap_display_frame, width=150, height=150, bg='black')
        self.minimap_canvas.grid(row=0, column=0, columnspan=3, pady=5)

        # 小地圖狀態
        self.minimap_status = ttk.Label(self.minimap_display_frame, text="小地圖: 未啟用")
        self.minimap_status.grid(row=1, column=0, columnspan=3, pady=2)

        # 小地圖輔助控件
        ttk.Checkbutton(self.minimap_display_frame, text="啟用小地圖偵測",
                        variable=self.minimap_var, command=self.toggle_minimap).grid(row=2, column=0, columnspan=3, pady=2)

        ttk.Button(self.minimap_display_frame, text="設定小地圖區域",
                   command=self.setup_minimap_region).grid(row=3, column=0, padx=2, pady=2)

        ttk.Button(self.minimap_display_frame, text="校準位置",
                   command=self.calibrate_minimap).grid(row=3, column=1, padx=2, pady=2)

        ttk.Label(self.minimap_display_frame, text="刷新(ms):").grid(row=4, column=0, padx=2, pady=2, sticky="e")
        self.minimap_interval_entry = ttk.Entry(self.minimap_display_frame, width=6)
        self.minimap_interval_entry.grid(row=4, column=1, padx=2, pady=2, sticky="w")
        self.minimap_interval_entry.insert(0, str(self.minimap_update_interval))
        ttk.Button(self.minimap_display_frame, text="套用", command=self.update_minimap_interval).grid(row=4, column=2, padx=2, pady=2)

        # 小地圖測試按鈕
        test_frame = ttk.Frame(self.minimap_display_frame)
        test_frame.grid(row=5, column=0, columnspan=3, pady=2)
        ttk.Button(test_frame, text="測試擷取", command=self.test_minimap_capture).pack(side="left", padx=2)
        ttk.Button(test_frame, text="開始監控", command=self.start_minimap_monitoring).pack(side="left", padx=2)
        ttk.Button(test_frame, text="停止監控", command=self.stop_minimap_monitoring).pack(side="left", padx=2)

        # 自動回程選項
        self.return_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(playback_frame, text="結束時回程", variable=self.return_var).grid(row=2, column=0, columnspan=2)

        # 位置驗證選項
        self.position_check_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(playback_frame, text="位置驗證", variable=self.position_check_var).grid(row=3, column=0, columnspan=2)

        # 設定框架權重
        self.main_frame.grid_columnconfigure(1, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)

    def find_maple_window(self):
        def callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                # 檢測包含指定關鍵字的視窗，並過濾掉非遊戲視窗
                if any(keyword in title for keyword in ["MapleStory", "幽靈谷"]):
                    try:
                        class_name = win32gui.GetClassName(hwnd)
                        # 排除 Discord、Chrome 等非遊戲視窗
                        if (class_name and class_name not in ["Shell_TrayWnd", "Button", "Chrome_WidgetWin_1"] 
                            and "Discord" not in title and "Chrome" not in title and "瀏覽器" not in title):
                            windows.append((title, hwnd, class_name))
                    except:
                        pass
        
        try:
            windows = []
            win32gui.EnumWindows(callback, windows)
            
            if windows:
                print("找到以下可能的視窗:")
                for title, hwnd, class_name in windows:
                    print(f"標題: {title}")
                    print(f"句柄: {hwnd}")
                    print(f"類名: {class_name}")
                    print("---")

                if len(windows) > 1:
                    choices = [f"{title} ({class_name})" for title, _, class_name in windows]
                    select_window = tk.Toplevel(self.root)
                    select_window.title("選擇遊戲視窗")
                    select_window.geometry("300x200")
                    
                    label = ttk.Label(select_window, text="請選擇正確的遊戲視窗:")
                    label.pack(pady=5)
                    
                    listbox = tk.Listbox(select_window, height=len(choices))
                    for choice in choices:
                        listbox.insert(tk.END, choice)
                    listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
                    
                    # 預設選擇第一個
                    if choices:
                        listbox.selection_set(0)
                    
                    def on_select():
                        selection = listbox.curselection()
                        if selection:
                            index = selection[0]
                            self.hooked_hwnd = windows[index][1]
                            self.window_status.config(text=f"視窗狀態: 已鎖定 ({windows[index][0]})")
                            print(f"✅ 已選擇視窗: {windows[index][0]} (句柄: {windows[index][1]})")
                            select_window.destroy()
                        else:
                            print("⚠️ 請選擇一個視窗")
                    
                    # 支援雙擊選擇
                    listbox.bind('<Double-Button-1>', lambda e: on_select())
                    
                    button_frame = ttk.Frame(select_window)
                    button_frame.pack(pady=5)
                    ttk.Button(button_frame, text="確定", command=on_select).pack(side=tk.LEFT, padx=5)
                    ttk.Button(button_frame, text="取消", command=select_window.destroy).pack(side=tk.LEFT, padx=5)
                    
                    self.root.wait_window(select_window)
                    return True
                else:
                    title, hwnd, _ = windows[0]
                    self.hooked_hwnd = hwnd
                    # 僅記錄視窗，不再附加記憶體
                    self.window_status.config(text=f"視窗狀態: 已鎖定 ({title})")
                    return True
            else:
                self.window_status.config(text="視窗狀態: 找不到遊戲視窗")
                return False
                
        except Exception as e:
            self.window_status.config(text=f"視窗狀態: 錯誤 ({str(e)})")
            return False

    def refresh_window(self):
        """手動刷新視窗檢測 (僅重新尋找視窗)"""
        def async_refresh():
            try:
                if self.find_maple_window():
                    self.root.after(0, lambda: self.window_status.config(text="視窗狀態: 已重新鎖定"))
                else:
                    self.root.after(0, lambda: messagebox.showwarning("警告", "找不到遊戲視窗"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("錯誤", f"刷新視窗時發生錯誤: {str(e)}"))
            finally:
                self.root.after(0, lambda: self.window_status.config(state='normal'))
        self.window_status.config(state='disabled')
        threading.Thread(target=async_refresh, daemon=True).start()
    
    # ====== 以小地圖追蹤替代記憶體讀取 ======
    def get_current_position(self):
        """取得目前角色位置(以小地圖像素座標表示)"""
        if not self.minimap_enabled or not self.minimap_region:
            return None, None
        try:
            px, py = self.get_minimap_player_position() if hasattr(self, 'get_minimap_player_position') else (None, None)
            if px is None or py is None:
                return None, None
            return float(px), float(py)
        except Exception:
            return None, None

    def update_position(self):
        """定期更新角色位置(小地圖)"""
        # 確保UI元件已經創建
        if not hasattr(self, 'position_label'):
            self.root.after(300, self.update_position)
            return
            
        x, y = self.get_current_position()
        if x is not None and y is not None:
            self.position_label.config(text=f"角色位置(小地圖): X={x:.0f}, Y={y:.0f}", foreground="green")
        else:
            if self.minimap_enabled and self.minimap_region:
                self.position_label.config(text="角色位置: 定位中...", foreground="blue")
            else:
                self.position_label.config(text="角色位置: 未設定小地圖", foreground="orange")
        self.root.after(300, self.update_position)

    def start_recording(self):
        """開始錄製按鍵"""
        if not self.find_maple_window():
            if not messagebox.askyesno("警告", "找不到遊戲視窗，是否繼續？"):
                return

        if not hasattr(self, 'events'):
            self.events = []

        # 記錄起始位置（小地圖像素座標）
        start_x, start_y = self.get_current_position()
        if start_x is not None and start_y is not None:
            self.start_position = {'x': start_x, 'y': start_y}
            print(f"錄製起始位置: X={start_x:.1f}, Y={start_y:.1f}")
        else:
            self.start_position = None
            print("無法獲取起始位置 (尚未啟用/設定小地圖)")

        # 清空之前的路徑記錄
        self.minimap_path_points = []

        # 如果啟用小地圖，開始記錄路徑
        if self.minimap_enabled:
            self.start_minimap_path_recording()

        self.recording = True
        self.start_button['state'] = 'disabled'
        self.stop_button['state'] = 'normal'
        self.recording_status.config(text=f"錄製狀態: 錄製中 | 事件數: {len(self.events)}")

        threading.Thread(target=self._recording_thread, daemon=True).start()

    def _recording_thread(self):
        """在背景執行錄製"""
        try:
            if self.hooked_hwnd:
                win32gui.ShowWindow(self.hooked_hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(self.hooked_hwnd)
                time.sleep(0.2)

            self.current_recorded_events = []
            last_state = set()
            key_press_times = {}  # 記錄每個按鍵的最後按下時間
            check_interval = 0.01
            last_event_time = None
            relative_time = 0
            continuous_press_interval = 0.05  # 持續按住的檢查間隔，縮短為50ms
            
            print("🎯 開始錄製 - 請在遊戲窗口中操作")
            print("⚠️ 注意：請避免在錄製期間點擊本程序界面")
            
            def check_keys():
                nonlocal last_state, relative_time, last_event_time
                current_state = set()
                current_time = time.perf_counter()
                
                # 檢查當前活動窗口是否是遊戲窗口
                try:
                    current_hwnd = win32gui.GetForegroundWindow()
                    if current_hwnd != self.hooked_hwnd:
                        # 如果不是遊戲窗口，跳過這次檢查
                        return
                except:
                    # 如果檢查失敗，繼續錄制
                    pass
                
                # 修復小鍵盤按鍵名稱，使用正確的 keyboard 庫格式
                monitored_keys = [
                    # 方向鍵
                    'left', 'right', 'up', 'down',
                    # 修飾鍵
                    'space', 'alt', 'ctrl', 'shift', 'tab', 'enter', 'backspace', 'delete',
                    'insert', 'home', 'end', 'page up', 'page down', 'esc',
                    # 完整字母表
                    'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm',
                    'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z',
                    # 數字鍵
                    '1', '2', '3', '4', '5', '6', '7', '8', '9', '0',
                    # 符號鍵
                    '-', '=', '[', ']', '\\', ';', "'", ',', '.', '/', '`',
                    # 功能鍵
                    'f1', 'f2', 'f3', 'f4', 'f5', 'f6', 'f7', 'f8', 'f9', 'f10', 'f11', 'f12',
                    # 小鍵盤 - 使用正確的 keyboard 庫格式
                    'keypad 0', 'keypad 1', 'keypad 2', 'keypad 3', 'keypad 4', 
                    'keypad 5', 'keypad 6', 'keypad 7', 'keypad 8', 'keypad 9',
                    'keypad /', 'keypad *', 'keypad -', 'keypad +', 'keypad .', 'keypad enter',
                    'num lock',
                ]
                
                # 檢測目前按下的按鍵
                for key in monitored_keys:
                    try:
                        if keyboard.is_pressed(key):
                            current_state.add(key)
                    except Exception as e:
                        if 'not mapped' not in str(e):
                            print(f"⚠️ 按鍵檢測錯誤 {key}: {e}")
                        continue
                
                # 處理持續按住的按鍵
                for key in current_state:
                    if key in last_state:
                        # 按鍵持續按住中
                        last_press_time = key_press_times.get(key, 0)
                        if current_time - last_press_time >= continuous_press_interval:
                            # 生成持續按住事件
                            key_press_times[key] = current_time
                            
                            # 計算相對時間
                            if last_event_time is None:
                                relative_time = 0
                            else:
                                time_diff = current_time - last_event_time
                                relative_time = self.current_recorded_events[-1]['time'] + time_diff if self.current_recorded_events else time_diff
                            
                            current_x, current_y = self.get_current_position()
                            event_data = {
                                'type': 'keyboard',
                                'event': key,
                                'event_type': 'hold',
                                'time': round(relative_time, 3),
                                'pressed_keys': list(current_state),
                                'position': {'x': current_x, 'y': current_y} if current_x is not None else None
                            }
                            self.current_recorded_events.append(event_data)
                            print(f"🔄 持續按住 {key} (時間: {relative_time:.3f}s, 間隔: {current_time - last_press_time:.3f}s)")
                            last_event_time = current_time
                            
                            # 更新狀態顯示
                            self.root.after(0, lambda count=len(self.current_recorded_events): self.recording_status.config(
                                text=f"錄製狀態: 錄製中 | 事件數: {count} | 🔄 持續按住 {key}"
                            ))
                    else:
                        # 新按下的按鍵
                        key_press_times[key] = current_time

                if current_state != last_state:
                    if last_event_time is None:
                        relative_time = 0
                    else:
                        time_diff = current_time - last_event_time
                        relative_time = self.current_recorded_events[-1]['time'] + time_diff if self.current_recorded_events else time_diff
                    
                    last_event_time = current_time
                    new_pressed = current_state - last_state
                    new_released = last_state - current_state
                    
                    for key in new_pressed:
                        # 獲取當前位置
                        current_x, current_y = self.get_current_position()
                        
                        event_data = {
                            'type': 'keyboard',
                            'event': key,
                            'event_type': 'down',
                            'time': round(relative_time, 3),
                            'pressed_keys': list(current_state),
                            'position': {'x': current_x, 'y': current_y} if current_x is not None else None
                        }
                        self.current_recorded_events.append(event_data)
                        if current_x is not None and current_y is not None:
                            print(f"🎯 錄製 {key} 按下 - 位置: X={current_x:.1f}, Y={current_y:.1f}")
                            # 更新狀態顯示包含位置資訊
                            self.root.after(0, lambda x=current_x, y=current_y, count=len(self.current_recorded_events): self.recording_status.config(
                                text=f"錄製狀態: 錄製中 | 事件數: {count} | 位置: X={x:.1f}, Y={y:.1f}"
                            ))
                        else:
                            print(f"⚠️  錄製 {key} 按下 - 無法獲取位置 | 所有按鍵: {list(current_state)}")
                            self.root.after(0, lambda count=len(self.current_recorded_events): self.recording_status.config(
                                text=f"錄製狀態: 錄製中 | 事件數: {count} | ⚠️ 無位置資訊"
                            ))
                    
                    for key in new_released:
                        # 獲取當前位置
                        current_x, current_y = self.get_current_position()
                        
                        event_data = {
                            'type': 'keyboard',
                            'event': key,
                            'event_type': 'up',
                            'time': round(relative_time, 3),
                            'pressed_keys': list(current_state),
                            'position': {'x': current_x, 'y': current_y} if current_x is not None else None
                        }
                        self.current_recorded_events.append(event_data)
                        if current_x is not None and current_y is not None:
                            print(f"🎯 錄製 {key} 放開 - 位置: X={current_x:.1f}, Y={current_y:.1f}")
                            # 更新狀態顯示包含位置資訊
                            self.root.after(0, lambda x=current_x, y=current_y, count=len(self.current_recorded_events): self.recording_status.config(
                                text=f"錄製狀態: 錄製中 | 事件數: {count} | 位置: X={x:.1f}, Y={y:.1f}"
                            ))
                        else:
                            print(f"⚠️ 錄製 {key} 放開 - 無法獲取位置")
                            self.root.after(0, lambda count=len(self.current_recorded_events): self.recording_status.config(
                                text=f"錄製狀態: 錄製中 | 事件數: {count} | ⚠️ 無位置資訊"
                            ))
                    
                    last_state = current_state.copy()
                    
                    self.events = self.current_recorded_events
            
                if self.recording:
                    self.root.after(int(check_interval * 1000), check_keys)
            
            check_keys()
            
            while self.recording:
                time.sleep(0.1)
                
        except Exception as e:
            messagebox.showerror("錯誤", f"錄製過程中發生錯誤: {str(e)}")
            self.events = []
        finally:
            keyboard.unhook_all()
            if self.recording:
                self.recording = False
                self.root.after(0, self.stop_recording)

    def stop_recording(self):
        """停止錄製"""
        self.recording = False
        keyboard.unhook_all()
        
        if hasattr(self, 'pressed_keys'):
            for key in self.pressed_keys:
                try:
                    pydirectinput.keyUp(key)
                except Exception as e:
                    print(f"釋放按鍵錯誤 {key}: {e}")
            self.pressed_keys.clear()
        
        self.start_button['state'] = 'normal'
        self.stop_button['state'] = 'disabled'
        
        if not hasattr(self, 'events'):
            self.events = []
        
        total_events = len(self.events)
        self.recording_status.config(text=f"錄製狀態: 已停止 | 事件數: {total_events}")

    def start_playback(self):
        if self.recording:
            messagebox.showwarning("警告", "請先停止錄製再播放")
            return
            
        if not self.events:
            messagebox.showwarning("警告", "沒有可播放的事件")
            return
        
        try:
            self.total_loops = int(self.loop_count.get())
            if self.total_loops < 1:
                raise ValueError("迴圈次數必須大於0")
        except ValueError as e:
            messagebox.showerror("錯誤", str(e))
            return

        if not self.find_maple_window():
            if not messagebox.askyesno("警告", "找不到遊戲視窗，是否繼續？"):
                return

        # 重置偏離追蹤
        self.deviation_start_time = None
        self.is_currently_deviating = False

        self.playing = True
        self.current_loop = 0
        self.play_button['state'] = 'disabled'
        self.stop_play_button['state'] = 'normal'
        # 保存基準腳本 (僅第一次播放時複製)
        if self.baseline_events is None:
            try:
                import copy
                self.baseline_events = copy.deepcopy(self.events)
            except Exception:
                self.baseline_events = list(self.events)
        threading.Thread(target=self._playback_thread, daemon=True).start()

    def _playback_thread(self):
        completed_normally = False
        # 追蹤目前按下的按鍵狀態，避免重複按下/釋放
        currently_pressed_keys = set()
        
        try:
            if self.hooked_hwnd:
                for _ in range(3):
                    try:
                        win32gui.ShowWindow(self.hooked_hwnd, win32con.SW_RESTORE)
                        win32gui.SetForegroundWindow(self.hooked_hwnd)
                        break
                    except:
                        time.sleep(0.1)
                time.sleep(0.2)
            
            self.paused_for_focus = False

            while self.playing and self.current_loop < self.total_loops:
                self.current_loop += 1
                # 迴圈開始時恢復跳躍事件
                self.suppress_space_until_loop_end = False
                self.playback_status.config(text=f"播放狀態: 執行中 (迴圈 {self.current_loop}/{self.total_loops})")
                
                last_event_time = 0
                playback_start = time.perf_counter()
                self.current_playback_step = 0  # 追蹤當前播放步驟
                
                for event_index, event in enumerate(self.events):
                    if not self.playing:
                        break
                    
                    self.current_playback_step = event_index  # 更新當前步驟
                    
                    # 檢查是否需要等待位置修正完成
                    if not self.correction_pause_event.is_set():
                        if not self.is_correcting:
                            self.playback_status.config(text="播放狀態: 暫停中 (位置修正)")
                        self.correction_pause_event.wait()  # 等待修正完成
                        if self.playing:  # 修正完成後恢復狀態顯示
                            self.playback_status.config(text=f"播放狀態: 進行中 (迴圈 {self.current_loop}/{self.total_loops})")
                    
                    if not self.check_window_focus():
                        if not self.paused_for_focus:
                            self.paused_for_focus = True
                            self.playback_status.config(text="播放狀態: 已暫停 (視窗失焦)")
                        time.sleep(0.1)
                        continue
                    elif self.paused_for_focus:
                        self.paused_for_focus = False
                        self.playback_status.config(text=f"播放狀態: 執行中 (迴圈 {self.current_loop}/{self.total_loops})")
                        playback_start = time.perf_counter() - last_event_time
                        continue

                    current_time = time.perf_counter()
                    elapsed = current_time - playback_start
                    wait_time = max(0, event['time'] - elapsed)
                    
                    if wait_time > 0.001:
                        time.sleep(wait_time)
                    
                    try:
                        if event['type'] == 'keyboard':
                            # 檢查位置（如果事件有位置信息且啟用了位置驗證）
                            if (self.position_check_var.get() and 'position' in event and 
                                event['position'] is not None):
                                current_x, current_y = self.get_current_position()
                                
                                # 如果記憶體位置無效，嘗試使用小地圖輔助
                                if (current_x is None or current_y is None) and self.minimap_enabled:
                                    minimap_offset_x, minimap_offset_y = self.get_minimap_position_offset()
                                    if minimap_offset_x is not None and minimap_offset_y is not None:
                                        # 使用小地圖偏移量計算當前位置
                                        # 已全改用小地圖座標，直接採用
                                        current_x, current_y = minimap_offset_x, minimap_offset_y
                                        print(f"🗺️ 小地圖定位: X={current_x:.1f}, Y={current_y:.1f}")
                                
                                expected_x = event['position']['x']
                                expected_y = event['position']['y']
                                
                                if current_x is not None and current_y is not None:
                                    # 過濾明顯異常的座標
                                    if (abs(current_x) > 10000 or abs(current_y) > 10000 or
                                        (abs(current_x) < 1e-30 and abs(current_y) < 1e-30)):
                                        print(f"⚠️ 忽略異常座標: X={current_x:.1f}, Y={current_y:.1f}")
                                        continue  # 跳過這次檢查
                                    
                                    # 計算位置偏差
                                    x_diff = abs(current_x - expected_x) if expected_x is not None else 0
                                    y_diff = abs(current_y - expected_y) if expected_y is not None else 0
                                    
                                    # 根據按鍵類型調整容忍度
                                    key_name = event['event']
                                    # 動態容忍度 (依小地圖尺寸調整)
                                    map_w = self.minimap_region[2] if self.minimap_region else 200
                                    map_h = self.minimap_region[3] if self.minimap_region else 150
                                    if key_name in ['space', 'shift']:  # 跳躍類按鍵
                                        tolerance_x = max(5, map_w * 0.25)
                                        tolerance_y = max(8, map_h * 0.35)
                                    elif key_name in ['left', 'right', 'up', 'down']:  # 移動類按鍵
                                        tolerance_x = max(4, map_w * 0.18)
                                        tolerance_y = max(6, map_h * 0.25)
                                    else:  # 其他技能
                                        tolerance_x = max(6, map_w * 0.30)
                                        tolerance_y = max(8, map_h * 0.40)
                                    
                                    # 檢查是否需要位置修正（更嚴格的條件）
                                    needs_major_correction = (x_diff > tolerance_x * 2 or y_diff > tolerance_y * 2)
                                    current_time = time.time()
                                    
                                    if x_diff > tolerance_x or y_diff > tolerance_y:
                                        # 開始追蹤偏離時間
                                        if not self.is_currently_deviating:
                                            self.deviation_start_time = current_time
                                            self.is_currently_deviating = True
                                            print(f"📍 開始追蹤位置偏離: X偏差={x_diff:.1f}, Y偏差={y_diff:.1f}")
                                        
                                        # 計算偏離持續時間
                                        deviation_duration = current_time - self.deviation_start_time
                                        
                                        # 動態門檻: 第1迴圈不修正 (觀察/建立基準), 第2迴圈起 >1秒才修正
                                        dyn_threshold = (self.loop2_correction_threshold if self.current_loop >= 2 else 10**9)
                                        if needs_major_correction and deviation_duration >= dyn_threshold:
                                            print(f"🚨 嚴重位置偏差持續 {deviation_duration:.1f}秒 ({key_name})")
                                            print(f"   預期 X={expected_x:.1f}, Y={expected_y:.1f}, 實際 X={current_x:.1f}, Y={current_y:.1f}")
                                            print(f"   偏差量: X={x_diff:.1f}, Y={y_diff:.1f} - 暫停腳本進行修正")
                                            
                                            # 暫停腳本進行位置修正
                                            self.pause_for_correction(current_x, current_y, expected_x, expected_y, x_diff, y_diff)
                                            
                                        else:
                                            # 偏離但未達到修正條件
                                            if needs_major_correction:
                                                if self.current_loop >= 2:
                                                    remaining_time = self.loop2_correction_threshold - deviation_duration
                                                    if remaining_time < 0: remaining_time = 0
                                                    print(f"⏳ 嚴重偏差等待修正 (還需 {remaining_time:.1f}秒): X偏差={x_diff:.1f}, Y偏差={y_diff:.1f}")
                                                    self.playback_status.config(text=f"播放狀態: 偏差等待修正 {remaining_time:.1f}s (迴圈 {self.current_loop}/{self.total_loops})")
                                                else:
                                                    self.playback_status.config(text=f"播放狀態: 基準觀察中 (迴圈 1/{self.total_loops})")
                                            else:
                                                print(f"📍 輕微位置偏差持續 {deviation_duration:.1f}秒: X偏差={x_diff:.1f}, Y偏差={y_diff:.1f}")
                                                self.playback_status.config(text=f"播放狀態: 輕微偏差 (迴圈 {self.current_loop}/{self.total_loops})")
                                    else:
                                        # 位置正常，重置偏離追蹤
                                        if self.is_currently_deviating:
                                            print("✅ 位置已恢復正常，重置偏離追蹤")
                                            self.deviation_start_time = None
                                            self.is_currently_deviating = False
                            
                            # 處理持續按住事件
                            if event['event_type'] == 'hold':
                                # 執行快速連發來模擬持續按住
                                current_key = event['event']
                                try:
                                    print(f"🔄 執行hold連發: {current_key}")
                                    # 減少連發次數，讓效果更接近實際按住
                                    for i in range(2):  # 從3次改為2次
                                        pydirectinput.keyDown(current_key)
                                        time.sleep(0.005)
                                        pydirectinput.keyUp(current_key)
                                        time.sleep(0.015)
                                    print(f"⚡ Hold連發完成: {current_key} (2次)")
                                except Exception as e:
                                    print(f"❌ Hold事件執行錯誤: {e}")
                                continue
                            
                            key_mapping = {
                                # 修飾鍵
                                'space': 'space',
                                'shift': 'shiftleft',
                                'right shift': 'shiftright',
                                'ctrl': 'ctrlleft',
                                'right ctrl': 'ctrlright',
                                'alt': 'altleft',
                                'right alt': 'altright',
                                'enter': 'enter',
                                'tab': 'tab',
                                'backspace': 'backspace',
                                'delete': 'delete',
                                'insert': 'insert',
                                'home': 'home',
                                'end': 'end',
                                'page up': 'pageup',
                                'page down': 'pagedown',
                                'esc': 'esc',
                                
                                # 功能鍵
                                'f1': 'f1', 'f2': 'f2', 'f3': 'f3', 'f4': 'f4',
                                'f5': 'f5', 'f6': 'f6', 'f7': 'f7', 'f8': 'f8',
                                'f9': 'f9', 'f10': 'f10', 'f11': 'f11', 'f12': 'f12',
                                
                                # 方向鍵
                                'left': 'left', 'right': 'right', 'up': 'up', 'down': 'down',
                                'arrow left': 'left', 'arrow right': 'right', 
                                'arrow up': 'up', 'arrow down': 'down',
                                'left arrow': 'left', 'right arrow': 'right',
                                'up arrow': 'up', 'down arrow': 'down',
                                
                                # 字母鍵 (完整字母表)
                                'a': 'a', 'b': 'b', 'c': 'c', 'd': 'd', 'e': 'e',
                                'f': 'f', 'g': 'g', 'h': 'h', 'i': 'i', 'j': 'j',
                                'k': 'k', 'l': 'l', 'm': 'm', 'n': 'n', 'o': 'o',
                                'p': 'p', 'q': 'q', 'r': 'r', 's': 's', 't': 't',
                                'u': 'u', 'v': 'v', 'w': 'w', 'x': 'x', 'y': 'y', 'z': 'z',
                                
                                # 數字鍵
                                '1': '1', '2': '2', '3': '3', '4': '4', '5': '5',
                                '6': '6', '7': '7', '8': '8', '9': '9', '0': '0',
                                
                                # 符號鍵
                                '-': '-', '=': '=', '[': '[', ']': ']', '\\': '\\',
                                ';': ';', "'": "'", ',': ',', '.': '.', '/': '/',
                                '`': '`',
                                
                                # 小鍵盤
                                'num lock': 'numlock',
                                'num 0': 'num0', 'num 1': 'num1', 'num 2': 'num2',
                                'num 3': 'num3', 'num 4': 'num4', 'num 5': 'num5',
                                'num 6': 'num6', 'num 7': 'num7', 'num 8': 'num8',
                                'num 9': 'num9',
                                'num /': 'divide', 'num *': 'multiply',
                                'num -': 'subtract', 'num +': 'add',
                                'num .': 'decimal', 'num enter': 'enter'
                            }
                            
                            current_key = key_mapping.get(event['event'], event['event'])
                            pressed_keys = [key_mapping.get(k, k) for k in event.get('pressed_keys', [])]
                            
                            # 調試：顯示原始按鍵和映射後的按鍵
                            if event['event'] in ['left', 'right', 'up', 'down'] or current_key in ['left', 'right', 'up', 'down']:
                                print(f"🎯 方向鍵調試: 原始='{event['event']}' -> 映射='{current_key}'")
                            
                            # 如果是數字鍵卻被當作方向鍵，輸出警告
                            if event['event'] in ['1', '2', '3', '4', '5', '6'] and current_key in ['left', 'right', 'up', 'down']:
                                print(f"⚠️ 警告: 數字鍵 '{event['event']}' 被錯誤映射為方向鍵 '{current_key}'")
                            
                            if event['event_type'] == 'down':
                                # 若本迴圈被標記抑制跳躍且當前為 space，直接跳過
                                if self.suppress_space_until_loop_end and current_key == 'space':
                                    print("⏭️ 抑制 space (down)")
                                else:
                                    # 調試：顯示將要執行的按鍵
                                    if current_key in ['left', 'right', 'up', 'down']:
                                        print(f"🎮 執行方向鍵: {current_key}")
                                    elif current_key in ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0']:
                                        print(f"🔢 執行數字鍵: {current_key}")
                                    
                                    # 確保按鍵沒有重複按下
                                    if current_key not in currently_pressed_keys:
                                        currently_pressed_keys.add(current_key)
                                        
                                        if current_key in self.skill_keys:
                                            self.execute_skill_with_repeat(current_key, pressed_keys)
                                        else:
                                            try:
                                                pydirectinput.keyDown(current_key)
                                            except Exception:
                                                pass
                                            for key in pressed_keys:
                                                if key != current_key and key not in currently_pressed_keys:
                                                    try:
                                                        pydirectinput.keyDown(key)
                                                        currently_pressed_keys.add(key)
                                                    except Exception:
                                                        pass
                            elif event['event_type'] == 'hold':
                                # 處理持續按住事件 - 使用有效的快速連發
                                if self.suppress_space_until_loop_end and current_key == 'space':
                                    print("⏭️ 抑制 space (hold)")
                                else:
                                    try:
                                        print(f"🔄 執行連發: {current_key}")
                                        
                                        # 減少連發次數，讓效果更接近實際按住
                                        for i in range(2):  # 從3次改為2次
                                            pydirectinput.keyDown(current_key)
                                            time.sleep(0.005)
                                            pydirectinput.keyUp(current_key)
                                            time.sleep(0.015)
                                        
                                        print(f"⚡ 連發完成: {current_key} (2次)")
                                        
                                    except Exception as e:
                                        print(f"❌ Hold事件執行錯誤: {e}")
                            else:  # event_type == 'up'
                                if self.suppress_space_until_loop_end and current_key == 'space':
                                    print("⏭️ 抑制 space (up)")
                                else:
                                    # 釋放按鍵
                                    if current_key in currently_pressed_keys:
                                        try:
                                            pydirectinput.keyUp(current_key)
                                            currently_pressed_keys.remove(current_key)
                                            print(f"🔓 釋放按鍵: {current_key}")
                                        except Exception:
                                            pass
                            
                            print(f"Playing: {event['event']} {event['event_type']}")
                            print(f"All pressed keys: {pressed_keys}")
                            
                            time.sleep(0.02)
                            
                    except Exception as e:
                        print(f"按鍵播放錯誤 {event['event']}: {str(e)}")
                    
                    last_event_time = event['time']
                
                if self.playing and self.current_loop < self.total_loops:
                    time.sleep(0.5)
            
            completed_normally = True
            
            # 清理所有按鍵狀態
            print("🧹 清理按鍵狀態...")
            for key in list(currently_pressed_keys):
                try:
                    pydirectinput.keyUp(key)
                    print(f"🔓 釋放殘留按鍵: {key}")
                except Exception:
                    pass
            currently_pressed_keys.clear()
            
            self.playing = False
            self.root.after(0, lambda: self._update_after_playback(completed_normally))
            
        except Exception as e:
            # 異常情況下也要清理按鍵狀態
            print("🧹 異常情況下清理按鍵狀態...")
            for key in list(currently_pressed_keys):
                try:
                    pydirectinput.keyUp(key)
                    print(f"🔓 釋放殘留按鍵: {key}")
                except Exception:
                    pass
            currently_pressed_keys.clear()
            
            self.playing = False
            messagebox.showerror("錯誤", f"播放過程中發生錯誤: {str(e)}")
            self.root.after(0, lambda: self._update_after_playback(False))

    def _update_after_playback(self, completed_normally):
        self.play_button['state'] = 'normal'
        self.stop_play_button['state'] = 'disabled'
        
        if completed_normally:
            self.playback_status.config(text="播放狀態: 完成")
            if self.return_var.get():
                print("執行自動回程...")
                self.return_to_town()
        else:
            self.playback_status.config(text="播放狀態: 已停止")

    def stop_playback(self):
        self.playing = False
        
        # 重置偏離追蹤和修正狀態
        self.deviation_start_time = None
        self.is_currently_deviating = False
        self.is_correcting = False
        self.correction_pause_event.set()  # 確保不會卡在暫停狀態
        
        self.playback_status.config(text="播放狀態: 停止中...")
        if self.return_var.get():
            if self.start_position is not None:
                self.return_to_start_position()
            else:
                self.return_to_town()
    
    def return_to_start_position(self):
        """自動回到錄製開始的位置"""
        try:
            if not self.start_position:
                print("沒有記錄起始位置，使用傳統回程")
                self.return_to_town()
                return
                
            current_x, current_y = self.get_current_position()
            if current_x is None or current_y is None:
                print("無法獲取當前位置，使用傳統回程")
                self.return_to_town()
                return
                
            target_x = self.start_position['x']
            target_y = self.start_position['y']
            
            print(f"自動回程: 從 X={current_x:.1f}, Y={current_y:.1f} 到 X={target_x:.1f}, Y={target_y:.1f}")
            
            # 確保視窗焦點
            if self.hooked_hwnd:
                win32gui.ShowWindow(self.hooked_hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(self.hooked_hwnd)
                time.sleep(0.5)
            
            # 計算距離和方向
            x_diff = target_x - current_x
            y_diff = target_y - current_y
            distance = (x_diff**2 + y_diff**2)**0.5
            
            if distance < 10:  # 已經很接近了
                print("已經在起始位置附近")
                self.playback_status.config(text="播放狀態: 已回到起始位置")
                return
                
            # 簡單的方向移動邏輯
            move_duration = min(distance / 100, 3.0)  # 根據距離決定移動時間，最多3秒
            
            self.playback_status.config(text="播放狀態: 自動回程中...")
            
            # 水平移動
            if abs(x_diff) > 10:
                if x_diff > 0:
                    print("向右移動")
                    pydirectinput.keyDown('right')
                    time.sleep(move_duration * 0.6)
                    pydirectinput.keyUp('right')
                else:
                    print("向左移動")
                    pydirectinput.keyDown('left')
                    time.sleep(move_duration * 0.6)
                    pydirectinput.keyUp('left')
            
            # 垂直移動（跳躍或下移）
            if abs(y_diff) > 10:
                if y_diff < 0:  # 需要向上
                    print("跳躍向上")
                    for _ in range(int(abs(y_diff) / 50) + 1):
                        pydirectinput.press('space')
                        time.sleep(0.3)
                        if not self.playing:  # 檢查是否被停止
                            break
            
            time.sleep(0.5)
            
            # 檢查是否到達目標位置
            final_x, final_y = self.get_current_position()
            if final_x is not None and final_y is not None:
                final_distance = ((target_x - final_x)**2 + (target_y - final_y)**2)**0.5
                if final_distance < 20:
                    print("成功回到起始位置")
                    self.playback_status.config(text="播放狀態: 已回到起始位置")
                else:
                    print(f"回程不完全，剩餘距離: {final_distance:.1f}")
                    self.playback_status.config(text="播放狀態: 回程完成（可能有偏差）")
            
        except Exception as e:
            print(f"自動回程失敗: {e}")
            self.playback_status.config(text="播放狀態: 回程失敗，嘗試傳統回程")
            self.return_to_town()

    def return_to_town(self):
        try:
            if not self.hooked_hwnd:
                if not self.find_maple_window():
                    messagebox.showwarning("警告", "找不到遊戲視窗，無法執行回程")
                    return False

            try:
                win32gui.ShowWindow(self.hooked_hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(self.hooked_hwnd)
                time.sleep(0.5)
            except Exception as e:
                print(f"視窗切換出錯: {e}")
                return False

            print("🏠 開始執行回程指令...")
            
            # 第一步：按 Enter 打開聊天框
            print("  1. 按 Enter 打開聊天框")
            pydirectinput.press('enter')
            time.sleep(0.3)  # 增加等待時間

            try:
                # 第二步：複製回程指令到剪貼簿
                print("  2. 複製回程指令到剪貼簿")
                pyperclip.copy('@FM')  # 範例指令，請根據實際遊戲修改
                time.sleep(0.2)
                
                # 第三步：貼上指令 (Ctrl+V)
                print("  3. 貼上指令 (Ctrl+V)")
                pydirectinput.keyDown('ctrl')
                time.sleep(0.05)
                pydirectinput.press('v')
                time.sleep(0.05)
                pydirectinput.keyUp('ctrl')
                time.sleep(0.3)  # 等待貼上完成
                
                # 第四步：按 Enter 執行指令
                print("  4. 按 Enter 執行回程指令")
                pydirectinput.press('enter')
                time.sleep(0.2)
                
                print("✅ 回程指令執行完成")

            except Exception as e:
                print(f"❌ 回程指令執行出錯: {e}")
                # 清理可能卡住的按鍵
                try:
                    pydirectinput.keyUp('ctrl')
                    pydirectinput.keyUp('shift')
                    pydirectinput.keyUp('alt')
                except:
                    pass
                return False
            finally:
                if hasattr(self, 'playback_status'):
                    self.playback_status.config(text="播放狀態: 已執行回程")

            return True

        except Exception as e:
            print(f"❌ 回程功能整體出錯: {e}")
            messagebox.showerror("錯誤", f"執行回程時發生錯誤: {str(e)}")
            return False

    def check_window_focus(self):
        try:
            if self.hooked_hwnd:
                foreground_hwnd = win32gui.GetForegroundWindow()
                return foreground_hwnd == self.hooked_hwnd
        except:
            pass
        return False

    def save_macro(self):
        if not self.events:
            messagebox.showwarning("警告", "沒有可保存的事件")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")]
        )
        if filename:
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(self.events, f, ensure_ascii=False, indent=2)
            self.recording_status.config(text=f"已保存到: {filename}")

    def load_macro(self):
        filename = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json")]
        )
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    self.events = json.load(f)
                self.recording_status.config(text=f"已載入: {filename} | 事件數: {len(self.events)}")
            except Exception as e:
                messagebox.showerror("錯誤", f"載入失敗: {str(e)}")

    def debug_events(self):
        """調試功能：顯示錄製的事件"""
        if not self.events:
            messagebox.showinfo("調試", "沒有錄製的事件")
            return
        
        debug_window = tk.Toplevel(self.root)
        debug_window.title("事件調試")
        debug_window.geometry("600x400")
        
        # 創建文本框顯示事件
        text_frame = ttk.Frame(debug_window)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        text_widget = tk.Text(text_frame, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 分析事件統計
        event_stats = {}
        hold_events = []
        
        for i, event in enumerate(self.events):
            event_type = event['event_type']
            key = event['event']
            
            if event_type not in event_stats:
                event_stats[event_type] = 0
            event_stats[event_type] += 1
            
            if event_type == 'hold':
                hold_events.append((i, event))
        
        # 顯示統計信息
        debug_text = f"=== 事件統計 ===\n"
        debug_text += f"總事件數: {len(self.events)}\n"
        for event_type, count in event_stats.items():
            debug_text += f"{event_type} 事件: {count} 個\n"
        
        debug_text += f"\n=== Hold事件詳情 ===\n"
        for idx, (event_idx, event) in enumerate(hold_events[:10]):  # 只顯示前10個
            debug_text += f"Hold #{idx+1} (索引{event_idx}): {event['event']} 在 {event['time']:.3f}s\n"
        
        if len(hold_events) > 10:
            debug_text += f"... 還有 {len(hold_events) - 10} 個 hold 事件\n"
        
        debug_text += f"\n=== 前20個事件詳情 ===\n"
        for i, event in enumerate(self.events[:20]):
            debug_text += f"{i:3d}: {event['time']:7.3f}s - {event['event']:8s} {event['event_type']:5s}\n"
        
        if len(self.events) > 20:
            debug_text += f"... 還有 {len(self.events) - 20} 個事件\n"
        
        text_widget.insert(tk.END, debug_text)
        text_widget.config(state=tk.DISABLED)

    def test_rapid_fire(self):
        """測試連發功能"""
        test_window = tk.Toplevel(self.root)
        test_window.title("連發測試")
        test_window.geometry("300x200")
        
        ttk.Label(test_window, text="選擇要測試的按鍵:").pack(pady=10)
        
        key_var = tk.StringVar(value="z")
        key_entry = ttk.Entry(test_window, textvariable=key_var, width=10)
        key_entry.pack(pady=5)
        
        def do_test():
            key = key_var.get().lower()
            print(f"開始測試連發: {key}")
            
            # 等待5秒讓用戶切換到遊戲窗口
            for i in range(5, 0, -1):
                print(f"倒數 {i} 秒...")
                time.sleep(1)
            
            print("開始連發測試!")
            
            # 測試不同的連發方式，每種之間有明顯間隔
            try:
                # 方式1: 慢速測試 - 先讓用戶看到單次按鍵效果
                print("=== 方式1: 單次按鍵測試 (5次，間隔1秒) ===")
                for i in range(5):
                    print(f"  單次按鍵 {i+1}/5")
                    pydirectinput.press(key)
                    time.sleep(1)
                print("=== 方式1: 完成 ===")
                
                time.sleep(3)  # 長間隔便於區分
                
                # 方式2: 中速連發
                print("=== 方式2: 中速連發 (持續5秒，每秒約10次) ===")
                start_time = time.time()
                count = 0
                while time.time() - start_time < 5.0:
                    pydirectinput.press(key)
                    count += 1
                    time.sleep(0.1)  # 每100ms一次
                print(f"=== 方式2: 完成 (共 {count} 次) ===")
                
                time.sleep(3)
                
                # 方式3: 快速連發
                print("=== 方式3: 快速連發 (持續5秒，每秒約50次) ===")
                start_time = time.time()
                count = 0
                while time.time() - start_time < 5.0:
                    pydirectinput.keyDown(key)
                    time.sleep(0.005)
                    pydirectinput.keyUp(key)
                    time.sleep(0.015)
                    count += 1
                print(f"=== 方式3: 完成 (共 {count} 次) ===")
                
                time.sleep(3)
                
                # 方式4: 極速連發
                print("=== 方式4: 極速連發 (持續5秒，每秒約100次) ===")
                start_time = time.time()
                count = 0
                while time.time() - start_time < 5.0:
                    pydirectinput.keyDown(key)
                    time.sleep(0.001)
                    pydirectinput.keyUp(key)
                    time.sleep(0.009)
                    count += 1
                print(f"=== 方式4: 完成 (共 {count} 次) ===")
                
                time.sleep(3)
                
                # 方式5: 持續按住
                print("=== 方式5: 持續按住不放 (5秒) ===")
                pydirectinput.keyDown(key)
                time.sleep(5.0)
                pydirectinput.keyUp(key)
                print("=== 方式5: 完成 ===")
                
                print("所有測試完成! 請告訴我哪種方式效果最好")
                
            except Exception as e:
                print(f"測試錯誤: {e}")
                # 確保釋放按鍵
                try:
                    pydirectinput.keyUp(key)
                except:
                    pass
        
        ttk.Button(test_window, text="開始測試 (3秒後)", 
                  command=lambda: threading.Thread(target=do_test, daemon=True).start()).pack(pady=10)
        
        ttk.Label(test_window, text="請在測試開始前切換到遊戲窗口", 
                 foreground="red").pack(pady=5)

    def clear_macro(self):
        if self.events:
            if messagebox.askyesno("確認", "確定要清除目前暫存的腳本嗎？"):
                self.events = []
                self.start_position = None  # 同時清除起始位置
                self.recording_status.config(text="錄製狀態: 就緒 | 事件數: 0")
        else:
            messagebox.showinfo("提示", "目前沒有暫存的腳本")
    
    def attempt_position_correction(self, current_x, current_y, expected_x, expected_y, x_diff, y_diff):
        """智能位置修正：根據位置偏差類型選擇修正策略（保守版本）"""
        try:
            correction_made = False
            
            # 更保守的Y軸偏差修正 - 避免干擾正常跳躍
            if y_diff > 400 and current_y > expected_y:  # 角色位置太高，提高觸發門檻
                print(f"🔧 嘗試向下跳躍修正 (Y偏差: {y_diff:.1f}) - 偏差過大")
                
                # 按下+alt往下跳（楓之谷的向下跳）
                pydirectinput.keyDown('down')
                time.sleep(0.05)
                pydirectinput.keyDown('alt')
                time.sleep(0.1)
                pydirectinput.keyUp('alt')
                pydirectinput.keyUp('down')
                time.sleep(0.3)  # 等待跳躍完成
                
                correction_made = True
                
            elif y_diff > 300 and current_y < expected_y:  # 角色位置太低，需要向上
                print(f"🔧 嘗試多次跳躍修正 (Y偏差: {y_diff:.1f}) - 偏差過大")
                
                # 多次跳躍以到達更高位置
                jump_count = min(int(y_diff / 150), 3)  # 根據偏差決定跳躍次數，最多3次
                for i in range(jump_count):
                    print(f"   跳躍 {i+1}/{jump_count}")
                    pydirectinput.keyDown('space')
                    time.sleep(0.08)  # 稍微縮短按鍵時間，提高響應
                    pydirectinput.keyUp('space')
                    time.sleep(0.15)  # 減少等待時間
                
                correction_made = True
            
            # 更保守的X軸偏差修正 - 只在偏差非常大時才修正
            if x_diff > 200:  # 提高X軸修正門檻
                direction = 'right' if current_x < expected_x else 'left'
                # 減少移動時間，避免過度修正
                move_time = min(x_diff / 400, 0.5)  # 最多0.5秒，移動更保守
                
                print(f"🔧 嘗試橫向移動修正 ({direction}, 時間: {move_time:.2f}秒, X偏差: {x_diff:.1f}) - 偏差過大")
                
                pydirectinput.keyDown(direction)
                time.sleep(move_time)
                pydirectinput.keyUp(direction)
                time.sleep(0.2)
                
                correction_made = True
            
            # 如果沒有進行修正，直接返回
            if not correction_made:
                print(f"📍 位置偏差在可接受範圍內 (X: {x_diff:.1f}, Y: {y_diff:.1f})，繼續執行")
                return False
            
            # 等待位置穩定
            time.sleep(0.3)  # 減少等待時間，提高響應
            
            # 驗證修正效果
            new_x, new_y = self.get_current_position()
            if new_x is not None and new_y is not None:
                new_x_diff = abs(new_x - expected_x) if expected_x is not None else 0
                new_y_diff = abs(new_y - expected_y) if expected_y is not None else 0
                
                print(f"   修正後位置: X={new_x:.1f}, Y={new_y:.1f}")
                print(f"   修正後偏差: X={new_x_diff:.1f}, Y={new_y_diff:.1f}")
                
                # 如果修正有改善，返回成功
                if new_x_diff < x_diff * 0.7 or new_y_diff < y_diff * 0.7:
                    print("✅ 位置修正有效果")
                    return True
                else:
                    print("⚠️ 位置修正效果有限")
                    
            return True  # 至少嘗試了修正
            
        except Exception as e:
            print(f"❌ 位置修正時發生錯誤: {e}")
            return False
    
    def reposition_to(self, target_x, target_y, fallback=False, max_time=3.0):
        """利用短脈衝按鍵把角色朝目標 minimap (x,y) 靠近。
        fallback=True 表示是回起始點的補救策略。
        回傳 True 表示成功接近 (在容忍內) 或顯著改善; False 表示失敗。
        """
        try:
            if target_x is None or target_y is None:
                return False
            if not self.minimap_enabled or not self.minimap_region:
                return False
            map_w = self.minimap_region[2]
            map_h = self.minimap_region[3]
            tol_x = max(5, map_w * 0.06)
            tol_y = max(6, map_h * 0.08)
            start_time = time.time()
            last_improve = start_time
            prev_dist = None
            while time.time() - start_time < max_time:
                cx, cy = self.get_current_position()
                if cx is None or cy is None:
                    time.sleep(0.1)
                    continue
                dx = target_x - cx
                dy = target_y - cy
                dist = (dx*dx + dy*dy) ** 0.5
                if prev_dist is None or dist < prev_dist - 1:
                    prev_dist = dist
                    last_improve = time.time()
                if abs(dx) <= tol_x and abs(dy) <= tol_y:
                    print(f"🎯 已到達目標附近: 現在({cx:.1f},{cy:.1f}) 目標({target_x:.1f},{target_y:.1f})")
                    return True
                if time.time() - last_improve > 1.2:
                    print("⚠️ 長時間無改善，停止精準移動")
                    return False
                # 撞牆檢測邏輯
                if self.detect_collision():
                    print("🚧 檢測到撞牆，取消當次移動")
                    return False
                # 水平 pulse
                if abs(dx) > tol_x:
                    dir_key = 'right' if dx > 0 else 'left'
                    pulse = min(0.3, max(0.1, abs(dx) / map_w * 0.25))  # 長按邏輯
                    pydirectinput.keyDown(dir_key)
                    time.sleep(pulse)
                    pydirectinput.keyUp(dir_key)
                # 垂直 pulse (假設 minimap y 向下增加)
                if abs(dy) > tol_y:
                    if dy < 0:  # 需要向上
                        pydirectinput.keyDown('space')
                        time.sleep(0.1)  # 長按邏輯
                        pydirectinput.keyUp('space')
                    else:  # 需要向下 (向下跳)
                        pydirectinput.keyDown('down')
                        time.sleep(0.04)
                        pydirectinput.keyDown('alt')
                        time.sleep(0.05)
                        pydirectinput.keyUp('alt')
                        pydirectinput.keyUp('down')
                time.sleep(0.10)
            print(f"⚠️ 未能在 {max_time:.1f}s 內靠近目標 ({target_x:.1f},{target_y:.1f})")
            return False
        except Exception as e:
            print(f"❌ reposition 發生錯誤: {e}")
            return False

    def detect_collision(self):
        """簡單的撞牆檢測邏輯"""
        # 此處可以加入更複雜的撞牆檢測邏輯，例如檢測角色位置是否長時間未變化
        return False

    def pause_for_correction(self, current_x, current_y, expected_x, expected_y, x_diff, y_diff):
        """暫停腳本進行位置修正，類似失焦處理機制"""
        def correction_thread():
            try:
                self.is_correcting = True
                self.correction_pause_event.clear()  # 暫停腳本執行
                
                print("⏸️ 腳本已暫停，開始位置修正...")
                self.playback_status.config(text=f"播放狀態: 位置修正中 (迴圈 {self.current_loop}/{self.total_loops})")
                
                # 嘗試位置修正
                correction_success = self.attempt_position_correction(current_x, current_y, expected_x, expected_y, x_diff, y_diff)
                # 無論快速修正結果，嘗試精準回到事件預期位置
                precise_ok = False
                if expected_x is not None and expected_y is not None:
                    precise_ok = self.reposition_to(expected_x, expected_y, fallback=False, max_time=2.5)
                if not precise_ok and self.start_position:
                    print("↩️ 回到錄製起始點嘗試重新對齊")
                    self.reposition_to(self.start_position['x'], self.start_position['y'], fallback=True, max_time=3.0)
                # 若已經接近預期位置則本迴圈抑制 space 以降低再次偏離
                final_x, final_y = self.get_current_position()
                if final_x is not None and final_y is not None and expected_x is not None and expected_y is not None:
                    map_w = self.minimap_region[2] if self.minimap_region else 200
                    map_h = self.minimap_region[3] if self.minimap_region else 150
                    tol_x = max(5, map_w * 0.20)
                    tol_y = max(6, map_h * 0.28)
                    if abs(final_x - expected_x) <= tol_x and abs(final_y - expected_y) <= tol_y:
                        self.suppress_space_until_loop_end = True
                        print("🔕 已啟用本迴圈跳躍抑制 (校正後穩定)\n")
                # 重置偏離追蹤 (避免立刻再次觸發)
                self.deviation_start_time = None
                self.is_currently_deviating = False
                
                # 恢復腳本執行
                time.sleep(0.5)  # 短暫等待確保修正完成
                print("▶️ 位置修正完成，恢復腳本執行")
                self.is_correcting = False
                self.correction_pause_event.set()  # 恢復腳本執行
                
            except Exception as e:
                print(f"❌ 位置修正過程發生錯誤: {e}")
                # 即使出錯也要恢復腳本執行
                self.is_correcting = False
                self.correction_pause_event.set()
        
        # 在後台線程中執行修正，避免阻塞主播放線程
        threading.Thread(target=correction_thread, daemon=True).start()
    
    def execute_skill_with_repeat(self, skill_key, pressed_keys):
        """執行技能並支援連發（最多2次）"""
        current_time = time.time()
        
        # 檢查是否在連發冷卻時間內
        if skill_key in self.last_skill_time:
            time_since_last = current_time - self.last_skill_time[skill_key]
            if time_since_last < self.skill_repeat_interval:
                # 在冷卻時間內，執行第二次
                print(f"🔥 技能連發: {skill_key} (第2次)")
                pydirectinput.keyDown(skill_key)
                for key in pressed_keys:
                    if key != skill_key:
                        pydirectinput.keyDown(key)
                
                # 重置時間，避免第三次連發
                self.last_skill_time[skill_key] = current_time - self.skill_repeat_interval * 2
                return
        
        # 正常執行第一次
        print(f"⚔️ 技能施放: {skill_key}")
        pydirectinput.keyDown(skill_key)
        for key in pressed_keys:
            if key != skill_key:
                pydirectinput.keyDown(key)
        
        # 記錄施放時間
        self.last_skill_time[skill_key] = current_time
    
    def toggle_minimap(self):
        """切換小地圖偵測"""
        self.minimap_enabled = self.minimap_var.get()
        if self.minimap_enabled:
            print("✅ 小地圖偵測已啟用")
        else:
            print("❌ 小地圖偵測已關閉")
    
    def get_minimap_position_offset(self):
        """回傳目前小地圖位置 (暫作偏移用)"""
        return self.get_current_position()
    
    def setup_minimap_region(self):
        """設定小地圖區域 - 直接一氣呵成的流程"""
        if not self.find_maple_window():
            messagebox.showerror("錯誤", "找不到遊戲視窗")
            return

        # 創建簡化的設定對話框
        setup_dialog = tk.Toplevel(self.root)
        setup_dialog.title("設定小地圖區域")
        setup_dialog.geometry("450x300")
        setup_dialog.transient(self.root)
        setup_dialog.grab_set()

        info_text = """小地圖一氣呵成設定：

🎯 一鍵完成流程
1. 點擊「開始設定」按鈕
2. 用滑鼠框選小地圖區域
3. 系統自動截取並顯示小地圖
4. 在放大圖片上點擊人物位置
5. 完成設定，開始精確追蹤

⭐ 或單獨標定人物位置
如果已經設定過小地圖區域，
可以直接點擊「標定人物位置」重新標定"""

        info_label = ttk.Label(setup_dialog, text=info_text, justify="left")
        info_label.pack(pady=15, padx=15)

        button_frame = ttk.Frame(setup_dialog)
        button_frame.pack(pady=20)

        def start_complete_setup():
            setup_dialog.destroy()
            self.select_minimap_region_and_calibrate()

        def start_player_calibration_only():
            if not self.minimap_region:
                messagebox.showwarning("提醒", "請先設定小地圖區域")
                return
            setup_dialog.destroy()
            self.calibrate_player_position()

        ttk.Button(button_frame, text="🎯 開始設定 (框選+標定)", 
                   command=start_complete_setup).pack(side=tk.TOP, pady=5, fill=tk.X)
        ttk.Button(button_frame, text="⭐ 僅標定人物位置", 
                   command=start_player_calibration_only).pack(side=tk.TOP, pady=5, fill=tk.X)
        ttk.Button(button_frame, text="取消", 
                   command=setup_dialog.destroy).pack(side=tk.TOP, pady=10)
    
    def manual_coordinate_input(self):
        """手動輸入精確座標"""
        coord_window = tk.Toplevel(self.root)
        coord_window.title("手動輸入小地圖座標")
        coord_window.geometry("400x350")
        coord_window.transient(self.root)
        coord_window.grab_set()
        
        info_text = """輸入小地圖的精確座標：

提示：
• 使用 Windows 剪取工具查看座標
• 小地圖通常在遊戲視窗右上角
• 建議先截圖查看遊戲視窗位置"""
        
        ttk.Label(coord_window, text=info_text, justify="left").pack(pady=10)
        
        # 顯示當前螢幕資訊
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        current_mouse_x, current_mouse_y = pyautogui.position()
        
        screen_info = f"螢幕解析度: {screen_width} x {screen_height}\n當前滑鼠位置: ({current_mouse_x}, {current_mouse_y})"
        ttk.Label(coord_window, text=screen_info, foreground="blue").pack(pady=5)
        
        # 座標輸入框
        coord_frame = ttk.Frame(coord_window)
        coord_frame.pack(pady=15)
        
        ttk.Label(coord_frame, text="左上角 X:").grid(row=0, column=0, padx=5, sticky="e")
        x_entry = ttk.Entry(coord_frame, width=10)
        x_entry.grid(row=0, column=1, padx=5)
        x_entry.insert(0, str(current_mouse_x - 100))  # 預設值
        
        ttk.Label(coord_frame, text="左上角 Y:").grid(row=0, column=2, padx=5, sticky="e")
        y_entry = ttk.Entry(coord_frame, width=10)
        y_entry.grid(row=0, column=3, padx=5)
        y_entry.insert(0, str(current_mouse_y - 100))
        
        ttk.Label(coord_frame, text="寬度:").grid(row=1, column=0, padx=5, sticky="e")
        w_entry = ttk.Entry(coord_frame, width=10)
        w_entry.grid(row=1, column=1, padx=5)
        w_entry.insert(0, "200")
        
        ttk.Label(coord_frame, text="高度:").grid(row=1, column=2, padx=5, sticky="e")
        h_entry = ttk.Entry(coord_frame, width=10)
        h_entry.grid(row=1, column=3, padx=5)
        h_entry.insert(0, "150")
        
        def update_mouse_pos():
            x, y = pyautogui.position()
            mouse_label.config(text=f"即時滑鼠位置: ({x}, {y})")
            coord_window.after(100, update_mouse_pos)
        
        mouse_label = ttk.Label(coord_window, text="", foreground="green")
        mouse_label.pack(pady=5)
        update_mouse_pos()
        
        def confirm_coordinates():
            try:
                x = int(x_entry.get())
                y = int(y_entry.get())
                w = int(w_entry.get())
                h = int(h_entry.get())
                
                self.minimap_region = (x, y, w, h)
                print(f"📐 手動設定區域: ({x}, {y}, {w}, {h})")
                
                # 立即測試
                test_screenshot = pyautogui.screenshot(region=(x, y, w, h))
                test_path = os.path.join(os.path.dirname(__file__), 'minimap_manual_test.png')
                test_screenshot.save(test_path)
                print(f"💾 測試圖片已保存: {test_path}")
                
                self.minimap_status.config(text=f"小地圖: 手動設定 {w}x{h}")
                coord_window.destroy()

                # 保存設定
                self._save_minimap_config()
                
                messagebox.showinfo("設定完成", f"小地圖區域已設定完成！\n測試圖片: minimap_manual_test.png\n\n接下來可以標定人物位置以獲得最佳精確度")
                
            except ValueError:
                messagebox.showerror("錯誤", "請輸入有效的數字")
            except Exception as e:
                messagebox.showerror("錯誤", f"測試失敗: {e}")
        
        button_frame = ttk.Frame(coord_window)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="確認設定", command=confirm_coordinates).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=coord_window.destroy).pack(side=tk.LEFT, padx=5)
    
    def calibrate_player_position(self):
        """標定人物位置 - 手動點擊指定人物在小地圖上的位置並建立追蹤模板"""
        if not self.minimap_region:
            messagebox.showwarning("提醒", "請先設定小地圖區域")
            return
        try:
            current_minimap = self.capture_minimap()
            if current_minimap is None:
                messagebox.showerror("錯誤", "無法擷取小地圖")
                return
            pil_image = Image.fromarray(current_minimap)
            enlarged_image = pil_image.resize((pil_image.width * 4, pil_image.height * 4), Image.Resampling.NEAREST)
            tk_image = ImageTk.PhotoImage(enlarged_image)
            win = tk.Toplevel(self.root); win.title("標定人物位置 (點擊人物)"); win.transient(self.root); win.grab_set()
            canvas = tk.Canvas(win, width=enlarged_image.width, height=enlarged_image.height)
            canvas.pack(padx=5, pady=5)
            canvas.create_image(0, 0, anchor=tk.NW, image=tk_image); canvas.image = tk_image
            status_var = tk.StringVar(value="請點擊人物 (放大 4x)")
            ttk.Label(win, textvariable=status_var, foreground='blue').pack(pady=4)
            marker = {'id': None}
            selection = {'pt': None}
            def on_click(e):
                ox, oy = int(e.x/4), int(e.y/4)
                selection['pt'] = (ox, oy)
                if marker['id']: canvas.delete(marker['id'])
                marker['id'] = canvas.create_oval(e.x-8, e.y-8, e.x+8, e.y+8, outline='red', width=3, fill='yellow')
                status_var.set(f"選取人物座標: ({ox},{oy}) 再次點擊可重選，按確認建立追蹤")
            canvas.bind('<Button-1>', on_click)
            btn_frame = ttk.Frame(win); btn_frame.pack(pady=6)
            def confirm():
                if not selection['pt']:
                    messagebox.showwarning("提醒", "請先點擊人物位置")
                    return
                px, py = selection['pt']
                # 建立模板 (從原圖裁 9x9 區域，超出邊界自動裁剪)
                half = 5
                y1, y2 = max(py-half,0), min(py+half+1, current_minimap.shape[0])
                x1, x2 = max(px-half,0), min(px+half+1, current_minimap.shape[1])
                template = current_minimap[y1:y2, x1:x2].copy()
                self.player_template = template
                self.player_template_offset = (px - x1, py - y1)  # 回推中心偏移
                self.last_player_pos = (px, py)
                self.use_manual_position = True
                self.enable_template_tracking = True
                self.player_lost_frames = 0
                print(f"🎯 人物標定完成: ({px},{py}) 模板大小 {template.shape[:2]}")
                win.destroy()
                messagebox.showinfo("標定完成", "人物模板已建立，將自動追蹤移動。")
            ttk.Button(btn_frame, text='確認', command=confirm).pack(side=tk.LEFT, padx=4)
            ttk.Button(btn_frame, text='取消', command=win.destroy).pack(side=tk.LEFT, padx=4)
        except Exception as e:
            print(f"❌ 人物位置標定失敗: {e}")
            messagebox.showerror("錯誤", f"標定失敗: {e}")

    def get_minimap_player_position(self):
        """獲取小地圖上人物的位置 - 模板追蹤優先"""
        if not self.minimap_region:
            return None, None
        try:
            minimap_image = self.capture_minimap()
            if minimap_image is None:
                return None, None
            # 若啟用模板追蹤
            if getattr(self, 'enable_template_tracking', False) and hasattr(self, 'player_template'):
                tpl = self.player_template
                if tpl is not None and tpl.size > 0 and minimap_image.shape[0] >= tpl.shape[0] and minimap_image.shape[1] >= tpl.shape[1]:
                    method = cv2.TM_CCOEFF_NORMED
                    res = cv2.matchTemplate(minimap_image, tpl, method)
                    _, max_val, _, max_loc = cv2.minMaxLoc(res)
                    threshold = 0.75
                    if max_val >= threshold:
                        x, y = max_loc[0] + self.player_template_offset[0], max_loc[1] + self.player_template_offset[1]
                        self.last_player_pos = (x, y)
                        self.player_lost_frames = 0
                        return x, y
                    else:
                        # 容錯：使用最後位置 (短暫失去追蹤)
                        self.player_lost_frames = getattr(self, 'player_lost_frames', 0) + 1
                        if self.player_lost_frames <= 5 and self.last_player_pos:
                            return self.last_player_pos
                # 如果模板失效，回退到顏色偵測
            # 顏色偵測作為備援
            x, y = self.find_player_dot_on_minimap(minimap_image)
            return x, y
        except Exception as e:
            print(f"❌ 獲取小地圖人物位置失敗: {e}")
            return None, None

    def calibrate_minimap(self):
        """校準小地圖（若需對齊參考影像，可在這裡擴充）"""
        try:
            if not getattr(self, 'minimap_region', None):
                print("❌ 小地圖區域未設定，無法校準")
                return
            img = self.capture_minimap()
            if img is None:
                print("❌ 擷取小地圖失敗，無法校準")
                return
            self.minimap_reference = img
            print("✅ 小地圖參考影像已更新 (校準完成)")
        except Exception as e:
            print(f"❌ 小地圖校準失敗: {e}")

    def test_minimap_capture(self):
        """測試小地圖擷取，保存一張截圖檢查區域是否正確"""
        print("\n=== 測試小地圖擷取 ===")
        if not getattr(self, 'minimap_region', None):
            print("❌ 未設定小地圖區域，先開啟區域選擇")
            self.select_minimap_region()
            if not getattr(self, 'minimap_region', None):
                return
        img = self.capture_minimap()
        if img is None:
            print("❌ 擷取失敗")
            return
        try:
            out_path = os.path.join(os.path.dirname(__file__), 'minimap_test_capture.png')
            Image.fromarray(img).save(out_path)
            print(f"✅ 已保存測試截圖: {out_path}")
            if hasattr(self, 'minimap_status'):
                self.minimap_status.config(text="小地圖: 測試截圖已保存")
        except Exception as e:
            print(f"❌ 保存測試截圖失敗: {e}")

    def select_minimap_region_and_calibrate(self):
        """一氣呵成：框選區域 + 自動跳轉人物標定"""
        try:
            overlay = tk.Toplevel(self.root)
            overlay.attributes('-fullscreen', True)
            try:
                overlay.attributes('-alpha', 0.3)
            except:
                pass
            overlay.attributes('-topmost', True)
            overlay.config(bg='black')
            canvas = tk.Canvas(overlay, cursor='cross')
            canvas.pack(fill='both', expand=True)

            start = {'x': 0, 'y': 0}
            rect = {'id': None}

            def on_press(e):
                start['x'], start['y'] = e.x_root, e.y_root
                if rect['id']:
                    canvas.delete(rect['id'])
                rect['id'] = canvas.create_rectangle(e.x_root, e.y_root, e.x_root, e.y_root, outline='red', width=2)

            def on_move(e):
                if rect['id']:
                    canvas.coords(rect['id'], start['x'], start['y'], e.x_root, e.y_root)

            def on_release(e):
                x1, y1 = start['x'], start['y']
                x2, y2 = e.x_root, e.y_root
                left, top = min(x1, x2), min(y1, y2)
                right, bottom = max(x1, x2), max(y1, y2)
                w, h = right - left, bottom - top
                overlay.destroy()
                
                if w < 10 or h < 10:
                    messagebox.showwarning('提醒', '選取區域過小，請重新選取')
                    return
                
                # 設定小地圖區域
                self.minimap_region = (left, top, w, h)
                print(f"🖼️ 已選取小地圖區域: {self.minimap_region}")
                
                # 測試截圖
                try:
                    shot = pyautogui.screenshot(region=self.minimap_region)
                    test_path = os.path.join(os.path.dirname(__file__), 'minimap_select_test.png')
                    shot.save(test_path)
                    print(f"💾 區域測試截圖: {test_path}")
                except Exception as ce:
                    print(f"截圖失敗: {ce}")
                
                # 更新狀態
                if hasattr(self, 'minimap_status'):
                    self.minimap_status.config(text=f"小地圖: 已選取 {w}x{h}")
                
                # 保存設定
                self._save_minimap_config()
                
                # 🎯 自動跳轉到人物標定階段
                print("🎯 區域選取完成，自動進入人物標定...")
                self.root.after(500, self.calibrate_player_position)  # 延遲500ms自動跳轉

            def on_cancel(e):
                overlay.destroy()
                
            canvas.bind('<ButtonPress-1>', on_press)
            canvas.bind('<B1-Motion>', on_move)
            canvas.bind('<ButtonRelease-1>', on_release)
            canvas.bind('<Escape>', on_cancel)
            
        except Exception as e:
            print(f"❌ 框選小地圖區域失敗: {e}")

    def select_minimap_region(self):
        """使用滑鼠框選小地圖區域 (舊版本，保留相容性)"""
        try:
            overlay = tk.Toplevel(self.root)
            overlay.attributes('-fullscreen', True)
            try:
                overlay.attributes('-alpha', 0.3)
            except:
                pass
            overlay.attributes('-topmost', True)
            overlay.config(bg='black')
            canvas = tk.Canvas(overlay, cursor='cross')
            canvas.pack(fill='both', expand=True)

            start = {'x': 0, 'y': 0}
            rect = {'id': None}

            def on_press(e):
                start['x'], start['y'] = e.x_root, e.y_root
                if rect['id']:
                    canvas.delete(rect['id'])
                rect['id'] = canvas.create_rectangle(e.x_root, e.y_root, e.x_root, e.y_root, outline='red', width=2)

            def on_move(e):
                if rect['id']:
                    canvas.coords(rect['id'], start['x'], start['y'], e.x_root, e.y_root)

            def on_release(e):
                x1, y1 = start['x'], start['y']
                x2, y2 = e.x_root, e.y_root
                left, top = min(x1, x2), min(y1, y2)
                right, bottom = max(x1, x2), max(y1, y2)
                w, h = right - left, bottom - top
                overlay.destroy()
                if w < 10 or h < 10:
                    messagebox.showwarning('提醒', '選取區域過小，請重新選取')
                    return
                self.minimap_region = (left, top, w, h)
                print(f"🖼️ 已選取小地圖區域: {self.minimap_region}")
                # 測試截圖
                try:
                    shot = pyautogui.screenshot(region=self.minimap_region)
                    test_path = os.path.join(os.path.dirname(__file__), 'minimap_select_test.png')
                    shot.save(test_path)
                    print(f"💾 區域測試截圖: {test_path}")
                except Exception as ce:
                    print(f"截圖失敗: {ce}")
                if hasattr(self, 'minimap_status'):
                    self.minimap_status.config(text=f"小地圖: 已選取 {w}x{h}")
                self._save_minimap_config()

            def on_cancel(e):
                overlay.destroy()
            canvas.bind('<ButtonPress-1>', on_press)
            canvas.bind('<B1-Motion>', on_move)
            canvas.bind('<ButtonRelease-1>', on_release)
            canvas.bind('<Escape>', on_cancel)
        except Exception as e:
            print(f"❌ 框選小地圖區域失敗: {e}")

    # ================== 小地圖擷取與玩家偵測 ==================
    def capture_minimap(self):
        """擷取目前設定區域的小地圖畫面，回傳 numpy array (RGB)"""
        if not self.minimap_region:
            return None
        try:
            x, y, w, h = self.minimap_region
            img = pyautogui.screenshot(region=(x, y, w, h))  # PIL Image (RGB)
            return np.array(img)
        except Exception as e:
            print(f"擷取小地圖失敗: {e}")
            return None

    def find_player_dot_on_minimap(self, minimap_image):
        """簡單顏色偵測人物點 (備援)"""
        try:
            arr = minimap_image
            # 建立遮罩：菱形黃色點 (RGB: fff88)
            yellow_diamond_mask = (
                (arr[:, :, 0] > 250) &  # R
                (arr[:, :, 1] > 250) &  # G
                (arr[:, :, 2] < 140)    # B
            )
            mask = yellow_diamond_mask
            coords = np.column_stack(np.where(mask))  # (y, x)
            if coords.size == 0:
                return None, None
            # 取平均中心
            mean_y, mean_x = coords.mean(axis=0)
            return int(mean_x), int(mean_y)
        except Exception as e:
            print(f"顏色偵測失敗: {e}")
            return None, None

    # ================== 小地圖路徑記錄 (簡易佔位) ==================
    def start_minimap_path_recording(self):
        """目前僅佔位，後續可擴充為實際路徑點紀錄"""
        self.minimap_path_points = []
        print("📝 開始小地圖路徑記錄 (尚未實作詳盡功能)")

    def start_minimap_monitoring(self):
        """開始小地圖監控 (顯示並追蹤位置)"""
        if not getattr(self, 'minimap_region', None):
            messagebox.showwarning("提醒", "請先設定小地圖區域")
            return
        self.minimap_enabled = True
        if not hasattr(self, 'minimap_canvas'):
            # 建一個簡單預設顯示區 (可與原UI整合)
            self.minimap_canvas = tk.Canvas(self.root, width=150, height=150, bg='black')
            self.minimap_canvas.place(x=10, y=10)
        if hasattr(self, 'minimap_status'):
            self.minimap_status.config(text='小地圖: 監控中')
        self._schedule_minimap_update()
        print('🗺️ 小地圖監控啟動')

    def stop_minimap_monitoring(self):
        self.minimap_enabled = False
        if hasattr(self, 'minimap_status'):
            self.minimap_status.config(text='小地圖: 已停止')
        print('⏹️ 小地圖監控停止')

    def _schedule_minimap_update(self):
        if getattr(self, 'minimap_enabled', False):
            self.update_minimap_display()
            self.root.after(self.minimap_update_interval, self._schedule_minimap_update)

    def update_minimap_display(self):
        """更新小地圖顯示 + 標記 (簡化版)"""
        try:
            if not getattr(self, 'minimap_enabled', False):
                return
            img = self.capture_minimap()
            if img is None:
                return
            # 顯示縮放
            disp = Image.fromarray(img).resize((150,150), Image.Resampling.NEAREST)
            tk_img = ImageTk.PhotoImage(disp)
            self.minimap_canvas.delete('all')
            self.minimap_canvas.create_image(0,0, anchor=tk.NW, image=tk_img)
            self.minimap_canvas.image = tk_img
            # 取得人物位置
            px, py = self.get_minimap_player_position()
            if px is not None and py is not None:
                scale_x = 150 / img.shape[1]
                scale_y = 150 / img.shape[0]
                sx, sy = px*scale_x, py*scale_y
                self.minimap_canvas.create_oval(sx-4, sy-4, sx+4, sy+4, outline='red', width=2)
                if hasattr(self, 'minimap_status'):
                    self.minimap_status.config(text=f'小地圖: 位置 ({px},{py})')
        except Exception as e:
            print(f'❌ 更新小地圖失敗: {e}')

    def update_minimap_interval(self):
        """更新小地圖監控刷新頻率"""
        try:
            val = int(self.minimap_interval_entry.get())
            if val < 50:
                val = 50  # 設定最小值避免過高負載
            if val > 5000:
                val = 5000
            self.minimap_update_interval = val
            print(f"⏱️ 小地圖刷新頻率已設定為 {val} ms")
            if self.minimap_enabled and hasattr(self, 'minimap_status'):
                self.minimap_status.config(text=f'小地圖: 監控中 ({val}ms)')
        except ValueError:
            messagebox.showerror("錯誤", "請輸入有效的整數毫秒值")

# ================== 程式入口 ==================
if __name__ == "__main__":
    print("=== 程序開始執行 ===")
    
    # 檢查管理員權限
    if not is_admin():
        print("⚠️ 檢測到非管理員權限，嘗試提升權限...")
        if not run_as_admin():
            sys.exit(1)
        sys.exit(0)
    
    print("✅ 管理員權限確認")
    
    try:
        print("啟動程序...")
        # 避免 pydirectinput 自動延遲
        try:
            pydirectinput.PAUSE = 0
            print("✓ pydirectinput 設定完成")
        except Exception as e:
            print(f"✗ pydirectinput 設定失敗: {e}")
        
        print("創建 Tkinter 根視窗...")
        root = tk.Tk()
        print("✓ Tkinter 主視窗已建立")
        
        print("初始化 MacroApp...")
        app = MacroApp(root)
        print("✓ MacroApp 已啟動")

        # 強制將 GUI 視窗置於最前方
        print("設定視窗置頂...")
        try:
            root.attributes('-topmost', True)
            root.update()
            root.attributes('-topmost', False)
            print("✓ 視窗置頂設定完成")
        except Exception as e:
            print(f"⚠️ 視窗置頂設定失敗: {e}")

        print("🚀 宏工具已啟動 - 若視窗未顯示在最前，可手動切換")
        print("進入 Tkinter 主迴圈...")
        root.mainloop()
        print("=== 主迴圈結束 ===")
    except Exception as e:
        print(f"❌ 啟動失敗: {e}")
        import traceback
        traceback.print_exc()
