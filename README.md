# 楓之谷自動化腳本

作者：**SchwarzeKatze_R**  
版本：**1.0**

---

## 📖 專案簡介
此專案為楓之谷自動化腳本，提供錄製、播放按鍵動作及小地圖輔助功能。  
採用 Python + Tkinter GUI，支援自動回城、位置校正與技能連發。

---

## ⚡ 功能特色
- **錄製與播放**：支援鍵盤動作錄製、迴圈播放  
- **小地圖偵測**：定位角色、偏移校正、回到起始點  
- **技能連發**：多按鍵技能自動重複施放  
- **回程功能**：支援傳統回城或自動回到錄製位置  
- **GUI 操作**：直覺化控制面板與狀態顯示  

---

## 🛠️ 安裝需求
- Python 3.x
- 依賴套件：
  - `tkinter`
  - `keyboard`
  - `pydirectinput`
  - `pyautogui`
  - `pyperclip`
  - `opencv-python`
  - `numpy`
  - `Pillow`
  - `pywin32`

安裝：
```bash
pip install keyboard pydirectinput pyautogui pyperclip opencv-python numpy pillow pywin32
```

---

## 🚀 使用方式
1. 以 **管理員權限** 執行程式：
   ```bash
   python maple_macro.py
   ```
2. 在 GUI 介面中：
   - **開始錄製** → 執行遊戲動作  
   - **停止錄製** → 儲存腳本  
   - **播放腳本** → 自動執行操作  

---

## ⚠️ 注意事項
- 請自行承擔使用風險，可能違反遊戲規範  
- 建議僅作學習與研究用途  
- 請務必以 **管理員權限** 執行  

---

## 📄 授權
本專案採用 [MIT License](https://github.com/kuroneko11375/MapleStoryMacro/blob/main/LICENSE)。
