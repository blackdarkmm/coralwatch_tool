# coralwatch_tool

![Python Version](https://shields.io/badge/python-3.13.5-blue)
>這是一個專門針對coralwatch的工具，用於讓珊瑚資料輸入excel時不用一個一個輸入。

## 🚀 使用條件
- 有一個coralwatch專用的excel。

- 有一個圖片資料夾(只支援:".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"，包括大寫副檔名)

## 📦 下載步驟

1. 到releases下載coralwatch_tool.zip
2. 解壓縮

## 🎯 使用說明

- 介面上的比色卡每個顏色各自代表各自顏色代號的按鈕(請以照片內的比色卡顏色為準)。

- 沒有選擇excel檔案跟圖片檔案前點擊比色卡按鈕不會輸入。

- 使用時選擇excel檔案時直接選取你的excel檔案，可以手動輸入**絕對路徑**。

- 使用時選擇圖片時是**選取你的圖片資料夾**，若要手動輸入則輸入圖片**絕對路徑**。

- 一個資料夾內的圖片輸入完後要重新選擇excel檔案跟圖片檔案。

## 🛠️ 環境需求
**非開發人員無須閱讀**
請確保您的電腦已安裝以下環境（適用於想修改原始碼的開發者）：
- Python 3.13.5 (不確定更高版本是否相容)

<details>
<summary><b>點擊展開/摺疊：詳細 Python 環境套件清單 (requirements.txt)</b></summary>

```text
altgraph==0.17.5
et_xmlfile==2.0.0
openpyxl==3.1.5
packaging==26.3
pefile==2024.8.26
pyinstaller==6.22.3
pyinstaller-hooks-contrib==2026.7
PyQt6==6.11.0
PyQt6-Qt6==6.11.2
PyQt6_sip==13.12.0
pywin32==312
pywin32-ctypes==0.2.3
setuptools==84.0.0
xlwings==0.37.4
```

**執行主程式：**
```bash
python main.py
```

## 📂 專案結構
```text
├── main.py          # 程式主入口
├── function.py        # 函式庫
├── coralwatch.ico      # 圖標
├── main.ui         #ui代碼
├── requirements.txt # 環境套件清單
└── README.md        # 本說明文件
```

## 📝 授權條款
本專案採用 [MIT](LICENSE) 授權條款。