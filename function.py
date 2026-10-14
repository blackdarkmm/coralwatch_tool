from PyQt6.QtWidgets import QFileDialog
import os
import hashlib
from PyQt6.QtCore import Qt , QTimer , QObject, pyqtSignal
from PyQt6.QtGui import QFont , QImage , QPixmap
import tomllib
import xlwings as xw
import threading
import queue
import time


class VarManager:

    def __init__(self,main_app):
        self.main = main_app
        self.lock = threading.Lock()
        self.cache_lists = []
        self.cache_pictures = []

    def picture_list(self,lists = None):
        with self.lock:
            if isinstance(lists,list):
                self.picture_lists = lists
            elif lists:
                self.picture_lists.append(lists)
            else:
                try:
                    return self.picture_lists
                except:
                    return

    def cache_list(self,lists = None):
        with self.lock:
            if isinstance(lists,list):
                self.cache_lists = lists
            elif lists:
                self.cache_lists.append(lists)
            else:
                try:
                    return self.cache_lists
                except:
                    return

    def cache_picture(self,lists = None):
        with self.lock:
            if isinstance(lists,list):
                self.cache_pictures = lists
            elif lists:
                self.cache_pictures.append(lists)
            else:
                try:
                    return self.cache_pictures
                except:
                    return
        
    def picture_folder(self,path = None):
        with self.lock:
            if path is not None:
                self.picture_folder_path = path
            else:
                try:
                    return self.picture_folder_path
                except:
                    return

    def picture_file(self,path = None):
        with self.lock:
            if path is not None:
                self.picture_file_name = path
            else:
                try:
                    return self.picture_file_name
                except:
                    return

    def excel_file(self,path = None):
        with self.lock:
            if path is not None:
                self.excel_file_path = path
            else:
                try:
                    return self.excel_file_path
                except:
                    return

    def files_num(self,num: int |None = None) -> int | None:
        with self.lock:
            if num is not None:
                num = int(num)
                self.file_num = num
            else:
                try:
                    return self.file_num
                except:
                    return

    def end_picture_num(self,num: int |None = None) -> int | None:
        with self.lock:
            if num is not None:
                num = int(num)
                self.end_picture_nums = num
            else:
                try:
                    return self.end_picture_nums
                except:
                    return

    def ui_text_control(self,uiobj_name,content = None):
        uiobj = getattr(self.main,uiobj_name)
        if content == "":
            uiobj.clear()
        elif not content:
            return uiobj.text()
        elif content:
            uiobj.setText(content)

class Function:
    def __init__(self,main_app,var):
        self.main = main_app
        self.var = var
        self.os_path = os.path
        self.before_folder_hash = None
        self.var.files_num(0)
        self.coral_type = ""
        self.var.end_picture_num(0)
        self.color_type = {"B1", "B2", "B3", "B4", "B5", "B6","C1", "C2", "C3", "C4", "C5", "C6","D1", "D2", "D3", "D4", "D5", "D6","E1", "E2", "E3", "E4", "E5", "E6"}

    def init(self):
        self.main.todo.excel_close()
        self.before_folder_hash = None
        self.var.picture_list([])
        self.var.cache_list([])
        self.var.cache_picture([])
        self.var.picture_folder("")
        self.var.picture_file("")
        self.var.files_num(0)
        self.var.end_picture_num(0)
        self.var.ui_text_control("excel_source","")
        self.var.ui_text_control("picture_source","")

    def findfile_or_floder(self,code):
        file_dialog = QFileDialog()
        if code == "excel":
            file_path, _ = file_dialog.getOpenFileName(None, "選擇 Excel 檔案", "", "Excel 檔案 (*.xlsx *.xls)")    
            self.var.excel_file(file_path)
        elif code == "picture":
            folder_path = file_dialog.getExistingDirectory(None, "選擇圖片資料夾")
            self.var.picture_folder(folder_path)
        else:
            raise TypeError("code only excel or picture")

    def serch_excel_file(self,file_path = None,staus = None):
        if not file_path:
            if not self.var.ui_text_control("excel_source"):
                self.var.ui_text_control("coral_picture","excel路徑不可為空")
                return None
            path = self.var.ui_text_control("excel_source")
        elif file_path:
            path = file_path
            self.var.ui_text_control("excel_source", path)
        real_excel = self.os_path.exists(path) and self.os_path.isfile(path) and path.lower().endswith((".xlsx", ".xls"))
        if real_excel:
            self.picture_or_floder()
            if staus == "test":
                pass
            else:
                self.main.tcfuncs.revise_picture()
            return path
        elif not path:
            return None
        else:
            self.var.ui_text_control("coral_picture","excel檔案不存在")
            return None

    def picture_or_floder(self,folder_path = None):
        if not folder_path and not self.var.ui_text_control("picture_source"):
            self.var.ui_text_control("coral_picture","圖片路徑不可為空")
            return None
        elif self.os_path.isabs(self.var.ui_text_control("picture_source")) and not folder_path:
            self.var.picture_file(self.var.ui_text_control("picture_source"))
            folder_path, name = self.os_path.split(self.var.picture_file())
            self.var.picture_folder(folder_path)
            self.var.picture_file(name)
            return True
        elif self.os_path.isdir(folder_path):
            first_image = next((f for f in sorted(os.listdir(folder_path)) if f.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff")) and self.os_path.isfile(self.os_path.join(folder_path, f))),None)
            if first_image is None:
                self.var.ui_text_control("coral_picture","資料夾下不存在圖片")
                return None
            else:
                self.var.picture_file(self.os_path.join(folder_path, first_image))
                self.var.ui_text_control("picture_source",self.var.picture_file())
                return True
        elif not self.var.ui_text_control("picture_source"):
            return None
        elif not self.os_path.isabs(self.var.ui_text_control("picture_source")):
            self.var.ui_text_control("coral_picture","圖片資料夾不存在或著不是絕對路徑")
            return None

    def paste_picture(self): 

        hasher = hashlib.sha256()
        if not self.var.ui_text_control("picture_source"):
            return None
        
        folder_path, _ = self.os_path.split(self.var.ui_text_control("picture_source"))

        if not self.os_path.isdir(folder_path):
            return None
        try:
            folder_stat = os.stat(folder_path)
            file_meta = f"{folder_path}|{folder_stat.st_mtime}|{folder_stat.st_size}"
        except Exception:
            return None
                
        hasher.update(file_meta.encode('utf-8'))
    
        current_folder_hash = hasher.hexdigest()

        if not self.before_folder_hash or current_folder_hash != self.before_folder_hash:

            self.var.picture_list([])
            for file in sorted(os.listdir(folder_path)):
                if file.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff")) and self.os_path.isfile(self.os_path.join(folder_path, file)):
                    self.var.picture_list(file)
            self.var.files_num(self.var.picture_list().index(os.path.basename(self.var.ui_text_control("picture_source"))))
            self.before_folder_hash = current_folder_hash
            self.var.end_picture_num(len(self.var.picture_list()))
            self.var.picture_file(self.var.ui_text_control("picture_source"))
            self.var.picture_folder(folder_path)
            self.var.cache_list([])
            self.var.cache_picture([])

        if self.var.end_picture_num() == self.var.files_num() and self.var.end_picture_num() != 0:
            self.main.coral_picture.clear()
            self.init()
            self.var.ui_text_control("coral_picture", "本次所選範圍已經沒有圖片了，請重新選擇資料夾")
            return 0
        self.main.coral_picture.clear()

    def picture_load(self,picture):
        pixmap = QPixmap(picture)
        self.main.coral_picture.setPixmap(pixmap)

    def coralwatchcarch(self,code):

        if not self.serch_excel_file(staus="test"):
            return
    
        if not self.picture_or_floder(): 
            return
    
        if code in {"BR", "BO", "PL", "SO"}:
            self.coral_type = code
        elif code in self.color_type:
            if not self.var.ui_text_control("lightest_color"):
                self.var.ui_text_control("lightest_color", code)
            elif not self.var.ui_text_control("darkest_color"):
                self.var.ui_text_control("darkest_color", code)

        if self.var.ui_text_control("lightest_color") and self.var.ui_text_control("darkest_color") and self.coral_type:
            data =  (self.coral_type, self.var.ui_text_control("lightest_color"), self.var.ui_text_control("darkest_color"),self.main.remark.toPlainText())
            self.var.ui_text_control("lightest_color", "")
            self.var.ui_text_control("darkest_color", "")
            self.coral_type = ""
            self.main.remark.clear()
            return data

    def confirm(self):

        if not self.var.ui_text_control("first") and not self.var.ui_text_control("last") and self.serch_excel_file() and self.picture_or_floder():
            self.main.coral_picture.clear()
            self.var.ui_text_control("coral_picture", "如要設定範圍，請至少輸入第幾張開始")
            QTimer.singleShot(1000, lambda: (self.ui_text_control("coral_picture", "clear()"), self.picture_next()))
        
        elif self.var.ui_text_control("first") and not self.var.ui_text_control("last") and self.serch_excel_file() and self.picture_or_floder():
            self.var.ui_text_control("last", str(self.var.end_picture_num()))
            self.var.files_num(int(self.var.ui_text_control("first")) - 1)
            file_path = self.os_path.join(self.var.picture_folder(), self.var.picture_list()[self.var.files_num()])
            self.var.cache_picture([])
            self.var.cache_list([])
            self.var.ui_text_control("picture_source", file_path)


        elif int(self.var.ui_text_control("last")) < int(self.var.ui_text_control("first")):
            self.main.coral_picture.clear()
            self.ui_text_control("coral_picture", "結束張數不可小於開始張數")
            QTimer.singleShot(1000, lambda: (self.var.ui_text_control("coral_picture", "clear()"), self.picture_next()))
        
        elif self.var.ui_text_control("first") and self.var.ui_text_control("last") and self.serch_excel_file() and self.picture_or_floder() and self.var.end_picture_num() != 0:
            self.var.files_num(int(self.var.ui_text_control("first")) - 1)
            self.var.end_picture_num(int(self.var.ui_text_control("last")))
            file_path = self.os_path.join(self.var.picture_folder(), self.var.picture_list()[self.var.files_num()])
            self.var.cache_picture([])
            self.var.cache_list([])
            self.var.ui_text_control("picture_source", file_path)


    def font_size(self):
        current_width = self.main.width()
        dynamic_size = int((current_width - 640) / (1200 - 640) * (24 - 9) + 9)
        final_font_size = min(max(9, dynamic_size), 24)
        my_font = QFont()
        my_font.setPointSize(final_font_size)
        self.main.B1.setFont(my_font)
        self.main.B2.setFont(my_font)
        self.main.B3.setFont(my_font)
        self.main.B4.setFont(my_font)
        self.main.B5.setFont(my_font)
        self.main.B6.setFont(my_font)
        self.main.C1.setFont(my_font)
        self.main.C2.setFont(my_font)
        self.main.C3.setFont(my_font)
        self.main.C4.setFont(my_font)
        self.main.C5.setFont(my_font)
        self.main.C6.setFont(my_font)
        self.main.D1.setFont(my_font)
        self.main.D2.setFont(my_font)
        self.main.D3.setFont(my_font)
        self.main.D4.setFont(my_font)
        self.main.D5.setFont(my_font)
        self.main.D6.setFont(my_font)
        self.main.E1.setFont(my_font)
        self.main.E2.setFont(my_font)
        self.main.E3.setFont(my_font)
        self.main.E4.setFont(my_font)
        self.main.E5.setFont(my_font)
        self.main.E6.setFont(my_font)
        self.main.coral_picture.setFont(my_font)

    def picture_click(self):
        self.findfile_or_floder("picture")
        self.picture_or_floder(self.var.picture_folder())

class TimeConsumingFunction(QObject):

    image_loaded = pyqtSignal(QImage)

    def __init__(self,main_app,var):
        super().__init__()
        self.os_path = os.path
        self.main = main_app
        self.var = var
        self.wb = None
        self.app = None
        self.task_queue = queue.Queue()
        self.worker_thread = threading.Thread(target=self._thread_loop, daemon=True)
        self.worker_thread.start()
        self._config_cache = None
        self.excel_min_raw = 4
        self.image_loaded.connect(self.on_image_ready_handle)

    def _thread_loop(self):

        import pythoncom


        while True:

            task_type, data = self.task_queue.get()


            try:
                if task_type == "open_excel":
                    pythoncom.CoInitialize()
                    self.open_excel(data)
                elif task_type == "excel_add":
                    self.excel_add(data)
                elif task_type == "picture_cache":
                    self.picture_cache()
                elif task_type == "excel_close":
                    self.excel_close(data)
                    pythoncom.CoUninitialize()
                    self.task_queue.task_done()


            except Exception as e:
                print(f"{task_type}⚠️ 背景 Excel 執行緒發生錯誤: {e}")

                self.task_queue.task_done()

    def open_excel(self,data):
        excel_path = data
        self.app = xw.App(visible=False, add_book=False)
        self.wb = self.app.books.open(excel_path)

    def excel_add(self,input_data):

        type,lightestcolorcode,darkestcolorletter,remark = input_data
        picture_path = self.var.ui_text_control("picture_source")
        _,picture_name = self.os_path.split(picture_path)
        if not self._config_cache and self.os_path.exists("button_setting.toml"):
            self._config_cache = {}
            with open("button_setting.toml", "rb") as f:
                self._config_cache.update(tomllib.load(f))

        char_list_lightest = list(lightestcolorcode)
        char_list_darkest = list(darkestcolorletter)

        if type in self._config_cache:
            real_type_name = self._config_cache[type].get("output", type)
        else:
            real_type_name = type

        now_row = self.excel_min_raw + self.var.files_num()


        ws = self.wb.sheets['Data entry']

        existing_g_value = ws.range(f"G{now_row}").value

        row_data = [
            str(picture_name),              # F 欄
            existing_g_value,                           # G 欄 (關鍵！直接跳過不寫入，保持原樣)
            str(real_type_name),            # H 欄
            str(char_list_lightest[0]),     # I 欄
            int(char_list_lightest[1]),     # J 欄
            str(char_list_darkest[0]),      # K 欄
            int(char_list_darkest[1]),      # L 欄
            remark                          # M 欄
        ]

        ws.range(f"F{now_row}").value = row_data


        self.save_file(self.wb)

        self.var.files_num(self.var.files_num() + 1)

        del self.var.cache_picture()[0]
        del self.var.cache_list()[0]

        staus = self.main.funcs.paste_picture()
        if staus == 0:
            return

        self.var.ui_text_control("picture_source", self.os_path.join(self.var.picture_folder(), self.var.picture_list()[self.var.files_num()]))

    def save_file(self,wb):
        if wb:
            wb.save()

    def excel_close(self,data):
        wb,app = data
        self.save_file(wb)
        if wb:
            if wb is not None:
                wb.save()
                wb.close()
                app.quit()
                self.wb = None
                self.app = None
                self.var.excel_file("")

    def picture_cache(self):
        current_w = self.main.coral_picture.width()
        current_h = self.main.coral_picture.height()

        if not self.var.picture_folder() or not self.var.picture_list() or self.var.files_num() is None:
            return
        if len(self.var.cache_list()) < 1:
            first_picture = QImage(self.os_path.join(self.var.picture_folder(), self.var.picture_list()[self.var.files_num()]))
            self.var.cache_picture(first_picture)
            first_pixmap = self.var.cache_picture()[0].scaled(current_w, current_h, Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation)
            self.var.cache_list(first_pixmap)
        self.image_loaded.emit(self.var.cache_list()[0])
        if (self.var.files_num() + 1 != self.var.end_picture_num()):
            next_picture = QImage(self.os_path.join(self.var.picture_folder(), self.var.picture_list()[self.var.files_num() + 1]))
            self.var.cache_picture(next_picture)
            next_pixmap = self.var.cache_picture()[1].scaled(current_w, current_h, Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation)
            self.var.cache_list(next_pixmap)


    def revise_picture(self):
        if not len(self.var.cache_list()):
            return
        current_w = self.main.coral_picture.width()
        current_h = self.main.coral_picture.height()
        self.var.cache_list([])
        first_pixmap = self.var.cache_picture()[0].scaled(current_w, current_h, Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation)
        self.var.cache_list(first_pixmap)
        self.main.funcs.picture_load(self.var.cache_list()[0])
        if (self.var.files_num() + 1 != self.var.end_picture_num()):
            next_pixmap = self.var.cache_picture()[1].scaled(current_w, current_h, Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation)
            self.var.cache_list(next_pixmap)

    def on_image_ready_handle(self,qimage):
        self.main.funcs.picture_load(qimage)

class ToDo:

    def __init__(self,main_app,var):
        self.main = main_app
        self.var = var

    def excel_click(self):
        self.main.funcs.findfile_or_floder("excel")
        self.var.excel_file(self.main.funcs.serch_excel_file(self.var.excel_file()))

    def excel_opne(self):
        path = self.var.excel_file()
        self.main.tcfuncs.task_queue.put(("excel_close",(self.main.tcfuncs.wb,self.main.tcfuncs.app)))
        if path:
            self.main.tcfuncs.task_queue.put(("open_excel",path))

    def picture_click(self):
        self.main.funcs.findfile_or_floder("picture")
        self.main.funcs.picture_or_floder(self.var.picture_folder())

    def picture_next(self):
        staus = self.main.funcs.paste_picture()
        if staus == 0:
           return
        self.main.tcfuncs.task_queue.put(("picture_cache",None))
        
    def button_click(self, code):
        data = self.main.funcs.coralwatchcarch(code)
        if data:
            self.main.tcfuncs.task_queue.put(("excel_add",data))
            
    def excel_close(self,*arg):
        self.main.tcfuncs.task_queue.put(("excel_close",(self.main.tcfuncs.wb,self.main.tcfuncs.app)))
