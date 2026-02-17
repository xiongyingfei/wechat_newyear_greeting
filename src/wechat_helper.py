# -*- coding: utf-8 -*-
import win32gui
import win32con
import win32api
import pyperclip
import time
import re
from typing import Optional
from dataclasses import dataclass


@dataclass
class ContactInfo:
    full_name: str
    surname: str
    given_name: str


# 全局缓存 RapidOCR 实例（初始化很慢，只做一次）
_rapidocr_instance = None


def _get_rapidocr():
    global _rapidocr_instance
    if _rapidocr_instance is None:
        from rapidocr_onnxruntime import RapidOCR
        _rapidocr_instance = RapidOCR()
    return _rapidocr_instance


class WeChatHelper:
    def __init__(self):
        self.wechat_hwnd: Optional[int] = None
        self.current_contact_name: Optional[str] = None

    def find_wechat_window(self) -> bool:
        """查找微信窗口句柄"""
        def enum_callback(hwnd: int, windows: list):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if "微信" in title or "WeChat" in title:
                    windows.append((hwnd, title))
            return True

        windows: list = []
        win32gui.EnumWindows(enum_callback, windows)
        if not windows:
            return False
        self.wechat_hwnd = windows[0][0]
        return True

    def init_wechat(self) -> bool:
        """初始化微信连接"""
        if self.find_wechat_window():
            # 预热 RapidOCR（后台加载模型，避免第一次按键卡顿）
            try:
                _get_rapidocr()
                print("[就绪] OCR 引擎已加载")
            except Exception:
                print("[提示] RapidOCR 不可用，将使用 Tesseract")
            return True
        print("[错误] 未找到微信窗口，请确保微信已打开")
        return False

    def _parse_name(self, full_name: str) -> ContactInfo:
        """解析姓名，分离姓和名"""
        full_name = full_name.strip()
        if any('\u4e00' <= c <= '\u9fff' for c in full_name):
            if len(full_name) >= 2:
                return ContactInfo(full_name, full_name[0], full_name[1:])
            return ContactInfo(full_name, full_name, "")
        parts = full_name.split()
        if len(parts) >= 2:
            return ContactInfo(full_name, parts[-1], " ".join(parts[:-1]))
        return ContactInfo(full_name, full_name, "")

    def get_current_contact(self) -> Optional[ContactInfo]:
        """获取当前联系人：先 OCR，失败则用缓存"""
        try:
            if not self.wechat_hwnd:
                return None

            contact_name = self._ocr_contact_name()
            if contact_name:
                self.current_contact_name = contact_name
                info = self._parse_name(contact_name)
                print(f"[调试] OCR结果='{contact_name}' -> 全名='{info.full_name}' 姓='{info.surname}' 名='{info.given_name}'", flush=True)
                return info

            if self.current_contact_name:
                return self._parse_name(self.current_contact_name)

            print("[提示] 未能识别联系人，请用 Ctrl+Alt+N 手动输入")
            return None

        except Exception as e:
            print(f"[错误] 获取联系人出错: {e}")
            if self.current_contact_name:
                return self._parse_name(self.current_contact_name)
            return None

    def _ocr_contact_name(self) -> Optional[str]:
        """截图 + 预处理 + OCR 识别标题栏联系人"""
        try:
            import cv2
            import numpy as np
            import mss

            if not self.wechat_hwnd:
                return None

            rect = win32gui.GetWindowRect(self.wechat_hwnd)
            left, top, right, bottom = rect

            title_left = left + 300  # 跳过左侧边栏
            title_width = right - title_left
            title_height = 100

            if title_width <= 0:
                return None

            with mss.mss() as sct:
                monitor = {
                    "top": top, "left": title_left,
                    "width": title_width, "height": title_height,
                    "mon": 0,
                }
                screenshot = sct.grab(monitor)

            img = np.array(screenshot)[:, :, :3]
            gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)

            # 预处理：裁剪左侧 → 反色 → 放大3倍 → 二值化
            cropped = gray[:, :min(200, gray.shape[1])]
            inverted = 255 - cropped
            scaled = cv2.resize(inverted, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)
            _, binary = cv2.threshold(scaled, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

            # 优先 RapidOCR，失败用 Tesseract
            name = self._try_rapidocr(binary)
            if name:
                return name
            return self._try_tesseract(binary)

        except Exception as e:
            print(f"[调试] OCR 出错: {e}")
            return None

    def _try_rapidocr(self, img) -> Optional[str]:
        try:
            ocr = _get_rapidocr()
            result, _ = ocr(img)
            if result:
                for line in result:
                    name = self._extract_chinese_name(line[1])
                    if name:
                        return name
        except Exception:
            pass
        return None

    def _try_tesseract(self, img) -> Optional[str]:
        try:
            import pytesseract
            import os
            os.environ['TESSDATA_PREFIX'] = 'tessdata'
            pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
            text = pytesseract.image_to_string(img, lang='chi_sim', config='--psm 7').strip()
            if text:
                return self._extract_chinese_name(text)
        except Exception:
            pass
        return None

    def _extract_chinese_name(self, text: str) -> Optional[str]:
        """从 OCR 文本中提取中文姓名"""
        text = text.replace(" ", "")
        for name in re.findall(r'[\u4e00-\u9fff]+', text):
            if 2 <= len(name) <= 8 and name not in ("微信", "聊天", "通讯录", "搜索", "文件传输助手"):
                return name
        return None

    def set_contact_name(self, name: str) -> None:
        """手动设置联系人姓名"""
        self.current_contact_name = name.strip() if name else None
        if self.current_contact_name:
            print(f"[成功] 已设置联系人: {self.current_contact_name}")

    def send_message(self, message: str) -> bool:
        """将消息粘贴到当前焦点窗口的输入框"""
        try:
            old_clipboard = ""
            try:
                old_clipboard = pyperclip.paste()
            except:
                pass

            pyperclip.copy(message)

            # 短暂等待确保剪贴板就绪
            time.sleep(0.1)

            # Ctrl+V 粘贴（缩短按键间隔以加快响应）
            win32api.keybd_event(17, 0, 0, 0)       # Ctrl down
            time.sleep(0.01)
            win32api.keybd_event(86, 0, 0, 0)       # V down
            time.sleep(0.01)
            win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0)  # V up
            time.sleep(0.01)
            win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)  # Ctrl up

            time.sleep(0.15)  # 等待粘贴完成

            # 恢复剪贴板
            try:
                pyperclip.copy(old_clipboard)
            except:
                pass

            return True

        except Exception as e:
            print(f"[错误] 发送消息失败: {e}")
            return False
