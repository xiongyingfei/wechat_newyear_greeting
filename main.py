# -*- coding: utf-8 -*-
import os
import sys
import random
import yaml
import keyboard
import time
import threading
import ctypes
from ctypes import wintypes
from src.wechat_helper import WeChatHelper

# Windows 常量
WM_HOTKEY = 0x0312
MOD_ALT = 0x0001
MOD_CTRL = 0x0002
MOD_SHIFT = 0x0004
MOD_NOREPEAT = 0x4000  # 防止按住时重复触发


def load_config(config_path="config/replies.yaml"):
    config_file = os.path.join(os.path.dirname(__file__), config_path)
    if not os.path.exists(config_file):
        print(f"配置文件不存在: {config_file}")
        sys.exit(1)
    with open(config_file, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def parse_name(full_name):
    if not full_name:
        return "", ""
    full_name = full_name.strip()

    # 处理中文姓名（默认）
    if any('\u4e00' <= c <= '\u9fff' for c in full_name):
        if len(full_name) >= 2:
            surname = full_name[0]
            given_name = full_name[1:]
        else:
            surname = full_name
            given_name = ""
        return surname, given_name
    else:
        # 处理英文姓名（按空格分隔）
        parts = full_name.split()
        if len(parts) >= 2:
            surname = parts[-1]
            given_name = " ".join(parts[:-1])
        elif len(parts) == 1:
            surname = parts[0]
            given_name = ""
        else:
            surname = ""
            given_name = ""
        return surname, given_name


def generate_reply(replies, contact_info):
    if not replies:
        return None
    template = random.choice(replies)
    if contact_info:
        result = template.replace("{name}", contact_info.full_name)
        result = result.replace("{surname}", contact_info.surname)
        result = result.replace("{given_name}", contact_info.given_name)
        return result
    result = template.replace("{name}", "").replace("{surname}", "").replace("{given_name}", "")
    return result


def run_hotkey_loop(hotkeys_config, callback):
    """
    使用 Windows RegisterHotKey 注册系统级快捷键
    优势：不受键盘钩子干扰，自动抑制按键传递，不会出现残留字符
    """
    id_to_hotkey = {}

    for hotkey_str in hotkeys_config.keys():
        parts = hotkey_str.lower().split('+')
        modifiers = MOD_NOREPEAT  # 防止按住时重复触发
        vk = None

        for part in parts:
            if part == 'alt':
                modifiers |= MOD_ALT
            elif part == 'ctrl':
                modifiers |= MOD_CTRL
            elif part == 'shift':
                modifiers |= MOD_SHIFT
            elif len(part) == 1:
                # 转换字符为虚拟键码
                if part.isdigit():
                    vk = ord(part)  # '1' -> 0x31
                else:
                    vk = ord(part.upper())

        if vk is not None:
            hk_id = len(id_to_hotkey) + 100  # 从 100 开始编号
            if ctypes.windll.user32.RegisterHotKey(None, hk_id, modifiers, vk):
                id_to_hotkey[hk_id] = hotkey_str
                print(f"[就绪] 快捷键已注册: {hotkey_str}")
            else:
                print(f"[警告] 快捷键注册失败: {hotkey_str} (可能被其他程序占用)")

    # Windows 消息循环，接收 WM_HOTKEY 消息
    msg = wintypes.MSG()
    while True:
        result = ctypes.windll.user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
        if result == 0 or result == -1:
            break
        if msg.message == WM_HOTKEY:
            hk_id = msg.wParam
            if hk_id in id_to_hotkey:
                callback(id_to_hotkey[hk_id])
        ctypes.windll.user32.TranslateMessage(ctypes.byref(msg))
        ctypes.windll.user32.DispatchMessageW(ctypes.byref(msg))

    # 程序退出时注销快捷键
    for hk_id in id_to_hotkey:
        ctypes.windll.user32.UnregisterHotKey(None, hk_id)


def main():
    config = load_config()
    wechat = WeChatHelper()
    
    if not wechat.init_wechat():
        print("请确保微信已登录并打开微信PC客户端")
        return
    
    print("=" * 55)
    print("       微信拜年回复助手 - 自动化版本")
    print("=" * 55)
    print("快捷键说明:")
    for hotkey, data in config.get("hotkeys", {}).items():
        print(f"  {hotkey} → {data['name']}")
    print("  Ctrl+Alt+N → 手动输入联系人姓名（备用）")
    print("=" * 55)
    print("按 Ctrl+Shift+Q 停止程序")
    print("=" * 55)
    print("\n使用方式：在微信聊天窗口中直接按快捷键")
    print("  程序会自动截图识别联系人姓名，生成回复并粘贴到输入框")
    print("  识别失败时可用 Ctrl+Alt+N 手动输入联系人姓名\n")
    
    def input_contact_name():
        print("\n请输入联系人姓名: ", end="", flush=True)
        try:
            name = input().strip()
            if name:
                wechat.set_contact_name(name)
            else:
                print("未输入姓名")
        except Exception as e:
            print(f"输入错误: {e}")
    
    # 锁：防止多次快捷键同时处理导致冲突
    hotkey_lock = threading.Lock()

    def _handle_hotkey(hotkey):
        """在独立线程中处理快捷键事件"""
        if not hotkey_lock.acquire(blocking=False):
            print("\n[提示] 上一条消息还在处理中，请稍候...")
            return
        try:
            hotkey_data = config.get("hotkeys", {}).get(hotkey)
            if not hotkey_data:
                return

            contact_info = wechat.get_current_contact()

            if contact_info:
                print(f"\n[获取] 联系人: {contact_info.full_name}")
            else:
                print("\n[警告] 未能获取联系人姓名，将使用默认回复")

            reply = generate_reply(hotkey_data.get("replies", []), contact_info)

            if reply:
                wechat.send_message(reply)
                target = contact_info.full_name if contact_info else "当前聊天"
                print(f"[{hotkey_data['name']}] -> {target}: {reply}\n")
        except Exception as e:
            print(f"\n[错误] 处理快捷键出错: {e}")
        finally:
            hotkey_lock.release()

    def on_hotkey(hotkey):
        """快捷键回调：启动新线程处理，避免阻塞消息循环"""
        threading.Thread(target=_handle_hotkey, args=(hotkey,), daemon=True).start()

    def set_contact():
        """手动设置联系人"""
        input_contact_name()

    def stop_program():
        print("\n程序已停止")
        sys.exit(0)

    # 用 RegisterHotKey 注册 Alt+1~5（系统级，不受键盘钩子干扰）
    hotkey_thread = threading.Thread(
        target=run_hotkey_loop,
        args=(config.get("hotkeys", {}), on_hotkey),
        daemon=True
    )
    hotkey_thread.start()

    # Ctrl+Alt+N 和 Ctrl+Shift+Q 继续用 keyboard 库（它们不涉及合成按键冲突）
    keyboard.add_hotkey("ctrl+alt+n", set_contact)
    keyboard.add_hotkey("ctrl+shift+q", stop_program)

    # 主循环
    try:
        keyboard.wait()
    except KeyboardInterrupt:
        stop_program()


if __name__ == "__main__":
    main()
