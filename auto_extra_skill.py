from win32api import GetModuleHandle
import ctypes
from ctypes import wintypes
import time
import json
import threading
import random
import os
import sys
import pystray
from PIL import Image, ImageDraw
import logging
import pydirectinput

logging.basicConfig(
    filename='app_error.log',
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s'
)


# 定义必要的Windows API结构
class KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ('vkCode', wintypes.DWORD),
        ('scanCode', wintypes.DWORD),
        ('flags', wintypes.DWORD),
        ('time', wintypes.DWORD),
        ('dwExtraInfo', ctypes.POINTER(ctypes.c_ulong))
    ]


# 设置钩子回调函数原型
HOOKPROC = ctypes.WINFUNCTYPE(
    ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.POINTER(KBDLLHOOKSTRUCT))

# 加载用户32.dll
user32 = ctypes.WinDLL('user32', use_last_error=True)

# 定义API函数
user32.SetWindowsHookExA.restype = ctypes.c_void_p
user32.SetWindowsHookExA.argtypes = (
    ctypes.c_int, HOOKPROC, wintypes.HINSTANCE, wintypes.DWORD)
user32.CallNextHookEx.restype = ctypes.c_long
user32.CallNextHookEx.argtypes = (
    ctypes.c_void_p, ctypes.c_int, ctypes.c_int, ctypes.POINTER(KBDLLHOOKSTRUCT))
user32.UnhookWindowsHookEx.restype = wintypes.BOOL
user32.UnhookWindowsHookEx.argtypes = (ctypes.c_void_p,)

# 定义常量
WH_KEYBOARD_LL = 13
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
WM_SYSKEYDOWN = 0x0104
WM_SYSKEYUP = 0x0105

# 虚拟键码字典
VK_CODES = {
    "a": 0x41, "b": 0x42, "c": 0x43, "d": 0x44, "e": 0x45, "f": 0x46, "g": 0x47,
    "h": 0x48, "i": 0x49, "j": 0x4A, "k": 0x4B, "l": 0x4C, "m": 0x4D, "n": 0x4E,
    "o": 0x4F, "p": 0x50, "q": 0x51, "r": 0x52, "s": 0x53, "t": 0x54, "u": 0x55,
    "v": 0x56, "w": 0x57, "x": 0x58, "y": 0x59, "z": 0x5A,
    "0": 0x30, "1": 0x31, "2": 0x32, "3": 0x33, "4": 0x34, "5": 0x35, "6": 0x36,
    "7": 0x37, "8": 0x38, "9": 0x39,
    "f1": 0x70, "f2": 0x71, "f3": 0x72, "f4": 0x73, "f5": 0x74, "f6": 0x75,
    "f7": 0x76, "f8": 0x77, "f9": 0x78, "f10": 0x79, "f11": 0x7A, "f12": 0x7B,
    "esc": 0x1B, "enter": 0x0D, "tab": 0x09, "space": 0x20, "backspace": 0x08,
    "delete": 0x2E, "insert": 0x2D, "home": 0x24, "end": 0x23, "pageup": 0x21, "pagedown": 0x22,
    "up": 0x26, "down": 0x28, "left": 0x25, "right": 0x27,
    "ctrl": 0x11, "alt": 0x12, "shift": 0x10, "win": 0x5B,
    "printscreen": 0x2C, "scrolllock": 0x91, "pause": 0x13,
    "capslock": 0x14, "numlock": 0x90,
    "numpad0": 0x60, "numpad1": 0x61, "numpad2": 0x62, "numpad3": 0x63, "numpad4": 0x64,
    "numpad5": 0x65, "numpad6": 0x66, "numpad7": 0x67, "numpad8": 0x68, "numpad9": 0x69,
    "add": 0x6B, "subtract": 0x6D, "multiply": 0x6A, "divide": 0x6F, "decimal": 0x6E,
    "comma": 0xBC, "period": 0xBE, "semicolon": 0xBA, "quote": 0xDE, "bracket_open": 0xDB,
    "bracket_close": 0xDD, "backslash": 0xDC, "slash": 0xBF, "grave": 0xC0, "dash": 0xBD,
    "equal": 0xBB
}

# 反向映射：虚拟键码 -> 键名
CODE_TO_NAME = {v: k for k, v in VK_CODES.items()}

# pydirectinput 键名映射（解决特殊键兼容性问题）
PDI_KEY_MAP = {
    "win": "winleft",
    "ctrl": "ctrlleft",
    "alt": "altleft",
    "shift": "shiftleft",
    "enter": "enter",
    "esc": "esc",
    "space": "space",
    "tab": "tab",
    "backspace": "backspace",
    "delete": "delete",
    "insert": "insert",
    "home": "home",
    "end": "end",
    "pageup": "pageup",
    "pagedown": "pagedown",
    "up": "up",
    "down": "down",
    "left": "left",
    "right": "right",
    "f1": "f1",
    "f2": "f2",
    "f3": "f3",
    "f4": "f4",
    "f5": "f5",
    "f6": "f6",
    "f7": "f7",
    "f8": "f8",
    "f9": "f9",
    "f10": "f10",
    "f11": "f11",
    "f12": "f12"
}

# 全局钩子句柄
hook_id = None

# 全局状态变量
active = False
hotkey_pressed = False
exit_flag = False
ignore_next_keys = False

# 配置变量
toggle_key_vk = None
trigger_keys = []  # 触发键序列
extra_sequence = []  # 额外按键序列
sequence_delay = 10
# 存储虚拟键码的集合，用于快速判断
trigger_vks = set()
extra_sequence_vks = set()


# 托盘图标功能
def resource_path(relative_path):
    try:
        if hasattr(sys, '_MEIPASS'):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
        path = os.path.join(base_path, relative_path)
        if not os.path.exists(path):
            logging.warning(f"资源文件不存在: {path}")
        return path
    except Exception as e:
        logging.error(f"资源路径获取失败: {e}")
        return relative_path


def create_tray_icon():
    try:
        icon_path = resource_path("keyboard.ico")
        if not os.path.exists(icon_path):
            logging.warning(f"图标文件不存在: {icon_path}")
            raise FileNotFoundError
        icon_image = Image.open(icon_path)
        logging.info("图标加载成功")
    except Exception as e:
        logging.error(f"图标加载失败: {e}, 使用备用图标")
        image = Image.new('RGB', (64, 64), (30, 144, 255))
        dc = ImageDraw.Draw(image)
        dc.rectangle((20, 20, 44, 44), fill=(255, 255, 255))
        icon_image = image

    def on_quit(icon):
        global exit_flag
        try:
            logging.info("收到退出请求")
            icon.stop()
            exit_flag = True
            if hook_id:
                user32.UnhookWindowsHookEx(hook_id)
            thread_id = threading.get_ident()
            logging.info(f"发送退出消息到线程: {thread_id}")
            user32.PostThreadMessageA(ctypes.wintypes.DWORD(thread_id), 0x0012, 0, 0)
        except Exception as e:
            logging.error(f"退出处理失败: {e}")

    menu = pystray.Menu(
        pystray.MenuItem('退出程序', on_quit)
    )
    try:
        icon = pystray.Icon(
            "AutoExtraSkill",
            icon=icon_image,
            title="AutoExtraSkill",
            menu=menu
        )
        logging.info("托盘图标创建成功")
        icon.run()
    except Exception as e:
        logging.critical(f"托盘图标运行失败: {e}")


# 加载配置
def load_config(config_file="config.json"):
    global toggle_key_vk, trigger_keys, extra_sequence, sequence_delay
    global trigger_vks, extra_sequence_vks

    # 默认配置
    default_toggle_key = "numpad9"
    default_trigger_keys = ["a", "s", "d", "f", "g", "h"]
    default_sequence = ["e", "q", "z"]
    default_delay = 100

    if not os.path.exists(config_file):
        print(f"配置文件 {config_file} 不存在，使用默认配置")
        toggle_key_vk = VK_CODES.get(default_toggle_key)
        trigger_keys = default_trigger_keys
        extra_sequence = default_sequence
        sequence_delay = default_delay
    else:
        try:
            with open(config_file, "r") as f:
                config = json.load(f)

                # 加载切换键
                toggle_key = config.get("toggle_key", default_toggle_key).lower()
                toggle_key_vk = VK_CODES.get(toggle_key)
                if toggle_key_vk is None:
                    print(f"警告: 无效的切换键 '{toggle_key}'，使用默认值 '{default_toggle_key}'")
                    toggle_key_vk = VK_CODES.get(default_toggle_key)

                # 加载触发键序列
                trigger_keys = config.get("trigger_keys", default_trigger_keys)
                # 验证触发键是否有效
                valid_triggers = []
                for key in trigger_keys:
                    key_lower = key.lower()
                    if key_lower in VK_CODES:
                        valid_triggers.append(key_lower)
                    else:
                        print(f"警告: 无效的触发键 '{key}'，已从序列中移除")
                trigger_keys = valid_triggers
                if not trigger_keys:
                    print(f"警告: 有效触发键序列为空，使用默认序列")
                    trigger_keys = default_trigger_keys

                # 加载额外按键序列
                extra_sequence = config.get("extra_sequence", default_sequence)
                # 验证序列中的键是否有效
                valid_sequence = []
                for key in extra_sequence:
                    key_lower = key.lower()
                    if key_lower in VK_CODES:
                        valid_sequence.append(key_lower)
                    else:
                        print(f"警告: 无效的额外键 '{key}'，已从序列中移除")
                extra_sequence = valid_sequence
                if not extra_sequence:
                    print(f"警告: 有效额外按键序列为空，使用默认序列 {default_sequence}")
                    extra_sequence = default_sequence

                # 加载延迟时间
                sequence_delay = config.get("sequence_delay", default_delay)
                if not isinstance(sequence_delay, (int, float)) or sequence_delay < 0:
                    print(f"警告: 无效的延迟值 {sequence_delay}，使用默认值 {default_delay}")
                    sequence_delay = default_delay

        except Exception as e:
            print(f"加载配置失败: {e}，使用默认配置")
            toggle_key_vk = VK_CODES.get(default_toggle_key)
            trigger_keys = default_trigger_keys
            extra_sequence = default_sequence
            sequence_delay = default_delay

    # 生成虚拟键码集合，用于快速判断
    trigger_vks.clear()
    for key in trigger_keys:
        trigger_vks.add(VK_CODES[key])

    extra_sequence_vks.clear()
    for key in extra_sequence:
        extra_sequence_vks.add(VK_CODES[key])

    return True


# 随机化
def randomize(t):
    return t * (1 + 0.2 *random.random())


# 执行额外按键序列
def execute_extra_sequence():
    global extra_sequence, sequence_delay, ignore_next_keys
    # 设置标志，忽略接下来由程序自动触发的按键
    ignore_next_keys = True
    try:
        for key_name in extra_sequence:
            # 转换特殊键名称
            pdi_key = PDI_KEY_MAP.get(key_name, key_name)
            # 点击按键
            time.sleep(randomize(sequence_delay) / 1000.0)
            pydirectinput.keyDown(pdi_key)
            time.sleep(randomize(0.01))
            pydirectinput.keyUp(pdi_key)
            # 按键之间的延迟
            time.sleep(randomize(sequence_delay)/ 1000.0)
    finally:
        # 重置标志，恢复正常检测
        ignore_next_keys = False


# 键盘钩子处理函数
def low_level_keyboard_handler(nCode, wParam, lParam):
    global active, hotkey_pressed, exit_flag, toggle_key_vk, ignore_next_keys
    global trigger_vks, extra_sequence_vks

    if exit_flag:
        return user32.CallNextHookEx(hook_id, nCode, wParam, lParam)
    if nCode != 0:
        return user32.CallNextHookEx(hook_id, nCode, wParam, lParam)

    kbd_struct = ctypes.cast(lParam, ctypes.POINTER(KBDLLHOOKSTRUCT)).contents

    # 处理切换键（启动/暂停）
    if kbd_struct.vkCode == toggle_key_vk:
        if wParam in [WM_KEYDOWN, WM_SYSKEYDOWN]:
            if not hotkey_pressed:
                active = not active
                status = "激活" if active else "暂停"
                print(f"状态切换: {status}")
                hotkey_pressed = True
            return 1
        elif wParam in [WM_KEYUP, WM_SYSKEYUP]:
            hotkey_pressed = False
            return 1

    # 如果程序未激活，不做额外处理
    if not active:
        return user32.CallNextHookEx(hook_id, nCode, wParam, lParam)

    # 如果是程序自动触发的按键，忽略它
    if ignore_next_keys:
        return user32.CallNextHookEx(hook_id, nCode, wParam, lParam)

    # 检测到触发键被按下，执行额外序列
    if (wParam in [WM_KEYDOWN, WM_SYSKEYDOWN] and
            kbd_struct.vkCode != toggle_key_vk and
            kbd_struct.vkCode in trigger_vks):

        # 使用线程执行，避免阻塞
        threading.Thread(
            target=execute_extra_sequence,
            daemon=True
        ).start()

    return user32.CallNextHookEx(hook_id, nCode, wParam, lParam)


def main():
    global hook_id, exit_flag
    exit_flag = False

    # 加载配置
    load_config()
    toggle_key_name = CODE_TO_NAME.get(toggle_key_vk, "未知键")
    print(f"程序切换键设置为: {toggle_key_name}")
    print(f"触发键序列: {trigger_keys}")
    print(f"额外按键序列: {extra_sequence}")
    print(f"按键间隔延迟: {sequence_delay}ms")

    # 设置键盘钩子
    callback = HOOKPROC(low_level_keyboard_handler)
    hook_id = user32.SetWindowsHookExA(
        WH_KEYBOARD_LL, callback, GetModuleHandle(None), 0)
    if not hook_id:
        print("设置键盘钩子失败")
        return

    try:
        # 启动托盘图标
        tray_thread = threading.Thread(target=create_tray_icon, daemon=True)
        tray_thread.start()
        time.sleep(1)

        print(f"程序已启动，按 {toggle_key_name.upper()} 切换激活状态")
        print("仅当按下触发键列表中的键时会触发额外按键序列")
        print("可通过右下角托盘图标退出程序")

        # 消息循环
        msg = wintypes.MSG()
        while True:
            if exit_flag:
                print("检测到退出标志，退出循环")
                break
            ret = user32.PeekMessageA(ctypes.byref(msg), 0, 0, 0, 1)
            if ret != 0:
                if msg.message == 0x0012:  # WM_QUIT
                    break
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageA(ctypes.byref(msg))
            else:
                time.sleep(0.001)
    except Exception as e:
        import traceback
        error_msg = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] ERROR: {str(e)}"
        with open("error.log", "a") as f:
            f.write(error_msg + "\n")
            f.write(traceback.format_exc())
    finally:
        if hook_id:
            user32.UnhookWindowsHookEx(hook_id)
        print("程序已退出")
        os._exit(0)


if __name__ == "__main__":
    pydirectinput.FAILSAFE = False
    pydirectinput.PAUSE = 0.001
    main()
