import pyautogui
import pygetwindow as gw
import win32gui  # 添加 win32gui 支持
import win32con  # 添加 win32con 支持
import time
import requests
import os
import json
import pyperclip  # 添加剪贴板支持
from typing import Optional
from dotenv import load_dotenv

# 加载.env文件
load_dotenv()

# 配置PyAutoGUI的安全设置
pyautogui.FAILSAFE = True  # 启用故障安全
pyautogui.PAUSE = 0.5  # 增加操作间隔时间

# DeepSeek API配置
DEEPSEEK_API_KEY = os.getenv('DEEPSEEK_API_KEY')
DEEPSEEK_API_URL = "https://api.siliconflow.cn/v1/chat/completions"

# 配置参数
JOKE_INTERVAL = 10  # 每次请求笑话的间隔时间（秒）
TYPING_DELAY = 0.1  # 打字间隔时间（秒）

# 全局变量：上次成功发送笑话的时间
last_joke_time = 0
# 全局变量：是否正在处理笑话
is_processing = False

def paste_text(text: str):
    """使用剪贴板粘贴文本"""
    if not text:
        print("警告: 尝试输入空文本")
        return
        
    print(f"准备通过剪贴板输入文本: {text}")
    try:
        # 保存当前剪贴板内容
        original = pyperclip.paste()
        
        # 将文本复制到剪贴板
        pyperclip.copy(text)
        time.sleep(0.2)  # 等待剪贴板更新
        
        # 执行粘贴操作
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.2)  # 等待粘贴完成
        
        # 恢复原始剪贴板内容
        pyperclip.copy(original)
        
        print("文本输入完成")
    except Exception as e:
        print(f"粘贴文本时出错: {e}")

def get_joke() -> Optional[str]:
    """调用DeepSeek API获取一个随机笑话"""
    if not DEEPSEEK_API_KEY:
        print("未找到DEEPSEEK_API_KEY环境变量，请检查.env文件")
        return None
    
    print(f"正在发送API请求到 {DEEPSEEK_API_URL}")
    print(f"使用的API密钥: {DEEPSEEK_API_KEY[:8]}...")
    
    headers = {
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
        "Content-Type": "application/json"
    }
    
    data = {
        "messages": [
            {
                "role": "user",
                "content": "请讲一个简短有趣的笑话，直接给出笑话内容，不要有多余的话"
            }
        ],
        "model": "deepseek-ai/DeepSeek-V3",
        "temperature": 0.7,
        "max_tokens": 150
    }
    
    try:
        print("发送请求数据:", json.dumps(data, ensure_ascii=False, indent=2))
        response = requests.post(DEEPSEEK_API_URL, headers=headers, json=data)
        print("API响应状态码:", response.status_code)
        
        # 尝试解析响应JSON
        try:
            response_json = response.json()
            print("API响应JSON:", json.dumps(response_json, ensure_ascii=False, indent=2))
            
            if 'choices' in response_json and len(response_json['choices']) > 0:
                joke = response_json['choices'][0]['message']['content']
                print(f"成功提取笑话内容: {joke}")
                return joke.strip()
            else:
                print("API响应中未找到笑话内容")
                return None
                
        except json.JSONDecodeError as je:
            print(f"解析JSON响应失败: {je}")
            print("原始响应内容:", response.text)
            return None
            
    except Exception as e:
        print(f"获取笑话时出错: {e}")
        if 'response' in locals():
            print(f"错误响应内容: {response.text if response else '无响应'}")
        return None

def is_notepad_open():
    """检查记事本是否打开"""
    try:
        windows = gw.getWindowsWithTitle('Notepad')
        return len(windows) > 0
    except Exception as e:
        print(f"检查记事本窗口时出错: {e}")
        return False

def get_notepad_window():
    """获取记事本窗口"""
    try:
        # 尝试多种方式获取记事本窗口
        windows = gw.getWindowsWithTitle('Notepad')
        if not windows:
            # 尝试获取中文标题的记事本
            windows = gw.getWindowsWithTitle('记事本')
        if not windows:
            # 尝试获取带有文件名的记事本窗口
            all_windows = gw.getAllWindows()
            windows = [w for w in all_windows if 'Notepad' in w.title or '记事本' in w.title]
        
        if windows:
            return windows[0]
        print("未找到记事本窗口")
    except Exception as e:
        print(f"获取记事本窗口时出错: {e}")
    return None

def activate_notepad_window(window):
    """激活记事本窗口"""
    if not window:
        print("无效的窗口对象")
        return False
        
    max_retries = 3
    retry_count = 0
    
    while retry_count < max_retries:
        try:
            # 获取窗口句柄
            hwnd = win32gui.FindWindow(None, window.title)
            if hwnd:
                # 确保窗口未最小化
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    time.sleep(0.5)
                
                # 将窗口置前
                win32gui.SetForegroundWindow(hwnd)
                
                # 等待窗口激活
                time.sleep(0.5)
                
                # 验证窗口是否真的激活
                active_hwnd = win32gui.GetForegroundWindow()
                if active_hwnd == hwnd:
                    print("记事本窗口已成功激活")
                    return True
                    
            retry_count += 1
            if retry_count < max_retries:
                print(f"激活窗口重试中... ({retry_count}/{max_retries})")
                time.sleep(1)
            
        except Exception as e:
            print(f"激活窗口时出错 (尝试 {retry_count + 1}/{max_retries}): {e}")
            retry_count += 1
            if retry_count < max_retries:
                time.sleep(1)
    
    print(f"在 {max_retries} 次尝试后仍无法激活窗口")
    return False

def is_cursor_in_window(window):
    """检查鼠标是否在窗口内"""
    try:
        x, y = pyautogui.position()
        in_window = (window.left < x < window.right and 
                    window.top < y < window.bottom)
        if in_window:
            print("鼠标在记事本窗口内")
        return in_window
    except Exception as e:
        print(f"检查鼠标位置时出错: {e}")
        return False

print("程序已启动，等待记事本打开...")
print(f"笑话获取间隔设置为 {JOKE_INTERVAL} 秒")

# 主循环
while True:
    try:
        current_time = time.time()
        time_since_last = current_time - last_joke_time
        
        # 首先检查时间间隔
        if time_since_last < JOKE_INTERVAL or is_processing:
            remaining = JOKE_INTERVAL - time_since_last
            if remaining > 0:
                print(f"\r等待中... 还需 {remaining:.1f} 秒", end='')
                time.sleep(1)
                continue
        
        # 重置处理标志
        is_processing = False
        
        if is_notepad_open():
            notepad_window = get_notepad_window()
            if notepad_window:
                # 检查鼠标位置
                x, y = pyautogui.position()
                print(f"\n当前鼠标位置: x={x}, y={y}")
                print(f"记事本窗口位置: 左={notepad_window.left}, 上={notepad_window.top}, "
                      f"右={notepad_window.right}, 下={notepad_window.bottom}")
                
                # 确保窗口处于激活状态
                if activate_notepad_window(notepad_window):
                    if is_cursor_in_window(notepad_window):
                        print("\n开始新一轮笑话获取...")
                        is_processing = True  # 标记正在处理
                        
                        joke = get_joke()
                        if joke:
                            print(f"\n准备输入笑话: {joke}")
                            print("开始输入...")
                            
                            try:
                                # 输入时间戳
                                timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
                                paste_text(f"[{timestamp}]\n")
                                time.sleep(0.5)
                                
                                # 输入笑话内容
                                paste_text(joke)
                                time.sleep(0.5)
                                
                                # 添加换行
                                pyautogui.press('enter')
                                time.sleep(0.2)
                                pyautogui.press('enter')
                                
                                last_joke_time = current_time
                                print("笑话输入完成")
                            except Exception as e:
                                print(f"输入文本时出错: {e}")
                                is_processing = False  # 重置处理标志
                        else:
                            print("未能获取到笑话，跳过本次输入")
                            is_processing = False  # 重置处理标志
                    else:
                        print("\r鼠标不在记事本窗口内", end='')
                else:
                    print("\r无法激活记事本窗口", end='')
        else:
            print("\r等待记事本打开...", end='')
        
        time.sleep(1)  # 每秒检查一次
        
    except Exception as e:
        print(f"主循环出错: {e}")
        is_processing = False  # 发生错误时重置处理标志
        
