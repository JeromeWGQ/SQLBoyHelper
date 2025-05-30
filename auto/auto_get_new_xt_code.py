import pyautogui
import time
import subprocess
import pyperclip
import sys

# 1. 用 Chrome 打开指定网址
url = "https://xt.sankuai.com/workbench/task/hmart_bikedw.app_bike_new_fault_detail"
subprocess.Popen(['open', '-a', 'Google Chrome', url])

# 2. 等待网页加载完成
time.sleep(5)

# 3. 执行操作序列
pyautogui.write('-')
pyautogui.press('backspace')
pyautogui.hotkey('command', 'a')
time.sleep(1)
pyautogui.hotkey('command', 'c')
time.sleep(1)

# 4. 获取并输出剪贴板内容
clipboard_content = pyperclip.paste()
print(clipboard_content)
