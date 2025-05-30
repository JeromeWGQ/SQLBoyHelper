import pyautogui
import time
import subprocess
import pyperclip
import sys

# 确保安装 OpenCV
try:
    import cv2
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "opencv-python"])
    import cv2

# 1. 用 Chrome 打开指定网址
url = "https://xt.sankuai.com/workbench/task/hmart_bikedw.app_bike_new_fault_detail"
subprocess.Popen(['open', '-a', 'Google Chrome', url])  # macOS 使用 'open'

# 2. 等待网页加载完成
time.sleep(5)  # 根据网络情况调整等待时间

# 3. 根据 line1.png 寻找点击位置并点击
image_path = 'line1-new.png'
location = None
while location is None:
    location = pyautogui.locateOnScreen(image_path, confidence=0.8)
    time.sleep(1)

# 保存当前屏幕截图以便调试
screenshot = pyautogui.screenshot()
screenshot.save('current_screen.png')

# 高分屏缩放比例
scale_factor = 2  # Retina 屏通常是2倍缩放

# 调整点击位置
center = pyautogui.center(location)
click_x = center.x / scale_factor
click_y = center.y / scale_factor

pyautogui.click(click_x, click_y)

# 4. 按 Command+A
pyautogui.hotkey('command', 'a')

# 5. 按 Command+C
pyautogui.hotkey('command', 'c')

# 6. 控制台输出剪贴板内容
time.sleep(1)  # 等待剪贴板内容更新
clipboard_content = pyperclip.paste()
print(clipboard_content)
