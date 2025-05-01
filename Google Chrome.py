
from selenium.webdriver.chrome import options
from DrissionPage import ChromiumOptions

# 设置 Chrome 可执行文件路径
path = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'  # 请改为你电脑内 chrome 可执行文件路径

# 设置 Chromium 选项并保存
ChromiumOptions().set_browser_path(path).save()



