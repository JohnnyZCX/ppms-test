import json
import logging
import os
import sys
import time
import traceback
from configparser import ConfigParser
from io import BytesIO
from logging import handlers

from PIL import Image
from selenium import webdriver
from selenium.common import WebDriverException
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

CUR_PATH = os.path.abspath(os.path.dirname(__file__))


def init_logger():
    logger = logging.getLogger()
    logging.logThreads = False
    # logger.setLevel(logging.INFO)
    sh = logging.StreamHandler(sys.stdout)
    sh.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    sh.setFormatter(formatter)
    logger.addHandler(sh)

    rq = time.strftime('%m-%d', time.localtime(time.time()))
    log_path = CUR_PATH + '/logs/'
    if not os.path.exists(log_path):
        os.mkdir(log_path)
    logfile = log_path + rq + '.log'
    # logging.FileHandler
    fh = handlers.TimedRotatingFileHandler(logfile, when='d', interval=1, backupCount=365)
    fh.suffix = "%Y-%m-%d.log"
    fh.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s - %(filename)s[line:%(lineno)d] - %(levelname)s: %(message)s")
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    logger.setLevel(logging.INFO)

    return logger


g_logger = init_logger()


def read_json_data(path):
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data


def write_json_data(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, separators=(',', ':'))


def load_cfg(part_name=None, encoding="utf-8"):
    cp = ConfigParser()
    cp.read(CUR_PATH + "\\system.cfg", encoding=encoding)

    if part_name:
        # print(part_name, path)
        if part_name in cp:
            return cp[part_name]
        return None
    return cp


def json2str(obj):
    return json.dumps(obj, ensure_ascii=False, separators=(',', ':'))


def rstrip(s: str, p: str):
    if s.endswith(p):
        return s[:-len(p)]
    else:
        return s


def get_chrome():
    driver_config = load_cfg("driver")
    options = ChromeOptions()
    if driver_config.getboolean("disableExtensions"):
        options.add_argument("--disable-extensions")
    if driver_config.get("windowSize"):
        options.add_argument(f"window-size={driver_config.get('windowSize')}")
    if driver_config.getboolean("headless"):
        options.add_argument("--headless")
    service = Service(driver_config.get("path"))
    return Chrome(service=service, options=options)


def new_chrome():
    chrome_options = Options()
    # chrome_options.add_argument('--headless')  # 启用无头模式
    chrome_options.add_argument("--force-device-scale-factor=0.85")  # 设置为 85% 缩放
    driver = webdriver.Chrome(options=chrome_options, service=ChromeService(ChromeDriverManager().install()))
    return driver


def page_screenshot(driver, image_path, document, image_name):
    """

    :param driver: 浏览器驱动对象
    :param image_path: 图片存放路径
    :param doc: word文件对象
    :param image_name: 图片名称
    :return:
    """
    # 获取网页截图
    screenshot = driver.get_screenshot_as_png()
    image = Image.open(BytesIO(screenshot))
    image.save(image_path)  # 保存截图到文件
    # 将截图文件写入到word文档
    document.add_heading(image_name, 4)
    document.add_picture(image_path)


def retry(max_retries):
    def wrapper1(func):
        def wrapper(*args, **kwargs):

            try:
                return func(*args, **kwargs)
            except Exception as e:
                g_logger.error(f"运行错误-{e}")
                connect_ok = False
                try_number = 0
                res = None
                if try_number >= max_retries:
                    sys.exit(1)  # 程序退出
                while not connect_ok and try_number < max_retries:
                    time.sleep(5)
                    try:
                        res = func(*args, **kwargs)
                        connect_ok = True
                    except Exception as e:
                        if isinstance(e, WebDriverException):
                            self = args[0]
                            self.re_connet()
                            print(f'DevTools connection attempt {try_number + 1} failed, retrying...')
                        print(traceback.format_exc())
                        g_logger.error(f"运行错误重试{try_number + 1}次-{e}")
                        try_number += 1
                return res

        return wrapper

    return wrapper1
