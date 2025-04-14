import time
import unittest

import ddddocr
import docx
import openpyxl
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

import utils

MAX_TRIES = 4

# cfg = utils.load_cfg()
wb = openpyxl.Workbook()
# 创建一个sheet并加上名称和所在位置，第一个位置索引号是0
wb.create_sheet("湖北省病虫疫情信息调度指挥中心", 1)
sheet = wb["湖北省病虫疫情信息调度指挥中心"]
# 写入表头
headers = ["页面", "检测结果"]
sheet.append(headers)

# 创建Word文档对象
doc = docx.Document()


class PPMSHB():

    def __init__(self):
        # self.cookies = {}
        self.init_chrome()

    def init_chrome(self):
        self.driver = utils.new_chrome()
        self.driver.maximize_window()

    def re_connet(self):
        self.driver.quit()
        self.driver = utils.new_chrome()
        self.driver.maximize_window()
        print("系统重启")

    @utils.retry(MAX_TRIES)
    def test_shouye(self):
        username = input("湖北省级系统生产环境巡检开始\n请输入登录用户名：")
        password = input("请输入登录密码：")
        self.driver.maximize_window()
        self.driver.get("https://nyt.hubei.gov.cn/pestiot/")
        self.driver.implicitly_wait(5)

        # 登录
        self.driver.find_element(By.XPATH, '//input[@placeholder="请输入用户名"]').send_keys(username)
        self.driver.find_element(By.XPATH, '//input[@name="password"]').send_keys(password)
        yanzhengma_image = self.driver.find_element(By.XPATH, '//img[contains(@src,"data:image/png;base64")]')
        img_bytes = yanzhengma_image.screenshot_as_png
        yzm = ddddocr.DdddOcr(show_ad=False).classification(img_bytes)
        self.driver.find_element(By.XPATH, '//input[@placeholder="请输入验证码"]').send_keys(yzm)
        time.sleep(2)
        self.driver.find_element(By.XPATH, '//button[@type="button"]').click()
        try:
            while self.driver.find_element(By.XPATH, "//*[text()='图片验证码错误']"):
                self.driver.find_element(By.XPATH, "//input[@placeholder='请输入验证码']").clear()
                yanzhengma_image = self.driver.find_element(By.XPATH, '//img[contains(@src,"data:image/png;base64")]')
                img_bytes = yanzhengma_image.screenshot_as_png
                yzm = ddddocr.DdddOcr(show_ad=False).classification(img_bytes)
                self.driver.find_element(By.XPATH, '//input[@placeholder="请输入验证码"]').send_keys(yzm)
                time.sleep(2)
                self.driver.find_element(By.XPATH, '//button[@type="button"]').click()
        except Exception:
            pass

        # 首页元素校验
        try:
            # 等待页面加载完成
            time.sleep(10)
            # 获取页面总高度和宽度
            total_height = self.driver.execute_script("return document.body.parentElement.scrollHeight")
            total_width = self.driver.execute_script("return document.body.parentElement.scrollWidth")
            # 调整窗口尺寸
            self.driver.set_window_size(total_width, total_height)
            element = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//div[@class="legendBox"]')))
            unittest.TestCase.assertTrue(element is not None, "登录成功，成功打开首页，且指定元素存在")
            utils.g_logger.info(f"登录成功，成功打开首页")
            sheet.append(["首页", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/首页.png", doc, "首页")
            self.driver.find_element(By.XPATH, "//div[@class='shrink-container']/i").click()
        except Exception as e:
            utils.g_logger.info("登录成功，但首页显示异常")
            sheet.append(["首页", "异常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/首页.png", doc, "首页")

    @utils.retry(MAX_TRIES)
    def test_jianceyubao(self):
        # 数据填报工作平台页
        try:
            time.sleep(3)
            self.driver.find_element(By.XPATH,
                                     "//div[@class='navigation-list']/div[@class='navigation-item']/span[text()=' 监测预报']").click()
            time.sleep(8)
            element = WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located(
                (By.XPATH, '//div[@class="widget_title widget_title_heading"]')))
            unittest.TestCase.assertTrue(element is not None, "成功打开数据填报工作平台，且指定元素存在")
            utils.g_logger.info("数据填报-工作平台页显示正常")
            sheet.append(["数据填报-工作平台", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据填报_工作平台.png", doc, "数据填报_工作平台")
        except Exception as e:
            utils.g_logger.info("数据填报-工作平台页显示异常")
            sheet.append(["数据填报-工作平台", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据填报_工作平台.png", doc, "数据填报_工作平台")

        # 数据填报任务填报页
        try:
            time.sleep(3)
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'el-menu-item first-menu')]/span[@style='margin-left: 20px;'][text()='数据填报']").click()
            time.sleep(10)
            self.driver.find_element(By.XPATH, "//i[@class='el-icon-arrow-down']").click()
            tree_element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//li[starts-with(@id,'reportTree')][1]")))
            unittest.TestCase.assertTrue(tree_element is not None, "成功打开任务填报页，且指定元素存在")
            utils.g_logger.info("数据填报-任务填报页显示正常")
            sheet.append(["数据填报-任务填报", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据填报_任务填报.png", doc, "数据填报_任务填报")
        except Exception as e:
            utils.g_logger.info("数据填报-任务填报页显示异常")
            sheet.append(["数据填报-任务填报", "异常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据填报_任务填报.png", doc, "数据填报_任务填报")

        # 数据填报数据查询页
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'el-menu-item first-menu')]/span[@style='margin-left: 20px;'][text()='数据查询']").click()
            time.sleep(10)
            self.driver.find_element(By.XPATH, "//div[@id='reportboxTree']//i[@class='el-icon-arrow-down']").click()
            tree_element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//li[starts-with(@id,'reportTree')][1]")))
            unittest.TestCase.assertTrue(tree_element is not None, "成功打开数据填报-数据查询页，且报表列表存在")
            self.driver.find_element(By.XPATH, "//div[@id='orgboxzTree']//i[@class='el-icon-arrow-down']").click()
            orgnization_tree_element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//ul[starts-with(@id,'orgTree') and @class='ztree']")))
            unittest.TestCase.assertTrue(orgnization_tree_element is not None,
                                         "成功打开数据填报-数据查询页，且站点列表存在")
            utils.g_logger.info("数据填报-数据查询页显示正常")
            sheet.append(["数据填报-数据查询", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据填报_数据查询.png", doc, "数据填报_数据查询")
        except Exception as e:
            utils.g_logger.info("数据填报-数据查询页显示异常")
            sheet.append(["数据填报-数据查询", "异常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据填报_数据查询.png", doc, "数据填报_数据查询")

        # 数据填报数据汇总页
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[@style='margin-left: 20px;'][text()='数据汇总']").click()
            time.sleep(10)
            self.driver.find_element(By.XPATH, "//div[@id='reportboxTree']//i[@class='el-icon-arrow-down']").click()
            tree_element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//li[starts-with(@id,'reportTree')][1]")))
            unittest.TestCase.assertTrue(tree_element is not None, "成功打开数据汇总页，且报表列表存在")
            self.driver.find_element(By.XPATH, "//div[@id='orgboxzTree']//i[@class='el-icon-arrow-down']").click()
            orgnization_tree_element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//ul[starts-with(@id,'orgTree') and @class='ztree']")))
            unittest.TestCase.assertTrue(orgnization_tree_element is not None,
                                         "成功打开数据汇总页，且站点列表存在")
            utils.g_logger.info("数据填报-数据汇总页显示正常")
            sheet.append(["数据填报-数据汇总", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据填报_数据汇总.png", doc, "数据填报_数据汇总")

        except Exception as e:
            utils.g_logger.info("数据填报-数据汇总页显示异常")
            sheet.append(["数据填报-数据汇总", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据填报_数据汇总.png", doc, "数据填报_数据汇总")

        # 数据填报催报查询页
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[@style='margin-left: 20px;'][text()='催报查询']").click()
            time.sleep(10)
            self.driver.find_element(By.XPATH, "//div[@id='reportboxTree']//i[@class='el-icon-arrow-down']").click()
            tree_element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//li[starts-with(@id,'reportTree')][1]")))
            unittest.TestCase.assertTrue(tree_element is not None, "成功打开催报查询页，且报表列表存在")
            self.driver.find_element(By.XPATH, "//div[@id='orgboxzTree']//i[@class='el-icon-arrow-down']").click()
            orgnization_tree_element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//ul[starts-with(@id,'orgTree') and @class='ztree']")))
            unittest.TestCase.assertTrue(orgnization_tree_element is not None,
                                         "成功打开催报查询页，且站点列表存在")
            utils.g_logger.info("数据填报-催报查询页显示正常")
            sheet.append(["数据填报-催报查询", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据填报_催报查询.png", doc, "数据填报_催报查询")
        except Exception as e:
            utils.g_logger.info("数据填报-催报查询页显示异常")
            sheet.append(["数据填报-催报查询", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据填报_催报查询.png", doc, "数据填报_催报查询")

        # 数据填报报送评价页
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[@style='margin-left: 20px;'][text()='报送评价']").click()
            time.sleep(10)
            self.driver.find_element(By.XPATH, "//div[@id='reportboxTree']//i[@class='el-icon-arrow-down']").click()
            tree_element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//li[starts-with(@id,'reportTree')][1]")))
            unittest.TestCase.assertTrue(tree_element is not None, "成功打开报送评价页，且报表列表存在")
            self.driver.find_element(By.XPATH, "//div[@id='orgboxzTree']//i[@class='el-icon-arrow-down']").click()
            orgnization_tree_element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//ul[starts-with(@id,'orgTree') and @class='ztree']")))
            unittest.TestCase.assertTrue(orgnization_tree_element is not None,
                                         "成功打开报送评价页，且站点列表存在")
            utils.g_logger.info("数据填报-报送评价页显示正常")
            sheet.append(["数据填报-报送评价", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据填报_报送评价.png", doc, "数据填报_报送评价")
        except Exception as e:
            utils.g_logger.info("数据填报-报送评价页显示异常")
            sheet.append(["数据填报-报送评价", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据填报_报送评价.png", doc, "数据填报_报送评价")

        # 数据填报填报任务一览页
        try:
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='填报任务一览']").click()
            time.sleep(10)
            self.driver.find_element(By.XPATH, "//div[@id='orgboxzTree']//i[@class='el-icon-arrow-down']").click()
            orgnization_tree_element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//ul[starts-with(@id,'orgTree') and @class='ztree']")))
            unittest.TestCase.assertTrue(orgnization_tree_element is not None,
                                         "成功打开填报任务一览页，且站点列表存在")
            utils.g_logger.info("数据填报-填报任务一览页显示正常")
            sheet.append(["数据填报-填报任务一览", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/填报任务一览.png", doc, "填报任务一览")
        except Exception as e:
            utils.g_logger.info("数据填报-填报任务一览页显示异常")
            sheet.append(["数据填报-填报任务一览", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/填报任务一览.png", doc, "填报任务一览")

        # 数据填报填报任务设置页
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[@style='margin-left: 20px;'][text()='填报任务设置']").click()
            time.sleep(10)
            self.driver.find_element(By.XPATH, "//div[@id='orgboxzTree']//i[@class='el-icon-arrow-down']").click()
            orgnization_tree_element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//ul[starts-with(@id,'orgTree') and @class='ztree']")))
            unittest.TestCase.assertTrue(orgnization_tree_element is not None,
                                         "成功打开数据填报-填报任务设置页，且站点列表存在")
            utils.g_logger.info("数据填报-填报任务设置页显示正常")
            sheet.append(["数据填报-填报任务设置", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据填报_填报任务设置.png", doc,
                                  "数据填报_填报任务设置")
        except Exception as e:
            utils.g_logger.info("数据填报-填报任务设置页异常")
            sheet.append(["数据填报-填报任务设置", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据填报_填报任务设置.png", doc,
                                  "数据填报_填报任务设置")

        # 数据填报任务审核页
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[@style='margin-left: 20px;'][text()='任务审核']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//div[@class='search-container']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开数据填报-任务审核页，且指定元素存在")
            utils.g_logger.info("数据填报-任务审核页显示正常")
            sheet.append(["数据填报-任务审核", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据填报_任务审核.png", doc, "数据填报_任务审核")
        except Exception as e:
            utils.g_logger.info("数据填报-任务审核页异常")
            sheet.append(["数据填报-任务审核", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据填报_任务审核.png", doc, "数据填报_任务审核")

        # 执行”关闭所有“页面操作
        try:
            self.driver.find_element(By.XPATH, "//div[@class='close-con']").click()
            time.sleep(2)
            element = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//ul[@class='el-dropdown-menu el-popper']/li[text()='关闭所有']"))
            )
            element.click()
            utils.g_logger.info("成功关闭所有页面")
        except Exception as e:
            utils.g_logger.info("关闭所有页面失败")

    @utils.retry(MAX_TRIES)
    def test_xitongguanli(self):
        try:
            # 系统管理工作平台页
            time.sleep(3)
            # 点击菜单“更多”>“系统管理”
            self.driver.find_element(By.XPATH, "//span[@class='navigation-more'][text()='更多']").click()
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//ul[contains(@id,'dropdown-menu')]//span[text()=' 系统管理']").click()
            time.sleep(10)
            report_chart_element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//div[text()=' 按报表周期统计 ']/following-sibling::div//canvas")))
            unittest.TestCase.assertTrue(report_chart_element is not None,
                                         "成功打开系统管理工作平台页")
            utils.g_logger.info("系统管理-工作平台页显示正常")
            sheet.append(["系统管理-工作平台", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理工作平台.png", doc, "系统管理工作平台")
        except Exception as e:
            utils.g_logger.info("系统管理-工作平台页异常")
            sheet.append(["系统管理-工作平台", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理工作平台.png", doc, "系统管理工作平台")

        # 系统管理报表权限管理页
        time.sleep(5)
        self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='报表权限管理']").click()
        try:
            time.sleep(10)
            self.driver.find_element(By.XPATH, "//div[@id='orgboxzTree']//i[@class='el-icon-arrow-down']").click()
            orgnization_tree_element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//ul[starts-with(@id,'orgTree') and @class='ztree']")))
            unittest.TestCase.assertTrue(orgnization_tree_element is not None,
                                         "成功打开报表权限管理页，且站点列表存在")
            utils.g_logger.info("系统管理-报表权限管理页显示正常")
            sheet.append(["系统管理-报表权限管理", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/报表权限管理.png", doc, "报表权限管理")
        except Exception as e:
            utils.g_logger.info("系统管理-报表权限管理页异常")
            sheet.append(["系统管理-报表权限管理", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/报表权限管理.png", doc, "报表权限管理")

        # 系统管理机构管理页
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[@style='margin-left: 20px;'][text()='机构管理']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//label[@for='orglevel']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开系统管理-机构管理页，且指定元素存在")
            utils.g_logger.info("系统管理-机构管理页显示正常")
            sheet.append(["系统管理-机构管理", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/机构管理.png", doc, "机构管理")
        except Exception as e:
            utils.g_logger.info("系统管理-机构管理页异常")
            sheet.append(["系统管理-机构管理", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/机构管理.png", doc, "机构管理")

        # 系统管理用户管理页
        self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='用户管理']").click()
        try:
            time.sleep(13)
            element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//label[@for='username']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开系统管理-用户管理页，且指定元素存在")
            utils.g_logger.info("系统管理-用户管理页显示正常")
            sheet.append(["系统管理-用户管理", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_用户管理.png", doc, "系统管理_用户管理")
        except Exception as e:
            utils.g_logger.info("系统管理-用户管理页异常")
            sheet.append(["系统管理-用户管理", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_用户管理.png", doc, "系统管理_用户管理")

        # 系统管理权限管理页
        try:
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='权限管理']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//div[@class='cell el-tooltip'][text()='超级管理员']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开系统管理-权限管理页，且指定元素存在")
            utils.g_logger.info("系统管理-权限管理页显示正常")
            sheet.append(["系统管理-权限管理", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_权限管理.png", doc, "系统管理_权限管理")
        except Exception as e:
            utils.g_logger.info("系统管理-权限管理页异常")
            sheet.append(["系统管理-权限管理", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_权限管理.png", doc, "系统管理_权限管理")

        # 系统管理菜单管理页
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[@style='margin-left: 20px;'][text()='菜单管理']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//label[@for='menuid']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开系统管理-菜单管理页，且指定元素存在")
            utils.g_logger.info("系统管理-菜单管理页显示正常")
            sheet.append(["系统管理-菜单管理", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_菜单管理.png", doc, "系统管理_菜单管理")
        except Exception as e:
            utils.g_logger.info("系统管理-菜单管理页异常")
            sheet.append(["系统管理-菜单管理", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_菜单管理.png", doc, "系统管理_菜单管理")

        # 系统管理-字典表管理-系统字典表管理页
        try:
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='字典表管理']").click()
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//li[@role='menuitem'][text()='系统字典表管理 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//div[@class='el-table__fixed-right']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开系统管理-字典表管理页，且指定元素存在")
            utils.g_logger.info("系统管理-字典表管理页显示正常")
            sheet.append(["系统管理-字典表管理", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_字典表管理.png", doc, "系统管理_字典表管理")
        except Exception as e:
            utils.g_logger.info("系统管理-字典表管理页异常")
            sheet.append(["系统管理-字典表管理", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_字典表管理_系统字典表管理.png", doc,
                                  "系统管理_字典表管理_系统字典表管理")

        # 系统管理-字典表管理-专题分析类型管理页
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//li[@role='menuitem'][text()='专题分析类型管理 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//div[@class='search-container']//label[text()='分析类型名称']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开系统管理-字典表管理-专题分析类型管理页")
            utils.g_logger.info("系统管理-字典表管理-专题分析类型管理页显示正常")
            sheet.append(["系统管理-字典表管理-专题分析类型管理", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_字典表管理-专题分析类型管理.png", doc,
                                  "系统管理_字典表管理_专题分析类型管理")
        except Exception as e:
            utils.g_logger.info("系统管理-字典表管理-专题分析类型管理页异常")
            sheet.append(["系统管理-字典表管理-专题分析类型管理", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_字典表管理_专题分析类型管理.png", doc,
                                  "系统管理_字典表管理_专题分析类型管理")

        # 系统管理-字典表管理-报表分类字典表管理页
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//li[@role='menuitem'][text()='报表分类字典表管理 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//div[@class='search-container']//label[text()='报表分类名称']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开系统管理-字典表管理-报表分类字典表管理页")
            utils.g_logger.info("系统管理-字典表管理-报表分类字典表管理页显示正常")
            sheet.append(["系统管理-字典表管理-报表分类字典表管理", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_字典表管理-报表分类字典表管理.png", doc,
                                  "系统管理_字典表管理_报表分类字典表管理")
        except Exception as e:
            utils.g_logger.info("系统管理-字典表管理-报表分类字典表管理页异常")
            sheet.append(["系统管理-字典表管理-报表分类字典表管理", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_字典表管理_报表分类字典表管理.png", doc,
                                  "系统管理_字典表管理_报表分类字典表管理")

        # 系统管理-字典表管理-机构权限管理页
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//li[@role='menuitem'][text()='机构权限管理 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//div[@class='search-container']//label[text()='机构名称']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开系统管理-字典表管理-机构权限管理页")
            utils.g_logger.info("系统管理-字典表管理-机构权限管理页显示正常")
            sheet.append(["系统管理-字典表管理-机构权限管理", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_字典表管理-机构权限管理.png", doc,
                                  "系统管理_字典表管理_机构权限管理")
        except Exception as e:
            utils.g_logger.info("系统管理-字典表管理-机构权限管理页异常")
            sheet.append(["系统管理-字典表管理-机构权限管理", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_字典表管理_机构权限管理.png", doc,
                                  "系统管理_字典表管理_机构权限管理")

        # 系统管理日志管理登录日志页
        try:
            time.sleep(5)
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='日志管理']").click()
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//li[@role='menuitem'][text()='登录日志 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//label[@class='el-form-item__label'][text()='登录时间']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开系统管理-登录日志页")
            utils.g_logger.info("系统管理-登录日志页显示正常")
            sheet.append(["系统管理-登录日志", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_日志管理_登录日志.png", doc,
                                  "系统管理_日志管理_登录日志")
        except Exception as e:
            utils.g_logger.info("系统管理-登录日志页异常")
            sheet.append(["系统管理-登录日志", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_日志管理_登录日志.png", doc,
                                  "系统管理_日志管理_登录日志")

        # 系统管理日志管理操作日志页
        try:
            self.driver.find_element(By.XPATH, "//li[@role='menuitem'][text()='操作日志 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//label[@class='el-form-item__label'][text()='操作类型']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开系统管理-操作日志页，且指定元素存在")
            utils.g_logger.info("系统管理-操作日志页显示正常")
            sheet.append(["系统管理-操作日志", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_日志管理_操作日志.png", doc,
                                  "系统管理_日志管理_操作日志")
        except Exception as e:
            utils.g_logger.info("系统管理-操作日志页异常")
            sheet.append(["系统管理-操作日志", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_日志管理_操作日志.png", doc,
                                  "系统管理_日志管理_操作日志")

        # 系统管理日志管理上报日志页
        try:
            self.driver.find_element(By.XPATH, "//li[@role='menuitem'][text()='上报日志 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//label[@class='el-form-item__label'][text()='上报国家状态']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开系统管理-上报日志页，且指定元素存在")
            utils.g_logger.info("系统管理-上报日志页显示正常")
            sheet.append(["系统管理-上报日志", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_日志管理_上报日志.png", doc,
                                  "系统管理_日志管理_上报日志")
        except Exception as e:
            utils.g_logger.info("系统管理-上报日志页异常")
            sheet.append(["系统管理-上报日志", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_日志管理_上报日志.png", doc,
                                  "系统管理_日志管理_上报日志")

        # 系统管理日志管理同步日志页
        try:
            self.driver.find_element(By.XPATH, "//li[@role='menuitem'][text()='同步日志 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//label[@class='el-form-item__label'][text()='同步时间']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开系统管理-同步日志页")
            utils.g_logger.info("系统管理-同步日志页显示正常")
            sheet.append(["系统管理-同步日志", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_日志管理_同步日志.png", doc,
                                  "系统管理_日志管理_同步日志")
        except Exception as e:
            utils.g_logger.info("系统管理-同步日志页异常")
            sheet.append(["系统管理-同步日志", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_日志管理_同步日志.png", doc,
                                  "系统管理_日志管理_同步日志")

        # 系统管理报表配置页
        try:
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='报表配置']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//div[@class='search-container']//label[text()='报表周期']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开系统管理-报表配置页")

            utils.g_logger.info("系统管理-报表配置页显示正常")
            sheet.append(["系统管理-报表配置", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_报表配置.png", doc,
                                  "系统管理_报表配置")
        except Exception as e:
            utils.g_logger.info("系统管理-报表配置页异常")
            sheet.append(["系统管理-报表配置", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_报表设置.png", doc,
                                  "系统管理_报表配置")

        # 系统管理专业分析配置页
        try:
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='专业分析配置']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//label[@class='el-form-item__label'][text()='分析类型']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开系统管理-专业分析配置页")

            utils.g_logger.info("系统管理-专业分析配置页显示正常")
            sheet.append(["系统管理-专业分析配置", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_专业分析配置.png", doc,
                                  "系统管理_专业分析配置")
        except Exception as e:
            utils.g_logger.info("系统管理-专业分析配置页异常")
            sheet.append(["系统管理-专业分析配置", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_专业分析配置.png", doc,
                                  "系统管理_专业分析配置")

        # 系统管理县级用户关系绑定页
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[@style='margin-left: 20px;'][text()='县级用户关系绑定']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//th[starts-with(@class,'el-table_')]/div[text()='县级系统用户登录名']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开系统管理-县级用户关系绑定页")
            utils.g_logger.info("系统管理-县级用户关系绑定页显示正常")
            sheet.append(["系统管理-县级用户关系绑定", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_县级用户关系绑定.png", doc,
                                  "系统管理_县级用户关系绑定")
        except Exception as e:
            utils.g_logger.info("系统管理-县级用户关系绑定页异常")
            sheet.append(["系统管理-县级用户关系绑定", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_县级用户关系绑定.png", doc,
                                  "系统管理_县级用户关系绑定")

        # 系统管理微信用户管理页
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[@style='margin-left: 20px;'][text()='微信用户管理']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//thead[@class='has-gutter']//div[text()='微信OPEN-ID']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开系统管理-微信用户管理页，且指定元素存在")
            utils.g_logger.info("系统管理-微信用户管理页显示正常")
            sheet.append(["系统管理-微信用户管理", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_微信用户管理.png", doc,
                                  "系统管理_微信用户管理")
        except Exception as e:
            utils.g_logger.info("系统管理-微信用户管理页异常")
            sheet.append(["系统管理-微信用户管理", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/系统管理_微信用户管理.png", doc,
                                  "系统管理_微信用户管理")

        # 关闭所有页面
        self.driver.find_element(By.XPATH, "//div[@class='close-con']").click()
        try:
            element = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//ul[@class='el-dropdown-menu el-popper']/li[text()='关闭所有']"))
            )
            element.click()
            utils.g_logger.info("成功关闭所有页面")
        except Exception as e:
            utils.g_logger.info("关闭所有页面失败")

    @utils.retry(MAX_TRIES)
    def test_wulianwang(self):
        # 物联网工作平台页
        try:
            time.sleep(5)
            self.driver.find_element(By.XPATH,
                                     "//div[@class='navigation-list']/div[@class='navigation-item']/span[text()=' 物联网']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(
                (By.XPATH, "//div[@class='lenged']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开物联网-工作平台页")
            utils.g_logger.info("物联网-工作平台页显示正常")
            sheet.append(["物联网-工作平台", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_工作平台.png", doc, "物联网_工作平台")
        except Exception as e:
            utils.g_logger.info("物联网-工作平台页异常")
            sheet.append(["物联网-工作平台", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_工作平台.png", doc, "物联网_工作平台")

        # 物联网-监测点分布页面
        try:
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='监测点分布']").click()
            time.sleep(10)
            tree_element = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, "//button[@title='Zoom in']")))
            unittest.TestCase.assertTrue(tree_element is not None, "成功打开监测点分布页面")
            utils.g_logger.info("物联网-监测点分布显示正常")
            sheet.append(["物联网-监测点分布", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_监测点分布.png", doc, "物联网_监测点分布")
        except Exception as e:
            utils.g_logger.info("物联网-监测点分布页异常")
            sheet.append(["物联网-监测点分布", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_监测点分布.png", doc, "物联网_监测点分布")

        # 物联网-设备分布页面
        try:
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='设备分布']").click()
            time.sleep(4)
            self.driver.find_element(By.XPATH, "//*[@id='tab-table']").click()
            time.sleep(10)
            tree_element = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, "//div[@class='basetable-wrapper']")))
            unittest.TestCase.assertTrue(tree_element is not None, "成功打开设备分布页面")
            utils.g_logger.info("物联网-设备分布显示正常")
            sheet.append(["物联网-设备分布", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_设备分布.png", doc, "物联网_设备分布")
        except Exception as e:
            utils.g_logger.info("物联网-设备分布页异常")
            sheet.append(["物联网-设备分布", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_设备分布.png", doc, "物联网_设备分布")

        # 虫量对比分析页
        try:
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='虫量对比分析']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(
                (By.XPATH, "//i[@class='iot-el-icon iconfont icon-zhuzhuang']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开物联网-虫量对比分析页，且虫害类型选项存在")
            utils.g_logger.info("物联网-虫量对比分析页显示正常")
            sheet.append(["物联网-虫量对比分析", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_虫量对比分析.png", doc, "物联网_虫量对比分析")
        except Exception as e:
            utils.g_logger.info("物联网-虫量对比分析页异常")
            sheet.append(["物联网-虫量对比分析", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_虫量对比分析.png", doc, "物联网_虫量对比分析")

        # 物联网-环境气象-趋势分析
        try:
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='环境气象']").click()
            time.sleep(2)
            self.driver.find_element(By.XPATH,
                                     "//li[@class='el-submenu is-opened']//li[text()='趋势分析 ']").click()
            time.sleep(10)
            tree_element = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(
                (By.XPATH,
                 "//div[@class='iot-el-form-item asterisk-left iot-el-form-item--label-right']//div[@class='iot-el-radio-group radio-group-container']")))
            unittest.TestCase.assertTrue(tree_element is not None, "成功打开环境气象-趋势分析页面")
            utils.g_logger.info("环境气象-趋势分析显示正常")
            sheet.append(["环境气象-趋势分析", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_环境气象_趋势分析.png", doc,
                                  "物联网_环境气象_趋势分析")
        except Exception as e:
            utils.g_logger.info("环境气象-趋势分析页异常")
            sheet.append(["环境气象-趋势分析", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_环境气象_趋势分析.png", doc,
                                  "物联网_环境气象_趋势分析")

        # 物联网-环境气象-实时数据列表
        try:
            self.driver.find_element(By.XPATH, "//li[text()='实时数据列表 ']").click()
            time.sleep(10)
            tree_element = WebDriverWait(self.driver, 30).until(EC.presence_of_element_located(
                (By.XPATH, "//table[@class='iot-el-table__body']")))
            unittest.TestCase.assertTrue(tree_element is not None, "成功打开环境气象-实时数据列表页面，且指定元素存在")
            utils.g_logger.info("环境气象-实时数据列表显示正常")
            sheet.append(["环境气象-实时数据列表", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_环境气象_实时数据列表.png", doc,
                                  "物联网_环境气象_实时数据列表")
        except Exception as e:
            utils.g_logger.info("环境气象-实时数据列表显示正常")
            sheet.append(["环境气象-实时数据列表", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_环境气象_实时数据列表.png", doc,
                                  "物联网_环境气象_实时数据列表")

        # 物联网-环境气象-实时数据统计
        time.s.sleep(5)
        self.driver.find_element(By.XPATH, "//li[text()='实时数据统计 ']").click()
        try:
            time.sleep(15)
            tree_element = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, "//button[@class='iot-el-button iot-el-button--primary']/span[text()=' 可选指标 ']")))
            unittest.TestCase.assertTrue(tree_element is not None, "成功打开环境气象-实时数据统计页面，且指定元素存在")
            utils.g_logger.info("环境气象-实时数据统计显示正常")
            sheet.append(["环境气象-实时数据统计", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_环境气象_实时数据统计.png", doc,
                                  "物联网_环境气象_实时数据统计")
        except Exception as e:
            utils.g_logger.info("环境气象-实时数据统计页异常")
            sheet.append(["环境气象-实时数据统计", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_环境气象_实时数据统计.png", doc,
                                  "物联网_环境气象_实时数据统计")

        # 物联网-环境气象-逐日数据统计
        try:
            self.driver.find_element(By.XPATH, "//li[text()='逐日数据统计 ']").click()
            time.sleep(10)
            tree_element = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, "//button[@class='iot-el-button iot-el-button--primary']/span[text()=' 可选指标 ']")))
            unittest.TestCase.assertTrue(tree_element is not None, "成功打开环境气象-逐日数据统计页面")
            utils.g_logger.info("环境气象-逐日数据统计显示正常")
            sheet.append(["环境气象-逐日数据统计", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_环境气象_逐日数据统计.png", doc,
                                  "物联网_环境气象_逐日数据统计")
        except Exception as e:
            utils.g_logger.info("环境气象-逐日数据统计页异常")
            sheet.append(["环境气象-逐日数据统计", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_环境气象_逐日数据统计.png", doc,
                                  "物联网_环境气象_逐日数据统计")

        # 物联网-环境气象-逐日数据列表
        try:
            self.driver.find_element(By.XPATH, "//li[text()='逐日数据列表 ']").click()
            time.sleep(10)
            wendu_element = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(
                (By.XPATH, "//div[@class='cell'][text()='日平均温度(℃)']")))
            unittest.TestCase.assertTrue(wendu_element is not None,
                                         "成功打开环境气象-逐日数据列表页，且站点列表存在")
            utils.g_logger.info("环境气象-逐日数据列表显示正常")
            sheet.append(["环境气象-逐日数据列表", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_环境气象_逐日数据列表.png", doc,
                                  "物联网_环境气象_逐日数据列表")
        except Exception as e:
            utils.g_logger.info("环境气象-逐日数据列表页异常")
            sheet.append(["环境气象-逐日数据列表", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_环境气象_逐日数据列表.png", doc,
                                  "物联网_环境气象_逐日数据列表")

        # 性诱监测-逐日数据统计
        try:
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='性诱监测']").click()
            time.sleep(2)
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='逐日数据统计 ']").click()
            time.sleep(10)
            tree_element = WebDriverWait(self.driver, 15).until(EC.presence_of_element_located(
                (By.XPATH, "//button[@class='iot-el-button iot-el-button--primary']/*[text()='可选指标']")))
            unittest.TestCase.assertTrue(tree_element is not None, "成功打开性诱监测-逐日数据统计页面，且指定元素存在")
            utils.g_logger.info("性诱监测-逐日数据统计显示正常")
            sheet.append(["性诱监测-逐日数据统计", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_性诱监测_逐日数据统计.png", doc,
                                  "物联网_性诱监测_逐日数据统计")
        except Exception as e:
            utils.g_logger.info("性诱监测-逐日数据统计页异常")
            sheet.append(["性诱监测-逐日数据统计", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_性诱监测_逐日数据统计.png", doc,
                                  "物联网_性诱监测_逐日数据统计")

        # 性诱监测-数据统计列表
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[@class='el-submenu is-active is-opened']//li[text()='数据统计列表 ']").click()
            time.sleep(10)
            tree_element = WebDriverWait(self.driver, 15).until(EC.presence_of_element_located(
                (By.XPATH, "//div[@class='cell'][text()='日累计虫量']")))
            unittest.TestCase.assertTrue(tree_element is not None, "成功打开性诱监测-数据统计列表页面，且指定元素存在")
            utils.g_logger.info("性诱监测-数据统计列表显示正常")
            sheet.append(["性诱监测-数据统计列表", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_性诱监测_数据统计列表.png", doc,
                                  "物联网_性诱监测_数据统计列表")
        except Exception as e:
            utils.g_logger.info("性诱监测-数据统计列表页异常")
            sheet.append(["性诱监测-数据统计列表", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_性诱监测_数据统计列表.png", doc,
                                  "物联网_性诱监测_数据统计列表")

        # 性诱监测-识别结果统计
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[@class='el-submenu is-active is-opened']//li[text()='识别结果统计 ']").click()
            time.sleep(10)
            orgnization_tree_element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//div[@class='iot-el-table__body-wrapper']")))
            unittest.TestCase.assertTrue(orgnization_tree_element is not None,
                                         "成功打开性诱监测-识别结果统计页")
            utils.g_logger.info("性诱监测-识别结果统计显示正常")
            sheet.append(["性诱监测-识别结果统计", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_性诱监测_识别结果统计.png", doc,
                                  "物联网_性诱监测_识别结果统计")
        except Exception as e:
            utils.g_logger.info("性诱监测-识别结果统计页异常")
            sheet.append(["性诱监测-识别结果统计", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_性诱监测_识别结果统计.png", doc,
                                  "物联网_性诱监测_识别结果统计")

        # 性诱监测-趋势分析
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[@class='el-submenu is-active is-opened']//li[text()='性诱趋势分析 ']").click()
            time.sleep(10)
            chart_element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, "//i[@class='iot-el-icon iconfont icon-zhuzhuang']")))
            unittest.TestCase.assertTrue(chart_element is not None,
                                         "成功性诱监测-趋势分析页")
            utils.g_logger.info("性诱监测-趋势分析显示正常")
            sheet.append(["性诱监测-趋势分析", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_性诱监测_趋势分析.png", doc,
                                  "物联网_性诱监测_趋势分析")
        except Exception as e:
            utils.g_logger.info("性诱监测-趋势分析页异常")
            sheet.append(["性诱监测-趋势分析", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_性诱监测_趋势分析.png", doc,
                                  "物联网_性诱监测_趋势分析")

        # 灯诱监测-灯诱数据分析页
        try:
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//span[text()='灯诱监测']").click()
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//li[@role='menuitem'][text()='灯诱数据分析 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(
                (By.XPATH, "//button[@class='iot-el-button iot-el-button--primary']/*[text()='可选指标']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开物联网-灯诱监测-灯诱数据分析页")
            utils.g_logger.info("物联网-灯诱监测-灯诱数据分析页显示正常")
            sheet.append(["物联网-灯诱监测-灯诱数据分析", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_灯诱监测_灯诱数据分析.png", doc,
                                  "物联网_灯诱监测_灯诱数据分析")
        except Exception as e:
            utils.g_logger.info("物联网-灯诱监测-灯诱数据分析页异常")
            sheet.append(["物联网-灯诱监测-灯诱数据分析", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_灯诱监测_灯诱数据分析.png", doc,
                                  "物联网_灯诱监测_灯诱数据分析")

        # 灯诱监测-数据统计列表
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[@slot='title' and text()='灯诱监测']/parent::div/parent::li//li[text()='数据统计列表 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(
                (By.XPATH, "//thead//div[text()='日累计虫量']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开物联网-灯诱监测-数据统计列表页")
            utils.g_logger.info("物联网-灯诱监测-数据统计列表页显示正常")
            sheet.append(["物联网-灯诱监测-数据统计列表", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_灯诱监测_数据统计列表.png", doc,
                                  "物联网_灯诱监测_数据统计列表")
        except Exception as e:
            utils.g_logger.info("物联网-灯诱监测-数据统计列表页异常")
            sheet.append(["物联网-灯诱监测-数据统计列表", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_灯诱监测_数据统计列表.png", doc,
                                  "物联网_灯诱监测_数据统计列表")

        # 灯诱监测-灯诱图片展示
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[@role='menuitem'][text()='灯诱图片展示 ']").click()
            time.sleep(10)
            calendar_element = WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located(
                (By.XPATH, "//span[text()='仅看识别后图片']")))
            unittest.TestCase.assertTrue(calendar_element is not None,
                                         "成功打开物联网-灯诱监测-灯诱图片展示页")
            utils.g_logger.info("物联网-灯诱监测-灯诱图片展示页显示正常")
            sheet.append(["物联网-灯诱监测-灯诱图片展示", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_灯诱监测_灯诱图片展示.png", doc,
                                  "物联网_灯诱监测_灯诱图片展示")
        except Exception as e:
            utils.g_logger.info("物联网-灯诱监测-灯诱图片展示页异常")
            sheet.append(["物联网-灯诱监测-灯诱图片展示", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_灯诱监测_灯诱图片展示.png", doc,
                                  "物联网_灯诱监测_灯诱图片展示")

        # 灯诱监测-灯诱识别结果统计
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[@role='menuitem'][text()='灯诱识别结果统计 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(
                (By.XPATH, "//label[text()='虫害类型']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开物联网-灯诱监测-灯诱识别结果统计页，且虫害类型选项存在")
            utils.g_logger.info("物联网-灯诱监测-灯诱识别结果统计页显示正常")
            sheet.append(["物联网-灯诱监测-灯诱识别结果统计", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_灯诱监测_灯诱识别结果统计.png", doc,
                                  "物联网_灯诱监测_灯诱识别结果统计")
        except Exception as e:
            utils.g_logger.info("物联网-灯诱监测-灯诱识别结果统计页异常")
            sheet.append(["物联网-灯诱监测-灯诱识别结果统计", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_灯诱监测_灯诱识别结果统计.png", doc,
                                  "物联网_灯诱监测_灯诱识别结果统计")

        # 灯诱监测-趋势分析
        try:
            time.sleep(5)
            self.driver.find_element(By.XPATH, "//li[@role='menuitem'][text()='灯诱趋势分析 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(
                (By.XPATH, "//label[text()='统计类型']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开物联网-灯诱监测-趋势分析页")
            utils.g_logger.info("物联网-灯诱监测-趋势分析页显示正常")
            sheet.append(["物联网-灯诱监测-趋势分析", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_灯诱监测_趋势分析.png", doc,
                                  "物联网_灯诱监测_趋势分析")
        except Exception as e:
            utils.g_logger.info("物联网-灯诱监测-趋势分析页异常")
            sheet.append(["物联网-灯诱监测-趋势分析", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_灯诱监测_趋势分析.png", doc,
                                  "物联网_灯诱监测_趋势分析")

        # 病害监测-孢子监测页
        try:
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//span[@slot='title' and text()='病害监测']").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//li[@role='menuitem'][text()='孢子监测 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//span[text()='仅看有孢子']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开物联网-病害监测-孢子监测页")
            utils.g_logger.info("物联网-病害监测-孢子监测页显示正常")
            sheet.append(["物联网-病害监测-孢子监测", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_病害监测_孢子监测.png", doc,
                                  "物联网_病害监测_孢子监测")
        except Exception as e:
            utils.g_logger.info("物联网-病害监测-孢子监测页异常")
            sheet.append(["物联网-病害监测-孢子监测", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_病害监测_孢子监测.png", doc,
                                  "物联网_病害监测_孢子监测")

        # 病害监测-病害GIS分析页
        try:
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//li[@role='menuitem'][text()='病害GIS分析 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//label[text()='图形类型']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开物联网-病害监测-病害GIS分析页")
            utils.g_logger.info("物联网-病害监测-病害GIS分析页显示正常")
            sheet.append(["物联网-病害监测-病害GIS分析", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_病害监测_病害GIS分析.png", doc,
                                  "物联网_病害监测_病害GIS分析")
        except Exception as e:
            utils.g_logger.info("物联网-病害监测-病害GIS分析页异常")
            sheet.append(["物联网-病害监测-病害GIS分析", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_病害监测_病害GIS分析.png", doc,
                                  "物联网_病害监测_病害GIS分析")

        # 关闭所有页面
        try:
            self.driver.find_element(By.XPATH, "//div[@class='close-con']").click()
            element = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//ul[@class='el-dropdown-menu el-popper']/li[text()='关闭所有']"))
            )
            element.click()
            utils.g_logger.info("成功关闭所有页面")
        except Exception as e:
            utils.g_logger.info("关闭所有页面失败")

        # 物联网-物联网管理-设备管理页
        try:
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//span[@slot='title' and text()='物联网管理']").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//li[@role='menuitem'][text()='设备管理 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(
                (By.XPATH, "//label[text()='设备名称']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开物联网-物联网管理-设备管理页")
            utils.g_logger.info("物联网-物联网管理-设备管理页显示正常")
            sheet.append(["物联网-物联网管理-设备管理", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_物联网管理_设备管理.png", doc,
                                  "物联网_物联网管理_设备管理")
        except Exception as e:
            utils.g_logger.info("物联网-物联网管理-设备管理页异常")
            sheet.append(["物联网-物联网管理-设备管理", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_物联网管理_设备管理.png", doc,
                                  "物联网_物联网管理_设备管理")

        # 物联网-物联网管理-监测点管理页
        try:
            self.driver.find_element(By.XPATH, "//li[@role='menuitem'][text()='监测点管理 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(
                (By.XPATH, "//label[text()='监测点名称']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开物联网-物联网管理-监测点管理页")
            utils.g_logger.info("物联网-物联网管理-监测点管理页显示正常")
            sheet.append(["物联网-物联网管理-监测点管理", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_物联网管理_监测点管理.png", doc,
                                  "物联网_物联网管理_监测点管理")
        except Exception as e:
            utils.g_logger.info("物联网-物联网管理-监测点管理页异常")
            sheet.append(["物联网-物联网管理-监测点管理", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_物联网管理_监测点管理.png", doc,
                                  "物联网_物联网管理_监测点管理")

        # 物联网-视频监控-视频监控分布
        try:
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//span[@slot='title' and text()='视频监控']").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//li[@role='menuitem'][text()='视频监控分布 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(
                (By.XPATH, "//label[text()='站点']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开物联网-视频监控分布页，且站点列表存在")
            utils.g_logger.info("物联网-视频监控-视频监控分布页显示正常")
            sheet.append(["物联网-视频监控-视频监控分布", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_视频监控_视频监控分布.png", doc,
                                  "物联网_视频监控_视频监控分布")
        except Exception as e:
            utils.g_logger.info("物联网-视频监控-视频监控分布页异常")
            sheet.append(["物联网-视频监控-视频监控分布", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_视频监控_视频监控分布.png", doc,
                                  "物联网_视频监控_视频监控分布")

        # 物联网-视频监控-视频图片展示
        try:
            self.driver.find_element(By.XPATH, "//li[@role='menuitem'][text()='视频图片展示 ']").click()
            time.sleep(10)
            calendar_element = WebDriverWait(self.driver, 5).until(EC.presence_of_all_elements_located(
                (By.XPATH, "//div[@class='imgCard']/p[text()=' 拍摄 ']")))
            unittest.TestCase.assertTrue(calendar_element is not None, "成功打开物联网-视频监控-视频图片展示页")
            utils.g_logger.info("物联网-视频监控-视频图片展示页显示正常")
            sheet.append(["物联网-视频监控-视频图片展示", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_视频监控_视频图片展示.png", doc,
                                  "物联网_视频监控_视频图片展示")
        except Exception as e:
            utils.g_logger.info("物联网-视频监控-视频图片展示页异常")
            sheet.append(["物联网-视频监控-视频图片展示", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_视频监控_视频图片展示.png", doc,
                                  "物联网_视频监控_视频图片展示")

        # 物联网评价-设备数据管理
        try:
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//span[@slot='title' and text()='物联网评价']").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//li[@role='menuitem'][text()='设备数据管理 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(
                (By.XPATH, "//div[text()='设备运行情况']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开物联网-物联网评价-设备数据管理页")
            utils.g_logger.info("物联网-物联网评价-设备数据管理页显示正常")
            sheet.append(["物联网-物联网评价-设备数据管理", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_物联网评价-设备数据管理.png", doc,
                                  "物联网_物联网评价-设备数据管理")
        except Exception as e:
            utils.g_logger.info("物联网-物联网评价-设备数据管理")
            sheet.append(["物联网-物联网评价-设备数据管理", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_物联网评价-设备数据管理.png", doc,
                                  "物联网_物联网评价-设备数据管理")

        # 物联网评价-厂商考核评价-厂商维护情况
        try:
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//li[@role='menuitem']//span[text()='厂商考核评价']").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//li[@role='menuitem'][text()='厂商维护情况 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(
                (By.XPATH, "//div[@class='cell'][text()='维护成功次数']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开物联网-物联网评价-厂商考核评价-厂商维护情况页")
            utils.g_logger.info("物联网-物联网评价-厂商考核评价-厂商维护情况页显示正常")
            sheet.append(["物联网-物联网评价-厂商考核评价-厂商维护情况", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_物联网评价-厂商考核评价-厂商维护情况.png",
                                  doc,
                                  "物联网_物联网评价-厂商考核评价-厂商维护情况")
        except Exception as e:
            utils.g_logger.info("物联网-物联网评价-厂商考核评价-厂商维护情况")
            sheet.append(["物联网-物联网评价-厂商考核评价-厂商维护情况", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_物联网评价-厂商考核评价-厂商维护情况.png",
                                  doc,
                                  "物联网_物联网评价-厂商考核评价-厂商维护情况")

        # 物联网评价-厂商考核评价-维护失败设备清单
        try:
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//li[@role='menuitem'][text()='维护失败设备清单 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(
                (By.XPATH, "//div[@class='cell'][text()='维护失败设备数量']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开物联网-物联网评价-厂商考核评价-维护失败设备清单页")
            utils.g_logger.info("物联网-物联网评价-厂商考核评价-维护失败设备清单页显示正常")
            sheet.append(["物联网-物联网评价-厂商考核评价-维护失败设备清单", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_物联网评价-厂商考核评价-维护失败设备清单.png",
                                  doc,
                                  "物联网_物联网评价-厂商考核评价-维护失败设备清单")
        except Exception as e:
            utils.g_logger.info("物联网-物联网评价-厂商考核评价-维护失败设备清单")
            sheet.append(["物联网-物联网评价-厂商考核评价-维护失败设备清单", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_物联网评价-厂商考核评价-维护失败设备清单.png",
                                  doc,
                                  "物联网_物联网评价-厂商考核评价-维护失败设备清单")

        # 物联网-高空灯监测-高空灯数据列表
        try:
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//li[@role='menuitem']//span[text()='高空灯监测']").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//li[@role='menuitem'][text()='高空灯数据列表 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(
                (By.XPATH, "//label[text()='虫量']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开物联网-高空灯监测-高空灯数据列表页")
            utils.g_logger.info("物联网-高空灯监测-高空灯数据列表页显示正常")
            sheet.append(["物联网-高空灯监测-高空灯数据列表", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_高空灯监测_高空灯数据列表.png", doc,
                                  "物联网_高空灯监测_高空灯数据列表")
        except Exception as e:
            utils.g_logger.info("物联网-高空灯监测-高空灯数据列表")
            sheet.append(["物联网-高空灯监测-高空灯数据列表", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_高空灯监测_高空灯数据列表.png", doc,
                                  "物联网_高空灯监测_高空灯数据列表")

        # 物联网-高空灯监测-高空灯图片
        try:
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//li[@role='menuitem'][text()='高空灯图片 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.presence_of_all_elements_located(
                (By.XPATH, "//div[@class='imgCard']/p[text()=' 拍摄 ']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开物联网-高空灯监测-高空灯图片页")
            utils.g_logger.info("物联网-高空灯监测-高空灯图片页显示正常")
            sheet.append(["物联网-高空灯监测-高空灯图片", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_高空灯监测_高空灯图片.png", doc,
                                  "物联网_高空灯监测_高空灯图片")
        except Exception as e:
            utils.g_logger.info("物联网-高空灯监测-高空灯图片")
            sheet.append(["物联网-高空灯监测-高空灯图片", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/物联网_高空灯监测_高空灯图片.png", doc,
                                  "物联网_高空灯监测_高空灯图片")
        # 关闭所有页面
        self.driver.find_element(By.XPATH, "//div[@class='close-con']").click()
        try:
            element = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//ul[@class='el-dropdown-menu el-popper']/li[text()='关闭所有']"))
            )
            element.click()
            utils.g_logger.info("成功关闭所有页面")
        except Exception as e:
            utils.g_logger.info("关闭所有页面执行失败")

    @utils.retry(MAX_TRIES)
    def test_zhibaotongji(self):
        # 植保统计-任务管理
        try:
            time.sleep(5)
            self.driver.find_element(By.XPATH,
                                     "//div[@class='navigation-list']/div[@class='navigation-item']/span[text()=' 植保统计']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located(
                (By.XPATH, "//thead[@class='has-gutter']//div[@class='cell'][text()='任务数']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植保统计-任务管理页")
            utils.g_logger.info("植保统计-任务管理页显示正常")
            sheet.append(["植保统计-任务管理", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保统计_任务管理.png", doc, "植保统计_任务管理")
        except Exception as e:
            utils.g_logger.info("植保统计-任务管理页异常")
            sheet.append(["植保统计-任务管理", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保统计_任务管理.png", doc, "植保统计_任务管理")

        # 植保统计-查询统计-任务查询
        try:
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='查询统计']").click()
            time.sleep(2)
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='任务查询 ']").click()
            time.sleep(10)
            tree_element = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, "//div[@class='cell'][text()='任务开始时间']")))
            unittest.TestCase.assertTrue(tree_element is not None, "成功打开植保统计-查询统计-任务查询页面")
            utils.g_logger.info("植保统计-查询统计-任务查询显示正常")
            sheet.append(["植保统计-查询统计-任务查询", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保统计_查询统计_任务查询.png", doc,
                                  "植保统计_查询统计_任务查询")
        except Exception as e:
            utils.g_logger.info("植保统计-查询统计-任务查询")
            sheet.append(["植保统计-查询统计-任务查询", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保统计_查询统计_任务查询.png", doc,
                                  "植保统计_查询统计_任务查询")

        # 植保统计-查询统计-原表查询
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='原表查询 ']").click()
            time.sleep(10)
            tree_element = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, "//label[text()='排序方式:']")))
            unittest.TestCase.assertTrue(tree_element is not None, "成功打开植保统计-查询统计-原表查询页面")
            utils.g_logger.info("植保统计-查询统计-原表查询显示正常")
            sheet.append(["植保统计-查询统计-原表查询", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保统计_查询统计_原表查询.png", doc,
                                  "植保统计_查询统计_原表查询")
        except Exception as e:
            utils.g_logger.info("植保统计-查询统计-原表查询")
            sheet.append(["植保统计-查询统计-原表查询", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保统计_查询统计_原表查询.png", doc,
                                  "植保统计_查询统计_原表查询")

        # 植保统计-查询统计-杂食害虫查询
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='杂食害虫查询 ']").click()
            time.sleep(10)
            tree_element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//label[text()='选择害虫']")))
            unittest.TestCase.assertTrue(tree_element is not None, "成功打开植保统计-查询统计-杂食害虫查询页面")
            utils.g_logger.info("植保统计-查询统计-杂食害虫查询显示正常")
            sheet.append(["植保统计-查询统计-杂食害虫查询", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保统计_查询统计_杂食害虫查询.png", doc,
                                  "植保统计_查询统计_杂食害虫查询")
        except Exception as e:
            utils.g_logger.info("植保统计-查询统计-杂食害虫查询")
            sheet.append(["植保统计-查询统计-杂食害虫查询", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保统计_查询统计_杂食害虫查询.png", doc,
                                  "植保统计_查询统计_杂食害虫查询")

        # 关闭所有页面
        self.driver.find_element(By.XPATH, "//div[@class='close-con']").click()
        try:
            element = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, "//ul[@class='el-dropdown-menu el-popper']/li[text()='关闭所有']")))
            element.click()
            utils.g_logger.info("成功关闭所有页面")
        except Exception as e:
            utils.g_logger.info("关闭所有页面执行失败")

        # 数据分析-综合分析页
        time.sleep(3)
        self.driver.find_element(By.XPATH, "//div[starts-with(@class,'navi-item')][text()='数据分析 ']").click()
        try:
            time.sleep(15)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@role='radiogroup']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开数据分析-综合分析页，且指定元素存在")
            select_date_element = self.driver.find_element(By.XPATH, '//input[@placeholder="请选择日期"]')
            select_date_element.clear()
            select_date_element.send_keys('2024-10-08')
            time.sleep(3)
            # 模拟按下键盘回车键
            select_date_element.send_keys(Keys.RETURN)

            time.sleep(15)
            chart_element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//div[text()=' 当前发生面积与常年对比 ']/parent::div//canvas")))
            unittest.TestCase.assertTrue(chart_element is not None, "成功打开数据分析-综合分析页，且指定元素存在")
            utils.g_logger.info("数据分析-综合分析页显示正常")
            sheet.append(["数据分析-综合分析", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据分析_综合分析.png", doc, "数据分析_综合分析")
        except Exception as e:
            utils.g_logger.info("数据分析-综合分析页异常")
            sheet.append(["数据分析-综合分析", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据分析_综合分析.png", doc, "数据分析_综合分析")

        # 数据分析-专题分析页
        time.sleep(3)
        self.driver.find_element(By.XPATH, "//span[text()='专题分析']").click()
        try:
            time.sleep(15)
            elements = WebDriverWait(self.driver, 10).until(EC.presence_of_all_elements_located(
                (By.XPATH, "//div[@class='el-image']")))
            unittest.TestCase.assertTrue(elements is not None,
                                         "成功打开数据分析-专题分析页，且指定元素存在")
            self.driver.find_element(By.XPATH, "//div[@class='row-title'][text()=' 稻纵卷叶螟 ']").click()
            time.sleep(5)
            indicator_element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//div[@class='section_left section_left_light']/ul/li[1]")))
            unittest.TestCase.assertTrue(indicator_element is not None,
                                         "成功打开数据分析-专题分析页，且指定元素存在")
            utils.g_logger.info("数据分析-专题分析页显示正常")
            sheet.append(["数据分析-专题分析", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据分析_专题分析.png", doc, "数据分析_专题分析")
            self.driver.find_element(By.XPATH, "//div[@class='el-form-item__content']//span[text()='关闭']").click()
        except Exception as e:
            utils.g_logger.info("数据分析-专题分析页异常")
            sheet.append(["数据分析-专题分析", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据分析_专题分析.png", doc, "数据分析_专题分析")

        # 数据分析-GIS分析页
        time.sleep(3)
        self.driver.find_element(By.XPATH, "//span[text()='GIS分析']").click()
        try:
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(
                (By.XPATH, "//label[text()='分析指标']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开数据分析-GIS分析页，且指定元素存在")
            map_element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//div[@class='map']")))
            unittest.TestCase.assertTrue(map_element is not None,
                                         "成功打开数据分析-GIS分析页，且指定元素存在")
            utils.g_logger.info("数据分析-GIS分析页显示正常")
            sheet.append(["数据分析-GIS分析", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据分析_GIS分析.png", doc, "数据分析_GIS分析")
        except Exception as e:
            utils.g_logger.info("数据分析-GIS分析页异常")
            sheet.append(["数据分析-GIS分析", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据分析_GIS分析.png", doc, "数据分析_GIS分析")

        # 数据分析-自定义分析页
        time.sleep(3)
        self.driver.find_element(By.XPATH, "//span[text()='自定义分析']").click()
        try:
            time.sleep(8)
            self.driver.find_element(By.XPATH, "//div[@class='report-selector']").click()
            elements = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//img[contains(@src,'/upload/specialtype/')]")))
            unittest.TestCase.assertTrue(elements is not None, "成功打开数据分析-自定义分析页，且指定元素存在")
            utils.g_logger.info("数据分析-自定义分析页显示正常")
            sheet.append(["数据分析-自定义分析", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据分析_自定义分析.png", doc, "数据分析_自定义分析")
        except Exception as e:
            utils.g_logger.info("数据分析-自定义分析页异常")
            sheet.append(["数据分析-自定义分析", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据分析_自定义分析.png", doc, "数据分析_自定义分析")

        # 数据分析-数据报告页
        time.sleep(3)
        self.driver.find_element(By.XPATH, "//span[text()='数据报告']").click()
        try:
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(
                (By.XPATH, "//label[text()='周期']/following-sibling::div/div[@role='radiogroup']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开数据分析-数据报告页，且指定元素存在")
            utils.g_logger.info("数据分析-数据报告页显示正常")
            sheet.append(["数据分析-数据报告", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据分析_数据报告.png", doc, "数据分析_数据报告")
        except Exception as e:
            utils.g_logger.info("数据分析-数据报告页异常")
            sheet.append(["数据分析-数据报告", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数据分析_数据报告.png", doc, "数据分析_数据报告")

        # 关闭所有页面
        self.driver.find_element(By.XPATH, "//div[@class='close-con']").click()
        try:
            element = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//ul[@class='el-dropdown-menu el-popper']/li[text()='关闭所有']"))
            )
            element.click()
            utils.g_logger.info("成功关闭所有页面")
        except Exception as e:
            utils.g_logger.info(e)

        # 知识库-工作平台页
        self.driver.find_element(By.XPATH, "//div[starts-with(@class,'navi-item')][text()='知识库 ']").click()
        try:
            time.sleep(12)
            elements = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, "//div[@class='title'][text()='病虫害知识库']/parent::div/following-sibling::div//img")))
            unittest.TestCase.assertTrue(elements is not None, "成功打开知识库-工作平台页，且指定元素存在")
            utils.g_logger.info("知识库-工作平台页显示正常")
            sheet.append(["知识库-工作平台", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_工作平台.png", doc, "知识库_工作平台")
        except Exception as e:
            utils.g_logger.info("知识库-工作平台页异常")
            sheet.append(["知识库-工作平台", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_工作平台.png", doc, "知识库_工作平台")

        # 知识库-病虫害知识库-知识浏览页
        time.sleep(3)
        self.driver.find_element(By.XPATH, "//span[text()='病虫害知识库']").click()
        self.driver.find_element(By.XPATH,
                                 "//span[text()='病虫害知识库']/parent::div/following-sibling::ul//li[text()='知识浏览 ']").click()
        try:
            time.sleep(12)
            elements = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//div[@class='el-image img']/img")))
            unittest.TestCase.assertTrue(elements is not None, "成功打开知识库-病虫害知识库-知识浏览页，且指定元素存在")
            utils.g_logger.info("知识库-病虫害知识库-知识浏览页显示正常")
            sheet.append(["知识库-病虫害知识库-知识浏览", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_病虫害知识库_知识浏览.png", doc,
                                  "知识库_病虫害知识库_知识浏览")
        except Exception as e:
            utils.g_logger.info("知识库-病虫害知识库-知识浏览页异常")
            sheet.append(["知识库-病虫害知识库-知识浏览", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_病虫害知识库_知识浏览.png", doc,
                                  "知识库_病虫害知识库_知识浏览")

        # 知识库-病虫害知识库-知识维护页
        time.sleep(3)
        self.driver.find_element(By.XPATH,
                                 "//span[text()='病虫害知识库']/parent::div/following-sibling::ul//li[text()='知识维护 ']").click()
        try:
            time.sleep(9)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@class='cell'][text()='修改时间']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开知识库-病虫害知识库-知识维护页，且指定元素存在")
            utils.g_logger.info("知识库-病虫害知识库-知识维护页显示正常")
            sheet.append(["知识库-病虫害知识库-知识维护", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_病虫害知识库_知识维护.png", doc,
                                  "知识库_病虫害知识库_知识维护")
        except Exception as e:
            utils.g_logger.info("知识库-病虫害知识库-知识维护页异常")
            sheet.append(["知识库-病虫害知识库-知识维护", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_病虫害知识库_知识维护.png", doc,
                                  "知识库_病虫害知识库_知识维护")

        # 植保知识库-知识浏览页
        time.sleep(3)
        self.driver.find_element(By.XPATH, "//span[text()='植保知识库']").click()
        self.driver.find_element(By.XPATH,
                                 "//span[text()='植保知识库']/parent::div/following-sibling::ul//li[text()='知识浏览 ']").click()
        try:
            time.sleep(8)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@role='radiogroup']//span[text()='植物检疫']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开知识库-植保知识库-知识浏览页，且指定元素存在")
            utils.g_logger.info("知识库-植保知识库-知识浏览页显示正常")
            sheet.append(["知识库-植保知识库-知识浏览", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_植保知识库_知识浏览.png", doc,
                                  "知识库_植保知识库_知识浏览")
        except Exception as e:
            utils.g_logger.info("知识库-植保知识库-知识浏览页异常")
            sheet.append(["知识库-植保知识库-知识浏览", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_植保知识库_知识浏览.png", doc,
                                  "知识库_植保知识库_知识浏览")

        # 植保知识库-知识审核页
        time.sleep(3)
        self.driver.find_element(By.XPATH,
                                 "//span[text()='植保知识库']/parent::div/following-sibling::ul//li[text()='知识审核 ']").click()
        try:
            time.sleep(8)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@class='cell'][text()='发布时间']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开知识库-植保知识库-知识审核页，且指定元素存在")
            utils.g_logger.info("知识库-植保知识库-知识审核页显示正常")
            sheet.append(["知识库-植保知识库-知识审核", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_植保知识库_知识审核.png", doc,
                                  "知识库_植保知识库_知识审核")
        except Exception as e:
            utils.g_logger.info("知识库-植保知识库-知识审核页异常")
            sheet.append(["知识库-植保知识库-知识审核", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_植保知识库_知识审核.png", doc,
                                  "知识库_植保知识库_知识审核")

        # 植知识库-知识上传页
        time.sleep(3)
        self.driver.find_element(By.XPATH,
                                 "//span[text()='植保知识库']/parent::div/following-sibling::ul//li[text()='知识上传 ']").click()
        try:
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@class='cell'][text()='发布时间']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开知识库-植保知识库-知识上传页，且指定元素存在")
            utils.g_logger.info("知识库-植保知识库-知识上传页显示正常")
            sheet.append(["知识库-植保知识库-知识上传", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_植保知识库_知识上传.png", doc,
                                  "知识库_植保知识库_知识上传")
        except Exception as e:
            utils.g_logger.info("知识库-植保知识库-知识上传页异常")
            sheet.append(["知识库-植保知识库-知识上传", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_植保知识库_知识上传.png", doc,
                                  "知识库_植保知识库_知识上传")

        # 资料库页
        time.sleep(3)
        self.driver.find_element(By.XPATH, "//span[text()='资料库']").click()
        try:
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@class='cell'][text()='目录']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开知识库-资料库页，且指定元素存在")
            utils.g_logger.info("知识库-资料库页显示正常")
            sheet.append(["知识库-资料库", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_资料库.png", doc,
                                  "知识库_资料库")
        except Exception as e:
            utils.g_logger.info("知识库-资料库页异常")
            sheet.append(["知识库-资料库", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_资料库.png", doc,
                                  "知识库_资料库")

        # 作物知识库-知识浏览
        time.sleep(3)
        self.driver.find_element(By.XPATH, "//span[text()='作物知识库']").click()
        self.driver.find_element(By.XPATH,
                                 "//span[text()='作物知识库']/parent::div/following-sibling::ul//li[text()='知识浏览 ']").click()
        try:
            time.sleep(6)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@role='radiogroup']//span[text()='粮食作物']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开知识库-作物知识库-知识浏览页，且指定元素存在")
            image_elements = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//div[@id='tupianshow']//img")))
            unittest.TestCase.assertTrue(image_elements is not None,
                                         "成功打开知识库-作物知识库-知识浏览页，且指定元素存在")
            utils.g_logger.info("知识库-作物知识库-知识浏览页显示正常")
            sheet.append(["知识库-作物知识库-知识浏览", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/作物知识库_知识浏览.png", doc, "作物知识库_知识浏览")
        except Exception as e:
            utils.g_logger.info("知识库-作物知识库-知识浏览页异常")
            sheet.append(["知识库-作物知识库-知识浏览", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/作物知识库_知识浏览.png", doc, "作物知识库_知识浏览")

        # 作物知识库-知识维护
        time.sleep(3)
        self.driver.find_element(By.XPATH,
                                 "//span[text()='作物知识库']/parent::div/following-sibling::ul//li[text()='知识维护 ']").click()
        try:
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@class='cell'][text()='中文名']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开知识库-作物知识库-知识维护页，且指定元素存在")
            utils.g_logger.info("知识库-作物知识库-知识维护页显示正常")
            sheet.append(["知识库-作物知识库-知识维护", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/作物知识库_知识维护.png", doc, "作物知识库_知识维护")
        except Exception as e:
            utils.g_logger.info("知识库-作物知识库-知识维护页异常")
            sheet.append(["知识库-作物知识库-知识维护", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/作物知识库_知识维护.png", doc, "作物知识库_知识维护")

        # 关闭所有页面
        self.driver.find_element(By.XPATH, "//div[@class='close-con']").click()
        try:
            element = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//ul[@class='el-dropdown-menu el-popper']/li[text()='关闭所有']")))
            element.click()
            utils.g_logger.info("成功关闭所有页面")
        except Exception as e:
            utils.g_logger.error(e)

    @utils.retry(MAX_TRIES)
    def test_zhiwujianyi(self):
        # 植物检疫-产地检疫-产检申请
        try:
            time.sleep(5)
            self.driver.find_element(By.XPATH,
                                     "//div[@class='navigation-list']/div[@class='navigation-item']/span[text()=' 植物检疫']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//div[text()=' 农业植物产地检疫申请书 ']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植物检疫-产地检疫-产检申请页")
            utils.g_logger.info("植物检疫-产地检疫-产检申请页显示正常")
            sheet.append(["植物检疫-产地检疫-产检申请页", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/植物检疫-产地检疫-产检申请.png", doc,
                                  "植物检疫-产地检疫-产检申请")
        except Exception as e:
            utils.g_logger.info("植物检疫-产地检疫-产检申请页显示异常")
            sheet.append(["植物检疫-产地检疫-产检申请页", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植物检疫-产地检疫-产检申请页.png", doc,
                                  "植物检疫-产地检疫-产检申请页")

        # 植物检疫-产地检疫-产检受理
        try:
            self.driver.find_element(By.XPATH, "//li[contains(@class,'is-opened')]//li[text()='产检受理 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//div[@id='pane-产检受理']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植物检疫-产地检疫-产检受理页")
            utils.g_logger.info("植物检疫-产地检疫-产检受理页显示正常")
            sheet.append(["植物检疫-产地检疫-产检受理页", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/植物检疫-产地检疫-产检受理.png", doc,
                                  "植物检疫-产地检疫-产检受理")
        except Exception as e:
            utils.g_logger.info("植物检疫-产地检疫-产检受理页显示异常")
            sheet.append(["植物检疫-产地检疫-产检受理页", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植物检疫-产地检疫-产检受理页.png", doc,
                                  "植物检疫-产地检疫-产检受理页")

        # 植物检疫-产地检疫-田间调查
        try:
            self.driver.find_element(By.XPATH, "//li[contains(@class,'is-opened')]//li[text()='田间调查 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//div[@id='pane-田间调查']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植物检疫-产地检疫-田间调查页")
            utils.g_logger.info("植物检疫-产地检疫-田间调查页显示正常")
            sheet.append(["植物检疫-产地检疫-田间调查页", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/植物检疫-产地检疫-田间调查.png", doc,
                                  "植物检疫-产地检疫-田间调查")
        except Exception as e:
            utils.g_logger.info("植物检疫-产地检疫-田间调查页显示异常")
            sheet.append(["植物检疫-产地检疫-田间调查页", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植物检疫-产地检疫-田间调查页.png", doc,
                                  "植物检疫-产地检疫-田间调查页")

        # 植物检疫-产地检疫-实验室检验
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='实验室检验 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//label[text()='受检单位:']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植物检疫-产地检疫-实验室检验页")
            utils.g_logger.info("植物检疫-产地检疫-实验室检验页显示正常")
            sheet.append(["植物检疫-产地检疫-实验室检验页", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/植物检疫-产地检疫-实验室检验.png", doc,
                                  "植物检疫-产地检疫-实验室检验")
        except Exception as e:
            utils.g_logger.info("植物检疫-产地检疫-实验室检验页显示异常")
            sheet.append(["植物检疫-产地检疫-实验室检验页", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植物检疫-产地检疫-实验室检验页.png", doc,
                                  "植物检疫-产地检疫-实验室检验页")

        # 植物检疫-产地检疫-签发证书
        try:
            self.driver.find_element(By.XPATH, "//li[contains(@class,'is-opened')]//li[text()='签发证书 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//thead[@class='has-gutter']//div[text()='田调完成时间']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植物检疫-产地检疫-签发证书页")
            utils.g_logger.info("植物检疫-产地检疫-签发证书页显示正常")
            sheet.append(["植物检疫-产地检疫-签发证书页", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/植物检疫-产地检疫-签发证书.png", doc,
                                  "植物检疫-产地检疫-签发证书")
        except Exception as e:
            utils.g_logger.info("植物检疫-产地检疫-签发证书页显示异常")
            sheet.append(["植物检疫-产地检疫-签发证书页", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植物检疫-产地检疫-签发证书页.png", doc,
                                  "植物检疫-产地检疫-签发证书页")

        # 植物检疫-产地检疫-产检查询
        try:
            self.driver.find_element(By.XPATH, "//li[contains(@class,'is-opened')]//li[text()='产检查询 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//thead[@class='has-gutter']//div[text()='办理状态']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植物检疫-产地检疫-产检查询页")
            utils.g_logger.info("植物检疫-产地检疫-产检查询页显示正常")
            sheet.append(["植物检疫-产地检疫-产检查询页", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/植物检疫-产地检疫-产检查询.png", doc,
                                  "植物检疫-产地检疫-产检查询")
        except Exception as e:
            utils.g_logger.info("植物检疫-产地检疫-产检查询页显示异常")
            sheet.append(["植物检疫-产地检疫-产检查询页", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植物检疫-产地检疫-产检查询页.png", doc,
                                  "植物检疫-产地检疫-产检查询页")

        # 植物检疫-产地检疫-综合查询
        try:
            self.driver.find_element(By.XPATH, "//li[contains(@class,'is-opened')]//li[text()='综合查询 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//div[text()='快速查询']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植物检疫-产地检疫-综合查询页")
            utils.g_logger.info("植物检疫-产地检疫-综合查询页显示正常")
            sheet.append(["植物检疫-产地检疫-综合查询页", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/植物检疫-产地检疫-综合查询.png", doc,
                                  "植物检疫-产地检疫-综合查询")
        except Exception as e:
            utils.g_logger.info("植物检疫-产地检疫-综合查询页显示异常")
            sheet.append(["植物检疫-产地检疫-综合查询页", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植物检疫-产地检疫-综合查询页.png", doc,
                                  "植物检疫-产地检疫-综合查询页")

        # 调运检疫-检疫要求书
        try:
            self.driver.find_element(By.XPATH, "//span[text()='调运检疫']").click()
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='检疫要求书 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//button/span[text()='发检疫要求书']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开调运检疫-检疫要求书页")
            utils.g_logger.info("调运检疫-检疫要求书页显示正常")
            sheet.append(["调运检疫-检疫要求书", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/调运检疫-检疫要求书.png", doc,
                                  "调运检疫-检疫要求书")
        except Exception as e:
            utils.g_logger.info("调运检疫-检疫要求书页显示异常")
            sheet.append(["调运检疫-检疫要求书", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/调运检疫-检疫要求书.png", doc,
                                  "调运检疫-检疫要求书")

        # 调检申请-产检换证
        try:
            self.driver.find_element(By.XPATH, "//span[text()='调检申请']").click()
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='产检换证 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//div[@class='el-message-box__header']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开调运检疫-调检申请-产检换证页")
            utils.g_logger.info("调运检疫-调检申请-产检换证页显示正常")
            sheet.append(["调运检疫-调检申请-产检换证", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/调运检疫-调检申请-产检换证.png", doc,
                                  "调运检疫-调检申请-产检换证")
            self.driver.find_element(By.XPATH, "//i[@class='el-message-box__close el-icon-close']").click()
        except Exception as e:
            utils.g_logger.info("调运检疫-调检申请-产检换证页显示异常")
            sheet.append(["调运检疫-调检申请-产检换证", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/调运检疫-调检申请-产检换证.png", doc,
                                  "调运检疫-调检申请-产检换证")

        # 调检申请-再次调运
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='再次调运 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located(
                    (By.XPATH,
                     "//div[@class='el-message-box__header']//span[text()='请输入再调运的植物检疫证书号']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开调运检疫-调检申请-再次调运页")
            utils.g_logger.info("调运检疫-调检申请-再次调运页显示正常")
            sheet.append(["调运检疫-调检申请-再次调运", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/调运检疫-调检申请-再次调运.png", doc,
                                  "调运检疫-调检申请-再次调运")
            self.driver.find_element(By.XPATH, "//i[@class='el-message-box__close el-icon-close']").click()
        except Exception as e:
            utils.g_logger.info("调运检疫-调检申请-再次调运页显示异常")
            sheet.append(["调运检疫-调检申请-再次调运", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/调运检疫-调检申请-再次调运.png", doc,
                                  "调运检疫-调检申请-再次调运")

        # 调检申请-直接调运
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='直接调运 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//div[text()=' 农业植物调运检疫申请书 ']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开调运检疫-调检申请-直接调运页")
            utils.g_logger.info("调运检疫-调检申请-直接调运页显示正常")
            sheet.append(["调运检疫-调检申请-直接调运", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/调运检疫-调检申请-直接调运.png", doc,
                                  "调运检疫-调检申请-直接调运")
        except Exception as e:
            utils.g_logger.info("调运检疫-调检申请-直接调运页显示异常")
            sheet.append(["调运检疫-调检申请-直接调运", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/调运检疫-调检申请-直接调运.png", doc,
                                  "调运检疫-调检申请-直接调运")

        # 调运检疫-调检受理
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='调检受理 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//thead[@class='has-gutter']//div[@class='cell'][text()='申请单位或个人']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开调运检疫-调检受理页")
            utils.g_logger.info("调运检疫-调检受理页显示正常")
            sheet.append(["调运检疫-调检受理", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/调运检疫-调检受理.png", doc, "调运检疫-调检受理")
        except Exception as e:
            utils.g_logger.info("调运检疫-调检受理页显示异常")
            sheet.append(["调运检疫-调检受理", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/调运检疫-调检受理.png", doc,
                                  "调运检疫-调检受理")

        # 调运检疫-检疫检查
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='检疫检查 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//label[text()='检疫单证号:']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开调运检疫-检疫检查页")
            utils.g_logger.info("调运检疫-检疫检查页显示正常")
            sheet.append(["调运检疫-检疫检查", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/调运检疫-检疫检查.png", doc, "调运检疫-检疫检查")
        except Exception as e:
            utils.g_logger.info("调运检疫-检疫检查页显示异常")
            sheet.append(["调运检疫-检疫检查", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/调运检疫-检疫检查.png", doc,
                                  "调运检疫-检疫检查")

        # 调运检疫-实验室检验
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='实验室检验 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//thead[@class='has-gutter']//div[@class='cell'][text()='检验状态']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开调运检疫-实验室检验页")
            utils.g_logger.info("调运检疫-实验室检验页显示正常")
            sheet.append(["调运检疫-实验室检验", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/调运检疫-实验室检验.png", doc,
                                  "调运检疫-实验室检验")
        except Exception as e:
            utils.g_logger.info("调运检疫-实验室检验页显示异常")
            sheet.append(["调运检疫-实验室检验", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/调运检疫-实验室检验.png", doc,
                                  "调运检疫-实验室检验")

        # 调运检疫-签发证书
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='签发证书 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//thead[@class='has-gutter']//div[@class='cell'][text()='调运量']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开调运检疫-签发证书")
            utils.g_logger.info("调运检疫-签发证书页显示正常")
            sheet.append(["调运检疫-签发证书", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/调运检疫-签发证书.png", doc, "调运检疫-签发证书")
        except Exception as e:
            utils.g_logger.info("调运检疫-签发证书页显示异常")
            sheet.append(["调运检疫-签发证书", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/调运检疫-签发证书.png", doc,
                                  "调运检疫-签发证书")

        # 调运检疫-调检查询
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='调检查询 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//thead[@class='has-gutter']//div[@class='cell'][text()='调运类别']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开调运检疫-调检查询")
            utils.g_logger.info("调运检疫-调检查询页显示正常")
            sheet.append(["调运检疫-调检查询", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/调运检疫-调检查询.png", doc, "调运检疫-调检查询")
        except Exception as e:
            utils.g_logger.info("调运检疫-调检查询页显示异常")
            sheet.append(["调运检疫-调检查询", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/调运检疫-调检查询.png", doc,
                                  "调运检疫-调检查询")

        # 调运检疫-综合查询
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='综合查询 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//div[text()='换证查询']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开调运检疫-综合查询")
            utils.g_logger.info("调运检疫-综合查询页显示正常")
            sheet.append(["调运检疫-综合查询", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/调运检疫-综合查询.png", doc, "调运检疫-综合查询")
        except Exception as e:
            utils.g_logger.info("调运检疫-综合查询页显示异常")
            sheet.append(["调运检疫-综合查询", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/调运检疫-综合查询.png", doc,
                                  "调运检疫-综合查询")

        # 国外引种-引种申请
        try:
            self.driver.find_element(By.XPATH, "//span[text()='国外引种']").click()
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='引种申请 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//div[text()=' 国（境）外引进农业种苗检疫审批申请书 ']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开国外引种-引种申请页")
            utils.g_logger.info("国外引种-引种申请书页显示正常")
            sheet.append(["国外引种-引种申请", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/国外引种-引种申请.png", doc, "国外引种-引种申请")
        except Exception as e:
            utils.g_logger.info("国外引种-引种申请页显示异常")
            sheet.append(["国外引种-引种申请", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/国外引种-引种申请.png", doc,
                                  "国外引种-引种申请")

        # 国外引种-引种预审
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='引种预审 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//thead[@class='has-gutter']//div[@class='cell'][text()='引种单位']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开国外引种-引种预审页")
            utils.g_logger.info("国外引种-引种预审书页显示正常")
            sheet.append(["国外引种-引种预审", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/国外引种-引种预审.png", doc, "国外引种-引种预审")
        except Exception as e:
            utils.g_logger.info("国外引种-引种预审页显示异常")
            sheet.append(["国外引种-引种预审", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/国外引种-引种预审.png", doc,
                                  "国外引种-引种预审")

        # 国外引种-引种受理
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='引种受理 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//label[text()='引种类型：']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开国外引种-引种受理页")
            utils.g_logger.info("国外引种-引种受理书页显示正常")
            sheet.append(["国外引种-引种受理", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/国外引种-引种受理.png", doc, "国外引种-引种受理")
        except Exception as e:
            utils.g_logger.info("国外引种-引种受理页显示异常")
            sheet.append(["国外引种-引种受理", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/国外引种-引种受理.png", doc,
                                  "国外引种-引种受理")

        # 国外引种-审批签发
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='审批签发 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//thead[@class='has-gutter']//div[@class='cell'][text()='承诺截止时间']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开国外引种-审批签发页")
            utils.g_logger.info("国外引种-审批签发书页显示正常")
            sheet.append(["国外引种-审批签发", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/国外引种-审批签发.png", doc, "国外引种-审批签发")
        except Exception as e:
            utils.g_logger.info("国外引种-审批签发页显示异常")
            sheet.append(["国外引种-审批签发", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/国外引种-审批签发.png", doc,
                                  "国外引种-审批签发")

        # 国外引种-引种查询
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='引种查询 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//thead[@class='has-gutter']//div[@class='cell'][text()='拟种植省份']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开国外引种-引种查询页")
            utils.g_logger.info("国外引种-引种查询页显示正常")
            sheet.append(["国外引种-引种查询", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/国外引种-引种查询.png", doc, "国外引种-引种查询")
        except Exception as e:
            utils.g_logger.info("国外引种-引种查询页显示异常")
            sheet.append(["国外引种-引种查询", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/国外引种-引种查询.png", doc,
                                  "国外引种-引种查询")

        # 国外引种-引种跟踪
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='引种跟踪 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//label[text()='是否入境登记：']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开国外引种-引种跟踪页")
            utils.g_logger.info("国外引种-引种跟踪页显示正常")
            sheet.append(["国外引种-引种跟踪", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/国外引种-引种跟踪.png", doc, "国外引种-引种跟踪")
        except Exception as e:
            utils.g_logger.info("国外引种-引种跟踪页显示异常")
            sheet.append(["国外引种-引种跟踪", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/国外引种-引种跟踪.png", doc,
                                  "国外引种-引种跟踪")

        # 国外引种-引种田调
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='引种田调 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//label[text()='检疫单证号:']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开国外引种-引种田调页")
            utils.g_logger.info("国外引种-引种田调页显示正常")
            sheet.append(["国外引种-引种田调", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/国外引种-引种田调.png", doc, "国外引种-引种田调")
        except Exception as e:
            utils.g_logger.info("国外引种-引种田调页显示异常")
            sheet.append(["国外引种-引种田调", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/国外引种-引种田调.png", doc,
                                  "国外引种-引种田调")

        # 国外引种-引种检验
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='引种检验 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//label[text()='送样单位:']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开国外引种-引种检验页")
            utils.g_logger.info("国外引种-引种检验页显示正常")
            sheet.append(["国外引种-引种检验", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/国外引种-引种检验.png", doc, "国外引种-引种检验")
        except Exception as e:
            utils.g_logger.info("国外引种-引种检验页显示异常")
            sheet.append(["国外引种-引种检验", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/国外引种-引种检验.png", doc,
                                  "国外引种-引种检验")

        # 国外引种-综合查询
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='综合查询 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//div[text()='查询引种许可信息']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开国外引种-综合查询页")
            utils.g_logger.info("国外引种-综合查询页显示正常")
            sheet.append(["国外引种-综合查询", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/国外引种-引种检验.png", doc, "国外引种-综合查询")
        except Exception as e:
            utils.g_logger.info("国外引种-综合查询页显示异常")
            sheet.append(["国外引种-综合查询", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/国外引种-综合查询.png", doc,
                                  "国外引种-综合查询")

        # 疫情报告-快报
        try:
            self.driver.find_element(By.XPATH, "//span[text()='疫情报告']").click()
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='快报 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//label[text()='快报编号:']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开疫情报告-快报页")
            utils.g_logger.info("疫情报告-快报页显示正常")
            sheet.append(["疫情报告-快报", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/疫情报告-快报.png", doc, "疫情报告-快报")
        except Exception as e:
            utils.g_logger.info("疫情报告-快报页显示异常")
            sheet.append(["疫情报告-快报", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/疫情报告-快报.png", doc,
                                  "疫情报告-快报")

        # 疫情报告-月报
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='月报 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//label[text()='月报编号:']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开疫情报告-月报页")
            utils.g_logger.info("疫情报告-月报页显示正常")
            sheet.append(["疫情报告-月报", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/疫情报告-月报.png", doc, "疫情报告-月报")
        except Exception as e:
            utils.g_logger.info("疫情报告-月报页显示异常")
            sheet.append(["疫情报告-月报", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/疫情报告-月报.png", doc,
                                  "疫情报告-月报")

        # 疫情报告-年报
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='年报 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//label[text()='年报编号:']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开疫情报告-年报页")
            utils.g_logger.info("疫情报告-年报页显示正常")
            sheet.append(["疫情报告-年报", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/疫情报告-年报.png", doc, "疫情报告-年报")
        except Exception as e:
            utils.g_logger.info("疫情报告-年报页显示异常")
            sheet.append(["疫情报告-年报", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/疫情报告-年报.png", doc,
                                  "疫情报告-年报")

        # 检疫员-检疫员查询
        try:
            self.driver.find_element(By.XPATH, "//span[text()='检疫员']").click()
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='检疫员查询 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//div[@class='search-container']//label[text()='检疫员编号:']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开检疫员-检疫员查询页")
            utils.g_logger.info("检疫员-检疫员查询页显示正常")
            sheet.append(["检疫员-检疫员查询", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/检疫员-检疫员查询.png", doc, "检疫员-检疫员查询")
        except Exception as e:
            utils.g_logger.info("检疫员-检疫员查询页显示异常")
            sheet.append(["检疫员-检疫员查询", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/检疫员-检疫员查询.png", doc,
                                  "检疫员-检疫员查询")

        # 检疫员-新增检疫员
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='新增检疫员 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//button/span[text()='新增检疫员']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开检疫员-新增检疫员页")
            utils.g_logger.info("检疫员-新增检疫员页显示正常")
            sheet.append(["检疫员-新增检疫员", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/检疫员-新增检疫员.png", doc, "检疫员-新增检疫员")
        except Exception as e:
            utils.g_logger.info("检疫员-新增检疫员页显示异常")
            sheet.append(["检疫员-新增检疫员", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/检疫员-新增检疫员.png", doc,
                                  "检疫员-新增检疫员")

        # 检疫员-检疫员管理
        try:
            self.driver.find_element(By.XPATH,
                                     "//li[contains(@class,'is-opened')]//li[text()='检疫员管理 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//div[@class='search-container']//label[text()='检疫员编号:']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开检疫员-检疫员管理页")
            utils.g_logger.info("检疫员-检疫员管理页显示正常")
            sheet.append(["检疫员-检疫员管理", "正常"])
            # 获取网页截图并保存至word文档
            utils.page_screenshot(self.driver, "outputs/imagefiles/检疫员-检疫员管理.png", doc, "检疫员-检疫员管理")
        except Exception as e:
            utils.g_logger.info("检疫员-检疫员管理页显示异常")
            sheet.append(["检疫员-检疫员管理", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/检疫员-检疫员管理.png", doc,
                                  "检疫员-检疫员管理")

        # 执行“关闭所有页面”操作
        self.driver.find_element(By.XPATH, "//div[@class='close-con']").click()
        try:
            element = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, "//ul[@class='el-dropdown-menu el-popper']/li[text()='关闭所有']")))
            element.click()
            utils.g_logger.info("成功关闭所有页面")
        except Exception:
            utils.g_logger.info("关闭所有页面失败")

    @utils.retry(MAX_TRIES)
    def test_zhibaotixi(self):
        # 植保体系-上报机构信息
        try:
            time.sleep(5)
            self.driver.find_element(By.XPATH,
                                     "//div[@class='navigation-list']/div[@class='navigation-item']/span[text()=' 植保体系']").click()
            time.sleep(5)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//label[text()='植保机构名称']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植保体系-植保机构信息上报页")
            utils.g_logger.info("植保体系-植保机构信息上报页显示正常")
            sheet.append(["植保体系-植保机构信息上报", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-植保机构信息上报.png", doc,
                                  "植保体系-植保机构信息上报")
        except Exception as e:
            utils.g_logger.info("植保体系-植保机构信息上报页异常")
            sheet.append(["植保体系-植保机构信息上报", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-植保机构信息上报.png", doc,
                                  "植保体系-植保机构信息上报")

        # 植保体系-植保人员管理-植保人员信息上报
        try:
            self.driver.find_element(By.XPATH, "//span[text()='植保人员管理']").click()
            time.sleep(2)
            self.driver.find_element(By.XPATH,
                                     "//span[text()='植保人员管理']/parent::div/following-sibling::ul//li[text()='植保人员信息上报 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//thead[@class='has-gutter']//div[@class='cell'][text()='编制情况']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植保体系-植保人员管理-植保人员信息上报页")
            utils.g_logger.info("植保体系-植保人员管理-植保人员信息上报页显示正常")
            sheet.append(["植保体系-植保人员管理-植保人员信息上报", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-植保人员管理-植保人员信息上报.png", doc,
                                  "植保体系-植保人员管理-植保人员信息上报")
        except Exception as e:
            utils.g_logger.info("植保体系-植保人员管理-植保人员信息上报页异常")
            sheet.append(["植保体系-植保人员管理-植保人员信息上报", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-植保人员管理-植保人员信息上报.png", doc,
                                  "植保体系-植保人员管理-植保人员信息上报")

        # 植保体系-植保人员管理-乡镇植保人员信息上报

        try:
            self.driver.find_element(By.XPATH,
                                     "//span[text()='植保人员管理']/parent::div/following-sibling::ul//li[text()='乡镇植保人员信息上报 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//thead[@class='has-gutter']//div[@class='cell'][text()='所属乡镇机构名称']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植保体系-植保人员管理-乡镇植保人员信息上报页")
            utils.g_logger.info("植保体系-植保人员管理-乡镇植保人员信息上报页显示正常")
            sheet.append(["植保体系-植保人员管理-乡镇植保人员信息上报", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-植保人员管理-乡镇植保人员信息上报.png", doc,
                                  "植保体系-植保人员管理-乡镇植保人员信息上报")
        except Exception as e:
            utils.g_logger.info("植保体系-植保人员管理-乡镇植保人员信息上报页异常")
            sheet.append(["植保体系-植保人员管理-乡镇植保人员信息上报", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-植保人员管理-乡镇植保人员信息上报.png", doc,
                                  "植保体系-植保人员管理-乡镇植保人员信息上报")

        # 植保体系-植保人员管理-村级植保员信息上报
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[text()='植保人员管理']/parent::div/following-sibling::ul//li[text()='村级植保员信息上报 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//thead[@class='has-gutter']//div[@class='cell'][text()='服务行政村数量（个）']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植保体系-植保人员管理-村级植保员信息上报页")
            utils.g_logger.info("植保体系-植保人员管理-村级植保员信息上报页显示正常")
            sheet.append(["植保体系-植保人员管理-村级植保员信息上报", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-植保人员管理-村级植保员信息上报.png", doc,
                                  "植保体系-植保人员管理-村级植保员信息上报")
        except Exception as e:
            utils.g_logger.info("植保体系-植保人员管理-村级植保员信息上报页异常")
            sheet.append(["植保体系-植保人员管理-村级植保员信息上报", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-植保人员管理-村级植保员信息上报.png", doc,
                                  "植保体系-植保人员管理-村级植保员信息上报")

        # 植保体系-数据审核
        try:
            self.driver.find_element(By.XPATH, "//span[contains(@style,'margin-left')][text()='数据审核']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//div[@class='search-container']//label[text()='审核状态']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植保体系-数据审核页")
            utils.g_logger.info("植保体系-数据审核页显示正常")
            sheet.append(["植保体系-数据审核", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据审核.png", doc,
                                  "植保体系-数据审核")
        except Exception as e:
            utils.g_logger.info("植保体系-数据审核页异常")
            sheet.append(["植保体系-数据审核", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据审核.png", doc,
                                  "植保体系-数据审核")

        # 植保体系-数据查询-机构查询
        try:
            self.driver.find_element(By.XPATH, "//span[text()='数据查询']").click()
            time.sleep(2)
            self.driver.find_element(By.XPATH,
                                     "//span[text()='数据查询']/parent::div/following-sibling::ul//li[text()='机构查询 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//thead[@class='has-gutter']//div[@class='cell'][text()='植保机构名称']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植保体系-数据查询-机构查询页")
            utils.g_logger.info("植保体系-数据查询-机构查询页显示正常")
            sheet.append(["植保体系-数据查询-机构查询", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据查询-机构查询.png", doc,
                                  "植保体系-数据查询-机构查询")
        except Exception as e:
            utils.g_logger.info("植保体系-数据查询-机构查询页异常")
            sheet.append(["植保体系-数据查询-机构查询", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据查询-机构查询.png", doc,
                                  "植保体系-数据查询-机构查询")

        # 植保体系-数据查询-人员查询
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[text()='数据查询']/parent::div/following-sibling::ul//li[text()='人员查询 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//thead[@class='has-gutter']//div[@class='cell'][text()='人员类型']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植保体系-数据查询-人员查询页")
            utils.g_logger.info("植保体系-数据查询-人员查询页显示正常")
            sheet.append(["植保体系-数据查询-人员查询", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据查询-人员查询.png", doc,
                                  "植保体系-数据查询-人员查询")
        except Exception as e:
            utils.g_logger.info("植保体系-数据查询-人员查询页异常")
            sheet.append(["植保体系-数据查询-人员查询", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据查询-人员查询.png", doc,
                                  "植保体系-数据查询-人员查询")

        # 植保体系-数据查询-通讯录查询
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[text()='数据查询']/parent::div/following-sibling::ul//li[text()='通讯录查询 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//thead[@class='has-gutter']//div[@class='cell'][text()='手机号码']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植保体系-数据查询-通讯录查询页")
            utils.g_logger.info("植保体系-数据查询-通讯录查询页显示正常")
            sheet.append(["植保体系-数据查询-通讯录查询", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据查询-通讯录查询.png", doc,
                                  "植保体系-数据查询-通讯录查询")
        except Exception as e:
            utils.g_logger.info("植保体系-数据查询-通讯录查询页异常")
            sheet.append(["植保体系-数据查询-通讯录查询", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据查询-通讯录查询.png", doc,
                                  "植保体系-数据查询-通讯录查询")

        # 植保体系-数据统计-数据统计
        try:
            self.driver.find_element(By.XPATH, "//span[text()='数据统计']").click()
            time.sleep(2)
            self.driver.find_element(By.XPATH,
                                     "//span[text()='数据统计']/parent::div/following-sibling::ul//li[text()='数据统计 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//thead[contains(@class,'has-gutter')]//div[@class='cell'][text()='机构数量（个）']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植保体系-数据统计-数据统计页")
            utils.g_logger.info("植保体系-数据统计-数据统计页显示正常")
            sheet.append(["植保体系-数据统计-数据统计", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据统计-数据统计.png", doc,
                                  "植保体系-数据统计-数据统计")
        except Exception as e:
            utils.g_logger.info("植保体系-数据统计-数据统计页异常")
            sheet.append(["植保体系-数据统计-数据统计", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据统计-数据统计.png", doc,
                                  "植保体系-数据统计-数据统计")

        # 植保体系-数据统计-汇总统计-机构信息统计
        try:
            self.driver.find_element(By.XPATH, "//span[text()='汇总统计']").click()
            time.sleep(2)
            self.driver.find_element(By.XPATH,
                                     "//span[text()='汇总统计']/parent::div/following-sibling::ul//li[text()='机构信息统计 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//thead[contains(@class,'has-gutter')]//div[@class='cell'][text()='植保机构个数']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植保体系-数据统计-汇总统计-机构信息统计页")
            utils.g_logger.info("植保体系-数据统计-汇总统计-机构信息统计页显示正常")
            sheet.append(["植保体系-数据统计-汇总统计-机构信息统计", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据统计-汇总统计-机构信息统计.png", doc,
                                  "植保体系-数据统计-汇总统计-机构信息统计")
        except Exception as e:
            utils.g_logger.info("植保体系-数据统计-汇总统计-机构信息统计页异常")
            sheet.append(["植保体系-数据统计-汇总统计-机构信息统计", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据统计-汇总统计-机构信息统计.png", doc,
                                  "植保体系-数据统计-汇总统计-机构信息统计")

        # 植保体系-数据统计-汇总统计-植保机构统计
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[text()='汇总统计']/parent::div/following-sibling::ul//li[text()='植保机构统计 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//label[text()='行政区划层级']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植保体系-数据统计-汇总统计-植保机构统计页")
            utils.g_logger.info("植保体系-数据统计-汇总统计-植保机构统计页显示正常")
            sheet.append(["植保体系-数据统计-汇总统计-植保机构统计", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据统计-汇总统计-植保机构统计.png", doc,
                                  "植保体系-数据统计-汇总统计-植保机构统计")
        except Exception as e:
            utils.g_logger.info("植保体系-数据统计-汇总统计-植保机构统计页异常")
            sheet.append(["植保体系-数据统计-汇总统计-植保机构统计", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据统计-汇总统计-植保机构统计.png", doc,
                                  "植保体系-数据统计-汇总统计-植保机构统计")

        # 植保体系-数据统计-汇总统计-机构类别统计
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[text()='汇总统计']/parent::div/following-sibling::ul//li[text()='机构类别统计 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//label[text()='机构类型']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植保体系-数据统计-汇总统计-机构类别统计页")
            utils.g_logger.info("植保体系-数据统计-汇总统计-机构类别统计页显示正常")
            sheet.append(["植保体系-数据统计-汇总统计-机构类别统计", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据统计-汇总统计-机构类别统计.png", doc,
                                  "植保体系-数据统计-汇总统计-机构类别统计")
        except Exception as e:
            utils.g_logger.info("植保体系-数据统计-汇总统计-机构类别统计页异常")
            sheet.append(["植保体系-数据统计-汇总统计-机构类别统计", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据统计-汇总统计-机构类别统计.png", doc,
                                  "植保体系-数据统计-汇总统计-机构类别统计")

        # 植保体系-数据统计-汇总统计-人员年龄统计
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[text()='汇总统计']/parent::div/following-sibling::ul//li[text()='人员年龄统计 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//span[text()='人员年龄汇总（植保机构人员）：']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植保体系-数据统计-汇总统计-人员年龄统计页")
            utils.g_logger.info("植保体系-数据统计-汇总统计-人员年龄统计页显示正常")
            sheet.append(["植保体系-数据统计-汇总统计-人员年龄统计", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据统计-汇总统计-人员年龄统计.png", doc,
                                  "植保体系-数据统计-汇总统计-人员年龄统计")
        except Exception as e:
            utils.g_logger.info("植保体系-数据统计-汇总统计-人员年龄统计页异常")
            sheet.append(["植保体系-数据统计-汇总统计-人员年龄统计", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据统计-汇总统计-人员年龄统计.png", doc,
                                  "植保体系-数据统计-汇总统计-人员年龄统计")

        # 植保体系-数据统计-汇总统计-人员学历统计
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[text()='汇总统计']/parent::div/following-sibling::ul//li[text()='人员学历统计 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//span[text()='人员学历汇总（植保机构人员）：']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植保体系-数据统计-汇总统计-人员学历统计页")
            utils.g_logger.info("植保体系-数据统计-汇总统计-人员学历统计页显示正常")
            sheet.append(["植保体系-数据统计-汇总统计-人员学历统计", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据统计-汇总统计-人员学历统计.png", doc,
                                  "植保体系-数据统计-汇总统计-人员学历统计")
        except Exception as e:
            utils.g_logger.info("植保体系-数据统计-汇总统计-人员学历统计页异常")
            sheet.append(["植保体系-数据统计-汇总统计-人员学历统计", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据统计-汇总统计-人员学历统计.png", doc,
                                  "植保体系-数据统计-汇总统计-人员学历统计")

        # 植保体系-数据统计-汇总统计-填报进度统计
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[text()='汇总统计']/parent::div/following-sibling::ul//li[text()='填报进度统计 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//thead[contains(@class,'has-gutter')]//div[@class='cell'][text()='填报率（%）']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开植保体系-数据统计-汇总统计-填报进度统计页")
            utils.g_logger.info("植保体系-数据统计-汇总统计-填报进度统计页显示正常")
            sheet.append(["植保体系-数据统计-汇总统计-填报进度统计", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据统计-汇总统计-填报进度统计.png", doc,
                                  "植保体系-数据统计-汇总统计-填报进度统计")
        except Exception as e:
            utils.g_logger.info("植保体系-数据统计-汇总统计-填报进度统计页异常")
            sheet.append(["植保体系-数据统计-汇总统计-填报进度统计", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/植保体系-数据统计-汇总统计-填报进度统计.png", doc,
                                  "植保体系-数据统计-汇总统计-填报进度统计")
        # 执行“关闭所有页面”操作
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//div[@class='close-con']").click()
            time.sleep(2)
            element = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, "//ul[@class='el-dropdown-menu el-popper']/li[text()='关闭所有']")))
            element.click()
            utils.g_logger.info("成功关闭所有页面")
        except Exception:
            utils.g_logger.info("关闭所有页面失败")

    @utils.retry(MAX_TRIES)
    def test_bingchongfangzhi(self):
        # 病虫防治-工作平台页
        try:
            time.sleep(5)
            self.driver.find_element(By.XPATH,
                                     "//div[@class='navigation-list']/div[@class='navigation-item']/span[text()=' 病虫防治']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[text()=' 本站点任务设置情况 ']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫防治-工作平台页")
            utils.g_logger.info("病虫防治-工作平台页显示正常")
            sheet.append(["病虫防治-工作平台", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫防治_工作平台.png", doc,
                                  "病虫防治_工作平台")
        except Exception as e:
            utils.g_logger.info("病虫防治-工作平台页异常")
            sheet.append(["病虫防治-工作平台", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫防治_工作平台.png", doc,
                                  "病虫防治_工作平台")

        # 病虫防治-数据填报页
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='数据填报']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@class='cell'][text()='最迟填报时间']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫防治-数据填报页")
            utils.g_logger.info("病虫防治-数据填报页显示正常")
            sheet.append(["病虫防治-数据填报", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫防治_数据填报.png", doc,
                                  "病虫防治_数据填报")
        except Exception as e:
            utils.g_logger.info("病虫防治-数据填报页异常")
            sheet.append(["病虫防治-数据填报", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫防治_数据填报.png", doc,
                                  "病虫防治_数据填报")

        # 病虫防治-数据查询页
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='数据查询']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//div[@class='search-container']//button/span[text()='退回']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫防治-数据查询页")
            utils.g_logger.info("病虫防治-数据查询页显示正常")
            sheet.append(["病虫防治-数据查询", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫防治_数据查询.png", doc,
                                  "病虫防治_数据查询")
        except Exception as e:
            utils.g_logger.info("病虫防治-数据查询页异常")
            sheet.append(["病虫防治-数据查询", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫防治_数据查询.png", doc,
                                  "病虫防治_数据查询")

        # 病虫防治-数据汇总页
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='数据汇总']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[text()='显示字段']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫防治-数据汇总页")
            utils.g_logger.info("病虫防治-数据汇总页显示正常")
            sheet.append(["病虫防治-数据汇总", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫防治_数据汇总.png", doc,
                                  "病虫防治_数据汇总")
        except Exception as e:
            utils.g_logger.info("病虫防治-数据汇总页异常")
            sheet.append(["病虫防治-数据汇总", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫防治_数据汇总.png", doc,
                                  "病虫防治_数据汇总")

        # 病虫防治-催报查询页
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='催报查询']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//thead[@class='has-gutter']//div[@class='cell'][text()='催报报表']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫防治-催报查询页")
            utils.g_logger.info("病虫防治-催报查询页显示正常")
            sheet.append(["病虫防治-催报查询", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫防治_催报查询.png", doc,
                                  "病虫防治_催报查询")
        except Exception as e:
            utils.g_logger.info("病虫防治-催报查询页异常")
            sheet.append(["病虫防治-催报查询", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫防治_催报查询.png", doc,
                                  "病虫防治_催报查询")

        # 病虫防治-填报任务设置页
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='填报任务设置']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//button/span[text()=' 新增任务 ']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫防治-填报任务设置页")
            utils.g_logger.info("病虫防治-填报任务设置页显示正常")
            sheet.append(["病虫防治-填报任务设置", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫防治_填报任务设置.png", doc,
                                  "病虫防治_填报任务设置")
        except Exception as e:
            utils.g_logger.info("病虫防治-填报任务设置页异常")
            sheet.append(["病虫防治-填报任务设置", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫防治_填报任务设置.png", doc,
                                  "病虫防治_填报任务设置")

        # 病虫防治-任务审核页
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='任务审核']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//button/span[text()='审核通过']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫防治-任务审核页")
            utils.g_logger.info("病虫防治-任务审核页显示正常")
            sheet.append(["病虫防治-任务审核", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫防治_任务审核.png", doc,
                                  "病虫防治_任务审核")
        except Exception as e:
            utils.g_logger.info("病虫防治-任务审核页异常")
            sheet.append(["病虫防治-任务审核", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫防治_任务审核.png", doc,
                                  "病虫防治_任务审核")
        # 关闭所有页面
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//div[@class='close-con']").click()
            time.sleep(2)
            element = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, "//ul[@class='el-dropdown-menu el-popper']/li[text()='关闭所有']")))
            element.click()
            utils.g_logger.info("成功关闭所有页面")
        except Exception:
            utils.g_logger.info("关闭所有页面失败")

    @utils.retry(MAX_TRIES)
    def test_nongyaoxie(self):
        # 农药械-工作平台页
        try:
            time.sleep(5)
            self.driver.find_element(By.XPATH,
                                     "//div[@class='navigation-list']/div[@class='navigation-item']/span[text()=' 农药械']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[text()=' 本站点任务设置情况 ']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开农药械-工作平台页")
            utils.g_logger.info("农药械-工作平台页显示正常")
            sheet.append(["农药械-工作平台", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/农药械_工作平台.png", doc,
                                  "农药械_工作平台")
        except Exception as e:
            utils.g_logger.info("农药械-工作平台页异常")
            sheet.append(["农药械-工作平台", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/农药械_工作平台.png", doc,
                                  "农药械_工作平台")

        # 农药械-统防统治页
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='统防统治']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@class='button-center bigTitle'][text()='农药械']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开农药械-统防统治页")
            utils.g_logger.info("农药械-统防统治页显示正常")
            sheet.append(["农药械-统防统治", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/农药械_统防统治.png", doc,
                                  "农药械_统防统治")
        except Exception as e:
            utils.g_logger.info("农药械-统防统治页异常")
            sheet.append(["农药械-统防统治", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/农药械_统防统治.png", doc,
                                  "农药械_统防统治")

        # 农药械-数据填报页
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='数据填报']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@class='cell'][text()='最迟填报时间']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开农药械-数据填报页")
            utils.g_logger.info("农药械-数据填报页显示正常")
            sheet.append(["农药械-数据填报", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/农药械_数据填报.png", doc,
                                  "农药械_数据填报")
        except Exception as e:
            utils.g_logger.info("农药械-数据填报页异常")
            sheet.append(["农药械-数据填报", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/农药械_数据填报.png", doc,
                                  "农药械_数据填报")
        # 农药械-数据查询页
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='数据查询']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//div[@class='search-container']//button/span[text()='退回']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开农药械-数据查询页")
            utils.g_logger.info("农药械-数据查询页显示正常")
            sheet.append(["农药械-数据查询", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/农药械_数据查询.png", doc,
                                  "农药械_数据查询")
        except Exception as e:
            utils.g_logger.info("农药械-数据查询页异常")
            sheet.append(["农药械-数据查询", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/农药械_数据查询.png", doc,
                                  "农药械_数据查询")

        # 农药械-数据汇总页
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='数据汇总']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[text()='显示字段']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开农药械-数据汇总页")
            utils.g_logger.info("农药械-数据汇总页显示正常")
            sheet.append(["农药械-数据汇总", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/农药械_数据汇总.png", doc,
                                  "农药械_数据汇总")
        except Exception as e:
            utils.g_logger.info("农药械-数据汇总页异常")
            sheet.append(["农药械-数据汇总", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/农药械_数据汇总.png", doc,
                                  "农药械_数据汇总")

        # 农药械-催报查询页
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='催报查询']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//thead[@class='has-gutter']//div[@class='cell'][text()='催报报表']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开农药械-催报查询页")
            utils.g_logger.info("农药械-催报查询页显示正常")
            sheet.append(["农药械-催报查询", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/农药械_催报查询.png", doc,
                                  "农药械_催报查询")
        except Exception as e:
            utils.g_logger.info("农药械-催报查询页异常")
            sheet.append(["农药械-催报查询", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/农药械_催报查询.png", doc,
                                  "农药械_催报查询")

        # 农药械-填报任务设置页
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='填报任务设置']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//button/span[text()=' 新增任务 ']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开农药械-填报任务设置页")
            utils.g_logger.info("农药械-填报任务设置页显示正常")
            sheet.append(["农药械-填报任务设置", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/农药械_填报任务设置.png", doc,
                                  "农药械_填报任务设置")
        except Exception as e:
            utils.g_logger.info("农药械-填报任务设置页异常")
            sheet.append(["农药械-填报任务设置", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/农药械_填报任务设置.png", doc,
                                  "农药械_填报任务设置")

        # 农药械-随时报送页
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='随时报送']").click()
            time.sleep(6)
            self.driver.find_element(By.XPATH, "//input[@placeholder='-请选择-']").click()

            elements = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, "//ul[starts-with(@id,'reportTree')]//span[contains(text(),'发生防治及损失情况')]")))
            unittest.TestCase.assertTrue(elements is not None, "成功打开农药械-随时报送页")
            utils.g_logger.info("农药械-随时报送页显示正常")
            sheet.append(["农药械-随时报送", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/农药械_随时报送.png", doc,
                                  "农药械_随时报送")
        except Exception as e:
            utils.g_logger.info("农药械-随时报送页异常")
            sheet.append(["农药械-随时报送", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/农药械_随时报送.png", doc,
                                  "农药械_随时报送")

        # 关闭所有页面
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//div[@class='close-con']").click()
            time.sleep(2)
            element = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, "//ul[@class='el-dropdown-menu el-popper']/li[text()='关闭所有']")))
            element.click()
            utils.g_logger.info("成功关闭所有页面")
        except Exception:
            utils.g_logger.info("关闭所有页面失败")

    @utils.retry(MAX_TRIES)
    def test_bingchongzhuanti(self):
        # 病虫专题-数据分析-专题分析
        try:
            time.sleep(5)
            self.driver.find_element(By.XPATH,
                                     "//div[@class='navigation-list']/div[@class='navigation-item']/span[text()=' 病虫专题']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//div[@class='zbcard']//img")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-数据分析-专题分析页")
            utils.g_logger.info("病虫专题-数据分析-专题分析页显示正常")
            sheet.append(["病虫专题-数据分析-专题分析", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_专题分析.png", doc,
                                  "病虫专题_数据分析_专题分析")
        except Exception as e:
            utils.g_logger.info("病虫专题-数据分析-专题分析页异常")
            sheet.append(["病虫专题-数据分析-专题分析", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_专题分析.png", doc,
                                  "病虫专题_数据分析_专题分析")

        # 病虫专题-数据分析-GIS分析
        try:
            self.driver.find_element(By.XPATH, "//li[contains(@class,'is-opened')]//li[text()='GIS分析 ']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//span[text()='插值图']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-数据分析-GIS分析页")
            utils.g_logger.info("病虫专题-数据分析-GIS分析页显示正常")
            sheet.append(["病虫专题-数据分析-GIS分析", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_GIS分析.png", doc,
                                  "病虫专题_数据分析_GIS分析")
        except Exception as e:
            utils.g_logger.info("病虫专题-数据分析-GIS分析页异常")
            sheet.append(["病虫专题-数据分析-GIS分析", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_GIS分析.png", doc,
                                  "病虫专题_数据分析_GIS分析")

        # 病虫专题-数据分析-自定义分析
        try:
            self.driver.find_element(By.XPATH, "//li[contains(@class,'is-opened')]//li[text()='自定义分析 ']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//label[text()='分析类型']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-数据分析-自定义分析页")
            utils.g_logger.info("病虫专题-数据分析-自定义分析页显示正常")
            sheet.append(["病虫专题-数据分析-自定义分析", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_自定义分析.png", doc,
                                  "病虫专题_数据分析_自定义分析")
        except Exception as e:
            utils.g_logger.info("病虫专题-数据分析-自定义分析页异常")
            sheet.append(["病虫专题-数据分析-自定义分析", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_自定义分析.png", doc,
                                  "病虫专题_数据分析_自定义分析")

        # 病虫专题-数据分析-数据报告
        try:
            self.driver.find_element(By.XPATH, "//li[contains(@class,'is-opened')]//li[text()='数据报告 ']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//div[@role='radiogroup']//span[text()=' 年报 ']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-数据分析-数据报告页")
            utils.g_logger.info("病虫专题-数据分析-数据报告页显示正常")
            sheet.append(["病虫专题-数据分析-数据报告", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_数据报告.png", doc,
                                  "病虫专题_数据分析_数据报告")
        except Exception as e:
            utils.g_logger.info("病虫专题-数据分析-数据报告页异常")
            sheet.append(["病虫专题-数据分析-数据报告", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_数据报告.png", doc,
                                  "病虫专题_数据分析_数据报告")

        # 病虫专题-数据分析-扩展推演
        try:
            self.driver.find_element(By.XPATH, "//li[contains(@class,'is-opened')]//li[text()='扩展推演 ']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//div[@class='legendBox']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-数据分析-扩展推演页")
            utils.g_logger.info("病虫专题-数据分析-扩展推演页显示正常")
            sheet.append(["病虫专题-数据分析-扩展推演", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_扩展推演.png", doc,
                                  "病虫专题_数据分析_扩展推演")
        except Exception as e:
            utils.g_logger.info("病虫专题-数据分析-扩展推演页异常")
            sheet.append(["病虫专题-数据分析-扩展推演", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_扩展推演.png", doc,
                                  "病虫专题_数据分析_扩展推演")

        # 病虫专题-数据分析-在线作图
        try:
            self.driver.find_element(By.XPATH, "//li[contains(@class,'is-opened')]//li[text()='在线作图 ']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//button/span[text()='图形设置']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-数据分析-在线作图页")
            utils.g_logger.info("病虫专题-数据分析-在线作图页显示正常")
            sheet.append(["病虫专题-数据分析-在线作图", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_在线作图.png", doc,
                                  "病虫专题_数据分析_在线作图")
        except Exception as e:
            utils.g_logger.info("病虫专题-数据分析-在线作图页异常")
            sheet.append(["病虫专题-数据分析-在线作图", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_在线作图.png", doc,
                                  "病虫专题_数据分析_在线作图")

        # 病虫专题-数据分析-定制报告
        try:
            self.driver.find_element(By.XPATH, "//li[contains(@class,'is-opened')]//li[text()='定制报告 ']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//div[@role='radiogroup']//span[text()=' 年报 ']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-数据分析-定制报告页")
            utils.g_logger.info("病虫专题-数据分析-定制报告页显示正常")
            sheet.append(["病虫专题-数据分析-定制报告", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_定制报告.png", doc,
                                  "病虫专题_数据分析_定制报告")
        except Exception as e:
            utils.g_logger.info("病虫专题-数据分析-定制报告页异常")
            sheet.append(["病虫专题-数据分析-定制报告", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_定制报告.png", doc,
                                  "病虫专题_数据分析_定制报告")

        # 病虫专题-数据分析-数据治理-数据地图配置
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//li[contains(@class,'is-opened')]//span[text()='数据治理']").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//ul//li[text()='数据地图配置 ']").click()
            time.sleep(4)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//div[@class='cell'][text()='来源_数据集名称']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-数据分析-数据治理-数据地图配置")
            utils.g_logger.info("病虫专题-数据分析-数据治理-数据地图配置页显示正常")
            sheet.append(["病虫专题-数据分析-数据治理-数据地图配置", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_数据治理_数据地图配置.png", doc,
                                  "病虫专题_数据分析_数据治理_数据地图配置")
        except Exception as e:
            utils.g_logger.info("病虫专题-数据分析-数据治理-数据地图配置页异常")
            sheet.append(["病虫专题-数据分析-数据治理-数据地图配置", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_数据治理_数据地图配置.png", doc,
                                  "病虫专题_数据分析_数据治理_数据地图配置")

        # 病虫专题-数据分析-数据治理-数据资源管理
        try:
            self.driver.find_element(By.XPATH, "//ul//li[text()='数据资源管理 ']").click()
            time.sleep(4)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//button/span[text()='新建数据集']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-数据分析-数据治理-数据资源管理")
            utils.g_logger.info("病虫专题-数据分析-数据治理-数据资源管理页显示正常")
            sheet.append(["病虫专题-数据分析-数据治理-数据资源管理", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_数据治理_数据资源管理.png", doc,
                                  "病虫专题_数据分析_数据治理_数据资源管理")
        except Exception as e:
            utils.g_logger.info("病虫专题-数据分析-数据治理-数据资源管理页异常")
            sheet.append(["病虫专题-数据分析-数据治理-数据资源管理", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_数据治理_数据资源管理.png", doc,
                                  "病虫专题_数据分析_数据治理_数据资源管理")

        # 病虫专题-数据分析-数据治理-数据分析配置
        try:
            self.driver.find_element(By.XPATH, "//ul//li[text()='数据分析配置 ']").click()
            time.sleep(4)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//button/span[text()='新建分析主题']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-数据分析-数据治理-数据分析配置页面")
            utils.g_logger.info("病虫专题-数据分析-数据治理-数据分析配置页显示正常")
            sheet.append(["病虫专题-数据分析-数据治理-数据分析配置", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_数据治理_数据分析配置.png", doc,
                                  "病虫专题_数据分析_数据治理_数据分析配置")
        except Exception as e:
            utils.g_logger.info("病虫专题-数据分析-数据治理-数据分析配置页异常")
            sheet.append(["病虫专题-数据分析-数据治理-数据分析配置", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_数据治理_数据分析配置.png", doc,
                                  "病虫专题_数据分析_数据治理_数据分析配置")

        # 病虫专题-数据分析-数据治理-数据分析浏览
        try:
            self.driver.find_element(By.XPATH, "//ul//li[text()='数据分析浏览 ']").click()
            time.sleep(4)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, "//div[@id='dashboardTree']//span[text()='检疫审批数据挖掘']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-数据分析-数据治理-数据分析浏览页面")
            utils.g_logger.info("病虫专题-数据分析-数据治理-数据分析浏览页显示正常")
            sheet.append(["病虫专题-数据分析-数据治理-数据分析浏览", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_数据治理_数据分析浏览.png", doc,
                                  "病虫专题_数据分析_数据治理_数据分析浏览")
        except Exception as e:
            utils.g_logger.info("病虫专题-数据分析-数据治理-数据分析浏览页异常")
            sheet.append(["病虫专题-数据分析-数据治理-数据分析浏览", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_数据治理_数据分析浏览.png", doc,
                                  "病虫专题_数据分析_数据治理_数据分析浏览")

        # 病虫专题-数据分析-数据治理-数据资源查看
        try:
            self.driver.find_element(By.XPATH, "//ul//li[text()='数据资源查看 ']").click()
            time.sleep(4)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//div[@class='relationgraphbox']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-数据分析-数据治理-数据资源查看页面")
            utils.g_logger.info("病虫专题-数据分析-数据治理-数据资源查看页显示正常")
            sheet.append(["病虫专题-数据分析-数据治理-数据资源查看", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_数据治理_数据资源查看.png", doc,
                                  "病虫专题_数据分析_数据治理_数据资源查看")
        except Exception as e:
            utils.g_logger.info("病虫专题-数据分析-数据治理-数据资源查看页异常")
            sheet.append(["病虫专题-数据分析-数据治理-数据资源查看", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_数据分析_数据治理_数据资源查看.png", doc,
                                  "病虫专题_数据分析_数据治理_数据资源查看")

        # 病虫专题-稻飞虱
        try:
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='稻飞虱']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//div[@id='navTitle'][text()='稻飞虱 · 发生分布']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-稻飞虱页")
            utils.g_logger.info("病虫专题-稻飞虱页显示正常")
            sheet.append(["病虫专题-稻飞虱", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_稻飞虱.png", doc,
                                  "病虫专题_稻飞虱")
        except Exception as e:
            utils.g_logger.info("病虫专题-稻飞虱页异常")
            sheet.append(["病虫专题-稻飞虱", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_稻飞虱.png", doc,
                                  "病虫专题_稻飞虱")

        # 病虫专题-赤霉病
        try:
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='赤霉病']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, "//div[@id='navTitle'][text()='小麦赤霉病专题 · 发生分布']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-赤霉病页")
            utils.g_logger.info("病虫专题-赤霉病页显示正常")
            sheet.append(["病虫专题-赤霉病", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_赤霉病.png", doc,
                                  "病虫专题_赤霉病")
        except Exception as e:
            utils.g_logger.info("病虫专题-赤霉病页异常")
            sheet.append(["病虫专题-赤霉病", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_赤霉病.png", doc,
                                  "病虫专题_赤霉病")

        # 病虫专题-条锈病
        try:
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='条锈病']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, "//div[@id='navTitle'][text()='条锈病湖北 · 发生分布']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-条锈病页")
            utils.g_logger.info("病虫专题-条锈病页显示正常")
            sheet.append(["病虫专题-条锈病", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_条锈病.png", doc,
                                  "病虫专题_条锈病")
        except Exception as e:
            utils.g_logger.info("病虫专题-条锈病页异常")
            sheet.append(["病虫专题-条锈病", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_条锈病.png", doc,
                                  "病虫专题_条锈病")

        # 病虫专题-马铃薯晚疫病-GIS分析
        try:
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='马铃薯晚疫病']").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//li[contains(@class,'is-opened')]//li[text()='GIS分析 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//div[@class='legend']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-马铃薯晚疫病-GIS分析页")
            utils.g_logger.info("病虫专题-马铃薯晚疫病-GIS分析页显示正常")
            sheet.append(["病虫专题-马铃薯晚疫病-GIS分析", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_马铃薯晚疫病_GIS分析.png", doc,
                                  "病虫专题_马铃薯晚疫病_GIS分析")
        except Exception as e:
            utils.g_logger.info("病虫专题-马铃薯晚疫病-GIS分析页异常")
            sheet.append(["病虫专题-马铃薯晚疫病-GIS分析", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_马铃薯晚疫病_GIS分析.png", doc,
                                  "病虫专题_马铃薯晚疫病_GIS分析")

        # 病虫专题-马铃薯晚疫病-监测期设置
        try:
            self.driver.find_element(By.XPATH, "//li[contains(@class,'is-opened')]//li[text()='监测期设置 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//th[text()='出苗期']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-马铃薯晚疫病-监测期设置页")
            utils.g_logger.info("病虫专题-马铃薯晚疫病-监测期设置页显示正常")
            sheet.append(["病虫专题-马铃薯晚疫病-监测期设置", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_马铃薯晚疫病_监测期设置.png", doc,
                                  "病虫专题_马铃薯晚疫病_监测期设置")
        except Exception as e:
            utils.g_logger.info("病虫专题-马铃薯晚疫病-监测期设置页异常")
            sheet.append(["病虫专题-马铃薯晚疫病-监测期设置", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_马铃薯晚疫病_监测期设置.png", doc,
                                  "病虫专题_马铃薯晚疫病_监测期设置")

        # 病虫专题-马铃薯晚疫病-侵染曲线
        try:
            self.driver.find_element(By.XPATH, "//li[contains(@class,'is-opened')]//li[text()='侵染曲线 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//div[text()='侵染程度：']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-马铃薯晚疫病-侵染曲线页")
            utils.g_logger.info("病虫专题-马铃薯晚疫病-侵染曲线页显示正常")
            sheet.append(["病虫专题-马铃薯晚疫病-侵染曲线", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_马铃薯晚疫病_侵染曲线.png", doc,
                                  "病虫专题_马铃薯晚疫病_侵染曲线")
        except Exception as e:
            utils.g_logger.info("病虫专题-马铃薯晚疫病-侵染曲线页异常")
            sheet.append(["病虫专题-马铃薯晚疫病-侵染曲线", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_马铃薯晚疫病_侵染曲线.png", doc,
                                  "病虫专题_马铃薯晚疫病_侵染曲线")

        # 病虫专题-马铃薯晚疫病-湿润期统计
        try:
            self.driver.find_element(By.XPATH, "//li[contains(@class,'is-opened')]//li[text()='湿润期统计 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//label[text()='监测期']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-马铃薯晚疫病-湿润期统计页")
            utils.g_logger.info("病虫专题-马铃薯晚疫病-湿润期统计页显示正常")
            sheet.append(["病虫专题-马铃薯晚疫病-湿润期统计", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_马铃薯晚疫病_湿润期统计.png", doc,
                                  "病虫专题_马铃薯晚疫病_湿润期统计")
        except Exception as e:
            utils.g_logger.info("病虫专题-马铃薯晚疫病-湿润期统计页异常")
            sheet.append(["病虫专题-马铃薯晚疫病-湿润期统计", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_马铃薯晚疫病_湿润期统计.png", doc,
                                  "病虫专题_马铃薯晚疫病_湿润期统计")

        # 病虫专题-马铃薯晚疫病-数据查询
        try:
            self.driver.find_element(By.XPATH, "//li[contains(@class,'is-opened')]//li[text()='数据查询 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//th[text()='数据时间']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-马铃薯晚疫病-数据查询页")
            utils.g_logger.info("病虫专题-马铃薯晚疫病-数据查询页显示正常")
            sheet.append(["病虫专题-马铃薯晚疫病-数据查询", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_马铃薯晚疫病_数据查询.png", doc,
                                  "病虫专题_马铃薯晚疫病_数据查询")
        except Exception as e:
            utils.g_logger.info("病虫专题-马铃薯晚疫病-数据查询页异常")
            sheet.append(["病虫专题-马铃薯晚疫病-数据查询", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_马铃薯晚疫病_数据查询.png", doc,
                                  "病虫专题_马铃薯晚疫病_数据查询")

        # 病虫专题-马铃薯晚疫病-数据统计
        try:
            self.driver.find_element(By.XPATH, "//li[contains(@class,'is-opened')]//li[text()='数据统计 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//th[text()='数据时间']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-马铃薯晚疫病-数据统计页")
            utils.g_logger.info("病虫专题-马铃薯晚疫病-数据统计页显示正常")
            sheet.append(["病虫专题-马铃薯晚疫病-数据统计", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_马铃薯晚疫病_数据统计.png", doc,
                                  "病虫专题_马铃薯晚疫病_数据统计")
        except Exception as e:
            utils.g_logger.info("病虫专题-马铃薯晚疫病-数据统计页异常")
            sheet.append(["病虫专题-马铃薯晚疫病-数据统计", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_马铃薯晚疫病_数据统计.png", doc,
                                  "病虫专题_马铃薯晚疫病_数据统计")

        # 病虫专题-稻纵卷叶螟
        try:
            self.driver.find_element(By.XPATH, "//span[@style='margin-left: 20px;'][text()='稻纵卷叶螟']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//div[text()='卷叶率预测模型']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开病虫专题-稻纵卷叶螟页")
            utils.g_logger.info("病虫专题-稻纵卷叶螟页显示正常")
            sheet.append(["病虫专题-稻纵卷叶螟", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_稻纵卷叶螟.png", doc,
                                  "病虫专题_稻纵卷叶螟")
        except Exception as e:
            utils.g_logger.info("病虫专题-稻纵卷叶螟页异常")
            sheet.append(["病虫专题-稻纵卷叶螟", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/病虫专题_稻纵卷叶螟.png", doc,
                                  "病虫专题_稻纵卷叶螟")

        # 关闭所有页面
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//div[@class='close-con']").click()
            time.sleep(2)
            element = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, "//ul[@class='el-dropdown-menu el-popper']/li[text()='关闭所有']")))
            element.click()
            utils.g_logger.info("成功关闭所有页面")
        except Exception:
            utils.g_logger.info("关闭所有页面失败")

    @utils.retry(MAX_TRIES)
    def test_shuzituku(self):
        # 数字图库页
        try:
            time.sleep(4)
            self.driver.find_element(By.XPATH,
                                     "//div[@class='navigation-list']/div[@class='navigation-item']/span[text()=' 数字图库']").click()
            time.sleep(5)
            # 获取所有窗口的句柄
            windows = self.driver.window_handles
            # 切换到新窗口，通常新窗口是最后一个句柄
            self.driver.switch_to.window(windows[-1])
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//li[text()=' 标本数据 ']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开数字图库页")
            utils.g_logger.info("数字图库页显示正常")
            sheet.append(["数字图库", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数字图库.png", doc,
                                  "数字图库")
            self.driver.close()
            self.driver.switch_to.window(windows[0])
        except Exception as e:
            utils.g_logger.info("数字图库页异常")
            sheet.append(["数字图库", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/数字图库.png", doc,
                                  "数字图库")

    @utils.retry(MAX_TRIES)
    def test_bangongyingyong(self):
        # 办公应用-文件收发管理-收件箱
        try:
            time.sleep(4)
            self.driver.find_element(By.XPATH,
                                     "//div[@class='navigation-list']/div[@class='navigation-item']/span[text()=' 办公应用']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@class='cell'][text()='发件单位']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开办公应用-文件收发管理-收件箱页")
            utils.g_logger.info("办公应用-文件收发管理-收件箱页显示正常")
            sheet.append(["办公应用-文件收发管理-收件箱", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用_文件收发管理_收件箱.png", doc,
                                  "办公应用_文件收发管理_收件箱")
        except Exception as e:
            utils.g_logger.info("办公应用-文件收发管理-收件箱页异常")
            sheet.append(["办公应用-文件收发管理-收件箱", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用-文件收发管理-收件箱.png", doc,
                                  "办公应用-文件收发管理-收件箱")

        # 办公应用-文件收发管理-草稿箱
        try:
            self.driver.find_element(By.XPATH, "//li[text()='草稿箱 ']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//div[@class='search-container']//label[text()='消息类别']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开办公应用-文件收发管理-草稿箱页")
            utils.g_logger.info("办公应用-文件收发管理-草稿箱页显示正常")
            sheet.append(["办公应用-文件收发管理-草稿箱", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用_文件收发管理_草稿箱.png", doc,
                                  "办公应用_文件收发管理_草稿箱")
        except Exception as e:
            utils.g_logger.info("办公应用-文件收发管理-草稿箱页异常")
            sheet.append(["办公应用-文件收发管理-草稿箱", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用-文件收发管理-草稿箱.png", doc,
                                  "办公应用-文件收发管理-草稿箱")

        # 办公应用-文件收发管理-发件箱
        try:
            self.driver.find_element(By.XPATH, "//li[text()='发件箱 ']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//label[text()='发送时间']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开办公应用-文件收发管理-发件箱页")
            utils.g_logger.info("办公应用-文件收发管理-发件箱页显示正常")
            sheet.append(["办公应用-文件收发管理-发件箱", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用_文件收发管理_发件箱.png", doc,
                                  "办公应用_文件收发管理_发件箱")
        except Exception as e:
            utils.g_logger.info("办公应用-文件收发管理-发件箱页异常")
            sheet.append(["办公应用-文件收发管理-发件箱", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用-文件收发管理-发件箱.png", doc,
                                  "办公应用-文件收发管理-发件箱")

        # 办公应用-辅助工具-方差分析
        try:
            self.driver.find_element(By.XPATH, "//span[@slot='title'][text()='辅助工具']").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//ul//li[text()='方差分析 ']").click()
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//button/span[text()='计算']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开办公应用-辅助工具-方差分析页")
            utils.g_logger.info("办公应用-辅助工具-方差分析页显示正常")
            sheet.append(["办公应用-辅助工具-方差分析", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用_辅助工具_方差分析.png", doc,
                                  "办公应用_辅助工具_方差分析")
        except Exception as e:
            utils.g_logger.info("办公应用-辅助工具-方差分析页异常")
            sheet.append(["办公应用-辅助工具-方差分析", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用-辅助工具-方差分析.png", doc,
                                  "办公应用-辅助工具-方差分析")

        # 办公应用-辅助工具-线性回归
        try:
            self.driver.find_element(By.XPATH, "//ul//li[text()='线性回归 ']").click()
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//button/span[text()='计算']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开办公应用-辅助工具-线性回归页")
            utils.g_logger.info("办公应用-辅助工具-线性回归页显示正常")
            sheet.append(["办公应用-辅助工具-线性回归", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用_辅助工具_线性回归.png", doc,
                                  "办公应用_辅助工具_线性回归")
        except Exception as e:
            utils.g_logger.info("办公应用-辅助工具-线性回归页异常")
            sheet.append(["办公应用-辅助工具-线性回归", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用-辅助工具-线性回归.png", doc,
                                  "办公应用-辅助工具-线性回归")

        # 办公应用-辅助工具-T检验
        try:
            self.driver.find_element(By.XPATH, "//ul//li[text()='T检验 ']").click()
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//label[text()='给定常量']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开办公应用-辅助工具-T检验页")
            utils.g_logger.info("办公应用-辅助工具-T检验页显示正常")
            sheet.append(["办公应用-辅助工具-T检验", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用_辅助工具_T检验.png", doc,
                                  "办公应用_辅助工具_T检验")
        except Exception as e:
            utils.g_logger.info("办公应用-辅助工具-T检验页异常")
            sheet.append(["办公应用-辅助工具-T检验", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用-辅助工具-T检验.png", doc,
                                  "办公应用-辅助工具-T检验")

        # 办公应用-辅助工具-非线性回归
        try:
            self.driver.find_element(By.XPATH, "//ul//li[text()='非线性回归 ']").click()
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//button/span[text()='计算']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开办公应用-辅助工具-非线性回归页")
            utils.g_logger.info("办公应用-辅助工具-非线性回归页显示正常")
            sheet.append(["办公应用-辅助工具-非线性回归", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用_辅助工具_非线性回归.png", doc,
                                  "办公应用_辅助工具_非线性回归")
        except Exception as e:
            utils.g_logger.info("办公应用-辅助工具-非线性回归页异常")
            sheet.append(["办公应用-辅助工具-非线性回归", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用-辅助工具-非线性回归.png", doc,
                                  "办公应用-辅助工具-非线性回归")

        # 办公应用-辅助工具-多因方差分析
        try:
            self.driver.find_element(By.XPATH, "//ul//li[text()='多因方差分析 ']").click()
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//button/span[text()='计算']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开办公应用-辅助工具-多因方差分析页")
            utils.g_logger.info("办公应用-辅助工具-多因方差分析页显示正常")
            sheet.append(["办公应用-辅助工具-多因方差分析", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用_辅助工具_多因方差分析.png", doc,
                                  "办公应用_辅助工具_多因方差分析")
        except Exception as e:
            utils.g_logger.info("办公应用-辅助工具-多因方差分析页异常")
            sheet.append(["办公应用-辅助工具-多因方差分析", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用-辅助工具-多因方差分析.png", doc,
                                  "办公应用-辅助工具-多因方差分析")

        # 办公应用-病虫害情报库-情报上传
        try:
            self.driver.find_element(By.XPATH, "//span[@slot='title'][text()='病虫害情报库']").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//ul//li[text()='情报上传 ']").click()
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//button/span[text()='新增']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开办公应用-病虫害情报库-情报上传页")
            utils.g_logger.info("办公应用-病虫害情报库-情报上传页显示正常")
            sheet.append(["办公应用-病虫害情报库-情报上传", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用_病虫害情报库-情报上传.png", doc,
                                  "办公应用_病虫害情报库-情报上传")
        except Exception as e:
            utils.g_logger.info("办公应用-病虫害情报库-情报上传页异常")
            sheet.append(["办公应用-病虫害情报库-情报上传", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用-病虫害情报库-情报上传.png", doc,
                                  "办公应用-病虫害情报库-情报上传")

        # 办公应用-病虫害情报库-情报浏览
        try:
            self.driver.find_element(By.XPATH, "//ul//li[text()='情报浏览 ']").click()
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//label[text()='关键字']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开办公应用-病虫害情报库-情报浏览页")
            utils.g_logger.info("办公应用-病虫害情报库-情报浏览页显示正常")
            sheet.append(["办公应用-病虫害情报库-情报浏览", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用_病虫害情报库-情报浏览.png", doc,
                                  "办公应用_病虫害情报库-情报浏览")
        except Exception as e:
            utils.g_logger.info("办公应用-病虫害情报库-情报浏览页异常")
            sheet.append(["办公应用-病虫害情报库-情报浏览", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用-病虫害情报库-情报浏览.png", doc,
                                  "办公应用-病虫害情报库-情报浏览")

        # 办公应用-病虫害情报库-情报统计
        try:
            self.driver.find_element(By.XPATH, "//ul//li[text()='情报统计 ']").click()
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@class='cell'][text()='情报标题']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开办公应用-病虫害情报库-情报统计页")
            utils.g_logger.info("办公应用-病虫害情报库-情报统计页显示正常")
            sheet.append(["办公应用-病虫害情报库-情报统计", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用_病虫害情报库-情报统计.png", doc,
                                  "办公应用_病虫害情报库-情报统计")
        except Exception as e:
            utils.g_logger.info("办公应用-病虫害情报库-情报统计页异常")
            sheet.append(["办公应用-病虫害情报库-情报统计", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用-病虫害情报库-情报统计.png", doc,
                                  "办公应用-病虫害情报库-情报统计")

        # 办公应用-APP新闻管理-新闻浏览
        try:
            self.driver.find_element(By.XPATH, "//span[@slot='title'][text()='APP新闻管理']").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//ul//li[text()='新闻浏览 ']").click()
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//label[text()='关键字']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开办公应用-APP新闻管理-新闻浏览页")
            utils.g_logger.info("办公应用-APP新闻管理-新闻浏览页显示正常")
            sheet.append(["办公应用-APP新闻管理-新闻浏览", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用_APP新闻管理_新闻浏览.png", doc,
                                  "办公应用_APP新闻管理_新闻浏览")
        except Exception as e:
            utils.g_logger.info("办公应用-APP新闻管理_新闻浏览页异常")
            sheet.append(["办公应用-APP新闻管理_新闻浏览", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用-APP新闻管理-新闻浏览.png", doc,
                                  "办公应用-APP新闻管理-新闻浏览")

        # 办公应用-APP新闻管理-新闻上传
        try:
            self.driver.find_element(By.XPATH, "//ul//li[text()='新闻上传 ']").click()
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//button/span[text()='新增']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开办公应用-APP新闻管理-新闻上传页")
            utils.g_logger.info("办公应用-APP新闻管理-新闻上传页显示正常")
            sheet.append(["办公应用-APP新闻管理-新闻上传", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用_APP新闻管理_新闻上传.png", doc,
                                  "办公应用_APP新闻管理_新闻上传")
        except Exception as e:
            utils.g_logger.info("办公应用-APP新闻管理_新闻上传页异常")
            sheet.append(["办公应用-APP新闻管理_新闻上传", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用-APP新闻管理-新闻上传.png", doc,
                                  "办公应用-APP新闻管理-新闻上传")

        # 办公应用-通知公告-公告管理
        try:
            self.driver.find_element(By.XPATH, "//span[@slot='title'][text()='通知公告']").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//ul//li[text()='公告管理 ']").click()
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//div[@class='search-container']//label[text()='公告类别']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开办公应用-通知公告-公告管理页")
            utils.g_logger.info("办公应用-通知公告-公告管理页显示正常")
            sheet.append(["办公应用-通知公告-公告管理", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用_通知公告-公告管理.png", doc,
                                  "办公应用_通知公告-公告管理")
        except Exception as e:
            utils.g_logger.info("办公应用-通知公告-公告管理页异常")
            sheet.append(["办公应用-通知公告-公告管理", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用-通知公告-公告管理.png", doc,
                                  "办公应用-通知公告-公告管理")

        # 办公应用-通知公告-公告查阅
        try:
            self.driver.find_element(By.XPATH, "//ul//li[text()='公告查阅 ']").click()
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//div[@class='search-container']//label[text()='公告类别']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开办公应用-通知公告-公告查阅页")
            utils.g_logger.info("办公应用-通知公告-公告查阅页显示正常")
            sheet.append(["办公应用-通知公告-公告查阅", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用_通知公告-公告查阅.png", doc,
                                  "办公应用_通知公告-公告查阅")
        except Exception as e:
            utils.g_logger.info("办公应用-通知公告-公告查阅页异常")
            sheet.append(["办公应用-通知公告-公告查阅", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/办公应用-通知公告-公告查阅.png", doc,
                                  "办公应用-通知公告-公告查阅")

        # 办公应用-知识与情报-工作平台
        try:
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//span[@slot='title'][text()='知识与情报']").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//ul//li[text()='工作平台 ']").click()
            time.sleep(10)
            elements = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, "//div[@class='title'][text()='病虫害知识库']/parent::div/following-sibling::div//img")))
            unittest.TestCase.assertTrue(elements is not None, "成功打开知识库-工作平台页")
            utils.g_logger.info("知识库-工作平台页显示正常")
            sheet.append(["知识库-工作平台", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_工作平台.png", doc, "知识库_工作平台")
        except Exception as e:
            utils.g_logger.info("知识库-工作平台页异常")
            sheet.append(["知识库-工作平台", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_工作平台.png", doc, "知识库_工作平台")

        # 知识库-病虫害知识库-知识浏览页
        try:
            self.driver.find_element(By.XPATH, "//span[text()='病虫害知识库']").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH,
                                     "//span[text()='病虫害知识库']/parent::div/following-sibling::ul//li[text()='知识浏览 ']").click()
            time.sleep(10)
            elements = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//div[@class='el-image img']/img")))
            unittest.TestCase.assertTrue(elements is not None, "成功打开知识库-病虫害知识库-知识浏览页，且指定元素存在")
            utils.g_logger.info("知识库-病虫害知识库-知识浏览页显示正常")
            sheet.append(["知识库-病虫害知识库-知识浏览", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_病虫害知识库_知识浏览.png", doc,
                                  "知识库_病虫害知识库_知识浏览")
        except Exception as e:
            utils.g_logger.info("知识库-病虫害知识库-知识浏览页异常")
            sheet.append(["知识库-病虫害知识库-知识浏览", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_病虫害知识库_知识浏览.png", doc,
                                  "知识库_病虫害知识库_知识浏览")

        # 知识库-病虫害知识库-知识维护页
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[text()='病虫害知识库']/parent::div/following-sibling::ul//li[text()='知识维护 ']").click()
            time.sleep(9)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@class='cell'][text()='修改时间']")))
            unittest.TestCase.assertTrue(element is not None,
                                         "成功打开知识库-病虫害知识库-知识维护页，且指定元素存在")
            utils.g_logger.info("知识库-病虫害知识库-知识维护页显示正常")
            sheet.append(["知识库-病虫害知识库-知识维护", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_病虫害知识库_知识维护.png", doc,
                                  "知识库_病虫害知识库_知识维护")
        except Exception as e:
            utils.g_logger.info("知识库-病虫害知识库-知识维护页异常")
            sheet.append(["知识库-病虫害知识库-知识维护", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_病虫害知识库_知识维护.png", doc,
                                  "知识库_病虫害知识库_知识维护")

        # 植保知识库-知识浏览页
        try:
            self.driver.find_element(By.XPATH, "//span[text()='植保知识库']").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH,
                                     "//span[text()='植保知识库']/parent::div/following-sibling::ul//li[text()='知识浏览 ']").click()
            time.sleep(8)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@role='radiogroup']//span[text()='病虫测报']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开知识库-植保知识库-知识浏览页，且指定元素存在")
            utils.g_logger.info("知识库-植保知识库-知识浏览页显示正常")
            sheet.append(["知识库-植保知识库-知识浏览", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_植保知识库_知识浏览.png", doc,
                                  "知识库_植保知识库_知识浏览")
        except Exception as e:
            utils.g_logger.info("知识库-植保知识库-知识浏览页异常")
            sheet.append(["知识库-植保知识库-知识浏览", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_植保知识库_知识浏览.png", doc,
                                  "知识库_植保知识库_知识浏览")

        # 植保知识库-知识审核页
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[text()='植保知识库']/parent::div/following-sibling::ul//li[text()='知识审核 ']").click()
            time.sleep(8)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@class='cell'][text()='发布时间']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开知识库-植保知识库-知识审核页，且指定元素存在")
            utils.g_logger.info("知识库-植保知识库-知识审核页显示正常")
            sheet.append(["知识库-植保知识库-知识审核", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_植保知识库_知识审核.png", doc,
                                  "知识库_植保知识库_知识审核")
        except Exception as e:
            utils.g_logger.info("知识库-植保知识库-知识审核页异常")
            sheet.append(["知识库-植保知识库-知识审核", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_植保知识库_知识审核.png", doc,
                                  "知识库_植保知识库_知识审核")

        # 植知识库-知识上传页
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[text()='植保知识库']/parent::div/following-sibling::ul//li[text()='知识上传 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@class='cell'][text()='发布时间']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开知识库-植保知识库-知识上传页")
            utils.g_logger.info("知识库-植保知识库-知识上传页显示正常")
            sheet.append(["知识库-植保知识库-知识上传", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_植保知识库_知识上传.png", doc,
                                  "知识库_植保知识库_知识上传")
        except Exception as e:
            utils.g_logger.info("知识库-植保知识库-知识上传页异常")
            sheet.append(["知识库-植保知识库-知识上传", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_植保知识库_知识上传.png", doc,
                                  "知识库_植保知识库_知识上传")

        # 资料库页
        try:
            self.driver.find_element(By.XPATH, "//li[text()='资料库 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@class='cell'][text()='目录']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开知识库-资料库页，且指定元素存在")
            utils.g_logger.info("知识库-资料库页显示正常")
            sheet.append(["知识库-资料库", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_资料库.png", doc,
                                  "知识库_资料库")
        except Exception as e:
            utils.g_logger.info("知识库-资料库页异常")
            sheet.append(["知识库-资料库", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/知识库_资料库.png", doc,
                                  "知识库_资料库")

        # 作物知识库-知识浏览
        try:
            self.driver.find_element(By.XPATH, "//span[text()='作物知识库']").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH,
                                     "//span[text()='作物知识库']/parent::div/following-sibling::ul//li[text()='知识浏览 ']").click()
            time.sleep(8)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//label[text()='作物类型']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开知识库-作物知识库-知识浏览页")
            utils.g_logger.info("知识库-作物知识库-知识浏览页显示正常")
            sheet.append(["知识库-作物知识库-知识浏览", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/作物知识库_知识浏览.png", doc,
                                  "作物知识库_知识浏览")
        except Exception as e:
            utils.g_logger.info("知识库-作物知识库-知识浏览页异常")
            sheet.append(["知识库-作物知识库-知识浏览", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/作物知识库_知识浏览.png", doc,
                                  "作物知识库_知识浏览")

        # 作物知识库-知识维护
        try:
            self.driver.find_element(By.XPATH,
                                     "//span[text()='作物知识库']/parent::div/following-sibling::ul//li[text()='知识维护 ']").click()
            time.sleep(10)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@class='cell'][text()='中文名']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开知识库-作物知识库-知识维护页")
            utils.g_logger.info("知识库-作物知识库-知识维护页显示正常")
            sheet.append(["知识库-作物知识库-知识维护", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/作物知识库_知识维护.png", doc,
                                  "作物知识库_知识维护")
        except Exception as e:
            utils.g_logger.info("知识库-作物知识库-知识维护页异常")
            sheet.append(["知识库-作物知识库-知识维护", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/作物知识库_知识维护.png", doc,
                                  "作物知识库_知识维护")

        # 关闭所有页面
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//div[@class='close-con']").click()
            time.sleep(2)
            element = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, "//ul[@class='el-dropdown-menu el-popper']/li[text()='关闭所有']")))
            element.click()
            utils.g_logger.info("成功关闭所有页面")
        except Exception:
            utils.g_logger.info("关闭所有页面失败")

    @utils.retry(MAX_TRIES)
    def test_shipinhuiyi(self):
        # 视频会议-会议管理
        try:
            time.sleep(5)
            self.driver.find_element(By.XPATH,
                                     "//div[@class='navigation-list']/div[@class='navigation-item']/span[text()=' 视频会议']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//button/span[text()='新增']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开视频会议-会议管理页")
            utils.g_logger.info("视频会议-会议管理页显示正常")
            sheet.append(["视频会议-会议管理", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/视频会议_会议管理.png", doc,
                                  "视频会议_会议管理")
        except Exception as e:
            utils.g_logger.info("视频会议-会议管理页异常")
            sheet.append(["视频会议-会议管理", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/视频会议-会议管理.png", doc,
                                  "视频会议-会议管理")

        # 视频会议-我的会议
        try:
            self.driver.find_element(By.XPATH,
                                     "//ul//span[@style='margin-left: 20px;'][text()='我的会议']").click()
            time.sleep(3)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//div[@class='search-container']//label[text()='会议名称']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开视频会议-我的会议页")
            utils.g_logger.info("视频会议-我的会议页显示正常")
            sheet.append(["视频会议-我的会议", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/视频会议_我的会议.png", doc,
                                  "视频会议_我的会议")
        except Exception as e:
            utils.g_logger.info("视频会议-我的会议页异常")
            sheet.append(["视频会议-我的会议", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/视频会议-我的会议.png", doc,
                                  "视频会议-我的会议")

        # 关闭所有页面
        try:
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//div[@class='close-con']").click()
            time.sleep(2)
            element = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, "//ul[@class='el-dropdown-menu el-popper']/li[text()='关闭所有']")))
            element.click()
            utils.g_logger.info("成功关闭所有页面")
        except Exception:
            utils.g_logger.info("关闭所有页面失败")

    def test_zhihuidiaodu(self):
        # 指挥调度-病虫测防
        try:
            time.sleep(5)
            self.driver.find_element(By.XPATH,
                                     "//div[@class='command-card']//div[text()='指挥调度']").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH,
                                     "//div[@class='name-box']/div[text()='病虫测防']").click()
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//div[@class='buttton-center bigTitle'][text()='病虫测防信息调度']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开指挥调度-病虫测防大屏")
            utils.g_logger.info("指挥调度-病虫测防大屏显示正常")
            sheet.append(["指挥调度-病虫测防大屏", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/指挥调度_病虫测防大屏.png", doc,
                                  "指挥调度_病虫测防大屏")
        except Exception as e:
            utils.g_logger.info("指挥调度-病虫测防大屏异常")
            sheet.append(["指挥调度-病虫测防大屏", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/指挥调度-病虫测防大屏.png", doc,
                                  "指挥调度-病虫测防大屏")

        # 指挥调度-植物检疫
        try:
            self.driver.find_element(By.XPATH,
                                     "//div[@class='el-form-item__content']/button//i[contains(@class,'home')]").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH,
                                     "//div[@class='name-box']/div[text()='植物检疫']").click()
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//div[@class='buttton-center bigTitle'][text()='植物检疫信息调度']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开指挥调度-植物检疫大屏")
            utils.g_logger.info("指挥调度-植物检疫大屏显示正常")
            sheet.append(["指挥调度-植物检疫大屏", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/指挥调度_植物检疫大屏.png", doc,
                                  "指挥调度_植物检疫大屏")
        except Exception as e:
            utils.g_logger.info("指挥调度-植物检疫大屏异常")
            sheet.append(["指挥调度-植物检疫大屏", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/指挥调度-植物检疫大屏.png", doc,
                                  "指挥调度-植物检疫大屏")

        # 指挥调度-物联网
        try:
            self.driver.find_element(By.XPATH,
                                     "//div[@class='el-form-item__content']/button//i[contains(@class,'home')]").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH,
                                     "//div[@class='name-box']/div[text()='物联网']").click()
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//div[@class='buttton-center bigTitle'][text()='物联网数据展示平台']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开指挥调度-物联网大屏")
            utils.g_logger.info("指挥调度-物联网大屏显示正常")
            sheet.append(["指挥调度-物联网大屏", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/指挥调度_物联网大屏.png", doc,
                                  "指挥调度_物联网大屏")
        except Exception as e:
            utils.g_logger.info("指挥调度-物联网大屏异常")
            sheet.append(["指挥调度-物联网大屏", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/指挥调度-物联网大屏.png", doc,
                                  "指挥调度-物联网大屏")

        # 指挥调度-植保体系
        try:
            self.driver.find_element(By.XPATH,
                                     "//div[@class='el-form-item__content']/button//i[contains(@class,'home')]").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH,
                                     "//div[@class='name-box']/div[text()='植保体系']").click()
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//div[text()='植保体系数据展示平台']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开指挥调度-植保体系大屏")
            utils.g_logger.info("指挥调度-植保体系大屏显示正常")
            sheet.append(["指挥调度-植保体系大屏", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/指挥调度_植保体系大屏.png", doc,
                                  "指挥调度_植保体系大屏")
        except Exception as e:
            utils.g_logger.info("指挥调度-植保体系大屏异常")
            sheet.append(["指挥调度-植保体系大屏", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/指挥调度-植保体系大屏.png", doc,
                                  "指挥调度-植保体系大屏")

        # 指挥调度-绿色示范区
        try:
            self.driver.find_element(By.XPATH,
                                     "//div[@class='el-form-item__content']/button//i[contains(@class,'home')]").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH,
                                     "//div[@class='name-box']/div[text()='绿色示范区']").click()
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//div[text()='湖北省绿色防控示范区']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开指挥调度-绿色示范区大屏")
            utils.g_logger.info("指挥调度-绿色示范区大屏显示正常")
            sheet.append(["指挥调度-绿色示范区大屏", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/指挥调度_绿色示范区大屏.png", doc,
                                  "指挥调度_绿色示范区大屏")
        except Exception as e:
            utils.g_logger.info("指挥调度-绿色示范区大屏异常")
            sheet.append(["指挥调度-绿色示范区大屏", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/指挥调度-绿色示范区大屏.png", doc,
                                  "指挥调度-绿色示范区大屏")

        # 指挥调度-农药械
        try:
            self.driver.find_element(By.XPATH,
                                     "//div[@class='el-form-item__content']/button//i[contains(@class,'home')]").click()
            time.sleep(3)
            self.driver.find_element(By.XPATH,
                                     "//div[@class='name-box']/div[text()='农药械']").click()
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//div[text()='农药械']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开指挥调度-绿色示范区大屏")
            self.driver.find_element(By.XPATH,
                                     "//div[@class='el-form-item__content']/button//i[contains(@class,'back')]").click()
            utils.g_logger.info("指挥调度-绿色示范区大屏显示正常")
            sheet.append(["指挥调度-绿色示范区大屏", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/指挥调度_绿色示范区大屏.png", doc,
                                  "指挥调度_绿色示范区大屏")
        except Exception as e:
            utils.g_logger.info("指挥调度-绿色示范区大屏异常")
            sheet.append(["指挥调度-绿色示范区大屏", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/指挥调度-绿色示范区大屏.png", doc,
                                  "指挥调度-绿色示范区大屏")

        # 昆虫雷达
        try:
            self.driver.find_element(By.XPATH, "//i[@class='iconfont icon-insectpicture']").click()
            time.sleep(3)
            # 获取所有窗口的句柄
            windows = self.driver.window_handles
            # 切换到新窗口，通常新窗口是最后一个句柄
            self.driver.switch_to.window(windows[-1])
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//*[text()='昆虫雷达']")))
            unittest.TestCase.assertTrue(element is not None, "成功打开昆虫雷达大屏")
            utils.g_logger.info("昆虫雷达大屏显示正常")
            sheet.append(["昆虫雷达大屏", "正常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/昆虫雷达大屏.png", doc,
                                  "昆虫雷达大屏")
            self.driver.close()
            self.driver.switch_to.window(windows[0])
        except Exception as e:
            utils.g_logger.info("昆虫雷达大屏异常")
            sheet.append(["昆虫雷达大屏", "异常"])
            utils.page_screenshot(self.driver, "outputs/imagefiles/昆虫雷达大屏.png", doc,
                                  "昆虫雷达大屏")

    def export_excel(self):
        # 退出浏览器
        self.driver.quit()
        # 调整列宽
        for col in sheet.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 20)
            sheet.column_dimensions[column].width = adjusted_width
        wb.save('outputs/省级功能巡检.xlsx')
        doc.save('outputs/湖北截图.docx')
        utils.g_logger.info("省级系统功能巡检结束，请检查输出的巡检文件内容。")


if __name__ == '__main__':
    p = PPMSHB()
    # 广西省级系统页面巡检
    p.test_shouye()
    p.test_jianceyubao()
    p.test_zhibaotixi()
    # 植物检疫只有使用admin账号登录时才起作用
    p.test_zhiwujianyi()
    p.test_bingchongfangzhi()
    p.test_nongyaoxie()
    p.test_wulianwang()
    p.test_zhibaotongji()
    p.test_bingchongzhuanti()
    p.test_bangongyingyong()
    p.test_shipinhuiyi()
    p.test_zhihuidiaodu()
    p.test_shuzituku()
    # 将所有巡检结果导出excel文件
    p.export_excel()
