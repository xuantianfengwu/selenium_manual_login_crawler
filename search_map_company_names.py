from complete_example import AdvancedCrawler

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
# from selenium.webdriver.chrome.service import Service as ChromeService
# from selenium.webdriver.edge.service import Service as EdgeService
# from selenium.webdriver.chrome.options import Options as ChromeOptions
# from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import json
import os
import sys
import time
import logging
import pandas as pd
from openpyxl import Workbook, load_workbook
from bs4 import BeautifulSoup

# 导入tkinter用于文件选择对话框
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

logger = logging.getLogger(__name__)

def search_map_company_names():
    """
        搜索并映射公司名称
    """
    # --------------------------
    # 1. 文件选择阶段
    # --------------------------
    logger.info("===== 开始演示完整爬虫流程 =====")
    logger.info("\n----- 文件选择阶段 -----")
    # 创建tkinter窗口（隐藏）
    root = tk.Tk()
    root.withdraw()
    # 弹出文件选择对话框
    input_file_path = filedialog.askopenfilename(
        title="选择公司列表文件",
        filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
    )
    # 检查用户是否选择了文件
    if not input_file_path:
        logger.error("用户未选择输入文件，程序终止")
        messagebox.showerror("错误", "未选择输入文件，程序终止")
        return
        # 验证文件是否存在且为.xlsx格式
    if not os.path.exists(input_file_path):
        logger.error(f"所选文件不存在: {input_file_path}")
        messagebox.showerror("错误", f"所选文件不存在: {input_file_path}")
        return
    if not input_file_path.endswith('.xlsx'):
        logger.error(f"所选文件不是有效的Excel文件(.xlsx): {input_file_path}")
        messagebox.showerror("错误", f"所选文件不是有效的Excel文件(.xlsx): {input_file_path}")
        return

    # 1. 读取公司列表
    logger.info(f"已选择输入文件: {input_file_path}")
    print(f"已选择输入文件: {input_file_path}")
    df = pd.read_excel(input_file_path)
    fund_names = df['基金名称'].dropna().tolist()
    print(f'基金名称共 {len(fund_names)} 家')

    # 网站URL示例（使用百度作为演示）
    target_url = "https://xunkebao.baidu.com"

    # Step1：首次访问并保存Cookies
    logger.info("\n----- 第一阶段：首次访问并保存Cookies -----")
    crawler = AdvancedCrawler()
    # chrome_driver_path = './static/chromedriver-win64/chromedriver.exe'
    # if crawler.start_browser(chrome_driver_path=chrome_driver_path):
    crawler.start_browser()
    crawler.navigate_to(target_url)  # 1-1: 访问网站
    input("请在浏览器中执行任何需要的操作（如登录），完成后按Enter键继续...")  # 1-2: 提示用户可以手动登录（如果需要）

    # Step2：使用保存的Cookies重新访问
    new_crawler = crawler

    # Step3: 顺序搜索基金信息，并保存到output_file_path
    output_file_path = 'static/complete_example/map_company_names.xlsx'
    print(f'产出文件路径: {output_file_path}')
    if not os.path.exists(output_file_path):
        logger.info('未发现产出文件，将新建')
        wb = Workbook()
        ws_company_name = wb.active
        ws_company_name.title = "公司名称"
        wb.save(output_file_path)  # 保存工作簿
        existed_company_names = []
    else:
        logger.info('已发现产出文件，将直接读取')
        existed_company_names = set()
        for tab in ('公司名称',):
            tmp_df = pd.read_excel(output_file_path, sheet_name=tab, header=None)
            if len(tmp_df) > 0:
                existed_company_names = existed_company_names | set([v.split('|')[0] for v in tmp_df[0]])
        logger.info(f'已存在 {len(existed_company_names)} 家公司名称')

    # 3. 顺序遍历爬取
    for i, fund_name in enumerate(fund_names):
        if fund_name in existed_company_names:
            logger.info(f'基金 {fund_name} 的公司名称已存在，跳过')
            continue
        logger.info(f'开始处理基金 {i + 1}/{len(fund_names)}: {fund_name}')
        # 3-1. 搜索公司
        is_search = 0
        for search_fund_name in [f'{fund_name}投资管理', fund_name]:
            if is_search == 0:
                search_input = new_crawler.driver.find_element(By.CSS_SELECTOR,
                                                               "div.search-input-wrap > section > div > div > div > div > input")
                search_btn = new_crawler.driver.find_element(By.CSS_SELECTOR,
                                                             "div.search-input-wrap > section > div > button.el-button.el-button--primary.search-btn")
                search_input.clear()
                search_input.send_keys(search_fund_name)
                search_btn.click()
                # 3-2. 等待搜索结果加载
                time.sleep(3)

                search_res_num = new_crawler.driver.find_element(By.CSS_SELECTOR,
                                                                 'div.middle-bar > div.info > span:nth-child(1) > em')
                # 搜索结果为0，company_name=''
                if search_res_num.text.strip() == '0':
                    company_name = ''
                # 搜索结果不为0，选择第一条记录的名称作为公司名称
                else:
                    company_name_items = new_crawler.driver.find_elements(By.CSS_SELECTOR, 'h6.company-name')
                    # 网站有个BUG，比如搜索“Lightspeed Venture Partners光速全球”，会显示“网络开小差了，请稍后重试”，并保持公司数为上次结果，
                    # 因此要二次兜底判断下
                    if len(company_name_items) > 0:
                        company_name = company_name_items[0].text.strip()
                        is_search = 1
                    else:
                        company_name = ''

        # 3-5. 保存基金的公司名称到结果文件
        wb = load_workbook(output_file_path)
        ws = wb['公司名称']
        ws.append([f'{fund_name}|{company_name}'])
        wb.save(output_file_path)
        time.sleep(2)

if __name__ == "__main__":
    # 运行演示流程
    search_map_company_names()