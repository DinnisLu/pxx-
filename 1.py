import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from functools import partial
import threading
import logging
import platform  # 导入platform模块
import os  # 导入os模块
import openpyxl
import pandas as pd
import selenium.webdriver as webdriver
from pandas import DataFrame
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image as PILImage
import sys
import datetime
import time  # 导入time模块

def check_license():
    license_key = "Kami6688"
    user_input = simpledialog.askstring("License Key", "请输入密匙：")
    if user_input != license_key:
        messagebox.showerror("错误", "密匙错误，程序将退出。")
        sys.exit()

def show_help():
    help_text = """
    使用帮助:

    1. 运行程序后，请先输入正确的密匙。
    2. 选择要处理的 Excel 文件和输出位置。
    3. 点击“执行处理”按钮进行处理。
    4. 处理完成后，程序会生成带有截图的 Excel 文件。

    注意：确保你的 Python 环境中已安装必要的库，如 openpyxl、pandas、selenium、Pillow 等。

    如果需要更详细的帮助，请联系开发者勇哥和帅牛。
    """
    messagebox.showinfo("使用帮助", help_text)

def main():
    check_license()
    run_gui()

def process_excel(excel_file_path, output_path):
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    store_name = '拼多多咕噜季家居官方旗舰店'
    excel_file = excel_file_path
    sheet_name = "售后信息"
    order_id = "订单编号"
    sh_id = "售后编号"
    save_df = DataFrame(columns=[""])
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    sh_id_list = df[sh_id].tolist()

    pdd_chrome_options = Options()

    # 获取操作系统类型
    system_type = platform.system()
    if system_type == "Windows":
        # 如果是Windows系统
        pdd_chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9223")
    else:
        # 如果是其他系统，如Linux或macOS
        pdd_chrome_options.add_argument("--remote-debugging-port=9223")
        pdd_chrome_options.add_argument("--user-data-dir=./chrome_data")

    driver = webdriver.Chrome(options=pdd_chrome_options)
    new_df = DataFrame(columns=["日期", "店铺名", "订单号", "快递公司", "快递单号", "金额", "丢件原因", "退款截图", "备注"])
    sh_count = len(df)
    sh_index = 0

    while sh_index < sh_count:
        sh_id = df.iloc[sh_index]['售后编号']
        order_id = df.iloc[sh_index]['订单编号']

        url_str = 'https://mms.pinduoduo.com/aftersales-ssr/detail?id={}&orderSn={}'.format(sh_id, order_id, )
        driver.get('https://mms.pinduoduo.com/aftersales-ssr/detail?id={}&orderSn={}'.format(sh_id, order_id))
        time.sleep(1)  # 等待3秒钟，让界面加载完成
        driver.save_screenshot('E:/自动化表格处理/售后_急速退款/截图/{}.png'.format(sh_id))
        logging.info("售后单{}截图处理完毕...".format(sh_id))

        express_info = driver.find_element(by=By.ID, value='detail-express-box')
        express_all_button = express_info.find_element(by=By.CSS_SELECTOR,
                                                       value='div.mui-steps-item.mui-steps-item-wait')
        express_all_button.click()
        express_company = ''
        if '湖北武汉东西湖区径河公司' in express_info.text:
            logging.error("售后单{},武汉硚口申通".format(sh_id))
            express_company = '武汉硚口申通'
        elif '湖北武汉东西湖区径河公司' in express_info.text:
            logging.info("湖北武汉东西湖区径河公司")
            express_company = '武汉泾河申通'
        else:
            logging.error("售后单{},无法匹配快递网点".format(sh_id))

        new_df.loc[sh_index, '日期'] = datetime.date.today()
        new_df.loc[sh_index, '店铺名'] = '拼多多咕噜季家居官方旗舰店'
        new_df.loc[sh_index, '订单号'] = df.iloc[sh_index]['订单编号']
        new_df.loc[sh_index, '快递公司'] = express_company
        new_df.loc[sh_index, '快递单号'] = df.iloc[sh_index]['发货运单号']
        new_df.loc[sh_index, '金额'] = df.iloc[sh_index]['退款金额']
        new_df.loc[sh_index, '丢件原因'] = ''
        new_df.loc[sh_index, '退款截图'] = ''
        new_df.loc[sh_index, '备注'] = ''

        sh_index = sh_index + 1

    output_filename = output_path + '/{}_{}_截图13.xlsx'.format(store_name, datetime.date.today())

    new_df.to_excel(output_filename, sheet_name='售后信息')

    sh_index = 0
    while sh_index < sh_count:
        workbook = openpyxl.load_workbook(output_filename)
        worksheet = workbook['售后信息']
        image_file_name = 'E:/自动化表格处理/售后_急速退款/截图/{}.png'.format(df.iloc[sh_index]['售后编号'])
        img = OpenpyxlImage(image_file_name)
        image_loc = 'I{}'.format(sh_index + 1)
        worksheet.add_image(img, image_loc)

        workbook.save(output_filename)
        workbook.close()

        sh_index = sh_index + 1

    driver.quit()
    messagebox.showinfo("提示", "运行结束")

def select_file(entry):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
    entry.delete(0, tk.END)
    entry.insert(0, file_path)

def select_folder(entry):
    folder_path = filedialog.askdirectory()
    entry.delete(0, tk.END)
    entry.insert(0, folder_path)

def run_process(entry_excel, entry_output):
    excel_file_path = entry_excel.get()
    output_path = entry_output.get()

    if not excel_file_path or not output_path:
        messagebox.showwarning("警告", "请选择Excel文件和输出位置")
        return

    thread = threading.Thread(target=process_excel, args=(excel_file_path, output_path))
    thread.start()

def run_gui():
    root = tk.Tk()
    root.title("大菠萝拼夕夕Excel处理工具")

    tk.Label(root, text="Excel文件路径:").grid(row=0, column=0)
    entry_excel = tk.Entry(root, width=50)
    entry_excel.grid(row=0, column=1)
    btn_select_excel = tk.Button(root, text="选择文件", command=partial(select_file, entry_excel))
    btn_select_excel.grid(row=0, column=2)

    tk.Label(root, text="输出位置:").grid(row=1, column=0)
    entry_output = tk.Entry(root, width=50)
    entry_output.grid(row=1, column=1)
    btn_select_output = tk.Button(root, text="选择文件夹", command=partial(select_folder, entry_output))
    btn_select_output.grid(row=1, column=2)

    btn_run = tk.Button(root, text="执行处理", command=partial(run_process, entry_excel, entry_output))
    btn_run.grid(row=2, column=0, columnspan=3, pady=10)

    btn_help = tk.Button(root, text="使用帮助", command=show_help)
    btn_help.grid(row=3, column=0, columnspan=3, pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()