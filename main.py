import requests
import re
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
import os


# 初始化Excel
def init_excel():
    """初始化Excel，创建一个名为'中药信息'的Excel文件，并创建一个名为'中药信息'的Sheet
    第一行写入药名，史载于，别名，性味归经，药材简介，用法用量，注意事项
    """
    # 检测是否存在中药信息.xlsx文件，存在则跳过
    if os.path.exists("中药信息.xlsx"):
        print("中药信息.xlsx文件已存在，跳过初始化")
        return
    # 创建一个名为'中药信息'的Excel文件
    workbook = openpyxl.Workbook()
    # 创建一个名为'中药信息'的Sheet
    sheet = workbook.create_sheet("中药信息", 0)
    # 第一行写入药名，史载于，别名，性味归经，药材简介，用法用量，注意事项
    sheet.append(
        ["药名", "史载于", "别名", "性味归经", "药材简介", "用法用量", "注意事项"]
    )
    # 保存Excel文件
    workbook.save("中药信息.xlsx")
    print("中药信息.xlsx文件初始化完成")


# 添加内容到Excel
def add_content_to_excel(content):
    """添加内容到Excel"""
    workbook = openpyxl.load_workbook("中药信息.xlsx")
    sheet = workbook["中药信息"]  # 直接通过sheet名称获取
    sheet.append(content)
    workbook.save("中药信息.xlsx")


def fetch_webpage(url):
    """获取网页内容并设置编码"""
    response = requests.get(url)
    response.encoding = "utf-8"
    return response.text


def save_response(content, filename="response.html"):
    """保存响应内容到文件"""
    with open(filename, "w", encoding="utf-8") as f:
        f.write(content)


def extract_medicine_info(text_content):
    """提取药品ID和名称"""
    pattern = re.compile(
        r'<a href="/traditionaldetails\?id=(\d+)" target="_blank" data-v-bce4c468>(.*?)</a>'
    )
    return pattern.findall(text_content.replace("\n", ""))


def process_matches(matches):
    """处理匹配结果并返回分离的ID和药名列表"""
    ids = [match[0].strip() for match in matches]
    medicine_names = [match[1].strip() for match in matches]
    return ids, medicine_names


def print_medicine_info(matches):
    """打印药品信息"""
    for id_num, medicine_name in matches:
        print(f"ID: {id_num}, 药名: {medicine_name.strip()}")


def get_medicine_id_map():
    """获取药品和id映射"""
    url = "https://www.zhiyuanzhongyi.com/traditional"
    # 获取内容
    content = fetch_webpage(url)
    # 提取药品信息
    matches = extract_medicine_info(content)
    # 获取分离的ID和药名列表
    ids, medicine_names = process_matches(matches)
    return ids, medicine_names


def process_medicine_info(ids, medicine_names):
    """处理药品信息"""
    base_url = "https://www.zhiyuanzhongyi.com/traditionaldetails?id="
    print("开始处理药品信息")
    for id, medicine_name in zip(ids, medicine_names):
        url = base_url + id
        print(f"开始处理药品信息：{medicine_name}")
        response = fetch_webpage(url)
        # 提取药品名称
        pattern = re.compile(r'<div class="right_msg" data-v-0bab2978>(.*?)</div>')
        medicine_name = pattern.search(response)
        if medicine_name:
            medicine_name = medicine_name.group(1)


if __name__ == "__main__":
    print("-" * 50)
    print("开始执行，Writed by W1ndys，https://github.com/W1ndys")
    print("-" * 50)
    init_excel()
    print("-" * 50)
    ids, medicine_names = get_medicine_id_map()
    print(f"获取药品ID和名称映射完成，共获取到{len(ids)}个药品")
    print("-" * 50)
    process_medicine_info(ids, medicine_names)
    print("-" * 50)
