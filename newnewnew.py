# 请求模块
import requests
# 处理文件路径模块
import os
# 导入正则表达式模块
import re
# 导入解码模块方法
from urllib.parse import unquote
# 导入json模块
import json
# 导入格式化输出模块
from pprint import pprint
# 导入 pandas 模块
import pandas as pd
# 导入 openpyxl 模块
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
# 实例化浏览器对象
# from DrissionPage import ChromiumPage
from selenium import webdriver
from selenium.webdriver.chrome import options
# 模拟请求头
from url_headers1 import headers, url_list1, url_list2, url_list3, url_list4
import speech_recognition as sr
from pydub import AudioSegment

def get_video(response):
    try:
        info_json = response.json()
        # 提取视频信息所在列表
        lis = info_json['aweme_list']

        # 指定保存路径
        save_path = os.path.join('excle', 'douyin_video_info.xlsx')

        # 确保目录存在
        os.makedirs(os.path.dirname(save_path), exist_ok=True)

        # 创建一个空的 DataFrame 用于存储所有视频信息
        all_data = []

        for index in lis:
            video_id = index['aweme_id']  # 提取视频ID

            # 模拟浏览器发送请求
            url = f'https://www.douyin.com/user/MS4wLjABAAAAz6JIUcQ-QVpqux6m09KOMI-GIB_fh5SSJcvhL1l8n41x2NCLGvzkOXD3zm6PVys8?from_tab_name=main&modal_id={video_id}'

            # 发送请求
            response = requests.get(url=url, headers=headers)
            response.raise_for_status()  # 检查请求是否成功

            # 获取响应的文本数据
            html = response.text

            # 提取视频相关信息(经过编码的数据)
            info = re.findall('<script id="RENDER_DATA" type="application/json">(.*?)</script>', html, re.S)
            if not info:
                print(f"未找到视频信息: {video_id}")
                continue

            # 解码
            json_str = unquote(info[0])

            # 将字符串转换为字典
            json_data = json.loads(json_str)

            # 提取视频链接
            video_url = json_data['app']['videoDetail']['video']['bitRateList'][0]['playAddr'][1]['src']
            # 提取音频链接
            audio_url = json_data['app']['videoDetail']['music']['playUrl']['uri']
            # 账号
            nickname = json_data['app']['videoDetail']['authorInfo']['nickname']
            # 提取视频标题
            title = json_data['app']['videoDetail']['desc']
            # 文本
            ocrContent = json_data['app']['videoDetail']['seoInfo']['ocrContent']
            # 提取红心个数
            like_count = json_data['app']['videoDetail']['stats']['diggCount']
            # 提取转发量
            share_count = json_data['app']['videoDetail']['stats']['shareCount']

            # 将数据保存到列表中
            all_data.append({
                '账号昵称': nickname,
                '主题': title,
                '原文案': ocrContent,
                '点赞': like_count,
                '转发量': share_count,
                '视频链接': video_url,
                '音频链接': audio_url,
            })

        # 创建 DataFrame
        df = pd.DataFrame(all_data)

        # 如果 Excel 文件已经存在，追加数据；否则创建新文件
        if os.path.exists(save_path):
            with pd.ExcelWriter(save_path, engine='openpyxl', mode='a') as writer:
                df.to_excel(writer, index=False, header=False)  # 追加数据，不写入表头
        else:
            df.to_excel(save_path, index=False)  # 创建新文件

        # 使用 openpyxl 加载刚刚创建的 Excel 文件
        wb = load_workbook(save_path)
        ws = wb.active

        # 设置列标题颜色为黄色
        yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        for cell in ws[1]:  # ws[1] 是第一行，即列标题
            cell.fill = yellow_fill

        # 自适应列宽
        for column in ws.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)  # 加2是为了给点额外的空间
            ws.column_dimensions[column[0].column_letter].width = adjusted_width

        # 保存修改后的 Excel 文件
        wb.save(save_path)
        print(f"数据已保存到 {save_path}")

    except Exception as e:
        print(f"处理视频信息时发生错误: {e}")


# 实例化浏览器对象
options = options.Options()
options.add_argument("--headless")
driver = webdriver.Chrome(options=options)
# dp = ChromiumPage()
# 监听数据包
# dp.listen.start('web/aweme/post/')
# 打开网页
driver.get(url="https://www.douyin.com/user/MS4wLjABAAAAz6JIUcQ-QVpqux6m09KOMI-GIB_fh5SSJcvhL1l8n41x2NCLGvzkOXD3zm6PVys8")
# dp.get("https://www.douyin.com/user/MS4wLjABAAAAz6JIUcQ-QVpqux6m09KOMI-GIB_fh5SSJcvhL1l8n41x2NCLGvzkOXD3zm6PVys8")
#
# 构建翻页
url_list = [url_list1, url_list2, url_list3, url_list4]

for page in range(4):
    try:
        print(f'正在爬取第{page + 1}页数据')
        response = requests.get(url=url_list[page], headers=headers)
        response.raise_for_status()  # 检查请求是否成功
        get_video(response)
    except Exception as e:
        print(f"爬取第{page + 1}页数据时发生错误: {e}")








