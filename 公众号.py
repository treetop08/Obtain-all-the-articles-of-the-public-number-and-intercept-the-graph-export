# -*- coding: utf-8 -*-
import requests
import pdfkit
import time
import xlwt
from selenium import webdriver
from selenium.webdriver.chrome.options import Options


def get_image(url, pic_name):
    """
    #设置chrome开启的模式，headless就是无界面模式
    # 创建一个参数对象，用来控制chrome以无界面模式打开
    :param url:             获取获取网页的地址
    :param pic_name:        需要保存的文件名或路径＋文件名
    :return:
    """
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    # 创建浏览器对象
    # driver = webdriver.Chrome(executable_path='./chromedriver', chrome_options=chrome_options)
    driver = webdriver.Chrome(executable_path="D:/python/chromedriver_win32/chromedriver.exe",chrome_options=chrome_options)
    # 打开网页
    driver.get(url)
    # driver.maximize_window()
    # 加延时 防止未加载完就截图
    time.sleep(1)

    # 用js获取页面的宽高，如果有其他需要用js的部分也可以用这个方法
    width = driver.execute_script("return document.documentElement.scrollWidth")
    height = driver.execute_script("return document.documentElement.scrollHeight")

    # 获取页面宽度及其宽度
    print(width, height)

    # 将浏览器的宽高设置成刚刚获取的宽高
    driver.set_window_size(width, height)

    time.sleep(1)

    # 截图并关掉浏览器
    driver.get_screenshot_as_file(pic_name)

    driver.quit()


# 你输入的参数
url_str = 'http://www.cq.gov.cn'
pic_name = r'qwq.png'

get_image(url_str, pic_name)

# 做表
book = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = book.add_sheet('19级', cell_overwrite_ok=True)
sheet.write(0, 0, '名称')
sheet.write(0, 1, '链接')
n = 1

headers = {
    # "cookie": "appmsglist_action_3889613222=card; ua_id=Q1Dfu2THA6T9Qr1HAAAAAN_KYa5xTwNmiuqj1Mkl6PY=; wxuin=18828715020059; openid2ticket_opsnW57Rn-HfZYURLULM6KyG49U4=; mm_lang=zh_CN; pgv_info=ssid=s3846515381; pgv_pvid=2250932512; rewardsn=; wxtokenkey=777; uuid=892882700c3649c61d997e65c9c391b1; rand_info=CAESIJOel5DqVg0rCHkL+nlhk4tMuse1Fc3Cqo1nlKt7VANx; slave_bizuin=3889613222; data_bizuin=3889613222; bizuin=3889613222; data_ticket=wb7Ew50VibgcG8efcvxnV0iYnw6mmwWlhpvUu/YPr7AVfs1vK8xooK9Wo+7w0MaW; slave_sid=WnNVZV93OFJlZ0VXTF9QODQ5M1Y2c0hibTdwNjZQV2VuTHpsSVZhQVplN3NvbklPM0p0NGJ2NUU0bXpKNGlkRlA0a28zQ2s1WkhpeVdyaFhiYmg5dTZ6OWpLTFByMlQ3SjNqMzNnM3Y4RkdqWXVKNUpOQnI4ZElTaXpWZnRKbEZmUGgyTm52QkYzd2JQZE11; slave_user=gh_bd2fc8d28eb3; xid=a5c7612f529374b74deb4178e7ff4ca7",
    # "cookie": "appmsglist_action_3866430480=card; appmsglist_action_3890620645=card; "
    #           "ua_id=Jr1GAkXiuik4EkthAAAAAD1Hayq-3G3IFZ-Whto4vWE=; wxuin=25411726488246; mm_lang=zh_CN; RK=/rpEYayaR+; "
    #           "ptcz=f0453cb9e02a0937eff3a5e021f0b8aff99a6155e39e25c7d06075b9a89ddd4e; Qs_lvt_323937=1628263969; "
    #           "Qs_pv_323937=2294544851859949300; ts_uid=3881336320; o_cookie=1610795342; pac_uid=1_1610795342; "
    #           "eas_sid=u1L6d3F6i6g5n1M4o3V4X726V3; pgv_pvid=1800158520; LW_uid=o1L6L348q1y9E6X1F4u0j9n6H9; "
    #           "tvfe_boss_uuid=eff8e567ac8e9f31; LW_sid=q1B6d4p222h4O5k343t349a0T6; "
    #           "fqm_pvqid=fad43520-e3af-42cc-aa27-ae1a9f1f0dda; _ga=GA1.2.1367857272.1644401878; "
    #           "rand_info=CAESILU+XCClRzACjhubS5r7+RobNdIgNz03rasO2G+MW69a; slave_bizuin=3890620645; "
    #           "data_bizuin=3890620645; bizuin=3890620645; "
    #           "data_ticket=6HJlKNywdfhxz61aG6JrVEiFoXo/ehycCJfXWvW79mR3lAFq2i1HuHE+M/WRILDD; "
    #           "slave_sid"
    #           "=S25vUWhHZkx2M1JlZzRzeHdsWmxLWEdyNlpYU0pvTHVSUTcwcmFMM3daS3ZlTFNtQTVhZUN0WTJsaF9rVkZab29IQ1JMUlFFS04xZXFheG01Z3l5NGNGZ0ZZTjNYSXdVM3ozRGlVbFJzRW9ocEdiSUZVUWQycUFOdzlaTWVybktHTkpKemxiQVoyMnA4SEdU; slave_user=gh_40ca67b64574; xid=59ac67ff50e340c29d63257d6675f829; rewardsn=; wxtokenkey=777",
    # "cookie": "appmsglist_action_3866430480=card; appmsglist_action_3890620645=card; ua_id=Jr1GAkXiuik4EkthAAAAAD1Hayq-3G3IFZ-Whto4vWE=; wxuin=25411726488246; mm_lang=zh_CN; RK=/rpEYayaR+; ptcz=f0453cb9e02a0937eff3a5e021f0b8aff99a6155e39e25c7d06075b9a89ddd4e; Qs_lvt_323937=1628263969; Qs_pv_323937=2294544851859949300; ts_uid=3881336320; o_cookie=1610795342; pac_uid=1_1610795342; eas_sid=u1L6d3F6i6g5n1M4o3V4X726V3; pgv_pvid=1800158520; LW_uid=o1L6L348q1y9E6X1F4u0j9n6H9; tvfe_boss_uuid=eff8e567ac8e9f31; LW_sid=q1B6d4p222h4O5k343t349a0T6; fqm_pvqid=fad43520-e3af-42cc-aa27-ae1a9f1f0dda; _ga=GA1.2.1367857272.1644401878; uuid=73bb89e0d66f947cb3403dfc8087018c; rand_info=CAESIOnZer1VV2udvf/Zu7y76IjCmGI4hKH62XI7DYZLui/y; slave_bizuin=3866430480; data_bizuin=3866430480; bizuin=3866430480; data_ticket=GFHbaJLvM46uJVnE5YKWjioKlPggd+t1zfnpYsh+Bg9v7Uac5T1I6cnsaLHhrSW6; slave_sid=aE9BRXU5RlBqRnpnWHlRazNfbUR6bHZ0ZjZ1T1c5VTZtTHczX0dWUUpDRlFFQzRIcWdtNXFQamxPUlRtZnVfVmo4S3l4OEVUQXp4OER2MEcxQjFITjh6TkQyZ0p4VXF4VnVOdjJCZU8yQXRDMWVSdHk1Q0JPcHVtV1pIT3QybUUzeVI2NDBzZWFCN1lRR0pI; slave_user=gh_39c206e4bc2d; xid=10e721b79c034e49e91336afb7e125eb",
    "cookie":"appmsglist_action_3866430480=card; appmsglist_action_3890620645=card; "
             "ua_id=Jr1GAkXiuik4EkthAAAAAD1Hayq-3G3IFZ-Whto4vWE=; wxuin=25411726488246; mm_lang=zh_CN; RK=/rpEYayaR+; "
             "ptcz=f0453cb9e02a0937eff3a5e021f0b8aff99a6155e39e25c7d06075b9a89ddd4e; Qs_lvt_323937=1628263969; "
             "Qs_pv_323937=2294544851859949300; ts_uid=3881336320; o_cookie=1610795342; pac_uid=1_1610795342; "
             "eas_sid=u1L6d3F6i6g5n1M4o3V4X726V3; pgv_pvid=1800158520; LW_uid=o1L6L348q1y9E6X1F4u0j9n6H9; "
             "tvfe_boss_uuid=eff8e567ac8e9f31; LW_sid=q1B6d4p222h4O5k343t349a0T6; "
             "fqm_pvqid=fad43520-e3af-42cc-aa27-ae1a9f1f0dda; _ga=GA1.2.1367857272.1644401878; "
             "uuid=3ed86cd4a7ef7e8182e719b2261fc36c; rand_info=CAESIKgDpPbXLZm2X07WATmFpe5OAfEUmLEpYPEftZYhWZNF; "
             "slave_bizuin=3866430480; data_bizuin=3866430480; bizuin=3866430480; "
             "data_ticket=3Fj9DK1MO6r+GwP/cpIyuStW9p3o04KGVwyZ1cvhFXTkWUA7kGhTVXaec9WSTMQ7; "
             "slave_sid"
             "=SFdQc2dNczdJRDc1NFNYcE43dm1tanc1R3lkY3RvWXNZV3l1WFVkbjNrNDZUMDgzU0RsSHROY3hyOEVJWDZDdUhqZW5uRmNqelgwNER2a1pqamg5VjFrWnRwdmxVNk9udVZiUVU5MFNQazJieXhCSjdVNmp4S1VvRDNsa0FrZGxma1hIWmR6U05ZVzlqVXlD; slave_user=gh_39c206e4bc2d; xid=99ab6eb5abe6fd043a09a9e300933d56",
    # "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.116 Safari/537.36"
    # "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.51 Safari/537.36 Edg/99.0.1150.36 "
    "user-agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                 "Chrome/99.0.4844.51 Safari/537.36 Edg/99.0.1150.39 "
}
url = 'https://mp.weixin.qq.com/cgi-bin/appmsg'
fad = 'Mzg5MDYyMDY0NQ=='  # 爬不同公众号只需要更改fakeid


def page(num=8):  # 要请求的文章页数
    title = []
    link = []
    for i in range(num):
        data = {
            'action': 'list_ex',
            'begin': i * 5,  # 页数
            'count': '5',
            'fakeid': fad,
            'type': '9',
            'query': '',
            'token': '407828752',
            'lang': 'zh_CN',
            'f': 'json',
            'ajax': '1',
        }

        session = requests.Session()
        session.trust_env = False
        r = session.get(url, headers=headers, params=data)
        dic = r.json()

        for i in dic['app_msg_list']:  # 遍历dic['app_msg_list']中所有内容
            title.append(i['title'])  # 取 key键 为‘title’的 value值
            link.append(i['link'])  # 去 key键 为‘link’的 value值
    return title, link


if __name__ == '__main__':
    (tle, lik) = page(8)
    wk_path = r'D:\python\wkhtmltopdf\bin\wkhtmltopdf.exe'
    # config = pdfkit.configuration(wkhtmltopdf=wk_path)
    for x, y in zip(tle, lik):
        # pdfkit.from_url(y, 'D:/python项目NEW/东大软件2020级公众号/' + x + '.pdf',configuration=config)
        sheet.write(n, 0, x)
        sheet.write(n, 1, y)
        n = n+1
        get_image(y, x + '.png')

        print(x, y)

book.save(u'深挚吟公众号.xlsx')