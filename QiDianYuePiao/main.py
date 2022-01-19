# 这是一个示例 Python 脚本。

# 按 Shift+F10 执行或将其替换为您的代码。
# 按 双击 Shift 在所有地方搜索类、文件、工具窗口、操作和设置。


import requests
import bs4
from bs4 import BeautifulSoup
import xlwt

def request_qidian(url):
    # 访问得到起点对应网站的html界面
    r = requests.get(url)
    return r.text

book = xlwt.Workbook(encoding="utf-8", style_compression=0)
sheet = book.add_sheet("起点月票榜", cell_overwrite_ok=True)
sheet.write(0, 0, "名称")
sheet.write(0, 1, "排名")
sheet.write(0, 2, "图片")
sheet.write(0, 3, "简介")
sheet.write(0, 4, "更新信息")
i = 1

def save(soup):


    items = soup.find(id="book-img-text").find_all("li")
    global i
    for item in items:
        # for string in item.find("span").stripped_strings:
        rank = item.find("span").next_element.string
        img = item.find("img").get("src")
        name = item.find("h2").string
        info = item.find("p", class_= "intro").string
        update_info = item.find("p", class_= "update").a.string
        print("排名: "+ rank+ "  "+ " 图片: "+ img+ " 书名: "+ name+ " 简介: "+ info+ " 更新信息: "+ update_info)

        sheet.write(int(rank), 0, name)
        sheet.write(int(rank), 1, rank)
        sheet.write(int(rank), 2, img)
        sheet.write(int(rank), 3, info)
        sheet.write(int(rank), 4, update_info)
        i = i + 1

    return items


def main(page):
    url = "https://www.qidian.com/rank/yuepiao/year2022-month01-page"+ str(page)
    html = request_qidian(url)
    soup = BeautifulSoup(html, "lxml")
    save(soup)






# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':
    for i in range(1,6):
        main(i)

book.save("E:/BookRank/起点月票榜.xlsx")
# 访问 https://www.jetbrains.com/help/pycharm/ 获取 PyCharm 帮助
