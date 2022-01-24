# 这是一个示例 Python 脚本。

# 按 Shift+F10 执行或将其替换为您的代码。
# 按 双击 Shift 在所有地方搜索类、文件、工具窗口、操作和设置。

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import xlwt

book = xlwt.Workbook(encoding="utf-8", style_compression=0)
sheet = book.add_sheet("起点月票榜", cell_overwrite_ok=True)
sheet.write(0, 0, "名称")
sheet.write(0, 1, "排名")
sheet.write(0, 2, "图片")
sheet.write(0, 3, "简介")
sheet.write(0, 4, "更新信息")
i = 1
chrome = webdriver.Chrome()
WAIT = WebDriverWait(chrome, 1000)
chrome.get("https://www.qidian.com/")

def driver():

    link = chrome.find_elements(By.CLASS_NAME, "nav-li")
    link[1].click()

    bangdan = chrome.find_element(By.CLASS_NAME, "list_type_detective")
    yuepiao1 = bangdan.find_element(By.TAG_NAME, "li")
    yuepiao = yuepiao1.find_element(By.TAG_NAME, "a")
    yuepiao.click()
    get_source(chrome)

    for i in range (2, 5):
        next_page = chrome.find_element(By.CLASS_NAME, "lbf-pagination-next")
        next_page.click()
        all_l = chrome.window_handles
        chrome.switch_to.window(all_l[-1])
        get_source(chrome)
    next_page = chrome.find_element(By.LINK_TEXT, "5")
    next_page.click()
    all_l = chrome.window_handles
    chrome.switch_to.window(all_l[-1])
    get_source(chrome)
    return


def save_excel(soup):
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

def get_source(driver):
    html = driver.page_source
    soup = BeautifulSoup(html, "lxml")
    save_excel(soup)

def main():
    driver()

if __name__ == '__main__':
    main()


book.save("E:/BookRank/起点月票榜前100.xlsx")