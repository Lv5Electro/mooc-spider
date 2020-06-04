from selenium import webdriver
import time
import xlwt


test = '代码测试'

wb = xlwt.Workbook()
ws = wb.add_sheet('answer')
style = xlwt.XFStyle()
al = xlwt.Alignment()
# 垂直居中
al.vert = 0x01
style.alignment = al
# 设置宽度
ws.col(0).width = 256 * 50
# 选择题数量
num1 = 45
# 填空题数量
num2 = 10
# 判断题数量
num3 = 10

# 账号
unameId = "15073230912"
# 密码
passwordId = "548926520."
# url
login_url = 'http://mooc1.hnsyu.net/work/doHomeWorkNew?courseId=204033648&classId=14307514&workId=9309566&workAnswerId=17933921&reEdit=1&isdisplaytable=2&mooc=1&enc=5afaaf9625c85f480732d0a3ab166286&workSystem=0&cpi=68491570&standardEnc=009f4eb1819faea9a7eaa4ef2588dbd4'
option = webdriver.ChromeOptions()
driver = webdriver.Chrome(chrome_options=option)
driver.maximize_window()
driver.get(login_url)
driver.implicitly_wait(10)
elem = driver.find_element_by_id("unameId")
elem.send_keys(unameId)
time.sleep(1)
elem = driver.find_element_by_id("passwordId")
elem.send_keys(passwordId)
time.sleep(3)
elem = driver.find_element_by_class_name("zl_btn_right")
elem.click()
time.sleep(3)
parent_path = "//*[@id='ZyBottom']/div[@class='TiMu']"
parent_elements = driver.find_elements_by_xpath(parent_path)


# list转数组
def listToString(list):
    myStr = ''
    for item in list:
        myStr += str(item)
    return myStr


# 选择题
def choice():
    options = element.find_elements_by_xpath(choice_answer_path)
    for option in options:
        option_text_list.append(option.text)


# 填空题
def completion():
    # 进入iframe框架
    driver.switch_to.frame(i - num1)
    option = driver.find_element_by_tag_name('p')
    option_text_list.append(option.text)
    # 切换到主文档
    driver.switch_to.default_content()


# 判断题
def judge():
    options = element.find_elements_by_xpath(judge_answer_path)
    for option in options:
        if option.is_selected():
            my_dict = {"true": "√", "false": "×"}
            option_text_list.append(my_dict.get(option.get_attribute('value')))


# 每个题目的文本xPath
question_path = ".//div[1]/div"
# 选择题答案的xPath
choice_answer_path = ".//div[2]/ul/li[@class='Hover']"
# 判断题答案的xPath
judge_answer_path = ".//ul/li/input"
# 根据父路径遍历其下的所有子路径
for i, element in enumerate(parent_elements):
    # 找到题目的DOM
    question = element.find_element_by_xpath(question_path)
    # 写入
    ws.write(i, 0, question.text, style)
    # 设置高度
    ws.row(i).set_style(xlwt.easyxf('font:height 480'))
    # 答案以及其选项的文本list
    option_text_list = []
    # 答案部分
    # 分别是选择题 填空题 判断题 顺序可以手动改
    if i < num1:
        choice()
    elif i < num1 + num2:
        completion()
    elif i < num1 + num2 + num3:
        judge()
    ws.write(i, 1, '     ' + listToString(option_text_list), style)

wb.save('./answer.xls')
