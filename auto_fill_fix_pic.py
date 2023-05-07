import openpyxl
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By

# 1、获得号码

# 2、将号码输入系统

# 3、复制图片粘贴到excel

# 获取工作簿对象
wb = openpyxl.load_workbook('4月零公里新增3-9.xlsx')

# 拿到第一个sheet
sheet = wb.worksheets[0]

# 获取工作表总行数
rows = sheet.max_row
nums = []
for row in range(rows):
    num = sheet.cell(row=row + 2, column=5).value
    nums.append(num)
print(nums)

# 创建一个Service对象
service = Service(r'./chromedriver.exe')

# 创建一个EdgeOptions对象，并设置一些选项
options = webdriver.EdgeOptions()

# 将Service对象传递给EdgeOptions的service属性
options.service = service

# 定义chrome驱动去地址
path = Service('./chromedriver.exe')
# 创建浏览器操作对象
driver = webdriver.Chrome(service=service)
driver.maximize_window()

# 访问网页
driver.get('https://dms.cfmoto-online.cn/cfdms/index.html')

# 输入
driver.find_element(by=By.XPATH, value="//*[@id='in-index-username']").send_keys("")
driver.find_element(by=By.XPATH, value="//*[@id='in-index-password']").send_keys("")

driver.find_element(by=By.XPATH, value="/html/body/div[1]/div[1]/div/div[2]/ul/li[2]/a").click()
driver.find_element(by=By.XPATH, value="/html/body/div[1]/div[1]/div/div[2]/ul/li[2]/ul/li[2]/div/div/div/div/div[3]/ul/li[2]/a").click()


for num in nums:
    # 输入索赔单号
    driver.find_element(by=By.XPATH, value="//*[@id='in-oemClaimOrderMngQuery-mng-claimno']").send_keys(num)
    driver.find_element(by=By.XPATH, value="//*[@id='btn-oemClaimOrderMngQuery-mng-search']").click()
    driver.find_element(by=By.XPATH, value="//*[@id='tab-claimOrderReportIndexQuery-list_content']/tbody/tr/td[3]/div/a").click()
    driver.find_element(by=By.XPATH, value="//*[@id='dia-tab-work-order-previewPic']").click()


print("Done")

