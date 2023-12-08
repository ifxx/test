from selenium import webdriver
from selenium.webdriver import ChromeService
from selenium.webdriver.common.by import By
import pandas as pd

df=pd.read_excel('Audit Communication Contacts.xlsx')
column_data=df.iloc[:,0]

for user in column_data:
    print (user)

service= ChromeService(executable_path=r"chromedriver.exe")

# 创建 ChromeDriver 实例
driver = webdriver.Chrome(service=service)

# 导入浏览器的 cookie
driver.get('https://people.wdf.sap.corp/profiles/I342184#?profile_tab=organization')  # 打开网站登录页面
# 在打开的网页上手动登录，并确保登录成功
cookies = driver.get_cookies()

for cookie in cookies:
    driver.add_cookie(cookie)

# 刷新页面，确认登录状态已经被保留
driver.refresh()
#body > div.wrap > div.search-container.sort_by_score.ready > div.content > div > div > div > div.main-profile-info > div > div.inline_content > div.info.search-info > div.container.mobile-more-less > div.row.collection.corporate > div.col-lg-2.col-md-4.col-sm-4.col-xs-12.customize.manager_link > div > div:nth-child(2) > span.value > a:nth-child(1)
#//*[@id="uhHtml"]/body/div[2]/div[1]/div[2]/div/div/div/div[1]/div/div[2]/div[2]/div[2]/div[1]/div[6]/div/div[2]/span[1]/a[1]
#/html/body/div[2]/div[1]/div[2]/div/div/div/div[1]/div/div[2]/div[2]/div[2]/div[1]/div[6]/div/div[2]/span[1]/a[1]

manager = driver.find_element(By.CSS_SELECTOR, "body > div.wrap > div.search-container.sort_by_score.ready > div.content > div > div > div > div.main-profile-info > div > div.inline_content > div.info.search-info > div.container.mobile-more-less > div.row.collection.corporate > div.col-lg-2.col-md-4.col-sm-4.col-xs-12.customize.manager_link > div > div:nth-child(2) > span.value > a:nth-child(1)")
#eid = driver.find_element(By.CSS_SELECTOR,'body > div.wrap > div.search-container.sort_by_score.ready > div.content > div > div > div > div.main-profile-info > div > div.inline_content > div.info.search-info > div.container.mobile-more-less > div.row.collection.corporate > div.col-lg-2.col-md-4.col-sm-4.col-xs-12.customize.manager_link > div > div:nth-child(2) > span.value > a:nth-child(1)')
#element = driver.find_element_by_class_name("col-lg-2 col-md-4 col-sm-4 col-xs-12 customize manager_link")
#element = driver.find_element_by_xpath('//div[@class="col-lg-2 col-md-4 col-sm-4 col-xs-12 customize manager_link"]')  # 修改这里的xpath表达式，使用适合你情况的class名称  
print(manager.text)

attributes = manager.get_attribute("href")
print(attributes)
  # 打印元素的文本内容  

#page_source = driver.page_source  
#print("页面源代码:\n", page_source)  

# 关闭浏览器实例
#driver.quit()
