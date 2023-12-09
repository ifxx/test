from selenium import webdriver
from selenium.webdriver import ChromeService
from selenium.webdriver.common.by import By
import pandas as pd

service= ChromeService(executable_path=r"chromedriver.exe")

# 创建 ChromeDriver 实例
driver = webdriver.Chrome(service=service)

df=pd.read_excel('test\Audit Communication Contacts.xlsx')
column_data=df.iloc[:,0]

for user in column_data:
    eid=user[0:7]
    peopleUrl= "https://people.wdf.sap.corp/profiles/"+eid+"#?profile_tab=organization"
    driver.get(peopleUrl)  # 打开网站登录页面
    cookies = driver.get_cookies()
      for cookie in cookies:
        driver.add_cookie(cookie)
    driver.refresh()
    manager = driver.find_element(By.CSS_SELECTOR, "body > div.wrap > div.search-container.sort_by_score.ready > div.content > div > div > div > div.main-profile-info > div > div.inline_content > div.info.search-info > div.container.mobile-more-less > div.row.collection.corporate > div.col-lg-2.col-md-4.col-sm-4.col-xs-12.customize.manager_link > div > div:nth-child(2) > span.value > a:nth-child(1)")
    print(manager.text)

    attributes = manager.get_attribute("href")
    print(attributes)
  # 打印元素的文本内容  

#page_source = driver.page_source  
#print("页面源代码:\n", page_source)  

# 关闭浏览器实例
driver.quit()
