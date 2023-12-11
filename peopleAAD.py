import requests  
import pandas as pd  

# get the token by az login and az account get-access-token --resource 'https://graph.microsoft.com'  should use a spn token later  
token = "eyJ0eXAiOiJKV1QiLCJub25jZSI6IjFNU2RFR3BvcDBBaDNhaGZHTnhaR1lwaW1scUFObVQ1VE04d0lhT2szV3ciLCJhbGciOiJSUzI1NiIsIng1dCI6IlQxU3QtZExUdnlXUmd4Ql82NzZ1OGtyWFMtSSIsImtpZCI6IlQxU3QtZExUdnlXUmd4Ql82NzZ1OGtyWFMtSSJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC80MmY3Njc2Yy1mNDU1LTQyM2MtODJmNi1kYzJkOTk3OTFhZjcvIiwiaWF0IjoxNzAyMjY4ODE1LCJuYmYiOjE3MDIyNjg4MTUsImV4cCI6MTcwMjI3MjcxNSwiYWlvIjoiRTJWZ1lKakE4blc1NVNjKzA5MVAzNFczc0xTL0JRQT0iLCJhcHBfZGlzcGxheW5hbWUiOiJhenJzcG5faHliX3BlZGV2X2Fwal8wMSIsImFwcGlkIjoiOGM5ZjlhNTEtYmVmZi00YzZhLWJkOGYtMzI0ZDA2ZDVlNDAzIiwiYXBwaWRhY3IiOiIxIiwiaWRwIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvNDJmNzY3NmMtZjQ1NS00MjNjLTgyZjYtZGMyZDk5NzkxYWY3LyIsImlkdHlwIjoiYXBwIiwib2lkIjoiMjhhMTFlOWYtZDZiZS00Yjk4LThiYmMtMzQ3MTBhMTA4Yzg4IiwicmgiOiIwLkFRa0FiR2YzUWxYMFBFS0M5dHd0bVhrYTl3TUFBQUFBQUFBQXdBQUFBQUFBQUFBSkFBQS4iLCJzdWIiOiIyOGExMWU5Zi1kNmJlLTRiOTgtOGJiYy0zNDcxMGExMDhjODgiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiRVUiLCJ0aWQiOiI0MmY3Njc2Yy1mNDU1LTQyM2MtODJmNi1kYzJkOTk3OTFhZjciLCJ1dGkiOiJ4cHZRRXU5clVrQzF3dFVwNzhrdUFBIiwidmVyIjoiMS4wIiwid2lkcyI6WyIwOTk3YTFkMC0wZDFkLTRhY2ItYjQwOC1kNWNhNzMxMjFlOTAiXSwieG1zX3RjZHQiOjEzNzE4Mzg1NTAsInhtc190ZGJyIjoiRVUifQ.qxQfouTnspla3JR8vxWRXCWHw6xgL4GfyI8qJQNX7QgoUTpbySQgCwfSIL22-b9clLZavtgqYiJ7JD6SSidbpseZ_OV0X_boEbeB5BAh1IIXpQHgq-DXnL6meoUkpLVXsrKNXxljZY9ZrGhcBjbDu_RTCRE0X0nag0DnTd_oHHbkTpPAm-kWDNO2n7Ez8yPlhGEmjg4wHjbNyFqb1jH1KBBcDcj2wTCQMkiYcwZ_q10PRrCkXytFBwQvgEDnDojpT7rfCOErfGPwWpZPbwBaDj4uxsO3jVH7mq3YiuipFjRJGcEOShm9QsMr8rPajt18wXiEJSjr3AnQIcKUZ6kv2g"  
# 读取Excel文件中的数据  
df=pd.read_excel('aad.xlsx')
column_data=df.iloc[:,0]
managerColumnList=[]
for user in column_data:
  try:
      eid=str(user)[0:7]
      url = "https://graph.microsoft.com/v1.0/users?$filter=mailNickname eq '"+eid+"'&$expand=manager($select=displayName)"  
      headers = {  
      "Authorization": "Bearer " + token  
      }   
      response = requests.get(url, headers=headers)
      data=response.json()
      manager_name = data['value'][0]['manager']['displayName'] 
      print (eid,manager_name)
      managerColumnList.append(manager_name)
      #print(managerColumnList)
  except IndexError:
      print("Cannot find Manager")
      managerColumnList.append('NA')

df.loc[:, 'Manager'] =  managerColumnList
  #else:
  #print("job finished")
  #finally:
  #df.loc[:, 'B'] =  managerColumnList
df.to_excel('aad.xlsx', index=False)