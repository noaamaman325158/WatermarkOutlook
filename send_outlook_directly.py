# Send Plain text email

import win32com.client as client

outlook = client.Dispatch('Outlook.Application')

message = outlook.CreateItem(0)



message.Display()
