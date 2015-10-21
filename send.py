import win32com.client	
import os

qi = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
computer_name = os.getenv('COMPUTERNAME')
qi.FormatName="direct=os:"+computer_name+"\\private$\\Tasks"

from constants import *
queue = qi.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)

msg = win32com.client.Dispatch("MSMQ.MSMQMessage")

for i in range(0,20):

	msg.Label = "Task " + str(i+1)
	msg.Body = "{report:" + str(i+1) + "}"

	msg.Send( queue )