import win32com.client	
import os

qi = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
computer_name = os.getenv('COMPUTERNAME')
qi.FormatName="direct=os:"+computer_name+"\\private$\\Tasks"

from constants import *
queue = qi.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)

while True:

	msg = queue.Receive(0,True,2000,0)

	if msg: 

		print( msg.Label )
		print( msg.Body )
		
	else: 

		print ( 'waiting for messages' )
