# -*- coding: utf-8 -*-
import win32com.client	
import os

qi = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
computer_name = os.getenv('COMPUTERNAME')
qi.FormatName="direct=os:"+computer_name+"\\private$\\pythontest;journal"
#qi.PathName = r".\Private$\Tasks;journal"

from constants import *
queue = qi.Open(MQ_PEEK_ACCESS, MQ_DENY_NONE)

while True:


    msg = queue.PeekCurrent(WantBody=True,ReceiveTimeout='0')


    while msg: 

        print( msg.Label )
        print( msg.Body )

        msg = queue.PeekNext(WantBody=True,ReceiveTimeout='0')

    break
    #else: 

    #    print ( 'waiting for messages' )


