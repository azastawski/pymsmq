# -*- coding: utf-8 -*-
import win32com.client	
import os

qi = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
computer_name = os.getenv('COMPUTERNAME')
qi.FormatName="direct=os:"+computer_name+"\\private$\\pythontest;journal"
strFind = "{report:0}"

from constants import *
queue = qi.Open(MQ_PEEK_ACCESS, MQ_DENY_NONE)

while True:


    msg = queue.PeekCurrent(ReceiveTimeout='1000')


    while msg: 

        print( msg.Label )
        print( msg.Body )
        #test Body for search string and write it out to disk
        if strFind in msg.Body:
            with open("Output.txt", "w") as text_file:
                text_file.write("\n{}".format(msg.Body))
        msg = queue.PeekNext(ReceiveTimeout='1000')

    break


