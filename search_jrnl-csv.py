# -*- coding: utf-8 -*-
"""
Created on Wed Oct 21 09:14:52 2015

@author: Vijay

Created on Tue Oct 13 22:25:52 2015

"""
import win32com.client                
import os
import codecs
import csv


qi = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
computer_name = os.getenv('COMPUTERNAME')
qi.FormatName="direct=os:"+computer_name+"\\private$\\Tasks;journal"

List = []
rmList = []

with open("TestSearchMsmq.csv") as csvfile:
    readCSV = csv.reader(csvfile,delimiter=',')    
    List = list(readCSV)

for val in List:


   for i in range(len(val)):
         print("\n")
         strval = val[i]
         print(strval)
         queue = qi.Open(32,0)
         strFind = codecs.decode(strval,'unicode_escape')
         strOutputPath = strFind.replace('/','-').replace(':','-')+"_messages.txt"
         path = codecs.decode(strOutputPath,'unicode_escape')
           
         with open(path, "w") as txt_file:
             with open(path,"a") as text_file:
                          text_file.write("Value '"'{}'"' found in the below messages :\n".format(strFind))
                          text_file.close()
             txt_file.close()
             
             while True:
                 Dict = {}
                 msg = queue.PeekCurrent(0,True,1000,0)
                 if msg:         
                     
                      if strFind in msg.Body:                            
                          Dict = {'Message SentTime :': msg.SentTime, 'Message Body :': msg.Body}                                                  
                      msg = queue.PeekNext(0,True,1000,0)
                      sorted(Dict.keys(),reverse=True)                     
                      for key,value in Dict.items():
                          print(key+"  "+str(value))
                          with open(path,"a") as text_file:
                             text_file.write("\n{}\n{}\n".format(key,value))                              
                      rmList.append(path)
                      
                 else:
                      if 'Body' not in open(path).read():
                          os.remove(path)
                      print("No More Messages in Queue")              
                      break
                      
         queue.Close()


     
     
     
     
     
