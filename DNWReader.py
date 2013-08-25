# -*- coding: UTF-8 -*-
import struct
import serial
import time
import lookupTable
import TemplateEngine
import win32com.client
import win32com.client.dynamic
import numpy as np
class ComDev:
    def __init__(self):
        self.com = None
        self.command_startup = (-1,':0002003C',32) #32
        self.command_calibration = (-1,':02021856',2)
        self.command_readTemp = [(i,(':020710020'+hex(i).replace('0x','')+'010004'+hex((90+i)).replace('0x','')).upper(),8) for i in range(16)]
    def Open(self,com):
        try:
            self.com = serial.Serial(com,baudrate=9600, bytesize=8,parity='N',
            stopbits=1,xonxoff=0,rtscts=0, timeout=2)
        except:
            self.com = None
            print 'Open %s fail!' %com

    def Close(self):
        if type(self.com) != type(None):
            self.com.close()
            self.com = None
            return True
        return False

    def ReadData(self,RevBytes):
        if type(self.com) != type(None):
            try:
                data = self.com.read(RevBytes)
                return data
            except:
                print 'ReadData fail!'
                self.Close()
                return None
        return None

    def SendData(self,Data):
        if type(self.com) != type(None):
            try:
                self.com.write(Data)
                return True
            except:
                print 'SendData fail!'
                self.Close()
                return False
        return False

    def Transform(self,inData):
        mbf_4bytestring = inData.lower().decode('hex')
        msbin = struct.unpack('4B', mbf_4bytestring)
##        ieee = [0] * 4
##        sign =1 if  (msbin[3] & 0x80)==0 else -1
##        int_part = (msbin[2]|0x80)>>1
##        print(int_part)
##        digits = ((msbin[2]&0x01)<<16)+(msbin[1]<<8)+msbin[0]
##        a = [2**-(i+1) for i in range(17)]
##        b = [(digits>>(16-i))&1 for i in range(17)]
##        c = map(lambda x,y:x*y,a,b)
##        digit_part = sum(c)
##        print(digit_part)
##        ieee_exp = (msbin[3]<<1) + ((msbin[2] & 0x80)>>7)-127
##        return sign*(int_part+digit_part)
        ieee = [0] * 4
        int_part = (msbin[1]|0x80)>>1
        digits = ((msbin[1]&0x01)*256)+msbin[2]
        a = [2**-(i+1) for i in range(9)]
        b = [(digits>>(8-i))&1 for i in range(9)]
        c = map(lambda x,y:x*y,a,b)
        digit_part = sum(c)
        return int_part+digit_part

    def Stop(self):
        self.keepGoing = False

    def IsRunning(self):
        return self.running

    def IsOpen(self):
            return type(self.com) != type(None)



def ReadWrod(fileName):
    app  = win32com.client.Dispatch("Word.Application")
    #app.Visible = True
    app.Documents.Open(fileName)

    return app

def MS_Word_Find_Replace(app, Search_Word, replace_str):
    wdStory = 6
    app.Selection.HomeKey(Unit=wdStory)
    find = app.Selection.Find
    find.Text = Search_Word
    while app.Selection.Find.Execute():
        app.Selection.TypeText(Text=replace_str)
        print "Find It ", Search_Word


def MS_Wrod_SaveAS(app, fileName):
    print fileName
    app.ActiveDocument.SaveAs(fileName)
    app.ActiveDocument.Close()

if __name__ == '__main__':

    #---------设置---------------------------------
    loops = 15   #循环次数
    channelRange= [7,8]  #待测量的通道

    maxLoops = 16  #最大循环次数
    maxChannels = 16  #最大通道数
    sleepTime = 0  #间隔时间 s
    templateFilePath = 'E:\\TaoJin\\Portable Python 2.7.5.1\\Template.doc' #模板文件
    outputFilePath = 'E:\\TaoJin\\Portable Python 2.7.5.1\\output.doc' #输出文件
    #-----------------------------------------------
    assert maxLoops>loops
    assert maxChannels > max(channelRange)

    AppWord = ReadWrod(templateFilePath)
    AppWord.ActiveDocument.SaveAs(outputFilePath) #try to save to avoid conflict after acquisition

    dev = ComDev()
##    print dev.Transform('07DC7020') #07DC7020
##    exit()
    dev.Open('COM1')
    dev.com.setRTS(0)
    dev.com.setDTR(1)
    ret = dev.SendData(dev.command_startup[1])
    time.sleep(0.4)
    value = dev.ReadData(dev.command_startup[2])
    print(value)
    print('The following channels will checked:' + ",".join(map(str,channelRange)))
    channelRange = map(lambda x:x-1,channelRange)
    channelsToCheck = [dev.command_readTemp[i] for i in channelRange if i <len(dev.command_readTemp)]
    tempratures = np.zeros((maxLoops,maxChannels))
    tempratures[:,:] = np.nan
    for i in range(loops):
        print('loop:'+str(i+1))
        print('calibration...')
        ret = dev.SendData(dev.command_calibration[1])
        if not ret:
            break
        time.sleep(5) #等待校准完成
        value = dev.ReadData(dev.command_calibration[2])
        if value is None:
            dev.Close()
            exit()
        else:
            print(value)

        for j,command,retLen in channelsToCheck:
            print('read ch'+ str(j+1) + '...')
            ret = dev.SendData(command)
            value = dev.ReadData(retLen)
            res = dev.Transform(value)
            temp = lookupTable.lookup(res)                                       #gongshi
            print(value,res,temp)
            tempratures[i,j]= temp
        time.sleep(sleepTime)
    dev.Close()
    print(tempratures)
    replace =lambda findStr,replaceStr: MS_Word_Find_Replace(AppWord, findStr, replaceStr)
    TemplateEngine.generate(replace,tempratures)
    MS_Wrod_SaveAS(AppWord, outputFilePath)
    print('Done!')


