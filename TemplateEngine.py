# -*- coding: UTF-8 -*-
import numpy as np
tempPattern = '{{loop%dch%d}}' #替换模板
maxPattern = '{{max%d}}' #替换模板
channels = 9
loops = 15
def generate(replace,tempratures):
    for loop in range(loops):
        for ch in range(channels):
            key = tempPattern%(loop+1,ch+1)
            value = tempratures[loop,ch]
            replace(key,value)
            print('replacing '+ key + ' to ' + str(value))
    for loop in range(loops):
        print(tempratures[loop,:].max())
if __name__ == '__main__':
    pass
