# -*- coding: utf-8 -*-
"""
Created on Mon Jul  1 13:03:42 2019

@author: afanning
"""
import numpy as np
import pandas as pd

def converttime(time):
    #offset = time & 0xFFF
    cycle1 = (time >> 12) & 0x1FFF
    cycle2 = (time >> 25) & 0x7F
    seconds = cycle2 + cycle1 / 8000.
    return seconds

def uncycle(time):
    cycles = np.insert(np.diff(time) < 0, 0, False)
    cycleindex = np.cumsum(cycles)
    return time + cycleindex * 128

if __name__ == "__main__":
    time = pd.read_csv('C:/Users/afanning/Documents/AnalysisScripts/data0705.csv', dtype='int64', header=None)
    time_converted = uncycle(converttime(np.array(time[0])))