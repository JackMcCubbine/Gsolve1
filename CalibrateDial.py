# -*- coding: utf-8 -*-
"""
Created on Sat Feb 14 14:07:44 2015

@author: jackm
"""
def Calibrate(Dial,Table):
    import numpy
    N=Table.nrows
    Lower=numpy.zeros(N)
    Upper=numpy.zeros(N)
    Factor=numpy.zeros(N)
    M=numpy.size(Dial)
    P=numpy.zeros(M)
    DialCal=numpy.zeros(M)
    for i in range (0,N):
        Lower[i]=Table.cell(i,0).value
        Upper[i]=Table.cell(i,1).value
        Factor[i]=Table.cell(i,2).value
    HASNEGs=0
    for i in range(0,N):
        if Lower[i]<0:
            HASNEGs=1
    if HASNEGs==0:
        Diff=Upper-Lower
        FactorDiff=numpy.multiply(Diff,Factor)
        FactorDiffCS=numpy.cumsum(FactorDiff)
        for i in range (0,M):
            for j in range (0,N):
                if Dial[i]>=Lower[j] and Dial[i]<Upper[j]:
                    P[i]=int(j)
        for i in range (0,M):
            DialCal[i]=FactorDiffCS[int(P[i]-1)]+(Dial[i]-Upper[int(P[i]-1)])*Factor[int(P[i])]
#        print(numpy.sum())
    if HASNEGs==1 and N==1:
        for i in range(0,M):
            DialCal[i]=Dial[i]*Factor[0]
    print(N)
           
    return DialCal