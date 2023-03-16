# -*- coding: utf-8 -*-
"""
Created on Fri Feb 13 13:01:39 2015

@author: jackm
"""
def TIDEFF(BLATA,BLONG,DD,MM,YYYY,hh,mm):
    import numpy
    [TIM,JT]=TIMANDJT(DD,MM,YYYY,hh,mm)
        
    TIDAMP=1.2
    BLATA=-BLATA
    ONE=1
    PI=numpy.pi
    RDN=PI/180
    RDNS=RDN/3600
    TT=JT/(36525*1440)
    LAT=-BLATA*RDN
    NN=((259-numpy.mod(5*TT,ONE)*360)*3600+657.12+TT*(-482912.63+TT*(7.58+TT*0.008)))*RDNS
    S=((270.0+numpy.mod(1336.0*TT,ONE)*360.0)*3600.0+1571.72+TT*(1108406.05+TT*(7.128+TT*0.0072)))*RDNS
    P=((334.0+numpy.mod(11*TT,ONE)*360.0)*3600.0+1186.42+TT*(392522.51+TT*(-37.15-TT*0.036)))*RDNS
    H=((279.0+numpy.mod(100*TT,ONE)*360.0)*3600.0+2508.05+TT*(2768.11+TT*1.08))*RDNS 
    CSN=numpy.cos(NN)
    SNN=numpy.sin(NN)
    I=5.145*RDN
    OMEGA=23.452*RDN
    SNI=numpy.sin(I)
    CSI=numpy.cos(I)
    SNO=numpy.sin(OMEGA)
    CSO=numpy.cos(OMEGA)
    E=0.05490
    M=0.074804
    CSII=CSO*CSI-SNO*SNI*CSN
    SNII=numpy.sqrt(ONE-CSII*CSII)
    SNNU=SNI*SNN/SNII
    CSNU=numpy.sqrt(ONE-SNNU*SNNU)
    SIG=S-NN+2*numpy.arctan((SNO*SNN/SNII)/(ONE+CSN*CSNU+SNN*SNNU*CSO))
    L=SIG+2*E*numpy.sin(S-P)+1.25*E*E*numpy.sin(2*(S-P))+3.75*M*E*numpy.sin(S-2*H+P)+1.375*M*M*numpy.sin(2*(S-H))
    E1=(16751.04+TT*(-41.80-TT*0.126))*pow(10,-6)
    P1=((281.0*60+13)*60+15+TT*(6189.03+TT*(1.63+TT*0.012)))*RDNS
    L1=H+2*E1*numpy.sin(H-P1)     
    A=numpy.floor(TIM/100)
    B=(TIM-A*100)
    T0=A+B/60-12.0
    T=(15*(T0-12)+BLONG)*RDN
    CHI=T+H-numpy.arctan(SNNU/CSNU)
    CHI1=T+H
    SNLAT=numpy.sin(LAT)
    CSLAT=numpy.cos(LAT)
    SNL=numpy.sin(L)
    CSL=numpy.cos(L)
    SNL1=numpy.sin(L1)
    CSL1=numpy.cos(L1)
    CM=SNLAT*SNII*SNL+CSLAT*(CSL*numpy.cos(CHI)+SNL*numpy.sin(CHI)*CSII)
    CS=SNLAT*SNO*SNL1+CSLAT*(CSL1*numpy.cos(CHI1)+SNL1*numpy.sin(CHI1)*CSO)

    C=3.84402*pow(10,10)
    C1=1.495*pow(10,13)
    AD=pow((C*(1-E*E)),-1)
    AD1=pow((C1*(ONE-E1*E1)),-1)
    D=pow(C,-1)+AD*(E*numpy.cos(S-P)+E*E*numpy.cos(2*(S-P))+1.875*M*E*numpy.cos(S-2*H+P)+M*M*numpy.cos(2*(S-H)))
    D1=pow(C1,-1)+AD1*E1*numpy.cos(H-P1)
    R=6.37827*pow(10,8)/numpy.sqrt(ONE+0.006738*SNLAT*SNLAT)
    TID=-6.670*pow(10,-8)*((7.3537*pow(10,25)*D*D*R*D*(3.0*CM*CM-ONE+1.5*R*D*CM*(5*CM*CM-3)))+(1.993*pow(10,33)*D1*D1*R*D1*(3*CS*CS-1)))*pow(10,3)
    TIDAL=TID*TIDAMP
    return TIDAL
    
def TIMANDJT(DD,MM,YYYY,hh,mm):
    import numpy
    TIM=numpy.mod(hh*100.0+mm+1200.0,2400.0)
    addday=numpy.floor((hh*100+mm+1200)/2400)
    from datetime import date
    datenum1900=date(1900,01,01).toordinal()
    datenumtoday=date(YYYY,MM,DD).toordinal()
    dayspast=datenumtoday-datenum1900+addday
    hourspast=dayspast*24
    HH1=numpy.floor(TIM/100)
    MM1=TIM-HH1*100
    JT=hourspast*60.0+HH1*60.0+MM1
    return [TIM,JT]
    
    
    
    
    
    
    
    
    