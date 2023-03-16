# -*- coding: utf-8 -*-
"""
Created on Mon Nov 30 09:28:35 2015

@author: jackm
"""

# -*- coding: utf-8 -*-
"""
Created on Fri Feb 13 13:01:39 2015

@author: jackm
"""
def TerrCorrg(Eastingv,Northingv,Heightv,rho,reso):
    TypeTC=1 
    import numpy
    if TypeTC==1:
        # Here we are calculating the terr correction by naggy prisms
        Eastingc1=Eastingv-reso/2
        Eastingc2=Eastingv+reso/2
        Northing1=Northingv-reso/2
        Northing2=Northingv+reso/2
        ### First put all the xs in the correct quadrant
        IDSy=numpy.empty(1)
        IDSy=numpy.where(numpy.sign(Eastingc1)!=numpy.sign(Eastingc2))
        IDSn=numpy.where(numpy.sign(Eastingc1)==numpy.sign(Eastingc2))
        if len(IDSy)!=0:
            Eastingc12=numpy.array([numpy.zeros([numpy.size(IDSy)]),numpy.abs(Eastingc1[IDSy])])
            Eastingc22=numpy.array([Eastingc2[IDSy],numpy.zeros([numpy.size(IDSy)])])
            EastingsTopc1=Eastingc1[IDSn]
            EastingsTopc2=Eastingc2[IDSn]
            Eastingc1mod1=numpy.append(EastingsTopc1,Eastingc12)
            Eastingc2mod1=numpy.append(EastingsTopc2,Eastingc22)
        # outputs for x swap
            Eastingc1mod1a=numpy.minimum(Eastingc1mod1,Eastingc2mod1)
            Eastingc2mod1a=numpy.maximum(Eastingc1mod1,Eastingc2mod1)
            Northingc1mod1=numpy.append(Northing1[IDSn],numpy.append(Northing1[IDSy],Northing1[IDSy]))
            Northingc2mod1=numpy.append(Northing2[IDSn],numpy.append(Northing2[IDSy],Northing2[IDSy]))
            Heightmod1=numpy.append(Heightv[IDSn],numpy.append(Heightv[IDSy],Heightv[IDSy]))
        if len(IDSy)==0:
            Eastingc1mod1a=Eastingc1
            Eastingc2mod1a=Eastingc2
            Northingc1mod1=Northing1
            Northingc2mod1=Northing2
            Heightmod1=Heightv
        ### Next put all the ys in the correct quadrant
        IDSy=numpy.where(numpy.sign(Northingc1mod1)!=numpy.sign(Northingc2mod1))
        IDSn=numpy.where(numpy.sign(Northingc1mod1)==numpy.sign(Northingc2mod1))
        if len(IDSy)!=0:
            Northingc1mod12=numpy.array([numpy.zeros([numpy.size(IDSy)]),numpy.abs(Northingc1mod1[IDSy])])
            Northingc2mod12=numpy.array([Northingc2mod1[IDSy],numpy.zeros([numpy.size(IDSy)])])
            Northingmod1Topc1=Northingc1mod1[IDSn]
            Northingmod1Topc2=Northingc2mod1[IDSn]
            Northingc1mod2=numpy.append(Northingmod1Topc1,Northingc1mod12)
            Northingc2mod2=numpy.append(Northingmod1Topc2,Northingc2mod12)
            # outputs for y swap
            Northingc1mod2a=numpy.minimum(Northingc1mod2,Northingc2mod2)
            Northingc2mod2a=numpy.maximum(Northingc1mod2,Northingc2mod2)
        
            Eastingc1mod2=numpy.append(Eastingc1mod1a[IDSn],numpy.append(Eastingc1mod1a[IDSy],Eastingc1mod1a[IDSy]))
            Eastingc2mod2=numpy.append(Eastingc2mod1a[IDSn],numpy.append(Eastingc2mod1a[IDSy],Eastingc2mod1a[IDSy]))
            Heightmod2=numpy.append(Heightmod1[IDSn],numpy.append(Heightmod1[IDSy],Heightmod1[IDSy]))
        
        if len(IDSy)==0:
            Northingc1mod2a=Northingc1mod1
            Northingc2mod2a=Northingc2mod1
            Eastingc1mod2=Eastingc1mod1a
            Eastingc2mod2=Eastingc2mod1a
            Heightmod2=Heightmod1
        TerrC=0
        for i in range (0,numpy.size(Eastingc1mod2)):
            TCadd=float(NAGGYTC(float(numpy.abs(Eastingc1mod2[i])),numpy.abs(float(Eastingc2mod2[i])),numpy.abs(float(Northingc1mod2a[i])),numpy.abs(float(Northingc2mod2a[i])),float(Heightmod2[i]),float(rho)))
            
            if numpy.isnan(TCadd)==0:
                TerrC=TerrC+TCadd
    return TerrC
    
def NAGGYTC(x1,x2,y1,y2,h,rho):
    import numpy
    twopiG = 0.0419
    G = twopiG / (2 * numpy.pi)
    fac11 = numpy.sqrt(x1** 2 + y1** 2)
    fac11h = numpy.sqrt(x1** 2 + y1** 2 + h** 2)
    fac12 = numpy.sqrt(x1** 2 + y2** 2)
    fac12h = numpy.sqrt(x1** 2 + y2** 2 + h** 2)
    fac21 = numpy.sqrt(x2** 2 + y1** 2)
    fac21h = numpy.sqrt(x2** 2 + y1** 2 + h** 2)
    fac22 = numpy.sqrt(x2** 2 + y2** 2)
    fac22h = numpy.sqrt(x2** 2 + y2** 2 + h** 2)
    fac2h = numpy.sqrt(y2** 2 + h** 2)
    fac1h = numpy.sqrt(y1** 2 + h** 2)
    y2h = y2** 2 + h** 2
    y1h = y1** 2 + h** 2
    terrc = x2 * (numpy.log((y2 + fac22)/ (y2 + fac22h)) - numpy.log((y1 + fac21)/ (y1 + fac21h))) - x1 * (numpy.log((y2 + fac12)/ (y2 + fac12h)) - numpy.log((y1 + fac11)/ (y1 + fac11h))) + y2 * (numpy.log((x2 + fac22)/ (x2 + fac22h)) - numpy.log((x1 + fac12)/ (x1 + fac12h))) - y1 * (numpy.log((x2 + fac21)/ (x2 + fac21h)) - numpy.log((x1 + fac11)/ (x1 + fac11h))) + h * (numpy.arcsin((y2h + y2 * fac22h)/ ((y2 + fac22h) * fac2h)) - numpy.arcsin((y2h + y2 * fac12h)/ ((y2 + fac12h) * fac2h)) - numpy.arcsin((y1h + y1 * fac21h)/ ((y1 + fac21h) * fac1h)) + numpy.arcsin((y1h + y1 * fac11h)/ ((y1 + fac11h) * fac1h)))
    terrc = numpy.real(G * rho * terrc)
    return terrc
