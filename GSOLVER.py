# -*- coding: utf-8 -*-
"""
Created on Sat Feb 14 18:57:32 2015

@author: jackm
"""
# Reilies original Method
def GsolverM2(TideFreeCalDial,SurveyedNames,SiteNames,TIM,TIENames,TIEVals,Loops,UL):
    import numpy
    N1=numpy.size(TideFreeCalDial)
    N2=numpy.size(SiteNames) 
    N3=numpy.size(TIENames)
    Nloops=numpy.max(numpy.unique(Loops))
    SLoops=numpy.unique(Loops)
    P=numpy.zeros((N1,N2))
    if UL==1:
        T=numpy.zeros((N1,2*Nloops))    
    else:
        T=numpy.zeros((N1,2))    
    Q=numpy.zeros((N2,N2))
    h=numpy.zeros((N2,1))
    x=numpy.zeros((N1,1))
    for i in range(0,N1):
        x[i,0]=TideFreeCalDial[i]
        for j in range(0,N2):
            if SurveyedNames[i]==SiteNames[j]:
                P[i,j]=1
                # Loop to decouple Tie Stations from the rest of the equations
            for k in range(0,N3):
                if SurveyedNames[i]==TIENames[k]:
                    P[i,j]=0
                    x[i,0]=TideFreeCalDial[i]-TIEVals[k]
        if UL==1:
            for k in range(0,int(Nloops)):
                if Loops[i]==SLoops[k]:
                    T[i,2*k]=1
                    T[i,2*k+1]=float(TIM[i])
        else:
            T[i,0]=1
            T[i,1]=float(TIM[i])
            
        
    for k in range(0,N2):
        for j in range(0,N3):
            if SiteNames[k]==TIENames[j]:
                Q[k,k]=1
                h[k]=TIEVals[j]
    PdashP=numpy.dot(numpy.transpose(P),P)
    QdashQ=numpy.dot(numpy.transpose(Q),Q)
    PdashT=numpy.dot(numpy.transpose(P),T)
    TdashP=numpy.dot(numpy.transpose(T),P)
    TdashT=numpy.dot(numpy.transpose(T),T)
    Pdashx=numpy.dot(numpy.transpose(P),x)
    Qdashh=numpy.dot(numpy.transpose(Q),h)
    Tdashx=numpy.dot(numpy.transpose(T),x)
    A=numpy.hstack((numpy.vstack((PdashP+QdashQ,-TdashP)),numpy.vstack((-PdashT,TdashT))))
    if UL==1:
        Al=numpy.hstack((numpy.vstack((P,Q)),numpy.vstack((-T,numpy.zeros((N2,2*Nloops))))))    
    else:
        Al=numpy.hstack((numpy.vstack((P,Q)),numpy.vstack((-T,numpy.zeros((N2,2))))))    
 
    b=numpy.vstack((Pdashx+Qdashh,-Tdashx))
    b2=numpy.vstack((x,h))
    for i in range(0,N2+1):
        for j in range(0,N2+1):
            if numpy.array_equal(A[i,],A[j,]) and i!=j: 
                print((i,j))
        if numpy.sum(A[i,])==0:
            print[i]
    Sol=numpy.linalg.solve(A,b)
    RES=b2-numpy.dot(Al,Sol)
    RESMeasured=numpy.zeros((N1,1))
    RESMeasureds=numpy.zeros((N1,1))
    for i in range(0,N1):
        RESMeasured[i]=RES[i]
        RESMeasureds[i]=RES[i]*RES[i]
    
    Ainv=numpy.linalg.inv(A)
    Pv=numpy.diagonal(Ainv)
    C=numpy.sum(RESMeasureds)/(N1+N3-N2-1-1)
    PvC=C*Pv
    if UL==1:
       Drift=numpy.zeros((Nloops,1))
       BL=numpy.zeros((Nloops,1))
       for i in range(0,int(Nloops)):
           Drift[i]=Sol[int(N2+2*i+2-1)]
           BL[i]=Sol[int(N2+2*i)]
           
    else:
        Drift=Sol[int(N2+1)]
        BL=Sol[int(N2)]
    return [Sol,RES,PvC,Drift,BL]

# Reilies original Method with beta factor
def GsolverM2Beta(TideFreeCalDial,SurveyedNames,SiteNames,TIM,TIENames,TIEVals,Loops,UL,CalDials):
    import numpy
    N1=numpy.size(TideFreeCalDial)
    N2=numpy.size(SiteNames) 
    N3=numpy.size(TIENames)
    Nloops=numpy.max(numpy.unique(Loops))
    SLoops=numpy.unique(Loops)
    P=numpy.zeros((N1,N2))
    if UL==1:
        T=numpy.zeros((N1,2*Nloops+1))    
    else:
        T=numpy.zeros((N1,2+1))    
    Q=numpy.zeros((N2,N2))
    h=numpy.zeros((N2,1))
    x=numpy.zeros((N1,1))
    for i in range(0,N1):
        x[i,0]=TideFreeCalDial[i]
        for j in range(0,N2):
            if SurveyedNames[i]==SiteNames[j]:
                P[i,j]=1
                # Loop to decouple Tie Stations from the rest of the equations
            for k in range(0,N3):
                if SurveyedNames[i]==TIENames[k]:
                    P[i,j]=0
                    x[i,0]=TideFreeCalDial[i]-TIEVals[k]
        if UL==1:
            for k in range(0,int(Nloops)):
                if Loops[i]==SLoops[k]:
                    T[i,2*k]=1
                    T[i,2*k+1]=float(TIM[i])
            T[i,2*Nloops]=float(CalDials[i])
 
        else:
            T[i,0]=1
            T[i,1]=float(TIM[i])
            T[i,2]=float(CalDials[i])
            
        
    for k in range(0,N2):
        for j in range(0,N3):
            if SiteNames[k]==TIENames[j]:
                Q[k,k]=1
                h[k]=TIEVals[j]
    PdashP=numpy.dot(numpy.transpose(P),P)
    QdashQ=numpy.dot(numpy.transpose(Q),Q)
    PdashT=numpy.dot(numpy.transpose(P),T)
    TdashP=numpy.dot(numpy.transpose(T),P)
    TdashT=numpy.dot(numpy.transpose(T),T)
    Pdashx=numpy.dot(numpy.transpose(P),x)
    Qdashh=numpy.dot(numpy.transpose(Q),h)
    Tdashx=numpy.dot(numpy.transpose(T),x)
    A=numpy.hstack((numpy.vstack((PdashP+QdashQ,-TdashP)),numpy.vstack((-PdashT,TdashT))))
    if UL==1:
        Al=numpy.hstack((numpy.vstack((P,Q)),numpy.vstack((-T,numpy.zeros((N2,2*Nloops+1))))))    
    else:
        Al=numpy.hstack((numpy.vstack((P,Q)),numpy.vstack((-T,numpy.zeros((N2,2+1))))))    
 
    b=numpy.vstack((Pdashx+Qdashh,-Tdashx))
    b2=numpy.vstack((x,h))
    for i in range(0,N2+1):
        for j in range(0,N2+1):
            if numpy.array_equal(A[i,],A[j,]) and i!=j: 
                print((i,j))
        if numpy.sum(A[i,])==0:
            print[i]
    Sol=numpy.linalg.solve(A,b)
    RES=b2-numpy.dot(Al,Sol)
    RESMeasured=numpy.zeros((N1,1))
    RESMeasureds=numpy.zeros((N1,1))
    for i in range(0,N1):
        RESMeasured[i]=RES[i]
        RESMeasureds[i]=RES[i]*RES[i]
    
    Ainv=numpy.linalg.inv(A)
    Pv=numpy.diagonal(Ainv)
    C=numpy.sum(RESMeasureds)/(N1+N3-N2-1-2)
    PvC=C*Pv
    if UL==1:
        BetaCal=Sol[N2+2*Nloops]
    else:
        BetaCal=Sol[N2+2]
    
    if UL==1:
       Drift=numpy.zeros((Nloops,1))
       BL=numpy.zeros((Nloops,1))
       for i in range(0,int(Nloops)):
           Drift[i]=Sol[int(N2+2*i+2-1)]
           BL[i]=Sol[int(N2+2*i)]
           
    else:
        Drift=Sol[int(N2+1)]
        BL=Sol[int(N2)]
    return [Sol,RES,PvC,BetaCal,Drift,BL]
#    
#    if UL==1:
#       Drift=numpy.zeros((Nloops,1))
#       for i in range(0,int(Nloops)):
#           Drift[i]=Sol[int(N2+2*i+2-1)]
#    else:
#        Drift=Sol[int(N2+1)]        
#    return [Sol,RES,PvC,BetaCal,Drift]

# Method one not decoupled - reweight equations by Tie variance and variance of other stations
def GsolverM1(TideFreeCalDial,SurveyedNames,SiteNames,TIM,TIENames,TIEVals,Loops,UL):
    import numpy
    N1=numpy.size(TideFreeCalDial)
    N2=numpy.size(SiteNames) 
    N3=numpy.size(TIENames)
    Nloops=numpy.max(numpy.unique(Loops))
    SLoops=numpy.unique(Loops)
    P=numpy.zeros((N1,N2))
    if UL==1:
        T=numpy.zeros((N1,2*Nloops))    
    else:
        T=numpy.zeros((N1,2))    
    Q=numpy.zeros((N2,N2))
    h=numpy.zeros((N2,1))
    x=numpy.zeros((N1,1))
    for i in range(0,N1):
        x[i,0]=TideFreeCalDial[i]
        for j in range(0,N2):
            if SurveyedNames[i]==SiteNames[j]:
                P[i,j]=1
                x[i,0]=TideFreeCalDial[i]
        if UL==1:
            for k in range(0,int(Nloops)):
                if Loops[i]==SLoops[k]:
                    T[i,2*k]=1
                    T[i,2*k+1]=float(TIM[i])
        else:
            T[i,0]=1
            T[i,1]=float(TIM[i])
            
        
    for k in range(0,N2):
        for j in range(0,N3):
            if SiteNames[k]==TIENames[j]:
                Q[k,k]=1
                h[k]=TIEVals[j]
    PdashP=numpy.dot(numpy.transpose(P),P)
    QdashQ=numpy.dot(numpy.transpose(Q),Q)
    PdashT=numpy.dot(numpy.transpose(P),T)
    TdashP=numpy.dot(numpy.transpose(T),P)
    TdashT=numpy.dot(numpy.transpose(T),T)
    Pdashx=numpy.dot(numpy.transpose(P),x)
    Qdashh=numpy.dot(numpy.transpose(Q),h)
    Tdashx=numpy.dot(numpy.transpose(T),x)
    A=numpy.hstack((numpy.vstack((PdashP+QdashQ,-TdashP)),numpy.vstack((-PdashT,TdashT))))
    if UL==1:
        Al=numpy.hstack((numpy.vstack((P,Q)),numpy.vstack((-T,numpy.zeros((N2,2*Nloops))))))    
    else:
        Al=numpy.hstack((numpy.vstack((P,Q)),numpy.vstack((-T,numpy.zeros((N2,2))))))    
 
    b=numpy.vstack((Pdashx+Qdashh,-Tdashx))
    b2=numpy.vstack((x,h))
    for i in range(0,N2+1):
        for j in range(0,N2+1):
            if numpy.array_equal(A[i,],A[j,]) and i!=j: 
                print((i,j))
        if numpy.sum(A[i,])==0:
            print[i]
    Sol=numpy.linalg.solve(A,b)
    RES=b2-numpy.dot(Al,Sol)
    # Calclualte Tie station variances and measurment variances
    Epps=numpy.zeros((N1,1))
    E=numpy.zeros((N3,1))  
    for i in range(0,N1):
        Epps[i]=RES[i]*RES[i]
    k=-1  
    for i in range(0,N2):
        if int(Q[i,i])==1:
            k=k+1
            E[k]=RES[i+N1]*RES[i+N1]   
    VEPPS=numpy.sum(Epps)/(N1-(N1*N2/(N3+N1)+1+1))
    VE=numpy.sum(Epps)/(N3-N3*N2/(N3+N1))
    # Re-weight LS equations and recompute solution
    #PdashP=1/VEPPS*numpy.dot(numpy.transpose(P),P)
    #QdashQ=1/VE*numpy.dot(numpy.transpose(Q),Q)
    #PdashT=1/VEPPS*numpy.dot(numpy.transpose(P),T)
    #TdashP=1/VEPPS*numpy.dot(numpy.transpose(T),P)
    #TdashT=1/VEPPS*numpy.dot(numpy.transpose(T),T)
    #Pdashx=1/VEPPS*numpy.dot(numpy.transpose(P),x)
    #Qdashh=1/VE*numpy.dot(numpy.transpose(Q),h)
    #Tdashx=1/VEPPS*numpy.dot(numpy.transpose(T),x)
    #A=numpy.hstack((numpy.vstack((PdashP+QdashQ,-TdashP)),numpy.vstack((-PdashT,TdashT))))
    #if UL==1:
    #    Al=numpy.hstack((numpy.vstack((P,Q)),numpy.vstack((-T,numpy.zeros((N2,2*Nloops))))))    
    #else:
    #    Al=numpy.hstack((numpy.vstack((P,Q)),numpy.vstack((-T,numpy.zeros((N2,2))))))    
 
    #b=numpy.vstack((Pdashx+Qdashh,-Tdashx))
    #b2=numpy.vstack((x,h))
    #for i in range(0,N2+1):
    #    for j in range(0,N2+1):
    #        if numpy.array_equal(A[i,],A[j,]) and i!=j: 
    #            print((i,j))
    #    if numpy.sum(A[i,])==0:
    #        print[i]
    #RES=b2-numpy.dot(Al,Sol)
    RESMeasured=numpy.zeros((N1,1))
    RESMeasuredsq=numpy.zeros((N1,1))
    for i in range(0,N1):
        RESMeasured[i]=RES[i]
        RESMeasuredsq[i]=RES[i]*RES[i]

    Sol=numpy.linalg.solve(A,b)
    RES=b2-numpy.dot(Al,Sol)
    
    Ainv=numpy.linalg.inv(A)
    Pv=numpy.diagonal(Ainv)
    C=numpy.sum(RESMeasuredsq)/(N1+N3-N2-1-1)
    PvC=C*Pv
    
#    if UL==1:
#       Drift=numpy.zeros((Nloops,1))
#       for i in range(0,int(Nloops)):
#           Drift[i]=Sol[int(N2+2*i+2-1)]
#    else:
#        Drift=Sol[int(N2+1)]
#    return [Sol,RES,PvC,Drift]
    if UL==1:
       Drift=numpy.zeros((Nloops,1))
       BL=numpy.zeros((Nloops,1))
       for i in range(0,int(Nloops)):
           Drift[i]=Sol[int(N2+2*i+2-1)]
           BL[i]=Sol[int(N2+2*i)]
           
    else:
        Drift=Sol[int(N2+1)]
        BL=Sol[int(N2)]
    return [Sol,RES,PvC,Drift,BL]
# Method one not decoupled - reweight equations by Tie variance and variance of other stations  
# With Beta
def GsolverM1Beta(TideFreeCalDial,SurveyedNames,SiteNames,TIM,TIENames,TIEVals,Loops,UL,CalDials):
    import numpy
    N1=numpy.size(TideFreeCalDial)
    N2=numpy.size(SiteNames) 
    N3=numpy.size(TIENames)
    Nloops=numpy.max(numpy.unique(Loops))
    SLoops=numpy.unique(Loops)
    P=numpy.zeros((N1,N2))
    if UL==1:
        T=numpy.zeros((N1,2*Nloops+1))    
    else:
        T=numpy.zeros((N1,2+1))    
    Q=numpy.zeros((N2,N2))
    h=numpy.zeros((N2,1))
    x=numpy.zeros((N1,1))
    for i in range(0,N1):
        x[i,0]=TideFreeCalDial[i]
        for j in range(0,N2):
            if SurveyedNames[i]==SiteNames[j]:
                P[i,j]=1
                x[i,0]=TideFreeCalDial[i]
        if UL==1:
            for k in range(0,int(Nloops)):
                if Loops[i]==SLoops[k]:
                    T[i,2*k]=1
                    T[i,2*k+1]=float(TIM[i])
            T[i,2*Nloops]=float(CalDials[i])

        else:
            T[i,0]=1
            T[i,1]=float(TIM[i])
            T[i,2]=float(CalDials[i])
            
        
    for k in range(0,N2):
        for j in range(0,N3):
            if SiteNames[k]==TIENames[j]:
                Q[k,k]=1
                h[k]=TIEVals[j]
    PdashP=numpy.dot(numpy.transpose(P),P)
    QdashQ=numpy.dot(numpy.transpose(Q),Q)
    PdashT=numpy.dot(numpy.transpose(P),T)
    TdashP=numpy.dot(numpy.transpose(T),P)
    TdashT=numpy.dot(numpy.transpose(T),T)
    Pdashx=numpy.dot(numpy.transpose(P),x)
    Qdashh=numpy.dot(numpy.transpose(Q),h)
    Tdashx=numpy.dot(numpy.transpose(T),x)
    A=numpy.hstack((numpy.vstack((PdashP+QdashQ,-TdashP)),numpy.vstack((-PdashT,TdashT))))
    if UL==1:
        Al=numpy.hstack((numpy.vstack((P,Q)),numpy.vstack((-T,numpy.zeros((N2,2*Nloops+1))))))    
    else:
        Al=numpy.hstack((numpy.vstack((P,Q)),numpy.vstack((-T,numpy.zeros((N2,2+1))))))    
 
    b=numpy.vstack((Pdashx+Qdashh,-Tdashx))
    b2=numpy.vstack((x,h))
    for i in range(0,N2+1):
        for j in range(0,N2+1):
            if numpy.array_equal(A[i,],A[j,]) and i!=j: 
                print((i,j))
        if numpy.sum(A[i,])==0:
            print[i]
    Sol=numpy.linalg.solve(A,b)
    RES=b2-numpy.dot(Al,Sol)
    # Calclualte Tie station variances and measurment variances
    Epps=numpy.zeros((N1,1))
    E=numpy.zeros((N3,1))  
    for i in range(0,N1):
        Epps[i]=RES[i]*RES[i]
    k=-1  
    for i in range(0,N2):
        if int(Q[i,i])==1:
            k=k+1
            E[k]=RES[i+N1]*RES[i+N1]   
    VEPPS=numpy.sum(Epps)/(N1-(N1*N2/(N3+N1)+1+2))
    VE=numpy.sum(Epps)/(N3-N3*N2/(N3+N1))
    # Re-weight LS equations and recompute solution
    #PdashP=1/VEPPS*numpy.dot(numpy.transpose(P),P)
    #QdashQ=1/VE*numpy.dot(numpy.transpose(Q),Q)
    #PdashT=1/VEPPS*numpy.dot(numpy.transpose(P),T)
    #TdashP=1/VEPPS*numpy.dot(numpy.transpose(T),P)
    #TdashT=1/VEPPS*numpy.dot(numpy.transpose(T),T)
    #Pdashx=1/VEPPS*numpy.dot(numpy.transpose(P),x)
   #Qdashh=1/VE*numpy.dot(numpy.transpose(Q),h)
   # Tdashx=1/VEPPS*numpy.dot(numpy.transpose(T),x)
   # A=numpy.hstack((numpy.vstack((PdashP+QdashQ,-TdashP)),numpy.vstack((-PdashT,TdashT))))
   # if UL==1:
     #   Al=numpy.hstack((numpy.vstack((P,Q)),numpy.vstack((-T,numpy.zeros((N2,2*Nloops+1))))))    
    #else:
     #   Al=numpy.hstack((numpy.vstack((P,Q)),numpy.vstack((-T,numpy.zeros((N2,2+1))))))    
 
    #b=numpy.vstack((Pdashx+Qdashh,-Tdashx))
    #b2=numpy.vstack((x,h))
    #for i in range(0,N2+1):
      #  for j in range(0,N2+1):
      #      if numpy.array_equal(A[i,],A[j,]) and i!=j: 
     #           print((i,j))
     #   if numpy.sum(A[i,])==0:
    #        print[i]
    #RES=b2-numpy.dot(Al,Sol)
    RESMeasured=numpy.zeros((N1,1))
    RESMeasuredsq=numpy.zeros((N1,1))
    for i in range(0,N1):
        RESMeasured[i]=RES[i]
        RESMeasuredsq[i]=RES[i]*RES[i]
#
    #Sol=numpy.linalg.solve(A,b)
    #RES=b2-numpy.dot(Al,Sol)
    
    Ainv=numpy.linalg.inv(A)
    Pv=numpy.diagonal(Ainv)
    C=numpy.sum(RESMeasuredsq)/(N1+N3-N2-1-2)
    PvC=C*Pv

    if UL==1:
        BetaCal=Sol[N2+2*Nloops]
    else:
        BetaCal=Sol[N2+2]    
    if UL==1:
        Drift=numpy.zeros((Nloops,1))
        BL=numpy.zeros((Nloops,1))
        for i in range(0,int(Nloops)):
            Drift[i]=Sol[int(N2+2*i+2-1)]
            BL[i]=Sol[int(N2+2*i)]
    else:
        Drift=Sol[int(N2+1)]
        BL=Sol[int(N2)]
    return [Sol,RES,PvC,BetaCal,Drift,BL]
   
#    if UL==1:
#       Drift=numpy.zeros((Nloops,1))
#       for i in range(0,int(Nloops)):
#           Drift[i]=Sol[int(N2+2*i+2-1)]
#    else:
#        Drift=Sol[int(N2+1)]
#    return [Sol,RES,PvC,BetaCal,Drift]
#    
    
def GsolverM3(TideFreeCalDial,SurveyedNames,SiteNames,TIM,TIENames,TIEVals,Loops,UL):
    import numpy
    N1=numpy.size(TideFreeCalDial)
    N2=numpy.size(SiteNames) 
    N3=numpy.size(TIENames)
    Nloops=numpy.max(numpy.unique(Loops))
    SLoops=numpy.unique(Loops)
    P=numpy.zeros((N1,N2))
    if UL==1:
        T=numpy.zeros((N1,2*Nloops))    
        K=numpy.zeros((N3,N2+2*Nloops))
    else:
        T=numpy.zeros((N1,2))    
        K=numpy.zeros((N3,N2+2))
    x=numpy.zeros((N1,1))
    for i in range(0,N1):
        x[i,0]=TideFreeCalDial[i]
        for j in range(0,N2):
            if SurveyedNames[i]==SiteNames[j]:
                P[i,j]=1
                x[i,0]=TideFreeCalDial[i]
        if UL==1:
            for k in range(0,int(Nloops)):
                if Loops[i]==SLoops[k]:
                    T[i,2*k]=1
                    T[i,2*k+1]=float(TIM[i])
        else:
            T[i,0]=1
            T[i,1]=float(TIM[i])
            
    c=numpy.zeros((N3,1))    
    for k in range(0,N2):
        for j in range(0,N3):
            if SiteNames[k]==TIENames[j]:
                K[j,k]=1
                c[j]=TIEVals[j]
    X=numpy.hstack((P,-T))
    # LHS matrix
    Ar1=numpy.hstack((numpy.dot(numpy.transpose(X),X),numpy.transpose(K)))
    Ar2=numpy.hstack((K,numpy.zeros((N3,N3))))         
    A=numpy.vstack((Ar1,Ar2))
    RHS=numpy.vstack((numpy.dot(numpy.transpose(X),x),c))
    Sol=numpy.linalg.solve(A,RHS)
    if UL==1:
        Beta=numpy.zeros((int(N2+2*Nloops),1))

        for k in range(0,int(N2+2*Nloops)):
            Beta[k]=Sol[k]
    else:
        Beta=numpy.zeros((int(N2+2),1))
        for k in range(0,int(N2+2)):
            Beta[k]=Sol[k]
            
    RES=x-numpy.dot(X,Beta)
    
    RESMeasured=numpy.zeros((N1,1))
    RESMeasuredsq=numpy.zeros((N1,1))
    for i in range(0,N1):
        RESMeasured[i]=RES[i]
        RESMeasuredsq[i]=RES[i]*RES[i]
    Ainv=numpy.linalg.inv(A)
    Pv=numpy.diagonal(Ainv)
    C=numpy.sum(RESMeasuredsq)/(N1+N3-N2-1-1)
    PvC=C*Pv
    
#    if UL==1:
#       Drift=numpy.zeros((Nloops,1))
#       for i in range(0,int(Nloops)):
#           Drift[i]=Sol[int(N2+2*i+2-1)]
#    else:
#        Drift=Sol[int(N2+1)]
#    return [Sol,RES,PvC,Drift]
    if UL==1:
       Drift=numpy.zeros((Nloops,1))
       BL=numpy.zeros((Nloops,1))
       for i in range(0,int(Nloops)):
           Drift[i]=Sol[int(N2+2*i+2-1)]
           BL[i]=Sol[int(N2+2*i)]
           
    else:
        Drift=Sol[int(N2+1)]
        BL=Sol[int(N2)]
    return [Sol,RES,PvC,Drift,BL]

def GsolverM3Beta(TideFreeCalDial,SurveyedNames,SiteNames,TIM,TIENames,TIEVals,Loops,UL,CalDials):
    import numpy
    N1=numpy.size(TideFreeCalDial)
    N2=numpy.size(SiteNames) 
    N3=numpy.size(TIENames)
    Nloops=numpy.max(numpy.unique(Loops))
    SLoops=numpy.unique(Loops)
    P=numpy.zeros((N1,N2))
    if UL==1:
        T=numpy.zeros((N1,2*Nloops+1))    
        K=numpy.zeros((N3,N2+2*Nloops+1))
    else:
        T=numpy.zeros((N1,2+1))    
        K=numpy.zeros((N3,N2+2+1))
    x=numpy.zeros((N1,1))
    for i in range(0,N1):
        x[i,0]=TideFreeCalDial[i]
        for j in range(0,N2):
            if SurveyedNames[i]==SiteNames[j]:
                P[i,j]=1
                x[i,0]=TideFreeCalDial[i]
        if UL==1:
            for k in range(0,int(Nloops)):
                if Loops[i]==SLoops[k]:
                    T[i,2*k]=1
                    T[i,2*k+1]=float(TIM[i])
            T[i,2*Nloops]=float(CalDials[i])
        else:
            T[i,0]=1
            T[i,1]=float(TIM[i])
            T[i,2]=float(CalDials[i])
            
    c=numpy.zeros((N3,1))    
    for k in range(0,N2):
        for j in range(0,N3):
            if SiteNames[k]==TIENames[j]:
                K[j,k]=1
                c[j]=TIEVals[j]
    X=numpy.hstack((P,-T))
    # LHS matrix
    Ar1=numpy.hstack((numpy.dot(numpy.transpose(X),X),numpy.transpose(K)))
    Ar2=numpy.hstack((K,numpy.zeros((N3,N3))))         
    A=numpy.vstack((Ar1,Ar2))
    RHS=numpy.vstack((numpy.dot(numpy.transpose(X),x),c))
    Sol=numpy.linalg.solve(A,RHS)
    if UL==1:
        Beta=numpy.zeros((int(N2+2*Nloops+1),1))

        for k in range(0,int(N2+2*Nloops+1)):
            Beta[k]=Sol[k]
    else:
        Beta=numpy.zeros((int(N2+2+1),1))
        for k in range(0,int(N2+2+1)):
            Beta[k]=Sol[k]
            
    RES=x-numpy.dot(X,Beta)
    
    RESMeasured=numpy.zeros((N1,1))
    RESMeasuredsq=numpy.zeros((N1,1))
    for i in range(0,N1):
        RESMeasured[i]=RES[i]
        RESMeasuredsq[i]=RES[i]*RES[i]
    Ainv=numpy.linalg.inv(A)
    Pv=numpy.diagonal(Ainv)
    C=numpy.sum(RESMeasuredsq)/(N1+N3-N2-1-2)
    PvC=C*Pv
    if UL==1:
        BetaCal=Sol[N2+2*Nloops]
    else:
        BetaCal=Sol[N2+2]    
    if UL==1:
       Drift=numpy.zeros((Nloops,1))
       BL=numpy.zeros((Nloops,1))
       for i in range(0,int(Nloops)):
           Drift[i]=Sol[int(N2+2*i+2-1)]
           BL[i]=Sol[int(N2+2*i)]
           
    else:
        Drift=Sol[int(N2+1)]
        BL=Sol[int(N2)]
    return [Sol,RES,PvC,BetaCal,Drift,BL]    
#    if UL==1:
#       Drift=numpy.zeros((Nloops,1))
#       for i in range(0,int(Nloops)):
#           Drift[i]=Sol[int(N2+2*i+2-1)]
#    else:
#        Drift=Sol[int(N2+1)]
#    
#    return [Sol,RES,PvC,BetaCal,Drift] 
   