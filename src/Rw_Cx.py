# -*- coding: utf-8 -*-
"""
Created on Thu Jul 28 15:55:16 2016
Calculate the minimum required Rw+Ctr or Rw+C
@author: Weigang Wei
"""
import numpy as np
import random
import matplotlib.pylab as plt


class C_Ctr():
    def __init__(self):
        """ the ranges for C and Ctr testing are based on the Pilkinton and Guardian
            test data.
        """
        self.C = np.asarray([-21,-14,-8,-5,-4])
        self.Ctr = np.asarray([-14,-10,-7,-4,-6])
        self.CRange = np.asarray([6,5,5,11,9])/2. # 125 to 2k average deviation
        self.CtrRange = np.asarray([6,4,6,11,12])/2.# 125 to 2k average deviation  
    
class CalculationInit():
    def __int__(self, refSpec=[0.,0.,0.,0.,0.], vairation=[6.,5.,6.,11.,11.]):
        self.refSpec = refSpec
        self.variation = vairation
        
class CalcSNRRandom(C_Ctr, CalculationInit):
    def __init__(self, sourceSpec,V, S, T, n,L2LimitWin=30, L2LimitVent=30, refSpec=[0.,0.,0.,0.,0.], vairation=[6.,5.,6.,11.,11.]):
        C_Ctr.__init__(self)
        CalculationInit.__int__(self, refSpec, vairation)
        self.sp = sourceSpec
        self.L2LimitWin = L2LimitWin
        self.L2LimitVent = L2LimitVent
        self.V = V
        self.S = S
        self.T = T
        self.n = n
        self.NUM = 10000
        self.Deltai_Ctr = self.sp - self.Ctr
        self.Deltai_C = self.sp - self.C
        self.condi = 10.*np.log10(T) + 10.*np.log10(S/V) + 11
        self.condi2 = 10.*np.log10(T) + 10.*np.log10(n/V) + 21
        
    def _generate_L2_spec(self, L2limit):   
        print("L2,i variation: ", self.variation)
        specs = []
        for num in range(self.NUM):
            spec = []
            for s,v in zip(self.refSpec, self.variation):
                spec.append(round(random.uniform(s-v, s+v))) # use the closest integer
            
            specA = np.asarray(spec)
            total = 10.*np.log10(np.sum(10**(specA/10)))
            specA = specA - total + L2limit
            specs.append(specA)
        return specs
    
    def _run_test(self):
        RwCtr, RwC, DnewCtr,DnewC = [], [], [], []
        L2isWin = self._generate_L2_spec(self.L2LimitWin)
        L2isVent = self._generate_L2_spec(self.L2LimitVent)
#            
        for L2iw, L2iv in zip(L2isWin, L2isVent):
            vari  = 10.*np.log10(np.sum(10**((L2iw - self.Deltai_Ctr)/10)))
            var2 = 10.*np.log10(np.sum(10**((L2iw - self.Deltai_C)/10)))
            RwCtr += [self.condi - vari]
            RwC += [self.condi - var2]
            
            var3  = 10.*np.log10(np.sum(10**((L2iv - self.Deltai_Ctr)/10)))
            var4 = 10.*np.log10(np.sum(10**((L2iv - self.Deltai_C)/10))) 
            DnewCtr += [self.condi2 - var3]
            DnewC += [self.condi2 - var4]
#       
        
        RwCtr = np.sort(RwCtr)
        RwC = np.sort(RwC)
        DnewCtr = np.sort(DnewCtr)
        DnewC = np.sort(DnewC)
        for n, x in enumerate([RwCtr, RwC, DnewCtr, DnewC]):
            if n== 0:
                print("\nRw+Ctr: ")
            elif n==1:
                print("\nRw+C: ")
            elif n==2:
                print("\nDnew+Ctr: ")
            else:
                print("\nDnew+C: ")
            print("5% :        ", "%0.1f" %x[int(len(x)*0.95)])
            print("25% to 75%: ", "%0.2f" %(x[int(len(x)*0.75)] - x[int(len(x)*0.25)]))
        
        
        # print required Rw+Ctr, Rw+C
        plt.figure()
        bt = min(min(RwCtr), min(RwC))
        top = max(max(RwCtr), max(RwC))
        plt.boxplot([np.round(RwCtr), np.round(RwC)])
        plt.xticks([1,2],['Rw+Ctr','Rw+C'])
        plt.ylim([bt-5,top+5])
        plt.ylabel('dB')
        plt.grid()
#        plt.savefig('Statistics.png')
        
        plt.figure()
        plt.subplot(1,2,1)
        plt.hist(RwCtr, bins=2*int(max(RwCtr)-min(RwCtr)))
        plt.subplot(1,2,2)
        plt.hist(RwC, bins=2*int(max(RwC)-min(RwC)))
#        plt.savefig('density-function.png')
        
        # print required Dnew+Ctr, Dnew+C
        plt.figure()
        bt = min(min(DnewCtr), min(DnewC))
        top = max(max(DnewCtr), max(DnewC))
        plt.boxplot([np.round(DnewCtr), np.round(DnewC)])
        plt.xticks([1,2],['Dnew+Ctr','Dnew+C'])
        plt.ylim([bt-5,top+5])
        plt.ylabel('dB')
        plt.grid()
#        plt.savefig('Statistics.png')
        
        plt.figure()
        plt.subplot(1,2,1)
        plt.hist(DnewCtr, bins=2*int(max(DnewCtr)-min(DnewCtr)))
        plt.subplot(1,2,2)
        plt.hist(DnewC, bins=2*int(max(DnewC)-min(DnewC)))
#        plt.savefig('density-function.png')
        
    
if __name__=='__main__':
    sourceSpec = [39, 43, 47, 53, 45]
    V, S, T, n, L2LimitWin, L2LimitVent = 22.5, 4.1, 0.5, 1, 28, 30
    objtest = CalcSNRRandom(sourceSpec, V, S, T, n, L2LimitWin, L2LimitVent)
    objtest._run_test()
    plt.show()