# -*- coding: utf-8 -*-
"""
Created on Sun Aug 24 16:23:30 2014

@author: Alexander van Kemenade
"""

import statsmodels.api as sm
import pandas as pd
#import numpy as np
import itertools as iter
#from operator import itemgetter
from patsy import dmatrices
import pickle
import math


def excelDF(filepath):
    """Returns a pandas dataframe object from Sheet1 of Excel file"""
    print "Creating dataframe..."    
    df = pd.io.excel.read_excel(filepath)
    pickle.dump( df, open( "save.p", "wb" ) )
    print "Dataframe created from Excel file"
    return df

def varcombs(filepath):
    """Returns list of series lables of dependent and independent variables"""    
    df = excelDF(filepath)
    vars = list(df.columns.values)
    print len(vars),"variables loaded"
    indep = vars[1:]  
    comb = (iter.combinations(indep, l) for l in range(1,len(indep) + 1))
    print "Variable combinations established"
    combs = list(iter.chain.from_iterable(comb))
    return combs, vars

def specify(filepath):
    "Returns a string with a specified equation for estimation"    
    combs, vars = varcombs(filepath)
    equations = []
    numStrings = 0
    for comb in combs:
        equation = str(vars[0]+" ~ ")
        for var in comb:
            equation += var+" + "
        equation = equation[:len(equation)-3]
        equations.append(equation)
        numStrings += 1
    print numStrings, "equation strings specified"
    return equations
    
def estimate(equation,data):
    """Takes an estimable equation as a string and returns OLS result
    s"""
    y, X = dmatrices(equation, data = data, return_type = 'dataframe')
    mod = sm.OLS(y, X)
    res = mod.fit()
    print res.summary()
#    return 
#    return res.resid
#    print res.fittedvalues
#    del y, X
#    print equation, "gives R-squared of", res.rsquared
#    return res
        
def getresids():
    df = excelDF('C:\Users\Joe\Documents\Work\Daimler\monthly.xlsx')
    df2 = excelDF('C:\Users\Joe\Documents\Work\Daimler\FAI_NSA.xlsx')
    for s in list(df2)[1:3]:
        print type(df2[s])
        df2["log"] = math.log(df2[s])
        equation = str(s + " ~ d1+d2+d3+d4+d5+d6+d7+d8+d9+d10+d11+d12")
        df[s] = estimate(equation,df2)
    writer = pd.ExcelWriter('C:\Users\Joe\Documents\Work\Daimler\output.xlsx')
    df.to_excel(writer,'Sheet1')
    writer.save()
    return df

#def maxRsquared(filepath,sheet):
#    """Takes an Excel table as an argument, takes the first data column as the 
#    dependent variable and returns results of the R-squared-maximising 
#    specification"""
#    equations = specify(filepath)
#    df = pickle.load(open("save.p", "rb"))
#    numEquations = 1
#    print "Running estimations...\n"
#    first = estimate(equations[0],df)
#    highest = first, first.rsquared    
#    for equation in equations[1:]:
#        res = estimate(equation,df)
#        r2 = res.rsquared
#        if r2 > highest[1]:
#            highest = res, r2
#        numEquations += 1
#    print "Estimated", numEquations, "equations"
#    print
#    print "R-squared maximising specification:\n"
#    print 
#    print highest[0].summary()maxRsquared("C:\Users\Joe\Documents\Python\OLS\OLS.xlsx")
    
#estimate("rev ~ l_visits_premium + l_clothing + l_malls + l_nndoor",df)
    
#estimate("rev ~ tz_area + l_malls + l_shopping + l_nndoor +  timesector_14_17y + Affordable_y + young + ISLOCAL_0_y + d_t1 + d_t2 + d_t3",df)

#estimate("rev ~ l_malls + l_shopping + l_nndoor + timesector_14_17y + young + Affordable_y + FEMALE_y",priority)
#estimate("rev ~ l_malls + l_shopping + l_nndoor + timesector_14_17y + young + FEMALE_y + Premium_y + d_beijing + d_shanghai",top)
#estimate("rev ~ l_malls + l_clothing + l_ndoor + sellspace + timesector_14_17y + FEMALE_y + d_beijing + d_shanghai",all)
#estimate("rev ~ l_visits + timesector_14_17y + young + Affordable_y + d_beijing + d_shanghai",all)