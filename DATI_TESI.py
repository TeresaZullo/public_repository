import numpy as np
import scipy.stats as spss
import scipy.optimize as spopt
import seaborn as sns
from QuantLib import *
import matplotlib.pyplot as plt
import math
import xlwings as xw
import pandas as pd

# prendo la curva in tempo reale'

def GetData():
    wb = xw.Book(r'C:\Users\T004697\Desktop\TESI\DATI _TESI.xlsx')
    sh1 = wb.sheets['Eur6m']
    sh1 = wb.sheets['ESTR']
    sh2 = wb.sheets['Vol']
  
    #curve
    df1 = sh1.range('A1').options(pd.DataFrame, expand='table').value
    df2 = sh2.range('A1').options(pd.DataFrame, expand='table').value


    
    Eur6m = DiscountCurve([Date.from_date(d) for d in df1.index.to_list()],
                                                         df1['CURVE_QUOTE'].to_list(),
                                                         Actual360())
    ESTR = DiscountCurve([Date.from_date(d) for d in df2.index.to_list()],
                                                         df2['CURVE_QUOTE'].to_list(),
                                                         Actual360())

   # abilito estrapolazione per entrambe le curve
    eur6m_curve.enableExtrapolation()
    estr_curve.enableExtrapolation()   
    
    # Crea gli handle delle curve
    eur6m_handle = RelinkableYieldTermStructureHandle(eur6m_curve)
    estr_handle = RelinkableYieldTermStructureHandle(estr_curve)
    
    return eur6m_handle, estr_handle

    # Utilizzo della funzione GetData
    eur6m_handle, estr_handle = GetData()

'generalizzo formula Annuity'
'Sulla base dell eq 28 è necessario annuity di R1, annuity di R2, annuit ymidcurve ' 

class Annuity:
    def __init__(self, start_date, end_date, discount_curve, frequency=Semiannual):
        self.start_date = start_date
        self.end_date = end_date
        self.discount_curve = discount_curve
        self.frequency = frequency

    def calculate_annuity(self):
        schedule = Schedule(self.start_date, self.end_date, Period(self.frequency), NullCalendar(),
                               ModifiedFollowing, ModifiedFollowing, DateGeneration.Forward, False)
        
        annuity = 0.0
        for i in range(1, len(schedule)):
            delta_t = Actual360().yearFraction(schedule[i-1], schedule[i])
            discount_factor = self.discount_curve.discount(schedule[i])
            annuity += delta_t * discount_factor
        
        return annuity

# Data di valutazione
evaluation_date = Date(10, 7, 2024)
Settings.instance().evaluationDate = evaluation_date

# Definizione delle date per rate1, rate2 e mid-curve
rate1_start = Date(10, 7, 2024)
rate1_end = Date(10, 7, 2025)

rate2_start = Date(10, 7, 2025)
rate2_end = Date(10, 7, 2027)

mid_curve_start = Date(10, 7, 2024)
mid_curve_end = Date(10, 7, 2027)

# Creazione delle istanze Annuity
annuity_rate1 = Annuity(rate1_start, rate1_end, eur6m_handle)
annuity_rate2 = Annuity(rate2_start, rate2_end, eur6m_handle)
annuity_mid_curve = Annuity(mid_curve_start, mid_curve_end, estr_handle)

def calculate_annuity(start_date, end_date, discount_curve, frequency = Semiannual):
    
    
   schedule = Schedule(start_date, end_date, Period(frequency), NullCalendar(),
                          ModifiedFollowing, ModifiedFollowing, DateGeneration.Forward, False)
   
   annuity = 0.0
   for i in range(1, len(schedule)):
       delta_t = Actual360().yearFraction(schedule[i-1], schedule[i])
       discount_factor = discount_curve.discount(schedule[i])
       annuity += delta_t * discount_factor
   
   return annuity


# @xw.sub

#     "generalizziamo la simulazione montecarlo per n basket con 2 elementi"
#     'dobbiamo creare 4 random variables'

# 'payoff blackbasket per midcurve'
# midcurve_price=delta_1*R1 + delta_2*R2 - K
# K = k - R_mc

# 'definiamo i coefficienti come ratio tra annuities'
# delta_1 = -(annuity_1/annuity_mc) 
# delta_2 = annuity_2/annuity_mc

# def calculate_annuity(rate, n_periods, frequency):
#     """

