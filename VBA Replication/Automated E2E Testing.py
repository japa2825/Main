import num
py as np
np.set_printoptions(suppress=True) #Disable scientific notations, which makes thins obscure.
from scipy.stats import norm
import datetime


def BSOptionPrice(S_0, K, r, t, T, sigma, cp):
    if T <= t:
        raise ValueError("Expiry must be later than trade!") ##Newly added feature that corresponds to the Excel error cases.
    if sigma < 0:
        raise ValueError("Volatility must be non-negative!")
    r_cont = np.log(1 + r)  #convert the annual rate to the continuously compounding rate
    tau = (T - t).days / 365 #convert day count into a fraction of years as the time to maturity
    d1 = (np.log(S_0 / K) + (r_cont + 0.5 * sigma ** 2) * tau) / (sigma * np.sqrt(tau))
    d2 = (np.log(S_0 / K) + (r_cont - 0.5 * sigma ** 2) * tau) / (sigma * np.sqrt(tau))
    if cp == "Call":
        return S_0  * norm.cdf(d1) - K * norm.cdf(d2) * np.exp(-r_cont * tau)
    if cp == "Put":
        return -S_0 * norm.cdf(-d1) + K * norm.cdf(-d2) * np.exp(-r_cont * tau)
    else:
        raise ValueError("Field 'cp' can only be 'Call' or 'Put' (case sensitive).")
        

import random
import xlwings as xw
import numpy as np
xw.App().visible = False ##Suppresses actually opening Excel files.
bk = "Option and VaR.xlsx"
wb = xw.Book(bk) #Load the workbook.
ws = wb.sheets[0] #The worksheet "Option".

base_date = datetime.date(2023, 1, 1)##I pick random days by adding random integers to a certain date.

import unittest
class test_calc(unittest.TestCase):
    def test_E2E(self):
        for i in range(1000):
            print(f'Test Case {i+1} begins.')
            trade_delta = random.randint(0, 360)
            trade_date = base_date + datetime.timedelta(days = trade_delta) ##Randomly set a trade date.
            expiry_delta = random.randint(0, 1800)
            expiry_date = base_date + datetime.timedelta(days = expiry_delta) ##Here we give a ~10% chance of expiry date no later than trade date, so that the error cases are also included. If the two dates are identical, cell C22 will have a dividing-by-zero problem.
            ws.range('C4').value = trade_date ##Rewrite the value on the cell C4. Note the formula cells automatically update.
            ws.range('C5').value = expiry_date
            S_0 = np.random.poisson(100) ##Set spot to follow Poisson distribution of mean 100.
            ws.range('C6').value = S_0
            K = np.random.geometric(0.01) ##Set strike to follow geometric distribution of mean 100.
            ws.range('C7').value = K
            r = np.random.normal(0.05, 0.02) ##Interest rate to follow some normal distribution. Note it could be negative.
            ws.range('C9').value = r
            sigma = np.random.uniform(0.01, 0.5) ##Volatility to follow uniform distribution.
            ws.range('C12').value = sigma
            cp = random.choice(['Call', 'Put']) ##Randomly select a cp label for the above BS formula.
            
            if cp == 'Call':
                val_excel = ws.range('C27').value ##Save the output from Excel if option is call.
            if cp == 'Put':
                val_excel = ws.range('C32').value
            
            try:
                val_py = BSOptionPrice(S_0, K, r, trade_date, expiry_date, sigma, cp) ##Compute option price with Python if expiry is later than trade. 
                self.assertAlmostEqual(val_py, val_excel, delta = 0.00000001) ##Compare the results, with error tolerance delta.
                #If the above assert function is called smoothly it means Python is consistent with Excel on the regular case, i.e., trade date earlier than expiry date.
            except:
                with self.assertRaises(ValueError) as ve:
                    BSOptionPrice(S_0, K, r, trade_date, expiry_date, sigma, cp) ##Note if expiry is earlier than trade, error message will be raised and the output has numpy.nan type. On the other hand, the output on Excel is an integer -2xxxxxxxxx in place of a non-value.
                    self.assertEqual(str(ve.exception), "Expiry must be later than trade!") ##Error message for Python.
                if isinstance(val_excel, (float, int)): ##Things bifurcate here. On different devices, it has been observed that "ErrValue" on Excel, once captured by xlwings, can either be None type or a 10-digit negative integer like -252882552 on Python. The reason can either be different versions of Spyder, Excel, or xlwings package.
                    self.assertLess(val_excel, 0) ##If numeric, then it means xlwings thinks the Excel output is a negative number -- which cannot be true under normal conditions.
                else:
                    self.assertIsNone(val_excel) ##Otherwise, the value is of None type. 
                #If both of the above two assert functions are called successfully, it means Python and Excel agree on the case trade date not earlier than expiry date.
                
            print(f'Test Case {i+1} ends.')
            
if __name__ == "__main__":
    unittest.main()