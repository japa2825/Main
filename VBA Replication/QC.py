#This function converts a regulatory-defined quantile to the np-library-defined quantile.
#Specifically, for a sorted array of lenth n, regulatory-defined quantiles for the i-th entry is i/n.
#While in Version 1.26.4, np-defined quantile for the i-th entry is (i - 1)/(n - 1).
#If the wanted quantile is not any of the fractions, both interpolate linearly.
#Thus, for n = 260 and quantile q = 0.01, the 2nd and 3rd entries have quantiles 2/260 and 3/260 in the regulatory, respectively.
#So the expression 0.4 * 2nd + 0.6 * 3rd linearly yields the '2.6-th' entry in the array of 260, which has quantile 0.01.
#But since the np.quantile function computes with a different assignment of quantiles to the entries, we cannot directly use the function.
#So, we wish to have a function f, so that for n and regulatory-defined q, f(n, q) = the equivalent quantile in the np-context.
#The benefit is, instead of computing the 1%-VaR manually on the one-year PnL vector, we can directly compute with np.quantile(the PnL vector, f(0.01)).
#Disclamer: We note that the PnL vector's length is 259 instead of 260. Not having found any authoritative document that supports this convention, we reserve our doubt.
import numpy as np
import unittest
import random

def quantile_converter(n, q): #n = actual length of the array.
    if n < 10: #Due to the length inconsistency, we set this to avoid unnecessary errors.
        raise ValueError("The array is too short!")
    
    #Note that, our present methodology "pretends" there were n + 1 entries in the array, each of which has quantile i / (n + 1). Thus, when there are only n elements in fact, nothing less than 1 / (n + 1) or greater than n / (n + 1) can be interpolated from what we have.
    if q < 1 / (n + 1):
        raise ValueError("Nothing can have such a small quantile in the regulatory-defined formula!")
    if q > n / (n + 1): #This is a theoretical bound from the length inconsistency.
        raise ValueError("The quantile is too large!")
    length_in_theory = n + 1
    rank_in_need = length_in_theory * q #What's the "rank" of the (imagined) entry, that exactly has quantile q in the array?
    return (rank_in_need - 1) / (n - 1) #The np-quantile of the imaginary entry with the needed rank in the array of length (n + 1). If regulatory requires PnL vector to have length 260 instead of 259, the last (n + 1) will be replaced with n.


#We validate this quantile_converter function with two tests.
#First, we compute the 1%-VaR for randomly generated arrays of 259 numbers. As the result shows, in numpy's algorithms, the converted quantile 0.006201550387596899 for any vector of length 259, always equals to 0.4 * the second smallest + 0.6 * the third smallest elements.
#Second, we test on (almost) arbitary VaR and length, but only on the sequence of positive integers 1, 2, 3, ..., len. It can be easily proved, that the choice of entries does not affect the relation between the quantiles under different definitions.
#Setting the entries in this way simplifies the computation, as VaR's quantile times length equals the VaR. As the test shows, the conversion is accurate in all cases.
class TestCalc(unittest.TestCase):
    def test_qc_on_1y_pnl_vectors(self): 
        n = 260
        q = 0.01
        #Then the 1%-VaR will always be 0.4 * second smallest + 0.6 * third smallest in the regulatory defination.
        qc = quantile_converter(n - 1, q) #0.006201550387596899 -- recall it's the ACTUAL length
        for i in range(1, 1001):
            if i % 100 == 0:
                print(f"Test 1 Case {i} begins.")
                
            pnl = np.random.normal(0, 1000, 259)
            pnl = sorted(pnl)
            expected_result = 0.4 * pnl[1] + 0.6 * pnl[2]
            np_result = np.quantile(pnl, qc)
            self.assertAlmostEqual(np_result, expected_result, delta = 0.00000001)
        
            if i % 100 == 0:
                print(f"Test 1 Case {i} ends.")
                
    def test_qc_on_natural_numbers(self): #The wanted entries with wanted rank will equal to the rank itself, in the regulatory-defined quantile formula.
        for i in range(1, 1001):
            if i % 100 == 0:
                print(f"Test 2 Case {i} begins.")
            
            n = random.randint(0, 500) #Note that the array will be empty if n = 0, 1, and there should be a value error.
            q = random.uniform(0, 1)
            arr = np.array(range(1, n)) #If n = 260, arr will be from 1 to 259, and its 1%-VaR in out context will be n * q = 2.6.

            if n <= 10:
                # Test for too small n
                with self.assertRaises(ValueError) as ve:
                    quantile_converter(len(arr), q)
                    self.assertEqual(str(ve.exception), "The array is too short!")
            
            elif q < 1 / (len(arr) + 1):
                # Test for too small q
                with self.assertRaises(ValueError) as ve:
                    quantile_converter(len(arr), q)
                    self.assertEqual(str(ve.exception), "Nothing can have such a small quantile in the regulatory-defined formula!")

            elif q > len(arr) / (len(arr) + 1):
                # Test for too large q
                with self.assertRaises(ValueError) as ve:
                    quantile_converter(len(arr), q)
                    self.assertEqual(str(ve.exception), "The quantile is too large!")
           
            else:
                qc = quantile_converter(len(arr), q) #The converted quantile for the np function.
                np_result = np.quantile(arr, qc)
                expected_result = n * q 
                self.assertAlmostEqual(np_result, expected_result, delta = 0.00000001)
            
            if i % 100 == 0:
                print(f"Test 2 Case {i} ends.")
                
    
    
if __name__ == "__main__":
    unittest.main(buffer=False)
