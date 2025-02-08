import xlwings as xw
from QC import quantile_converter
import pandas as pd
import numpy as np
import unittest

horizon = 1 #Quantity in the assignemnt's formula.
temp_xlsm = 'temp.xlsm' #The path of the macro enabled excel file (in Python default dir).
temp_csv = 'temp.csv' #The path of the .csv file (same directory).

#This function gets the Python 1%-VaR from Excel worksheets, in two methods configed by QC.
def pyed_var(path, csv_path, QC = False):
    app = xw.App(visible=False)  # Ensure Excel is not visible to avoid interference
    app.display_alerts = False  # Disable any alerts from Excel (like read-only)
    app.screen_updating = False  # Turn off screen updating for performance
    # Open the workbook
    df = pd.read_excel(path)
    df.to_csv(csv_path)
    
    df_csv = pd.read_csv(csv_path) #.csv files work better on dataframes.
    asset_1 = float(df_csv.iloc[1, 3])
    asset_2 = float(df_csv.iloc[2, 3])

    df_csv = df_csv.iloc[5:, 6:8] 
    #FX rates are on the F,G columns, beginning from the seventh row--
    #(pd treats the first row as header and doesn't deem it as a row).

    rate_1 = df_csv.iloc[:, 0].values #This converts a column to an np array.
    rate_2 = df_csv.iloc[:, 1].values
    quotient_1 = rate_1[: -1] / rate_1[1:] #The quotient up_cell/down_cell.
    quotient_2 = rate_2[: -1] / rate_2[1:]
    return_1 = np.exp(np.log(quotient_1) * np.sqrt(horizon)) - 1 
    #The daily pnl of ccy_1 asset per unit of domestic currency.
    return_2 = np.exp(np.log(quotient_2) * np.sqrt(horizon)) - 1
    pnl = return_1 * asset_1 + return_2 * asset_2 
    #The total daily pnl of the portfolio in domestic currency.
    
    if QC == False: #Applying the traditional approach.
        arr = np.sort(pnl)
        val = 0.4 * arr[1] + 0.6 * arr[2]
    else: #Applying the np.quantile function with converted quantile of 0.01.
        qc = quantile_converter(len(pnl), 0.01) #0.006201550387596899
        val = np.quantile(pnl, qc)
        
    app.quit() 
    #Terminates the Excel application, so that reopening the file does not cause problems.
    return val

#This generates the input data for Python and VBA to compute VaR.
def sampling_fx_rates(path): 
    #Each time the function runs, a sample portfolio of two FX assets is built.
    #It serves as the input of the two methods of VaR calculation,
    #and the output will have the same shape and position as the asset price (core) part-- 
    #in the assignment sheet.
    app = xw.App(visible=False)  # Ensure Excel is not visible to avoid interference
    app.display_alerts = False  # Disable any alerts from Excel (like read-only)
    app.screen_updating = False  # Turn off screen updating for performance

    # Open the workbook
    workbook = app.books.open(path)
    sheet = workbook.sheets[0] #The temp file has only one sheet named "VaR Calculation".
    sheet.clear_contents() #Start from a blank sheet.
    
    sheet.range('C3').value = np.random.uniform(10000, 100000) #Rebase the portfolio.
    sheet.range('C4').value = np.random.uniform(10000, 100000)
    sheet.range('F7').value = np.random.normal(1, 0.08) 
    #Sample the fx rates on the last day.
    sheet.range('G7').value = np.random.normal(120, 30)
    

    for i in range(259):
        w1, w2 = np.random.normal(1, 0.01, 2) #Set the daily multiplicative variation.
        upper_cells = 'F' + str(7 + i), 'G' + str(7 + i)
        lower_cells = 'F' + str(7 + i + 1), 'G' + str(7 + i + 1)
        sheet.range(lower_cells[0]).value = sheet.range(upper_cells[0]).value * w1 
        #Here we are simulating the FX rates for the remaining days in the reverse order. 
        #The second last day's price is assumed to be the last day's price times--
        #(1 + epsilon), where epsilon = some random walk. 
        #This is NOT essential -- the user can choose any dynamic for the FX rates.
        sheet.range(lower_cells[1]).value = sheet.range(upper_cells[1]).value * w2
    
    workbook.save() #Save the worksheet, and turn off the Excel app for the next step.
    workbook.close()
    app.quit()
    return

#This function obtains the macro's result of 1%-VaR.
def macroed_var(path):
    app = xw.App(visible=False)  # Ensure Excel is not visible to avoid interference
    app.display_alerts = False  # Disable any alerts from Excel (like read-only)
    app.screen_updating = False  # Turn off screen updating for performance
    
    workbook = app.books.open(path)
    macro = workbook.macro("Calculate_1pct_VaR") #Name of the macro module.
    macro() #Run the macro, which writes the 1%-VaR on Cell O6, same as the assignment.
    
    worksheet = workbook.sheets[0]
    VaR_vba = worksheet.range('O6').value 
    #The result from the macro, against which we compare the Python computation result.
    
    workbook.save() #Save the worksheet, and turn off the Excel app for the next step.
    workbook.close()
    app.quit()
    return VaR_vba

#E2E testing on the consistency between Py (both old and new methods) and VBA.
class test_calc(unittest.TestCase):
    # def test_macro_vs_py(self): #Python approach directly computes the 0.4/0.6 combination.
    #     path_temp = temp_xlsm
    #     csv_path = temp_csv
    #     for i in range(1, 1001):
    #         if i % 100 == 0:
    #             print(f"Test 1 Case {i} begins.")
    #         sampling_fx_rates(path_temp) #Shuffle the input data.
    #         VaR_macro = macroed_var(path_temp) #Save the macro result.
    #         VaR_py = pyed_var(path_temp, csv_path) #Save the Python result.
    #         self.assertAlmostEqual(VaR_macro, VaR_py, delta = 0.00000001) #Compare.
    #         if i % 100 == 0:
    #             print(f"Test 1 Case {i} ends.")
                
    def test_macro_vs_py_with_QC(self): #Python uses the function quantile_converter.
        path_temp = temp_xlsm
        csv_path = temp_csv
        for i in range(1, 1001):
            if i % 100 == 0:
                print(f"Test 2 Case {i} begins.")
            sampling_fx_rates(path_temp)
            VaR_macro = macroed_var(path_temp)
            VaR_py = pyed_var(path_temp, csv_path, QC = True) #Use quantile_converter.
            self.assertAlmostEqual(VaR_macro, VaR_py, delta = 0.00000001)
            if i % 100 == 0:
                print(f"Test 2 Case {i} ends.")


if __name__ == "__main__":
    unittest.main(buffer = False) #This prints the testing progress messages orderly.