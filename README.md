# BTS-Data-Anlysis-Programs
Useful python programs (Python 2.7) for analyzing electrochemical cell performance data exported from BTSDA software, a common software used  by energy storage device researchers. Programs allow users to select mulptiple files at once and compare the performance of different cells.

Example data provided as TestData1, TestData2, TestData3.

Programs capable of analyzing excel or text files exported from BTSDA software cotaining the discharge capacity values for each charge/discharge cycle.

Required Python Libraries/Modules: Numpy, Pandas, Openpyxl, Tkinter, TKconstants, TkFileFialog, Os, Collections, Operator

For either program, user must specify number of cycles to be analyzed across all selected data files.

Sort_by_Stability.py allows users to sort the performance data of selected cells by the average percent decrease in discharge capacity per charge/discharge cycle. Upon running the program, a dialog box will appear and allow users to select the data file for each cell to be analyzed. The discharge capacity and coulombic efficiency of the first cycle for each cell are also presented in the output excel file, "sort_by_stability.xlsx" 

Compare_Discharge_Capacities.py allows users to compare the discharge capacity values for each cycle for all selected cells. Upon running the program, a dialog box will appear and allow users to select the data file for each cell to be analyzed. The results will appear in the output excel file, "compare_discharge_capacities.xlsx" 
