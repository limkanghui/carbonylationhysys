import numpy as np
import pandas as pd
import itertools, math
import time
import openpyxl

from others import create_excel_file, print_df_to_excel


class CSTR:
    def __init__(self, Hycase, reactor_name, sprd_name):
        self.Hycase = Hycase
        self.Reactor = Hycase.Flowsheet.Operations.Item(reactor_name)

        #Decision Variables
        self.spreadsheetdata = Hycase.Flowsheet.Operations.Item(sprd_name)
        self.inlettemp = self.spreadsheetdata.Cell('B2').CellValue
        self.catalystweight = self.spreadsheetdata.Cell('B3').CellValue
        self.residencetime = self.spreadsheetdata.Cell('B4').CellValue
        self.reactorP = self.spreadsheetdata.Cell('B5').CellValue

        #Constraints
        self.vaporFrac = self.spreadsheetdata.Cell('D2').CellValue

        #Other variables
        self.duty = self.spreadsheetdata.Cell('B9').CellValue
        self.beforeinlettemp = self.spreadsheetdata.Cell('B10').CellValue
        self.reactorsize = self.spreadsheetdata.Cell('B11').CellValue
        self.reactortemp = self.spreadsheetdata.Cell('B12').CellValue

        #Objective
        self.conversion = self.spreadsheetdata.Cell('D5').CellValue
        self.MFproduction = self.spreadsheetdata.Cell(D6).CellValue

    def solve_reactor(self, inlettemp, catatlystweight, residencetime, reactorP):
        self.inlettemp = inlettemp
        self.catalystweight = catatlystweight
        self.residencetime = residencetime
        self.reactorP = reactorP

    def reactor_results(self):
        # Electricity cost for heating/cooling
        if self.beforeinlettemp < self.inlettemp:
            #Heating is required
            cost_of_heating = 0.10*abs(self.duty)*0.000277778 #cost of heating per hour

        else:
            cost_of_cooling = 0.02*abs(self.duty)*0.000277778 #cost of cooling per hour

    def reactor_cost(self,):
        # CSTR modelled as a pressure vessel
        # Costing based on Towler's Book
        operatingtemp = self.reactortemp
        operatingP = self.reactorP

        # Design Pressure
        pressureinpsig = operatingP*0.145038-14.7
        if pressureinpsig >= 0 & pressureinpsig <= 10:
            designP = 10
        elif pressureinpsig > 10 & pressureinpsig <= 1000:
            designP = math.exp(0.60608+0.91615*np.log(operatingP)+0.0015655*np.log(operatingP)**2)
        else:
            designP = operatingP*1.1

        # Design Temperature from Turton
        designTemp = operatingtemp + 25 # in degree celsius

        # Maximum Allowable Stress

        designTemp_in_F = designTemp * (9/5) + 32
        if designTemp_in_F >= -20 & designTemp_in_F <= 650:
            # Use carbon steel, SA-285, grade C
            maxstress = 13750 # in psi
        elif designTemp_in_F > 650 & designTemp_in_F <= 750:
            # Use low-alloy (1% Cr and 0.5% Mo) steel, SA-387B
            maxstress = 15000 # in psi
        elif designTemp_in_F > 750 & designTemp_in_F <= 800:
            # Use low-alloy (1% Cr and 0.5% Mo) steel, SA-387B
            maxstress = 14750 # in psi
        elif designTemp_in_F > 800 & designTemp_in_F <= 850:
            # Use low-alloy (1% Cr and 0.5% Mo) steel, SA-387B
            maxstress = 14200 # in psi
        elif designTemp_in_F > 850 & designTemp_in_F <= 900:
            # Use low-alloy (1% Cr and 0.5% Mo) steel, SA-387B
            maxstress = 13100 # in psi

        # Weld Efficiency
        if





