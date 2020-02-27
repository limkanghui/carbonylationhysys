import numpy as np
import pandas as pd
import itertools, math
import time
import openpyxl
from scipy.optimize import fsolve

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
        self.MFproduction = self.spreadsheetdata.Cell('D6').CellValue

    def solve_reactor(self, inlettemp, catatlystweight, residencetime, reactorP):
        self.inlettemp = inlettemp
        self.catalystweight = catatlystweight
        self.residencetime = residencetime
        self.reactorP = reactorP

    def reactor_design(self,):
        # CSTR modelled as a pressure vessel
        # Costing based on Towler's Book
        operatingtemp = self.reactortemp
        operatingP = self.reactorP

        # Design Pressure in psig
        pressureinpsig = operatingP*0.145038-14.7
        if pressureinpsig >= 0 and pressureinpsig <= 10:
            designP = 10
        elif pressureinpsig > 10 and pressureinpsig <= 1000:
            designP = math.exp(0.60608+0.91615*np.log(operatingP)+0.0015655*np.log(operatingP)**2)
        else:
            designP = operatingP*1.1

        # Design Temperature from Turton
        designTemp = operatingtemp + 25 # in degree celsius

        # Maximum Allowable Stress

        designTemp_in_F = designTemp * (9/5) + 32
        if designTemp_in_F >= -20 and designTemp_in_F <= 650:
            # Use carbon steel, SA-285, grade C
            maxstress = 13750 # in psi
        elif designTemp_in_F > 650 and designTemp_in_F <= 750:
            # Use low-alloy (1% Cr and 0.5% Mo) steel, SA-387B
            maxstress = 15000 # in psi
        elif designTemp_in_F > 750 and designTemp_in_F <= 800:
            # Use low-alloy (1% Cr and 0.5% Mo) steel, SA-387B
            maxstress = 14750 # in psi
        elif designTemp_in_F > 800 and designTemp_in_F <= 850:
            # Use low-alloy (1% Cr and 0.5% Mo) steel, SA-387B
            maxstress = 14200 # in psi
        elif designTemp_in_F > 850 and designTemp_in_F <= 900:
            # Use low-alloy (1% Cr and 0.5% Mo) steel, SA-387B
            maxstress = 13100 # in psi




        # Assume cylindrical height to diameter is 3:1

        def internalDfunction(D):
            return math.pi*(D/2)**2*(3*D) - self.reactorsize

        Di = fsolve(internalDfunction, 0.01)

        # Shell thickness tp calculation
        shell_thickness0 = designP*Di*39.3701/(2*maxstress*0.85-1.2*designP)

        if shell_thickness0 <= 1.25:
            shell_thickness = shell_thickness0
        else:
            shell_thickness = designP*Di*39.3701/(2*maxstress*1-1.2*designP)

        if Di*3.28084 <= 4:
            shell_thickness = 1/4
        elif Di*3.28084 <= 6:
            shell_thickness = 5/16
        elif Di*3.28084 <= 8:
            shell_thickness = 3/8
        elif Di*3.28084 <= 10:
            shell_thickness = 7/16
        elif Di*3.28084 <= 12:
            shell_thickness = 1/2

        # Consider wind and earthquake for vertical column

        def twfunc(tw):
            return tw - 0.22*(Di+2*shell_thickness+tw+1/4+18)*((Di)*3)**2/(maxstress*(Di+2*shell_thickness+tw+1/4)**2)

        tw_solved = fsolve(twfunc, 0.2)
        tv = (2*shell_thickness+tw_solved)/2
        tc = 1/8
        ts = tv + tc

        if ts >= 3/16 and ts <= 1/2:
            ts = math.ceil(ts/(1/6))
        elif ts >= 5/8 and ts <= 2:
            ts = math.ceil(ts/(1/8))
        elif ts >= 9/4 and ts <= 3:
            ts = math.ceil(ts/(1/4))

        # weight of vessel
        # where ρ is the density of carbon steel SA-285 grade C which is 0.284 lbm/in3

        weight = math.pi*(Di+ts)*(3*Di+0.8*Di)*ts*0.284

        return ts, weight, Di

    def reactor_cost(self):
        # Using Reactor-mixer values from Appendix A of turton's tb
        # Volume between 0.04 and 60 m3

        k1 = 4.7116
        k2 = 0.4479
        k3 = 0.0004

        S = self.reactorsize

        cp0_2001 = 10**(k1+k2*math.log10(S)+k3*(math.log10(S))**2)
        cp0_2018 = cp0_2001*603.1/397

        # Pressure factor for process vessels
        operatingP = self.reactorP
        pressureinpsig = operatingP * 0.145038 - 14.7
        pressureinbarg = pressureinpsig * 0.0689476
        ts, weight, Di = self.reactor_design()

        if pressureinbarg <= -0.5:
            Fp = 1.25
        elif pressureinpsig > -0.5 and ts <= 1/8:
            Fp = 1
        elif pressureinpsig > -0.5 and ts > 1/4:
            Fp = (((pressureinbarg+1)*Di)/(2*(850-0.6*(pressureinbarg+1)))+0.00315)/0.0063

        # Material Factor
        Fm = 1 # Carbon steel, ID no. 18 from Table A.3

        # Bare module factor for vertical process vessel
        B1 = 2.25
        B2 = 1.82

        # Bare module cost of reactor
        Cbm = cp0_2018*(B1+(B2*Fp*Fm))

        return Cbm

    def reactor_results(self):
        # Electricity cost for heating/cooling
        if self.beforeinlettemp < self.inlettemp:
            # Heating is required
            cost_of_heating = 0.10 * abs(self.duty) * 0.000277778  # cost of heating per hour
            Cbm = self.reactor_cost()
            objective = (cost_of_heating+Cbm)/self.MFproduction
        else:
            cost_of_cooling = 0.02 * abs(self.duty) * 0.000277778  # cost of cooling per hour
            Cbm = self.reactor_cost()
            objective = (cost_of_cooling + Cbm)/self.MFproduction
        return objective
















