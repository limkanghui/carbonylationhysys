import numpy as np
import pandas as pd
import itertools, math
import time, pickle
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

        #Other variables
        self.E101duty = self.spreadsheetdata.Cell('B9').CellValue *3600
        self.beforeinlettemp = self.spreadsheetdata.Cell('B10').CellValue
        self.reactorsize = self.Reactor.VolumeValue
        self.reactortemp = self.spreadsheetdata.Cell('B12').CellValue
        self.E100duty = self.spreadsheetdata.Cell('D9').CellValue *3600
        self.E102duty = self.spreadsheetdata.Cell('D10').CellValue *3600
        self.E104duty = self.spreadsheetdata.Cell('D11').CellValue *3600
        self.E106duty = self.spreadsheetdata.Cell('D12').CellValue *3600
        self.E111duty = self.spreadsheetdata.Cell('D13').CellValue *3600
        self.P8duty = self.spreadsheetdata.Cell('B13').CellValue *3600
        self.P106duty = self.spreadsheetdata.Cell('B14').CellValue *3600
        self.C101duty = self.spreadsheetdata.Cell('D14').CellValue *3600
        self.C103duty = self.spreadsheetdata.Cell('D15').CellValue *3600
        self.C104duty = self.spreadsheetdata.Cell('D16').CellValue *3600
        self.C100duty = self.spreadsheetdata.Cell('B15').CellValue *3600
        self.C102duty = self.spreadsheetdata.Cell('B16').CellValue *3600
        self.beforeinlet8_1_temp = self.spreadsheetdata.Cell('B19').CellValue
        self.catalystmassflow = self.spreadsheetdata.Cell('D19').CellValue *3600
        self.comassflow = self.spreadsheetdata.Cell('D17').CellValue * 3600
        self.MFin1 = self.spreadsheetdata.Cell('B17').CellValue *3600
        self.MFin2 = self.spreadsheetdata.Cell('B18').CellValue * 3600

        #Objective
        self.conversion = self.spreadsheetdata.Cell('D5').CellValue
        self.MFproduction = self.spreadsheetdata.Cell('D6').CellValue *3600

        # Used to store all results evaulated from .solve_column to pickle save at the end of an optimization run
        self.data_store = []
        self.data_store_columns = ['inlettemp', 'catalystweight', 'residencetime', 'reactorP', 'reactorsize', 'reactortemp', 'conversion', 'MFproduction','cost of heating','cost of cooling','cost of comp and pump','FCI','COMd','objective']


    def solve_reactor(self, inlettemp, catatlystweight, residencetime, reactorP, sleep):

        self.spreadsheetdata.Cell('B2').CellValue = inlettemp
        self.spreadsheetdata.Cell('B3').CellValue = catatlystweight
        self.spreadsheetdata.Cell('B4').CellValue = residencetime
        self.spreadsheetdata.Cell('B5').CellValue = reactorP

        self.inlettemp = inlettemp
        self.catalystweight = catatlystweight
        self.residencetime = residencetime
        self.reactorP = reactorP

        time.sleep(sleep)

        # Constraints

        # Other variables
        self.E101duty = self.spreadsheetdata.Cell('B9').CellValue * 3600
        self.beforeinlettemp = self.spreadsheetdata.Cell('B10').CellValue
        self.reactorsize = self.Reactor.VolumeValue
        self.reactortemp = self.spreadsheetdata.Cell('B12').CellValue
        self.E100duty = self.spreadsheetdata.Cell('D9').CellValue * 3600
        self.E102duty = self.spreadsheetdata.Cell('D10').CellValue * 3600
        self.E104duty = self.spreadsheetdata.Cell('D11').CellValue * 3600
        self.E106duty = self.spreadsheetdata.Cell('D12').CellValue * 3600
        self.E111duty = self.spreadsheetdata.Cell('D13').CellValue * 3600
        self.P8duty = self.spreadsheetdata.Cell('B13').CellValue * 3600
        self.P106duty = self.spreadsheetdata.Cell('B14').CellValue * 3600
        self.C101duty = self.spreadsheetdata.Cell('D14').CellValue * 3600
        self.C103duty = self.spreadsheetdata.Cell('D15').CellValue * 3600
        self.C104duty = self.spreadsheetdata.Cell('D16').CellValue * 3600
        self.C100duty = self.spreadsheetdata.Cell('B15').CellValue * 3600
        self.C102duty = self.spreadsheetdata.Cell('B16').CellValue * 3600
        self.beforeinlet8_1_temp = self.spreadsheetdata.Cell('B19').CellValue
        self.catalystmassflow = self.spreadsheetdata.Cell('D19').CellValue * 3600
        self.comassflow = self.spreadsheetdata.Cell('D17').CellValue *3600
        self.MFin1 = self.spreadsheetdata.Cell('B17').CellValue * 3600
        self.MFin2 = self.spreadsheetdata.Cell('B18').CellValue * 3600

        # Objective
        self.conversion = self.spreadsheetdata.Cell('D5').CellValue
        self.MFproduction = self.spreadsheetdata.Cell('D6').CellValue * 3600

        self.store_to_data_store()

    def reactor_design(self,):
        # CSTR modelled as a pressure vessel
        # Reactor design based on Towler's Book

        operatingtemp = self.reactortemp
        operatingP = self.reactorP

        # Design Pressure in psig
        pressureinpsig = operatingP*0.145038-14.7
        if pressureinpsig >= 0 and pressureinpsig < 5:
            designP = 10
        elif pressureinpsig >= 5 and pressureinpsig < 10:
            designP = 15
        elif pressureinpsig >= 10 and pressureinpsig <= 1000:
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
            shell_thickness = max(1/4, shell_thickness)
        elif Di*3.28084 <= 6:
            shell_thickness = max(5/16, shell_thickness)
        elif Di*3.28084 <= 8:
            shell_thickness = max(3/8, shell_thickness)
        elif Di*3.28084 <= 10:
            shell_thickness = max(7/16,shell_thickness)
        elif Di*3.28084 <= 12:
            shell_thickness = max(1/2, shell_thickness)

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

        weight = math.pi*(Di*39.3701+ts)*(3*Di*39.3701+0.8*Di)*ts*0.284*0.453592 # weight in kg

        return ts, weight, Di

    def reactor_cost(self):
        # Using Reactor-mixer values from Appendix A of turton's tb
        # Volume between 0.04 and 60 m3

        k1 = 4.7116
        k2 = 0.4479
        k3 = 0.0004

        S_init = self.reactorsize

        if S_init <= 60:
            counter = 1
        else:
            counter = 2

        S = S_init
        while S >= 60:
            S = S_init/(counter)
            counter+=1

        cp0_2001 = 10**(k1+k2*math.log10(S)+k3*(math.log10(S))**2)*counter
        cp0_2018 = cp0_2001*603.1/397

        # Pressure factor for process vessels
        operatingP = self.reactorP
        pressureinpsig = operatingP * 0.145038 - 14.7
        pressureinbarg = pressureinpsig * 0.0689476
        ts, weight, Di = self.reactor_design()

        if pressureinbarg <= -0.5:
            Fp = 1.25
        elif pressureinpsig > -0.5 and ts <= 1/4:
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

    def reactor_results(self, storedata):
        # Electricity cost for heating/cooling
        if self.beforeinlettemp < self.inlettemp and self.beforeinlet8_1_temp < self.inlettemp:
            # Heating is required
            cost_of_heating = 0.10 * abs(self.E101duty+self.E102duty) * 0.000277778  # cost of heating per hour
            # Combined cooling costs
            cooling_duties = self.E100duty+self.E104duty+self.E106duty+self.E111duty
            cost_of_cooling = 0.02 * cooling_duties * 0.000277778
            # Combined Compressor and Pump Electricity Costs
            compressor_duties = self.C100duty+self.C101duty+self.C102duty+self.C103duty+self.C104duty
            pump_duties = self.P8duty+self.P106duty
            cost_of_comp_and_pump_duties = 0.2 * (compressor_duties+pump_duties) * 0.000277778
            # Cost of Manufacture w/o depreciation: COMd = 0.18 FCI + 2.73C_OL + 1.23(C_RM + C_WT + C_UT)
            # Ignore waste treatment cost C_WT
            # FCI
            Cbm = self.reactor_cost()
            FCI = Cbm
            # Cost of utilities per annual (8000 hours a year)
            C_UT = (cost_of_heating+cost_of_comp_and_pump_duties+cost_of_cooling)*8000
            # Cost of raw materials, consider CO feed and catalyst top up
            # Cost of catalyst $226/kg (sigma aldrich), 8000 hours a year
            cost_of_catalyst = self.catalystmassflow*226*8000
            # Cost of CO $10/m3 (alibaba), 8000 hours a year
            cost_of_CO = self.comassflow/10*8000
            C_RM = cost_of_catalyst+cost_of_CO
            # Cost of labour, C_OL
            # C_OL = wage * 4.5(6.29 + 31.7P^2 + 0.23 N_np)^0.5
            # P = 0, N_np = 1 reactor, 5 compressors, 6 heat exchangers
            # Wage of chemical plant operator = $62170/pear
            C_OL = 62170 * 4.5*(6.29 + 31.7*0**2 + 0.23*12)

            COMd = 0.18*FCI+2.73*C_OL+1.23*(C_RM+C_UT)

            # Yield of MF wrt CO, (MF_out - MF_in)/CO
            yield_of_MF = abs(self.MFproduction-(self.MFin1+self.MFin2))/self.comassflow

            objective = COMd/yield_of_MF

            if storedata == True:
                data = self.store_to_data_store()
                data.extend([cost_of_heating])
                data.extend([cost_of_cooling])w
                data.extend([cost_of_comp_and_pump_duties])
                data.extend(FCI)
                data.extend(COMd)
                data.extend(objective)
                self.data_store.append(data)
                self.save_data_store_pkl(self.data_store)

        elif self.beforeinlettemp < self.inlettemp and self.beforeinlet8_1_temp > self.inlettemp:
            # Heating is required
            cost_of_heating = 0.10 * abs(self.E101duty) * 0.000277778  # cost of heating per hour
            # Combined cooling costs
            cooling_duties = self.E100duty + self.E102duty + self.E104duty + self.E106duty + self.E111duty
            cost_of_cooling = 0.02 * cooling_duties * 0.000277778
            # Combined Compressor and Pump Electricity Costs
            compressor_duties = self.C100duty + self.C101duty + self.C102duty + self.C103duty + self.C104duty
            pump_duties = self.P8duty + self.P106duty
            cost_of_comp_and_pump_duties = 0.2 * (compressor_duties + pump_duties) * 0.000277778
            # Cost of Manufacture w/o depreciation: COMd = 0.18 FCI + 2.73C_OL + 1.23(C_RM + C_WT + C_UT)
            # Ignore waste treatment cost C_WT
            # FCI
            Cbm = self.reactor_cost()
            FCI = Cbm
            # Cost of utilities per annual (8000 hours a year)
            C_UT = (cost_of_heating + cost_of_comp_and_pump_duties + cost_of_cooling) * 8000
            # Cost of raw materials, consider CO feed and catalyst top up
            # Cost of catalyst $226/kg (sigma aldrich), 8000 hours a year
            cost_of_catalyst = self.catalystmassflow * 226 * 8000
            # Cost of CO $10/m3 (alibaba), 8000 hours a year
            cost_of_CO = self.comassflow / 10 * 8000
            C_RM = cost_of_catalyst + cost_of_CO
            # Cost of labour, C_OL
            # C_OL = wage * 4.5(6.29 + 31.7P^2 + 0.23 N_np)^0.5
            # P = 0, N_np = 1 reactor, 5 compressors, 6 heat exchangers
            # Wage of chemical plant operator = $62170/pear
            C_OL = 62170 * 4.5*(6.29 + 31.7 * 0 ** 2 + 0.23 * 12)

            COMd = 0.18 * FCI + 2.73 * C_OL + 1.23 * (C_RM + C_UT)

            # Yield of MF wrt CO, (MF_out - MF_in)/CO
            yield_of_MF = abs(self.MFproduction - (self.MFin1 + self.MFin2)) / self.comassflow

            objective = COMd / yield_of_MF

            if storedata == True:
                data = self.store_to_data_store()
                data.extend([cost_of_heating])
                data.extend([cost_of_cooling])
                data.extend([cost_of_comp_and_pump_duties])
                data.extend(FCI)
                data.extend(COMd)
                data.extend(objective)
                self.data_store.append(data)
                self.save_data_store_pkl(self.data_store)

        elif self.beforeinlettemp > self.inlettemp and self.beforeinlet8_1_temp < self.inlettemp:
            # Heating is required
            cost_of_heating = 0.10 * abs(self.E102duty) * 0.000277778  # cost of heating per hour
            # Combined cooling costs
            cooling_duties = self.E100duty + self.E101duty + self.E104duty + self.E106duty + self.E111duty
            cost_of_cooling = 0.02 * cooling_duties * 0.000277778
            # Combined Compressor and Pump Electricity Costs
            compressor_duties = self.C100duty + self.C101duty + self.C102duty + self.C103duty + self.C104duty
            pump_duties = self.P8duty + self.P106duty
            cost_of_comp_and_pump_duties = 0.2 * (compressor_duties + pump_duties) * 0.000277778
            # Cost of Manufacture w/o depreciation: COMd = 0.18 FCI + 2.73C_OL + 1.23(C_RM + C_WT + C_UT)
            # Ignore waste treatment cost C_WT
            # FCI
            Cbm = self.reactor_cost()
            FCI = Cbm
            # Cost of utilities per annual (8000 hours a year)
            C_UT = (cost_of_heating + cost_of_comp_and_pump_duties + cost_of_cooling) * 8000
            # Cost of raw materials, consider CO feed and catalyst top up
            # Cost of catalyst $226/kg (sigma aldrich), 8000 hours a year
            cost_of_catalyst = self.catalystmassflow * 226 * 8000
            # Cost of CO $10/m3 (alibaba), 8000 hours a year
            cost_of_CO = self.comassflow / 10 * 8000
            C_RM = cost_of_catalyst + cost_of_CO
            # Cost of labour, C_OL
            # C_OL = wage * 4.5(6.29 + 31.7P^2 + 0.23 N_np)^0.5
            # P = 0, N_np = 1 reactor, 5 compressors, 6 heat exchangers
            # Wage of chemical plant operator = $62170/pear
            C_OL = 62170 * 4.5*(6.29 + 31.7 * 0 ** 2 + 0.23 * 12)

            COMd = 0.18 * FCI + 2.73 * C_OL + 1.23 * (C_RM + C_UT)

            # Yield of MF wrt CO, (MF_out - MF_in)/CO
            yield_of_MF = abs(self.MFproduction - (self.MFin1 + self.MFin2)) / self.comassflow

            objective = COMd / yield_of_MF

            if storedata == True:
                data = self.store_to_data_store()
                data.extend([cost_of_heating])
                data.extend([cost_of_cooling])
                data.extend([cost_of_comp_and_pump_duties])
                data.extend(FCI)
                data.extend(COMd)
                data.extend(objective)
                self.data_store.append(data)
                self.save_data_store_pkl(self.data_store)

        else:
            # No heating is required at all
            cost_of_heating = 0
            # Combine all cooling costs
            cooling_duties = self.E100duty + self.E101duty + self.E102duty + self.E104duty + self.E106duty + self.E111duty
            cost_of_cooling = 0.02 * cooling_duties * 0.000277778
            # Combined Compressor and Pump Electricity Costs
            compressor_duties = self.C100duty + self.C101duty + self.C102duty + self.C103duty + self.C104duty
            pump_duties = self.P8duty + self.P106duty
            cost_of_comp_and_pump_duties = 0.2 * (compressor_duties + pump_duties) * 0.000277778
            # Cost of Manufacture w/o depreciation: COMd = 0.18 FCI + 2.73C_OL + 1.23(C_RM + C_WT + C_UT)
            # Ignore waste treatment cost C_WT
            # FCI
            Cbm = self.reactor_cost()
            FCI = Cbm
            # Cost of utilities per annual (8000 hours a year)
            C_UT = (cost_of_heating + cost_of_comp_and_pump_duties + cost_of_cooling) * 8000
            # Cost of raw materials, consider CO feed and catalyst top up
            # Cost of catalyst $226/kg (sigma aldrich), 8000 hours a year
            cost_of_catalyst = self.catalystmassflow * 226 * 8000
            # Cost of CO $10/m3 (alibaba), 8000 hours a year
            cost_of_CO = self.comassflow / 10 * 8000
            C_RM = cost_of_catalyst + cost_of_CO
            # Cost of labour, C_OL
            # C_OL = wage * 4.5(6.29 + 31.7P^2 + 0.23 N_np)^0.5
            # P = 0, N_np = 1 reactor, 5 compressors, 6 heat exchangers
            # Wage of chemical plant operator = $62170/pear
            C_OL = 62170 * 4.5*(6.29 + 31.7 * 0 ** 2 + 0.23 * 12)

            COMd = 0.18 * FCI + 2.73 * C_OL + 1.23 * (C_RM + C_UT)

            # Yield of MF wrt CO, (MF_out - MF_in)/CO
            yield_of_MF = abs(self.MFproduction - (self.MFin1 + self.MFin2)) / self.comassflow

            objective = COMd / yield_of_MF

            if storedata == True:
                data = self.store_to_data_store()
                data.extend([cost_of_heating])
                data.extend([cost_of_cooling])
                data.extend([cost_of_comp_and_pump_duties])
                data.extend(FCI)
                data.extend(COMd)
                data.extend(objective)
                self.data_store.append(data)
                self.save_data_store_pkl(self.data_store)

        return objective

    def store_to_data_store(self):
        # Decision Variables
        inlettemp = self.inlettemp
        catalystweight = self.catalystweight
        residencetime = self.residencetime
        reactorP = self.reactorP

        # Constraints

        # Other variables
        reactorsize = self.reactorsize
        reactortemp = self.reactortemp

        # Objective
        conversion = self.conversion
        MFproduction = self.MFproduction

        return [inlettemp, catalystweight, residencetime, reactorP, reactorsize, reactortemp, conversion, MFproduction]

    def save_data_store_pkl(self, data):
        with open('data_store.pkl', 'wb') as handle:
            pickle.dump([self.data_store_columns, data], handle, protocol=pickle.HIGHEST_PROTOCOL)









