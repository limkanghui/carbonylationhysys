import numpy as np
import pandas as pd
import itertools, math
import time
import openpyxl

from others import create_excel_file, print_df_to_excel


class CSTR:
    def __init__(self, Hycase, column_name, sprd_name, max_iter=500):
        self.Hycase = Hycase
        self.Reactor = Hycase.Flowsheet.Operations.Item(column_name)

        # Access to Column  objects
        self.Specifications = self.ColumnFlowsheet.Specifications  # RR/BR/....
        self.Operations = self.ColumnFlowsheet.Operations  # Main TS/Reboiler/Condenser

        # Access to optimization parameters
        self.ColumnFlowsheet.MaximumIterations = max_iter  # if too large, will take very long for optimization
        # Active Specs. Use object.GoalValue = XXX to modify spec value.
        self.active_specs = [self.ColumnFlowsheet.ActiveSpecifications.Item(idx) for (idx, _) in enumerate(self.ColumnFlowsheet.ActiveSpecifications.Names)]
        # Other decision variables for the column
        # self.dv_number_of_trays = self.Main_Tower.NumberOfTrays
        # self.Main_TS.SpecifyFeedLocation(self.FeedMainTS, max(round(specs[-1] * specs[-2]), 1))
        self.sprd_pressure = Hycase.Flowsheet.Operations.Item(sprd_name)
        self.max_trays = max_trays
        # Setting up Material Stream objects
        self.feed = [self.ColumnFlowsheet.FeedStreams.Item(x+1) for x in range(number_of_feed)]

        # Used to store all results evaulated from .solve_column to pickle save at the end of an optimization run
        self.data_store = []

        if number_of_draws == 3:
            self.partial_condenser = True
            self.vap = self.ColumnFlowsheet.MaterialStreams.Item(4)
            self.distillate = self.ColumnFlowsheet.MaterialStreams.Item(5)
            self.btm = self.ColumnFlowsheet.MaterialStreams.Item(6)
        else:
            self.partial_condenser = False
            self.vap = None
            self.distillate = self.ColumnFlowsheet.MaterialStreams.Item(4)
            self.btm = self.ColumnFlowsheet.MaterialStreams.Item(5)

        # Setting up energy streams
        self.cond = self.ColumnFlowsheet.EnergyStreams.Item(0)
        self.rebo = self.ColumnFlowsheet.EnergyStreams.Item(1)


    def solve_column(self, active_spec_goals, del_p, rebo_p, number_of_trays, feed_frac):
        '''
        Input specs -> Set Hysys column according to specs -> Run column -> Report converge or not and results
        :param specs: [active spec 1, active spec 2, active spec 3 (if have)]
        :param feed_frac: Int -> 1 feed frac only, List -> 2 or more feed stages. 0.1 ==> 10% from the top
        :return:
        '''
        self.ColumnFlowsheet.Reset()
        for active_spec, active_spec_goal in zip(self.active_specs, active_spec_goals):
            active_spec.GoalValue = active_spec_goal

        self.Main_Tower.NumberOfTrays = number_of_trays
        try:
            for idx, single_feed_frac in enumerate(feed_frac):
                self.Main_Tower.SpecifyFeedLocation(self.feed[idx], max(round(single_feed_frac * number_of_trays), 1))
        except TypeError:
            # If feed_frac is scalar ==> one feed only. But self.feed will still be [XXX] list with single feed obj.
            self.Main_Tower.SpecifyFeedLocation(self.feed[0], max(round(feed_frac * number_of_trays), 1))

        for i in range(2,self.max_trays+2+1):
            # Clear all the old pressure values. B1 no need to clear since condenser is always at stage 0.
            self.sprd_pressure.Cell('B{}'.format(i)).Erase()

        self.sprd_pressure.Cell('B1').CellValue = rebo_p - del_p * number_of_trays
        rebo_stage = 1 + number_of_trays + 1
        self.sprd_pressure.Cell('B{}'.format(rebo_stage)).CellValue = rebo_p
        self.ColumnFlowsheet.Run()
        self.status = self.ColumnFlowsheet.CfsConverged

    def store_to_data_store(self, nt, feedstage):
        '''

        :param nt: MUST BE INT
        :param feedstage: MUST BE LIST
        :return:
        '''
        # Specification
        spec_names = self.ColumnFlowsheet.Specifications.Names
        spec_values = [self.ColumnFlowsheet.Specifications(x) for x in range(1, len(spec_names+1))]
        active_spec_names = self.ColumnFlowsheet.ActiveSpecifications.Names

        # Column parameters
        cond_P = self.Main_Tower.PressureValue[0]
        rebo_P = self.Main_Tower.PressureValue[-1]

        # Distillate

        return [[spec_names + ['Active Spec {}'.format(x) for x in range(1, len(active_spec_names)+1)] +
                 ['NT'] + ['FS {}'.format(x) for x in range(1, len(feedstage)+1)] + ['Cond P', 'Rebo_P']],
                [spec_values + active_spec_names + [nt] + feedstage + [cond_P, rebo_P]]]


    def column_results(self):
        if self.status:
            top_vap_mol_flow = self.ColumnFlowsheet.NetMolarVapourFlowsValue[1] * 3600  # kmol/h, Hysys is in secs

            # Operating Cost
            cond_duty = self.cond.HeatFlowValue * 3600  # change to per hours
            rebo_duty = self.rebo.HeatFlowValue * 3600  # change to per hour

            tac = (cond_duty+rebo_duty)/1e3 + top_vap_mol_flow

        else:
            tac = 1e20

        return tac

    def get_capital_cost(self, pres, leng, dim):
        '''Assume head thickness same as shell'''
        '''p in kpa, l in m & dim in m & W in kg & C in $/kg'''

        di = dim * 39.3701
        dift = dim * 3.28084
        l = leng * 39.3701
        p = (pres - 101.3) * 0.145038
        if p < 5:
            dp = 10
        else:
            dp = math.exp(0.60608 + 0.91615 * (math.log(p)) + 0.0015655 * (math.log(p)) ** 2)
        S = 13750
        E = 0.85
        C = 100
        tp = (dp * di) / (2 * S * E - 1.2 * dp)
        if tp > 1.25:
            E = 1
            tp = (dp * di) / (2 * S * E - 1.2 * dp)
        if dift <= 4 and tp < 1 / 4:
            tp = 1 / 4
        if 6 >= dift >= 4 and tp < 5 / 16:
            tp = 5 / 16
        if 6 <= dift <= 8 and tp < 3 / 8:
            tp = 3 / 8
        if 8 <= dift <= 10 and tp < 7 / 16:
            tp = 7 / 16
        if 10 <= dift <= 12 and tp < 1 / 2:
            tp = 1 / 2
        if dift > 12:
            return 'error1'
        ts = 1.0
        tv = 3 / 16
        while ts != tv:
            ts = tv
            do = di + 2 * ts
            tw = (0.22 * (do + 18) * l ** 2) / (S * do ** 2)
            tv = ((tp * 2 + tw) / 2) + 1 / 8
            if 3 / 16 <= tv <= 8 / 16:
                for tvv in np.arange(3 / 16, 1 / 2, 1 / 16):
                    if tv <= tvv:
                        tv = tvv
                        break
            if 8 / 16 < tv <= 2.0:
                for tvv in np.arange(5 / 8, 2.0, 1 / 8):
                    if tv <= tvv:
                        tv = tvv
                        break
            if 2.0 < tv <= 3.0:
                for tvv in np.arange(2.25, 3.0, 1 / 4):
                    if tv <= tvv:
                        tv = tvv
                        break
            if tv > 3.0:
                return 'error2'
                break
        W = math.pi * (di + ts) * (l + 0.8 * di) * ts * 490 * 0.453592
        return W*C


    def feaible_converge_column(self, specs_name, specs_bounds, trials, trays_bounds, trays_trials, cuhu_name, write_dir):
        specs_linspace = [np.linspace(start=bound[0], stop=bound[1], num=trial) for (bound, trial) in
                          zip(specs_bounds, trials)]
        tray_spec = np.unique(np.rint(np.linspace(start=trays_bounds[0], stop=trays_bounds[1], num=trays_trials[0])))
        feed_tray = np.linspace(0,1, num=trays_trials[1])
        specs_linspace.extend([tray_spec, feed_tray])
        specs_linspace = [x.tolist() for x in specs_linspace]
        specs_combi = itertools.product(*specs_linspace)

        feasible_combi = []
        feasible_score = []
        for idx, specs in enumerate(specs_combi):
            self.ColumnFlowsheet.Reset()
            self.Main_TS.NumberOfTrays = specs[-2]
            self.Main_TS.SpecifyFeedLocation(self.FeedMainTS, max(round(specs[-1]*specs[-2]),1))
            for (spec_name, spec) in zip(specs_name, specs[:-2]):
                self.Specifications.Item(spec_name).GoalValue = spec

            self.ColumnFlowsheet.Run()
            #time.sleep(2)
            status = self.ColumnFlowsheet.CfsConverged
            print('Trial {}. Converged = {}'.format(idx+1, status))
            if status == True:
                feasible_combi.append(specs)
                cu = self.Hycase.Flowsheet.EnergyStreams.Item(cuhu_name[0]).HeatFlow.GetValue('kJ/h')
                hu = self.Hycase.Flowsheet.EnergyStreams.Item(cuhu_name[0]).HeatFlow.GetValue('kJ/h')
                distillate_mdot = self.distillate.MassFlowValue * 3600
                btm_mdot = self.btm.MassFlowValue * 3600
                feasible_score.append([cu, hu, distillate_mdot, btm_mdot])

        excel_file = create_excel_file('{}/feasible_combi'.format(write_dir))
        data = [x + y for (x, y) in zip(feasible_combi, feasible_score)]
        df = pd.DataFrame(data=data, columns=specs_name + ['NT', 'Feed Stage', 'CU', 'HU', 'Distillate mdot', 'Btm mdot'])
        wb = openpyxl.load_workbook(excel_file)
        ws = wb[wb.sheetnames[-1]]
        print_df_to_excel(df=df, ws=ws)
        wb.save(excel_file)
