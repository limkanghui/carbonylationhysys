from own_package.hysys.hysys_CSTR import CSTR
from own_package.hysys.hysys_link import init_hysys
from own_package.pso_ga import pso_ga
from own_package.others import create_excel_file, print_df_to_excel
import openpyxl
import pickle
import pandas as pd
import numpy as np

def optimize_CSTR(storedata,sleep,pso_gen,ga):
    Hycase = init_hysys()
    cstr = CSTR(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt')
    b_inlettemp = [60, 110]
    b_catalystweight = [0.0001, 0.05]
    b_residencetime = [0.05, 2]
    b_reactorP = [2000, 4000]
    p_store = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP]

    params = {'c1': 1.5, 'c2': 1.5, 'wmin': 0.4, 'wmax': 0.9,
              'ga_iter_min': 5, 'ga_iter_max': 20, 'iter_gamma': 10,
              'ga_num_min': 10, 'ga_num_max': 20, 'num_beta': 15,
              'tourn_size': 3, 'cxpd': 0.5, 'mutpd': 0.05, 'indpd': 0.5, 'eta': 0.5,
              'pso_iter': pso_gen, 'swarm_size': 50}

    pmin = [x[0] for x in p_store]
    pmax = [x[1] for x in p_store]

    smin = [abs(x - y) * 0.01 for x, y in zip(pmin, pmax)]
    smax = [abs(x - y) * 0.5 for x, y in zip(pmin, pmax)]

    def func(individual):
        nonlocal cstr
        inlettemp, catalystweight, residencetime, reactorP = individual
        cstr.solve_reactor(inlettemp=individual[0], catatlystweight=individual[1],
                           residencetime=individual[2], reactorP=individual[3], sleep=sleep)
        return (cstr.reactor_results(storedata),)

    pop, logbook, best = pso_ga(func=func, pmin=pmin, pmax=pmax,
                                smin=smin, smax=smax,
                                int_idx=None, params=params, ga=ga)
    return best

def read_col_data_store(name):
    with open('./data_store.pkl', 'rb') as handle:
        data_store = pickle.load(handle)

    write_excel = create_excel_file('./results/{}_results.xlsx'.format(name))
    wb = openpyxl.load_workbook(write_excel)
    ws = wb[wb.sheetnames[-1]]
    print_df_to_excel(df=pd.DataFrame(data=data_store[1], columns=data_store[0]), ws=ws)
    wb.save(write_excel)

def get_data_from_hysys(best,sleep):
    Hycase = init_hysys()
    cstr = CSTR(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt')
    cstr.solve_reactor(inlettemp=best[0], catatlystweight=best[1], residencetime=best[2], reactorP=best[3], sleep=sleep)
    cstr.reactor_results(storedata=True)
    read_col_data_store()

def run_CSTROpt(storedata, sleep, pso_gen, ga):
    if storedata == True:
        best = optimize_CSTR(storedata=storedata, sleep=sleep, pso_gen=pso_gen, ga=ga)
        read_col_data_store(name='cstr')
    else:
        best = optimize_CSTR(storedata=storedata, sleep=sleep, pso_gen=pso_gen, ga=ga)
        get_data_from_hysys(best=best, sleep=0.3)

def run_sensitivity_analysis(sleep):
    Hycase = init_hysys()
    cstr = CSTR(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt')
    b_inlettemp = 60
    b_catalystweight = 0.025
    b_residencetime = 0.707
    b_reactorP = 4000
    while b_inlettemp <= 110:
        DVvector = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP]
        cstr.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2], reactorP=DVvector[3],
                           sleep=sleep)
        cstr.reactor_results(storedata=True)
        b_inlettemp += 1
    read_col_data_store(name='TempSensiAnalysis')
    cstr = CSTR(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt')
    b_inlettemp = 80
    b_catalystweight = 0.0001
    while b_catalystweight <= 0.05:
        DVvector = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP]
        cstr.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                           reactorP=DVvector[3],
                           sleep=sleep)
        cstr.reactor_results(storedata=True)
        b_catalystweight += 0.001
    read_col_data_store(name='CatalystSensiAnalysis')
    cstr = CSTR(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt')
    b_catalystweight = 0.025
    b_residencetime = 0.050
    while b_residencetime <= 2:
        DVvector = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP]
        cstr.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                           reactorP=DVvector[3],
                           sleep=sleep)
        cstr.reactor_results(storedata=True)
        b_residencetime += 0.01
    read_col_data_store(name='ResidenceTSensiAnalysis')
    cstr = CSTR(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt')
    b_residencetime = 0.707
    b_reactorP = 2000
    while b_reactorP <= 4000:
        DVvector = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP]
        cstr.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                           reactorP=DVvector[3],
                           sleep=sleep)
        cstr.reactor_results(storedata=True)
        b_reactorP += 10
    read_col_data_store(name='ReactorPSensiAnalysis')

run_CSTROpt(storedata=False, sleep=0.3, pso_gen=200, ga=False)
#run_sensitivity_analysis(sleep=0.3)