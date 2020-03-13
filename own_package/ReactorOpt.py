from own_package.hysys.hysys_CSTR import Reactor
from own_package.hysys.hysys_link import init_hysys
from own_package.pso_ga import pso_ga
from own_package.others import create_excel_file, print_df_to_excel
import openpyxl
import pickle
import pandas as pd
import numpy as np


def optimize_reactor(storedata, sleep, pso_gen, ga, type):
    Hycase = init_hysys()
    if type == 'cstr':
        reactor = Reactor(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt', type=type)
        b_inlettemp = [60, 110]
        b_catalystweight = [0.0001, 0.05]
        # b_residencetime = [0.0015, 0.1723]
        b_residencetime = [0.0015, 2]
        b_reactorP = [2000, 4000]
        b_methanolCOratio = [1, 100]
        p_store = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP, b_methanolCOratio]
    elif type == 'pfr':
        reactor = Reactor(Hycase=Hycase, reactor_name='PFR-100', sprd_name='PFR_opt', type=type)
        b_inlettemp = [60, 110]
        b_catalystweight = [0.0001, 0.05]
        b_residencetime = [0.0015, 2]
        b_reactorP = [2000, 4000]
        b_methanolCOratio = [1, 100]
        p_store = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP, b_methanolCOratio]
    elif type == 'cstr2':
        reactor = Reactor(Hycase=Hycase, reactor_name='R-100-2', sprd_name='CSTR_opt-2', type=type)
        b_inlettemp = [60, 110]
        b_catalystweight = [0.0001, 0.05]
        # b_residencetime = [0.0015, 0.1723]
        b_residencetime = [0.0015, 2]
        b_reactorP = [2000, 4000]
        b_methanolCOratio = [1, 100]
        p_store = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP, b_methanolCOratio]
    elif type == 'isothermalcstr':
        reactor = Reactor(Hycase=Hycase, reactor_name='R-100-3', sprd_name='CSTR_opt-3', type=type)
        b_inlettemp = [60, 110]
        b_catalystweight = [0.0001, 0.05]
        # b_residencetime = [0.0015, 0.1723]
        b_residencetime = [0.0015, 2]
        b_reactorP = [2000, 4000]
        b_methanolCOratio = [1, 100]
        p_store = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP, b_methanolCOratio]

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
        nonlocal reactor
        inlettemp, catalystweight, residencetime, reactorP, methanolCOratio = individual
        reactor.solve_reactor(inlettemp=individual[0], catatlystweight=individual[1],
                              residencetime=individual[2], reactorP=individual[3], methanolCOratio=individual[4], sleep=sleep)
        return (reactor.reactor_results(storedata, type=type),)

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


def get_data_from_hysys(best, sleep, type):
    Hycase = init_hysys()
    if type == 'cstr':
        reactor = Reactor(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt', type=type)
        reactor.solve_reactor(inlettemp=best[0], catatlystweight=best[1], residencetime=best[2], reactorP=best[3],
                              methanolCOratio=best[4], sleep=sleep)
        reactor.reactor_results(storedata=True, type=type)
        read_col_data_store(name='cstr')
    elif type == 'pfr':
        reactor = Reactor(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt', type=type)
        reactor.solve_reactor(inlettemp=best[0], catatlystweight=best[1], residencetime=best[2], reactorP=best[3],
                              methanolCOratio=best[4], sleep=sleep)
        reactor.reactor_results(storedata=True, type=type)
        read_col_data_store(name='pfr')
    elif type == 'cstr2':
        reactor = Reactor(Hycase=Hycase, reactor_name='R-100-2', sprd_name='CSTR_opt-2', type=type)
        reactor.solve_reactor(inlettemp=best[0], catatlystweight=best[1], residencetime=best[2], reactorP=best[3],
                              methanolCOratio=best[4], sleep=sleep)
        reactor.reactor_results(storedata=True, type=type)
        read_col_data_store(name='cstr2')
    elif type == 'isothermalcstr':
        reactor = Reactor(Hycase=Hycase, reactor_name='R-100-3', sprd_name='CSTR_opt-3', type=type)
        reactor.solve_reactor(inlettemp=best[0], catatlystweight=best[1], residencetime=best[2], reactorP=best[3],
                              methanolCOratio=best[4], sleep=sleep)
        reactor.reactor_results(storedata=True, type=type)
        read_col_data_store(name='isothermalcstr')

def run_ReactorOpt(storedata, sleep, pso_gen, ga, type):
    if storedata:
        best = optimize_reactor(storedata=storedata, sleep=sleep, pso_gen=pso_gen, ga=ga, type=type)
        read_col_data_store(name=type)
    else:
        best = optimize_reactor(storedata=storedata, sleep=sleep, pso_gen=pso_gen, ga=ga, type=type)
        get_data_from_hysys(best=best, sleep=0.3, type=type)


def run_sensitivity_analysis(sleep):
    Hycase = init_hysys()
    reactor = Reactor(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt', type='cstr')
    b_inlettemp = 60
    b_catalystweight = 0.025
    b_residencetime = 0.707
    b_reactorP = 4000
    b_methanolCOratio = 70.74
    while b_inlettemp <= 110:
        DVvector = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP, b_methanolCOratio]
        reactor.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                              reactorP=DVvector[3], methanolCOratio=DVvector[4],
                              sleep=sleep)
        reactor.reactor_results(storedata=True, type='cstr')
        b_inlettemp += 1
    read_col_data_store(name='TempSensiAnalysis')
    reactor = Reactor(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt', type='cstr')
    b_inlettemp = 80
    b_catalystweight = 0.0001
    while b_catalystweight <= 0.05:
        DVvector = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP, b_methanolCOratio]
        reactor.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                              reactorP=DVvector[3], methanolCOratio=DVvector[4],
                              sleep=sleep)
        reactor.reactor_results(storedata=True, type='cstr')
        b_catalystweight += 0.001
    read_col_data_store(name='CatalystSensiAnalysis')
    reactor = Reactor(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt', type='cstr')
    b_catalystweight = 0.025
    b_residencetime = 0.050
    while b_residencetime <= 2:
        DVvector = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP, b_methanolCOratio]
        reactor.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                              reactorP=DVvector[3], methanolCOratio=DVvector[4],
                              sleep=sleep)
        reactor.reactor_results(storedata=True, type='cstr')
        b_residencetime += 0.01
    read_col_data_store(name='ResidenceTSensiAnalysis')
    reactor = Reactor(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt', type='cstr')
    b_residencetime = 0.707
    b_reactorP = 2000
    while b_reactorP <= 4000:
        DVvector = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP, b_methanolCOratio]
        reactor.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                              reactorP=DVvector[3], methanolCOratio=DVvector[4],
                              sleep=sleep)
        reactor.reactor_results(storedata=True, type='cstr')
        b_reactorP += 10
    read_col_data_store(name='ReactorPSensiAnalysis')
    reactor = Reactor(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt', type='cstr')
    b_reactorP = 4000
    b_methanolCOratio = 1
    while b_methanolCOratio <= 100:
        DVvector = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP, b_methanolCOratio]
        reactor.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                              reactorP=DVvector[3], methanolCOratio=DVvector[4],
                              sleep=sleep)
        reactor.reactor_results(storedata=True, type='cstr')
        b_methanolCOratio += 5
    read_col_data_store(name='MethanolCOratioSensiAnalysis')


def run_sensitivity_analysis_bestVector(sleep, best, type):
    Hycase = init_hysys()
    reactor = Reactor(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt', type='cstr')
    b_inlettemp = best[0]
    b_catalystweight = best[1]
    b_residencetime = best[2]
    b_reactorP = best[3]
    b_methanolCOratio = best[4]
    lowerbound = max(b_inlettemp - 10, 50)
    upperbound = min(b_inlettemp + 10, 110)
    for i in np.arange(lowerbound, upperbound, 1):
        DVvector = [i, b_catalystweight, b_residencetime, b_reactorP, b_methanolCOratio]
        reactor.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                              reactorP=DVvector[3], methanolCOratio=DVvector[4],
                              sleep=sleep)
        reactor.reactor_results(storedata=True, type=type)
    read_col_data_store(name='TempSensiAnalysisforBEST')
    reactor = Reactor(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt', type='cstr')
    lowerbound = max(b_catalystweight - 0.01, 0.0001)
    upperbound = min(b_catalystweight + 0.01, 0.05)
    for i in np.arange(lowerbound, upperbound, 0.0001):
        DVvector = [b_inlettemp, i, b_residencetime, b_reactorP, b_methanolCOratio]
        reactor.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                              reactorP=DVvector[3], methanolCOratio=DVvector[4],
                              sleep=sleep)
        reactor.reactor_results(storedata=True, type=type)
    read_col_data_store(name='CatalystSensiAnalysisforBEST')
    reactor = Reactor(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt', type='cstr')
    lowerbound = max(b_residencetime - 0.2, 0.05)
    upperbound = min(b_residencetime + 0.2, 2)
    for i in np.arange(lowerbound, upperbound, 0.01):
        DVvector = [b_inlettemp, b_catalystweight, i, b_reactorP, b_methanolCOratio]
        reactor.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                              reactorP=DVvector[3], methanolCOratio=DVvector[4],
                              sleep=sleep)
        reactor.reactor_results(storedata=True, type=type)
    read_col_data_store(name='ResidenceTSensiAnalysisforBEST')
    reactor = Reactor(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt', type='cstr')
    lowerbound = max(b_reactorP - 200, 2000)
    upperbound = min(b_reactorP + 200, 4000)
    for i in np.arange(lowerbound, upperbound, 10):
        DVvector = [b_inlettemp, b_catalystweight, b_residencetime, i, b_methanolCOratio]
        reactor.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                              reactorP=DVvector[3], methanolCOratio=DVvector[4],
                              sleep=sleep)
        reactor.reactor_results(storedata=True, type=type)
    read_col_data_store(name='ReactorPSensiAnalysisforBEST')
    reactor = Reactor(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt', type='cstr')
    lowerbound = max(b_methanolCOratio - 10, 1)
    upperbound = min(b_methanolCOratio + 10, 100)
    for i in np.arange(lowerbound, upperbound, 1):
        DVvector = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP, i]
        reactor.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                              reactorP=DVvector[3], methanolCOratio=DVvector[4],
                              sleep=sleep)
        reactor.reactor_results(storedata=True, type=type)
    read_col_data_store(name='MethanolCOratioSensiAnalysisforBEST')

run_ReactorOpt(storedata=True, sleep=0.5, pso_gen=50, ga=True, type='cstr')
# run_sensitivity_analysis(sleep=0.3)
#best = [89.83665875336027, 0.008469789091178227, 1.894972583420758, 4000, 12.956305834759327]
#get_data_from_hysys(best=best, sleep=0.3, type='pfr')
# run_sensitivity_analysis_bestVector(sleep=0.3, best=best, type='cstr')
