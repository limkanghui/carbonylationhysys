from own_package.hysys.hysys_CSTR import Reactor
from own_package.hysys.hysys_link import init_hysys
from own_package.pso_ga import pso_ga
from own_package.others import create_excel_file, print_df_to_excel
from timeit import default_timer as timer
import openpyxl
import pickle
import pandas as pd
import numpy as np


def optimize_reactor(storedata, sleep, pso_gen, ga, type):
    global p_store
    Hycase = init_hysys()
    if type == 'cstr':
        reactor = Reactor(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt', type=type)
        b_inlettemp = [60, 110]
        b_catalystweight = [0.0001, 0.05]
        # b_residencetime = [0.0015, 0.1723]
        b_residencetime = [0.0015, 2]
        b_reactorP = [2000, 4000]
        b_methanolCOratio = [2, 100]
        p_store = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP, b_methanolCOratio]
    elif type == 'pfr':
        reactor = Reactor(Hycase=Hycase, reactor_name='PFR-100', sprd_name='PFR_opt', type=type)
        b_inlettemp = [60, 110]
        b_catalystweight = [0.0001, 0.05]
        b_residencetime = [0.0015, 2]
        b_reactorP = [2000, 4000]
        b_methanolCOratio = [2, 100]
        p_store = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP, b_methanolCOratio]
    elif type == 'cstr2':
        reactor = Reactor(Hycase=Hycase, reactor_name='R-100-2', sprd_name='CSTR_opt-2', type=type)
        b_inlettemp = [60, 110]
        b_catalystweight = [0.0001, 0.05]
        # b_residencetime = [0.0015, 0.1723]
        b_residencetime = [0.0015, 2]
        b_reactorP = [2000, 4000]
        b_methanolCOratio = [2, 100]
        p_store = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP, b_methanolCOratio]
    elif type == 'isothermalcstr':
        reactor = Reactor(Hycase=Hycase, reactor_name='R-100-3', sprd_name='CSTR_opt-3', type=type)
        b_inlettemp = [60, 110]
        b_catalystweight = [0.0001, 0.05]
        # b_residencetime = [0.0015, 0.1723]
        b_residencetime = [0.0015, 2]
        b_reactorP = [2000, 4000]
        b_methanolCOratio = [2, 100]
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
                              residencetime=individual[2], reactorP=individual[3], methanolCOratio=individual[4], sleep=sleep, type=type)
        return (reactor.reactor_results(storedata, type=type),)

    start = timer()
    pop, logbook, best = pso_ga(func=func, pmin=pmin, pmax=pmax,
                                smin=smin, smax=smax,
                                int_idx=None, params=params, ga=ga, type=type)
    end = timer()
    timetaken = (end - start)/3600
    print('time taken: {}h'.format(timetaken))
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
                              methanolCOratio=best[4], sleep=sleep, type=type)
        reactor.reactor_results(storedata=True, type=type)
        read_col_data_store(name='cstr')
    elif type == 'pfr':
        reactor = Reactor(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt', type=type)
        reactor.solve_reactor(inlettemp=best[0], catatlystweight=best[1], residencetime=best[2], reactorP=best[3],
                              methanolCOratio=best[4], sleep=sleep, type=type)
        reactor.reactor_results(storedata=True, type=type)
        read_col_data_store(name='pfr')
    elif type == 'cstr2':
        reactor = Reactor(Hycase=Hycase, reactor_name='R-100-2', sprd_name='CSTR_opt-2', type=type)
        reactor.solve_reactor(inlettemp=best[0], catatlystweight=best[1], residencetime=best[2], reactorP=best[3],
                              methanolCOratio=best[4], sleep=sleep, type=type)
        reactor.reactor_results(storedata=True, type=type)
        read_col_data_store(name='cstr2')
    elif type == 'isothermalcstr':
        reactor = Reactor(Hycase=Hycase, reactor_name='R-100-3', sprd_name='CSTR_opt-3', type=type)
        reactor.solve_reactor(inlettemp=best[0], catatlystweight=best[1], residencetime=best[2], reactorP=best[3],
                              methanolCOratio=best[4], sleep=sleep, type=type)
        reactor.reactor_results(storedata=True, type=type)
        read_col_data_store(name='isothermalcstr')

def run_ReactorOpt(storedata, sleep, pso_gen, ga, type, sensitivityanalysis):
    if storedata:
        best = optimize_reactor(storedata=storedata, sleep=sleep, pso_gen=pso_gen, ga=ga, type=type)
        read_col_data_store(name=type)
    else:
        best = optimize_reactor(storedata=storedata, sleep=sleep, pso_gen=pso_gen, ga=ga, type=type)
        get_data_from_hysys(best=best, sleep=0.5, type=type)
        if sensitivityanalysis:
            run_sensitivity_analysis_bestVector(sleep=sleep, best=best, type=type)


def run_sensitivity_analysis(sleep, type):
    Hycase = init_hysys()
    if type == 'cstr':
        reactor_name = 'R-100'
        sprd_name = 'CSTR_opt'
    elif type == 'pfr':
        reactor_name = 'PFR-100'
        sprd_name = 'PRF_opt'
    elif type == 'cstr2':
        reactor_name = 'R-100-2'
        sprd_name = 'CSTR_opt-2'
    elif type == 'isothermalcstr':
        reactor_name = 'R-100-3'
        sprd_name = 'CSTR_opt-3'
    reactor = Reactor(Hycase=Hycase, reactor_name=reactor_name, sprd_name=sprd_name, type=type)
    b_inlettemp = 60
    b_catalystweight = 0.025
    b_residencetime = 0.707
    b_reactorP = 4000
    b_methanolCOratio = 70.74
    while b_inlettemp <= 110:
        DVvector = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP, b_methanolCOratio]
        reactor.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                              reactorP=DVvector[3], methanolCOratio=DVvector[4],
                              sleep=sleep, type=type)
        reactor.reactor_results(storedata=True, type=type)
        b_inlettemp += 1
    read_col_data_store(name='TempSensiAnalysis_{}'.format(type))
    reactor = Reactor(Hycase=Hycase, reactor_name=reactor_name, sprd_name=sprd_name, type=type)
    b_inlettemp = 80
    b_catalystweight = 0.0001
    while b_catalystweight <= 0.05:
        DVvector = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP, b_methanolCOratio]
        reactor.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                              reactorP=DVvector[3], methanolCOratio=DVvector[4],
                              sleep=sleep, type=type)
        reactor.reactor_results(storedata=True, type=type)
        b_catalystweight += 0.001
    read_col_data_store(name='CatalystSensiAnalysis_{}'.format(type))
    reactor = Reactor(Hycase=Hycase, reactor_name=reactor_name, sprd_name=sprd_name, type=type)
    b_catalystweight = 0.025
    b_residencetime = 0.050
    while b_residencetime <= 2:
        DVvector = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP, b_methanolCOratio]
        reactor.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                              reactorP=DVvector[3], methanolCOratio=DVvector[4],
                              sleep=sleep, type=type)
        reactor.reactor_results(storedata=True, type=type)
        b_residencetime += 0.01
    read_col_data_store(name='ResidenceTSensiAnalysis_{}'.format(type))
    reactor = Reactor(Hycase=Hycase, reactor_name=reactor_name, sprd_name=sprd_name, type=type)
    b_residencetime = 0.707
    b_reactorP = 2000
    while b_reactorP <= 4000:
        DVvector = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP, b_methanolCOratio]
        reactor.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                              reactorP=DVvector[3], methanolCOratio=DVvector[4],
                              sleep=sleep, type=type)
        reactor.reactor_results(storedata=True, type=type)
        b_reactorP += 10
    read_col_data_store(name='ReactorPSensiAnalysis_{}'.format(type))
    reactor = Reactor(Hycase=Hycase, reactor_name=reactor_name, sprd_name=sprd_name, type=type)
    b_reactorP = 4000
    b_methanolCOratio = 2
    while b_methanolCOratio <= 100:
        DVvector = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP, b_methanolCOratio]
        reactor.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                              reactorP=DVvector[3], methanolCOratio=DVvector[4],
                              sleep=sleep, type=type)
        reactor.reactor_results(storedata=True, type=type)
        b_methanolCOratio += 5
    read_col_data_store(name='MethanolCOratioSensiAnalysis_{}'.format(type))


def run_sensitivity_analysis_bestVector(sleep, best, type):
    Hycase = init_hysys()
    if type == 'cstr':
        reactor_name = 'R-100'
        sprd_name = 'CSTR_opt'
    elif type == 'pfr':
        reactor_name = 'PFR-100'
        sprd_name = 'PRF_opt'
    elif type == 'cstr2':
        reactor_name = 'R-100-2'
        sprd_name = 'CSTR_opt-2'
    elif type == 'isothermalcstr':
        reactor_name = 'R-100-3'
        sprd_name = 'CSTR_opt-3'
    reactor = Reactor(Hycase=Hycase, reactor_name=reactor_name, sprd_name=sprd_name, type=type)
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
                              sleep=sleep, type=type)
        reactor.reactor_results(storedata=True, type=type)
    read_col_data_store(name='TempSensiAnalysisforBEST_{}'.format(type))
    reactor = Reactor(Hycase=Hycase, reactor_name=reactor_name, sprd_name=sprd_name, type=type)
    lowerbound = max(b_catalystweight - 0.01, 0.0001)
    upperbound = min(b_catalystweight + 0.01, 0.05)
    for i in np.arange(lowerbound, upperbound, 0.0001):
        DVvector = [b_inlettemp, i, b_residencetime, b_reactorP, b_methanolCOratio]
        reactor.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                              reactorP=DVvector[3], methanolCOratio=DVvector[4],
                              sleep=sleep, type=type)
        reactor.reactor_results(storedata=True, type=type)
    read_col_data_store(name='CatalystSensiAnalysisforBEST_{}'.format(type))
    reactor = Reactor(Hycase=Hycase, reactor_name=reactor_name, sprd_name=sprd_name, type=type)
    lowerbound = max(b_residencetime - 0.2, 0.05)
    upperbound = min(b_residencetime + 0.2, 2)
    for i in np.arange(lowerbound, upperbound, 0.01):
        DVvector = [b_inlettemp, b_catalystweight, i, b_reactorP, b_methanolCOratio]
        reactor.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                              reactorP=DVvector[3], methanolCOratio=DVvector[4],
                              sleep=sleep, type=type)
        reactor.reactor_results(storedata=True, type=type)
    read_col_data_store(name='ResidenceTSensiAnalysisforBEST_{}'.format(type))
    reactor = Reactor(Hycase=Hycase, reactor_name=reactor_name, sprd_name=sprd_name, type=type)
    lowerbound = max(b_reactorP - 200, 2000)
    upperbound = min(b_reactorP + 200, 4000)
    for i in np.arange(lowerbound, upperbound, 10):
        DVvector = [b_inlettemp, b_catalystweight, b_residencetime, i, b_methanolCOratio]
        reactor.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                              reactorP=DVvector[3], methanolCOratio=DVvector[4],
                              sleep=sleep, type=type)
        reactor.reactor_results(storedata=True, type=type)
    read_col_data_store(name='ReactorPSensiAnalysisforBEST_{}'.format(type))
    reactor = Reactor(Hycase=Hycase, reactor_name=reactor_name, sprd_name=sprd_name, type=type)
    lowerbound = max(b_methanolCOratio - 10, 2)
    upperbound = min(b_methanolCOratio + 10, 100)
    for i in np.arange(lowerbound, upperbound, 1):
        DVvector = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP, i]
        reactor.solve_reactor(inlettemp=DVvector[0], catatlystweight=DVvector[1], residencetime=DVvector[2],
                              reactorP=DVvector[3], methanolCOratio=DVvector[4],
                              sleep=sleep, type=type)
        reactor.reactor_results(storedata=True, type=type)
    read_col_data_store(name='MethanolCOratioSensiAnalysisforBEST_{}'.format(type))

run_sensitivity_analysis(sleep=0.5, type='cstr')
#run_ReactorOpt(storedata=False, sleep=0.5, pso_gen=100, ga=True, type='cstr', sensitivityanalysis=True)
#run_ReactorOpt(storedata=False, sleep=1, pso_gen=100, ga=True, type='pfr', sensitivityanalysis=True)
#run_ReactorOpt(storedata=False, sleep=0.5, pso_gen=100, ga=True, type='cstr2', sensitivityanalysis=True)
#run_ReactorOpt(storedata=False, sleep=0.5, pso_gen=100, ga=True, type='isothermalcstr', sensitivityanalysis=True)
# run_sensitivity_analysis(sleep=0.3)
#best = [89.83665875336027, 0.008469789091178227, 1.894972583420758, 4000, 12.956305834759327]
#get_data_from_hysys(best=best, sleep=0.3, type='pfr')
# run_sensitivity_analysis_bestVector(sleep=0.3, best=best, type='cstr')
