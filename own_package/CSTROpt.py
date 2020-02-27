from own_package.hysys.hysys_CSTR import CSTR
from own_package.hysys.hysys_link import init_hysys
from own_package.pso_ga import pso_ga

def optimize_CSTR():
    Hycase = init_hysys()
    cstr = CSTR(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt')
    b_inlettemp = [50, 150]
    b_catalystweight = [0.001, 0.05]
    b_residencetime = [0.05, 2]
    b_reactorP = [2000,4000]
    p_store = [b_inlettemp, b_catalystweight, b_residencetime, b_reactorP]

    params = {'c1': 1.5, 'c2': 1.5, 'wmin': 0.4, 'wmax': 0.9,
              'ga_iter_min': 5, 'ga_iter_max': 20, 'iter_gamma': 10,
              'ga_num_min': 10, 'ga_num_max': 20, 'num_beta': 15,
              'tourn_size': 3, 'cxpd': 0.5, 'mutpd': 0.05, 'indpd': 0.5, 'eta': 0.5,
              'pso_iter': 200, 'swarm_size': 50}

    pmin = [x[0] for x in p_store]
    pmax = [x[1] for x in p_store]

    smin = [abs(x - y) * 0.01 for x, y in zip(pmin, pmax)]
    smax = [abs(x - y) * 0.5 for x, y in zip(pmin, pmax)]

    def func(individual):
        nonlocal CSTR
        inlettemp, catalystweight, residencetime, reactorP = individual
        CSTR.solve_reactor(inlettemp=individual[0], catatlystweight=individual[1], residencetime=individual[2], reactorP=individual[3])
        return (CSTR.reactor_results(),)

    pso_ga(func=func, pmin=pmin, pmax=pmax,
           smin=smin, smax=smax,
           int_idx=None, params=params, ga=True)


optimize_CSTR()

