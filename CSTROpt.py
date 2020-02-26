import numpy as np

from hysys.hysys_CSTR import CSTR
from hysys.hysys_link import init_hysys
import pso_ga


def optimize_CSTR():
    Hycase = init_hysys()
    cstr = CSTR(Hycase=Hycase, reactor_name='R-100', sprd_name='CSTR_opt')
    b_inlettemp = [50, 150]
    b_catalystweight = [0.001, 0.05]
    b_nt = [3, 10]
    b_feedloc = [0,1]
    p_store = [b_active_spec_1, b_rebo_p, b_nt, b_feedloc]

    params = {'c1': 1.5, 'c2': 1.5, 'wmin': 0.4, 'wmax': 0.9,
              'ga_iter_min': 5, 'ga_iter_max': 20, 'iter_gamma': 10,
              'ga_num_min': 10, 'ga_num_max': 20, 'num_beta': 15,
              'tourn_size': 3, 'cxpd': 0.5, 'mutpd': 0.05, 'indpd': 0.5, 'eta': 0.5,
              'pso_iter': 200, 'swarm_size': 50}

    pmin = [x[0] for x in p_store]
    pmax = [x[1] for x in p_store]

    smin = [abs(x - y) * 0.01 for x, y in zip(pmin, pmax)]
    smax = [abs(x - y) * 0.5 for x, y in zip(pmin, pmax)]
    num_evals = 0
    def func(individual):
        nonlocal num_evals
        num_evals += 1
        active_specs_1 = individual[0]
        active_specs_2 = 1
        del_p = 0.3  #kPa
        distcolumn.solve_column(active_spec_goals=[active_specs_1, active_specs_2], del_p=del_p,
                                rebo_p=individual[1], number_of_trays=individual[2], feed_frac=individual[3])
        if num_evals % 10 == 0:
            print('Total evals: {}. Current Eval Converged: {}'.format(num_evals, distcolumn.status))
        return (distcolumn.column_results(),)

    pso_ga(func=func, pmin=pmin, pmax=pmax,
           smin=smin, smax=smax,
           int_idx=[2], params=params, ga=True)


if __name__=='__main__':
    optimize_column()

