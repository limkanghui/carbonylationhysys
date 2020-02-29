import win32com.client as win32
import numpy as np

def init_hysys():
    '''
    Connects to hysys and returns the HyCase object.
    Will link to the currently open Hysys case file.
    :return: HyCase object
    '''
    print('Connecting to the Aspen Hysys App ... ')
    HyApp = win32.Dispatch('HYSYS.Application')

    HyCase = HyApp.ActiveDocument

    # 04 Aspen Hysys Environment Visible
    HyCase.Visible = 1

    # 05 Aspen Hysys File Name
    HySysFile = HyCase.Title.Value
    print(' ')
    print('HySys File: ----------  ', HySysFile)

    # 06 Aspen Hysys Fluid Package Name
    package_name = HyCase.Flowsheet.FluidPackage.PropertyPackageName
    print('HySys Fluid Package: ---  ', package_name)
    print(' ')
    return HyCase







