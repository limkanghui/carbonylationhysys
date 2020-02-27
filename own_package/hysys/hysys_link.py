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


def get_carbonylation_feed(Hycase, feed_name):
    feed = Hycase.Flowsheet.Streams.Item(feed_name)
    T0 = feed.TemperatureValue + 273.15
    p0 = feed.PressureValue
    F0_temp = np.array(feed.ComponentMolarFLowValue) * 3600  # kmol/h
    F0 = np.concatenate((F0_temp[[2,0,3,1]], np.array(F0_temp[4:].sum())[None]))[:-1]
    mdot0_temp = np.array(feed.ComponentMassFLowValue) * 3600
    mdot0 = np.concatenate((mdot0_temp[[2,0,3,1]], np.array(mdot0_temp[4:].sum())[None]))[:-1]
    mr = mdot0/F0
    rho0 = feed.MassDensityValue

    return prepare_flow_data(Hycase=None, feed_name=feed_name, F0=F0, mdot0=mdot0, mr=mr, rho0=rho0, T0=T0, p0=p0)


def set_hydrolysis_reactor(output_eq, sprdsht):
    sprdsht.Cell("B1").CellValue = output_eq['X_eq']*100  # Set percentage conversion for conversion reactor
    sprdsht.Cell("B2").CellValue = output_eq['T_eq'] - 273.15  # Set output temperature in degrees C







