'''
This is python application to:
- read a time series data that consist of multiple column i.e time series data of multiple sensor
- plot all data in a time series plot graph
- select a segment of time series data (all column)
- calculate mean of the segment (all column)
- write all mean data to an open active excel file (need ms office excel)
'''


import PySimpleGUI as sg
import matplotlib.pyplot as plt
import matplotlib.pyplot as plt
from statistics import mean, StatisticsError
import numpy as np
from matplotlib.widgets import SpanSelector
import xlwings as xw #library to export to active open excel file
import matplotlib.colors as mcolors

import warnings
warnings.filterwarnings(action='ignore', category=UserWarning) # setting ignore as a parameter and further adding category

# write to excel file start from 4th row
# you can change this
counterrow = 4

AppFont = 'Any 8'
sg.theme('green')

layoutinputoutput = [
                    [sg.FileBrowse("TimeSeries Data File",file_types=(('Text Files', '*.txt'),)),
                    sg.Input(key='_FILE_DATATIMESERIES_BROWSE_',readonly=True,size=(100,100),enable_events=True)]
                    ]

windows = sg.Window('Get Mean Multiple Column Time Series Data',
                    layoutinputoutput,
                    size=(900,100), 
                    finalize=True,
                    resizable=False,
                    location=(300, 300),
                    )

# pyplot initial setting
plt.style.use('fivethirtyeight')
plt.rcParams['font.size'] = '7'
fig_floating, axs_floating = plt.subplots(layout="constrained", figsize=(35, 30), num='Mean of Selected Multiple Column Data')

# ----------------------------------------------------------------------------- #

def datagraph():
    
    # Load file data for each of WHS from calibration data file then plot in an axes
    data_timeseries=np.loadtxt(values['_FILE_DATATIMESERIES_BROWSE_'],skiprows=24)
    
    # choice of color for line graph 
    colorlist = list(mcolors.TABLEAU_COLORS.values())+list(mcolors.BASE_COLORS.values()) + list(mcolors.CSS4_COLORS.values())
     
    # jumlah kolom dalam file data kalibrasi = data_kalibrasi.shape[1]
    # plot all data in one axes
    for i in range(1, data_timeseries.shape[1]):
        axs_floating.plot(data_timeseries[:,0], data_timeseries[:,i], linewidth=0.75, color=colorlist[i-1])

    axs_floating.set_ylabel('Amplitude')
    axs_floating.set_xlabel('Data')
    axs_floating.set_title('Press left mouse button and drag '
                 'to select a region in thist chart',fontsize=8)
    axs_floating.grid(False)
    
    plt.grid(False)
    plt.show(block=False)
    
    return data_timeseries

def onselect(xmin, xmax):

    global counterrow
    outputdatagraph = datagraph() #output func datagraph
    columncounter = outputdatagraph.shape[1]-1 #jumlah column
    
    meanSliceSensorData = [i for i in range(columncounter)] #initialize array for mean every WHS data
    
    # get index x axis (time data) of slice data
    indmin, indmax = np.searchsorted(outputdatagraph[:,0], (xmin, xmax))
    
    try:
        for i in range( columncounter):
            # Slice each of column data based on index slice data of time data
            # then calculate mean of the slice data of each column
            # example:
            # for i = 0
            # sensor_1 = outputdatagraph[:,i+1] or 0 + 1 = 1
            # meanSliceSensorData[0] = mean of slice sensor 1 data
            meanSliceSensorData[i] = mean(outputdatagraph[:,i+1] [indmin:indmax])

        try:
            # write to excel file
            # start at 4th row, column 3 or C4
            # opsional, you can change this
            xw.Range((counterrow,3), (counterrow,7)).value = meanSliceSensorData
            counterrow = counterrow + 1

        except xw.XlwingsError as e:
            #print(e)
            e = str(e) #convert e type to string
            if e == "Couldn't find any active App!":
                string1 = "Open an analisa excel file \n"
                string2 = "Click Pause then click Quit to close this Debug Window"
                string3 = "Then reload TimeSeries Data File again"
                e = "\n" + e + "\n" + string1 + string2 + string3            
            else:
                e = e

            sg.Print('Error', sg.__file__, e, keep_on_top=False, wait=False)
            

    except StatisticsError as f:
        # This error (mean error) will raise if we only click one point
        # because mean need at least two point
        # one point click is to trigger clear the selector
        # so if this happened, counter row wont be added
        counterrow = counterrow


span = SpanSelector(
        axs_floating,
        onselect,
        "horizontal",
        useblit=True,
        props=dict(alpha=0.1, facecolor="tab:red"),
        interactive=True,
        drag_from_anywhere=True,

        # Press and release events triggered at the same coordinates
        # outside the selection will clear 
        # the selector, except when ignore_event_outside=True.
        # but this could raise statistic mean error
        ignore_event_outside=False
)

while True:
    
    #event, values = _VARS['window'].read()
    event, values = windows.read()
    if event in (sg.WIN_CLOSED, 'Cancel'):
        break
                       
    if event == '_FILE_DATATIMESERIES_BROWSE_':
        counterrow = 4 #reset 
        axs_floating.clear()
        datagraph()  
            
windows.close()
