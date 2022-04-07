# import libraries
import paramiko
from pywinauto.application import Application
from time import sleep, time
from math import ceil



############## Connect to Raspberry Pi and open OMNIC ##############

# connect to Raspberry Pi
pi = paramiko.SSHClient()
# accept the Pi's host key automatically 
pi.set_missing_host_key_policy(paramiko.AutoAddPolicy())
# connect to the Pi (change username and password if needed)
pi.connect(hostname = "raspberrypi.local",
                   username = "pi", password = "spectraflowpi")

# connect to OMNIC; open it if not found
try:
    omnic = Application().connect(path='C:/Program Files (x86)/omnic/omnic32.exe',timeout=0)
except:
    omnic = Application().start('C:/Program Files (x86)/omnic/omnic32.exe',timeout=2)
main_dlg = omnic.window(title_re='OMNIC*') # locate the main window/dialog
main_dlg.maximize() # maximize the screen



############## Define several functions for later use ##############

# function to send command from PC to Pi
def run_Pi(cmd):
    # variables are not used, mainly for sending back signal to PC
    stdin,stdout,stderr = pi.exec_command(cmd)
    stdout.read() # needed to let PC know when Pi finishes the command
# Note: format for exec_command is (cd path; python program.py funx)
# cd path: change directory to where the program is on Rasberry Pi
# python program.py: access the Python program with that file name
# funx: name of the function in that Python program


    
# function that stores and retrieves scan time when given no. of scans and resolution
def scan_time(no_of_scans,resolution):
    switcher = {
        '0.125.': 6.7,
        '0.25.': 3.53,
        '0.5.': 1.96,
        '1.': 1.16,
        '2.': 1.12,
        '4.': 0.72,
        '6.': 0.57,
        '8.': 0.52,
        '16.': 0.42,
        '32.': 0.37}
    # pad 5s to the calculated scan time
    return ceil(3+no_of_scans*switcher.get(resolution))
    

# function close all popup windows
def close_popup():
    try:
        omnic.top_window().Cancel.click()
    except:
        pass


# function to create a new window
def new_window():
    main_dlg.menu_select('Window->New Window')
    omnic.Dialog.OK.click()


# function to collect reference/analyte
def collect(loops,sample):
    main_dlg.menu_select('Collect->Experiment Setup')
    omnic.Dialog.ComboBox4.Edit.set_edit_text(sample+'_')
    omnic.Dialog.OK.click()
    for j in range(loops):
        main_dlg.menu_select('Collect->Loop1')
        sleep(scan_time(no_of_scans,resolution))


# function to ask for user's inputs
def get_user_inputs():
    # remind user to set up Experiment Setup before continue
    print('\nPlease ensure to choose where to save file and set up the experiment.')
    input('Press ENTER when done to continue...')
    
    close_popup() # close all popup windows before we start
    
    # ask for number of samples
    while True:
        try:
            no_of_samples = int(input('Enter the number of samples you want to collect: '))
        except ValueError:
            print('The number of samples has to be an integer.')
            continue
        else:
            break
        
    # ask for the number of loops for reference and analyte / aka LoopN
    while True:
        try:
            loops = int(input('Enter the number of loops you want for each sample: '))
        except ValueError:
            print('The number of loops has to be an integer.')
            continue
        else:
            break
    
    # ask for time interval between each run
    while True:
        try:
            time_int = float(input('Enter the time inteval between each run in minutes: '))
        except ValueError:
            print('Time interval has to be a number.')
            continue
        else:
            break
    
    # open Experiment Setup to grab no. of scans and resolution
    close_popup() # close any popup in case user has one opened to avoid error 
    main_dlg.menu_select('Collect->Experiment Setup')
    no_of_scans = int(omnic.Dialog.Edit8.window_text())
    resolution = omnic.Dialog.ComboBox1.window_text()
    omnic.Dialog.OK.click()
    
    return no_of_scans,resolution,no_of_samples,loops,time_int

input('Hit ENTER when you are ready')

close_popup() # close any popup in case user has one opened to avoid error 
main_dlg.menu_select('Collect->Experiment Setup')
no_of_scans = int(omnic.Dialog.Edit8.window_text())
resolution = omnic.Dialog.ComboBox1.window_text()
omnic.Dialog.OK.click()
no_of_samples = 100
loops = 7
time_int = 26
n = 1

############## Get user's input and calibrate if necessary ##############

# call function to ask for user's inputs
no_of_scans,resolution,no_of_samples,loops,time_int = get_user_inputs()

# call function for initial fill up of cell
run_Pi('cd UTDesign; python pump.py start')


############## Automatically control pump and collect spectra ##############


# loop for the number of runs
for i in range(n,no_of_samples+n):
    run_Pi('cd UTDesign; python pump.py EM_on_2nd_pump_on')
    run_Pi('cd UTDesign; python pump.py main_pump_on')
    print('Collecting analyte #{}...'.format(i))
   
    if i == n:
        #get reference run
        run_Pi('cd UTDesign; python pump.py EM_off')
        run_Pi('cd UTDesign; python pump.py sonicate_2nd_pump_on')
        collect(loops,'analyte_before'+str(i))
    
        
    else:
       #get results of several test with EM on and EM off
        run_Pi('cd UTDesign; python pump.py EM_on_2nd_pump_on')
        collect(loops,'analyte_after'+str(i))
        run_Pi('cd UTDesign; python pump.py EM_off')
        run_Pi('cd UTDesign; python pump.py sonicate_2nd_pump_on')
        collect(loops,'analyte_before'+str(i))
    sleep(60*time_int) # sleep between runs
    
    # create new window for new run, except the last
    if i != no_of_samples:
        new_window()

# end SSH connection before closing
pi.close()