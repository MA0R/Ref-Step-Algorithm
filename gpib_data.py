"""
Everything about GPIB control for the Ref Step algorithm.
The GUI should only allow one GPIB thread at a time.
Thread also writes raw data file to an excel sheet.
Log files are written at event termination from graphframe.py.
"""

import stuff
import csv
import time
import wx
import numpy as np
import gpib_inst
import openpyxl

class GPIBThreadF(stuff.WorkerThread):
    """
    Ref step main thread
    """
    def __init__(self, notify_window, EVT, param, data,start_time,OverideSafety):
        stuff.WorkerThread.__init__(self, notify_window, EVT, param, data,start_time,OverideSafety)
        #param[0] holds self.inst_bus, the chosen real or simulated visa
        self.inst_bus = param[0] #visa or simulated visa2
        self.grid = param[1]
        self.start_row = param[2]
        self.stop_row = param[3]
        self.dvm_nordgs_col = param[4]
        self.voltmeter = param[5]
        self.dvm_range_col = param[6]
        self.sourceX = param[7]
        self.sourceX_range_col = param[8]
        self.sourceX_setting_col = param[9]
        self.sourceS = param[10]
        self.sourceS_range_col = param[11]
        self.sourceS_setting_col = param[12]
        self.delay_col = param[13]
        self.Analysis_file_name = param[14]
        self.OverideSafety = OverideSafety
        self.MadeSafe = False

        time_bit = time.strftime("%Y.%m.%d.%H.%M.%S",time.localtime())
        log_file_name = 'log.'+time_bit+".txt"
        self.raw_file_name = 'raw.'+time_bit
        self.wb = openpyxl.Workbook()
        self.sh = self.wb.active
        self.logfile = open(log_file_name,'w')
        first_line = [self.read_grid_cell(6, i) for i in range(self.grid.GetNumberCols())]
        first_line = first_line+['start time','end time','readings...']

        for i in range(6):
            self.sh.append( [self.read_grid_cell(i,column) for column in range(10)])
            #append empty things, but they can contain more useful info later
            #dates, instruments, start time finish time?
            #this skeeps the analysis sheets compatible with the normal input files.
        self.sh.append(first_line)
        #must be the last command:
        self.start()

        
    def PrintSave(self, text):
        """
        Prints a string, and saves to the log file too.
        """
        if self._want_abort: #stop the pointless printing if the thread wants to abort
            return
        else:
            self.logfile.write(str(text)+"\n")
            print(str(text))
            
    def com(self, command,send_item=None):
        """
        Combine the instrument command with an item (if one is specified).
        All commands go through this function, as it wraps each command
        with a check to see if the thread should abort. Saves having to
        surround each instrument send command with an if statement.
        Will also recieve the status of the command from the instrument
        class, and if it failes this will flag the thread to abort.
        """
        if not self._want_abort:
            if send_item != None:
                sucess,val,string = command(send_item)
                self.PrintSave(string)
                if sucess == False:
                    self._want_abort = 1
                return val
            else:
                sucess,val,string = command()
                self.PrintSave(string)
                if sucess == False:
                    self._want_abort = 1
                return val
        else:
            return 0
        
    def MakeSafe(self):
        """
        Make the instruments safe, bypass the self.com
        command in order to pass it directly to the instruments,
        that way recording weather or not the making safe worked
        for the instruments. Reports to GUI if it is unsafe.
        """
        sucessX,valX,stringX = self.sourceX.make_safe()
        sucessS,valS,stringS = self.sourceS.make_safe()
        sucessM,valM,stringM = self.meter.make_safe()
        self.PrintSave("Make safe sent. status is:")
        self.PrintSave("SourceX {}\nSourceS {}\nMeter {}".format(sucessX,sucessS,sucessM))
        if not all([sucessX,sucessS,sucessM]):
            wx.PostEvent(self._notify_window, stuff.ResultEvent(self.EVT, "UNSAFE"))
        else:
            wx.PostEvent(self._notify_window, stuff.ResultEvent(self.EVT, None))
            
    def Error_string_maker(self):
        """
        Reads all errors in the instruments, and appends them together in a string.
        If there are errors during the running of the code, they will be printed
        on the left of the table as a warning flag.
        """
        #somehow still prints "0" when the instruments have no error.
        string = " "
        #query instrument errors, and save individual error strings.
        self.com(self.sourceX.query_error)
        x_esr = str(self.com(self.sourceX.read_instrument))
        self.PrintSave('sourceX ESR = '+x_esr)
        if x_esr != self.sourceX.com['NoError']: string = 'Source X: '+x_esr
        self.com(self.sourceS.query_error)
        s_esr = str(self.com(self.sourceS.read_instrument))
        self.PrintSave('sourceS ESR = '+s_esr)
        if s_esr != self.sourceS.com['NoError']: string = string +' source S: '+s_esr
        self.com(self.voltmeter.query_error)
        m_esr = str(self.com(self.voltmeter.read_instrument))
        self.PrintSave('meter ESR = '+m_esr)
        if m_esr != self.voltmeter.com['NoError']: string = string +' meter: ' +m_esr
        self.PrintSave('')
        return string

    def set_grid_val(self,row,col,data):
        """ Set the value of a grid cell"""
        wx.CallAfter(self.grid.SetCellValue, row, col, str(data))
        
    def read_grid_cell(self,row,col):
        """ Read a grid cell"""
        value = 0
        if not self._want_abort:
            value = self.grid.GetCellValue(row, col)
        return value
        
    def end(self):
        """
        Function to close files and post an event to the main grid,
        letting it know that the thread has ended.
        """
        self.com(self.sourceS.Standby)
        self.com(self.sourceX.Standby)
        self.logfile.close()
        self.wb.save(filename = str(self.raw_file_name+'.xlsx'))
        wx.CallAfter(self.Analysis_file_name.SetValue,self.raw_file_name+'.xlsx')
        wx.PostEvent(self._notify_window, stuff.ResultEvent(self.EVT, 'GPIB ended'))

    def wait(self,wait_time):
        """
        safely loop until time is up, checking for abort
        """
        t=time.time()
        while time.time()<t+wait_time:
            if self._want_abort: #has the stop button been pushed?
                break
        return
    
    def initialise_instruments(self):
        self.com(self.voltmeter.reset_instrument)
        self.com(self.voltmeter.initialise_instrument)
        self.com(self.sourceX.reset_instrument)
        self.com(self.sourceX.initialise_instrument)
        self.com(self.sourceS.reset_instrument)
        self.com(self.sourceS.initialise_instrument)
        
        self.PrintSave('')
        self.com(self.sourceX.query_error)
        self.PrintSave('sourceX ESR = '+str(self.com(self.sourceX.read_instrument)))
        self.com(self.sourceS.query_error)
        self.PrintSave('sourceS ESR = '+str(self.com(self.sourceS.read_instrument)))
        self.com(self.voltmeter.query_error)
        self.PrintSave('meter ESR = '+str(self.com(self.voltmeter.read_instrument)))
        self.PrintSave('')
        
    def set_source_ranges(self,row):
        """
        Read the grid and obtain values for the ranges of the sources.
        If the ranges are different to the ranges set previously,
        set the current range (avoid needlessly sending range commands).
        """
        #read ranges of instruments.
        #If current range is not equal to the previous, apply range setting.
        sourceS_range = self.read_grid_cell(row, self.sourceS_range_col)
        if sourceS_range != self.read_grid_cell(row-1, self.sourceS_range_col):
            self.com(self.sourceS.set_DCrange,sourceS_range)
        sourceS_value = self.read_grid_cell(row, self.sourceS_setting_col)
        self.com(self.sourceS.set_DCvalue,sourceS_value)
        
        sourceX_range = self.read_grid_cell(row, self.sourceX_range_col)
        if sourceX_range != self.read_grid_cell(row-1, self.sourceX_range_col):
            self.com(self.sourceX.set_DCrange,sourceX_range)
        sourceX_value = self.read_grid_cell(row, self.sourceX_setting_col)
        self.com(self.sourceX.set_DCvalue,sourceX_value)
        
    def set_meter_range(self,row):
        """
        Set the range for the meter, but also return the value
        of the max of the range. used in safety checks.
        """
        dvm_range = self.read_grid_cell(row, self.dvm_range_col)
        if dvm_range != self.read_grid_cell(row-1, self.dvm_range_col):
            self.com(self.voltmeter.set_DCrange,dvm_range)
            #fixes range as specified in sheet
        range_max = 0 # so it returns SOMEthing
        try:
            range_index = [some_range[2] for some_range in self.voltmeter.range].index(dvm_range)
            range_max = self.voltmeter.range[range_index][1]
        except:
            self.PrintSave("could not identify "+str(dvm_range)+" as a valid range")
            self._want_abort = 1
        return range_max
    
    def do_readings(self,delay_time2,range_max,nominal_reading,nordgs):
        """
        Do a single reading on the meter, including safety checks:
        report a reading as faulty if it is 0.1% of the top of the current
        range away from the nominal value. If the difference of the nominal
        reading to the actual reading is 5% of the max range value of more then abort.
        """
        these_readings = []
        error_strings = " "
        
        for i in range(nordgs):
            if self._want_abort:
                break
            self.wait(delay_time2)
            #the pause is placed before the reading (retreaving of value),
            #as first the DVM must integrate (or do wahtever it does) then you can read
            
            if self.voltmeter.com["SingleMsmntSetup"] != 'SKIP':
                #only send the extra command if the voltmeter has somethings to be sent
                self.com(self.voltmeter.SingleMsmntSetup)

            a = float(self.com(self.voltmeter.read_instrument))
            #meter class provides string for evaluation

            if self.OverideSafety == False: #Only do these checks if we dont overide safety?
                if abs((a-nominal_reading)) > abs(0.05*range_max):
                    #taking abs to allow for negative ranging.
                    #readings is more than 5% away from nominal, abort
                    error_strings = error_strings+' reading '+str(i)+' is more than 5% from nominal.'
                    self.PrintSave("Readings exceeded 5% from nominal, aborting. Check settle times, ranges, and connections")
                    self._want_abort = 1
                elif abs((a-nominal_reading)) > abs(0.001*range_max):
                    #readings is more than 0.1% of range away from nominal, issue a warning
                    error_strings = error_strings+' reading '+str(i)+' is more than 0.1% from nominal.'
                    self.PrintSave("Readings exceeded 0.1% from nominal. Check settle times, ranges, and connections")
            
            self.wait(self.voltmeter.measure_seperation) #wait if voltmeter needs to not be read continuously
            self.PrintSave(time.strftime("%Y.%m.%d.%H.%M.%S, ",time.localtime()) + repr(a) + '\n')
            self.data.add_to_list(a)
            these_readings.append(a)
        return these_readings,error_strings
    
    def print_instrument_status(self,row):
        """Read the status of all instruments, print to grid"""
        self.com(self.voltmeter.inst_status)
        self.set_grid_val(row, self.dvm_nordgs_col+3, str(self.com(self.voltmeter.read_instrument)))
        #force string of value as it needs to go into the grid and grid only takes in strings.
        
        self.com(self.sourceS.inst_status)
        self.set_grid_val(row, self.dvm_nordgs_col+4, str(self.com(self.sourceS.read_instrument)))
        
        self.com(self.sourceX.inst_status)
        self.set_grid_val(row, self.dvm_nordgs_col+5, str(self.com(self.sourceX.read_instrument)))

    def run(self):
        """
        Main thread, reads through the table and executes commands.
        """
        self.com(self.voltmeter.create_instrument)
        self.com(self.sourceX.create_instrument)
        self.com(self.sourceS.create_instrument)
        #If the initialisations failed, stop the thread immediatly.
        #There is no point running make safe if the instruments arent there.
        if self._want_abort:
            #Notify the window that the thread ended, so buttons are enabled.
            wx.PostEvent(self._notify_window, stuff.ResultEvent(self.EVT, 'GPIB ended'))
            return
        
        if self.OverideSafety == False: #Then do the safety checks.
            state = self.SafetyCheck()
            if state != 'clear': self._want_abort = True
            self.PrintSave('safety checks failed, making safe, aborting')
            
        #initialise instruments
        self.initialise_instruments()
        
        for row in range(self.start_row,self.stop_row + 1):
            if self._want_abort:
                break #Breaks and skips to the end, where it runs "self.end()".
            
            #do the row highlighting, force a refresh.
            self.PrintSave("Spread sheet row "+str(int(row)+1))
            wx.CallAfter(self.grid.SelectRow,row)
            wx.CallAfter(self.grid.ForceRefresh)
            wx.CallAfter(self.grid.Update)

            #prepare ranges for sources
            self.set_source_ranges(row)
            #operate sources
            self.com(self.sourceS.Operate)
            self.com(self.sourceX.Operate)

            #read errors in meters. The error string later has more warnings appended to it. 
            error_strings = self.Error_string_maker()
            
            #need to determine maximum of DVM range to find accuracy threshold
            range_max = self.set_meter_range(row)

            delay_time = int(float(self.read_grid_cell(row,self.delay_col)))
            #delay for settle time
            delay_time2 = float(self.read_grid_cell(row,self.delay_col+1))
            #delay for DVM between mesmnts
            nominal_reading  = float(self.read_grid_cell(row,self.dvm_range_col+1))
            #nominal reading for comparison to actual reading
            
            self.com(self.voltmeter.MeasureSetup)
            
            #wait desired time for instruments to settle
            self.wait(delay_time)
            
            these_readings = [] #array for readings from a single table row.
            before_msmnt_time =time.time() #start time of measuremtns.
            nordgs = int(float(self.read_grid_cell(row, self.dvm_nordgs_col) ))
            #Do the readings
            these_readings,issues = self.do_readings(delay_time2,range_max,nominal_reading,nordgs)
            
            #Update the full error string
            error_strings = error_strings + issues
            #print the error report of this data sequence to the table
            self.set_grid_val(row, 0,str(error_strings))
            
            if len(these_readings)>1:
                data_stdev = np.std(these_readings)
                data_mean = np.mean(these_readings)
            else:
                data_stdev = 'No readings' #avoid infinity for stdev
                data_mean = 'No readings'

            #create array to send to csv data sheet
            after_msmnt_time = time.time()
            Time = time.localtime() #time at end of msmnt
            row_time = str(Time[0])+str(Time[1])+str(Time[2])+str(Time[3])+str(Time[4])
 
            #print results
            self.set_grid_val(row, self.dvm_nordgs_col + 6,repr(data_mean))
            self.set_grid_val(row, self.dvm_nordgs_col + 7,repr(data_stdev))
            
            self.print_instrument_status(row)
                
            #put csv reading at end so all values are on table.
            csv_line = [self.read_grid_cell(row, i) for i in range(self.grid.GetNumberCols())]
            csv_line = csv_line+[before_msmnt_time,after_msmnt_time]+these_readings
            self.sh.append(csv_line)
     
        self.end()
        

    def initialchecks(self,instrument):
        """
        Short function used in the safety check routine, given an instrument it resets
        the instrument, does initialisation and reads errors. Will continuously read
        errors until the meter returns the no-error string. This can cause problems if
        the no-error string is inputed wrongly.
        """
        self.com(instrument.reset_instrument) #reset voltmeter
        self.com(instrument.initialise_instrument)
        self.com(instrument.query_error)
        error = self.com(instrument.read_instrument)
        self.PrintSave(error)
        if error != instrument.com['NoError']:
            self.PrintSave('Error '+str(error)+' while resetting instrument, or queriying status')
            return 'failed'
        else: return 'clear'
        
    def CheckInstruments(self,*args):
        """
        Given any number of instruments, will read through each and PrintSave errors.
        Each instrumetn is continuously interrogated until it returns its no-error string.
        """
        for instrument in args:
            self.com(instrument.query_error)
            error = self.com(instrument.read_instrument())
            if error != instrument.com['NoError']:
                while error != instrument.com['NoError'] and not self._want_abort:
                    #loop as there could be several errors
                    self.PrintSave('Error '+str(error)+' in instrument '+instrument.com['label'])
                    self.com(instrument.query_error)
                    error = self.com(instrument.read_instrument)
                return 'failed'
        return 'clear'

            
    def SafetyCheck(self):
        """
        A linear function for the determination of some basic safety settings.
        It can be overriden.
        Will check: errors, resetting, outputting zero volts, ouputting 1 volt.
        If this is sucessful it will also determine if the wires are connected the wrong way.
        """
        #determine how long to wait for an output of 1V to settle. do this
        #by reading ranges that include 1V, and finding maximum delay

        
        SettleTime = 5
        
        state = 'clear'
        state = self.initialchecks(self.voltmeter)
        if state == 'clear': state = self.initialchecks(self.sourceX)
        if state == 'clear': state = self.initialchecks(self.sourceS)
        if state == 'clear':
            self.com(self.sourceX.Standby)
            self.com(self.sourceS.Standby)
            state = self.CheckInstruments(self.sourceX,self.sourceS)
        if state == 'clear':
            self.com(self.sourceX.set_DCvalue,0)
            self.com(self.sourceS.set_DCvalue,0)
            state = self.CheckInstruments(self.sourceX,self.sourceS)
        if state == 'clear':
            self.PrintSave('testing voltage is setting correctly, and zeros')
            
            self.com(self.sourceX.Operate)
            self.com(self.sourceS.Operate)
            if self.voltmeter.com["SingleMsmntSetup"] != 'SKIP':
                #only send the extra command if the voltmeter has somethings to be sent
                self.com(self.voltmeter.SingleMsmntSetup)
            self.wait(SettleTime) #let instrumetns settle
            reading = float(self.com(self.voltmeter.read_instrument))
            if reading>5e-3:
                self.PrintSave('Expect reading to be near zero (less than 5e-3), but it is '+str(reading)+' check zeros')
                state = 'failed'
                self.PrintSave(state)
            state = self.CheckInstruments(self.sourceX,self.sourceS)
        if state == 'clear':
            self.PrintSave('testing voltage setting correctly')
            self.com(self.sourceX.Standby)
            #find correct range for 1volt
            for r in self.sourceX.range[::-1]:
                if r[1]>=1:
                    r_use = r
            self.com(self.sourceX.set_DCrange,r_use[2])
            self.com(self.sourceX.set_DCvalue,1)
            self.com(self.sourceX.Operate)
            if self.voltmeter.com["SingleMsmntSetup"] != 'SKIP':
                #only send the extra command if the voltmeter has somethings to be sent
                self.com(self.voltmeter.SingleMsmntSetup)
            self.wait(SettleTime) #let instrumetns settle
            reading = float(self.com(self.voltmeter.read_instrument))
            if round(reading,2)!= 1.00:
                self.PrintSave('Source X set to 1, expect reading to be near 1, but it is '+str(reading)+' check output load on X and SettleTime')
                if round(reading,2) == -1:
                    self.PrintSave('reading negative, check if wires are the right way')
                state = 'failed'
                return state
            state = self.CheckInstruments(self.sourceX,self.sourceS)
        if state == 'clear':
            self.com(self.sourceX.Standby)
            self.com(self.sourceX.set_DCvalue,0)
            self.com(self.sourceX.Operate)
            self.com(self.sourceS.Standby)
            for r in self.sourceS.range[::-1]:
                if r[1]>=1:
                    r_use = r
            self.com(self.sourceS.set_DCrange,r_use[2])
            self.com(self.sourceS.set_DCvalue,1)
            self.com(self.sourceS.Operate)
            if self.voltmeter.com["SingleMsmntSetup"] != 'SKIP':
                #only send the extra command if the voltmeter has somethings to be sent
                self.com(self.voltmeter.SingleMsmntSetup)
            self.wait(SettleTime) #let instrumetns settle
            reading = float(self.com(self.voltmeter.read_instrument))
            if round(reading,2)!= -1.00:
                self.PrintSave('Set source S to 1, expect reading to be near -1, but it is '+str(reading)+' output load on S and SettleTime')
                state = 'failed'
                return state
            state = self.CheckInstruments(self.sourceX,self.sourceS)
            
        return state
        
