"""
Everything about GPIB control for the Ref Step algorithm. The GUI should only allow one GPIB thread at a
time.
Thread also writes raw data file to an excel sheet. Log files are written at event termination from graphframe.py.
"""

import stuff
import csv
import time
import wx
import visa #use inst_bus instead of visa so the calling application
#can provide the simulated (visa2) or real version of visa
import visa2
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

        log_file_name = 'log.'+str(start_time[0])+'.'+str(start_time[1])+'.'+str(start_time[2])+'.'+str(start_time[3])+'.'+str(start_time[4])+".txt"
        self.raw_file_name = 'raw.'+str(start_time[0])+'.'+str(start_time[1])+'.'+str(start_time[2])+'.'+str(start_time[3])+'.'+str(start_time[4])
        self.wb = openpyxl.Workbook()
        self.sh = self.wb.active
        self.logfile = open(log_file_name,'w')
        first_line = [self.grid.GetCellValue(6, i) for i in range(self.grid.GetNumberCols())]+['start time','end time','readings...']

        for i in range(6):
            self.sh.append( [ self.grid.GetCellValue(i,column) for column in range(10) ])
            #append empty things, but they can contain more useful info later
            #dates, instruments, start time finish time? this skeeps the analysis sheets compatible with the normal input files.
        self.sh.append(first_line)
            
        self.rm = self.inst_bus.ResourceManager() #one resource manager for this thread
        eval(self.voltmeter.create_instrument()) #create voltmeter in resource manager, before thread start
        eval(self.sourceX.create_instrument()) #create sourceX
        eval(self.sourceS.create_instrument())

        self.start() #important that this is the last statement of initialisation. goes to run()

        
    def PrintSave(self, text):
        """
        Prints a string, and saves to the log file too.
        """
        if self._want_abort: #stop the pointless printing if the thread wants to abort
            return
        else:
            self.logfile.write(str(text))
            print(str(text))

    def MakeSafe(self):
        """
        MakeSafe executes the make safe commands straight to the instruments, bypassing the try_command structure.
        Incase it it failed making safe, the Make safe button will turn red (posts "UNSAFE" to the main thread) and state which instruments the make safe failed at.
        The program only attempts the make safe once and terminates the thread.
        """
        if self.MadeSafe ==False:
            sucess = True
            #executes the make safe sequence on each instrument, bypassing the try_command
            insts = [self.voltmeter.com['label'],self.sourceX.com['label'],self.sourceS.com['label']] #meter label, x label, s label
            safes = [self.voltmeter.com['Make_Safe'],self.sourceX.com['Make_Safe'],self.sourceS.com['Make_Safe']] #make ssafe routines for each
            for l,s in zip(insts, safes):
                try:
                    exec('self.'+str(l)+'.write("'+str(s)+'")')
                except self.inst_bus.VisaIOError:
                    sucess = False
                    self.PrintSave("DID NOT SUCESFULLY MAKE "+str(l)+" SAFE")
            if sucess == False:
                wx.PostEvent(self._notify_window, stuff.ResultEvent(self.EVT, "UNSAFE"))
            else:
                wx.PostEvent(self._notify_window, stuff.ResultEvent(self.EVT, None))
            
        self.MadeSafe = True
        
    def try_command(self, command, error_message, choice,SettleTime):
        """
        Uses a try: except structure whenever GPIB is called by 'command' and
        returns 'error_message' if the try fails. This should give more readable
        instrument control code by removing the clutter of the try block. The
        returned gpib needs to be assigned to self.xxx if xxx is to be used later
        in try_command. choice is either 'ev' for using eval or 'ex' for using exec.
        exec executes the string as a statement while eval returns a value.
        In case of a visa failure, calls MakeSafe so if it is possible some instruments will be safe.
        """
        
        try:
            if choice == 'ev':
                if self._want_abort:
                    if self.MadeSafe == False:
                        self.MakeSafe()
                    return 0
                else:
                    #when reading the instrument, there is no settle time required.
                    gpib = eval(command)
                    string = str(time.strftime("%Y.%m.%d.%H.%M.%S, ",time.localtime())) + str(command)
                    self.PrintSave(string)
                    return str(gpib)
            elif choice == 'ex':
                if self._want_abort:
                    if self.MadeSafe == False:
                        self.MakeSafe()
                    return
                else:
                    exec command
                    time.sleep(SettleTime)
                    string = str(time.strftime("%Y.%m.%d.%H.%M.%S, ",time.localtime())) + str(command)
                    self.PrintSave(string)
                    return
        except self.inst_bus.VisaIOError:
            if self.MadeSafe == False:
                self.MakeSafe()
            string = str(time.strftime("%Y.%m.%d.%H.%M.%S, ",time.localtime()) + error_message)
            self.PrintSave(string)
            #wx.PostEvent(self._notify_window, stuff.ResultEvent(self.EVT, None))
            self._want_abort = 1
            return
        
    def Error_string_maker(self):
        """
        Reads all errors in the instruments, and appends them together in a string.
        If there are errors during the running of the code, they will be printed
        on the left of the table as a warning flag.
        """
        #somehow still prints "0" when the instruments have no error.
        string = " "
        #query instrument errors, and save individual error strings.
        eval(self.sourceX.query_error())
        x_esr = str(eval(self.sourceX.read_instrument()))
        self.PrintSave('sourceX ESR = '+x_esr)
        if x_esr != self.sourceX.com['NoError']: string = 'Source X: '+x_esr
        eval(self.sourceS.query_error())
        s_esr = str(eval(self.sourceS.read_instrument()))
        self.PrintSave('sourceS ESR = '+s_esr)
        if s_esr != self.sourceS.com['NoError']: string = string +' source S: '+s_esr
        eval(self.voltmeter.query_error())
        m_esr = str(eval(self.voltmeter.read_instrument()))
        self.PrintSave('meter ESR = '+m_esr)
        if m_esr != self.voltmeter.com['NoError']: string = string +' meter: ' +m_esr
        self.PrintSave('')
        return string
        
    def run(self):
        """
        Main thread, reads through the table and executes commands.
        """
        state = 'clear'
        #check integer/float values are correct type ()
        if self.OverideSafety == False:
            state = self.SafetyCheck()

        if state != 'clear': self._want_abort = 1
        
        if self._want_abort:
            self.PrintSave('safety checks failed, making safe, aborting')
            if self.MadeSafe == False:
                self.MakeSafe()
            #wx.PostEvent(self._notify_window, stuff.ResultEvent(self.EVT, "Thread halted at safety checks"))
            return
        
        ####################INITIALISE INSTRUMENTS##########################
        
        eval(self.voltmeter.reset_instrument()) #reset voltmeter
        eval(self.voltmeter.initialise_instrument()) #initialise the voltmeter for reading
        eval(self.sourceX.reset_instrument())
        eval(self.sourceX.initialise_instrument())
        eval(self.sourceS.reset_instrument())
        eval(self.sourceS.initialise_instrument())
        #error string for sourceS?
        
        self.PrintSave('')
        eval(self.sourceX.query_error())
        self.PrintSave('sourceX ESR = '+str(eval(self.sourceX.read_instrument())))
        eval(self.sourceS.query_error())
        self.PrintSave('sourceS ESR = '+str(eval(self.sourceS.read_instrument())))
        eval(self.voltmeter.query_error())
        self.PrintSave('meter ESR = '+str(eval(self.voltmeter.read_instrument())))
        self.PrintSave('')

            ##read all instruments status and print
            
        ######################END of INITIALISATION##########################
        
        ######################MEASUREMENT PROCESS############################
        
        for row in range(self.start_row,self.stop_row + 1):
            self.PrintSave("Spread sheet row "+str(int(row)+1))
            wx.CallAfter(self.grid.SelectRow,row)
            wx.CallAfter(self.grid.ForceRefresh)
            wx.CallAfter(self.grid.Update)

            #read ranges of instruments. If current range is not equal to the previous, apply range setting.
            sourceS_range = self.grid.GetCellValue(row, self.sourceS_range_col)
            if sourceS_range != self.grid.GetCellValue(row-1, self.sourceS_range_col): eval(self.sourceS.set_DCrange(sourceS_range))
            sourceS_value = self.grid.GetCellValue(row, self.sourceS_setting_col)
            eval(self.sourceS.set_DCvalue(sourceS_value))
            
            sourceX_range = self.grid.GetCellValue(row, self.sourceX_range_col)
            if sourceX_range != self.grid.GetCellValue(row-1, self.sourceX_range_col): eval(self.sourceX.set_DCrange(sourceX_range))
            sourceX_value = self.grid.GetCellValue(row, self.sourceX_setting_col)
            eval(self.sourceX.set_DCvalue(sourceX_value))

            eval(self.sourceS.Operate())
            eval(self.sourceX.Operate())
            self.PrintSave('')

            #read errors in meters. The error string later has more warnings appended to it. 
            error_strings = self.Error_string_maker()
            
            #measurement settings, and DVM range setting
            nordgs = int(float(self.grid.GetCellValue(row, self.dvm_nordgs_col) ))
            dvm_range = self.grid.GetCellValue(row, self.dvm_range_col)
            #need to determine maximum of DVM range to find accuracy threshold
            try:
                range_index = [some_range[2] for some_range in self.voltmeter.range].index(dvm_range)
                range_max = self.voltmeter.range[range_index][1]
            except:
                self.PrintSave("could not identify "+str(dvm_range)+" as a valid range")
                self._want_abort = 1
                
            delay_time = int(float(self.grid.GetCellValue(row,self.delay_col))) #delay for settle time
            delay_time2 = float(self.grid.GetCellValue(row,self.delay_col+1)) #delay for DVM between mesmnts
            nominal_reading  = float(self.grid.GetCellValue(row,self.dvm_range_col+1)) #nominal reading for comparison to actual reading
            
            if dvm_range != self.grid.GetCellValue(row-1, self.dvm_range_col): eval(self.voltmeter.set_DCrange(dvm_range)) #fixes range as specified in sheet
            eval(self.voltmeter.MeasureSetup())
            
            #wait desired time for instruments to settle
            #a simple sleep(time) was not used as then the thread cannot be aborted while waiting.
            t=time.time()
            while time.time()<t+delay_time:
                if self._want_abort: #has the stop button been pushed?
                    delay_time = 0
                    
            these_readings = [] #array for readings from a single table row.
            before_msmnt_time =time.time() #start time of measuremtns. 
            for i in range(nordgs):

                t=time.time()
                while time.time()<t+delay_time2: #delay_time2 is the time to wait between measuremnts, if there is such a delay necessary.
                    if self._want_abort: #has the stop button been pushed?
                        delay_time2 = 0
                #the pause is placed before the reading (retreaving of value), as first the DVM must integrate (or do wahtever it does) then you can read
                
                if self._want_abort: #has the stop button been pushed?
                    #wx.PostEvent(self._notify_window, stuff.ResultEvent(self.EVT, "halting thread mid measurements"))
                    #eval(self.sourceS.Standby())
                    #eval(self.sourceX.Standby())
                    i=nordgs #make loop exit
                    if self.MadeSafe == False:
                        self.MakeSafe()
                    self.logfile.close()
                    self.wb.save(filename = str(self.raw_file_name+'.xlsx'))
                    #save the files, otherwise the 'return' will skip the save at the end and data is lost.
                    return #is it best to return, or to set row=self.stop_row+1 in order to exit the for loop.
                
               
                if self.voltmeter.com["SingleMsmntSetup"] != 'SKIP': #only send the extra command if the voltmeter has somethings to be sent
                    eval(self.voltmeter.SingleMsmntSetup())
                try:
                    a = float(eval(self.voltmeter.read_instrument())) #meter class provides string for evaluation
                    if nominal_reading != 0 and abs((a-nominal_reading)) > abs(0.05*range_max): #taking abs to allow for negative ranging.
                        #readings is more than 5% away from nominal, abort
                        error_strings = error_strings+' reading '+str(i)+' is more than 50% from nominal.'
                        self.PrintSave("Readings exceeded 5% from nominal, aborting. Check settle times, ranges, and connections")
                        self._want_abort = 1
                    elif nominal_reading !=0 and abs((a-nominal_reading)) > abs(0.001*range_max):
                        #readings is more than 0.1% away from nominal, issue a warning
                        error_strings = error_strings+' reading '+str(i)+' is more than 0.1% from nominal.'
                        self.PrintSave("Readings exceeded 0.1% from nominal. Check settle times, ranges, and connections")
                    
                    time.sleep(self.voltmeter.measure_seperation) #wait if voltmeter needs to not be read continuously
                    self.PrintSave(time.strftime("%Y.%m.%d.%H.%M.%S, ",time.localtime()) + repr(a) + '\n')
                    self.data.add_to_list(a)
                    these_readings.append(a)
                       
                except:
                    a=0 #so that functions before can do something if there are no previous values.
                    #this does add non-data information to the data, but only to the last value and only in a faulty set.
                    self.PrintSave("Failed to read instrument, aborting. Check: pause time, initialisation and connection")
                    self._want_abort = 1
                
            data_mean = np.mean(these_readings)
            if len(these_readings)>1:
                data_stdev = np.std(these_readings)
            else:
                data_stdev = '0' #avoid infinity for stdev

            #create array to send to csv data sheet
            after_msmnt_time = time.time()
            Time = time.localtime() #time at end of msmnt
            row_time = str(Time[0])+str(Time[1])+str(Time[2])+str(Time[3])+str(Time[4])
 
            #print results
            self.grid.SetCellValue(row, self.dvm_nordgs_col + 6,repr(data_mean))
            self.grid.SetCellValue(row, self.dvm_nordgs_col + 7,repr(data_stdev))
            
            eval(self.voltmeter.inst_status())
            self.grid.SetCellValue(row, self.dvm_nordgs_col + 3,str(eval(self.voltmeter.read_instrument()))) #force string of value as it needs to go into the grid and grid only takes in strings.
            
            eval(self.sourceS.inst_status())
            self.grid.SetCellValue(row, self.dvm_nordgs_col + 4,str(eval(self.sourceS.read_instrument())))
            
            eval(self.sourceX.inst_status())
            self.grid.SetCellValue(row, self.dvm_nordgs_col + 5,str(eval(self.sourceX.read_instrument())))

            #if the error string is not empty, print it to the table on the far left as a warning flag.
            if error_strings != " ":
                self.grid.SetCellValue(row, 0,str(error_strings))
                
            #put csv reading at end so all values are on table.
            csv_line = [self.grid.GetCellValue(row, i) for i in range(self.grid.GetNumberCols())]+[before_msmnt_time,after_msmnt_time]+these_readings
            self.sh.append(csv_line)

            #time.sleep(1.0) #to make sure graph finishes

        ####################### END of MEASUREMENT PROCESS #####################
     
        eval(self.sourceS.Standby())
        eval(self.sourceX.Standby())
        self.logfile.close()
        self.wb.save(filename = str(self.raw_file_name+'.xlsx'))
        wx.CallAfter(self.Analysis_file_name.SetValue,self.raw_file_name+'.xlsx')
        wx.PostEvent(self._notify_window, stuff.ResultEvent(self.EVT, 'GPIB ended'))
        
    ###################################################
    #######        SAFETY CHECK PROCEDURES      #######
    ###################################################
    def initialchecks(self,instrument):
        """
        Short function used in the safety check routine, given an instrument it resets
        the instrument, does initialisation and reads errors. Will continuously read
        errors until the meter returns the no-error string. This can cause problems if
        the no-error string is inputed wrongly.
        """
        eval(instrument.reset_instrument()) #reset voltmeter
        eval(instrument.initialise_instrument())
        eval(instrument.query_error())
        error = eval(instrument.read_instrument())
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
            eval(instrument.query_error())
            error = eval(instrument.read_instrument())
            if error != instrument.com['NoError']:
                while error != instrument.com['NoError'] and not self._want_abort: #loop as there could be several errors
                    self.PrintSave('Error '+str(error)+' in instrument '+instrument.com['label'])
                    eval(instrument.query_error())
                    error = eval(instrument.read_instrument())
                return 'failed'
        return 'clear'

            
    def SafetyCheck(self):
        """
        A linear function for the determination of some basic safety settings. It can be overriden.
        Will check: errors, resetting, outputting zero volts, ouputting 1 volt. If this is sucessful it will also
        determine if the wires are connected the wrong way.
        """
        #determine how long to wait for an output of 1V to settle. do this
        #by reading ranges that include 1V, and finding maximum delay

        
        SettleTime = 5
        
        state = 'clear'
        state = self.initialchecks(self.voltmeter)
        if state == 'clear': state = self.initialchecks(self.sourceX)
        if state == 'clear': state = self.initialchecks(self.sourceS)
        if state == 'clear':
            eval(self.sourceX.Standby())
            eval(self.sourceS.Standby())
            state = self.CheckInstruments(self.sourceX,self.sourceS)
        if state == 'clear':
            eval(self.sourceX.set_DCvalue(0))
            eval(self.sourceS.set_DCvalue(0))
            state = self.CheckInstruments(self.sourceX,self.sourceS)
        if state == 'clear':
            self.PrintSave('testing voltage is setting correctly, and zeros')
            
            eval(self.sourceX.Operate())
            eval(self.sourceS.Operate())
            if self.voltmeter.com["SingleMsmntSetup"] != 'SKIP': #only send the extra command if the voltmeter has somethings to be sent
                eval(self.voltmeter.SingleMsmntSetup())
            time.sleep(SettleTime) #let instrumetns settle
            reading = float(eval(self.voltmeter.read_instrument()))
            if reading>5e-3:
                self.PrintSave('Expect reading to be near zero (less than 5e-3), but it is '+str(reading)+' check zeros')
                state = 'failed'
                self.PrintSave(state)
            state = self.CheckInstruments(self.sourceX,self.sourceS)
        if state == 'clear':
            self.PrintSave('testing voltage setting correctly')
            eval(self.sourceX.Standby())
            #find correct range for 1volt
            for r in self.sourceX.range[::-1]:
                if r[1]>=1:
                    r_use = r
            eval(self.sourceX.set_DCrange(r_use[2]))
            eval(self.sourceX.set_DCvalue(1))
            eval(self.sourceX.Operate())
            if self.voltmeter.com["SingleMsmntSetup"] != 'SKIP': #only send the extra command if the voltmeter has somethings to be sent
                eval(self.voltmeter.SingleMsmntSetup())
            time.sleep(SettleTime) #let instrumetns settle
            reading = float(eval(self.voltmeter.read_instrument()))
            if round(reading,2)!= 1.00:
                self.PrintSave('Source X set to 1, expect reading to be near 1, but it is '+str(reading)+' check output load on X and SettleTime')
                if round(reading,2) == -1:
                    self.PrintSave('reading negative, check if wires are the right way')
                state = 'failed'
                return state
            state = self.CheckInstruments(self.sourceX,self.sourceS)
        if state == 'clear':
            eval(self.sourceX.Standby())
            eval(self.sourceX.set_DCvalue(0))
            eval(self.sourceX.Operate())
            eval(self.sourceS.Standby())
            for r in self.sourceS.range[::-1]:
                if r[1]>=1:
                    r_use = r
            eval(self.sourceS.set_DCrange(r_use[2]))
            eval(self.sourceS.set_DCvalue(1))
            eval(self.sourceS.Operate())
            if self.voltmeter.com["SingleMsmntSetup"] != 'SKIP': #only send the extra command if the voltmeter has somethings to be sent
                eval(self.voltmeter.SingleMsmntSetup())
            time.sleep(SettleTime) #let instrumetns settle
            reading = float(eval(self.voltmeter.read_instrument()))
            if round(reading,2)!= -1.00:
                self.PrintSave('Set source S to 1, expect reading to be near -1, but it is '+str(reading)+' output load on S and SettleTime')
                state = 'failed'
                return state
            state = self.CheckInstruments(self.sourceX,self.sourceS)
        return state
        
