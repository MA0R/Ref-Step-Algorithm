"""
A general instrument class that returns a status for each command sent
or recieved from its instrument. This allows it to be used with the "com"
function in the main algorithm, when visa fails it does not halt the
whole program but reports the failure instead.
"""
import visa
import time
class INSTRUMENT(object):
    """ This instrument class really is only a function to read and write to
    some instrument, using pyvisa. It has a general 'dictionary' to which
    more key word arguments can be added, and more sub functions can be
    used to make instruments specific. It can be used very generally, as just a semd
    and recieve class which also wraps each communication with a check to see if the
    communication was sucesful. """
    def __init__(self,inst_bus,letter, **kwargs):
        self.com = {'label':'','adress':'', 'Ranges':[], 'measure_seperation':'0', 'NoError':'',\
                    'reset':'','status':'','init':'','Make_Safe':'', 'error':'', \
                    'SettleTime':'0', 'DCVRange':'', 'SetVoltage':'', 'operate':'', \
                    'standby':'','MeasureSetup':'','SingleMsmntSetup':''} #command dictionary
        self.com.update(kwargs) #update dictionary to include all sent commands.
        self.label = self.com["label"]
        self.com.update(label=str(letter)+str(kwargs['label']) )
        self.range = eval(self.com['Ranges']) #Use eval here or string operations? Like split multiple times.
        self.adress = self.com['adress']
        #ensure values are ints
        try:
            self.com_settle_time = float(self.com['SettleTime'])
        except:
            print("settle time made into 1 on "+str(self.com['label']+", from unreadable: "+str(self.com['SettleTime'])))
            self.com_settle_time = 1
        try:
            self.measure_seperation = float(self.com['measure_seperation'])
        except:
            print("measure seperation made into 0 on "+str(self.com['label']+", from unreadable: "+str(self.com['measure_seperation'])))
            self.measure_seperation = 0
            
        self.inst_bus = inst_bus #save the instrument bus, either visa or the simulated visa

    def create_instrument(self):
        """
        Needs to be called prior to any commands being sent or recieved.
        Creates the visa instrument object, to which commands will be sent
        and recieved. 
        """
        sucess = False
        string = str(time.strftime("%Y.%m.%d.%H.%M.%S, ", time.localtime()))+' Creating '+self.label+': '
        try:
            self.rm = self.inst_bus.ResourceManager()
            self.inst = self.rm.open_resource(self.adress)
            string = string+"sucess"
            sucess = True
        except: #There are a number of issues visa might raise?
            string = string+"visa failed at adress "+str(self.adress)
        return [sucess,None,string]
    
    def send(self,command):
        """
        From here a command is sent to the instrument, surrounded by the try block.
        If the command fails, it does not halt the problem but sends back a failed status.
        """
        sucess = False #did we read sucessfully
        #string to be printed and saved in log file
        string = str(time.strftime("%Y.%m.%d.%H.%M.%S, ", time.localtime()))+' writing to '+self.label+': ' 
        try:
            self.inst.write(command)
            time.sleep(self.com_settle_time)
            string = string+str(command)
            sucess = True
        except self.inst_bus.VisaIOError:
            string = string+"visa failed"
        return [sucess,None,string]
    
    def read_instrument(self):
        """
        Similar to the send function, but reads and expects a return value too.
        """
        val = '0' #value to be returned, string-type like instruments
        sucess = False #did we read sucessfully
        #string to be printed and saved in log file
        string = str(time.strftime("%Y.%m.%d.%H.%M.%S, ", time.localtime()))+' reading '+self.label+': ' 
        try:
            time.sleep(self.measure_seperation)
            val = self.inst.read()
            string = string+str(val)
            sucess = True
        except self.inst_bus.VisaIOError:
            string = string+"visa failed"
        return [sucess,val,string]

    def initialise_instrument(self):
        """A specific instrument command to the ref-step algorithm,
initialises instruments with a set of commands"""
        return self.send(self.com['init'])

    def make_safe(self):
        """specific to the ref-step algorithm, should turn instruments off"""
        return self.send(self.com['Make_Safe'])
    
    def inst_status(self):
        """specific to the ref-step algorithm, used for reading status"""
        return self.send(self.com['status'])

    def reset_instrument(self):
        """specific to the ref-step algorithm, reset routine"""
        return self.send(self.com['reset'])
        
    def set_DCrange(self, value):
        """specific to the ref-step algorithm, setting a DC voltage"""
        line = str(self.com['DCVRange'])
        line = line.replace("$",str(value))
        return self.send(line)
    
    def query_error(self):
        """specific to the ref-step algorithm, reading the instruments error"""
        return self.send(self.com['error'])
            
    def set_DCvalue(self, value):
        """specific to the ref-step algorithm, set a DC value for sources"""
        line = str(self.com['SetVoltage'])
        line = line.replace("$",str(value))
        return self.send(line)

    def Operate(self):
        """specific to the ref-step algorithm, operates sources"""
        return self.send(self.com['operate'])
    
    def Standby(self):
        """specific to the ref-step algorithm, puts sources on standby"""
        return self.send(self.com['standby'])

    def MeasureSetup(self):
        """specific to the ref-step algorithm, pre measurement sequence set up"""
        return self.send(self.com['MeasureSetup'])

    def SingleMsmntSetup(self):
        """specific to the ref-step algorithm, should any commands be sent prior to an individual measurement"""
        return self.send(self.com['SingleMsmntSetup'])

