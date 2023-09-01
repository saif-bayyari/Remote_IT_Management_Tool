import PySimpleGUI as sg
from time import sleep
from pypsexec.client import Client
import subprocess


#^^^^^LIBRARIES USED

#pypsexec library defines a client for you to remotely connect to, 
# I decided it would be best to just connect to the local admin account because those are already linked with LAPS

#the c variable represents the current remote connection session. 
#obviously when you first run the program, there is no session so c is set to None (Null)
#c needs to be a global variable so we can reference it when we handle disconnecting from the session
c = None

def connectCommand(pcName,command,win):
    global c
    if c == None: #if there is no current remote session going on, start a new session and run the inputted command
        
        try:
        
            #the output variable runs a powershell command that retrieves the admin LAPs password from the inputted PC name.
            #we then take that powershell output and remove all unnecessary white space and headers
            #the final result is just the LAPS password and nothing else, we set that to LAPsPassword variable
            output = subprocess.Popen(['powershell.exe', "Get-ADComputer -Identity '"+pcName+"' -Properties ms-Mcs-AdmPwd | select-object -expandproperty ms-Mcs-AdmPwd"], shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            LAPsPassword = output.stdout.read().decode().strip()
            "".join(LAPsPassword.split())

            #We then define a connection to the specified PC and select the local admin as the user we log into when connecting
            c = Client(pcName, username=".\Administrator", password=LAPsPassword)
            c.connect()
            c.create_service()

        except:
            print("invalid logon")
            c = None
            return

    #if the connection is successful, we then make an attempt to process a command during that session
    try:
        stdout, stderr, rc = c.run_executable("cmd.exe", arguments="/c "+command)
        if stdout:
            win['-OUTPUT-'].update(stdout)
        else:
            win['-OUTPUT-'].update(stderr)

    except:
        print("command failed")

   
#when the "disconnect from session" button is pressed, the session is terminated and that global variable is set back to None
#you must also remove the service when you're disconnecting
def disconnectFromSession(win):
    global c
    if c != None:
        c.remove_service()
        c.disconnect()
        c = None
    

    
        

    






#Next couple lines of code are the GUI components such as buttons, input fields, labels, etc.

remoteCommandLabel = [sg.Text("Remote Powershell Command Tool:")]
connectToPC =  [sg.Text("PC Name:"),sg.Input(key="-IN-" ,change_submits=True, size=(20, 10))]
consoleOutput = [sg.Output(size=(50,10), key='-OUTPUT-'), sg.Button("Run Remote Command"), sg.Button("Disconnect Current Session")]
userCreationDepartureLabel = [sg.Text("Automated User Creation and Departure:")]
userCreationDeparture1 = [sg.Text("First Name:"), sg.Input(key="-IN4-"), sg.Text("Last Name:"), sg.Input(key="-IN5-")]
userCreationDeparture2 = [sg.Text("Title:"), sg.Input(key="-IN6-"), sg.Text("Department:"), sg.Input(key="-IN7-")]
userCreationDeparture3 = [sg.Text("Manager:"), sg.Input(key="-IN8-"), sg.Text("Status:"), sg.Combo(["Full-Time Employee", "Contractor"], font=('Arial Bold', 12),  expand_x=True, enable_events=True,  readonly=True, key='-COMBO-')]


#the template commands you can specify for the command combo box
commands = ["whoami", "ipconfig", "ipconfig /flushdns", "netsh winsock reset", 
            "ipconfig /renew", "gpupdate /force", "shutdown -r"]

commandComboBox = [sg.Combo(commands, font=('Arial Bold', 12),  expand_x=True, enable_events=True,  readonly=False, key='-COMBO2-')]



#defines the color theme of the GUI and puts all the components all together into a single window layout
sg.theme("DarkTeal2")
layout = [remoteCommandLabel, connectToPC, commandComboBox,
          consoleOutput, userCreationDepartureLabel, userCreationDeparture1, userCreationDeparture2, userCreationDeparture3]

#Building Window
window = sg.Window('Remote IT Management Tool, By Saif Bayyari', layout, size=(800,600))



#BELOW IS THE LOOP TO ACTUALLY RUN THE WINDOW AND LISTEN FOR BUTTON CLICKS
    
while True:


    event, values = window.read()
    if event == sg.WIN_CLOSED or event=="Exit":
        break
    elif event == "Run Remote Command":
        window["-OUTPUT-"].Update('')
        connectCommand(values["-IN-"],values["-COMBO2-"], window)
    elif event == "Disconnect Current Session":
        window["-OUTPUT-"].Update('')
        disconnectFromSession(window)
    

    

         
        
