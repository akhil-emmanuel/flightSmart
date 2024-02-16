import openpyxl
import PySimpleGUI as sg
from openpyxl import Workbook, load_workbook
import webbrowser
import math
from datetime import timedelta
from datetime import datetime
import random
import pytz

book = load_workbook('data/IFAircraftData.xlsx')
Checklist = load_workbook('data/Checklist.xlsx')


def round_nrst_10(number):
    return round(number / 10) * 10

def InputLoadtoDataRow(acLoad):
    acLoad = int(acLoad)
    if acLoad == 23 or acLoad == 24 or acLoad == 25 or acLoad == 26 or acLoad == 27:
        row = 10
    elif acLoad == 73 or acLoad == 74 or acLoad == 75 or acLoad == 76 or acLoad == 77:
        row = 4
    else:
        acLoad = round_nrst_10(acLoad)
        if acLoad == 0 or acLoad == 10:
            row = 11
        elif acLoad == 20:
            row = 10
        elif acLoad == 30:
            row = 9
        elif acLoad == 40:
            row = 8
        elif acLoad == 50:
            row = 7
        elif acLoad == 60:
            row = 6
        elif acLoad == 70:
            row = 5
        elif acLoad == 80:
            row = 4
        elif acLoad == 90:
            row = 3
        elif acLoad == 100:
            row = 2

    return str(row)

def getDepatureData(row):
    DepPower = sheet['B' + row].value
    DepFlaps = sheet['C' + row].value
    DepRotate = sheet['D' + row].value
    DepAirBy = sheet['E' + row].value
    return DepPower, DepFlaps, DepRotate, DepAirBy

def getArrivalData(row):
    LdgFlaps = sheet['F'+ str(row)].value
    LdgApprSpd  = sheet['G'+ str(row)].value
    LdgFlareSpd = sheet['H'+ str(row)].value
    FlapsSpd = sheet['A14'].value
    return LdgFlaps, LdgApprSpd, LdgFlareSpd, FlapsSpd

def getFuelBurnData(row):
    Even = sheet['I'+ str(row)].value
    Odd = sheet['J'+ str(row)].value
    MedBurn = sheet['K'+ str(row)].value
    RecWest = sheet['E14'].value
    RecEast = sheet['F14'].value
    return Even, Odd, MedBurn, RecWest, RecEast

def getOtherData(row):
    Ceiling = sheet['B14'].value
    Cruise = sheet['C14'].value
    MMO = sheet['D14'].value
    Range = sheet['A13'].value
    return Ceiling, Cruise, MMO, Range

def CL350(oat):
    if oat == -30:
        return "80% = 81.4% N1"
    elif oat == -25:
        return "81% = 82.3% N1"
    elif oat == -20:
        return "82% = 83.2% N1"
    elif oat == -15:
        return "83% = 84.0% N1"
    elif oat == -10:
        return "84% = 84.7% N1"
    elif oat == -5:
        return "85% = 85.5% N1"
    elif oat == 0:
        return "87% = 86.3% N1"
    elif oat == 5:
        return "87% = 87.0% N1"
    elif oat == 10:
        return "88% = 87.8% N1"
    elif oat == 15:
        return "90% = 88.7% N1"
    elif oat == 20:
        return "91% = 89.5% N1"
    elif oat == 25:
        return "92% = 90.3% N1"
    elif oat == 30 or oat == 35:
        return "93% = 91.1% N1"
    elif oat == 40:
        return "90% = 88.7% N1"
    else:
        return "Invalid input"

def getChecklistItem(page, itemNumber):
    sheetName = Checklist.worksheets[page]
    output = [sheetName['A' + str(itemNumber)].value, sheetName['B' + str(itemNumber)].value, sheetName['C' + str(itemNumber)].value, sheetName['D' + str(itemNumber)].value]
    if output[0] is None:
        output[0] = ' '
    if output[1] is None:
        output[1] = ' '
    return (output)

def updateChecklistItem(page, itemNumber, newComplete):
    sheetName = Checklist.worksheets[page]
    sheetName['C' + str(itemNumber)] = newComplete

def collapse(layout, key):
    return sg.pin(sg.Column(layout, key=key))

def calculateWindCompents(rwNumber, wdDirection, wdSpd):
        rwDirection = rwNumber * 10
        runway_dir_rad = math.radians(rwDirection)
        wind_dir_rad = math.radians(wdDirection)
        angle_diff_rad = wind_dir_rad - runway_dir_rad
        crosswind_component = wdSpd * math.sin(angle_diff_rad)
        headwind_component = wdSpd * math.cos(angle_diff_rad)
        return crosswind_component, headwind_component


def calculateDescentSpeed(min, sec, initial, final):
    totalTime = min + (sec / 60)
    altDif = initial - final
    vspd = altDif/totalTime
    return vspd

def calculate_time(hours1, minutes1, operation, hours2, minutes2):
    time1 = timedelta(hours=hours1, minutes=minutes1)
    time2 = timedelta(hours=hours2, minutes=minutes2)

    if operation == "addition":
        result = time1 + time2
    elif operation == "subtraction":
        result = time1 - time2
    else:
        return "Invalid operation. Please choose 'addition' or 'subtraction'."

    result_hours = result.days * 24 + result.seconds // 3600
    result_minutes = (result.seconds % 3600) // 60

    return result_hours, result_minutes


def get_current_time(time_format):
    current_time = datetime.now()
    time_strings = []

    time_zones = ['Australia/Sydney', 'America/New_York', 'Asia/Dubai', 'Asia/Tokyo', 'UTC']
    for tz_name in time_zones:
        tz = pytz.timezone(tz_name)
        time_str = current_time.astimezone(tz).strftime(time_format)
        time_strings.append(f"{tz_name.split('/')[-1]}: {time_str}")


    return time_strings


dark_mode = {
    'BACKGROUND': '#1C1C1D', #color1
    'TEXT': '#FFFFFF', #color3
    'INPUT': '#404040', #color2
    'TEXT_INPUT': '#FFFFFF',
    'SCROLL': '#404040',
    'BUTTON': ('#FFFFFF', '#404040'),
    'PROGRESS': ('#FFFFFF', '#404040'),
    'BORDER': 1,
    'SLIDER_DEPTH': 0,
    'PROGRESS_DEPTH': 0,
}
sg.theme_add_new('darkMode', dark_mode)





light_mode = {
    'BACKGROUND': '#F9F9F9', #color1
    'TEXT': '#5c5c5c', #color2
    'INPUT': '#E0E1E2', #color3
    'TEXT_INPUT': '#5c5c5c',
    'SCROLL': '#E0E1E2',
    'BUTTON': ('#5c5c5c', '#E0E1E2'),
    'PROGRESS': ('#5c5c5c', '#E0E1E2'),
    'BORDER': 1,
    'SLIDER_DEPTH': 0,
    'PROGRESS_DEPTH': 0,
}
sg.theme_add_new('lightMode', light_mode)

mode = 'dark'

if mode == 'light':
    color1 = '#F9F9F9'
    color2 = '#E0E1E2'
    color3 = '#5c5c5c'
    filename = 'img/flightSmartTitleAppVerisonLight.png'
    sg.theme('lightMode')
elif mode =='dark':
    color1 = '#1C1C1D'
    color2 = '#404040'
    color3 = '#FFFFFF'
    filename = 'img/flightSmartTitleAppVerison.png'
    sg.theme('darkMode')




title_row = [
        [sg.Image(filename=filename ,size=(940, 50))],
        [sg.HorizontalSeparator(color=color1)]
]



sectionNotepad = [
            [sg.Multiline(size=(60, 60), key='notepadData', no_scrollbar=False, pad=(0,0))]
        ]

sectionXwind = [
            [sg.Text("Runaway Number:", pad=(0,2)), sg.Input('', key='-rwDirInput-', pad=(0,2))],
            [sg.Text("Wind Direction:", pad=(0,2)), sg.Input('', key='-wdDirInput-',  pad=(0,2))],
            [sg.Text("Wind Speed:", pad=(0,2)), sg.Input('', key='-wdSpdInput-',  pad=(0,2))],
            [sg.Button('Calculate Wind Components', auto_size_button=False, size=(50,1),key='-XwindGo-', pad=(0,3))],
            [sg.Text(key="-XwindOutputLine1-",pad=(0,1))],
            [sg.Text(key="-XwindOutputLine2-",pad=(0,1))],
        ]

sectionDescent = [
            [sg.Text("ETE:    ", pad=(0,1)), sg.Input('', key='-minInput-', size=(2,1), pad=(0,1)), sg.Text("mins :", pad=(0,1)), sg.Input('', key='-secInput-', size=(2,1), pad=(0,1)), sg.Text("secs", pad=(0,1),)],
            [sg.Text("Initial Altitude:", pad=(0,1)), sg.Input('', key='-initialInput-', pad=(0,1))],
            [sg.Text("Desired Altitude", pad=(0,1)), sg.Input('', key='-finalInput-', pad=(0,1))],
            [sg.Button('Calculate Vertical Speed', auto_size_button=False, size=(50,1),key='-DescentGo-', pad=(0,3))],
            [sg.Text(key="-DescentOutput-",pad=(0,1))],
        ]

sectionTime = [
            [sg.Text("Hour: ", pad=(0,1)), sg.Input('', key='-hr1Input-', size=(4,1), pad=(0,1)), sg.Text("       Min:", pad=(0,1)), sg.Input('', key='-min1Input-', size=(4,1), pad=(0,1))],
            [sg.Text("Hour: ", pad=(0,1)), sg.Input('', key='-hr2Input-', size=(4,1), pad=(0,1)), sg.Text("       Min:", pad=(0,1)), sg.Input('', key='-min2Input-', size=(4,1), pad=(0,1))],
            [sg.Button('Addition', auto_size_button=False, size=(17,1),key='-TimeAdd-', pad=(0,3)), sg.Button('Subtraction', auto_size_button=False, size=(17,1),key='-TimeSub-', pad=(0,3))],
            [sg.Text(key="-TimeOutput-",pad=(0,1))],


        ]

sectionPaxSim = [
            [sg.Button('Widebody', auto_size_button=False, size=(17,1),key='-PSwide-', pad=(0,1)), sg.Button('Narrowbody', auto_size_button=False, size=(17,1),key='-PSnarrow-', pad=(0,1))],
            [sg.Button("Load Passengers", auto_size_button=False, size=(15,1), pad=(0,1)), sg.ProgressBar(100, orientation='h', size=(20, 10), key='-BoardingProgress-')],
            [sg.Button("Load Cargo", auto_size_button=False ,size=(15,1), pad=(0,1)), sg.ProgressBar(100, orientation='h', size=(20, 10), key='-CargoProgress-')],
            [sg.Button("Load Catering", auto_size_button=False, size=(15,1), pad=(0,1)), sg.ProgressBar(100, orientation='h', size=(20, 10), key='-CateringProgress-')],
            [sg.Multiline("MSG LOG:\n",size=(40,20), pad=(0,3), key='PaxSimOutput')]
        ]




col1 = sg.Column([
    [sg.Frame('Input:', layout=[
        [sg.Text('Aircraft Manufacturer:'), sg.Button('Airbus'), sg.Button('Boeing'), sg.Button('Bombardier'),
         sg.Button('Embraer'), sg.Button('McDonnell Douglas')],
        [sg.Text('Aircraft Type:'), sg.DropDown(['      '], key='selected_aircraft')],
        [sg.Text('Aircraft Load:'), sg.InputText(key='load', size=(3, 1)), sg.Text('%'),
        # CL350 ONLY
        sg.Text("Select OAT (°C):", key='askOAT', visible=False), sg.Slider(range=(-30, 40), default_value=0, orientation="h", key="OATinput", size=(20, 20), resolution=5, visible=False)],
        [sg.Text('Request:'), sg.Button('Departure Data'),  sg.Button('Fuel Burn Data'), sg.Button('Arrival Data'), sg.Button('Other Data')]], size=(600, 160) )],

    [sg.Frame('Output:', layout=[
        [sg.Multiline(size=(60, 60), key='output', disabled=True, no_scrollbar=True)],
        ], size=(295, 300)),

    sg.Frame('Other:', layout=[
            [sg.Button('Notes', enable_events=True, button_color=(color2,color3), pad=(3,0), k='-OPEN NOTEPAD-', font=('Helvetica Neue', 14)),
            sg.Button('PaxSim', enable_events=True, pad=(3, 0), k='-OPEN PAXSIM-', font=('Helvetica Neue', 14)),
            sg.Button('Wind', enable_events=True, pad=(3,0), k='-OPEN XWIND-', font=('Helvetica Neue', 14)),
            sg.Button('Descent', enable_events=True, pad=(3,0), k='-OPEN DESCENT-', font=('Helvetica Neue', 14)),
            sg.Button('Time', enable_events=True, pad=(3,0), k='-OPEN TIME-', font=('Helvetica Neue', 14))
            ],
            [sg.HorizontalSeparator(color=color1)],
            [collapse(sectionNotepad, '-Notepad-')],
            [collapse(sectionPaxSim, '-PaxSim-')],
            [collapse(sectionXwind, '-Xwind-')],
            [collapse(sectionDescent, '-Descent-')],
            [collapse(sectionTime, '-Time-')],



        ], size=(295, 300))],
    [sg.Button("Credits"), sg.Button("App Details ↗"), sg.Button("Quit")],
], expand_y=True)



col2 = sg.Column([
        [sg.Frame('Checklist:', layout=[
        [sg.Text(key='checklist-Header', text_color=color3, background_color=color2, size=(65, 1), pad=(0,0), justification='center', enable_events=True, visible=False,)],
        [sg.Text(key='checklist-line1a', text_color=color3, background_color=color1, size=(16, 1), pad=(0,0), enable_events=True), sg.Text(key='checklist-line1b', text_color=color3, background_color=color1, expand_x=True, size=(16, 1), pad=(0,0), justification='right',enable_events=True)],
        [sg.Text(key='checklist-line2a', text_color=color3, background_color=color1, size=(16, 1), pad=(0,0), enable_events=True), sg.Text(key='checklist-line2b', text_color=color3, background_color=color1, expand_x=True, size=(16, 1), pad=(0,0), justification='right',enable_events=True)],
        [sg.Text(key='checklist-line3a', text_color=color3, background_color=color1, size=(16, 1), pad=(0,0), enable_events=True), sg.Text(key='checklist-line3b', text_color=color3, background_color=color1, expand_x=True, size=(16, 1), pad=(0,0), justification='right',enable_events=True)],
        [sg.Text(key='checklist-line4a', text_color=color3, background_color=color1, size=(16, 1), pad=(0,0), enable_events=True), sg.Text(key='checklist-line4b', text_color=color3, background_color=color1, expand_x=True, size=(16, 1), pad=(0,0), justification='right',enable_events=True)],
        [sg.Text(key='checklist-line5a', text_color=color3, background_color=color1, size=(16, 1), pad=(0,0), enable_events=True), sg.Text(key='checklist-line5b', text_color=color3, background_color=color1, expand_x=True, size=(16, 1), pad=(0,0), justification='right',enable_events=True)],
        [sg.Text(key='checklist-line6a', text_color=color3, background_color=color1, size=(16, 1), pad=(0,0), enable_events=True), sg.Text(key='checklist-line6b', text_color=color3, background_color=color1, expand_x=True, size=(16, 1), pad=(0,0), justification='right',enable_events=True)],
        [sg.Text(key='checklist-line7a', text_color=color3, background_color=color1, size=(16, 1), pad=(0,0), enable_events=True), sg.Text(key='checklist-line7b', text_color=color3, background_color=color1, expand_x=True, size=(16, 1), pad=(0,0), justification='right',enable_events=True)],
        [sg.Text(key='checklist-line8a', text_color=color3, background_color=color1, size=(16, 1), pad=(0,0), enable_events=True), sg.Text(key='checklist-line8b', text_color=color3, background_color=color1, expand_x=True, size=(16, 1), pad=(0,0), justification='right',enable_events=True)],
        [sg.Text(key='checklist-line9a', text_color=color3, background_color=color1, size=(16, 1), pad=(0,0), enable_events=True), sg.Text(key='checklist-line9b', text_color=color3, background_color=color1, expand_x=True, size=(16, 1), pad=(0,0), justification='right',enable_events=True)],
        [sg.Text(key='checklist-line10a', text_color=color3, background_color=color1, size=(16, 1), pad=(0,0), enable_events=True), sg.Text(key='checklist-line10b', text_color=color3, background_color=color1, expand_x=True, size=(16, 1), pad=(0,0), justification='right',enable_events=True)],
        [sg.HorizontalSeparator(color=color1)],
        [sg.Button("BACK", visible=False),sg.Button("OK", size=(19,1), auto_size_button=False, visible=False), sg.Button("NEXT", visible=False)],
        [sg.Button("Start", size=(35,1), auto_size_button=False, visible=True)]
        ], size=(295, 300))],
        [sg.Frame('Hyperlinks: ↗', layout=[
        [sg.Button('IFC', size=(15,1), auto_size_button=False,), sg.Text(''), sg.Button('IF Pro Discord', size=(15,1), auto_size_button=False,)],
        [sg.Button('Live Flight', size=(15,1), auto_size_button=False,), sg.Text(''), sg.Button('Map Flight', size=(15,1), auto_size_button=False,)],
        [sg.Button('FlightAware', size=(15,1), auto_size_button=False,), sg.Text(''), sg.Button('Flightradar24', size=(15,1), auto_size_button=False,)],
        [sg.Button('Windy', size=(15,1), auto_size_button=False,), sg.Text(''), sg.Button('FPL to IF', size=(15,1), auto_size_button=False,)],

        ], size=(295, 160))],

], expand_y=True)


layout = [
    title_row,
    [col1, col2],
]



window = sg.Window("flightSmart", layout, font=('Helvetica Neue', 14), size=(955, 600), resizable=True, finalize=True)  # size: width, height

sheetNumber = -1

currentPage = -1
currentItem = 0
itemNumber = 0


window.bind('<Right>', 'NEXT')
window.bind('<Down>', 'OK')
window.bind('<Left>', 'BACK')
window.bind('<Space>', 'OK')

PaxTimer_running = False
PaxTimer_value = 0
CargoTimer_running = False
CargoTimer_value = 0
CateringTimer_running = False
CateringTimer_value = 0

PStype = 'none'

is_24_hour_format = False




while True:
    event, value = window.read(timeout=1000)

    if event == "Quit" or event == sg.WIN_CLOSED:
        break



    if event == '-OPEN NOTEPAD-':
        window['-Notepad-'].update(visible=True)
        window['-Xwind-'].update(visible=False)
        window['-Descent-'].update(visible=False)
        window['-Time-'].update(visible=False)
        window['-PaxSim-'].update(visible=False)

        window['-OPEN NOTEPAD-'].update(button_color=(color2,color3))
        window['-OPEN XWIND-'].update(button_color=(color3, color2))
        window['-OPEN DESCENT-'].update(button_color=(color3, color2))
        window['-OPEN TIME-'].update(button_color=(color3, color2))
        window['-OPEN PAXSIM-'].update(button_color=(color3, color2))

    if event == '-OPEN XWIND-':
        window['-Notepad-'].update(visible=False)
        window['-Xwind-'].update(visible=True)
        window['-Descent-'].update(visible=False)
        window['-Time-'].update(visible=False)
        window['-PaxSim-'].update(visible=False)

        window['-OPEN NOTEPAD-'].update(button_color=(color3, color2))
        window['-OPEN XWIND-'].update(button_color=(color2,color3))
        window['-OPEN DESCENT-'].update(button_color=(color3, color2))
        window['-OPEN TIME-'].update(button_color=(color3, color2))
        window['-OPEN PAXSIM-'].update(button_color=(color3, color2))

    if event == '-OPEN DESCENT-':
        window['-Notepad-'].update(visible=False)
        window['-Xwind-'].update(visible=False)
        window['-Descent-'].update(visible=True)
        window['-Time-'].update(visible=False)
        window['-PaxSim-'].update(visible=False)

        window['-OPEN NOTEPAD-'].update(button_color=(color3, color2))
        window['-OPEN XWIND-'].update(button_color=(color3, color2))
        window['-OPEN DESCENT-'].update(button_color=(color2,color3))
        window['-OPEN TIME-'].update(button_color=(color3, color2))
        window['-OPEN PAXSIM-'].update(button_color=(color3, color2))

    if event == '-OPEN TIME-':
        window['-Notepad-'].update(visible=False)
        window['-Xwind-'].update(visible=False)
        window['-Descent-'].update(visible=False)
        window['-Time-'].update(visible=True)
        window['-PaxSim-'].update(visible=False)

        window['-OPEN NOTEPAD-'].update(button_color=(color3,color2))
        window['-OPEN XWIND-'].update(button_color=(color3, color2))
        window['-OPEN DESCENT-'].update(button_color=(color3, color2))
        window['-OPEN TIME-'].update(button_color=(color2, color3))
        window['-OPEN PAXSIM-'].update(button_color=(color3, color2))

    if event == '-OPEN PAXSIM-':

        window['-Notepad-'].update(visible=False)
        window['-Xwind-'].update(visible=False)
        window['-Descent-'].update(visible=False)
        window['-Time-'].update(visible=False)
        window['-PaxSim-'].update(visible=True)

        window['-OPEN NOTEPAD-'].update(button_color=(color3,color2))
        window['-OPEN XWIND-'].update(button_color=(color3, color2))
        window['-OPEN DESCENT-'].update(button_color=(color3, color2))
        window['-OPEN TIME-'].update(button_color=(color3, color2))
        window['-OPEN PAXSIM-'].update(button_color=(color2, color3))


    if event == ('-XwindGo-'):
        try:
            rwDir = value['-rwDirInput-']
            wdDir = value['-wdDirInput-']
            wdSpd = value['-wdSpdInput-']
            if int(rwDir) > 36:
                window['-XwindOutputLine1-'].update('Runway number must be less than 36')
            elif int(wdDir) <= 0 or int(wdDir) > 360:
                window['-XwindOutputLine1-'].update('Wind direction must be less than 360')
            else:
                leftRightWind, headTailWind = (calculateWindCompents(float(rwDir), float(wdDir), float(wdSpd)))
                if headTailWind >= 0:
                    ouputLine1 = "Headwind Component:  ↓ " + str(round(headTailWind, 2))
                else:
                    ouputLine1 = "Tailwind Component:  ↑ " + str(abs(round(headTailWind, 2)))

                if leftRightWind >= 0:
                    ouputLine2 = "Crosswind Component: ←   " + str(round(leftRightWind, 2))
                else:
                    ouputLine2 = "Crosswind Component: →  " + str(abs(round(leftRightWind, 2)))
                window['-XwindOutputLine1-'].update(ouputLine1)
                window['-XwindOutputLine2-'].update(ouputLine2)
        except:
            window['-XwindOutputLine1-'].update('Error, try again')

    if event == ('-DescentGo-'):
        try:
            min = value['-minInput-']
            sec = value['-secInput-']
            initial = value['-initialInput-']
            final = value['-finalInput-']
            output = calculateDescentSpeed(int(min), int(sec), int(initial), int(final))
            output = round(output, 0)
            output = "Descend at -" + str(output) + " fpm"
            window['-DescentOutput-'].update(output)
        except:
            window['-DescentOutput-'].update('Error, try again')

    if event == ('-TimeAdd-'):
        try:
            hr1 = value['-hr1Input-']
            min1 = value['-min1Input-']
            hr2 = value['-hr2Input-']
            min2 = value['-min2Input-']
            outputHrs, outputMins = calculate_time(int(hr1), int(min1), "addition", int(hr2), int(min2))
            output = str(outputHrs) + " hours and " + str(outputMins) + " minutes"
            window['-TimeOutput-'].update(output)
        except:
            window['-TimeOutput-'].update('Error, try again')
    if event == ('-TimeSub-'):
        try:
            hr1 = value['-hr1Input-']
            min1 = value['-min1Input-']
            hr2 = value['-hr2Input-']
            min2 = value['-min2Input-']
            outputHrs, outputMins = calculate_time(int(hr1), int(min1), "subtraction", int(hr2), int(min2))
            output = str(outputHrs) + " hours and " + str(outputMins) + " minutes"
            window['-TimeOutput-'].update(output)
        except:
            window['-TimeOutput-'].update('Error, try again')


    if event == "-PSwide-":
        PStype = 'wide'
        window['-PSwide-'].update(button_color=(color2,color3))
        window['-PSnarrow-'].update(button_color=(color3, color2))

    if event == "-PSnarrow-":
        PStype = 'narrow'
        window['-PSwide-'].update(button_color=(color3, color2))
        window['-PSnarrow-'].update(button_color=(color2, color3))
    if event == 'Load Passengers' and PStype != 'none':
        if not PaxTimer_running:
            if PStype == 'wide':
                PaxTimer_value = random.randint(300, 420) #/ 30
            elif PStype == 'narrow':
                PaxTimer_value = random.randint(240, 300) #/ 30
            PaxTimer_start = PaxTimer_value
            PaxTimer_running = True
            window['-BoardingProgress-'].update_bar(100)
            output1 = "FLIGHT ATTENDANTS: Cpt, we have starting boarding, expected duration " + str(round(PaxTimer_start / 60)) + " minutes."
            output2 = "FLIGHT ATTENDANTS: Cpt, we're about to begin boarding, and it should take approximately " + str(round(PaxTimer_start / 60)) + " minutes."
            output3 = "FLIGHT ATTENDANTS: Cpt, we have just connected the jetbridge. Boarding should be done in about " + str(round(PaxTimer_start / 60)) + " minutes."
            window['PaxSimOutput'].print(random.choice([output1, output2, output3]),text_color="cyan")
    if event == 'Load Passengers' and PStype == 'none':
        window['PaxSimOutput'].print('Select Widebody or Narrowbody!', text_color="red")
    if PaxTimer_running and PStype != 'none':
        PaxTimer_value -= 1
        PaxProgress = (PaxTimer_value / PaxTimer_start) * 100
        window['-BoardingProgress-'].update_bar(PaxProgress)
        if PaxTimer_value <= 0:
            output1 = "FLIGHT ATTENDANTS: Cpt, boarding is complete, you may disconnect the jetbridge when you're ready."
            output2 = "FLIGHT ATTENDANTS: Cpt, everyone is on board, you can retract the jetbridge at your convenience."
            window['PaxSimOutput'].print(random.choice([output1, output2,]), text_color="cyan")
            PaxTimer_running = False

    if event == 'Load Cargo' and PStype != 'none':
        if not CargoTimer_running:
            if PStype == 'wide':
                CargoTimer_value = random.randint(180, 300) #/ 30
            elif PStype == 'narrow':
                CargoTimer_value = random.randint(120, 180) #/ 30
            CargoTimer_start = CargoTimer_value
            CargoTimer_running = True
            window['-CargoProgress-'].update_bar(100)
            output1 = "GROUND CREW: Cpt, we have started loading the baggage, we should be done in " + str(round(CargoTimer_start / 60)) + " minutes."
            output2 = "GROUND CREW: Cpt, we have starting loading the cargo, and it should take approximately " + str(round(CargoTimer_start / 60)) + " minutes."
            output3 = "FLIGHT ATTENDANTS: Cpt, we have just connected the cargo truck. Loading should be done in about " + str(round(CargoTimer_start / 60)) + " minutes."
            window['PaxSimOutput'].print(random.choice([output1, output2, output3]), text_color="yellow")
    if event == 'Load Cargo' and PStype == 'none':
        window['PaxSimOutput'].print('Select Widebody or Narrowbody!', text_color="red")
    if CargoTimer_running and PStype != 'none':
        CargoTimer_value -= 1
        CargoProgress = (CargoTimer_value / CargoTimer_start) * 100
        window['-CargoProgress-'].update_bar(CargoProgress)
        if CargoTimer_value <= 0:
            output1 = "GROUND CREW: Cpt, we finished loading the cargo."
            output2 = "GROUND CREW: Cpt, all the cargo is on board."
            window['PaxSimOutput'].print(random.choice([output1, output2]), text_color="yellow")
            CargoTimer_running = False


    if event == 'Load Catering' and PStype != 'none':
        if not CateringTimer_running:
            if PStype == 'wide':
                CateringTimer_value = random.randint(120, 180) #/ 30
            elif PStype == 'narrow':
                CateringTimer_value = random.randint(60, 120) #/ 30
            CateringTimer_start = CateringTimer_value
            CateringTimer_running = True
            window['-CateringProgress-'].update_bar(100)
            output1 = "GROUND CREW: Cpt, we have starting loading the catering, expected duration " + str(round(CateringTimer_start / 60)) + " minutes."
            output2 = "GROUND CREW: Cpt, we have starting loading the food and drinks, and it should take approximately " + str(round(CateringTimer_start / 60)) + " minutes."
            output3 = "GROUND CREW: Cpt, we have just connected the catering truck. We should be done in about " + str(round(CateringTimer_start / 60)) + " minutes."
            window['PaxSimOutput'].print(random.choice([output1, output2, output3]), text_color="magenta")
    if event == 'Load Catering' and PStype == 'none':
        window['PaxSimOutput'].print('Select Widebody or Narrowbody!', text_color="red")
    if CateringTimer_running and PStype != 'none':
        CateringTimer_value -= 1
        CateringProgress = (CateringTimer_value / CateringTimer_start) * 100
        window['-CateringProgress-'].update_bar(CateringProgress)
        if CateringTimer_value <= 0:
            output1 = "GROUND CREW: Cpt, we have finshed loading the catering."
            output2 = "GROUND CREW: Cpt, all foods and drinks are on board."
            output3 = "GROUND CREW: Cpt, all the catering is on board."
            window['PaxSimOutput'].print(random.choice([output1, output2, output3]), text_color="magenta")
            CateringTimer_running = False



    if event == 'NEXT' or event =='Start':
        window['Start'].update(visible=False)
        window['checklist-Header'].update(visible=True)
        window['BACK'].update(visible=True)
        window['OK'].update(visible=True)
        window['NEXT'].update(visible=True)

        try:
            currentPage += 1

            sheetName = Checklist.worksheets[currentPage]
            window['checklist-Header'].update(sheetName.title)

            while itemNumber <= 9:
                itemNumber += 1
                itemData = getChecklistItem(currentPage, itemNumber)
                if itemData[2] == 'uncompleted':
                    window['checklist-line' + str(itemNumber) +'a'].update(itemData[0], text_color=color3, background_color=color1)
                    window['checklist-line' + str(itemNumber) + 'b'].update(itemData[1], text_color=color3, background_color=color1)
                if itemData[2] == 'completed':
                    window['checklist-line' + str(itemNumber) +'a'].update(itemData[0], text_color='green', background_color=color1)
                    window['checklist-line' + str(itemNumber) + 'b'].update(itemData[1], text_color='green',background_color=color1)

            itemNumber = 0
            currentItem = 0
        except:
            currentPage -= 1

    elif event == 'OK':
        currentItem += 1
        previousItem = currentItem - 1
        if currentItem <= 10:
            itemData = getChecklistItem(currentPage, currentItem)
            if itemData[1] == ' ':
                window['checklist-line' + str(currentItem) + 'a'].update(itemData[0], text_color='green', background_color=color1)
                window['checklist-line' + str(currentItem) + 'b'].update(itemData[1], text_color='green', background_color=color1)
                previousItemData = getChecklistItem(currentPage, previousItem)
                window['checklist-line' + str(previousItem) + 'a'].update(previousItemData[0], text_color='green', background_color=color1)
                window['checklist-line' + str(previousItem) + 'b'].update(previousItemData[1], text_color='green', background_color=color1)
                updateChecklistItem(currentPage, currentItem, 'completed')
                updateChecklistItem(currentPage, previousItem, 'completed')
            else:
                nextItem = currentItem + 1
                window['checklist-line' + str(currentItem) + 'a'].update(itemData[0], text_color=color3, background_color='green')
                window['checklist-line' + str(currentItem) + 'b'].update(itemData[1], text_color=color3, background_color='green')
                updateChecklistItem(currentPage, currentItem, 'uncompleted')
                if previousItem != 0:
                    previousItemData = getChecklistItem(currentPage, previousItem)
                    window['checklist-line' + str(previousItem) + 'a'].update(previousItemData[0], text_color='green', background_color=color1)
                    window['checklist-line' + str(previousItem) + 'b'].update(previousItemData[1], text_color='green', background_color=color1)
                    updateChecklistItem(currentPage, previousItem, 'completed')
        elif currentItem == 11:
            currentItem = 10
            window['checklist-line' + str(currentItem) + 'a'].update(itemData[0], text_color='green', background_color=color1)
            window['checklist-line' + str(currentItem) + 'b'].update(itemData[1], text_color='green', background_color=color1)
            updateChecklistItem(currentPage, currentItem, 'completed')

    elif event == 'BACK':
        try:
            currentPage -= 1

            sheetName = Checklist.worksheets[currentPage]
            window['checklist-Header'].update(sheetName.title)

            while itemNumber <= 9:
                itemNumber += 1
                itemData = getChecklistItem(currentPage, itemNumber)
                if itemData[2] == 'uncompleted':
                    window['checklist-line' + str(itemNumber) +'a'].update(itemData[0], text_color=color3, background_color=color1)
                    window['checklist-line' + str(itemNumber) + 'b'].update(itemData[1], text_color=color3, background_color=color1)
                if itemData[2] == 'completed':
                    window['checklist-line' + str(itemNumber) +'a'].update(itemData[0], text_color='green', background_color=color1)
                    window['checklist-line' + str(itemNumber) + 'b'].update(itemData[1], text_color='green',background_color=color1)
            itemNumber = 0
            currentItem = 0
        except:
            currentPage += 1

    elif event == 'Airbus':
        window['Airbus'].update(button_color = (color2,color3 )); window['Boeing'].update(button_color=(color3, color2)); window['Bombardier'].update(button_color=(color3, color2)); window['Embraer'].update(button_color=(color3, color2)); window['McDonnell Douglas'].update(button_color=(color3, color2))
        new_choices = ['A220', 'A318', 'A319', 'A320', 'A321', 'A332', 'A333', 'A339', 'A346', 'A359', 'A388']
        window['selected_aircraft'].update(values=new_choices)

    elif event == 'Boeing':
        window['Boeing'].update(button_color = (color2,color3 )); window['Airbus'].update(button_color=(color3, color2)); window['Bombardier'].update(button_color=(color3, color2)); window['Embraer'].update(button_color=(color3, color2)); window['McDonnell Douglas'].update(button_color=(color3, color2))
        new_choices = ['B712', 'B737', 'B738', 'B739', 'B742', 'B744', 'B748', 'B752', 'B763', 'B772', 'B77L', 'B77W', 'B77F',
               'B788', 'B789', 'B78X']
        window['selected_aircraft'].update(values=new_choices)

    elif event == 'Bombardier':
        window['Bombardier'].update(button_color = (color2,color3 )); window['Boeing'].update(button_color=(color3, color2)); window['Airbus'].update(button_color=(color3, color2)); window['Embraer'].update(button_color=(color3, color2)); window['McDonnell Douglas'].update(button_color=(color3, color2))
        window['Bombardier'].update(button_color=(color2, color3))
        new_choices = ['CL350', 'CRJ2', 'CRJ7', 'CRJ9', 'CRJX', 'DH8D']
        window['selected_aircraft'].update(values=new_choices)

    elif event == 'Embraer':
        window['Embraer'].update(button_color = (color2,color3 )); window['Boeing'].update(button_color=(color3, color2)); window['Bombardier'].update(button_color=(color3, color2)); window['Airbus'].update(button_color=(color3, color2)); window['McDonnell Douglas'].update(button_color=(color3, color2))
        new_choices = ['E175', 'E190']
        window['selected_aircraft'].update(values=new_choices)

    elif event == 'McDonnell Douglas':
        window['McDonnell Douglas'].update(button_color = (color2,color3 )); window['Boeing'].update(button_color=(color3, color2)); window['Bombardier'].update(button_color=(color3, color2)); window['Embraer'].update(button_color=(color3, color2)); window['Airbus'].update(button_color=(color3, color2))
        new_choices = ['DC10', 'DC1F', 'MD11', 'MD1F']
        window['selected_aircraft'].update(values=new_choices)

    elif event == 'App Details ↗':
        webbrowser.open("https://community.infiniteflight.com/t/the-unofficial-infinite-aircraft-calculator-using-community-data/869648")

    elif event == 'Credits':
        window['Credits'].update(button_color = (color2,color3 )); window['Departure Data'].update(button_color=(color3, color2)); window['Arrival Data'].update(button_color=(color3, color2)); window['Fuel Burn Data'].update(button_color=(color3, color2)); window['Other Data'].update(button_color=(color3, color2));
        output_text = '\n'.join([
            f"DeerCrusher: Takeoff and Landing Profile Data for Reworked Aircraft",
            f"Kuba_Jaroszczyk: Takeoff and Landing Profile Data for Older Aircraft",
            f"AndrewWu: Fuel Burn Data and Recommended Flight Profiles",
            f"Jan: Ceiling, Normal Range, Cruise Spd, MMO Spd Data",
            f"darkeyes: This app. Thank you for using.",
            f"\nThe data here isn't perfectly accurate. It is simply intended to offer a basic guidance for your flight. ",
            f"\nVerison 1.01 ",

        ])
        window['output'].update(output_text, text_color=color3)

    elif event == 'Departure Data':
        window['Departure Data'].update(button_color = (color2,color3 )); window['Credits'].update(button_color=(color3, color2)); window['Arrival Data'].update(button_color=(color3, color2)); window['Fuel Burn Data'].update(button_color=(color3, color2)); window['Other Data'].update(button_color=(color3, color2));
        try:
            acType = value["selected_aircraft"]
            sheet = book[acType]
            ac_load = value['load']
            row = InputLoadtoDataRow(ac_load)
            DepPower, DepFlaps, DepRotate, DepAirBy = getDepatureData(row)
            if value["selected_aircraft"] == 'CL350':
                window['askOAT'].update(visible=True)
                window['OATinput'].update(visible=True)
                oat = value["OATinput"]
                DepPower = CL350(oat)
                output_text = '\n'.join([f"Flap Setting: {DepFlaps} ", f"\nPower: \n{DepPower} ", f"\nRotate: {DepRotate} ", f"Airborne By: {DepAirBy}"])
            else:
                window['askOAT'].update(visible=False)
                window['OATinput'].update(visible=False)
                output_text = '\n'.join(
                    [f"Flap Setting: {DepFlaps} ", f"\nPower: {DepPower} ", f"\nRotate: {DepRotate} ",
                     f"Airborne By: {DepAirBy}"])
            window['output'].update(output_text, text_color=color3)
        except KeyError:
            errorMessage = 'Select Aircraft Type!'
            window['output'].update(errorMessage, text_color='red')
        except ValueError:
            errorMessage = 'Enter Aircraft Load as an Integer!'
            window['output'].update(errorMessage, text_color='red')
        except:
            errorMessage = 'Error, try again'
            mycolor = 'red'
            window['output'].update(errorMessage, text_color='red')

    elif event == 'Arrival Data':
        window['Arrival Data'].update(button_color = (color2,color3 )); window['Credits'].update(button_color=(color3, color2)); window['Departure Data'].update(button_color=(color3, color2)); window['Fuel Burn Data'].update(button_color=(color3, color2)); window['Other Data'].update(button_color=(color3, color2));
        try:
            acType = value["selected_aircraft"]
            sheet = book[acType]
            ac_load = value['load']
            row = InputLoadtoDataRow(ac_load)
            ArrFlaps, ArrApprSpd, ArrFlareSpd, Flaps = getArrivalData(row)
            output_text = '\n'.join([f"Flap Setting: {ArrFlaps} ", f"\nApproach Speed: {ArrApprSpd}", f"Flare Speed: {ArrFlareSpd} ", f"\nFlap Spds: \n{Flaps} "])
            window['output'].update(output_text, text_color=color3)
        except KeyError:
            errorMessage = 'Select Aircraft Type!'
            mycolor = 'red'
            window['output'].update(errorMessage, text_color='red')
        except ValueError:
            errorMessage = 'Enter Aircraft Load as an Integer!'
            mycolor = 'red'
            window['output'].update(errorMessage, text_color='red')
        except:
            errorMessage = 'Error, try again'
            mycolor = 'red'
            window['output'].update(errorMessage, text_color='red')

    elif event == 'Fuel Burn Data':
         window['Fuel Burn Data'].update(button_color = (color2,color3 )); window['Credits'].update(button_color=(color3, color2)); window['Departure Data'].update(button_color=(color3, color2)); window['Arrival Data'].update(button_color=(color3, color2)); window['Other Data'].update(button_color=(color3, color2));
         try:
            acType = value["selected_aircraft"]
            sheet = book[acType]
            ac_load = value['load']
            row = InputLoadtoDataRow(ac_load)
            Even, Odd, Med, RecWest, RecEast = getFuelBurnData(row)
            output_text = '\n'.join([f"West/Even Cruise Alt: {Even} ", f"East/Odd Cruise Alt: {Odd} ", f"High Fuel Burn: {Med} ",
                                      f"\n\nRecommend Flight Profile West/Even: {RecWest}", f"\nRecommend Flight Profile East/Odd: {RecEast}"])
            window['output'].update(output_text, text_color=color3)
         except KeyError:
            errorMessage = 'Select Aircraft Type!'
            mycolor = 'red'
            window['output'].update(errorMessage, text_color='red')
         except ValueError:
             errorMessage = 'Enter Aircraft Load as an Integer!'
             mycolor = 'red'
             window['output'].update(errorMessage, text_color='red')
         except:
             errorMessage = 'Error, try again'
             mycolor = 'red'
             window['output'].update(errorMessage, text_color='red')

    elif event == 'Other Data':
        window['Other Data'].update(button_color = (color2,color3 )); window['Credits'].update(button_color=(color3, color2)); window['Departure Data'].update(button_color=(color3, color2)); window['Arrival Data'].update(button_color=(color3, color2)); window['Fuel Burn Data'].update(button_color=(color3, color2));
        try:
            acType = value["selected_aircraft"]
            sheet = book[acType]
            ac_load = value['load']
            row = InputLoadtoDataRow(ac_load)
            Ceiling, Cruise, MMO, Range = getOtherData(row)
            output_text = '\n'.join([f"Ceiling: {Ceiling}", f"Normal Range: {Range}", f"\nCruise Spd: {Cruise}", f"MMO Spd: {MMO} \n"])
            window['output'].update(output_text, text_color=color3)
        except KeyError:
            errorMessage = 'Select Aircraft Type!'
            mycolor = 'red'
            window['output'].update(errorMessage, text_color='red')
        except ValueError:
            errorMessage = 'Enter Aircraft Load as an Integer!'
            mycolor = 'red'
            window['output'].update(errorMessage, text_color='red')
        except:
            errorMessage = 'Error, try again'
            mycolor = 'red'
            window['output'].update(errorMessage, text_color='red')

    elif event == 'IFC':
        webbrowser.open("https://community.infiniteflight.com")
    elif event == 'IF Pro Discord':
        webbrowser.open("https://discord.com/channels/914868776758038528/962704044873379880")
    elif event == 'Live Flight':
        webbrowser.open("https://liveflight.app/")
    elif event == 'Map Flight':
        webbrowser.open("https://en.map-flight.com/")
    elif event == 'FlightAware':
        webbrowser.open("https://www.flightaware.com/")
    elif event == 'Flightradar24':
        webbrowser.open("https://www.flightradar24.com/")
    elif event == 'Windy':
        webbrowser.open("https://www.windy.com/")
    elif event == 'FPL to IF':
        webbrowser.open("https://fpltoif.com/")

window.close()
