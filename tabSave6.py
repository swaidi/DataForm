
import tkinter as tk
from tkinter import ttk
from tkinter.font import BOLD
from tkinter import scrolledtext
from tkinter.ttk import Style

import pandas as pd

from sys import platform as _platform




def saveinfo():

    valor1 = nameEntry.get()
    valor2 = mobileEntry.get()
    valor3 = emailEntry.get()
    valor4 = collegechoosen.get()
    valor5 = locationchoosen.get()

    valor6 = titleEntry.get()
    valor7 = detailsEntry.get()

    valor8 = budgetEntry.get()
    valor9 = dateEntry.get()
    valor10 = durationEntry.get()
    valor11 = setupEntry.get()
    valor12 = wrapEntry.get()
    valor13 = startEntry.get()
    valor14 = endEntry.get()
    valor15 = anticipatedEntry.get()
    valor16 = expectedEntry.get()

    valor17 = radioCME.get()
    valor18 = radioAV.get()
    valor19 = chkStudents.state()

    # this works with ttk.checkbutton but gives you : selected , alternate
    # valor20 = tk.IntVar()
    # valor20 = chkStudents.val.get() #this works with tk.checkbutton but gives you : 1, 0

    valor20 = chkFaculty.state()

    valor21 = chkStaff.state()
    valor22 = chkAlumni.state()
    valor23 = chkCommunity.state()
    valor24 = chkPublic.state()
    valor25 = radioStudRequirted.get()

    valor26 = vipEntry.get(1.0, tk.END)

    valor27 = chkCampus.state()
    valor28 = chkMedia.state()
    valor29 = chkOtherAD.state()
    valor30 = otherAdtextEntry.get()

    valor31 = radioSafety.get()

    valor32 = chkMonitor.state()
    valor33 = chkCheckIDs.state()
    valor34 = chkVIPsafety.state()
    valor35 = chkPatrol.state()
    valor36 = chkOtherSafety.state()
    valor37 = OtherSafetyTextEntry.get()

    valor38 = chkWiFi.state()

    valor39 = wifiITEntry.get()
    valor40 = chkDevices.state()
    valor41 = deviceITEntry.get()
    valor42 = chkOtherIT.state()
    valor43 = otherITtextEntry.get()

    valor44 = chkInstall.state()

    valor45 = chkCheckup.state()
    valor46 = chkOtherTech.state()
    valor47 = otherTechtextEntry.get()

    valor48 = chkTables.state()

    valor49 = tablesSerEntry.get()
    valor50 = chkChairs.state()
    valor51 = chairsSerEntry.get()
    valor52 = chkSignages.state()
    valor53 = signageSerEntry.get()
    valor54 = chkOtherSer.state()
    valor55 = otherSertextEntry.get()

    valor56 = radiocatering.get()

    valor57 = addRequirEntry.get(1.0, tk.END)
    valor58 = chkAgree.state()

    data.append([valor1, valor2, valor3, valor4, valor5, valor6, valor7, valor8, valor9, valor10, valor11, valor12, valor13, valor14, valor15, valor16, valor17, valor18, valor19, valor20, valor21, valor22, valor23, valor24, valor25, valor26, valor27, valor28, valor29, valor30, valor31, valor32, valor33, valor34, valor35, valor36, valor37, valor38, valor39, valor40, valor41, valor42, valor43, valor44, valor45, valor46, valor47, valor48, valor49, valor50, valor51, valor52, valor53, valor54, valor55, valor56, valor57, valor58])
    print(data)

    valor1 = nameEntry.delete(0, "end")
    valor2 = mobileEntry.delete(0, "end")
    valor3 = emailEntry.delete(0, "end")
    valor4 = collegechoosen.delete(0, "end")
    valor5 = locationchoosen.delete(0, "end")

    valor6 = titleEntry.delete(0, "end")
    valor7 = detailsEntry.delete(0, "end")

    valor8 = budgetEntry.delete(0, "end")
    valor9 = dateEntry.delete(0, "end")
    valor10 = durationEntry.delete(0, "end")
    valor11 = setupEntry.delete(0, "end")
    valor12 = wrapEntry.delete(0, "end")
    valor13 = startEntry.delete(0, "end")
    valor14 = endEntry.delete(0, "end")
    valor15 = anticipatedEntry.delete(0, "end")
    valor16 = expectedEntry.delete(0, "end")

    valor17 = radioCME.set("0")
    valor18 = radioAV.set("0")
    # valor19 = radioRegist.set("0")



    valor21 = chkFaculty.state()
    valor22 = chkStaff.state()
    valor23 = chkAlumni.state()
    valor24 = chkCommunity.state()
    valor25 = chkPublic.state()
    valor26 = radioStudRequirted.set("0")
    valor27 = vipEntry.delete('1.0', tk.END)

    valor28 = chkCampus.state()
    valor29 = chkMedia.state()
    valor30 = chkOtherAD.state()
    valor31 = otherAdtextEntry.delete(0, "end")

    valor32 = radioSafety.set("0")
    valor33 = chkMonitor.state()
    valor34 = chkCheckIDs.state()
    valor35 = chkVIPsafety.state()
    valor36 = chkPatrol.state()
    valor37 = chkOtherSafety.state()
    valor38 = OtherSafetyTextEntry.delete(0, "end")

    valor39 = chkWiFi.state()
    valor40 = wifiITEntry.delete(0, "end")
    valor41 = chkDevices.state()
    valor42 = deviceITEntry.delete(0, "end")
    valor43 = chkOtherIT.state()
    valor44 = otherITtextEntry.delete(0, "end")

    valor45 = chkInstall.state()
    valor46 = chkCheckup.state()
    valor47 = chkOtherTech.state()
    valor48 = otherTechtextEntry.delete(0, "end")

    valor49 = chkTables.state()
    valor50 = tablesSerEntry.delete(0, "end")
    valor51 = chkChairs.state()
    valor52 = chairsSerEntry.delete(0, "end")
    valor53 = chkSignages.state()
    valor54 = signageSerEntry.delete(0, "end")
    valor55 = chkOtherSer.state()
    valor56 = otherSertextEntry.delete(0, "end")

    valor57 = radiocatering.set("0")
    valor58 = addRequirEntry.delete('1.0', tk.END)
    valor59 = chkAgree.state()

def export():

    df = pd.DataFrame(data)
    df.to_excel("DataBase.xlsx")




# --- main ---
df = pd.DataFrame
data = []

# intializing the window
window = tk.Tk()
window.title("Event Form Request")



# configuring size of the window
window.geometry('450x750')

#Create Tab Control
TAB_CONTROL = ttk.Notebook(window)

#Tab1
TAB1 = ttk.Frame(TAB_CONTROL)
TAB_CONTROL.add(TAB1, text='  1 / 3  ')

#Tab2
TAB2 = ttk.Frame(TAB_CONTROL)
TAB_CONTROL.add(TAB2, text='  2 / 3  ')

#Tab3
TAB3 = ttk.Frame(TAB_CONTROL)
TAB_CONTROL.add(TAB3, text='  3 / 3  ')

TAB_CONTROL.pack(expand=1, fill="both")


###############
#TAB 1
###############

ttk.Label(TAB1, text=" College and Event Details", font=("arial", 10, BOLD)).place(x=100, y=20)

ttk.Label(TAB1, text="Name of Incharge Person:").place(x=30, y=50, width=200)
nameEntry = ttk.Entry(TAB1)
nameEntry.place(x=200, y=50, width=160)


ttk.Label(TAB1, text="Mobile:").place(x=30, y=80, width=80)
mobileEntry = ttk.Entry(TAB1)
mobileEntry.place(x=200, y=80, width=160)

ttk.Label(TAB1, text="Email:").place(x=30, y=110, width=80)
emailEntry = ttk.Entry(TAB1)
emailEntry.place(x=200, y=110, width=160)



ttk.Label(TAB1, text="College:", state="readonly").place(x=30, y=140)
# Combobox creation

collegechoosen = ttk.Combobox(TAB1, width=23, state="readonly")

# Adding combobox drop down list
collegechoosen['values'] = (' COM-R',
                          ' COM-J',
                          ' CON-R',
                          ' CON-J',
                          ' CON-A',
                          ' COP',
                          ' COD',
                          ' COAMS-R',
                          ' COAMS-J',
                          ' COAMS-A',
                          ' COSHP-R',
                          ' COSHP-J')
collegechoosen.set('Please Select ..')
collegechoosen.place(x=200, y=140)
collegechoosen.current()


ttk.Label(TAB1, text="Location:").place(x=30, y=170)
# Combobox creation

locationchoosen = ttk.Combobox(TAB1, width=23, state="readonly")

# Adding combobox drop down list
locationchoosen['values'] = (
                          ' CONF. ROOM 12TH FLOOR',
                          ' DINING ROOM 12 FLOOR',
                          ' MAJLIS 12TH FLOOR',
                          ' ROOM 21 M. FLOOR',)
locationchoosen.set('Please Select ..')
locationchoosen.place(x=200, y=170)
locationchoosen.current()

ttk.Label(TAB1, text="Event Title:").place(x=30, y=200, width=80)
titleEntry = ttk.Entry(TAB1)
titleEntry.place(x=200, y=200, width=160)

ttk.Label(TAB1, text="Event Details:").place(x=30, y=230, width=80)
detailsEntry = ttk.Entry(TAB1)
detailsEntry.place(x=200, y=230, width=160)

ttk.Label(TAB1, text="Event Budget:").place(x=30, y=260, width=80)
budgetEntry = ttk.Entry(TAB1)
budgetEntry.place(x=200, y=260, width=160)


ttk.Label(TAB1, text="Event Date:").place(x=30, y=290, width=80)
dateEntry = ttk.Entry(TAB1)
dateEntry.place(x=200, y=290, width=160)

ttk.Label(TAB1, text="Duration:").place(x=30, y=320, width=80)
durationEntry = ttk.Entry(TAB1)
durationEntry.place(x=200, y=320, width=160)

ttk.Label(TAB1, text="Setup Date:").place(x=30, y=350, width=80)
setupEntry = ttk.Entry(TAB1)
setupEntry.place(x=200, y=350, width=160)

ttk.Label(TAB1, text="Wrap Date:").place(x=30, y=380, width=80)
wrapEntry = ttk.Entry(TAB1)
wrapEntry.place(x=200, y=380, width=160)

ttk.Label(TAB1, text="Start Time:").place(x=30, y=410, width=80)
startEntry = ttk.Entry(TAB1)
startEntry.place(x=200, y=410, width=160)

ttk.Label(TAB1, text="End Time:").place(x=30, y=440, width=80)
endEntry = ttk.Entry(TAB1)
endEntry.place(x=200, y=440, width=160)


ttk.Label(TAB1, text="Number of Anticipated:").place(x=30, y=470, width=180)
anticipatedEntry = ttk.Entry(TAB1)
anticipatedEntry.place(x=200, y=470, width=160)

ttk.Label(TAB1, text="Number of Expected\nAttendees:").place(x=30, y=495, width=180)
expectedEntry = ttk.Entry(TAB1)
expectedEntry.place(x=200, y=500, width=160)



######################################### RADIOPOINT

ttk.Label(TAB1, text="Is this Event a CME?").place(x=30, y=540)
radioCME = tk.IntVar()
radioOne = ttk.Radiobutton(TAB1, text='Yes',
                         variable=radioCME, value=1)
radioTwo = ttk.Radiobutton(TAB1, text='No',
                         variable=radioCME, value=2)
labelValue = ttk.Label(TAB1, textvariable=radioCME.get())
radioOne.place(x=230, y=540)
radioTwo.place(x=300, y=540)






ttk.Label(TAB1, text="Required Audio/ Visual?").place(x=30, y=570)
radioAV = tk.IntVar()
radioOne = ttk.Radiobutton(TAB1, text='Yes',
                         variable=radioAV, value=1)
radioTwo = ttk.Radiobutton(TAB1, text='No',
                         variable=radioAV, value=2)
labelValue = ttk.Label(TAB1, textvariable=radioAV.get())
radioOne.place(x=230, y=570)
radioTwo.place(x=300, y=570)




###############
#TAB 2
###############


ttk.Label(TAB2, text="General Requirement", font=("arial", 10, BOLD)).place(x=100, y=20)




ttk.Label(TAB2, text='Select the Targeted Audience:').place(x=30, y=50, width=160)


#######################################################################


valor20 = tk.IntVar()
chkStudents = ttk.Checkbutton(TAB2, text='Students', variable=valor20)
chkStudents.place( x=30, y=80, width=80)

valor21 = tk.IntVar()
chkFaculty = ttk.Checkbutton(TAB2, text='Faculty', variable=valor21)
chkFaculty.place(x=130, y=80, width=80)

valor22 = tk.IntVar()
chkStaff = ttk.Checkbutton(TAB2, text='Staff',variable=valor22)
chkStaff.place(x=230, y=80, width=80)

valor23 = tk.IntVar()
chkAlumni = ttk.Checkbutton(TAB2, text='Alumni',variable=valor23)
chkAlumni.place(x=330, y=80, width=80)

valor24 = tk.IntVar()
chkCommunity = ttk.Checkbutton(TAB2, text='Healthcare Community',variable=valor24)
chkCommunity.place(x=30, y=110, width=200)

valor25 = tk.IntVar()
chkPublic = ttk.Checkbutton(TAB2, text='Public',variable=valor25)
chkPublic.place(x=230, y=110, width=80)

ttk.Label(TAB2, text="Required Students Attending?").place(x=30, y=150)
radioStudRequirted = tk.IntVar()
radioOne = ttk.Radiobutton(TAB2, text='Yes',
                         variable=radioStudRequirted, value=1)
radioTwo = ttk.Radiobutton(TAB2, text='No',
                         variable=radioStudRequirted, value=2)
labelValue = ttk.Label(TAB2, textvariable=radioStudRequirted.get())

radioOne.place(x=230, y=150)
radioTwo.place(x=300, y=150)

ttk.Label(TAB2, text="Please list any dignitaries, VIPs who may attend "
                     "as guest speakers,\npanelists, etc. Or invited guests"
                    "with title and place of employment:").place(x=30, y=200, width=300)
vipEntry = scrolledtext.ScrolledText(TAB2, width=20, height=4, wrap=tk.WORD)

# vipEntry = ttk.Entry(TAB2)
vipEntry.place(x=30, y=240, width=300)

ttk.Label(TAB2, text="Advertisment & Marketing", font=("arial", 10, BOLD)).place(x=30, y=320, width=400)

ttk.Label(TAB2, text='What will be used to advertise this event?').place(x=30, y=350, width=300)

valor28 = tk.IntVar()
chkCampus = ttk.Checkbutton(TAB2, text='On Campus', variable= valor28)
chkCampus.place(x=30, y=380, width=100)
valor29 = tk.IntVar()
chkMedia = ttk.Checkbutton(TAB2, text='KSAU-HS Social Media', variable= valor29 )
chkMedia.place(x=130, y=380, width=160)
valor30 = tk.IntVar()
chkOtherAD = ttk.Checkbutton(TAB2, text='Other',variable= valor30 )
chkOtherAD.place(x=30, y=410, width=120)

ttk.Label(TAB2, text='Specify:').place(x=100, y=410, width=200)

otherAdtextEntry = ttk.Entry(TAB2)
otherAdtextEntry.place(x=150, y=410, width=200)


ttk.Label(TAB2, text="This event requires a public \nsafety presence?").place(x=30, y=450, width=200)
radioSafety = tk.IntVar()
radioOne = ttk.Radiobutton(TAB2, text='Yes',
                         variable=radioSafety, value=1)
radioTwo = ttk.Radiobutton(TAB2, text='No',
                         variable=radioSafety, value=2)
labelValue = ttk.Label(TAB2, textvariable=radioSafety.get())
radioOne.place(x=230, y=450)
radioTwo.place(x=300, y=450)



ttk.Label(TAB2, text='For what purpose:').place(x=30, y=500, width=200)

valor33 = tk.IntVar()
chkMonitor = ttk.Checkbutton(TAB2, text='To Monitor the Event Entrance', variable=valor33)
chkMonitor.place(x=30, y=530, width=200)

valor34 = tk.IntVar()
chkCheckIDs = ttk.Checkbutton(TAB2, text='Check IDs', variable=valor34)
chkCheckIDs.place(x=230, y=530, width=80)

valor35 = tk.IntVar()
chkVIPsafety = ttk.Checkbutton(TAB2, text='VIP Safety', variable=valor35)
chkVIPsafety.place(x=230, y=560, width=80)

valor36 = tk.IntVar()
chkPatrol = ttk.Checkbutton(TAB2, text='Patrol the Event', variable=valor36)
chkPatrol.place(x=30, y=560, width=200)

valor37 = tk.IntVar()
chkOtherSafety = ttk.Checkbutton(TAB2, text='Other', variable=valor37)
chkOtherSafety.place(x=30, y=590, width=120)

ttk.Label(TAB2, text='Specify:').place(x=100, y=590, width=200)

OtherSafetyTextEntry = ttk.Entry(TAB2)
OtherSafetyTextEntry.place(x=150, y=590, width=200)








###############
#TAB 3
###############


ttk.Label(TAB3, text="Technical Requirement", font=("arial", 10, BOLD)).place(x=100, y=20, width=160)

ttk.Label(TAB3, text='IT support requested:').place(x=30, y=55, width=200)

valor39 = tk.IntVar()
chkWiFi = ttk.Checkbutton(TAB3, text='Wi-Fi', variable=valor39)
chkWiFi.place(x=30, y=80, width=80)

ttk.Label(TAB3, text='Qty:').place(x=120, y=80, width=100)
wifiITEntry = ttk.Entry(TAB3)
wifiITEntry.place(x=150, y=80, width=50)

valor41 = tk.IntVar()
chkDevices = ttk.Checkbutton(TAB3, text='Devices', variable=valor41)
chkDevices.place(x=30, y=105, width=60)

ttk.Label(TAB3, text='Qty:').place(x=120, y=105, width=100)
deviceITEntry = ttk.Entry(TAB3)
deviceITEntry.place(x=150, y=105, width=50)

valor43 = tk.IntVar()
chkOtherIT = ttk.Checkbutton(TAB3, text='Other', variable=valor43)
chkOtherIT.place(x=30, y=130, width=80)

ttk.Label(TAB3, text='Specify:').place(x=100, y=130, width=100)

otherITtextEntry = ttk.Entry(TAB3)
otherITtextEntry.place(x=150, y=130, width=200)





ttk.Label(TAB3, text='Technical support requested:').place(x=30, y=170, width=200)

valor45 = tk.IntVar()
chkInstall = ttk.Checkbutton(TAB3, text='Installation', variable=valor45)
chkInstall.place(x=30, y=190, width=80)

valor46 = tk.IntVar()
chkCheckup = ttk.Checkbutton(TAB3, text='Check UP', variable=valor46)
chkCheckup.place(x=30, y=215, width=80)

valor47 = tk.IntVar()
chkOtherTech = ttk.Checkbutton(TAB3, text='Other', variable=valor47)
chkOtherTech.place(x=30, y=240, width=80)

ttk.Label(TAB3, text='Specify:').place(x=100, y=240, width=100)

otherTechtextEntry = ttk.Entry(TAB3)
otherTechtextEntry.place(x=150, y=240, width=200)


ttk.Label(TAB3, text='Pick Up Services:').place(x=30, y=280, width=200)

valor49 = tk.IntVar()
chkTables = ttk.Checkbutton(TAB3, text='Tables', variable=valor49)
chkTables.place(x=30, y=305, width=80)

ttk.Label(TAB3, text='Qty:').place(x=120, y=305, width=100)
tablesSerEntry = ttk.Entry(TAB3)
tablesSerEntry.place(x=150, y=305, width=50)

valor51 = tk.IntVar()
chkChairs = ttk.Checkbutton(TAB3, text='Chairs', variable=valor51)
chkChairs.place(x=30, y=330, width=80)

ttk.Label(TAB3, text='Qty:').place(x=120, y=330, width=100)
chairsSerEntry = ttk.Entry(TAB3)
chairsSerEntry.place(x=150, y=330, width=50)

valor53 = tk.IntVar()
chkSignages = ttk.Checkbutton(TAB3, text='Signages', variable=valor53)
chkSignages.place(x=30, y=355, width=80)

ttk.Label(TAB3, text='Qty:').place(x=120, y=355, width=100)
signageSerEntry = ttk.Entry(TAB3)
signageSerEntry.place(x=150, y=355, width=50)

valor55 = tk.IntVar()
chkOtherSer = ttk.Checkbutton(TAB3, text='Other', variable=valor55)
chkOtherSer.place(x=30, y=380, width=80)

ttk.Label(TAB3, text='Specify:').place(x=100, y=380, width=100)

otherSertextEntry = ttk.Entry(TAB3)
otherSertextEntry.place(x=150, y=380, width=200)



ttk.Label(TAB3, text="Catering Requested?").place(x=30, y=435)
radiocatering = tk.IntVar()
radioOne = ttk.Radiobutton(TAB3, text='Yes',
                         variable=radiocatering, value=1)
radioTwo = ttk.Radiobutton(TAB3, text='No',
                         variable=radiocatering, value=2)
labelValue = ttk.Label(TAB3, textvariable=radiocatering.get())
radioOne.place(x=170, y=435)
radioTwo.place(x=250, y=435)






ttk.Label(TAB3, text="Additional Requirement", ).place(x=30, y=470, width=300)
addRequirEntry = scrolledtext.ScrolledText(TAB3, width=300, height=4, wrap=tk.WORD)
addRequirEntry.place(x=30, y=490, width=300)

ttk.Label(TAB3, text="Agreeing of taking Responsibility for good condition and the cleanness \n of the Event Location as received:", font=("arial", 8, BOLD) ).place(x=30, y=600, width=400)


valor59 = tk.IntVar()

chkAgree = ttk.Checkbutton(TAB3, text='Agree',variable=valor59)
chkAgree.place(x=30, y=650, width=200)








ttk.Button(TAB3, text="Save", command=saveinfo, ).place(x=150, y=650, width=50)
ttk.Button(TAB3, text="Export", command=export).place(x=250, y=650, width=50)
ttk.Button(TAB3, text="Quit", command=window.destroy).place(x=350, y=650, width=50)


ttk.Label(TAB3, text="KSAU-HS/MUR Dept.", font=("arial", 5, BOLD)).place(x=30, y=700, width=160)


#Calling Main()
window.mainloop()
