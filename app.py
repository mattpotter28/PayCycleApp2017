# Pay Cycle Application 2017
# Created by Matt Potter
# March 13 2017

import configparser
import base64
import pypyodbc
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter.filedialog import asksaveasfilename
from tkinter import messagebox
from datetime import datetime

class MainApplication:
    def __init__(self, master, conn, serv):
        # region constructor

        # SQL Connection
        MainApplication.connection = conn
        MainApplication.cursor = self.connection.cursor()
        self.SQLCommand = ("")
        self.serverName = serv

        # window creation
        self.master = master
        self.frame = ttk.Frame(self.master)

        # sets titles
        self.master.wm_title("Pay Cycle Setup")
        self.serverLab = tk.Label(self.frame, text="Connected to: " + self.serverName)
        self.serverLab.grid(column=0, row=0, padx=5, pady=5, sticky='W')

        # create widgets on window
        self.createWidgets()

        # places completed frame
        self.frame.grid(column=0, row=0)

    def createWidgets(self):
        self.fields = ('Location', 'NBO Pay Group', 'ADP Store Code', 'Pay Cycle', 'Tip Share', 'Begin Date', 'End Date')
        self.days = list(range(1, 32))
        self.months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        self.monthsDict = {'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06', 'Jul': '07',
                           'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'}
        self.years = list(range(2017, 2080))

        for field in self.fields:
            if field == 'Location':
                self.LocationVariable = tk.StringVar()
                self.SQLCommand = ("SELECT strLocationReportHeading FROM POSLabor.dbo.tbl_UnitDescription UNION SELECT SiteName FROM POSLabor.dbo.NBO_Sites WHERE SiteNumber not in (select MLSiteNumber from POSLabor.dbo.tbl_UnitDescription)")
                MainApplication.cursor.execute(self.SQLCommand)
                self.siteNames = MainApplication.cursor.fetchall()
                self.siteNames = [str(s).strip('{(\"\',)}') for s in self.siteNames]
                self.siteNames.sort()

                self.lLab = tk.Label(self.frame, text="Location: ")
                self.lLab.grid(column=0, row=1, padx=5, pady=5, sticky='W')
                self.lEntry = ttk.Combobox(self.frame, textvariable=self.LocationVariable, value=self.siteNames,
                                           state="readonly", width=30)
                self.lEntry.grid(column=2, columnspan=5, row=1, padx=5, pady=5, sticky='W')

                self.lEntry.bind('<<ComboboxSelected>>', self.locationSelect)

            elif field == 'NBO Pay Group':
                MainApplication.PayGroupVariable = tk.StringVar()
                self.SQLCommand = ("select distinct [PayrollGroupName] from [POSLabor].[dbo].[NBO_PayGroup]")
                MainApplication.cursor.execute(self.SQLCommand)
                MainApplication.payGroups = MainApplication.cursor.fetchall()
                MainApplication.payGroups = [str(s).strip('{(\"\',)}') for s in MainApplication.payGroups]

                self.gLab = tk.Label(self.frame, text="NBO Pay Group: ")
                self.gLab.grid(column=0, row=2, padx=5, pady=5, sticky='W')
                MainApplication.gEntry = ttk.Combobox(self.frame, textvariable=MainApplication.PayGroupVariable, values=MainApplication.payGroups,
                                           state="readonly", width=25)
                MainApplication.gEntry.grid(column=2, columnspan=4, row=2, padx=5, pady=5, sticky='W')

            elif field == 'ADP Store Code':
                self.ADPStoreCodeVariable = tk.StringVar()
                self.sLab = tk.Label(self.frame, text="ADP Store Code: ")
                self.sLab.grid(column=0, row=3, padx=5, pady=5, sticky='W')
                self.sEntry = tk.Entry(self.frame, textvariable=self.ADPStoreCodeVariable, width=30)
                self.sEntry.grid(column=2, columnspan=4, row=3, padx=5, pady=5, sticky='W')

            elif field == 'Pay Cycle':
                self.PayCycleVariable = tk.StringVar()
                self.SQLCommand = ("SELECT LTRIM(RTRIM(RIGHT(BusinessCalendarName, 2 ))) as CycleNumber "
                                   "FROM [NBO_TRAIN].[dbo].[BusinessCalendars] "
                                   "WHERE BusinessCalendarName LIKE 'Payroll Cycle%'")
                MainApplication.cursor.execute(self.SQLCommand)
                self.payCycles = MainApplication.cursor.fetchall()

                self.cLab = tk.Label(self.frame, text="Pay Cycle: ")
                self.cLab.grid(column=0, row=4, padx=5, pady=5, sticky='W')
                self.cEntry = ttk.Combobox(self.frame, textvariable=self.PayCycleVariable, values=self.payCycles,
                                           state="readonly")
                self.cEntry.grid(column=2, columnspan=4, row=4, padx=5, pady=5, sticky='W')

            elif field == 'Tip Share':
                self.TipShareVariable = tk.StringVar()
                self.tLab = tk.Label(self.frame, text="Tip Share: ")
                self.tLab.grid(column=0, row=5, padx=5, pady=5, sticky='W')
                self.tEntry = ttk.Combobox(self.frame, textvariable=self.TipShareVariable,
                                           values=("No Tip Share", "Landrys Tip Share", "NBO Tip Share"))
                self.tEntry.grid(column=2, columnspan=4, row=5, padx=5, pady=5, sticky='W')

            elif field == 'Begin Date':
                self.bmonthVariable = tk.StringVar()
                self.bdayVariable = tk.StringVar()
                self.byearVariable = tk.StringVar()
                self.bLab = tk.Label(self.frame, text="Begin Date: ")
                self.bLab.grid(column=0, row=6, padx=5, pady=5, sticky='W')
                self.bmonthEntry = ttk.Combobox(self.frame, textvariable=self.bmonthVariable, values=self.months,
                                                width=4)
                self.bmonthEntry.grid(column=2, row=6, padx=5, pady=5, sticky='W')
                self.bdayEntry = ttk.Combobox(self.frame, textvariable=self.bdayVariable, values=self.days, width=3)
                self.bdayEntry.grid(column=3, row=6, pady=5, sticky='W')
                self.byearEntry = ttk.Combobox(self.frame, textvariable=self.byearVariable, values=self.years, width=5)
                self.byearEntry.grid(column=4, row=6, padx=5, pady=5, sticky='W')

            elif field == 'End Date':
                self.emonthVariable = tk.StringVar()
                self.edayVariable = tk.StringVar()
                self.eyearVariable = tk.StringVar()
                self.eLab = tk.Label(self.frame, text="End Date: ")
                self.eLab.grid(column=0, row=7, padx=5, pady=5, sticky='W')
                self.emonthEntry = ttk.Combobox(self.frame, textvariable=self.emonthVariable, values=self.months,
                                                width=4)
                self.emonthEntry.grid(column=2, row=7, padx=5, pady=5, sticky='W')
                self.edayEntry = ttk.Combobox(self.frame, textvariable=self.edayVariable, values=self.days, width=3)
                self.edayEntry.grid(column=3, row=7, pady=5, sticky='W')
                self.eyearEntry = ttk.Combobox(self.frame, textvariable=self.eyearVariable, values=self.years, width=5)
                self.eyearEntry.grid(column=4, row=7, padx=5, pady=5, sticky='W')

        self.tableButton = tk.Button(self.frame, text="View Table", command=self.newTableWindow)
        self.tableButton.grid(column=0, columnspan=2, row=8, padx=5, pady=5, sticky='E')
        self.addButton = tk.Button(self.frame, text="Add Pay Group", command=self.newAddWindow)
        self.addButton.grid(column=2, columnspan=2, row=8, padx=5, pady=5)
        self.editButton = tk.Button(self.frame, text="Edit Pay Group", command=self.newEditWindow)
        self.editButton.grid(column=4, row=8, padx=5, pady=5)
        self.submitButton = tk.Button(self.frame, text="Submit", command=self.submit)
        self.submitButton.grid(column=5, row=8, padx=5, pady=5, sticky='E')
        self.cancelButton = tk.Button(self.frame, text="Close", command=self.master.destroy)
        self.cancelButton.grid(column=0, row=8, padx=5, pady=5, sticky='W')

    def locationSelect(self, oth):
        # check current location selection
        location = self.siteNames[self.lEntry.current()]
        location = str(location).replace("'", "%")

        # clear current values from entries
        MainApplication.gEntry.set('')
        self.sEntry.delete(0, 'end')
        self.cEntry.set('')
        self.tEntry.set('')
        self.bmonthEntry.set('')
        self.bdayEntry.set('')
        self.byearEntry.set('')
        self.emonthEntry.set('')
        self.edayEntry.set('')
        self.eyearEntry.set('')

        # find corresponding site number to name
        self.SQLCommand = ("select POSLabor.dbo.tbl_UnitDescription.MLSiteNumber, POSLabor.dbo.NBO_Sites.SiteNumber "
                           "from POSLabor.dbo.tbl_UnitDescription "
                           "full join POSLabor.dbo.NBO_Sites "
                           "ON POSLabor.dbo.tbl_UnitDescription.strAcctNbr=POSLabor.dbo.NBO_Sites.SiteID "
                           "WHERE POSLabor.dbo.tbl_UnitDescription.strLocationReportHeading "
                           "like '%"+location+"%' "
                           "OR POSLabor.dbo.NBO_Sites.SiteName "
                           "like '%"+location+"%' ")
        MainApplication.cursor.execute(self.SQLCommand)
        self.siteNumbers = MainApplication.cursor.fetchall()
        for row in self.siteNumbers:
            if row[0] is None:
                self.siteNumber = row[1]
            else:
                self.siteNumber = row[0]

        # take in values based on site number
        self.SQLCommand = ("SELECT * FROM POSLabor.dbo.PayrollPayCycle where SiteNumber = "+str(self.siteNumber))
        MainApplication.cursor.execute(self.SQLCommand)

        # fill entries with values
        for row in MainApplication.cursor:
            # row[0] - iLocationID
            # row[1] - SiteNumber
            # row[2] - NBO_Payroll
            # row[3] - ADPStoreCode -------> sEntry
            # row[4] - PayCycle -----------> cEntry
            # row[5] - NBOCalendarNumber
            # row[6] - PayGroupID ---------> translate to PayGroupName for gEntry
            # row[7] - TipShareLocation ---> translate to TipShareName for tEntry
            # row[8] - BeginDate ----------> bmonthEntry, bdayEntry, byearEntry
            # row[9] - EndDate ------------> emonthEntry, edayEntry, eyearEntry

            # 1 - PayGroupID: translate ID to name and match with combobox option
            self.SQLCommand = ("SELECT [PayrollGroupName] FROM [POSLabor].[dbo].[NBO_PayGroup] where [PayGroupID] = "+str(row[6]))
            MainApplication.cursor.execute(self.SQLCommand)
            groupName = str(MainApplication.cursor.fetchall()).strip("[(\',)]")
            MainApplication.gEntry.set(groupName)

            # 2 - ADPStoreCode: fill sEntry with string
            self.sEntry.insert(0, str(row[3]))

            # 3 - Pay Cycle: fill cEntry with integer
            self.cEntry.set(row[4])

            # 4 - TipShareLocation: translate to string to apply to tEntry
            if int(row[7]) == 0:
                self.tEntry.set("No Tip Share")
            elif int(row[7]) == 1:
                self.tEntry.set("Landrys Tip Share")
            elif int(row[7]) == 2:
                self.tEntry.set("NBO Tip Share")

            # 5 - BeginDate
            self.bmonthEntry.set(self.months[row[8].month - 1])
            self.bdayEntry.set(row[8].day)
            self.byearEntry.set(row[8].year)

            # 5 - EndDate
            self.emonthEntry.set(self.months[row[9].month - 1])
            self.edayEntry.set(row[9].day)
            self.eyearEntry.set(row[9].year)

    def newTableWindow(self):
        # region opens Table Window
        self.newWindow = tk.Toplevel(self.master)
        self.app = tableWindow(self.newWindow)
        # endregion opens Table Window

    def newAddWindow(self):
        # region opens Add Window
        self.newWindow = tk.Toplevel(self.master)
        self.app = addWindow(self.newWindow)
        # endregion opens Add Window

    def newEditWindow(self):
        # region opens Edit Window
        self.newWindow = tk.Toplevel(self.master)
        self.app = editWindow(self.newWindow)
        # endregion region opens Edit Window

    def submit(self):
        # iLocationID - take from tbl_UnitDescription or NBO_Sites depending on where siteNumber came from
        location = self.LocationVariable.get()
        self.SQLCommand = ("select POSLabor.dbo.tbl_UnitDescription.iLocationID, POSLabor.dbo.NBO_Sites.SiteID "
                           "from POSLabor.dbo.tbl_UnitDescription "
                           "full join POSLabor.dbo.NBO_Sites "
                           "ON POSLabor.dbo.tbl_UnitDescription.strAcctNbr=POSLabor.dbo.NBO_Sites.SiteID "
                           "WHERE POSLabor.dbo.tbl_UnitDescription.strLocationReportHeading "
                           "like '%" + location + "%' "
                           "OR POSLabor.dbo.NBO_Sites.SiteName "
                           "like '%" + location + "%' ")
        MainApplication.cursor.execute(self.SQLCommand)
        self.ids = MainApplication.cursor.fetchall()
        for row in self.ids:
            if row[0] is None:
                self.id = int(row[1])
            else:
                self.id = int(row[0])

        # SiteNumber - self.siteNumber
        location = self.LocationVariable.get()
        self.SQLCommand = ("select POSLabor.dbo.tbl_UnitDescription.MLSiteNumber, POSLabor.dbo.NBO_Sites.SiteNumber "
                           "from POSLabor.dbo.tbl_UnitDescription "
                           "full join POSLabor.dbo.NBO_Sites "
                           "ON POSLabor.dbo.tbl_UnitDescription.strAcctNbr=POSLabor.dbo.NBO_Sites.SiteID "
                           "WHERE POSLabor.dbo.tbl_UnitDescription.strLocationReportHeading "
                           "like '%" + location + "%' "
                           "OR POSLabor.dbo.NBO_Sites.SiteName "
                           "like '%" + location + "%' ")
        MainApplication.cursor.execute(self.SQLCommand)
        self.siteNumbers = MainApplication.cursor.fetchall()
        for row in self.siteNumbers:
            if row[0] is None:
                self.siteNumber = row[1]
            else:
                self.siteNumber = row[0]

        # NBO_Payroll - if there is tipshare, set to 1
        self.nbo = 0
        if self.TipShareVariable.get()!= "No Tip Share":
            self.nbo = 1

        # NBOCalendarNumber - translate based on PayCycle
        self.pc = 0
        self.calNum = 0
        if self.PayCycleVariable.get() != '':
            self.pc = int(self.PayCycleVariable.get())
        if self.pc == 1:
            self.calNum = 1693843
        elif self.pc == 2:
            self.calNum = 1693844
        elif self.pc == 3:
            self.calNum = 1693845
        elif self.pc == 4:
            self.calNum = 1837816
        elif self.pc == 5:
            self.calNum = 1908266

        # PayGroupID - translate from MainApplication.PayGroupVariable.get()
        self.SQLCommand = "select PayGroupID from [POSLabor].[dbo].[NBO_PayGroup] " \
                          "where PayrollGroupName = '"+MainApplication.PayGroupVariable.get()+"'"
        MainApplication.cursor.execute(self.SQLCommand)
        self.pgID = MainApplication.cursor.fetchall()
        if self.pgID == []:
            self.payGroupID = "None"
        else:
            self.payGroupID = self.pgID[0][0]

        # TipShareLocation - translate from self.TipShareVariable.get()
        self.tsl = 0
        if self.TipShareVariable.get() == "Landrys Tip Share":
            self.tsl = 1
        elif self.TipShareVariable.get() == "NBO Tip Share":
            self.tsl = 2

        # Dates Conversion and Check
        dateError = False;
        max = datetime(month=6, day=7, year=2079)  # maximum date of June 6, 2079

        # BeginDate - import, check dates, then format into datetime
        self.bdate = ""
        if self.bmonthVariable.get() != '':
            self.bmonth = self.monthsDict[self.bmonthVariable.get()]
            self.bday = "%02d" % int(self.bdayVariable.get())
            self.byear = self.byearVariable.get()

            self.bdatetime = datetime(month=int(self.bmonth), day=int(self.bday), year=int(self.byear))
            if self.bdatetime > max:
                dateError = True;
                self.mBox = tk.messagebox.showinfo("Error!", "Invalid Begin Date")
            else:
                self.bdate = "{}-{}-{} 00:00:00".format(self.byear, self.bmonth, self.bday)

        # EndDate - import, check dates, then format into datetime
        self.edate = ""
        if self.emonthVariable.get() != '':
            self.emonth = self.monthsDict[self.emonthVariable.get()]
            self.eday = "%02d" % int(self.edayVariable.get())
            self.eyear = self.eyearVariable.get()

            self.edatetime = datetime(month=int(self.emonth), day=int(self.eday), year=int(self.eyear))
            if self.edatetime > max:
                dateError = True;
                self.mBox = tk.messagebox.showinfo("Error!", "Invalid End Date")
            elif self.edatetime < self.bdatetime:
                dateError = True;
                self.mBox = tk.messagebox.showinfo("Error!", "Invalid Date Range")
            else:
                self.edate = "{}-{}-{} 00:00:00".format(self.eyear, self.emonth, self.eday)

        if dateError == False:
            self.executeProcedure(self.id, self.siteNumber, self.nbo, self.ADPStoreCodeVariable.get(), self.PayCycleVariable.get(),
                                  self.calNum, self.payGroupID, self.tsl, self.bdate, self.edate)

        return 0

    def executeProcedure(self, id, siteNum, nbo, adp, pcv, calNum, pgID, tsl, bdate, edate):
        print("iLocationID:", id)
        print("Site Number:", siteNum)
        print("NBO_Payroll:", nbo)
        print("ADPStoreCode:", adp)
        print("PayCycle:", pcv)
        print("NBOCalendarNumber:", calNum)
        print("PayGroupID:", pgID)
        print("TipShareLocation:", tsl)
        print("BeginDate:", bdate)
        print("EndDate:", edate)

        self.SQLCommand = ("DECLARE @RT INT EXECUTE @RT = POSLabor.dbo.pr_NBO_PayrollPayCycle " + str(id) + ", " +
                           str(siteNum) + ", " + str(nbo) + ", '" + str(adp) + "', " + str(pcv) + ", " + str(calNum) +
                           ", " + str(pgID) + ", " + str(tsl) + ", '" + str(bdate) + "', '" + str(edate)+"'")
        try:
            # region Statement submittion

            print(self.SQLCommand)
            MainApplication.cursor.execute(self.SQLCommand)
            MainApplication.connection.commit()
            self.mBox = tk.messagebox.showinfo("Success!","Import Complete")
            # endregion Statement submittion

        except:
            # region submition error
            self.mBox = tk.messagebox.showinfo("Error!","Import Failed")
            print("Failed command:", self.SQLCommand)
            MainApplication.connection.rollback()
            # endregion submition error

        self.master.destroy()

        return 0



class tableWindow:
    def __init__(self, master):
        # region constructor
        self.master = master
        self.frame = tk.Frame(self.master)
        self.master.wm_title("View Table")
        self.createWidgets()
        self.frame.grid(column=0, row=0)
        # endregion constructor

    def createWidgets(self):
        self.table = ttk.Treeview(self.frame)
        self.table.grid(column=1, columnspan=10, row=1, rowspan=5, padx=(5,0), pady=5)
        self.table['columns'] = ('siteNum', 'nbo', 'adp', 'payCycle', 'nboCal', 'pgID', 'tipShare', 'begin', 'end')
        self.table.heading('#0', text="iLocationID")
        self.table.column('#0', width=75)
        self.table.heading('siteNum', text="Site Number")
        self.table.column('siteNum', width=100)
        self.table.heading('nbo', text="NBO_Payroll")
        self.table.column('nbo', width=75)
        self.table.heading('adp', text="ADP Store Code")
        self.table.column('adp', width=100)
        self.table.heading('payCycle', text="PayCycle")
        self.table.column('payCycle', width=75)
        self.table.heading('nboCal', text="NBOCalendarNumber")
        self.table.column('nboCal', width=125)
        self.table.heading('pgID', text="PayGroupID")
        self.table.column('pgID', width=75)
        self.table.heading('tipShare', text="Tip Share")
        self.table.column('tipShare', width=75)
        self.table.heading('begin', text="BeginDate")
        self.table.column('begin', width=125)
        self.table.heading('end', text="EndDate")
        self.table.column('end', width=125)

        self.SQLCommand = ("select * from POSLabor.dbo.PayrollPayCycle")
        MainApplication.cursor.execute(self.SQLCommand)

        for row in MainApplication.cursor:
            self.table.insert('', 'end', text=str(row[0]), values=(row[1], row[2], row[3], row[4], row[5],
                                                                   row[6], row[7], row[8], row[9]))

        sb = ttk.Scrollbar(self.frame, orient='vertical', command=self.table.yview)
        self.table.configure(yscroll=sb.set)
        sb.grid(row=1, column=11, rowspan=5, sticky="NS", pady=5)

        self.exportButton = tk.Button(self.frame, text="Export...", command=self.export)
        self.exportButton.grid(column=3, columnspan=3, row=6, padx=5, pady=5, sticky='E')
        self.closeButton = tk.Button(self.frame, text="Close", command=self.close_windows)
        self.closeButton.grid(column=6, columnspan=3, row=6, padx=5, pady=5, sticky='W')

    def export(self):
        # prepare text file
        ftypes = [('Text file', '.txt'), ('All files', '*')]
        name = filedialog.asksaveasfilename(filetypes=ftypes, defaultextension=".txt")
        outFile = open(name, 'w')
        # prepare sql query
        self.SQLCommand = ("select * from POSLabor.dbo.PayrollPayCycle")
        MainApplication.cursor.execute(self.SQLCommand)
        # fill text file
        for row in MainApplication.cursor:
            outFile.write(str(row[0])+"\t"+str(row[1])+"\t"+str(row[2])+"\t"+str(row[3])+"\t"+str(row[4])+"\t"+
                          str(row[5])+"\t"+str(row[6])+"\t"+str(row[7])+"\t"+str(row[8])+"\t"+str(row[9])+"\n")
        outFile.close()

    def close_windows(self):
        self.master.destroy()

class addWindow:
    def __init__(self, master):
        # region constructor
        self.master = master
        self.frame = tk.Frame(self.master)
        self.master.wm_title("Add Pay Group")
        self.newPayGroupName = tk.StringVar()
        self.createWidgets()
        self.frame.grid(column=0, row=0)
        # endregion constructor

    def createWidgets(self):
        # region Label/Entry creation
        self.lab = tk.Label(self.frame, text="New Pay Group Name")
        self.lab.grid(column=1, columnspan = 4, row=1, padx=5, pady=5, sticky='W')
        self.entry = tk.Entry(self.frame, textvariable=self.newPayGroupName, width = 39)
        self.entry.grid(column=1, columnspan=4, row=2, padx=5, pady=5, sticky='W')
        # endregion Label/Entry creation

        # region Button creation
        self.cancelButton = tk.Button(self.frame, text="Cancel", command=self.close_windows)
        self.cancelButton.grid(column=1, row=3, padx=5, pady=5, sticky='W')
        self.submitButton = tk.Button(self.frame, text="Submit", command=self.submit)
        self.submitButton.grid(column=4, row=3, padx=5, pady=5, sticky='E')
        # endregion Button creation

    def submit(self):
        # region name validation
        self.flag = False
        for name in MainApplication.payGroups:
            if str(self.newPayGroupName.get()) == name:
                self.mBox = tk.messagebox.showinfo("Error!","Name Already in Use")
                self.flag = True
        # endregion name validation

        # region SQL submittion
        if self.flag == False:
            MainApplication.payGroups.insert((len(MainApplication.payGroups) + 1), self.newPayGroupName.get())
            self.SQLCommand = ("INSERT INTO [POSLabor].[dbo].[NBO_PayGroup] (PayrollGroupName, PayGroupID) " \
                          "VALUES ('" + str(self.newPayGroupName.get()) + "', " + str(len(MainApplication.payGroups)) + " );")
            MainApplication.cursor.execute(self.SQLCommand)
            MainApplication.connection.commit()

            # region resetting menus
            MainApplication.payGroups.insert(0,self.newPayGroupName)
            MainApplication.gEntry['values'] = MainApplication.payGroups
            MainApplication.PayGroupVariable.set(self.newPayGroupName.get())
            # endregion resetting menus

        # endregion SQL submittion

        self.master.destroy()

    def close_windows(self):
        self.master.destroy()

class editWindow:
    def __init__(self, master):
        # region constructor
        self.master = master
        self.frame = tk.Frame(self.master)
        self.master.wm_title("Edit Pay Group")
        self.createWidgets()
        self.frame.grid(column=0, row=0)
        # endregion constructor

    def createWidgets(self):
        # region Label/Entry creation

        self.lab1 = tk.Label(self.frame, text="Pay Group to Edit:")
        self.lab1.grid(column=1, columnspan=5, row=1, sticky='W', padx=5, pady=5)

        ## nameList
        self.nameList = tk.Listbox(self.frame, width=25)
        self.count = 0
        for name in MainApplication.payGroups:
            name = str(name).strip("(,)")
            self.nameList.insert(self.count, name)
            self.count = self.count + 1
        self.nameList.grid(row=2, rowspan=6, column=1, padx=(5,0), pady=5)
        sb = ttk.Scrollbar(self.frame, orient='vertical', command=self.nameList.yview)
        self.nameList.configure(yscroll=sb.set)
        sb.grid(row=2, column=2, rowspan=6, sticky="NS", pady=5)
        self.lab2 = tk.Label(self.frame, text="New Pay Group Name: ", anchor='w', width=30)
        self.lab2.grid(row=3, column=3, columnspan=4, padx=10, pady=5, sticky='S')
        self.NewPayGroupName = tk.StringVar()
        self.entry = tk.Entry(self.frame, textvariable=self.NewPayGroupName, width=30)
        self.entry.grid(row=4, column=3, columnspan=4, padx=10, sticky='NW')
        self.entry.grid_columnconfigure(3, weight=1)
        self.submitButton = tk.Button(self.frame, text="Submit Edit", command=lambda: self.submit(self.nameList.get(tk.ACTIVE), self.NewPayGroupName))
        self.submitButton.grid(column=6, columnspan=2, row=7, padx=5, pady=5, sticky='SE')
        self.cancelButton = tk.Button(self.frame, text="Cancel", command=self.close_windows)
        self.cancelButton.grid(column=4, columnspan=2, row=7, padx=(5,0), pady=5, sticky='SE')
        # endregion Label/Entry creation

    def submit(self, OldPayGroupName, NewPayGroupName):
        # region name validation
        self.flag = False
        for name in MainApplication.payGroups:
            if str(NewPayGroupName.get()) in name:
                self.mBox = tk.messagebox.showinfo("Error!","Name Already in Use")
                self.flag = True
        # endregion name validation

        # region SQL submittion
        if self.flag == False:
            # use sql statement to UPDATE values to new values
            OldPayGroupName = str(OldPayGroupName).replace("'", "%")
            self.SQLCommand = ("UPDATE [POSLabor].[dbo].[NBO_PayGroup] " \
                          "SET [PayrollGroupName]='" + str(NewPayGroupName.get()) + "' " + \
                          "WHERE [PayrollGroupName] like '" + OldPayGroupName + "';")
            MainApplication.cursor.execute(self.SQLCommand)
            MainApplication.connection.commit()

            # region resetting menus
            OldPayGroupName = OldPayGroupName.strip('%')
            for index, item in enumerate(MainApplication.payGroups):
                if (OldPayGroupName in item):
                    MainApplication.payGroups[index] = str(NewPayGroupName.get())
                    MainApplication.gEntry.set(MainApplication.payGroups[index])
            MainApplication.gEntry['values'] = MainApplication.payGroups
            MainApplication.PayGroupVariable.set(NewPayGroupName.get())
            # endregion resetting menus

            self.master.destroy()

        # endregion SQL submittion

    def close_windows(self):
        self.master.destroy()

def main():
    # region login
    # opens config file
    Config = configparser.ConfigParser()
    Config.read("config.ini")

    # reading base64 login from config.ini
    driver =    Config.get("Login","Driver")
    server =    Config.get("Login","Server")
    database =  Config.get("Login","Database")
    uid =       Config.get("Login","uid")
    pwd =       Config.get("Login","pwd")

    # decoding credentials
    driver =    str(base64.b64decode(driver)).strip('b\'')
    server =    str(base64.b64decode(server)).strip('b\'')
    ind =       server.index("\\\\")
    server =    server[:ind] + server[ind+1:]
    database =  str(base64.b64decode(database)).strip('b\'')
    uid =       str(base64.b64decode(uid)).strip('b\'')
    pwd =       str(base64.b64decode(pwd)).strip('b\'')

    login =     ("Driver=%s;Server=%s;Database=%s;uid=%s;pwd=%s" % (driver, server, database, uid, pwd))

    connection = pypyodbc.connect(login)
    # endregion login

    # region open tkinter window
    root = tk.Tk()
    app = MainApplication(root, connection, server)
    root.mainloop()
    connection.close()
    # endregion open tkinter window

if __name__ == '__main__':
    main()