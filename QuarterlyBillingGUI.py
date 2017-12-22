import pandas as pd
import tkinter as tk
LARGE_FONT = ('Verdana', 24)
from tkinter import messagebox
import tkinter.ttk as ttk
import matplotlib
matplotlib.use('TkAgg')
from pylab import *
from PIL import Image, ImageTk
from tkinter import filedialog
import datetime
from datetime import datetime
from io import StringIO



"""Unless Artwork is disabled this program needs to contain in the same folder the BWMIcon.ico and StartPageBackground.jpg files.
If not there will be an error. Artwork can be disabled by removing lines 67-71"""


### This Class Drives the Graphical User Interface (GUI)
class BWMCommandCenterApp(tk.Tk):

    def __init__(self, *args, **kwargs):

        tk.Tk.__init__(self, *args, **kwargs)

        #tk.Tk.iconbitmap(self, default='BWMIcon.ico')
        tk.Tk.wm_title(self, "BWM Quarterly Billing Application")
        container = tk.Frame(self)

        container.pack(side='top', fill='both', expand = True)

        container.grid_rowconfigure(0, weight = 1)
        container.grid_columnconfigure(0, weight = 1)

        self.shared_data = {
            "holdings_file": '',
            "billing_file": '',
            "transactions_file": '',
            "portaccount_file": '',
            "cashavailable_file": '',
        }


        self.frames = {}

        # All pages need to be listed here
        for F in (StartPage, QuarterlyBilling, FileSelection):

            frame = F(container, self)

            self.frames[F] = frame

            frame.grid(row=0, column=0, sticky='nsew')

        self.show_frame(StartPage)


    def show_frame(self, cont):

        frame = self.frames[cont]
        frame.tkraise()

    def get_page(self, page_class):
        return self.frames[page_class]


### This class is used for the start page where the quarterly billing program is located.
class StartPage(tk.Frame):

    def __init__(self, parent, controller):

        tk.Frame.__init__(self, parent)

        # Adds the artwork to the GUI. Can be stripped out if desired by deleting lines 67-71.
        photo = Image.open("StartPageBackground.jpg")
        self.background_image = ImageTk.PhotoImage(photo)
        background_label = tk.Label(self, image=self.background_image)
        background_label.image = photo
        background_label.place(x=0, y=0, relwidth=1, relheight=1)


        buttonFrame = tk.Frame(self, width = 150, height = 150)
        button = tk.Button(buttonFrame, text='Quarterly Billing',
                           command = lambda: controller.show_frame(QuarterlyBilling), bg = '#0b0ea3', foreground = 'white', relief = tk.GROOVE)
        buttonFrame.grid_propagate(False)
        buttonFrame.columnconfigure(0, weight = 1)
        buttonFrame.rowconfigure(0, weight = 1)
        buttonFrame.place(x=-350, y = 0, relx = .5, relheight = .2)
        button.grid(sticky='wens')

class FileSelection(tk.Frame):

    def __init__(self,parent,controller):
        tk.Frame.__init__(self,parent)
        self.controller = controller

        ### Dictionary that stores information for the files selected. File paths are copied to the shared_data dictionary in the controller.
        files_needed = {'Holdings': ["holdings_file",''],
                        'Billing': ["billing_file",''],
                        'Transactions': ["transactions_file",''],
                        'PortAccount#': ["portaccount_file",''],
                        'Cash Available': ["cashavailable_file",''],
                        }

        ### Function to browse for the appropriate files.
        def file_entry(file,text_variable,entry_gui):
            entry_gui.delete(0,tk.END)
            text_variable = filedialog.askopenfilename(initialdir = "/", title = "Select {} File".format(file.split('_')[0].title()))
            entry_gui.insert(tk.END,text_variable)
            self.controller.shared_data[file] = text_variable

        ### Loop that creates the GUI fields for file selection
        row = 2
        for i in files_needed:
            files_needed[i].append(tk.Label(self,text = "What is the name of the {} File: ".format(i)))
            files_needed[i][2].grid(row = row, column = 0)
            files_needed[i].append(tk.Entry(self,textvariable=files_needed[i][0],width=60))
            files_needed[i][3].grid(row = row, column = 1)
            files_needed[i].append(tk.Button(self,text='Browse',command=lambda i=i: file_entry(files_needed[i][0],files_needed[i][1],files_needed[i][3]),width=12,bd=3))
            files_needed[i][4].grid(row = row, column = 2)
            row += 1

        ### Function that makes sure you've selected a file for all 6 required files.
        def submit_check(listerthing):
            error = 0
            for i in listerthing:
                if len(listerthing[i]) < 1:
                    messagebox.showwarning("Error", "Please make sure to enter in a file for all 6 required files")
                    error +=1
                    break
                else:
                    pass
            if error == 0:
                controller.show_frame(QuarterlyBilling)

        ### Submit button that runs the submit_check function and then takes you to run billing
        submit_button = tk.Button(self,text = 'Submit', command = lambda: submit_check(self.controller.shared_data),width=12,bd=3)
        submit_button.grid(row = row, column = 2, sticky = tk.E)
        submit_button.config(bd = 5)

        ### Button that takes you back to run billing page.
        backButton = ttk.Button(self, text="Back to Billing", width = 17, command=lambda: controller.show_frame(QuarterlyBilling))
        backButton.grid(row = 0)

        ### Directions
        directionsLabel = tk.Label(self, text = "Directions:", font = ("Times", 25, "bold", "underline"))
        directionsLabel.grid(row=row+1)
        directionsFrame = tk.Frame(self)
        directionsFrame.grid(row=row+2,columnspan=4)
        directionsFrame.columnconfigure(0, weight = 1)
        directionsFrame.rowconfigure(0, weight = 1)
        directions = tk.Text(directionsFrame,wrap='word')
        directions.grid(row = 0, column = 0, sticky='nsew')
        scroll = tk.Scrollbar(directionsFrame, command = directions.yview)
        scroll.grid(row = 0, column = 1, sticky = 'nsew')
        directions['yscrollcommand'] = scroll.set


        directions.insert(tk.END,"""The files necessary for billing can be found in the following location:\n\nHoldings file: Black Diamond Rebalancer Report. Find this file by clicking on the Rebalancer Report under the Skypages tab. Run the report on the All Portfolios Select Level, and download the report to CSV. \n\nBilling file: Black Diamond Run Billing Report. To access the file, Run Client Billing in Setup. Make sure to run billing for the Billing Family Accounts only advisor. Export the csv file and if given the option of type to export, make sure to export the Black Diamond Export.\n\nTransactions file: Black Diamond Transactions report. Find this file by clicking on Transactions under the Skypages tab. Change the date range to the last 30 days and run the report. You can then export the csv file. IMPORTANT: You need to open the Transactions CSV in Excel and re-save it as a csv to fix an error with how the csv is structured directly from Black Diamond.\n\nPortAccount file: Black Diamond Data Mining. Find this file by clicking on Data Mining under the Skypages tab. select the account number and portfolio name options, then process the data mining. export the csv.\n\nCashavailable file: Wealth Scape All Accounts Summary Page. Export the csv for the All Accounts summary page. Make sure your summary page includes the account number column and the cash available to withdraw column.\n""")

        directions.config(state='disabled')

### Quarterly Billing Algorithm. CSV handling is done using Pandas DataFrames.
class QuarterlyBilling(tk.Frame):

    ### Initiliaze the GUI page for quarterly billing.
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        self.printoutInfo = []

        ### This function runs the quarterly billing algorithm.
        def runBilling():

            output.delete("0.0",tk.END)
            #try:
            hd = pd.read_csv(self.controller.shared_data["holdings_file"])
            tb = pd.read_csv(self.controller.shared_data["billing_file"], thousands = ',')
            tt = pd.read_csv(self.controller.shared_data["transactions_file"])
            pa = pd.read_csv(self.controller.shared_data["portaccount_file"])
            ca = pd.read_csv(self.controller.shared_data["cashavailable_file"])

            #except:
            #             messagebox.showwarning("Error","Make sure to input the files in the input files section before running billing.")

            try:

                ### This for loop prints out the files used for each required input. If you accidently import the holdings file for billing you will be able to see this.
                for value in self.controller.shared_data:
                    output.insert(tk.END, '{}: {}\n'.format(value,self.controller.shared_data[value]))
                    output.update()
                output.insert(tk.END, '\n')

                ### Rename columns in PortAccount#, Billing, and CashAvailable to make the merge function possible.
                pa.rename(columns={'Portfolio Name':'Portfolio','Account Number': 'AccountNumber'},inplace=True)
                tb.rename(columns={'Acct Number': 'AccountNumber'},inplace=True)
                ca.rename(columns={'Cash Avail. to Withdraw': 'WcCashAvailable','Account #': 'AccountNumber'},inplace=True)
                ### Fill NaN values and change the column names for Trasactions DataFrame.
                tt.fillna('',inplace=True)
                try:
                    tt.columns = tt.loc[1]
                    ### Convert the Market Value from a string to a float.
                    tt['MarketValue'] = pd.to_numeric(tt['MarketValue'],errors='coerce')
                except:
                    try:
                        tt.columns=tt.loc[0]
                        ### Convert the Market Value from a string to a float.
                        tt['MarketValue'] = pd.to_numeric(tt['MarketValue'],errors='coerce')
                    except:
                        output.insert(tk.END,'There was an issue with the Transactions csv. Make sure you opened the csv export from Black Diamond in Excel and re-saved it as a new csv file to fix a format error from the export Black Diamond gives you. Re-import the fixed file and re-run billing.')
                        return

                ### Drop the Portfolios that do not have holdings at the bottom of the Holding Dataframe.
                hd = hd[0:len(hd['Symbol'].dropna())]
                ca = ca[0:len(ca['WcCashAvailable'].dropna())]

                ### Convert Fee from a string to a float. $ needs to be removed.
                tb['Fee'] = tb['Fee'].apply(lambda x: x.replace('$','').replace(',',''))
                tb['Fee'] = pd.to_numeric(tb['Fee'],errors='coerce')

                ### Convert Cash Available from a string to a float. $ needs to be removed.
                ca['WcCashAvailable'] = ca['WcCashAvailable'].apply(lambda x: x.replace('$','').replace(',','').replace('--',''))
                ca['WcCashAvailable'] = pd.to_numeric(ca['WcCashAvailable'],errors='raise')




                """ Create the Transaction DataFrame that only includes buy orders greater than 999. Sum it to double Check with
                 manually doing the Transaction sheet in Excel."""

                tt = tt[(tt['Action']=='Buy') & (tt['MarketValue'][3:]>999)]
                sum(tt['MarketValue'])

                ### Create the Yes value for the RoundTrip column

                tt['RoundTrip'] = 'Yes'

                ### This removes the 403(b) accounts which are longer than 9 characters
                tt = tt[tt['AccountNumber'].apply(len)<10]

                ### Prep the Transactions DataFrame for the merge with holdings. Keep only relevant columns

                tt = tt[['AccountNumber','Ticker','RoundTrip']]

                ### Rename the columns to match with the holdings excel sheet

                tt.columns = ['AccountNumber','Symbol','RoundTrip']

                ### Remove duplicates from the Transaction dataframe. Since only 1 roundtrip matters we don't need duplicates.
                tt.drop_duplicates(inplace=True)

                ### Remove trailing whitespace from the Portfolio column and then merge with Data mined Portfolio and Account# CSV

                hd['Portfolio'] = hd.Portfolio.apply(lambda x: str(x).rstrip())
                hda = pd.merge(hd,pa,how='left',on=['Portfolio'])

                ### Merge the holdings DataFrame on account number and symbol with the Transactions Dataframe

                Trades1 = pd.merge(hda,tt,how = 'left',on=['AccountNumber','Symbol'])

                ### Replace the Trades NaN with No

                Trades1.fillna('No',inplace = True)

                ### Remove Duplicate values from Trades1

                Trades1_drop_duplicated = Trades1.drop_duplicates()

                ### This tests to make sure that the roundtrip additions to the Holding DataFrame equal the amount that originated from the Transaction DataFrame
                print('''---Double Checks that the Merged DataFrame with RoundTrips Equals the Original DataFrame with RoundTrip Information---
                ------------------------------------------------------------------------------------------------------------------------''')
                if Trades1_drop_duplicated[Trades1_drop_duplicated['RoundTrip'] == 'Yes'].shape[0] != tt.shape[0]:
                    roundtripLine = '''\nMerged DataFrame has {} instances of RoundTrips and Original Transactions DataFrame has {} instances. There is a difference between the amount of holdings showing roundtrip and the non-duplicated RoundTrip dataframe tt. Total difference = {}\n'''.format(
                        Trades1_drop_duplicated[Trades1_drop_duplicated['RoundTrip'] =='Yes'].shape[0],
                        tt.shape[0], abs(Trades1_drop_duplicated[Trades1_drop_duplicated['RoundTrip'] == 'Yes'].shape[0]-tt.shape[0]))

                else:
                    roundtripLine = '\nMerged DataFrame has {} instances of RoundTrips and Original Transactons DataFrame has {} instances. Everything is good\n'.format(
                        Trades1_drop_duplicated[Trades1_drop_duplicated['RoundTrip'] =='Yes'].shape[0],tt.shape[0])

                self.printoutInfo.append(roundtripLine)

                """Find values that aren't in merged dataframe. This will show the instances where a position was bought into, but potentially sold
                before the round-trip was over or for positions that were converted to advantage class funds (i.e FSEMX to FSEVX)"""

                missing = pd.DataFrame()
                missing['OriginalAccountNumber'] = tt['AccountNumber']
                missing['OriginalSymbol'] = tt['Symbol']
                missing['MergeFrame'] = missing['OriginalAccountNumber'] + missing['OriginalSymbol']
                missing2 = pd.DataFrame()
                missing2 = Trades1_drop_duplicated[Trades1_drop_duplicated['RoundTrip'] == 'Yes'][['AccountNumber','Symbol']]
                missing2['MergeFrame'] = missing2['AccountNumber'] + missing2['Symbol']
                findmissing = pd.merge(missing, missing2, how = 'outer', on = ['MergeFrame'])

                ### Fills Symbol and AccountNumber column in all positions that are not contained in both DataFrames.

                findmissing.fillna('ERROR',inplace = True)
                print(findmissing[findmissing['Symbol'] == 'ERROR'])
                missingRoundTripLine1 = """These Accounts are the accounts missing in merged dataframe that added RoundTrip information.More often than not this is the result of the seed accounts that do not have a portfolio or an account that bought a fund that was either sold during the round-trip or that was converted.Conversions occur for a lower fee fund of the same positions (ie. FSEMX to FSEVX)\n"""
                missingRoundTripLine2 = findmissing[findmissing['Symbol'] == 'ERROR']

                self.printoutInfo.append(missingRoundTripLine1)
                self.printoutInfo.append(missingRoundTripLine2)

                ### Replaces the hyphon in the account numbers that are in the cash available Dataframe.
                ca['AccountNumber'] = ca['AccountNumber'].apply(lambda x: x.replace('-',''))
                ### Merge Short-term shares onto the Dataframe that includes holding and round-trips. This is the final DataFrame for Trade Prep

                FinalTradePrep = Trades1_drop_duplicated.copy()

                ### Replace Shares With Redemtion Fees values of Nan with 0

                ### Reduce columns in CashAvailable to desired columns.
                ca = ca[['AccountNumber','WcCashAvailable']]

                ### Remove Cash from Holdings so that it doesn not impact trades since these holdings would be used as regular mutual funds.
                FinalTradePrep = FinalTradePrep[FinalTradePrep['Level Name'] != 'Cash & Equivalents*']



                ####### PART 1 IS COMPLETE AT THIS POINT. ALL HOLDINGS HAVE BEEN PREPPED TO SHOW SHORT-TERM AND ROUNDTRIP HOLDINGS. ##########

                ### Calculate cumulative fees by billing account.
                cumfees = pd.DataFrame(tb.groupby("BillingAccount")["Fee"].sum())
                cumfees['AccountNumber']=cumfees.index
                cumfees.rename(columns={'Fee':'CumulativeFee'},inplace=True)

                ### Add cumulative fees and WC CashAvailable to the billing DataFrame.
                tb = pd.merge(tb,cumfees,how = 'left',on=['AccountNumber'])
                tb = pd.merge(tb,ca,how='left',on=['AccountNumber'])

                ### Fill all NaN cumulative fee values with 0
                tb.fillna(0,inplace=True)

                ### Create the Fee trade for all accounts. Cumulative Fee minus the cash available from WC Cash Data and then add .5 for rounding errors.
                tb['FeeTrade'] = tb['CumulativeFee'] - tb['WcCashAvailable']+.05

                ### Create A DataFrame for Pay By Invoice clients. Used to double check Trade Total + Invoice Total after trading and for reference. Remove Clients from DataFrame
                PayByInvoice = tb[tb['BillingAccount'] == 'Pay By Invoice']
                PayByInvoice = PayByInvoice[['Portfolio','ScheduleName','AccountNumber','Account Name','BillingAccount','AccountValue','BilledValue','Fee','Billing Group']].copy()
                tb_no_invoice = tb[tb['BillingAccount'] != 'Pay By Invoice']

                ### Create length value to loop through billing dataframe
                tblength = len(tb.index.values)

                ### Add TradeValue to Final Trade Prep DataFrame
                FinalTradePrep['TradeValue']=0

                """ This Loop applies the trade values to each holding that needs to be traded for fees. I first make sure that the cummulative fee is greater than 0 to avoid applying
                trades for accounts with negative fee trades due to the calculation method. I then check to make sure the fee trade is greater than 0. If either of these if functions fail
                make sure the FeeTrade column value = 0. I take the account number from billing and then the feetrade amount calculated earlier. After this I create a new mini DataFrame
                that only contains the information for the accountnumber and then I clean it to remove holdings with redemption fees, targets that have a value of zero since these tend
                to be special holdings. New clients will have waived fees and they should be the only other clients holding positions with a target of 0. To clean for models that account
                for stock holdings on a clients request I have filtered the dataframe to remove symbols that are less than length 5 which denotes mutual fund. Finally I sort the dataframe
                by the difference from target and actual holding values. This places overweight holdings on top. After this I create a variable to test whether we can sell multiple holdings to
                achieve the fee or if we need to sell a large amount of one holding. This then get's processed through some if statements and it creates the final trade instructions."""
                for i in range(0,tblength):
                    try:
                        if tb_no_invoice.loc[i,'CumulativeFee'] > 0:
                            if tb_no_invoice.loc[i,'FeeTrade'] > 0:
                                AccountNumber = tb_no_invoice.loc[i,'AccountNumber']
                                print(AccountNumber,tb_no_invoice.loc[i,'CumulativeFee'],tb_no_invoice.loc[i,'WcCashAvailable'])
                                FeeTrade = tb_no_invoice.loc[i,'FeeTrade']
                                TradingFrame = FinalTradePrep[(FinalTradePrep['AccountNumber'] == AccountNumber) &
                                                              (FinalTradePrep['Target'] > 0) & (FinalTradePrep['Symbol'].apply(len) > 4)].sort_values(by = 'Difference',ascending = False)
                                TradingFrame.sort_values(by = 'Difference',ascending=False,inplace = True)
                                TradingFramelength = len(TradingFrame.index.values)
                                TradingFrameTotal = TradingFramelength*999
                                IndexValues = TradingFrame.index.values
                                if TradingFrameTotal < FeeTrade:
                                    TradingFrame = TradingFrame[TradingFrame['RoundTrip'] != 'Yes']
                                    FinalTradePrep.loc[IndexValues[0], 'TradeValue'] = FeeTrade
                                elif FeeTrade > 999 and TradingFrameTotal > FeeTrade:
                                    count = 0
                                    while FeeTrade > 999:
                                        FinalTradePrep.loc[IndexValues[count], 'TradeValue'] = 999
                                        FeeTrade = FeeTrade - 999
                                        count = count + 1
                                    FinalTradePrep.loc[IndexValues[count], 'TradeValue'] = FeeTrade
                                else:
                                    FinalTradePrep.loc[IndexValues[0],'TradeValue'] = FeeTrade
                            else:
                                tb_no_invoice.loc[i,'FeeTrade'] = 0
                        else:
                            tb_no_invoice.loc[i,'FeeTrade'] = 0
                            pass
                    except Exception as e:
                        print(str(e))
                        try:
                            output.insert(tk.END,"Location {} is associated with a client or account that is to be paid by invoice. Portfolio name is: {}".format(i, PayByInvoice.loc[i,'Portfolio']))
                            output.insert(tk.END,'\n')
                            output.update()
                        except:
                            print("This account had an error {}".format(tb_no_invoice.loc[i, "Portfolio"]))

                ####### PART 2 IS COMPLETE #######
                print('-'*50)
                ###Create New DataFrame that will be used to import trade information to WC.
                WC_Import = FinalTradePrep[FinalTradePrep['TradeValue'] > 0]

                if round(sum(FinalTradePrep['TradeValue']),4) == round(sum(WC_Import['TradeValue']),4):
                    tradeAuditLine = 'Trade values match up: FinalTradePrep = {}, WC_Import = {}'.format(sum(FinalTradePrep['TradeValue']),sum(WC_Import['TradeValue']))
                else:
                    tradeAuditLine = 'Trade values do not match-up: FinalTradePrep = {}, WC_Import = {}'.format(sum(FinalTradePrep['TradeValue']),sum(WC_Import['TradeValue']))


                self.printoutInfo.append(tradeAuditLine)

                ###Changes the Import File to just contain the 3 columns that are needed.
                WC_Import_Final = WC_Import[['AccountNumber','TradeValue','Symbol']]

                ###Fills in the columns with the information that Fidelity uses to create mutual fund trades. All information can be seen in WC. Import/Export > Import > Question Mark in the top right.
                WC_Import_Final['AccountType'] = 1
                WC_Import_Final['OrderAction'] = 'SF'
                WC_Import_Final['PriceType'] = ''
                WC_Import_Final['TimeInForce'] = ''
                WC_Import_Final['StopPrice/LimitPrice'] = ''
                WC_Import_Final['QuantityType'] = 'D'

                ### Creates the Finalized DataFrame with the columns in the correct order.
                WC_Import_Final = WC_Import_Final[['AccountNumber','AccountType','OrderAction','TradeValue','Symbol','PriceType','TimeInForce','StopPrice/LimitPrice','QuantityType']]

                ### Rounds the trade values to 2 decimal places to avoid an error that is created when the value has goes to too many decimal places.
                WC_Import_Final['TradeValue'] = WC_Import_Final['TradeValue'].apply(lambda x: round(x,2))


                ### Prints diagnostic information to the diagnostics text box in the window.
                for i in self.printoutInfo:
                    output.insert(tk.END, i)
                    output.insert(tk.END, '\n')
                    output.update()

                ### Creation of the CSV file that will be uploaded.
                now = datetime.now()
                date = str(now.month) + str(now.day) + str(now.year)
                getDirectoryName = filedialog.askdirectory(title = 'Choose Folder to Save Files')
                PayByInvoice.to_csv(getDirectoryName + "\PaybyInvoiceClients{}.csv".format(date), index = False)
                FinalTradePrep.to_csv(getDirectoryName + '\FinalTradesPrep{}.csv'.format(date), index = False)
                WC_Import_Final.to_csv(getDirectoryName + "\WCImport{}.csv".format(date), index = False, header=False)

            except Exception as e:
                import traceback
                output.insert(tk.END, traceback.format_exc())
                output.update()

                ####### PART 3 IS COMPLETE #######



        ### GUI buttons to run the program.
        backHomeButton = ttk.Button(self, text="Back to Home", width = 17, command=lambda: controller.show_frame(StartPage))
        backHomeButton.pack()
        # importFilesButton = ttk.Button(self, text = "Import Files", command=get_files)
        importFilesButton = ttk.Button(self, text = "Import Files", command=lambda: controller.show_frame(FileSelection))
        importFilesButton.pack()
        runBillingButton = ttk.Button(self, text = "Run Billing", command = runBilling)
        runBillingButton.pack()

        ### Gui code for the diagnostic window where information is displayed.
        outputLabel = tk.Label(self, text = "Output", font = ("Times", 25, "bold", "underline"))
        outputLabel.pack()
        outputFrame = tk.Frame(self)
        outputFrame.pack(fill = 'both', expand = True)
        outputFrame.grid_propagate(False)
        outputFrame.columnconfigure(0, weight = 1)
        outputFrame.rowconfigure(0, weight = 1)
        output = tk.Text(outputFrame,wrap='word')
        output.grid(row = 0, column = 0, sticky='nsew')
        scroll = tk.Scrollbar(outputFrame, command = output.yview)
        scroll.grid(row = 0, column = 1, sticky = 'nsew')
        output['yscrollcommand'] = scroll.set



### Calls the application to run.
app = BWMCommandCenterApp()
app.geometry("1045x620")
app.mainloop()

