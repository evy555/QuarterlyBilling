import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
### Import Files and clean the files


class QuarterlyBilling():

    def __init__(self):
        self.master = tk.Tk()

        def get_files():
            try:
                self.hd = pd.read_csv(filedialog.askopenfilename(initialdir = "/", title = "Select Holdings File"))
                self.tb = pd.read_csv(filedialog.askopenfilename(initialdir = "/", title = "Select Billing File"),thousands = ',')
                self.tt = pd.read_csv(filedialog.askopenfilename(initialdir = "/", title = "Select Transactions File"))
                self.pa = pd.read_csv(filedialog.askopenfilename(initialdir = "/", title = "Select PortAccount# File"))
                self.ca = pd.read_csv(filedialog.askopenfilename(initialdir = "/", title = "Select Cash Available File"))
                self.st = pd.read_csv(filedialog.askopenfilename(initialdir = "/", title = "Select Short Term Shares File"), thousands = ',')

                runBillingButton.configure(state = 'active')
                runBillingButton.update()

            except:
                messagebox.showwarning("Error", "There was an error with one of the input files.")

        def runBilling():
            hd = self.hd
            tb = self.tb
            tt = self.tt
            pa = self.pa
            ca = self.ca

            ### While reading the ShortTerm csv I need
            st = self.st

            ### Rename columns in PortAccount#, ShortTerm, Billing, and CashAvailable to make the merge function possible.
            pa.rename(columns={'Portfolio Name':'Portfolio','Account Number': 'AccountNumber'},inplace=True)
            st.rename(columns={'Account Number': 'AccountNumber'},inplace=True)
            tb.rename(columns={'Acct Number': 'AccountNumber'},inplace=True)
            ca.rename(columns={'Cash Avail. to Withdraw': 'WcCashAvailable','Account #': 'AccountNumber'},inplace=True)

            ### Fill NaN values and change the column names for Trasactions DataFrame.
            tt.fillna('',inplace=True)
            tt.columns = tt.loc[1]

            ### Remove all unnessesary columns from ShortTerm DataFrame
            st = st[['AccountNumber','Shares With Redemption Fees', 'Symbol']]
            st.fillna(0,inplace = True)

            ### Drop the Portfolios that do not have holdings at the bottom of the Holding Dataframe.
            hd = hd[0:len(hd['Symbol'].dropna())]

            ### Convert Fee from $ as a string to a float
            tb['Fee'] = tb['Fee'].apply(lambda x: x.replace('$','').replace(',',''))
            tb['Fee'] = pd.to_numeric(tb['Fee'],errors='coerce')

            ### Convert Cash Available from $ as a string to a float
            ca['WcCashAvailable'] = ca['WcCashAvailable'].apply(lambda x: x.replace('$','').replace(',',''))
            ca['WcCashAvailable'] = pd.to_numeric(ca['WcCashAvailable'],errors='raise')

            ### Convert the Market Value to a float

            tt['MarketValue'] = pd.to_numeric(tt['MarketValue'],errors='coerce')

            ### Create the Transaction DataFrame that only includes buy orders greater than 999. Sum it to double Check with
            ### manually doing the Transaction sheet in Excel. Total Should equal 21793739.190000002

            tt = tt[(tt['Action']=='Buy') & (tt['MarketValue'][3:]>999)]
            sum(tt['MarketValue'])

            ### Create the Yes value for the RoundTrip column

            tt['RoundTrip'] = 'Yes'

            ### This removes the 403(b) accounts which are longer than 9 characters
            tt = tt[tt['AccountNumber'].apply(len)<10]

            ### Prep the Transactions DataFrame for the merge with holidngs. Keep only relevant columns

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
                messagebox.showwarning("RoundTrips",'''Merged DataFrame has {} instances of RoundTrips and Original Transactions DataFrame has {} instances.\n
There is a difference between the amount of holdings showing roundtrip and the non-duplicated RoundTrip dataframe tt\n
Total difference = {}'''.format(Trades1_drop_duplicated[Trades1_drop_duplicated['RoundTrip'] =='Yes'].shape[0],
                                                 tt.shape[0], abs(Trades1_drop_duplicated[Trades1_drop_duplicated['RoundTrip'] == 'Yes'].shape[0]-tt.shape[0])))
            else:
                messagebox.showinfo("RoundTrips",'Merged DataFrame has {} instances of RoundTrips and Original Transactons DataFrame has {} instances. Everything is good'.format(
                                                Trades1_drop_duplicated[Trades1_drop_duplicated['RoundTrip'] =='Yes'].shape[0],tt.shape[0]))

            ### Find values that aren't in merged dataframe. This will show the instances where a position was bought into, but potentially sold
            ### before the round-trip was over or for positions that were converted to advantage class funds (i.e FSEMX to FSEVX)
            missing = pd.DataFrame()
            missing['OriginalAccountNumber'] = tt['AccountNumber']
            missing['OriginalSymbol'] = tt['Symbol']
            missing['MergeFrame'] = missing['OriginalAccountNumber'] + missing['OriginalSymbol']
            missing2 = pd.DataFrame()
            missing2 = Trades1_drop_duplicated[Trades1_drop_duplicated['RoundTrip'] == 'Yes'][['AccountNumber','Symbol']]
            missing2['MergeFrame'] = missing2['AccountNumber'] + missing2['Symbol']
            findmissing = pd.merge(missing, missing2, how = 'outer', on = ['MergeFrame'])

            ### Fills Symbol and AccountNumber column in all positions that are not contained in both DataFrames.
            print('''---Prints Out the Accounts That Did Not Merge RoundTrip Information---------------------------------------------------
            ----------------------------------------------------------------------------------------------------------------------''')
            findmissing.fillna('ERROR',inplace = True)
            print(findmissing[findmissing['Symbol'] == 'ERROR'])
            messagebox.showwarning('Missing Round-Trips',"""These Accounts are the accounts missing in merged dataframe that added RoundTrip information.\n
More often than not this is the result of the seed accounts that do not have a portfolio or an account that bought a fund that was either sold during the round-trip or that was converted.
Conversions occur for a lower fee fund of the same positions (ie. FSEMX to FSEVX)
{}""".format(findmissing[findmissing['Symbol'] == 'ERROR']))

            ### Creation of a ShortTermCheck varible that calculates the total shares with redemption fees for a later double-check
            ShortTermCheck = st['Shares With Redemption Fees'].sum()

            ### Creation of the new ShortTerm Dataframe that only shows value greater than 0
            st = st[st['Shares With Redemption Fees'] > 0]

            ### Creation of a variable that is used to make sure that the short term shares have been calculated correctly.
            PostManipulatedShortTermCheck = st['Shares With Redemption Fees'].sum()

            if PostManipulatedShortTermCheck != round(ShortTermCheck,4):
                messagebox.showwarning('Short-Term Shares','''The amount of short-term shares with the premanipulated DataFrame is different than the frame after checking for Shares > 0\n
Dataframe Values: Original = {} and New = {}'''.format(ShortTermCheck,PostManipulatedShortTermCheck))
            else:
                messagebox.showinfo('Short-Term Shares','Short term shares are calculated correctly. Dataframe Values: Orignal = {} and New = {}'.format(
                    round(ShortTermCheck,4),PostManipulatedShortTermCheck))

            ### Replaces the hyphon in the account numbers that are in the Short Term shares Dataframe.
            st['AccountNumber'] = st['AccountNumber'].apply(lambda x: x.replace('-',''))

            ### Merge Short-term shares onto the Dataframe that includes holding and round-trips. This is the final DataFrame for Trade Prep

            FinalTradePrep = pd.merge(Trades1_drop_duplicated,st,how = 'left',on=['AccountNumber','Symbol'])

            ### Replace Shares With Redemtion Fees values of Nan with 0

            FinalTradePrep['Shares With Redemption Fees'].fillna(0,inplace = True)

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
            tb['FeeTrade'] = tb['CumulativeFee'] - tb['WcCashAvailable']+.5

            tblength = len(tb.index.values)

            ### Create A DataFrame for Pay By Invoice clients. Used to double check Trade Total + Invoice Total after trading and for reference. Remove Clients from DataFrame
            PayByInvoice = tb[tb['BillingAccount'] == 'Pay By Invoice']
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
                            FeeTrade = tb_no_invoice.loc[i,'FeeTrade']
                            TradingFrame = FinalTradePrep[(FinalTradePrep['AccountNumber'] == AccountNumber) &
                            (FinalTradePrep['Shares With Redemption Fees'] < 1) & (FinalTradePrep['Target'] > 0) & (FinalTradePrep['Symbol'].apply(len) > 4)].sort_values(by = 'Difference',ascending = False)
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
                except:
                    print("Location {} is associated with a client or account that is to be paid by invoice. Portfolio name is: {}".format(i, PayByInvoice.loc[i,'Portfolio']))

            ### Part 2 Is Complete ###

            ###Create New DataFrame that will be used to import trade information to WC.
            WC_Import = FinalTradePrep[FinalTradePrep['TradeValue'] > 0]

            print('''--Double Check That The Trade Values Match Up From the Trade Prep Sheet and the WC_Import Sheet------------------
            ----------------------------------------------------------------------------------------------------------------------''')
            if sum(FinalTradePrep['TradeValue']) == sum(WC_Import['TradeValue']):
                messagebox.showinfo("Trade Value Audit",'Trade values match up: FinalTradePrep = {}, WC_Import = {}'.format(sum(FinalTradePrep['TradeValue']),sum(WC_Import['TradeValue'])))
            else:
                messagebox.showwarning("Trade Value Audit", 'Trade values do not match-up: FinalTradePrep = {}, WC_Import = {}'.format(sum(FinalTradePrep['TradeValue']),sum(WC_Import['TradeValue'])))

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

            ### Creation of the CSV file that will be uploaded.
            getDirectoryName = filedialog.askdirectory(title = 'Choose Folder to Save Files')
            PayByInvoice.to_csv(getDirectoryName + "\PaybyInvoiceClients.csv", index = False)
            FinalTradePrep.to_csv(getDirectoryName + '\FinalTradesPrep.csv', index = False)
            WC_Import_Final.to_csv(getDirectoryName + "\WCImport.csv", index = False, header=False)


            ### Part 3 Is Complete ###

            ### Creation Log


        importFilesButton = tk.Button(self.master, text = "Import Files", command=get_files)
        importFilesButton.pack()
        runBillingButton = tk.Button(self.master, text = "Run Billing", command = runBilling, state = 'disabled')
        runBillingButton.pack()

        self.master.mainloop()

QuarterlyBilling()

"""12/20/2016 Currently my Holdings file with merged Round-Trip data is not adding up to equal the total amount of round trips there
are in the tt DataFrame. Make sure to remove duplicates from the tt DataFrame and account for the fact that
my account and the Leiferman 403(b)s are not shown in the holdings file. Need to get Short-Term File to print from WC
tomorrow as well. After that I can move onto creating which funds will need to be traded based off fee size and whether the
fund is over or under weighted.

12/21/2016 I solved the holdings file issue. I found that Jack Matthews had a duplicate portfolio where his name had just one t and this
was causing roundtrips to get counted twice. I also removed the Leiferman accounts by filtering the dataframe where account numbers are
greater than 9. I still need to get the short-term data to print from WC. Once that section is complete all trading prep will be done.
Then I can move onto part 2 which is developing the code that actually places the trades.

12/22/2016 I have completely finished Part 1. All prep work is completed. I was able to get the ShortTerm share information and then
add that to the Trades1_Drop_Dulicated DataFrame to create the FinalTradePrep dataframe. I have added various checks to make sure
RoundTrip and Short-term info gets calculated correctly. I created the cumulative fee for billing accounts from
the BD billing export. Now I need to figure out how to create the appropriate trades. I need to apply the billed fee to each account
and then loop through the holdings to pick which funds will be sold. Two types of trades will be placed; multiple trades for the same
account under $1000 or one trade that covers the entire fee if the total of multiple trades is not equal to or greater than the fee amount.

12/27/2016 I've finished part 2 designating which positions to sell. Tomorrow I need to go through and make sure that everything matches up
correctly. I need to double check large fee accounts to make sure it is calculating correctly.
I should create a test to make sure that the sum of trading equals fee totals. But a lot of this can be calculated after trades are placed.
The next step is to create the output file for WC upload.

12/29/2016 I've finished Part 3. The import file works correctly. Fidelity has a description of fields required for trading by going to
IMPORT/EXPORT in WC and then clicking import > orders > The question mark in the upper right. But the file exported and imported correctly.
After double checking trade values it lined up correctly with the FeeTrade information from the tb_no_invoice DataFrame. This auditing procedure
will be completed every quarter to make sure that the trades have been completed correctly.

Files that are needed before the Program can be run.

1. CashAvailabe - Cash Available from WC. Get this by going to Groups > All Accounts > Export into CSV.
4. ShortTerm - The WC output of all holdings with short-term shares. Get this by going to Groups > All Accounts > Positions Open/Unrealized > Export
2. Holdings - The Rebalance Report from BD.
3. PortAccount# - Data mining file form BD that just grabs Account# and Portfolio Name
5. Billing - The Billing output from BD. Make sure to run billing on the advisor for Billing Family Accounts Only.
6. Transactions - The transactions output for all holdings the last 30 day from BD.

*** All Files should be saved as .csv files.

Files that are created from Program.

1. Import.csv - To be used for WC upload of trade information.

403(b) and Annuity accounts will have to be billed seperately.

12/29/2016 After trying this again with new information I hit a road block where in my account I bought FUSEX, which Fidelity later converted
to FUSVX which is the same fund but advantage class. I'm going to need to find a way to fix this. I also added a printout which will show Jack
and Karen's accounts if they have any round-trips since the if test will be tripped by them since they don't have portfolios to show up in the
Rebalance Report

12/30/2016 Program is completed and ready to be used for 1st Quarter Billing. I created code today that will show which accounts did not have their round-trip inforamtion
merged to the tradeprep dataframe. That solves the issues that section was giving me.

03/07/2017 I changed the billing into a class structure with some tkinter GUI features. I am adding it to the BWM Command Center Application."""
