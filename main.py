from tkinter import*
from tkinter import ttk
from tkinter import filedialog
import xlrd
import pyodbc
import datetime
import os

class Root(Tk):
    def __init__(self):
        super(Root, self).__init__()
        self.title("AC Data Upload")
        self.minsize(440, 200)

        self.labelFrame = ttk.LabelFrame(self, text = "Select File & Upload")
        self.labelFrame.grid(row = 0,  column = 5, padx=20, pady=40)

        self.labelFrameButton=ttk.Button(self.labelFrame, text="Click to Append Data",command = self.AppendQuery)
        self.labelFrameButton.grid(row=2, columnspan=5, padx=10, pady=10)
        self.SelectFile()



    def BrowseFile(self):
        global filelocation
        self.filelocation = filedialog.askopenfilename(initialdir="/", title="Select File", filetypes=(("text file","*.txt"), ("All Files" , "*.*")))
        self.SelectedFile = ttk.Label(self.labelFrame, text = "")
        self.SelectedFile.grid(row=1, columnspan=5)
        self.SelectedFile.configure(text=self.filelocation)

        self.statusbar = ttk.Label(self.labelFrame, text="Ready…")
        self.statusbar.grid(row=3, column=5)
        print(self.filelocation)



    def SelectFile(self):
        self.FileName = StringVar()
        self.labelSelectFile = ttk.Label(self.labelFrame, text="Select File ")
        self.labelSelectFile.grid(row=0, column=0,sticky=W)
        self.BrowseButton = ttk.Button(self.labelFrame, text="Browser A File", command = self.BrowseFile)
        self.BrowseButton.grid(row=0, column = 1, padx=5, pady=5)



    def AppendQuery(self):
        conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\BaijnathKumar\Documents\ALC3\INVREG.accdb;PWD=Ibm100')
        cursor = conn.cursor()
        print(self.filelocation)
        book = xlrd.open_workbook(self.filelocation)
        sheet = book.sheet_by_name("Sheet1")
        query = """INSERT INTO INVREG ( Account,VendorName,CompanyCode,FiscalYear,DocumentType,DocumentDate,PostingKey,
                      PostingDate,DocumentNumber,Reference,	Amountinlocalcurrency,LocalCurrency,Amountindoccurr,
                      Documentcurrency,ClearingDocument,Clearingdate,DocumentHeaderText,Assignment,PaymentMethod,
                      TermsofPayment,Username,PaymentBlock,Lineitem,BusinessArea,ReferenceKey1,ReferenceKey2,ReferenceKey3) 
                      values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"""
        TodayDate=datetime.date.today()
        Filename = 'Log ' + str(TodayDate) + '.txt'
        myfile = open(Filename, 'a')

        DataCount=0
        for r in range(1, sheet.nrows):
            dateandtimelog = datetime.datetime.now()
            fromatdate=dateandtimelog.strftime("%m-%d-%Y, %H:%M:%S")
            Account = sheet.cell(r, 0).value
            VendorName = ""
            CompanyCode = sheet.cell(r, 1).value
            FiscalYear = sheet.cell(r, 2).value
            DocumentType = sheet.cell(r, 3).value
            DocumentDate = sheet.cell(r, 4).value
            PostingKey = sheet.cell(r, 5).value
            PostingDate = sheet.cell(r, 6).value
            DocumentNumber = sheet.cell(r, 7).value
            Reference = sheet.cell(r, 8).value
            Amountinlocalcurrency = sheet.cell(r, 9).value
            LocalCurrency = sheet.cell(r, 10).value
            Amountindoccurr = sheet.cell(r, 11).value
            Documentcurrency = sheet.cell(r, 12).value
            ClearingDocument = sheet.cell(r, 13).value
            Clearingdate = sheet.cell(r, 14).value
            DocumentHeaderText = sheet.cell(r, 15).value
            Assignment = sheet.cell(r, 16).value
            PaymentMethod = sheet.cell(r, 17).value
            TermsofPayment = sheet.cell(r, 18).value
            Username = sheet.cell(r, 19).value
            PaymentBlock = sheet.cell(r, 20).value
            Lineitem = sheet.cell(r, 21).value
            BusinessArea = sheet.cell(r, 22).value
            ReferenceKey1 = sheet.cell(r, 23).value
            ReferenceKey2 = sheet.cell(r, 24).value
            ReferenceKey3 = sheet.cell(r, 25).value

            values = (Account, VendorName, CompanyCode, FiscalYear, DocumentType, DocumentDate, PostingKey,
                      PostingDate, DocumentNumber, Reference, Amountinlocalcurrency, LocalCurrency, Amountindoccurr,
                      Documentcurrency, ClearingDocument, Clearingdate, DocumentHeaderText, Assignment, PaymentMethod,
                      TermsofPayment, Username, PaymentBlock, Lineitem, BusinessArea, ReferenceKey1, ReferenceKey2,
                      ReferenceKey3)

            try:
                cursor.execute(query, values)
                DataCount=DataCount+1
                print("Added")
                print(fromatdate)
                myfile.write(str(fromatdate) +"--"+ str(Account) + " " + str(VendorName) + " " + str(CompanyCode)+ " " + str(FiscalYear)+ " " + str(DocumentType)+ " " + str(DocumentDate)+ " " + str(PostingKey)+" "+
                      str(PostingDate)+ " " + str(DocumentNumber)+ " " + str(Reference)+ " " + str(Amountinlocalcurrency)+ " " + str(LocalCurrency)+ " " + str(Amountindoccurr)+ " " +
                      str(Documentcurrency)+ " " + str(ClearingDocument)+ " " + str(Clearingdate)+ " " + str(DocumentHeaderText)+ " " + str(Assignment)+ " " + str(PaymentMethod)+ " " +
                      str(TermsofPayment)+ " " + str(Username)+ " " + str(PaymentBlock)+ " " + str(Lineitem)+ " " + str(BusinessArea)+ " " + str(ReferenceKey1)+ " " + str(ReferenceKey2)+ " " +
                      str(ReferenceKey3) + '\n')


            except:
                print("Unable to add")

        myfile.close()
        self.statusbar = ttk.Label(self.labelFrame, text="Data Appended…")
        self.statusbar.grid(row=3, column=5)

        self.statusbar = ttk.Label(self.labelFrame, text="Number of rows added:-")
        self.statusbar.grid(row=3, column=0)

        self.statusbar = ttk.Label(self.labelFrame, text=DataCount)
        self.statusbar.grid(row=3, column=1)
        cursor.close()
        conn.commit()
        conn.close()


if __name__ == "__main__":
    root = Root()
    root.mainloop()