from Tkinter import*
from xlrd import open_workbook
import tkFileDialog,xlrd,xlsxwriter,datetime,time,tkMessageBox
root = Tk()
root.title("SteelBox Inc. Calculator")
global volume_calculation
global cost_calculation
global weight_calculation
global height_size
global width_size
global length_size                                                                      #THIS GLOBAL PART USED FOR WHEN VARIABLES HAVE TO PASSED OTHER FUNCTION
global thickness_size
global surface_area_calculation
global try_usd_exhnage_rate_size
global current_steel_price_size
global A
class GUI(Frame):
    def __init__(self,parent):
        Frame.__init__(self,parent)
        self.grid(columnspan=1)
        self.widget()
    def widget(self):
        self.var=IntVar()
        self.var1=IntVar()
        self.var2=IntVar()
        self.var3=IntVar()
        self.var4=IntVar()                                                          #HERE IS THE INTVAR'S FOR ALL OF THE TEXTVARIABLES AND VARIABLES
        self.var5=IntVar()
        self.var6=IntVar()
        self.var7=IntVar()
        self.frame1 = Frame(root).grid()
        self.label = Label(self.frame1,text="STEELBOX INC. CALCULATOR",fg="black").grid(sticky=E,padx=5,pady=5,row=0,column=2)
        self.button1 = Button(self.frame1,width=20,text="Import",command=self.import_data).grid(padx=5,pady=5)
        self.label_width = Label(self.frame1,text="Width").grid(padx=5,pady=5)                                                  #HERE IS THE BUTTON/ENTRY/LABEL PART
        self.label_length = Label(self.frame1,text="Length").grid(padx=5,pady=5)
        self.label_height = Label(self.frame1,text="Height").grid(padx=5,pady=5)
        self.label_thickness = Label(self.frame1,text="Thickness").grid(padx=5,pady=5)
        self.entry_width= Entry(self.frame1,textvariable=self.var)
        self.entry_width.grid(row=3,column=1,padx=5,pady=5)
        self.entry_length= Entry(self.frame1,textvariable=self.var1)
        self.entry_length.grid(row=4,column=1,padx=5,pady=5)
        self.entry_height= Entry(self.frame1,textvariable=self.var2)
        self.entry_height.grid(row=5,column=1,padx=5,pady=5)
        self.entry_thickness= Entry(self.frame1,textvariable=self.var3)
        self.entry_thickness.grid(row=6,column=1,padx=5,pady=5)
        self.button_calculate= Button(self.frame1,text="Calculate",width=20,command=self.final_calculation).grid(row=9,column=0,padx=5,pady=5)
        self.total_weight_label= Label(self.frame1,text="Total Weight").grid(row=8,column=1,padx=5,pady=5)
        self.total_weight_entry = Entry(self.frame1,textvariable=self.var6)
        self.total_weight_entry.grid(row=9,column=1,padx=5,pady=5)
        self.total_price_label = Label(self.frame1,text="Total Price").grid(row=8,column=2,padx=5,pady=5)
        self.total_price_entry = Entry(self.frame1,textvariable=self.var7)
        self.total_price_entry.grid(row=9,column=2,padx=5,pady=5)
        self.export_button = Button(self.frame1,text="Export",width=20,command=self.export_data).grid(row=9,column=3,padx=5,pady=5)
        self.current_steel_price_label= Label(self.frame1,text="Current Steel Price").grid(row=6,column=2,padx=5,pady=5)
        self.TRY_USD_Exhnage_Rate_label= Label(self.frame1,text="TRY/USD Exhnage Rate").grid(row=7,column=2,padx=5,pady=5)
        self.current_steel_price_entry= Entry(self.frame1,textvariable=self.var4)
        self.current_steel_price_entry.grid(row=6,column=3,padx=5,pady=5)
        self.TRY_USD_Exhnage_Rate_entry= Entry(self.frame1,textvariable=self.var5)
        self.TRY_USD_Exhnage_Rate_entry.grid(row=7,column=3,padx=5,pady=5)
        self.lid_label = Label(self.frame1,text="Lid?").grid(row=3,column=2,padx=5,pady=5)
        self.separator_label = Label(self.frame1,text="Separator?").grid(row=3,column=3,padx=5,pady=5)
        self.varcb = IntVar()
        self.lid_checkbutton = Checkbutton(self.frame1,var=self.varcb).grid(row=4,column=2,padx=5,pady=5)
        self.varcb2 = IntVar()
        self.separator = Checkbutton(self.frame1,var=self.varcb2).grid(row=4,column=3,padx=5,pady=5)
        self.get_button = Button(self.frame1,text="Get",width=15,).grid(row=6,column=4,padx=5,pady=5)                               #THIS BUTTON HAVE NO FUNCTION BECAUSE OF THE CAN NOT DOWNLOAD DOLAR/TL FUNCTION ON INTERNET
    def import_data(self):
        file_path = tkFileDialog.askopenfilename()
        workbook = xlrd.open_workbook(file_path)
        sheet = workbook.sheet_by_index(0)
        width_size = sheet.cell_value(0,1)
        self.var.set(width_size)                                                                                    #THIS PART FOR EXCEL READING AND SHOWING VARIABLES WHAT IT IS READ IN EXCEL
        length_size = sheet.cell_value(1,1)
        self.var1.set(length_size)
        height_size = sheet.cell_value(2,1)
        self.var2.set(height_size)
        thickness_size = sheet.cell_value(3,1)
        self.var3.set(thickness_size)
        self.varcb.set(sheet.cell_value(4,1))
        if self.varcb.get() == 1:
            self.varcb.set(True)
        if self.varcb.get() == 0:
            self.varcb.set(False)
        self.varcb2.set(sheet.cell_value(5,1))
        if self.varcb2.get() == 1:
            self.varcb2.set(True)
        if self.varcb2.get() == 0:
            self.varcb2.set(False)
        current_steel_price_size = sheet.cell_value(6,1)
        self.var4.set(current_steel_price_size)
        try_usd_exhnage_rate_size= sheet.cell_value(7, 1)
        self.var5.set(try_usd_exhnage_rate_size)
    def final_calculation(self):
        width_size = self.entry_width.get()
        length_size = self.entry_length.get()
        if width_size > length_size:
            width_size = self.entry_length.get()
            length_size = self.entry_width.get()                                                                #THIS PART FOR EXCEL VARIABLES OR USER'S VARIABLES CALCULATION
        height_size = self.entry_height.get()
        thickness_size = self.entry_thickness.get()
        current_steel_price_size = self.current_steel_price_entry.get()
        try_usd_exhnage_rate_size = self.TRY_USD_Exhnage_Rate_entry.get()
        A = 2*(float(width_size)*float(height_size) +  float(length_size)*float(height_size)) + (float(width_size)*float(length_size))
        if self.varcb.get() == 1:
            A = A + float(width_size)*float(length_size)
        if self.varcb2.get() == 1:
            A+=float(width_size)*float(height_size)
        if self.varcb.get() == 0:
            A*=float(1)
        if  self.varcb2.get() == 0:
            A*= float(1)
        surface_area_calculation = A
        volume_calculation = float(surface_area_calculation)*float(thickness_size)
        weight_calculation = float(volume_calculation) * float(0.00785)
        cost_calculation = float(weight_calculation) * float(current_steel_price_size) * float(try_usd_exhnage_rate_size) / (1000)
        self.var6.set(weight_calculation)
        self.var7.set(cost_calculation)
    def export_data(self):
        workbook = xlsxwriter.Workbook('output_data.xlsx')
        worksheet = workbook.add_worksheet()
        worksheet.set_column("A:A",15)
        worksheet.write("A1","Date")
        worksheet.write("A2","Time")
        worksheet.write("A3","Width")                                                                                   #THIS PART FOR WRITING TAKEN INFORMATIONS'S ON EXCEL FILE
        worksheet.write("A4","Length")
        worksheet.write("A5","Height")
        worksheet.write("A6","Thickness")
        worksheet.write("A7","Lid Exist? ")
        worksheet.write("A8","Separator Exist?")
        worksheet.write("A9","Steels Price In USD(per ton)")
        worksheet.write("A10","USD/TRY")
        worksheet.write("A11","Weight")
        worksheet.write("A12","Total Cost")
        time = datetime.date.today()
        time1 = datetime.datetime.now()
        time_clock = datetime.datetime.strftime(time1, '%X')
        worksheet.write("B1",str(time))
        worksheet.write("B2",time_clock)
        worksheet.write("B3",self.entry_width.get())
        worksheet.write("B4",self.entry_length.get())
        worksheet.write("B5",self.entry_height.get())
        worksheet.write("B6",self.entry_thickness.get())
        worksheet.write("B7",self.varcb.get())
        worksheet.write("B8",self.varcb2.get())
        worksheet.write("B9",self.current_steel_price_entry.get())
        worksheet.write("B10",self.TRY_USD_Exhnage_Rate_entry.get())
        worksheet.write("B11",self.total_weight_entry.get())
        worksheet.write("B12",self.total_price_entry.get())
        workbook.close()
        tkMessageBox.showinfo("SteelBox Inc. Calculator","Done!******Congratulations!")                             #THIS PART IS EXTRA FOR SHOWING WHAT WHOLE CALCULATION AND PROGRAM FINISHED
app = GUI(root)
root.mainloop()