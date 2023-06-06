import numpy
import pandas
import datetime
import time
import matplotlib.pyplot as plt
import plotly.express as px
import plotly.graph_objects as go
import json
import seaborn
import mplcursors
from tkinter import *
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)

main_data=[['Date', 'Total Sales - Total', 'Total Sales - In-store', 'Total Sales - Online', 'Total sales - Third-party', 'Same store sales - Total', 'Same store sales - Instore', 'Same store sales - Online', 'Same store sales - Third-party', 'Total traffic', 'Paid traffic', 'Spend', 'Conversions', 'CPA'],[0,0,0,0,0,0,0,0,0,0,0,0,0,0]]
sales=[['Date','Store name','Total','In-store','Online','Total'],[0,0,0,0,0,0],[0,0,0,0,0,0],[0,0,0,0,0,0]]
platform_data=['Platform','Date','Campaign name','Ad set name','Ad name','Impressions','Clicks','Purchases','Amount spent']
platform_metrics=['Impressions','Clicks','Purchases','Amount spent']
grp_metric_list = []    #List of all the metrics that denominate the grouping parameters
platform_files={'Facebook':'FB_data','Snapchat':'SC_data','Tiktok':'TT_data'}
main_sheet=pandas.DataFrame(main_data)
sales_sheet=pandas.DataFrame(sales)
platform_sheet=pandas.DataFrame()
read_data=pandas.DataFrame()

Metrics_list={'Facebook':['Day','Campaign name','Ad set name','Ad name','Impressions','Link clicks','Adds of payment info','Amount spent (USD)'],
              'Snapchat':['Start Time','Campaign Name','Ad Set Name','Ad Name','Paid Impressions','Swipe-Ups','Purchases','Amount Spent'],
              'Tiktok':['Date','Campaign name','Ad Group Name','Ad Name','Impression','Clicks','Conversions','Cost']}

XA_list=  ['Date','LI Name'      ,'Split Name'   ,'Creative Name','Imps'            ,'Clicks'        ,'Total Cost (USD)']
AZ_list=  ['Date','Order'        ,'Line item'    ,'Creative'     ,'Impressions'     ,'Click-throughs','Total cost']
GAds_list=['Date','Campaign'     ,'Ad group'     ,'Impr.'        ,'Clicks'          ,'Conversions'   ,'Cost']

class Dash:
    def __init__(self):
        #Metrics chosen for the graph
        self.grouping_metric = "Platform"
        self.grp_list = {}   #Dictionary of all the metrics that denominate the grouping parameter, along with their values

        self.root = Tk()
        self.x_metric = StringVar()
        self.x_metric.set("Clicks")  #x_metric stores the metric name for the X axis (Tkinter object)
        self.y_metric = StringVar()
        self.y_metric.set("Purchases")   #y_metric stores the metric name for the Y axis (Tkinter object)
        
        self.sorted_data = platform_sheet.groupby(by=self.grouping_metric,as_index = False)  #Groups by the particular grouping metric
        print("No. of groups = ",self.sorted_data.ngroups)
        
        for groups in self.sorted_data:
            self.grp_list[groups[0]]=pandas.DataFrame(groups[1])    #Assign datasets according to names of the grouping metric
            grp_metric_list.append(groups[0])  

        self.grp_metric = StringVar()
        self.grp_metric.set(grp_metric_list[0])   #This stores the option chosen after grouping

        
        #Tkinter x and y dropdown setup
        self.x_drop = OptionMenu(self.root, self.x_metric, *platform_metrics)   #Dropdown to select x metric
        self.y_drop = OptionMenu(self.root, self.y_metric, *platform_metrics)   #Dropdown to select y metric
        self.grp_drop = OptionMenu(self.root, self.grp_metric, *grp_metric_list) #Dropdown to select option after grouping
        self.x_drop.pack()
        self.y_drop.pack()
        self.grp_drop.pack()

        self.button=Button(self.root,text="Update metrics", command = self.tk_axis_val_update).pack()
        self.labelx=Label(self.root, text="X axis value : ")
        #self.labelx.place(relx=0.1, rely=0.1, anchor="sw")
        self.labelx.pack()
        self.labely=Label(self.root, text="Y axis value : ")
        self.labely.pack()
        self.label_grp=Label(self.root, text=self.grouping_metric+" : ")
        self.label_grp.pack() 
        self.hover_metric="Campaign name"
        
        self.fig = plt.figure()
        for i in range(0,self.sorted_data.ngroups):    
            self.axs = self.fig.add_subplot(1,self.sorted_data.ngroups,i+1)
            self.axs = seaborn.scatterplot(data=self.grp_list[grp_metric_list[i]],x=self.x_metric.get(),y=self.y_metric.get())
            print(i)
        #mplcursors.cursor(self.fig).connect("add", lambda sel: sel.annotation.set_text(self.grp_list[self.grp_metric.get()][self.hover_metric][sel.index]))
        
        self.canvas = FigureCanvasTkAgg(self.fig.figure,master=self.root)
        
    def tk_axis_val_update(self):
        #Update labels
        self.labelx.config(text = "X axis : " + self.x_metric.get())
        self.labely.config(text = "Y axis : " + self.y_metric.get())
        self.label_grp.config(text = self.grouping_metric+" : "+self.grp_metric.get())

        #Update figure
        self.fig = plt.gca()
        self.fig.clear()
        for i in range(0,self.sorted_data.ngroups):    
            #self.axs[i] = self.fig.add_subplot(1,self.sorted_data.ngroups,i+1)
            self.axs = seaborn.scatterplot(data=self.grp_list[grp_metric_list[i]],x=self.x_metric.get(),y=self.y_metric.get())
        #mplcursors.cursor(self.fig).connect("add", lambda sel: sel.annotation.set_text(self.grp_list[self.grp_metric.get()][self.hover_metric][sel.index]))
        self.canvas.draw()
        self.canvas.get_tk_widget().pack()
        

    def graph_rep(self):
        global platform_sheet
        global updated_plot

        #Tkinter window setup
        self.root.title("Dashboard options")
        #self.root.geometry("600x600")
        
        NavigationToolbar2Tk(self.canvas,self.root)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack()
        #plt.show()
        
        self.root.mainloop()

def excel_writer():
    with pandas.ExcelWriter('C:\\Users\\Tejal Shetty\\Documents\\Python programs\\Report files\\Dashboard_op.xlsx') as writer:
        #pandas.DataFrame(FB_list).to_excel(writer, sheet_name='Platform data', index=False, header=False)
        main_sheet.to_excel(writer, sheet_name = 'Combined', index = False, header = False)
        sales_sheet.to_excel(writer, sheet_name = 'Sales', index = False, header = False)
        platform_sheet.to_excel(writer, sheet_name = 'Platform data', index = False)     # Write the data into the excel sheet with platform data
    print('\n\nWrite complete. Data stored in Dashboard_op.xlsx')

def data_entry(Platform):
    global read_data
    global platform_sheet
    path = f'C:\\Users\\Tejal Shetty\\Documents\\Python programs\\Report Files\\{platform_files[Platform]}.xlsx'
    read_data=pandas.read_excel(path)    # Read data from platform sheet

    for col in read_data.columns:
        if col not in Metrics_list[Platform]:
            print("Removed : ", col)
            read_data = read_data.drop(col, axis=1)

    read_data.fillna(0,inplace = True)    #Fill blanks with 0
    read_data = read_data[Metrics_list[Platform]]   #Arrange the data according to the list
    read_data.insert(0,'Platform',Platform) #Add column with platform name  
    read_data.columns = platform_data     #Add the column names to the excel sheet
    platform_sheet=pandas.concat([platform_sheet, read_data], ignore_index = True)  #Add read data to dataframe with all values


def calc_metrics():
    platform_sheet['CTR'] = (platform_sheet['Clicks']/platform_sheet['Impressions'])*100
    platform_sheet['CPA'] = (platform_sheet['Amount spent']/platform_sheet['Purchases'])*100
    platform_metrics.append('CTR')
    platform_metrics.append('CPA')
    
def main():
    data_entry("Facebook")
    data_entry("Snapchat")
    data_entry("Tiktok")
    calc_metrics()
    excel_writer()
    test = Dash()
    test.graph_rep()

main()
