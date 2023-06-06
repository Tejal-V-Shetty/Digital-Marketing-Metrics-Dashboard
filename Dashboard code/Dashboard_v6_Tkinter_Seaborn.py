import numpy
import pandas
import datetime
import time
import matplotlib.pyplot as plt
import plotly.express as px
import plotly.graph_objects as go
import json
import seaborn
from tkinter import *
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)

main_data=[['Date', 'Total Sales - Total', 'Total Sales - In-store', 'Total Sales - Online', 'Total sales - Third-party', 'Same store sales - Total', 'Same store sales - Instore', 'Same store sales - Online', 'Same store sales - Third-party', 'Total traffic', 'Paid traffic', 'Spend', 'Conversions', 'CPA'],[0,0,0,0,0,0,0,0,0,0,0,0,0,0]]
sales=[['Date','Store name','Total','In-store','Online','Total'],[0,0,0,0,0,0],[0,0,0,0,0,0],[0,0,0,0,0,0]]
platform_data=['Platform','Date','Campaign name','Ad set name','Ad name','Impressions','Clicks','Purchases','Amount spent']
platform_metrics=['Impressions','Clicks','Purchases','Amount spent']
platform_files={'Facebook':'Grimaldis-Tej','Snapchat':'SC_Grimaldis','Tiktok':'TT_Grimaldis'}
main_sheet=pandas.DataFrame(main_data)
sales_sheet=pandas.DataFrame(sales)
platform_sheet=pandas.DataFrame()
read_data=pandas.DataFrame()
updated_plot=0

Metrics_list={'Facebook':['Day','Campaign name','Ad set name','Ad name','Impressions','Link clicks','Adds of payment info','Amount spent (USD)'],
              'Snapchat':['Start Time','Campaign Name','Ad Set Name','Ad Name','Paid Impressions','Swipe-Ups','Purchases','Amount Spent'],
              'Tiktok':['Date','Campaign name','Ad Group Name','Ad Name','Impression','Clicks','Conversions','Cost']}

XA_list=  ['Date','LI Name'      ,'Split Name'   ,'Creative Name','Imps'            ,'Clicks'        ,'Total Cost (USD)']
AZ_list=  ['Date','Order'        ,'Line item'    ,'Creative'     ,'Impressions'     ,'Click-throughs','Total cost']
SC_list=  ['Start Time','Campaign Name','Ad Set Name','Ad Name','Paid Impressions','Swipe-Ups','Purchases','Amount Spent']
TT_list=  ['Date','Campaign name','Ad Group Name','Ad Name'      ,'Impression'      ,'Click'         ,'Cost']
GAds_list=['Date','Campaign'     ,'Ad group'     ,'Impr.'        ,'Clicks'          ,'Conversions'   ,'Cost']

class Dash:
    def __init__(self):
        
        #Metrics chosen for the graph
        self.grp_metric = "Platform"
        self.grp_list = {}   #Dictionary of all the metrics that denominate the grouping parameter, along with their values
        self.grp_metric_list = []    #List of all the metrics that denominate the grouping parameters

        self.root = Tk()
        self.x_metric = StringVar()
        self.x_metric.set("Clicks")  #x_metric stores the metric name for the X axis (Tkinter object)
        self.y_metric = StringVar()
        self.y_metric.set("Purchases")   #y_metric stores the metric name for the Y axis (Tkinter object)

        #Tkinter x and y dropdown setup
        self.x_drop = OptionMenu(self.root, self.x_metric, *platform_metrics)
        self.y_drop = OptionMenu(self.root, self.y_metric, *platform_metrics)
        self.x_drop.pack()
        self.y_drop.pack()

        self.sorted_data = platform_sheet.groupby(by=self.grp_metric,as_index = False)  #Groups by the particular grouping metric
        print("No. of groups = ",self.sorted_data.ngroups)
        
        for groups in self.sorted_data:
            #repdf = pandas.DataFrame.from_dict(groups[1])   #Create a dataframe from the returned tuple
            self.grp_list[groups[0]]=pandas.DataFrame(groups[1])    #Assign datasets according to names of the grouping metric
            self.grp_metric_list.append(groups[0])  
        
        self.button=Button(self.root,text="Update metrics", command = self.tk_axis_val_update).pack()
        self.labelx=Label(self.root, text="X axis value : ")
        self.labelx.pack()
        self.labely=Label(self.root, text="Y axis value : ")
        self.labely.pack() 
        self.hover_metric="Campaign name"

        self.fig = plt.figure()
        self.fig = seaborn.scatterplot(data=self.grp_list[self.grp_metric_list[0]],x=self.x_metric.get(),y=self.y_metric.get())
        self.canvas = FigureCanvasTkAgg(self.fig.figure,master=self.root)
        
    def tk_axis_val_update(self):
        self.labelx.config(text="X axis : " + self.x_metric.get())
        self.labely.config(text="Y axis : " + self.y_metric.get())
        self.fig.clf() #= seaborn.scatterplot(data=self.grp_list[self.grp_metric_list[0]],x=self.x_metric.get(),y=self.y_metric.get())
        self.canvas.draw()
        self.canvas.get_tk_widget().pack()
        #self.fig.set_xdata(self.grp_list[self.grp_metric_list[0]][self.x_metric.get()])

    def graph_rep(self):
        global platform_sheet
        global updated_plot

        #Tkinter window setup
        self.root.title("Dashboard options")
        self.root.geometry("400x400")
        
        #Seaborn plot setup and show
        
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
    #print("\n\n")
    #print(platform_sheet)

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
