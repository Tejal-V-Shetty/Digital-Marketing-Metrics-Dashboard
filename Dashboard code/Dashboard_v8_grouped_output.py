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
import sklearn

#Example set to explain metrics, also the default values
#Grouping: Platform-wise

main_data=[['Date', 'Total Sales - Total', 'Total Sales - In-store', 'Total Sales - Online', 'Total sales - Third-party', 'Same store sales - Total', 'Same store sales - Instore', 'Same store sales - Online', 'Same store sales - Third-party', 'Total traffic', 'Paid traffic', 'Spend', 'Conversions', 'CPA'],[0,0,0,0,0,0,0,0,0,0,0,0,0,0]]
sales=[['Date','Store name','Total','In-store','Online','Total'],[0,0,0,0,0,0],[0,0,0,0,0,0],[0,0,0,0,0,0]]

platform_data=['Platform','Date','Campaign name','Ad set name','Ad name','Impressions','Clicks','Purchases','Amount spent']
platform_filters=['None','Platform','Date','Campaign name','Ad set name','Ad name']
platform_metrics=['Impressions','Clicks','Purchases','Amount spent']

grp_metric_list = []    #List of all the metrics that denominate the grouping parameters
platform_files={'Facebook':'FB_data','Snapchat':'SC_data','Tiktok':'TT_data','GoogleAds':'GAds_data'}

main_sheet=pandas.DataFrame(main_data)
sales_sheet=pandas.DataFrame(sales)
platform_sheet=pandas.DataFrame()
read_data=pandas.DataFrame()
limits_sheet=pandas.DataFrame()

Metrics_list={'Facebook':['Day','Campaign name','Ad set name','Ad name','Impressions','Link clicks','Adds of payment info','Amount spent (USD)'],
              'Snapchat':['Start Time','Campaign Name','Ad Set Name','Ad Name','Paid Impressions','Swipe-Ups','Purchases','Amount Spent'],
              'Tiktok':['Date','Campaign name','Ad Group Name','Ad Name','Impression','Clicks','Conversions','Cost'],
              'GoogleAds':['Day','Campaign','Ad group','Search keyword','Impr.','Clicks','Conversions','Cost']}

Limits_colour_list={'CTR':['r','g','tab:olive'],
                    'CPA':['tab:olive','g','r'],
                    'CPM':['tab:olive','g','r'],
                    'Amount spent':['tab:olive','g','tab:olive'],
                    'CPC':['tab:olive','g','r']}

XA_list=  ['Date','LI Name'      ,'Split Name'   ,'Creative Name','Imps'            ,'Clicks'        ,'Total Cost (USD)']
AZ_list=  ['Date','Order'        ,'Line item'    ,'Creative'     ,'Impressions'     ,'Click-throughs','Total cost']
GAds_list=['Date','Campaign'     ,'Ad group'     ,'Impr.'        ,'Clicks'          ,'Conversions'   ,'Cost']

class Dash:
    
    def calc_metrics(self):
        self.sorted_data['CTR'] = (self.sorted_data['Clicks']/self.sorted_data['Impressions'])*100
        self.sorted_data['CPA'] = (self.sorted_data['Amount spent']/self.sorted_data['Purchases'])*100
        self.sorted_data['CPM'] = (self.sorted_data['Amount spent']/self.sorted_data['Impressions'])*1000
        self.sorted_data['CPC'] = (self.sorted_data['Amount spent']/self.sorted_data['Clicks'])

        try:
            self.sorted_data['Optimization score_CPM']= (1-(abs(self.sorted_data['CPM']-self.sorted_data['Platform'].map(lambda x : self.metric_limits[x]['CPM_Lower']))/(self.sorted_data['Platform'].map(lambda x : self.metric_limits[x]['CPM_Upper'])*2-self.sorted_data['Platform'].map(lambda x : self.metric_limits[x]['CPM_Lower']))))*100
            self.sorted_data['Optimization score_CTR']= (1-(abs(self.sorted_data['CTR']-self.sorted_data['Platform'].map(lambda x : self.metric_limits[x]['CTR_Lower']))/(self.sorted_data['Platform'].map(lambda x : self.metric_limits[x]['CTR_Upper'])*2-self.sorted_data['Platform'].map(lambda x : self.metric_limits[x]['CTR_Lower']))))*100
            self.sorted_data['Optimization score_CPA']= (1-(abs(self.sorted_data['CPA']-self.sorted_data['Platform'].map(lambda x : self.metric_limits[x]['CPA_Lower']))/(self.sorted_data['Platform'].map(lambda x : self.metric_limits[x]['CPA_Upper'])*2-self.sorted_data['Platform'].map(lambda x : self.metric_limits[x]['CPA_Lower']))))*100
            self.sorted_data['Optimization score_CPC']= (1-(abs(self.sorted_data['CPC']-self.sorted_data['Platform'].map(lambda x : self.metric_limits[x]['CPC_Lower']))/(self.sorted_data['Platform'].map(lambda x : self.metric_limits[x]['CPC_Upper'])*2-self.sorted_data['Platform'].map(lambda x : self.metric_limits[x]['CPC_Lower']))))*100
            self.sorted_data['Optimization score_Overall'] = numpy.where(numpy.isnan(self.sorted_data['Optimization score_CPM']),0,self.sorted_data['Optimization score_CPM'])*self.sorted_data['Platform'].map(lambda x : self.metric_limits[x]['Score_weight_CPM'])\
                                                            +numpy.where(numpy.isnan(self.sorted_data['Optimization score_CTR']),0,self.sorted_data['Optimization score_CTR'])*self.sorted_data['Platform'].map(lambda x : self.metric_limits[x]['Score_weight_CTR'])\
                                                            +numpy.where(numpy.isnan(self.sorted_data['Optimization score_CPA']),0,self.sorted_data['Optimization score_CPC'])*self.sorted_data['Platform'].map(lambda x : self.metric_limits[x]['Score_weight_CPA'])\
                                                            +numpy.where(numpy.isnan(self.sorted_data['Optimization score_CPC']),0,self.sorted_data['Optimization score_CPC'])*self.sorted_data['Platform'].map(lambda x : self.metric_limits[x]['Score_weight_CPC'])
        except:
            pass
        
    def grouper(self):  #Creates grouping based on selected parameter(Eg. Platform, Campaign name)
        global grp_metric_list
        self.grp_list.clear()
        self.sorted_data = platform_sheet.groupby(by=[self.grouping_metric.get(),self.filter_metric.get()],as_index = False)[platform_metrics].sum().reset_index()  #Groups by the particular grouping metric
        self.calc_metrics()
        print(self.sorted_data)
        grp_metric_list = self.sorted_data[self.grouping_metric.get()].unique().tolist()
        for groups in grp_metric_list:
            self.grp_list[groups]=self.sorted_data[self.sorted_data[self.grouping_metric.get()]== groups]
        print(grp_metric_list)

        try:    #Update the options for the dropdown to select divisions
            menu=self.grp_drop["menu"]
            menu.delete(0,"end")
            for string in grp_metric_list:
                menu.add_command(label=string, command=lambda value=string: self.grp_metric.set(value))
        except:
            pass
        
        self.grp_metric.set(grp_metric_list[0])   #This stores the option chosen after grouping

    def annotation_maker(self, index):
        print(index)
        #Current annotation : Filter name + X_value + Y_value
        Annotation = self.grp_list[self.grp_metric.get()][self.filter_metric.get()][index]+"\n"+self.x_metric.get()+": "+str(self.grp_list[self.grp_metric.get()][self.x_metric.get()][index])+"\n"+self.y_metric.get()+": "+str(self.grp_list[self.grp_metric.get()][self.y_metric.get()][index])
        return Annotation

    def __init__(self):
        #Metrics chosen for the graph
        self.grp_list = {}   #Dictionary of all the metrics that denominate the grouping parameter, along with their values
        self.root = Tk()
        self.grouping_metric = StringVar()  #The metric for grouping values
        self.grouping_metric.set("Platform")
        self.filter_metric = StringVar()  #The metric for filtering values
        self.filter_metric.set("Campaign name")
        self.x_metric = StringVar()
        self.x_metric.set("Clicks")  #x_metric stores the metric name for the X axis (Tkinter object)
        self.y_metric = StringVar()
        self.y_metric.set("CPA")   #y_metric stores the metric name for the Y axis (Tkinter object)
        self.grp_metric = StringVar()   #Holds the name of one of the items chosen after grouping - Eg. Facebook

        self.filter_list = Text(self.root,height=100,width=80)#Holds a list of all the values that can be seen after applying the filter
        self.filter_list.place(relx=0.5,rely=0.25,anchor='nw')

        self.metric_limits = limits_sheet.set_index('Campaign type').to_dict('index')
        #Divide the data into groups
        self.grouper()
        
        #Tkinter dropdown setup
        self.x_drop = OptionMenu(self.root, self.x_metric, *platform_metrics)   #Dropdown to select x metric
        self.x_drop.place(relx=0.14,rely=0.1,anchor='w')
        
        self.y_drop = OptionMenu(self.root, self.y_metric, *platform_metrics)   #Dropdown to select y metric
        self.y_drop.place(relx=0.14,rely=0.15,anchor='w')
        
        self.grouper_drop = OptionMenu(self.root, self.grouping_metric, *platform_filters)  #Dropdown to select grouping parameter - Eg. Platform, Campaign name
        self.grouper_drop.place(relx=0.6,rely=0.05,anchor='w')

        self.filter_drop = OptionMenu(self.root, self.filter_metric, *platform_filters)  #Dropdown to select filter parameter - Eg. Platform, Campaign name
        self.filter_drop.place(relx=0.6,rely=0.1,anchor='w')

        self.grp_drop = OptionMenu(self.root, self.grp_metric, *grp_metric_list) #Dropdown to select option after grouping - Eg. Facebook, Snapchat
        self.grp_drop.place(relx=0.14,rely=0.05,anchor='w')

        self.button_metrics=Button(self.root,text="Update metrics", command = self.tk_axis_val_update).place(relx=0.1,rely=0.2,anchor='w')
        self.button_grouper=Button(self.root,text="Update grouping", command = self.grouper).place(relx=0.7,rely=0.05,anchor='w')

        self.labelx=Label(self.root, text="X axis : ")
        self.labelx.place(relx=0.1,rely=0.1,anchor='w')
        
        self.labely=Label(self.root, text="Y axis : ")
        self.labely.place(relx=0.1,rely=0.15,anchor='w')
        
        self.label_grp=Label(self.root, text=self.grouping_metric.get()+" : ")
        self.label_grp.place(relx=0.1,rely=0.05,anchor='w')
        
        self.hover_metric="Campaign name"

        self.fig = plt.figure()
        self.axs = seaborn.scatterplot(data=self.grp_list[self.grp_metric.get()],x=self.x_metric.get(),y=self.y_metric.get())
        try:    #Display limits for the selected metric along Y-axis as per pre-defined values
            self.axs.axhspan(0,self.metric_limits[self.grp_metric.get()][self.y_metric.get()+"_Lower"],facecolor=Limits_colour_list[self.y_metric.get()][0],alpha=0.3)
            self.axs.axhspan(self.metric_limits[self.grp_metric.get()][self.y_metric.get()+"_Lower"],self.metric_limits[self.grp_metric.get()][self.y_metric.get()+"_Upper"],facecolor=Limits_colour_list[self.y_metric.get()][1],alpha=0.3)
            self.axs.axhspan(self.metric_limits[self.grp_metric.get()][self.y_metric.get()+"_Upper"],self.metric_limits[self.grp_metric.get()][self.y_metric.get()+"_Upper"]*2,facecolor=Limits_colour_list[self.y_metric.get()][2],alpha=0.3)
        except:
            pass
        
        try:    #Display limits for the selected metric along X-axis as per pre-defined values
            self.axs.axvspan(0,self.metric_limits[self.grp_metric.get()][self.x_metric.get()+"_Lower"],facecolor=Limits_colour_list[self.x_metric.get()][0],alpha=0.5)
            self.axs.axvspan(self.metric_limits[self.grp_metric.get()][self.x_metric.get()+"_Lower"],self.metric_limits[self.grp_metric.get()][self.x_metric.get()+"_Upper"],facecolor=Limits_colour_list[self.x_metric.get()][1],alpha=0.5)
            self.axs.axvspan(self.metric_limits[self.grp_metric.get()][self.x_metric.get()+"_Upper"],self.metric_limits[self.grp_metric.get()][self.x_metric.get()+"_Upper"]*2,facecolor=Limits_colour_list[self.x_metric.get()][2],alpha=0.5)
        except:
            pass
        
        mplcursors.cursor(self.fig).connect("add", lambda sel: sel.annotation.set_text(self.annotation_maker(sel.index)))
        self.canvas = FigureCanvasTkAgg(self.fig.figure,master=self.root)
        
        
    def tk_axis_val_update(self):
        
        #Update labels
        self.labelx.config(text = "X axis : ")
        self.labely.config(text = "Y axis : ")
        self.label_grp.config(text = self.grouping_metric.get()+" : ")

        self.filter_list.delete("1.0",END)
        self.filter_list.insert(INSERT,self.grp_list[self.grp_metric.get()][self.filter_metric.get()])
        
        #Update figure
        self.fig = plt.gca()
        self.fig.clear()
        self.axs = seaborn.scatterplot(data=self.grp_list[self.grp_metric.get()],x=self.x_metric.get(),y=self.y_metric.get())
        try:    #Display limits for the selected metric along Y-axis as per pre-defined values
            self.axs.axhspan(0,self.metric_limits[self.grp_metric.get()][self.y_metric.get()+"_Lower"],facecolor=Limits_colour_list[self.y_metric.get()][0],alpha=0.5)
            self.axs.axhspan(self.metric_limits[self.grp_metric.get()][self.y_metric.get()+"_Lower"],self.metric_limits[self.grp_metric.get()][self.y_metric.get()+"_Upper"],facecolor=Limits_colour_list[self.y_metric.get()][1],alpha=0.5)
            self.axs.axhspan(self.metric_limits[self.grp_metric.get()][self.y_metric.get()+"_Upper"],self.metric_limits[self.grp_metric.get()][self.y_metric.get()+"_Upper"]*2,facecolor=Limits_colour_list[self.y_metric.get()][2],alpha=0.5)
        except:
            pass
        
        try:    #Display limits for the selected metric along X-axis as per pre-defined values
            self.axs.axvspan(0,self.metric_limits[self.grp_metric.get()][self.x_metric.get()+"_Lower"],facecolor=Limits_colour_list[self.x_metric.get()][0],alpha=0.5)
            self.axs.axvspan(self.metric_limits[self.grp_metric.get()][self.x_metric.get()+"_Lower"],self.metric_limits[self.grp_metric.get()][self.x_metric.get()+"_Upper"],facecolor=Limits_colour_list[self.x_metric.get()][1],alpha=0.5)
            self.axs.axvspan(self.metric_limits[self.grp_metric.get()][self.x_metric.get()+"_Upper"],self.metric_limits[self.grp_metric.get()][self.x_metric.get()+"_Upper"]*2,facecolor=Limits_colour_list[self.x_metric.get()][2],alpha=0.5)
        except:
            pass
        
        try:
            mplcursors.cursor(self.fig).connect("add", lambda sel: sel.annotation.set_text(self.annotation_maker(sel.index)))
        except:
            pass
        #mplcursors.cursor(self.fig).connect("add", lambda sel: sel.annotation.set_text(self.grp_list[self.grp_metric.get()][self.hover_metric][sel.index]))
        self.canvas.draw()
        self.canvas.get_tk_widget().place(relx=0,rely=0.25,anchor='nw')

    def graph_rep(self):
        global platform_sheet
        global updated_plot

        #Tkinter window setup
        self.root.title("Dashboard options")
        self.root.geometry("600x600")
        NavigationToolbar2Tk(self.canvas,self.root)
        self.canvas.draw()
        self.canvas.get_tk_widget().place(relx=0,rely=0.25,anchor='nw')#.pack()
        self.root.mainloop()
        
    def analysis_output(self):
        with pandas.ExcelWriter('C:\\Users\\Tejal Shetty\\Documents\\Python programs\\Report files\\Analysis_output.xlsx') as writer:
            #pandas.DataFrame(FB_list).to_excel(writer, sheet_name='Platform data', index=False, header=False)
            self.sorted_data.to_excel(writer, sheet_name = 'Platform data', index = False)     # Write the data into the excel sheet with sorted data
        print('\n\nWrite complete. Data stored in Analysis_output.xlsx')
        
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
    try:
        read_data['Date']=read_data['Date'].str.slice(0,10) #Get the correct date format - Mainly for Snapchat data as of FB,TT,SC implementation
    except:
        pass
    platform_sheet=pandas.concat([platform_sheet, read_data], ignore_index = True)  #Add read data to dataframe with all values

def calc_metrics():
    platform_sheet['CTR'] = (platform_sheet['Clicks']/platform_sheet['Impressions'])*100
    platform_sheet['CPA'] = (platform_sheet['Amount spent']/platform_sheet['Purchases'])*100
    platform_sheet['CPM'] = (platform_sheet['Amount spent']/platform_sheet['Impressions'])*1000
    platform_sheet['CPC'] = (platform_sheet['Amount spent']/platform_sheet['Clicks'])
    platform_sheet['Campaign type'] = numpy.where(platform_sheet['Campaign name'].str.contains("AW"),"AW","PE")
    platform_sheet['Platform']=platform_sheet['Platform']+"_"+platform_sheet['Campaign type']

    platform_metrics.append('CTR')
    platform_metrics.append('CPA')
    platform_metrics.append('CPM')
    platform_metrics.append('CPC')
    
def main():
    global limits_sheet
    
    data_entry("Facebook")
    data_entry("Snapchat")
    data_entry("Tiktok")
    data_entry("GoogleAds")
    
    limits_sheet=pandas.read_excel("C:\\Users\\Tejal Shetty\\Documents\\Python programs\\Report Files\\Training data limits_FandB.xlsx")
    calc_metrics()
    excel_writer()
    test = Dash()
    test.graph_rep()
main()
