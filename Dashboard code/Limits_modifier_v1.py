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
from tkinter import ttk
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
import sklearn
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import PolynomialFeatures
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error, r2_score
from dateutil.relativedelta import relativedelta

#Example set to explain metrics, also the default values
#Grouping: Platform-wise

main_data=[['Date', 'Total Sales - Total', 'Total Sales - In-store', 'Total Sales - Online', 'Total sales - Third-party', 'Same store sales - Total', 'Same store sales - Instore', 'Same store sales - Online', 'Same store sales - Third-party', 'Total traffic', 'Paid traffic', 'Spend', 'Conversions', 'CPA'],[0,0,0,0,0,0,0,0,0,0,0,0,0,0]]
sales=[['Date','Store name','Total','In-store','Online','Total'],[0,0,0,0,0,0],[0,0,0,0,0,0],[0,0,0,0,0,0]]

platform_data=['Platform','Date','Campaign name','Ad set name','Ad name','Impressions','Clicks','Purchases','Amount spent']
platform_filters=['None','Platform','Date','Campaign name','Ad set name','Ad name','Month','Month_num','Week_num']
platform_metrics=['Impressions','Clicks','Purchases','Amount spent']

grp_metric_list = []    #List of all the metrics that denominate the grouping parameters
platform_files={'Facebook':'FB_1year_data','Snapchat':'SC_1year_data','Tiktok':'TT_1year_data','GoogleAds':'GAds_1year_data'}

main_sheet=pandas.DataFrame(main_data)
sales_sheet=pandas.DataFrame(sales)
platform_sheet=pandas.DataFrame()
read_data=pandas.DataFrame()
limits_sheet=pandas.DataFrame()

Metrics_list={'Facebook':['Day','Campaign name','Ad set name','Ad name','Impressions','Link clicks','Adds of payment info','Amount spent (USD)'],
              'Snapchat':['Start Time','Campaign Name','Ad Set Name','Ad Name','Paid Impressions','Swipe-Ups','Purchases','Amount Spent'],
              'Tiktok':['Date','Campaign name','Ad Group Name','Ad Name','Impression','Clicks','Conversions','Cost'],
              'GoogleAds':['Day','Campaign','Ad group','Search keyword','Impr.','Clicks','Conversions','Cost']}

XA_list=  ['Date','LI Name'      ,'Split Name'   ,'Creative Name','Imps'            ,'Clicks'        ,'Total Cost (USD)']
AZ_list=  ['Date','Order'        ,'Line item'    ,'Creative'     ,'Impressions'     ,'Click-throughs','Total cost']

class Dash:
    
    def calc_metrics(self):
        self.sorted_data['CTR'] = (self.sorted_data['Clicks']/self.sorted_data['Impressions'])*100
        self.sorted_data['CPA'] = (self.sorted_data['Amount spent']/self.sorted_data['Purchases'])*100
        self.sorted_data['CPM'] = (self.sorted_data['Amount spent']/self.sorted_data['Impressions'])*1000
        self.sorted_data['CPC'] = (self.sorted_data['Amount spent']/self.sorted_data['Clicks'])

    def grouper(self):  #Creates grouping based on selected parameter(Eg. Platform, Campaign name)
        global grp_metric_list
        self.grp_list.clear()
        self.sorted_data = platform_sheet.groupby(by=[self.grouping_metric.get(),self.filter_metric.get()],as_index = False)[platform_metrics].sum().reset_index()  #Groups by the particular grouping metric
        self.calc_metrics()
        print(self.sorted_data)
        grp_metric_list = self.sorted_data[self.grouping_metric.get()].unique().tolist()    #Gets the new set of selectable items after grouping
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
    
    def model_predict(self):
        X = self.grp_list[self.grp_metric.get()][[self.filter_metric.get()]]
        y = self.grp_list[self.grp_metric.get()][self.y_metric.get()]
        print(X)
        print(y)
        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.5)
        degree=6
        poly_transformer = PolynomialFeatures(degree=degree)
        X_train_poly = poly_transformer.fit_transform(X_train)
        X_test_poly = poly_transformer.transform(X_test)
        model = LinearRegression()
        model.fit(X_train_poly, y_train)
        
        # Predict on the test set
        y_pred = model.predict(X_test_poly)
        
        # Calculate evaluation metrics
        mse = mean_squared_error(y_test, y_pred)
        r2 = r2_score(y_test, y_pred)

        print("mse :")
        print(mse)
        print("r2 :")
        print(r2)
        
        sorted_data = pandas.DataFrame({self.filter_metric.get(): X_test[self.filter_metric.get()], self.y_metric.get()+'_pred': y_pred})
        sorted_data = sorted_data.sort_values(self.filter_metric.get())

        next_val=[numpy.array(X.max()+1)]
        next_val_poly = poly_transformer.transform(next_val)
        print(X_test)
        print(next_val_poly)
        #next_val=numpy.array(numpy.repeat(X.max()+1,7,axis=1))
        #shaped_val=next_val.reshape(-1,1)
        predicted_val = model.predict(next_val_poly)    #Predict the value for current month/week
        print("Predicted value for current scenario : ")
        print(predicted_val)
        
        self.fig = plt.gca()
        self.fig.clear()
        
        # Plot the polynomial regression fit using Seaborn
        self.axs = seaborn.scatterplot(x=X_test[self.filter_metric.get()], y=y_test, label='True Values')
        self.axs = seaborn.lineplot(x=sorted_data[self.filter_metric.get()], y=sorted_data[self.y_metric.get()+'_pred'], color='red', label='Polynomial Regression Fit')
        plt.xlabel(self.filter_metric.get())
        plt.ylabel(self.y_metric.get())
        plt.title('Polynomial Regression Fit')
        plt.legend()
        self.canvas.draw()
        self.canvas.get_tk_widget().place(relx=0,rely=0.25,anchor='nw')

    def __init__(self):
        #Metrics chosen for the graph
        self.grp_list = {}   #Dictionary of all the metrics that denominate the grouping parameter, along with their values
        self.root = Tk()
        
        self.grouping_metric = StringVar()  #The metric for grouping values
        self.grouping_metric.set("Platform")
        self.filter_metric = StringVar()  #The metric for filtering values
        self.filter_metric.set("Week_num")
        self.y_metric = StringVar()
        self.y_metric.set("CTR")   #y_metric stores the metric name for the Y axis (Tkinter object)
        self.grp_metric = StringVar()   #Holds the name of one of the items chosen after grouping - Eg. Facebook

        self.filter_list = Text(self.root,height=100,width=80)#Holds a list of all the values that can be seen after applying the filter
        self.filter_list.place(relx=0.5,rely=0.25,anchor='nw')

        self.metric_limits = limits_sheet.set_index('Campaign type').to_dict('index')

        #Divide the data into groups
        self.grouper()
        
        #Tkinter dropdown setup        
        self.y_drop = OptionMenu(self.root, self.y_metric, *platform_metrics)   #Dropdown to select y metric
        self.y_drop.place(relx=0.14,rely=0.15,anchor='w')
        
        self.grouper_drop = OptionMenu(self.root, self.grouping_metric, *platform_filters)  #Dropdown to select grouping parameter - Eg. Platform, Campaign name
        self.grouper_drop.place(relx=0.6,rely=0.05,anchor='w')

        self.filter_drop = OptionMenu(self.root, self.filter_metric, *platform_filters)  #Dropdown to select filter parameter - Eg. Platform, Campaign name
        self.filter_drop.place(relx=0.6,rely=0.1,anchor='w')

        self.grp_drop = OptionMenu(self.root, self.grp_metric, *grp_metric_list) #Dropdown to select option after grouping - Eg. Facebook, Snapchat
        self.grp_drop.place(relx=0.14,rely=0.05,anchor='w')

        self.button_metrics=Button(self.root,text="Update metrics", command = self.tk_axis_val_update).place(relx=0.1,rely=0.2,anchor='w')#Buttons to update the grouping after selecting the parameters
        self.button_grouper=Button(self.root,text="Update grouping", command = self.grouper).place(relx=0.7,rely=0.05,anchor='w')
        try:
            self.button_grouper=Button(self.root,text="Predict", command = self.model_predict).place(relx=0.7,rely=0.15,anchor='w')
        except:
            print("Unable to form prediction model due to incomprehensible data.")
        self.labely=Label(self.root, text="Y axis : ")
        self.labely.place(relx=0.1,rely=0.15,anchor='w')
        
        self.label_grp=Label(self.root, text=self.grouping_metric.get()+" : ")
        self.label_grp.place(relx=0.1,rely=0.05,anchor='w')
        
        self.hover_metric="Campaign name"

        self.fig = plt.figure()
        self.axs = seaborn.scatterplot(data=self.grp_list[self.grp_metric.get()],x=self.filter_metric.get(),y=self.y_metric.get())
        
        mplcursors.cursor(self.fig).connect("add", lambda sel: sel.annotation.set_text(self.annotation_maker(sel.index)))
        self.canvas = FigureCanvasTkAgg(self.fig.figure,master=self.root)
        
    def tk_axis_val_update(self):
        
        #Update labels
        self.labely.config(text = "Y axis : ")
        self.label_grp.config(text = self.grouping_metric.get()+" : ")

        self.filter_list.delete("1.0",END)
        self.filter_list.insert(INSERT,self.grp_list[self.grp_metric.get()][self.filter_metric.get()])
        
        #Update figure
        self.fig = plt.gca()
        self.fig.clear()
        self.axs = seaborn.scatterplot(data=self.grp_list[self.grp_metric.get()],x=self.filter_metric.get(),y=self.y_metric.get())
        
        try:
            mplcursors.cursor(self.fig).connect("add", lambda sel: sel.annotation.set_text(self.annotation_maker(sel.index)))
        except:
            pass
        
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

def data_entry(Platform):
    global read_data
    global platform_sheet
    
    path = f'C:\\Users\\Tejal\\Documents\\Datawrkz\\Python programs\\Report files\\Year_data_client1\\{platform_files[Platform]}.xlsx'
    read_data=pandas.read_excel(path)    # Read data from platform sheet
    
    for col in read_data.columns:
        if col not in Metrics_list[Platform]:
            print("Removed : ", col)
            read_data = read_data.drop(col, axis=1)

    read_data.fillna(0,inplace = True)    #Fill blanks with 0
    read_data = read_data[Metrics_list[Platform]]   #Arrange the data according to the list
    read_data.insert(0,'Platform',Platform) #Add column with platform name
    read_data.columns = platform_data     #Add the column names to the excel sheet
    read_data['Date']=read_data['Date'].str.slice(0,10) #Get the correct date format - Mainly for Snapchat data as of FB,TT,SC implementation

    platform_sheet=pandas.concat([platform_sheet, read_data], ignore_index = True)  #Add read data to dataframe with all values

def calc_metrics():
    platform_sheet['CTR'] = (platform_sheet['Clicks']/platform_sheet['Impressions'])*100
    platform_sheet['CPA'] = (platform_sheet['Amount spent']/platform_sheet['Purchases'])*100
    platform_sheet['CPM'] = (platform_sheet['Amount spent']/platform_sheet['Impressions'])*1000
    platform_sheet['CPC'] = (platform_sheet['Amount spent']/platform_sheet['Clicks'])
    
    platform_sheet['Campaign type'] = numpy.where(platform_sheet['Campaign name'].str.contains("AW"),"AW","PE")
    platform_sheet['Platform'] = platform_sheet['Platform']+"_"+platform_sheet['Campaign type']
    platform_sheet['Month'] = platform_sheet['Date'].str[0:7]
    
    Month_num_list=platform_sheet['Month'].unique()
    Month_num_list.sort()
    month_num=1
    Month_num_dict={}
    for item in Month_num_list:
        Month_num_dict[item]=month_num
        month_num=month_num+1
        
    Day_num_list=platform_sheet['Date'].unique()
    Day_num_list.sort()
    week_num=1
    day_num=0
    Week_num_dict={}
    for item in Day_num_list:
        day_num=day_num+1
        if day_num%7==0:
            week_num=week_num+1
        Week_num_dict[item]=week_num
    
    platform_sheet['Month_num']=platform_sheet['Month'].map(Month_num_dict)
    platform_sheet['Week_num']=platform_sheet['Date'].map(Week_num_dict)
    
    platform_metrics.append('CTR')
    platform_metrics.append('CPA')
    platform_metrics.append('CPM')
    platform_metrics.append('CPC')
    
def main():
    global limits_sheet

    try:
        data_entry("Facebook")        
    except:
        print("No Facebook data available.")

    try:
        data_entry("Snapchat")      
    except:
        print("No Snapchat data available.")

    try:
        data_entry("Tiktok")        
    except:
        print("No Tiktok data available.")

    try:
        data_entry("GoogleAds")
    except:
        print("No GoogleAds data available.")
    limits_sheet=pandas.read_excel("C:\\Users\\Tejal\\Documents\\Datawrkz\\Python programs\\Report files\\Training data limits_FandB.xlsx")
    calc_metrics()
    test = Dash()
    test.graph_rep()
    
main()
