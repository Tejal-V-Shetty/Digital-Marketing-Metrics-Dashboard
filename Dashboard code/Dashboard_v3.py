import numpy
import pandas
import datetime
import time
import matplotlib.pyplot as plt
import plotly.express as px
import plotly.graph_objects as go
import json

main_data=[['Date', 'Total Sales - Total', 'Total Sales - In-store', 'Total Sales - Online', 'Total sales - Third-party', 'Same store sales - Total', 'Same store sales - Instore', 'Same store sales - Online', 'Same store sales - Third-party', 'Total traffic', 'Paid traffic', 'Spend', 'Conversions', 'CPA'],[0,0,0,0,0,0,0,0,0,0,0,0,0,0]]
sales=[['Date','Store name','Total','In-store','Online','Total'],[0,0,0,0,0,0],[0,0,0,0,0,0],[0,0,0,0,0,0]]
platform_data=['Platform','Date','Campaign name','Ad set name','Ad name','Impressions','Clicks','Purchases','Amount spent']
platform_metrics=['Impressions','Clicks','Purchases','Amount spent']
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
SC_list=  ['Start Time','Campaign Name','Ad Set Name','Ad Name','Paid Impressions','Swipe-Ups','Purchases','Amount Spent']
TT_list=  ['Date','Campaign name','Ad Group Name','Ad Name'      ,'Impression'      ,'Click'         ,'Cost']
GAds_list=['Date','Campaign'     ,'Ad group'     ,'Impr.'        ,'Clicks'          ,'Conversions'   ,'Cost']

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

def graph_rep():
    global platform_sheet

    #Metrics chosen for the graph
    grp_metric = "Platform"
    grp_list = {}   #Dictionary of all the metrics that denominate the grouping parameter, along with their values
    grp_metric_list = []    #List of all the metrics that denominate the grouping parameters
    x_metric = "Clicks"
    y_metric = "Purchases"
    hover_metric="Campaign name"
    
    sorted_data = platform_sheet.groupby(by=grp_metric,as_index = False)  #Groups by the particular grouping metric
    print("No. of groups = ",sorted_data.ngroups)
    
    for groups in sorted_data:
        #repdf = pandas.DataFrame.from_dict(groups[1])   #Create a dataframe from the returned tuple
        grp_list[groups[0]]=pandas.DataFrame(groups[1])
        grp_metric_list.append(groups[0])
        #print(groups)
        #try:
    
    fig = px.scatter(grp_list[grp_metric_list[0]],x=x_metric,y=y_metric,hover_data=[grp_metric, "Campaign name"])  #Plot the graph
    #x_metric="CTR"
    #y_metric="CPA"
    #fig.update_traces(overwrite=True, x=grp_list[grp_metric_list[0]][x_metric], y=grp_list[grp_metric_list[0]][y_metric],hovertemplate=hover_metric+': %{custom_data}<br>'+x_metric+': %{x}<br>'+y_metric+': %{y}<br>')   #Update x and y axis values
    #fig.update_xaxes(title_text=x_metric)   #Update x axis label
    #fig.update_yaxes(title_text=y_metric)   #Update y axis label
    fig.update_layout(title_text=x_metric+" vs."+y_metric,title_x=0.5)
    fig.update_layout(updatemenus=[{'buttons':[{'method':'update','label':key,'args':[{'x': [value[x_metric]],'y': [value[y_metric]]}]}
                                                  for key, value in grp_list.items()], 'y':1.2},
                                   {'buttons':[{'method':'update','label':param_x,'args':[{x_metric: param_x}]}
                                                  for param_x in platform_metrics], 'y':1.3}
                                  ])
    
    print(grp_list[grp_metric_list[0]][x_metric])
    print(x_metric)    
    fig.show()
    #except:
    #    print("------------Figure disabled-------------")

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
    graph_rep()

main()
