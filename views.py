from django.shortcuts import render , redirect
from django.shortcuts import render
from django.http import HttpResponse , HttpResponseRedirect
import pandas
import glob
import os
import numpy as np
from django.contrib.auth import logout, authenticate, login
from django.contrib.auth.decorators import login_required
from netmiko import ConnectHandler
import openpyxl
from io import StringIO
from openpyxl.styles import Font, Fill
from datetime import datetime
from pathlib import Path, PureWindowsPath


############################ Main Index for Ip plan shows ####################################
@login_required
def showTcodeTables(request):
 ###getting the information from user in HTML###
    x = request.POST.get('sitename')
    Province = request.POST.get('province_name')

 #Check the input is not just Alpha
    if x.isalpha() or len(x) < 4 :
        return render(request,"pagenotfound.html")

 ############ Region 5&10 #################
  ### Kerman-Nokia IP Plans Check ###
    elif Province == 'Kerman-Nokia' :     
        list_of_files = glob.glob('Y:\IP Plans\Region 5&10\Kerman\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        oo = pandas.read_excel(latest_file,sheet_name='Nokia-IPs')
        gf = oo.groupby(oo['Sites'].str.contains(x))
        hf = oo.groupby(oo['Sites-TDD'].str.contains(x))   
    # For the sites that not exist in both LTE and TDD sections
        if len(list(gf)) == 1 and len(list(hf)) == 1:
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
    # for the sites which dont have TDD traffic and just have 2G traffic
        elif len(list(hf)) == 1 :
            cf = list(gf)[1][1]
            final_df = pandas.DataFrame(data=None)
            for i in cf.index:
                j = i - (i%33) +1
                jj = oo.iloc[[j,j+1,i]]
                jj = jj[['Sites','O&M' ,'Iub','Abis', 'LTE']]
                final_df = pandas.concat([final_df , jj],sort=False)
                final_df.replace(to_replace = np.nan, value ="" , inplace=True) 
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)
    # For sites which just have TDD traffic
        elif len(list(gf)) == 1:
            cf2 = list(hf)[1][1]
            final_df2 = pandas.DataFrame(data=None)
            for i in cf2.index:
                j = i - (i%33) +1
                jj = oo.iloc[[j,j+1,i]]
                jj = jj[['Sites-TDD','LTE-TDD' ,'LTE-TDD(O&M)']]
                final_df2 = pandas.concat([final_df2 , jj],sort=False)
                final_df2.replace(to_replace = np.nan, value ="" , inplace=True) 
            table = final_df2.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)
    # For the sites which have both normal and TDD traffic
        else: 
            cf = list(gf)[1][1]
            cf2 = list(hf)[1][1]
            final_df = pandas.DataFrame(data=None)
            final_df2 = pandas.DataFrame(data=None)

            for i in cf.index:
                j = i - (i%33) +1
                jj = oo.iloc[[j,j+1,i]]
                jj = jj[['Sites','O&M' ,'Iub','Abis', 'LTE']]
                final_df = pandas.concat([final_df , jj],sort=False)
            for i in cf2.index:
                j = i - (i%33) +1
                jj = oo.iloc[[j,j+1,i]]
                jj = jj[['Sites-TDD','LTE-TDD' ,'LTE-TDD(O&M)']]
                final_df2 = pandas.concat([final_df2 , jj],sort=False)
            final_df3 = pandas.concat([final_df , final_df2],sort=False)
            final_df3.replace(to_replace = np.nan, value ="" , inplace=True) 
            
            table = final_df3.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

  #### Kerman-ZTE IP Plans Check ####
    elif Province == 'Kerman-ZTE' :
        sheet_names = ['ZTE-IPs']
        list_of_files = glob.glob('Y:\IP Plans\Region 5&10\Kerman\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                    jj = jj[['Sites','O&M' ,'Iub', 'Abis', 'LTE']]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
        }
        return render(request, "Showtable.html" , context)

  ### Esfahan OLD ###

    elif Province == 'Isfahan-Old' :
        list_of_files = glob.glob('Y:\IP Plans\Region 5&10\Esfahan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        oo = pandas.read_excel(latest_file,sheet_name='Esfahan (2G)')
        oo1 = pandas.read_excel(latest_file,sheet_name='Esfahan(3G)')
        oo2 = pandas.read_excel(latest_file,sheet_name='Esfahan(LTE)')
        oo3 = pandas.read_excel(latest_file,sheet_name='Esfahan LTE TDD')
        oo = oo[['Sites','Transmission node','2G IP Address','2G VLAN ID','2G O&M IP Address','2G O&M VLAN Traffic','Sync IP Address' , 'Sync VLAN ID']]
        oo1 = oo1[['Sites','3G IP Address', '3G VLAN ID','3G O&M IP Address' ,'3G O&M VLAN ID']]
        oo2 = oo2[['Sites','LTE IP Address', 'LTE VLAN ID','LTE  O&M IP Address','LTE O&M VLAN ID']]
        oo3 = oo3[['Sites','DCN', 'LTE 2600','LTE 2600 O&M','LTE 3500', 'LTE 3500 O&M']]
        gf = oo.groupby(oo['Sites'].str.contains(x))
        gf1 = oo1.groupby(oo1['Sites'].str.contains(x))
        gf2 = oo2.groupby(oo2['Sites'].str.contains(x))
        gf3 = oo3.groupby(oo3['Sites'].str.contains(x))
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf1 = list(gf)[1][1].index
            cf2 = list(gf1)[1][1].index
            cf3 = list(gf2)[1][1].index
            cf4 = list(gf3)[1][1].index
            final_df = pandas.DataFrame(data=None)
            final_df1 = pandas.DataFrame(data=None)
            final_df2 = pandas.DataFrame(data=None)
            final_df3 = pandas.DataFrame(data=None)
            final_df4 = pandas.DataFrame(data=None)
            for i in cf1:
                f = (i%30)
                j = i - f 
                jj = oo.iloc[[j,i]]
                final_df1 = pandas.concat([final_df1 , jj],sort=False) 
            for i in cf2:
                f = (i%30)
                j = i - f 
                jj = oo1.iloc[[j,i]]
                final_df2 = pandas.concat([final_df2 , jj],sort=False)     
            for i in cf3:
                f = (i%30)
                j = i - f 
                jj = oo2.iloc[[j,i]]
                final_df3 = pandas.concat([final_df3 , jj],sort=False)  
            for i in cf4:
                f = (i%33)
                j = i - f + 1
                jj = oo3.iloc[[j,j+1,i]]
                final_df4 = pandas.concat([final_df4 , jj],sort=False) 
            final_df1.reset_index(inplace=True)
            final_df [['Sites','Transmission node','2G IP Address','2G VLAN ID','2G O&M IP Address',
                        '2G O&M VLAN Traffic','Sync IP Address' , 'Sync VLAN ID']] = final_df1[['Sites','Transmission node',
                        '2G IP Address','2G VLAN ID','2G O&M IP Address',
                        '2G O&M VLAN Traffic','Sync IP Address' , 'Sync VLAN ID']]
            final_df2.reset_index(inplace=True)
            final_df3.reset_index(inplace=True)
            final_df4.reset_index(inplace=True)
            final_df [['3G IP Address','3G VLAN ID','3G O&M IP Address' ,'3G O&M VLAN ID']] = final_df2[['3G IP Address','3G VLAN ID','3G O&M IP Address' ,'3G O&M VLAN ID']]
            final_df [['LTE IP Address', 'LTE VLAN ID','LTE  O&M IP Address','LTE O&M VLAN ID']] = final_df3[['LTE IP Address', 'LTE VLAN ID','LTE  O&M IP Address','LTE O&M VLAN ID']]
            final_df [['Sites-TDD',	'DCN','LTE 2600','LTE 2600 O&M' ,'LTE 3500' ,'LTE 3500 O&M']] = final_df4[['Sites','DCN', 'LTE 2600','LTE 2600 O&M','LTE 3500', 'LTE 3500 O&M']]
            
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

  #### Esfahan DPs  ####
    elif Province == 'Isfahan-DPs' :
        sheet_names = ['Esfahan NEW PAO' , 'Ericsson Routers','Esfahan New DP']
        list_of_files = glob.glob('Y:\IP Plans\Region 5&10\Esfahan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+2,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
        }
        return render(request, "Showtable.html" , context)

  #### Hormozgan  ####
    elif Province == 'Hormozgan' :
        sheet_names = ['BandarAbbas-Old clusters' , 'B.Abbas New LTE','Bandar Abbas New POA' , 'Bandar Abbas U900' , 'Bandar Abbas New DP']
        list_of_files = glob.glob('Y:\IP Plans\Region 5&10\Hormozgan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+2,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
        }
        return render(request, "Showtable.html" , context)

  ### Yazd-Nokia IP Plans check ###
    elif Province == 'Yazd-Nokia' :
        list_of_files = glob.glob('Y:\IP Plans\Region 5&10\Yazd\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        oo = pandas.read_excel(latest_file,sheet_name='Nokia-IPs')
        gf = oo.groupby(oo['Sites'].str.contains(x))
        hf = oo.groupby(oo['Sites-TDD'].str.contains(x))   
    # For the sites that not exist in both LTE and TDD sections
        if len(list(gf)) == 1 and len(list(hf)) == 1:
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
    # for the sites which dont have TDD traffic and just have 2G traffic
        elif len(list(hf)) == 1 :
            cf = list(gf)[1][1]
            final_df = pandas.DataFrame(data=None)
            for i in cf.index:
                j = i - (i%33) +1
                jj = oo.iloc[[j,j+1,i]]
                jj = jj[['Sites','O&M' ,'Iub','Abis', 'LTE']]
                final_df = pandas.concat([final_df , jj],sort=False)
                final_df.replace(to_replace = np.nan, value ="" , inplace=True) 
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)
    # For sites which just have TDD traffic
        elif len(list(gf)) == 1:
            cf2 = list(hf)[1][1]
            final_df2 = pandas.DataFrame(data=None)
            for i in cf2.index:
                j = i - (i%33) +1
                jj = oo.iloc[[j,j+1,i]]
                jj = jj[['Sites-TDD','LTE-TDD' ,'LTE-TDD(O&M)']]
                final_df2 = pandas.concat([final_df2 , jj],sort=False)
                final_df2.replace(to_replace = np.nan, value ="" , inplace=True) 
            table = final_df2.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)    
        # For the sites which have both normal and TDD traffic
        else: 
            cf = list(gf)[1][1]
            cf2 = list(hf)[1][1]
            final_df = pandas.DataFrame(data=None)
            final_df2 = pandas.DataFrame(data=None)

            for i in cf.index:
                j = i - (i%33) +1
                jj = oo.iloc[[j,j+1,i]]
                jj = jj[['Sites','O&M' ,'Iub','Abis', 'LTE']]
                final_df = pandas.concat([final_df , jj],sort=False)
            for i in cf2.index:
                j = i - (i%33) +1
                jj = oo.iloc[[j,j+1,i]]
                jj = jj[['Sites-TDD','LTE-TDD' ,'LTE-TDD(O&M)']]
                final_df2 = pandas.concat([final_df2 , jj],sort=False)
            final_df3 = pandas.concat([final_df , final_df2],sort=False)
            final_df3.replace(to_replace = np.nan, value ="" , inplace=True) 

            
            table = final_df3.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)
  #### Yazd-ZTE IP Plans Check ####
    elif Province == 'Yazd-ZTE' :
        sheet_names = ['ZTE-IPs']
        list_of_files = glob.glob('Y:\IP Plans\Region 5&10\Yazd\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                    jj = jj[['Sites','O&M' ,'Iub', 'Abis', 'LTE']]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)
  ### Shahrekord ###
    elif Province == 'Chahar-Mahaal' :
        sheet_names = ['Shahrekord OLD' , 'Shahrekord NEW PAO']
        list_of_files = glob.glob('Y:\IP Plans\Region 5&10\Chahar Mahal Bakhtiari\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+2,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)
  ### Sistan-Nokia IP Plans check ###    
    elif Province == 'Sistan-Nokia' :     
        list_of_files = glob.glob('Y:\IP Plans\Region 5&10\Sistan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        oo = pandas.read_excel(latest_file,sheet_name='Nokia-IPs')
        gf = oo.groupby(oo['Sites'].str.contains(x))
        hf = oo.groupby(oo['Sites-TDD'].str.contains(x))       
    # For the sites that not exist in both LTE and TDD sections
        if len(list(gf)) == 1 and len(list(hf)) == 1:
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
    # for the sites which dont have TDD traffic and just have 2G traffic
        elif len(list(hf)) == 1 :
            cf = list(gf)[1][1]
            final_df = pandas.DataFrame(data=None)
            for i in cf.index:
                j = i - (i%33) +1
                jj = oo.iloc[[j,j+1,i]]
                jj = jj[['Sites','O&M' ,'Iub','Abis', 'LTE']]
                final_df = pandas.concat([final_df , jj],sort=False)
                final_df.replace(to_replace = np.nan, value ="" , inplace=True) 
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)    # For sites which just have TDD traffic
        elif len(list(gf)) == 1:
            cf2 = list(hf)[1][1]
            final_df2 = pandas.DataFrame(data=None)
            for i in cf2.index:
                j = i - (i%33) +1
                jj = oo.iloc[[j,j+1,i]]
                jj = jj[['Sites-TDD','LTE-TDD' ,'LTE-TDD(O&M)']]
                final_df2 = pandas.concat([final_df2 , jj],sort=False)
                final_df2.replace(to_replace = np.nan, value ="" , inplace=True) 
            table = final_df2.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)    # For the sites which have both normal and TDD traffic
        else: 
            cf = list(gf)[1][1]
            cf2 = list(hf)[1][1]
            final_df = pandas.DataFrame(data=None)
            final_df2 = pandas.DataFrame(data=None)

            for i in cf.index:
                j = i - (i%33) +1
                jj = oo.iloc[[j,j+1,i]]
                jj = jj[['Sites','O&M' ,'Iub','Abis', 'LTE']]
                final_df = pandas.concat([final_df , jj],sort=False)
            for i in cf2.index:
                j = i - (i%33) +1
                jj = oo.iloc[[j,j+1,i]]
                jj = jj[['Sites-TDD','LTE-TDD' ,'LTE-TDD(O&M)']]
                final_df2 = pandas.concat([final_df2 , jj],sort=False)
            final_df3 = pandas.concat([final_df , final_df2],sort=False)
            final_df3.replace(to_replace = np.nan, value ="" , inplace=True) 
            # final_df3 = final_df3.drop('Hubsite')
            
            table = final_df3.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)
    
    # else:
    #     # of = oo[(oo['Sites'] == "GW") | (oo['Sites'].str.contains(x))]
    #     # return HttpResponse(of.to_html())
    #     gf = oo.groupby(oo['Sites'].str.contains(x))      
    #     cf = list(gf)[1][1].Sites
    #     final_df = pandas.DataFrame(data=None)
    #     for i in cf.index:
    #         j = i - (i%33) +1
    #         jj = oo.iloc[[j,j+1,i]]
    #         final_df = pandas.concat([final_df , jj],sort=False)
    #     return HttpResponse(final_df.to_html(index=False ,classes="responstable"))

  #### Sistan-ZTE IP Plans Check ####
    elif Province == 'Sistan-ZTE' :
        sheet_names = ['ZTE-IPs']
        list_of_files = glob.glob('Y:\IP Plans\Region 5&10\Sistan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                    jj = jj[['Sites','O&M' ,'Iub', 'Abis', 'LTE']]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

 ############ Region 1&3 #################

  #### Gilan IP Plan Check ####
    elif Province == 'Gilan' :
        sheet_names = ['Gilan-IP-Plan']
        list_of_files = glob.glob('Y:\IP Plans\Region 1&3\Gilan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],sort=False)
        gf = oo.groupby(oo['Sites'].str.contains(x))
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1]
            final_df = pandas.DataFrame(data=None)
            for i in cf.index:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)
        final_df.fillna("" , inplace=True)
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
        }
        return render(request, "Showtable.html" , context)


  #### Golestan IP Plan Check ####
    elif Province == 'Golestan' :
        sheet_names = ['Golestan-IP-Plan', 'Golestan-IP-Plan-DPs']
        list_of_files = glob.glob('Y:\IP Plans\Region 1&3\Golestan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)


  #### South-Khorasan IP Plan Check ####
    elif Province == 'South-Khorasan' :
        sheet_names = ['Birjand-IP-Plan']
        list_of_files = glob.glob('Y:\IP Plans\Region 1&3\Kh. Jonoubi\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)


  #### Khorasan-Razavi IP Plan Check ####
    elif Province == 'Khorasan-Razavi' :
        sheet_names = ['Razavi-IP', 'Razavi-DP-IP']
        list_of_files = glob.glob('Y:\IP Plans\Region 1&3\Kh. Razavi\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]] 
                final_df = pandas.concat([final_df , jj],sort=False)
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)


  #### North-Khorasan IP Plan Check ####
    elif Province == 'North-Khorasan' :
        sheet_names = ['Bojnourd-IP-Plan']
        list_of_files = glob.glob('Y:\IP Plans\Region 1&3\Kh. Shomali\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

  #### Mazandaran IP Plan Check ####
    elif Province == 'Mazandaran' :
        sheet_names = ['Mazandaran-IP-Plan', 'Mazandaran-DP-IPs']
        list_of_files = glob.glob('Y:\IP Plans\Region 1&3\Mazandaran\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

 ############ Region 2&4 #################

  #### Ardebil IP Plan Check ####
    elif Province == 'Ardabil' :
        sheet_names = ['Ardebil-IP-Plan' , 'Ardebil-DP-IPs']
        list_of_files = glob.glob('Y:\IP Plans\Region 2&4\Ardebil\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    ##Check wheter the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
        # pandas.set_option('colheader_justify', 'center') 
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)


  #### East-Azarbayejan IP Plan Check ####
    elif Province == 'East-Azarbayejan' :
        sheet_names = ['Tabriz-IP-Plan']
        list_of_files = glob.glob('Y:\IP Plans\Region 2&4\E.Azarbaiejan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

  #### Hamedan IP Plan Check ####
    elif Province == 'Hamadan' :
        sheet_names = ['Hamedan-IPs']
        list_of_files = glob.glob('Y:\IP Plans\Region 2&4\Hamedan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

  #### Ilam IP Plan Check ####
    elif Province == 'Ilam' :
        sheet_names = ['Ilam-IPs']
        list_of_files = glob.glob('Y:\IP Plans\Region 2&4\Ilam\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

  #### Kermanshah IP Plan Check ####
    elif Province == 'Kermanshah' :
        sheet_names = ['Kermanshah-IP-Plan']
        list_of_files = glob.glob('Y:\IP Plans\Region 2&4\Kermanshah\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

  #### Kermanshah IP Plan Check ####
    elif Province == 'Kermanshah' :
        sheet_names = ['Kermanshah-IP-Plan']
        list_of_files = glob.glob('Y:\IP Plans\Region 2&4\Kermanshah\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

  #### Kordestan IP Plan Check ####
    elif Province == 'Kurdistan' :
        sheet_names = ['Sanandaj-IPs', 'Sanandaj-DP-IPs']
        list_of_files = glob.glob('Y:\IP Plans\Region 2&4\Kordestan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

  #### Lorestan IP Plan Check ####
    elif Province == 'Lorestan' :
        sheet_names = ['KhorramAbad-IPs', 'KhorramAbad-DP-IPs']
        list_of_files = glob.glob('Y:\IP Plans\Region 2&4\Lorestan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

  #### Markazi IP Plan Check ####
    elif Province == 'Markazi' :
        sheet_names = ['Arak-IPs']
        list_of_files = glob.glob('Y:\IP Plans\Region 2&4\Markazi\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)


  #### Qazvin IP Plan Check ####
    elif Province == 'Qazvin' :
        sheet_names = ['Qazvin-IPs', 'Qazvin-DP-IPs']
        list_of_files = glob.glob('Y:\IP Plans\Region 2&4\Qazvin\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 100 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

  #### West_azarbayejan IP Plan Check ####
    elif Province == 'West-Azarbayejan' :
        sheet_names = ['Oroumieh-IPs', 'Oroumieh-DP-IPs']
        list_of_files = glob.glob('Y:\IP Plans\Region 2&4\W.Azarbaiejan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

  #### Zanjan IP Plan Check ####
    elif Province == 'Zanjan' :
        sheet_names = ['Zanjan-IPs']
        list_of_files = glob.glob('Y:\IP Plans\Region 2&4\Zanjan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

 ############ Region 7&8 ################## 


  #### Qom IP Plan Check ####
    elif Province == 'Qom' :
        sheet_names = ['Qom-IPs']
        list_of_files = glob.glob('Y:\IP Plans\Region 7&8\Qom\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

  #### Karaj IP Plan Check ####
    elif Province == 'Alborz' :
        sheet_names = ['Alborz-IPs', 'Alborz-DP-IPs']
        list_of_files = glob.glob('Y:\IP Plans\Region 7&8\Karaj\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check wheter the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)
        # return HttpResponse("latest_file")


  #### Semnan IP Plan Check ####
    elif Province == 'Semnan' :
        sheet_names = ['Semnan-IPs']
        list_of_files = glob.glob('Y:\IP Plans\Region 7&8\Semnan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check wheter the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

  #### TE-OLD IP Plan Check ####
 
    elif Province == 'TE-Ericsson-OLD' :
        list_of_files = glob.glob('Y:\IP Plans\Region 7&8\Tehran-East\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        xls = pandas.ExcelFile(latest_file)
        oo = pandas.read_excel(xls,sheet_name='RBS TCU Abis-O&M TE')
        oo1 = pandas.read_excel(xls,sheet_name='Cluster-Iub-Mub TE')
        oo2 = pandas.read_excel(xls,sheet_name='RBS S1X2-O&M TE')
        oo = oo[['Sites','Transmission node','2G IP Address','2G VLAN ID','2G O&M IP Address','2G O&M VLAN Traffic','DCN IP Address' ,
                'DCN VLAN ID','NTP IP Address' , 'Sync VLAN ID']]
        oo1 = oo1[['Sites','3G IP Address', '3G VLAN ID','3G O&M IP Address', '3G O&M VLAN ID']]
        oo2 = oo2[['Sites','IP Address (1800) LTE', 'LTE VLAN ID','IP Address (1800) LTE O&M', 'LTE O&M VLAN ID' ,'IP Address (2600)FDD' ,'FDD VLAN ID' 
                ,'IP Address FDD (2600) O&M' , 'FDD O&M VLAN ID','IP Address (3500)TDD', 'TDD VLAN ID' ,'IP Address TDD (3500) O&M' ,'TDD O&M VLAN ID']]
        gf = oo.groupby(oo['Sites'].str.contains(x))
        gf1 = oo1.groupby(oo1['Sites'].str.contains(x))
        gf2 = oo2.groupby(oo2['Sites'].str.contains(x))
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf1 = list(gf)[1][1].index
            cf2 = list(gf1)[1][1].index
            cf3 = list(gf2)[1][1].index
            final_df = pandas.DataFrame(data=None)
            final_df1 = pandas.DataFrame(data=None)
            final_df2 = pandas.DataFrame(data=None)
            final_df3 = pandas.DataFrame(data=None)
            for i in cf1:
                f = (i%30)
                j = i - f 
                jj = oo.iloc[[j,i]]
                final_df1 = pandas.concat([final_df1 , jj],sort=False) 
            for i in cf2:
                f = (i%30)
                j = i - f 
                jj = oo1.iloc[[j,i]]
                final_df2 = pandas.concat([final_df2 , jj],sort=False)     
            for i in cf3:
                f = (i%31)
                j = i - f +1
                jj = oo2.iloc[[j,i]]
                final_df3 = pandas.concat([final_df3 , jj],sort=False)   
            final_df1.reset_index(inplace=True)
            final_df [['Sites','Transmission node','2G IP Address','2G VLAN ID','2G O&M IP Address','2G O&M VLAN Traffic','DCN IP Address' ,
                'DCN VLAN ID','NTP IP Address' , 'Sync VLAN ID']] = final_df1[['Sites','Transmission node','2G IP Address','2G VLAN ID','2G O&M IP Address','2G O&M VLAN Traffic','DCN IP Address' ,
                'DCN VLAN ID','NTP IP Address' , 'Sync VLAN ID']]

            final_df2.reset_index(inplace=True)
            final_df3.reset_index(inplace=True)
            final_df [['3G IP Address', '3G VLAN ID','3G O&M IP Address', '3G O&M VLAN ID']] = final_df2[['3G IP Address', '3G VLAN ID','3G O&M IP Address', '3G O&M VLAN ID']]
            final_df [['IP Address (1800) LTE', 'LTE VLAN ID','IP Address (1800) LTE O&M', 'LTE O&M VLAN ID' ,'IP Address (2600)FDD' ,'FDD VLAN ID' 
                    ,'IP Address FDD (2600) O&M' , 'FDD O&M VLAN ID','IP Address (3500)TDD', 'TDD VLAN ID' ,'IP Address TDD (3500) O&M' ,'TDD O&M VLAN ID']] = final_df3[['IP Address (1800) LTE', 'LTE VLAN ID','IP Address (1800) LTE O&M', 'LTE O&M VLAN ID' ,'IP Address (2600)FDD' ,'FDD VLAN ID' 
                    ,'IP Address FDD (2600) O&M' , 'FDD O&M VLAN ID','IP Address (3500)TDD', 'TDD VLAN ID' ,'IP Address TDD (3500) O&M' ,'TDD O&M VLAN ID']]

    
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)    

  #### TE Huawei IP Plan Check ####

    elif Province == 'TE-Huawei' :
        sheet_names = ['Huawei-TE-DP' , 'TE-Sub']
        list_of_files = glob.glob('Y:\IP Plans\Region 7&8\Tehran-East\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

  #### TE DPs IP Plan Check ####
    elif Province == 'TE-DPs' :
        sheet_names = ['TE-DP' , 'TE Ericsson DPs']
        list_of_files = glob.glob('Y:\IP Plans\Region 7&8\Tehran-East\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+2,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

  #### TW-OLD IP Plan Check ####
 
    elif Province == 'TW-Ericsson-OLD' :
        list_of_files = glob.glob('Y:\IP Plans\Region 7&8\Tehran-West\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        xls = pandas.ExcelFile(latest_file)
        oo = pandas.read_excel(xls,sheet_name='RBS TCU Abis-O&M TW')
        oo1 = pandas.read_excel(xls,sheet_name='Cluster-Iub-Mub TW')
        oo2 = pandas.read_excel(xls,sheet_name='RBS S1X2-O&M TW')
        oo = oo[['Sites','Transmission node','2G IP Address','2G VLAN ID','2G O&M IP Address','2G O&M VLAN Traffic','DCN IP Address' ,
                'DCN VLAN ID','NTP IP Address' , 'Sync VLAN ID']]
        oo1 = oo1[['Sites','3G IP Address', '3G VLAN ID','3G O&M IP Address', '3G O&M VLAN ID']]
        oo2 = oo2[['Sites','IP Address (1800) LTE', 'LTE VLAN ID','IP Address (1800) LTE O&M', 'LTE O&M VLAN ID' ,'IP Address (2600)FDD' ,'FDD VLAN ID' 
                ,'IP Address FDD (2600) O&M' , 'FDD O&M VLAN ID','IP Address (3500)TDD', 'TDD VLAN ID' ,'IP Address TDD (3500) O&M' ,'TDD O&M VLAN ID']]
        gf = oo.groupby(oo['Sites'].str.contains(x))
        gf1 = oo1.groupby(oo1['Sites'].str.contains(x))
        gf2 = oo2.groupby(oo2['Sites'].str.contains(x))
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf1 = list(gf)[1][1].index
            cf2 = list(gf1)[1][1].index
            cf3 = list(gf2)[1][1].index
            final_df = pandas.DataFrame(data=None)
            final_df1 = pandas.DataFrame(data=None)
            final_df2 = pandas.DataFrame(data=None)
            final_df3 = pandas.DataFrame(data=None)
            for i in cf1:
                f = (i%30)
                j = i - f 
                jj = oo.iloc[[j,i]]
                final_df1 = pandas.concat([final_df1 , jj],sort=False) 
            for i in cf2:
                f = (i%30)
                j = i - f 
                jj = oo1.iloc[[j,i]]
                final_df2 = pandas.concat([final_df2 , jj],sort=False)     
            for i in cf3:
                f = (i%31)
                j = i - f +1
                jj = oo2.iloc[[j,i]]
                final_df3 = pandas.concat([final_df3 , jj],sort=False)   
            final_df1.reset_index(inplace=True)
            final_df [['Sites','Transmission node','2G IP Address','2G VLAN ID','2G O&M IP Address','2G O&M VLAN Traffic','DCN IP Address' ,
                'DCN VLAN ID','NTP IP Address' , 'Sync VLAN ID']] = final_df1[['Sites','Transmission node','2G IP Address','2G VLAN ID','2G O&M IP Address','2G O&M VLAN Traffic','DCN IP Address' ,
                'DCN VLAN ID','NTP IP Address' , 'Sync VLAN ID']]

            final_df2.reset_index(inplace=True)
            final_df3.reset_index(inplace=True)
            final_df [['3G IP Address', '3G VLAN ID','3G O&M IP Address', '3G O&M VLAN ID']] = final_df2[['3G IP Address', '3G VLAN ID','3G O&M IP Address', '3G O&M VLAN ID']]
            final_df [['IP Address (1800) LTE', 'LTE VLAN ID','IP Address (1800) LTE O&M', 'LTE O&M VLAN ID' ,'IP Address (2600)FDD' ,'FDD VLAN ID' 
                    ,'IP Address FDD (2600) O&M' , 'FDD O&M VLAN ID','IP Address (3500)TDD', 'TDD VLAN ID' ,'IP Address TDD (3500) O&M' ,'TDD O&M VLAN ID']] = final_df3[['IP Address (1800) LTE', 'LTE VLAN ID','IP Address (1800) LTE O&M', 'LTE O&M VLAN ID' ,'IP Address (2600)FDD' ,'FDD VLAN ID' 
                    ,'IP Address FDD (2600) O&M' , 'FDD O&M VLAN ID','IP Address (3500)TDD', 'TDD VLAN ID' ,'IP Address TDD (3500) O&M' ,'TDD O&M VLAN ID']]

    
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)
 
  #### TW DPs IP Plan Check ####
    elif Province == 'TW-DPs' :
        sheet_names = ['TW-DP' , 'TW Ericsson DPs']
        list_of_files = glob.glob('Y:\IP Plans\Region 7&8\Tehran-West\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+2,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

  #### TE Huawei IP Plan Check ####

    elif Province == 'TW-Huawei' :
        sheet_names = ['Tehran-West (Huawei)']
        list_of_files = glob.glob('Y:\IP Plans\Region 7&8\Tehran-West\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

 ############ Region 6&9 ##################

  ### Kohkiloye Va boyer Ahmad OLD ###

    elif Province == 'Kohgiluyeh-Old' :
        list_of_files = glob.glob('Y:\IP Plans\Region 6&9\Kohkiloye Va boyer Ahmad\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        oo = pandas.read_excel(latest_file,sheet_name='Yasouj (2G)')
        oo1 = pandas.read_excel(latest_file,sheet_name='Yasouj(3G)')
        oo2 = pandas.read_excel(latest_file,sheet_name='Yasouj(LTE)')
        oo = oo[['Sites','Transmission node','2G IP Address','2G VLAN ID','2G O&M IP Address','2G O&M VLAN Traffic','Sync IP Address' , 'Sync VLAN ID']]
        oo1 = oo1[['Sites','3G IP Address', '3G VLAN ID','3G O&M IP Address' ,'3G O&M VLAN ID']]
        oo2 = oo2[['Sites','LTE IP Address', 'LTE VLAN ID','LTE  O&M IP Address','LTE O&M VLAN ID']]
        gf = oo.groupby(oo['Sites'].str.contains(x))
        gf1 = oo1.groupby(oo1['Sites'].str.contains(x))
        gf2 = oo2.groupby(oo2['Sites'].str.contains(x))
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf1 = list(gf)[1][1].index
            cf2 = list(gf1)[1][1].index
            cf3 = list(gf2)[1][1].index
            final_df = pandas.DataFrame(data=None)
            final_df1 = pandas.DataFrame(data=None)
            final_df2 = pandas.DataFrame(data=None)
            final_df3 = pandas.DataFrame(data=None)
            for i in cf1:
                f = (i%30)
                j = i - f 
                jj = oo.iloc[[j,i]]
                final_df1 = pandas.concat([final_df1 , jj],sort=False) 
            for i in cf2:
                f = (i%30)
                j = i - f 
                jj = oo1.iloc[[j,i]]
                final_df2 = pandas.concat([final_df2 , jj],sort=False)     
            for i in cf3:
                f = (i%30)
                j = i - f 
                jj = oo2.iloc[[j,i]]
                final_df3 = pandas.concat([final_df3 , jj],sort=False)   
            final_df1.reset_index(inplace=True)
            final_df [['Sites','Transmission node','2G IP Address','2G VLAN ID','2G O&M IP Address',
                        '2G O&M VLAN Traffic','Sync IP Address' , 'Sync VLAN ID']] = final_df1[['Sites','Transmission node',
                        '2G IP Address','2G VLAN ID','2G O&M IP Address',
                        '2G O&M VLAN Traffic','Sync IP Address' , 'Sync VLAN ID']]
            final_df2.reset_index(inplace=True)
            final_df3.reset_index(inplace=True)
            final_df [['3G IP Address','3G VLAN ID','3G O&M IP Address' ,'3G O&M VLAN ID']] = final_df2[['3G IP Address','3G VLAN ID','3G O&M IP Address' ,'3G O&M VLAN ID']]
            final_df [['LTE IP Address', 'LTE VLAN ID','LTE  O&M IP Address','LTE O&M VLAN ID']] = final_df3[['LTE IP Address', 'LTE VLAN ID','LTE  O&M IP Address','LTE O&M VLAN ID']]

            
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)
  ### Kohkiloye Va boyer Ahmad New ###
    elif Province == 'Kohgiluyeh-new' :
        sheet_names = ['Yasouj New POA']
        list_of_files = glob.glob('Y:\IP Plans\Region 6&9\Kohkiloye Va boyer Ahmad\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)
  #### Bushehr DPs  ####
    elif Province == 'Bushehr' :
        sheet_names = ['Bushehr OLD clusters' , 'Bushehr TDD',' Bushehr New PAO', 'Bushsher New DP']
        list_of_files = glob.glob('Y:\IP Plans\Region 6&9\Bushehr\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+2,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

  ### Fars OLD ###

    elif Province == 'Fars-Old' :
        list_of_files = glob.glob('Y:\IP Plans\Region 6&9\Fars\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        oo = pandas.read_excel(latest_file,sheet_name='Fars 2G')
        oo1 = pandas.read_excel(latest_file,sheet_name='Fars(3G)')
        oo2 = pandas.read_excel(latest_file,sheet_name='Fars(LTE)')
        oo3 = pandas.read_excel(latest_file,sheet_name='Fars TDD')
        oo = oo[['Sites','Transmission node','2G IP Address','2G VLAN ID','2G O&M IP Address','2G O&M VLAN Traffic','Sync IP Address' , 'Sync VLAN ID']]
        oo1 = oo1[['Sites','3G IP Address', '3G VLAN ID','3G O&M IP Address' ,'3G O&M VLAN ID' , 'U900' ,'U900 Vlan','U900 O&M','U900 O&M Vlan']]
        oo2 = oo2[['Sites','LTE IP Address', 'LTE VLAN ID','LTE  O&M IP Address','LTE O&M VLAN ID']]
        oo3 = oo3[['Sites','DCN', 'LTE 2600','LTE 2600 O&M','LTE 3500', 'LTE 3500 O&M']]
        gf = oo.groupby(oo['Sites'].str.contains(x))
        gf1 = oo1.groupby(oo1['Sites'].str.contains(x))
        gf2 = oo2.groupby(oo2['Sites'].str.contains(x))
        gf3 = oo3.groupby(oo3['Sites'].str.contains(x))
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf1 = list(gf)[1][1].index
            cf2 = list(gf1)[1][1].index
            cf3 = list(gf2)[1][1].index
            cf4 = list(gf3)[1][1].index
            final_df = pandas.DataFrame(data=None)
            final_df1 = pandas.DataFrame(data=None)
            final_df2 = pandas.DataFrame(data=None)
            final_df3 = pandas.DataFrame(data=None)
            final_df4 = pandas.DataFrame(data=None)
            for i in cf1:
                f = (i%30)
                j = i - f 
                jj = oo.iloc[[j,i]]
                final_df1 = pandas.concat([final_df1 , jj],sort=False) 
            for i in cf2:
                f = (i%30)
                j = i - f 
                jj = oo1.iloc[[j,i]]
                final_df2 = pandas.concat([final_df2 , jj],sort=False)     
            for i in cf3:
                f = (i%30)
                j = i - f 
                jj = oo2.iloc[[j,i]]
                final_df3 = pandas.concat([final_df3 , jj],sort=False)  
            for i in cf4:
                f = (i%33)
                j = i - f + 1 
                jj = oo3.iloc[[j,i]]
                final_df4 = pandas.concat([final_df4 , jj],sort=False) 
            final_df1.reset_index(inplace=True)
            final_df4.reset_index(inplace=True)
            final_df2.reset_index(inplace=True)
            final_df [['Sites','Transmission node','2G IP Address','2G VLAN ID','2G O&M IP Address',
                        '2G O&M VLAN Traffic','Sync IP Address' , 'Sync VLAN ID']] = final_df1[['Sites','Transmission node',
                        '2G IP Address','2G VLAN ID','2G O&M IP Address','2G O&M VLAN Traffic','Sync IP Address' , 'Sync VLAN ID']]
            final_df [['3G IP Address','3G VLAN ID','3G O&M IP Address' ,'3G O&M VLAN ID', 'U900' ,'U900 Vlan','U900 O&M','U900 O&M Vlan']] = final_df2[['3G IP Address'
                        ,'3G VLAN ID','3G O&M IP Address' ,'3G O&M VLAN ID', 'U900' ,'U900 Vlan','U900 O&M','U900 O&M Vlan']]
            final_df3.reset_index(inplace=True)
            final_df [['LTE IP Address', 'LTE VLAN ID','LTE  O&M IP Address','LTE O&M VLAN ID']] = final_df3[['LTE IP Address', 'LTE VLAN ID','LTE  O&M IP Address','LTE O&M VLAN ID']]
            final_df4.reset_index(inplace=True)
            final_df [['Sites-TDD',	'DCN','LTE 2600','LTE 2600 O&M' ,'LTE 3500' ,'LTE 3500 O&M']] = final_df4[['Sites','DCN', 'LTE 2600','LTE 2600 O&M','LTE 3500', 'LTE 3500 O&M']]
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)
  ### Fars DPs ###
    elif Province == 'Fars-DPs' :
        sheet_names = ['Fars new PAO' , 'Fars New DP','Fars Ericsson DP']
        list_of_files = glob.glob('Y:\IP Plans\Region 6&9\Fars\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)

  ### Ahwaz OLD ###

    elif Province == 'Khuzestan-OLD' :
        list_of_files = glob.glob('Y:\IP Plans\Region 6&9\Khuzestan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        oo = pandas.read_excel(latest_file,sheet_name='Khozestan 2G')
        oo1 = pandas.read_excel(latest_file,sheet_name='Khozestan(3G)')
        oo2 = pandas.read_excel(latest_file,sheet_name='Khozestan(LTE)')
        oo3 = pandas.read_excel(latest_file,sheet_name='Khozastan TDD')
        oo = oo[['Sites','Transmission node','2G IP Address','2G VLAN ID','2G O&M IP Address','2G O&M VLAN Traffic','Sync-IP-Address' , 'Sync-VLAN-ID']]
        oo1 = oo1[['Sites','3G IP Address', '3G VLAN ID','3G O&M IP Address' ,'3G O&M VLAN ID']]
        oo2 = oo2[['Sites','LTE IP Address', 'LTE VLAN ID','LTE  O&M IP Address','LTE O&M VLAN ID']]
        oo3 = oo3[['Sites','DCN', 'LTE 2600','LTE 2600 O&M','LTE 3500', 'LTE 3500 O&M']]
        gf = oo.groupby(oo['Sites'].str.contains(x))
        gf1 = oo1.groupby(oo1['Sites'].str.contains(x))
        gf2 = oo2.groupby(oo2['Sites'].str.contains(x))
        gf3 = oo3.groupby(oo3['Sites'].str.contains(x))
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf1 = list(gf)[1][1].index
            cf2 = list(gf1)[1][1].index
            cf3 = list(gf2)[1][1].index
            cf4 = list(gf3)[1][1].index
            final_df = pandas.DataFrame(data=None)
            final_df1 = pandas.DataFrame(data=None)
            final_df2 = pandas.DataFrame(data=None)
            final_df3 = pandas.DataFrame(data=None)
            final_df4 = pandas.DataFrame(data=None)
            for i in cf1:
                f = (i%30)
                j = i - f 
                jj = oo.iloc[[j,i]]
                final_df1 = pandas.concat([final_df1 , jj],sort=False) 
            for i in cf2:
                f = (i%30)
                j = i - f 
                jj = oo1.iloc[[j,i]]
                final_df2 = pandas.concat([final_df2 , jj],sort=False)     
            for i in cf3:
                f = (i%30)
                j = i - f 
                jj = oo2.iloc[[j,i]]
                final_df3 = pandas.concat([final_df3 , jj],sort=False)  
            for i in cf4:
                f = (i%33)
                j = i - f + 1 
                jj = oo3.iloc[[j,i]]
                final_df4 = pandas.concat([final_df4 , jj],sort=False) 
            final_df1.reset_index(inplace=True)
            final_df4.reset_index(inplace=True)
            final_df2.reset_index(inplace=True)
            final_df [['Sites','Transmission node','2G IP Address','2G VLAN ID','2G O&M IP Address',
                        '2G O&M VLAN Traffic','Sync-IP-Address' , 'Sync-VLAN-ID']] = final_df1[['Sites','Transmission node',
                        '2G IP Address','2G VLAN ID','2G O&M IP Address','2G O&M VLAN Traffic','Sync-IP-Address' , 'Sync-VLAN-ID']]
            final_df [['3G IP Address','3G VLAN ID','3G O&M IP Address' ,'3G O&M VLAN ID']] = final_df2[['3G IP Address'
                        ,'3G VLAN ID','3G O&M IP Address' ,'3G O&M VLAN ID']]
            final_df3.reset_index(inplace=True)
            final_df [['LTE IP Address', 'LTE VLAN ID','LTE  O&M IP Address','LTE O&M VLAN ID']] = final_df3[['LTE IP Address', 'LTE VLAN ID','LTE  O&M IP Address','LTE O&M VLAN ID']]
            final_df4.reset_index(inplace=True)
            final_df [['Sites-TDD',	'DCN','LTE 2600','LTE 2600 O&M' ,'LTE 3500' ,'LTE 3500 O&M']] = final_df4[['Sites','DCN', 'LTE 2600','LTE 2600 O&M','LTE 3500', 'LTE 3500 O&M']]
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)
  ### Ahwaz DPs ###
    elif Province == 'Khuzestan-DPs' :
        sheet_names = ['Ahvaz New PAO' , 'Khozestan new DP','Ericsson Routers']
        list_of_files = glob.glob('Y:\IP Plans\Region 6&9\Khuzestan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            context = {
            'Site' : x ,
            'province' : Province
                }
            return render(request , "Sitenotfound.html", context)
        else:
            cf = list(gf)[1][1].index
            final_df = pandas.DataFrame(data=None)
            for i in cf:
                f = (i%33)
                j = i - f +1
                if f <= 3:
                    continue
                else:
                    jj = df.iloc[[j,j+1,i]]
                final_df = pandas.concat([final_df , jj],sort=False)  
            final_df.fillna("" , inplace=True)
            table = final_df.to_html(index=False ,classes="responstable")
            context = {
            'table': table
            }
            return render(request, "Showtable.html" , context)


# test IP Plan Check
    elif Province == 'test':

        oo = pandas.read_excel('D:\Python\IPPLAN4.xlsx',sheet_name='2G')
        oo = oo.fillna('--')
        gf = oo[(oo['Sites'].str.contains(x))]
        gf = gf[['Sites','O&M' ,'Iub','Abis', 'LTE']]
        hf = oo[(oo['Sites.1'].str.contains(x))]
        hf = hf[['Sites.1','LTE-TDD' ,'LTE-TDD(O&M)']]
        kk =  pandas.concat([gf,hf],sort=False)
        return HttpResponse(kk.to_html())
        # return HttpResponse(gf.to_html())

    if x == '*':
       return HttpResponse(oo.to_html()) 


#Test Index
def index3(request):
    # list_sheetha = ['2G','3G']
    # df = pandas.DataFrame()
    # for i in list_sheetha:
    #     oo = pandas.read_excel('D:\Python\IPPLAN2.xlsx',sheet_name=i)
    #     df = pandas.concat([df,oo],sort=False)
    # return HttpResponse(df.to_html())
    return render(request, "CheckARPcluster.html")
   


#1st Page HTML
def Loginpage(request):
    # next_url = request.GET.get('next')
    if request.method == 'POST':
        username = request.POST.get('username')
        password = request.POST.get('password')
        user = authenticate(request, username=username, password=password)
        if user is not None:
            # Successful login
            login(request, user)
            # return render(request, 'index4.html')
            return HttpResponseRedirect("http://10.131.57.172/viewsites/Dashboard")
        else:
            # undefined user or wrong password
                return render(request,"UserPasswrong.html")
    else:
        context = {}
    return render(request, "loginpage.html", context)
        

#Show Tcode Search page
def SearchPage(request):
        return render(request,"SiteSearch.html")
    

#Show Dashboard page
@login_required
def dashboard(request):
        context = {
            'username' : request.user.username ,
            'email' : request.user.email
            }
        return render(request,"Dashboard.html",context)

#check ping of Sites home
def check_live_ip_home(request): 
    return render(request,"Checkping.html") 



#check ping of Sites result
def check_live_ip_result(request): 
    IP = request.POST.get('IP')
    rep = os.system('ping ' + IP + r" -n 3")
    if rep == 0:
        message = "This IP is UP"
    else:
        message = "This IP is Down"
    context ={
        "message" : message
    }
    
    return render(request,"checkping.html",context)

# Check GW of one IP
def check_router_GW(request): 
    
    if request.method == 'POST':
        try:
            IP2 = request.POST.get('IP')
            router = {

            'device_type' : 'cisco_ios',

            'ip' : IP2,

            'username' : 'mohammadreza.mahm',

            'password' : 'mash@142',

            }

            net_connect = ConnectHandler(**router)
            hostname = net_connect.find_prompt()
            hostname = hostname.replace("[local]","")
            hostname = "Router name is {0}".format(hostname)
        except:
            hostname = "This IP is not defined"
        context = {
            "message" : hostname

        }
        return render(request,"CheckGW.html",context)
    else:
        return render(request,"CheckGW.html")


#Check ARP Entry    
def CheckARP(request):
    if request.method == 'POST':
        GW = request.POST.get('GW')
        Vlan = request.POST.get('Vlan')
        VRF = request.POST.get('VRF_name')
    
        router = {

        'device_type' : 'cisco_ios',

        'ip' : GW,

        'username' : 'mohammadreza.mahm',

        'password' : 'mash@142',

        }
        try:
            net_connect = ConnectHandler(**router)
            hostname = net_connect.find_prompt()
            text1 = open(r'C:\Users\mohammadreza.mahm\Desktop\WebSite\mysite\Temp Files\Text1.csv' , 'w')
            if hostname.endswith('>'): # If Router is Huawei
                command = "dis arp int vlan {}".format(Vlan)
                net_connect.send_command("screen-length 0 temporary")
                outputt = net_connect.send_command(command)
                text1.write(outputt)
                text1.close()
                df = pandas.read_csv(r"C:\Users\mohammadreza.mahm\Desktop\WebSite\mysite\Temp Files\Text1.csv", delimiter="\s+" )
                df = df[["IP" , "ADDRESS" ]]
                df.columns = ["IP Address", "Mac Address"] 
                df = df[df['IP Address'].str.contains("10\.")]
                table = df.to_html(index=False ,classes="responstable")
            elif hostname.startswith("[local]"): # If Router is Ericsson
                command = "sh arp-cache all-context | inc {}".format(Vlan)
                net_connect.send_command("terminal length 0")
                outputt = net_connect.send_command(command)
                text1.write(outputt)
                text1.close()
                df = pandas.read_csv(r"C:\Users\mohammadreza.mahm\Desktop\WebSite\mysite\Temp Files\Text1.csv", delimiter="\s+",names=["IP Address", "Mac Address" ,"1","2","3","4","5"])
                df = df[["IP Address", "Mac Address"]]
                table = df.to_html(index=False ,classes="responstable")

            elif hostname.endswith('#'): # If Router is Cisco
                if VRF != "o&m" :
                    command = "sh ip arp vrf {} vlan {}".format(VRF,Vlan)

                else: 
                    command = "sh ip arp vlan {}".format(Vlan)
                net_connect.send_command("terminal length 0")
                outputt = net_connect.send_command(command)
                text1.write(outputt)
                text1.close()
                df = pandas.read_csv(r"C:\Users\mohammadreza.mahm\Desktop\WebSite\mysite\Temp Files\Text1.csv", delimiter="\s+")
                df = df[["Address" , "(min)" ]]
                df.columns = ["IP Address", "Mac Address"]
                table = df.to_html(index=False ,classes="responstable")

            context = {
            'table': table
            }
            return render(request, "ShowtableARP.html" , context)
        except:
            context = {
                "message" : "It Seems that Connection has Problem please contact with #NWG IP RAN planning"
            }
            return render(request, "connectionfailed.html" , context)
    else:
        return render(request,"CheckARP.html")

        # return HttpResponse(GW , Vlan , VRF)

#Cluster Cleanup
def ClusterCleanup(request):
    list_of_files = glob.glob('X:\*.xlsx') # * means all if need specific format then *.csv
    latest_file = max(list_of_files, key=os.path.getmtime)
    if request.method == 'POST':
        Clustertype = request.POST.get('clustertype')
        GW = request.POST.get('GW')
        df = pandas.read_excel(latest_file)
        wb = openpyxl.load_workbook(latest_file)
        sheet = wb["Sheet1"]
        try: 
            
            router = {

                    'device_type' : 'cisco_ios',

                    'ip' : str(GW) ,

                    'username' : 'mohammadreza.mahm',

                    'password' : 'mash@142',

                    }
            net_connect = ConnectHandler(**router)
            hostname = net_connect.find_prompt()
        except : 
            print("Can not connect to router")
  
            
        soton_dict = {
            1 : "B" ,
            2 : "C" ,
            3 : "D" ,
            4 : "E" ,
            5 : "F" ,
            6 : "G" ,
            7 : "H" ,
            8 : "I" ,
            9 : "J" ,
            10 : "K" ,
            11 : "L" ,
            12 : "M" ,
            13 : "N" ,
            14 : "O" ,
            15 : "P" ,
            16 : "Q" ,
            17 : "R" ,
            18 : "S" ,
            19 : "T" ,
            20 : "U" ,
            }

        if hostname.endswith('>'):      #Hu// Routers
            for soton in range (1 , df.shape[1]):
                coloumn = df.columns[soton]
                command = "dis arp int vlan {0}".format(df.at[1 ,coloumn])
                net_connect.send_command("screen-length 0 temporary")
                try :
                    outputt = net_connect.send_command(command)
                    df2 = pandas.read_csv(StringIO(outputt) , delimiter="\s+" , names=["IP", "1","2","3","4","5","6","7","8","9"])
                    df2 = df2[["IP"]]
                    df2 = df2[df2['IP'].str.contains("10\.")]
                    for k , j in df2.iterrows():
                        try :
                            location = df[df[coloumn]==j["IP"]].index.values
                            cell = sheet[soton_dict[soton] + str(location[0]+2)]
                            cell.font = Font(size=12 , bold=True , color="003366FF")
                        except:
                            continue
                except :
                    continue

        elif hostname.startswith("[local]"):     #Er// Routers
            for soton in range (1 , df.shape[1]):
                coloumn = df.columns[soton]
                command = "sh arp-cache all-context | inc {}".format(df.at[1 ,coloumn])
                net_connect.send_command("terminal length 0")
                try :
                    outputt = net_connect.send_command(command)
                    df2 = pandas.read_csv(StringIO(outputt), delimiter="\s+",names=["IP", "Mac Address" ,"1","2","3","4","5"])
                    df2 = df2[["IP"]]
                    df2 = df2[df2['IP'].str.contains("10\.")]
                    for k , j in df2.iterrows():
                        try :

                            location = df[df[coloumn]==j["IP"]].index.values
                            cell = sheet[soton_dict[soton] + str(location[0]+2)]
                            cell.font = Font(size=12 , bold=True , color="003366FF")
                        except:
                            continue
                except :
                    continue



        else:   #Cisco Routers

            for soton in range (1 , df.shape[1]):
                coloumn = df.columns[soton]

                if coloumn == "Abis":
                    if Clustertype == "Ericsson":
                        command = "sh arp vrf vpn_eabis | inc {}".format(df.at[1 ,coloumn])
                    if Clustertype == "Huawei":
                        command = "sh arp vrf vpn_habis | inc {}".format(df.at[1 ,coloumn])
                    if Clustertype == "ZTE":
                        command = "sh arp vrf vpn_zabis | inc {}".format(df.at[1 ,coloumn])
                    if Clustertype == "Nokia":
                        command = "sh arp vrf vpn_ziub | inc {}".format(df.at[1 ,coloumn])

                elif coloumn == " 2600 LTE" :
                    command = "sh arp vrf vpn_lte | inc {}".format(df.at[1 ,coloumn])
                elif coloumn ==" 3500 LTE" :
                    command = "sh arp vrf vpn_lte | inc {}".format(df.at[1 ,coloumn])
                elif coloumn == "LTE" :
                    command = "sh arp vrf vpn_lte | inc {}".format(df.at[1 ,coloumn])
                elif coloumn  == "SYNCE" :
                    if Clustertype == "Ericsson":   
                        command = "sh arp vrf vpn_eiub | inc {}".format(df.at[1 ,coloumn])
                    if Clustertype == "Huawei":
                        command = "sh arp vrf vpn_hiub | inc {}".format(df.at[1 ,coloumn])
                    if Clustertype == "ZTE":
                        command = "sh arp vrf vpn_ziub | inc {}".format(df.at[1 ,coloumn])
                    if Clustertype == "Nokia":
                        command = "sh arp vrf vpn_niub | inc {}".format(df.at[1 ,coloumn])
                elif coloumn  == "Iub" :
                    if Clustertype == "Ericsson":  
                        command = "sh arp vrf vpn_eiub | inc {}".format(df.at[1 ,coloumn])
                    if Clustertype == "Huawei":
                        command = "sh arp vrf vpn_hiub | inc {}".format(df.at[1 ,coloumn])
                    if Clustertype == "ZTE":
                        command = "sh arp vrf vpn_ziub | inc {}".format(df.at[1 ,coloumn])
                    if Clustertype == "Nokia":
                        command = "sh arp vrf vpn_niub | inc {}".format(df.at[1 ,coloumn])
                elif coloumn  == "U900":
                    if Clustertype == "Ericsson":  
                        command = "sh arp vrf vpn_eiub | inc {}".format(df.at[1 ,coloumn])
                    if Clustertype == "Huawei":
                        command = "sh arp vrf vpn_hiub | inc {}".format(df.at[1 ,coloumn])
                    if Clustertype == "ZTE":
                        command = "sh arp vrf vpn_ziub | inc {}".format(df.at[1 ,coloumn])
                    if Clustertype == "Nokia":
                        command = "sh arp vrf vpn_niub | inc {}".format(df.at[1 ,coloumn])
                else :
                    command = "sh arp | inc {}".format(df.at[1 ,coloumn])
                
                net_connect.send_command("terminal length 0")
                try :
                    outputt = net_connect.send_command(command)
                    df2 = pandas.read_csv(StringIO(outputt), delimiter="\s+" , names=["1","IP","2","3","4","5"])
                    df2 = df2[["IP"]]
                    df2 = df2[df2["IP"].str.contains("10\.")]
                    for k , j in df2.iterrows():
                        try :
                            location = df[df[coloumn]==j["IP"]].index.values
                            cell = sheet[soton_dict[soton] + str(location[0]+2)]
                            cell.font = Font(size=12 , bold=True , color="003366FF")
                        except : 
                            continue
                except :
                    continue


        day = datetime.strftime(datetime.now(), '%Y_%m_%d')
        hour = datetime.strftime(datetime.now(), '%H_%M_%p')
        wb.save(r'C:\Users\mohammadreza.mahm\Desktop\WebSite\mysite\viewsites\static\viewsites\ARPResult\Result({0})({1}).xlsx'.format(day,hour))
        list_of_files = glob.glob(r"C:\Users\mohammadreza.mahm\Desktop\WebSite\mysite\viewsites\static\viewsites\ARPResult\*.xlsx") # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        filename = os.path.basename(latest_file)
        context = {

            "Name" : filename
        }
        return render(request, "CheckARPclusterResult.html", context)
    
    else:
        return render(request, "CheckARPcluster.html")






#check the user groups
def is_in_multiple_groups(user):
    return user.groups.filter(name__in=['group1', 'group2']).exists()

    """ 
    from django.contrib.auth.decorators import login_required, user_passes_test
    @login_required
    @user_passes_test(is_member) # or @user_passes_test(is_in_multiple_groups)
    def myview(request):
    # Do your processing

    """