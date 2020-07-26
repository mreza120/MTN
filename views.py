from django.shortcuts import render , redirect
from django.shortcuts import render
from django.http import HttpResponse , HttpResponseRedirect
import pandas
import glob
import os
import numpy as np
from django.contrib.auth import logout, authenticate, login
from django.contrib.auth.decorators import login_required

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
        list_of_files = glob.glob('Z:\IP Plans\Region 5&10\Kerman\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        oo = pandas.read_excel(latest_file,sheet_name='Nokia-IPs')
        gf = oo.groupby(oo['Sites'].str.contains(x))
        hf = oo.groupby(oo['Sites-TDD'].str.contains(x))   
    # For the sites that not exist in both LTE and TDD sections
        if len(list(gf)) == 1 and len(list(hf)) == 1:
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)
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
        table = final_df2.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)
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
            
        table = final_df3.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)

  #### Kerman-ZTE IP Plans Check ####
    elif Province == 'Kerman-ZTE' :
        sheet_names = ['ZTE-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 5&10\Kerman\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)

  ### Esfahan OLD ###

    elif Province == 'Isfahan-Old' :
        list_of_files = glob.glob('Z:\IP Plans\Region 5&10\Esfahan\*.xlsx') # * means all if need specific format then *.csv
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
            return HttpResponse(" Site is not Valid!!!")
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
            final_df [['Sites-TDD',
            
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)

  #### Esfahan DPs  ####
    elif Province == 'Isfahan-DPs' :
        sheet_names = ['Esfahan NEW PAO' , 'Ericsson Routers','Esfahan New DP']
        list_of_files = glob.glob('Z:\IP Plans\Region 5&10\Esfahan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)

  #### Hormozgan  ####
    elif Province == 'Hormozgan' :
        sheet_names = ['BandarAbbas-Old clusters' , 'B.Abbas New LTE','Bandar Abbas New POA' , 'Bandar Abbas U900' , 'Bandar Abbas New DP']
        list_of_files = glob.glob('Z:\IP Plans\Region 5&10\Hormozgan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)

  ### Yazd-Nokia IP Plans check ###
    elif Province == 'Yazd-Nokia' :
        list_of_files = glob.glob('Z:\IP Plans\Region 5&10\Yazd\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        oo = pandas.read_excel(latest_file,sheet_name='Nokia-IPs')
        gf = oo.groupby(oo['Sites'].str.contains(x))
        hf = oo.groupby(oo['Sites-TDD'].str.contains(x))   
    # For the sites that not exist in both LTE and TDD sections
        if len(list(gf)) == 1 and len(list(hf)) == 1:
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)
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
        table = final_df2.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)
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
            # final_df3 = final_df3.drop('Hubsite')
            
        table = final_df3.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)

  #### Yazd-ZTE IP Plans Check ####
    elif Province == 'Yazd-ZTE' :
        sheet_names = ['ZTE-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 5&10\Yazd\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)

  ### Shahrekord ###
    elif Province == 'Chahar-Mahaal' :
        sheet_names = ['Shahrekord OLD' , 'Shahrekord NEW PAO']
        list_of_files = glob.glob('Z:\IP Plans\Region 5&10\Chahar Mahal Bakhtiari\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)

  ### Sistan-Nokia IP Plans check ###    
    elif Province == 'Sistan-Nokia' :     
        list_of_files = glob.glob('Z:\IP Plans\Region 5&10\Sistan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        oo = pandas.read_excel(latest_file,sheet_name='Nokia-IPs')
        gf = oo.groupby(oo['Sites'].str.contains(x))
        hf = oo.groupby(oo['Sites-TDD'].str.contains(x))       
    # For the sites that not exist in both LTE and TDD sections
        if len(list(gf)) == 1 and len(list(hf)) == 1:
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)
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
        table = final_df2.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)
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
            # final_df3 = final_df3.drop('Hubsite')
            
        table = final_df3.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)

    
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
        list_of_files = glob.glob('Z:\IP Plans\Region 5&10\Sistan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


 ############ Region 1&3 #################

  #### Gilan IP Plan Check ####
    elif Province == 'Gilan' :
        sheet_names = ['Gilan-IP-Plan']
        list_of_files = glob.glob('Z:\IP Plans\Region 1&3\Gilan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],sort=False)
        gf = oo.groupby(oo['Sites'].str.contains(x))
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)



  #### Golestan IP Plan Check ####
    elif Province == 'Golestan' :
        sheet_names = ['Golestan-IP-Plan', 'Golestan-IP-Plan-DPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 1&3\Golestan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)



  #### South-Khorasan IP Plan Check ####
    elif Province == 'South-Khorasan' :
        sheet_names = ['Birjand-IP-Plan']
        list_of_files = glob.glob('Z:\IP Plans\Region 1&3\Kh. Jonoubi\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)



  #### Khorasan-Razavi IP Plan Check ####
    elif Province == 'Khorasan-Razavi' :
        sheet_names = ['Razavi-IP', 'Razavi-DP-IP']
        list_of_files = glob.glob('Z:\IP Plans\Region 1&3\Kh. Razavi\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)



  #### North-Khorasan IP Plan Check ####
    elif Province == 'North-Khorasan' :
        sheet_names = ['Bojnourd-IP-Plan']
        list_of_files = glob.glob('Z:\IP Plans\Region 1&3\Kh. Shomali\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


  #### Mazandaran IP Plan Check ####
    elif Province == 'Mazandaran' :
        sheet_names = ['Mazandaran-IP-Plan', 'Mazandaran-DP-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 1&3\Mazandaran\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


 ############ Region 2&4 #################

  #### Ardebil IP Plan Check ####
    elif Province == 'Ardabil' :
        sheet_names = ['Ardebil-IP-Plan' , 'Ardebil-DP-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\Ardebil\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    ##Check wheter the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)



  #### East-Azarbayejan IP Plan Check ####
    elif Province == 'East-Azarbayejan' :
        sheet_names = ['Tabriz-IP-Plan']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\E.Azarbaiejan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


  #### Hamedan IP Plan Check ####
    elif Province == 'Hamadan' :
        sheet_names = ['Hamedan-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\Hamedan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


  #### Ilam IP Plan Check ####
    elif Province == 'Ilam' :
        sheet_names = ['Ilam-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\Ilam\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


  #### Kermanshah IP Plan Check ####
    elif Province == 'Kermanshah' :
        sheet_names = ['Kermanshah-IP-Plan']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\Kermanshah\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


  #### Kermanshah IP Plan Check ####
    elif Province == 'Kermanshah' :
        sheet_names = ['Kermanshah-IP-Plan']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\Kermanshah\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


  #### Kordestan IP Plan Check ####
    elif Province == 'Kurdistan' :
        sheet_names = ['Sanandaj-IPs', 'Sanandaj-DP-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\Kordestan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


  #### Lorestan IP Plan Check ####
    elif Province == 'Lorestan' :
        sheet_names = ['KhorramAbad-IPs', 'KhorramAbad-DP-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\Lorestan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


  #### Markazi IP Plan Check ####
    elif Province == 'Markazi' :
        sheet_names = ['Arak-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\Markazi\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)



  #### Qazvin IP Plan Check ####
    elif Province == 'Qazvin' :
        sheet_names = ['Qazvin-IPs', 'Qazvin-DP-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\Qazvin\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 100 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


  #### West_azarbayejan IP Plan Check ####
    elif Province == 'West-Azarbayejan' :
        sheet_names = ['Oroumieh-IPs', 'Oroumieh-DP-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\W.Azarbaiejan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


  #### Zanjan IP Plan Check ####
    elif Province == 'Zanjan' :
        sheet_names = ['Zanjan-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\Zanjan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


 ############ Region 7&8 ################## 


  #### Qom IP Plan Check ####
    elif Province == 'Qom' :
        sheet_names = ['Qom-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 7&8\Qom\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


  #### Karaj IP Plan Check ####
    elif Province == 'Alborz' :
        sheet_names = ['Alborz-IPs', 'Alborz-DP-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 7&8\Karaj\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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


  #### Semnan IP Plan Check ####
    elif Province == 'Semnan' :
        sheet_names = ['Semnan-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 7&8\Semnan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check wheter the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


  #### TE-OLD IP Plan Check ####
 
    elif Province == 'TE-Ericsson-OLD' :
        list_of_files = glob.glob('Z:\IP Plans\Region 7&8\Tehran-East\*.xlsx') # * means all if need specific format then *.csv
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
            return HttpResponse(" Site is not Valid!!!")
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

    
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)
    

  #### TE Huawei IP Plan Check ####

    elif Province == 'TE-Huawei' :
        sheet_names = ['Huawei-TE-DP' , 'TE-Sub']
        list_of_files = glob.glob('Z:\IP Plans\Region 7&8\Tehran-East\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


  #### TE DPs IP Plan Check ####
    elif Province == 'TE-DPs' :
        sheet_names = ['TE-DP' , 'TE Ericsson DPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 7&8\Tehran-East\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


  #### TW-OLD IP Plan Check ####
 
    elif Province == 'TW-Ericsson-OLD' :
        list_of_files = glob.glob('Z:\IP Plans\Region 7&8\Tehran-West\*.xlsx') # * means all if need specific format then *.csv
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
            return HttpResponse(" Site is not Valid!!!")
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

    
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)

 
  #### TW DPs IP Plan Check ####
    elif Province == 'TW-DPs' :
        sheet_names = ['TW-DP' , 'TW Ericsson DPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 7&8\Tehran-West\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


  #### TE Huawei IP Plan Check ####

    elif Province == 'TW-Huawei' :
        sheet_names = ['Tehran-West (Huawei)']
        list_of_files = glob.glob('Z:\IP Plans\Region 7&8\Tehran-West\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


 ############ Region 6&9 ##################

  ### Kohkiloye Va boyer Ahmad OLD ###

    elif Province == 'Kohgiluyeh-Old' :
        list_of_files = glob.glob('Z:\IP Plans\Region 6&9\Kohkiloye Va boyer Ahmad\*.xlsx') # * means all if need specific format then *.csv
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
            return HttpResponse(" Site is not Valid!!!")
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

            
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)

  ### Kohkiloye Va boyer Ahmad New ###
    elif Province == 'Kohgiluyeh-new' :
        sheet_names = ['Yasouj New POA']
        list_of_files = glob.glob('Z:\IP Plans\Region 6&9\Kohkiloye Va boyer Ahmad\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)

  #### Bushehr DPs  ####
    elif Province == 'Bushehr' :
        sheet_names = ['Bushehr OLD clusters' , 'Bushehr TDD',' Bushehr New PAO', 'Bushsher New DP']
        list_of_files = glob.glob('Z:\IP Plans\Region 6&9\Bushehr\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


  ### Fars OLD ###

    elif Province == 'Fars-Old' :
        list_of_files = glob.glob('Z:\IP Plans\Region 6&9\Fars\*.xlsx') # * means all if need specific format then *.csv
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
            return HttpResponse(" Site is not Valid!!!")
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
            final_df [['Sites-TDD',
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)

  ### Fars DPs ###
    elif Province == 'Fars-DPs' :
        sheet_names = ['Fars new PAO' , 'Fars New DP','Fars Ericsson DP']
        list_of_files = glob.glob('Z:\IP Plans\Region 6&9\Fars\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


  ### Ahwaz OLD ###

    elif Province == 'Khuzestan-OLD' :
        list_of_files = glob.glob('Z:\IP Plans\Region 6&9\Khuzestan\*.xlsx') # * means all if need specific format then *.csv
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
            return HttpResponse(" Site is not Valid!!!")
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
            final_df [['Sites-TDD',
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)

  ### Ahwaz DPs ###
    elif Province == 'Khuzestan-DPs' :
        sheet_names = ['Ahvaz New PAO' , 'Khozestan new DP','Ericsson Routers']
        list_of_files = glob.glob('Z:\IP Plans\Region 6&9\Khuzestan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True,sort=False)

        gf = df.groupby(df['Sites'].str.contains(x))
    #Check whete the Site is Valid or not
        if len(list(gf)) == 1 :
            return HttpResponse(" Site is not Valid!!!")
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
        table = final_df.to_html(index=False ,classes="responstable")
        context = {
        'table': table
         }
        return render(request, "Showtable.html" , context)


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
    import pandas
    list_sheetha = ['2G','3G']
    df = pandas.DataFrame()
    for i in list_sheetha:
        oo = pandas.read_excel('D:\Python\IPPLAN2.xlsx',sheet_name=i)
        df = pandas.concat([df,oo],sort=False)
    return HttpResponse(df.to_html())


#1st Page HTML
def Loginpage(request):
    # next_url = request.GET.get('next') >
    if request.method == 'POST':
        username = request.POST.get('username')
        password = request.POST.get('password')
        user = authenticate(request, username=username, password=password)
        if user is not None:
            # Successful login
            login(request, user)
            # return render(request, 'index4.html')
            return HttpResponseRedirect("http://10.131.57.172/viewsites/query")
        else:
            # undefined user or wrong password
                return render(request,"pagenotfound.html")
    else:
        context = {}
    return render(request, "index3.html", context)
        

#Show Tcode Search page
def SearchPage(request):
        return render(request,"SiteSearch.html")
    

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
















































