from django.shortcuts import render , redirect
from django.shortcuts import render
from django.http import HttpResponse
import pandas
import glob
import os
import numpy as np


# Create your views here.


############################ Main Index ####################################
def index(request):
###getting the information from user in HTML###
    x = request.POST.get('sitename')
    Province = request.POST.get('province_name')

 #Check the input is not just Alpha
    if x.isalpha() or len(x) < 4 :
        return HttpResponse("Site is not Valid")

 # test IP Plan Check
    elif Province == 'test':

        oo = pandas.read_excel('D:\Python\IPPLAN4.xlsx',sheet_name='2G')
        oo = oo.fillna('--')
        gf = oo[(oo['Sites'].str.contains(x))]
        gf = gf[['Sites','O&M' ,'Iub','Abis', 'LTE']]
        hf = oo[(oo['Sites.1'].str.contains(x))]
        hf = hf[['Sites.1','LTE-TDD' ,'LTE-TDD(O&M)']]
        kk =  pandas.concat([gf,hf])
        return HttpResponse(kk.to_html())
        # return HttpResponse(gf.to_html())

    if x == '*':
       return HttpResponse(oo.to_html()) 

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
                final_df = pandas.concat([final_df , jj])
                final_df.replace(to_replace = np.nan, value ="" , inplace=True) 
            return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))
    # For sites which just have TDD traffic
        elif len(list(gf)) == 1:
            cf2 = list(hf)[1][1]
            final_df2 = pandas.DataFrame(data=None)
            for i in cf2.index:
                j = i - (i%33) +1
                jj = oo.iloc[[j,j+1,i]]
                jj = jj[['Sites-TDD','LTE-TDD' ,'LTE-TDD(O&M)']]
                final_df2 = pandas.concat([final_df2 , jj])
                final_df2.replace(to_replace = np.nan, value ="" , inplace=True) 
            return HttpResponse(final_df2.to_html(index=False,justify='center',col_space='150'))
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
                final_df = pandas.concat([final_df , jj])
            for i in cf2.index:
                j = i - (i%33) +1
                jj = oo.iloc[[j,j+1,i]]
                jj = jj[['Sites-TDD','LTE-TDD' ,'LTE-TDD(O&M)']]
                final_df2 = pandas.concat([final_df2 , jj])
            final_df3 = pandas.concat([final_df , final_df2])
            final_df3.replace(to_replace = np.nan, value ="" , inplace=True) 
            
            return HttpResponse(final_df3.to_html(index=False,justify='center',col_space='150'))

 #### Kerman-ZTE IP Plans Check ####
    elif Province == 'Kerman-ZTE' :
        sheet_names = ['ZTE-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 5&10\Kerman\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])  
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))


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
                final_df = pandas.concat([final_df , jj])
                final_df.replace(to_replace = np.nan, value ="" , inplace=True) 
            return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))
    # For sites which just have TDD traffic
        elif len(list(gf)) == 1:
            cf2 = list(hf)[1][1]
            final_df2 = pandas.DataFrame(data=None)
            for i in cf2.index:
                j = i - (i%33) +1
                jj = oo.iloc[[j,j+1,i]]
                jj = jj[['Sites-TDD','LTE-TDD' ,'LTE-TDD(O&M)']]
                final_df2 = pandas.concat([final_df2 , jj])
                final_df2.replace(to_replace = np.nan, value ="" , inplace=True) 
            return HttpResponse(final_df2.to_html(index=False,justify='center',col_space='150'))
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
                final_df = pandas.concat([final_df , jj])
            for i in cf2.index:
                j = i - (i%33) +1
                jj = oo.iloc[[j,j+1,i]]
                jj = jj[['Sites-TDD','LTE-TDD' ,'LTE-TDD(O&M)']]
                final_df2 = pandas.concat([final_df2 , jj])
            final_df3 = pandas.concat([final_df , final_df2])
            final_df3.replace(to_replace = np.nan, value ="" , inplace=True) 
            # final_df3 = final_df3.drop('Hubsite')
            
            return HttpResponse(final_df3.to_html(index=False,justify='center',col_space='150'))

 #### Yazd-ZTE IP Plans Check ####
    elif Province == 'Yazd-ZTE' :
        sheet_names = ['ZTE-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 5&10\Yazd\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])  
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))

 ### Shahrekord ###
    elif Province == 'Chahar-Mahaal' :
        sheet_names = ['Shahrekord OLD' , 'Shahrekord NEW PAO']
        list_of_files = glob.glob('Z:\IP Plans\Region 5&10\Chahar Mahal Bakhtiari\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])  
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))


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
                final_df = pandas.concat([final_df , jj])
                final_df.replace(to_replace = np.nan, value ="" , inplace=True) 
            return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))
    # For sites which just have TDD traffic
        elif len(list(gf)) == 1:
            cf2 = list(hf)[1][1]
            final_df2 = pandas.DataFrame(data=None)
            for i in cf2.index:
                j = i - (i%33) +1
                jj = oo.iloc[[j,j+1,i]]
                jj = jj[['Sites-TDD','LTE-TDD' ,'LTE-TDD(O&M)']]
                final_df2 = pandas.concat([final_df2 , jj])
                final_df2.replace(to_replace = np.nan, value ="" , inplace=True) 
            return HttpResponse(final_df2.to_html(index=False,justify='center',col_space='150'))
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
                final_df = pandas.concat([final_df , jj])
            for i in cf2.index:
                j = i - (i%33) +1
                jj = oo.iloc[[j,j+1,i]]
                jj = jj[['Sites-TDD','LTE-TDD' ,'LTE-TDD(O&M)']]
                final_df2 = pandas.concat([final_df2 , jj])
            final_df3 = pandas.concat([final_df , final_df2])
            final_df3.replace(to_replace = np.nan, value ="" , inplace=True) 
            # final_df3 = final_df3.drop('Hubsite')
            
            return HttpResponse(final_df3.to_html(index=False,justify='center',col_space='150'))

    
    # else:
    #     # of = oo[(oo['Sites'] == "GW") | (oo['Sites'].str.contains(x))]
    #     # return HttpResponse(of.to_html())
    #     gf = oo.groupby(oo['Sites'].str.contains(x))      
    #     cf = list(gf)[1][1].Sites
    #     final_df = pandas.DataFrame(data=None)
    #     for i in cf.index:
    #         j = i - (i%33) +1
    #         jj = oo.iloc[[j,j+1,i]]
    #         final_df = pandas.concat([final_df , jj])
    #     return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))


 #### Sistan-ZTE IP Plans Check ####
    elif Province == 'Sistan-ZTE' :
        sheet_names = ['ZTE-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 5&10\Sistan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])  
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))


############ Region 1&3 #################

 #### Gilan IP Plan Check ####
    elif Province == 'Gilan' :
        sheet_names = ['Gilan-IP-Plan']
        list_of_files = glob.glob('Z:\IP Plans\Region 1&3\Gilan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo])
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
                final_df = pandas.concat([final_df , jj])
        final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))



 #### Golestan IP Plan Check ####
    elif Province == 'Golestan' :
        sheet_names = ['Golestan-IP-Plan', 'Golestan-IP-Plan-DPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 1&3\Golestan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))



 #### South-Khorasan IP Plan Check ####
    elif Province == 'South-Khorasan' :
        sheet_names = ['Birjand-IP-Plan']
        list_of_files = glob.glob('Z:\IP Plans\Region 1&3\Kh. Jonoubi\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))



 #### Khorasan-Razavi IP Plan Check ####
    elif Province == 'Khorasan-Razavi' :
        sheet_names = ['Razavi-IP', 'Razavi-DP-IP']
        list_of_files = glob.glob('Z:\IP Plans\Region 1&3\Kh. Razavi\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))



 #### North-Khorasan IP Plan Check ####
    elif Province == 'North-Khorasan' :
        sheet_names = ['Bojnourd-IP-Plan']
        list_of_files = glob.glob('Z:\IP Plans\Region 1&3\Kh. Shomali\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))


 #### Mazandaran IP Plan Check ####
    elif Province == 'Mazandaran' :
        sheet_names = ['Mazandaran-IP-Plan', 'Mazandaran-DP-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 1&3\Mazandaran\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])  
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))


############ Region 2&4 #################

 #### Ardebil IP Plan Check ####
    elif Province == 'Ardabil' :
        sheet_names = ['Ardebil-IP-Plan', 'Ardebil-DP-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\Ardebil\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])  
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))


 #### East-Azarbayejan IP Plan Check ####
    elif Province == 'East-Azarbayejan' :
        sheet_names = ['Tabriz-IP-Plan']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\E.Azarbaiejan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])  
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))


 #### Hamedan IP Plan Check ####
    elif Province == 'Hamadan' :
        sheet_names = ['Hamedan-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\Hamedan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])  
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))


 #### Ilam IP Plan Check ####
    elif Province == 'Ilam' :
        sheet_names = ['Ilam-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\Ilam\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])  
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))


 #### Kermanshah IP Plan Check ####
    elif Province == 'Kermanshah' :
        sheet_names = ['Kermanshah-IP-Plan']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\Kermanshah\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])  
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))


 #### Kermanshah IP Plan Check ####
    elif Province == 'Kermanshah' :
        sheet_names = ['Kermanshah-IP-Plan']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\Kermanshah\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])  
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))


 #### Kordestan IP Plan Check ####
    elif Province == 'Kurdistan' :
        sheet_names = ['Sanandaj-IPs', 'Sanandaj-DP-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\Kordestan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])  
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))


 #### Lorestan IP Plan Check ####
    elif Province == 'Lorestan' :
        sheet_names = ['KhorramAbad-IPs', 'KhorramAbad-DP-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\Lorestan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])  
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))


 #### Markazi IP Plan Check ####
    elif Province == 'Markazi' :
        sheet_names = ['Arak-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\Markazi\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])  
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))



 #### Qazvin IP Plan Check ####
    elif Province == 'Qazvin' :
        sheet_names = ['Qazvin-IPs', 'Qazvin-DP-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\Qazvin\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])  
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))


 #### West_azarbayejan IP Plan Check ####
    elif Province == 'West-Azarbayejan' :
        sheet_names = ['Oroumieh-IPs', 'Oroumieh-DP-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\W.Azarbaiejan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])  
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))


 #### Zanjan IP Plan Check ####
    elif Province == 'Zanjan' :
        sheet_names = ['Zanjan-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 2&4\Zanjan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])  
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))


############ Region 7&8 ##################


 #### Qom IP Plan Check ####
    elif Province == 'Qom' :
        sheet_names = ['Qom-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 7&8\Qom\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])  
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))


 #### Karaj IP Plan Check ####
    elif Province == 'Alborz' :
        sheet_names = ['Alborz-IPs', 'Alborz-DP-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 7&8\Karaj\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])  
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))


 #### Semnan IP Plan Check ####
    elif Province == 'Semnan' :
        sheet_names = ['Semnan-IPs']
        list_of_files = glob.glob('Z:\IP Plans\Region 7&8\Semnan\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])  
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))


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
                final_df1 = pandas.concat([final_df1 , jj]) 
            for i in cf2:
                f = (i%30)
                j = i - f 
                jj = oo1.iloc[[j,i]]
                final_df2 = pandas.concat([final_df2 , jj])     
            for i in cf3:
                f = (i%31)
                j = i - f 
                jj = oo2.iloc[[j,i]]
                final_df3 = pandas.concat([final_df3 , jj])   
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

    
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))



 #### TW-OLD IP Plan Check ####
 
    elif Province == 'TW-Ericsson-OLD' :
        list_of_files = glob.glob('Z:\IP Plans\Region 7&8\Tehran-West\*.xlsx') # * means all if need specific format then *.csv
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
                final_df1 = pandas.concat([final_df1 , jj]) 
            for i in cf2:
                f = (i%30)
                j = i - f 
                jj = oo1.iloc[[j,i]]
                final_df2 = pandas.concat([final_df2 , jj])     
            for i in cf3:
                f = (i%31)
                j = i - f 
                jj = oo2.iloc[[j,i]]
                final_df3 = pandas.concat([final_df3 , jj])   
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

    
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))



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
                final_df1 = pandas.concat([final_df1 , jj]) 
            for i in cf2:
                f = (i%30)
                j = i - f 
                jj = oo1.iloc[[j,i]]
                final_df2 = pandas.concat([final_df2 , jj])     
            for i in cf3:
                f = (i%30)
                j = i - f 
                jj = oo2.iloc[[j,i]]
                final_df3 = pandas.concat([final_df3 , jj])   
            final_df1.reset_index(inplace=True)
            final_df [['Sites','Transmission node','2G IP Address','2G VLAN ID','2G O&M IP Address',
                        '2G O&M VLAN Traffic','Sync IP Address' , 'Sync VLAN ID']] = final_df1[['Sites','Transmission node',
                        '2G IP Address','2G VLAN ID','2G O&M IP Address',
                        '2G O&M VLAN Traffic','Sync IP Address' , 'Sync VLAN ID']]
            final_df2.reset_index(inplace=True)
            final_df3.reset_index(inplace=True)
            final_df [['3G IP Address','3G VLAN ID','3G O&M IP Address' ,'3G O&M VLAN ID']] = final_df2[['3G IP Address','3G VLAN ID','3G O&M IP Address' ,'3G O&M VLAN ID']]
            final_df [['LTE IP Address', 'LTE VLAN ID','LTE  O&M IP Address','LTE O&M VLAN ID']] = final_df3[['LTE IP Address', 'LTE VLAN ID','LTE  O&M IP Address','LTE O&M VLAN ID']]

            
            return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))

 ### Kohkiloye Va boyer Ahmad New ###
    elif Province == 'Kohgiluyeh-new' :
        sheet_names = ['Yasouj New POA']
        list_of_files = glob.glob('Z:\IP Plans\Region 6&9\Kohkiloye Va boyer Ahmad\*.xlsx') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getmtime)
        df = pandas.DataFrame()
        for sheet in sheet_names:
            oo = pandas.read_excel(latest_file,sheet_name=sheet)
            df = pandas.concat([df,oo],ignore_index=True)

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
                final_df = pandas.concat([final_df , jj])  
            final_df.fillna("" , inplace=True)
        return HttpResponse(final_df.to_html(index=False,justify='center',col_space='150'))





#Test Index
def index3(request):
    import pandas
    list_sheetha = ['2G','3G']
    df = pandas.DataFrame()
    for i in list_sheetha:
        oo = pandas.read_excel('D:\Python\IPPLAN2.xlsx',sheet_name=i)
        df = pandas.concat([df,oo])
    return HttpResponse(df.to_html())


#1st Page HTML
def index4(request):
        return render(request,"Home2.html")
        


    


