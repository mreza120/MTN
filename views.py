from django.shortcuts import render , redirect

# Create your views here.

from django.http import HttpResponse
import pandas


def index(request):
    # print(request.POST['sitename'])
    # oo = pandas.read_excel('D:\Python\IPPLAN.xlsx')
    #getting the information from user
    x = request.POST.get('sitename')
    Province = request.POST.get('province_name')
    #Kerman IP Check
    if x.isalpha():
        return HttpResponse("its not Valid")
    #Kerman IP Check
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

    elif Province == 'Kerman' :
        oo = pandas.read_excel('D:\Python\MTNi_Kerman_SBTS_IPDatafill_V.7.6(2020-05-17).xlsx',sheet_name='SBTS_IP_Datafill_Nokia_MW')
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
            # final_df3 = final_df3.drop('Hubsite')
            
            return HttpResponse(final_df3.to_html(index=False,justify='center',col_space='150'))

    elif Province == 'Yazd' :
        oo = pandas.read_excel('D:\Python\MTNi_Kerman_SBTS_IPDatafill_V.7.6(2020-05-17).xlsx',sheet_name='SBTS_IP_Datafill_Nokia_MW')
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




def index3(request):
    import pandas
    list_sheetha = ['2G','3G']
    df = pandas.DataFrame()
    for i in list_sheetha:
        oo = pandas.read_excel('D:\Python\IPPLAN2.xlsx',sheet_name=i)
        df = pandas.concat([df,oo])
    return HttpResponse(df.to_html())

from django.shortcuts import render
def index4(request):
        return render(request,"Home2.html")
        


    


