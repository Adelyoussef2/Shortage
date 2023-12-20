import tkinter as tk
from tkinter import *
from PIL import ImageTk, Image
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import numpy as np



window = tk.Tk()
window.geometry("800x500")
window.title("Warehouse")
background_img = Image.open("Untitled.png")
icon = PhotoImage(file="Untitled.png")


window.iconphoto(False, icon)


label = Label(window, text="Hello, world!")
label.pack()


image = ImageTk.PhotoImage(background_img)

background_label = Label(window, image=image)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

file_path_1 = ''
file_path_2 = ''
file_path_3 = ''


def browse_file_1():
    global file_path_1
    file_path_1 = filedialog.askopenfilename()
    

def browse_file_2():
    global file_path_2
    file_path_2 = filedialog.askopenfilename()
   
def browse_file_3():
    global file_path_3
    file_path_3 = filedialog.askopenfilename()

def browse_file_4():
    global file_path_4
    file_path_4 = filedialog.askopenfilename()

def perform_analysis():
   
    pen1=pd.read_excel(file_path_1)
    sc1=pd.read_excel(file_path_2)
    ag1=pd.read_excel(file_path_3)
    ap=pd.read_excel(file_path_4)
    pen=pen1[['Repair Order No.','SC Code','SC Name','Engineer','Good Material Code','Good Material Subcategory','Marketing Model','Receive Date']]
    pen=pen.fillna(0)
    pen=pen1
    ag=ag1
    pen['Good Material Code']=pen['Good Material Code'].astype(str)
    pen['Con']=pen['SC Code']+pen['Good Material Code']
    pen['Receive Date']=pd.to_datetime(pen['Receive Date'])
    pen['Duration']=(pd.Timestamp.now().normalize() - pen['Receive Date'])/np.timedelta64(1,'D')
    pen['Con']=pen['Con'].astype(str)
    pen=pen[['Duration','Con','Repair Order No.','SC Code','SC Name','Engineer','Good Material Code','Good Material Subcategory','Marketing Model','Receive Date']]
    sc=sc1
    sc=sc[['Agent/SC code','Material code','Availible Qty for Good material']]
    sc.loc[:, 'Con'] = sc['Agent/SC code'] + sc['Material code']
    sc=sc[['Con','Availible Qty for Good material']]
    merge=pd.merge(pen,sc,on='Con',how='left')
    ag=ag[['Material code','Availible Qty for Good material']]
    ag=ag.rename(columns={'Material code':'Good Material Code','Availible Qty for Good material':'warehouse'})
    merge2=pd.merge(merge,ag,on='Good Material Code',how='left')
    merge2=merge2.fillna(0)
    ap=ap[['Applicant Code','Material code','Delivery Qty','Receive Time']]
    ap['Con']=ap['Applicant Code']+ap['Material code'].astype(str)
    ap['Receive Time']=pd.to_datetime(ap['Receive Time'])
    ap['days from receive']=(pd.Timestamp.now().normalize() - ap['Receive Time'])/np.timedelta64(1,'D')
    ap['days from receive'].fillna(0)
    ap=ap[['Con','Delivery Qty','days from receive']]
    ap=ap.drop_duplicates(['Con'])
    merge3=pd.merge(merge2,ap,on='Con',how='left')
    merge3=merge3.fillna(0)
    final=merge3[['SC Name','Engineer','SC Code','Repair Order No.','Good Material Code','Good Material Subcategory','Marketing Model','Duration','Availible Qty for Good material','warehouse','Delivery Qty','days from receive']]
    final=final.rename(columns={'Good Material Code':'Material code','Availible Qty for Good material':'SC Stock'})
    Row_data=final.set_index('SC Name')
    Total=final.drop_duplicates(['Repair Order No.'])
    Total=Total.pivot_table(index='SC Name',values='Repair Order No.',aggfunc='count',margins=True)
    Total=Total.sort_values('Repair Order No.', ascending=False)
    Total=Total.reset_index()
    agg=ag
    agg=agg.rename(columns={'Good Material Code': 'Material code'})
    sh1=final.loc[final['Good Material Subcategory']!='Auxiliary Material']
    sh1=sh1.pivot_table(index=['SC Name','Marketing Model','Good Material Subcategory','SC Code','Material code'],aggfunc={'Repair Order No.': 'count'})
    pt=sh1.sort_values('Repair Order No.', ascending=False)
    pt=pt.reset_index()
    pt['Con']=pt['SC Code']+pt['Material code']
    ptt=pd.merge(pt,sc,on='Con',how='left')
    pttt=pd.merge(ptt,agg,on='Material code',how='left')
    pttt['Need']=pttt['Repair Order No.']-pttt['Availible Qty for Good material']
    pttt['Result'] = np.where(ptt['Availible Qty for Good material'] <= ptt['Repair Order No.'], ptt['Availible Qty for Good material'], ptt['Repair Order No.'])
    tw=pttt[['Marketing Model','Good Material Subcategory','Material code','SC Name','Repair Order No.','Availible Qty for Good material','Need','warehouse']]
    tw1=tw.pivot_table(index='Material code',values='Need',aggfunc='sum')
    tw1=tw1.sort_values('Need', ascending=False)
    tw1=tw1.reset_index()
    tw2=pd.merge(tw,tw1,on='Material code',how='left')
    tw3=tw2
    tw3=tw3.rename(columns={'Good Material Subcategory': 'Subcategory','Availible Qty for Good material':'SC Stock','Repair Order No.':'Total Pending','Need_x':'SC need from warehouse','Need_y':'All SC need from warehouse'})
    Final_sheet_tw=tw3
    tw4=pttt[['SC Name','Marketing Model','Good Material Subcategory','Material code','Repair Order No.','Need','warehouse']]
    tw5=tw4.pivot_table(index='Material code',values='Need',aggfunc='sum')
    tw5=tw5.sort_values('Need', ascending=False)
    tw5=tw5.reset_index()
    tw6=pd.merge(tw4,tw5,on='Material code',how='left')
    tw7=tw6
    tw7=tw7.rename(columns={'Good Material Subcategory': 'Subcategory','Repair Order No.':'Total','Need_x':'Not in SC','Need_y':'Total need'})
    tw3['late']=np.where(tw3['SC need from warehouse']<=0,tw3['Total Pending'],tw3['Total Pending']-tw3['SC need from warehouse'])
    late=tw3.pivot_table(index='SC Name',values='late',aggfunc='sum')
    late=late.sort_values('late', ascending=False)
    late=late.reset_index()
    nf=final
    nf=nf.loc[(nf['warehouse']<=5)]
    nf=nf.pivot_table(index=['Marketing Model','Good Material Subcategory','Material code'],values='Repair Order No.',aggfunc='count')
    nf=nf.reset_index()
    nf=nf.sort_values('Repair Order No.', ascending=False)
    soft=pen1.drop_duplicates(['Repair Order No.'])
    soft=soft.loc[soft['Repair type']=='Upgrade']
    soft['Receive Date']=pd.to_datetime(soft['Receive Date'])
    soft['Duration']=(pd.Timestamp.now().normalize() - soft['Receive Date'])/np.timedelta64(1,'D')
    soft=soft[['Duration','Receive time','Repair Order No.','SC Name','Repair type']]
    cons=pen1.drop_duplicates(['Repair Order No.'])
    cons=cons.loc[cons['Repair type']=='Consultation']
    cons['Receive Date']=pd.to_datetime(cons['Receive Date'])
    cons['Duration']=(pd.Timestamp.now().normalize() - cons['Receive Date'])/np.timedelta64(1,'D')
    cons=cons[['Duration','Receive time','Repair Order No.','SC Name','Repair type']]
    tw=pttt[['Marketing Model','Good Material Subcategory','Material code','Repair Order No.','Availible Qty for Good material','Need','warehouse']]
    tw1=tw.pivot_table(index='Material code',values='Need',aggfunc='sum')
    tw1=tw1.sort_values('Need', ascending=False)
    tw1=tw1.reset_index()
    tw2=pd.merge(tw,tw1,on='Material code',how='left')
    tw3=tw2
    tw3=tw3.rename(columns={'Good Material Subcategory': 'Subcategory','Availible Qty for Good material':'SC Stock','Repair Order No.':'Total Pending','Need_x':'SC need from warehouse','Need_y':'All SC need from warehouse'})
    w=pttt
    w=w.pivot_table(index=['Good Material Subcategory','Marketing Model','Material code'],values='Repair Order No.',aggfunc='sum')
    w=w.sort_values('Repair Order No.', ascending=False)
    w=w.reset_index()
    s=sc1.pivot_table(index=['Material species','Model','Material code'],values='Availible Qty for Good material',aggfunc='sum')
    s=s.sort_values('Availible Qty for Good material', ascending=False)
    s=s.reset_index()
    s=s[['Material code','Availible Qty for Good material']]
    t=pd.merge(w,s,on='Material code',how='left')
    x=ag
    x=x.rename(columns={'Good Material Code':'Material code'})
    fb=pd.merge(t,x,on='Material code',how='left')
    fb=fb.rename(columns={'Repair Order No.':'Total Pending','Availible Qty for Good material':'Total SC Stock','warehouse':'Warehouse Stock'})
    fb=fb.loc[fb['Material code']!= 'nan']
    kl=sc1
    kl=kl.pivot_table(index='Material code',values='In-transit Qty for Good material',aggfunc='sum')
    kl=kl.sort_values('In-transit Qty for Good material',ascending=False)
    kl=kl.reset_index()
    kl=pd.merge(fb,kl,on='Material code',how='left')
    kl=kl.rename(columns={'In-transit Qty for Good material':'In-transit to SC'})
    int=ag1.pivot_table(index='Material code',values='In-transit Qty for Good material',aggfunc='sum')
    fbt=pd.merge(kl,int,on='Material code',how='left')
    fbt=fbt.rename(columns={'In-transit Qty for Good material':'In-transit From Cina'})
    lt1=final
    lt1=lt1.loc[lt1['Duration']<30]
    lt1=lt1.pivot_table(index='Material code',values='Repair Order No.',aggfunc='count')
    lt1=lt1.sort_values('Repair Order No.',ascending=False)
    lt1=lt1.reset_index()
    lt1=lt1.loc[lt1['Material code']!= 'nan']
    ft1=pd.merge(fbt,lt1,on='Material code',how='left')
    ft1=ft1.rename(columns={'Repair Order No.':'Less Than 1 Month'})
    lt2=final
    lt2=lt2.loc[lt2['Duration']>=30]
    lt2=lt2.pivot_table(index='Material code',values='Repair Order No.',aggfunc='count')
    lt2=lt2.sort_values('Repair Order No.',ascending=False)
    lt2=lt2.reset_index()
    lt2=lt2.loc[lt2['Material code']!= 'nan']
    ft2=pd.merge(ft1,lt2,on='Material code',how='left')
    ft2=ft2.rename(columns={'Repair Order No.':'More Than 1 Month'})
    jk=final
    jk=jk.loc[jk['Duration']>=60]
    jk=jk[['SC Name','Repair Order No.','Good Material Subcategory','Material code','Marketing Model','SC Stock','warehouse']]
    kl=pd.merge(jk,int,on='Material code',how='left')
    kl=kl.rename(columns={'In-transit Qty for Good material':'In transit from Cina'})

    
    
    save_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=(("Excel files", "*.xlsx"),))

  
    with pd.ExcelWriter(save_path) as writer:
        sheet_name = "Row_data"
        final.to_excel(writer, sheet_name=sheet_name, index=False)

        sheet_name = "Total"
        Total.to_excel(writer, sheet_name=sheet_name, index=False)
      
        sheet_name = "Late"
        late.to_excel(writer, sheet_name=sheet_name, index=False)

        sheet_name = "Final_sheet_tw"
        Final_sheet_tw.to_excel(writer, sheet_name=sheet_name, index=False)

        sheet_name = "Need from china "
        nf.to_excel(writer, sheet_name=sheet_name, index=False)

        sheet_name = "Consultation "
        cons.to_excel(writer, sheet_name=sheet_name, index=False)

        sheet_name = "Software "
        soft.to_excel(writer, sheet_name=sheet_name, index=False)
        
        sheet_name = "Pending Details "
        ft2.to_excel(writer, sheet_name=sheet_name, index=False)
       
        sheet_name = "More than 2 months "
        kl.to_excel(writer, sheet_name=sheet_name, index=False)







button_1 = tk.Button(window, text="Pending sheet", command=browse_file_1)
button_2 = tk.Button(window, text="SC Real time sheet", command=browse_file_2)
button_3 = tk.Button(window, text="Agent real time sheet", command=browse_file_3)
button_4 = tk.Button(window, text="SC Applications", command=browse_file_4)
button_5 = tk.Button(window, text="Start", command=perform_analysis)


button_1.pack()
button_1.place(x=150, y=30)
button_2.pack()
button_2.place(x=270, y=30)
button_3.pack()
button_3.place(x=410, y=30)
button_4.pack()
button_4.place(x=570, y=30)
button_5.pack()
button_5.place(x=380, y=410)

window.mainloop()
