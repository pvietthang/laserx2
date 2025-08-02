import datetime
import socket
import requests,json
import time
import warnings
from datetime import date
from tkinter import *
from tkinter import messagebox, ttk
from tkinter import filedialog
import pandas as pd
import pyautogui as py
import pyperclip
from fiprequest import *
import traceback
from ERROR import *
warnings.filterwarnings("ignore")

def disable_event():
    pass
def detect_IR(Ma:str):
    if Ma.find("C0I") == -1:
        return True
    return False
def synchronized():
    try:
        filetypes = (
            ('text files', '*.xlsx'),
            ('All files', '*.*')
        )

        filenames = filedialog.askopenfilenames(
            title='Open files',
            initialdir='/',
            filetypes=filetypes)
        data = pd.read_excel(filenames[0])
        for i, j in zip(data['Casting code'], data['NGÀY ĐÚC']):
            synchronized_data(str(i.replace('-', '')), str(j))
        messagebox.showinfo("ĐỒNG BỘ", "Đồng Bộ Dữ Liệu Thành Công")
    except:
        Systemp_log(traceback.format_exc()).append_new_line()
        messagebox.showinfo("LỖI HỆ THỐNG", "Đồng Bộ Dữ Liệu Lỗi")

def malo_str(malo):
    try:
        if 'A2312' in CB_mahang.get() or 'A2502135' in CB_mahang.get():
            return True
        if malo[0].isnumeric():
            year = '202'+malo[0]
        else:
            messagebox.showerror('Lỗi cú pháp','Mã năm sai quy định')
            return False
        if malo[1].upper() == 'X':
            month = '10'
        elif malo[1].upper() == 'Y':
            month = '11'
        elif malo[1].upper() == 'Z':
            month = '12'
        elif malo[1].isnumeric():
            month = '0'+malo[1]
        else:
            messagebox.showerror('Lỗi cú pháp','Mã tháng sai quy định')
            return False
        if malo[2:4].isnumeric():
            day = malo[2:4]
        else:
            messagebox.showerror('Lỗi cú pháp','Mã ngày sai quy định')
            return False
        if not malo[4].isalpha():
            messagebox.showerror('Lỗi cú pháp','Nhập sai ký tự')
            return False
        if not malo[5:7].isnumeric():
            messagebox.showerror('Lỗi cú pháp','Số mẻ đúc không hợp lệ')
            return False
        try:
            date = datetime.datetime.strptime(day+month+year,'%d%m%Y')
            if datetime.datetime.now()>date:
                return True
            else:
                messagebox.showerror('Lỗi ngày tháng','Mã ngày lớn hơn ngày hiện tại')
                return False
        except:
            messagebox.showerror('Lỗi ngày tháng','Ngày tháng không hợp lệ')
            return False
    except:
        Systemp_log(traceback.format_exc()).append_new_line()
def check_dmc2():
    def NG_pcs():
        def save_ng():
            try:
                if CB_ng_reason.get() == '':
                    messagebox.showerror('Lỗi hệ thống','Chưa lựa chọn nguyên nhân NG')
                    return False
                global rework_flag
                TimeBarcode = datetime.datetime.now()
                if rework_flag:
                    savedata(ET_tennv.get(), machine_no, mahang, maqr, ET_dmccheck.get(), rwdmcin, TimeDMCstart, TimeDMCFinish,TimeBarcode, 'NG', maloi, chatluong)
                    rework_flag = False
                    LB_dmccheck.configure(text="Kết quả:\n" + 'NG', bg='#bb0000')
                    tk_check_dmc.destroy()
                    tk_check_dmc.quit()
                    tk_rework.grab_set()
                else:
                    savedata(ET_tennv.get(), machine_no, mahang, maqr, ET_dmccheck.get(), '', TimeDMCstart, TimeDMCFinish,TimeBarcode, 'NG', CB_ng_reason.get(), chatluong)
                    LB_dmccheck.configure(text="Kết quả:\n" + 'NG', bg='#bb0000')
                    tk_check_dmc.destroy()
                    tk_check_dmc.quit()                
            except:
                Systemp_log(traceback.format_exc()).append_new_line()
        try:
            tk_ng_reason = Toplevel(tk_check_dmc)
            tk_ng_reason.geometry('300x100')
            tk_ng_reason.grab_set()
            CB_ng_reason = ttk.Combobox(tk_ng_reason, state='readonly', values=get_status(), font=(font_style, font_size))
            CB_ng_reason.place(x=10, y=10, width=280, height=30)
            Btn_confirm_ng = Button(tk_ng_reason,text='Xác nhận',command=save_ng).place(x=200,y=50)
        except:
            Systemp_log(traceback.format_exc()).append_new_line()
    def check_kq(i):
        try:
            global rework_flag
            if len(ET_dmccheck.get()) == 29 or len(ET_dmccheck.get()) == 25 or len(ET_dmccheck.get()) == 26 or len(ET_dmccheck.get()) == 27 or len(ET_dmccheck.get()) == 28 or len(ET_dmccheck.get()) == 39:
                if ET_dmccheck.get() == maqr:
                    TimeBarcode = datetime.datetime.now()
                    if rework_flag:
                        savedata(ET_tennv.get(),machine_no,mahang,maqr,ET_dmccheck.get(),rwdmcin,TimeDMCstart,TimeDMCFinish,TimeBarcode,'OK',maloi,chatluong)
                        rework_flag = False
                        LB_dmccheck.configure(text="Kết quả:\n"+'OK',bg='#00bb00')
                        tk_check_dmc.destroy()
                        tk_check_dmc.quit()
                        tk_rework.grab_set()
                    else:
                        savedata(ET_tennv.get(),machine_no,mahang,maqr,ET_dmccheck.get(),'',TimeDMCstart,TimeDMCFinish,TimeBarcode,'OK','',chatluong)
                        LB_dmccheck.configure(text="Kết quả:\n"+'OK',bg='#00bb00')
                        tk_check_dmc.destroy()
                        tk_check_dmc.quit()
                else:
                    if messagebox.askretrycancel('Mã không khớp','Sai mã, bạn có muốn thử lại?'):
                        ET_dmccheck.delete(0,END)
                    else:
                        TimeBarcode = datetime.datetime.now()
                        if rework_flag:
                            savedata(ET_tennv.get(),machine_no,mahang,maqr,ET_dmccheck.get(),rwdmcin,TimeDMCstart,TimeDMCFinish,TimeBarcode,'NG',maloi,chatluong)
                            rework_flag = False
                            LB_dmccheck.configure(text="Kết quả:\n" + 'NG', bg='#bb0000')
                            tk_check_dmc.destroy()
                            tk_check_dmc.quit() 
                            tk_rework.grab_set()
                        else:
                            savedata(ET_tennv.get(),machine_no,mahang,maqr,ET_dmccheck.get(),'',TimeDMCstart,TimeDMCFinish,TimeBarcode,'NG','Mã không khớp',chatluong)
                            LB_dmccheck.configure(text="Kết quả:\n" + 'NG', bg='#bb0000')
                            tk_check_dmc.destroy()
                            tk_check_dmc.quit() 
            else:
                if messagebox.askretrycancel('Sai mã','Mã sai số lượng, bạn có muốn thử lại?'):
                    ET_dmccheck.delete(0,END)
                else:
                    TimeBarcode = datetime.datetime.now()
                    if rework_flag:
                        savedata(ET_tennv.get(),machine_no,mahang,maqr,ET_dmccheck.get(),rwdmcin,TimeDMCstart,TimeDMCFinish,TimeBarcode,'NG',maloi,chatluong)
                        rework_flag = False
                        tk_rework.grab_set()
                    else:
                        savedata(ET_tennv.get(),machine_no,mahang,maqr,ET_dmccheck.get(),'',TimeDMCstart,TimeDMCFinish,TimeBarcode,'NG','Mã sai số lượng',chatluong)
                    LB_dmccheck.configure(text="Kết quả:\n" + 'NG', bg='#bb0000')
                    tk_check_dmc.destroy()
                    tk_check_dmc.quit()
        except:
            Systemp_log(traceback.format_exc()).append_new_line()
    try:
        tk_check_dmc = Toplevel(wk)
        tk_check_dmc.grab_set()
        tk_check_dmc.geometry('285x100+120+200')
        tk_check_dmc.protocol("WM_DELETE_WINDOW", disable_event)
        tk_check_dmc.configure(background='#add8e6')
        LB_dmccheck = Label(tk_check_dmc,bg='#D9D9D9', text = "DMC lỗi",font=(font_style,font_size),highlightbackground="#41719C", highlightcolor="#41719C",highlightthickness=1,anchor=W,width=19,height=2)
        LB_dmccheck.place(x=10,y=10, width=50, height=30)
        ET_dmccheck  = Entry(tk_check_dmc,font=(font_style,font_size+2),width=29)
        ET_dmccheck.place(x=60,y=10,width=280,height=30)
        # ET_dmccheck  = Entry(tk_check_dmc,font=(font_style,font_size),width=29)
        # ET_dmccheck.place(x=10,y=10)
        ET_dmccheck.focus()
        Btn_NG = Button(tk_check_dmc,text='NG',bg='#f47147',font=(font_style,font_size+5,'bold'),command=NG_pcs)
        Btn_NG.place(x=150,y=40,height=50,width=50)
        ET_dmccheck.bind('<Return>',check_kq)
        tk_check_dmc.mainloop()
    except:
        Systemp_log(traceback.format_exc()).append_new_line()
def checkmalo(Malo):
    try:
        if offline_mode:
            return True
        if check_castingname(Malo.replace('-', '')[:7]) == '"True"':
            return True
        else:
            messagebox.showerror("Lỗi mã","Mã không tồn tại")
            return False
    except:
        Systemp_log(traceback.format_exc()).append_new_line()
def MAQR(serial,Malo,Mabanve,PhienBan,Productname,xulynhiet):
    try:
        try:
            if int(Malo[7:]) >= 10:
                Makhuonsap = Malo[7:]
            else:
                Makhuonsap = "0" + Malo[7:]
        except:
            Makhuonsap = ""
        if PhienBan=="IR":
            Maphienban=" "
        else:
            Maphienban=PhienBan
        if Productname=="A2012003TQ":
            MaSX="NW"
        elif Productname=="A2012003VC" or Productname=="A2012004" or Productname=="A2107024":
            MaSX="VC"
        else:
            MaSX=""
        mangay = datetime.datetime.now().strftime('%j')
        # if x >= dt1 and x <= dt2:
        #     mangay=str((dt1 - dt).days+1)
        # else:
        #    mangay=str((dt2 - dt).days+1)
        nam=datetime.datetime.now().strftime("%y")
        #print(type(mangay))
        Malo = Malo.upper()
        if 'A2307075' in Productname:
            Qrcode = Mabanve[:10]+mangay+nam+Malo[:4]+'-'+Malo[4:]+serial
            Qrcode = Qrcode.upper()
            print(Qrcode)
            if len(Qrcode) != 27:
                messagebox.showerror('Lỗi DMC','Mã DMC sai số lượng')
                return 'Error'
        elif 'A210' in Productname:
            Qrcode = '555241244'+Mabanve+Maphienban+nam+mangay+serial+Malo[:4]+'-'+Malo[4:]
            print(Qrcode)
            if len(Qrcode) != 39:
                messagebox.showerror('Lỗi DMC','Mã DMC sai số lượng')
                return 'Error'
        elif 'A2310084' in Productname:
            Qrcode=mangay+nam+serial+Mabanve[:1]+Mabanve[-7:]+Maphienban+Malo[:7]+xulynhiet
            Qrcode = Qrcode.upper()
            print(Qrcode)
            if len(Qrcode) != 26:
                messagebox.showerror('Lỗi DMC','Mã DMC sai số lượng')
                return 'Error'
        elif 'A2312' in Productname or 'A2502' in Productname:
            if len(ET_gcdate.get()) != 6:
                messagebox.showerror('Lỗi ngày tháng','Ngày tháng không hợp lệ')
                return 'Error'
            Qrcode = Mabanve+Maphienban+Malo+ET_gcdate.get()+serial[-3:]
            Qrcode = Qrcode.upper() 
            print(Qrcode)
            if len(Qrcode) != 28:
                messagebox.showerror('Lỗi DMC','Mã DMC sai số lượng')
                return 'Error'
        else:
            Qrcode=mangay+nam+serial+Mabanve[:1]+Mabanve[-6:]+Maphienban+MaSX+Makhuonsap+Malo[:7]+xulynhiet
            Qrcode = Qrcode.upper()
            print(Qrcode)
            if len(Qrcode) != 29:
                messagebox.showerror('Lỗi DMC','Mã DMC sai số lượng')
                return 'Error'
        
        return Qrcode
    except:
        Systemp_log(traceback.format_exc()).append_new_line()
def offline_mode_tog():
    try:
        global offline_mode
        offline_mode = not offline_mode
        if offline_mode:
            Btn_test.configure(text="Offline",bg="#FFFF00")
        else:
            Btn_test.configure(text="Online",bg="#00ff00")
    except:
        Systemp_log(traceback.format_exc()).append_new_line()
def login():
    def set_security_level():
        try:
            global security_level
            global user_name
            try:
                user_data = get_user(ET_password.get())
                user_name = user_data[0]
                security_level = int(user_data[1])
                messagebox.showinfo("Đăng nhập", "Thành công")
                tk_login.destroy()
                tk_login.quit()
            except:
                security_level = 0
                messagebox.showerror("Đăng nhập","Sai mật khẩu")
                tk_login.destroy()
                tk_login.quit()
        except:
            Systemp_log(traceback.format_exc()).append_new_line()
    try:
        tk_login = Toplevel(wk)
        tk_login.title("Đăng nhập")
        tk_login.geometry("250x90+800+400")
        tk_login.grab_set()
        tk_login.protocol("WM_DELETE_WINDOW", disable_event)
        tk_login.configure(bg='#FFFFFF')
        LB_password = Label(tk_login, text= "Mật khẩu", font=(font_style,font_size), bg='#ffffff', anchor=W)
        LB_password.place(x=15,y=10, width=70,height=30)
        ET_password = Entry(tk_login, font=(font_style,font_size),show="*")
        ET_password.place(x=85,y=10,width=150,height=30)
        Btn_login = Button(tk_login, text="Xác nhận", command=set_security_level)
        Btn_login.place(x=85, y=50, width=80,height=30)
        tk_login.mainloop()
        return security_level
    except:
        Systemp_log(traceback.format_exc()).append_new_line()
        return 0
def str4(n): # sửa chuỗi có ít hơn 4 ký tự
    try:
        n = str(n)
        while len(n)<4:
            n = '0' + n
        return n
    except:
        Systemp_log(traceback.format_exc()).append_new_line()
def calibPos(): # chỉnh sửa vị trí click
    def savePos(): # lưu vị trí sau khi chỉnh sửa
        try:
            global  click_pos
            list_pos.to_csv(file_position)
            click_pos = pd.read_csv(file_position, index_col=0)
            tk_calib.destroy()
            tk_calib.quit()
        except:
            Systemp_log(traceback.format_exc()).append_new_line()
    def curPos():
        while True:
            try:
                x,y = py.position()
                time.sleep(1)
                if (x,y) == py.position():
                    list_pos.loc[CB_list_pos.get(),'X'] = x
                    list_pos.loc[CB_list_pos.get(),'Y'] = y
                    break
            except:
                Systemp_log(traceback.format_exc()).append_new_line()
    try:
        list_pos = pd.read_csv(file_position,index_col=0)
        tk_calib = Toplevel(wk)
        tk_calib.geometry("120x180")
        tk_calib.grab_set()
        tk_calib.resizable(False,False)
        
        CB_list_pos = ttk.Combobox(tk_calib,state='readonly',font=font_size,values=pd.read_csv(file_position,index_col=0).index.to_list()[1:])
        CB_list_pos.place(x=5,y=5,width=100,height=30)
        # CB_list_pos.bind("<<ComboBoxSelected>>",curPos)
        btn2 = Button(tk_calib,text="Change",command=curPos,activebackground="red").place(x=5,y=40,width=100,height=30)
        btn5 = Button(tk_calib,text="Save",command=savePos).place(x=5,y=145,width=100,height=30)
        tk_calib.mainloop()
    except:
        Systemp_log(traceback.format_exc()).append_new_line()
def clicked(maqr): # auto click theo tọa độ
    try:
        global flag_working
        flag_working = True
        pyperclip.copy(maqr)
        click_pos = pd.read_csv('click_position.csv',index_col=0)
        time.sleep(0.1)
        # py.click(int(click_pos.loc['Text','X']),int(click_pos.loc['Text','Y']))
        # py.rightClick(int(click_pos.loc['Text','X']),int(click_pos.loc['Text','Y']))
        py.rightClick(int(click_pos.loc['Text','X']),int(click_pos.loc['Text','Y']))
        # py.doubleClick(int(click_pos.loc['Text','X']),int(click_pos.loc['Text','Y']))
        # py.move(70,140)
        time.sleep(0.4)
        py.press('a')
        time.sleep(0.1)
        py.hotkey('ctrl','v')
        time.sleep(0.3)
        py.click(int(click_pos.loc['Apply','X']),int(click_pos.loc['Apply','Y']))
        time.sleep(0.3) 
        py.press('f2')
        time.sleep(0.1)
        ET_malo.delete(0, END)
        pyperclip.copy('')
        py.click(int(click_pos.loc['Mã Lò','X']),int(click_pos.loc['Mã Lò','Y']))
    except:
        Systemp_log(traceback.format_exc()).append_new_line()
def again(): # update Label thời gian
    click_pos = pd.read_csv(file_position, index_col=0)
    try:
            # time.sleep(1)
            global today
            today = datetime.datetime.now()
            nextday = today + datetime.timedelta(days=1)
            strtoday = today.strftime("%Y-%m-%d")
            strnextday = nextday.strftime("%Y-%m-%d")
            LB_ng.configure(text="Số lượng khắc\nNG của máy:\n"+count_history(machine_no,strtoday,strnextday,'NG'))
            LB_ok.configure(text="Số lượng khắc\nOK của máy:\n"+count_history(machine_no,strtoday,strnextday,'OK'))
            def update_top10_data():
                try:
                    row1=laser_result(machine_no,mahang,'OK')
                    row2=laser_result(machine_no,mahang,'NG')
                    trv.delete(*trv.get_children())
                    trv1.delete(*trv1.get_children())
                    for i in range(len(row1)):
                        trv.insert(parent='', index=i, values=(row1['DMCin'][i], row1['Quality'][i], str(row1['TimeOutBarcode'][i]).replace('T',' ')))
                    for i in range(len(row2)):
                        trv1.insert(parent='', index=i, values=(row2['DMCin'][i], row2['Quality'][i], str(row2['TimeOutBarcode'][i]).replace('T',' ')))
                except:
                    Systemp_log(traceback.format_exc()).append_new_line()
            global serial_no
            global LB_serial
            global isWorking
            global TimeDMCFinish
            global TimeDMCstart
            global TimeBarcode
            global savedone
            global flag_working
            if CB_mahang.get() != '':
                serial_no = str4(getserial(CB_mahang.get()))
            r,g,b = py.pixel(int(click_pos.loc['Check','X']),int(click_pos.loc['Check','Y']))
            if savedone:
                click_pos = pd.read_csv('click_position.csv',index_col=0)
                #py.rightClick(int(click_pos.loc['Text','X']),int(click_pos.loc['Text','Y']))
                #time.sleep(0.4)
                #py.press('a')
                #time.sleep(0.1)
                #py.write('A')
                #time.sleep(0.3)
                #py.click(int(click_pos.loc['Apply','X']),int(click_pos.loc['Apply','Y']))
                time.sleep(0.1) 
                py.click(int(click_pos.loc['Mã Lò','X']),int(click_pos.loc['Mã Lò','Y']))
                savedone = False
            if r==240 and g==240 and b==240 and flag_working:        
                isWorking = True
                print('wait')
            if isWorking:
                if py.pixel(int(click_pos.loc['Check','X']),int(click_pos.loc['Check','Y'])) == (255,255,255):
                    TimeDMCFinish = datetime.datetime.now()
                    print('done')
                    check_dmc2()
                    update_top10_data()
                    isWorking = False
                    flag_working = False
                    savedone = True
            # serial_no = str4(stt_count)
            LB_serial.configure(text = "Serial No. " + serial_no)
            nowtime.configure(text=f"{datetime.datetime.now():%d/%m/%Y %H:%M:%S}")

    except :
        Systemp_log(traceback.format_exc()).append_new_line()
    nowtime.after(1000, again)


def popup_data(): # hiện thị cửa sổ danh sách dữ liệu
    def export_excel():
        try:
            foder=filedialog.askdirectory()
            # filename=foder+'/QC_report_'+datetime.datetime.now().strftime('%y_%m_%d-%H_%M')+'.xlsx'
            filename=foder+'/QC_report_'+datetime.datetime.now().strftime('%y_%m_%d-%H_%M')+'.xlsx'
            # filename=r'filename
            # print(data1)
            data1.to_excel(filename,index=False)
            messagebox.showinfo('Xuất file','Xuất thành công')
        except:
            Systemp_log(traceback.format_exc()).append_new_line()
    def showdata():
        try:
            global data1
            Productname= productname.get()
            machinename= Machinename.get()
            name=Name.get()
            data1=laser_all_data(machinename,name,Productname,start_time.get(),end_time.get())
            trv2.delete(*trv2.get_children())
            for i in range(0, len(data1)):
                trv2.insert(parent='', index=0,values=(data1['MachineNo'][len(data1)-1 - i], data1['NameOperator'][len(data1)-1 - i], data1['NameProduct'][len(data1)-1 - i],data1['DMCin'][len(data1)-1 - i],data1['DMCout'][len(data1)-1 - i],str(data1['TimeInDMC'][len(data1)-1 - i]).replace('T',' '),str(data1['TimeOutBarcode'][len(data1)-1 - i]).replace('T',' '),data1['Result'][len(data1)-1 - i],data1['Quality'][len(data1)-1 - i]))
        except:
            Systemp_log(traceback.format_exc()).append_new_line()
    try:
        top = Toplevel(wk)
        top.geometry("1600x800")
        top.state('zoomed')
        check1 = StringVar()
        check2 = StringVar()
        check3 = StringVar()
        wrapper1 = LabelFrame(top, text="OK", bg='#2596be', width=1600, height=800)
        wrapper1.place(x=0, y=0)
        trv2 = ttk.Treeview(wrapper1,selectmode='browse', columns=(1,2, 3,4,5,6,7,8,9), show="headings", height="25")
        trv2.place(x=0, y=0)

        trv2.column(1, anchor=CENTER, width=90)
        trv2.column(2, anchor=CENTER, width=190)
        trv2.column(3, anchor=CENTER, width=150)
        trv2.column(4, anchor=CENTER, width=210)
        trv2.column(5, anchor=CENTER, width=210)
        trv2.column(6, anchor=CENTER, width=250)
        trv2.column(7, anchor=CENTER, width=250)
        trv2.column(8, anchor=CENTER, width=90)
        trv2.column(9, anchor=CENTER, width=90)


        trv2.heading(1, text="Machineno")
        trv2.heading(2, text="NameOperator")
        trv2.heading(3, text="NameProduct")
        trv2.heading(4, text="DMC_in")
        trv2.heading(5, text="DMC_out")
        trv2.heading(6, text="Time_DMCin")
        trv2.heading(7, text="Time_DMCout")
        trv2.heading(8, text="Result")
        trv2.heading(9, text="Quality")

        checkbox = Checkbutton(top, text="Productname", font=("Times New Roman", 12, "bold"), variable=check1)
        checkbox.deselect()
        checkbox.place(x=80, y=545)

        checkbox = Checkbutton(top, text="Machineno", font=("Times New Roman", 12, "bold"), variable=check2)
        checkbox.deselect()
        checkbox.place(x=80, y=600)

        checkbox = Checkbutton(top, text="Name", font=("Times New Roman", 12, "bold"), variable=check3)
        checkbox.deselect()
        checkbox.place(x=80, y=655)

        productname = ttk.Combobox(top, width=27, values=('A2012003VC','A2012003TQ', 'A2012004','A2107024', 'A2303121','A2303123','A2307075',''),state = 'readonly')
        productname.place(x=200, y=550)

        Machinename = ttk.Combobox(top, width=27,values=('MÁY 1', 'MÁY 2','MÁY 3','MÁY 4','MÁY 5',''),state = 'readonly')
        Machinename.place(x=200, y=600)

        Name = Entry(top, width=27)
        Name.place(x=200, y=660)

        button = Button(top, text="Show_all", font=("Times New Roman", 16, "bold"), command=showdata)
        button.place(x=500, y=560, width=140, height=30)

        start_time = Entry(wrapper1, width=13, font=('calibre', 16, 'normal'))
        start_time.place(x=700, y=545)
        end_time = Entry(wrapper1, width=13, font=('calibre', 16, 'normal'))
        end_time.place(x=700, y=590)
        btn_export_excel = Button(wrapper1,text='Xuat du lieu', font=("Times New Roman", 16, "bold"),command=export_excel).place(x=880,y=590,height=30)
    except:
        Systemp_log(traceback.format_exc()).append_new_line()
def rework(): # nút sửa mã
    def fix_dmc():
        try:
            if CB_mahang.get() == '':
                messagebox.showerror("Lỗi hệ thống","Chưa chọn sản phẩm")
            elif isWorking == True or savedone == True:
                messagebox.showerror("Lỗi hệ thống","Chưa kiểm tra DMC")
            elif len(ET_malorw.get())<8 or len(ET_malorw.get())>9:
                messagebox.showerror("Lỗi cú pháp","Mã sai số lượng")
            elif not ET_malorw.get()[7:].isnumeric():
                messagebox.showerror("Lỗi mã","Mã khuôn sáp không hợp lệ")
            elif int(ET_malorw.get()[7:]) > int(getwax(CB_mahang.get())):
                messagebox.showerror("Lỗi mã","Mã khuôn sáp không tồn tại")
            else:
                if checkmalo(ET_malorw.get()) and malo_str(ET_malorw.get()):
                    global maqr
                    global TimeDMCstart
                    global TimeBarcode
                    global rwdmcin
                    rwdmcin = ET_DMCrework.get()
                    TimeDMCstart = datetime.datetime.now()
                    mabanve = pd.read_csv(file_serial,index_col=0).loc[CB_mahang.get(),'Model']
                    print(mabanve)
                    Phienban = pd.read_csv(file_serial,index_col=0).loc[CB_mahang.get(),'Phiên Bản']
                    xulynhiet = pd.read_csv(file_serial,index_col=0).loc[CB_mahang.get(),'Xử Lý Nhiệt']
                    maqr = MAQR(str4(update_serial(CB_mahang.get())),ET_malorw.get(),mabanve,Phienban,CB_mahang.get(),xulynhiet)
                    if duplicate(maqr,mahang) != 0:
                        if not messagebox.askyesno('Cảnh báo trùng mã','Đã tồn tại mã cùng STT, có muốn tiếp tục hay không'):
                            return False
                    ET_DMCrework.delete(0, END)
                    ET_malorw.delete(0,END)
                    clicked(maqr)
        except:
            Systemp_log(traceback.format_exc()).append_new_line()
    def rework_dmc():
        try:
            global maqr
            global TimeDMCstart
            global rework_flag
            global maloi
            maloi = CB_maloi.get()
            rework_flag = True
            if int(type_error(CB_maloi.get())) == 1:
                if len(ET_DMCrework.get()) == 29 or len(ET_DMCrework.get()) == 25:
                    global rwdmcin
                    rwdmcin = ET_DMCrework.get()
                    maqr = ET_DMCrework.get()
                    TimeDMCstart = datetime.datetime.now()
                    clicked(maqr)
                else:
                    messagebox.showerror('Lỗi mã DMC','Mã DMC không đúng số lượng')
                    return
            elif int(type_error(CB_maloi.get())) == 2:
                if len(ET_DMCrework.get()) == 0:
                    messagebox.showerror('Lỗi mã','Scan DMC lỗi trước')
                    return
                fix_dmc()
            elif int(type_error(CB_maloi.get())) == 3:
                fix_dmc()
        except:
            Systemp_log(traceback.format_exc()).append_new_line()
    def check_error_code(i):
        try:
            if int(type_error(CB_maloi.get())) == 1:
                ET_malorw.delete(0, END)
                ET_DMCrework.delete(0, END)
                ET_malorw.configure(state='readonly')
                ET_DMCrework.configure(state='normal')
            elif int(type_error(CB_maloi.get())) == 2:
                ET_malorw.delete(0, END)
                ET_DMCrework.delete(0, END)
                ET_DMCrework.configure(state='normal')
                ET_malorw.configure(state='normal')
            else:
                ET_malorw.delete(0, END)
                ET_DMCrework.delete(0, END)
                ET_DMCrework.configure(state='readonly')
                ET_malorw.configure(state='normal')    
        except:
            Systemp_log(traceback.format_exc()).append_new_line() 
    try:  
        global tk_rework
        tk_rework = Toplevel(wk)
        tk_rework.grab_set()
        tk_rework.geometry('350x190+120+200')
        tk_rework.configure(background='#add8e6')
        LB_dmcloi = Label(tk_rework,bg='#D9D9D9', text = "DMC lỗi",font=(font_style,font_size),highlightbackground="#41719C", highlightcolor="#41719C",highlightthickness=1,anchor=W,width=19,height=2)
        LB_dmcloi.place(x=10,y=10, width=50, height=30)
        ET_DMCrework  = Entry(tk_rework,state='readonly',font=(font_style,font_size+2),width=29)
        ET_DMCrework.place(x=60,y=10,width=280,height=30)
        ET_DMCrework.focus()
        LB_malorw = Label(tk_rework,bg='#D9D9D9', text = "Mã lò",font=(font_style,font_size),highlightbackground="#41719C", highlightcolor="#41719C",highlightthickness=1,anchor=W,width=19,height=2)
        LB_malorw.place(x=10,y=50, width=50, height=30)
        ET_malorw  = Entry(tk_rework,state='readonly',font=(font_style,font_size+2),width=29)
        ET_malorw.place(x=60,y=50,width=280,height=30)
        LB_maloi = Label(tk_rework,bg='#D9D9D9', text = "Mã lỗi",font=(font_style,font_size),highlightbackground="#41719C", highlightcolor="#41719C",highlightthickness=1,anchor=W,width=19,height=2)
        LB_maloi.place(x=10,y=90, width=50, height=30)
        CB_maloi = ttk.Combobox(tk_rework,state='readonly',values=get_status(),font=(font_style,font_size))
        CB_maloi.place(x=60,y=90,width=280,height=30)
        CB_maloi.bind("<<ComboboxSelected>>", check_error_code)
        Btn_NG = Button(tk_rework,text='Rework',bg='#f47147',font=(font_style,font_size+2,'bold'),command=rework_dmc)
        Btn_NG.place(x=150,y=130,height=50,width=80)
        # ET_dmccheck.bind('<Return>',check_kq)
        tk_rework.mainloop()
    except:
        Systemp_log(traceback.format_exc()).append_new_line()
def setup_data(): # nút thiết lập
    if login() > 1:
        def save_setup():
            try:
                mahang = CB_mhsetup.get()
                f = pd.read_csv(file_serial,index_col=0)
                if messagebox.askyesno('Lưu thiết lập','Bạn có muốn đổi thông số của mã hàng ' + CB_mhsetup.get() + ':\nMã Bản Vẽ thành ' + ET_MaBanVe.get() +'\nPhiên Bản thành ' + ET_PhienBan.get() + ' không?'):
                    mabanvesau = ET_MaBanVe.get()
                    phienbansau = ET_PhienBan.get()            
                    if str(dmc_setup_history(datetime.datetime.now(),user_name,CB_mhsetup.get(),mabanvetruoc,mabanvesau,phienbantruoc,phienbansau)).strip() == '"OK"':
                        f.loc[mahang,'Model'] = ET_MaBanVe.get()
                        f.loc[mahang,'Phiên Bản'] = ET_PhienBan.get()
                        messagebox.showinfo('Lưu thiết lập','Thay đổi thiết lập thành công')
                    else:
                        messagebox.showerror('Lưu thiết lập','Lưu thiết lập thất bại')
                    f.loc[mahang,'Xử Lý Nhiệt'] = ET_XuLyNhiet.get()  
                f.to_csv(file_serial)
                tk_setup.destroy()
                tk_setup.quit()
            except:
                Systemp_log(traceback.format_exc()).append_new_line()
        def get_setup(self):
            try:
                mahang = CB_mhsetup.get()
                global mabanvetruoc
                global phienbantruoc
                mabanvetruoc = pd.read_csv(file_serial,index_col=0).loc[mahang,'Model']
                phienbantruoc = pd.read_csv(file_serial,index_col=0).loc[mahang,'Phiên Bản']
                ET_MaBanVe.delete(0,END)
                ET_MaBanVe.insert(0,mabanvetruoc)
                ET_PhienBan.delete(0,END)
                ET_PhienBan.insert(0,phienbantruoc)
                ET_XuLyNhiet.delete(0,END)
                ET_XuLyNhiet.insert(0,pd.read_csv(file_serial,index_col=0).loc[mahang,'Xử Lý Nhiệt'])
            except:
                Systemp_log(traceback.format_exc()).append_new_line()
        def delete_setup():
            f = pd.read_csv(file_serial,index_col=0)
            f = f.drop(CB_mhsetup.get())
            f.to_csv(file_serial)
            CB_mhsetup.config(values=pd.read_csv(file_serial, index_col=0).index.to_list())         
        def rename_setup():
            def save_rename():
                f = pd.read_csv(file_serial,index_col=0)
                f = f.rename(index={CB_mhsetup.get():ET_rename.get()})
                f.to_csv(file_serial)
                rename_box.destroy()
                CB_mhsetup.config(values=pd.read_csv(file_serial, index_col=0).index.to_list())
            rename_box = Toplevel(wk)
            ET_rename = Entry(rename_box, font=font_size)
            ET_rename.place(x=60,y=5,width=100,height=25)
            ET_rename.insert(0,CB_mhsetup.get())
            Btn_save_rename = Button(rename_box, text="Save", font=(font_style,font_size-1),command=save_rename)
            Btn_save_rename.place(x=60,y=35,width=60,height=40)
        def history_setup():
            try:
                tk_history_setup = Toplevel(tk_setup)
                tk_history_setup.geometry("786x600")
                tk_history_setup.resizable(False,False)
                tk_history_setup.grab_set()
                his = ttk.Treeview(tk_history_setup, columns=(1,2,3,4,5,6,7), show="headings", height="12")
                his.place(x=3, y=0,height=600)
                style = ttk.Style()
                style.configure("Treeview.treearea", font=font_size+5)
                his.column(1, anchor=CENTER, width=120)
                his.column(2, anchor=CENTER, width=150)
                his.column(3, anchor=CENTER, width=90)
                his.column(4, anchor=CENTER, width=120)
                his.column(5, anchor=CENTER, width=120)
                his.column(6, anchor=CENTER, width=90)
                his.column(7, anchor=CENTER, width=90)

                his.heading(1, text="Thời gian")
                his.heading(2, text="Người Thay Đổi")
                his.heading(3, text="Mã Hàng")
                his.heading(4, text="Mã Bản Vẽ cũ")
                his.heading(5, text="Mã Bản Vẽ mới")
                his.heading(6, text="Phiên Bản cũ")
                his.heading(7, text="Phiên Bản mới")
                histo=dmc_change_history()
                print(histo)
                his.delete(*his.get_children())
                for i in range(len(histo)):
                    his.insert(parent='', index=i, values=(str(histo['Date'][i]).replace('T',' '), histo['NguoiThayDoi'][i], histo['MaHang'][i], histo['MaBanVeTruoc'][i], histo['MaBanVeSau'][i], histo['PhienBanTruoc'][i], histo['PhienBanSau'][i]))
            except:
                Systemp_log(traceback.format_exc()).append_new_line()
        try:
            tk_setup = Toplevel(wk)
            tk_setup.geometry("220x200")
            tk_setup.resizable(False,False)
            tk_setup.grab_set()
            LB_mhsetup = Label(tk_setup,text="Mã hàng ", font=(font_style,font_size), anchor='w')
            LB_mhsetup.place(x=5,y=5,width=70,height=20)
            CB_mhsetup = ttk.Combobox(tk_setup,state='readonly', font=(font_style,font_size), values=pd.read_csv(file_serial, index_col=0).index.to_list())
            CB_mhsetup.place(x=65,y=5,width=110,height=20)
            CB_mhsetup.bind("<<ComboboxSelected>>", get_setup)
            if security_level>4:
                CB_mhsetup.configure(state='normal')
                Btn_get_setup = Button(tk_setup,text="Delete", font=(font_style,font_size),command=delete_setup)
                Btn_get_setup.place(x=115, y=30,width=60,height=25)
                Btn_get_setup = Button(tk_setup,text="Rename", font=(font_style,font_size),command=rename_setup)
                Btn_get_setup.place(x=30, y=30,width=60,height=25)

            LB_MaBanVe = Label(tk_setup,text="Mã Bản Vẽ", font=(font_style,font_size),borderwidth=2, anchor='w')
            LB_MaBanVe.place(x=5,y=60,width=110,height=20)
            ET_MaBanVe = Entry(tk_setup, font=font_size)
            ET_MaBanVe.place(x=115,y=60,width=100,height=20)
            LB_PhienBan = Label(tk_setup,text="Phiên Bản", font=(font_style,font_size),anchor='w')
            LB_PhienBan.place(x=5,y=85,width=100,height=20)
            ET_PhienBan = Entry(tk_setup, font=font_size)
            ET_PhienBan.place(x=115,y=85,width=60,height=20)
            LB_XuLyNhiet = Label(tk_setup,text="Xử Lý Nhiệt", font=(font_style,font_size), anchor='w')
            LB_XuLyNhiet.place(x=5,y=110,width=110,height=20)
            ET_XuLyNhiet = Entry(tk_setup, font=(font_style,font_size))
            ET_XuLyNhiet.place(x=115,y=110,width=60,height=20)

            Btn_history_setup = Button(tk_setup,text="History", font=(font_style,font_size),command=history_setup)
            Btn_history_setup.place(x=25, y=160,width=50,height=30)
            Btn_save_setup = Button(tk_setup,text="Save", font=(font_style,font_size),command=save_setup)
            Btn_save_setup.place(x=85, y=160,width=50,height=30)
            tk_setup.mainloop()
        except:
            Systemp_log(traceback.format_exc()).append_new_line()
def update_data(): # nút cập nhật
    synchronized()
def error_quantity():
    quantity = len(ET_malo.get())
    print(quantity,CB_mahang.get())
    if 'A230' in CB_mahang.get():
        if quantity != 7:
            return True
    elif 'A210' in CB_mahang.get():
        if quantity != 7:
            return True
    elif 'A2310084' in CB_mahang.get():
        if quantity != 7:
            return True
    elif 'A2312' in CB_mahang.get() or 'A2502' in CB_mahang.get():
        print('zo')
        if quantity != 8:
            return True
    else:
        if quantity != 8 or quantity != 9:
            return True
    return False
def error_wax():
    if 'A201200' in CB_mahang.get():
        return not ET_malo.get()[7:].isnumeric()
def fill_malo(i): # thực hiện tự động chuyển mã sang phần mềm
    try:
        global stt_count
        
        if CB_mahang.get() == '':
            messagebox.showerror("Lỗi hệ thống","Chưa chọn sản phẩm")
        elif ET_tennv.get() == '':
            messagebox.showerror('Lỗi hệ thống','Chưa nhập tên nhân viên')
        elif isWorking == True or savedone == True:
            messagebox.showerror("Lỗi hệ thống","Chưa kiểm tra DMC")
        # elif ('A230' in CB_mahang.get() or 'A210' in CB_mahang.get()) and len(ET_malo.get())!=7:
        #     messagebox.showerror("Lỗi cú pháp","Mã sai số lượng")
        # elif ('A230' not in CB_mahang.get() and 'A210' not in CB_mahang.get())  and (len(ET_malo.get())<8 or len(ET_malo.get())>9):
        #     messagebox.showerror("Lỗi cú pháp","Mã sai số lượng")
        elif error_quantity():
            messagebox.showerror("Lỗi cú pháp","Mã sai số lượng")
        # elif ('A230' not in CB_mahang.get() and 'A210' not in CB_mahang.get())  and (not ET_malo.get()[7:].isnumeric()):
        elif error_wax():
            messagebox.showerror("Lỗi mã","Mã khuôn sáp không hợp lệ")
        elif len(ET_malo.get())>7 and int(ET_malo.get()[7:]) > int(getwax(CB_mahang.get())):
            messagebox.showerror("Lỗi mã","Mã khuôn sáp không tồn tại")
        else:
            if checkmalo(ET_malo.get()) and malo_str(ET_malo.get()):
                global maqr
                global TimeDMCstart
                global TimeBarcode

                TimeDMCstart = datetime.datetime.now()
                mabanve = pd.read_csv(file_serial,index_col=0).loc[CB_mahang.get(),'Model']
                Phienban = pd.read_csv(file_serial,index_col=0).loc[CB_mahang.get(),'Phiên Bản']
                xulynhiet = pd.read_csv(file_serial,index_col=0).loc[CB_mahang.get(),'Xử Lý Nhiệt']
                maqr = MAQR(str4(update_serial(CB_mahang.get())),ET_malo.get(),mabanve,Phienban,CB_mahang.get(),xulynhiet)
                if maqr == 'Error':
                    return False
                if duplicate(maqr,mahang) != 0:
                    if not messagebox.askyesno('Cảnh báo trùng mã','Đã tồn tại mã cùng STT, có muốn tiếp tục hay không'):
                        return False
                ET_malo.delete(0, END)
                clicked(maqr)
    except:
        Systemp_log(traceback.format_exc()).append_new_line()
def change_machine_no(): # thay đổi tên máy thực hiện
    def save_mcno(): # lưu tên máy
        try:
            global machine_no
            machine_no = ET_mcno.get()
            data = pd.read_csv(file_position, index_col=0)
            data.loc['machine no','X'] = ET_mcno.get()
            data.to_csv(file_position)
            LB_machine.configure(text="Machine No. " + machine_no)
            tkmachine_no.destroy()
        except:
            Systemp_log(traceback.format_exc()).append_new_line()
    try:
        tkmachine_no = Toplevel(wk)
        tkmachine_no.geometry("200x100")
        tkmachine_no.grab_set()
        LB_mcno = Label(tkmachine_no,text="Số máy\nMachine no.",height=2).grid(row=0,column=0)
        ET_mcno = Entry(tkmachine_no)
        ET_mcno.grid(row=0,column=1)
        ET_mcno.insert(0,machine_no)
        btn = Button(tkmachine_no,text="Save",command=save_mcno)
        btn.grid(row=1,column=1)
        tkmachine_no.mainloop()
    except:
        Systemp_log(traceback.format_exc()).append_new_line()
def select_mahang(self):
    try:
        global mahang
        global mahang_model
        if messagebox.askyesno('Đổi mã hàng','Bạn có muốn đổi sang mã hàng '+CB_mahang.get()+' không?') and login() > 0:
            mahang = CB_mahang.get()
            mahang_model = pd.read_csv('serial_no.csv',index_col=0).loc[mahang,'Model']
        else:
            CB_mahang.set(mahang)
    except:
        Systemp_log(traceback.format_exc()).append_new_line()
        CB_mahang.set('')
def serial_change():
    def get_serial(self):
        try:
            ET_setserial.delete(0,END)
            ET_setserial.insert(0,getserial(CB_serialmahang.get()))
        except:
            Systemp_log(traceback.format_exc()).append_new_line()
    def set_serial():
        try:
            global stt_count
            global stt_max
            if messagebox.askyesno('Lưu STT','Bạn có muốn lưu thay đổi STT không?'):
                if ET_setserial.get().isnumeric():
                    setserial(CB_serialmahang.get(),ET_setserial.get())
                    tk_setserial.destroy()
                    tk_setserial.quit()
                else:
                    messagebox.showerror('Lỗi','Nhập đúng số')
            else:
                tk_setserial.destroy()
                tk_setserial.quit()
        except:
            Systemp_log(traceback.format_exc()).append_new_line()
    try:
        pdserial = pd.read_csv(file_serial,index_col=0)
        tk_setserial = Toplevel(wk)
        tk_setserial.iconbitmap('abc.ico')
        tk_setserial.grab_set()
        tk_setserial.geometry('160x100+300+300')
        LB_serialmahang = Label(tk_setserial,text='Mã',font=(font_style,font_size), anchor=W)
        LB_serialmahang.place(x=10,y=10,width=40,height=25)
        CB_serialmahang = ttk.Combobox(tk_setserial,value=pdserial.index.to_list(),state='readonly', font=(font_style,font_size))
        CB_serialmahang.place(x=50,y=10,width=100,height=25)
        CB_serialmahang.bind('<<ComboboxSelected>>',get_serial)
        LB_setserial = Label(tk_setserial,text='STT Hiện Tại',font=(font_style,font_size), anchor=W)
        LB_setserial.place(x=10,y=35,width=90,height=25)
        ET_setserial = Entry(tk_setserial,font=(font_style,font_size+1))
        ET_setserial.place(x=100,y=35,width=50,height=25)
        
        # ET_setserialmax = Entry(tk_setserial,font=(font_style,font_size+1))
        # ET_setserialmax.place(x=100,y=35,width=50,height=25)
        Btn_setserial = Button(tk_setserial,text='OK',font=(font_style,font_size+1),command=set_serial)
        Btn_setserial.place(x=108,y=70,width=35,height=25)
        tk_setserial.mainloop()
    except:
        Systemp_log(traceback.format_exc()).append_new_line()
def wax_change():
    def get_wax(self):
        try:
            ET_setwax.delete(0,END)
            ET_setwax.insert(0,getwax(CB_waxmahang.get()))
        except:
            Systemp_log(traceback.format_exc()).append_new_line()
    def set_wax():
        try:
            global stt_count
            global stt_max
            if messagebox.askyesno('Lưu mã khuôn sáp','Bạn có muốn lưu thay đổi giới hạn mã khuôn sáp không?'):
                if ET_setwax.get().isnumeric():
                    setwax(CB_waxmahang.get(),ET_setwax.get())
                    tk_setwax.destroy()
                    tk_setwax.quit()
                else:
                    messagebox.showerror('Lỗi','Nhập đúng số')
            else:
                tk_setwax.destroy()
                tk_setwax.quit()
        except:
            Systemp_log(traceback.format_exc()).append_new_line()
    try:
        pdserial = pd.read_csv(file_serial,index_col=0)
        tk_setwax = Toplevel(wk)
        tk_setwax.iconbitmap('abc.ico')
        tk_setwax.grab_set()
        tk_setwax.geometry('160x100+300+300')
        LB_waxmahang = Label(tk_setwax,text='Mã',font=(font_style,font_size), anchor=W)
        LB_waxmahang.place(x=10,y=10,width=40,height=25)
        CB_waxmahang = ttk.Combobox(tk_setwax,value=pdserial.index.to_list(),state='readonly', font=(font_style,font_size))
        CB_waxmahang.place(x=50,y=10,width=100,height=25)
        CB_waxmahang.bind('<<ComboboxSelected>>',get_wax)
        LB_setwax = Label(tk_setwax,text='Mã lớn nhất:',font=(font_style,font_size), anchor=W)
        LB_setwax.place(x=10,y=35,width=90,height=25)
        ET_setwax = Entry(tk_setwax,font=(font_style,font_size+1))
        ET_setwax.place(x=100,y=35,width=50,height=25)
        
        # ET_setwaxmax = Entry(tk_setwax,font=(font_style,font_size+1))
        # ET_setwaxmax.place(x=100,y=35,width=50,height=25)
        Btn_setwax = Button(tk_setwax,text='OK',font=(font_style,font_size+1),command=set_wax)
        Btn_setwax.place(x=108,y=70,width=35,height=25)
        tk_setwax.mainloop()
    except:
        Systemp_log(traceback.format_exc()).append_new_line()

def test_ip():
    try:
        ip = socket.gethostbyname(socket.gethostname())
        s = ''
        for i in ip:
            s =str(ord(i))+s
        s = str(int(s)*int(s[:1]))
        url = 'http://192.168.8.59:5000/api/v1/Laser/Security/'+s
        Req = requests.get(url)
        get =  json.loads( Req.text )
        return get
    except:
        return 'False'
try:
    # if test_ip() == 'False':
    #     messagebox.showerror('Lỗi máy chủ','Thiết bị chưa được cho phép truy cập')
    #     exit()


    py.FAILSAFE = False
    # Khởi tạo biến global
    rwdmcin = ''
    rework_flag = False
    savedone = False
    flag_working = False
    today = datetime.datetime.now()
    maqr = ''
    chatluong = ['']*15
    file_position = 'click_position.csv'
    machine_no = pd.read_csv(file_position, index_col=0).loc['machine no','X']
    file_serial = 'serial_no.csv'
    stt_count = 1
    stt_max = 1
    ng_count = 0
    ok_count = 0
    font_style = "Calibri"
    # font_style = "Segoe UI Variable"
    font_size = 9
    dmc_check = True
    serial_no = ''
    mahang = ''
    mahang_model = ''
    security_level = 0
    server = '192.168.8.59'
    uid = 'sa'
    pwd = '1234'
    database = 'QC'
    table = 'Castingproduct'
    isWorking = False
    TimeDMCFinish = datetime.datetime.now()
    TimeDMCstart = datetime.datetime.now()
    TimeBarcode = datetime.datetime.now()
    offline_mode = False

    ip_sr2000 = '192.168.172.116'
    port_sr2000 = 9004
    # Tạo cửa sổ giao diện chính
    wk = Tk()
    wk.iconbitmap('abc.ico')
    wk.title("Laser Automation")
    wk.geometry("740x560+0+0")
    wk.configure(bg='#DAE3F3')
    wk.resizable(False,True)

    # Tạo thanh menu
    main_menu = Menu(wk)
    wk.configure(menu=main_menu)
    edit_menu = Menu(main_menu, tearoff=0)
    connect_menu = Menu(main_menu, tearoff=0)
    main_menu.add_cascade(label="Chỉnh Sửa", menu=edit_menu)
    edit_menu.add_command(label="Vị Trí Nhấn", command=calibPos)
    edit_menu.add_command(label="Tên Máy", command=change_machine_no)
    edit_menu.add_command(label="Số Thứ Tự", command=serial_change)
    edit_menu.add_command(label="Mã Khuôn Sáp", command=wax_change)
    main_menu.add_cascade(label="Kết nối", menu=connect_menu)

    # Tạo frame tittle
    Frame_Tittle = Frame(wk,width=740, height=40,bg= '#EA6B14', highlightbackground="#41719C", highlightcolor="#41719C",highlightthickness=1)
    Frame_Tittle.place(x=0,y=0)
    LB_tittle = Label(Frame_Tittle,bg='#EA6B14',fg="#FFFFFF", text="ỨNG DỤNG KHẮC LASER",font=(font_style,font_size+2),anchor=CENTER)
    LB_tittle.place(x=200,y=5,width=340,height=30)
    nowtime = Label(Frame_Tittle,bg='#EA6B14',fg='#DDDDDD', text = f"{datetime.datetime.now():%d/%m/%Y %H:%M:%S}",font=(font_style,font_size+1))
    nowtime.place(x=580,y=12)


    # Tạo frame chính
    frame1 = Frame(wk,width=740, height=860, highlightbackground="#41719C", highlightcolor="#41719C",highlightthickness=1)
    frame1.configure(bg = '#DAE3F3')
    frame1.place(x=0,y=40)

    LB_mahang = Label(frame1,bg='#D9D9D9', text = "Mã hàng đang chạy (Chọn mã cần đổi)",font=(font_style,font_size),highlightbackground="#41719C", highlightcolor="#41719C",highlightthickness=1, width=30, height=2, anchor=W)
    LB_mahang.place(x=20,y=3, width=250, height=42)
    listmahang = ['003','004','Other']
    CB_mahang = ttk.Combobox(frame1, state='readonly', values=pd.read_csv(file_serial, index_col=0).index.to_list(), font=(font_style,2*font_size), width=15)
    CB_mahang.place(x=265,y=3, width=240, height=42)
    CB_mahang.bind("<<ComboboxSelected>>",select_mahang)

    LB_ng = Label(frame1, bg='#FFFF00', text="Số lượng khắc\nNG của máy:\n"+str(ng_count), font=(font_style,font_size,"bold"), highlightbackground="#41719C", highlightcolor="#41719C",highlightthickness=1,anchor=CENTER,width=12,height=3)
    LB_ng.place(x=515, y=2, width= 100, height= 63)
    LB_ok = Label(frame1, bg='#92D050', text="Số lượng khắc\nOK của máy:\n"+str(ok_count), font=(font_style,font_size,"bold"),highlightbackground="#41719C", highlightcolor="#41719C",highlightthickness=1,anchor=CENTER,width=12,height=3)
    LB_ok.place(x=625, y=2, width= 100, height= 63)
    LB_serial = Label(frame1,bg='#D9D9D9', text = "Serial No. " + serial_no,font=(font_style,font_size),borderwidth=1,relief="solid", width=16, height=1, anchor=W)
    LB_serial.place(x=186, y=52, width=100, height=30)
    LB_machine = Label(frame1,bg='#D9D9D9', text = " Machine No. " + machine_no,font=(font_style,font_size),borderwidth=1,relief="solid", anchor=W)
    LB_machine.place(x=20, y=52, width=160, height=30)
    LB_tennv = Label(frame1,bg='#D9D9D9', text = "Tên Nhân Viên \nOperator Name",font=(font_style,font_size),highlightbackground="#41719C", highlightcolor="#41719C",highlightthickness=1,anchor=W,width=19,height=2)
    LB_tennv.place(x=20,y=90, width=165, height=45)
    ET_tennv = Entry(frame1, font=(font_style,font_size), width=22, highlightbackground="#41719C", highlightcolor="#000000",highlightthickness=1)
    ET_tennv.place(x=182,y=90, width=160,  height=45)
    LB_gcdate = Label(frame1,bg='#D9D9D9', text = "Ngày Gia Công\nMachining Date",font=(font_style,font_size),highlightbackground="#41719C", highlightcolor="#41719C",highlightthickness=1,anchor=W,width=19,height=2)
    LB_gcdate.place(x=20,y=140, width=165, height=45)
    ET_gcdate = Entry(frame1, font=(font_style,font_size), width=22, highlightbackground="#41719C", highlightcolor="#000000",highlightthickness=1)
    ET_gcdate.place(x=182,y=140, width=160,  height=45)
    LB_malo = Label(frame1,bg='#D9D9D9', text = "Mã Lò\nCasting code",font=(font_style,font_size),highlightbackground="#41719C", highlightcolor="#41719C",highlightthickness=1,anchor=W,width=19,height=2)
    LB_malo.place(x=20,y=190, width=165, height=45)
    ET_malo = Entry(frame1, font=(font_style,font_size), width=22, highlightbackground="#41719C", highlightcolor="#000000",highlightthickness=1)
    ET_malo.bind('<Shift_R>', fill_malo)
    ET_malo.place(x=182,y=190, width=160,  height=45)
    LB_hint = Label(frame1, bg='#DAE3F3',fg='#BFBFBF', text="Nhấn Enter\nđể khắc", font=(font_style,font_size,"bold"),anchor=CENTER,width=8,height=2)
    LB_hint.place(x=350, y=140, height=45)

    Btn_test = Button(frame1, text="Online",font=(font_style,font_size,"bold"),bg='#00FF00',width=9,highlightbackground="#41719C", highlightcolor="#000000",highlightthickness=1, command = offline_mode_tog)
    Btn_test.place(x=350, y=90, width=60,  height=50)

    LB_dmccheck = Label(frame1, bg='#BDD7EE', text="Kết quả:\n"+' ', font=(font_style,font_size,"bold"),highlightbackground="#41719C", highlightcolor="#41719C",highlightthickness=1,anchor=CENTER,width=7,height=2)
    LB_dmccheck.place(x=352, y=190, width=60,  height=45)
    Btn_suahang = Button(frame1, text="Sửa hàng\nRework",font=(font_style,font_size,"bold"),bg='#FFFF00',width=9,highlightbackground="#41719C", highlightcolor="#000000",highlightthickness=1, command = rework)
    Btn_suahang.place(x=425, y=90, width=85,  height=50)
    Btn_xemdulieu = Button(frame1, text="Xem dữ liệu\nData",font=(font_style,font_size,"bold"),bg='#8FAADC',fg='#FFFFFF',width=9,highlightbackground="#41719C", highlightcolor="#000000",highlightthickness=1, command = popup_data)
    Btn_xemdulieu.place(x=425, y=160, width=85,  height=50)
    Btn_setup = Button(frame1, text="Thiết lập mục tiêu\nchất lượng & sản lượng",font=(font_style,font_size-1,"bold"),bg='#C55A11',fg= '#FFFFFF',borderwidth=2, command = setup_data)
    Btn_setup.place(x=525, y=90, width=150,  height=50)
    Btn_update = Button(frame1, text="Cập nhật dữ liệu\nMã lò đúc",font=(font_style,font_size,"bold"),bg='#C55A11',fg='#FFFFFF',width=18,highlightbackground="#41719C", highlightcolor="#000000",highlightthickness=1, command = update_data)
    Btn_update.place(x=525, y=160, width=150,  height=50)

    # Tạo frame dữ liệu NG OK
    frame_top10 = Frame(frame1)
    frame_top10.configure(bg = '#DAE3F3')
    frame_top10.place(x=5,y=240,width=730, height=260)
    wrapper1 = LabelFrame(frame_top10, text="OK", bg='#C5E0B4', width=365, height=255)
    wrapper2 = LabelFrame(frame_top10, text="NG",  bg='#FFF2CC', width=365, height=255)
    wrapper1.place(x=0, y=0)
    wrapper2.place(x=365, y=0)
    trv = ttk.Treeview(wrapper1, columns=(1, 2,3), show="headings", height="12")
    trv.place(x=3, y=0,height=230)
    style = ttk.Style()
    style.configure("Treeview.treearea", font=font_size)
    trv.column(1, anchor=CENTER, width=190)
    trv.column(2, anchor=CENTER, width=50)
    trv.column(3, anchor=CENTER, width=112)

    trv.heading(1, text="DMC")
    trv.heading(2, text="Quality")
    trv.heading(3, text="Time_Finish")

    trv1 = ttk.Treeview(wrapper2, columns=(1, 2,3), show="headings", height="12")
    trv1.place(x=3, y=0,height=230)
    trv1.column(1, anchor=CENTER, width=190)
    trv1.column(2, anchor=CENTER, width=50)
    trv1.column(3, anchor=CENTER, width=112)

    trv1.heading(1, text="DMC")
    trv1.heading(2, text="Quality")
    trv1.heading(3, text="Time_Finish")

    again()
    wk.mainloop()
except Exception:
    Systemp_log(traceback.format_exc()).append_new_line()
