import requests,json
import pandas as pd
import env_variable as env
api_host_url = env.get_env("API_HOST_URL")
def synchronized_data(Castingname,TimeCasting):
    url = f'{api_host_url}Laser/synchronized'
    payl = {"Castingname": Castingname, "TimeCasting": TimeCasting}
    Req = requests.post(url, json=payl)
def save_data2(NV,Machineno,Productname,DMCin,DMCout,TimeDMCstart,TimeDMCFinish):
    url = f'{api_host_url}Laser/Save_Data'
    payl = {
            "DMCin": DMCin,
            "DMCout": DMCout,
            "MODE": "2",
            "MachineNo": Machineno,
            "NameOperator": NV,
            "NameProduct": Productname,
            "TimeInDMC": str(TimeDMCstart)[:-3],
            "TimeOutDMC": str(TimeDMCFinish)[:-3]
            }
    Req = requests.post(url,json=payl)
    return Req.text.strip()
def savedata(NV,Machineno,Productname,DMCin,DMCout,DMCrework,TimeDMCstart,TimeDMCFinish,TimeBarcode,Result,status,chatluong):
    url = f'{api_host_url}Laser/SaveData'
    payl = {
            "Axial_Nonuniformity": chatluong[10],
            "DMCRework": DMCrework,
            "DMCin": DMCin,
            "DMCout": DMCout,
            "Decode": chatluong[3],
            "Fixed_Pattern_Damage": chatluong[7],
            "Format_Info_Damage": chatluong[8],
            "Grid_Nonuniformity": chatluong[11],
            "MachineNo": Machineno,
            "Modulation": chatluong[5],
            "NameOperator": NV,
            "NameProduct": Productname,
            "Print_Growth_Horizontal": chatluong[13],
            "Print_Growth_Vertical": chatluong[14],
            "Quality": chatluong[2],
            "Reflectance_Margin": chatluong[6],
            "Result": Result,
            "Status": status,
            "Symbol_Contrast": chatluong[4],
            "TimeInDMC": str(TimeDMCstart)[:-3],
            "TimeOutBarcode": str(TimeBarcode)[:-3],
            "TimeOutDMC": str(TimeDMCFinish)[:-3],
            "Unused_Err_Correction": chatluong[12],
            "Version_Info_Damage": chatluong[9]
            }
    Req = requests.post(url,json=payl)
    return Req.text.strip()
def getserial(mahang):
    url = f'{api_host_url}Laser/get_serial/'+mahang
    Req = requests.get(url)
    get =  json.loads( Req.text )['data'][0]['Serial']
    return get
def update_serial(mahang):
    url = f'{api_host_url}Laser/update_serial/'+mahang
    Req = requests.get(url)
    get =  json.loads( Req.text )['data'][0]['Serial']
    return get
def setserial(mahang,serial):
    url = f'{api_host_url}Laser/set_serial'
    payl = {"NameProduct": mahang,"Serial": serial}
    Req = requests.post(url,json=payl)    
    # get =  json.loads( Req.text )
    return Req.text
def getwax(mahang):
    url = f'{api_host_url}Laser/get_waxmold/'+mahang
    Req = requests.get(url)
    get =  json.loads( Req.text )['data'][0]['Waxmold']
    return get
def setwax(mahang,ma):
    url = f'{api_host_url}Laser/set_waxmold'
    payl = {"NameProduct": mahang,"Waxmold": ma}
    Req = requests.post(url,json=payl)    
    # get =  json.loads( Req.text )
    return Req.text
def update_result(Machineno,ProductName,Result,TimeBarcode,pcs_count):
    url = f'{api_host_url}Laser/Update_result'
    payl = {
            "MachineNo": Machineno,
            "NameProduct": ProductName,
            "Result": Result,
            "TimeOutBarcode": str(TimeBarcode)[:-3],
            "row":pcs_count
            }
    print(payl)
    Req = requests.post(url,json=payl)
    return Req.text.strip()
def get_count_result(Machineno,ProductName):
    url = f'{api_host_url}Laser/get_count_result'
    payl =  {
            "MachineNo": Machineno,
            "NameProduct": ProductName
            }
    Req = requests.post(url,json=payl)
    return Req.text
def count_history(Machineno,strtoday,strnextday,Result):
    url = f'{api_host_url}Laser/count_history'
    payl = {
            "MachineNo": Machineno,
            "Result": Result,
            "strnextday": strnextday,
            "strtoday": strtoday
            }
    Req = requests.post(url,json=payl)
    get =  json.loads( Req.text )
    return str(get)
def get_user(password):
    url = f'{api_host_url}Laser/Get_user_laser/'+password
    Req = requests.get(url)
    user =  json.loads( Req.text ) ['data'][0]['Name']
    scrt =  json.loads( Req.text ) ['data'][0]['Security']
    return user,scrt
def dmc_setup_history(date,nv,mahang,mbvtruoc,mbvsau,pbtruoc,pbsau):
    url = f'{api_host_url}Laser/DMC_setup_history'
    payl = {
            "Date": str(date)[:-3],
            "MaBanVeSau": mbvsau,
            "MaBanVeTruoc": mbvtruoc,
            "MaHang": mahang,
            "NguoiThayDoi": nv,
            "PhienBanSau": pbsau,
            "PhienBanTruoc": pbtruoc
            }
    Req = requests.post(url,json=payl)
    # get =  json.loads( Req.text )
    return Req.text
def dmc_change_history():
    url = f'{api_host_url}Laser/DMC_change_history'
    Req = requests.get(url)
    get =  json.loads( Req.text )['data']
    get = pd.DataFrame(get)
    return get
def duplicate(DMCin,ProductName):
    url = f'{api_host_url}Laser/fill_malo'
    payl =  {
            "DMCout": DMCin,
            "NameProduct": ProductName
            }
    print(payl)
    Req = requests.post(url,json=payl)
    get =  json.loads( Req.text )
    print(get)
    return get
def check_castingname(malo):
    url = f'{api_host_url}Laser/Check_castingname/'+malo
    Req = requests.get(url)
    return Req.text.strip()
def laser_result(Machineno,ProductName,Result):
    url = f'{api_host_url}Laser/Laser_result_history'
    payl =  {
            "MachineNo": Machineno,
            "NameProduct": ProductName,
            "Result": Result
            }
    Req = requests.post(url,json=payl)
    get =  json.loads( Req.text )['data']
    get = pd.DataFrame(get)
    return get
def laser_all_data(Machineno,nv,ProductName,TimeStart,TimeEnd):
    url = f'{api_host_url}Laser/Laser_all_data'
    payl =  {
            "MachineNo": Machineno,
            "NameOperator": nv,
            "NameProduct": ProductName,
            "TimeFinish": TimeEnd,
            "TimeStart": TimeStart
            }
    Req = requests.post(url,json=payl)
    get =  json.loads( Req.text )['data']
    get = pd.DataFrame(get)
    for i in range(len(get)):
        get['TimeInDMC'][i]=str(get['TimeInDMC'][i]).replace('T',' ')
        get['TimeOutDMC'][i] = str(get['TimeOutDMC'][i]).replace('T', ' ')
        get['TimeOutBarcode'][i] = str(get['TimeOutBarcode'][i]).replace('T', ' ')
    get=get[[  "MachineNo"
            ,"NameOperator"
            ,"NameProduct"
            ,"DMCin"
            ,"TimeInDMC"
            ,"TimeOutDMC"
            ,"DMCout"
            ,"TimeOutBarcode"
            ,"DMCRework"
            ,"Result"
            ,"Quality"
            ,"Status"
            ,"Decode"
            ,"Symbol_Contrast"
            ,"Modulation"
            ,"Reflectance_Margin"
            ,"Fixed_Pattern_Damage"
            ,"Format_Info_Damage"
            ,"Version_Info_Damage"
            ,"Axial_Nonuniformity"
            ,"Grid_Nonuniformity"
            ,"Unused_Err_Correction"
            ,"Print_Growth_Horizontal"
            ,"Print_Growth_Vertical"]]
    return get
def get_status():
    url=f'{api_host_url}Laser/List_error'
    Req = requests.get(url)
    get =  json.loads( Req.text )['data']
    get = pd.DataFrame(get).set_index('Status_error').index.to_list()
    return get
def type_error(Status_error):
    url = f'{api_host_url}Laser/Type_error'
    payl =  {
            "Status_error": Status_error
            }
    Req = requests.post(url,json=payl)
    return Req.text
