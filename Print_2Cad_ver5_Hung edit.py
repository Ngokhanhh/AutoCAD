## Copyright (C)
## Power System Analysis and Planning Department
## National Load Dispatch Centre (NLDC)
## Vietnam Electricity(EVN)
## Address: Floor 10, Tower A, 11 Cua Bac, Ba Dinh, Ha Noi
##
##
## DOC-ME PLEASE!
## read file : 'Huong dan su dung xuat PSSE_2_CAD.docx'
##
## HISTORY
## 
##       
## Jui 17, 2014 : First Version (PhuongPQ)
## Oct 01, 2014 : Thu ve 1 file Print_2Cad.py, bo file Tools_ppt.py (PhuongPQ)
##
## 

import psspy 
import string
import Tkinter 
import tkMessageBox
import tkFileDialog


## OPTIONS OF Print_2Cad.py
#  ------------------------------------------------------------------------------------------------------------------------
prefix_PQ = 'get_PQ:'
nFormat_PQ = 1 
unit_PQ =''   ## ='' or = 'MVA'
#
prefix_U = 'get_U:'
nFormat_U = 1 
unit_U = 'kV' ## ='' or = 'kV'
#
prefix_RATE = 'get_RATE:'
nFormat_RATE = 1 
unit_RATE = '%' ## ='' or = '%'
## END OF OPTIONS-----------------------------------------------------------------------------------------------------------

##
##Function General
##--------------------------------------------------------------------------------------------------------------------------
## Convert a complex to a String
## Inputs :  c : the complex
##           n : number of format (after decimal)
## Ouputs : a String result
#  ------------------------------------------------------------------------------------------------------------------------
def complex_To_String(c,nFormat):             
    if c.imag >= 0:
        return number_To_String(c.real,nFormat) + "+j" + number_To_String(c.imag  , nFormat) 
    else :
        return number_To_String(c.real,nFormat) + "-j" + number_To_String(-c.imag , nFormat)               

## Convert a double number to a String
## Inputs :  v : the double number
##           n : number of format (after decimal)
## Ouputs : a String result
#  ------------------------------------------------------------------------------------------------------------------------
def number_To_String(v,nFormat):
    if abs(v)< 1e-4 :
        return "0"
    nF = nFormat
    if (nF < 0):
        nF = 0       
    return ("{0:." + str(nF) + "f}").format(v)

def window_error(s):
    root = Tkinter.Tk()
    root.withdraw() 
    tkMessageBox.showerror("USER ERROR",s)
    root.destroy()
    root.mainloop()       
    return

def window_warning(s):
    root = Tkinter.Tk()
    root.withdraw()       
    tkMessageBox.showwarning("USER WARNING",s)
    root.destroy()
    root.mainloop()        
    return

def window_info(s):
    root = Tkinter.Tk()
    root.withdraw()
    tkMessageBox.showinfo("INFOS",s)
    root.destroy()
    root.mainloop()        
    return

def ask_open_File_CAD_dxf():
    root = Tkinter.Tk()
    root.withdraw()
    s1 = tkFileDialog.askopenfilename(defaultextension=("AutoCad DXF File",".dxf"),\
                            filetypes=[('All Files', '.*'),("AutoCad DXF File",".dxf")],title='Choose AutoCad input file')
    root.destroy()
    root.mainloop()
    return s1

#
## Pick up an user error: The python will be stopped with an explicit message to user
#  -------------------------------------------------------------------------------------------------------------------------
def error(s):      
    s1= "\n\nUSER ERROR  --- USER ERROR --- USER ERROR --- USER ERROR\n" + s + "\n\n"
    window_error(s)       
    raise StandardError, s1
    return

def error_0():    
    raise StandardError
    return

## Read a txt file
## Inputs : a txt file
## Ouputs : an array txt
#  -----------------------------------------------------------------------------------------------------------------------
def read_File_text (sfile):
    try:
        ins = open( sfile, "r" )
        print ("Open file in:\n'" + sfile + "'")
        array = []
        for line in ins:
            array.append( line )
        ins.close()
        return array
    #
    except: error("\tImpossible to open file in:\n'" + sfile +"'")       

## Write an array txt to a file
## Inputs : name of the file to write
##        : an array txt
## Ouputs : null
#  -----------------------------------------------------------------------------------------------------------------------
def write_File_text (sfile , array):
    n = 0
    try:
        f = open(sfile, "w")
        try:
            f.writelines(array) # Write a sequence of strings to a file
        except:
            error('Error in write_File_text')      
        finally:
            f.close()
            s1 = "File result created in:\n" + sfile
            window_info("SUCCESS !!!\n"+s1)
            print s1
    except :
        n += 1
        sn = modif_nameFile (sfile,"_1")               
        write_File_text(sn , array)
        #
        if(n>10): error("\tImpossible to write file in: " + sn) 

## get all string in the array string 'array' with start by the substring 'sub'
## Inputs : aray  : aray string
##        : sub   : a string 
## Ouput  : array of string with start with sub
#  -----------------------------------------------------------------------------------------------------------------------
def get_String_With_Start_By_SubString(aray, sub):
    res = []
    idp = []
    n = 0
    for line in aray:
        l1 = line.replace(" ","")
        if l1.startswith(sub):                        
            res.append(line)
            idp.append(n)
        n += 1
    #
    return idp,res

def modif_nameFile(s,sadd):
    try:
        t1 = string.rindex(s, ".")
        return s[0:t1] + sadd + "." + s[t1+1:]
    except:
        error ("\tProblem of file name:"+s)
##--------------------------------------------------------------------------------------------------------------------------
## END Function General
##--------------------------------------------------------------------------------------------------------------------------               


# get inputs et outputs files
cadInput = ask_open_File_CAD_dxf()
if cadInput =='':
    window_error('File input is NULL')
    
    


cadOutput = modif_nameFile(cadInput,'_res')
a_res = read_File_text(cadInput)
#ini
repList = ''
errList = ''
ne = 0

##
## For extract PQ [MVA]
###-----------------------------------------------------------------------------------------------------------
id_pq,line_pq = get_String_With_Start_By_SubString(a_res,prefix_PQ)
n = -1
for li in line_pq:
    n += 1
    l1 = li.replace(" ","")
    l1 = l1.replace("\n","")
    l1 = l1[len(prefix_PQ):]
    bra = l1.split('+')
    vi = 0
    vr = 0.0
    #Hung edit #########################
    i=0


    ####################################
    #
    try:
        for bri in bra:
            # Hung edit####################
            i = i +1

            ################################
            b = bri.split(',')
            # for the cas get_PQ: -2012,3119,1 <=> get_PQ: 3119,2012,1 =
            vpn = 1.0
            if int(b[0])<0:
                vpn = -1.0
            #        
            if len(b) == 3:            
                ier,vb = psspy.brnflo(abs(int(b[0])),int(b[1]),b[2])# get power flow of normal branch
                #
                if (ier == 0)|(ier == 3):
                    vi += vb
                else :
                    error_0()
            #
# Hung Edit#########################################################################################################
                ier, br_rate1 = psspy.brnmsc(abs(int(b[0])),int(b[1]),b[2],"PCTRTA")#get rate of branch
                if ier != 0:
                    if (ier == 1) | (ier == 2):
                        error_0()
                    #
                    br_rate1 = 0.0              
                #    
                ier, br_rate2 = psspy.brnmsc(abs(int(b[1])),int(b[0]),b[2],"PCTRTA")#get rate of branch
                #
                if ier != 0:
                    br_rate2 = 0.0
                #
                vr = max(br_rate1 , br_rate2)
                if i ==1:
                    H_add = number_To_String(vr,nFormat_RATE) + unit_RATE
                    H_add = " (" + H_add + ")"
                elif i>1:
                    H_add =""

##################################################################################################################


            elif len(b) == 4:
                ier,vb = psspy.wnddt2(abs(int(b[0])),int(b[1]),int(b[2]),b[3],"FLOW")# get power flow of three winding transformer
                #
                if (ier == 0)|(ier == 3):
                    vi += vb
                else:
                    error_0()
# Hung edit ######################################################################################################
                ier, trf_rate1 = psspy.wnddat(abs(int(b[0])),int(b[1]),int(b[2]),b[3],"PCTRTA")#get rate of trf3
                if ier != 0:
                    if (ier == 1) | (ier == 2):
                        error_0()
                    #
                    trf_rate1 = 0.0
                #
                ier, trf_rate2 = psspy.wnddat(abs(int(b[1])),int(b[0]),int(b[2]),b[3],"PCTRTA")#get rate of trf3
                if ier != 0:
                    trf_rate2 = 0.0
                #
                ier, trf_rate3 = psspy.wnddat(abs(int(b[2])),int(b[1]),int(b[0]),b[3],"PCTRTA")#get rate of trf3
                if ier != 0:
                    trf_rate3 = 0.0
                #
                vr = max(trf_rate1,trf_rate2,trf_rate3)

                ############# Option Get rate for Trans
                if i ==1:
                    H_add = number_To_String(vr,nFormat_RATE) + unit_RATE
                    H_add = " (" + H_add + ")"
                elif i>1:
                    H_add =""

###################################################################################################################
            #
            else : error_0()
        #        
        # convert to String with format        
        vsi = complex_To_String( vi * vpn , nFormat_PQ) + unit_PQ + H_add
        repList += vsi + " = " + li
        a_res[id_pq[n]] = vsi+"\n"
    #
    except :
        ne += 1
        errList += '\t' + li   


##
## For extract Voltage [kV]
###-----------------------------------------------------------------------------------------------------------
id_u,line_u = get_String_With_Start_By_SubString(a_res,prefix_U)
n = -1
for li in line_u:
    n += 1
    l1 = li.replace(" ","")
    l1 = l1.replace("\n","")
    l1 = l1[len(prefix_U):]
    vb = 0
    try: 
        ier,vb =  psspy.busdat(int(l1),'KV') # get bus voltage
        if ier > 0:
            error_0()
        #               
        vsi = number_To_String(vb,nFormat_U) + unit_U
        repList += vsi + " = " + li        
        a_res[id_u[n]] = vsi+"\n"
    #
    except :
        ne += 1
        errList += '\t' + li
      
   
##
## For extract RATE [%]
###-----------------------------------------------------------------------------------------------------------
id_rate,line_rate = get_String_With_Start_By_SubString(a_res,prefix_RATE)
n = -1
for li in line_rate:
    n += 1
    l1 = li.replace(" ","")
    l1 = l1.replace("\n","")
    l1 = l1[len(prefix_RATE):]
    b = l1.split(',')
    vr = 0.0
    #
    try:
        if len(b) == 3:                
            ier, br_rate1 = psspy.brnmsc(int(b[0]),int(b[1]),b[2],"PCTRTA")#get rate of branch
            if ier != 0:
                if (ier == 1) | (ier == 2):
                    error_0()
                #
                br_rate1 = 0.0              
            #    
            ier, br_rate2 = psspy.brnmsc(int(b[1]),int(b[0]),b[2],"PCTRTA")#get rate of branch
            #
            if ier != 0:
                br_rate2 = 0.0
            #
            vr = max(br_rate1 , br_rate2)
        # RATE OF TRANSFORMER THREE WINDING
        elif len(b) == 4:                
            ier, trf_rate1 = psspy.wnddat(int(b[0]),int(b[1]),int(b[2]),b[3],"PCTRTA")#get rate of trf3
            if ier != 0:
                if (ier == 1) | (ier == 2):
                    error_0()
                #
                trf_rate1 = 0.0
            #
            ier, trf_rate2 = psspy.wnddat(int(b[1]),int(b[0]),int(b[2]),b[3],"PCTRTA")#get rate of trf3
            if ier != 0:
                trf_rate2 = 0.0
            #
            ier, trf_rate3 = psspy.wnddat(int(b[2]),int(b[1]),int(b[0]),b[3],"PCTRTA")#get rate of trf3
            if ier != 0:
                trf_rate3 = 0.0
            #
            vr = max(trf_rate1,trf_rate2,trf_rate3)
        #
        else:
            error_0()
        #               
        vsi = number_To_String(vr,nFormat_RATE) + unit_RATE
        repList += vsi + " = " + li 
        a_res[id_rate[n]]= vsi+"\n"
    except :
        ne += 1
        errList += '\t' +li           

#    
## --------------------------------------------------------------------------------------------------------------
if (ne > 0):      
    if ne ==1 :
        errList =  str(ne) + " ERROR of definion of \n\t branch \n\t three-winding transformer \n\t bus number \nin:\n" + errList
    else      :
        errList =  str(ne) + " ERRORS of definion of \n\t branch \n\t three-winding transformer \n\t bus number \nin:\n" + errList
    #
    error(errList)
    
print ("List of replacement:")
print repList+'SUCCESSFUL'

##    
## create file final
## ---------------------------------------------------------------------------------------------------------------
write_File_text (cadOutput,a_res)











