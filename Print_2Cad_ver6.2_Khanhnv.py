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
## Feb 05, 2026: Migrated to Python 3.9, updated to bypass definition errors, continue processing, and save result file (KhanhNV)


import psspy 
import string
# Khanhnv edit: Updated tkinter library for Python 3.9
import tkinter as Tkinter 
from tkinter import messagebox as tkMessageBox
from tkinter import filedialog as tkFileDialog
import sys # Khanhnv edit: Added system library

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
    # root.mainloop() # Khanhnv edit: Removed mainloop to prevent script hanging
    return

def window_warning(s):
    root = Tkinter.Tk()
    root.withdraw()       
    tkMessageBox.showwarning("USER WARNING",s)
    root.destroy()
    # root.mainloop() # Khanhnv edit: Removed mainloop
    return

def window_info(s):
    root = Tkinter.Tk()
    root.withdraw()
    tkMessageBox.showinfo("INFOS",s)
    root.destroy()
    # root.mainloop() # Khanhnv edit: Removed mainloop
    return

def ask_open_File_CAD_dxf():
    root = Tkinter.Tk()
    root.withdraw()
    # Khanhnv edit: Fixed defaultextension and file types for Python 3 compatibility
    s1 = tkFileDialog.askopenfilename(
        defaultextension=".dxf",
        filetypes=[('AutoCad DXF File', '*.dxf'), ('All Files', '*.*')],
        title='Choose AutoCad input file'
    )
    root.destroy()
    # root.mainloop() # Khanhnv edit: Removed mainloop
    return s1

#
## Pick up an user error: The python will be stopped with an explicit message to user
#  -------------------------------------------------------------------------------------------------------------------------
def error(s):      
    s1= "\n\nUSER ERROR  --- USER ERROR --- USER ERROR --- USER ERROR\n" + s + "\n\n"
    window_error(s)       
    # Khanhnv edit: Updated StandardError to Exception for Python 3
    raise Exception(s1) 
    return

def error_0():    
    raise Exception() # Khanhnv edit: Updated Exception
    return

## Read a txt file
## Inputs : a txt file
## Ouputs : an array txt
#  -----------------------------------------------------------------------------------------------------------------------
def read_File_text(sfile):
    try:
        # Khanhnv edit: Added utf-8 encoding for reading DXF files in Python 3.9
        ins = open(sfile, "r", encoding='utf-8', errors='ignore') 
        print("Open file in:\n'" + sfile + "'")
        array = ins.readlines()
        ins.close()
        return array
    except:
        window_error("\tImpossible to open file in:\n'" + sfile + "'")
        sys.exit() # Khanhnv edit: Exit if file cannot be opened

## Write an array txt to a file
## Inputs : name of the file to write
##        : an array txt
## Ouputs : null
#  -----------------------------------------------------------------------------------------------------------------------
def write_File_text (sfile , array):
    # Khanhnv edit: Updated file writing logic to support encoding
    try:
        f = open(sfile, "w", encoding='utf-8')
        f.writelines(array) # Write a sequence of strings to a file
        f.close()
        return True
    except Exception as e:
        window_error("Error in write_File_text: " + str(e))
        return False

## get all string in the array string 'array' with start by the substring 'sub'
## Inputs : aray  : aray string
##        : sub   : a string 
## Ouput  : array of string with start with sub
#  -----------------------------------------------------------------------------------------------------------------------
def get_String_With_Start_By_SubString(aray, sub):
    res = []
    idp = []
    # Khanhnv edit: Used enumerate to get index more efficiently
    for n, line in enumerate(aray):
        l1 = line.replace(" ","")
        if l1.startswith(sub):                        
            res.append(line)
            idp.append(n)
    return idp,res

def modif_nameFile(s,sadd):
    try:
        # Khanhnv edit: Used rindex directly from string
        t1 = s.rindex(".")  
        return s[0:t1] + sadd + "." + s[t1+1:]
    except:
        error ("\tProblem of file name:"+s)
##--------------------------------------------------------------------------------------------------------------------------
## END Function General
##--------------------------------------------------------------------------------------------------------------------------               


# get inputs et outputs files
cadInput = ask_open_File_CAD_dxf()
if not cadInput: # Khanhnv edit: Null check
    sys.exit()

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
# Khanhnv edit: Used enumerate for the line loop
for n, li in enumerate(line_pq):
    # Khanhnv edit: Used try...except for each line to bypass partial errors
    try:
        H_add = ""
        l1 = li.replace(" ","").replace("\n","")
        l1 = l1[len(prefix_PQ):]
        bra = l1.split('+')
        # Khanhnv edit: Initialized complex number as 0j
        vi = 0j
        vr_max = 0.0
        
        for i, bri in enumerate(bra):
            b = bri.split(',')
            # for the cas get_PQ: -2012,3119,1 <=> get_PQ: 3119,2012,1 =
            vpn = 1.0
            if int(b[0])<0:
                vpn = -1.0
            #        
            if len(b) == 3: # get power flow of normal branch           
                ier, vb = psspy.brnflo(abs(int(b[0])),int(b[1]),b[2])
                if (ier != 0) and (ier != 3): raise Exception()
                vi += vb * vpn
                
                # Khanhnv edit: Faster way to get rate
                ier, r1 = psspy.brnmsc(abs(int(b[0])),int(b[1]),b[2],"PCTRTA")
                ier, r2 = psspy.brnmsc(abs(int(b[1])),int(b[0]),b[2],"PCTRTA")
                vr_max = max(vr_max, r1 if ier == 0 else 0, r2 if ier == 0 else 0)

            elif len(b) == 4: # get power flow of three winding transformer
                ier, vb = psspy.wnddt2(abs(int(b[0])),int(b[1]),int(b[2]),b[3],"FLOW")
                if (ier != 0) and (ier != 3): raise Exception()
                vi += vb * vpn
                
                # Khanhnv edit: Get the maximum rate of the 3-winding transformer
                for nodes in [(b[0],b[1],b[2]), (b[1],b[0],b[2]), (b[2],b[1],b[0])]:
                    ier, r = psspy.wnddat(abs(int(nodes[0])),int(nodes[1]),int(nodes[2]),b[3],"PCTRTA")
                    vr_max = max(vr_max, r if ier == 0 else 0)
            else:
                raise Exception()
            
            # Khanhnv edit: Attach rate to the first label
            if i == 0:
                H_add = " (" + number_To_String(vr_max, nFormat_RATE) + unit_RATE + ")"

        # convert to String with format        
        vsi = complex_To_String(vi, nFormat_PQ) + unit_PQ + H_add
        a_res[id_pq[n]] = vsi + "\n"
        repList += vsi + " = " + li
    except:
        ne += 1
        errList += "\t" + li # Khanhnv edit: Log error but continue processing

##
## For extract Voltage [kV]
###-----------------------------------------------------------------------------------------------------------
id_u,line_u = get_String_With_Start_By_SubString(a_res,prefix_U)
for n, li in enumerate(line_u):
    try:
        l1 = li.replace(" ","").replace("\n","")
        bus_num = int(l1[len(prefix_U):])
        ier, vb = psspy.busdat(bus_num, 'KV') # get bus voltage
        if ier > 0: raise Exception()
        
        vsi = number_To_String(vb, nFormat_U) + unit_U
        a_res[id_u[n]] = vsi + "\n"
        repList += vsi + " = " + li
    except:
        ne += 1
        errList += "\t" + li
      
##
## For extract RATE [%]
###-----------------------------------------------------------------------------------------------------------
id_rate,line_rate = get_String_With_Start_By_SubString(a_res,prefix_RATE)
for n, li in enumerate(line_rate):
    try:
        l1 = li.replace(" ","").replace("\n","")
        b = l1[len(prefix_RATE):].split(',')
        vr = 0.0
        if len(b) == 3: # RATE OF BRANCH
            ier, r1 = psspy.brnmsc(int(b[0]),int(b[1]),b[2],"PCTRTA")
            ier, r2 = psspy.brnmsc(int(b[1]),int(b[0]),b[2],"PCTRTA")
            vr = max(r1 if ier == 0 else 0, r2 if ier == 0 else 0)
        elif len(b) == 4: # RATE OF TRANSFORMER THREE WINDING
            for nodes in [(b[0],b[1],b[2]), (b[1],b[0],b[2]), (b[2],b[1],b[0])]:
                ier, r = psspy.wnddat(int(nodes[0]),int(nodes[1]),int(nodes[2]),b[3],"PCTRTA")
                vr = max(vr, r if ier == 0 else 0)
        else:
            raise Exception()
        
        vsi = number_To_String(vr, nFormat_RATE) + unit_RATE
        a_res[id_rate[n]] = vsi + "\n"
        repList += vsi + " = " + li
    except:
        ne += 1
        errList += "\t" + li

# Khanhnv edit: Save results
write_File_text(cadOutput, a_res)

# Khanhnv edit: Print List of replacement
print("List of replacement:")

# Khanhnv edit: Updated print for Python 3
print(repList + 'Successful')
print("File result created in: " + cadOutput)

# Khanhnv edit: Print ERRORS REPORT
# --------------------------------------------------------------------------------------------------------------
if (ne > 0):      
    # Maintained the original error message format from version 5
    if ne == 1 :
        error_msg = str(ne) + " ERROR of definion of \n\t branch \n\t three-winding transformer \n\t bus number \nin:\n" + errList
    else      :
        error_msg = str(ne) + " ERRORS of definion of \n\t branch \n\t three-winding transformer \n\t bus number \nin:\n" + errList
    
    # Khanhnv edit: Print error message to Console with separator
    print("\n" + "="*50)
    print("Error:")
    print(error_msg)
    print("="*50 + "\n")
    
    # Show Warning Popup
    window_warning(error_msg)
else:
    # Khanhnv edit: Success notification
    window_info("Success !!!\nFile result created in:\n" + cadOutput)