#--------------------Imports-----------------------------------
import pandas as pd
import numpy as np
import shutil
from tkinter import *
from tkinter import ttk
from PIL import ImageTk, Image
import os
import pickle
import xlsxwriter
from re import split, sub
from decimal import Decimal
import tkinter.filedialog
import datetime as dt
import winsound
from calendar import monthrange
from difflib import SequenceMatcher
from tkinter import messagebox
from firebase import  firebase
from datetime import date
from datetime import datetime
import multiprocessing
#---------------------Archive read------------------------------
def Sales(diff):
    global fact
    global total
    global reg
    global vendors
    global date_range
    global init
    try:
        if(diff == 0):
            new_file = Tk()
            new_file.withdraw()
            new_file.filename = tkinter.filedialog.askopenfilename(initialdir = "/",
                                title = "Seleccione archivo de Costo de ventas por vendedor",filetypes = (("Excel","*.xlsx"),
                                ))
            mypath = new_file.filename
            original = mypath
            target = 'file1.xlsx'
            shutil.copyfile(original, target)
        else:
            mypath = 'file1.xlsx'
        init = str(mypath).replace(str(os.path.basename(mypath)), "")
        df = pd.read_excel(mypath, header=None)
        df[0,0] = 'Total empresa'
        edit = df[df[0] == "MOTORLIGHTS S.A.S"].index.values
        df.loc[edit[0]] = "Total Vendedor"
        list = df[0].unique()
        date_range = str(list[2]).replace("Entre", "")
        vendors = [x for x in list if
                   'Total' not in str(x) and 'Costo' not in str(x) and 'Entre' not in str(x) and
                   'Vendedor' not in str(x) and str(x) != 'nan']
        lastIndex = int(df[df[0] == 'Total general'].index.values)
        index = []
        for vendor in vendors:
            index.append(int(df[df[0] == vendor].index.values))
        new_list = index[1:]
        new_list.append(lastIndex)
        array1 = np.array(index)
        array2 = np.array(new_list)
        reg = np.subtract(array2, array1)
        reg = [x - 1 for x in reg]
        newcolumns = df.loc[df[0] == 'Vendedor']
        c = newcolumns.values.tolist()
        df.columns = c[0]
        vdf = df['Ventas']
        ddf = df['Doc']
        total = []
        fact = []
        flete = []
        fdf = df['CódigoInventario']
        flete_id = "0001 Flete Mercancia"
        for i in range(len(array1)):
            v = vdf[array1[i]:array2[i] - 1]
            d = ddf[array1[i]:array2[i] - 1]
            f = fdf[array1[i]:array2[i] - 1][fdf == flete_id].index.values.astype(int)
            if(len(f) > 0):
                flt = 0
                for j in f:
                    flt = vdf.loc[j] + flt
            else:
                flt = 0
            cc = d.unique()
            actualFact = [x for x in cc if
                       'FV' not in str(x)]
            total.append(v.sum() - flt)
            fact.append(d.nunique()-len(actualFact))
        recaudo.config(state = NORMAL)
        winsound.MessageBeep()
    except Exception as e:
        messagebox.showerror("Ocurrió un error leyendo el archivo", "Favor revise que sea el correcto o contacte al administrador")
        print(e)

def Collection(diff):
    global nrdf
    global a
    try:
        if (diff == 0):
            new_file = Tk()
            new_file.withdraw()
            new_file.filename = tkinter.filedialog.askopenfilename(initialdir=init,
                                                                   title="Seleccione archivo de Libro Auxiliar",
                                                                   filetypes=(("Excel", "*.xlsx"),
                                                                              ))
            mypath = new_file.filename
            original = mypath
            target = 'file2.xlsx'
            shutil.copyfile(original, target)
        else:
            mypath = 'file2.xlsx'
        rdf = pd.read_excel(mypath, header=None)
        list = rdf[0].unique()
        check_date = str(list[1]).replace("Libro Auxiliar entre el ", "").replace("y el", "Y").replace(" ", "")
        comp = similar(date_range.strip().replace(" ",""), check_date.strip())
        if(comp > 0.9):
            pass
        else:
            print(comp)
            messagebox.showerror("Error de fecha", "Las fechas de los archivos no coincien")
            return
        nrdf = rdf[[5,6]]
        header = nrdf.columns.values
        nrdf = nrdf[pd.notna(nrdf[header[0]])]
        newheader = nrdf.iloc[0]
        nrdf = nrdf.iloc[1:]
        nrdf.columns = newheader
        a =  newheader.tolist()
        cobranza.config(state = NORMAL)
        winsound.MessageBeep()
    except Exception as e:
        messagebox.showerror("Ocurrió un error leyendo el archivo", "Favor revise que sea el correcto o contacte al administrador")
        print(e)

def Charge(diff):
    global result
    global resume
    try:
        if (diff == 0):
            new_file = Tk()
            new_file.withdraw()
            new_file.filename = tkinter.filedialog.askopenfilename(initialdir=init,
                                                                   title="Seleccione archivo de Gestión de cobranza",
                                                                   filetypes=(("Excel", "*.xlsx"),
                                                                              ))
            mypath = new_file.filename

            original = mypath
            target = 'file3.xlsx'
            shutil.copyfile(original, target)
        else:
            mypath = 'file3.xlsx'

        cdf = pd.read_excel(mypath, header=None)

        list = cdf[0].unique()
        check_date = str(list[1]).replace("Gestión de Cobranza entre el ", "").replace("y el", "Y").replace(" ", "")
        comp = similar(date_range.strip().replace(" ", ""), check_date.strip())
        if(comp > 0.9):
            pass
        else:
            messagebox.showerror("Error de fecha", "Las fechas de los archivos no coincien")
            return
        ncdf = cdf[[1, 4]]
        header = ncdf.columns.values
        resume = []
        nr = []
        iva = 1.19
        #Missing Vendor
        checkVendor = ncdf.copy()
        checkVendor.dropna(subset = [1], inplace=True)
        checkVendor.dropna(subset = [4], inplace=True)
        for newVendor in checkVendor[1]:
            if(newVendor in vendors):
                pass
            else:
                if(newVendor != 'Vendedor'):
                    print(newVendor)
                    vendors.append(newVendor)
                    total.append(0)
                    fact.append(0)
                    reg.append(0)
        for vendor in vendors:
            sum = 0
            try:
                print('Total ' + vendor)
                print(ncdf[ncdf[header[0]] == 'Total ' + vendor].index.values)
                index = int(ncdf[ncdf[header[0]] == vendor].index.values)
                _index = int(ncdf[ncdf[header[0]] == 'Total ' + vendor].index.values)
                r = ncdf[index:_index]
                search = r[header[1]].unique()
                for value in search:
                    _value = '(MS) rc ' + str(value)
                    i = nrdf[nrdf[a[0]] == _value].index.values
                    if (len(i) > 1):
                        for j in i:
                            sum = float(nrdf[a[1]][int(j)]) + sum
                    else:
                        if(len(i) != 0):
                            sum = float(nrdf[a[1]][int(i)]) + sum
                resume.append(sum / float(iva))
            except Exception as e:
                print(e)
                nr.append(vendor)
                resume.append(0)
        goals, goals_ = GetGoals()
        d = {'Vendedor': vendors, 'Registros':reg, 'Facturas':fact, 'Metas':total, 'Ventas':total, 'Recaudo (sin IVA)':resume}
        my_tree.tag_configure('oddrow', background="white")
        my_tree.tag_configure('evenrow', background="lightblue")
        count = 0
        for record in vendors:
            if (count % 2 == 0):
                my_tree.insert(parent='', index='end', iid=count, text='', values=(vendors[count], reg[count], fact[count],
                                                                                   "${:,.2f}".format(goals[count]), "${:,.2f}".format(total[count])
                                                                                   , "${:,.2f}".format(goals_[count]), "${:,.2f}".format(resume[count])), tags=('evenrow',))
            else:
                my_tree.insert(parent='', index='end', iid=count, text='', values=(vendors[count], reg[count], fact[count],
                                                                                   "${:,.2f}".format(goals[count]), "${:,.2f}".format(total[count])
                                                                                   , "${:,.2f}".format(goals_[count]), "${:,.2f}".format(resume[count])),tags=('oddrow',))
            count += 1
        result = pd.DataFrame(d)
        total_resume.config(text = result['Vendedor'].count())
        sale_resume.config(text = "${:,.2f}".format(result['Ventas'].sum()))
        fact_resume.config(text = result['Facturas'].sum())
        reca_resume.config(text = "${:,.2f}".format(result['Recaudo (sin IVA)'].sum()))
        dateLabel.config(text = 'Rango de fecha'+date_range)
        confirm.config(state = NORMAL)
        confirm_.config(state = NORMAL)
        generate.config(state = NORMAL)
        recaudo.config(state=DISABLED)
        cobranza.config(state = DISABLED)
        sales.config(state = DISABLED)
        clear.config(state = NORMAL)
        winsound.MessageBeep()
    except Exception as e:
        messagebox.showerror("Ocurrió un error leyendo el archivo", "Favor revise que sea el correcto o contacte al administrador")
        print(e)

def similar(string1, string2):
    return SequenceMatcher(None, string1, string2).ratio()

def Update():
    selected = my_tree.focus()
    values = my_tree.item(selected, 'values')
    lst = list(values)
    lst[3] = meta.get()
    my_tree.item(selected, text = "", values = (lst[0],lst[1],lst[2],"${:,.2f}".format(float(lst[3])),lst[4],lst[5], lst[6]))
    database.put('Target/Venta', lst[0], "${:,.2f}".format(float(lst[3])))

def GetGoals():
    metasVentas = database.get('/Target/Venta', None)
    metasRecaudo = database.get('/Target/Recaudo', None)
    goal = []
    goal_ = []
    for v in vendors:
        try:
            value = Decimal(sub(r'[^\d.]', '', metasVentas[str(v)]))
            goal.append(value)
        except:
            goal.append(0)
        try: 
            value_ = Decimal(sub(r'[^\d.]', '', metasRecaudo[str(v)]))
            goal_.append(value_)
        except:
            goal_.append(0)
    return  goal, goal_
            
def WriteGoals():
    print('de1')
    filename = 'targets'
    infile = open(filename, 'rb')
    new_dict = pickle.load(infile)
    check = list(new_dict['Vendedor'])
    check_goal = list(new_dict['Metas'])
    check_goal_ = list(new_dict['Metas_'])
    new_goal = []
    new_reca_goal = []
    for child in my_tree.get_children():
        param = my_tree.item(child)["values"]
        index = check.index(str(param[0]))
        check_goal[index] = param[3]
        check_goal_[index] = param[5]
        new_goal.append(param[3])
        new_reca_goal.append(param[5])
    update_dic = {'Vendedor':check, 'Metas':check_goal,'Metas_': check_goal_}
    #datos = {'id':'12', 'dato1':'testing', 'dato25':'working'}
    outfile = open(filename, 'wb')
    pickle.dump(update_dic, outfile)
    outfile.close()
    return new_goal, new_reca_goal

def Update_():
    selected = my_tree.focus()
    values = my_tree.item(selected, 'values')
    lst = list(values)
    lst[5] = meta.get()
    my_tree.item(selected, text = "", values = (lst[0],lst[1],lst[2],lst[3],lst[4],"${:,.2f}".format(float(lst[5])), lst[6]))
    database.put('Target/Recaudo', lst[0], "${:,.2f}".format(float(lst[5])))

def Callback(event):

    index = str(my_tree.selection()).replace("(","").replace("',)","").replace("'","")
    row = my_tree.item(index,'values')
    selected_vendor.config(text = str(row[0]))

def ExpenseCallback(event):
    deleteEntry.config(state=NORMAL)
    global mySelection, mytreeSelection
    mytreeSelection = ExpenseTree.selection()
    index = str(ExpenseTree.selection()).replace("(","").replace("',)","").replace("'","")
    row = list(ExpenseTree.item(index,'values'))
    print(str(row[1]).replace('$', '').replace('.00',''))
    row[1] = str(row[1]).replace('$', '').replace('.00','').replace(',','')
    mySelection = ';'.join(row)
    print(mySelection)
    #selected_vendor.config(text = str(row[0]))

def MonthName(month):
    if (month == 1):
        return 'ENERO'
    if (month == 2):
        return 'FEBRERO'
    if (month == 3):
        return 'MARZO'
    if (month == 4):
        return 'ABRIL'
    if (month == 5):
        return 'MAYO'
    if (month == 6):
        return 'JUNIO'
    if (month == 7):
        return 'JULIO'
    if (month == 8):
        return 'AGOSTO'
    if (month == 9):
        return 'SEPTIEMBRE'
    if (month == 10):
        return 'OCTUBRE'
    if (month == 11):
        return 'NOVIEMBRE'
    if (month == 12):
        return 'DICIEMBRE'

def Generate():
    try:
        totalVenta = 0
        totalFacturas = 0
        totalMetaVenta = 0
        totalRecaudo = 0
        totalMetaRecaudo = 0
        save_path = tkinter.filedialog.asksaveasfile(defaultextension=".xlsx", filetypes=(("Excel", "*.xlsx"),))
        #print(save_path.name)
        range = date_range.split("Y")
        days = LaborDays(range[0], range[1])+2
        cmonth, cyear, cday = ActiveDays(range[1])
        mes = MonthName(int(cmonth))
        final_date = str(cday)+'/'+str(cmonth)+'/'+str(cyear)
        days_ = int(LaborDays(range[0], final_date))
        ugoal, rgoal = GetGoals()
        print(int(rgoal[0]))
        image_width = 310.0
        image_height = 182.0
        cell_width = 64.0
        cell_height = 20.0
        x_scale = cell_width / image_width
        y_scale = cell_height / image_height
        workbook = xlsxwriter.Workbook(save_path.name)
        normal_format = workbook.add_format({'bold': True, 'border':1})
        normal_money_format =  workbook.add_format({'bold': True, 'num_format': '$#,##0', 'border':1})
        normal_percentage_format = workbook.add_format({'bold': True, 'num_format': '0%', 'border': 1})
        title_format = workbook.add_format({'bold': True, 'bg_color': 'yellow', 'border':1})
        main_format = workbook.add_format({'bold': True, 'bg_color': '#C2C2C2', 'border':1})
        result_format = workbook.add_format({'bold': True, 'bg_color': 'black', 'font_color': 'white', 'border':1})
        result_money_format = workbook.add_format({'bold': True, 'num_format': '$#,##0' ,'bg_color': 'black', 'font_color': 'white', 'border':1})
        main_money_format = workbook.add_format({'num_format': '$#,##0', 'bg_color': 'yellow', 'bold': True, 'border':1})
        money_format = workbook.add_format({'num_format': '$#,##0', 'bg_color': '#C2C2C2', 'bold': True, 'border':1})
        percentage_format = workbook.add_format({'num_format': '0%', 'bg_color': '#C2C2C2', 'bold': True, 'border':1})
        main_percentage_format = workbook.add_format({'num_format': '0%', 'bg_color': 'yellow', 'bold': True, 'border': 1})
        # add resume data titles
        resume_worksheet = workbook.add_worksheet(name="RESUMEN")
        resume_worksheet.set_column('B:B', 25)
        resume_worksheet.write('B2', mes, title_format)
        resume_worksheet.write('B3','=C21', main_percentage_format)
        resume_worksheet.write('B4', 'TOTAL VENTA', title_format)
        resume_worksheet.write('B5', 'Cantidad de facturas', main_format)
        resume_worksheet.write('B6', 'Promedio de venta', main_format)
        resume_worksheet.write('B7', 'Meta de venta', result_format)
        resume_worksheet.write('B8', '%VENTA', normal_format)
        resume_worksheet.write('B9', 'Pendiente', normal_format)
        resume_worksheet.write('B10','',result_format)
        resume_worksheet.write('B11', 'TOTAL RECAUDO',title_format)
        resume_worksheet.write('B12', 'META RECAUDO SIN IVA', result_format)
        resume_worksheet.write('B13', '% RECAUDO', normal_format)
        resume_worksheet.write('B14', 'Pendiente', normal_format)
        resume_worksheet.write('B19', 'Días hábiles del mes', normal_format)
        resume_worksheet.write('B20', 'Transcurridos', normal_format)
        resume_worksheet.write('D18', 'Días festivos', normal_format)
        resume_worksheet.write('D19', '0', normal_format)
        resume_worksheet.write('C19', '='+str(days_ + 1)+'-D19', normal_format)
        resume_worksheet.write('C20', '='+str(days-1)+'-D19', normal_format)
        resume_worksheet.write('C21', '=C20/C19', percentage_format)
        resume_worksheet.write('F19', date_range, normal_format)
        resume_worksheet.write('B24', 'AVANCE DEL MES', title_format)
        resume_worksheet.write('C24', '=C21', main_percentage_format)
        resume_worksheet.write('B25', '=B2', title_format)
        resume_worksheet.write('B26', '%VENTA', normal_format)
        resume_worksheet.write('B27', '%RECAUDO', normal_format)
        resume_worksheet.set_header('&L&G', {'image_left': 'logo.png'})
        resume_worksheet.hide_gridlines(2)
        counter = len(vendors)
        i = 0
        for v in vendors:
            # add data per vendor
            sheet_name = str(vendors[i]).split(" ")
            worksheet = workbook.add_worksheet(name = sheet_name[0] + " " + sheet_name[1])
            worksheet.hide_gridlines(2)
            worksheet.set_column('A:A', 30)
            worksheet.set_column('B:B', 30)
            worksheet.write('A1', 'EJECUTIVO', title_format)
            worksheet.write('B1', sheet_name[0] + " " + sheet_name[1] , title_format)
            worksheet.write('A2', '=RESUMEN!B2', title_format)
            worksheet.write('B2', '=RESUMEN!C21', main_percentage_format)
            worksheet.write('A3', 'TOTAL VENTA', title_format)
            worksheet.write('B3', total[i], main_money_format)
            worksheet.write('A4', 'Cantidad de facturas', main_format)
            worksheet.write('B4', fact[i], main_format)
            worksheet.write('A5', 'Promedio de venta', main_format)
            worksheet.write('B5', '=B3/B4', money_format)
            worksheet.write('A6', 'Meta de venta', main_format)
            goal = ugoal[i]
            #goal = Decimal(sub(r'[^\d.]', '', int(ugoal[i])))
            worksheet.write('B6', goal, money_format)
            worksheet.write('A7', '% de venta', main_format)
            worksheet.write('B7', '=B3/B6', percentage_format)
            worksheet.write('A8', 'Pendiente', main_format)
            worksheet.write('B8', '=B6-B3', main_money_format)
            worksheet.write('A10', 'TOTAL RECAUDO', title_format)
            worksheet.write('B10', resume[i], main_money_format)
            worksheet.write('A11', 'META RECAUDO SIN IVA', main_format)
            goal_=rgoal[i]
            #goal_ = Decimal(sub(r'[^\d.]', '', int(rgoal[i])))
            worksheet.write('B11', goal_, money_format)
            worksheet.write('A12', 'PORCENTAJE DE RECAUDO', main_format)
            worksheet.write('B12', '=B10/B11', percentage_format)
            worksheet.write('A13', 'Pendiente', main_format)
            worksheet.write('B13', '=B11-B10', main_money_format)
            worksheet.insert_image('E2', 'logo.png',
                                   {'x_scale': x_scale, 'y_scale': y_scale})
            #add data to resume
            resume_worksheet.set_column(i + 2, i + 2, 20)
            resume_worksheet.write(2, i + 2, sheet_name[0] + " " + sheet_name[1], result_format)
            resume_worksheet.write(3, i + 2, total[i], main_money_format)
            resume_worksheet.write(4, i + 2, fact[i], main_format)
            if(fact[i] == 0):
                resume_worksheet.write(5, i + 2, 'DIV/0', money_format)
            else:
                resume_worksheet.write(5, i + 2, total[i]/fact[i], money_format)
            resume_worksheet.write(6, i + 2, goal, result_money_format)
            if (goal != 0):
                per = float(total[i]) / float(goal)
            else:
                per = "DIV/0"
            pendiente = float(goal) - float(total[i])
            resume_worksheet.write(7, i + 2, per, normal_percentage_format)
            resume_worksheet.write(8, i + 2, pendiente,  normal_money_format)
            resume_worksheet.write(9, i + 2, '', result_format)
            resume_worksheet.write(10,i + 2, resume[i], main_money_format)
            resume_worksheet.write(11, i + 2, goal_, result_money_format)
            if (goal_ != 0):
                per_ = float(resume[i]) / float(goal_)
            else:
                per_ = "DIV/0"
            resume_worksheet.write(12, i + 2, per_, normal_percentage_format)
            pendiente = float(goal_) - float(resume[i])
            resume_worksheet.write(13, i + 2, pendiente, normal_money_format)
            resume_worksheet.write(24, i + 2, sheet_name[0] + " " + sheet_name[1], result_format)
            resume_worksheet.write(25, i + 2, per, normal_percentage_format)
            resume_worksheet.write(26, i + 2, per_, normal_percentage_format)
            totalVenta = totalVenta + total[i]
            totalFacturas = totalFacturas + fact[i]
            totalMetaVenta = totalMetaVenta + goal
            totalRecaudo = totalRecaudo + resume[i]
            totalMetaRecaudo = totalMetaRecaudo + goal_
            totalPendienteRecaudo = 0
            i += 1
        totalPromedioVentas = float(totalVenta)/float(totalFacturas)
        totalPendiente = float(totalMetaVenta) - float(totalVenta)
        totalPendienteRecaudo = float(totalMetaRecaudo) - float(totalRecaudo)
        resume_worksheet.set_column(i + 2, i + 2, 20)
        resume_worksheet.write(2, i + 2, 'TOTAL', result_format)
        resume_worksheet.write(3, i + 2, totalVenta, main_money_format)
        resume_worksheet.write(4, i + 2, totalFacturas, main_format)
        resume_worksheet.write(5, i + 2, totalPromedioVentas, money_format)
        resume_worksheet.write(6, i + 2, totalMetaVenta, result_money_format)
        if(totalMetaVenta == 0):
            resume_worksheet.write(7, i + 2, "DIV/0", percentage_format)
        else:
            resume_worksheet.write(7, i + 2, float(totalVenta)/float(totalMetaVenta), percentage_format)
        resume_worksheet.write(8, i + 2, totalPendiente, normal_money_format)
        resume_worksheet.write(9, i + 2, "", result_format)
        resume_worksheet.write(10, i + 2, totalRecaudo, main_money_format)
        resume_worksheet.write(11, i + 2, totalMetaRecaudo, result_money_format)
        if(float(totalMetaRecaudo) == 0):
            resume_worksheet.write(12, i + 2, "DIV/0", percentage_format)
        else:
            resume_worksheet.write(12, i + 2, float(totalRecaudo)/float(totalMetaRecaudo), percentage_format)
        resume_worksheet.write(13, i + 2, totalPendienteRecaudo, normal_money_format)
        resume_worksheet.insert_image('H19', 'logo.png',
                               {'x_scale': x_scale, 'y_scale': y_scale})
        workbook.close()
        winsound.MessageBeep()
    except AssertionError as error:
        messagebox.showerror("Se detectó un error", str(error))

def Clear():
    selection = my_tree.get_children()
    for item in selection:
        my_tree.delete(item)
    selection = my_tree.get_children()
    for item in selection:
        my_tree.delete(item)
    sales.config(state = NORMAL)
    confirm.config(state = DISABLED)
    confirm_.config(state = DISABLED)
    generate.config(state = DISABLED)
    total_resume.config(text = "")
    sale_resume.config(text = "")
    fact_resume.config(text = "")
    reca_resume.config(text = "")

def LaborDays(date, date2):
    first_date = str(date).split("/")
    start = dt.date(int(first_date[2]), int(first_date[1]), int(first_date[0]))
    second_date = str(date2).split("/")
    end = dt.date(int(second_date[2]), int(second_date[1]), int(second_date[0]))
    days = np.busday_count(start, end, weekmask='Mon Tue Wed Thu Fri Sat')
    return days

def ActiveDays(current_date):
    datee = dt.datetime.strptime(str(current_date).strip(), "%d/%m/%Y")
    num_days = monthrange(datee.year, datee.month)[1]
    return datee.month, datee.year, num_days

def GetVendor():
    lists = database.get('Target/','')
    if(len(lists['Recaudo']) > len(lists['Venta'])):
        vendedores = lists['Recaudo']
    else:
        vendedores = lists['Venta']
    print(list(vendedores.keys()))
    return  list(vendedores.keys())

def SelectionBonus():
   selection = "You selected the option " + str(var.get())
   print(selection)

def SelectedVendor(opt):
    global val
    val = string_variable.get()

def GetBonus():
    global val, bonusOne, bonusTwo, bonusThree, bonusFour, firstRange, string_variable
    firstRange.delete(0, END)
    bonusOne.delete(0, END)
    bonusTwo.delete(0, END)
    bonusThree.delete(0, END)
    bonusFour.delete(0, END)
    bonus = database.get('Bonus/', str(string_variable.get()))
    if (bonus != 'none'):
        data = str(bonus).split(',')
        firstRange.insert(0,data[0])
        bonusOne.insert(0,data[1])
        bonusTwo.insert(0,data[2])
        bonusThree.insert(0,data[3])
        bonusFour.insert(0,data[4])

def SaveBonus():
    global val, bonusOne, bonusTwo, bonusThree, bonusFour, firstRange, string_variable
    database.put('Bonus/', str(string_variable.get()), str(firstRange.get())+','+str(bonusOne.get())+','+str(bonusTwo.get())+','+str(bonusThree.get())+','+str(bonusFour.get()))
    
def ValidateEmpty():
    if (var.get() == 1):
        if(min0.get() != "" and min1.get() != "" and min2.get() != "" and min3.get() != "" and min4.get() != ""
                and max0.get() != "" and max1.get() != "" and max2.get() != "" and max3.get() != "" and max4.get() != "" and max5.get() != ""
                and zero_bonus.get() != "" and first_bonus.get() != "" and second_bonus.get() != "" and third_bonus.get() != ""
                and fourth_bonus.get() != "" and fifth_bonus.get() != ""):
            return True
        else:
            return False
    else:
        if (min0.get() != "" and min1.get() != "" and min2.get() != ""  and max0.get() != ""
                and max1.get() != "" and max2.get() != "" and max3.get() != ""
                and zero_bonus.get() != "" and first_bonus.get() != "" and second_bonus.get() != "" and third_bonus.get() != ""):
            return  True
        else:
            return False

def RangesGui():
    top = Toplevel()
    top.title("MotorLight - Rangos de incentivos por vendedor")
    top.geometry(str(int(sw*0.4))+"x"+str(int(sh*0.40)))
    top.resizable(0, 0)
    my_notebook = ttk.Notebook(top)
    my_notebook.pack(pady = 15, fill='both', expand = 1)
    my_frame1 = Frame(my_notebook, width = int(sw*0.3), height = int(sh*0.40))
    my_frame1.pack(fill='both', expand = 1)
    my_notebook.add(my_frame1, text = 'Incentivos de Vendedores')
    #FRAME VENTAS GUI

    OPTIONS = GetVendor()
    global string_variable, var, downloadBonus
    string_variable = StringVar(top)
    var = IntVar()
    string_variable.set(OPTIONS[1])  # default value
    staticLabel = Label(my_frame1, text = "Seleccione vendedor")
    staticLabel.grid(row = 0, column = 1)
    w = OptionMenu(my_frame1, string_variable, *OPTIONS, command = SelectedVendor)
    w.grid(row = 1, column = 1, padx=int(sw*0.003255))
    w.config(width = 28, font=('Helvetica',7), anchor = W)
    actionType = LabelFrame(my_frame1, text ="Seleccione acción")
    actionType.grid(row = 0, column = 2, rowspan = 2, columnspan=4)
    #IMAGE

    canvas = Canvas(my_frame1, width=int(sw * 0.2), height=int(sh * 0.35))
    canvas.grid(row = 3, column = 0, columnspan = 6, sticky="nsew")
    img = Image.open("logo.png")  # PIL solution
    img = img.resize((int(sw * 0.15), int(sh * 0.1)), Image.ANTIALIAS)  # The (x, y) is (height, width)
    img = ImageTk.PhotoImage(img)  # convert to PhotoImage
    canvas.create_image(5, 5, anchor=NW, image=img)

    #ACTION FRAME
    insertBonus = Button(actionType, text = "Ingresar incentivos",  command = lambda : SalesBonusInsert(my_frame1, canvas))
    insertBonus.pack(anchor = W, padx = 2, pady= 2)
    downloadBonus = Button(actionType, text = "Ver incentivos", state = DISABLED, command = lambda : GetBonus())
    downloadBonus.pack(anchor = W, padx = 2, pady = 2)

    top.mainloop()

def SalesBonusInsert(my_frame1, canvas):
    global bonusOne, bonusTwo, bonusThree, bonusFour, firstRange
    global collectionBonus, saleBonus, check
    canvas.grid_forget()
    downloadBonus.config(state = NORMAL)
    # SALES FORM FRAME
    sales_form = LabelFrame(my_frame1, text="Estructura de incentivos")
    sales_form.grid(row = 3, column = 0, columnspan = 3, sticky="nsew")
    sales_form.config(width = int(sw*0.4))
    #BONUS
    bonusTitle = Label(sales_form, text ="")
    bonusTitle.grid(row = 0, column = 0)
    premise = Label(sales_form, text = "* PREMISA GENERAL: CUMPLIMIENTO SUPERIOR A: ")
    premise.grid(row = 1, column = 0, columnspan=5)
    firstRange = Entry(sales_form, justify=RIGHT, width=8)
    firstRange.grid(row = 1, column=5)
    p1 = Label(sales_form, text ="% DEL RECAUDO")
    p1.grid(row=1, column=6)
    secondTitle = Label(sales_form, text="        -> UMBRAL DE VENTA $", justify=LEFT)
    secondTitle.grid(row = 2, column = 0)
    bonusOne = Entry(sales_form, width=15)
    bonusOne.grid(row=2, column=1)
    thirTitle = Label(sales_form, text=" DE PESOS ", justify=LEFT)
    thirTitle.grid(row = 2, column = 2)
    title_4 = Label(sales_form, text="--> MAYOR:    ")
    title_4.grid(row = 3, column = 0)
    title_4_ = Label(sales_form, text="% DEL RECAUDO")
    title_4_.grid(row = 3, column = 2)
    title_5 = Label(sales_form, text="--> MENOR:    ")
    title_5.grid(row = 4, column = 0)
    title_5_ = Label(sales_form, text="% DEL RECAUDO")
    title_5_.grid(row = 4, column = 2)
    title_6 = Label(sales_form, text="* INCUMPLIMIENTO UMBRAL DE RECUADO            ")
    title_6.grid(row = 5, column = 0, columnspan=5)
    title_7 = Label(sales_form, text="--> RECAUDO:  ")
    title_7.grid(row = 6, column = 0)
    title_7_ = Label(sales_form, text="%")
    title_7_.grid(row = 6, column = 2)
    bonusTwo = Entry(sales_form, width=8)
    bonusTwo.grid(row=3, column=1)
    bonusThree = Entry(sales_form, width=8)
    bonusThree.grid(row=4, column=1)
    bonusFour = Entry(sales_form, width=8)
    bonusFour.grid(row=6, column=1)
    addBonus = Button(sales_form, text='Aprobar', width=20, command=lambda: SaveBonus())
    addBonus.grid(row=7, column=0, columnspan=6, sticky=NSEW, pady=20, padx=30)

def CalculateBonus():
    save_path = tkinter.filedialog.asksaveasfile(defaultextension=".xlsx", filetypes=(("Excel", "*.xlsx"),))
    workbook = xlsxwriter.Workbook(save_path.name)
    normal_money_format =  workbook.add_format({'bold': True, 'num_format': '$#,##0', 'border':1})
    normal_format = workbook.add_format({'bold': True, 'border':1})
    main_money_format = workbook.add_format({'num_format': '$#,##0', 'bg_color': 'yellow', 'bold': True, 'border':1})
    title_format = workbook.add_format({'bold': True, 'bg_color': 'yellow', 'border':1})
    percentage_format = workbook.add_format({'num_format': '0%', 'bg_color': '#C2C2C2', 'bold': True, 'border':1})
    resume_worksheet = workbook.add_worksheet(name="INCENTIVO")
    resume_worksheet.set_column('A:A', 30)
    resume_worksheet.set_column('B:B', 8)
    resume_worksheet.set_column('C:C', 8)
    resume_worksheet.set_column('D:D', 15)
    resume_worksheet.set_column('E:E', 15)
    resume_worksheet.set_column('F:F', 15)
    resume_worksheet.set_column('G:G', 15)
    resume_worksheet.set_column('H:H', 8)
    resume_worksheet.set_column('I:I', 8)
    resume_worksheet.set_column('J:J', 8)
    resume_worksheet.write('A1', 'Vendedor', title_format)
    resume_worksheet.write('B1', 'Registros', title_format)
    resume_worksheet.write('C1', 'Facturas', title_format)
    resume_worksheet.write('D1', 'Meta de Venta', title_format)
    resume_worksheet.write('E1', 'Venta', title_format)
    resume_worksheet.write('F1', 'Meta de Recaudo', title_format)
    resume_worksheet.write('G1', 'Recaudo', title_format)
    resume_worksheet.write('H1', '% Venta', title_format)
    resume_worksheet.write('I1', '% Recaudo', title_format)
    resume_worksheet.write('J1', 'Incentivo', title_format)
    ids = my_tree.get_children()
    export = []
    counter = 2
    for _id in ids:
        row = my_tree.item(_id)
        values = row['values']
        resume_worksheet.write('A'+str(counter), values[0], normal_format)
        resume_worksheet.write('B'+str(counter), values[1], normal_format)
        resume_worksheet.write('C'+str(counter), values[2], normal_format)
        metaVenta = float(str(values[3]).replace('$', '').replace(',', ''))
        venta = float(str(values[4]).replace('$', '').replace(',', ''))
        metaRecaudo = float(str(values[5]).replace('$', '').replace(',', ''))
        recaudo = float(str(values[6]).replace('$', '').replace(',', ''))
        incentivo = GetVendorBonus(values[0], metaVenta, metaRecaudo, venta, recaudo)
        print(incentivo)
        print(metaVenta)
        resume_worksheet.write('D'+str(counter), metaVenta, normal_money_format)
        resume_worksheet.write('E'+str(counter), venta, normal_money_format)
        resume_worksheet.write('F'+str(counter), metaRecaudo, normal_money_format)
        resume_worksheet.write('G'+str(counter), recaudo, normal_money_format)
        resume_worksheet.write('H'+str(counter), float(venta/metaVenta), percentage_format)
        resume_worksheet.write('I'+str(counter), float(recaudo/metaRecaudo), percentage_format)
        resume_worksheet.write('J'+str(counter), incentivo, main_money_format)
        counter = counter +1
    workbook.close()

def Bonus(meta, valor, min, max, tipo):
    min = float(min)/100
    max = float(max)/100
    valor = float(valor)
    meta = float(meta)
    if(tipo == 1):
        if(meta == 0):
            return False
        else:
            if(tipo == 1):
                reach = valor/meta
            else:
                reach = valor
            print("Cumplimiento "+str(reach))
            if(reach >= min and reach <= max):
                return True
            else:
                return False
    else:
        if (valor >= min and valor <= max):
            return True
        else:
            return False

def GenerateBonus():
    CalculateBonus()

def GetVendorBonus(nombre, metaVenta, metaRecaudo, venta, recaudo):
    bonus = database.get('Bonus/', str(nombre))
    print(bonus)
    if (bonus != 'none'):
        data = str(bonus).split(',')
        umbralRecaudo = float(data[0])/100
        umbralVenta = float(data[1])
        ventaMayor = float(data[2])/100
        ventaMenor = float(data[3])/100
        recaudoMenor= float(data[4])/100
        umbral = float(recaudo/metaRecaudo)
        if(umbral > umbralRecaudo):
            if(venta > umbralVenta):
                incentivo = float(ventaMayor*recaudo)
            else:
                incentivo = float(ventaMenor*recaudo)
        else:
            incentivo = float(recaudoMenor*recaudo)
    else:
        incentivo = 0
    return incentivo

def Load():
    Sales(1)
    Collection(1)
    Charge(1)

def Expenses():
    global tableName, confirm1, yearSelect, motnhSelect, loadMonthselect, confirmLoad, ExpenseTree, TreeIndex, entry3, entry2, entry4, ExpenseTree, deleteEntry
    test = 0
    TreeIndex = 0
    loadMonth = []
    loadList = []
    result = database.get('/', 'Gastos')
    if(result != None):
        for r in result:
            r1 = str(r)
            loadMonth.append(r1)
        loadList = unique(loadMonth)
        print(loadList)
        status = NORMAL
    else:
        status = DISABLED
        loadlist = ['Vacío']
    master = Toplevel()
    master.title("MotorLight - Control de gastos")
    master.geometry(str(int(sw*0.6))+"x"+str(int(sh*0.55)))
    master.resizable(0, 0)
    main = LabelFrame(master, text = 'Seleccione opción')
    main.grid(row= 0, column= 0, pady=int(sh*0.003472), padx=int(sw*0.001953125))
    newButton = Button(main, text='Crear', width = int(sw*0.01171875), command= lambda : CreateControl())
    newButton.pack(pady=int(sh*0.003472), padx=int(sw*0.001953125))
    loadButton = Button(main, text ='Cargar', width = int(sw*0.01171875), command= lambda : LoadControl())
    loadButton.pack(pady=int(sh*0.003472), padx=int(sw*0.001953125))
    create = LabelFrame(master, text = 'Seleccione fecha')
    create.grid(row = 0, column=1, pady=int(sh*0.003472), padx=int(sw*0.001953125))
    MONTHS = [
                "Enero",
                "Febrero",
                "Marzo",
                "Abril",
                "Mayo",
                "Junio",
                "Julio",
                "Agosto",
                "Septiembre",
                "Octubre",
                "Noviembre",
                "Diciembre"
                ]
    variable = StringVar(master)
    variable.set(MONTHS[0]) # default value
    motnhSelect = OptionMenu(create, variable, *MONTHS)
    motnhSelect.configure(state=DISABLED)
    motnhSelect.grid(row = 0, column=0, pady=int(sh*0.003472), padx=int(sw*0.001953125))
    YEARS = [2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030]
    _variable = IntVar(master)
    _variable.set(YEARS[0])
    yearSelect = OptionMenu(create, _variable, *YEARS)
    yearSelect.configure(state=DISABLED)
    yearSelect.grid(row = 0, column=1,pady=int(sh*0.003472), padx=int(sw*0.001953125))
    confirm1 = Button(create, text='Confirmar',  width = int(sw*0.01171875), state=DISABLED, command= lambda : FirebaseCreate(variable, _variable))
    confirm1.grid(row=1, column=0, columnspan=2, pady=int(sh*0.003472), padx=int(sw*0.001953125))
    loadCanvas = LabelFrame(master, text = 'Seleccione archvio a cargar')
    loadCanvas.grid(row = 0, column=2, pady=int(sh*0.003472), padx=int(sw*0.001953125))
    loadVariable = StringVar(master)
    loadVariable.set(loadList[0])
    loadMonthselect = OptionMenu(loadCanvas, loadVariable, *loadList)
    loadMonthselect.configure(state=DISABLED)
    loadMonthselect.grid(row = 0, column=0, pady=int(sh*0.003472), padx=int(sw*0.001953125))
    confirmLoad = Button(loadCanvas, text = 'Confirmar', width = int(sw*0.01171875), state=DISABLED, command= lambda: FirebaseLoad(loadVariable))
    confirmLoad.grid(row=1, column=0, columnspan=2, pady=int(sh*0.003472), padx=int(sw*0.001953125))
    #TreeView Expenses
    ExpensesFrame = Frame(master)
    ExpenseScroll = Scrollbar(ExpensesFrame)
    ExpenseScroll.pack(side = RIGHT, fill = Y)
    ExpensesFrame.grid(row = 1, column = 0, columnspan = 3, pady=5, padx=5)
    ExpenseTree = ttk.Treeview(ExpensesFrame, yscrollcommand = ExpenseScroll.set)
    ExpenseTree.pack(pady= 5, padx=5)
    ExpenseTree.bind('<<TreeviewSelect>>', ExpenseCallback)
    ExpenseTree.config(selectmode = 'browse')
    ExpenseScroll.config(command = ExpenseTree.yview)
    ExpenseTree['columns'] = ('Concepto', 'Detalle', 'Valor', 'Banco', 'Descripción', 'Fecha')
    ExpenseTree.column('#0', width = 0)
    ExpenseTree.column('Concepto', anchor = W, width = int(sw*0.09))
    ExpenseTree.column('Detalle', anchor = W, width = int(sw*0.09))
    ExpenseTree.column('Valor', anchor = CENTER, width = int(sw*0.09))
    ExpenseTree.column('Banco', anchor = CENTER, width = int(sw*0.09))
    ExpenseTree.column('Descripción', anchor = CENTER, width = int(sw*0.12))
    ExpenseTree.column('Fecha', anchor = CENTER, width = int(sw*0.07))
    #Headings
    ExpenseTree.heading('#0', text = 'ID')
    ExpenseTree.heading('Concepto', text = 'Concepto')
    ExpenseTree.heading('Detalle', text = 'Detalle')
    ExpenseTree.heading('Valor', text = 'Valor')
    ExpenseTree.heading('Banco', text = 'Banco')
    ExpenseTree.heading('Descripción', text = 'Descripción')
    ExpenseTree.heading('Fecha', text = 'Fecha')
    #Data
    #Acton buttons
    action = LabelFrame(master, text = 'Acciones',)
    action.grid(row = 2, column = 0, columnspan = 3, padx=5, pady=5)
    send = Button(action, text = 'Resumen', width = int(sw*0.01171875), command= lambda: ExportResume())
    send.grid(row = 0, column = 0, padx = 5, pady = 2)
    deleteEntry = Button(action, text = 'Eliminar', width = int(sw*0.01171875), state=DISABLED, command = lambda: DeleteFromDataBase())
    deleteEntry.grid(row = 0, column = 1, padx = 5, pady = 2)
    ExportData = Button(action, text = 'Exportar', width = int(sw*0.01171875), command= lambda: ExportdataToExcel())
    ExportData.grid(row = 0, column = 2, padx = 5, pady = 2)
    #Import
    importData = Button(action, text = 'Importar', width = int(sw*0.01171875), command=lambda: ImportData())
    importData.grid(row = 0, column= 3, padx=5, pady=2)

def ImportData():
    global TreeIndex, tableName
    ExpenseTree.tag_configure('oddrow', background="white")
    ExpenseTree.tag_configure('evenrow', background="lightblue")
    new_file = Tk()
    new_file.withdraw()
    new_file.filename = tkinter.filedialog.askopenfilename(initialdir = "/",
                        title = "Seleccione archivo de control de gastos",filetypes = (("Excel","*.xlsx"),
                        ))
    mypath = new_file.filename
    df = pd.read_excel(mypath, sheet_name='Formato')
    for index, row in df.iterrows():
        data = row.values.tolist()
        myconcept = data[0]
        mydetail = data[1]
        myvalue = data[2]
        mybank = data[3]
        mydesc = data[4]
        mydate = data[5]
        myiid = str(TreeIndex)
        if (TreeIndex % 2 == 0):
            ExpenseTree.insert(parent='', index='end', iid=myiid, text='', values=(myconcept, mydetail,"${:,.2f}".format(myvalue), mybank, mydesc, mydate), tags=('evenrow',))
        else:
            ExpenseTree.insert(parent='', index='end', iid=myiid, text='', values=(myconcept, mydetail ,"${:,.2f}".format(myvalue), mybank, mydesc, mydate), tags=('oddrow',))
        TreeIndex += 1
        pushdata = str(myconcept)+';'+str(mydetail)+';'+str(myvalue)+';'+str(mybank)+';'+str(mydesc)+';'+str(mydate)
        print(pushdata)
        key = datetime.now().strftime("%Y%m%d_%H:%M:%S")
        database.put('/Gastos/'+tableName, str(key).replace(' ', '')+'_'+str(index), pushdata)

def InsertExpense(cnp, bnk):
    global TreeIndex, tableName
    ExpenseTree.tag_configure('oddrow', background="white")
    ExpenseTree.tag_configure('evenrow', background="lightblue")
    myconcept = cnp.get()
    myvalue = int(entry2.get())
    mybank = bnk.get()
    mydesc = entry3.get()
    mydate = entry4.get()
    myiid = str(TreeIndex)
    if (TreeIndex % 2 == 0):
        ExpenseTree.insert(parent='', index='end', iid=myiid, text='', values=(myconcept,"","${:,.2f}".format(myvalue), mybank, mydesc, mydate), tags=('evenrow',))
    else:
         ExpenseTree.insert(parent='', index='end', iid=myiid, text='', values=(myconcept,"" ,"${:,.2f}".format(myvalue), mybank, mydesc, mydate), tags=('oddrow',))
    TreeIndex += 1
    pushdata = str(myconcept)+';'+str(myvalue)+';'+str(mybank)+';'+str(mydesc)+';'+str(mydate)
    print(pushdata)
    key = datetime.now().strftime("%Y%m%d_%H:%M:%S")
    database.put('/Gastos/'+tableName, str(key).replace(' ', ''), pushdata)

def CreateControl():
    global TreeIndex
    print('create')
    yearSelect.config(state=NORMAL)
    motnhSelect.config(state=NORMAL)
    confirm1.config(state=NORMAL)
    loadMonthselect.config(state=DISABLED)
    confirmLoad.config(state=DISABLED)
    TreeIndex = 0

def LoadControl():
    global TreeIndex
    print('Load')
    loadMonthselect.config(state=NORMAL)
    confirmLoad.config(state=NORMAL)
    yearSelect.config(state=DISABLED)
    motnhSelect.config(state=DISABLED)
    confirm1.config(state=DISABLED)
    TreeIndex = 0
    
def unique(list1):
    myset = set(list1)
    mynewlist = list(myset)
    return mynewlist

def FirebaseCreate(month, year):
    global tableName
    print(month.get())
    result = database.get('Gastos/', str(month.get())+'_'+str(year.get()))
    if (result == None):
        database.put('Gastos/'+str(month.get())+'_'+str(year.get()), 'toDelete', 'toDelete1')
    else:
        messagebox.showerror("Esta tabla ya existe", "Favor use la opción cargar")
    tableName = str(month.get())+'_'+str(year.get())
    fecha = datetime.now().strftime("%d/%m/%Y")
    entry4.insert(0, str(fecha))
    selection = ExpenseTree.get_children()
    for item in selection:
        ExpenseTree.delete(item)
    selection = ExpenseTree.get_children()
    for item in selection:
        ExpenseTree.delete(item)

def FirebaseLoad(table):
    global tableName, TreeIndex, fireBaseData
    tableName = table.get()
    selection = ExpenseTree.get_children()
    print(selection)
    for item in selection:
        ExpenseTree.delete(item)
    selection = ExpenseTree.get_children()
    for item in selection:
        ExpenseTree.delete(item)
    ExpenseTree.tag_configure('oddrow', background="white")
    ExpenseTree.tag_configure('evenrow', background="lightblue")
    fireBaseData = database.get('Gastos/',str(table.get()))
    for guide in fireBaseData:
        print(guide)
        try:
            c1, c2, c3, c4, c5, c6 = str(fireBaseData[guide]).split(';')
            if (c3 == ''):
                c3 = 0
            if (TreeIndex % 2 == 0):
                ExpenseTree.insert(parent='', index='end', iid=TreeIndex, text='', values=(c1, c2,"${:,.2f}".format(int(c3)), c4, c5, c6), tags=('evenrow',))
            else:
                ExpenseTree.insert(parent='', index='end', iid=TreeIndex, text='', values=(c1, c2,"${:,.2f}".format(int(c3)), c4, c5, c6), tags=('oddrow',))
            TreeIndex += 1
        except:
            if(guide == 'toDelete'):
                dname = 'Gastos/'+str(table.get())
                print(dname)
                database.delete(dname, guide)
                print('Default record deleted')
    fecha = datetime.now().strftime("%d/%m/%Y")

def LoadDataBase():
    loadYear = []
    loadMonth = []
    result = database.get('/', 'Gastos')
    for r in result:
        r1, r2 = str(r).split('_')
        loadMonth.append(r1)
        loadYear.append(r2)
    unique(loadMonth)
    unique(loadYear)

def ExportdataToExcel():
    save_path = tkinter.filedialog.asksaveasfile(defaultextension=".xlsx", filetypes=(("Excel", "*.xlsx"),))
    ids = ExpenseTree.get_children()
    export = []
    for _id in ids:
        row = ExpenseTree.item(_id)
        export.append(list(row['values']))
    export_df = pd.DataFrame(export)
    export_df.columns = ['Concepto', 'Total', 'Banco', 'Descripción', 'Fecha']
    #export_df.to_excel('writer.xlsx', sheet_name='GASTOS')
    writer = pd.ExcelWriter(save_path.name) 
    export_df.to_excel(writer, sheet_name='GASTOS', index=False, na_rep='NaN')
    writer.sheets['GASTOS'].set_column(0, 0, 35)
    writer.sheets['GASTOS'].set_column(1, 1, 35)
    writer.sheets['GASTOS'].set_column(2, 2, 35)
    writer.sheets['GASTOS'].set_column(3, 3, 45)
    writer.sheets['GASTOS'].set_column(4, 4, 35)
    writer.save()

def ExportResume():
    save_path = tkinter.filedialog.asksaveasfile(defaultextension=".xlsx", filetypes=(("Excel", "*.xlsx"),))
    workbook = xlsxwriter.Workbook(save_path.name)
    normal_money_format =  workbook.add_format({'bold': True, 'num_format': '$#,##0', 'border':1})
    normal_format = workbook.add_format({'bold': True, 'border':1})
    main_money_format = workbook.add_format({'num_format': '$#,##0', 'bg_color': 'yellow', 'bold': True, 'border':1})
    title_format = workbook.add_format({'bold': True, 'bg_color': 'yellow', 'border':1})
    resume_worksheet = workbook.add_worksheet(name="RESUMEN")
    resume_worksheet.set_column('A:A', 30)
    resume_worksheet.set_column('B:B', 30)
    ids = ExpenseTree.get_children()
    export = []
    for _id in ids:
        row = ExpenseTree.item(_id)
        export.append(list(row['values']))
    export_df = pd.DataFrame(export)
    export_df.columns = ['A', 'B', 'C', 'D', 'E', 'F']
    mainList = ['GASTOS DE VENTAS', 'GASTOS NO OPERACIONALES', 'GENERAL', 'OPERACIONALES DE ADMON', 'OTRO']
    concept_one =[ 'Viatícos Ana Mayra', 'Viatícos Andrés', 'Viatícos Edwin', 'Viatícos Melany', 'Viatícos Sneyder', 
    'Viatícos MA Victoria', 'Viatícos Jhonatan', 'Viatícos Silvia - Lissette (Tiquetes, viajes etc)', 'Viatícos René', 
    'Viatícos Andrés Arias', 'Viatícos Karen', 'Comisiones Ana Mayra', 'Comisiones Andrés', 'Comisiones Edwin','Comisiones Melany', 
    'Comisiones Sneyder', 'Comisiones MA Victoria', 'Comisiones Ronald (Trimestrales)', 'Comisiones Karina', 'Comisiones Karen', 
    'Comisiones Bonos fin de año','Comisiones Ajustes de comisión 2019', 'Comisiones Rene', 'Comisiones Silvia y Liss', 'Cesantias',
    'Industria y Comercio', 'Transportes, Fletes y Acarreos', 'Utililes papeleria y copias','Uniformes','Eventos, rifas, detalles, examenes, etc']
    concept_two = ['Gravamen bancos + comisiones + iva', 'Impuestos Asumidos','Constitución nueva sociedad',
    'Pago cuota de manejo cta cte', 'Movilizacion', 'Aseo', 'Pagos Sw y/o Página, MTTOS', 'Compra Equipos (Celulares, tablets etc)', 
    'Gastos de Representación Y asesorias', 'Otros','Retiro de caja para ingresar a banco', 'Anticipos (recaudos descontados a vendedores)', 
    'Marketing']
    concept_three = ['COMPRA DE MERCANCIA', 'CREDITOS BANCOS', 'DECLARACION DE RENTA', 'COMPRAS INTERNACIONALES' ]
    concept_four = ['Salarios', 'Básicos vendedores', 'Honorarios Lissette', 'Honararios Silvia', 'Primas' ,
    'Otros pagos (Horas Extras etc)', 'Planillas', 'Pagos Auxiliar Bodegaje', 'Honorarios Servicios Contables',
    'Arrendamientos', 'Servicios Generales', 'Acueducto y Alcantarillado', 'Energia Electrica','Vigilancia','Telefono e Internet', 
    'Celular','Envia', '4./72','Covinoc','Giitic','Seguros', 'Impuestos IVA', 'Impuestos Retefuente', 'Camara de comercio']
    concept_five = ['Otro']
    index = 0
    counter = 1
    labels = []
    info = []
    for item in mainList:
        filterData = export_df.loc[export_df['A'] == item]
        values = filterData['C'].replace( '[\$,]','', regex=True ).astype(float)
        resume_worksheet.write('A'+str(counter), item, title_format)
        resume_worksheet.write('B'+str(counter), values.sum(), main_money_format)
        labels.append(item)
        info.append(values.sum())
        counter = counter + 1
        if (index == 0):
            secondList = concept_one
        elif (index == 1):
            secondList = concept_two
        elif (index == 2):
            secondList = concept_three
        elif (index == 3):
            secondList = concept_four
        elif (index == 4):
            secondList = concept_five
        for _item in secondList:
                    try:
                        _filterData = filterData.loc[filterData['B'] == _item]
                        _values = _filterData['C'].replace( '[\$,]','', regex=True ).astype(float)
                        resume_worksheet.write('A'+str(counter), _item, normal_format)
                        resume_worksheet.write('B'+str(counter), _values.sum(), normal_money_format)
                        counter=counter+1
                        labels.append(_item)
                        info.append(_values.sum())
                    except:
                        resume_worksheet.write('A'+str(counter), _item, normal_format)
                        resume_worksheet.write('B'+str(counter), 0, normal_money_format)
                        counter=counter+1
                        labels.append(_item)
                        info.append(0)
        index = index + 1
    workbook.close()
    resume = pd.DataFrame(list(zip(labels, info)),
               columns =['Name', 'val'])

def DeleteFromDataBase():
    todelete = dict(fireBaseData)
    for key, data in todelete.items():
        if(data == mySelection):
            myKey = key
    ExpenseTree.delete(mytreeSelection)
    database.delete('Gastos/'+str(tableName), myKey)
# -----------------------GUI------------------------------------
multiprocessing.freeze_support()
root = Tk()
root.title("MotorLights - Seguimiento diario de vendedores")
try:
    database = firebase.FirebaseApplication('https://pruebas-6d911.firebaseio.com/', None)
    print(database)
except Exception as e:
    messagebox.showerror("Ocurrió un error conectandose a la base de datos", e)
sh = root.winfo_screenheight()
sw = root.winfo_screenwidth()
root.resizable(0,0)
root.geometry(str(int(sw*0.55338541))+"x"+str(int(sh*0.491898)))
menubar = Menu(root)
filemenu = Menu(menubar, tearoff=0)
salesmenu = Menu(menubar, tearoff=0)
salesmenu.add_command(label = "Control de gastos por mes", command = lambda : Expenses())
filemenu.add_command(label = "Generar seguimiento diario", command = lambda : Generate())
filemenu.add_cascade(label = "Ingresar rangos de incentivos", command = lambda : RangesGui())
filemenu.add_cascade(label = "Calcular incentivos",command = lambda : GenerateBonus())
filemenu.add_cascade(label = "Cargar último",command = lambda : Load())
filemenu.add_separator()
filemenu.add_cascade(label = "Salir", command = root.destroy)
menubar.add_cascade(label = "Archivo", menu = filemenu)
menubar.add_cascade(label = "Gastos", menu = salesmenu)
filemenu.entryconfig("Calcular incentivos")
filemenu.entryconfig("Ingresar rangos de incentivos")
root.config(menu = menubar)
style = ttk.Style()
style.theme_use("vista")
style.configure("Treeview", background = "#FFFFFF", foreground = "black", fieldbackgound = "#FFFFFF")
style.map('Treeview', background = [('selected', '#98F5FF')])
tree_frame = Frame(root)
tree_frame.place(x = int(sw*0.013), y = int(sh*0.023))
tree_scroll = Scrollbar(tree_frame)
tree_scroll.pack(side = RIGHT, fill = Y)
my_tree = ttk.Treeview(tree_frame, yscrollcommand = tree_scroll.set)
my_tree.pack()
my_tree.bind('<<TreeviewSelect>>', Callback)
my_tree.config(selectmode = 'browse')
tree_scroll.config(command = my_tree.yview)
my_tree['columns'] = ('Vendedor', 'Registros', 'Facturas', 'Meta', 'Venta', 'Meta_','Recaudo')
my_tree.column('#0', width = 0)
my_tree.column('Vendedor', anchor = W, width = int(sw*0.097))
my_tree.column('Registros', anchor = CENTER, width = int(sw*0.052))
my_tree.column('Facturas', anchor = CENTER, width = int(sw*0.052))
my_tree.column('Meta', anchor = CENTER, width = int(sw*0.078))
my_tree.column('Venta', anchor = CENTER, width = int(sw*0.078))
my_tree.column('Meta_', anchor = CENTER, width = int(sw*0.078))
my_tree.column('Recaudo', anchor = CENTER, width = int(sw*0.078))
#Headings
my_tree.heading('#0', text = 'ID')
my_tree.heading('Vendedor', text = 'Vendedor')
my_tree.heading('Registros', text = 'Registros')
my_tree.heading('Facturas', text = 'Facturas')
my_tree.heading('Meta', text = 'Meta de venta')
my_tree.heading('Venta', text = 'Venta')
my_tree.heading('Meta_', text = 'Meta de recaudo')
my_tree.heading('Recaudo', text = 'Recaudo*')
#Entradas
source = LabelFrame(root, text="Archivos de entrada")
source.place(x =int(sw*0.013), y = int(sh*0.3125))
sales = Button(source, text="Costo de ventas por vendedor", command = lambda: Sales(0))
sales.pack(pady=int(sh*0.005787), padx=int(sw*0.003255), fill = X)
recaudo = Button(source, text="Libro auxiliar", command = lambda: Collection(0), state = DISABLED)
recaudo.pack(pady=int(sh*0.005787), padx=int(sw*0.003255), fill = X)
cobranza = Button(source, text="Gestión de cobranza", command = lambda: Charge(0), state = DISABLED)
cobranza.pack(pady=int(sh*0.005787), padx=int(sw*0.003255), fill = X)
#Meta
target = LabelFrame(root, text="Agregar metas", width=int(sw*0.1757), height=int(sh*0.1157))
target.place(x = int(sw*0.1412), y = int(sh*0.3125))
meta = Entry(target, text="", width = int(sw*0.01302), justify=RIGHT)
meta.grid(row = 0, column =0, pady=int(sh*0.003472), padx=int(sw*0.001953))
confirm = Button(target, text="Meta de venta", command = lambda: Update(), width = int(sw*0.01171875), state = DISABLED)
confirm.grid(row=0, column = 1, pady=int(sh*0.003472), padx=int(sw*0.001953125))
confirm_ = Button(target, text="Meta de recaudo", command = lambda: Update_(), width = int(sw*0.01171875), state = DISABLED)
confirm_.grid(row = 1, column = 1, pady=int(sh*0.003472), padx=int(sw*0.001953125))
title = Label(target, text = "Vendedor", width = int(sw*0.00520833), font=("Arial", 8))
title.grid(row = 1, column = 0, pady=int(sh*0.003472), padx=int(sw*0.001953125))
selected_vendor = Label(target, text = "", width = int(sw*0.0162760), font=("Arial", 8), anchor = W)
selected_vendor.grid(row =3 , column = 0, columnspan = 2, pady=int(sh*0.004629), padx=int(sw*0.00130208))
dateLabel = Label(root, text = "")
dateLabel.place(x = int(sw*0.141276), y = int(sh*0.43981481))
#resumen
resume = LabelFrame(root, text="Resumen", width = int(sw*0.130208))
resume.place(x = int(sw*0.325520), y = int(sh*0.3125))
total_vendor = Label(resume, text="Total vendedores: ", width = int(sw*0.00846354), anchor = W)
total_vendor.grid(pady=int(sh*0.00231481),row = 0, column = 0)
total_resume = Label(resume, text ="", width = int(sw*0.003255208), anchor = E)
total_resume.grid(row = 0, column = 1)
total_sale = Label(resume, text="Total ventas: ", width = int(sw*0.00846354), anchor = W)
total_sale.grid(pady=int(sh*0.00231481), padx=int(sw*0.00130208), row = 1, column = 0)
sale_resume = Label(resume, text = "", width = int(sw*0.0078125), anchor = W)
sale_resume.grid(row = 1, column = 1)
total_fact = Label(resume, text="Total facturas: ", width = int(sw*0.00846354), anchor = W)
total_fact.grid(pady=int(sh*0.00231481), padx=int(sw*0.00130208), row = 2, column = 0)
fact_resume = Label(resume, text = "", width = int(sw*0.003255208), anchor = E)
fact_resume.grid(row = 2, column = 1)
total_reca = Label(resume, text="Total recaudo*: ", width = int(sw*0.00846354), anchor = W)
total_reca.grid(pady=int(sh*0.00231481), padx=int(sw*0.00130208), row = 3, column = 0)
reca_resume = Label(resume, text = "", width = int(sw*0.008463541), anchor = W)
reca_resume.grid(row=3, column=1)
#
ind = Label(root, text = "*Sin iva")
ind.place(x = int(sw*0.325520833), y = int(sh*0.4513888))
# logo display
canvas = Canvas(root, width = int(sw*0.1953125), height = int(sh*0.3472222))
canvas.place(x=int(sw*0.46875), y = int(sh*0.28935185))
img = Image.open("logo.png")  # PIL solution
img = img.resize((int(sw*0.045572916), int(sh*0.0810185)), Image.ANTIALIAS) #The (x, y) is (height, width)
img = ImageTk.PhotoImage(img) # convert to PhotoImage
canvas.create_image(20,20, anchor=NW, image=img)
#GEN
generate = Button(root, text = "Generar reporte", width = int(sw*0.009765625), command = lambda : Generate(), state = DISABLED)
generate.place(x = int(sw*0.46875), y = int(sh*0.4050925))
clear = Button(root, text = 'limpiar', width = int(sw*0.009765625), command = lambda  : Clear(), state = NORMAL)
clear.place(x=int(sw*0.46875), y = int(sh*0.439814))
root.mainloop()
#----------------------------------------------------------------------------------------