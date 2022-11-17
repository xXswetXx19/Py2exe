import tkinter
from tkinter import *
from tkinter import ttk, Toplevel, messagebox
from Clibs import Autocomplete as ac
from datetime import datetime, timedelta
import time
import sqlite3 as sql
import pandas as pd
import os
import xlsxwriter

class APP:
    db_name = 'Registro.db'
    def __init__(self, Ventana):
        self.fechaActual = datetime.now()
        self.Fecha = datetime.strftime(self.fechaActual, '%Y-%m-%d')
        self.Hora = datetime.strftime(self.fechaActual, "%H:%M:%S")
# Listas
        QProductos = self.run_query("SELECT Producto FROM Productos")
        self.Productos = [i[0] for i in list(QProductos)]

        QPresentaciones = self.run_query("SELECT Presentacion FROM Presentaciones")
        self.Presentaciones = [i[0] for i in list(QPresentaciones)]
        
        QFragancias = self.run_query("SELECT Fragancia FROM Fragancias")
        self.Fragancias = [i[0] for i in list(QFragancias)]

# INTERFAZ
        self.frame = Frame(Ventana,bg="#9BBB59")
        self.frame.place(relx=0.1, rely=0.1, relwidth=0.8, relheight=0.8)
        self.canvas = Canvas(self.frame, bg="#9BBB59", width=450, height=125, borderwidth=0, highlightthickness = 0)
        self.img = tkinter.PhotoImage(file="Archivos/imgs/Logo.png") 
        self.canvas.create_image(250, 65, image=self.img)

        self.EHora1=Label(self.frame,text="Hora", bg="#92D050", borderwidth=1, relief="solid", font="Bahnschrift 12")
        self.EHora2=Label(self.frame,text="00:00:00", bg="#00B050", borderwidth=1, relief="solid", font="Bahnschrift 12")
        self.EFecha1=Label(self.frame,text="Fecha", bg="#92D050", borderwidth=1, relief="solid", font="Bahnschrift 12")
        self.EFecha2=Label(self.frame,text=self.Fecha, bg="#00B050", borderwidth=1, relief="solid", font="Bahnschrift 12")
        self.EProducto=Label(self.frame,text="Producto", bg="#00B050", borderwidth=1, relief="solid", font="Bahnschrift 12")
        self.LProducto = ac.AutocompleteCombobox(self.frame, width=20,completevalues=self.Productos)
        self.EPresentacion=Label(self.frame,text="Presentacion", bg="#00B050", borderwidth=1, relief="solid", font="Bahnschrift 12")
        self.LPresentacion= ac.AutocompleteCombobox(self.frame, completevalues=self.Presentaciones, width=20)
        self.EUnidades=Label(self.frame,text="Unidades", bg="#00B050", borderwidth=1, relief="solid", font="Bahnschrift 12")
        self.CUnidades = tkinter.Entry(self.frame, borderwidth=1, relief="solid")
        self.EFragancia=Label(self.frame,text="Fragancia", bg="#00B050", borderwidth=1, relief="solid", font="Bahnschrift 12")
        self.LFragancia= ac.AutocompleteCombobox(self.frame, completevalues=self.Fragancias, width=20)
        self.EPrecio=Label(self.frame,text="PVP C/U", bg="#00B050", borderwidth=1, relief="solid", font="Bahnschrift 12")
        self.CPrecio = tkinter.Entry(self.frame, borderwidth=1, relief="solid")
        self.LObservaciones=Label(self.frame,text="Observaciones", bg="#00B050", borderwidth=1, relief="solid", font="Bahnschrift 11")
        self.CObservaciones = tkinter.Entry(self.frame, borderwidth=1, relief="solid")
        self.boton1 = tkinter.Button(self.frame, text = "Limpiar", command = self.clear_input, bg= "#FF0000", font=('Arial', '11', 'bold'), fg="White")
        self.boton2 = tkinter.Button(self.frame, text = "Registrar", command = self.add_product, bg= "#00B050", font=('Arial', '11', 'bold'), fg="White")        
        self.boton3 = tkinter.Button(self.frame, text = "Ventas", command = self.Ventana2, bg= "#0877CE", font=('Arial', '11', 'bold'), fg="White")        
        # Widgets place
        self.canvas.place(relx=0.3, rely=0, relwidth=0.4, relheight=0.180)
        self.EHora1.place(relx=0.180, rely=0.3, relwidth=0.1, relheight=0.03)
        self.EHora2.place(relx=0.279, rely=0.3, relwidth=0.12, relheight=0.03)
        self.EProducto.place(relx=0.180, rely=0.4, relwidth=0.1, relheight=0.03)
        self.LProducto.place(relx=0.279, rely=0.4, relwidth=0.12, relheight=0.03)
        self.EPresentacion.place(relx=0.180, rely=0.5, relwidth=0.1, relheight=0.03)
        self.LPresentacion.place(relx=0.279, rely=0.5, relwidth=0.12, relheight=0.03)
        self.EUnidades.place(relx=0.180, rely=0.6, relwidth=0.1, relheight=0.03)
        self.CUnidades.place(relx=0.279, rely=0.6, relwidth=0.12, relheight=0.03)
        self.EFecha1.place(relx=0.6, rely=0.3, relwidth=0.1, relheight=0.03)
        self.EFecha2.place(relx=0.699, rely=0.3, relwidth=0.12, relheight=0.03)
        self.EFragancia.place( relx=0.6, rely=0.4, relwidth=0.1, relheight=0.03)
        self.LFragancia.place(relx=0.699, rely=0.4, relwidth=0.12, relheight=0.03)
        self.EPrecio.place( relx=0.6, rely=0.5, relwidth=0.1, relheight=0.03)
        self.CPrecio.place(relx=0.699, rely=0.5, relwidth=0.12, relheight=0.03)
        self.LObservaciones.place(relx=0.6, rely=0.6, relwidth=0.1, relheight=0.03)
        self.CObservaciones.place(relx=0.699, rely=0.6, relwidth=0.12, relheight=0.03)
        self.boton1.place(relx=0.3, rely=0.7, width=120, height=30)
        self.boton2.place(relx=0.6, rely=0.7, width=120, height=30)
        self.boton3.place(relx=0.450, rely=0.7, width=120, height=30)

# Functions
        def createDB():
            conn = sql.connect("Registro.db")
            conn.commit()
            conn.close()
            
        def createTable():
            conn = sql.connect("Registro.db")
            cursor = conn.cursor()
            cursor.execute(
              """CREATE TABLE IF NOT EXISTS REGISTRO (
                  Id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
                  Fecha text,
                  Hora text,
                  Producto text,
                  Presentacion text,
                  Cantidad integer,
                  Fragancia text,
                  Observaciones text,
                  PVPU float,
                  PVPT float
            )"""
            )
            conn.commit()
            conn.close()
        def update_clock():
            hours = time.strftime("%I")
            minutes = time.strftime("%M")
            seconds = time.strftime("%S")
            am_or_pm = time.strftime("%p")
            time_text = hours + ":" + minutes + ":" + seconds + " " + am_or_pm
            self.EHora2.config(text=time_text)
            self.EHora2.after(1000, update_clock)
# Startup functions
        createDB()
        createTable() 
        update_clock()
# Utilities functions
    def clear_input(self):
         self.LProducto.set('')
         self.LPresentacion.set('')
         self.CUnidades.delete(0, END)
         self.LFragancia.set('')
         self.CObservaciones.delete(0, END)
         self.CPrecio.delete(0, END)
    def run_query(self, query, parametros = ()):
        with sql.connect(self.db_name) as conn:
            cursor = conn.cursor()
            result = cursor.execute(query, parametros)
            conn.commit()
        return result
    def Validacion(self):
        if len(self.LProducto.get()) == 0 and len(self.CUnidades.get()) == 0 and len(self.CPrecio.get()) == 0:
            messagebox.showerror(title= 'Registro Sumliprob', message = "Debes Escribir el Producto las unidades y el Precio.")
            return False
        if not self.CUnidades.get().isnumeric() and self.CPrecio.get().replace(",", ".").split(".")[0].isdigit():
            messagebox.showerror(title= 'Registro Sumliprob', message = "Verifica que las unidades y precio sean numeros")
            return False
        if not self.LProducto.get() in self.Productos:
            messagebox.showerror(title= 'Registro Sumliprob', message = "Producto no registrado")
            return False
        if not self.LPresentacion.get() == "" and not self.LPresentacion.get() in self.Presentaciones:
            messagebox.showerror(title= 'Registro Sumliprob', message = "Presentacion no valida")
            return False
        if not self.LFragancia.get() == "" and not self.LFragancia.get() in self.Fragancias:
            messagebox.showerror(title= 'Registro Sumliprob', message = "Fragancia no Registrada")
            return False
        return True
    def Ventana2(self):
        if not any(isinstance(x, Toplevel) for x in Ventana.winfo_children()):
            self.wind = tkinter.Toplevel(Ventana)
            app = Registro(self.wind)
            self.wind.title('Ventana de Registro')
            self.wind.resizable(width=False, height=False)
            self.wind.geometry("900x322")
            self.wind.iconbitmap("Archivos/imgs/Icono.ico")
        else:
            messagebox.showinfo(title= 'Registro Sumliprob', message = "Ya tienes una ventana de registro abierta")
    def add_product(self):
        if not self.Validacion():
            return
        query = 'INSERT INTO Registro VALUES(NULL, ?, ?, ?, ?, ?, ?, ?, ?, ?)'
        self.PVPMT = (int(self.CUnidades.get()) * float(self.CPrecio.get().replace(",",".")))
        Observaciones = "N/A" if self.CObservaciones.get() == "" else self.CObservaciones.get()
        Fragancia = "N/A" if self.LFragancia.get() == "" else self.LFragancia.get()
        Presentacion = "N/E" if self.LPresentacion.get() == "" else self.LPresentacion.get()
        parametros = (self.Fecha,self.LProducto.get(), Presentacion,self.CUnidades.get(), Fragancia, Observaciones.capitalize(), self.CPrecio.get().replace(",","."), self.PVPMT, self.Hora)
        self.run_query(query, parametros)
        self.clear_input()
        try:
            Regclass = Registro(self.wind)
            Regclass.get_products()
            Regclass.suma()
        except:
            pass    

class Registro:
    def __init__(self, ventana):
        self.db_name = 'Registro.db'
        self.table_name = 'Registro'
        self.fechaActual = datetime.now()
        self.Fechahoy = datetime.strftime(self.fechaActual, "%Y-%m-%d")
        self.ventanaR = ventana

        self.frame = LabelFrame(self.ventanaR, text = '', borderwidth = 0)
        self.frame.grid(row = 0, column = 0, columnspan = 3, pady = 30)
        self.message = Label(self.frame,text = '', fg = 'black')
        self.message.grid(row = 3, column = 0, columnspan=2, sticky= W + E) 
        self.ventashoymsg = Label(self.ventanaR,text = '' ,borderwidth=1, font = ('Arial', 9, 'bold'))
        self.ventashoymsg.place(width=140, x=380, y=15) 
      
        self.ventasayermsg = Label(self.ventanaR,text = '' ,borderwidth=1, font = ('Arial', 9, 'bold'))
        self.ventasayermsg.place(width=140, x=180, y=15) 
       
        self.ventassemanamsg = Label(self.ventanaR,text = '' ,borderwidth=1, font = ('Arial', 9, 'bold'))
        self.ventassemanamsg.place(width=180, x=560, y=15)

        self.tree = ttk.Treeview(self.frame, height = 11 ,columns = ('#0', '#1','#2','#3', '#4', '#5', '#6', '#7', '#8'))
        self.tree.grid(row = 4, column = 0, columnspan = 2)
        self.tree.heading('#0', text = 'ID', anchor = CENTER)
        self.tree.column("#0",minwidth=0,width=0, stretch=NO, anchor= CENTER)
        self.tree.heading('#1', text = 'Fecha', anchor = CENTER)
        self.tree.column("#1",minwidth=0,width=100, stretch=NO, anchor= CENTER)
        self.tree.column("#2",minwidth=0,width=100, stretch=NO, anchor= CENTER)
        self.tree.heading('#2', text = 'Hora', anchor = CENTER)
        self.tree.heading('#3', text = 'Producto', anchor = CENTER)
        self.tree.column("#3",minwidth=0,width=100, stretch=NO, anchor= CENTER)
        self.tree.heading('#4', text = 'Presentacion', anchor = CENTER)
        self.tree.column("#4",minwidth=0,width=100, stretch=NO, anchor= CENTER)
        self.tree.heading('#5', text = 'Cant', anchor = CENTER)
        self.tree.column("#5",minwidth=0,width=100, stretch=NO, anchor= CENTER)
        self.tree.heading('#6', text = 'Fragancia', anchor = CENTER)
        self.tree.column("#6",minwidth=0,width=100, stretch=NO, anchor= CENTER)
        self.tree.heading('#7', text = 'Observaciones', anchor = CENTER)
        self.tree.column("#7",minwidth=0,width=100, stretch=NO, anchor= CENTER)
        self.tree.heading('#8', text = 'PVP C/U', anchor = CENTER)
        self.tree.column("#8",minwidth=0,width=100, stretch=NO, anchor= CENTER)
        self.tree.heading('#9', text = 'Total', anchor = CENTER)
        self.tree.column("#9",minwidth=0,width=100, stretch=NO, anchor= CENTER)
        ttk.Button(self.frame, text = 'BORRAR', command = self.delete_product).grid(row = 5, column = 0, sticky = W+E)
        ttk.Button(self.frame, text = 'EDITAR', command = self.edit_product).grid(row = 5, column = 1, sticky = W+E)
        ttk.Button(self.ventanaR, text = 'EXPORTAR', command = self.exportar).place(x=812, y=15)
        #llenando filas
        self.get_products()

    def run_query(self, query, parametros = ()):
        with sql.connect(self.db_name) as conn:
            cursor = conn.cursor()
            result = cursor.execute(query, parametros)
            conn.commit()
        return result
    def get_products(self):
        #Limpiando tabla
        records = self.tree.get_children()
        for element in records:
            self.tree.delete(element)
        #Consultando datos
        query = 'SELECT * FROM Registro WHERE Fecha >= ? ORDER BY Fecha ASC'
        Fechalunes = (datetime.today() - timedelta(days=datetime.today().weekday())).strftime("%Y-%m-%d")
        parametros = ([Fechalunes])
        db_rows = self.run_query(query, parametros)
        #Llenando datos
        for row in db_rows:
            self.tree.insert('', 0, text = row[0], values = [row[1], row[9], row[2], row[3], row[4], row[5], row[6], row[7], row[8]], tags=('fg', 'bg'))
        self.suma()
    def suma(self):
        query = 'SELECT SUM(PVPT) FROM Registro WHERE Fecha >= ?'

        result = self.run_query(query, ([self.Fechahoy]))
        total = list(result)[0][0] or 0
        self.ventashoymsg['text'] = f"Ventas de hoy {total}$"

        FechaAyer = datetime.strftime(self.fechaActual - timedelta(days=1), "%Y-%m-%d")
        result = self.run_query('SELECT SUM(PVPT) FROM Registro WHERE Fecha = ?', ([FechaAyer]))
        total = list(result)[0][0] or 0
        self.ventasayermsg['text'] = f"Ventas de ayer {total}$"

        Fechalunes = (datetime.today() - timedelta(days=datetime.today().weekday())).strftime("%Y-%m-%d")
        result = self.run_query(query, ([Fechalunes]))
        total = list(result)[0][0] or 0 
        self.ventassemanamsg['text'] = f"Ventas de la semana {total}$"

    def delete_product(self):
        self.message['text'] = ''
        try:
           self.tree.item(self.tree.selection())['text']
        except IndexError as e:
            self.message['text'] = 'Por favor selecciona la venta que deseas borrar.'
            return
        self.message['text'] = ''
        ID = self.tree.item(self.tree.selection())['text']
        query = 'DELETE FROM Registro WHERE ID = ?'
        self.run_query(query, (ID, ))
        self.message['text'] = 'Venta borrada correctamente.'
        self.ventashoymsg['text'] = ''
        self.get_products()
    def edit_product(self):
        self.message['text'] = ''
        try:
            ID = self.tree.item(self.tree.selection())['text']
            Fecha = self.tree.item(self.tree.selection())['values'][0]
            Hora = self.tree.item(self.tree.selection())['values'][1]
            Producto = self.tree.item(self.tree.selection())['values'][2]
            Presentacion = self.tree.item(self.tree.selection())['values'][3]
            Cantidad = self.tree.item(self.tree.selection())['values'][4]
            Fragancia = self.tree.item(self.tree.selection())['values'][5]
            Observaciones = self.tree.item(self.tree.selection())['values'][6]
            PVPU = self.tree.item(self.tree.selection())['values'][7]
            PVPT = self.tree.item(self.tree.selection())['values'][8]
        except IndexError as e:
            self.message['text'] = 'Por favor selecciona la venta que deseas editar.'
            return
        self.edit_wind = Toplevel()
        self.edit_wind.resizable(width=False, height=False)
        self.edit_wind.title = 'Editar Producto'
        #Fecha
        Label(self.edit_wind, text = 'Fecha: ').grid(row = 1, column = 1) # row = 1, column = 1
        CFecha = Entry(self.edit_wind, textvariable = StringVar(self.edit_wind, value = Fecha))
        CFecha.grid(row = 1, column = 2)
        #Hora
        Label(self.edit_wind, text = 'Hora: ').grid(row = 2, column = 1) # row = 1, column = 1
        CHora = Entry(self.edit_wind, textvariable = StringVar(self.edit_wind, value = Hora))
        CHora.grid(row = 2, column = 2)
        #Nombre Nuevo
        Label(self.edit_wind, text = 'Producto: ').grid(row = 3, column = 1) # row = 1, column = 1
        CProducto = Entry(self.edit_wind, textvariable = StringVar(self.edit_wind, value = Producto))
        CProducto.grid(row = 3, column = 2)
        #Presentacion Nuevo
        Label(self.edit_wind, text = 'Presentacion: ').grid(row = 4, column = 1)
        CPresentacion = Entry(self.edit_wind, textvariable = StringVar(self.edit_wind, value = Presentacion))
        CPresentacion.grid(row = 4, column = 2)
        #Cantitades Nuevo
        Label(self.edit_wind, text = 'Cantidad: ').grid(row = 5, column = 1)
        CCantidad = Entry(self.edit_wind, textvariable = StringVar(self.edit_wind, value = Cantidad))
        CCantidad.grid(row = 5, column = 2)
        #Fragancia Nuevo
        Label(self.edit_wind, text = 'Fragancia: ').grid(row = 6, column = 1)
        CFragancia = Entry(self.edit_wind, textvariable = StringVar(self.edit_wind, value = Fragancia))
        CFragancia.grid(row = 6, column = 2)
        #Observaciones Nuevo
        Label(self.edit_wind, text = 'Observaciones: ').grid(row = 7, column = 1)
        CObservaciones = Entry(self.edit_wind, textvariable = StringVar(self.edit_wind, value = Observaciones))
        CObservaciones.grid(row = 7, column = 2)
        #PVPU Nuevo
        Label(self.edit_wind, text = 'PVP C/U: ').grid(row = 8, column = 1)
        CPVPU = Entry(self.edit_wind, textvariable = StringVar(self.edit_wind, value = PVPU))
        CPVPU.grid(row = 8, column = 2)
        #PVPT Nuevo
        Label(self.edit_wind, text = 'PVP Total: ').grid(row = 9, column = 1)
        CPVPT = Entry(self.edit_wind, textvariable = StringVar(self.edit_wind, value = PVPT), state = 'readonly')
        CPVPT.grid(row = 9, column = 2)

        Button(self.edit_wind, text = 'Editar', command = lambda: self.edit_records(CFecha.get(), CHora.get(), CProducto.get(), CPresentacion.get(), CCantidad.get(), CFragancia.get(), CObservaciones.get(), CPVPU.get(), CPVPT.get(), ID)).grid(row = 9, column = 2, sticky = W)
    def edit_records(self, CFecha,CHora, CProducto, CPresentacion, CCantidad, CFragancia, CObservaciones, CPVPU, CPVPT, ID):
        PVPMT= int(CCantidad) * float(CPVPU)
        query = 'UPDATE Registro SET Fecha = ?, Hora = ?, Producto = ?, Presentacion = ?, Cantidad = ?, Fragancia = ?, Observaciones = ?, PVPU = ?, PVPT = ? WHERE ID = ?'
        parameters = (CFecha, CHora, CProducto, CPresentacion, CCantidad, CFragancia, CObservaciones, CPVPU, PVPMT, ID)
        self.run_query(query, parameters)
        self.edit_wind.destroy()
        self.message['text'] = 'El Registro ha sido Actualizado Correctamente'
        self.get_products()

    def exportar(self):
        dt = datetime.now()
        FechaActual = datetime.strftime(dt, '%d/%m/%Y')
        HoraActual = datetime.strftime(dt, '%H:%M')
        am_or_pm = time.strftime("%p")
        desktopath = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')    
        conn = sql.connect(self.db_name) 
        cursor = conn.cursor()
        query = cursor.execute("select * from {} ORDER BY Fecha DESC, Hora DESC".format(self.table_name))
        conn.commit()
        Nlist = []
        for row in query:
            row = list(row)
            Fechaobj = datetime.strptime(row[1], "%Y-%m-%d")
            Horaobj = datetime.strptime(row[9], "%H:%M:%S")
            Fecha = Fechaobj.toordinal() - datetime(1900, 1, 1).toordinal() + 2
            Hora = (Horaobj - datetime(1900, 1, 1)).total_seconds() / 86400  
            row = Fecha, Hora, row[2], row[3], row[4], row[5], row[6], row[7], row[8] 
            Nlist.append(tuple(row))
        
        columnas = ['Fecha', "Hora", 'Producto', 'Presentacion', 'Cantidad', 'Fragancia','Observaciones', 'PVPU', 'PVPT']
        df = pd.DataFrame(Nlist, columns = columnas)
        writer = pd.ExcelWriter(desktopath+'/Registro de ventas.xlsx', engine='xlsxwriter')
        workbook  = writer.book   
        df.to_excel(writer, sheet_name='Registro de ventas', startcol = 0, startrow = 13, index = False, na_rep='N/A')
        worksheet = writer.sheets['Registro de ventas']
        worksheet2 = workbook.add_worksheet('Estadisticas')
        Generado = f'GENERADO EL {FechaActual} A las {HoraActual} {am_or_pm}'
        merge_format1 = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font': 'Arial Black',
        'font_color': '#FFFFFF',
        'font_size': '22',
        'fg_color': '#76933C'})
        merge_format2 = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_color': '#000000',
        'font_size': '11',
        'fg_color': '#EEECE1'})
        MoneyFormat = workbook.add_format({'num_format': '$ #,##0.00'})
        TimeFormat = workbook.add_format({'num_format': 'hh:mm:ss AM/PM'})
        DateFormat = workbook.add_format({'num_format': 'yyyy-mm-dd'})
        worksheet.merge_range('A1:I12', '', merge_format1)
        worksheet.merge_range('A13:I13', Generado, merge_format2)
        worksheet.set_column('A:A', 17, DateFormat)
        worksheet.set_column('B:B', 12, TimeFormat)
        worksheet.set_column('C:C', 30)
        worksheet.set_column('D:D', 29)
        worksheet.set_column('E:E', 20)
        worksheet.set_column('F:F', 11)
        worksheet.set_column('G:G', 55)
        worksheet.set_column('H:H', 10, MoneyFormat)
        worksheet.set_column('I:I', 10, MoneyFormat)
        worksheet.autofilter("A14:I14")
        worksheet.insert_image(0,3, 'Archivos/imgs/Registro.png')
        ValueFormat = workbook.add_format({ 
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_color': '#000000',
        'font_size': '11',
        'fg_color': '#E4DFEC'}) 
        worksheet2.write('A1', 'Ventas AM', ValueFormat)
        worksheet2.set_column('A:A', 10)
        worksheet2.write('A2', 'Ventas PM', ValueFormat)

        maxcolumnhora = len(df["Hora"]) + 14

        worksheet2.write_formula('B1', f'''=COUNTIF('Registro de ventas'!B15:B{maxcolumnhora}, "<0,5") - COUNTIF('Registro de ventas'!B15:B4124,0)''', ValueFormat)
        worksheet2.write_formula('B2', f'''=COUNTIF('Registro de ventas'!B15:B{maxcolumnhora}, ">0,5")''', ValueFormat)

        chart1 = workbook.add_chart({'type': 'pie'})
        chart1.add_series({
        'name': 'Ventas AM/PM',
        'categories': '=Estadisticas!$A$1:$A$2',
        'values':     '=Estadisticas!$B$1:$B$2',
        }) 
        chart1.set_style(37)
        chart1.set_size({'width': 384, 'height': 360})
        worksheet2.insert_chart(5, 2, chart1) 
        worksheet2.hide_gridlines(2) 
        writer.save()
        os.startfile(desktopath+'/Registro de ventas.xlsx')


if __name__ == '__main__':
    Ventana = Tk()
    Ventana.title('Registro de ventas de Sumliprob')
    Ventana.config(bg = "#9BBB59")
    Ventana.iconbitmap("Archivos/imgs/Icono.ico")
    Ventana.wm_attributes("-transparentcolor", 'grey')
    Ventana.state('zoomed')
    aplication = APP(Ventana)
    aplication.LProducto.focus()
    Ventana.mainloop()
    
    