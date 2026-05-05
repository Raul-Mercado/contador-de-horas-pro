import os
import json
import tempfile
from datetime import datetime, timedelta
import sqlite3
import platform
import subprocess

import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import messagebox, filedialog, simpledialog, ttk
from tkcalendar import DateEntry

import openpyxl
from openpyxl.styles import Font, Alignment

try:
    import win32api
    import win32print
except ImportError:
    win32api = None
    win32print = None

CONFIG_FILE = "config.json"
DB_FILE = "horas.db"


def cargar_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    else:
        conf = {
            "usuarios": {},
            "usuarios_orden": [],
            "ultimo_usuario": None
        }
        with open(CONFIG_FILE, "w") as f:
            json.dump(conf, f)
        return conf


def guardar_config(conf):
    with open(CONFIG_FILE, "w") as f:
        json.dump(conf, f, indent=4)


def crear_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS registros (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
 usuario TEXT NOT NULL,
            fecha TEXT NOT NULL,
            hora_inicio TEXT NOT NULL,
            hora_fin TEXT NOT NULL,
            duracion REAL NOT NULL,
            valor_hora REAL DEFAULT 0,
            total REAL DEFAULT 0,
            notas TEXT
        )
    ''')
    conn.commit()
    conn.close()


class VentanaUsuarios(tb.Toplevel):
    def __init__(self, root, config, *args, **kwargs):
        super().__init__(root, *args, **kwargs)
        self.title("Seleccione Usuario o Cree uno Nuevo")
        self.geometry("400x220")
        self.config_obj = config
        self.usuario_seleccionado = None

        tb.Label(self, text="Usuarios existentes:", font=("Arial", 12, "bold")).pack(pady=8)

        self.combo_usuarios = tb.Combobox(self, values=self.config_obj.get("usuarios_orden", []), state='readonly')
        if self.config_obj.get("ultimo_usuario"):
            self.combo_usuarios.set(self.config_obj["ultimo_usuario"])
        else:
            if self.config_obj.get("usuarios_orden"):
                self.combo_usuarios.current(0)
        self.combo_usuarios.pack(pady=8, fill='x', padx=20)

        frame_botones = tb.Frame(self)
        frame_botones.pack(pady=10)

        btn_sel = tb.Button(frame_botones, text="Seleccionar", bootstyle=SUCCESS, width=12, command=self.seleccionar)
        btn_sel.grid(row=0, column=0, padx=10)

        btn_nuevo = tb.Button(frame_botones, text="Nuevo Usuario", bootstyle=INFO, width=12, command=self.nuevo_usuario)
        btn_nuevo.grid(row=0, column=1, padx=10)

        self.transient(root)
        self.grab_set()
        self.focus_set()

    def seleccionar(self):
        elegido = self.combo_usuarios.get()
        if elegido:
            self.usuario_seleccionado = elegido
            self.destroy()
        else:
            messagebox.showerror("Error", "Seleccione un usuario.")

    def nuevo_usuario(self):
        nombre = simpledialog.askstring("Nuevo Usuario", "Ingrese nombre de usuario:", parent=self)
        if nombre:
            nombre = nombre.strip()
            if nombre in self.config_obj.get("usuarios", {}):
                messagebox.showerror("Error", "El usuario ya existe.")
                return
            self.config_obj["usuarios"][nombre] = {"valor_hora_base": 0.0}
            self.config_obj["usuarios_orden"].append(nombre)
            guardar_config(self.config_obj)
            self.combo_usuarios['values'] = self.config_obj["usuarios_orden"]
            self.combo_usuarios.set(nombre)
        else:
            messagebox.showerror("Error", "Nombre inválido.")


class PreguntarValorHora(tb.Toplevel):
    def __init__(self, root, valor_guardado, *args, **kwargs):
        super().__init__(root, *args, **kwargs)
        self.title("Valor Hora")
        self.geometry("400x150")
        self.valor_base = valor_guardado
        self.valor_resultado = None

        tb.Label(self, text=f"Valor hora actual: ${self.valor_base:.2f}", font=('Arial', 12)).pack(pady=8)
        tb.Label(self, text="¿Desea conservar este valor o modificarlo?", font=('Arial', 11)).pack()

        frame_botones = tb.Frame(self)
        frame_botones.pack(pady=15)

        btn1 = tb.Button(frame_botones, text="Conservar", bootstyle=PRIMARY, width=12, command=self.conservar)
        btn1.grid(row=0, column=0, padx=10)

        btn2 = tb.Button(frame_botones, text="Modificar", bootstyle=WARNING, width=12, command=self.modificar)
        btn2.grid(row=0, column=1, padx=10)

        self.transient(root)
        self.grab_set()
 self.focus_set()

    def conservar(self):
        self.valor_resultado = self.valor_base
        self.destroy()

    def modificar(self):
        valor = simpledialog.askfloat("Modificar Valor Hora", "Ingrese nuevo valor por hora:", minvalue=0.0, parent=self)
        if valor is not None:
            self.valor_resultado = valor
            self.destroy()


class ControlHorasApp(tb.Window):
    def __init__(self):
        super().__init__(themename="flatly")
        self.title("Control Horas Pro")
        self.minsize(1020, 720)
        self.geometry('1020x720')
        self.eval('tk::PlaceWindow . center')  # Centrar ventana
        self.resizable(True, True)

        self.configuracion = cargar_config()
        crear_db()

        self.usuario_activo = None  # Se inicializa sin usuario
        self.valor_hora_base = 0.0
        self.corriendo = False
        self.pausado = False
        self.inicio_tiempo = None
        self.tiempo_acumulado = 0

        self._crear_menu()
        self.crear_widgets()

        self.cambiar_usuario()  # En lugar de proceso en init para poder repetir logout/login

        self.actualizar_reloj()

    def _crear_menu(self):
        menu = tb.Menu(self)
        self.config(menu=menu)
        menu_ayuda = tb.Menu(menu, tearoff=0)
        menu_usuarios = tb.Menu(menu, tearoff=0)

        menu.add_cascade(label="Archivo", menu=menu_usuarios)
        menu_usuarios.add_command(label="Cerrar Sesión", command=self.cerrar_sesion)
        menu_usuarios.add_separator()
        menu_usuarios.add_command(label="Salir", command=self.destroy)

        menu.add_cascade(label="Ayuda", menu=menu_ayuda)
        menu_ayuda.add_command(label="Tutorial", command=self.mostrar_tutorial)
        menu_ayuda.add_separator()
        menu_ayuda.add_command(label="Acerca de", command=self.mostrar_acerca_de)

    def mostrar_tutorial(self):
        texto = (
            "Tutorial de Control Horas Pro:\n\n"
            "1. Seleccione o cree un usuario al iniciar o al cerrar sesión.\n"
            "2. El valor hora se puede conservar o modificar.\n"
            "3. Use Inicio, Pausa, Fin para controlar la jornada.\n"
            "4. Puede ingresar horas manualmente con la ayuda de los selectores horarios.\n"
            "5. Visualice últimos registros y pueda editarlos o borrarlos.\n"
            "6. Genere informes por fechas, exporte a Excel o TXT, o imprima.\n"
            "\nSi tiene dudas consulte el manual o contacte al soporte."
        )
        messagebox.showinfo("Tutorial", texto)

    def mostrar_acerca_de(self):
        texto = (
            "Control Horas Pro\n"
            "Versión 1.1\n\n"
            "Desarrollador: Raúl Mercado\n"
            "Teléfono: +5492644980791\n"
            "Email: raul.roberto.mercado@gmail.com\n"
        )
        messagebox.showinfo("Acerca de", texto)

    def crear_widgets(self):
        frame_reloj = tb.Labelframe(self, text="Reloj y Cronómetro", bootstyle="primary")
        frame_reloj.pack(padx=10, pady=5, fill='x')

        self.label_usuario = tb.Label(frame_reloj, text="Usuario: Sin sesión", font=('Arial', 14, 'bold'))
        self.label_usuario.pack(pady=3)

        self.label_reloj = tb.Label(frame_reloj, text="", font=('Arial', 24))
        self.label_reloj.pack(pady=5)

        self.label_cronometro = tb.Label(frame_reloj, text="00:00:00", font=('Arial', 20), foreground='blue')
        self.label_cronometro.pack(pady=5)

        frame_botones = tb.Frame(frame_reloj)
        frame_botones.pack(pady=5)

        self.btn_inicio = tb.Button(frame_botones, text="Inicio", bootstyle=SUCCESS, width=10, command=self.iniciar_jornada, state='disabled')
        self.btn_inicio.grid(row=0, column=0, padx=5)

        self.btn_pausa = tb.Button(frame_botones, text="Pausa", bootstyle=WARNING, width=10,
                                  command=self.pausar_jornada, state='disabled')
        self.btn_pausa.grid(row=0, column=1, padx=5)

        self.btn_fin = tb.Button(frame_botones, text="Fin", bootstyle=DANGER, width=10, command=self.finalizar_jornada,
                                state='disabled')
        self.btn_fin.grid(row=0, column=2, padx=5)

        frame_manual = tb.Labelframe(self, text="Ingreso / Edición Manual", bootstyle="info")
        frame_manual.pack(padx=10, pady=5, fill='x')

        tb.Label(frame_manual, text="Fecha (dd/mm/yyyy):").grid(row=0, column=0, sticky='e', padx=5, pady=3)
        self.fecha_manual = DateEntry(frame_manual, date_pattern='dd/mm/yyyy')
        self.fecha_manual.grid(row=0, column=1, padx=5, pady=3)

        tb.Label(frame_manual, text="Hora inicio (HH:MM):").grid(row=0, column=2, sticky='e', padx=5, pady=3)
        self.hora_inicio_h = tb.Combobox(frame_manual, width=3, values=[f"{h:02d}" for h in range(24)], state="readonly")
        self.hora_inicio_h.grid(row=0, column=3, sticky='w', padx=(5,0), pady=3)
        self.hora_inicio_h.set("00")
        self.hora_inicio_m = tb.Combobox(frame_manual, width=3, values=[f"{m:02d}" for m in range(0,60,5)], state="readonly")
        self.hora_inicio_m.grid(row=0, column=3, sticky='e', padx=(0,5), pady=3)
        self.hora_inicio_m.set("00")

        tb.Label(frame_manual, text="Hora fin (HH:MM):").grid(row=0, column=4, sticky='e', padx=5, pady=3)
        self.hora_fin_h = tb.Combobox(frame_manual, width=3, values=[f"{h:02d}" for h in range(24)], state="readonly")
        self.hora_fin_h.grid(row=0, column=5, sticky='w', padx=(5,0), pady=3)
        self.hora_fin_h.set("00")
        self.hora_fin_m = tb.Combobox(frame_manual, width=3, values=[f"{m:02d}" for m in range(0,60,5)], state="readonly")
        self.hora_fin_m.grid(row=0, column=5, sticky='e', padx=(0,5), pady=3)
        self.hora_fin_m.set("00")

        tb.Label(frame_manual, text="Valor hora ($):").grid(row=1, column=0, sticky='e', padx=5, pady=3)
        self.valor_hora_manual = tb.Entry(frame_manual, width=12)
        self.valor_hora_manual.grid(row=1, column=1, padx=5, pady=3)

        tb.Label(frame_manual, text="Notas:").grid(row=1, column=2, sticky='e', padx=5, pady=3)
        self.notas_manual = tb.Entry(frame_manual, width=30)
        self.notas_manual.grid(row=1, column=3, columnspan=3, padx=5, pady=3)

        self.btn_guardar = tb.Button(frame_manual, text="Guardar Registro", bootstyle=PRIMARY,
                                   command=self.guardar_registro_manual)
        self.btn_guardar.grid(row=2, column=3, pady=10, sticky='e')

        self.btn_actualizar = tb.Button(frame_manual, text="Actualizar Registro", bootstyle=WARNING,
                                      command=self.actualizar_registro, state='disabled')
        self.btn_actualizar.grid(row=2, column=4, pady=10, sticky='w')

        self.btn_borrar = tb.Button(frame_manual, text="Borrar Registro", bootstyle=DANGER,
                                  command=self.borrar_registro, state='disabled')
        self.btn_borrar.grid(row=2, column=5, pady=10, sticky='w')

        frame_tabla = tb.Labelframe(self, text="Últimos 3 Registros de Horas", bootstyle="secondary")
        frame_tabla.pack(padx=10, pady=5, fill='both', expand=True)

        columns = ("ID", "Fecha", "Inicio", "Fin", "Duración", "Valor $", "Total $", "Notas")
        self.tree = tb.Treeview(frame_tabla, columns=columns, show='headings')

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor='center')

        self.tree.pack(side='left', fill='both', expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.seleccionar_registro)

        scrollbar = tb.Scrollbar(frame_tabla, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=scrollbar.set)

        frame_informe = tb.Labelframe(self, text="Informes y Exportación", bootstyle="success")
        frame_informe.pack(padx=10, pady=5, fill='x')

        tb.Label(frame_informe, text="Desde:").grid(row=0, column=0, padx=5, pady=3)
        self.informe_fecha_desde = DateEntry(frame_informe, date_pattern='dd/mm/yyyy')
        self.informe_fecha_desde.grid(row=0, column=1, padx=5, pady=3)

        tb.Label(frame_informe, text="Hasta:").grid(row=0, column=2, padx=5, pady=3)
        self.informe_fecha_hasta = DateEntry(frame_informe, date_pattern='dd/mm/yyyy')
        self.informe_fecha_hasta.grid(row=0, column=3, padx=5, pady=3)

        self.var_mostrar_valor = tb.IntVar(value=1)
        self.var_mostrar_total = tb.IntVar(value=1)
        cb_valor = tb.Checkbutton(frame_informe, text="Mostrar Valor Hora", variable=self.var_mostrar_valor)
        cb_valor.grid(row=0, column=4, padx=5, pady=3)
        cb_total = tb.Checkbutton(frame_informe, text="Mostrar Total $", variable=self.var_mostrar_total)
        cb_total.grid(row=0, column=5, padx=5, pady=3)

        self.btn_generar_informe = tb.Button(frame_informe, text="Generar Informe", bootstyle=PRIMARY,
                                          command.ventana_informe_generar)
        self.btn_generar_informe.grid(row=0, column=6, padx=10)

    def actualizar_reloj(self):
        ahora = datetime.now().strftime("%d/%m/%Y  %H:%M:%S")
        self.label_reloj.config(text=ahora)
        self.label_usuario.config(text=f"Usuario: {self.usuario_activo}")

        if self.corriendo and not self.pausado:
            tiempo_actual = self.tiempo_acumulado + (datetime.now() - self.inicio_tiempo).total_seconds()
            horas = int(tiempo_actual // 3600)
            minutos = int((tiempo_actual % 3600) // 60)
            segundos = int(tiempo_actual % 60)
            self.label_cronometro.config(text=f"{horas:02d}:{minutos:02d}:{segundos:02d}")

        self.after(1000, self.actualizar_reloj)

    def iniciar_jornada(self):
        if not self.corriendo:
            self.corriendo = True
            self.pausado = False
            self.inicio_tiempo = datetime.now()
            self.btn_inicio.config(state='disabled')
            self.btn_pausa.config(state='normal', text='Pausa')
            self.btn_fin.config(state='normal')
            self.tiempo_acumulado = 0

    def pausar_jornada(self):
        if self.corriendo:
            if not self.pausado:
                self.tiempo_acumulado += (datetime.now() - self.inicio_tiempo).total_seconds()
                self.pausado = True
                self.btn_pausa.config(text='Reanudar')
            else:
                self.inicio_tiempo = datetime.now()
                self.pausado = False
                self.btn_pausa.config(text='Pausa')

    def finalizar_jornada(self):
        if self.corriendo:
            if not self.pausado:
                self.tiempo_acumulado += (datetime.now() - self.inicio_tiempo).total_seconds()
            self.corriendo = False
            self.pausado = False
            self.btn_inicio.config(state='normal')
            self.btn_pausa.config(state='disabled', text='Pausa')
            self.btn_fin.config(state='disabled')
            self.label_cronometro.config(text='00:00:00')

            fecha = datetime.now().strftime("%d/%m/%Y")
            hora_inicio_dt = datetime.now() - timedelta(seconds=self.tiempo_acumulado)
            hora_inicio = hora_inicio_dt.strftime("%H:%M")
            hora_fin = datetime.now().strftime("%H:%M")
            duracion = round(self.tiempo_acumulado / 3600, 4)
            valor_hora = self.valor_hora_base
            total = round(duracion * valor_hora, 2)
            notas = "Registro automático desde cronómetro"

            self.guardar_registro(fecha, hora_inicio, hora_fin, duracion, valor_hora, total, notas)
            self.cargar_registros()

            self.tiempo_acumulado = 0

    def guardar_registro(self, fecha, hora_inicio, hora_fin, duracion, valor_hora, total, notas):
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO registros (usuario, fecha, hora_inicio, hora_fin, duracion, valor_hora, total, notas)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (self.usuario_activo, fecha, hora_inicio, hora_fin, duracion, valor_hora, total, notas))
        conn.commit()
        conn.close()

    def cargar_registros(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT id, fecha, hora_inicio, hora_fin, duracion, valor_hora, total, notas FROM registros
            WHERE usuario = ?
            ORDER BY date(substr(fecha,7,4)||'-'||substr(fecha,4,2)||'-'||substr(fecha,1,2)) DESC, hora_inicio DESC
            3
        ''', (self.usuario_activo,))
        filas = cursor.fetchall()
        conn.close()

        for fila in filas:
            id_, fecha, h_inicio, h_fin, duracion_decimal, valor, total, notas = fila
            duracion_str = self.formato_hhmm(duracion_decimal)
            self.tree.insert("", "end", values=(id_, fecha, h_inicio, h_fin, duracion_str,
                                                f"{valor:.2f}", f"{total:.2f}", notas))
        self._ajustar_columnas()
        self.limpiar_campos_manual()

    def _ajustar_columnas(self):
        for col in self.tree["columns"]:
            max_len = max([len(str(self.tree.set(k, col))) for k in self.tree.get_children()] + [len(col)])
            self.tree.column(col, width=max(80, max_len * 10))

    def borrar_registro(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Advertencia", "Seleccione un registro para borrar.")
            return
        = messagebox.askyesno("Confirmar", "¿Confirma borrar el registro seleccionado?")
        if res:
            item = self.tree.item(selected[0])
            id_reg = item['values'][0]
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute('DELETE FROM registros WHERE id=?', (id_reg,))
            conn.commit()
            conn.close()
            messagebox.showinfo("Borrado", "Registro eliminado correctamente.")
            self.cargar_registros()

    def formato_hhmm(self, duracion_decimal):
        horas = int(duracion_decimal)
        minutos = int(round((duracion_decimal - horas) * 60))
        if minutos == 60:
            horas += 1
            minutos = 0
        return f"{horas:02d}:{minutos:02d}"

    def obtener_hora_inicio(self):
        return f"{self.hora_inicio_h.get()}:{self.hora_inicio_m.get()}"

    def obtener_hora_fin(self):
        return f"{self.hora_fin_h.get()}:{self.hora_fin_m.get()}"

    def guardar_registro_manual(self):
        try:
            fecha = self.fecha_manual.get()
            hora_inicio = self.obtener_hora_inicio()
            hora_fin = self.obtener_hora_fin()
            valor_hora = float(self.valor_hora_manual.get())
            notas = self.notas_manual.get()

            hi = datetime.strptime(hora_inicio, "%H:%M")
            hf = datetime.strptime(hora_fin, "%H:%M")

            duracion_segundos = (hf - hi).total_seconds()
            if duracion_segundos < 0:
                duracion_segundos += 24 * 3600
            duracion = duracion_segundos / 3600

            total = round(duracion * valor_hora, 2)

            self.guardar_registro(fecha, hora_inicio, hora_fin, duracion, valor_hora, total, notas)
            self.cargar_registros()
            self.limpiar_campos_manual()
            messagebox.showinfo("Éxito", "Registro guardado correctamente.")
        except ValueError:
            messagebox.showerror("Error", "Formato de hora inválido o campos vacíos. Use HH:MM para horas.")

    def limpiar_campos_manual(self):
        self.hora_inicio_h.set("00")
        self.hora_inicio_m.set("00")
        self.hora_fin_h.set("00")
        self.hora_fin_m.set("00")
        self.valor_hora_manual.delete(0, tb.END)
        self.valor_hora_manual.insert(0, f"{self.valor_hora_base:.2f}")
        self.notas_manual.delete(0, tb.END)
        self.fecha_manual.set_date(datetime.now())
        self.btn_actualizar.config(state='disabled')
        self.btn_guardar.config(state='normal')
        self.btn_borrar.config(state='disabled')

    def seleccionar_registro(self, event):
        selected = self.tree.selection()
        if selected:
            item = self.tree.item(selected[0])
            valores = item['values']
            self.registro_seleccionado_id = valores[0]
            self.fecha_manual.set_date(datetime.strptime(valores[1], "%d/%m/%Y"))
 hh, mm = valores[2].split(":")
            self.hora_inicio_h.set(hh)
            self.hora_inicio_m.set(mm)
            hh, mm = valores[3].split(":")
            self.hora_fin_h.set(hh)
            self.hora_fin_m.set(mm)
            self.valor_hora_manual.delete(0, tb.END)
            self.valor_hora_manual.insert(0, valores[5])
            self.notas_manual.delete(0, tb.END)
            self.notas_manual.insert(0, valores[7])
            self.btn_actualizar.config(state='normal')
            self.btn_guardar.config(state='disabled')
            self.btn_borrar.config(state='normal')
        else:
            self.limpiar_campos_manual()

    def actualizar_registro(self):
        try:
            fecha = self.fecha_manual.get()
            hora_inicio = self.obtener_hora_inicio()
            hora_fin = self.obtener_hora_fin()
            valor_hora = float(self.valor_hora_manual.get())
            notas = self.notas_manual.get()

            hi = datetime.strptime(hora_inicio, "%H:%M")
            hf = datetime.strptime(hora_fin, "%H:%M")

            duracion_segundos = (hf - hi).total_seconds()
            if duracion_segundos < 0:
                duracion_segundos += 24 * 3600
            duracion = duracion_segundos / 3600
            total = round(duracion * valor_hora, 2)

            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute('''
               UPDATE registros
               SET fecha=?, hora_inicio=?, hora_fin=?, duracion=?, valor_hora=?, total=?, notas=?
               WHERE id=? AND usuario=?
            ''', (fecha, hora_inicio, hora_fin, duracion, valor_hora, total, notas,
                  self.registro_seleccionado_id, self.usuario_activo))
            conn.commit()
            conn.close()

            self.cargar_registros()
            self.limpiar_campos_manual()
            messagebox.showinfo("Éxito", "Registro actualizado correctamente.")
        except ValueError:
            messagebox.showerror("Error", "Formato de hora inválido o campos vacíos. Use HH:MM para horas.")

    def generar_informe(self):
        desde = self.informe_fecha_desde.get_date()
        hasta = self.informe_fecha_hasta.get_date()
        if desde > hasta:
            messagebox.showerror("Error", "La fecha 'Desde' no puede ser mayor que 'Hasta'")
            return None

        mostrar_valor = self.var_mostrar_valor.get()
        mostrar_total = self.var_mostrar_total.get()

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()

        desde_str = desde.strftime("%Y-%m-%d")
        hasta_str = hasta.strftime("%Y-%m-%d")

        query = '''
            SELECT fecha, hora_inicio, hora_fin, duracion, valor_hora, total, notas
            FROM registros 
            WHERE usuario = ? AND date(substr(fecha,7,4)||'-'||substr(fecha,4,2)||'-'||substr(fecha,1,2)) BETWEEN ? AND ?
            ORDER BY date(substr(fecha,7,4)||'-'||substr(fecha,4,2)||'-'||substr(fecha,1,2)) ASC, hora_inicio ASC
        '''
        cursor.execute(query, (self.usuario_activo, desde_str, hasta_str))
        filas = cursor.fetchall()
        conn.close()

        texto = f"Informe de Horas para usuario {self.usuario_activo} desde {desde.strftime('%d/%m/%Y')} hasta {hasta.strftime('%d/%m/%Y')}\n\n"

        texto += f"{'Fecha':10} {'Inicio':6} {'Fin':6} {'Durac. (HH:MM)':15}"
        if mostrar_valor:
            texto += f" {'Valor Hora ($)':15}"
        if mostrar_total:
            texto += f" {'Total ($)':10}"
        texto += " Notas\n"
        texto += "-" * 90 + "\n"

        suma_duracion = 0
        suma_total = 0

        for fila in filas:
            fecha, hi, hf, dur, val, tot, notas = fila
            dur_formato = self.formato_hhmm(dur)
            suma_duracion += dur
            suma_total += tot
            linea = f"{fecha:10} {hi:6} {hf:6} {dur_formato:15}"
            if mostrar_valor:
                linea += f" {val:15.2f}"
            if mostrar_total:
                linea += f" {tot:10.2f}"
            linea += f" {notas}"
            texto += linea + "\n"

        texto += "\n" + "-" * 90 + "\n"
        texto += f"Total horas: {self.formato_hhmm(suma_duracion)} h\n"
        if mostrar_total:
            texto += f"Total pagado: $ {suma_total:.2f}\n"

        return texto

    def ventana_informe_generar(self):
        texto = self.generar_informe()
        if texto is None:
            return
        self.ventana_informe = VentanaInforme(
            self,
            texto,
            export_func=self.exportar_excel_a4,
            imprimir_func=self.imprimir_con_config,
            exportar_txt_func=self.exportar_txt
        )

    def exportar_txt(self):
        texto = self.ventana_informe.text_informe.get('1.0', tb.END).strip()
        if not texto:
            messagebox.showwarning("Advertencia", "No hay para exportar.")
            return

        default_name = f"informe_{self.usuario_activo}_{datetime.now().strftime('%Y%m%d')}.txt"
        filepath = filedialog.asksaveasfilename(defaultextension=".txt",
                                                filetypes=[("Archivos TXT", "*.txt")],
                                                initialfile=default_name)
        if not filepath:
            return

        try:
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(texto            messagebox.showinfo("Éxito", f"Informe TXT exportado a:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo TXT:\n{e}")

    def exportar_excel_a4(self):
        texto = self.ventana_informe.text_informe.get('1.0', tb.END).strip()
        if not texto:
            messagebox.showwarning("Advertencia", "No hay informe para exportar.")
            return

        default_name = f"informe_{self.usuario_activo}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                filetypes=[("Archivos Excel", "*.xlsx")],
                                                initialfile=default_name)
        if not filepath:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Informe Horas"

        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_margins.left = 0.75
        ws.page_margins.right = 0.75
        ws.page_margins.top = 1.0
        ws.page_margins.bottom = 1.0

        fuente_encabezado = Font(bold=True)
        alineacion_centrada = Alignment(horizontal="center")

        lineas = texto.split("\n")

        for row_idx, linea in enumerate(lineas, 1):
            columnas = linea.split()
            for col_idx, valor in enumerate(columnas, 1):
                celda = ws.cell(row=row_idx, column=col_idx, value=valor)
                if row_idx == 1 or row_idx == 3:
                    celda.font = fuente_encabezado
                    celda.alignment = alineacion_centrada
                if row_idx > 3 and col_idx in (4, 5, 6):
                    celda.alignment = Alignment(horizontal="right")

        try:
            wb.save(filepath)
            messagebox.showinfo("Éxito", f"Informe Excel exportado a:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo Excel:\n{e}")

    def imprimir_con_config(self):
        if win32api is None or win32print is None:
            messagebox.showerror("Error", "Para la impresión con diálogo en Windows debe instalar pywin32:\npip install pywin32")
            return

        texto = self.ventana_informe.text_informe.get('1.0', tb.END).strip()
        if not texto:
            messagebox.showwarning("Advertencia", "No hay informe para imprimir.")
            return

        try:
            tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode='w', encodingutf-8')
            tmp_file.write(texto)
            tmp_file.close()

            win32api.ShellExecute(
                0,
                "printto",
                tmp_file.name,
                '"%s"' % win32print.GetDefaultPrinter(),
                ".",
                0
            )
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo imprimir:\n{e}")

    def cambiar_usuario(self):
        # Desactiva controles mientras no hay usuario
        self.usuario_activo = None
        self.valor_hora_base = 0
        self.corriendo = False
        self.pausado = False
        self.inicio_tiempo = None
        self.tiempo_acumulado = 0

        self.label_usuario.config(text="Usuario: Sin sesión")
        self.btn_inicio.config(state='disabled')
        self.btn_pausa.config(state='disabled')
        self.btn_fin.config(state='disabled')

        # Loop para seleccionar usuario al cambiar
        usuario_valido = False
        while not usuario_valido:
            self.ventana_usuarios = VentanaUsuarios(self, self.configuracion)
            self.wait_window(self.ventana_usuarios)
            if not self.ventana_usuarios.usuario_seleccionado:
                respuesta = messagebox.askyesno("Salir", "No se seleccionó usuario.\n¿Desea salir de la aplicación?")
                if respuesta:
                    self.destroy()
                    return
            else:
                usuario_valido = True

        self.usuario_activo = self.ventana_usuarios.usuario_seleccionado
        self.configuracion["ultimo_usuario"] = self.usuario_activo
        guardar_config(self.configuracion)
        self.valor_hora_base = self.configuracion["usuarios"].get(self.usuario_activo, {}).get("valor_hora_base", 0.0)

        ventana_valor = PreguntarValorHora(self, self.valor_hora_base)
        self.wait_window(ventana_valor)
        if ventana_valor.valor_resultado is not None:
            self.valor_hora_base = ventana_valor.valor_resultado
            self.configuracion["usuarios"][self.usuario_activo]["valor_hora_base"] = self.valor_hora_base
            guardar_config(self.configuracion)

        self.label_usuario.config(text=f"Usuario: {self.usuario_activo}")

        self.btn_inicio.config(state='normal')
        self.btn_pausa.config(state='disabled')
        self.btn_fin.config(state='disabled')

        self.cargar_registros()
        self.limpiar_campos_manual()
        self.label_cronometro.config(text='00:00:00')

    def cerrar_sesion(self):
        if self.corriendo:
            respuesta = messagebox.askyesno("Cerrar sesión", "Hay una jornada en curso. ¿Desea finalizarla antes de cerrar sesión?")
            if respuesta:
                self.finalizar_jornada()
            else:
                return
        self.cambiar_usuario()


class VentanaInforme(tb.Toplevel):
    def __init__(self, root, texto, export_func, imprimir_func, exportar_txt_func):
        super().__init__(root)
        self.title("Informe de Horas")
        self.geometry("900x600")
        self.resizable(True, True)

        self.text_informe = tb.Text(self, font=('Arial', 10), wrap='none')
        self.text_informe.pack(fill='both', expand=True)

        scroll_v = tb.Scrollbar(self, orient="vertical", command=self.text_informe.yview)
        scroll_v.pack(side='right', fill='y')
        self.text_informe.configure(yscrollcommand=scroll_v.set)

        scroll_h = tb.Scrollbar(self, orient="horizontal", command=self.text_informe.xview)
        scroll_h.pack(side='bottom', fill='x')
        self.text_informe.configure(xscrollcommand=scroll_h.set)

        self.text_informe.insert('1.0', texto)
        self.text_informe.config(state='disabled')

        frame_botones = tb.Frame(self)
        frame_botones.pack(pady=5)

        btn_exportar_excel = tb.Button(frame_botones, text="Exportar a Excel A4", bootstyle=PRIMARY, command=export_func)
        btn_exportar_excel.pack(side='left', padx=10)

        btn_imprimir = tb.Button(frame_botones, text="Imprimir Informe", bootstyle=SECONDARY, command=imprimir_func)
        btn_imprimir.pack(side='left', padx=10)

        btn_exportar_txt = tb.Button(frame_botones, text="Exportar a TXT", bootstyle=INFO, command=exportar_txt_func)
        btn_exportar_txt.pack(side='left', padx=10)


if __name__ == "__main__":
    app = ControlHorasApp()
    app.mainloop()
