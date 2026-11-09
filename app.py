# -*- coding: utf-8 -*-
"""ayuda-colegio: ventana simple para ordenar archivos y actualizar el Excel de casos.
Todo ocurre en este equipo: no se envía ninguna información a internet."""

import os
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

sys.path.insert(0, str(Path(__file__).resolve().parent))
import nucleo_excel as nx
import nucleo_ordenar as no

TITULO = "ayuda-colegio"


def abrir_archivo(ruta):
    """Abre un archivo con el programa predeterminado del sistema."""
    if sys.platform == "darwin":
        subprocess.Popen(["open", str(ruta)])
    elif sys.platform.startswith("win"):
        os.startfile(str(ruta))  # noqa
    else:
        subprocess.Popen(["xdg-open", str(ruta)])


class Aviso(tk.Toplevel):
    """Aviso de documento escaneado: Continuar / Continuar con todos / Abrir."""

    def __init__(self, padre, nombre):
        super().__init__(padre)
        self.title("Leer manualmente")
        self.resultado = "continuar"
        self.resizable(False, False)
        self.transient(padre)
        ttk.Label(self, padding=16, justify="left", wraplength=420,
                  text=f"Este archivo parece escaneado y no se puede leer solo:\n\n{nombre}\n\n"
                       "Queda anotado como «LEER MANUALMENTE».").pack()
        fila = ttk.Frame(self, padding=(16, 0, 16, 16))
        fila.pack()
        for texto, valor in (("Continuar", "continuar"), ("Continuar con todos", "todos"),
                             ("Abrir (se detiene)", "abrir")):
            ttk.Button(fila, text=texto, command=lambda v=valor: self._cerrar(v)).pack(side="left", padx=4)
        self.protocol("WM_DELETE_WINDOW", lambda: self._cerrar("continuar"))
        self.update_idletasks()
        self.geometry(f"+{padre.winfo_rootx() + 60}+{padre.winfo_rooty() + 80}")
        self.grab_set()

    def _cerrar(self, valor):
        self.resultado = valor
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(TITULO)
        self.geometry("720x520")
        self.minsize(600, 440)
        self.carpeta = tk.StringVar()
        fr = ttk.Frame(self, padding=14)
        fr.pack(fill="both", expand=True)

        ttk.Label(fr, text="Paso 1. Elige la carpeta de trabajo:", font=("", 12, "bold")).pack(anchor="w")
        fila = ttk.Frame(fr)
        fila.pack(fill="x", pady=(6, 14))
        ttk.Entry(fila, textvariable=self.carpeta).pack(side="left", fill="x", expand=True)
        ttk.Button(fila, text="Examinar...", command=self.elegir).pack(side="left", padx=(6, 0))

        ttk.Label(fr, text="Paso 2. Pulsa lo que necesites:", font=("", 12, "bold")).pack(anchor="w")
        botones = ttk.Frame(fr)
        botones.pack(fill="x", pady=(6, 12))
        self.b_ordenar = ttk.Button(botones, text="Ordenar archivos", command=self.ordenar)
        self.b_ordenar.pack(side="left", padx=(0, 8), ipadx=10, ipady=6)
        self.b_excel = ttk.Button(botones, text="Actualizar Excel", command=self.actualizar_excel)
        self.b_excel.pack(side="left", ipadx=10, ipady=6)

        self.salida = scrolledtext.ScrolledText(fr, wrap="word", height=16)
        self.salida.pack(fill="both", expand=True)
        self.escribir("Elige la carpeta y pulsa un botón.\n\n"
                      "• Ordenar archivos: mueve cada archivo a inicial, primaria o secundaria.\n"
                      "• Actualizar Excel: crea o actualiza «casos-niños.xlsx» en esa carpeta.\n\n"
                      "Nunca se borra ni se sobrescribe ningún archivo.")

    # ----- utilidades de ventana
    def escribir(self, txt):
        self.salida.configure(state="normal")
        self.salida.delete("1.0", "end")
        self.salida.insert("end", txt)
        self.salida.configure(state="disabled")

    def _botones(self, activos):
        for b in (self.b_ordenar, self.b_excel):
            b.configure(state="normal" if activos else "disabled")

    def elegir(self):
        c = filedialog.askdirectory(title="Carpeta de trabajo")
        if c:
            self.carpeta.set(c)

    def _carpeta_valida(self):
        c = self.carpeta.get().strip()
        if not c or not Path(c).is_dir():
            messagebox.showwarning(TITULO, "Primero elige la carpeta (Paso 1).")
            return None
        return Path(c)

    def _en_hilo(self, trabajo, fin):
        self._botones(False)

        def run():
            try:
                r = trabajo()
                self.after(0, lambda: (self._botones(True), fin(r, None)))
            except Exception as e:
                self.after(0, lambda err=e: (self._botones(True), fin(None, err)))
        threading.Thread(target=run, daemon=True).start()

    # ----- Ordenar
    def ordenar(self):
        base = self._carpeta_valida()
        if not base:
            return
        self.escribir("Revisando archivos...")
        self._en_hilo(lambda: no.planificar(base), lambda plan, error: self._confirmar_ordenar(base, plan, error))

    def _confirmar_ordenar(self, base, plan, error):
        if error:
            self.escribir("No se pudo revisar la carpeta.")
            messagebox.showerror(TITULO, f"No se pudo revisar la carpeta:\n{error}")
            return
        if not plan:
            self.escribir("No hay archivos sueltos por ordenar en esta carpeta.")
            return
        a_rev = sum(1 for _, d in plan if d == no.REVISAR)
        msg = f"Se moverán {len(plan)} archivos."
        if a_rev:
            msg += f"\n{a_rev} no se pueden clasificar y irán a la carpeta {no.REVISAR}."
        if not messagebox.askyesno(TITULO, msg + "\n\n¿Continuar?"):
            self.escribir("Cancelado. No se movió nada.")
            return
        self.escribir("Ordenando...")
        self._en_hilo(lambda: no.mover(base, plan), self._fin_ordenar)

    def _fin_ordenar(self, r, error):
        if error:
            self.escribir(f"Ocurrió un error: {error}")
            return
        movidos, errores = r
        lineas = ["Listo. Archivos movidos:"]
        for k in no.NIVELES + [no.REVISAR]:
            if movidos[k]:
                lineas.append(f"  {k}: {len(movidos[k])}")
        if movidos[no.REVISAR]:
            lineas += ["", f"Revisar a mano (carpeta {no.REVISAR}):"] + [f"  - {n}" for n in movidos[no.REVISAR]]
        if errores:
            lineas += ["", "No se pudieron mover:"] + [f"  - {e}" for e in errores]
        lineas += ["", "Cuando quieras, pulsa «Actualizar Excel»."]
        self.escribir("\n".join(lineas))

    # ----- Excel
    def _preguntar(self, nombre, ruta):
        """Se llama desde el hilo de trabajo: muestra el aviso en la ventana y espera la respuesta."""
        hecho, caja = threading.Event(), {}

        def mostrar():
            try:
                d = Aviso(self, nombre)
                self.wait_window(d)
                caja["r"] = d.resultado
                if d.resultado == "abrir":
                    try:
                        abrir_archivo(ruta)
                    except Exception:
                        pass
            except Exception:
                pass
            finally:  # el hilo de trabajo nunca debe quedar esperando
                caja.setdefault("r", "continuar")
                hecho.set()
        self.after(0, mostrar)
        hecho.wait()
        return caja["r"]

    def actualizar_excel(self):
        base = self._carpeta_valida()
        if not base:
            return
        self.escribir("Leyendo archivos y actualizando el Excel...\n(puede tardar un poco)")
        self._en_hilo(lambda: nx.actualizar(base, self._preguntar), self._fin_excel)

    def _fin_excel(self, r, error):
        if isinstance(error, nx.ExcelAbierto):
            self.escribir("Tienes abierto un Excel en el cual vamos a trabajar.\nCiérralo y vuelve a intentar.")
            messagebox.showwarning(TITULO, "Tienes abierto un Excel en el cual vamos a trabajar.\n"
                                           "Ciérralo y vuelve a intentar.")
            return
        if isinstance(error, (nx.ErrorCarpeta, nx.ExcelInvalido)):
            self.escribir(str(error))
            messagebox.showwarning(TITULO, str(error))
            return
        if error:
            self.escribir(f"Ocurrió un error: {error}")
            return
        l = [("Excel creado." if r["creado"] else "Excel actualizado.") + f"  ({r['ruta']})", "",
             f"Archivos revisados: {r['total']}",
             f"  filas nuevas: {r['nuevos']}",
             f"  filas actualizadas: {r['actualizados']}"]
        if r["movidos_hoja"]:
            l.append(f"  filas que cambiaron de hoja: {r['movidos_hoja']}")
        if r["no_encontrados"]:
            l.append(f"  marcados «Archivo no encontrado»: {r['no_encontrados']}")
        if r["huerfanos"]:
            l += ["", "Para leer manualmente (aparecen en el Excel):"] + [f"  - {n}" for n in r["huerfanos"]]
        if r["detenido"]:
            l += ["", "Se detuvo para que abras el archivo. Cuando termines, pulsa de nuevo «Actualizar Excel»."]
        if r["errores"]:
            l += ["", "No se pudieron leer:"] + [f"  - {e}" for e in r["errores"]]
        self.escribir("\n".join(l))


if __name__ == "__main__":
    App().mainloop()
