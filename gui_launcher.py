import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import datetime as _dt
import logging
import threading

from generator_core import generate_docs
from document_generator import TestCaseDocumentConfig, FontConfig
from processor import regenerate_single


def _setup_logging(dest: Path) -> Path:
    """Configura logging hacia un archivo dentro de la carpeta destino.
    Devuelve la ruta del log para mostrarla al usuario."""
    log_path = dest / "testdocgen.log"
    root = logging.getLogger()
    # Limpiar handlers previos para evitar duplicados entre ejecuciones
    for h in list(root.handlers):
        root.removeHandler(h)
    handler = logging.FileHandler(log_path, mode="a", encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s: %(message)s"))
    root.addHandler(handler)
    root.setLevel(logging.INFO)
    return log_path


DEFAULT_FONT = FontConfig("Source Sans Pro", 12,
                          "Source Sans Pro", 11,
                          "Source Sans Pro", 11,
                          "Source Sans Pro", 11,
                          "Source Sans Pro", 12)

PLACE_FONT = ("Segoe UI", 9, "italic")
EDIT_FONT  = ("Segoe UI", 10, "normal")


def add_placeholder(entry: tk.Entry, placeholder: str) -> None:
    """Añade placeholder gris + cursiva si el Entry está vacío."""
    def _put_placeholder():
        entry.insert(0, placeholder)
        entry.config(foreground="grey", font=PLACE_FONT)

    def on_focus_in(_):
        if entry.get() == placeholder:
            entry.delete(0, "end")
            entry.config(foreground="black", font=EDIT_FONT)

    def on_focus_out(_):
        if not entry.get():
            _put_placeholder()

    # Sólo insertar si realmente está vacío
    if not entry.get():
        _put_placeholder()

    entry.bind("<FocusIn>", on_focus_in)
    entry.bind("<FocusOut>", on_focus_out)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Generador De Evidencias De Casos de Prueba - Azure DevOps")
        self.geometry("800x460")
        self.resizable(False, False)

        # ----- StringVars vacíos -----
        self.folder_src = tk.StringVar(value=".")
        self.folder_dst = tk.StringVar(value=".")
        self.project    = tk.StringVar(value="")
        self.analyst    = tk.StringVar(value="")
        self.date       = tk.StringVar(value=_dt.date.today().strftime("%d/%m/%Y"))
        self.header_img = tk.StringVar(value="")
        self.footer_img = tk.StringVar(value="")
        self.evidencia  = tk.StringVar(value="")
        self.version    = tk.StringVar(value="001")
        self.generate_pdf = tk.BooleanVar(value=True)
        
        # Dictionary to track Entry widgets for later access
        self.entries = {}

        self._build_form()

    # ---------------- interfaz ----------------
    def _build_form(self):
        pad = {"padx": 6, "pady": 2}

        def row(label: str, var: tk.StringVar, placeholder: str, tooltip: str,
                browse=False, is_path=False):
            frame = tk.Frame(self)
            frame.pack(fill="x", **pad)

            ttk.Label(frame, text=label, width=20).pack(side="left")
            entry = ttk.Entry(frame, textvariable=var)
            entry.pack(side="left", fill="x", expand=1)
            add_placeholder(entry, placeholder)
            
            # Store entry widget reference for later access (use id() since StringVar is not hashable)
            self.entries[id(var)] = entry

            if browse:
                ttk.Button(frame, text="…",
                           command=lambda v=var, p=is_path: self._browse(v, p)
                           ).pack(side="left", padx=(4, 2))

            ttk.Button(frame, text="?", width=2,
                       command=lambda txt=tooltip: messagebox.showinfo("Ayuda", txt)
                       ).pack(side="left")

        # filas
        row("Carpeta origen:", self.folder_src, ".", 
            "Selecciona la carpeta que contiene los .xlsx exportados desde Azure DevOps.",
            browse=True, is_path=True)

        row("Carpeta destino:", self.folder_dst, ".", 
            "Aquí se crearán las subcarpetas de cada plan y los .docx generados.",
            browse=True, is_path=True)

        row("Proyecto:", self.project, "Nombre del Proyecto",
            "Nombre del proyecto que aparecerá en la tabla.")

        row("Analista QA:", self.analyst, "Nombre del Analista de Calidad",
            "Nombre de la persona responsable de QA.")

        row("Fecha:", self.date, _dt.date.today().strftime("%d/%m/%Y"),
            "Fecha que se mostrará en el documento (DD/MM/AAAA).")

        row("Imagen encabezado:", self.header_img, "header.png",
            "PNG/JPG que se insertará en la cabecera.  Si no se especifica, se usará la imagen de ICETEX por defecto.",
            browse=True, is_path=False)

        row("Imagen pie:", self.footer_img, "footer.png",
            "PNG/JPG que se insertará en el pie de página.  Si no se especifica, se usará la imagen de ICETEX por defecto.",
            browse=True, is_path=False)

        row("Evidencias:", self.evidencia, "Texto/ruta/descripción de la evidencia",
            "Información (ruta, descripción, hash, etc.) que se inscribirá en la celda 'Evidencia' del documento.")

        row("Versión:", self.version, "001",
            "Versión del documento (campo V en el nombre del archivo).")

        # Checkbox PDF
        pdf_frame = tk.Frame(self)
        pdf_frame.pack(fill="x", **pad)
        ttk.Label(pdf_frame, text="", width=20).pack(side="left")
        ttk.Checkbutton(pdf_frame, text="Generar también copia PDF",
                        variable=self.generate_pdf).pack(side="left")
        ttk.Button(pdf_frame, text="?", width=2,
                   command=lambda: messagebox.showinfo(
                       "Ayuda",
                       "Si está activo, se intenta generar un .pdf junto al .docx usando "
                       "Microsoft Word (o LibreOffice). Desmárcalo si Word no está disponible "
                       "o no quieres el PDF.")).pack(side="left", padx=(4, 0))

        self.generate_btn = ttk.Button(self, text="Generar", command=self._run)
        self.generate_btn.pack(pady=12)

        # Barra de progreso + status detallado
        progress_frame = tk.Frame(self)
        progress_frame.pack(fill="x", padx=20, pady=(2, 4))
        self.progress = ttk.Progressbar(progress_frame, mode="determinate", length=600)
        self.progress.pack(fill="x")
        info_frame = tk.Frame(self)
        info_frame.pack(fill="x", padx=20)
        self.counter = ttk.Label(info_frame, text="", foreground="black",
                                 font=("Segoe UI", 9))
        self.counter.pack(side="right")
        self.status = ttk.Label(info_frame, text="", foreground="green",
                                font=("Segoe UI", 9))
        self.status.pack(side="left")

    # --------------- helper browse ---------------
    def _browse(self, var: tk.StringVar, path_mode: bool):
        path = filedialog.askdirectory() if path_mode else filedialog.askopenfilename(
            title="Selecciona una imagen",
            filetypes=[("Imágenes", "*.png *.jpg *.jpeg *.bmp *.gif"), ("Todos", "*.*")]
        )
        if path:
            var.set(path)
            var_id = id(var)
            if var_id in self.entries:
                entry = self.entries[var_id]
                entry.config(foreground="black", font=EDIT_FONT)

    # --------------- generar ---------------------
    def _run(self):
        for img in (self.header_img.get(), self.footer_img.get()):
            if img and not img.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".gif")):
                messagebox.showerror("Imagen no válida",
                                     "Las imágenes deben ser .png, .jpg, .jpeg, .bmp o .gif")
                return
        self.generate_btn.config(state="disabled")
        self.status.config(text="Iniciando...", foreground="blue")
        self.counter.config(text="")
        self.progress.config(mode="determinate", value=0, maximum=100)
        self.update_idletasks()
        threading.Thread(target=self._worker, daemon=True).start()

    def _finish(self, msg, color):
        """Update status label on the main thread; reactivate Generar button."""
        def _do():
            self.status.config(text=msg, foreground=color)
            self.generate_btn.config(state="normal")
            # Si terminamos OK, dejar la barra al 100 %
            if color == "green" and self.progress["maximum"] > 0:
                self.progress["value"] = self.progress["maximum"]
                self.counter.config(text="100%")
        self.after(0, _do)

    def _on_progress(self, msg: str, current: int, total: int):
        """Callback que el worker llama desde su hilo. Despachamos al main."""
        def _do():
            self.status.config(text=msg, foreground="blue")
            if total > 0:
                self.progress["maximum"] = total
                self.progress["value"] = current
                pct = int(current * 100 / total) if total else 0
                self.counter.config(text=f"{current}/{total}  ({pct}%)")
            else:
                self.progress["value"] = 0
                self.counter.config(text="")
        self.after(0, _do)

    def _worker(self):
        try:
            ev_default = "Texto/ruta/descripción de la evidencia"
            evidencia_val = self.evidencia.get()
            if evidencia_val == ev_default:
                evidencia_val = ""

            version_val = (self.version.get() or "001").strip() or "001"
            gen_pdf = self.generate_pdf.get()
            dest_path = Path(self.folder_dst.get() or ".")
            dest_path.mkdir(parents=True, exist_ok=True)
            log_path = _setup_logging(dest_path)
            logging.info("=== Inicio de generación ===")
            logging.info("PDF activado: %s", gen_pdf)

            cfg = TestCaseDocumentConfig(
                header_image=self.header_img.get() or "header.png",
                footer_image=self.footer_img.get() or "footer.png",
                project_name=self.project.get()  or "Nombre del Proyecto",
                analyst_name=self.analyst.get()  or "Nombre del Analista de Calidad",
                date=self.date.get(),
                success_message="Resultado del caso de prueba: Éxito",
                font_config=DEFAULT_FONT,
                template_path=None,
                privacy_classification="DOCUMENTO PRIVADO",
                resultado="Exito",
                evidencia=evidencia_val,
                version=version_val,
            )

            report = generate_docs(
                Path(self.folder_src.get()),
                cfg,
                dest_root=dest_path,
                overwrite=False,
                generate_pdf=gen_pdf,
                progress_callback=self._on_progress,
            )

            if report.collisions:
                decision = self._ask_global(len(report.collisions))
                if decision == "cancel":
                    self._finish("Operación cancelada.", "red"); return
                if decision == "skip_all":
                    self._finish(self._summary(report) + " (existentes omitidos)", "green"); return
                if decision == "replace_all":
                    report = generate_docs(
                        Path(self.folder_src.get()),
                        cfg,
                        dest_root=dest_path,
                        overwrite=True,
                        generate_pdf=gen_pdf,
                        progress_callback=self._on_progress,
                    )
                elif decision == "ask_each":
                    self._ask_each(report.collisions, cfg)

            logging.info("=== Fin de generación: %s ===", self._summary(report))

        except Exception as exc:  # noqa: BLE001
            logging.exception("Error inesperado durante la generación")
            self._finish(f"Error: {exc}", "red")
        else:
            self._finish(self._summary(report), "green")

    def _summary(self, report) -> str:
        if report.docx_count == 0:
            src = Path(self.folder_src.get()).resolve()
            return (
                f"⚠ No se encontró ningún .xlsx en '{src}'. "
                "Selecciona la 'Carpeta origen' que contiene los Excel."
            )
        pdf_part = ""
        if self.generate_pdf.get():
            pdf_part = f"; {report.pdf_count} PDFs"
            if report.pdf_failed:
                pdf_part += f" ({report.pdf_failed} fallidos)"
        return f"OK: {report.docx_count} documentos Word{pdf_part}."

    # ---------- diálogo global (4 opciones) -----------------------------------
    def _ask_global(self, n: int) -> str:
        dlg = tk.Toplevel(self); dlg.title("Archivos existentes")
        ttk.Label(dlg, text=f"Hay {n} archivos Word que ya existen.\n"
                            "¿Qué deseas hacer?").pack(padx=20, pady=10)

        ans = tk.StringVar(value="cancel")

        def set_and_close(val):
            ans.set(val)
            dlg.destroy()

        btns = [("Reemplazar todos", "replace_all"),
                ("Preguntar uno-a-uno", "ask_each"),
                ("Omitir todos", "skip_all"),
                ("Cancelar", "cancel")]
        for text, val in btns:
            ttk.Button(dlg, text=text, command=lambda v=val: set_and_close(v))\
                .pack(fill="x", padx=20, pady=2)

        dlg.grab_set(); self.wait_window(dlg)
        return ans.get()

    # ---------- diálogo por-archivo ------------------------------------------
    def _ask_each(self, paths: list[Path], cfg):
        replace_all = False
        skip_all    = False

        for p in paths:
            if replace_all:
                regenerate_single(p, cfg, generate_pdf=self.generate_pdf.get()); continue
            if skip_all:
                continue

            dlg = tk.Toplevel(self); dlg.title("Sobrescribir")
            ttk.Label(dlg, text=f"El archivo ya existe:\n{p.name}").pack(padx=20, pady=10)

            ans = tk.StringVar()

            def choose(v):
                ans.set(v); dlg.destroy()

            row = tk.Frame(dlg); row.pack(pady=4)
            ttk.Button(row, text="Reemplazar",     command=lambda: choose("yes")).pack(side="left", padx=2)
            ttk.Button(row, text="Omitir",         command=lambda: choose("no")).pack(side="left", padx=2)
            ttk.Button(row, text="Sí a todos",     command=lambda: choose("yes_all")).pack(side="left", padx=2)
            ttk.Button(row, text="No a todos",     command=lambda: choose("no_all")).pack(side="left", padx=2)
            ttk.Button(row, text="Cancelar",       command=lambda: choose("cancel")).pack(side="left", padx=2)

            dlg.grab_set(); self.wait_window(dlg)
            choice = ans.get()

            if choice == "cancel":
                break
            if choice == "yes":
                regenerate_single(p, cfg, generate_pdf=self.generate_pdf.get())
            elif choice == "yes_all":
                replace_all = True
                regenerate_single(p, cfg, generate_pdf=self.generate_pdf.get())
            elif choice == "no_all":
                skip_all = True


if __name__ == "__main__":
    App().mainloop()
