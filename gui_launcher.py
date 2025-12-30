import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import datetime as _dt
import threading

from generator_core import generate_docs
from document_generator import TestCaseDocumentConfig, FontConfig
from processor import regenerate_single


DEFAULT_FONT = FontConfig("Segoe UI", 12,
                          "Segoe UI", 11,
                          "Segoe UI", 11,
                          "Segoe UI", 11,
                          "Segoe UI", 12)

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
        self.geometry("800x300")
        self.resizable(False, False)

        # ----- StringVars vacíos -----
        self.folder_src = tk.StringVar(value=".")
        self.folder_dst = tk.StringVar(value=".")
        self.project    = tk.StringVar(value="")
        self.analyst    = tk.StringVar(value="")
        self.date       = tk.StringVar(value=_dt.date.today().strftime("%d/%m/%Y"))
        self.header_img = tk.StringVar(value="")
        self.footer_img = tk.StringVar(value="")

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

        ttk.Button(self, text="Generar", command=self._run).pack(pady=12)
        self.status = ttk.Label(self, text="", foreground="green")
        self.status.pack()

    # --------------- helper browse ---------------
    def _browse(self, var: tk.StringVar, path_mode: bool):
        path = filedialog.askdirectory() if path_mode else filedialog.askopenfilename(
            title="Selecciona una imagen",
            filetypes=[("Imágenes", "*.png *.jpg *.jpeg *.bmp *.gif"), ("Todos", "*.*")]
        )
        if path:
            var.set(path)
            entry = self.entries[var]
            entry.config(foreground="black", font=EDIT_FONT)

    # --------------- generar ---------------------
    def _run(self):
        for img in (self.header_img.get(), self.footer_img.get()):
            if img and not img.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".gif")):
                messagebox.showerror("Imagen no válida",
                                     "Las imágenes deben ser .png, .jpg, .jpeg, .bmp o .gif")
                return
        self.status.config(text="Procesando...")
        threading.Thread(target=self._worker, daemon=True).start()

    # def _worker(self):
    #     try:
    #         cfg = TestCaseDocumentConfig(
    #             header_image=self.header_img.get() or "header.png",
    #             footer_image=self.footer_img.get() or "footer.png",
    #             project_name=self.project.get() or "Nombre del Proyecto",
    #             analyst_name=self.analyst.get() or "Nombre del Analista de Calidad",
    #             date=self.date.get(),
    #             success_message="Resultado del caso de prueba: Éxito",
    #             font_config=DEFAULT_FONT,
    #         )
    #         generate_docs(
    #             Path(self.folder_src.get()),
    #             cfg,
    #             dest_root=Path(self.folder_dst.get()) or None
    #         )
    #     except Exception as exc:  # noqa: BLE001
    #         self._finish(f"Error: {exc}", "red")
    #     else:
    #         self._finish("¡Documentos de Word generados con éxito!", "green")

    # def _finish(self, msg, color):
    #     self.after(0, lambda: self.status.config(text=msg, foreground=color))

    # def _worker(self):
    #     try:
    #         cfg = TestCaseDocumentConfig(
    #             header_image=self.header_img.get() or "header.png",
    #             footer_image=self.footer_img.get() or "footer.png",
    #             project_name=self.project.get()  or "Nombre del Proyecto",
    #             analyst_name=self.analyst.get()  or "Nombre del Analista de Calidad",
    #             date=self.date.get(),
    #             success_message="Resultado del caso de prueba: Éxito",
    #             font_config=DEFAULT_FONT,
    #         )

    #         # 1ª pasada: NO sobrescribe
    #         collisions = generate_docs(
    #             Path(self.folder_src.get()), cfg,
    #             dest_root=Path(self.folder_dst.get()) or None,
    #             overwrite=False
    #         )

    #         # Si hubo colisiones, preguntar
    #         if collisions:
    #             overwrite = self._ask_overwrite(len(collisions))
    #             if overwrite:
    #                 # 2ª pasada con overwrite=True
    #                 generate_docs(
    #                     Path(self.folder_src.get()), cfg,
    #                     dest_root=Path(self.folder_dst.get()) or None,
    #                     overwrite=True
    #                 )
    #             else:
    #                 self._finish("Generación cancelada por usuario.", "red")
    #                 return

    #     except Exception as exc:   # noqa: BLE001
    #         self._finish(f"Error: {exc}", "red")
    #     else:
    #         self._finish("¡Documentos de Word generados con éxito!", "green")

    # # ---------- diálogo sí / no ------------------
    # def _ask_overwrite(self, n: int) -> bool:
    #     var = tk.BooleanVar()

    #     def _ask():
    #         res = messagebox.askyesno(
    #             "Archivos existentes",
    #             f"Se encontraron {n} archivos Word ya existentes.\n"
    #             "¿Deseas sobrescribirlos?"
    #         )
    #         var.set(res)

    #     self.after(0, _ask)
    #     self.wait_variable(var)
    #     return var.get()
    def _worker(self):
        try:
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
                evidencia="",
                version="001",
            )

            collisions = generate_docs(
                Path(self.folder_src.get()),
                cfg,
                dest_root=Path(self.folder_dst.get()) or None,
                overwrite=False
            )

            if collisions:
                decision = self._ask_global(len(collisions))
                if decision == "cancel":
                    self._finish("Operación cancelada.", "red"); return
                if decision == "skip_all":
                    self._finish("Generación terminada (se omitieron archivos existentes).", "green"); return
                if decision == "replace_all":
                    generate_docs(
                        Path(self.folder_src.get()),
                        cfg,
                        dest_root=Path(self.folder_dst.get()) or None,
                        overwrite=True
                    )
                elif decision == "ask_each":
                    self._ask_each(collisions, cfg)

        except Exception as exc:  # noqa: BLE001
            self._finish(f"Error: {exc}", "red")
        else:
            self._finish("¡Documentos de Word generados con éxito!", "green")

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
                regenerate_single(p, cfg); continue
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
                regenerate_single(p, cfg)
            elif choice == "yes_all":
                replace_all = True
                regenerate_single(p, cfg)
            elif choice == "no_all":
                skip_all = True


if __name__ == "__main__":
    App().mainloop()
