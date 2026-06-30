import tkinter as tk
from tkinter import messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import os

from master_script import generar_output   # Import backend directo


class DragDropUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Procesador NOTA–PEDIDO")
        self.root.geometry("600x440")
        self.root.configure(bg="#f5f5f5")

        self.file1_path = ""
        self.file2_path = ""

        # ---- TÍTULO ----
        tk.Label(
            root,
            text="Arrastra tus archivos",
            font=("Segoe UI", 16, "bold"),
            bg="#f5f5f5"
        ).pack(pady=(18, 8))

        # ---- CUADRO 1: Archivo Output ----
        self.box1 = tk.Label(
            root,
            text="Arrastrar Output (solo se usa el nombre)",
            width=56,
            height=4,
            relief="ridge",
            bg="#ffffff"
        )
        self.box1.pack(pady=10)
        self.box1.drop_target_register(DND_FILES)
        self.box1.dnd_bind('<<Drop>>', self.drop_file1)

        # ---- CUADRO 2: Archivo F ----
        self.box2 = tk.Label(
            root,
            text="Arrastra archivo F",
            width=56,
            height=4,
            relief="ridge",
            bg="#ffffff"
        )
        self.box2.pack(pady=10)
        self.box2.drop_target_register(DND_FILES)
        self.box2.dnd_bind('<<Drop>>', self.drop_file2)

        # ---- PERIODO ----
        periodo_frame = tk.Frame(root, bg="#f5f5f5")
        periodo_frame.pack(pady=(20, 12))

        tk.Label(
            periodo_frame,
            text="PERIODO:",
            font=("Segoe UI", 11),
            bg="#f5f5f5"
        ).grid(row=0, column=0, padx=(0, 10))

        self.periodo_entry = tk.Entry(periodo_frame, width=20, font=("Segoe UI", 12))
        self.periodo_entry.grid(row=0, column=1)
        self.periodo_entry.insert(0, "")

        # ---- BOTÓN EJECUTAR ----
        tk.Button(
            root,
            text="Ejecutar",
            font=("Segoe UI", 13, "bold"),
            bg="#4CAF50",
            fg="white",
            width=16,
            command=self.run_script
        ).pack(pady=20)

        # ---- Nota ----
        tk.Label(
            root,
            text="El archivo final se creará en una carpeta 'Resultado' junto al archivo Output.",
            font=("Segoe UI", 9),
            bg="#f5f5f5",
            fg="#555"
        ).pack(pady=(4, 8))

    # ---------------------------------------------
    # Normalizar ruta
    # ---------------------------------------------
    def _normalize_drop_path(self, data):
        path = data.strip()
        if path.startswith("{") and path.endswith("}"):
            path = path[1:-1]
        parts = path.split()
        for p in parts:
            if os.path.exists(p):
                return p
        return path

    # ---------------------------------------------
    # Drag & Drop handlers
    # ---------------------------------------------
    def drop_file1(self, event):
        self.file1_path = self._normalize_drop_path(event.data)
        self.box1.config(text=os.path.basename(self.file1_path))

    def drop_file2(self, event):
        self.file2_path = self._normalize_drop_path(event.data)
        self.box2.config(text=os.path.basename(self.file2_path))

    # ---------------------------------------------
    # Ejecutar proceso completo
    # ---------------------------------------------
    def run_script(self):
        if not self.file1_path or not self.file2_path:
            messagebox.showerror("Error", "Por favor arrastra ambos archivos (Output y F).")
            return

        periodo_val = self.periodo_entry.get().strip()
        if periodo_val == "":
            messagebox.showerror("Error", "Por favor ingresa un valor para PERIODO.")
            return

        # -------------------------------
        # NUEVA LÓGICA DEL ARCHIVO FINAL
        # -------------------------------
        dir_output = os.path.dirname(self.file1_path)
        base_name = os.path.splitext(os.path.basename(self.file1_path))[0]

        # Crear carpeta Resultado si no existe
        resultado_dir = os.path.join(dir_output, "Resultado")
        os.makedirs(resultado_dir, exist_ok=True)

        # Archivo final
        final_path = os.path.join(resultado_dir, f"{base_name}.xlsx")

        try:
            generar_output(self.file1_path, self.file2_path, final_path, periodo_val)

            messagebox.showinfo(
                "Éxito",
                f"Archivo generado:\n\n{os.path.basename(final_path)}\n\n"
                f"Ubicación:\n{resultado_dir}"
            )
        except Exception as e:
            messagebox.showerror("Error inesperado", str(e))


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = DragDropUI(root)
    root.mainloop()
