#!/usr/bin/env python3
"""
GUI em Tkinter para gerar declarações a partir de um DOCX modelo.
"""
from __future__ import annotations

import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from docx_engine import DeclarationData, apply_declaration


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Gerador de Declaração")
        self.geometry("700x520")
        self.resizable(False, False)
        self._build_widgets()

    def _build_widgets(self) -> None:
        padding = {"padx": 10, "pady": 6}

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        tab_modelo = ttk.Frame(notebook)
        tab_dados = ttk.Frame(notebook)
        tab_saida = ttk.Frame(notebook)

        notebook.add(tab_modelo, text="1. Modelo")
        notebook.add(tab_dados, text="2. Dados")
        notebook.add(tab_saida, text="3. Saída")

        # Modelo
        frm_modelo = ttk.Frame(tab_modelo)
        frm_modelo.pack(fill="x", **padding)
        ttk.Label(frm_modelo, text="Modelo DOCX:").pack(side="left")
        self.model_path_var = tk.StringVar()
        ttk.Entry(frm_modelo, textvariable=self.model_path_var, width=60).pack(side="left", padx=5)
        ttk.Button(frm_modelo, text="Escolher...", command=self._choose_model).pack(side="left")

        # Campos
        frm_fields = ttk.Frame(tab_dados)
        frm_fields.pack(fill="both", expand=True, **padding)

        self.papel_var = tk.StringVar(value="parecerista")
        ttk.Label(frm_fields, text="Papel/Título:").grid(row=0, column=0, sticky="e", padx=4, pady=4)
        ttk.Combobox(
            frm_fields,
            textvariable=self.papel_var,
            values=["parecerista", "orientador", "coorientador"],
            state="readonly",
            width=25,
        ).grid(row=0, column=1, sticky="w", padx=4, pady=4)

        # texto entries
        self.prof_var = tk.StringVar()
        self.titulo_var = tk.StringVar()
        self.autor_var = tk.StringVar()
        self.resp_var = tk.StringVar()
        self.cargo_var = tk.StringVar()
        self.setor_var = tk.StringVar()
        self.cidade_var = tk.StringVar()
        self.data_var = tk.StringVar(value=datetime.now().strftime("%d de %B de %Y"))

        labels = [
            ("Nome do(a) professor(a):", self.prof_var),
            ("Título do TCC:", self.titulo_var),
            ("Autor do TCC:", self.autor_var),
            ("Responsável (assinatura):", self.resp_var),
            ("Cargo do responsável:", self.cargo_var),
            ("Setor/Unidade:", self.setor_var),
            ("Cidade:", self.cidade_var),
            ("Data por extenso:", self.data_var),
        ]
        for idx, (text, var) in enumerate(labels, start=1):
            ttk.Label(frm_fields, text=text).grid(row=idx, column=0, sticky="e", padx=4, pady=4)
            ttk.Entry(frm_fields, textvariable=var, width=50).grid(row=idx, column=1, sticky="w", padx=4, pady=4)

        # Saída
        frm_saida = ttk.Frame(tab_saida)
        frm_saida.pack(fill="x", **padding)
        ttk.Label(frm_saida, text="Pasta de saída:").pack(side="left")
        self.output_dir_var = tk.StringVar(value=str(Path("saida_docx").resolve()))
        ttk.Entry(frm_saida, textvariable=self.output_dir_var, width=60).pack(side="left", padx=5)
        ttk.Button(frm_saida, text="Escolher...", command=self._choose_output_dir).pack(side="left")

        # Botões
        frm_buttons = ttk.Frame(tab_saida)
        frm_buttons.pack(fill="x", pady=12)
        ttk.Button(frm_buttons, text="Gerar DOCX", command=self._on_generate).pack(side="left", padx=10)
        ttk.Button(frm_buttons, text="Sair", command=self.destroy).pack(side="left", padx=10)

    def _choose_model(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Documentos Word", "*.docx")])
        if path:
            self.model_path_var.set(path)

    def _choose_output_dir(self) -> None:
        path = filedialog.askdirectory()
        if path:
            self.output_dir_var.set(path)

    def _collect_data(self) -> DeclarationData:
        return DeclarationData(
            papel=self.papel_var.get().strip() or "parecerista",
            professor=self.prof_var.get().strip(),
            titulo_tcc=self.titulo_var.get().strip(),
            autor_tcc=self.autor_var.get().strip(),
            responsavel=self.resp_var.get().strip(),
            cargo=self.cargo_var.get().strip(),
            setor=self.setor_var.get().strip(),
            cidade=self.cidade_var.get().strip(),
            data_extenso=self.data_var.get().strip(),
        )

    def _on_generate(self) -> None:
        try:
            template_path = Path(self.model_path_var.get())
            output_dir = Path(self.output_dir_var.get()) if self.output_dir_var.get() else Path("saida_docx")
            data = self._collect_data()

            if not template_path.exists():
                messagebox.showerror("Erro", f"Modelo não encontrado: {template_path}")
                return

            output_path = apply_declaration(template_path, output_dir, data)
            messagebox.showinfo("Sucesso", f"Arquivo gerado em:\n{output_path}")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Erro", str(exc))


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
