#!/usr/bin/env python3
"""Painel grafico do RobotToWord.

Interface simples para a equipe usar sem linha de comando: escolher os PDFs do
Robot, ajustar duas ou tres opcoes e gerar os Words com um clique.

    python painel.py
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from robot_to_word import Options, convert

# Rotulos amigaveis -> valores que o conversor entende.
QUALITIES = {
    "Normal (150 DPI)": 150,
    "Alta (200 DPI)": 200,
    "Muito alta (300 DPI)": 300,
}
CROP_MODES = {
    "Largura uniforme (recomendado)": "uniform",
    "Colado no conteúdo de cada página": "tight",
    "Sem recorte (página inteira do PDF)": "none",
}
FORMATS = {
    "PNG (mais nítido)": "png",
    "JPEG (arquivo menor)": "jpeg",
}


def open_in_explorer(path: Path) -> None:
    """Abre a pasta no gerenciador de arquivos do sistema."""
    if sys.platform == "win32":
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.run(["open", str(path)], check=False)
    else:
        subprocess.run(["xdg-open", str(path)], check=False)


class Panel(ttk.Frame):
    def __init__(self, master: tk.Tk) -> None:
        super().__init__(master, padding=12)
        self.grid(sticky="nsew")
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.pdfs: list[Path] = []
        self.messages: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.last_output_dir: Path | None = None

        self._build_file_section(row=0)
        self._build_output_section(row=1)
        self._build_options_section(row=2)
        self._build_run_section(row=3)

        self._refresh_output_state()
        self.after(100, self._drain_messages)

    # ------------------------------------------------------------------ layout

    def _build_file_section(self, row: int) -> None:
        box = ttk.LabelFrame(self, text="1. PDFs do Robot", padding=8)
        box.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        box.columnconfigure(0, weight=1)
        box.rowconfigure(0, weight=1)

        self.listbox = tk.Listbox(box, height=6, selectmode=tk.EXTENDED,
                                  activestyle="none")
        self.listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(box, orient="vertical", command=self.listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.listbox.configure(yscrollcommand=scrollbar.set)

        buttons = ttk.Frame(box)
        buttons.grid(row=0, column=2, sticky="n", padx=(8, 0))
        ttk.Button(buttons, text="Adicionar PDFs...", width=18,
                   command=self.add_files).grid(row=0, column=0, pady=2)
        ttk.Button(buttons, text="Adicionar pasta...", width=18,
                   command=self.add_folder).grid(row=1, column=0, pady=2)
        ttk.Button(buttons, text="Remover selecionados", width=18,
                   command=self.remove_selected).grid(row=2, column=0, pady=2)
        ttk.Button(buttons, text="Limpar lista", width=18,
                   command=self.clear_files).grid(row=3, column=0, pady=2)

    def _build_output_section(self, row: int) -> None:
        box = ttk.LabelFrame(self, text="2. Saída", padding=8)
        box.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        box.columnconfigure(1, weight=1)

        self.merge = tk.BooleanVar(value=False)
        ttk.Radiobutton(box, text="Um Word para cada PDF", variable=self.merge,
                        value=False, command=self._refresh_output_state
                        ).grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Radiobutton(box, text="Um Word único com todos os PDFs",
                        variable=self.merge, value=True,
                        command=self._refresh_output_state
                        ).grid(row=1, column=0, columnspan=3, sticky="w")

        ttk.Label(box, text="Nome do arquivo:").grid(row=2, column=0, sticky="w",
                                                     padx=(20, 6), pady=(4, 0))
        self.merge_name = tk.StringVar(value="Memorial de Ligações.docx")
        self.merge_entry = ttk.Entry(box, textvariable=self.merge_name)
        self.merge_entry.grid(row=2, column=1, columnspan=2, sticky="ew", pady=(4, 0))

        ttk.Label(box, text="Pasta de saída:").grid(row=3, column=0, sticky="w",
                                                    pady=(6, 0))
        self.out_dir = tk.StringVar()
        ttk.Entry(box, textvariable=self.out_dir).grid(row=3, column=1, sticky="ew",
                                                       pady=(6, 0))
        ttk.Button(box, text="...", width=4, command=self.choose_out_dir
                   ).grid(row=3, column=2, padx=(4, 0), pady=(6, 0))
        ttk.Label(box, text="Em branco: grava na mesma pasta de cada PDF.",
                  foreground="gray40").grid(row=4, column=1, sticky="w")

    def _build_options_section(self, row: int) -> None:
        box = ttk.LabelFrame(self, text="3. Opções", padding=8)
        box.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        box.columnconfigure(1, weight=1)
        box.columnconfigure(3, weight=1)

        ttk.Label(box, text="Qualidade:").grid(row=0, column=0, sticky="w")
        self.quality = tk.StringVar(value="Alta (200 DPI)")
        ttk.Combobox(box, textvariable=self.quality, values=list(QUALITIES),
                     state="readonly", width=20).grid(row=0, column=1, sticky="ew")

        ttk.Label(box, text="Imagem:").grid(row=0, column=2, sticky="w", padx=(12, 6))
        self.image_format = tk.StringVar(value="PNG (mais nítido)")
        ttk.Combobox(box, textvariable=self.image_format, values=list(FORMATS),
                     state="readonly", width=20).grid(row=0, column=3, sticky="ew")

        ttk.Label(box, text="Recorte:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.crop_mode = tk.StringVar(value="Largura uniforme (recomendado)")
        ttk.Combobox(box, textvariable=self.crop_mode, values=list(CROP_MODES),
                     state="readonly").grid(row=1, column=1, columnspan=3,
                                            sticky="ew", pady=(6, 0))

        self.page_break = tk.BooleanVar(value=True)
        ttk.Checkbutton(box, text="Um print por página", variable=self.page_break
                        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))
        self.title = tk.BooleanVar(value=False)
        ttk.Checkbutton(box, text="Título com o nome do arquivo", variable=self.title
                        ).grid(row=2, column=2, columnspan=2, sticky="w", pady=(6, 0))

        ttk.Label(box, text="Modelo .docx:").grid(row=3, column=0, sticky="w",
                                                  pady=(6, 0))
        self.template = tk.StringVar()
        ttk.Entry(box, textvariable=self.template).grid(row=3, column=1, columnspan=2,
                                                        sticky="ew", pady=(6, 0))
        picker = ttk.Frame(box)
        picker.grid(row=3, column=3, sticky="e", padx=(4, 0), pady=(6, 0))
        ttk.Button(picker, text="...", width=4, command=self.choose_template
                   ).grid(row=0, column=0)
        ttk.Button(picker, text="Limpar", width=7,
                   command=lambda: self.template.set("")).grid(row=0, column=1,
                                                               padx=(4, 0))

    def _build_run_section(self, row: int) -> None:
        box = ttk.Frame(self)
        box.grid(row=row, column=0, sticky="nsew")
        box.columnconfigure(0, weight=1)
        self.rowconfigure(row, weight=1)
        box.rowconfigure(1, weight=1)

        self.progress = ttk.Progressbar(box, mode="determinate")
        self.progress.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 6))

        self.log = tk.Text(box, height=7, wrap="word", state="disabled")
        self.log.grid(row=1, column=0, columnspan=2, sticky="nsew")
        log_scroll = ttk.Scrollbar(box, orient="vertical", command=self.log.yview)
        log_scroll.grid(row=1, column=2, sticky="ns")
        self.log.configure(yscrollcommand=log_scroll.set)

        actions = ttk.Frame(box)
        actions.grid(row=2, column=0, columnspan=3, sticky="e", pady=(8, 0))
        self.open_button = ttk.Button(actions, text="Abrir pasta de saída",
                                      command=self.open_output, state="disabled")
        self.open_button.grid(row=0, column=0, padx=(0, 8))
        self.run_button = ttk.Button(actions, text="Gerar Word", command=self.run)
        self.run_button.grid(row=0, column=1)

    # ----------------------------------------------------------------- arquivos

    def add_files(self) -> None:
        chosen = filedialog.askopenfilenames(
            title="Selecione os PDFs do Robot",
            filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")])
        self._add(Path(item) for item in chosen)

    def add_folder(self) -> None:
        folder = filedialog.askdirectory(title="Selecione a pasta com os PDFs")
        if folder:
            found = sorted(Path(folder).glob("*.pdf"))
            if not found:
                messagebox.showinfo("RobotToWord", "Nenhum PDF nessa pasta.")
            self._add(found)

    def _add(self, paths) -> None:
        added = 0
        for path in paths:
            if path not in self.pdfs:
                self.pdfs.append(path)
                self.listbox.insert(tk.END, path.name)
                added += 1
        if added:
            self.write_log(f"{added} PDF(s) adicionado(s).")

    def remove_selected(self) -> None:
        for index in sorted(self.listbox.curselection(), reverse=True):
            self.listbox.delete(index)
            del self.pdfs[index]

    def clear_files(self) -> None:
        self.listbox.delete(0, tk.END)
        self.pdfs.clear()

    def choose_out_dir(self) -> None:
        folder = filedialog.askdirectory(title="Pasta onde gravar os Words")
        if folder:
            self.out_dir.set(folder)

    def choose_template(self) -> None:
        chosen = filedialog.askopenfilename(
            title="Modelo .docx", filetypes=[("Documento Word", "*.docx *.dotx")])
        if chosen:
            self.template.set(chosen)

    def _refresh_output_state(self) -> None:
        self.merge_entry.configure(state="normal" if self.merge.get() else "disabled")

    # ---------------------------------------------------------------- conversao

    def write_log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)
        self.log.configure(state="disabled")

    def run(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        if not self.pdfs:
            messagebox.showwarning("RobotToWord", "Adicione pelo menos um PDF.")
            return

        missing = [p for p in self.pdfs if not p.is_file()]
        if missing:
            messagebox.showerror(
                "RobotToWord",
                "Arquivo(s) não encontrado(s):\n" + "\n".join(str(p) for p in missing))
            return

        out_dir = Path(self.out_dir.get()) if self.out_dir.get().strip() else None
        if self.merge.get():
            name = self.merge_name.get().strip() or "Memorial de Ligações.docx"
            if not name.lower().endswith(".docx"):
                name += ".docx"
            base = out_dir or self.pdfs[0].parent
            jobs = [(base / name, list(self.pdfs))]
        else:
            jobs = [((out_dir or pdf.parent) / f"{pdf.stem}.docx", [pdf])
                    for pdf in self.pdfs]

        template_text = self.template.get().strip()
        template = Path(template_text) if template_text else None
        if template and not template.is_file():
            messagebox.showerror("RobotToWord", f"Modelo não encontrado:\n{template}")
            return

        options = Options(
            dpi=QUALITIES[self.quality.get()],
            crop_mode=CROP_MODES[self.crop_mode.get()],
            crop_threshold=245,
            crop_padding_pt=4.0,
            image_format=FORMATS[self.image_format.get()],
            jpeg_quality=85,
            page_break=self.page_break.get(),
            title=self.title.get(),
        )

        self.run_button.configure(state="disabled", text="Gerando...")
        self.open_button.configure(state="disabled")
        self.progress.configure(value=0, maximum=100)
        self.write_log("")

        self.worker = threading.Thread(target=self._work, args=(jobs, options, template),
                                       daemon=True)
        self.worker.start()

    def _work(self, jobs, options: Options, template: Path | None) -> None:
        """Roda na thread de trabalho - so fala com a interface pela fila."""
        try:
            for output, sources in jobs:
                self.messages.put(("log", f"Gerando {output.name}..."))

                def progress(done: int, total: int, message: str) -> None:
                    self.messages.put(("progress", (done, total, message)))

                prints = convert(sources, output, options, template, progress)
                self.messages.put(("log", f"   {prints} print(s) em {output}"))

            self.messages.put(("done", jobs[-1][0].parent))
        except Exception as error:  # a interface nao pode morrer com o erro
            self.messages.put(("error", f"{error}\n\n{traceback.format_exc()}"))

    def _drain_messages(self) -> None:
        """Le a fila da thread de trabalho e atualiza a interface."""
        try:
            while True:
                kind, payload = self.messages.get_nowait()
                if kind == "log":
                    self.write_log(payload)
                elif kind == "progress":
                    done, total, message = payload
                    self.progress.configure(value=done, maximum=max(total, 1))
                    self.write_log(f"   [{done}/{total}] {message}")
                elif kind == "done":
                    self.last_output_dir = payload
                    self.progress.configure(value=self.progress["maximum"])
                    self.write_log("Concluído.")
                    self.run_button.configure(state="normal", text="Gerar Word")
                    self.open_button.configure(state="normal")
                elif kind == "error":
                    self.write_log("ERRO: " + payload.splitlines()[0])
                    self.run_button.configure(state="normal", text="Gerar Word")
                    self.progress.configure(value=0)
                    messagebox.showerror("RobotToWord", payload)
        except queue.Empty:
            pass
        self.after(100, self._drain_messages)

    def open_output(self) -> None:
        if self.last_output_dir:
            open_in_explorer(self.last_output_dir)


def main() -> int:
    root = tk.Tk()
    root.title("RobotToWord - prints do Robot para o Word")
    root.minsize(720, 640)
    try:  # tema nativo no Windows; nos outros sistemas cai no padrao
        ttk.Style().theme_use("vista")
    except tk.TclError:
        pass
    Panel(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
