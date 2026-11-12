
import os
import sys
import glob
import threading
import subprocess
import traceback
from datetime import datetime

if getattr(sys, 'frozen', False):
    BASE = os.path.dirname(sys.executable)
else:
    BASE = os.path.dirname(os.path.abspath(__file__))

DADOS_PATH = os.path.join(BASE, 'dados.xlsx')
MODELOS_DIR = BASE
OUTPUT_DIR = os.path.join(BASE, 'impressoes')
LOG_PATH = os.path.join(BASE, 'erro.log')


def erro_fatal(titulo, detalhe, tecnico=''):
    try:
        with open(LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(f'[{datetime.now():%d/%m/%Y %H:%M:%S}] {titulo}\n{detalhe}\n{tecnico}\n\n')
    except Exception:
        pass
    print(f'{titulo}\n{detalhe}\n{tecnico}', file=sys.stderr)
    try:
        import tkinter as _tk
        from tkinter import messagebox as _mb
        root = _tk.Tk()
        root.withdraw()
        _mb.showerror(titulo, f'{detalhe}\n\n(detalhes técnicos salvos em erro.log)')
        root.destroy()
    except Exception:
        pass
    sys.exit(1)


try:
    import pandas as pd
    import tkinter as tk
    from tkinter import ttk, messagebox
    from docxtpl import DocxTemplate
    from docx2pdf import convert
except ImportError as e:
    erro_fatal(
        'Falta instalar uma biblioteca',
        f'{e}\n\nPython em uso:\n{sys.executable}\n\n'
        'Instale com:\n"{}" -m pip install pandas openpyxl docxtpl docx2pdf'.format(sys.executable)
    )

COLUNAS_OBRIGATORIAS = ['Artigo', 'Desc. Artigo', 'De:', 'Por:', 'CI Anterior']

MODELO_PROMO = 'cartaz_promo.docx'
MODELO_CHEIO = 'cartaz_cheio.docx'
CAMPOS_OBRIGATORIOS = {'Desc_Artigo', 'Artigo', 'Por'}


def carregar_produtos():
    if not os.path.exists(DADOS_PATH):
        raise FileNotFoundError(f'Não encontrei o dados.xlsx nesta pasta:\n{BASE}')

    ci = pd.read_excel(DADOS_PATH, header=4, usecols='G:K')

    faltando = [c for c in COLUNAS_OBRIGATORIAS if c not in ci.columns]
    if faltando:
        raise ValueError(
            f'A planilha (1ª aba) não está no layout esperado.\n'
            f'Colunas que faltam: {faltando}\n'
            f'Colunas lidas (G:K, cabeçalho na linha 5): {list(ci.columns)}'
        )

    filtrado = ci[ci['CI Anterior'] != 1]
    filtrado = filtrado[~filtrado['Desc. Artigo'].str.contains('220V', na=False, case=False)]
    filtrado = filtrado.dropna(subset=['Desc. Artigo', 'Artigo'])
    return filtrado.reset_index(drop=True)


def tipo_cartaz(row):
    de, por, ci = row['De:'], row['Por:'], row['CI Anterior']

    if pd.notna(de) and pd.notna(por):
        if por < de:
            return 'promo'
        if por > de:
            return 'cheio'

    if pd.notna(ci):
        if ci < 1:
            return 'promo'
        if ci > 1:
            return 'cheio'
    return None


def detectar_modelos():
    modelos = []
    for f in glob.glob(os.path.join(MODELOS_DIR, '*.docx')):
        nome = os.path.basename(f)
        if not nome.startswith('~') and 'temp_print' not in nome:
            modelos.append(nome)
    return sorted(modelos)


def gerar_e_imprimir(produtos, modelos_escolhidos, btn, status_label):
    btn.config(state='disabled')
    status_label.config(text='⏳ Gerando PDFs...')

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    erros = []
    pdfs_gerados = []
    modelos_ruins = set()

    for i, (idx, row) in enumerate(produtos.iterrows()):
        modelo_nome = modelos_escolhidos[i].get()
        if modelo_nome == '-- Pular --' or modelo_nome in modelos_ruins:
            continue

        modelo_path = os.path.join(MODELOS_DIR, modelo_nome)
        try:
            doc = DocxTemplate(modelo_path)

            faltam = CAMPOS_OBRIGATORIOS - doc.get_undeclared_template_variables()
            if faltam:
                modelos_ruins.add(modelo_nome)
                nomes = ', '.join('{{' + c + '}}' for c in sorted(faltam))
                erros.append(f"Modelo '{modelo_nome}' não tem os campos: {nomes}. "
                             f"Nenhum cartaz foi gerado com ele.")
                continue

            doc.render({
                'Desc_Artigo': str(row['Desc. Artigo']),
                'Artigo': str(row['Artigo']),
                'De': f"R$ {row['De:']:.2f}".replace('.', ',') if pd.notna(row['De:']) else '',
                'Por': f"R$ {row['Por:']:.2f}".replace('.', ','),
            })
            nome_docx = os.path.join(OUTPUT_DIR, f"cartaz_{i+1:03d}_{int(row['Artigo'])}.docx")
            doc.save(nome_docx)

            nome_pdf = nome_docx.replace('.docx', '.pdf')
            convert(nome_docx, nome_pdf)
            if not os.path.exists(nome_pdf):
                raise RuntimeError('o PDF não foi criado (a conversão falhou)')
            os.remove(nome_docx)
            pdfs_gerados.append(nome_pdf)

        except Exception as e:
            erros.append(f"Produto {row['Artigo']}: {e}")

    status_label.config(text=f'🖨️ Imprimindo {len(pdfs_gerados)} cartaz(es)...')
    for pdf in pdfs_gerados:
        try:
            if sys.platform == 'win32':
                os.startfile(pdf, 'print')
            else:
                subprocess.run(['lp', pdf], check=True)
        except Exception as e:
            erros.append(f"Erro ao imprimir {os.path.basename(pdf)}: {e}")

    if erros:
        messagebox.showerror('Erros encontrados', '\n'.join(erros))

    status_label.config(
        text=f'✅ Concluído! {len(pdfs_gerados)} cartaz(es) enviados para impressão.'
    )
    btn.config(state='normal')


class AutoImpressApp(tk.Tk):
    def __init__(self, produtos):
        super().__init__()
        self.title('AutoImpress')
        self.resizable(True, True)
        self.configure(bg='#1e1e2e')
        self._build(produtos)

    def report_callback_exception(self, exc, val, tb):
        detalhe = ''.join(traceback.format_exception(exc, val, tb))
        try:
            with open(LOG_PATH, 'a', encoding='utf-8') as f:
                f.write(f'[{datetime.now():%d/%m/%Y %H:%M:%S}] erro na interface\n{detalhe}\n')
        except Exception:
            pass
        messagebox.showerror('Erro', f'{val}\n\n(detalhes salvos em erro.log)')

    def _build(self, produtos):
        BG = '#1e1e2e'
        CARD = '#2a2a3e'
        ACCENT = '#00dc82'
        FG = '#e0e0f0'
        FG2 = '#8888aa'
        FONT = ('Segoe UI', 10)
        FONT_B = ('Segoe UI', 10, 'bold')
        FONT_H = ('Segoe UI', 13, 'bold')

        header = tk.Frame(self, bg=BG, pady=16)
        header.pack(fill='x', padx=24)
        tk.Label(header, text='AutoImpress', font=('Segoe UI', 18, 'bold'),
                 bg=BG, fg=ACCENT).pack(side='left')
        self.status_label = tk.Label(header, text='', font=FONT,
                                     bg=BG, fg=FG2)
        self.status_label.pack(side='right', padx=8)

        self.produtos = produtos

        self.modelos_disponiveis = detectar_modelos()
        opcoes = ['-- Pular --'] + self.modelos_disponiveis

        tipos = [tipo_cartaz(row) for _, row in self.produtos.iterrows()]
        n_promo = tipos.count('promo')
        n_cheio = tipos.count('cheio')

        info = tk.Label(self,
            text=(f'{len(self.produtos)} produto(s) com mudança de preço esta semana  •  '
                  f'{n_promo} promo (baixou)  •  {n_cheio} cheio (subiu)'),
            font=FONT, bg=BG, fg=FG2)
        info.pack(padx=24, anchor='w', pady=(0, 8))

        frame_tbl = tk.Frame(self, bg=BG)
        frame_tbl.pack(fill='both', expand=True, padx=24)

        canvas = tk.Canvas(frame_tbl, bg=BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame_tbl, orient='vertical', command=canvas.yview)
        self.inner = tk.Frame(canvas, bg=BG)

        self.inner.bind('<Configure>',
            lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=self.inner, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        headers = ['#', 'Cód.', 'Descrição', 'De', 'Por', 'Tipo', 'Modelo']
        widths = [3, 7, 42, 10, 10, 9, 24]
        for col, (h, w) in enumerate(zip(headers, widths)):
            tk.Label(self.inner, text=h, font=FONT_B, bg=CARD, fg=ACCENT,
                     width=w, anchor='w', padx=6, pady=6,
                     relief='flat').grid(row=0, column=col, sticky='ew',
                                        padx=1, pady=1)

        self.modelos_escolhidos = []
        modelos_faltando = set()
        for i, row in self.produtos.iterrows():
            bg_row = BG if i % 2 == 0 else CARD
            tipo = tipos[i]
            de_val = f"R$ {row['De:']:.2f}".replace('.', ',') if pd.notna(row['De:']) else '—'
            por_val = f"R$ {row['Por:']:.2f}".replace('.', ',')

            tipo_txt = {'promo': '↓ Promo', 'cheio': '↑ Cheio'}.get(tipo, '?')
            vals = [i+1, int(row['Artigo']),
                    str(row['Desc. Artigo'])[:45],
                    de_val, por_val, tipo_txt]
            widths_row = [3, 7, 42, 10, 10, 9]

            for col, (v, w) in enumerate(zip(vals, widths_row)):
                tk.Label(self.inner, text=str(v), font=FONT, bg=bg_row,
                         fg=FG, width=w, anchor='w', padx=6, pady=5
                         ).grid(row=i+1, column=col, sticky='ew', padx=1, pady=0)

            alvo = {'promo': MODELO_PROMO, 'cheio': MODELO_CHEIO}.get(tipo)
            if alvo in self.modelos_disponiveis:
                padrao = alvo
            else:
                padrao = '-- Pular --'
                if alvo:
                    modelos_faltando.add(alvo)

            var = tk.StringVar(value=padrao)
            cb = ttk.Combobox(self.inner, textvariable=var,
                              values=opcoes, width=24, state='readonly',
                              font=FONT)
            cb.grid(row=i+1, column=6, sticky='ew', padx=4, pady=2)
            self.modelos_escolhidos.append(var)

        if modelos_faltando:
            self.after(300, lambda: messagebox.showwarning(
                'Modelo não encontrado',
                'Não encontrei na pasta: ' + ', '.join(sorted(modelos_faltando)) + '\n\n'
                'Os produtos que usariam esse modelo ficaram como "-- Pular --".\n\n'
                'O programa só lê .docx. Se o modelo está em .doc, abra no Word e use '
                'Salvar como -> .docx (mantendo o mesmo nome).'))

        footer = tk.Frame(self, bg=BG, pady=16)
        footer.pack(fill='x', padx=24)

        self.btn_imprimir = tk.Button(
            footer, text='🖨️  Gerar e Imprimir todos',
            font=FONT_B, bg=ACCENT, fg='#0a1a10',
            activebackground='#00b86a', relief='flat',
            padx=20, pady=8, cursor='hand2',
            command=self._iniciar_impressao)
        self.btn_imprimir.pack(side='right')

        tk.Button(footer, text='↺  Recarregar planilha',
                  font=FONT, bg=CARD, fg=FG2,
                  relief='flat', padx=12, pady=8, cursor='hand2',
                  command=self._recarregar).pack(side='right', padx=8)

    def _iniciar_impressao(self):
        ativos = [v for v in self.modelos_escolhidos if v.get() != '-- Pular --']
        if not ativos:
            messagebox.showwarning('Atenção', 'Todos os produtos estão marcados como "Pular".')
            return

        threading.Thread(
            target=gerar_e_imprimir,
            args=(self.produtos, self.modelos_escolhidos,
                  self.btn_imprimir, self.status_label),
            daemon=True
        ).start()

    def _recarregar(self):
        try:
            produtos = carregar_produtos()
        except Exception as e:
            messagebox.showerror('Erro', f'Não foi possível ler dados.xlsx:\n{e}')
            return
        for widget in self.winfo_children():
            widget.destroy()
        self._build(produtos)


if __name__ == '__main__':
    try:
        produtos = carregar_produtos()
        app = AutoImpressApp(produtos)
        app.geometry('1040x560')
    except Exception as e:
        erro_fatal('Não foi possível abrir o AutoImpress', str(e),
                   traceback.format_exc())
    app.mainloop()
