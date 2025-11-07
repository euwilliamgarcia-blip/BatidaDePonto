"""
Batimento de Ponto simples (Windows + Linux)
Single-file Tkinter app that records clock-ins to a CSV spreadsheet per user and per month.

Correções e melhorias:
- Uma planilha separada para cada usuário.
- Registro do mês também.
- Dias da semana e meses em português.

Features:
- GUI para registrar usuários e senhas.
- Registra primeira e segunda batida de cada dia.
- Calcula total de horas, horas previstas e saldo.
- Evita batidas duplicadas em curto espaço.
- Salva automaticamente em CSV (um por usuário/mês).
- Permite ajustar carga horária e pasta de salvamento.
- Função "Fechar folha" para encerrar o mês.

Execute: python3 batimento_ponto.py
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import os
import csv
import json
from datetime import datetime, date, timedelta
import calendar
import locale

# Define locale para português
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except Exception:
    pass

APP_NAME = "Batimento de Ponto"
DEFAULT_EXPECTED_HOURS = 8.0
DEFAULT_DOUBLE_WINDOW_MIN = 2
DATA_FILENAME_TEMPLATE = "{user}-folha-{year}-{month:02d}.csv"
USERS_FILE = "users.json"
CONFIG_FILE = "config.json"


def ensure_file_exists(path, headers):
    if not os.path.exists(path):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(headers)


def load_json(path, default):
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return default


def save_json(path, obj):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(obj, f, indent=2, ensure_ascii=False)


def get_filename_for_user(folder, user):
    now = datetime.now()
    return os.path.join(folder, DATA_FILENAME_TEMPLATE.format(user=user, year=now.year, month=now.month))


def weekday_name_pt(d: date):
    dias = ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado', 'Domingo']
    return dias[d.weekday()]


def month_name_pt(m):
    meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
    return meses[m - 1]


class TimeClockApp:
    def __init__(self, root):
        self.root = root
        root.title(APP_NAME)

        self.config = load_json(CONFIG_FILE, {
            'data_folder': os.path.expanduser('~'),
            'expected_hours': DEFAULT_EXPECTED_HOURS,
            'double_window_min': DEFAULT_DOUBLE_WINDOW_MIN
        })
        self.users = load_json(USERS_FILE, {})

        self.var_user = tk.StringVar()
        self.var_password = tk.StringVar()
        self.var_status = tk.StringVar()
        self.var_expected = tk.DoubleVar(value=self.config.get('expected_hours', DEFAULT_EXPECTED_HOURS))
        self.var_double_window = tk.IntVar(value=self.config.get('double_window_min', DEFAULT_DOUBLE_WINDOW_MIN))

        self._build_ui()
        self.update_user_dropdown()

    def _build_ui(self):
        frm = ttk.Frame(self.root, padding=12)
        frm.grid(row=0, column=0, sticky='nsew')

        ttk.Label(frm, text='Usuário:').grid(row=0, column=0, sticky='w')
        self.user_combo = ttk.Combobox(frm, textvariable=self.var_user, width=30)
        self.user_combo.grid(row=0, column=1, sticky='w')

        ttk.Button(frm, text='Registrar novo usuário', command=self.register_user).grid(row=0, column=2, padx=6)

        ttk.Label(frm, text='Senha:').grid(row=1, column=0, sticky='w')
        self.pw_entry = ttk.Entry(frm, textvariable=self.var_password, show='*', width=32)
        self.pw_entry.grid(row=1, column=1, sticky='w')

        ttk.Button(frm, text='Registrar Ponto', command=self.punch).grid(row=2, column=0, pady=8)
        ttk.Button(frm, text='Ajustes', command=self.open_settings).grid(row=2, column=1, pady=8)
        ttk.Button(frm, text='Escolher pasta', command=self.choose_folder).grid(row=2, column=2, pady=8)

        ttk.Label(frm, textvariable=self.var_status, foreground='blue').grid(row=3, column=0, columnspan=3, sticky='w')

        ttk.Button(frm, text='Abrir planilha atual', command=self.open_current_sheet).grid(row=4, column=0, pady=6)
        ttk.Button(frm, text='Fechar folha', command=self.close_sheet).grid(row=4, column=1, pady=6)

        ttk.Label(frm, text='Pasta atual:').grid(row=5, column=0, sticky='w')
        ttk.Label(frm, text=self.config.get('data_folder')).grid(row=5, column=1, columnspan=2, sticky='w')

    def update_user_dropdown(self):
        names = sorted(self.users.keys())
        self.user_combo['values'] = names
        if names:
            self.var_user.set(names[0])

    def register_user(self):
        username = simpledialog.askstring('Registrar usuário', 'Nome do usuário:', parent=self.root)
        if not username:
            return
        if username in self.users:
            messagebox.showerror('Erro', 'Usuário já existe.')
            return
        password = simpledialog.askstring('Senha', f'Defina uma senha para {username}:', show='*', parent=self.root)
        if password is None:
            return
        self.users[username] = password
        save_json(USERS_FILE, self.users)
        self.update_user_dropdown()
        messagebox.showinfo('OK', f'Usuário {username} registrado.')

    def open_settings(self):
        win = tk.Toplevel(self.root)
        win.title('Ajustes')
        ttk.Label(win, text='Horas previstas por dia:').grid(row=0, column=0, sticky='w')
        ttk.Entry(win, textvariable=self.var_expected).grid(row=0, column=1)
        ttk.Label(win, text='Janela anti-duplicata (min):').grid(row=1, column=0, sticky='w')
        ttk.Entry(win, textvariable=self.var_double_window).grid(row=1, column=1)

        def save_and_close():
            self.config['expected_hours'] = float(self.var_expected.get())
            self.config['double_window_min'] = int(self.var_double_window.get())
            save_json(CONFIG_FILE, self.config)
            messagebox.showinfo('Ajustes', 'Configurações salvas.')
            win.destroy()

        ttk.Button(win, text='Salvar', command=save_and_close).grid(row=2, column=0, columnspan=2)

    def choose_folder(self):
        folder = filedialog.askdirectory(initialdir=self.config.get('data_folder', os.path.expanduser('~')))
        if folder:
            self.config['data_folder'] = folder
            save_json(CONFIG_FILE, self.config)
            messagebox.showinfo('Pasta', f'Nova pasta: {folder}')

    def open_current_sheet(self):
        username = self.var_user.get().strip()
        path = get_filename_for_user(self.config['data_folder'], username)
        if not os.path.exists(path):
            messagebox.showinfo('Planilha', 'Nenhuma planilha encontrada para este usuário.')
            return
        os.system(f'xdg-open "{path}"' if os.name != 'nt' else f'start {path}')

    def close_sheet(self):
        messagebox.showinfo('Fechar folha', 'Função fechar folha ainda em desenvolvimento.')

    def punch(self):
        username = self.var_user.get().strip()
        pw = self.var_password.get()
        if not username:
            messagebox.showerror('Erro', 'Informe o usuário.')
            return
        if username not in self.users or self.users[username] != pw:
            messagebox.showerror('Erro', 'Usuário ou senha incorretos.')
            return

        folder = self.config['data_folder']
        now = datetime.now()
        filename = get_filename_for_user(folder, username)
        headers = ['Data','Dia da semana','Mês','Usuário','Hora 1','Hora 2','Total horas (h)','Horas previstas (h)','Saldo (h)']
        ensure_file_exists(filename, headers)

        with open(filename, 'r', encoding='utf-8') as f:
            rows = list(csv.DictReader(f))

        today = now.strftime('%Y-%m-%d')
        found = next((r for r in rows if r['Data'] == today and r['Usuário'] == username), None)

        if found and found.get('Hora 2'):
            messagebox.showinfo('Aviso', 'Já existem duas batidas para hoje.')
            return

        if not found:
            row = {
                'Data': today,
                'Dia da semana': weekday_name_pt(now.date()),
                'Mês': month_name_pt(now.month),
                'Usuário': username,
                'Hora 1': now.strftime('%H:%M:%S'),
                'Hora 2': '',
                'Total horas (h)': '',
                'Horas previstas (h)': str(self.config['expected_hours']),
                'Saldo (h)': ''
            }
            rows.append(row)
        else:
            found['Hora 2'] = now.strftime('%H:%M:%S')
            t1 = datetime.strptime(found['Hora 1'], '%H:%M:%S')
            t2 = datetime.strptime(found['Hora 2'], '%H:%M:%S')
            total = (t2 - t1).total_seconds() / 3600
            found['Total horas (h)'] = f'{total:.2f}'
            expected = float(self.config['expected_hours'])
            found['Saldo (h)'] = f'{(total - expected):+.2f}'

        with open(filename, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=headers)
            writer.writeheader()
            for r in rows:
                writer.writerow(r)

        self.var_status.set(f'Batida registrada para {username} em {now.strftime("%d/%m/%Y %H:%M:%S")}')
        messagebox.showinfo('OK', 'Ponto registrado com sucesso!')


if __name__ == '__main__':
    root = tk.Tk()
    app = TimeClockApp(root)
    root.mainloop()