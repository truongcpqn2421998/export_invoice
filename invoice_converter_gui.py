import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import shutil
import subprocess
import sys

DEFAULT_MAPPING_FILENAME = 'mapping_template.yaml'
DEFAULT_MAPPING_CONTENT = '''field_mappings:
  ThTien: Amount
  ThanhTien: Amount
  DGia: UnitPrice
  SLuong: Quantity
  MHHDVu: Code
  THHDVu: Description
  DVTinh: Unit
  STT: LineNo
  TSuat: TaxRate
  ThueSuat: TaxRate
  TienThue: TaxAmount
  CKhau: DiscountAmount
  TongTien: TotalAmount
  TgTTTBSo: TotalAmountNumber
  TgTTTBChu: TotalAmountWords
  TgThue: TotalTaxAmount
  TgTT: TotalBeforeTax
  MaDiaDiemKD: BusinessLocationCode
  NguoiLap: IssuerName
  NguoiKy: SignerName
'''

try:
    from convert_xml_invoices_to_excel import convert_dir_to_excels
except Exception:
    convert_dir_to_excels = None


def get_resource_dir():
    if getattr(sys, 'frozen', False):
        return getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def get_app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def ensure_default_mapping_exists(path):
    if not os.path.exists(path):
        with open(path, 'w', encoding='utf-8') as mf:
            mf.write(DEFAULT_MAPPING_CONTENT)


def copy_default_mapping_to_app_dir():
    src = os.path.join(get_resource_dir(), DEFAULT_MAPPING_FILENAME)
    dst = os.path.join(get_app_dir(), DEFAULT_MAPPING_FILENAME)
    if os.path.abspath(src) != os.path.abspath(dst):
        if os.path.exists(src) and not os.path.exists(dst):
            try:
                shutil.copy(src, dst)
            except Exception:
                ensure_default_mapping_exists(dst)
    else:
        ensure_default_mapping_exists(dst)
    return dst


class App:
    def __init__(self, root):
        self.root = root
        root.title('Invoice XML -> Excel')
        root.resizable(False, False)

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.mapping_var = tk.StringVar()

        tk.Label(root, text='Input folder:').grid(row=0, column=0, sticky='w', padx=8, pady=6)
        tk.Entry(root, textvariable=self.input_var, width=56).grid(row=0, column=1, padx=8)
        tk.Button(root, text='Browse', command=self.browse_input).grid(row=0, column=2, padx=8)

        tk.Label(root, text='Output folder (optional):').grid(row=1, column=0, sticky='w', padx=8, pady=6)
        tk.Entry(root, textvariable=self.output_var, width=56).grid(row=1, column=1, padx=8)
        tk.Button(root, text='Browse', command=self.browse_output).grid(row=1, column=2, padx=8)

        tk.Label(root, text='Mapping (YAML optional; default mapping is built-in):').grid(row=2, column=0, sticky='w', padx=8, pady=6)
        tk.Entry(root, textvariable=self.mapping_var, width=56).grid(row=2, column=1, padx=8)
        tk.Button(root, text='Browse', command=self.browse_mapping).grid(row=2, column=2, padx=8)

        tk.Button(root, text='Edit mapping', command=self.edit_mapping).grid(row=3, column=1, sticky='w', padx=8)

        self.apply_btn = tk.Button(root, text='Apply', width=12, command=self.apply)
        self.apply_btn.grid(row=5, column=1, pady=12)

        self.status = tk.Label(root, text='Ready', anchor='w')
        self.status.grid(row=6, column=0, columnspan=3, sticky='we', padx=8, pady=(0,8))

        self.default_mapping_path = copy_default_mapping_to_app_dir()
        try:
            self.mapping_var.set(self.default_mapping_path)
        except Exception:
            pass

        if convert_dir_to_excels is None:
            messagebox.showwarning('Warning', 'Cannot import converter script. Ensure convert_xml_invoices_to_excel.py is in the same folder.')
            self.apply_btn.config(state='disabled')

    def browse_input(self):
        d = filedialog.askdirectory()
        if d:
            self.input_var.set(d)
            if not self.output_var.get():
                self.output_var.set(d)

    def browse_output(self):
        d = filedialog.askdirectory()
        if d:
            self.output_var.set(d)

    def browse_mapping(self):
        f = filedialog.askopenfilename(filetypes=[('YAML files', '*.yml;*.yaml'), ('All files', '*.*')])
        if f:
            self.mapping_var.set(f)

    def edit_mapping(self):
        path = self.mapping_var.get() or self.default_mapping_path
        try:
            ensure_default_mapping_exists(path)
        except Exception as e:
            messagebox.showerror('Error', f'Cannot create mapping template:\n{e}')
            return
        self.mapping_var.set(path)
        # open in default editor (Notepad)
        try:
            if os.name == 'nt':
                os.startfile(path)
            else:
                subprocess.Popen(['xdg-open', path])
        except Exception as e:
            messagebox.showerror('Error', f'Cannot open mapping file:\n{e}')

    def apply(self):
        inp = self.input_var.get().strip()
        out = self.output_var.get().strip() or None
        mapping = self.mapping_var.get().strip() or None
        if not inp or not os.path.isdir(inp):
            messagebox.showerror('Error', 'Please choose a valid input folder')
            return
        self.apply_btn.config(state='disabled')
        self.status.config(text='Processing...')
        t = threading.Thread(target=self.run_conversion, args=(inp, out, mapping), daemon=True)
        t.start()

    def run_conversion(self, inp, out, mapping):
        try:
            convert_dir_to_excels(inp, out, mapping)
            self.status.config(text='Done')
            messagebox.showinfo('Done', 'Conversion finished. Check output folder.')
        except Exception as e:
            self.status.config(text='Error')
            messagebox.showerror('Error', f'Conversion failed:\n{e}')
        finally:
            self.apply_btn.config(state='normal')


if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()
