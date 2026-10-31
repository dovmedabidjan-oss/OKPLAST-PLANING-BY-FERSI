import sys
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

class ApplicationAdaptationMoules:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestion & Adaptation Moules / Machines (2025-2026)")
        self.root.geometry("1200x720")
        self.root.minsize(980, 600)
        
        self.df = None
        self.setup_styles()
        self.create_widgets()
        
        default_file = 'liste changement2025-2026.xlsx'
        if os.path.exists(default_file):
            self.load_excel_data(default_file)

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TFrame', background='#F8FAFC')
        style.configure('TLabelframe', background='#F8FAFC', font=('Segoe UI', 10, 'bold'))
        style.configure('TLabelframe.Label', background='#F8FAFC', foreground='#1E293B')
        style.configure('Treeview', font=('Segoe UI', 10), rowheight=28)
        style.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'), background='#0284C7', foreground='white')
        style.map('Treeview.Heading', background=[('active', '#0369A1')])

    def create_widgets(self):
        header = tk.Frame(self.root, bg='#0284C7', height=70)
        header.pack(fill=tk.X, side=tk.TOP)
        
        title_label = tk.Label(header, text="🏭 ADAPTATION MOULES & MACHINES", 
                               font=('Segoe UI', 14, 'bold'), fg='white', bg='#0284C7')
        title_label.pack(side=tk.LEFT, padx=20, pady=15)
        
        btn_frame = tk.Frame(header, bg='#0284C7')
        btn_frame.pack(side=tk.RIGHT, padx=15, pady=15)

        btn_import = tk.Button(btn_frame, text="📁 Charger Excel", command=self.import_file,
                               font=('Segoe UI', 9, 'bold'), bg='#0369A1', fg='white', 
                               activebackground='#075985', activeforeground='white', bd=0, padx=12, pady=6, cursor='hand2')
        btn_import.pack(side=tk.LEFT, padx=5)

        btn_save = tk.Button(btn_frame, text="💾 Enregistrer sous", command=self.save_as,
                             font=('Segoe UI', 9, 'bold'), bg='#0D9488', fg='white', 
                             activebackground='#0F766E', activeforeground='white', bd=0, padx=12, pady=6, cursor='hand2')
        btn_save.pack(side=tk.LEFT, padx=5)

        btn_print = tk.Button(btn_frame, text="🖨️ Imprimer", command=self.print_results,
                              font=('Segoe UI', 9, 'bold'), bg='#475569', fg='white', 
                              activebackground='#334155', activeforeground='white', bd=0, padx=12, pady=6, cursor='hand2')
        btn_print.pack(side=tk.LEFT, padx=5)

        main_frame = tk.Frame(self.root, bg='#F8FAFC')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        search_frame = tk.LabelFrame(main_frame, text=" 🔍 Moteur de recherche (3 Critères) ", bg='#F8FAFC', fg='#0284C7')
        search_frame.pack(fill=tk.X, side=tk.TOP, pady=(0, 15), ipadx=10, ipady=10)

        tk.Label(search_frame, text="Rechercher par :", font=('Segoe UI', 10, 'bold'), bg='#F8FAFC', fg='#334155').grid(row=0, column=0, padx=10, pady=5, sticky='w')
        
        self.search_type = tk.StringVar(value="MOUL")
        rb1 = tk.Radiobutton(search_frame, text="1. Code Moule (OKPLT / MO...)", variable=self.search_type, value="MOUL",
                             bg='#F8FAFC', font=('Segoe UI', 9), command=self.perform_search)
        rb1.grid(row=0, column=1, padx=10, pady=5, sticky='w')
        
        rb2 = tk.Radiobutton(search_frame, text="2. Nom de produit (Désignation)", variable=self.search_type, value="PROD",
                             bg='#F8FAFC', font=('Segoe UI', 9), command=self.perform_search)
        rb2.grid(row=0, column=2, padx=10, pady=5, sticky='w')

        rb3 = tk.Radiobutton(search_frame, text="3. Numéro de Machine (ex: M07, M28)", variable=self.search_type, value="MACH",
                             bg='#F8FAFC', font=('Segoe UI', 9), command=self.perform_search)
        rb3.grid(row=0, column=3, padx=10, pady=5, sticky='w')

        tk.Label(search_frame, text="Saisie :", font=('Segoe UI', 10, 'bold'), bg='#F8FAFC', fg='#334155').grid(row=1, column=0, padx=10, pady=10, sticky='w')
        
        self.entry_search = tk.Entry(search_frame, font=('Segoe UI', 11), width=45, bd=1, relief=tk.SOLID)
        self.entry_search.grid(row=1, column=1, columnspan=2, padx=10, pady=10, sticky='w')
        self.entry_search.bind('<KeyRelease>', lambda e: self.perform_search())

        btn_reset = tk.Button(search_frame, text="Réinitialiser", command=self.reset_search,
                              font=('Segoe UI', 9), bg='#64748B', fg='white', bd=0, padx=12, pady=4, cursor='hand2')
        btn_reset.grid(row=1, column=3, padx=10, pady=10, sticky='w')

        self.lbl_stats = tk.Label(main_frame, text="Statut : Chargement des données...", font=('Segoe UI', 9, 'italic'), fg='#64748B', bg='#F8FAFC')
        self.lbl_stats.pack(fill=tk.X, side=tk.TOP, pady=(0, 5))

        hint_label = tk.Label(main_frame, text="💡 Conseil : Double-cliquez sur une ligne pour voir le détail des dates de montage et de démontage.", 
                              font=('Segoe UI', 9, 'bold'), fg='#0284C7', bg='#F8FAFC')
        hint_label.pack(fill=tk.X, side=tk.TOP, pady=(0, 10))

        tree_frame = tk.Frame(main_frame, bg='#F8FAFC')
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = ('col1', 'col2', 'col3', 'col4', 'col5')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', selectmode='browse')
        
        self.tree.heading('col1', text='Code Moule / Machine')
        self.tree.heading('col2', text='Désignation Produit')
        self.tree.heading('col3', text='Machine / Moule Associé')
        self.tree.heading('col4', text='Nb de fois (Montage + Démontage)')
        self.tree.heading('col5', text='Adaptation / Recommandation')

        self.tree.column('col1', width=130, anchor='center')
        self.tree.column('col2', width=300, anchor='w')
        self.tree.column('col3', width=160, anchor='center')
        self.tree.column('col4', width=200, anchor='center')
        self.tree.column('col5', width=180, anchor='center')

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<Double-1>", self.show_date_details)

    def import_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx *.xls")])
        if file_path:
            self.load_excel_data(file_path)

    def load_excel_data(self, file_path):
        try:
            df_raw = pd.read_excel(file_path, header=None)
            df_raw.columns = ['DATE', 'OKPLT', 'PROG', 'DEPROG', 'MACH']
            
            df_raw['DATE_STR'] = pd.to_datetime(df_raw['DATE'], errors='coerce').dt.strftime('%d/%m/%Y').fillna(str(df_raw['DATE']))
            df_raw['DESIGNATION'] = df_raw['PROG'].combine_first(df_raw['DEPROG']).fillna('N/A')
            df_raw['OKPLT'] = df_raw['OKPLT'].astype(str).str.strip()
            df_raw['MACH'] = df_raw['MACH'].astype(str).str.strip()
            df_raw['DESIGNATION'] = df_raw['DESIGNATION'].astype(str).str.strip()
            
            self.df = df_raw[~df_raw['OKPLT'].str.contains('AUCUN|JAMAIS|DATE|PAS!', case=False, na=False)].copy()
            
            nb_m = self.df['OKPLT'].nunique()
            nb_mac = self.df['MACH'].nunique()
            nb_total = len(self.df)
            
            self.lbl_stats.config(text=f"Base active : {nb_total} enregistrements | {nb_m} Moules uniques | {nb_mac} Machines uniques")
            self.perform_search()
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur de lecture du fichier Excel : {str(e)}")

    def reset_search(self):
        self.entry_search.delete(0, tk.END)
        self.perform_search()

    def perform_search(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        if self.df is None:
            return

        term = self.entry_search.get().strip().upper()
        mode = self.search_type.get()

        if mode == "MOUL":
            filtered = self.df[self.df['OKPLT'].str.upper().str.contains(term, na=False)]
            grouped = filtered.groupby(['OKPLT', 'DESIGNATION', 'MACH']).size().reset_index(name='raw_count')
            
            for code in grouped['OKPLT'].unique():
                sub = grouped[grouped['OKPLT'] == code].sort_values(by='raw_count', ascending=False)
                total_raw = sub['raw_count'].sum()
                for idx, (_, row) in enumerate(sub.iterrows()):
                    nb_fois = max(1, row['raw_count'] // 2)
                    recom = "⭐ Choix Optimal" if idx == 0 else ("✔️ Alternative" if idx == 1 else "🔹 Occasionnel")
                    pct = int((row['raw_count'] / max(1, total_raw)) * 100)
                    self.tree.insert('', tk.END, values=(
                        row['OKPLT'], row['DESIGNATION'], f"Machine {row['MACH']}", 
                        f"{nb_fois} fois", recom
                    ))

        elif mode == "PROD":
            filtered = self.df[self.df['DESIGNATION'].str.upper().str.contains(term, na=False)]
            grouped = filtered.groupby(['OKPLT', 'DESIGNATION', 'MACH']).size().reset_index(name='raw_count')
            
            for code in grouped['OKPLT'].unique():
                sub = grouped[grouped['OKPLT'] == code].sort_values(by='raw_count', ascending=False)
                total_raw = sub['raw_count'].sum()
                for idx, (_, row) in enumerate(sub.iterrows()):
                    nb_fois = max(1, row['raw_count'] // 2)
                    recom = "⭐ Choix Optimal" if idx == 0 else "Alternative"
                    self.tree.insert('', tk.END, values=(
                        row['OKPLT'], row['DESIGNATION'], f"Machine {row['MACH']}", 
                        f"{nb_fois} fois", recom
                    ))

        elif mode == "MACH":
            filtered = self.df[self.df['MACH'].str.upper().str.contains(term, na=False)]
            grouped = filtered.groupby(['MACH', 'OKPLT', 'DESIGNATION']).size().reset_index(name='raw_count')
            
            for mach in grouped['MACH'].unique():
                sub = grouped[grouped['MACH'] == mach].sort_values(by='raw_count', ascending=False)
                for _, row in sub.iterrows():
                    nb_fois = max(1, row['raw_count'] // 2)
                    self.tree.insert('', tk.END, values=(
                        f"Machine {row['MACH']}", row['DESIGNATION'], row['OKPLT'], 
                        f"{nb_fois} fois", "Moule Compatible"
                    ))

    def show_date_details(self, event):
        selected_item = self.tree.selection()
        if not selected_item or self.df is None:
            return

        item_values = self.tree.item(selected_item[0])['values']
        col1 = str(item_values[0]).replace('Machine ', '')
        col3 = str(item_values[2]).replace('Machine ', '')
        
        if col1.startswith('MO'):
            moule_code = col1
            machine_code = col3
        elif col3.startswith('MO'):
            moule_code = col3
            machine_code = col1
        else:
            moule_code = col1
            machine_code = col3

        events = self.df[
            (self.df['OKPLT'].str.upper() == moule_code.upper()) & 
            (self.df['MACH'].str.upper() == machine_code.upper())
        ].copy()

        if events.empty:
            events = self.df[self.df['OKPLT'].str.upper() == moule_code.upper()].copy()

        top = tk.Toplevel(self.root)
        top.title(f"📅 Détail Montage / Démontage — {moule_code} / Machine {machine_code}")
        top.geometry("900x480")
        top.configure(bg='#F8FAFC')

        lbl_top = tk.Label(top, text=f"Suivi des opérations : Moule {moule_code} sur Machine {machine_code}",
                           font=('Segoe UI', 11, 'bold'), bg='#0284C7', fg='white', pady=10)
        lbl_top.pack(fill=tk.X)

        detail_frame = tk.Frame(top, bg='#F8FAFC')
        detail_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        cols = ('d_date', 'd_prog', 'd_deprog', 'd_prod', 'd_mach')
        dtree = ttk.Treeview(detail_frame, columns=cols, show='headings', height=12)
        
        dtree.heading('d_date', text='Date')
        dtree.heading('d_prog', text='Montage (PROG)')
        dtree.heading('d_deprog', text='Démontage (DEPROG)')
        dtree.heading('d_prod', text='Désignation Produit')
        dtree.heading('d_mach', text='Machine')

        dtree.column('d_date', width=110, anchor='center')
        dtree.column('d_prog', width=180, anchor='center')
        dtree.column('d_deprog', width=180, anchor='center')
        dtree.column('d_prod', width=260, anchor='w')
        dtree.column('d_mach', width=90, anchor='center')

        sb = ttk.Scrollbar(detail_frame, orient=tk.VERTICAL, command=dtree.yview)
        dtree.configure(yscroll=sb.set)

        dtree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        for _, row in events.iterrows():
            prog_val = str(row['PROG']) if pd.notna(row['PROG']) and str(row['PROG']).strip() != 'nan' and str(row['PROG']).strip() != '' else "—"
            deprog_val = str(row['DEPROG']) if pd.notna(row['DEPROG']) and str(row['DEPROG']).strip() != 'nan' and str(row['DEPROG']).strip() != '' else "—"
            
            dtree.insert('', tk.END, values=(
                row['DATE_STR'],
                f"🟢 {prog_val}" if prog_val != "—" else "—",
                f"🔴 {deprog_val}" if deprog_val != "—" else "—",
                row['DESIGNATION'],
                f"Machine {row['MACH']}"
            ))

        btn_close = tk.Button(top, text="Fermer", command=top.destroy, font=('Segoe UI', 9, 'bold'),
                              bg='#64748B', fg='white', bd=0, padx=15, pady=6, cursor='hand2')
        btn_close.pack(pady=10)

    def get_tree_data(self):
        rows = []
        for child in self.tree.get_children():
            rows.append(self.tree.item(child)['values'])
        return pd.DataFrame(rows, columns=['Code / Machine', 'Désignation Produit', 'Association', 'Nb de fois', 'Recommandation'])

    def save_as(self):
        df_export = self.get_tree_data()
        if df_export.empty:
            messagebox.showwarning("Attention", "Aucun résultat à enregistrer.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV (séparateur virgule)", "*.csv")]
        )
        if file_path:
            try:
                if file_path.endswith('.csv'):
                    df_export.to_csv(file_path, index=False, encoding='utf-8-sig')
                else:
                    df_export.to_excel(file_path, index=False)
                messagebox.showinfo("Succès", f"Fichier enregistré avec succès :\n{file_path}")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de l'enregistrement : {str(e)}")

    def print_results(self):
        df_export = self.get_tree_data()
        if df_export.empty:
            messagebox.showwarning("Attention", "Aucun résultat à imprimer.")
            return

        temp_file = "impression_resultats_moules.txt"
        with open(temp_file, "w", encoding="utf-8") as f:
            f.write("=== RAPPORT D'ADAPTATION MOULES / MACHINES ===\n\n")
            f.write(df_export.to_string(index=False))
        
        try:
            if sys.platform == "win32":
                os.startfile(temp_file, "print")
            else:
                os.system(f"lpr {temp_file}")
            messagebox.showinfo("Impression", "Le document a été envoyé à l'imprimante.")
        except Exception as e:
            os.system(f"notepad {temp_file}")

if __name__ == '__main__':
    root = tk.Tk()
    app = ApplicationAdaptationMoules(root)
    root.mainloop()