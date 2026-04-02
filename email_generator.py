#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Email Generator για Ακίνητα
Εφαρμογή που διαβάζει Excel από Google Forms και δημιουργεί έτοιμα emails
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import re


class EmailGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Generator - Ακίνητα")
        self.root.geometry("1000x800")

        # Δεδομένα
        self.df = None  # DataFrame από το Excel
        self.selected_client = None  # Επιλεγμένος πελάτης

        # Δημιουργία GUI
        self.create_widgets()

    def create_widgets(self):
        """Δημιουργία όλων των στοιχείων του GUI"""

        # Frame για το κουμπί φόρτωσης
        top_frame = tk.Frame(self.root, padx=10, pady=10)
        top_frame.pack(fill=tk.X)

        self.load_btn = tk.Button(
            top_frame,
            text="Φόρτωση Excel",
            command=self.load_excel,
            font=("Arial", 12, "bold"),
            bg="#4CAF50",
            fg="white",
            padx=20,
            pady=10
        )
        self.load_btn.pack()

        # Label για πληροφορίες
        self.info_label = tk.Label(
            top_frame,
            text="Φορτώστε το Excel αρχείο για να ξεκινήσετε",
            font=("Arial", 10)
        )
        self.info_label.pack(pady=5)

        # Container frame για λίστα πελατών και στοιχεία (side by side)
        clients_container = tk.Frame(self.root)
        clients_container.pack(fill=tk.BOTH, expand=False, padx=10, pady=5)

        # Frame για λίστα πελατών (αριστερά)
        clients_frame = tk.LabelFrame(
            clients_container,
            text="Λίστα Πελατών",
            padx=10,
            pady=10
        )
        clients_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Scrollbar για τη λίστα
        scrollbar = tk.Scrollbar(clients_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Listbox για πελάτες
        self.clients_listbox = tk.Listbox(
            clients_frame,
            height=8,
            font=("Arial", 10),
            yscrollcommand=scrollbar.set
        )
        self.clients_listbox.pack(fill=tk.BOTH, expand=True)
        self.clients_listbox.bind('<<ListboxSelect>>', self.on_client_select)
        scrollbar.config(command=self.clients_listbox.yview)

        # Frame για στοιχεία πελάτη (δεξιά)
        client_info_frame = tk.LabelFrame(
            clients_container,
            text="Στοιχεία Επιλεγμένου Πελάτη",
            padx=10,
            pady=10
        )
        client_info_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0))

        # ScrolledText για στοιχεία πελάτη με scrollbar
        self.client_info_text = scrolledtext.ScrolledText(
            client_info_frame,
            height=8,
            font=("Arial", 9),
            state=tk.DISABLED,
            bg="#f0f0f0",
            wrap=tk.WORD
        )
        self.client_info_text.pack(fill=tk.BOTH, expand=True)

        # Frame για αγγελίες
        ads_frame = tk.LabelFrame(
            self.root,
            text="Αγγελίες (επικολλήστε εδώ - χωρίστε με κενή γραμμή)",
            padx=10,
            pady=10
        )
        ads_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.ads_text = scrolledtext.ScrolledText(
            ads_frame,
            height=10,
            font=("Arial", 10),
            wrap=tk.WORD
        )
        self.ads_text.pack(fill=tk.BOTH, expand=True)
        self.ads_text.bind('<KeyRelease>', self.count_ads)

        # Ενεργοποίηση paste με Ctrl+V και right-click
        self.ads_text.bind('<Control-v>', self.on_paste)
        self.ads_text.bind('<Button-3>', self.show_paste_menu)

        # Label για μέτρηση αγγελιών
        self.ads_count_label = tk.Label(
            ads_frame,
            text="Αριθμός αγγελιών: 0",
            font=("Arial", 9, "italic")
        )
        self.ads_count_label.pack(anchor=tk.W)

        # Frame για κουμπιά ενεργειών
        actions_frame = tk.Frame(self.root, padx=10, pady=10)
        actions_frame.pack(fill=tk.X)

        self.generate_btn = tk.Button(
            actions_frame,
            text="Δημιουργία Email",
            command=self.generate_email,
            font=("Arial", 12, "bold"),
            bg="#2196F3",
            fg="white",
            padx=20,
            pady=10,
            state=tk.DISABLED
        )
        self.generate_btn.pack(side=tk.LEFT, padx=5)

        self.copy_btn = tk.Button(
            actions_frame,
            text="Αντιγραφή Email",
            command=self.copy_to_clipboard,
            font=("Arial", 12, "bold"),
            bg="#FF9800",
            fg="white",
            padx=20,
            pady=10,
            state=tk.DISABLED
        )
        self.copy_btn.pack(side=tk.LEFT, padx=5)

        # Frame για το email
        email_frame = tk.LabelFrame(
            self.root,
            text="Έτοιμο Email",
            padx=10,
            pady=10
        )
        email_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.email_text = scrolledtext.ScrolledText(
            email_frame,
            height=15,
            font=("Arial", 10),
            wrap=tk.WORD,
            state=tk.DISABLED
        )
        self.email_text.pack(fill=tk.BOTH, expand=True)

    def load_excel(self):
        """Φορτώνει το Excel αρχείο"""
        file_path = filedialog.askopenfilename(
            title="Επιλέξτε το Excel αρχείο",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            # Διάβασμα Excel με υποστήριξη ελληνικών
            self.df = pd.read_excel(file_path)

            # Έλεγχος αν έχει δεδομένα
            if self.df.empty:
                messagebox.showerror("Σφάλμα", "Το αρχείο είναι κενό!")
                return

            # Καθαρισμός κενών γραμμών
            self.df = self.df.dropna(how='all')

            # Ενημέρωση της λίστας πελατών
            self.populate_clients_list()

            # Ενημέρωση UI
            self.info_label.config(
                text=f"Φορτώθηκαν {len(self.df)} πελάτες",
                fg="green"
            )

        except Exception as e:
            messagebox.showerror("Σφάλμα", f"Αποτυχία ανάγνωσης αρχείου:\n{str(e)}")

    def populate_clients_list(self):
        """Γεμίζει τη λίστα με τους πελάτες"""
        self.clients_listbox.delete(0, tk.END)

        # Προσπάθεια να βρούμε τη στήλη ονόματος
        name_column = None
        for col in self.df.columns:
            if any(keyword in str(col).lower() for keyword in ['όνομα', 'επώνυμο', 'name']):
                name_column = col
                break

        if name_column is None:
            name_column = self.df.columns[0]

        # Προσθήκη πελατών στη λίστα
        for idx, row in self.df.iterrows():
            name = str(row[name_column]) if pd.notna(row[name_column]) else f"Πελάτης {idx+1}"

            # Προσθήκη βασικών στοιχείων
            display_text = f"{idx+1}. {name}"

            # Προσθήκη τύπου ακινήτου αν υπάρχει
            for col in self.df.columns:
                if 'ψάχν' in str(col).lower() or 'τι ψάχν' in str(col).lower():
                    if pd.notna(row[col]):
                        property_type = str(row[col])[:30]
                        display_text += f" - {property_type}"
                    break

            self.clients_listbox.insert(tk.END, display_text)

    def on_client_select(self, event):
        """Χειρισμός επιλογής πελάτη"""
        selection = self.clients_listbox.curselection()
        if not selection:
            return

        idx = selection[0]
        self.selected_client = self.df.iloc[idx]

        # Ενημέρωση πληροφοριών πελάτη
        self.display_client_info()

        # Ενεργοποίηση κουμπιού δημιουργίας
        self.generate_btn.config(state=tk.NORMAL)

    def display_client_info(self):
        """Εμφανίζει τις πληροφορίες του επιλεγμένου πελάτη"""
        if self.selected_client is None:
            return

        self.client_info_text.config(state=tk.NORMAL)
        self.client_info_text.delete(1.0, tk.END)

        info_lines = []
        for col in self.df.columns:
            # Φιλτράρουμε τη στήλη συγκατάθεσης
            if any(keyword in str(col).lower() for keyword in ['συγκατάθεση', 'συγκαταθεση', 'gdpr', 'consent', 'checkbox']):
                continue

            value = self.selected_client[col]
            if pd.notna(value):
                info_lines.append(f"{col}: {value}")

        self.client_info_text.insert(1.0, "\n".join(info_lines))
        self.client_info_text.config(state=tk.DISABLED)

    def count_ads(self, event=None):
        """Μετράει τον αριθμό αγγελιών"""
        ads_content = self.ads_text.get(1.0, tk.END).strip()

        if not ads_content:
            self.ads_count_label.config(text="Αριθμός αγγελιών: 0")
            return

        # Χωρισμός με διπλή αλλαγή γραμμής (κενή γραμμή)
        ads = [ad.strip() for ad in re.split(r'\n\s*\n', ads_content) if ad.strip()]
        count = len(ads)

        self.ads_count_label.config(text=f"Αριθμός αγγελιών: {count}")

    def on_paste(self, event=None):
        """Χειρίζεται το paste event"""
        try:
            # Παίρνουμε το κείμενο από το clipboard
            clipboard_content = self.root.clipboard_get()
            # Εισάγουμε στο cursor position
            self.ads_text.insert(tk.INSERT, clipboard_content)
            # Μετράμε ξανά τις αγγελίες
            self.count_ads()
            return "break"  # Αποτρέπουμε το default paste behavior
        except:
            pass

    def show_paste_menu(self, event):
        """Εμφανίζει μενού right-click για paste"""
        try:
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="Επικόλληση", command=lambda: self.on_paste())
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def get_client_name(self):
        """Επιστρέφει το όνομα του πελάτη"""
        for col in self.df.columns:
            if any(keyword in str(col).lower() for keyword in ['όνομα', 'name']):
                value = self.selected_client[col]
                if pd.notna(value):
                    # Παίρνουμε μόνο το όνομα (πρώτη λέξη)
                    return str(value).split()[0]
        return "κύριε/κυρία"

    def get_property_type(self):
        """Επιστρέφει τον τύπο ακινήτου"""
        for col in self.df.columns:
            if 'ψάχν' in str(col).lower():
                value = self.selected_client[col]
                if pd.notna(value):
                    return str(value)
        return "ακίνητο"

    def extract_property_type_for_subject(self, full_property_type):
        """Εξάγει τον τύπο ακινήτου για το θέμα (π.χ. 'μονοκατοικίας' από 'Αγορά μονοκατοικίας')"""
        # Αφαιρούμε το "Αγορά" ή "Ενοικίαση" και κρατάμε το υπόλοιπο
        property_type = full_property_type.lower()
        property_type = property_type.replace('αγορά', '').replace('ενοικίαση', '').strip()
        return property_type

    def get_area(self):
        """Επιστρέφει την περιοχή"""
        for col in self.df.columns:
            if 'περιοχ' in str(col).lower() or 'area' in str(col).lower():
                value = self.selected_client[col]
                if pd.notna(value):
                    return str(value)
        return ""

    def get_budget(self):
        """Επιστρέφει το budget μορφοποιημένο"""
        for col in self.df.columns:
            if any(keyword in str(col).lower() for keyword in ['τιμή', 'ενοίκιο', 'budget', 'μέγιστο', 'επιθυμητό']):
                value = self.selected_client[col]
                if pd.notna(value):
                    # Μετατροπή σε αριθμό και μορφοποίηση
                    try:
                        # Καθαρισμός από σύμβολα
                        clean_value = str(value).replace('€', '').replace('.', '').replace(',', '').strip()
                        budget = int(float(clean_value))
                        # Μορφοποίηση: 100000 -> 100.000
                        formatted = f"{budget:,}".replace(',', '.')
                        return f"{formatted}€"
                    except:
                        return str(value)
        return ""

    def get_square_meters(self):
        """Επιστρέφει τα τετραγωνικά αν υπάρχουν"""
        for col in self.df.columns:
            if 'τετραγωνικ' in str(col).lower() or 'τ.μ' in str(col).lower():
                value = self.selected_client[col]
                if pd.notna(value) and str(value).strip():
                    return str(value)
        return None

    def generate_email(self):
        """Δημιουργεί το email"""
        if self.selected_client is None:
            messagebox.showwarning("Προσοχή", "Παρακαλώ επιλέξτε πελάτη!")
            return

        ads_content = self.ads_text.get(1.0, tk.END).strip()
        if not ads_content:
            messagebox.showwarning("Προσοχή", "Παρακαλώ προσθέστε τουλάχιστον μία αγγελία!")
            return

        # Χωρισμός αγγελιών
        ads = [ad.strip() for ad in re.split(r'\n\s*\n', ads_content) if ad.strip()]
        ads_count = len(ads)

        # Συλλογή στοιχείων πελάτη
        name = self.get_client_name()
        property_type_full = self.get_property_type()
        property_type = self.extract_property_type_for_subject(property_type_full)
        area = self.get_area()
        budget = self.get_budget()
        sqm = self.get_square_meters()

        # Δημιουργία θέματος
        subject = f"Πρόταση {property_type} σύμφωνα με τα κριτήριά σας"

        # Δημιουργία περιγραφής αιτήματος
        request_description = property_type_full
        if area:
            request_description += f" στη {area}"
        if budget:
            request_description += f" έως {budget}"
        if sqm:
            request_description += f" και {sqm} τ.μ."

        # Επιλογή σωστής διατύπωσης ανάλογα με τον αριθμό αγγελιών
        if ads_count == 1:
            proposal_text = "μία διαθέσιμη επιλογή που βρίσκεται πολύ κοντά στα κριτήριά σας:"
        elif ads_count == 2:
            proposal_text = "δύο διαθέσιμες επιλογές που βρίσκονται πολύ κοντά στα κριτήριά σας:"
        else:
            proposal_text = "τις παρακάτω διαθέσιμες επιλογές που βρίσκονται πολύ κοντά στα κριτήριά σας:"

        # Δημιουργία σώματος email
        email_body = f"""Καλησπέρα σας {name},

Σας ευχαριστώ για τη συμπλήρωση της φόρμας και το ενδιαφέρον σας.

Με βάση το αίτημά σας ({request_description}), σας προτείνω {proposal_text}

{ads_content}

Τηλέφωνο επικοινωνίας για τα παραπάνω ακίνητα: 6977917523

Αν ενδιαφέρεστε, μπορείτε να επικοινωνήσετε για ραντεβού ή για περισσότερες επιλογές που ταιριάζουν στα κριτήριά σας."""

        # Συνδυασμός θέματος και κειμένου
        full_email = f"Θέμα: {subject}\n\n{email_body}"

        # Εμφάνιση στο text widget
        self.email_text.config(state=tk.NORMAL)
        self.email_text.delete(1.0, tk.END)
        self.email_text.insert(1.0, full_email)
        self.email_text.config(state=tk.DISABLED)

        # Ενεργοποίηση κουμπιού αντιγραφής
        self.copy_btn.config(state=tk.NORMAL)

        messagebox.showinfo("Επιτυχία", "Το email δημιουργήθηκε με επιτυχία!")

    def copy_to_clipboard(self):
        """Αντιγράφει το email στο clipboard"""
        email_content = self.email_text.get(1.0, tk.END).strip()

        if not email_content:
            messagebox.showwarning("Προσοχή", "Δεν υπάρχει email για αντιγραφή!")
            return

        # Αντιγραφή στο clipboard
        self.root.clipboard_clear()
        self.root.clipboard_append(email_content)
        self.root.update()

        messagebox.showinfo("Επιτυχία", "Το email αντιγράφηκε στο clipboard!")


def main():
    """Κύρια συνάρτηση"""
    root = tk.Tk()
    app = EmailGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
