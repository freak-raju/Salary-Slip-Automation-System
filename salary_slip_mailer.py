# =========================================
# PROFESSIONAL SALARY SLIP MAILER V4
# =========================================

# INSTALL:
# pip install pandas openpyxl pywin32

# =========================================
# IMPORTS
# =========================================

import os
import threading
import pandas as pd
import tkinter as tk

from tkinter import (
    filedialog,
    messagebox,
    ttk
)

from datetime import datetime

import win32com.client as win32

from pypdf import PdfReader, PdfWriter
import tempfile

# =========================================
# MONTHS
# =========================================

MONTHS = [
    ("Jan", "jan"),
    ("Feb", "feb"),
    ("Mar", "mar"),
    ("Apr", "apr"),
    ("May", "may"),
    ("Jun", "jun"),
    ("Jul", "jul"),
    ("Aug", "aug"),
    ("Sep", "sep"),
    ("Oct", "oct"),
    ("Nov", "nov"),
    ("Dec", "dec")
]

# =========================================
# MAIN CLASS
# =========================================

class SalarySlipMailer:

    # =========================================
    # LOAD OUTLOOK ACCOUNTS
    # =========================================

    def load_outlook_accounts(self):

        try:

            outlook = win32.Dispatch(
                "Outlook.Application"
            )

            session = outlook.Session

            accounts = []

            for account in session.Accounts:

                accounts.append(
                    account.SmtpAddress
                )

            return accounts

        except:

            return []

    # =========================================
    # INIT
    # =========================================

    def __init__(self, root):

        self.root = root

        self.root.title(
            "Professional Salary Slip Mailer V4"
        )

        self.root.geometry("1400x850")

        self.root.state("zoomed")

        self.root.configure(bg="#1E1E1E")

        self.excel_path = tk.StringVar()

        self.total_count = tk.StringVar(value="0")
        self.sent_count = tk.StringVar(value="0")
        self.failed_count = tk.StringVar(value="0")
        self.pending_count = tk.StringVar(value="0")

        # =========================================
        # TITLE
        # =========================================

        title = tk.Label(
            root,
            text="Professional Salary Slip Mailer",
            font=("Arial", 24, "bold"),
            bg="#1E1E1E",
            fg="#00C853"
        )

        title.pack(pady=15)

        # =========================================
        # TOP FRAME
        # =========================================

        top_frame = tk.Frame(
            root,
            bg="#1E1E1E"
        )

        top_frame.pack(pady=10)

        tk.Label(
            top_frame,
            text="Excel File:",
            font=("Arial", 12),
            bg="#1E1E1E",
            fg="white"
        ).grid(
            row=0,
            column=0,
            padx=5
        )

        tk.Entry(
            top_frame,
            textvariable=self.excel_path,
            width=80,
            font=("Arial", 10)
        ).grid(
            row=0,
            column=1,
            padx=5
        )

        tk.Button(
            top_frame,
            text="Browse",
            width=12,
            bg="#2962FF",
            fg="white",
            command=self.browse_file
        ).grid(
            row=0,
            column=2,
            padx=5
        )

        # =========================================
        # DASHBOARD
        # =========================================

        dashboard = tk.Frame(
            root,
            bg="#1E1E1E"
        )

        dashboard.pack(pady=15)

        self.create_card(
            dashboard,
            "Total",
            self.total_count,
            0
        )

        self.create_card(
            dashboard,
            "Sent",
            self.sent_count,
            1
        )

        self.create_card(
            dashboard,
            "Failed",
            self.failed_count,
            2
        )

        self.create_card(
            dashboard,
            "Pending",
            self.pending_count,
            3
        )

        # =========================================
        # MONTH FRAME
        # =========================================

        month_frame = tk.LabelFrame(
            root,
            text="Select Months",
            bg="#1E1E1E",
            fg="white",
            font=("Arial", 11)
        )

        month_frame.pack(
            pady=10,
            padx=10
        )

        self.month_vars = {}

        for i, (display, key) in enumerate(MONTHS):

            var = tk.BooleanVar(value=True)

            self.month_vars[key] = var

            tk.Checkbutton(
                month_frame,
                text=display,
                variable=var,
                bg="#1E1E1E",
                fg="white",
                selectcolor="#333333"
            ).grid(
                row=i // 4,
                column=i % 4,
                padx=30,
                pady=10,
                sticky="w"
            )

        # =========================================
        # YEAR DROPDOWN
        # =========================================

        year_frame = tk.Frame(
            root,
            bg="#1E1E1E"
        )

        year_frame.pack(pady=10)

        tk.Label(
            year_frame,
            text="Select Year:",
            font=("Arial", 12),
            bg="#1E1E1E",
            fg="white"
        ).pack(side="left", padx=10)

        self.selected_year = tk.StringVar()

        self.year_dropdown = ttk.Combobox(
            year_frame,
            textvariable=self.selected_year,
            width=20,
            state="readonly"
        )

        self.year_dropdown.pack(side="left")

        self.load_dynamic_years()

        # =========================================
        # OUTLOOK ACCOUNT DROPDOWN
        # =========================================

        account_frame = tk.Frame(
            root,
            bg="#1E1E1E"
        )

        account_frame.pack(pady=10)

        tk.Label(
            account_frame,
            text="Send From:",
            font=("Arial", 12),
            bg="#1E1E1E",
            fg="white"
        ).pack(
            side="left",
            padx=10
        )

        self.selected_account = tk.StringVar()

        accounts = self.load_outlook_accounts()

        self.account_dropdown = ttk.Combobox(
            account_frame,
            textvariable=self.selected_account,
            values=accounts,
            width=40,
            state="readonly"
        )

        self.account_dropdown.pack(
            side="left"
        )

        if accounts:

            self.account_dropdown.current(0)

        # =========================================
        # SEARCH BOX
        # =========================================

        search_frame = tk.Frame(
            root,
            bg="#1E1E1E"
        )

        search_frame.pack(pady=10)

        tk.Label(
            search_frame,
            text="Search Employee:",
            bg="#1E1E1E",
            fg="white",
            font=("Arial", 11)
        ).pack(side="left")

        self.search_var = tk.StringVar()

        tk.Entry(
            search_frame,
            textvariable=self.search_var,
            width=30
        ).pack(side="left", padx=10)

        # =========================================
        # BUTTONS
        # =========================================

        btn_frame = tk.Frame(
            root,
            bg="#1E1E1E"
        )

        btn_frame.pack(pady=15)

        self.send_btn = tk.Button(
            btn_frame,
            text="Send Salary Slips",
            font=("Arial", 13, "bold"),
            bg="#00C853",
            fg="white",
            width=25,
            height=2,
            command=self.start_sending
        )

        self.send_btn.grid(
            row=0,
            column=0,
            padx=10
        )

        tk.Button(
            btn_frame,
            text="Open History",
            font=("Arial", 11),
            bg="#2962FF",
            fg="white",
            width=18,
            command=self.open_history
        ).grid(
            row=0,
            column=1,
            padx=10
        )

        # =========================================
        # PROGRESS BAR
        # =========================================

        self.progress = ttk.Progressbar(
            root,
            orient="horizontal",
            length=1000,
            mode="determinate"
        )

        self.progress.pack(pady=10)

        # =========================================
        # LOG BOX
        # =========================================

        log_frame = tk.Frame(
            root,
            bg="#1E1E1E"
        )

        log_frame.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=10
        )

        scrollbar = tk.Scrollbar(log_frame)

        scrollbar.pack(
            side="right",
            fill="y"
        )

        self.log_box = tk.Text(
            log_frame,
            bg="#121212",
            fg="#00FF7F",
            insertbackground="white",
            font=("Consolas", 10)
        )

        self.log_box.pack(
            side="left",
            fill="both",
            expand=True
        )

        self.log_box.config(
            yscrollcommand=scrollbar.set
        )

        scrollbar.config(
            command=self.log_box.yview
        )

    # =========================================
    # DASHBOARD CARD
    # =========================================

    def create_card(
        self,
        parent,
        title,
        variable,
        col
    ):

        frame = tk.Frame(
            parent,
            bg="#2C2C2C",
            width=200,
            height=90
        )

        frame.grid(
            row=0,
            column=col,
            padx=15
        )

        frame.grid_propagate(False)

        tk.Label(
            frame,
            text=title,
            font=("Arial", 12, "bold"),
            bg="#2C2C2C",
            fg="white"
        ).pack(pady=10)

        tk.Label(
            frame,
            textvariable=variable,
            font=("Arial", 18, "bold"),
            bg="#2C2C2C",
            fg="#00E676"
        ).pack()

    # =========================================
    # LOG FUNCTION
    # =========================================

    def log(self, message):

        self.log_box.insert(
            tk.END,
            message + "\n"
        )

        self.log_box.see(tk.END)

    # =========================================
    # LOAD YEARS
    # =========================================

    def load_dynamic_years(self):

        current_year = datetime.now().year

        years = []

        for i in range(
            current_year - 2,
            current_year + 3
        ):

            years.append(str(i))

        self.year_dropdown["values"] = years

        self.year_dropdown.current(2)

    # =========================================
    # BROWSE FILE
    # =========================================

    def browse_file(self):

        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")]
        )

        if file_path:

            self.excel_path.set(file_path)

    # =========================================
    # OPEN HISTORY
    # =========================================

    def open_history(self):

        history_file = "Salary_Mail_History.xlsx"

        if os.path.exists(history_file):

            os.startfile(history_file)

        else:

            messagebox.showerror(
                "Error",
                "History file not found"
            )

    # =========================================
    # START THREAD
    # =========================================

    def start_sending(self):

        thread = threading.Thread(
            target=self.send_mails
        )

        thread.daemon = True

        thread.start()

    # =========================================
    # GET SELECTED MONTHS
    # =========================================

    def get_selected_months(self):

        selected = []

        for display, key in MONTHS:

            if self.month_vars[key].get():

                selected.append(key)

        return selected

    # =========================================
    # GET PDF FILES
    # =========================================

    def get_pdf_files(
        self,
        folder_path,
        selected_months
    ):

        pdf_files = []

        year = self.selected_year.get()

        year_folder = os.path.join(
            folder_path,
            year
        )

        if not os.path.exists(year_folder):

            return pdf_files

        for file in os.listdir(year_folder):

            if file.lower().endswith(".pdf"):

                lower_file = file.lower()

                for month in selected_months:

                    if month.lower() in lower_file:

                        full_path = os.path.join(
                            year_folder,
                            file
                        )

                        pdf_files.append(full_path)

                        break

        return pdf_files

    # =========================================
    # CREATE PASSWORD PROTECTED PDF
    # =========================================

    def create_password_protected_pdf(self, input_pdf, password):

        temp_folder = tempfile.gettempdir()

        output_pdf = os.path.join(
            temp_folder,
             f"protected_{os.path.basename(input_pdf)}"
            )

        reader = PdfReader(input_pdf)
        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

        writer.encrypt(password)

        with open(output_pdf, "wb") as f:
            writer.write(f)

        return output_pdf
        
    # =========================================
    # SEND MAILS
    # =========================================

    def send_mails(self):

        self.send_btn.config(
            state="disabled"
        )

        excel_file = self.excel_path.get()

        if not excel_file:

            messagebox.showerror(
                "Error",
                "Please select Excel file"
            )

            self.send_btn.config(
                state="normal"
            )

            return

        selected_months = self.get_selected_months()

        if not selected_months:

            messagebox.showerror(
                "Error",
                "Select months"
            )

            self.send_btn.config(
                state="normal"
            )

            return

        try:

            df = pd.read_excel(excel_file)

            search_text = (
                self.search_var.get()
                .strip()
                .lower()
            )

            if search_text:

                df = df[
                    df["Employee Name"]
                    .str.lower()
                    .str.contains(search_text)
                ]

            total = len(df)

            self.total_count.set(str(total))

            self.pending_count.set(str(total))

            self.progress["maximum"] = total

            sent = 0
            failed = 0

            history = []

            for index, row in df.iterrows():

                name = str(
                    row["Employee Name"]
                ).strip()

                email = str(
                    row["Email ID"]
                ).strip()

                folder = str(
                    row["PDF Folder Path"]
                ).strip()

                self.log(
                    f"Processing: {name}"
                )

                pdf_files = self.get_pdf_files(
                    folder,
                    selected_months
                )

                if not pdf_files:

                    self.log(
                        f"No PDFs found: {name}"
                    )

                    failed += 1

                    continue

                try:

                    outlook = win32.Dispatch(
                        "Outlook.Application"
                    )

                    mail = outlook.CreateItem(0)

                    # =========================================
                    # SELECT SENDER ACCOUNT
                    # =========================================

                    selected_sender = (
                        self.selected_account.get()
                    )

                    for account in outlook.Session.Accounts:

                        if (
                            account.SmtpAddress.lower()
                            ==
                            selected_sender.lower()
                        ):

                            mail._oleobj_.Invoke(
                                *(64209, 0, 8, 0, account)
                            )

                            break

                    mail.To = email

                    month_text = ", ".join(
                        [m.upper() for m in selected_months]
                    )

                    year_text = self.selected_year.get()

                    mail.Subject = (
                        f"Salary Slips "
                        f"{month_text} {year_text}"
                    )

                    mail.Body = f"""
Dear {name},

Please find attached your password protected salary slips.

Password: Employee ID (Example: EMP12345)

Regards,
HR Team
"""

                    emp_id = str(row["Employee ID"]).strip()

                    for pdf in pdf_files:

                        protected_pdf = self.create_password_protected_pdf(
                            pdf,
                            emp_id
                        )

                        mail.Attachments.Add(protected_pdf)

                    mail.Send()

                    sent += 1

                    self.sent_count.set(str(sent))

                    self.log(
                        f"Mail sent: {email}"
                    )

                    status = "Success"

                except Exception as e:

                    failed += 1

                    self.failed_count.set(str(failed))

                    self.log(str(e))

                    status = f"Failed - {str(e)}"

                pending = total - (
                    sent + failed
                )

                self.pending_count.set(
                    str(pending)
                )

                history.append({

                    "Employee Name": name,

                    "Email ID": email,

                    "Years": year_text,

                    "Months": month_text,

                    "Status": status,

                    "Date Time": datetime.now().strftime(
                        "%d-%m-%Y %H:%M:%S"
                    )

                })

                self.progress["value"] = (
                    index + 1
                )

            # =========================================
            # SAVE HISTORY
            # =========================================

            history_file = (
                "Salary_Mail_History.xlsx"
            )

            new_df = pd.DataFrame(history)

            if os.path.exists(history_file):

                old_df = pd.read_excel(
                    history_file
                )

                final_df = pd.concat(
                    [old_df, new_df],
                    ignore_index=True
                )

            else:

                final_df = new_df

            final_df.to_excel(
                history_file,
                index=False
            )

            self.log(
                "History Saved Successfully"
            )

            messagebox.showinfo(
                "Completed",
                f"Sent: {sent}\n"
                f"Failed: {failed}"
            )

        except Exception as e:

            messagebox.showerror(
                "Error",
                str(e)
            )

        self.send_btn.config(
            state="normal"
        )

# =========================================
# MAIN
# =========================================

if __name__ == "__main__":

    root = tk.Tk()

    app = SalarySlipMailer(root)

    root.mainloop()