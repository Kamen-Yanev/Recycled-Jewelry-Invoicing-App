from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import os

from invoice import invoice, customers
from config import VAT_FORMS_DIR


class App:
    def __init__(self):
        self.root = Tk()
        self.cpm_document = ''
        self.client = ''
        self.vat = ''
        self.client_rates = customers()
        self.build_ui()
        self.root.mainloop()

    def upload_cpm_doc(self):
        path = filedialog.askopenfilename(parent=self.root)
        if path:
            self.cpm_document = path

    def build_ui(self):
        self.root.geometry('640x200')
        self.root.title('Invoice Helper')

        title_frame = ttk.Frame(self.root)
        title_frame.pack(pady=(20, 0))

        title_label = ttk.Label(title_frame, text='Invoice Helper', font=('TkDefaultFont', 16, 'bold'))
        title_label.pack()

        subtitle_label = ttk.Label(title_frame, text='Generate client invoices from CPM documents', foreground='gray')
        subtitle_label.pack()

        sep_frame = ttk.Frame(self.root)
        sep_frame.pack(fill='x', padx=20, pady=10)
        ttk.Separator(sep_frame, orient='horizontal').pack(fill='x')

        mainframe = ttk.Frame(self.root, padding=(20, 0, 20, 20))
        mainframe.pack()

        ttk.Label(mainframe, text='Client:').grid(column=0, row=0, sticky='e', padx=(0, 5))

        client_rates_list = [i for i in self.client_rates.keys()]
        client_rates_list.append('New client')
        n = StringVar()
        self.combobox = ttk.Combobox(mainframe, width=25, textvariable=n, values=client_rates_list, state='readonly')
        self.combobox.grid(column=1, row=0, padx=(0, 10))

        button_frame = ttk.Frame(mainframe)
        button_frame.grid(column=2, row=0)
        ttk.Button(button_frame, text='Upload CPM', command=self.upload_cpm_doc).grid(column=0, row=0, padx=(0, 5))
        ttk.Button(
            button_frame,
            text='Generate Invoice',
            command=lambda: invoice(self.root, self.cpm_document, self.client, self.vat, self.client_rates)
        ).grid(column=1, row=0)

        self.combobox.bind('<<ComboboxSelected>>', self.select_client)

    def select_client(self, event=None):
        if self.combobox.get() == 'New client':
            self.open_new_client_dialog()
        else:
            self.client = self.combobox.get()
            self.vat = os.path.join(VAT_FORMS_DIR, f'{self.client}.pdf')

    def open_new_client_dialog(self):
        top = Toplevel()
        top.grab_set()
        top.geometry("340x270")
        top.title("New client details")
        top.configure(bg="#f9f9f9")
        top.resizable(False, False)

        frame = Frame(top, bg="#f9f9f9", padx=18, pady=16)
        frame.pack(fill='both', expand=True)

        title_label = Label(frame, text="New client details",
                            font=('TkDefaultFont', 11, 'bold'),
                            bg="#f9f9f9", fg="#1a1a1a")
        title_label.grid(column=0, row=0, columnspan=2, sticky='w', pady=(0, 4))

        separator = Frame(frame, height=1, bg="#e0e0e0")
        separator.grid(column=0, row=1, columnspan=2, sticky='ew', pady=(0, 12))

        metal_label = Label(frame, text="Metal:", font=('TkDefaultFont', 10),
                            bg="#f9f9f9", fg="#555555", anchor='e', width=9)
        discount_label = Label(frame, text="Discount:", font=('TkDefaultFont', 10),
                               bg="#f9f9f9", fg="#555555", anchor='e', width=9)

        metal_var = StringVar()
        discount_var = StringVar()

        metal_entry = Entry(frame, textvariable=metal_var,
                            relief='flat', bd=1, highlightthickness=1,
                            highlightbackground="#cccccc", highlightcolor="#888888",
                            font=('TkDefaultFont', 10))
        discount_entry = Entry(frame, textvariable=discount_var,
                               relief='flat', bd=1, highlightthickness=1,
                               highlightbackground="#cccccc", highlightcolor="#888888",
                               font=('TkDefaultFont', 10))

        metal_label.grid(column=0, row=2, sticky='e', pady=4)
        metal_entry.grid(column=1, row=2, sticky='ew', padx=(8, 0), pady=4)
        discount_label.grid(column=0, row=3, sticky='e', pady=4)
        discount_entry.grid(column=1, row=3, sticky='ew', padx=(8, 0), pady=4)

        frame.columnconfigure(1, weight=1)

        warn_frame = Frame(frame, bg="#fffbf0",
                           highlightthickness=1, highlightbackground="#f0d080")
        warn_frame.grid(column=0, row=4, columnspan=2, sticky='ew', pady=(12, 0))

        warn_inner = Frame(warn_frame, bg="#fffbf0", padx=10, pady=8)
        warn_inner.pack(fill='both')

        Label(warn_inner, text="If rate of day, enter 0.",
              font=('TkDefaultFont', 9, 'bold'), bg="#fffbf0", fg="#7a5500",
              anchor='w').pack(fill='x')
        Label(warn_inner, text="Metal must be spelled exactly as AU, AG, PT or PD",
              font=('TkDefaultFont', 9), bg="#fffbf0", fg="#555555",
              anchor='w').pack(fill='x')
        Label(warn_inner, text="For multiple metals, submit and repeat process",
              font=('TkDefaultFont', 9), bg="#fffbf0", fg="#555555",
              anchor='w').pack(fill='x')

        self.client = self.combobox.get()

        def submit():
            client = 'New client'
            metal = metal_var.get()
            discount = float(discount_var.get())
            if client in self.client_rates:
                self.client_rates[client][metal] = discount
            else:
                self.client_rates[client] = {metal: discount}
            top.destroy()

        submit_button = Button(frame, text="Submit information", command=submit,
                               font=('TkDefaultFont', 10), relief='flat',
                               bg="#e8e8e8", activebackground="#d4d4d4",
                               cursor="hand2", pady=6)
        submit_button.grid(column=0, row=5, columnspan=2, sticky='ew', pady=(14, 14))


if __name__ == '__main__':
    App()
