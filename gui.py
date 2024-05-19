import tkinter as tk
from tkinter import ttk
import nikitooos

def send_reference():
    employee_id = id_entry.get()
    try:
        nikitooos.process_reference(employee_id)
        result_label.config(text="Справка успешно отправлена", foreground="green")
    except Exception as e:
        result_label.config(text=f"Ошибка при отправке справки: {e}", foreground="red")

def create_employee_tab(notebook):
    employee_tab = ttk.Frame(notebook)
    notebook.add(employee_tab, text='Информация о сотрудниках')

    employee_data = nikitooos.load_employee_data()

    tree = ttk.Treeview(employee_tab, columns=("№", "Должность", "ФИО", "Дата", "№ договора"), show="headings")
    tree.heading("№", text="№")
    tree.heading("Должность", text="Должность")
    tree.heading("ФИО", text="ФИО")
    tree.heading("Дата", text="Дата")
    tree.heading("№ договора", text="№ договора")

    for row in employee_data:
        tree.insert("", "end", values=row)

    tree.pack(expand=True, fill="both")

def create_issued_references_tab(notebook):
    issued_references_tab = ttk.Frame(notebook)
    notebook.add(issued_references_tab, text='Реестр выданных справок')

    issued_references_data = nikitooos.load_issued_references()

    tree = ttk.Treeview(issued_references_tab, columns=("Reference Number", "Name", "Issue Date"), show="headings")
    tree.heading("Reference Number", text="Reference Number")
    tree.heading("Name", text="Name")
    tree.heading("Issue Date", text="Issue Date")

    for row in issued_references_data:
        tree.insert("", "end", values=row)

    tree.pack(expand=True, fill="both")

root = tk.Tk()
root.title("Выдача справок")

frame = ttk.Frame(root, padding="20")
frame.grid(column=0, row=0, sticky=(tk.W, tk.E, tk.N, tk.S))
frame.columnconfigure(0, weight=1)
frame.rowconfigure(0, weight=1)

id_label = ttk.Label(frame, text="Идентификатор сотрудника:")
id_label.grid(column=0, row=0, sticky=tk.W)

id_entry = ttk.Entry(frame, width=20)
id_entry.grid(column=1, row=0, sticky=tk.W)

send_button = ttk.Button(frame, text="Отправить справку", command=send_reference)
send_button.grid(column=0, row=1, columnspan=2, pady=10)

result_label = ttk.Label(frame, text="", foreground="green")
result_label.grid(column=0, row=2, columnspan=2)

notebook = ttk.Notebook(root)
notebook.grid(row=1, column=0, columnspan=2, sticky="NESW")

create_employee_tab(notebook)
create_issued_references_tab(notebook)

root.mainloop()
