import tkinter
import os
from datetime import date

import warnings
import datetime
# Suppress specific warnings
# warnings.filterwarnings("ignore", category=UserWarning, message="Duplicate name: 'docProps/core.xml'")

from tkinter import ttk
from docxtpl import DocxTemplate

# kkdd
if __name__ == '__main__':

    def generate_invoice():

        save_directory = "/Users/rahatbhatti/Desktop/save_files"

        file =DocxTemplate("InvoiceTemplate.docx")
        try:
            # date = date_entry.get()
            name = To_name_entry.get()
            today = date.today()
            # print("Today's date:", today)
            print(date)
            file.render({
                "date" : today,
                "name":name,
                "invoice":invoiceNumber_entry.get(),
                "street":to_street_entry.get(),
                "city" : to_cityp_entry.get(),
                "invoice_list": invoice_list,
                "title_combobox":title_combobox.get()

            })
            file_name = to_cityp_entry.get()+".docx"
            full_file_path = os.path.join(save_directory, file_name)

            file.save(full_file_path)
            print(name)
            # file = DocxTemplate(full_file_path)
            print(file)
        except Exception as e:
            print(f"An error occurred: {e}")

    def clear_invoice():
        date_entry.delete(0, tkinter.END)
        invoiceNumber_entry.delete(0, tkinter.END)
        To_name_entry.delete(0,tkinter.END)
        to_street_entry.delete(0, tkinter.END)
        to_cityp_entry.delete(0, tkinter.END)
        clear()
        tree.delete(*tree.get_children())
        print("pressed")
        invoice_list.clear()

    def clear():
        user_name_entry.delete(0, tkinter.END)
        user_ticket_entry.delete(0, tkinter.END)
        billed_to_entry.delete(0, tkinter.END)
        price_label_entry.delete(0, tkinter.END)

    invoice_list=[]
    def addCustomer():
        name = user_name_entry.get()
        tktnumber = user_ticket_entry.get()
        billed_to = billed_to_entry.get()
        price = price_label_entry.get()

        items = [name, tktnumber, billed_to, price]
        tree.insert("", 0, values=items)
        clear()

        invoice_list.append(items)
        print(invoice_list)


    window = tkinter.Tk()
    window.title("Invoice Application")
    window.geometry('1000x700')

    frame = tkinter.Frame(window)
    frame.pack()


    ticket_info_frame = tkinter.LabelFrame(frame,text = "Ticket Information")
    ticket_info_frame.grid(row = 0, column = 0)
    # invoice number , Date , To ,

    # decleration
    invoiceNumber_label = tkinter.Label(ticket_info_frame, text="Invoice Number")
    date_label = tkinter.Label(ticket_info_frame, text="Date")

    To_NAME_label = tkinter.Label(ticket_info_frame, text="To")
    To_Street_label = tkinter.Label(ticket_info_frame, text="Street")
    invoiceNumber_entry = tkinter.Entry(ticket_info_frame)

    To_name_entry = tkinter.Entry(ticket_info_frame)
    date_entry = tkinter.Entry(ticket_info_frame)
    to_street_entry = tkinter.Entry(ticket_info_frame)

    # title_label = tkinter.Label(ticket_info_frame, text = "Title")
    title_combobox= ttk.Combobox(ticket_info_frame,values = ["Mr","Ms","Mrs","Miss"])

    To_cityp_label = tkinter.Label(ticket_info_frame, text="City-Province")
    to_cityp_entry = tkinter.Entry(ticket_info_frame)

    # show inside the screen
    date_label.grid(row = 0 , column = 0)
    date_entry.grid(row = 0 , column = 1)
    invoiceNumber_label.grid(row = 0 , column = 2)
    invoiceNumber_entry.grid(row = 0 , column = 3)

    # title_label.grid(row = 1,column = 0)

    To_NAME_label.grid(row = 1 , column = 0)
    title_combobox.grid(row=1, column=1)
    To_name_entry.grid(row = 1 , column = 3)

    To_Street_label.grid(row = 2 , column = 0)
    to_street_entry.grid(row = 2 , column = 1)
    To_cityp_label.grid(row = 2 , column = 2)
    to_cityp_entry.grid(row = 2 , column = 3)
    # have to add more option for date

    # for widget in  ticket_info_frame.winfo_children():
    #     widget.grid_configure(padx = 10,pady=5)

    user_info_frame = tkinter.LabelFrame(frame,text ="User Information")
    user_info_frame.grid(row=1 ,column = 0,sticky = "news" ,padx = 20 , pady= 20)

    user_name_label = tkinter.Label(user_info_frame,text = "Name")
    user_name_entry = tkinter.Entry(user_info_frame)
    user_ticket = tkinter.Label(user_info_frame,text = "Ticket Number")
    user_ticket_entry = tkinter.Entry(user_info_frame)
    billed_to_label = tkinter.Label(user_info_frame,text = "Billed to")
    billed_to_entry = tkinter.Entry(user_info_frame)
    price_label = tkinter.Label(user_info_frame,text = "Ticket Price")
    price_label_entry = tkinter.Entry(user_info_frame)
    add_button  = tkinter.Button(user_info_frame,text = "Add Customer",command=addCustomer)

    user_name_label.grid(row = 0 , column = 0)
    user_name_entry.grid(row = 0, column= 1)
    user_ticket.grid(row = 0 , column =2)
    user_ticket_entry.grid(row = 0 ,column=3)

    billed_to_label.grid(row = 1,column = 0)
    billed_to_entry.grid(row=1,column = 1)
    price_label.grid(row=1,column = 2)
    price_label_entry.grid(row=1,column = 3)
    add_button.grid(row=2,column=3)

    # for widget in  user_info_frame.winfo_children():
    #     widget.grid_configure(padx = 5,pady=5)

    treeframe = tkinter.LabelFrame(frame,text ="User Information")
    treeframe.grid(row=2 ,column = 0,sticky = "news" ,padx = 20 ,)


    column = ("Name","Ticket", "Billed to" , "Ticket Price")
    tree = ttk.Treeview(treeframe,columns= column,show="headings")
    tree.grid(row = 5,column = 0, columnspan = 2,padx = 20, pady=10)

    tree.heading("Name",text = "Name")
    tree.heading("Ticket",text = "Ticket Number")
    tree.heading("Billed to",text = "Billed to")
    tree.heading("Ticket Price",text = "Ticket Price")

    generate_invoice_button = tkinter.Button(treeframe,text = "Generate Invoice",command= generate_invoice)
    generate_invoice_button.grid(row = 6,column = 0,columnspan = 3,padx =20 ,pady =5 )

    clear_invoice_button = tkinter.Button(treeframe, text="Clear Invoice",command = clear_invoice)
    clear_invoice_button.grid(row=7, column=0, columnspan=3, padx=20, pady = 5)


    window.mainloop()

# from tkinter import*
# window = Tk()
# window.title("Invoice Application")
# window.geometry('600x400')
#
#
# date_label = Label(window, text="Date")
# date_label.pack()
# window.mainloop()
