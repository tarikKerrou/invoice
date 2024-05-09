import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
import os
import docx2pdf



def clear_item():
    NBRJR_sprinBox.delete(0,tkinter.END)
    NBRJR_sprinBox.insert(0,"1")
    DATE_enter.delete(0,tkinter.END)
    PrixParJr_sprinbox.delete(0,tkinter.END)
    PrixParJr_sprinbox.insert(0,"1")
    DESIGNATION_entry.delete(0,tkinter.END)


invoice_items=[]
def add_item():
    nbr_jr=int(NBRJR_sprinBox.get())
    date=DATE_enter.get()
    designation=DESIGNATION_entry.get()
    prix_par_jour=float(PrixParJr_sprinbox.get())
    line_total = round(nbr_jr * prix_par_jour, 2)
    facteur_item=[nbr_jr,date,designation,prix_par_jour,line_total]
    tree.insert('',0,values=facteur_item)
    clear_item()
    invoice_items.append(facteur_item)

def new_invoice():
    client_address_entry.delete(0,tkinter.END)
    client_ice_entry.delete(0,tkinter.END)
    client_name_entry.delete(0,tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())
    invoice_items.clear()

def generate_invoice():
    doc = DocxTemplate("invoice_template.docx")
    name_client = client_name_entry.get()
    ice = client_ice_entry.get()
    adress = client_address_entry.get()
    totalHt = round(sum(item[4] for item in invoice_items),2)
    tva = round(totalHt * 0.13,2)
    total = round(totalHt + tva ,2)
    date = datetime.datetime.now().strftime("%Y/%m/%d")
    
    with open("config.txt", "r") as file:
        numero_facture = int(file.read().strip())
    
    doc.render({
        "name_client": name_client,
        "adress": adress,
        "ice": ice,
        "invoice_lst": invoice_items,
        "totalHt": totalHt,
        "tva": tva,
        "Totalc": total,
        "date": date,
        "fac_num": numero_facture
    })

    doc_name = f"invoice_{name_client}_{datetime.datetime.now().strftime('%Y-%m-%d-%H-h-%M-min-%S-s')}.docx"
    doc.save(doc_name)

    # Convertir le fichier .docx en PDF
    pdf_name = doc_name.replace(".docx", ".pdf")
    docx2pdf.convert(doc_name, pdf_name)

    # Supprimer le fichier .docx après conversion
    os.remove(doc_name)

    # Incrémenter le numéro de facture
    numero_facture += 1

    # Sauvegarde du nouveau numéro de facture dans config.txt
    with open("config.txt", "w") as file:
        file.write(str(numero_facture))
    
    year_directory = datetime.datetime.now().strftime('%Y')
    if not os.path.exists(year_directory):
        os.makedirs(year_directory)

    # Déplacer le fichier PDF dans le répertoire de l'année
    destination_pdf = os.path.join(year_directory, os.path.basename(pdf_name))
    os.rename(pdf_name, destination_pdf)
    
    new_invoice()
   
window=tkinter.Tk()
window.title("Generate invoive form")

frame=tkinter.Frame(window)
frame.pack()
client_name_label= tkinter.Label(frame,text="Client name")
client_name_label.grid(row=1,column=0)

client_address_label= tkinter.Label(frame,text="Client address")
client_address_label.grid(row=1,column=1)
client_ice_label= tkinter.Label(frame,text="Client ice")
client_ice_label.grid(row=1,column=2)

client_name_entry=tkinter.Entry(frame)
client_address_entry=tkinter.Entry(frame)
client_ice_entry=tkinter.Entry(frame)

client_name_entry.grid(row=2,column=0)
client_address_entry.grid(row=2,column=1)
client_ice_entry.grid(row=2,column=2)

NBR_jr_label=tkinter.Label(frame,text="Nombre des jours")
NBR_jr_label.grid(row=3,column=0)

DATE_label=tkinter.Label(frame,text="DATE")
DATE_label.grid(row=3,column=1)

DESIGNATION_label=tkinter.Label(frame,text="DESIGNATION")
DESIGNATION_label.grid(row=3,column=2)

PrixParJr_label=tkinter.Label(frame,text="Prix Par Jour")
PrixParJr_label.grid(row=3,column=3)

NBRJR_sprinBox=tkinter.Spinbox(frame,from_=1,to=1000)
NBRJR_sprinBox.grid(row=4,column=0)

PrixParJr_sprinbox=tkinter.Spinbox(frame,from_=1,to=100000000)
PrixParJr_sprinbox.grid(row=4,column=3)

DATE_enter=tkinter.Entry(frame)
DESIGNATION_entry=tkinter.Entry(frame)

DATE_enter.grid(row=4,column=1)
DESIGNATION_entry.grid(row=4,column=2)

add_item_button=tkinter.Button(frame,text="Add Item",command=add_item)
add_item_button.grid(row=5,column=2,pady=5)

columns=('nbr_jr','date','designation','Prix_Par_Jour','total')
tree= ttk.Treeview(frame,columns=columns,show="headings")
tree.grid(row=6,column=0,columnspan=3,padx=20,pady=10)

tree.heading('nbr_jr',text='nbr_jr')
tree.heading('date',text='date')
tree.heading('designation',text='designation')
tree.heading('Prix_Par_Jour',text='Prix_Par_Jour')
tree.heading('total',text='Prix_Par_Jour')


save_facteur_button=tkinter.Button(frame,text="generate invoice",command=generate_invoice)
save_facteur_button.grid(row=7,column=0,columnspan=3,sticky="news",padx=20,pady=5)

new_facteur_button=tkinter.Button(frame,text="new invoice",command=new_invoice)
new_facteur_button.grid(row=8,column=0,columnspan=3,sticky="news",padx=20,pady=5)




window.mainloop()