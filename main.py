from tkinter import *
from tkinter import messagebox
import os
import pandas
import docx
import datetime

FONT = ("Arial", 20, "normal")
WORKING_DIRECTORY = os.getcwd()


# noinspection PyBroadException
class Prescription:
    # Create the prescription page
    def __init__(self):
        self.patient_id = 0
        os.chdir(WORKING_DIRECTORY)
        self.list_of_indexes = []
        self.window = Tk()
        self.doctor_name = str
        self.window.title("Prescription Writer 2.0")
        self.window.config(padx=50, pady=50)
        self.welcome = Label(text="Welcome", font=("Arial", 78, "bold"), fg="blue")
        self.welcome.grid(row=0, column=0, columnspan=2)
        self.canvas = Canvas(width=600, height=50)
        self.canvas.create_line(0, 50, 600, 50)
        self.canvas.grid(row=1, column=0, columnspan=2, pady=20)
        self.name_label = Label(text="Name: ", font=FONT)
        self.name_label.grid(row=2, column=0)
        self.name_entry = Entry(width=50)
        self.name_entry.grid(row=2, column=1)
        self.age_label = Label(text="Age: ", font=FONT)
        self.age_label.grid(row=3, column=0)
        self.age_entry = Entry(width=50)
        self.age_entry.grid(row=3, column=1)
        self.gender_label = Label(text="Gender: ", font=FONT)
        self.gender_label.grid(row=4, column=0)
        self.gender_entry = Entry(width=50)
        self.gender_entry.grid(row=4, column=1)
        self.issues_label = Label(text="Issues: ", font=FONT)
        self.issues_label.grid(row=5, column=0)
        self.issues_text_area = Text(width=33, height=5, font=("Arial", 12, "normal"))
        self.issues_text_area.grid(row=5, column=1)
        self.prescription_label = Label(text="Prescription: ", font=FONT)
        self.prescription_label.grid(row=6, column=0)
        self.prescription_text_area = Text(width=33, height=5, font=("Arial", 12, "normal"))
        self.prescription_text_area.grid(row=6, column=1, pady=20)
        self.billing_label = Label(text="Total amount (₹): ", font=FONT)
        self.billing_label.grid(row=7, column=0)
        self.billing_entry = Entry(width=50)
        self.billing_entry.grid(row=7, column=1)
        # self.doctor_name_label = Label(text="Dr.", font=FONT)
        # self.doctor_name_label.grid(row=3, column=2)
        # self.doctor_name_entry = Entry(width=50)
        # self.doctor_name_entry.grid(row=3, column=3)
        self.show_index = Button(text="Show available indexes", font=FONT, padx=2, pady=2, command=self.show_indexes)
        self.show_index.grid(column=3, row=5)
        self.generate = Button(text="Generate", font=FONT, padx=2, pady=2, command=self.generate_prescription)
        self.generate.grid(column=3, row=6)
        self.p_index = Button(text="Use index", font=FONT, padx=2, pady=2, command=self.use_index)
        self.p_index.grid(column=3, row=7)
        self.add_index_button = Button(text="Add an index", font=FONT, padx=2, pady=2, command=self.add_indexes)
        self.add_index_button.grid(column=3, row=4)
        self.define_indexes()
        self.window.mainloop()

    def generate_prescription(self):
        try:
            # Get patient ID number
            self.get_patient_id()
            # Read data from CSV file and then convert
            prescription_data = pandas.read_excel("Prescriptions/Prescriptions.xlsx").to_dict(orient="records")
            name = self.name_entry.get()
            age = self.age_entry.get()
            gender = self.gender_entry.get()
            bill = self.billing_entry.get()
            issues = self.issues_text_area.get(1.0, END)
            prescription_result = self.prescription_text_area.get(1.0, END)
            # self.doctor_name = self.doctor_name_entry.get().title()
            new_item = {
                "Date": datetime.datetime.now().date().strftime("%d/%m/%Y"),
                "Name": name.title(),
                "Age": age,
                "Gender": gender.title(),
                "Total Bill": f"₹{int(bill)}",
                "Doctor Name": f"Dr. Kandarp Vidyarthi",
            }
            prescription_data.append(new_item)
            data_ready = pandas.DataFrame(prescription_data)
            data_ready.to_excel("Prescriptions/Prescriptions.xlsx", index=False)
            # Create new docx file containing data
            document = docx.Document()
            document.add_paragraph(f"\n\n\n\nDate: {datetime.datetime.now().date().strftime("%d/%m/%Y")}\nPatient ID: {self.patient_id}\n"
                f"Name: {name.title()}\nAge: {age}\nGender: {gender.title()}\n\n\nHistory: \n{issues.title()}\n"
                f"Prescription: \n{prescription_result.title()}"
                f"\n\n\nDr. Kandarp Vidyarthi"
            )
            document.save(f"Prescriptions/{name}.docx")
            # Reset the fields
            self.reset()
        except Exception as e:
            messagebox.showerror("Something went wrong", str(e))

    def reset(self):
        self.name_entry.delete(0, 'end')
        self.age_entry.delete(0, 'end')
        self.gender_entry.delete(0, 'end')
        self.issues_text_area.delete('1.0', 'end')
        self.prescription_text_area.delete('1.0', 'end')
        self.billing_entry.delete(0, 'end')
        # self.doctor_name_entry.delete(0, 'end')
        self.set_patient_id()
        self.name_entry.focus()

    # Patient ID functions

    @staticmethod
    def set_patient_id():
        try:
            with open("patientID.txt", mode="r+") as patient_id_file:
                new_num = int(patient_id_file.read()) + 1
                patient_id_file.seek(0)
                patient_id_file.truncate(0)
                patient_id_file.write(str(new_num))
        except Exception as e:
            messagebox.showerror("Something went wrong", str(e))

    def get_patient_id(self):
        try:
            with open("patientID.txt", mode="r") as patient_id_file:
                self.patient_id = patient_id_file.read()
        except Exception as e:
            messagebox.showerror("Something went wrong", str(e))

    # Index functions

    def use_index(self):
        try:
            index_for_change = self.issues_text_area.get(1.0, END).strip()
            self.issues_text_area.delete(1.0, END)
            self.prescription_text_area.delete(1.0, END)
            print(f"{index_for_change}")
            # Backache
            if index_for_change == "Backache\n":
                self.issues_text_area.insert(index=1.0, chars="Backache\nComplaint of severe low backache with radiation "
                                                              "to lower limb since {replace with time}, \nNo History of "
                                                              "trauma, fever\nOn examination - severe lumber "
                                                              "spasm\nStraight leg raising test\nNo neurovascular deficit")
                self.prescription_text_area.insert(index=1.0, chars="Tab Flexura D (diclofenac and metaxalone) 12 hrly "
                                                                    "for 5 days after meals (might cause acidity and "
                                                                    "sedation)\nTab Pantocid (pantoprazole) 40 mg once "
                                                                    "daily empty stomach for 5 days in case of "
                                                                    "acidity\nTab Pregabid (pregabalin) 75 mg bed time "
                                                                    "for 15 days (might cause sedation)\nTab Dezed ("
                                                                    "deflazocort) 30 mg once daily for 10 days\nTab "
                                                                    "medrol (methylprednisolone) 8mg once daily for 10 "
                                                                    "days\nTab vertin 8 mg 8 hrly for 10 days\nTab Lumia "
                                                                    "(Vitamin D3) 60000 units once a week for 2 "
                                                                    "months\nTab Osteofit HD (Calcium, Vitamin D3 and "
                                                                    "Vitamin B12) once daily for 3 months\nPrecautions – "
                                                                    "avoid bending forward, lifting heavy weight\nSpinal "
                                                                    "extension exercises (when pain "
                                                                    "decreases)\nPhysiotherapy - Interferential therapy, "
                                                                    "ultrasonic therapy and TENS therapy for 10 "
                                                                    "sittings\nVolitra APS spray and hot packs\n\nXray - "
                                                                    "Lumbosacral spine - AP and lateral view\nMRI L-S "
                                                                    "spine")
                self.issues_text_area.update()
                self.prescription_text_area.update()
            # Neck Pain
            elif index_for_change == "Neck Pain\n":
                self.issues_text_area.insert(index=1.0, chars="Neck Pain\nComplaint of neck pain with radiation to upper "
                                                              "limb, vertigo, headache, nausea\nOn examination - "
                                                              "Tenderness and spasm in cervical region\nNo neurological "
                                                              "deficit")
                self.prescription_text_area.insert(index=1.0, chars="Xray cervical spine - AP and lateral\nTab Flexura D "
                                                                    "(diclofenac and metaxalone) 12 hrly for 5 days after "
                                                                    "meals (might cause acidity and sedation)\nTab "
                                                                    "Pantocid (pantoprazole) 40 mg once daily empty "
                                                                    "stomach for 5 days in case of acidity\nTab Pregabid "
                                                                    "(pregabalin) 75 mg bed time for 15 days (might cause "
                                                                    "sedation)\nTab Dezed (deflazocort) 30 mg once daily "
                                                                    "for 10 days\nTab medrol (methylprednisolone) 8mg "
                                                                    "once daily for 10 days\nTab vertin 8 mg 8 hrly for "
                                                                    "10 days\nTab Lumia (Vitamin D3) 60000 units once a "
                                                                    "week for 2 months\nTab Osteofit HD (Calcium, "
                                                                    "Vitamin D3 and Vitamin B12) once daily for 3 "
                                                                    "months\nPrecautions - use a light pillow, "
                                                                    "avoid bending at the neck and keep head straight "
                                                                    "while working on computer and phone\nDesk stretches, "
                                                                    "trapezius and levator scapulae stretches\nIsometric "
                                                                    "neck range of motion exercises (when pain "
                                                                    "decreases)\nPhysiotherapy - Interferential therapy, "
                                                                    "ultrasonic therapy and TENS therapy for 10 "
                                                                    "sittings\nVolitra APS spray and hot packs\nReview "
                                                                    "after 2 weeks")
                self.issues_text_area.update()
                self.prescription_text_area.update()
            else:
                e = True
                index_for_change.replace("\\n", "")
                # Other
                for i in self.list_of_indexes:
                    print(index_for_change, i[0].rstrip())
                    if index_for_change == i[0]:
                        self.issues_text_area.insert(index=1.0, chars=i[0])
                        self.prescription_text_area.insert(index=1.0, chars=i[1])
                        self.issues_text_area.update()
                        self.prescription_text_area.update()
                        e = False
                        break
                if e:
                    e2 = True
                    index_for_change += "\\n"
                    for i in self.list_of_indexes:
                        print(index_for_change, i[0].rstrip())
                        if index_for_change == i[0]:
                            self.issues_text_area.insert(index=1.0, chars=i[0])
                            self.prescription_text_area.insert(index=1.0, chars=i[1])
                            self.issues_text_area.update()
                            self.prescription_text_area.update()
                            e2 = False
                            break
                    if e2:
                        messagebox.showerror(title="Invalid", message="That is an invalid index. Please correct the entry.")
        except Exception as error:
            messagebox.showerror("Something went wrong", str(error))

    def define_indexes(self):
        try:
            with open("indexes.txt") as file:
                for line in file:
                    parts = line.split('|')
                    words, paragraph = parts[0].strip(), parts[1].strip().replace("\\n", "\n")
                    self.list_of_indexes.append((words, paragraph))
        except Exception as e:
            messagebox.showerror("Something went wrong", str(e))

    def show_indexes(self):
        try:
            string = "1) Neck Pain\n2) Backache\n"
            i = 3
            for item in self.list_of_indexes:
                string += f"{i}) {item[0]}\n"
                i += 1
            messagebox.showinfo("Indexes", string)
        except Exception as e:
            messagebox.showerror("Something went wrong", str(e))

    def add_indexes(self):
        try:
            prescription_to_be_added = self.prescription_text_area.get(1.0, END)
            index_to_be_added = self.issues_text_area.get(1.0, END).rstrip("\n")
            with open("indexes.txt", mode="a") as file:
                prescription_to_be_added = prescription_to_be_added.rstrip(prescription_to_be_added[-2])
                file.write("\n" + index_to_be_added + "|" + prescription_to_be_added)
            self.list_of_indexes = []
            self.define_indexes()
        except Exception as e:
            messagebox.showerror("Something went wrong", str(e))


prescription = Prescription()
