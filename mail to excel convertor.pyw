import pandas as pd
import customtkinter as Ctk
from customtkinter import filedialog
from openpyxl import load_workbook

EXCEL_PATH = ""
MAIL_PATHS = []


class Window(Ctk.CTk):
    def __init__(self):
        super().__init__()
        Ctk.set_default_color_theme("dark-blue")

        self.geometry("290x400")
        self.title("Mail to excel convertor")

        self.grid_rowconfigure((0, 1, 2, 3, 4), weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.main_button = Ctk.CTkButton(self,
                                        text="Prevod",
                                        fg_color="#E5B20B",
                                        hover_color="#E5870B",
                                        text_color="black",
                                        command=lambda: convert(EXCEL_PATH, MAIL_PATHS)
                                        ).grid(
                                        row=3,
                                        column=0,
                                        padx=20,
                                        pady=20,
                                        sticky="ew")

        self.mails_button = Ctk.CTkButton(self,
                                          text="vyberte maily",
                                          fg_color="#8CC6C9",
                                          hover_color="#8C97C9",  
                                          text_color="black",
                                          command=select_mails
                                          ).grid(
                                          row=0,
                                          column=0,
                                          padx=20,
                                          pady=20,
                                          sticky="ew"
                                          )

        self.table_button = Ctk.CTkButton(self,
                                          text="vyberte tabuľku",
                                          fg_color="#4DDD30",
                                          hover_color="#25AB25",
                                          text_color="black",
                                          command=select_table
                                          ).grid(
                                          row=1,
                                          column=0,
                                          padx=20,
                                          pady=20,
                                          sticky="ew"
                                          )
        
        self.textbox = Ctk.CTkTextbox(self,
                                     state="normal",
                                     height=100
                                     )
        self.textbox.grid(row=4, 
                          column=0,
                          padx=20,
                          pady=20,
                          sticky="ew"
                          )
        self.textbox.insert("0.0", "Zdravím :D\n"\
                                   "Vyberte si tabuľku a správy, ktoré by ste chceli pridať")
        self.textbox.configure(state="disabled")


def mail_parser(path: str) -> list[dict[str, str]]:
    students:list[dict[str, str]] = []
    result = {
        "mail na rodiča":    "<nie je uvedené>",
        "dátum narodenia":   "<nie je uvedené>",
        "meno":              "<nie je uvedené>",
        "rodič":             "<nie je uvedené>",
        "tel. čislo":        "<nie je uvedené>",
        "adresa":            "<nie je uvedené>",
        "deň":               "<nie je uvedené>",
        "škola":             "<nie je uvedené>",
        "trieda":            "<nie je uvedené>"}

    # custom constants, becasuse CASE won't let me use normal ones ;-;
    class con:
        PARENT_MAIL = 1
        NAME = 2
        PARENT_NAME = 3
        PHONE_NUMBER = 4
        ADRESS = 5
        DATE = 6
        SCHOOL = 7
        CLASS = 8

    flag = 0

    with open(path, encoding="utf-8") as mail:
        for line in mail:
            line = line.replace('\n', '')
            
            match flag:
                case con.PARENT_MAIL:
                    result["mail na rodiča"] = line

                case con.NAME:
                    info = line.split(' ')
                    if "." in info[-1]:
                        result["dátum narodenia"] = info.pop()
                    result["meno"] = ' '.join(info)

                case con.PARENT_NAME:
                    if line != "":
                        result["rodič"] = line

                case con.PHONE_NUMBER:
                    if line != "":
                        result["tel. čislo"] = line

                case con.ADRESS:                
                    if line != "":
                        result["adresa"] = line
                
                case con.DATE:
                    if line != "":
                        result["deň"] = line\
                
                case con.SCHOOL:
                    if line != "":
                        result["škola"] = line
                case con.CLASS:
                    if line != "":
                        result["trieda"] = line

            flag = 0

            match line:
                case "Z":
                    flag = con.PARENT_MAIL
                case "Meno, priezvisko a dátum narodenia (dd mm rrrr) dieťaťa":
                    flag = con.NAME
                case "Meno a priezvisko zákonného zástupcu dieťaťa":
                    flag = con.PARENT_NAME
                case "Telefónne číslo na zákonného zástupcu dieťaťa - ktoré používate pre whats app aplikáciu":
                    flag = con.PHONE_NUMBER
                case "Adresa zákonného zástupcu dieťaťa":
                    flag = con.ADRESS
                case "Vyberte si deň kurzu":
                    flag = con.DATE
                case "Názov školy":
                    flag = con.SCHOOL
                case "Trieda (príklad: 3 c)":
                    flag = con.CLASS
                case "(www.inkaart.sk, WebHouse)":
                    students.append(result.copy())
                    result = {
                        "mail na rodiča":    "<nie je uvedené>",
                        "dátum narodenia":   "<nie je uvedené>",
                        "meno":              "<nie je uvedené>",
                        "rodič":             "<nie je uvedené>",
                        "tel. čislo":        "<nie je uvedené>",
                        "adresa":            "<nie je uvedené>",
                        "deň":               "<nie je uvedené>",
                        "škola":             "<nie je uvedené>",
                        "trieda":            "<nie je uvedené>"}

    return students


def convert(path: str, mail_paths: list[str]) -> None:

    if path == "" or mail_paths == []:
        window.textbox.configure(state="normal")
        window.textbox.delete("1.0", "end")
        window.textbox.insert("1.0", "Chyba: neboli zvolené súbory")
        return
        
    # open for pandas
    table = pd.read_excel(path)

    # for every mail
    for mail_path in mail_paths:
        
        # parse the email and append every student from it to the sheet
        for student in mail_parser(mail_path):
            new_data = pd.DataFrame([student])
            table = pd.concat([table, new_data], ignore_index=True)

    # save update
    table.to_excel(path, index=False)

    # Open with openpyxl
    wb = load_workbook(path)
    ws = wb.active

    # Adjust column widths
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # A, B, C...
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = max_length + 2
        ws.column_dimensions[column].width = adjusted_width

    # save edited file
    wb.save(path)

    window.textbox.configure(state="normal")
    window.textbox.delete("1.0", "end")
    window.textbox.insert("1.0", "Správy boli úspešne pridané do tabuľky\n" \
                                 "Teraz môžete zavrieť toto okno :D")


def select_mails() -> None:
    paths = filedialog.askopenfilenames(
        title="Select files",
        filetypes=(("Mail files", "*.txt"), ("All files", "*.*")))

    global MAIL_PATHS
    MAIL_PATHS = paths


def select_table() -> None:
    path = filedialog.askopenfilename(
        title="Select a file",
        filetypes=(("Excel file", "*.xlsx"), ("All files", "*.*")))

    global EXCEL_PATH
    EXCEL_PATH = path


window = Window()
window.mainloop()
