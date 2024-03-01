from openpyxl import load_workbook
import customtkinter as ctk
from PIL import Image, ImageDraw, ImageFont
import textwrap
from datetime import datetime
import locale
from os import startfile

wb = load_workbook(
    "\\\\CIM-NAS\\Servidor_Recursos_Humanos\\RECURSOS HUMANOS 2023\\CREDENCIALES\\Informacion Empleados.xlsx"
)
ws = wb.active


#  ---Principal function that executes on click
def on_button_click():
    searchNumber()


def searchNumber():
    try:
        value = int(text_box.get())
    except ValueError:
        error.configure(text="Ingrese el No. de empleado")
        return
    column = "A"
    row = None
    for cell in ws[column]:
        if cell.value == value:
            row = cell.row
            searchData(row)
    if not row:
        error.configure(text="No se encontro empleado")


def searchData(row):
    number = str(ws[f"A{row}"].value)
    name = str(ws[f"B{row}"].value)
    NSS = str(ws[f"C{row}"].value)
    CURP = str(ws[f"D{row}"].value)
    RFC = str(ws[f"E{row}"].value)
    blood = str(ws[f"F{row}"].value)
    account = str(ws[f"G{row}"].value)
    emergencyName = str(ws[f"H{row}"].value)
    emergencyNumber = str(ws[f"I{row}"].value)
    date = str(ws[f"L{row}"].value)
    occupation = str(ws[f"J{row}"].value)
    clasification = str(ws[f"K{row}"].value)

    locale.setlocale(locale.LC_TIME, "es_ES")
    date = datetime.strptime(date, "%Y-%m-%d %H:%M:%S")
    date = date.strftime("%d/%B/%Y")

    if clasification == "N/A" or clasification == "":
        job = occupation
    else:
        job = occupation + " " + clasification

    generateImage(
        number,
        name,
        NSS,
        CURP,
        RFC,
        account,
        emergencyName,
        emergencyNumber,
        blood,
        job,
        date,
    )


def generateImage(
    number,
    name,
    NSS,
    CURP,
    RFC,
    account,
    emergencyName,
    emergencyNumber,
    blood,
    job,
    date,
):
    error.configure(text="")
    # resources
    font1 = ImageFont.truetype("arial.ttf", 80)
    font2 = ImageFont.truetype("arialbd.ttf", 100)
    font3 = ImageFont.truetype("arial.ttf", 60)
    font4 = ImageFont.truetype("arialbd.ttf", 60)
    font5 = ImageFont.truetype("arial.ttf", 40)
    black = (0, 0, 0)
    # ---first image
    image = Image.open("img\\front.png").convert("RGBA")
    draw = ImageDraw.Draw(image)
    try:
        photo = Image.open(
            f"\\\\CIM-NAS\\Servidor_Recursos_Humanos\\RECURSOS HUMANOS 2023\\CREDENCIALES\\FOTOS PARA CREDENCIALES\\Foto Sin fondo\\{number}.png"
        )
    except FileNotFoundError:
        try:
            photo = Image.open(
                f"\\\\CIM-NAS\\Servidor_Recursos_Humanos\\RECURSOS HUMANOS 2023\\CREDENCIALES\\FOTOS PARA CREDENCIALES\\Foto Sin fondo\\{name}.png"
            )
        except FileNotFoundError:
            return error.configure(text="No se encontro la foto")

    original_width, original_height = photo.size
    new_height = 800
    new_width = int(new_height * original_width / original_height)
    photo = photo.resize((new_width, new_height))
    centerx = int((image.width - new_width) / 2)
    image.paste(photo, (centerx, 450), photo)

    lines = textwrap.wrap(name, width=20)
    y_text = 1280
    for line in lines:
        centerx = (image.width - (draw.textsize(line, font=font2)[0])) / 2
        draw.text((centerx, y_text), line, font=font2, fill=black)
        y_text = 1380

    centerx = (image.width - (draw.textsize(job, font=font3)[0])) / 2
    draw.text((centerx, 1550), job, font=font3, fill=black)

    centerx = (image.width - (draw.textsize(number, font=font1)[0])) / 2
    draw.text((centerx, 1650), number, font=font1, fill=black)

    draw.text((40, 1925), f"Ingreso: {date}", font=font5, fill=black)
    image = image.convert("RGB")
    image.save("img\\front.jpg")

    # ---second image
    image = Image.open("img\\back.png")
    draw = ImageDraw.Draw(image)

    draw.text((100, 980), "CTA:", font=font4, fill=black)
    draw.text((100, 1070), "NSS:", font=font4, fill=black)
    draw.text((100, 1160), "T.SANGRE:", font=font4, fill=black)
    draw.text((100, 1250), "RFC:", font=font4, fill=black)
    draw.text((100, 1340), "CURP:", font=font4, fill=black)

    draw.text((250, 980), account, font=font3, fill=black)
    draw.text((255, 1070), NSS, font=font3, fill=black)
    draw.text((447, 1160), blood, font=font3, fill=black)
    draw.text((260, 1250), RFC, font=font3, fill=black)
    draw.text((310, 1340), CURP, font=font3, fill=black)

    centerx = (
        image.width - (draw.textsize("En caso de emergencia llamar a:", font=font3)[0])
    ) / 2
    draw.text(
        (centerx, 1540), "En caso de emergencia llamar a:", font=font3, fill=black
    )

    centerx = (
        image.width
        - (draw.textsize(f"{emergencyName} / {emergencyNumber}", font=font4)[0])
    ) / 2
    draw.text(
        (centerx, 1620), f"{emergencyName} / {emergencyNumber}", font=font4, fill=black
    )
    draw.text((780, 1050), "VIGENCIA", font=font1, fill=(255, 0, 0))

    image = image.convert("RGB")
    image.save("img\\back.jpg")

    startfile("img\\front.jpg")
    startfile("img\\back.jpg")


# ---Make the windows and widgets
app = ctk.CTk()
app.geometry("300x200")
app.title("Credenciales")
app.resizable(False, False)
ctk.set_appearance_mode("light")

frame = ctk.CTkFrame(app, bg_color="#FEFAFA", fg_color="#FEFAFA")
frame.pack(fill="both", expand=True)

label = ctk.CTkLabel(master=frame, text="Numero de empleado:")
label.pack(pady=10)

text_box = ctk.CTkEntry(master=frame)
text_box.pack(pady=10)

button = ctk.CTkButton(master=frame, text="Buscar", command=on_button_click)
button.configure(fg_color="#B42725", hover_color="#871e1c")
button.pack(pady=10)

error = ctk.CTkLabel(master=frame, text="")
error.pack(pady=10)

app.mainloop()
