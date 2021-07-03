from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from PIL import Image
from PIL import ImageTk

#My import
import ExcelReadWriteFechas as magic

#Callback functions that get file path, when the button request them
def PathCampañaPls():
    Temp = filedialog.askopenfilenames()
    PCampaña.set(Temp)

def PathOrdenesPls():
    Temp = filedialog.askopenfilenames()
    POrdenes.set(Temp)

#Magic happens
def ExcelCallIn():
    Temp = filedialog.askdirectory()
    POutput.set(Temp)
    magic.Initilization(POrdenes.get(), PCampaña.get())
    magic.ZoneDictConstructor()
    magic.WriteThis()
    magic.SaveMe(POutput.get())

#Callback functions that shows a thumbs up image
#and manege some logic to activate the confirm button,
#I could probably do this stuff inside the path functions
def ShowCampañaImage(*args):
    ttk.Label(mainframe, image=ImageDisplayable).grid(column=0, row=0)
    CFlag.set(True)
    Flag.set(CFlag.get() and OFlag.get())

def ShowOrdenesImage(*args):
    ttk.Label(mainframe, image=ImageDisplayable).grid(column=2, row=0)
    OFlag.set(True)
    Flag.set(CFlag.get() and OFlag.get())

def ConfirmButtonUpdate(*args):
    if Flag.get():
        ConfirmButton.state(["!disabled"])

#Initialization of the root and some variables
root = Tk()
root.title("Manejador de fechas de campaña")
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

PCampaña = StringVar()
POrdenes = StringVar()
POutput = StringVar()

CFlag = BooleanVar()
OFlag = BooleanVar()
Flag = BooleanVar()

#Mainframe and widgets initialization
mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
mainframe.columnconfigure([0, 1, 2], weight=1)
mainframe.rowconfigure([0, 1], weight=1)

ConfirmButton = ttk.Button(mainframe, text="Confirmar", command=ExcelCallIn, state="disabled")
ConfirmButton.grid(column=1, row=1)

SelectCampañaButton = ttk.Button(mainframe, text="Selecciona Archivo de Campaña", command=PathCampañaPls)
SelectCampañaButton.grid(column=0, row=1)
SelectOrdenesButton = ttk.Button(mainframe, text="Selecciona Archivo de Ordenes", command=PathOrdenesPls)
SelectOrdenesButton.grid(column=2, row=1)

#Image rezise
#OriginalImage = Image.open("¡estoy bien!.png") 
#I need to change this part, to not depend too much in the exact path to the image.
OriginalImage = Image.open("D:\PyMyScripts\Room\ResourcesIGuess\¡estoy bien!.png")
ReziseImage = OriginalImage.resize((120, 120))
ImageDisplayable = ImageTk.PhotoImage(ReziseImage)

#Adding tracers to call for the image display and the button activation
PCampaña.trace_add("write", ShowCampañaImage)
POrdenes.trace_add("write", ShowOrdenesImage)
Flag.trace_add("write", ConfirmButtonUpdate)

root.mainloop()