import pandas as pd
def UnpackPls(x):
    return x[2:-3]

#Creating dataframes and dictionaries
def Initilization(ONCPath, CPath):
    dfONC = pd.read_excel(UnpackPls(ONCPath), sheet_name="Ordenes_No_Colocadas.csv")
    dfC = pd.read_excel(UnpackPls(CPath), sheet_name="Hoja1")
    global dC
    dC = dfC.to_dict(orient = "index")
    global dONC
    dONC = dfONC.to_dict(orient = "index")

#ZoneDict Constructor (too dumb to dict comprehend)
def ZoneDictConstructor():
    global ZoneDict
    ZoneDict = {}
    for i in dC.keys():
        TempD = dC[i].copy()
        TempD.pop(" ")
        TempD.pop("CAMP.")
        TempD.pop("ÚLTIMO DIA EN BODEGA")
        TempD.pop("Unnamed: 6")
        ZoneDict.update({dC[i][" "]:TempD})

#Dict writer
def WriteThis():
    for i in dONC.keys():
        if dONC[i]["Zona"] == 1235:
            dONC[i].update({"CIERRE WEB": "", "FACTURACION": "", "DIAS DE REPARTO": ""})
        else:
            dONC[i].update(ZoneDict[dONC[i]["Zona"]])

#Parsing and saving
def SaveMe(OutPath):
    ColNames = dONC[0].keys()
    data = {i:list(dONC[i].values()) for i in dONC.keys()}
    dfOutput = pd.DataFrame.from_dict(data, orient = "index", columns = ColNames)
    writer = pd.ExcelWriter(OutPath+"/Output.xlsx", date_format = 'DD-MM-YYYY', datetime_format='DD-MM-YYYY', engine="xlsxwriter")
    dfOutput.to_excel(writer)
    writer.save()

if __name__ == "__main__":
    #Creating dataframes and dictionaries
    dfONC = pd.read_excel("D:\PyMyScripts\Room\ResourcesIGuess\Ordenes_No_Colocadas - 2021-07-02T080049.421.csv", sheet_name="Ordenes_No_Colocadas.csv")
    dfC = pd.read_excel("D:\PyMyScripts\Room\ResourcesIGuess\CAMPAÑA 12.xlsx", sheet_name="Hoja1")
    dC = dfC.to_dict(orient = "index")
    dONC = dfONC.to_dict(orient = "index")

    #ZoneDict Constructor (too dumb to dict comprehend)
    ZoneDict = {}
    for i in dC.keys():
        TempD = dC[i].copy()
        TempD.pop("ZONA")
        TempD.pop("CAMP.")
        TempD.pop("ÚLTIMO DIA EN BODEGA")
        #TempD.pop("Unnamed: 6")
        ZoneDict.update({dC[i]["ZONA"]:TempD})

    #Dict writer
    for i in dONC.keys():
        if dONC[i]["Zona"] == 1235:
            dONC[i].update({"CIERRE WEB": "", "FACTURACION": "", "DIAS DE REPARTO": ""})
        else:
            dONC[i].update(ZoneDict[dONC[i]["Zona"]])

    #Parsing and saving
    ColNames = dONC[0].keys()
    data = {i:list(dONC[i].values()) for i in dONC.keys()}
    dfOutput = pd.DataFrame.from_dict(data, orient = "index", columns = ColNames)
    writer = pd.ExcelWriter("D:\PyMyScripts\Room\ResourcesIGuess\Output.xlsx", date_format = 'DD-MM-YYYY', datetime_format='DD-MM-YYYY', engine="xlsxwriter")
    dfOutput.to_excel(writer)
    writer.save()