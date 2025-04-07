#Requires AutoHotkey v2.0

global datos := Map()  ; Crear un diccionario global
global EditControls := Map()  ; Guarda los controles Edit para acceder después
global casesObj := Map()
global NmOfSteps := 18
global fileDir1 := A_Desktop "\CasesJSON"  ; Solicitar la ruta del directorio




#Include  TeamViewer.ahk  
#Include Json.ahk
Persistent 1  ; Establece el script como persistente

; ---------------------------------------------------------------------------------------------------------------------------------

SkrvGui := Gui(, "SKRIVE")
SkrvGui.Opt("+Resize +MinSize400x300 +MaximizeBox +MinimizeBox") ; Hace la ventana redimensionable
SkrvGui.SetFont("s12")  

tab := SkrvGui.Add("Tab3", "x10 y10 w900 h600", ["Information","Remote Session", "Add Info", "Email", "2nd Line - Call Back"])

SkrvGui.OnEvent("Size", (*) => AutoResize())  ; Evento para ajustar tamaño dinámicamente
SkrWidth := 1250
SkrHeigh := 750 
BtnHeigh:=30
btnSze:= 200

SkrvGui.Show("w" SkrWidth " h" SkrHeigh)  ; Esto es correcto


AutoResize(*) {
    global SkrvGui, tab
    SkrvGui.GetPos(&x, &y, &w, &h)  ; Obtiene el tamaño actual de la ventana
    tab.Move(10, 10, w - 20, h - 40)  ; Ajusta tamaño y posición
}

InputLine(Name, x?, xedit?, widthedit?, yOffset?, btnSze?, BtnHeigh?, buttomDefltPaste?, editYplace?, editHeigh?) {
    Name2Text := Name

    ; Valores predeterminados
    x := IsSet(x) ? x : 50
    xedit := IsSet(xedit) ? xedit : 260
    widthedit := IsSet(widthedit) ? widthedit : 770
    yOffset := IsSet(yOffset) ? yOffset : 100
    btnSze := IsSet(btnSze) ? btnSze : 80
    BtnHeigh := IsSet(BtnHeigh) ? BtnHeigh : 30
    editYplace := IsSet(editYplace) ? editYplace : yOffset
    editHeigh := IsSet(editHeigh) ? editHeigh : BtnHeigh

    btn := SkrvGui.Add("Button", Format("x{} y{} w{} h{}", x, yOffset, btnSze, BtnHeigh), Name2Text)
    edit := SkrvGui.Add("Edit", Format("x{} y{} w{} h{}", xedit, editYplace, widthedit, editHeigh), "")
    if buttomDefltPaste == true {
        btn.OnEvent("Click", (*) => edit.Value := A_Clipboard)  ; Clic izquierdo: pega desde el portapapeles
        btn.OnEvent("ContextMenu", (*) =>Altclick() )    ; Clic derecho: limpia el Edit
        Altclick() {
            if GetKeyState("Alt"){  ; Verifica si Alt está presionado
                    A_Clipboard := edit.Value
                    return
            }
            else { 
                edit.Value := ""
                return
            }
        }
    } 
    if buttomDefltPaste == false {
        btn.OnEvent("Click", (*) => edit.Value := A_Clipboard)  ; Clic izquierdo: pega desde el portapapeles
        btn.OnEvent("ContextMenu", (*) => Ctrl_6())  ; Clic izquierdo: limpia el Edit
    }
    if buttomDefltPaste == "paste" {
        btn.OnEvent("Click", (*) =>   A_Clipboard:= edit.Value )  ; Clic izquierdo: pega desde el portapapeles
        btn.OnEvent("ContextMenu", (*) =>   A_Clipboard:= edit.Value )  ; Clic izquierdo: pega desde el portapapeles
        
    }
    ; Guardar la referencia del Edit
    global EditControls
    EditControls[Name2Text] := edit  

    ; Capturar cambios en el input
    global datos
    datos[Name2Text] := ""  
    edit.OnEvent("Change", (*) => datos[Name2Text] := edit.Value)
}

SoftwareVersionSelect(Name, yOffset, btnSze, BtnHeigh) {
    Name2Text := Name
    btnSoftware := SkrvGui.Add("Button", Format("x50 y{} w{} h{}", yOffset, btnSze, BtnHeigh), Name2Text)
    editSoftware := SkrvGui.Add("Edit", Format("x260 y{} w770 h{}", yOffset, BtnHeigh), "")
    
    btnSoftware.OnEvent("Click", (*) => SoftwareVersionGUI(editSoftware, Name2Text))
    ; Guardar la referencia del Edit
    global EditControls
    EditControls[Name2Text] := editSoftware  

    global datos
    datos[Name2Text] := editSoftware.Value
    UpdateDataFromEdits() ; 💡 Refresca `datos` con los valores actuales de los Edits

    editSoftware.OnEvent("Change", (*) => datos[Name2Text] := editSoftware.Value)
}

SoftwareVersionGUI(editSoftware, Name2Text) { 
        global datos, EditControls  ; Asegurar acceso a los datos y los Edit


    SoftwGui := Gui()
    SoftwGui.Opt("+AlwaysOnTop")
    SoftwGui.SetFont("s10")

    SoftwGui.Add("Text", "x10 y10", "Select Software Versions:")
    
    SoftwGui.Add("Text", "x10 y30", "Module:")
    comboBox1 := SoftwGui.Add("ComboBox", "x10 y50 w150", ["Unite", "TRIOS Module", "Dental Desktop", "Model Builder", "Automate"])
    
    SoftwGui.Add("Text", "x180 y30", "Version:")
    comboBox2 := SoftwGui.Add("ComboBox", "x180 y50 w150", ["1.7.83.0",  "1.8.8.0","1.7.8.1", "1.8.5.1","1.7.82.5"])

    ; Valores por defecto
    comboBox1.Text := "Dental Desktop"
    comboBox2.Text := "1.7.83.0"

    btnSelect := SoftwGui.Add("Button", "x10 y120 w320", "Select")
    
    btnSelect.OnEvent("Click", (*) => (
        editSoftware.Value := comboBox1.Text " version " comboBox2.Text,  ; Actualiza el Edit
        datos[Name2Text] := editSoftware.Value,  ; 🔥 También actualiza el Map()
        SoftwGui.Destroy()
    ))
    
    SoftwGui.Show()
}

OutputLine(Name,xOffset, yOffset, Insize ,btnSze,BtnHeigh){
    ; Name2Text := Format("{}",Name)
    Name2Text := Name
    Outbtn := SkrvGui.Add("Button", Format("x{} y{} w{} h{}",xOffset, yOffset,btnSze,BtnHeigh), Name2Text)
    outedit := SkrvGui.Add("Edit", Format("x{} y{} w{} h{}",xOffset, yOffset+BtnHeigh+8,Insize,BtnHeigh*2), "")
    ; Guardar la referencia del Edit
    global EditControls
    EditControls[Name2Text] := outedit  
    global datos
    datos[Name2Text] := outedit.Value
    outedit.OnEvent("Change", (*) => datos[Name2Text] := outedit.Value)
    UpdateDataFromEdits() ; 💡 Refresca `datos` con los valores actuales de los Edits
    Outbtn.OnEvent("ContextMenu", (*) => outedit.Value := "")       ; Clic derecho: limpia el Edit
    Outbtn.OnEvent("Click", (*) => A_Clipboard := "RC: " datos["RC"] "`n" "S: " datos["Solution"]  )       ; Clic derecho: limpia el Edit
    

}

Hotkeys(Name, yOffset){
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit

    ; Name2Text := Format("{}",Name)
    Name2Text := Name
    btnSze:= 200
    btnHot := SkrvGui.Add("Button", Format("x50 y{} w{} h{}", yOffset,btnSze,BtnHeigh), Name2Text)
    ; edit := SkrvGui.Add("Edit", Format("x260 y{} w700 h{}", yOffset), "")
    ; btnHot.OnEvent("Click", (*) => Ctrl_y())
    btnHot.OnEvent("ContextMenu", (*) => Ctrl_a() )
    btnHot.OnEvent("Click", (*) => altcheck())
    altcheck() {
        if GetKeyState("Alt"){  ; Verifica si Alt está presionado
                Ctrl_2()
                return
        }
        else { 
            Ctrl_y()
            return
        }
    }


    

}

IntPhBttm(IntBool) {
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit
    

    issue := datos.Has("Issue") ? datos["Issue"] : "No issue"
    softwareVersion := datos.Has("Software Version") ? datos["Software Version"] : "Unknown version"
    

    today := A_Now  ; Obtiene la fecha y hora actual en formato AAAAMMDDHHMMSS
    formattedDate := FormatTime(today, "yyyyMMdd")  ; Formatea la fecha

    PH := "PHONECALL " formattedDate
    Int := "INT " formattedDate

    ; Obtiene valores, si no existen usa "N/A"
    name := datos.Get("Name", "N/A")
    description := issue " on " softwareVersion
    phone := datos.Get("&Phone", "N/A")
    email := datos.Get("&Email", "N/A")
    dongle := datos.Get("&Dongle", "N/A")
    tv_id := datos.Get("TV ID", "N/A")
    tv_ps := datos.Get("TV PSS", "N/A")
    HJ := datos.Get("HJ", "N/A")

    IntNote := Format("This HelpJuices was used `n{}", HJ)

    PHNote := Format(
        "{} report {}`n"
        "Name: {}`n"
        "Phone N: {}`n"
        "Email: {}`n"
        "Dongle N: {}`n"
        "TV ID: {}`n"
        "TV PS: {}",
        name, description, name, phone, email, dongle, tv_id, tv_ps
    )
    ; Copia primero el título (PHONECALL o INT)
    A_Clipboard := (IntBool ? PHNote : IntNote)
    
    ; Espera un momento para asegurar que se copie
    Sleep(500)

    ; Luego copia la descripción (PHNote o IntNote)
    A_Clipboard := (IntBool ? PH : Int)


    Sleep(500)


    return
}

C1stAdd(C1Bool) {
        global datos, EditControls  ; Asegurar acceso a los datos y los Edit

    issue := datos.Has("Issue") ? datos["Issue"] : "No issue"
    softwareVersion := datos.Has("Software Version") ? datos["Software Version"] : "Unknown version"
    UpdateDataFromEdits() ; 💡 Refresca `datos` con los valores actuales de los Edits


    C1 := "C1st"

    ; Obtiene valores, si no existen usa "N/A"
    Cname := datos.Get("&Company Name", "N/A")
    SID := datos.Get("SID", "N/A")
    Descrp := datos.Get("Descrip&tion", "N/A")
    
    Note := Format("{} / SID: {} / {}" ,  Cname, SID, Descrp)

    C1Note :=  Format("{} / {}" , C1, Note)

    

    ; Copia primero el título (PHONECALL o INT)
    A_Clipboard := (C1Bool ? C1Note : Note)
    
    


    return
}

UpdateEditsFromData() {
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit

    for key, value in datos {
        if (EditControls.Has(key)) {  ; Usar Has() en lugar de HasKey()
            EditControls[key].Value := value
        }
    }
}

PrintMap(m) {
    if Type(m) != "Map" {
        MsgBox("El valor proporcionado no es un Map.")
        return
    }

    output := ""
    for clave, valor in m {
        output .= clave ": " valor "`n"
    }

    MsgBox(output)
}



; ---------------------------------------------------------------------------------------------------------------------------------

; PESTAÑA 1 (General)
tab.UseTab(1) ; (Information)
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit

imagePath := A_ScriptDir . "/LogoSmall.png"
scale := 2.65
wd:= 843
hi := 559
imgWidth1 := wd / scale
imgHeight1 := hi / scale

; Obtener el tamaño actual de la ventana para posicionar la imagen
SkrvGui.GetPos(&winX1, &winY1, &winWidth1, &winHeight1)
imgX1 := 1050
imgY1 := 550

imgControl := SkrvGui.Add("Picture", Format("x{} y{} w{} h{}", imgX1, imgY1, imgWidth1, imgHeight1), imagePath)



y := 50 ; Posición inicial en Y
x:= 260
spacing := 35 ; Espaciado entre elementos
Insize:=390

SkrvGui.Add("Text",)
tvsize := (750/2)-(spacing*2)
Complaints := Map() 
Request := Map()
Miscellaneous := Map()
CasesfolderPaths:= Map()

Caseopt := ["Complaints", "Request", "Miscellaneous"]
CaseTP := Map()



CaseTpC :=["New 3shape account/ New user", 
"Specific Case", 
"No reproducible",
"Country Network Restrictions",
"Slow Performance",
"Company ID Token",
"Cloud Storage Lost",
"Case not able to be sent (Case Corruption)",
"Color Calibration Expirated",
"Server Issue Ocurred",
"Black lines while scanning",
"Unable to Reach",
"TP Link Issue",
"Reimport Streams",
"Calibration Code 11",
"Server Issue Ocurred / Running on PC {0}",
"Unite Server not found",
"Send step is not present or greyed-out in the workflow bar",
"Slow Performance Trios 3-5 (BM-6026) IA",
"Configuration Download not working",
"Reimport Streams due to corruption after receiving case on a lab integration",
"Disabled Integration on Comunicate (Invisalign)",
"Unite not being updated automatically to Unite III on a client PC",
"EM: (Error code:10) when performing calibration on TRIOS 3",
"Abutment not selected EM",
"WIFI networks are not showing on a move +",
"Migration unsuccessfull",
"Clear Connect conection"
]
CaseTpR :=["Reimport Streams",
"Remote Routine Check",
"Unite Setup",
"Lab connection on Communicate",
"Unite Update (Install)",
"No access to the mail",
"Set Up 3Shape Acount",
"Request to have help to connect a TRIOS core scanner for the first time (replacement/loan)",
"How to export stl files in TRIOS software 1.x.x.x",
] 
CaseTpM :=["Customer Called for X reason"]


y := 50

; Primer ComboBox con opciones

comboBoxCaseSelect1 := SkrvGui.Add("ComboBox", Format("x{} y{} w140", 50 , y), Caseopt), y  ; Aumentar Y para el siguiente control

; Segundo ComboBox (empieza vacío)
comboBoxCaseSelect2 := SkrvGui.Add("ComboBox", Format("x{} y{} w570", 200 , y), ["El box de la izquierda esta vacio", "Por eso nada aparecera Aqui"]), y
; Valores por defecto

; Al cambiar la selección del primer ComboBox, se actualiza el segundo
comboBoxCaseSelect1.OnEvent("Change", (*) => check1())

check1(){
    global CSTP := Caseopt
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit
    for CaseIndex, CaseO in Caseopt {
        folderPath := A_WorkingDir "\" Caseopt[CaseIndex]
        fileList := []

        loop files folderPath "\*.*"
        {
            fileList.Push(A_LoopFileName)
        }

        ; Guardar la lista completa de archivos de esta carpeta
        CaseTP[CaseIndex] := fileList

        ; ; Mostrar los archivos de esta carpeta
        ; for fileIndex, fileName in fileList {
        ;     MsgBox "Carpeta " CaseIndex " - Archivo " fileIndex ": " fileName
        ; }
    }
    selectedText := comboBoxCaseSelect1.Text

    ; Switch con tus valores originales
    Switch selectedText {
        Case Caseopt[1]:
            CSTP := CaseTP[1]
        Case Caseopt[2]:
            CSTP := CaseTP[2]
        Case Caseopt[3]:
            CSTP := CaseTP[3]
        Default:
            CSTP := ["El box de la izquierda esta vacio", "Por eso nada aparecera Aqui"]
    }

    comboBoxCaseSelect2.Delete()
    comboBoxCaseSelect2.Add(CSTP)
    comboBoxCaseSelect2.Text := ""  ; Limpia selección previa
}



LoadExistingcase := SkrvGui.Add("Button", Format("x780 y{} w{} h{}", y, 250, 28), "Save - Load Existing Template")
y += spacing

LoadExistingcase.OnEvent("ContextMenu", (*) => eval() )
LoadExistingcase.OnEvent("Click", (*) =>  savecase())

savecase(){
    UpdateDataFromEdits()
    Caseoption1 := comboBoxCaseSelect1.Text
    Casetype1 := comboBoxCaseSelect2.Text
    if (Casetype1 == "" or Caseoption1 == ""){
        MsgBox "Coloca un Tipo y Nombre de caso para guardar"
        return
    }
    filePathOptions1 := A_WorkingDir "\" Caseoption1 
    savedpath := filePathOptions1 "\" Casetype1 ".json"
    if !DirExist(filePathOptions1) {
        DirCreate(filePathOptions1)
    }
    Datos2SavedCases := ["HJ","Probing Questions (Add Info)","Issue","RC","Solution",  "EmailInputEdit"]
    global NmOfSteps
    Loop NmOfSteps {
        Datos2SavedCases.Push(Format("Step {}", A_Index))
    }
    ; Crear un nuevo Map con solo los que están en Datos2SavedCases
    dats2 := Map()

    for _, key in Datos2SavedCases {
        if datos.Has(key)
            dats2[key] := datos[key]
    }
    
    
           ; Guardar el JSON
           ; AGREGAR SOLO LOS CAMPOS QUE SON NECESARIOS
        jsonText := Jxon_Dump(dats2, 4)
        FileAppend(jsonText, savedpath, "UTF-8")

        MsgBox("Guardado como:`n" savedpath)
    }

eval(){
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit
    
    Caseoption := comboBoxCaseSelect1.Text
    Casetype := comboBoxCaseSelect2.Text
    if (Casetype == "" or Caseoption == ""){
        MsgBox "Coloca un Tipo y Nombre de caso para cargar"
        return
    }
    dta1 := LoadMapFromFileEXISTINGCASE(Caseoption,Casetype) 
        if (Type(dta1) = "Map") {
            datos := dta1 ; Sobreescribe el mapa global con los datos cargados
            Sleep(500)
            UpdateEditsFromData() ; Llama a la función que actualiza los edits
        } else {
            MsgBox("Error: El archivo JSON no es válido.")
        }
        return
}

LoadMapFromFileEXISTINGCASE(Caseoption, Casetype) {
    filePathOptions := A_WorkingDir "\" Caseoption 
    filePath1 := filePathOptions "\" Casetype 

    ; Crear carpeta si no existe
    if !DirExist(filePathOptions) {
        DirCreate(filePathOptions)
    }

    if !FileExist(filePath1) {
        MsgBox("Archivo no encontrado: " filePath1)
        return
    }

    
    

    fileContent := FileRead(filePath1, "UTF-8")
    MsgBox("Cargado desde el archivo: " filePath1)

    return fileLoaded := Jxon_Load(&fileContent)
}


InputLine("TV ID",,(btnSze/2)+55,tvsize, y,btnSze/2,BtnHeigh, false), y
InputLine("TV PSS",465,570,tvsize, y,btnSze/2,BtnHeigh, false), y 
Int_PhBtt := SkrvGui.Add("Button", Format("x880 y{} w{} h{}", y, btnSze-100, BtnHeigh), "INT-PH"), y 
Int_PhBtt.OnEvent("Click", (*) => altcheck2())
    altcheck2() {
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit

    UpdateDataFromEdits() ; 💡 Refresca `datos` con los valores actuales de los Edits
    if GetKeyState("Alt") {  ; Verifica si Alt está presionado
        IntPhBttm(false)
        Sleep(1000)
        A_Clipboard := "Logs and images are here"
        return
    } else { 
        IntPhBttm(false)
        return
    }
}
UpdateDataFromEdits() {
    global datos, EditControls

    for key, control in EditControls {
        datos[key] := control.Value  ; Actualiza los valores de `datos`
    }
}



Int_PhBtt.OnEvent("ContextMenu", (*) => IntPhBttm(true))

c1stbttn := SkrvGui.Add("Button", Format("x985 y{} w{} h{}", y, 45, BtnHeigh), "C1st"), y += spacing
c1stbttn.OnEvent("Click", (*) => C1stAdd(false))
c1stbttn.OnEvent("ContextMenu", (*) => C1stAdd(true))

InputLine("Name",,,, y,btnSze,BtnHeigh, true,,), y += spacing
InputLine("&Phone",,,, y,btnSze,BtnHeigh, true,,), y += spacing
InputLine("&Email",,,, y,btnSze,BtnHeigh, true,,), y += spacing
InputLine("&Company Name",,,, y,btnSze,BtnHeigh, true,,), y += spacing
SoftwareVersionSelect("Software Version", y,btnSze,BtnHeigh), y += spacing

InputLine("&Dongle",,,, y,btnSze,BtnHeigh, true,,), y += spacing
InputLine("SID",,,, y,btnSze,BtnHeigh, true,,), y += spacing
InputLine("GUI",,,, y,btnSze,BtnHeigh, true,,), y += spacing
InputLine("Scanner S/N",,,, y,btnSze,BtnHeigh, true,,), y += spacing
InputLine("PC ID",,,, y,btnSze,BtnHeigh, true,,), y += spacing
InputLine("Case Number",,,, y,btnSze,BtnHeigh, true,,), y += spacing



InputLine("HJ",,,, y,btnSze,BtnHeigh, true,,), y += spacing
InputLine("Survey",,,, y,btnSze,BtnHeigh, true,,), y += spacing
InputLine("RC",,,, y,btnSze,BtnHeigh, true,,), y += spacing
; ---------------------------------------------------------------------------------------------------------------------------------
descripton:= SkrvGui.Add("Button", Format("x50 y{} w{} h{}", y,btnSze,BtnHeigh), "Descrip&tion"),
descripton.OnEvent("ContextMenu", (*) => DescriptionGUI(true))
descripton.OnEvent("Click", (*) => DescriptionGUI(false))
    

DescriptionGUI(bool) {
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit

    DescripGui := Gui()
    DescripGui.Opt("+AlwaysOnTop")
    DescripGui.SetFont("s10")

    ; Verifica si las claves existen antes de usarlas
    issue := datos.Has("Issue") ? datos["Issue"] : "No issue"
    softwareVersion := datos.Has("Software Version") ? datos["Software Version"] : "Unknown version"


    ; Agrega el texto correctamente formateado
    Descr := DescripGui.Add("Text", "x10 y10", "The description is:`n" issue " on " softwareVersion)
    btnCopy := DescripGui.Add("Button", "x10 y150 w320", "Copy")
    
    if bool == false{
        A_Clipboard := issue " on " softwareVersion
        return
    }
    ; Botón para copiar la información
    if bool == true {
        DescripGui.Show()
        btnCopy.OnEvent("Click", (*) => (
            A_Clipboard := issue " on " softwareVersion,  ; Copia al portapapeles
            DescripGui.Destroy()  )
        )
        return
    }



    
}
; ---------------------------------------------------------------------------------------------------------------------------------
OutputLine("Issue",x,y, Insize ,btnSze,BtnHeigh), 
OutputLine("Solution",(x+Insize +10),y, Insize ,btnSze,BtnHeigh), y += spacing
Hotkeys("Hotkeys", y), y


BtnSaveInfo := SkrvGui.Add("Button", "x50 y680 w150 h40", "&Save Info-Load Info")
clear := SkrvGui.Add("Button", " x205 y680  w45 h30", "Clear")

clear.OnEvent("Click", (*) => ClearAll())

ClearAll() {
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit

    respuesta := MsgBox("¿Está seguro de que desea borrar todos los campos?", "Borrar Datos", "YesNo")
    if (respuesta = "No")  ; Si el usuario elige "No", cancela la acción
        return
    Sleep(800)
    for key in datos {
        if (EditControls.Has(key)) {  
            EditControls[key].Value := ""  
        }
        datos[key] := ""  
    }

    UpdateEditsFromData()  
}




BtnSaveInfo.OnEvent("ContextMenu", (*) => SaveBttm(true,fileDir1))
BtnSaveInfo.OnEvent("Click", (*) => altcheck3())

altcheck3() {
    global fileDir1
    if GetKeyState("Alt") {  ; Verifica si Alt está presionado
        ; Permite al usuario escribir el nombre del archivo y carpeta donde guardar
        filePath := FileSelect("D", , "Guardar archivo JSON en carpeta", )
        
        ; Si el usuario canceló, termina
        if (!filePath) {
            MsgBox("Guardado cancelado.")
            return
        }
        SaveMapSmart(filePath)  ; Llama a tu función de guardado
    } else {
        SaveBttm(false,fileDir1)  ; Tu función alternativa si no se presionó Alt
    }
}

; Función para abrir una nueva ventana
SaveBttm(SvBool,fileDir) {

        global datos, EditControls  ; Asegurar acceso a los datos y los Edit

    C1stAdd(false) ; Actualiza el mapa con la descripción
    
    ; Recorrer el mapa y mostrar los valores
    if SvBool==false{
        SaveMapSmart(fileDir)  ; Guardar el mapa en un archivo JSON
    }


    if SvBool==true{
        caseNumber := datos.Get("Case Number")

        if caseNumber != "" {
            filePath := fileDir "\" caseNumber ".json"
        } else {
            filePath := FileSelect(, fileDir, , "JSON (*.json)")
        }

        if (!filePath) {
            MsgBox("Recuperación cancelada.")
            return
        }

        dta := LoadMapFromFile(filePath) 
        if (Type(dta) = "Map") {
            datos := dta ; Sobreescribe el mapa global con los datos cargados
            Sleep(500)
            UpdateEditsFromData() ; Llama a la función que actualiza los edits
        } else {
            MsgBox("Error: El archivo JSON no es válido.")
        }

        return
    }
    
}
SaveMapSmart(fileDir) {
        global datos, EditControls  ; Asegurar acceso a los datos y los Edit



    ; Asegurarse de que haya un Case Number
    caseNumber := datos.Get("Case Number")
    if !caseNumber or caseNumber = "" {
        MsgBox("No hay 'Case Number' definido, coloque uno.")
        return
    }

    ; Crear carpeta si no existe
    if !DirExist(fileDir) {
        DirCreate(fileDir)
    }

    ; Sanitizar el nombre (evita caracteres inválidos para nombre de archivo)
    caseNumber := RegExReplace(caseNumber, "[^\w\-]", "_")

    ; Buscar un nombre disponible
    i := 0
    loop {
        suffix := (i = 0) ? "" : i
        filePath := fileDir "\" caseNumber suffix ".json"
        if !FileExist(filePath)
            break
        i++
    }

    ; Guardar el JSON
    jsonText := Jxon_Dump(datos, 4)
    FileAppend(jsonText, filePath, "UTF-8")

    MsgBox("Guardado como:`n" filePath)
}

LoadMapFromFile(filePath) {
    if !FileExist(filePath) {
        MsgBox("Archivo no encontrado: " filePath)
        return   ; Devuelve mapa vacío
    }
    

    fileContent := FileRead(filePath, "UTF-8")
    
    MsgBox("Cargado desde el archivo: " filePath)

    return fileLoaded := Jxon_Load(&fileContent)
}




; ---------------------------------------------------------------------------------------------------------------------------------

; PESTAÑA 2 (View)
tab.UseTab(2) ; (Remote Session Desktop)
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit


Remote := SkrvGui.Add("Button", " x50 y50  w75 h30", "Remote")
Remote.OnEvent("ContextMenu", (*) => A_Clipboard:="Remote Session Desktop")
Remote.OnEvent("Click", (*) => RemoteSessionBuild())
global NmOfSteps
Loop NmOfSteps {
    i := A_Index
    yTab2 := 90 + (i - 1) * spacing  ; Ajuste para calcular yTab2 en cada iteración
    InputLine(Format("Step {}", i),,150,, yTab2,,, true,,,)
    ; Código a ejecutar en cada iteración
}

; Función para recopilar y mostrar los pasos
RemoteSessionBuild() {
    stepsList := ""
    global NmOfSteps
    Loop NmOfSteps {
        i := A_Index
        controlName := 'Step ' . i
        stepText := EditControls[controlName].Value
        stepsList .= Format("Step {}: {}`n", i, stepText)
    }
    stepsList := "Accessed to TV session `n" Format("Ask the customer to reproduce the issue {}`n", EditControls["Issue"].Value) stepsList 
    A_Clipboard := stepsList
    Sleep(500)
    return

}







tab.UseTab(3) ;(Add info)
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit

SkrvGui.GetPos(&x, &y, &SkrWidth, &SkrHeigh)  ; Obtiene tamaño de la ventana
offset:=55
InputLine("Probing Questions (Add Info)",,50,990 ,offset, 250,,"paste",50+45,580)



tab.UseTab(4) ;(Email)
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit

    Remote := SkrvGui.Add("Button", " x50 y50  w75 h30", "Email")
    Remoteedit := SkrvGui.Add("Edit", " x50 y100  w990 h580", "")
    ; Guardar la referencia del Edit
    global EditControls
    EditControls["EmailInputEdit"] := Remoteedit  

    ; Capturar cambios en el input
    global datos
    datos["EmailInputEdit"] := Remoteedit
    Remoteedit.OnEvent("Change", (*) => datos["EmailInputEdit"] := Remoteedit.Value)

    Remote.OnEvent("ContextMenu", (*) => A_Clipboard:="Regarding your case number " datos["Case Number"])

    Remote.OnEvent("Click", (*) =>  EmailButom())
    EmailButom() {
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit

    UpdateDataFromEdits() ; 💡 Refresca `datos` con los valores actuales de los Edits
    if GetKeyState("Alt") {  ; Verifica si Alt está presionado
        EmailBld(false,( datos["Issue"] "`n" datos["EmailInputEdit"]),false,false ,false)
        return
    } else { 
        EmailBld(false, (datos["Issue"]),false,false ,false)
        return
    }
}





EmailBld(Greeting?, Issue? , Body?, Recommend? ,CloseSurvey?){
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit

    UpdateDataFromEdits()
    greetingdflt := Format("Dear {}`n`n", datos["&Company Name"] ) "I hope this email finds you well.`n`n" "I am writing in reference to your recent call to our technical support center regarding your issue "

    bodydflt := "I'm pleased to inform you that we have resolved this request, during our interaction we have "
    Issuedflt := "Case got corrupted, we have reimported the streams and the case is now working as expected."

    RecommendationDflt := "Additionally,  to prevent similar issues in the future, we recommend shutting down your computer every night and ensuring it remains connected during the scanning process"

    CloseSurveyDfflt:= "Thank you for allowing me to assist you today. I would greatly appreciate it if you could take a moment to provide feedback on my service by completing the following survey: " "`n`n"  "https://" datos["Survey"] "`n`n" "At the end of the survey, please leave a note about your experience during this interaction. " "`n" "Your patience and understanding throughout the process are greatly appreciated, and your feedback will help us continue improving our services as we strive for excellence."
    
    Greeting := (Greeting? Greeting:greetingdflt )
    Issue := (Issue? Issue:Issuedflt )
    Body := (Body? Body:Bodydflt )
    Recommend := (Recommend? Recommend:RecommendationDflt )
    CloseSurvey := (CloseSurvey? CloseSurvey:CloseSurveyDfflt )
    emailFinal := Greeting  Issue "`n`n" Body datos["Solution"] "`n`n" Recommend "`n`n" CloseSurvey
    A_Clipboard := emailFinal

}
tab.UseTab(5)

    global datos, EditControls  ; Asegurar acceso a los datos y los Edit

    CallBack := SkrvGui.Add("Button", " x50 y50  w205 h30", "2nd Line - Call Back")
    CallBackEdit := SkrvGui.Add("Edit", " x50 y100  w990 h580", "")
    ; Guardar la referencia del Edit
    EditControls["Call Back"] := CallBackEdit  

    ; Capturar cambios en el input
    datos["Call Back"] := CallBackEdit
    CallBackEdit.OnEvent("Change", (*) => datos["Call Back"] := CallBackEdit.Value)

    CallBack.OnEvent("ContextMenu", (*) => A_Clipboard:= "Called back to this phone number " datos["&Phone"] " and no body answered the phone. Voice message leaved in order to coninue resolving the issue" )

    CallBack.OnEvent("Click", (*) =>  CallBackCuild())


    CallBackCuild() {
        global datos, EditControls  ; Asegurar acceso a los datos y los Edit
        UpdateDataFromEdits()
        A_Clipboard := "2nd Line Title"
        Sleep(500)
        A_Clipboard := "2nd Line Body email" EditControls["Case Number"].Value
        Sleep(500)
            ; Crear objeto Excel
            xl := ComObject("Excel.Application")
            xl.Visible := true ; Muestra Excel
            wb := xl.Workbooks.Add()
            ws := wb.Worksheets(1)
            UpdateDataFromEdits()

            ; Datos de ejemplo para la tabla
            datosTabla := [
                ["Item", "Value"],
                ["Name", EditControls["Name"].Value],
                ["Email", EditControls["&Email"].Value],
                ["Phone Number", EditControls["&Phone"].Value],
                ["Company Name", EditControls["&Company Name"].Value],
                ["Software Version", EditControls["Software Version"].Value],
                ["Dongle", EditControls["&Dongle"].Value],
                ["Case Number", EditControls["Case Number"].Value]
                ; ["GUI", EditControls["GUI"].Value],
                ; ["Scanner S/N", EditControls["Scanner S/N"].Value],
                ; ["PC ID", EditControls["PC ID"].Value],
                ; ["HJ", EditControls["HJ"].Value],
                ; ["Survey", EditControls["Survey"].Value],
                ; ["SID", EditControls["SID"].Value],
                ; ["Issue", EditControls["Issue"].Value],
                ; ["TV ID", EditControls["TV ID"].Value],
                ; ["TV PSS", EditControls["TV PSS"].Value]
                ; ["RC", EditControls["RC"].Value],
                ; ["Solution", EditControls["Solution"].Value],
                ; ["Call Back", EditControls["Call Back"].Value]
            ]

            ; Escribir los datosTabla en la hoja
            for filaIndex, fila in datosTabla {
                for colIndex, valor in fila {
                    ws.Cells(filaIndex, colIndex).Value := valor
                }
            }

            ; Convertir en tabla oficial de Excel
            lastRow := datosTabla.Length
            lastCol := datosTabla[1].Length
            rango := ws.Range("A1", ws.Cells(lastRow, lastCol))
            missing := ComValue(13, 0)
            ws.ListObjects.Add(1, rango, missing, 1).Name := "MiTabla"

            ; Construir texto de la tabla
            tablaTexto := ""
            Loop datosTabla.Length - 1 {
                fila := datosTabla[A_Index + 1]
                filaTexto := ""
                for colIndex, valor in fila {
                    filaTexto .= (colIndex > 1 ? "`t" : "") . valor
                }
                tablaTexto .= filaTexto . "`r`n"
            }

            ; Copiar al portapapeles
            A_Clipboard := tablaTexto
            Sleep(500)

            ; Copiar valor individual
            A_Clipboard := EditControls["Call Back"].Value
            Sleep(500)

            ; Liberar referencias (sin cerrar Excel)
            rango := "", ws := "", wb := ""

            MsgBox "Tabla Copiada al portapapeles"

                
    }


; Salir del modo de pestañas
tab.UseTab()


isSkrvVisible := true

^+u:: {
    global isSkrvVisible
    if isSkrvVisible {
        SkrvGui.Hide()
        isSkrvVisible := false
    } else {
        SkrvGui.Show("w" SkrWidth " h" SkrHeigh)  ; Esto es correcto
        isSkrvVisible := true
    }
}

; ^+u::SkrvGui.Show()  ; Ctrl + L para mostrar la GUI


; 🔹 Atajos de teclado para cambiar pestañas
^!i::tab.Value := 1  ; Ctrl + Alt + I -> information
^!r::tab.Value := 2  ; Ctrl + Alt + R -> Rmt Sess
^!a::tab.Value := 3  ; Ctrl + Alt + A -> Add Inf
^!e::tab.Value := 4  ; Ctrl + Alt + E -> Email
^!p::tab.Value := 5  ; Ctrl + Alt + P -> Call Back

/*  */