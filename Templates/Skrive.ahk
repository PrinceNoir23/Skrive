#Requires AutoHotkey v2.0
; Navid Rahmani DDS PC / SID: 8242255223 / complaint of an existing case, wants to have the scanns of onw patient under other one mistake when scanning on TRIOS Software version 1.18.6.6 [CAS-1243173-P3
; {
;     "User Name": "Aaa",
;     "Password": "Aaa",
;     "PasswordFive": "Aaa",
;     "StationNumber": "Aaa",
;     "Hora In": "12:00",
;     "Hora Out": "12:00",
;     "Brk 1": "12:00",
;     "Brk 2": "12:00",
;     "Lunch ": "12:00"
; }

; try:
;     subprocess.run(["taskkill", "/F", "/IM", "chrome.exe"], check=True)
;     print("Chrome cerrado correctamente.")
; except subprocess.CalledProcessError:
;     print("No se pudo cerrar Chrome o no estaba abierto.")


; pyinstaller --onefile --add-binary "chromedriver.exe;." Chrome_Activator2.py
; pyinstaller --onefile --add-binary "chromedriver.exe;." .\Five9_2.py
; pyinstaller --onefile --add-binary "chromedriver.exe;." .\CRM2_2.py
global InfoFile := A_WorkingDir "\A_Info.json"
; Logs
; -----------------------------------------------------------------------
; Crear la carpeta de logs si no existe
global logDir := A_WorkingDir "\logs"
if !DirExist(logDir)
    DirCreate(logDir)

; Ruta completa del archivo de log
global logPath := logDir "\log.txt"
global logPathPy := logDir "\logPy.txt"

; Función para registrar mensajes en el log
Log(msg) {
    timestamp := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
    FileAppend(timestamp " - " msg "`n", logPath)
}
Logpy(msg) {
    timestamp := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
    FileAppend(timestamp " - " msg "`n", logPathPy)
}

Log("Skrive Started")
; -----------------------------------------------------------------------

CreateDefaultInfoFile(filePath) {
    if !FileExist(filePath) {
        defaultData := Map(
            "Brk 1",        "23:59",
            "Brk 2",        "23:59",
            "PasswordFive", "",
            "Hora In",      "23:59",
            "Hora Out",     "23:59",
            "Lunch ",       "23:59",
            "Password",     "",
            "StationNumber","",
            "User Name",    ""
        )

        json := Jxon_Dump(defaultData, 4)  ; 4 = indentación bonita
        FileAppend(json, filePath, "UTF-8")
    }
}
CreateDefaultInfoFile(InfoFile)
global datos := Map()  ; Crear un diccionario global
global EditControls := Map()  ; Guarda los controles Edit para acceder después
global checkboxStates := Map()
global NmOfSteps := 18
global fileDir1 := A_Desktop "\CasesJSON"  ; Solicitar la ruta del directorio
global fileDir_CasesFinal := A_Desktop "\Cases_Final"  ; Solicitar la ruta del directorio
; global rutaPython := "C:\Python313\python.exe"
; global rutaPython := "C:\Users\Joel Hurtado\AppData\Local\Programs\Python\Python313\python.exe"
global rutaScript := A_WorkingDir "\CRM2.py"

; Ruta para python ------------------------------------------------------------
; pythontxt := A_WorkingDir "\python_path.txt"
; if FileExist(pythontxt){
;     FileDelete(pythontxt) 
; }

; Ejecutar "where python" y guardar la salida
try {
    ; tvPS debe estar definido; esto es solo un ejemplo
    RunWait(A_ComSpec ' /c where python > python_path.txt', , 'Hide')
    Logpy("Python path leido correctamente.")
    ; Leer archivo generado por `where python`
    raw := FileRead("python_path.txt")

    ; Separar por líneas
    lines := StrSplit(raw, "`n")

    ; Tomar la primera línea y quitar posibles espacios o saltos de carro (\r)
    global rutaPython := Trim(lines[1], "`r`n ")
}
catch  {   
    MsgBox('instalando python')
    Logpy("ERROR: python no estaba instalado, se va a instalar " )
    RunWait("*RunAs " A_WorkingDir "\PythonInstaller.bat")

}




; Ruta para python ------------------------------------------------------------


for _, dirPath in [fileDir_CasesFinal, fileDir1] {
    if !DirExist(dirPath)
        DirCreate(dirPath)
}




; ; tag v1.2.2.0 Msgboxes 
; ; tag v1.2.2.1 25 segundos

; ; https://benchmark.unigine.com/heaven

#Include Json.ahk
#Include TeamViewer.ahk  
#Include Klokken.ahk
#Include AutoLogin.ahk

Persistent 1  ; Establece el script como persistente
; ChatGPT - Google Chrome
; ahk_class Chrome_WidgetWin_1
; ahk_exe chrome.exe





ForwardToDynamics() {
;     Case: 3Shape v3: New Case - Dynamics 365 - Google Chrome
; ahk_class Chrome_WidgetWin_1
    

    if !WinExist("ahk_exe chrome.exe")  {
        MsgBox "Chrome no está abierto`n Abrelo y Corre el codigo nuevamente."
        Sleep(100)
        return
    }
    if !WinExist("Case: 3Shape v3: New Case - Dynamics 365 - Google Chrome"){
        MsgBox "El Tab de CRM (New Case) no esta abierta`n Abrela y Corre el codigo nuevamente."
        Sleep(100)
        return
    }
    Sleep(500)
    WinActivate("Case: 3Shape v3: New Case - Dynamics 365 - Google Chrome")  ; Activa la ventana de TeamViewer
    Sleep(500)
    WinWaitActive("Case: 3Shape v3: New Case - Dynamics 365 - Google Chrome")
    Sleep(50)
    exename := WinGetProcessName("A")
    if !(exename = "chrome.exe") {
    MsgBox A_ComputerName
    }
    Return false
    ; 
}

; ---------------------------------------------------------------------------------------------------------------------------------

SkrvGui := Gui(, "SKRIVE")


SkrvGui.Opt("+Resize +MinSize1280x730 +MaximizeBox +MinimizeBox") ; Hace la ventana redimensionable
SkrvGui.SetFont("s15 ", "Segoe UI")  ; Tamaño 14, negrita, fuente bonita
SkrvGui.SetFont("q5")
; SkrvGui.SetFont("cA53860")
SkrvGui.SetFont("s12 italic", "Segoe UI")  ; Tamaño 14, negrita, fuente bonita

tab := SkrvGui.Add("Tab3", "x10 y10 w900 h800", ["Information","Remote Session", "Email"])
SkrvGui.SetFont("s12 norm", "Segoe UI")  ; Tamaño 14, negrita, fuente bonita



SkrvGui.OnEvent("Size", (*) => AutoResize())  ; Evento para ajustar tamaño dinámicamente
SkrWidth := 1280
SkrHeigh := 730 
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
    editSoftware := SkrvGui.Add("Edit", Format("x260 y{} w255 h{}", yOffset, BtnHeigh), "")
    
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
    SoftwGui.SetFont("s12 ", "Segoe UI")
    SoftwGui.SetFont("q5")


    SoftwGui.Add("Text", "x10 y10", "Select Software Versions:")
    
    SoftwGui.Add("Text", "x10 y30", "Module:")
    comboBox1 := SoftwGui.Add("ComboBox", "x10 y50 w200", ["Unite","TRIOS", "TRIOS Software", "Dental Desktop (unite)","Unite Cloud", "Model Builder", "Automate","3shape Account"])
    
    SoftwGui.Add("Text", "x230 y30", "Version:")
    comboBox2 := SoftwGui.Add("ComboBox", "x230 y50 w150", ["1.7.83.0",  "1.8.8.0","1.8.10.1", "1.18.8.5","1.18.6.6" ,"1.18.7.6" ,"1.7.8.1", "1.8.5.1","1.7.82.5"])

    ; Valores por defecto
    comboBox1.Text := "Unite"
    comboBox2.Text := "1.7.83.0"

    btnSelect := SoftwGui.Add("Button", "x10 y120 w320", "Select")
    
    btnSelect.OnEvent("Click", (*) => (
        datos["Modulo"] := comboBox1.Text,
        datos["Version"] := comboBox2.Text,
        editSoftware.Value := datos["Modulo"] " version " datos["Version"],
        datos[Name2Text] := editSoftware.Value,
        SoftwGui.Destroy()
    ))
    
    
    SoftwGui.Show()
}

CategorySelect(Name, yOffset, btnSze, BtnHeigh) {
    Name2Text := Name
    btnCategory := SkrvGui.Add("Button", Format("x520 y{} w{} h{}", yOffset, btnSze/2, BtnHeigh), Name2Text)
    editCategory := SkrvGui.Add("Edit", Format("x625 y{} w405 h{}", yOffset, BtnHeigh), "")
    
    btnCategory.OnEvent("Click", (*) => CategoryGUI(editCategory, Name2Text))
    ; Guardar la referencia del Edit
    global EditControls
    EditControls[Name2Text] := editCategory  

    global datos
    datos[Name2Text] := editCategory.Value
    UpdateDataFromEdits() ; 💡 Refresca `datos` con los valores actuales de los Edits

    editCategory.OnEvent("Change", (*) => datos[Name2Text] := editCategory.Value)
}


CategoryGUI(editCategory, Name2Text) { 
        global datos, EditControls  ; Asegurar acceso a los datos y los Edit


    CategorGui := Gui()
    CategorGui.Opt("+AlwaysOnTop")
    CategorGui.SetFont("s12 ", "Segoe UI")
    CategorGui.SetFont("q5 ")

    CategorGui.Add("Text", "x10 y10", "Select Category:")
    
    CategorGui.Add("Text", "x10 y30", "Category:")
    comboBoxCat1 := CategorGui.Add("ComboBox", "x10 y50 w150", ["Complaint", "Request", "Miscellaneous"])
    
    CategorGui.Add("Text", "x180 y30", "Category Area:")
    comboBoxCat2 := CategorGui.Add("ComboBox", "x180 y50 w150", ["Solved Remotely",  "Support Fee Rejected","Data Migration",  "Installation","General Product Information","Onboarding","Partner","Remote" ])

    CategorGui.Add("Text", "x350 y30", "Case Type:")
    comboBoxCat3 := CategorGui.Add("ComboBox", "x350 y50 w150", ["PC/OS/Network",  "Expected Behaviour","Performance","Not Reproducible","Solved by customer","How to question", "Product Info"])



    ; Valores por defecto
    comboBoxCat1.Text := "Complaint"
    comboBoxCat2.Text := "Solved Remotely"
    comboBoxCat3.Text := "Expected Behaviour"

    btnSelect := CategorGui.Add("Button", "x10 y100 w510", "Select")
    
    btnSelect.OnEvent("Click", (*) => (
        datos["Categ"] := comboBoxCat1.Text,
        datos["CategoryArea"] := comboBoxCat2.Text,
        datos["CaseType"] := comboBoxCat3.Text,
        editCategory.Value := datos["Categ"] " || " datos["CategoryArea"] " || " datos["CaseType"],
        datos[Name2Text] := editCategory.Value,
        CategorGui.Destroy()
    ))
    
    
    CategorGui.Show()
}


LinkSelect(Name, yOffset, btnSze, BtnHeigh) {
    Name2Text := Name
    btnLink := SkrvGui.Add("Button", Format("x520 y{} w{} h{}", yOffset, btnSze/2, BtnHeigh), Name2Text)
    editLink := SkrvGui.Add("Edit", Format("x625 y{} w405 h{}", yOffset, BtnHeigh), "")
    
    btnLink.OnEvent("Click", (*) => editLink.Value := A_Clipboard)
    btnLink.OnEvent("ContextMenu", (*) => Altclick() )    ; Clic derecho: limpia el Edit
        Altclick() {
            if GetKeyState("Alt"){  ; Verifica si Alt está presionado
                    A_Clipboard := editLink.Value
                    return
            }
            else { 
                editLink.Value := ""
                return
            }
        }
    ; Guardar la referencia del Edit
    global EditControls
    EditControls[Name2Text] := editLink  

    global datos
    datos[Name2Text] := editLink.Value
    UpdateDataFromEdits() ; 💡 Refresca `datos` con los valores actuales de los Edits

    editLink.OnEvent("Change", (*) => datos[Name2Text] := editLink.Value)
}


OutputLine(Name,xOffset, yOffset, Insize ,btnSze,BtnHeigh){
    ; Name2Text := Format("{}",Name)
    Name2Text := Name
    Outbtn := SkrvGui.Add("Button", Format("x{} y{} w{} h{}",xOffset, yOffset,btnSze,BtnHeigh), Name2Text)
    outedit := SkrvGui.Add("Edit", Format("x{} y{} w{} h{}",xOffset, yOffset+BtnHeigh+8,Insize,BtnHeigh*2.6), "")
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
    UpdateDataFromEdits()
    

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
    HJ := datos.Get("HJ", "")

    if HJ == ""{
        IntNote := Format("No helpjuice used, as no article found for this process")

    }
    else{
        IntNote := Format("This HelpJuices was used `n{}", HJ)
    }


    

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

    datos["PhT"] := PH 
    datos["PhN"] := PHNote
    datos["InT"] := Int
    datos["InN"] := IntNote
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
    UpdateDataFromEdits()


    issue := datos.Has("Issue") ? datos["Issue"] : "No issue"
    softwareVersion := datos.Has("Software Version") ? datos["Software Version"] : "Unknown version"


    C1 := "C1st"

    ; Obtiene valores, si no existen usa "N/A"
    Cname := datos.Get("&Company Name", "N/A")
    SID := datos.Get("S&ID", "N/A")
    issue := datos.Has("Issue") ? datos["Issue"] : "No issue"
    softwareVersion := datos.Has("Software Version") ? datos["Software Version"] : "Unknown version"
    Descrp := issue " on " softwareVersion

    
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
    UpdateDataFromEdits()

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


/*  */
Automatic(){
    Sleep(1000)
    UpdateDataFromEdits() ; Refresca `datos` con los valores actuales de los Edits
    SaveBttm(false,fileDir1)
        Sleep(500)
    DescriptionGUI(false)
        Sleep(500)
    C1stAdd(false)
        Sleep(500)
    A_Clipboard:= datos["&Email"]
        Sleep(500)
    IntPhBttm(true)
        Sleep(800)
    A_Clipboard:= "Logs and images are here"
        Sleep(800)
    IntPhBttm(false)
        Sleep(500)


    A_Clipboard:="Remote Session Desktop"
        Sleep(500)
    RemoteSessionBuild()
        Sleep(500)

    A_Clipboard:= EditControls["Probing Questions (Add Info)"].Value
        Sleep(500)

    A_Clipboard := "RC: " datos["RC"] "`n" "S: " datos["Solution"]
        Sleep(500)

    A_Clipboard:="Regarding your case number " datos["C&ase Number"]
        Sleep(500)
    EmailButom()


        Sleep(500)
    MsgBox("Informacion Total del caso copiada al portapapeles", "Informacion Copiada Exitosamente","64")
        Sleep(500)
    A_Clipboard := datos["&Dongle"]
        Sleep(500)
    A_Clipboard := datos["&Email"]
        Sleep(500)
    
    return
}

FindBar(item){
    Sleep(800)
    Send('^f')
    Sleep(650)
    Send(item)
    Sleep(650)
    Send('{Enter}')
    Sleep(650)
    Send('{Esc}')
    Sleep(650)
    Send('{Enter}')
    Sleep(900)
    return

}

CreateNote(TitleNote,BodyNote,addlog){
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit
    
    Sleep(1000)
    Send(TitleNote)
    Sleep(800)
    Send('{Tab}')
    Sleep(800)
    A_Clipboard := BodyNote
    Sleep(500)
    Send('{Control Down}{V}{Control Up}')
    Sleep(500)
    if addlog == true {
        Loop 1 {
            Send('{Tab}')
            Sleep(445)
        }
        Sleep(445)
        Send('{Enter}')
        Sleep(1000)
        Loop 2 {
            Send('{Shift Down}{Tab Down}{Shift Up}{Tab Up}')
            Sleep(800)
        }
        Send('{Down}')
        Sleep(445)
        Send('{Up}')
        Sleep(445)
        Send('{Enter}')
        Sleep(6000)
        Loop 2 {
            Send('{Tab}')
            Sleep(445)
        }
        Send('{Enter}')
        Sleep(6000)
        return
    }
    if addlog == false {
        Loop 2   {
            Send('{Tab}')
            Sleep(445)
        }
        Send('{Enter}')
        Sleep(800)
        return
    }

}

; ; CRM(){
; ;     global datos,EditControls
; ;     UpdateDataFromEdits() ; 💡 Refresca `datos` con los valores actuales de los Edits
; ;         ; Sleep(200)
; ;     ; SaveBttm(false,fileDir1)
; ;     Sleep(1000)
    
; ;     FindBar("Summary")
; ;     Sleep(1000)
; ;     Loop 2 {
; ;         Send('{Tab}')
; ;         Sleep(445)
; ;     }

; ;         Sleep(1000)
; ;     DescriptionGUI(false)
; ;         Sleep(500)
; ;     Send('{Control Down}{v}{Control Up}')  
; ;         Sleep(500)
; ;     Send('{Tab}')

; ;         Sleep(500)
; ;     Send('{Tab}')
; ;         Sleep(500)
; ;     Send('{N}')
; ;         Sleep(500)
; ;     Send('{Enter}')
; ;     Sleep(400)
; ;     Loop 4 {
; ;         Send('{Tab}')
; ;         Sleep(445)
; ;     }
; ;         ; Sleep(500)
; ;     ; Send('{Enter}')
; ;         Sleep(1000)
; ;     A_Clipboard := datos["Modulo"] ;Product
; ;         Sleep(1000)
; ;     Send('{Control Down}{v}{Control Up}')  
; ;         Sleep(1500)
; ;     Send('{Enter}')
; ;     Sleep(1500)
; ;     Loop 4 {
; ;         Send('{Tab}')
; ;         Sleep(445)
; ;     }
    
; ;     Send('{S}') ;Serious
; ;     Sleep(445)
; ;     Send('{Enter}')
; ;     Sleep(445)
; ;     Send('{Tab}')
; ;     Sleep(445)
; ;     Send('{P}') ;Phone
; ;     Sleep(445)
; ;     Send('{Enter}')
; ;     Sleep(445)
; ;     Loop 4 {
; ;         Send('{Tab}')
; ;         Sleep(445)
; ;     }
; ;         Sleep(1000)
; ;     A_Clipboard:= datos["&Dongle"]
; ;         Sleep(445)
; ;     Send('^v')
; ;         Sleep(1500)
; ;     Send('{Enter}')
; ;     Loop 13 {
; ;         Send('{Tab}')
; ;         Sleep(445)
; ;     }
; ;         ; Sleep(500)
; ;     ; Send('{Enter}')
; ;         Sleep(1000)
; ;     A_Clipboard := datos["Version"] ;Version
; ;         Sleep(1000)
; ;     Send('{Control Down}{v}{Control Up}')   
; ;         Sleep(1500)
; ;     Send('{Enter}')
; ;         Sleep(1500)
; ;     Loop 3 {
; ;         Send('{Tab}')
; ;         Sleep(445)
; ;     }
; ;         Sleep(445)
; ;     Send('{Enter}')
; ;         Sleep(1000)
; ;     A_Clipboard := datos["Categ"] ;Category
; ;         Sleep(1000)
; ;     Send('{Control Down}{v}{Control Up}')  
; ;         Sleep(3500)
; ;     Send('{Enter}')
; ;     Sleep(1500)
; ;     Loop 3 {
; ;         Send('{Tab}')
; ;         Sleep(600)
; ;     }
; ;         Sleep(445)
; ;     Send('{Enter}')
; ;         Sleep(1000)
; ;     A_Clipboard := datos["CategoryArea"] ;Category Area
; ;         Sleep(1000)
; ;     Send('{Control Down}{v}{Control Up}') 
; ;         Sleep(3500)
; ;     Send('{Enter}')
; ;     Sleep(1500)
; ;     Loop 3 {
; ;         Send('{Tab}')
; ;         Sleep(600)
; ;     }
; ;         Sleep(445)
; ;     Send('{Enter}')
; ;         Sleep(1000)
; ;     A_Clipboard := datos["CaseType"] ;CaseType
; ;         Sleep(1000)
; ;     Send('{Control Down}{v}{Control Up}') 
; ;         Sleep(3500)
; ;     Send('{Enter}')

; ;     Sleep(2500)

; ;     Send('^f')
; ;     Sleep(650)
; ;     Send("Dynamics 365 Mobile – custom")
; ;     Sleep(650)
; ;     Send('{Enter}')
; ;     Sleep(650)
; ;     Send('{Esc}')
; ;     Sleep(650)


; ;     Loop 14 {
; ;             Send('{Shift Down}{Tab Down}{Shift Up}{Tab Up}')
; ;             Sleep(500)
; ;     }
; ;         Sleep(500)
; ;     Send('{Enter}')
; ;         Sleep(500)
; ;     A_Clipboard:= datos["&Email"]
; ;         Sleep(400)
; ;     Send('^v')
; ;         Sleep(2500)
; ;     Send('{Enter}')
; ;         Sleep(1000)


; ;     ; SAVE
; ;     Sleep(800)
; ;     Send('^f')
; ;     Sleep(650)
; ;     Send("save &")
; ;     Sleep(650)
; ;     Send('{Enter}')
; ;     Sleep(650)
; ;     Send('{Esc}')
; ;     Sleep(800)

; ;     Loop 1 {
; ;             Send('{Shift Down}{Tab Down}{Shift Up}{Tab Up}')
; ;             Sleep(1000)
; ;     }
; ;     Sleep(1000)
; ;     Send('{Enter}')
; ;     Sleep(500)

; ;     Sleep(18000)

; ;     ; ; ; MsgBoxSteps
; ;     ; R_save:= MsgBox("¿Se guardo el caso?", "Confirmacion Para Continuar Proceso", "36")
; ;     ; if (R_save = "No") {
; ;     ;      ; Si el usuario elige "No", cancela la acción
; ;     ;     Automatic()
; ;     ;     return

; ;     ; }


; ;     Sleep(1000)

; ;     FindBar("Summary")
; ;     Sleep(1000)
; ;     Loop 3 {
; ;         Send('{Tab}')
; ;         Sleep(445)
; ;     }
; ;         Sleep(500)
; ;     Send('^c') 
; ;         Sleep(1000)
; ;     Alt_a()

; ;         Sleep(1000)
; ;     Send('{Left}')
; ;         Sleep(500)    
; ;     C1stAdd(false)
; ;         Sleep(500)
; ;     Send('^v')
; ;         Sleep(400)

; ;     IntPhBttm(false)
; ;     RemoteSessionBuild() 

; ;     ; MsgBox datos["PhT"] "`n`n" datos["PhN"] "`n`n" datos["InT"] "`n`n" datos["InN"] "`n`n"

; ;     ; MsgBox Type(datos["PhT"]) "`n`n" Type(datos["PhN"]) "`n`n" Type(datos["InT"]) "`n`n" Type(datos["InN"]) "`n`n"
; ;     Ph1 := datos["PhT"]         

; ;     PHNote1 := datos["PhN"]
; ;     Int1 := datos["InT"]
; ;     IntNote1 := datos["InN"]

; ;     Sleep(500)
; ;     FindBar("Enter a note")
; ;     Sleep(100)
; ;     CreateNote(Ph1,PHNote1,false)
; ;     Sleep(500)
; ;     FindBar("Enter a note")
; ;     Sleep(100)
; ;     CreateNote(Int1,IntNote1,false)
; ;     Sleep(500)
; ;     FindBar("Enter a note")
; ;     Sleep(200)
; ;     rmtsession:= datos["RMTSS"]
; ;     Sleep(100)
; ;     CreateNote("Remote Session Desktop",rmtsession,false)
; ;     Sleep(500)
; ;     FindBar("Enter a note")
; ;     Sleep(100)
; ;     CreateNote(Int1,"Logs and images are here `n",true)
; ;     Sleep(1000)

; ; Sleep(18000)

; ; ; ; ; MsgBoxSteps
; ; ; ; ; Si el usuario elige "No", cancela la acción
; ;     ; R_logs:= MsgBox("¿Se guardaron los logs?", "Confirmacion Para Continuar Proceso", "36")
; ;     ; if (R_logs = "No") {
         
; ;     ;     Automatic()
; ;     ;     return

; ;     ; }
    
; ;     ; SAVE
; ;     Sleep(800)
; ;     Send('^f')
; ;     Sleep(650)
; ;     Send("save &")
; ;     Sleep(650)
; ;     Send('{Enter}')
; ;     Sleep(650)
; ;     Send('{Esc}')
; ;     Sleep(800)

; ;     Loop 1 {
; ;             Send('{Shift Down}{Tab Down}{Shift Up}{Tab Up}')
; ;             Sleep(1000)
; ;     }
; ;     Sleep(1000)
; ;     Send('{Enter}')


; ;     Sleep(15000)
; ;     Sleep(800)
; ;     Send('^f')
; ;     Sleep(650)
; ;     Send("merged cases")
; ;     Sleep(650)
; ;     Send('{Enter}')
; ;     Sleep(650)
; ;     Send('{Esc}')
; ;     Sleep(800)
; ;     Send('{Right}')
; ;     Sleep(500)
; ;     Send('{Enter}')
; ;     Sleep(650)
; ;     Loop 5 {
; ;             Send('{Tab}')
; ;             Sleep(500)
; ;     }

; ;         Sleep(500)
; ;     Send('^c') 
; ;         Sleep(500)
; ;     EditControls["Sur&vey"].Value := A_Clipboard
; ;     ; Send('¡v') 
; ;     Sleep(500)
; ;     FindBar("summary")
; ;     Sleep(1000)

; ;     FindBar("Search timeline")
; ;     Sleep(800)
; ;     Loop 5 {
; ;             Send('{Shift Down}{Tab Down}{Shift Up}{Tab Up}')
; ;             Sleep(500)
; ;     }
; ;     Sleep(445)
; ;     Send('{Enter}')
; ;     Sleep(445)
; ;     Send('{Enter}')
; ;     Sleep(3500)
; ;     Loop 13 {
; ;             Send('{Tab}')
; ;             Sleep(445)
; ;     }
; ;     Sleep(800)
; ;     A_Clipboard:="Regarding your case number " datos["C&ase Number"]
; ;     Sleep(445)
; ;     Send('^V')
; ;     Sleep(445)
; ;     Loop 2 {
; ;             Send('{Tab}')
; ;             Sleep(445)
; ;     }
; ;     Sleep(445)
; ;     EmailButom()
; ;     Sleep(445)
; ;     Send('^v') 
; ;     Sleep(1000)
; ;     FindBar("RECIPIENT INFO")
; ;     Sleep(445)
; ;     Loop 6 {
; ;             Send('{Shift Down}{Tab Down}{Shift Up}{Tab Up}')
; ;             Sleep(500)
; ;     }
; ;     Sleep(445)
; ;     Send('{Enter}')
; ;     Sleep(445)


; ;     Sleep(18000)

; ; ; ; ; MsgBoxSteps
; ;     ; R_CRM:= MsgBox("¿Ya fue enviado el email?", "Confirmacion Para Continuar Proceso", "36")
; ;     ; if (R_CRM = "No"){  ; Si el usuario elige "No", cancela la acción
; ;     ;     Automatic()
; ;     ;     return 
; ;     ; }
; ;     ; Sleep(5000)
; ;     ; ; ; ; |||


; ;     FindBar("Description &")
; ;     Sleep(445)
; ;     Send('{Tab}')
; ;     Sleep(445)
; ;     Send('{Enter}')
; ;     Sleep(445)
; ;     A_Clipboard:= EditControls["Probing Questions (Add Info)"].Value
; ;     Sleep(500)
; ;     Send('^v') 
; ;     Sleep(445)
; ;     Loop 10 {
; ;             Send('{Tab}')
; ;             Sleep(445)
; ;     }
; ;     Sleep(445)
; ;     A_Clipboard := "RC: " datos["RC"] "`n" "S: " datos["Solution"]
; ;     Sleep(500)
; ;     Send('^v') 
; ;     Sleep(1000)

; ;     R_Close:= MsgBox("¿Ya desea Terminar el caso?", "Confirmacion Para terminar flow", "36")
; ;     if (R_Close = "No"){  ; Si el usuario elige "No", cancela la acción
; ;         Automatic()
; ;         return 
; ;     }
; ;     Sleep(800)
; ;     Send('^f')
; ;     Sleep(650)
; ;     Send("americas 1st")
; ;     Sleep(650)
; ;     Send('{Enter}')
; ;     Sleep(650)
; ;     Send('{Esc}')
; ;     Sleep(650)
; ;     Loop 2 {
; ;             Send('{Tab}')
; ;             Sleep(445)
; ;     }
; ;     Sleep(650)
; ;     Send('{Enter}')
; ;     Sleep(2000)
; ;     FindBar("Next Stage")
; ;     Sleep(2000)
; ;     FindBar("Next Stage")
; ;     Sleep(2000)
; ;     Loop 3 {
; ;             Send('{Tab}')
; ;             Sleep(445)
; ;     }
; ;     Sleep(800)
; ;     A_Clipboard := "RC: " datos["RC"] "`n" "S: " datos["Solution"]
; ;     Sleep(500)
; ;     Send('^v') 
; ;     Sleep(1000)
; ;     FindBar("Finish")
; ;     Sleep(10000)
; ;     FindBar("Resolve Case")
; ;     Sleep(7000)
; ;     Loop 3 {
; ;             Send('{Tab}')
; ;             Sleep(445)
; ;     }
; ;     Sleep(500)
; ;     A_Clipboard := "RC: " datos["RC"] "`n" "S: " datos["Solution"]
; ;     Sleep(500)
; ;     Send('^v') 
; ;     Sleep(1000)
; ;     Loop 3 {
; ;             Send('{Shift Down}{Tab Down}{Shift Up}{Tab Up}')
; ;             Sleep(870)
; ;     }
; ;     Send('{Enter}')
    

; ;         Sleep(1500)
; ;     SaveBttm(false,fileDir1)

; ;     MsgBox("CRM Completado", "CRM AUTOMATIZATION","64")
; ;         Sleep(500)
    
; ;     Reload
    
; ;     return

; ; }

; ; CRM2(CRMBool){
; ;     global datos,EditControls
; ;     global pythonPath
; ;     datos["Case Link"] := ""

; ;     UpdateDataFromEdits() ; 💡 Refresca `datos` con los valores actuales de los Edits
; ;     IntPhBttm(false)
; ;     RemoteSessionBuild() 
; ;     EmailBld(false,( datos["Issue"] "`n" ),false ,datos["EmailInputEdit"] "`n" ,false)
    

; ;     if (CRMBool == true){
; ;         PyPath := A_WorkingDir . "\CRM2.py"
; ;     }else if (CRMBool == false){
; ;         PyPath := A_WorkingDir . "\CRM.py"
; ;     }
; ;     jsonPath := A_WorkingDir . "\NewCase.json"
; ;     if FileExist(jsonPath) {
; ;         FileDelete(jsonPath)
; ;     }
; ;     FileAppend(Jxon_Dump(datos,4), jsonPath, "UTF-8")

    


; ;     Sleep(200)
; ;     ; RunWait(Format('python.exe "{}" "{}" ',PyPath,jsonPath))
; ;     logPath := A_WorkingDir . "\error_log.txt"
; ;     SkrvGui.Minimize()
    

; ;     RunWait(Format('cmd.exe /c "{}" "{}" "{}" 2> "{}"', pythonPath, PyPath, jsonPath, logPath))


    
    
; ;     ; ; esto esconde la ventana de la consola de python
; ;     ; RunWait(Format('python.exe "{}" "{}"', PyPath, jsonPath), , "Hide")


; ;     ; FileDelete(jsonPath)

; ; }


; ---------------------------------------------------------------------------------------------------------------------------------

; PESTAÑA 1 (General)
tab.UseTab(1) ; (Information)
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit

datos["Modulo"] := "Unite"
datos["Version"] := "1.7.83.0"
datos["Categ"] := "Complaint"
datos["CategoryArea"] := "Solved Remotely" 
datos["CaseType"] := "Expected Behaviour" 

imagePath := A_ScriptDir . "/IMG_LogoSmall.png"
scale := 2.4
wd:= 843
hi := 559
imgWidth1 := wd / scale
imgHeight1 := hi / scale

; Obtener el tamaño actual de la ventana para posicionar la imagen
SkrvGui.GetPos(&winX1, &winY1, &winWidth1, &winHeight1)
imgX1 := 1060
imgY1 := 350

imgControl := SkrvGui.Add("Picture", Format("x{} y{} w{} h{}", imgX1, imgY1, imgWidth1, imgHeight1), imagePath)

imagePath_Logo := A_ScriptDir . "/IMG_SKRIVE_Trns.png"
scale_Logo := 3.3
wd_Logo:= 1024
hi_Logo := 1024
imgWidth1_Logo := wd_Logo / scale_Logo
imgHeight1_Logo := hi_Logo / scale_Logo


imgX1_Logo := 1005
imgY1_Logo := 50

imgControl_Logo := SkrvGui.Add("Picture", Format("x{} y{} w{} h{}", imgX1_Logo, imgY1_Logo, imgWidth1_Logo, imgHeight1_Logo), imagePath_Logo)




y := 50 ; Posición inicial en Y
x:= 260
spacing := 35 ; Espaciado entre elementos
Insize:=390

tvsize := (750/2)-(spacing*2)
Complaints := Map() 
Request := Map()
Miscellaneous := Map()

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

        loop files folderPath "\*.*" {
            SplitPath A_LoopFileName, , , , &NameNoExt
            fileList.Push(NameNoExt)
        }

        ; Guardar la lista completa de archivos de esta carpeta (sin extensiones)
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

savecase() {
    UpdateDataFromEdits()
    Caseoption1 := comboBoxCaseSelect1.Text
    Casetype1 := comboBoxCaseSelect2.Text
    if (Casetype1 == "" or Caseoption1 == "") {
        MsgBox("Coloca un Tipo y Nombre de caso para guardar", "Guardar Template", "48")
        return
    }
    filePathOptions1 := A_WorkingDir "\" Caseoption1
    savedpath := filePathOptions1 "\" Casetype1 ".json"
    if !DirExist(filePathOptions1) {
        DirCreate(filePathOptions1)
    }
    Datos2SavedCases := ["HJ", "Probing Questions (Add Info)", "Issue", "RC", "Solution", "EmailInputEdit", "Modulo","CaseType", "Categ", "Category", "CategoryArea", "Version"]

        

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
    if FileExist(savedpath){
        ; Eliminar el archivo existente si existe
        FileDelete(savedpath)
    }

    Rt:= MsgBox(Format("Seguro que desea guardar el caso {} como {}? ",Caseoption1,Casetype1), "Confirmacion Para Guardar Template", "292")
    if (Rt = "No")  ; Si el usuario elige "No", cancela la acción
        return
    ; Guardar el JSON
    jsonText := Jxon_Dump(dats2, 4)
    FileAppend(jsonText, savedpath, "UTF-8")
    MsgBox("Guardado como:`n" savedpath, "Guardando Template", "64")
    return
}


eval(){
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit
    
    Caseoption := comboBoxCaseSelect1.Text
    Casetype := comboBoxCaseSelect2.Text ".json"
    if (Casetype == "" or Caseoption == ""){
        MsgBox("Coloca un Tipo y Nombre de caso para cargar", "Cargar Template", "48")
        return
    }
    dta1 := LoadMapFromFileEXISTINGCASE(Caseoption,Casetype) 
        if (Type(dta1) = "Map") {
            datos := dta1 ; Sobreescribe el mapa global con los datos cargados
            Sleep(500)
            UpdateEditsFromData() ; Llama a la función que actualiza los edits
        } else {
            MsgBox("El archivo JSON no es válido.", "Error de Archivo","16")
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
        MsgBox("Archivo no encontrado: " filePath1,"Error de Path","16")
        return
    }

    
    

    fileContent := FileRead(filePath1, "UTF-8")
    MsgBox("Cargado desde el archivo: " filePath1,"Cargando Template","64")

    return fileLoaded := Jxon_Load(&fileContent)
}


InputLine("TV ID",,(btnSze/2)+55,tvsize, y,btnSze/2,BtnHeigh, false), y
InputLine("TV PSS",465,570,tvsize, y,btnSze/2,BtnHeigh, false), y 
Int_PhBtt := SkrvGui.Add("Button", Format("x880 y{} w{} h{}", y, btnSze-100, BtnHeigh), "INT-P&H"), y 
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
    return
}



Int_PhBtt.OnEvent("ContextMenu", (*) => IntPhBttm(true))

c1stbttn := SkrvGui.Add("Button", Format("x985 y{} w{} h{}", y, 45, BtnHeigh), "C&1st"), y += spacing
c1stbttn.OnEvent("Click", (*) => C1stAdd(false))
c1stbttn.OnEvent("ContextMenu", (*) => C1stAdd(true))

InputLine("Name",,,255, y,btnSze,BtnHeigh, true,,), y 
LinkSelect("Case Link",y,btnSze,30), y += spacing
InputLine("&Phone",,,200, y,btnSze,BtnHeigh, true,,),
InputLine("&Email",465,535,320, y,btnSze/3,BtnHeigh, true,,),
InputLine("PC ID",860,930,100, y,btnSze/3,BtnHeigh, true,,), y += spacing
InputLine("&Company Name",,,255, y,btnSze,BtnHeigh, true,,), 
InputLine("C&ase Number",520,660,370, y,btnSze/1.5,BtnHeigh, true,,), y += spacing
SoftwareVersionSelect("Software Version", y,btnSze,BtnHeigh), y 
CategorySelect("Category", y,btnSze,BtnHeigh), y += spacing

InputLine("&Dongle",,,110, y,btnSze,BtnHeigh, true,,), 
InputLine("S&ID",375,445,115, y,btnSze/3,BtnHeigh, true,,),
InputLine("Scanne&r S/N",565,670,180, y,btnSze/2,BtnHeigh, true,,),
InputLine("GUI",855,925,105, y,btnSze/3,BtnHeigh, true,,), y += spacing
InputLine("HJ",,,300, y,btnSze,BtnHeigh, true,,), 
InputLine("Sur&vey",565,670,360, y,btnSze/2,BtnHeigh, true,,), y += spacing
InputLine("RC",,,, y,btnSze,BtnHeigh, true,,), y += spacing
; ---------------------------------------------------------------------------------------------------------------------------------
descripton:= SkrvGui.Add("Button", Format("x50 y{} w{} h{}", y,btnSze,BtnHeigh), "Descrip&tion"),
descripton.OnEvent("ContextMenu", (*) => DescriptionGUI(true))
descripton.OnEvent("Click", (*) => DescriptionGUI(false))
    

DescriptionGUI(bool) {
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit

    DescripGui := Gui()
    DescripGui.Opt("+AlwaysOnTop")


    DescripGui.SetFont("s12 ", "Segoe UI")
    DescripGui.SetFont("q5 ")

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
Hotkeys("Hotke&ys", y), y


BtnSaveInfo := SkrvGui.Add("Button", "x50 y440 w150 h40", "&Save - Load Info")
clear := SkrvGui.Add("Button", " x205 y440  w45 h30", "Clear")

clear.OnEvent("Click", (*) => ClearAll())

ClearAll() {
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit

    respuesta := MsgBox("¿Está seguro de que desea borrar todos los campos?", "Borrar Datos", "292")
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
            MsgBox("Guardado cancelado.","Guardar Caso","64")
            return
        }
        SaveMapSmart(filePath)  ; Llama a tu función de guardado
        return
    } else {
        SaveBttm(false,fileDir1)  ; Tu función alternativa si no se presionó Alt
        return
    }
    return
}

; Función para abrir una nueva ventana
SaveBttm(SvBool,fileDir) {

    global datos, EditControls  ; Asegurar acceso a los datos y los Edit

    C1stAdd(false) ; Actualiza el mapa con la descripción
    
    ; Recorrer el mapa y mostrar los valores
    if SvBool==false{
        SaveMapSmart(fileDir)  ; Guardar el mapa en un archivo JSON
        return
    }


    if SvBool==true{
        
        caseNumber := datos.Get("C&ase Number")

        if caseNumber != "" {
            filePath := fileDir "\" caseNumber ".json"
        } else {
            filePath := FileSelect(, fileDir, , "JSON (*.json)")
        }

        if (!filePath) {
            MsgBox("Recuperación cancelada.", "Cargar Caso", "64")
            return
        }

        dta := LoadMapFromFile(filePath) 
        if (Type(dta) = "Map") {
            datos := dta ; Sobreescribe el mapa global con los datos cargados
            Sleep(500)
            UpdateEditsFromData() ; Llama a la función que actualiza los edits
        } else {
            MsgBox("El archivo JSON no es válido.","Error de Archivo","16")
        }

        return
    }
    return
    
}
SaveMapSmart(fileDir) {
        global datos, EditControls  ; Asegurar acceso a los datos y los Edit



    ; Asegurarse de que haya un Case Number
    caseNumber := datos.Get("C&ase Number")
    if !caseNumber or caseNumber = "" {
        MsgBox("No hay 'Case Number' definido, coloque uno.","Requiere Case Number", "48" )
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
        suffix := (i = 0) ? "" : Format("_{}",i)
        filePath := fileDir "\" caseNumber suffix ".json"
        if !FileExist(filePath)
            break
        i++
    }

    ; Guardar el JSON
    jsonText := Jxon_Dump(datos, 4)
    FileAppend(jsonText, filePath, "UTF-8")

    MsgBox("Guardado como:`n" filePath, "Guardando Caso", "64")
    return
}

LoadMapFromFile(filePath) {
    if !FileExist(filePath) {
        MsgBox("Archivo no encontrado: " filePath,"Error de Path","16")
        return   ; Devuelve mapa vacío
    }
    

    fileContent := FileRead(filePath, "UTF-8")
    
    MsgBox("Cargado desde el archivo: " filePath,"Cargando Caso","64")

    return fileLoaded := Jxon_Load(&fileContent)
}

SkrvGui.SetFont("s34 bold", "Segoe UI")  ; Tamaño 14, negrita, fuente bonita
BtnSKRIVE := SkrvGui.Add("Button", "x50 y500 w450 h60 ", "S&KRIVE!")



Ejecutar_Autom_Python() {
    global datos, EditControls,fileDir1,rutaPython,rutaScript, logPath
    UpdateDataFromEdits()
    IntPhBttm(false)
    RemoteSessionBuild() 
    Sleep(1000)
    EmailBld(false,( datos["Issue"] "`n" ),false ,datos["EmailInputEdit"] "`n" ,false)
    
    BackupSave := fileDir1 . Format("\DNG_{}.json", datos["&Dongle"] )
    if FileExist(BackupSave) {
        FileDelete(BackupSave)
    }
    FileAppend(Jxon_Dump(datos,4), BackupSave, "UTF-8" )


    
    UpdateDataFromEdits()
    
    ;  global datos, edits, myGui
    jsonPath := A_WorkingDir . "\NewCase.json"
    if FileExist(jsonPath) {
        FileDelete(jsonPath)
    }
    ; args := "--seccion1 --seccion2 --seccion3 --seccion4 --seccion5 "
    Loop 5 {
        seccions := Format("seccion{}", A_Index)
        datos[seccions] := 1
        arguments := Format("--seccion{} ", A_Index)
        args .= arguments
        
    }

    if (CheckBoxEmailManual.Value) {
        args .= Format("--seccionE ")
        
    }
    
    FileAppend(Jxon_Dump(datos,4), jsonPath, "UTF-8" )

    
    SkrvGui.Minimize()
    


    CRM2_Exe := A_WorkingDir "\CRM2.exe"
    if FileExist(CRM2_Exe) {
        comand := Format('"{}" {}', CRM2_Exe, args)
        A_Clipboard := comand
        RunWait(comand, , "Hide")

    } else {
        comand := Format('"{}" "{}" {}', rutaPython, rutaScript, args) 
        A_Clipboard := comand
        RunWait(comand, , "Hide")
    }
}
AutomGUI(){ 
    global datos, EditControls,fileDir1,rutaPython,rutaScript
    UpdateDataFromEdits()
    IntPhBttm(false)
    RemoteSessionBuild() 
    Sleep(1000)
    EmailBld(false,( datos["Issue"] "`n" ),false ,datos["EmailInputEdit"] "`n" ,false)
    
    BackupSave := fileDir1 . Format("\DNG_{}.json", datos["&Dongle"] )
    if FileExist(BackupSave) {
        FileDelete(BackupSave)
    }
    FileAppend(Jxon_Dump(datos,4), BackupSave, "UTF-8" )

    SectionsGui := Gui("+AlwaysOnTop", "Enviar a Python")
    SectionsGui.SetFont("s25 bold", "Segoe UI")
    SectionsGui.AddText("x20 y20", "Acciones:")
    SectionsGui.SetFont("s25 norm", "Segoe UI")

    checkboxStates := Map()
    sections := [
        "Sección 1 [Llenar Datos Cliente]",
        "Sección 2 [Guardar CaseNumber y Survey]",
        "Sección 3 [Agregar notas y enviar Email]",
        "Sección 4 [Terminar Workflow]",
        "Sección 5 [Resolver caso]"
    ]

    global tooltipText := ""  ; acumulador para todas las descripciones
    y := 60

    loop sections.Length {
        fullText := sections[A_Index]

        ; Extraer nombre visible (antes del [)
        visibleText := RegExReplace(fullText, "\s*\[.*\]", "")

        ; Extraer el contenido entre [ ]
        description := RegExReplace(fullText, ".*\[\s*(.*?)\s*\]", "$1")

        ; Agregar checkbox con texto limpio
        cbx := SectionsGui.AddCheckbox("x20 y" y, visibleText)
        checkboxStates["seccion" A_Index] := cbx

        tooltipText .= visibleText ": " description "`n"
        y += 50
    }

    ; Botón de ayuda "?" al final
    helpBtn := SectionsGui.AddButton("x20 y" y " w25 h45", "?")
    helpBtn.OnEvent("Click", (*) => ShowTooltip())

    ; Activar todos los checkboxes por defecto
    for claves in [ "seccion2", "seccion3", "seccion4"] {
        checkboxStates[claves].Value := true
    }
    SectionsGui.SetFont("s15 bold", "Segoe UI")  ; Tamaño 14, negrita, fuente bonita

    SectionsGui.AddButton("x+20", "SKRIVE!").OnEvent("Click", (*) => EjecutarPython())
    SectionsGui.SetFont("s15 norm", "Segoe UI")  ; Tamaño 14, negrita, fuente bonita

    SectionsGui.Show()
    ShowTooltip() {
        global tooltipText
        ToolTip(tooltipText)
        SetTimer(() => ToolTip(), -4000)  ; Oculta después de 4 segundos
    }
    EjecutarPython() {
        global datos, EditControls,fileDir1,rutaPython,rutaScript
        UpdateDataFromEdits()
        ; updatedataformedits
        ;  global datos, edits, myGui
        jsonPath := A_WorkingDir . "\NewCase.json"
        if FileExist(jsonPath) {
            FileDelete(jsonPath)
        }
        Loop 5{
            seccions := Format("seccion{}", A_Index)
            datos[seccions] := ""
        }

        args := ""
        
        SectionsGui.hide()


        if datos["Case Link"] != "" {
            if checkboxStates["seccion1"].Value{
                for clave, chk in checkboxStates {
                    if chk.Value and !(clave == "seccion1") {
                        datos[clave] := chk.Value
                        args .= "--" clave " "
                    }
                }
            }
            for clave, chk in checkboxStates {
                    if chk.Value {
                        datos[clave] := chk.Value
                        args .= "--" clave " "
                    }
            }
        }

        if datos["Case Link"] == ""{
            checkboxStates["seccion1"].Value := true
            for clave, chk in checkboxStates {
                if chk.Value{
                    datos[clave] := chk.Value
                    args .= "--" clave " "  
                }
            }
        }
        if (CheckBoxEmailManual.Value) {
            args .= Format("--seccionE ")
        }
        

        FileAppend(Jxon_Dump(datos,4), jsonPath, "UTF-8" )


        SkrvGui.Minimize()
        

        CRM2_Exe := A_WorkingDir "\CRM2.exe"
        if FileExist(CRM2_Exe) {
            comand := Format('"{}" {}', CRM2_Exe, args)
            A_Clipboard := comand
            ; RunWait(comand)
            RunWait(comand, , "Hide")

        } else {
            comand := Format('"{}" "{}" {}', rutaPython, rutaScript, args) 
            A_Clipboard := comand
            RunWait(comand, , "Hide")

        }
    }

}

AutomGUI_SinCondiciones(){  
    global datos, EditControls,fileDir1,rutaPython,rutaScript
    UpdateDataFromEdits()
    IntPhBttm(false)
    RemoteSessionBuild() 
    Sleep(1000)
    EmailBld(false,( datos["Issue"] "`n" ),false ,datos["EmailInputEdit"] "`n" ,false)
    
    BackupSave := fileDir1 . Format("\DNG_{}.json", datos["&Dongle"] )
    if FileExist(BackupSave) {
        FileDelete(BackupSave)
    }
    FileAppend(Jxon_Dump(datos,4), BackupSave, "UTF-8" )

    SectionsGui := Gui("+AlwaysOnTop", "Enviar a Python")
    SectionsGui.SetFont("s25 bold", "Segoe UI")
    SectionsGui.AddText("x20 y20", "Acciones:")
    SectionsGui.SetFont("s25 norm", "Segoe UI")

    checkboxStates := Map()
    sections := [
        "Sección 1 [Llenar Datos Cliente]",
        "Sección 2 [Guardar CaseNumber y Survey]",
        "Sección 3 [Agregar notas y enviar Email]",
        "Sección 4 [Terminar Workflow]",
        "Sección 5 [Resolver caso]"
    ]

    global tooltipText := ""  ; acumulador para todas las descripciones
    y := 60

    loop sections.Length {
        fullText := sections[A_Index]

        ; Extraer nombre visible (antes del [)
        visibleText := RegExReplace(fullText, "\s*\[.*\]", "")

        ; Extraer el contenido entre [ ]
        description := RegExReplace(fullText, ".*\[\s*(.*?)\s*\]", "$1")

        ; Agregar checkbox con texto limpio
        cbx := SectionsGui.AddCheckbox("x20 y" y, visibleText)
        checkboxStates["seccion" A_Index] := cbx

        tooltipText .= visibleText ": " description "`n"
        y += 50
    }

    ; Botón de ayuda "?" al final
    helpBtn := SectionsGui.AddButton("x20 y" y " w25 h45", "?")
    helpBtn.OnEvent("Click", (*) => ShowTooltip())

    ; Activar todos los checkboxes por defecto
    for _, checkBoxes in checkboxStates {
        checkBoxes.Value := true
    }
    SectionsGui.SetFont("s15 bold", "Segoe UI")  ; Tamaño 14, negrita, fuente bonita

    SectionsGui.AddButton("x+20", "SKRIVE!").OnEvent("Click", (*) => EjecutarPython_sinCondicionales())
    SectionsGui.SetFont("s15 norm", "Segoe UI")  ; Tamaño 14, negrita, fuente bonita

    SectionsGui.Show()

    ShowTooltip() {
        global tooltipText
        ToolTip(tooltipText)
        SetTimer(() => ToolTip(), -4000)  ; Oculta después de 4 segundos
    }

    
    EjecutarPython_sinCondicionales() {
        global datos, EditControls,fileDir1,rutaPython,rutaScript
        UpdateDataFromEdits()
        jsonPath := A_WorkingDir . "\NewCase.json"
        if FileExist(jsonPath) {
            FileDelete(jsonPath)
        }
        Loop 5{
            seccions := Format("seccion{}", A_Index)
            datos[seccions] := ""
        }

        args := ""
        
        SectionsGui.hide()

        for clave, chk in checkboxStates {
            if chk.Value {
                datos[clave] := chk.Value
                args .= "--" clave " "

            }
        }
        if (CheckBoxEmailManual.Value) {
            args .= Format("--seccionE ")
        }
        

        FileAppend(Jxon_Dump(datos,4), jsonPath, "UTF-8" )

        
        
        SkrvGui.Minimize()
        
        CRM2_Exe := A_WorkingDir "\CRM2.exe"
        if FileExist(CRM2_Exe) {
            comand := Format('"{}" {}', CRM2_Exe, args)
            A_Clipboard := comand
            RunWait(comand, , "Hide")

        } else {
            comand := Format('"{}" "{}" {}', rutaPython, rutaScript, args) 
            A_Clipboard := comand
            RunWait(comand, , "Hide")

        }

    }

}
Skrive_link_btn() {
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit

    UpdateDataFromEdits() ; 💡 Refresca `datos` con los valores actuales de los Edits
    if GetKeyState("Alt") {  ; Verifica si Alt está presionado
        AutomGUI_SinCondiciones()
        return
    } else { 
        AutomGUI()
        return
    }
}

Skrive_link_btn2() {
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit

    UpdateDataFromEdits() ; 💡 Refresca `datos` con los valores actuales de los Edits
    if GetKeyState("Alt") {  ; Verifica si Alt está presionado
        Automatic()
        return
    } else { 
        Ejecutar_Autom_Python()
        return
    }
}


BtnSKRIVE.OnEvent("ContextMenu", (*) => Skrive_link_btn())
BtnSKRIVE.OnEvent("Click", (*) => Skrive_link_btn2())


BtnKlokken := SkrvGui.Add("Button", "x600 y500 w200 h60", "Klokken!")
BtnKlokken.OnEvent("Click", (*) => Klokken())

Klokken() {
    SkrvGui.Minimize()

    ; Ejecutar un archivo .exe (ajusta la ruta según sea necesario)
    exePath := A_WorkingDir . "\Klokken_App.exe"
    if FileExist(exePath){
        Run(exePath)
    }
    else{
        global KlokkenGui  ; Asegúrate de que este GUI esté definido globalmente (probablemente en Klokken.ahk)
        KlKnWidth := 320
        KlKnHeigh := 170  
        KlokkenGui.Show("w" KlKnWidth " h" KlKnHeigh)
    }
}
SkrvGui.SetFont("s12 norm", "Segoe UI")  ; Tamaño 14, negrita, fuente bonita


CheckBoxEmailManual := SkrvGui.AddCheckbox("x850 y500 vCheckBoxEmailManual", "Manual Email?")




InputLine("Probing Questions (Add Info)",50,50,800 ,565, 250,,"paste",600,100)

UpdateDataFromEdits()
CallBack := SkrvGui.Add("Button", " x860 y600  w180 h30", "2nd Line - Call Back")
Earlyacces := SkrvGui.Add("Button", " x860 y635  w100 h30", "Early Access")

PhotosTab := SkrvGui.Add("Button", " x860 y670  w75 h30", "Photos")

PhotosTab.OnEvent("Click", (*) =>  PhotosOpenExcel())

PhotosOpenExcel(){
    UpdateDataFromEdits()

    filePath := A_WorkingDir ".\Photos.xlsm"

    excel := ComObject("Excel.Application")
    excel.Visible := true

    workbook := excel.Workbooks.Open(filePath)
    sheet := workbook.Sheets(1)

    sheet.Cells(1, 2).Value := datos["C&ase Number"]
    return

}



CallBack.OnEvent("ContextMenu", (*) =>  callbacktitle())

callbacktitle(){
    UpdateDataFromEdits()
    today := A_Now  ; Obtiene la fecha y hora actual en formato AAAAMMDDHHMMSS
    formattedDate := FormatTime(today, "yyyyMMdd")  ; Formatea la fecha
    A_Clipboard := "Called back to this phone number " datos["&Phone"] " and no body answered the phone. Voice message leaved in order to continue resolving the issue"
    Sleep(500)
    A_Clipboard:= "CALLBACK " formattedDate 
    Sleep(500)
    return

}

CallBack.OnEvent("Click", (*) =>  CallBackCuild())


CallBackCuild() {
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit
    UpdateDataFromEdits()

    A_Clipboard := ("The case will be escalated to second line to continue with the verification.`n User Info`n`n`n`n" "Name: " datos["Name"] "`n`n" "Phone: " datos["&Phone"] "`n`n" "Email: " datos["&Email"] )
    Sleep(500)
    today := A_Now  ; Obtiene la fecha y hora actual en formato AAAAMMDDHHMMSS
    formattedDate := FormatTime(today, "yyyyMMdd")  ; Formatea la fecha

    
    Int2 := "INT " formattedDate
    A_Clipboard := (Int2)
    Sleep(500)
    A_Clipboard := ("CB Escalation - 2nd Line Clinic - " datos["&Company Name"] " - DN:" datos["&Dongle"])
    Sleep(500)
    A_Clipboard := " Buenos días, ayuda para agendar este callback de escalación de 2nd Line Clinic. " 
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
            ["Case Number", datos["C&ase Number"]],
            ["Caller Name", datos["Name"]],
            ["Request/Issue", datos["Issue"]],
            ["Dongle", datos["&Dongle"]],
            ["Company Name", datos["&Company Name"]],
            ["Phone Number", datos["&Phone"]],
            ["BEST CALL BACK TIME WITH TIME ZONE (URGENCY)", "ASAP"]
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

        

        ; Liberar referencias (sin cerrar Excel)
        rango := "", ws := "", wb := ""

        MsgBox("Tabla Copiada al portapapeles", "2ndLine Clinic Info", "64")
        return

            
}

Earlyacces.OnEvent("Click", (*) =>  Earlyacces1())
Earlyacces1(){
    UpdateDataFromEdits()
    today := A_Now  ; Obtiene la fecha y hora actual en formato AAAAMMDDHHMMSS
    formattedDate := FormatTime(today, "yyyyMMdd")  ; Formatea la fecha
    A_Clipboard := "3Q " formattedDate "EARLY ACESS"
    Sleep(500)

    earlyBody := "Dear 3rd Line team`n" "I hope this message finds you well. I am writing to request early access to Unite III for this Company. Their licenses didn't migrate to the cloud, and the modules and labs are missing. the dongle has access to Unite III but the company it's still in Unite" datos["Software Version"] "`n" "Dongle ID " datos["&Dongle"] "`n" "SID " datos["S&ID"] "`n" "Company Name " datos["&Company Name"] "`n" "Email " datos["&Email"] "`n" "GUID " datos["GUI"] "`n" "Phone " datos["&Phone"] "`n" "TV " datos["TV ID"] "`n" "TVPSS " datos["TV PSS"] "`n"

    A_Clipboard := earlyBody
    Sleep(500)
    return

}




if FileExist(InfoFile){

    fileContent := FileRead(InfoFile, "UTF-8")

    hours := Jxon_Load(&fileContent)
}
else
    hours := Map(
        "Hora In",  "12:00",
        "Hora Out", "12:00",
        "Brk 1",     "12:00",
        "Brk 2",     "12:00",
        "Lunch",     "12:00"
    )

; GUI principal con botón flotante
HoursBtn := SkrvGui.AddButton("x1150 y675 w75 h30", "Shift")
HoursBtn.OnEvent("Click", ShowHourEditor)

ShowHourEditor(*) {
    global InfoFile

    hourGui := Gui("Resize", "Editar Horarios")
    hourGui.SetFont("s12 norm", "Segoe UI")  ; Tamaño 14, negrita, fuente bonita

    controls := Map()

    ; Crear campos de edición
    for k, v in hours {
        hourGui.AddText("", k)
        controls[k] := hourGui.AddEdit("w100", v)
    }

    ; Botón para guardar
    hourGui.AddButton("xm w100", "Guardar").OnEvent("Click", (*) =>guardarHoras(controls, hourGui)) 
    guardarHoras(controls, hourGui){
        for k, ctrl in controls {
            hours[k] := ctrl.Text
        }
        FileDelete(InfoFile)
        FileAppend(Jxon_Dump(hours, "`t"), InfoFile)
        hourGui.Destroy()
        MsgBox "Horas guardadas en 'A_Info.json'"
    }

    hourGui.Show()
}



; ; ; BtnSKRIVE.OnEvent("ContextMenu", (*) => SKRIVE_Autom(true))
; ; ; BtnSKRIVE.OnEvent("Click", (*) => SKRIVE_Autom(false))


; ; ; SKRIVE_Autom(SKR_Boool){
; ; ;     global datos, EditControls  ; Asegurar acceso a los datos y los Edit
    

; ; ;     ; Si el usuario elige "No", cancela la acción
; ; ;     if (SKR_Boool == true) {
; ; ;         if GetKeyState("Alt") {  ; Verifica si Alt está presionado
; ; ;             Automatic()
; ; ;             return
; ; ;         } else {
; ; ;             Sleep(6000)  ; Tu función alternativa si no se presionó Alt
; ; ;             CRM()
; ; ;             return
; ; ;         }
; ; ;     }
; ; ;     ForwardToDynamics()
; ; ;     return
; ; ; }


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
    global datos, EditControls, NmOfSteps ; Asegurar acceso a los datos y los Edit
    stepsList := ""
    Loop NmOfSteps
    {
        i := A_Index
        controlName := "Step " . i
        stepText := EditControls[controlName].Value
        if (stepText != "")
        {
            stepsList .= Format("- {}`n", stepText)
        }
    }
    ; Agregar información adicional
    issueText := EditControls["Issue"].Value

    if (issueText == "") {
        datos["RMTSS"] := "Accessed to TV session`n" "Logs and photos are in a .zip at an Internal NOTE `n" . Format("Ask the customer to reproduce the issue: {}`n")
        MsgBox("Introduce un Issue por favor.", "Error de Info", "16")
        return
    }
    stepsList := "Accessed to TV session`n" "Logs and photos are in a .zip at an Internal NOTE `n" . Format("Ask the customer to reproduce the issue: {}`n", issueText) . stepsList

    datos["RMTSS"] := stepsList
    ; Copiar al portapapeles
    A_Clipboard := stepsList
    Sleep 500
}
; === INICIO: Bloque de Checkboxes (Checklist de seguimiento) ===
items := [
    "Opening / Confirm Info", 
    "Previous Cases / Add Info", 
    "Paraphrasing", 
    "Reproduce Issue",
    "What Where When", 
    "Probing Questions", 
    "Waiting times / Reason of them", 
    "Engagement", 
    "Helpjuice / Phone in ContactInfo?",
    "Logs photos", 
    "Email", 
    "Expectations", 
    "Recap / Offer More assistance", 
    "Case Number / Survey"
]

checkboxes := []  ; Lista para guardar referencias

xStart := 940
yStart := 50
spacing := 35
SkrvGui.SetFont("s15 norm", "Segoe UI")  ; Tamaño 14, negrita, fuente bonita

for i, item in items {
    cb := SkrvGui.AddCheckbox("x" xStart " y" yStart + (i-1)*spacing " vcb" i, item)
    checkboxes.Push(cb)
}

; Botón para reiniciar todos los checkboxes
resetBtn := SkrvGui.AddButton("x" xStart " y" yStart + (items.Length)*spacing + 10, "Reiniciar Checkboxes")
resetBtn.OnEvent("Click", (*) => DelCheck())
DelCheck(){
    for _, cb in checkboxes {
        cb.Value := false
    }
}
; === FIN: Bloque de Checkboxes ===

SkrvGui.SetFont("s12 norm", "Segoe UI")  ; Tamaño 14, negrita, fuente bonita


tab.UseTab(3) ;(Email)
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit

    EmailTab := SkrvGui.Add("Button", " x50 y50  w75 h30", "Email")
    EmailTabedit := SkrvGui.Add("Edit", " x50 y100  w990 h580", "")
    ; Guardar la referencia del Edit
    global EditControls
    EditControls["EmailInputEdit"] := EmailTabedit  

    ; Capturar cambios en el input
    global datos
    datos["EmailInputEdit"] := EmailTabedit
    EmailTabedit.OnEvent("Change", (*) => datos["EmailInputEdit"] := EmailTabedit.Value)

    EmailTab.OnEvent("ContextMenu", (*) => emTit())

    emTit(){
        UpdateDataFromEdits()

        global datos, EditControls  ; Asegurar acceso a los datos y los Edit
        datos["EmTitl"] := ""

        emTITl:="Regarding your case number " datos["C&ase Number"]
        datos["EmTitl"] := emTITl


        A_Clipboard:= emTITl

    }

    EmailTab.OnEvent("Click", (*) =>  EmailButom())
    EmailButom() {
        global datos, EditControls  ; Asegurar acceso a los datos y los Edit

        UpdateDataFromEdits() ; 💡 Refresca `datos` con los valores actuales de los Edits
        if GetKeyState("Alt") {  ; Verifica si Alt está presionado
            EmailBld(false, (datos["Issue"] "`n"),false,false ,false)
            return
        } else { 
            EmailBld(false,( datos["Issue"] "`n" ),false ,datos["EmailInputEdit"] "`n" ,false)
            return
        }
    }





EmailBld(Greeting?, Issue? , Body?, Recommend? ,CloseSurvey?){
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit
    datos["EmailFinal"] := ""


    UpdateDataFromEdits()

    G1:=Format("Dear {}", datos["&Company Name"] )
    G2:="I hope this email finds you well."
    G3:= "I am writing in reference to your recent call to our technical support center regarding your issue "
    greetingdflt := (  G1 "`n`n`n`n" G2 "`n`n`n`n" G3)

    bodydflt := ("I'm pleased to inform you that we have resolved this complaint or request, during our interaction we have " datos["Solution"] "`n`n`n`nBringing to your attention, it occurred due to " datos["RC"] "." "`n`n`n`n") 
    Issuedflt := ("Case got corrupted, we have reimported the streams and the case is now working as expected.")

    RecommendationDflt := "Additionally,  to prevent similar issues in the future, we recommend shutting down your computer every night and ensuring it remains connected during the scanning process"

    CloseSurveyDfflt:= ("Thank you for allowing me to assist you today. I would greatly appreciate it if you could take a moment to provide feedback on my service by completing the following survey: `n`n`n`n"  "https://" datos["Sur&vey"] "`n`n`n`n At the end of the survey, please leave a note about your experience during this interaction. " "`n`n`n" "Your patience and understanding throughout the process are greatly appreciated, and your feedback will help us continue improving our services as we strive for excellence." "`n`n`n")

    

    
    
    Greeting := (Greeting? Greeting:greetingdflt )
    Issue := (Issue? Issue:Issuedflt )
    Body := (Body? Body:Bodydflt )
    Recommend := (Recommend? Recommend:RecommendationDflt )
    CloseSurvey := (CloseSurvey? CloseSurvey:CloseSurveyDfflt )
    emailFinal := (Greeting  Issue "`n`n`n`n" Body  "`n`n`n`n" Recommend "`n`n`n`n`n`n" CloseSurvey)
    emailFinal1 := (Greeting  Issue "`n`n`n`n" Body  "`n`n`n`n" Recommend "`n`n`n`n`n`n" )
    
    datos["EmailFinal"] := emailFinal1
    Sleep(1)
    A_Clipboard := emailFinal

}


tab.UseTab()


isSkrvVisible := true

^+u:: {
    global isSkrvVisible
    if isSkrvVisible {
        SkrvGui.Minimize()
        isSkrvVisible := false
        return
    } else {
        SkrvGui.Show("w" SkrWidth " h" SkrHeigh)  ; Esto es correcto
        isSkrvVisible := true
        return
    }
}




; ^+u::SkrvGui.Show()  ; Ctrl + L para mostrar la GUI


; 🔹 Atajos de teclado para cambiar pestañas
^!i::tab.Value := 1  ; Ctrl + Alt + I -> information
^!m::tab.Value := 2  ; Ctrl + Alt + R -> Rmt Sess
^!e::tab.Value := 3  ; Ctrl + Alt + E -> Email

^!f::Automatic()  ; Ctrl + Alt + F -> Copiar todo al portapapeles
^!s::Skrive_link_btn()  ; Ctrl + Alt + S -> Gui de Acciones
^!d::Skrive_link_btn2()  ; Ctrl + Alt + S -> Automatizar CRM


!d::EditControls["&Dongle"].Value := A_Clipboard  ; Ctrl + Alt + S -> Fwd2Skrive
!p::EditControls["&Phone"].Value := A_Clipboard  ; Ctrl + Alt + S -> Fwd2Skrive
!e::EditControls["&Email"].Value := A_Clipboard  ; Ctrl + Alt + S -> Fwd2Skrive
!c::EditControls["&Company Name"].Value := A_Clipboard  ; Ctrl + Alt + S -> Fwd2Skrive
!i::EditControls["S&ID"].Value := A_Clipboard  ; Ctrl + Alt + S -> Fwd2Skrive
!v::EditControls["Sur&vey"].Value := A_Clipboard  ; Ctrl + Alt + S -> Fwd2Skrive
!r::EditControls["Scanne&r S/N"].Value := A_Clipboard  ; Ctrl + Alt + S -> Fwd2Skrive
; !a::EditControls["C&ase Number"].Value := A_Clipboard  ; Ctrl + Alt + S -> Fwd2Skrive
!a:: Alt_a()

Alt_a(){
    ClipWait  ; Espera a que el portapapeles tenga contenido
    ; Elimina los corchetes y conserva solo el texto interno
    Clipbd := A_Clipboard
    texto := RegExReplace(Clipbd, ".*\[(.*?)\].*", "$1")
    ; Asigna el texto procesado al campo correspondiente
    A_Clipboard :=texto
    EditControls["C&ase Number"].Value := texto
    return
}




!t::DescriptionGUI(false)  ; Ctrl + Alt + S -> Fwd2Skrive
!h::{
    IntPhBttm(true) 
    Sleep(500) 
    IntPhBttm(false)
}  ; Ctrl + Alt + S -> Fwd2Skrive
!s::SaveBttm(false,fileDir1)  ; Ctrl + Alt + S -> Fwd2Skrive
!y::Ctrl_a()  ; Ctrl + Alt + S -> Fwd2Skrive
^!/:: {
    Run("taskkill /F /IM chrome.exe", , "Hide")
    Run("taskkill /F /IM Chrome_Activator.exe", , "Hide")
    Run("taskkill /F /IM chromedriver.exe", , "Hide")
    Run("taskkill /F /IM Five9.exe", , "Hide")
    Run("taskkill /F /IM CRM2.exe", , "Hide")
    Run("taskkill /F /IM Klokken_App.exe", , "Hide")
    Run("taskkill /F /IM python.exe", , "Hide")
    Run("taskkill /F /IM pythonw.exe", , "Hide")
    Run("taskkill /F /IM Skrive.exe", , "Hide")
}

ChrActive_Exe := A_WorkingDir "\Chrome_Activator.exe"
if FileExist(ChrActive_Exe) {
    ; Ejecutar el script con los argumentos
    RunWait(ChrActive_Exe, , "Hide")

} else {
    rutaScriptAct := A_WorkingDir "\Chrome_Activator.py"
    comand := Format('"{}" "{}"', rutaPython, rutaScriptAct) 
    A_Clipboard := comand
    RunWait(comand, , "Hide")

}
