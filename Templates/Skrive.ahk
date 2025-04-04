#Requires AutoHotkey v2.0

global datos := Map()  ; Crear un diccionario global
global EditControls := Map()  ; Guarda los controles Edit para acceder después


#Include  TeamViewer.ahk  
#Include Json.ahk
Persistent 1  ; Establece el script como persistente
fileDir := A_Desktop "\CasesJSON"  ; Solicitar la ruta del directorio
; ---------------------------------------------------------------------------------------------------------------------------------

SkrvGui := Gui(, "SKRIVE")
SkrvGui.Opt("+Resize +MinSize400x300 +MaximizeBox +MinimizeBox") ; Hace la ventana redimensionable
SkrvGui.SetFont("s12")

tab := SkrvGui.Add("Tab3", "x10 y10 w900 h600", ["Information","Remote Session", "Add Info", "Email", "Settings"])

SkrvGui.OnEvent("Size", (*) => AutoResize())  ; Evento para ajustar tamaño dinámicamente
SkrWidth := 1070
SkrHeigh := 700 
BtnHeigh:=30
btnSze:= 200

SkrvGui.Show("w" SkrWidth " h" SkrHeigh)  ; Esto es correcto


AutoResize(*) {
    global SkrvGui, tab
    SkrvGui.GetPos(&x, &y, &w, &h)  ; Obtiene el tamaño actual de la ventana
    tab.Move(10, 10, w - 20, h - 40)  ; Ajusta tamaño y posición
}

InputLine(Name, x?, xedit?, widthedit?, yOffset?, btnSze?, BtnHeigh?, buttomDefltPaste?) {
    Name2Text := Name

    ; Valores predeterminados
    x := IsSet(x) ? x : 50
    xedit := IsSet(xedit) ? xedit : 260
    widthedit := IsSet(widthedit) ? widthedit : 770
    yOffset := IsSet(yOffset) ? yOffset : 100
    btnSze := IsSet(btnSze) ? btnSze : 80
    BtnHeigh := IsSet(BtnHeigh) ? BtnHeigh : 30

    btn := SkrvGui.Add("Button", Format("x{} y{} w{} h{}", x, yOffset, btnSze, BtnHeigh), Name2Text)
    edit := SkrvGui.Add("Edit", Format("x{} y{} w{} h{}", xedit, yOffset, widthedit, BtnHeigh), "")

    if buttomDefltPaste == true {
        btn.OnEvent("Click", (*) => edit.Value := A_Clipboard)  ; Clic izquierdo: pega desde el portapapeles
        btn.OnEvent("ContextMenu", (*) => edit.Value := "")    ; Clic derecho: limpia el Edit
    } 
    if buttomDefltPaste == false {
        btn.OnEvent("Click", (*) => edit.Value := A_Clipboard)  ; Clic izquierdo: pega desde el portapapeles
        btn.OnEvent("ContextMenu", (*) => Ctrl_6())  ; Clic izquierdo: limpia el Edit
    }

    

    global datos
    datos[Name2Text] := ""  ; Asegura que el valor inicial sea vacío

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
    global datos

    SoftwGui := Gui()
    SoftwGui.Opt("+AlwaysOnTop")
    SoftwGui.SetFont("s10")

    SoftwGui.Add("Text", "x10 y10", "Select Software Versions:")
    
    SoftwGui.Add("Text", "x10 y30", "Module:")
    comboBox1 := SoftwGui.Add("ComboBox", "x10 y50 w150", ["Unite", "TRIOS Module", "Dental Desktop", "Model Builder", "Automate"])
    
    SoftwGui.Add("Text", "x180 y30", "Version:")
    comboBox2 := SoftwGui.Add("ComboBox", "x180 y50 w150", ["1.7.83.0", "1.7.8.1", "1.8.8.0", "1.8.5.1"])

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
    Outbtn.OnEvent("ContextMenu", (*) => outedit.Value := "")       ; Clic derecho: limpia el Edit
    global datos
    datos[Name2Text] := outedit.Value
    outedit.OnEvent("Change", (*) => datos[Name2Text] := outedit.Value)

}

Hotkeys(Name, yOffset){
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
    global datos
    today := A_Now  ; Obtiene la fecha y hora actual en formato AAAAMMDDHHMMSS
    formattedDate := FormatTime(today, "yyyyMMdd")  ; Formatea la fecha

    A_Clipboard := (IntBool ? "PHONECALL " : "INT ") formattedDate  ; Usa operador ternario

    return  ; Buena práctica incluirlo, aunque no es obligatorio
}
; ---------------------------------------------------------------------------------------------------------------------------------

; PESTAÑA 1 (General)
tab.UseTab(1)
y := 50 ; Posición inicial en Y
x:= 260
spacing := 35 ; Espaciado entre elementos
Insize:=390

SkrvGui.Add("Text",)
tvsize := (750/2)-(spacing*2)

InputLine("TV ID",,(btnSze/2)+55,tvsize, y,btnSze/2,BtnHeigh, false), y
InputLine("TV PSS",465,570,tvsize, y,btnSze/2,BtnHeigh, false), y 
Int_PhBtt := SkrvGui.Add("Button", Format("x880 y{} w{} h{}", y, btnSze-100, BtnHeigh), "INT-PH"), y += spacing
Int_PhBtt.OnEvent("Click", (*) => IntPhBttm(false))
Int_PhBtt.OnEvent("ContextMenu", (*) => IntPhBttm(true))







 
InputLine("Name",,,, y,btnSze,BtnHeigh, true), y += spacing
InputLine("&Phone",,,, y,btnSze,BtnHeigh, true), y += spacing
InputLine("&Email",,,, y,btnSze,BtnHeigh, true), y += spacing
InputLine("&Company Name",,,, y,btnSze,BtnHeigh, true), y += spacing
SoftwareVersionSelect("Software Version", y,btnSze,BtnHeigh), y += spacing

InputLine("&Dongle",,,, y,btnSze,BtnHeigh, true), y += spacing
InputLine("SID",,,, y,btnSze,BtnHeigh, true), y += spacing
InputLine("GUI",,,, y,btnSze,BtnHeigh, true), y += spacing
InputLine("Scanner S/N",,,, y,btnSze,BtnHeigh, true), y += spacing
InputLine("PC ID",,,, y,btnSze,BtnHeigh, true), y += spacing
InputLine("Case Number",,,, y,btnSze,BtnHeigh, true), y += spacing
; ---------------------------------------------------------------------------------------------------------------------------------
descripton:= SkrvGui.Add("Button", Format("x50 y{} w{} h{}", y,btnSze,BtnHeigh), "Descrip&tion"),
descripton.OnEvent("ContextMenu", (*) => DescriptionGUI(true))
descripton.OnEvent("Click", (*) => DescriptionGUI(false))
    

#Include CompareMaps.ahk
DescriptionGUI(bool) {
    global datos  ; Hacer accesible el mapa de datos

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


BtnSaveInfo := SkrvGui.Add("Button", "x50 y645 w150 h40", "&Save Info-Load Info")
clear := SkrvGui.Add("Button", " x205 y645  w45 h30", "Clear")

clear.OnEvent("Click", (*) => ClearAll())

ClearAll() {
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit

    respuesta := MsgBox("¿Está seguro de que desea borrar todos los campos?", "Borrar Datos", "YesNo")
    if (respuesta = "No")  ; Si el usuario elige "No", cancela la acción
        return
    Sleep(1000)
    for key in datos {
        if (EditControls.Has(key)) {  
            EditControls[key].Value := ""  
        }
        datos[key] := ""  
    }

    UpdateEditsFromData()  
}




BtnSaveInfo.OnEvent("ContextMenu", (*) => SaveBttm(true))
BtnSaveInfo.OnEvent("Click", (*) => altcheck3())

altcheck3() {
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
        SaveBttm(false)  ; Tu función alternativa si no se presionó Alt
    }
}

; Función para abrir una nueva ventana
SaveBttm(SvBool) {

    global datos
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
    global datos


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
UpdateEditsFromData() {
    global datos, EditControls  ; Asegurar acceso a los datos y los Edit

    for key, value in datos {
        if (EditControls.Has(key)) {  ; Usar Has() en lugar de HasKey()
            EditControls[key].Value := value
        }
    }
}





; ---------------------------------------------------------------------------------------------------------------------------------

; PESTAÑA 2 (View)
tab.UseTab(2)
SkrvGui.GetPos(&x, &y, &SkrWidth, &SkrHeigh)  ; Obtiene tamaño de la ventana
offset:=55
InputLine("Probing Questions (Add Info)",,50,990 ,offset, 250,,"paste",50+45,580)
SkrvGui.Show()



; PESTAÑA 3 (Settings)
tab.UseTab(4)


; Salir del modo de pestañas
tab.UseTab()

SkrvGui.Show()
^+u::SkrvGui.Show()  ; Ctrl + L para mostrar la GUI


; 🔹 Atajos de teclado para cambiar pestañas
^!i::tab.Value := 1  ; Alt + G -> information
^!r::tab.Value := 2  ; Alt + V -> Rmt Sess
^!a::tab.Value := 3  ; Alt + S -> Add Inf
^!e::tab.Value := 4  ; Alt + S -> Email







 



