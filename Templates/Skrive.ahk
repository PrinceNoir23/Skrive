#Requires AutoHotkey v2.0

global datos := Map()  ; Crear un diccionario global

#Include  TeamViewer.ahk  
#Include Json.ahk
Persistent 1  ; Establece el script como persistente

; ---------------------------------------------------------------------------------------------------------------------------------

saveData(key, value) {
    global datos
    datos[key] := value
}

; saveData("usuario", "Carlos")
; saveData("edad", 30)

; MsgBox "Usuario: " datos["usuario"] "`nEdad: " datos["edad"]


; ---------------------------------------------------------------------------------------------------------------------------------
; Definir las combinaciones de teclas y sus funciones


SkrvGui := Gui(, "SKRIVE")
SkrvGui.Opt("+Resize +MinSize400x300 +MaximizeBox +MinimizeBox") ; Hace la ventana redimensionable
SkrvGui.SetFont("s12")

tab := SkrvGui.Add("Tab3", "x10 y10 w1000 h600", ["Information", "Add Info", "Email", "Settings"])

SkrvGui.OnEvent("Size", (*) => AutoResize())  ; Evento para ajustar tamaño dinámicamente
SkrWidth := 1070
SkrHeigh := 650 
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
    
    ; Asignar valores predeterminados si los parámetros no están definidos
    x := IsSet(x) ? x : 50
    xedit := IsSet(xedit) ? xedit : 260
    widthedit := IsSet(widthedit) ? widthedit : 750
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
    edit.OnEvent("Change", (*) => datos[Name2Text] := edit.Value)
}



SoftwareVersionSelect(Name, yOffset, btnSze, BtnHeigh) {
    Name2Text := Name
    btnSoftware := SkrvGui.Add("Button", Format("x50 y{} w{} h{}", yOffset, btnSze, BtnHeigh), Name2Text)
    editSoftware := SkrvGui.Add("Edit", Format("x260 y{} w750 h{}", yOffset, BtnHeigh), "")
    
    btnSoftware.OnEvent("Click", (*) => SoftwareVersionGUI(editSoftware, Name2Text))

    global datos
    datos[Name2Text] := editSoftware.Value

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
Int_PhBtt.OnEvent("Click", (*) => IntPhBttm())
Int_PhBtt.OnEvent("ContextMenu", (*) => IntPhBttm())


IntPhBttm(){

}




 
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
descripton:= SkrvGui.Add("Button", Format("x50 y{} w{} h{}", y,btnSze,BtnHeigh), "Description"),
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
    DescripGui.Add("Text", "x10 y10", "The description is:`n" issue " on " softwareVersion)
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
Hotkeys("Hotkeys", y) 

BtnSaveInfo := SkrvGui.Add("Button", "w150 h30", "&Save Info")
BtnSaveInfo.OnEvent("ContextMenu", (*) => SaveBttm(true))
BtnSaveInfo.OnEvent("Click", (*) => SaveBttm(false))
; Función para abrir una nueva ventana
SaveBttm(SvBool) {

    global datos
    
    ; Recorrer el mapa y mostrar los valores
    if SvBool==false{

        
        textData2 := Jxon_dump(datos,4) ; ===> convert array to JSON
        ; MsgBox textData2
        newObj := Jxon_load(&textData2) ; ===> convert json back to array

        CompareMaps(newObj, datos)
        CompareMaps(datos, newObj)
                
    }


    if SvBool==true{
        lista := "Datos guardados:`n"
        for key, value in datos {
            lista .= key ": " value "`n"
        }
        MsgBox lista  ; Muestra todos los valores guardados
        return
    }
    
}
; ---------------------------------------------------------------------------------------------------------------------------------

; PESTAÑA 2 (View)
tab.UseTab(2)
SkrvGui.GetPos(&x, &y, &SkrWidth, &SkrHeigh)  ; Obtiene tamaño de la ventana
offset:=55
SkrvGui.Add("Edit", "w" (SkrWidth - offset) " h" (SkrHeigh - (offset*3.5)), "")
SkrvGui.Show()



; PESTAÑA 3 (Settings)
tab.UseTab(3)


; Salir del modo de pestañas
tab.UseTab()

SkrvGui.Show()


; 🔹 Atajos de teclado para cambiar pestañas
^!i::tab.Value := 1  ; Alt + G -> General
^!v::tab.Value := 2  ; Alt + V -> View
^!x::tab.Value := 3  ; Alt + S -> Settings







 



