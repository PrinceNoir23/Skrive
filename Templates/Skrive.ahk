#Requires AutoHotkey v2.0
#Include  TeamViewer.ahk
Persistent 1  ; Establece el script como persistente


; ---------------------------------------------------------------------------------------------------------------------------------
; Definir las combinaciones de teclas y sus funciones


SkrvGui := Gui(, "SKRIVE")
SkrvGui.Opt("+Resize +MinSize400x300 +MaximizeBox +MinimizeBox") ; Hace la ventana redimensionable
SkrvGui.SetFont("s12")

tab := SkrvGui.Add("Tab3", "x10 y10 w1000 h600", ["Information", "Add Info", "Email", "Settings"])

SkrvGui.OnEvent("Size", (*) => AutoResize())  ; Evento para ajustar tamaño dinámicamente
SkrWidth := 1050
SkrHeigh := 650 
BtnHeigh:=30
btnSze:= 200

SkrvGui.Show("w" SkrWidth " h" SkrHeigh)  ; Esto es correcto


AutoResize(*) {
    global SkrvGui, tab
    SkrvGui.GetPos(&x, &y, &w, &h)  ; Obtiene el tamaño actual de la ventana
    tab.Move(10, 10, w - 20, h - 40)  ; Ajusta tamaño y posición
}


InputLine(Name, yOffset, btnSze, BtnHeigh) {
    btn := SkrvGui.Add("Button", Format("x50 y{} w{} h{}", yOffset, btnSze, BtnHeigh), Name)
    edit := SkrvGui.Add("Edit", Format("x260 y{} w750 h{}", yOffset, BtnHeigh), "")
    
    btn.OnEvent("Click", (*) => edit.Value := A_Clipboard)    ; Clic izquierdo: pega desde el portapapeles
    btn.OnEvent("ContextMenu", (*) => edit.Value := "")       ; Clic derecho: limpia el Edit
}

SoftwareVersionSelect(Name, yOffset, btnSze, BtnHeigh) {
    btnSoftware := SkrvGui.Add("Button", Format("x50 y{} w{} h{}", yOffset, btnSze, BtnHeigh), Name)
    editSoftware := SkrvGui.Add("Edit", Format("x260 y{} w750 h{}", yOffset, BtnHeigh), "")
    
    btnSoftware.OnEvent("Click", (*) => SoftwareVersionGUI(editSoftware))
}

SoftwareVersionGUI(editSoftware) {
    SoftwGui := Gui()
    SoftwGui.Opt("+AlwaysOnTop")
    SoftwGui.SetFont("s10")

   
    SoftwGui.Add("Text", "x10 y10", "Select Software Versions:")
    
    SoftwGui.Add("Text", "x10 y30", "Module:")
    comboBox1 := SoftwGui.Add("ComboBox", "x10 y50 w150", ["Unite", "TRIOS Module", "Dental Desktop", "Model Builder", "Automate"])
    
    SoftwGui.Add("Text", "x180 y30", "Version:")
    comboBox2 := SoftwGui.Add("ComboBox", "x180 y50 w150", [ "1.7.83.0", "1.7.8.1","1.8.8.0","1.8.5.1"])
    
    ; *** Aquí establecemos los valores por defecto correctamente ***
    comboBox1.Text := "Dental Desktop"
    comboBox2.Text := "1.7.83.0"

    btnSelect := SoftwGui.Add("Button", "x10 y120 w320", "Select")
    
    btnSelect.OnEvent("Click", (*) => (
        editSoftware.Value := comboBox1.Text " version " comboBox2.Text, 
        SoftwGui.Destroy()
    ))
    
    SoftwGui.Show()
}



OutputLine(Name,xOffset, yOffset, Insize ,btnSze,BtnHeigh){
    Outbtn := SkrvGui.Add("Button", Format("x{} y{} w{} h{}",xOffset, yOffset,btnSze,BtnHeigh), Name)
    outedit := SkrvGui.Add("Edit", Format("x{} y{} w{} h{}",xOffset, yOffset+BtnHeigh+8,Insize,BtnHeigh*2), "")
    Outbtn.OnEvent("ContextMenu", (*) => outedit.Value := "")       ; Clic derecho: limpia el Edit
}


Hotkeys(Name, yOffset){
    btnSze:= 200
    btn := SkrvGui.Add("Button", Format("x50 y{} w{} h{}", yOffset,btnSze,BtnHeigh), Name)
    ; edit := SkrvGui.Add("Edit", Format("x260 y{} w700 h{}", yOffset), "")
    btn.OnEvent("Click", (*) => Ctrl_y())

}
; PESTAÑA 1 (General)
tab.UseTab(1)
y := 50 ; Posición inicial en Y
x:= 260
spacing := 35 ; Espaciado entre elementos
Insize:=390

SkrvGui.Add("Text",)
InputLine("Name", y,btnSze,BtnHeigh), y += spacing
InputLine("Phone", y,btnSze,BtnHeigh), y += spacing
InputLine("Email", y,btnSze,BtnHeigh), y += spacing
InputLine("Company Name", y,btnSze,BtnHeigh), y += spacing
SoftwareVersionSelect("Software Version", y,btnSze,BtnHeigh), y += spacing
InputLine("Dongle", y,btnSze,BtnHeigh), y += spacing
InputLine("SID", y,btnSze,BtnHeigh), y += spacing
InputLine("GUI", y,btnSze,BtnHeigh), y += spacing
InputLine("Scanner S/N", y,btnSze,BtnHeigh), y += spacing
InputLine("PC ID", y,btnSze,BtnHeigh), y += spacing
InputLine("Case Number", y,btnSze,BtnHeigh), y += spacing

descripton:= SkrvGui.Add("Button", Format("x50 y{} w{} h{}", y,btnSze,BtnHeigh), "Description"),

descripton.OnEvent("Click", (*) => DescriptionGUI())

DescriptionGUI() {
    DescripGui := Gui()
    DescripGui.Opt("+AlwaysOnTop")
    DescripGui.SetFont("s10")

   
    DescripGui.Add("Text", "x10 y10", "The desccription is:")
    ; FinalDesc:= outedit.value, "on" , editSoftware.value 

          
    

    ; btnCopy := DescripGui.Add("Button", "x10 y120 w320", "Copy")
    
    ; btnCopy.OnEvent("Click", (*) => (
    ;     editDescripare.Value := comboBox1.Text " version " comboBox2.Text, 
    ;     DescripGui.Destroy()
    ; ))
    
    DescripGui.Show()
}

OutputLine("Isuue",x,y, Insize ,btnSze,BtnHeigh), 
OutputLine("Solution",(x+Insize +10),y, Insize ,btnSze,BtnHeigh), y += spacing


Hotkeys("Hotkeys", y) 

BtnSaveInfo := SkrvGui.Add("Button", "w150 h30", "Save Info")
BtnSaveInfo.OnEvent("Click", SaveBttm)
; Función para abrir una nueva ventana
SaveBttm(*) {
    newGui := Gui(,"Nueva Ventana")
    newGui.Add("Text","w300 h300", "Informaton Saved")
    newGui.Show()
}

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
!i::tab.Value := 1  ; Alt + G -> General
!v::tab.Value := 2  ; Alt + V -> View
!x::tab.Value := 3  ; Alt + S -> Settings





 



