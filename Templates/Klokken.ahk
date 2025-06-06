#Requires AutoHotkey v2.0
global rutaScriptFive9 := A_ScriptDir . "\Five9.py" ; Cambia esto a la ruta de tu script Python
; global rutaPython := "C:\Python313\python.exe"
; ; global rutaPython := "C:\Users\Joel Hurtado\AppData\Local\Programs\Python\Python313\python.exe"
; ; ; Ruta para python ------------------------------------------------------------
; ; ; Ejecutar "where python" y guardar la salida
; ; RunWait A_ComSpec ' /c where python > python_path.txt', , 'Hide'

; ; ; Leer archivo generado por `where python`
; ; raw := FileRead("python_path.txt")

; ; ; Separar por líneas
; ; lines := StrSplit(raw, "`n")

; ; ; Tomar la primera línea y quitar posibles espacios o saltos de carro (\r)
; ; global rutaPython := Trim(lines[1], "`r`n ")


; ; ; Ruta para python ------------------------------------------------------------
; https://login.five9.com/index.htm
; ---------------------------------------------------------------------------------------------------------------------------------
global KlokkenGui  ; Asegúrate de que este GUI esté definido globalmente (probablemente en Klokken.ahk)

KlokkenGui := Gui(, "KLOKKEN")
KlokkenGui.Opt("+Resize +MinSize250x170 +MaximizeBox +MinimizeBox") ; Hace la ventana redimensionable
KlokkenGui.SetFont("s20 bold", "Segoe UI")  ; Tamaño 14, negrita, fuente bonita
KlokkenGui.SetFont("q5")
; Establecer el ícono de la ventana
BtnHeigh:=40
btnSze:= 130
; ; KlKnWidth := 320
; ; KlKnHeigh := 170
; ; KlokkenGui.Show("w" KlKnWidth " h" KlKnHeigh)


breakTime(){
    KlokkenGui.Minimize() ; Minimiza la ventana
    global rutaPython
    args := "--breaktime "

    Five9_Exe := A_WorkingDir "\Five9.exe"
    if FileExist(Five9_Exe) {
        comand := Format('"{}" {}', Five9_Exe, args)
        A_Clipboard := comand
        RunWait(comand, , "Hide")

    } else {
        rutaScriptFive9 := A_WorkingDir "\Five9.py"
        comand := Format('"{}" "{}" {}', rutaPython, rutaScriptFive9, args) 
        A_Clipboard := comand
        RunWait(comand, , "Hide")
    }

}
lunchTime(){
    KlokkenGui.Minimize() ; Minimiza la ventana

    global rutaPython

    args := "--lunchtime "
    Five9_Exe := A_WorkingDir "\Five9.exe"
    if FileExist(Five9_Exe) {
        comand := Format('"{}" {}', Five9_Exe, args)
        A_Clipboard := comand
        RunWait(comand, , "Hide")

    } else {
        rutaScriptFive9 := A_WorkingDir "\Five9.py"
        comand := Format('"{}" "{}" {}', rutaPython, rutaScriptFive9, args) 
        A_Clipboard := comand
        RunWait(comand, , "Hide")
    }
}


InputLineKloken(Name, x?, xedit?, widthedit?, yOffset?, btnSze?, BtnHeigh?, buttomDefltPaste?, editYplace?, editHeigh?) {
    global rutaScriptFive9

    Name2Text := Name

    ; Valores predeterminados
    x := IsSet(x) ? x : 50
    xedit := IsSet(xedit) ? xedit : 300
    widthedit := IsSet(widthedit) ? widthedit : 770
    yOffset := IsSet(yOffset) ? yOffset : 100
    btnSze := IsSet(btnSze) ? btnSze : 80
    BtnHeigh := IsSet(BtnHeigh) ? BtnHeigh : 30
    editYplace := IsSet(editYplace) ? editYplace : yOffset
    editHeigh := IsSet(editHeigh) ? editHeigh : BtnHeigh

    ; Botón con evento independiente
    btn := KlokkenGui.Add("Button", Format("x{} y{} w{} h{}", x, yOffset, btnSze, BtnHeigh), Name2Text)
    KlokkenGui.SetFont("s20 bold", "Segoe UI")  ; Tamaño 20, negrita, fuente bonita

    ; Asigna la función según el nombre del botón
    if (Name == "&Break") {
        btn.OnEvent("Click", (*) => breakTime())  ; Función para el botón "Break"
    } else if (Name == "&Lunch") {
        btn.OnEvent("Click", (*) => lunchTime())  ; Función para el botón "Lunch"
    }
}


; ---------------------------------------------------------------------------------------------------------------------------------

; PESTAÑA 1 (General)
; imagePath := A_ScriptDir . "/LogoSmall.png"
; scale := 5.5
; wd:= 843
; hi := 559

imagePath := A_ScriptDir . "/IMG_Klokken_Trns.png"
wd:= 1024
hi := 1024
scale := 5.5
imgWidth1 := wd / scale
imgHeight1 := hi / scale
imgX1 := 150
imgY1 := -10
imgControl := KlokkenGui.Add("Picture", Format("x{} y{} w{} h{}", imgX1, imgY1, imgWidth1, imgHeight1), imagePath)

y := 10 ; Posición inicial en Y
x:= 10
InputLineKloken("&Break",10,20+btnSze,200, y,btnSze,BtnHeigh, true,,), y += 50 
InputLineKloken("&Lunch",10,20+btnSze,200, y,btnSze,BtnHeigh, false,,), y 
BtnKlokken := KlokkenGui.Add("Button", "x10 y110 w160 h40", "KLOKKEN!")
BtnKlokken.OnEvent("Click", (*) => AbrirFive9())
BtnKlokken.OnEvent("ContextMenu", (*) => CerrarFive9())

AbrirFive9() {
    KlokkenGui.Minimize() ; Minimiza la ventana

    global rutaPython

    args := "--klokken "
    Five9_Exe := A_WorkingDir "\Five9.exe"
    if FileExist(Five9_Exe) {
        comand := Format('"{}" {}', Five9_Exe, args)
        A_Clipboard := comand
        RunWait(comand, , "Hide")

    } else {
        rutaScriptFive9 := A_WorkingDir "\Five9.py"
        comand := Format('"{}" "{}" {}', rutaPython, rutaScriptFive9, args) 
        A_Clipboard := comand
        RunWait(comand, , "Hide")
    }
   
}
CerrarFive9() {
    KlokkenGui.Minimize() ; Minimiza la ventana
    

    global rutaPython

    args := "--klokkenout "
    Five9_Exe := A_WorkingDir "\Five9.exe"
    if FileExist(Five9_Exe) {
        comand := Format('"{}" {}', Five9_Exe, args)
        A_Clipboard := comand
        RunWait(comand, , "Hide")

    } else {
        rutaScriptFive9 := A_WorkingDir "\Five9.py"
        comand := Format('"{}" "{}" {}', rutaPython, rutaScriptFive9, args) 
        A_Clipboard := comand
        RunWait(comand, , "Hide")
    }
   
}