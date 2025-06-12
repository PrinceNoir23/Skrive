#Requires AutoHotkey v2.0
#Include Json.ahk



Persistent

global eventos := Map()

loadInfo(){
    ; Ruta del JSON con horarios
    InfoPath := A_WorkingDir "\A_Info.json"

    ; Verifica existencia
    if !FileExist(InfoPath) {
        MsgBox("Archivo no encontrado: " InfoPath, "Error de Path", 16)
        ExitApp
    }

    ; Leer y parsear JSON con Jxon
    try {
        fileContent := FileRead(InfoPath, "UTF-8")
        horarios := Jxon_Load(&fileContent)
    } catch {
        MsgBox("Error al leer o parsear el archivo JSON.", "Error", 16)
        ExitApp
    }

    ; Claves esperadas
    claves := ["Hora In", "Hora Out", "Brk 1", "Brk 2", "Lunch "]

    ; Verifica que estén presentes y correctas
    for k in claves {
        if !horarios.Has(k) {
            MsgBox("Falta la clave '" k "' en el archivo JSON.", "Error", 16)
            ExitApp
        }
        hora := Trim(horarios[k])
        if !RegExMatch(hora, "^\d{1,2}:\d{2}$") {
            MsgBox("Hora inválida para '" k "': " hora, "Error de formato", 16)
            ExitApp
        }
        eventos[k] := hora
    }
}

; ===== FUNCIONES AUXILIARES =====

TimeToday(hhmm) {
    parts := StrSplit(hhmm, ":")
    return BuildDateTime(A_YYYY, A_MM, A_DD, parts[1], parts[2])
}

BuildDateTime(yyyy, mm, dd, hh := 0, min := 0, ss := 0) {
    return Format("{:04}{:02}{:02}{:02}{:02}{:02}", yyyy, mm, dd, hh, min, ss)
}
DateParse(dtStr) {
    return DateDiff(dtStr, "19700101000000", "Seconds")
}

DateDiffCustom(datetime1, datetime2, unidad := "Minutes") {
    if (unidad = "Seconds")
        return DateDiff(datetime1, datetime2, "Seconds")
    else if (unidad = "Minutes")
        return DateDiff(datetime1, datetime2, "Seconds") // 60
    else
        throw Error("Unidad inválida. Usa 'Seconds' o 'Minutes'.")
}




Access() {
    TrayTip("Access", "Ejecutando Login Access", 5)

    global rutaPython 
    global rutaScriptFive9 := A_ScriptDir . "\Five9.py" ; Cambia esto a la ruta de tu script Python
    args := "--klokken "


    comand := Format('"{}" "{}" {}', rutaPython, rutaScriptFive9, args) 
    A_Clipboard := comand
    ; Ejecutar el script con los argumentos
    RunWait(comand, , "Hide")
    ; RunWait(comand)
    ; Run("tu_programa.exe") ; si deseas ejecutar algo aquí
}
Logout() {
    TrayTip("Logout", "DESLOGUEATE TU TURNO ACABO", 5)

    global rutaPython 
    global rutaScriptFive9 := A_ScriptDir . "\Five9.py" ; Cambia esto a la ruta de tu script Python
    args := "--klokkenout "


    comand := Format('"{}" "{}" {}', rutaPython, rutaScriptFive9, args) 
    ; Ejecutar el script con los argumentos
    RunWait(comand, , "Hide")
    A_Clipboard := comand

    Sleep(60000)
    try {
        RunWait("*RunAs " A_Desktop "\ShutDown.bat")
    }
    catch {
        return ; Salir del script
    }



    ; RunWait(comand)
    ; Run("tu_programa.exe") ; si deseas ejecutar algo aquí
}

Klokken1() {
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
; Verificador de eventos
SetTimer(VerificarEventos, 60000)
SetTimer(loadInfo, 86400000)


VerificarEventos() {
    global eventos 
    now := A_Now
    weekday := FormatTime(now, "WDay")  ; 6 = Viernes

    for nombre, hora in eventos {
        evento := TimeToday(hora)

        diff := DateDiffCustom(evento, now, "Minutes")


        if (diff == 15) {
            TrayTip("Recordatorio", "Faltan 15 minutos para " . nombre . " (" . hora . ")", 5)
            Klokken1()
        }

        if (nombre == "Hora In" and diff == 1) {
            Access()
        }
        if (nombre == "Hora Out" and diff == 1) {
            Logout()
        }

        if (weekday == 6 and nombre == "Hora Out" and diff == 20) {
            TrayTip("Recordatorio", "Es viernes: edita tu horario para la próxima semana", 10)
            Klokken1()

        }
    }
}
; Llamar al inicio también
loadInfo()
VerificarEventos()
