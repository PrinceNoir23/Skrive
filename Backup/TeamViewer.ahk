#Requires AutoHotkey v2.0
Persistent 1  ; Establece el script como persistente

; Tu código aquí
; Mostrar el tooltip
; ---------------------------------------------------------------------------------------------------------------------------------
; Mostrar GUI al inicio y cerrarla después de 3 segundos
startupGui := Gui()
startupGui.Title := "Windows Keys"
startupGui.Opt("AlwaysOnTop -SysMenu")
startupGui.SetFont("s15")
startupGui.Add("Text", , "Windows Keys`nStarted`n")
startupGui.SetFont("s10 Bold")
startupGui.Add("Text", , "COMBOS ")
startupGui.SetFont("s10 ")
startupGui.Add("Text", , "Ctrl + Shift + Y")
startupGui.Show("AutoSize x0 y0")
SetTimer(() => startupGui.Destroy(), 3500) ; Ocultar el tooltip después de 3 segundos
; Variables globales
myGui := 0
guiVisible := false

; ---------------------------------------------------------------------------------------------------------------------------------
; Definir las combinaciones de teclas y sus funciones

^+z::{
    MsgBox("Key Stopper")
    Pause
    Suspend
}

^+e::{
    MsgBox("Key Killer")
    
    ExitApp
    
}

^+r::{
    ; MsgBox("Key Reloder")
    Reload
}

ForwardToTeamViewer(KeyToSend) { 
    Critical  ; Evita interrupciones

    Sleep(500)
    Send (KeyToSend)
    Return True
}


; ForwardToTeamViewer(KeyToSend) {
;     exename := WinGetProcessName("A")
;     if (exename = "TeamViewer.exe") {
;         Critical  ; Evita interrupciones
;         Sleep(1000)
;         SendInput(KeyToSend)
;         Return True
;     }
;     MsgBox A_ComputerName
;     Return false
;     ; 
; }


; ---------------------------------------------------------------------------------------------------------------------------------

Search_Focus(SearchObject){
    Sleep(800)
    Send("{Ctrl Down}{l}{Ctrl Up}")  ; Ctrl + L
    Sleep(800)
    Send(SearchObject)  ; Ctrl + L
    Sleep(500)
    Send '{Enter}'
    Sleep(100)
}
OpenTab(URL) {
    Sleep(1000)
    Send("{Ctrl Down}{t}{Ctrl Up}")  ; Ctrl + L
    Search_Focus(URL)
    Sleep(500)
    return
}
WndOpen(App_to_Open){
    Sleep(200)
    if ForwardToTeamViewer("{LWin Down}{LWin Up}"){
    Sleep(1000)
    Send(App_to_Open)
    Sleep(500)
    Send '{Enter}'
    Sleep(500)
    return
    }
}
WndRun(App_to_Open){
    Sleep(500)
    if ForwardToTeamViewer("{LWin Down}r{LWin Up}") {
    Sleep(1000)  ; Espera dos segundos antes de enviar el texto
    Send(App_to_Open)
    Sleep(500)
    Send '{Enter}'
    Sleep(200)
    return
    }
}
ExplorerOpen(Folder_To_Open){
    Sleep(100)
    if ForwardToTeamViewer("{LWin Down}e{LWin Up}") {
        Sleep(2000)
        Search_Focus(Folder_To_Open)
        Sleep(300)
    }    
    return

}

Ctrl_a() {
    IB := InputBox("Input the case number `nEXAMPLE :`n CAS-XXXXXX-XXXXXX  ", " Powershell Backup ", "w289 h180")
    if (IB.Result = "Cancel") {
        MsgBox("Ingresaste '" IB.Value "' pero luego cancelaste.")
    } else {
        MsgBox("Ingresaste '" IB.Value "'.")
        if ForwardToTeamViewer("{LWin Down}{x}{LWin Up}") {
        Sleep(500)  ; Espera 500 milisegundos para que el Explorador se abra completamente
        Send("{a}")
        Sleep(2000)  ; Espera 1 se
        Send("{Left}")
        Sleep(500)
        Send '{Enter}'
        Sleep(1000)
        ; Concatenar IB.Value al comando SendPlay
        PWS := (" cd ..\..`n $programData = [System.Environment]::GetFolderPath('CommonApplicationData') `n$public = $env:PUBLIC `n$source = Join-Path -Path $programData -ChildPath '3Shape\DentalDesktop\Logs' `n$destination = Join-Path -Path $public -ChildPath 'CAS_" IB.Value "'" " `n$destination1 = Join-Path -Path $public -ChildPath 'CAS_" IB.Value ".zip' `nCopy-Item -Path $source -Destination $destination -Recurse -Force `nCompress-Archive -Path $destination  -DestinationPath $destination1 `nRemove-Item -Path $destination -Force `nStart-Process explorer.exe $public")
        Sleep(500)
        A_Clipboard := PWS

        ; Send("{Raw}" PWS)
        ; Send( PWS)
        ; Sleep(500)
        ; ClipWait  ; Espera a que el contenido esté disponible en el portapapeles
        Send("{Ctrl Down}{v}{Ctrl Up}")  ; Ctrl + L
        }

        Return  ; Sale si el envío fue exitoso
    }
}


; Ctrl_a() {
;     if ForwardToTeamViewer("{LWin Down}{x}{LWin Up}") {
;         Sleep(500)  ; Espera 500 milisegundos para que el Explorador se abra completamente
;         Send("{a}")
;         Sleep(1000)  ; Espera 1 segundos
;         Send("{Left}")
;         Sleep(500)
;         Send '{Enter}'
;         Sleep(1000)
;         Send("{Ctrl Down}{v}{Ctrl Up}")  ; Envía Ctrl+V para pegar el contenido seleccionado 
;         return  ; Salir si el envío fue exitoso
;     }
    
;       ; Si no es TeamViewer, muestra el nombre del PC
; }


; Ctrl_a() {
;     if ForwardToTeamViewer("{LWin Down}{x}{LWin Up}") {
;         Sleep(500)  ; Espera 500 milisegundos para que el Explorador se abra completamente
;         Send("{a}")
;         Sleep(10000)
;         Send("{Ctrl Down}{v}{Ctrl Up}")  ; Envía Ctrl+V para pegar el contenido seleccionado 
;         return  ; Salir si el envío fue exitoso
;     }
    
;       ; Si no es TeamViewer, muestra el nombre del PC
; }


Ctrl_b() {
    if ForwardToTeamViewer(WndOpen("Edge")) {
        
        OpenTab("https://www.speedtest.net/")
        OpenTab("http://localhost:27027/")
        OpenTab("https://www.nvidia.com/en-us/drivers/")
        OpenTab("http://localhost:8000/3ShapeWirelessService//")
        OpenTab("updates.3shape.com")
        OpenTab("3shapeconfig.com")
        ; response := MsgBox(4, "Continue?", "¿You want to open Wireless Drivers?")
        ; If (response = 6)  ; 6 es el código de respuesta cuando se presiona "Sí"
        ; {

        ;     ; Aquí va el código para continuar con el proceso
        ;     MsgBox("El proceso continuará.")
        ; }
        ; Else
        ; {
        ;     ; Aquí puedes poner el código para detener el proceso o realizar otra acción
        ;     MsgBox("El proceso se detendrá.")
        ;     ExitApp()  ; Detiene el script si el usuario elige "No"
        ; }



        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_s() {
    if ForwardToTeamViewer("") {
        
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_h() {
    if ForwardToTeamViewer(ExplorerOpen( "%programfiles%\3Shape\Dental Desktop\Plugins\ThreeShape.TRIOS\ScanSuiteTrios")) {
        ExplorerOpen("%programdata%\3Shape\ScanSuite")
        return  ; Salir si el envío fue exitoso
    }
}

Ctrl_w() {
    if ForwardToTeamViewer(ExplorerOpen( "%programfiles%\3Shape\Wireless Service\drivers")) {
        return  ; Salir si el envío fue exitoso
    } 
}

Ctrl_q() {
     if ForwardToTeamViewer(ExplorerOpen("%programdata%\3Shape\DentalDesktop\ServerPackages") ) {
        
        return  ; Salir si el envío fue exitoso
    } 
}

; Ctrl_q() {
;      if ForwardToTeamViewer("{LWin Down}{e}{LWin Up}") {
;         ExplorerOpen("%programdata%\3Shape\DentalDesktop\ServerPackages")
;         return  ; Salir si el envío fue exitoso
;     } 
; }

Ctrl_d() {
    if ForwardToTeamViewer("{LWin Down}{x}{LWin Up}") {
        Sleep(500)  ; Espera 500 milisegundos para que el Explorador se abra completamente
        Send("{a}")
        Sleep(2000)  ; Espera 1 se
        Send("{Left}")
        Sleep(500)
        Send '{Enter}'
        Sleep(1000)   
        Sleep(500)
        PWS := "sc delete " . Chr(34) . "MigrationService" . Chr(34)
        A_Clipboard := PWS

        Send("{Ctrl Down}{v}{Ctrl Up}")  ; Ctrl + L
        Sleep(500)
        Send '{Enter}'
        Sleep(800)
        ExplorerOpen("%programfiles%\3Shape\Dental Desktop")
        Sleep(500)
        Send('{M}{i}{g}')
        
        response := MsgBox("¿Has eliminado la carpeta de migración?", "Continuar?", "YesNo")
        If (response = "Yes")
        {
            WndRun("regedit")
            Sleep(4000)
            Search_Focus("Computer\HKEY_LOCAL_MACHINE\SOFTWARE\3Shape\DentalDesktop\")
            Sleep(1000)
            Send('{M}{i}{g}')
            Sleep(1000)
            


            ; Aquí va el código para continuar con el proceso
            MsgBox("El proceso continuará.")
        }
        Else
        {
            ; Aquí puedes poner el código para detener el proceso o realizar otra acción
            MsgBox("El proceso se detendrá.")
            ExitApp()  ; Detiene el script si el usuario elige "No"
        }



    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_f() {
    if ForwardToTeamViewer("{LWin Down}{LWin Up}") {
        Sleep(2000)  ; Espera dos segundos antes de enviar el texto
        Send("cmd")
        Sleep(500)
        Send("{Right}")
        Sleep(500)
        Send("{Down}")
        Sleep(500)
        Send '{Enter}'
        ; Sleep(5000)
        Sleep(1500)
        Send("{Left}")
        Sleep(500)
        Send '{Enter}'
        Sleep(1500)

        PWS := ("@echo off`n"
                "cd ..`n"
                "del *.tmp /s /f /q`n"
                "del %tmp% /s /f /q`n"
                "DISM.exe /Online /Cleanup-image /Scanhealth`n"
                "DISM.exe /Online /Cleanup-image /Restorehealth`n"
                "sfc /scannow`n")
        Sleep(500)
        A_Clipboard := PWS
        Sleep(500)
        Send("{Ctrl Down}{v}{Ctrl Up}")  ; Ctrl + L
        Sleep(100)
        Send '{Enter}'
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

AbrirFirewall_1(port,Name) {
    PRT := port 
    
    Sleep(2000)
    
    if (Name = "3Shape Outbound Ports" or Name = "3Shape TRIOS 3 & 4 TCP_O" or Name = "3Shape TRIOS 5 UDP_O" or Name = "3Shape TRIOS 5_O" or Name="3Shape coDiagnostix Port_o") {
    Send '{o}'
    } else {
        Send '{i}'
    }



    Sleep(2000)
    Send '{AppsKey}'
    Sleep(2000)
    Send '{Down}'
    Sleep(2000)
    Send '{Enter}'
    Sleep(2000)
    
    Loop 4 {
        Send '{Tab}'
        Sleep(500)
    }
    
    Send '{Down}'
    Sleep(2000)
    Send '{Enter}'
    Sleep(2000)
    Send(PRT)
    Sleep(2000)
    Send '{Enter}'
    if (Name = "3Shape Outbound Ports" or Name = "3Shape TRIOS 3 & 4 TCP_O" or Name = "3Shape TRIOS 5 UDP_O" or Name = "3Shape TRIOS 5_O" or Name="3Shape coDiagnostix Port_o"){
        Loop 5 {
        Send '{Tab}'
        Sleep(500)
        }
        Sleep(2000)
        Send '{Down}'
    }
    Sleep(2000)
    Send '{Enter}'
    Sleep(2000)
    Send '{Enter}'
    Sleep(2000)
    Send(Name)
    Sleep(2000)
    Send '{Enter}'
    Sleep(2000)
    return
}
AbrirFirewall_2(port,Name,IP,Protocol) {

    PRT := port 
    
    Sleep(2000)
    if (Name = "3Shape Outbound Ports" or Name = "3Shape TRIOS 3 & 4 TCP_O" or Name = "3Shape TRIOS 5 UDP_O" or Name = "3Shape TRIOS 5_O") {
    Send '{o}'
    } else {
        Send '{i}'
    }

    Sleep(2000)
    Send '{AppsKey}'
    Sleep(1000)
    Send '{Down}'
    Sleep(1000)
    Send '{Enter}'
    Sleep(1000)
    
    Loop 4 {
        Send '{Tab}'
        Sleep(500)
    }
    
    Send '{Up}'
    Sleep(1000)
    Send '{Enter}'
    Sleep(1000)
    Send '{Enter}'

    Loop 5 {
        Send '{Tab}'
        Sleep(500)
    }

    Send(Protocol)
    Sleep(1000)
    loop 2 {
        Send '{Tab}'
        Sleep(500)
        Send '{Down}'
        Sleep(1000)
        Send(PRT)
    }
    
    Sleep(1000)
    Send '{Enter}'

    Loop 5 {
        Send '{Tab}'
        Sleep(50)
    }
    
    loop 2 {
        Send '{Down}'
        Sleep(1000)
        loop 2 {
            Send '{Tab}'
            Sleep(500)
        }
        Send '{Enter}'
        Sleep(1000)
        Send(IP) 
        Send '{Enter}'
        loop 2 {
            Send '{Tab}'
            Sleep(500)
        }
    }
    Send '{Enter}'
    Sleep(1000) 
    if (Name = "3Shape Outbound Ports" or Name = "3Shape TRIOS 3 & 4 TCP_O" or Name = "3Shape TRIOS 5 UDP_O" or Name = "3Shape TRIOS 5_O"){
        Loop 5 {
        Send '{Tab}'
        Sleep(500)
        }
        Sleep(1000)
        Send '{Down}'
    }
    Sleep(1000)
    Send '{Enter}'
    Sleep(1000)
    Send '{Enter}'
    Sleep(1000)
    Send(Name)
    Sleep(1000)
    Send '{Enter}'
    Sleep(1000)
    return

}

Ctrl_i() {
    if (ForwardToTeamViewer("{LWin Down}{LWin Up}")) {
        Sleep(2000)  ; Espera antes de enviar el texto

        ; MsgBox

        ; Send("Windows Defender Firewall with adv")
        ; Sleep(2000)
        ; Send("{Enter}")


        ; ShapePorts := "5353,20100 "
        ; AbrirFirewall_1(ShapePorts,"3Shape coDiagnostix Ports")
        ; Sleep(1000)
        ; AbrirFirewall_1(ShapePorts,"3Shape coDiagnostix Port_o")


        Send("Windows Defender Firewall with adv")
        Sleep(2000)
        Send("{Enter}")


        ShapePorts := "443, 8530, 80, 21, 5480, 5481, 5482, 5483, 27027, 1433, 5484, 4380, 2277, 2278, 1900, 2048, 2070, 8000, 1434, 1066, 1666, 51398, 27028, 27029, 27030, 27031, 27032, 27033"
        AbrirFirewall_1(ShapePorts,"3Shape Inbound Ports")
        Sleep(1000)
        AbrirFirewall_1(ShapePorts,"3Shape Outbound Ports")


        Sleep (2000)

        ; Send '{Alt}{Tab}' 
        
        Shape_IP := "10.33.3.1"


        ; Send("{LWin Down}{LWin Up}")  
        ; Sleep(2000)  ; Espera antes de enviar el texto
        ; Send("Windows Defender Firewall with adv")
        ; Sleep(2000)
        ; Send("{Enter}")

        TRIOS_3_4_Ports1:= "21, 80, 23796"
                
        AbrirFirewall_2(TRIOS_3_4_Ports1,"3Shape TRIOS 3 & 4 TCP",Shape_IP, '{T}')
        Sleep(1000)
        AbrirFirewall_2(TRIOS_3_4_Ports1,"3Shape TRIOS 3 & 4 TCP_O",Shape_IP, '{T}')


        TRIOS_5_Ports:= "58220-58230"

        Sleep (2000)
        
        ; Send("{LWin Down}{LWin Up}")
        ; Sleep(2000)  ; Espera antes de enviar el texto
        ; Send("Windows Defender Firewall with adv")
        ; Sleep(2000)
        ; Send("{Enter}")


        AbrirFirewall_2(TRIOS_5_Ports,"3Shape TRIOS 5 UDP",Shape_IP, '{U}') 
        Sleep(1000)
        AbrirFirewall_2(TRIOS_5_Ports,"3Shape TRIOS 5 UDP_O",Shape_IP, '{U}')


        ; Send("{LWin Down}{LWin Up}")
        ; Sleep(2000)  ; Espera antes de enviar el texto
        ; Send("Windows Defender Firewall with adv")
        ; Sleep(2000)
        ; Send("{Enter}")

        AbrirFirewall_1("23796","3Shape TRIOS 5")
        Sleep(1000)
        AbrirFirewall_1("23796","3Shape TRIOS 5_O")

        

        ; Mensaje de confirmación con botones Sí/No
        ; loop 3 {
        ;     Send '{Alt Down}{F4}{Alt Up}'  ; Cierra la ventana actual  ; Cierra la ventana actual
        ;     Sleep (2000)
        ; }
        
        MsgBox("Firewalls Setted Correctly", "Confirmación", "Ok")
        return  ; Salir si el envío fue exitoso
    }
    
    MsgBox(A_ComputerName)  ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_o() {
    if ForwardToTeamViewer("{LWin Down}{LWin Up}") {
        Sleep(2000)  ; Espera dos segundos antes de enviar el texto
        Send("cmd")
        Sleep(500)
        Send("{Right}")
        Sleep(500)
        Send("{Down}")
        Sleep(500)
        Send '{Enter}'
        ; Sleep(5000)
        Sleep(1500)
        Send("{Left}")
        Sleep(500)
        Send '{Enter}'
        Sleep(1000)
        Send("wmic bios get serialnumber")
        Send '{Enter}'
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_x() {
    if ForwardToTeamViewer(Search_Focus("%programdata%\3Shape\DentalDesktop\ThreeShape.TRIOS\")) {
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_n() {
    
    IB := InputBox("Input the PATH `nEXAMPLE :`n C:\Program Files (x86)\Google  ", " Explorer Path ", "w289 h180")
    if (IB.Result = "Cancel") {
        MsgBox("Ingresaste ( " IB.Value " ) pero luego cancelaste.")
    } else {
        MsgBox("Ingresaste ( " IB.Value " ) ")
        (ExplorerOpen(IB.value)) 
        Sleep(300)
        Send '{Enter}'
        

        Return  ; Sale si el envío fue exitoso
    }
}

Ctrl_g() {
    if ForwardToTeamViewer(Sleep(500)) {
        Send("joel.hurtado@3shape.com")
        Sleep(9000)  ; Espera dos segundos antes de enviar el texto
        Send("LaScarlata2024*")
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}
; Ctrl_g() {
;     if ForwardToTeamViewer(Sleep(500)) {
;           ; Espera dos segundos antes de enviar el texto
;         SendEvent "{Click 1758 959 Right Down}{Click 1758 959 Right Up}"
;         Sleep(1000)  ; Espera dos segundos antes de enviar el texto
;         SendEvent "{Click 956 590 Left Down}{Click 956 590 Left Up}"
;         Sleep(3000)  ; Espera dos segundos antes de enviar el texto
;         Send("joel.hurtado@3shape.com")
;         Sleep(1000)  ; Espera dos segundos antes de enviar el texto
;         SendEvent "{Click 917 509 Left Down}{Click 917 509 Left Up}"
;         Sleep(1000)  ; Espera dos segundos antes de enviar el texto
;         SendEvent "{Click 992 616 Left Down}{Click 992 616 Left Up}"
;         Sleep(1000)  ; Espera dos segundos antes de enviar el texto
;         SendEvent "{Click 915 488 Left Down}{Click 915 488 Left Up}"

;         Sleep(1000)  ; Espera dos segundos antes de enviar el texto
;         Send("LaScarlata2024*")
;         Sleep(1000)  ; Espera dos segundos antes de enviar el texto
;         SendEvent "{Click 958 630 Left Down}{Click 958 630 Left Up}"
;         Sleep(1000)  ; Espera dos segundos antes de enviar el texto
;         SendEvent "{Click 719 418 Left Down}{Click 719 418 Left Up}"
;         Sleep(1000)  ; Espera dos segundos antes de envia   r el texto
;         SendEvent "{Click 739 599 Left Down}{Click 739 599 Left Up}"
;         Sleep(1000)  ; Espera dos segundos antes de envia   r el texto
;         SendEvent "{Click 972 660 Left Down}{Click 972 660 Left Up}"
;         Sleep(2000)  ; Espera dos segundos antes de envia   r el texto
;         SendEvent "{Click 1188 615 Left Down}{Click 1188 615 Left Up}"
;         Sleep(1000)  ; Espera dos segundos antes de envia   r el texto
;         return  ; Salir si el envío fue exitoso
;     }
    
;       ; Si no es TeamViewer, muestra el nombre del PC
; }


ClearClipboard() {
    Sleep(500)
    Send("{LWin Down}{LWin Up}") 
    Sleep(2000)  ; Espera dos segundos antes de enviar el texto
    Send("clear clipboard data")
    Sleep(500)
    Send '{Enter}'
    Sleep(2000)
    Send '{Enter}'

    return  ; Salir si el envío fue exitoso
    }

; ---------------------------------------------------------------------------------------------------------------------------------
AddColoredText(gui, xPos, yPos, colWidth, rowHeight, text, index) {
    colors := ["FFBA08", "3F88C5", "032B43", "700548", "B2EF9B"]  ; Lista de colores
    colorIndex := Mod(index - 1, colors.Length) + 1  ; Ciclo a través de los colores

    control := gui.Add("Text", Format("x{} y{} w{} h{}", xPos, yPos, colWidth, rowHeight), text)
    control.SetFont(Format("c{}", colors[colorIndex]))  ; Asigna el color correspondiente
}

; KEY COMBS TOOLTIP 
Ctrl_y() {
    global myGui, guiVisible, aplhVisible, AlphGui
    if !myGui {
        myGui := Gui()
        myGui.Opt("+Resize +MinSize +MaximizeBox +MinimizeBox")
        myGui.BackColor := "00100B"
        myGui.Title := "COMBOS"
        myGui.Opt("AlwaysOnTop")
        myGui.SetFont("c00ffe1")
        myGui.SetFont("S14")
        myGui.SetFont(, "Arial")

        ; Medidas para la cuadrícula de botones
        btnWidth := 110
        btnHeight := 30
        btnSpacing := 10
        startX := 20
        startY := 10
        cols := 4  ; Cantidad de columnas en la cuadrícula

        ; Lista de botones con formato: [Texto, Función]
        buttons := [
            ["Transparent", ToggleTransparency],
            ["Title Hide", TitleHide],
            ["OACI", Alphabet],
            ; ["Extra 1", ExtraFunction1], ; Puedes agregar más funciones
            ; ["Extra 2", ExtraFunction2],
            ; ["Extra 3", ExtraFunction3]
        ]

        ; Función para agregar los botones en una cuadrícula
        AddButtonGrid(buttons, startX, startY, btnWidth, btnHeight, btnSpacing, cols)

        ; Variables de control
        global transparencyEnabled := false
        global titleHidden := false
        global aplhVisible := false
        global AlphGui := ""

        ; Agregar el texto debajo de los botones
        myGui.Add("Text", Format("x{} y+10", startX), "KEY COMBINATIONS:`n"
            . "  `n"
            . "All COMBOS = Ctrl + Shift + KEY`n"
            . "  `n"
            . "Pressing COMBO + Y again closes this window `n"
            . "  `n"
            . "-------------------------------`n"
            . "  `n"
            . "E - Exit Program`n"
            . "R - Reload Program`n"
            . "Z - Stop Process`n"
            . "  `n"
            . "-------------------------------`n"
            . "A - Logs - Backup code`n"
            . "B - LclHst 27027 / 8000 / Nvidia Drivers`n"
            . "D - Migration Failed Solution`n"
            . "F - *`n"
            . "G - Support Mode`n"
            . "H - Hardware Test* || Firmware / Scan Suit`n"
            . "I - Firewall Advanced`n"
            . "K - Clear Clipboard`n"
            . "N - Specified PATH `n"
            . "O - CMD - Service Tag`n"
            . "Q - Server Packages`n"
            . "S - `n"
            . "W - Wireless Drivers*`n"
            . "X - Streams`n"
            . "Y - TECLAS INFORMATION`n"
            . "-------------------------------`n"
            . "1 - Servers`n"
            . "2 - Logs folder`n"
            . "3 - SQL*`n"
            . "4 - Keys*`n"
            . "5 - Dxdiag / Appwiz / Sys Info / Device manager / Services`n"
            . "6 - `n"
            . "7 - `n"
            . "8 - `n"
            . "9 - `n"
            . "0 - ")
    }

    ; Alternar visibilidad de la ventana principal
    if guiVisible {
        myGui.Hide()
        guiVisible := false
    } else {
        myGui.Show("AutoSize Center")
        guiVisible := true
    }
}

; Función para crear botones en una cuadrícula 3x3
AddButtonGrid(buttons, startX, startY, btnWidth, btnHeight, btnSpacing, cols) {
    global myGui
    loop buttons.Length {
        row := Floor((A_Index - 1) / cols)  ; Fila actual
        col := Mod(A_Index - 1, cols)  ; Columna actual (Usamos Mod para evitar problemas)

        xPos := startX + col * (btnWidth + btnSpacing)
        yPos := startY + row * (btnHeight + btnSpacing)

        btn := myGui.AddButton(Format("x{} y{} w{} h{}", xPos, yPos, btnWidth, btnHeight), buttons[A_Index][1])

        ; Verificar que existe una función asociada antes de asignarla
        if IsObject(buttons[A_Index][2]) {
            btn.OnEvent("Click", buttons[A_Index][2])  ; Asignar función solo si existe
        }
    }
}

; Función para cambiar la transparencia
ToggleTransparency(*) {
    global transparencyEnabled
    if (transparencyEnabled) {
        WinSetTransColor "Off"
        transparencyEnabled := false
    } else {
        WinSetTransColor("00100B", (myGui and AlphGui), 180)
        transparencyEnabled := true
    }
}


; Función para ocultar/mostrar el título de la ventana
TitleHide(*) {
    global titleHidden
    if (titleHidden) {
        myGui.Opt("+Caption")
        titleHidden := false
    } else {
        myGui.Opt("-Caption")
        titleHidden := true
    }
}



; Función para mostrar/ocultar el AlphGui (OACI)
Alphabet(*) {
    global AlphGui, aplhVisible, myGui

    if !IsObject(AlphGui) {  
        AlphGui := Gui()
        AlphGui.Opt("+Resize +AlwaysOnTop +MinSize +MaximizeBox +MinimizeBox")  ; Se mantiene sobre myGui
        AlphGui.BackColor := "00100B"
        AlphGui.SetFont("cb700ff")
        AlphGui.SetFont("S20")

        ; Definir el alfabeto en una matriz de pares (letra - palabra)
        alphabet := [
            ["A", "Alfa"], ["B", "Bravo"], ["C", "Charlie"], ["D", "Delta"],
            ["E", "Echo"], ["F", "Foxtrot"], ["G", "Golf"], ["H", "Hotel"],
            ["I", "India"], ["J", "Julieta"], ["K", "Kilo"], ["L", "Lima"],
            ["M", "Mike"], ["N", "November"], ["O", "Oscar"], ["P", "Papa"],
            ["Q", "Quebec"], ["R", "Romeo"], ["S", "Sierra"], ["T", "Tango"],
            ["U", "Uniform"], ["V", "Víctor"], ["W", "Whiskey"], ["X", "X-ray"],
            ["Y", "Yankee"], ["Z", "Zulu"]
        ]
        colSize := Floor(alphabet.Length / 2)  ; Punto de corte para la segunda columna


        rowHeight := 30
        colWidth := 180
        xStart := 15
        yStart := 10
        spacingX := 10  ; Espaciado entre columnas
        spacingY := 10   ; Espaciado entre filas


        ; Agregar texto en dos columnas
        Loop alphabet.Length {
            if (A_Index <= colSize) {  ; Primera mitad en la primera columna
                xPos := xStart
                yPos := yStart + (A_Index - 1) * (rowHeight + spacingY)
            } else {  ; Segunda mitad en la segunda columna
                xPos := xStart + colWidth + spacingX
                yPos := yStart + (A_Index - colSize - 1) * (rowHeight + spacingY)
            }

            AlphGui.Add("Text", Format("x{} y{} w{} h{}", xPos, yPos, colWidth, rowHeight), 
                Format("{} - {}", alphabet[A_Index][1], alphabet[A_Index][2]))
        }
    }

    ; Alternar visibilidad y mover AlphGui a la esquina superior derecha de myGui
    if aplhVisible {
        AlphGui.Hide()
        aplhVisible := false
    } else {
        ; Obtener la posición y tamaño de myGui
        myGuiX := 0, myGuiY := 10, myGuiW := 0, myGuiH := 0
        myGui.GetPos(&myGuiX, &myGuiY, &myGuiW, &myGuiH)  ; Obtener posición y tamaño

        ; Calcular la posición de AlphGui en la esquina superior derecha
        alphX := myGuiX + myGuiW - 620  ; 320 = aprox. el ancho de AlphGui
        alphY := myGuiY  ; Se alinea con la parte superior de myGui

        AlphGui.Show("AutoSize")  
        AlphGui.Move(alphX, alphY)  ; Mueve AlphGui a la esquina superior derecha de myGui
        aplhVisible := true
    }
}



; Funciones adicionales para nuevos botones
ExtraFunction1(*) {
    MsgBox("Función Extra 1 ejecutada")
}

ExtraFunction2(*) {
    MsgBox("Función Extra 2 ejecutada")
}

ExtraFunction3(*) {
    MsgBox("Función Extra 3 ejecutada")
}


; ---------------------------------------------------------------------------------------------------------------------------------





Ctrl_1() {
    if ForwardToTeamViewer(ExplorerOpen("%programdata%\3Shape\DentalDesktop\ThreeShape.TRIOS\Servers")) {
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_2() {
    if ForwardToTeamViewer("{LWin Down}e{LWin Up}") {
        Sleep(2000)
        Send("{Ctrl Down}{l}{Ctrl Up}")  ; Ctrl + L
        Sleep(500)  ; Espera medio segundo antes de enviar el texto
        Send("%programdata%\3Shape\DentalDesktop\Logs")
        Sleep(500)
        Send '{Enter}'
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_3() {
    if ForwardToTeamViewer("{LWin Down}e{LWin Up}") {
        Sleep(2000)
        Send("{Ctrl Down}{l}{Ctrl Up}")  ; Ctrl + L
        Sleep(500)  ; Espera medio segundo antes de enviar el texto
        SQL := "%programfiles%\Microsoft SQL Server\"
        A_Clipboard := SQL
        Send("{Ctrl Down}{v}{Ctrl Up}")  ; Ctrl + L
        Sleep(500)
        Send '{Enter}'

        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_4() {
    if ForwardToTeamViewer("{LWin Down}e{LWin Up}") {
        Sleep(2000)
        Send("{Ctrl Down}{l}{Ctrl Up}")  ; Ctrl + L
        Sleep(500)  ; Espera medio segundo antes de enviar el texto
        Send("%programfiles% (x86)\3Shape\Dongle Server Service\Keys")
        Sleep(500)
        Send '{Enter}'
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_5() {
    if ForwardToTeamViewer(WndRun("dxdiag")) {
        WndRun("appwiz.cpl")
        WndOpen("system info")
        WndOpen("Device manager")
        WndOpen("services")
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_6() {
    
}

Ctrl_7() {
    if ForwardToTeamViewer("") {
        
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_8() {
    if ForwardToTeamViewer("") {
        
        
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_9() {
    if ForwardToTeamViewer("") {
        
        
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_0() {
    if ForwardToTeamViewer("") {
        
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}
; ---------------------------------------------------------------------------------------------------------------------------------

^+a::Ctrl_a()  ; (Logs - Código de respaldo)
^+b::Ctrl_b() ;(LclHst 27027)
^+s::Ctrl_s() ;(Scan Suit)
^+h::Ctrl_h() ;(Hardware Test*)
^+w::Ctrl_w() ;(Wireless Drivers*)
^+q::Ctrl_q() ;(Server Packages)
^+d::Ctrl_d() ;(Device manager)
^+f::Ctrl_f() ;(Firmware*)
^+i::Ctrl_i() ;(Firewall Advanced)
^+o::Ctrl_o() ;(CMD - Service Tag)
^+x::Ctrl_x() ;(Streams)
^+y::Ctrl_y() ;(INFORMATION KEYS)
^+n::Ctrl_n() ;(Specifed Path)
^+g::Ctrl_g() ;(Support Mode)

^+k::ClearClipboard() ;(Clipboard Cleanup)

; ---------------------------------------------------------------------------------------------------------------------------------
; * un asteristo significa que lleva a una carpeta con %programfiles%
^+1::Ctrl_1() ;(servers)
^+2::Ctrl_2() ;(logs folder)
^+3::Ctrl_3() ;(SQL*)
^+4::Ctrl_4() ;(keys*)
^+5::Ctrl_5() ;(dxdiag)
^+6::Ctrl_6() ;(appwiz)
^+7::Ctrl_7() ;(Sys Info)
^+8::Ctrl_8() ;(LclHst 8000)
^+9::Ctrl_9() ;(Nvidia Drivers)
^+0::Ctrl_0() ;(Services)