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
    Sleep(500)
    SendInput(KeyToSend)
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
    if ForwardToTeamViewer("{LWin Down}{LWin Up}") {
        Sleep(2000)  ; Espera dos segundos antes de enviar el texto
        Send("Edge")
        Sleep(500)
        Send '{Enter}'
        Sleep(2000)
        Send("{Ctrl Down}{t}{Ctrl Up}")  ; Ctrl + L
        Sleep(1500)
        Send("{Ctrl Down}{l}{Ctrl Up}")  ; Ctrl + L
        Sleep(1500)
        Send("http://localhost:27027/")  ; Ctrl + L
        Send '{Enter}'
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_s() {
    if ForwardToTeamViewer("{LWin Down}e{LWin Up}") {
        Sleep(2000)
        Send("{Ctrl Down}{l}{Ctrl Up}")  ; Ctrl + L
        Sleep(500)  ; Espera medio segundo antes de enviar el texto
        Send("%programdata%\3Shape\ScanSuite")
        Sleep(500)
        Send '{Enter}'
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_h() {
    if ForwardToTeamViewer("{LWin Down}e{LWin Up}") {
        Sleep(2000)
        Send("{Ctrl Down}{l}{Ctrl Up}")  ; Ctrl + L
        Sleep(500)  ; Espera medio segundo antes de enviar el texto
        Send("%programfiles%\3Shape\Dental Desktop\Plugins\ThreeShape.TRIOS\ScanSuiteTrios")
        Sleep(500)
        Send '{Enter}'
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_w() {
    if ForwardToTeamViewer("{LWin Down}e{LWin Up}") {
        Sleep(2000)
        Send("{Ctrl Down}{l}{Ctrl Up}")  ; Ctrl + L
        Sleep(500)  ; Espera medio segundo antes de enviar el texto
        Send("%programfiles%\3Shape\Wireless Service\drivers")
        Sleep(500)
        Send '{Enter}'
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_q() {
    if ForwardToTeamViewer("{LWin Down}e{LWin Up}") {
        Sleep(2000)
        Send("{Ctrl Down}{l}{Ctrl Up}")  ; Ctrl + L
        Sleep(500)  ; Espera medio segundo antes de enviar el texto
        Send("%programdata%\3Shape\DentalDesktop\ServerPackages")
        Sleep(500)
        Send '{Enter}'
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_d() {
    if ForwardToTeamViewer("{LWin Down}{LWin Up}") {
        Sleep(2000)  ; Espera dos segundos antes de enviar el texto
        Send("Device manager")
        Sleep(500)
        Send '{Enter}'
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_f() {
    if ForwardToTeamViewer("{LWin Down}e{LWin Up}") {
        Sleep(2000)
        Send("{Ctrl Down}{l}{Ctrl Up}")  ; Ctrl + L
        Sleep(500)  ; Espera medio segundo antes de enviar el texto
        Send("%programfiles%\3Shape\Dental Desktop\Plugins\ThreeShape.TRIOS\ScanSuiteTrios\Firmware")
        Sleep(500)
        Send '{Enter}'
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_i() {
    if ForwardToTeamViewer("{LWin Down}{LWin Up}") {
        Sleep(2000)  ; Espera dos segundos antes de enviar el texto
        Send("Windows Defender Firewall with adv")
        Sleep(500)
        Send '{Enter}'
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
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
    if ForwardToTeamViewer("{Ctrl Down}{l}{Ctrl Up}") {
        Sleep(500)  ; Espera medio segundo antes de enviar el texto
        Send("%programdata%\3Shape\DentalDesktop\ThreeShape.TRIOS\")
        Sleep(500)
        Send '{Enter}'
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_n() {
    
    IB := InputBox("Input the PATH `nEXAMPLE :`n C:\Program Files (x86)\Google  ", " Explorer Path ", "w289 h180")
    if (IB.Result = "Cancel") {
        MsgBox("Ingresaste '" IB.Value "' pero luego cancelaste.")
    } else {
        MsgBox("Ingresaste '" IB.Value "'.")
        if ForwardToTeamViewer("{LWin Down}{e}{LWin Up}") {
        Sleep(2000)  ; Espera 500 milisegundos para que el Explorador se abra completamente
        Send("{Ctrl Down}{l}{Ctrl Up}")  ; Ctrl + L
        Sleep(100)  ; Espera 1 se
        ; Concatenar IB.Value al comando SendPlay
        PWS :=  IB.Value 
        Sleep(1000)
        A_Clipboard := PWS

        ; Send("{Raw}" PWS)
        ; Send( PWS)
        ; Sleep(500)
        ; ClipWait  ; Espera a que el contenido esté disponible en el portapapeles
        Send("{Ctrl Down}{v}{Ctrl Up}")  ; Ctrl + L
        Sleep(300)
        Send '{Enter}'

        }

        Return  ; Sale si el envío fue exitoso
    }
}

Ctrl_g() {
    if ForwardToTeamViewer(Sleep(500)) {
        Send("joel.hurtado@3shape.com")
        Sleep(7500)  ; Espera dos segundos antes de enviar el texto
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
; KEY COMBS TOOLTIP 
Ctrl_y() {
    global myGui, guiVisible
    if !myGui {
        myGui := Gui()
        myGui.Title := "COMBOS"
        myGui.Opt("AlwaysOnTop -SysMenu")
        myGui.SetFont("s10")
        myGui.Add("Text", , "KEY COMBINATIONS:`n"
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
            . "B - LclHst 27027`n"
            . "D - Device manager`n"
            . "F - Firmware*`n"
            . "G - Support Mode`n"
            . "H - Hardware Test*`n"
            . "I - Firewall Advanced`n"
            . "K - Clear Clipboard`n"
            . "N - Specified PATH `n"
            . "O - CMD - Service Tag`n"
            . "Q - Server Packages`n"
            . "S - Scan Suit`n"
            . "W - Wireless Drivers*`n"
            . "X - Streams`n"
            . "Y - TECLAS INFORMATION`n"
            . "-------------------------------`n"
            . "1 - Servers`n"
            . "2 - Logs folder`n"
            . "3 - SQL*`n"
            . "4 - Keys*`n"
            . "5 - Dxdiag`n"
            . "6 - Appwiz`n"
            . "7 - Sys Info`n"
            . "8 - LclHst 8000`n"
            . "9 - Nvidia Drivers`n"
            . "0 - Services")
        ; Agregar la imagen a la derecha del texto
        myGui.Add("Picture", "x+20 y0 w400 h200", ".\Letter.png")
    }
    if guiVisible {
        myGui.Hide()
        guiVisible := false
        Return
    } else {
        myGui.Show("AutoSize Center")
        guiVisible := true
        ; SetTimer(() => myGui.Destroy(), 20000)
        Return
    }
}

; ---------------------------------------------------------------------------------------------------------------------------------





Ctrl_1() {
    if ForwardToTeamViewer("{LWin Down}e{LWin Up}") {
        Sleep(2000)
        Send("{Ctrl Down}{l}{Ctrl Up}")  ; Ctrl + L
        Sleep(500)  ; Espera medio segundo antes de enviar el texto
        Send("%programdata%\3Shape\DentalDesktop\ThreeShape.TRIOS\Servers")
        Sleep(500)
        Send '{Enter}'
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
    if ForwardToTeamViewer("{LWin Down}r{LWin Up}") {
        Sleep(2000)  ; Espera dos segundos antes de enviar el texto
        Send("dxdiag")
        Sleep(500)
        Send '{Enter}'
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_6() {
    if ForwardToTeamViewer("{LWin Down}r{LWin Up}") {
        Sleep(2000)  ; Espera dos segundos antes de enviar el texto
        Send("appwiz.cpl")
        Sleep(500)
        Send '{Enter}'
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_7() {
    if ForwardToTeamViewer("{LWin Down}{LWin Up}") {
        Sleep(2000)  ; Espera dos segundos antes de enviar el texto
        Send("system info")
        Sleep(500)
        Send '{Enter}'
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_8() {
    if ForwardToTeamViewer("{LWin Down}{LWin Up}") {
        Sleep(2000)  ; Espera dos segundos antes de enviar el texto
        Send("Edge")
        Sleep(500)
        Send '{Enter}'
        Sleep(2000)
        Send("{Ctrl Down}{t}{Ctrl Up}")  ; Ctrl + L
        Sleep(1500)
        Send("{Ctrl Down}{l}{Ctrl Up}")  ; Ctrl + L
        Sleep(1500)
        Send("http://localhost:8000/3ShapeWirelessService//")  ; Ctrl + L
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_9() {
    if ForwardToTeamViewer("{LWin Down}{LWin Up}") {
        Sleep(2000)  ; Espera dos segundos antes de enviar el texto
        Send("Edge")
        Sleep(500)
        Send '{Enter}'
        Sleep(2000)
        Send("{Ctrl Down}{t}{Ctrl Up}")  ; Ctrl + L
        Sleep(1500)
        Send("{Ctrl Down}{l}{Ctrl Up}")  ; Ctrl + L
        Sleep(1500)
        Send("https://www.nvidia.com/en-us/drivers/")  ; Ctrl + L
        return  ; Salir si el envío fue exitoso
    }
    
      ; Si no es TeamViewer, muestra el nombre del PC
}

Ctrl_0() {
    if ForwardToTeamViewer("{LWin Down}{LWin Up}") {
        Sleep(2000)  ; Espera dos segundos antes de enviar el texto
        Send("services")
        Sleep(500)
        Send '{Enter}'
        return  ; Salir si el envío fue exitoso
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