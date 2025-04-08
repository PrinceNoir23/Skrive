# Skrive
Estamos creando una aplicación para obtener la información de nuestros clientes.

{D} = Click Derecho 
{I} = Click Izquierdo
[ Boton ] = Nombre del Boton



La aplicacion tiene 5 tabs de navegacion, que pueden ser activados con las siguientes combinaciones de teclas 
Ctrl + Alt + I -> Information
Ctrl + Alt + R -> Rmt Session
Ctrl + Alt + A -> Add Info
Ctrl + Alt + E -> Email
Ctrl + Alt + 2 -> 2nd Line - Call Back


Ctrl + Shift + U -> Muestra y Oculta la herramienta
Ctrl + Shift + Y -> Muestra y Oculta la lista de hotkeys para procesos automatizados

+++ INFO TAB +++

[Save - Load Existing Template]

Las dos primeras listas de arriba permiten seleccionar una opcion y tipo de caso, para guardar o cargar un Template de un caso, va a guarda y a cargar solo las casillas que no poseen datos personales, y ambas casillas deben estar llenas para que funcione.
    - {D} --> Guarda el caso con los datos actuales
    - {I} --> Llama un template de los que se encuentren en las carpetas de Complaint Request Miscellaneous


[TVID-TVPSS]
    - {D} --> Pega el ultimo valor del clipboard
    - {I} --> Abre remote session automatico con los datos del tv si tv esta abierto
[INT-PH] 
    - {D} --> Internal Note + HelpJuice note 
    - {I} --> Phone Call Note 
    - Alt + {D} --> Internal Note + HelpJuice + Logs note

[C1st]
    - {D} --> Titulo
    - {I} --> Titulo + C1st 
[Name-Phone-Email-Company Name-Dongle-SID-GUI-Scanner S/N-PC ID-Case Number-HJ-Survey-RC]
    - {D} -->  pegar
    - {I} -->  eliminar
    - Alt + {I} --> copiar al portapapeles 
[SoftwareVersion]
    - {D} -->  Abre las opciones de modulo y version, se puede escribir lo que se desee tambien
[Description]
    - {D} -->  Copia el description note al portapapeles
    - {I} --> muestra como queda la descripcion 
[Issue-Solution]
    - {D} -->  crea conclsuion note (RC: S:)
    - {I} --> eliminar
[HotKey]
    - {D} --> COMBOS + Y :
        + abrira la pestaña que enseña la lista de comandos ¡¡¡ para la ultima version los botones de comandos !!!
    - {I} --> COMBOS + A:
        + si case number colocado en la casilla corre el comando que saca los logs desde powershell
        + si no case number colocado en la casilla va a pedirte que llenes la casilla (Permite ingresar cualquier valor e.i. "aaaaaaa")
    - Alt + {D} --> COMBOS + 2 :
        + abrira logs folder !!!
[Save - Load Info]

    -{D} --> Guarda en la carpeta por default Escritorio\CasesJSON, por case number
    -{I} --> Cargamos un JSON guardado para recuperar info de un caso, si hay un case number 
    -Alt + {D} --> Permite seleccionar una carpeta diferente para guardar

[Clear]
    - {D} -->  Limpia todas las casillas


+++ Remote Session +++

[Remote]
    - {D} --> pega al portapapeles los pasos de las casillas que no esten vacias
    - {I} --> pega al portapapeles remote session title


+++ Add Info +++

[ProbingQuestion]
    - {D} --> Copia al portapapeles lo escrito en ese box
    - {I} --> Copia al portapapeles lo escrito en ese box


+++ Email +++

[Email]
    - {D} --> default email 
    - {I} --> email title
    - Alt + {D} --> email con la informacion ingresada
    










