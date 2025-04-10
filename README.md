# Skrive

Estamos desarrollando una aplicación para obtener la información de nuestros clientes.

Si tienes algun error, diferente a los mostrados por la app, guarda el caso con un numero de caso y luego corre Ctrl + Shif + R, carga el numero de caso 

**Teclas de acceso rápido:**

- **{D}**: Click Derecho
- **{I}**: Click Izquierdo
- **[Botón]**: Nombre del Botón

La aplicación cuenta con 5 pestañas de navegación, que pueden activarse con las siguientes combinaciones de teclas:

- **Ctrl + Alt + I**: Información
- **Ctrl + Alt + R**: Sesión Remota
- **Ctrl + Alt + A**: Añadir Información
- **Ctrl + Alt + E**: Correo Electrónico
- **Ctrl + Alt + 2**: Segunda Línea - Callback
- **Ctrl + Alt + P**: Photos - Excel

Tenemos un shortcut que copia TODO lo necesario para el CRM en el portapaeles, se sugiere llenar todos los campos y llenar bien el issue, problem y root cause, tambien de dejar recommendaciones en el espacio del email.
Para no perderse entre casos se sugiere limpiar el portapapeles antes de correr el comando, para limpiarlo puedes usar  Ctrl + Shif + K, de la ventana de Hotkeys (solo para windows 10):

- **Ctrl + Alt + Q**: Copia todo lo necesario para CRM

Shortcuts para activar/ocultar las herramientas:

- **Ctrl + Shift + U**: Muestra y oculta la herramienta SKRIVE
- **Ctrl + Shift + Y**: Muestra y oculta la lista de Hotkeys rápidas para procesos automatizados
**COMBOS**: Se le llama asi a la combinacion *Ctrl + Shift*

---

**Atajos para pegar valores en SKRIVE desde el portapapeles**

Al presionar las siguientes combinaciones de teclas (`Alt` + letra), se pegan automáticamente el ultimo valor del portapapeles en los respectivos campos de SKRIVE, además de ejecutar acciones adicionales como guardar o correr comandos:

- **Alt + D** → Pega valor en **Dongle**
- **Alt + P** → Pega valor en **Phone**
- **Alt + E** → Pega valor en **Email**
- **Alt + C** → Pega valor en **Company**
- **Alt + I** → Pega valor en **SID**
- **Alt + A** → Pega valor en **Case Number**
- **Alt + V** → Pega valor en **Survey**
- **Alt + T** → Copia el valor de **Descripción**
- **Alt + H** → Acceso rápido a **Phone Call**
- **Alt + S** → Guarda caso por su **Info**
- **Alt + Y** → Ejecuta el comando que toma los logs (**Ctrl + A**)


## Pestaña de Información

### [Guardar - Cargar Plantilla Existente]

Las dos primeras listas permiten seleccionar una opción y tipo de caso para guardar o cargar una plantilla. Estas acciones afectan solo las casillas que no contienen datos personales, y ambas casillas deben estar llenas para que funcionen.

- **{D}**: Guarda el caso con los datos actuales.
- **{I}**: Carga una plantilla de las disponibles en las carpetas de "Complaint Request Miscellaneous".

### [TVID-TVPSS]

- **{D}**: Pega el último valor del portapapeles.
- **{I}**: Abre una sesión remota automática con los datos del televisor si está abierto.

### [INT-PH]

- **{D}**: Añade una nota interna y una nota en HelpJuice.
- **{I}**: Añade una nota de llamada telefónica.
- **Alt + {D}**: Añade una nota con nota interna, HelpJuice y registros.

### [C1st]

- **{D}**: Título.
- **{I}**: Título + C1st.

### [Campos: Nombre, Teléfono, Correo Electrónico, Nombre de la Empresa, Dongle, SID, GUI, Número de Serie del Escáner, ID de PC, Número de Caso, Encuesta HJ, RC]

- **{D}**: Pega el contenido.
- **{I}**: Elimina el contenido.
- **Alt + {I}**: Copia al portapapeles.

### [Versión de Software]

- **{D}**: Abre las opciones de módulo y versión; también se puede escribir lo que se desee.

### [Descripción]

- **{D}**: Copia la nota de descripción al portapapeles.
- **{I}**: Muestra cómo queda la descripción.

### [Problema-Solución]

- **{D}**: Crea una nota de conclusión (RC: S:).
- **{I}**: Elimina la nota.

### [Tecla Rápida]

- **{D}**: Abre la lista de comandos.
- **{I}**: Ejecuta comandos relacionados con el número de caso.
- **Alt + {D}**: Abre la carpeta de registros.

### [Guardar - Cargar Información]

- **{D}**: Guarda en la carpeta predeterminada `Escritorio\CasesJSON`, por número de caso.
- **{I}**: Carga un JSON guardado para recuperar información de un caso, si hay un número de caso.
- **Alt + {D}**: Permite seleccionar una carpeta diferente para guardar.

### [Limpiar]

- **{D}**: Limpia todas las casillas.

## Sesión Remota

### [Remoto]

- **{D}**: Copia al portapapeles los pasos de las casillas que no estén vacías.
- **{I}**: Copia al portapapeles el título de la sesión remota.

## Añadir Información

### [Pregunta de Sondeo]

- **{D}**: Copia al portapapeles lo escrito en este campo.
- **{I}**: Copia al portapapeles lo escrito en este campo.

## Correo Electrónico

### [Correo]

- **{D}**: Envía un correo con la información ingresada.
- **{I}**: Título del correo.
- **Alt + {D}**: Correo electrónico predeterminado.

## Segunda Línea - Callback

### [Segunda Línea - Callback]

- **{D}**: Título del correo, cuerpo del correo y abre un archivo de Excel.
- **{I}**: Nota de Callback.

### [Acceso Anticipado]

- **{D}**: Título y nota de Early Access.
