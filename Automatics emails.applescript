on splitText(theText, theDelimiter)
	set AppleScript's text item delimiters to theDelimiter
	set theTextItems to every text item of theText
	set AppleScript's text item delimiters to ""
	return theTextItems
end splitText

on findAndReplaceInText(theText, theSearchString, theReplacementString)
	set AppleScript's text item delimiters to theSearchString
	set theTextItems to every text item of theText
	set AppleScript's text item delimiters to theReplacementString
	set theText to theTextItems as string
	set AppleScript's text item delimiters to ""
	return theText
end findAndReplaceInText



set sharesFileName to display dialog "¿Cuál es la ruta del archivo de texto con el mensaje base?" default answer "/Users/cristianmerchan/Desktop/Mensaje.txt" buttons {"Cancelar", "Continuar"} default button "Continuar" cancel button "Cancelar"
set sharesFileName to text returned of sharesFileName
set sharesLines to read sharesFileName as «class utf8»
set sharesLines to splitText(sharesLines, "~")
set listVaribles to {}
set listColumns to {}

repeat with theCurrentListItem in sharesLines
	if theCurrentListItem starts with "&" then
		set theCurrentListItem to findAndReplaceInText(theCurrentListItem, "&", "")
		set ColumnVarible to display dialog "¿En cuál columna está la variable '" & theCurrentListItem & "'?" default answer "" buttons {"Cancelar", "Continuar"} default button "Continuar" cancel button "Cancelar"
		set ColumnVarible to text returned of ColumnVarible
		copy ColumnVarible to the end of the listColumns
		copy theCurrentListItem to the end of the listVaribles
	end if
end repeat

set numeroDeVariables to count of listVaribles

set rutaArchivo to display dialog "¿Cuál es la ruta del archivo de Excel?" default answer "/Users/cristianmerchan/Desktop/Notas.xlsx" buttons {"Cancelar", "Continuar"} default button "Continuar" cancel button "Cancelar"
set rutaArchivo to text returned of rutaArchivo

set minimoDeEstudiantes to display dialog "¿Cuál es la fila del primer estudiante?" default answer "1" buttons {"Cancelar", "Continuar"} default button "Continuar" cancel button "Cancelar"
set maximoDeEstudiantes to display dialog "¿Cuál es la fila del último estudiante?" default answer "1" buttons {"Cancelar", "Continuar"} default button "Continuar" cancel button "Cancelar"
set minimoDeEstudiantes to text returned of minimoDeEstudiantes
set maximoDeEstudiantes to text returned of maximoDeEstudiantes
set minimoDeEstudiantes to minimoDeEstudiantes as integer
set maximoDeEstudiantes to maximoDeEstudiantes as integer

set correoSalida to display dialog "¿Desde qué correo quieres enviar los mensajes?" default answer "cdmerchans@icloud.com" buttons {"Cancelar", "Continuar"} default button "Continuar" cancel button "Cancelar"
set correoSalida to text returned of correoSalida

set asunto to display dialog "¿Cuál es el asunto del mensaje?" default answer "Prueba" buttons {"Cancelar", "Continuar"} default button "Continuar" cancel button "Cancelar"
set asunto to text returned of asunto

set colCorreoEstudiente to display dialog "¿En qué columna están los correos de los estudientes?" default answer "" buttons {"Cancelar", "Continuar"} default button "Continuar" cancel button "Cancelar"
set colCorreoEstudiente to text returned of colCorreoEstudiente

set archivosAdjunto to {}
set banderaAdjunto to display dialog "¿Quieres enviar un archivo adjunto diferentes para cada correo?" buttons {"No enviar adjunto", "Sí", "No"} default button "No"

if button returned of banderaAdjunto = "No" then
	set archivoAdjuntoBoton to "Agregar"
	repeat while archivoAdjuntoBoton is "Agregar"
		set archivoAdjunto to display dialog "¿Cuál es la ruta del archivo adjunto?" default answer " " buttons {"Agregar", "Cancelar", "Continuar"} default button "Continuar" cancel button "Cancelar"
		set archivoAdjuntoBoton to button returned of archivoAdjunto
		set archivoAdjunto to text returned of archivoAdjunto
		copy archivoAdjunto to the end of the archivosAdjunto
	end repeat
end if

set banderaConfirmacion to display dialog "¿Quieres confirmar antes de enviar cada correo?" buttons {"Sí", "No"} default button "No"

set banderaConfirmarTodos to "No"
repeat with fila from minimoDeEstudiantes to maximoDeEstudiantes
	set estudiante to {}
    set archivosAdjunto to {}
	set sharesLines to read sharesFileName as «class utf8»
	set sharesLines to splitText(sharesLines, "~")
	repeat with j from 1 to numeroDeVariables
		tell application "Microsoft Excel"
			open rutaArchivo
			set readValue to string value of cell (item j of listColumns & fila)
			copy readValue to the end of the estudiante
			set correoEstudiante to string value of cell (colCorreoEstudiente & fila)
		end tell
	end repeat
	set i to 1
	set j to 1
	set mensaje to ""
	repeat with theCurrentListItem in sharesLines
		if theCurrentListItem starts with "&" then
			set item j of sharesLines to item i of estudiante
			set i to i + 1
		end if
		set mensaje to mensaje & item j of sharesLines
		set j to j + 1
	end repeat
	
	if button returned of banderaAdjunto = "Sí" then
		set archivoAdjuntoBoton to "Agregar"
		repeat while archivoAdjuntoBoton is "Agregar"
			set archivoAdjuntoBoton to display dialog "¿Cuál es la ruta del archivo adjunto para " & correoEstudiante & "?" default answer "" buttons {"Agregar","Cancelar", "Continuar"} default button "Continuar" cancel button "Cancelar"
			set archivoAdjunto to text returned of archivoAdjuntoBoton
			copy archivoAdjunto to the end of the archivosAdjunto
			set archivoAdjuntoBoton to button returned of archivoAdjuntoBoton
		end repeat
	end if
	
	set banderaConfirmarTodos to "No"
	
	if button returned of banderaConfirmacion = "Sí" and banderaConfirmarTodos is "No" and button returned of banderaAdjunto is not "Sí" then
		set confirmacion to display dialog "De: " & correoSalida & "
Para: " & correoEstudiante & "
Mensaje: 
" & mensaje & "
Adjunto: " & archivosAdjunto buttons {"Confirmar todos", "Cancelar", "Enviar"} default button "Enviar" cancel button "Cancelar"
		if button returned of confirmacion = "Confirmar todos" then
			set banderaConfirmarTodos to "Sí"
		end if
	end if
	
	if button returned of banderaConfirmacion = "Sí" and banderaConfirmarTodos is "No" and button returned of banderaAdjunto is "Sí" then
		set confirmacion to display dialog "De: " & correoSalida & "
Para: " & correoEstudiante & "
Mensaje: 
" & mensaje & "
Adjunto: " & archivosAdjunto buttons {"Cancelar", "Enviar"} default button "Enviar" cancel button "Cancelar"
	end if
	
	
	tell application "Mail"
		
		set theFrom to correoSalida
		set theTos to {correoEstudiante}
		set theCcs to {""}
		set theBccs to {""}
		
		set theSubject to asunto
		set theContent to mensaje
		
		set theSignature to ""
		set theDelay to 1
		
		set theMessage to make new outgoing message with properties {sender:theFrom, subject:theSubject, content:theContent, visible:false}
		tell theMessage
			repeat with theTo in theTos
				make new recipient at end of to recipients with properties {address:theTo}
			end repeat
			repeat with theCc in theCcs
				make new cc recipient at end of cc recipients with properties {address:theCc}
			end repeat
			repeat with theBcc in theBccs
				make new bcc recipient at end of bcc recipients with properties {address:theBcc}
			end repeat
		end tell
		tell content of theMessage
			repeat with theAttachment in archivosAdjunto
				make new attachment with properties {file name:theAttachment} at after last paragraph
				delay theDelay
			end repeat
		end tell
		send theMessage
		
	end tell
end repeat

set finalAplicacion to display dialog "Los mensajes fueron enviados."
