Correo automáticos

Este script sirve para enviar correos automáticamente en un computador con MacOS. La idea es poder 
enviar mensajes a varias personas en donde cada mensaje tenga palabras o números diferentes. Los
caracteres o valores de estas palabras o números deben estar en columnas de un archivo de Excel, al
igual que los correos electrónicos, ajustando la fila de cada correo electrónico con las palabras 
o valores que quieren ser enviados a dicho correo. 

Además del archivo de excel, se requiere un archivo de texto plano en donde se escriba el mensaje 
de forma general, escribiendo los nombres de las varibles que representen las columnas de las palabras 
o los números del archivo de Excel. Estos nombres de las variables deben ser escritos antecedidos por
un ampensand (&) y ambas cosas entre virgulillas (~).

Los cuadros de diálogo que aparecerán son lo siguientes:

    ¿Cuál es la ruta del archivo de texto con el mensaje base?

        - Se debe indicar la ruta del archivo de texto plano en donde se escribió el mensaje general. Esto
          se puede hacer arrastrando el archivo a la caja de texto o con click secundario sobre el archivo más
          la tecla "option" y escogiendo "copiar "--" como ruta de acceso".

    ¿En cuál columna está la variable '---'?  

        - Luego de agregar la ruta del archivo de texto, el programa buscará las variables que éste contenga.
          En esta pregunta hay que indicar en la caja de texto a qué columna hace referencia la varibles "---"
          que se escribió en el archivo de texto. Esta pregunta se repetirá tantas veces como variables haya en 
          el mensaje.

    ¿Cuál es la ruta del archivo de Excel?

        - Se le debe indicar la ruta del archivo de Excel en donde se encuentran los valores de las variables. Esto
          se puede hacer arrastrando el archivo a la caja de texto o con click secundario sobre el archivo más
          la tecla "option" y escogiendo "copiar "--" como ruta de acceso".

    ¿Cuál es la fila del primer estudiante?

        - Se debe indicar en la caja de texto desde qué fila se quiere empezar a leer los datos en el archivo de Excel.

    ¿Cuál es la fila del último estudiante?

        - Se debe indicar en la caja de texto hasta qué fila se quiere leer los datos en el archivo de Excel.

    ¿Cuál es la fila del último estudiante?

        - Se debe indicar en la caja de texto desde cuál correo se quiere enviar todos los mensajes. Este cuanta 
          debe estar agregada en "Mail".

    ¿Cuál es el asunto del mensaje?

        - Se debe escribir en la caja de texto con qué asunto se quiere enviar todos los mensaje.

    ¿En qué columna están los correos de los estudiante?

        - En esta pregunta hay que indicar en la caja de texto en qué columna del archivo de Excel están los 
          correos de los destinatarios.

    ¿Quieres enviar un archivos adjuntos diferentes para cada correo?

        - "No enviar adjunto": No se podrá enviar ningún archivo adjunto a todos los correos.

            ¿Quieres confirmar antes de enviar cada correo?

                - "Si": Se mostrará información de cada mensaje y pedirá una confirmación para enviarlo. Puede confirmarse 
                  el envio de todos los correos posteriores si se selecciona la opción "Confirmar todos".

                - "No": Se enviarán todos los correos sin mostrar información al respecto. 

        - "No": Se podrá enviar archivos adjuntos iguales para todos lo correos. 

            ¿Cuál es la ruta del archivo adjunto general?

                - En la caja de texto se debe indicar la ruta del archivo que se quiere mandar como adjunto. Pueden
                  agregarse varios archivos si se selecciona la opción "Agregar". Cuando se llegue al último archivo
                  se debe seleccionar "Continuar" y no "Agregar".

                    ¿Quieres confirmar antes de enviar cada correo?

                        - "Si": Se mostrará información de cada correo y pedirá una confirmación para enviarlo. Puede confirmarse 
                          el envio de todos los correos posteriores si se selecciona la opción "Confirmar todos".

                        - "No": Se enviarán todos los correos sin mostrar información al respecto. 

        - "Sí": Se podrá enviar archivos adjuntos diferentes para cada correo.

            ¿Quieres confirmar antes de enviar cada correo?

                - "Si": Se mostrará información de cada correo y pedirá una confirmación para enviarlo.

                    ¿Cuál es la ruta del archivo adjunto para "----"?

                        - En la caja de texto se debe indicar la ruta del archivo que se quiere mandar como adjunto al correo "----". 
                          Pueden agregarse varios archivos si se selecciona la opción "Agregar". Cuando se llegue al último archivo
                          se debe seleccionar "Continuar" y no "Agregar".    

                -"No": Se enviarán todos los correos sin mostrar información al respecto. 

                    ¿Cuál es la ruta del archivo adjunto para "----"?

                        - En la caja de texto se debe indicar la ruta del archivo que se quiere mandar como adjunto al correo "----". 
                         Pueden agregarse varios archivos si se selecciona la opción "Agregar". Cuando se llegue al último archivo
                         se debe seleccionar "Continuar" y no "Agregar". 

Como ejemplo, si se tienen los siguientes archivos de texto y Excel:

    - - - - - - - - - - - - - - - - - - -
    | Buenos días, ~&nombreEstudiante~. |
    |                                   |
    | Sus notas son:                    |
    | Nota 1: ~&nota1Estudiante~.       |
    | Nota 2: ~&nota2Estudiante~.       | 
    | Nota 3: ~&nota3Estudiante~.       | 
    | Nota 4: ~&nota4Estudiante~.       | 
    |                                   |
    | Hasta luego.                      |
    |                                   |
    - - - - - - - - - - - - - - - - - - -

y

    - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
    |Cristian Merchán	|cdmerchans@gmail.com	|1	|1	|1	|1 |
    |----------------------------------------------------------|
    |David Merchán	    |cdmerchan@hotmail.com	|2	|2	|2	|2 |
    |----------------------------------------------------------|
    |Cristian Sarmiento	|cdmerchans@unal.edu.co	|3	|3	|3	|3 |
    |----------------------------------------------------------|
    |David Sarmiento	|cdmerchans@icloud.com	|4	|4	|4	|4 |
    |----------------------------------------------------------|
    - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Se obtendrán los siguientes correos:

    - Para cdmerchans@gmail.com:

        - - - - - - - - - - - - - - - - - - -
        |Buenos días, Cristian Merchán.     |
        |                                   |
        |Sus notas son:                     |
        |Nota 1: 1.                         | 
        |Nota 2: 1.                         | 
        |Nota 3: 1.                         |
        |Nota 4: 1.                         |
        |                                   |
        |Hasta luego.                       |
        - - - - - - - - - - - - - - - - - - -

    - Para cdmerchan@hotmail.com:

        - - - - - - - - - - - - - - - - - - -
        |Buenos días, David Merchán.        |
        |                                   |
        |Sus notas son:                     |
        |Nota 1: 2.                         |
        |Nota 2: 2.                         |
        |Nota 3: 2.                         |
        |Nota 4: 2.                         | 
        |                                   |
        |Hasta luego.                       |
        - - - - - - - - - - - - - - - - - - -

    - Para cdmerchans@unal.edu.co:
        - - - - - - - - - - - - - - - - - - -
        |Buenos días, Cristian Sarmiento.   |
        |                                   |
        |Sus notas son:                     |
        |Nota 1: 3.                         |
        |Nota 2: 3.                         |
        |Nota 3: 3.                         |
        |Nota 4: 3.                         |
        |                                   |
        |Hasta luego.                       |
        - - - - - - - - - - - - - - - - - - - 

    - Para cdmerchans@icloud.com:
    
        - - - - - - - - - - - - - - - - - - -
        |Buenos días, David Sarmiento.      |
        |                                   |
        |Sus notas son:                     |
        |Nota 1: 4.                         |
        |Nota 2: 4.                         |
        |Nota 3: 4.                         |
        |Nota 4: 4.                         |
        |                                   |
        |Hasta luego.                       |
        - - - - - - - - - - - - - - - - - - -
