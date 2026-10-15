// Datos del informe - Guía de Laboratorio N.º 05 "Practica Linux"
// Estructura basada en el PDF original y adaptada al formato Word institucional (APA 7)

const FUENTE = "Fuente: Elaboración propia.";

const sections = [
    {
        key: "a",
        title: "a) Creación de un aplicativo de monitoreo del sistema",
        intro: [
            "En esta primera parte del laboratorio se trabaja sobre el directorio /opt de la máquina virtual con Linux Debian, empleando comandos de diagnóstico como free -h y df -h para conocer el estado de la memoria RAM, la memoria swap y el espacio disponible en el disco. A partir de esta información se construye progresivamente un archivo de texto denominado capacidades.txt, el cual funciona como bitácora o reporte de estado del sistema.",
            "Posteriormente, dicho reporte se automatiza mediante un script de Bash (reporte01.sh) y se programa su ejecución periódica utilizando la herramienta crontab, lo que permite comprender de manera práctica cómo se automatizan tareas administrativas recurrentes en un entorno Linux de producción."
        ],
        figures: [
            {
                lead: "Para iniciar el laboratorio, nos ubicamos dentro del directorio /opt utilizando el comando cd opt, y a continuación ejecutamos los comandos free -h y df -h para visualizar, respectivamente, el estado de la memoria del sistema y el espacio disponible en cada partición del disco.",
                cmd: ["root@JorgeC:~# cd opt", "root@JorgeC:~/opt# free -h", "root@JorgeC:~/opt# df -h"],
                title: "Verificación de memoria y espacio en disco desde el directorio /opt",
                desc: "La captura evidencia la salida combinada de free -h, que expresa los valores en unidades legibles (GiB/MiB), y de df -h, que detalla el sistema de archivos, el tamaño total, el espacio usado, el disponible y el porcentaje de uso de cada punto de montaje, entre ellos la partición principal /dev/sda1.",
                note: "El parámetro -h (human-readable) es fundamental en entornos de administración de sistemas porque evita tener que realizar conversiones manuales desde bytes a unidades más comprensibles como MB o GB."
            },
            {
                lead: "A continuación, se crea el archivo de texto que almacenará la información recolectada durante la práctica, empleando el comando touch.",
                cmd: ["root@JorgeC:~/opt# touch capacidades.txt"],
                title: "Creación del archivo capacidades.txt",
                desc: "El comando touch genera un archivo vacío con marca de tiempo actual si este no existe previamente en el directorio de trabajo; en caso de existir, simplemente actualiza su fecha de modificación sin alterar su contenido.",
                note: "Este archivo será el contenedor central de todos los datos de diagnóstico generados en las siguientes figuras, por lo que su correcta creación es un paso crítico para el resto del procedimiento."
            },
            {
                lead: "Empleamos el comando ls para confirmar el contenido actual del directorio /opt y verificar que el archivo capacidades.txt se haya creado correctamente junto al resto de archivos y carpetas preexistentes.",
                cmd: ["root@JorgeC:~/opt# ls"],
                title: "Listado del contenido del directorio /opt",
                desc: "La salida del comando ls muestra, en orden alfabético y organizados en columnas, tanto los archivos de texto de prueba (archivo01.txt a archivo05.txt) como las carpetas carpeta01 a carpeta05 y el archivo capacidades.txt recién creado.",
                note: "El comando ls sin parámetros adicionales no distingue entre archivos y directorios visualmente; para ello se recomienda emplear ls -l o ls --color en sesiones interactivas."
            },
            {
                lead: "Se ejecuta nuevamente el comando df -h con la finalidad de contrastar la información del espacio en disco y prepararla para su posterior almacenamiento en el archivo capacidades.txt.",
                cmd: ["root@JorgeC:~/opt# cat capacidades.txt", "root@JorgeC:~/opt# df -h"],
                title: "Confirmación de archivo vacío y nueva consulta de espacio en disco",
                desc: "El comando cat capacidades.txt no arroja ninguna salida visible, confirmando que el archivo se encuentra vacío tras su creación; inmediatamente después, df -h despliega el detalle habitual de particiones y sistemas de archivos montados.",
                note: "Verificar el estado vacío de un archivo antes de escribir en él es una buena práctica quNe evita sobrescribir información relevante de manera accidental en entornos de producción."
            },
            {
                lead: "Con el objetivo de persistir la información del disco en el archivo de texto, se utiliza el operador de redirección de anexado (>>) junto al comando df -h.",
                cmd: ["root@JorgeC:~/opt# df -h >> capacidades.txt"],
                title: "Redirección de la salida de df -h hacia capacidades.txt",
                desc: "El operador >> anexa la salida estándar del comando al final del archivo indicado sin eliminar su contenido previo, a diferencia del operador simple > que sobrescribiría por completo el archivo de destino.",
                note: "Comprender la diferencia entre > y >> resulta esencial para evitar la pérdida accidental de registros históricos en archivos de bitácora (logs)."
            },
            {
                lead: "Para verificar que la redirección se ejecutó correctamente, se visualiza el contenido íntegro del archivo capacidades.txt mediante el comando cat.",
                cmd: ["root@JorgeC:~/opt# cat capacidades.txt"],
                title: "Verificación del contenido almacenado en capacidades.txt",
                desc: "La salida del comando confirma que la información de df -h fue correctamente almacenada en el archivo, replicando exactamente el formato tabular original con las columnas Filesystem, Size, Used, Avail, Use% y Mounted on.",
                note: "El comando cat es una herramienta de concatenación de archivos que, utilizada sin parámetros adicionales, también sirve como visor rápido de contenido en texto plano."
            },
            {
                lead: "Se repite el procedimiento de redirección varias veces consecutivas para demostrar cómo el operador >> acumula información en lugar de sobrescribirla.",
                cmd: ["root@JorgeC:~/opt# df -h >> capacidades.txt", "root@JorgeC:~/opt# cat capacidades.txt"],
                title: "Acumulación de registros mediante ejecuciones sucesivas de df -h >> capacidades.txt",
                desc: "Al repetir el comando, el archivo capacidades.txt va incrementando su tamaño con bloques duplicados de la salida de df -h, cada uno de ellos anexado inmediatamente después del anterior, lo que ilustra de forma práctica el comportamiento acumulativo del operador de redirección.",
                note: "En un entorno real, este comportamiento se aprovecha para construir archivos de histórico o bitácoras de monitoreo que registran el estado del sistema en distintos momentos del tiempo."
            },
            {
                lead: "Para depurar el archivo y quedarnos únicamente con la información relevante de la partición principal, se combina una secuencia de comandos: eliminación, recreación y filtrado con grep.",
                cmd: [
                    "root@JorgeC:~/opt# rm capacidades.txt",
                    "root@JorgeC:~/opt# touch capacidades.txt",
                    "root@JorgeC:~/opt# df -h | grep /dev/sda1"
                ],
                title: "Depuración del archivo y filtrado de la partición /dev/sda1 con grep",
                desc: "La secuencia primero elimina el archivo capacidades.txt con rm, luego lo recrea vacío con touch y, finalmente, canaliza (pipe |) la salida de df -h hacia grep, el cual filtra únicamente la línea correspondiente a la partición /dev/sda1, mostrando su capacidad total, uso y disponibilidad.",
                note: "El operador de tubería (|) permite encadenar comandos de modo que la salida estándar de uno se convierta en la entrada estándar del siguiente, siendo la base de la filosofía Unix de 'herramientas pequeñas que hacen una sola cosa bien'."
            },
            {
                lead: "Se emplea el comando date sin parámetros para visualizar la fecha y hora actuales del sistema operativo.",
                cmd: ["root@JorgeC:~/opt# date"],
                title: "Consulta de la fecha y hora del sistema con date",
                desc: "El comando date, ejecutado sin argumentos, devuelve por defecto el día de la semana, el mes, el día del mes, la hora en formato de 12 horas con indicador AM/PM, la zona horaria y el año, tal como se observa en la salida 'Mon Apr 15 06:23:53 PM PDT 2024'.",
                note: "El comando date admite una amplia variedad de formatos de salida personalizados mediante el símbolo + seguido de especificadores como %A, %D o %r, los cuales se emplearán más adelante en esta guía."
            },
            {
                lead: "Se utiliza el comando free -m para obtener información detallada sobre el uso de memoria del sistema expresada en megabytes.",
                cmd: ["root@JorgeC:~/opt# free -m"],
                title: "Consulta de memoria del sistema en megabytes con free -m",
                desc: "A diferencia de free -h, el parámetro -m fuerza la representación de los valores exclusivamente en megabytes, sin redondeos automáticos a otras unidades, lo cual resulta útil cuando se requiere procesar la salida de manera programática mediante scripts.",
                note: "Elegir unidades fijas como -m o -k es una práctica recomendada al momento de automatizar el procesamiento de datos, pues evita inconsistencias que podrían surgir del formato variable de -h."
            },
            {
                lead: "Se construye una secuencia más elaborada que combina separadores visuales, fecha, información de disco y de memoria, todo ello almacenado ordenadamente en capacidades.txt.",
                cmd: [
                    "root@JorgeC:~/opt# rm capacidades.txt",
                    'root@JorgeC:~/opt# echo "----------------------------" >> capacidades.txt',
                    "root@JorgeC:~/opt# date >>capacidades.txt",
                    "root@JorgeC:~/opt# df -h | grep /dev/sda1 >>capacidades.txt",
                    "root@JorgeC:~/opt# free -m >>capacidades.txt",
                    'root@JorgeC:~/opt# echo "----------------------------" >> capacidades.txt',
                    "root@JorgeC:~/opt# cat capacidades.txt"
                ],
                title: "Generación de un reporte estructurado con separadores, fecha, disco y memoria",
                desc: "Esta secuencia demuestra cómo, mediante encadenamiento de comandos simples y redirecciones sucesivas, es posible construir un reporte legible y ordenado que incorpora líneas de guiones como separadores visuales, la marca temporal exacta de la consulta y los datos relevantes de disco y memoria en un solo archivo consolidado.",
                note: "Esta técnica constituye la base conceptual sobre la cual se construirá posteriormente el script reporte01.sh, automatizando exactamente esta misma secuencia de comandos."
            },
            {
                lead: "Se exploran distintas combinaciones del comando cut aplicado sobre la salida filtrada de free -m | grep Mem:, con el fin de identificar la posición exacta de los campos de interés dentro de la cadena de texto.",
                cmd: [
                    "root@JorgeC:~/opt# free -m >> capacidades.txt",
                    "root@JorgeC:~/opt# free -m | grep Mem:",
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d\" \" -f1",
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d\" \" -f1,2,3",
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d\" \" -f2",
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d\" \" -f3",
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d\" \" -f4",
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d\" \" -f1"
                ],
                title: "Exploración de campos con cut -d' ' sobre la línea Mem:",
                desc: "Debido a que la salida de free -m contiene múltiples espacios consecutivos entre columnas para lograr una alineación visual, el delimitador de espacio simple (-d' ') no separa los campos exactamente como columnas semánticas, por lo que muchas de las combinaciones de -f devuelven resultados vacíos o parciales, evidenciando una limitación práctica del comando cut ante campos de ancho variable.",
                note: "Este ejercicio resulta pedagógicamente valioso porque enseña a identificar, por prueba y error controlada, el número de campo correcto antes de integrarlo en un script definitivo."
            },
            {
                lead: "Se continúa la exploración incrementando progresivamente el número de campo solicitado a cut, hasta identificar aquellos que efectivamente contienen información útil.",
                cmd: [
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d' ' -f1",
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d' ' -f13",
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d' ' -f1,13",
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d' ' -f14",
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d' ' -f15",
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d' ' -f16",
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d' ' -f17",
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d' ' -f18",
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d' ' -f19",
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d' ' -f25"
                ],
                title: "Identificación del campo 13 como valor total de memoria",
                desc: "Tras un barrido sistemático de posiciones, se determina que el campo número 13 corresponde efectivamente al valor total de memoria (1967 MB), mientras que campos vecinos como 14 al 19 y 25 devuelven cadenas vacías por corresponder a espacios en blanco adicionales generados por el alineado de columnas de free.",
                note: "Este tipo de exploración manual es habitual durante el desarrollo de scripts de shell cuando no se dispone de documentación exacta sobre el formato interno de salida de un comando."
            },
            {
                lead: "Se combinan los campos previamente identificados (1 y 13) y se añade un tercer campo (30) para intentar capturar tres valores relevantes en una sola instrucción.",
                cmd: [
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d' ' -f1,13,30",
                    "root@JorgeC:~/opt# free -m | grep Mem:"
                ],
                title: "Combinación de campos 1, 13 y 30 con cut",
                desc: "La combinación -f1,13,30 devuelve 'Mem: 1967 1482', es decir, la etiqueta, el total de memoria y el valor disponible; se contrasta este resultado con la salida completa de grep Mem: para verificar visualmente la correspondencia entre las cifras extraídas y su posición original en la línea completa.",
                note: "Los números finales que aparecen en la salida representan, respectivamente, la cantidad total de memoria del sistema y la memoria disponible, valores clave para cualquier reporte de monitoreo."
            },
            {
                lead: "Se afinan las posiciones de los campos correspondientes a memoria usada y libre, probando distintas combinaciones hasta lograr capturar los tres valores de interés (total, usado y libre).",
                cmd: [
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d' ' -f1,13,20",
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d' ' -f1,13,21",
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d' ' -f1,13,22",
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d' ' -f1,13,22,30"
                ],
                title: "Determinación final de los campos total (13), usado (22) y libre (30)",
                desc: "Después de descartar los campos 20 y 21 por no arrojar valores numéricos, se comprueba que el campo 22 corresponde al valor de memoria usada (460 MB), por lo que la combinación final -f1,13,22,30 permite obtener en una sola línea el total, el uso y la disponibilidad de memoria del sistema.",
                note: "Este hallazgo será reutilizado en la construcción del script de reporte automatizado, evitando así repetir manualmente el proceso de exploración de campos."
            },
            {
                lead: "Con los campos ya identificados, se construye la secuencia definitiva que genera un reporte completo con encabezados descriptivos y se almacena en el archivo capacidades.txt.",
                cmd: [
                    "root@JorgeC:~/opt# rm capacidades.txt",
                    'root@JorgeC:~/opt# echo "----------------------------" >> capacidades.txt',
                    "root@JorgeC:~/opt# date >> capacidades.txt",
                    "root@JorgeC:~/opt# df -h | grep /dev/sda1 >>capacidades.txt",
                    'root@JorgeC:~/opt# echo "     Total Used Libre" >> capacidades.txt',
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d' ' -f1,13,22,30 >> capacidades.txt",
                    'root@JorgeC:~/opt# echo "----------------------------" >> capacidades.txt',
                    "root@JorgeC:~/opt# cat capacidades.txt"
                ],
                title: "Reporte consolidado con encabezados y valores de memoria y disco",
                desc: "El archivo capacidades.txt resultante integra, de forma ordenada, la fecha y hora de la consulta, la información de la partición /dev/sda1, un encabezado descriptivo para los valores de memoria y, finalmente, los valores totales, usados y libres extraídos previamente mediante cut.",
                note: "Esta estructura de reporte, aunque construida manualmente comando por comando, constituye exactamente la lógica que después se trasladará a un script de Bash reutilizable."
            },
            {
                lead: "Se identifica al último usuario registrado en el sistema consultando el archivo /etc/passwd, tomando su última línea con tail y extrayendo el nombre de usuario con cut.",
                cmd: ["root@JorgeC:~/opt# tail -1 /etc/passwd | cut -d: -f1"],
                title: "Obtención del último usuario registrado con tail y cut",
                desc: "El comando tail -1 extrae únicamente la última línea del archivo /etc/passwd, la cual corresponde al usuario creado más recientemente en el sistema; posteriormente, cut -d: -f1 utiliza los dos puntos como delimitador para aislar el primer campo, correspondiente al nombre de usuario (en este caso, admin02).",
                note: "El archivo /etc/passwd almacena las cuentas de usuario del sistema en formato de campos separados por dos puntos, siendo el primer campo el nombre de usuario y el tercero su identificador numérico (UID)."
            },
            {
                lead: "Se integra la consulta del último usuario a la secuencia general de reporte, agregando esta información al archivo capacidades.txt junto a los datos previamente recolectados.",
                cmd: [
                    "root@JorgeC:~/opt# rm capacidades.txt",
                    'root@JorgeC:~/opt# echo "----------------------------" >> capacidades.txt',
                    "root@JorgeC:~/opt# date >> capacidades.txt",
                    "root@JorgeC:~/opt# df -h | grep /dev/sda1 >>capacidades.txt",
                    'root@JorgeC:~/opt# echo "     Total Used Libre" >> capacidades.txt',
                    "root@JorgeC:~/opt# free -m | grep Mem: | cut -d' ' -f1,13,22,30 >> capacidades.txt",
                    "root@JorgeC:~/opt# tail -1 /etc/passwd | cut -d: -f1 >> capacidades.txt",
                    'root@JorgeC:~/opt# echo "----------------------------" >> capacidades.txt',
                    "root@JorgeC:~/opt# cat capacidades.txt"
                ],
                title: "Reporte final con fecha, disco, memoria y último usuario registrado",
                desc: "El archivo capacidades.txt queda conformado por cuatro bloques de información: un separador inicial, la fecha y hora, el estado del disco y la memoria, y finalmente el nombre del último usuario creado en el sistema, cerrando el reporte con un segundo separador de guiones.",
                note: "Esta versión constituye el reporte definitivo que será automatizado íntegramente dentro del script reporte01.sh en la siguiente etapa del laboratorio."
            },
            {
                lead: "Con el comando touch se crea un archivo con extensión .sh denominado reporte01, el cual posteriormente se abre con el editor de texto nano para incorporar el script.",
                cmd: ["root@JorgeC:~/opt# touch reporte01.sh", "root@JorgeC:~/opt# nano reporte01.sh"],
                title: "Creación y apertura del archivo de script reporte01.sh",
                desc: "El editor nano se abre mostrando un lienzo en blanco correspondiente al nuevo archivo reporte01.sh, listo para recibir el contenido del script que automatizará la generación del reporte de sistema previamente construido manualmente.",
                note: "La extensión .sh no es obligatoria para que un script de Bash funcione, pero constituye una convención ampliamente adoptada que facilita identificar el propósito del archivo."
            },
            {
                lead: "Se edita el archivo reporte01.sh incorporando la línea shebang (#!/bin/bash) y la secuencia completa de comandos previamente validada de forma manual.",
                cmd: [
                    "#!/bin/bash",
                    'echo "----------------------------" >> capacidades.txt',
                    "date >> capacidades.txt",
                    "df -h | grep /dev/sda1 >> capacidades.txt",
                    'echo "     Total Used Libre" >> capacidades.txt',
                    "free -m | grep Mem: | cut -d' ' -f1, 13, 22, 30 >> capacidades.txt",
                    "tail -1 /etc/passwd | cut -d: -f1 >> capacidades.txt",
                    'echo "----------------------------" >> capacidades.txt'
                ],
                title: "Contenido del script reporte01.sh",
                desc: "La línea shebang #!/bin/bash indica al sistema operativo que el archivo debe interpretarse utilizando el intérprete de comandos Bash, independientemente de cuál sea la shell activa del usuario que lo ejecute; las líneas siguientes replican exactamente la secuencia de comandos validada manualmente en las figuras anteriores.",
                note: "Es imprescindible otorgar permisos de ejecución al script (chmod +x reporte01.sh) antes de poder invocarlo directamente con ./reporte01.sh."
            },
            {
                lead: "Una vez guardado el script, se procede a ejecutarlo utilizando la notación ./ que indica al sistema que el archivo ejecutable se encuentra en el directorio actual.",
                cmd: ["root@JorgeC:~/opt# ./reporte01.sh"],
                title: "Ejecución del script reporte01.sh",
                desc: "Al ejecutarse, el script corre de manera secuencial cada uno de los comandos internos, generando nuevamente el archivo capacidades.txt con la información actualizada de fecha, disco, memoria y último usuario, sin requerir la intervención manual comando por comando.",
                note: "La notación ./ es necesaria en Linux porque, por razones de seguridad, el directorio actual no forma parte de la variable de entorno PATH por defecto."
            },
            {
                lead: "Se visualiza el resultado generado automáticamente por el script mediante el contenido actualizado del archivo capacidades.txt.",
                cmd: [
                    "----------------------------",
                    "Mon Apr 15 07:15:32 PM PDT 2024",
                    "/dev/sda1        8.9G  1.9G  6.6G  22% /",
                    "     Total Used Libre",
                    "admin02",
                    "----------------------------"
                ],
                title: "Salida generada automáticamente por reporte01.sh",
                desc: "El resultado obtenido replica fielmente el formato del reporte construido manualmente en las figuras previas, demostrando que la automatización mediante script produce resultados idénticos a la ejecución manual, pero con un considerable ahorro de tiempo y reducción de errores humanos.",
                note: "A partir de este punto, cualquier modificación futura al formato del reporte solo requerirá editar el script, en lugar de reescribir cada comando de forma individual."
            },
            {
                lead: "Se utiliza el comando crontab -e para editar la tabla de tareas programadas del usuario actual, seleccionando /bin/nano como editor por defecto ante la ausencia de una configuración previa.",
                cmd: ["root@JorgeC:~/opt# crontab -e", "no crontab for root - using an empty one", "Select an editor... Choose 1-3 [1]:"],
                title: "Edición del archivo crontab del usuario root",
                desc: "Al no existir una tabla cron previa para el usuario root, el sistema informa 'no crontab for root - using an empty one' y solicita seleccionar un editor de texto entre nano, mcedit o vim.tiny, optándose por la primera alternativa por su simplicidad de uso.",
                note: "Crontab es el mecanismo estándar en sistemas Unix/Linux para programar la ejecución periódica y desatendida de comandos o scripts, siendo la base de la automatización administrativa."
            },
            {
                lead: "Dentro del editor cron, se ingresa la línea de programación que indica al sistema ejecutar el script reporte01.sh de manera automática cada dos minutos.",
                cmd: ["*/2 * * * * /opt/reporte01.sh"],
                title: "Programación de la ejecución periódica cada 2 minutos",
                desc: "La sintaxis de crontab consta de cinco campos temporales (minuto, hora, día del mes, mes y día de la semana) seguidos del comando a ejecutar; la notación */2 en el campo de minutos indica que la tarea debe correr cada dos minutos, las veinticuatro horas del día, todos los días del mes, mes y año.",
                note: "Es indispensable especificar la ruta absoluta del script (/opt/reporte01.sh) dentro de crontab, ya que las tareas programadas se ejecutan sin un contexto de directorio de trabajo definido."
            },
            {
                lead: "Debido a que crontab ejecuta las tareas sin partir de un directorio de trabajo específico, es necesario corregir las rutas relativas dentro del propio script para que apunten a la ruta absoluta del archivo capacidades.txt.",
                cmd: [
                    "#!/bin/bash",
                    'echo "----------------------------" >> /opt/capacidades.txt',
                    "date >> /opt/capacidades.txt",
                    "df -h | grep /dev/sda1 >> /opt/capacidades.txt",
                    'echo "     Total Used Libre" >> /opt/capacidades.txt',
                    "free -m | grep Mem: | cut -d' ' -f1, 13, 22, 30 >> /opt/capacidades.txt",
                    "tail -1 /etc/passwd | cut -d: -f1 >> /opt/capacidades.txt",
                    'echo "----------------------------" >> /opt/capacidades.txt'
                ],
                title: "Corrección de rutas relativas a absolutas dentro de reporte01.sh",
                desc: "Se reemplaza cada referencia al archivo capacidades.txt por su ruta absoluta /opt/capacidades.txt, garantizando así que, independientemente del contexto desde el cual crontab dispare la ejecución del script, la información siempre se almacene en la ubicación correcta.",
                note: "Este es uno de los errores más comunes al automatizar scripts con crontab: omitir las rutas absolutas provoca que los archivos generados aparezcan en ubicaciones inesperadas o que el script falle silenciosamente."
            },
            {
                lead: "Se emplea el comando find para localizar, en todo el sistema de archivos, un archivo específico denominado error.log.",
                cmd: ["root@JorgeC:~# find / -type f -name error.log"],
                title: "Búsqueda global del archivo error.log con find",
                desc: "El comando find recorre recursivamente el árbol de directorios a partir de la raíz (/), filtrando exclusivamente archivos regulares (-type f) cuyo nombre coincida exactamente con 'error.log'; el resultado obtenido señala la ruta /var/log/apache2/error.log, correspondiente al registro de errores del servidor web Apache.",
                note: "Una búsqueda desde la raíz del sistema puede resultar considerablemente más lenta que una búsqueda acotada a un directorio específico, por lo que en entornos de producción se recomienda delimitar el alcance de find siempre que sea posible."
            },
            {
                lead: "Se utiliza grep con la opción recursiva (-r) para localizar todas las apariciones de la cadena 'torvalds' dentro de los archivos del sistema de archivos raíz.",
                cmd: ["root@JorgeC:~# grep -r torvalds /*"],
                title: "Búsqueda recursiva de la cadena 'torvalds' con grep -r",
                desc: "La búsqueda recursiva recorre archivos de configuración crítica como /etc/passwd, /etc/group, /etc/shadow, /etc/subuid y /etc/subgid, evidenciando que el usuario torvalds posee una cuenta activa en el sistema, con su respectivo UID, GID y directorio home, además de aparecer referenciado en archivos de código fuente de terceros.",
                note: "Realizar búsquedas recursivas sobre archivos sensibles como /etc/shadow debe manejarse con extrema cautela en entornos productivos, dado que dichos archivos contienen hashes de contraseñas y su exposición representa un riesgo de seguridad."
            },
            {
                lead: "Se visualiza el contenido completo del archivo /etc/passwd mediante el comando cat, con el propósito de identificar todas las cuentas de usuario existentes en el sistema.",
                cmd: ["root@JorgeC:~# cat /etc/passwd"],
                title: "Listado completo de cuentas de usuario mediante /etc/passwd",
                desc: "El archivo /etc/passwd contiene una línea por cada cuenta del sistema, estructurada en siete campos separados por dos puntos: nombre de usuario, contraseña (habitualmente representada por una x, ya que el hash real reside en /etc/shadow), UID, GID, comentario o nombre completo, directorio home y shell por defecto.",
                note: "Las cuentas con UID inferior a 1000 corresponden típicamente a servicios y procesos del sistema (daemon, bin, sys, entre otros), mientras que los UID iguales o superiores a 1000 corresponden a cuentas de usuarios reales creadas por el administrador."
            },
            {
                lead: "Se combina el comando cut con awk para extraer y filtrar de manera precisa la información de los usuarios cuyo identificador (UID) se encuentra entre 1000 y 2000.",
                cmd: ["root@JorgeC:~# cut -d: -f1,3,6 /etc/passwd | awk -F: '$2 >= 1000 && $2 <= 2000'"],
                title: "Filtrado combinado de usuarios con UID entre 1000 y 2000 mediante cut y awk",
                desc: "El comando cut selecciona primero los campos 1 (usuario), 3 (UID) y 6 (directorio home) del archivo /etc/passwd; a continuación, awk recibe dicha salida y aplica una condición lógica sobre el segundo campo resultante para filtrar únicamente aquellas líneas cuyo UID se ubique dentro del rango especificado.",
                note: "Encadenar cut y awk permite combinar la simplicidad de selección de columnas del primero con la potencia de evaluación de condiciones lógicas del segundo, patrón muy habitual en scripts de administración de sistemas."
            },
            {
                lead: "Se emplea awk de forma independiente, sin cut previo, para filtrar directamente sobre el archivo /etc/passwd aquellos usuarios cuyo UID sea igual o superior a 1000.",
                cmd: ["root@JorgeC:~# awk -F: '$3 >= 1000 {print $1,$3,$6}' /etc/passwd"],
                title: "Filtrado directo de usuarios no privilegiados con awk -F:",
                desc: "El parámetro -F: establece los dos puntos como delimitador de campos; la condición $3 >= 1000 filtra las líneas cuyo tercer campo (UID) cumpla dicha condición, e inmediatamente se imprime una selección personalizada de los campos 1, 3 y 6, correspondientes al nombre de usuario, el UID y el directorio home.",
                note: "Awk es considerablemente más versátil que la combinación cut+grep, dado que integra en un solo lenguaje de programación la capacidad de filtrar, transformar e imprimir información estructurada en una única instrucción."
            },
            {
                lead: "Se crea un nuevo script de Bash denominado menú.sh y se abre con el editor nano para construir una interfaz de texto interactiva.",
                cmd: ["root@JorgeC:~/opt# nano menú.sh"],
                title: "Apertura del editor nano para la creación de menú.sh",
                desc: "El comando nano abre el editor de texto en modo interactivo sobre el archivo menú.sh; si el archivo no existe previamente, nano lo crea automáticamente al momento de guardar los cambios por primera vez con la combinación de teclas Ctrl+O.",
                note: "Nano es uno de los editores de texto en modo consola más utilizados en la administración de sistemas Linux gracias a su curva de aprendizaje reducida en comparación con editores como vim."
            },
            {
                lead: "De forma alternativa, se confirma la creación explícita del archivo utilizando el comando touch antes de proceder a editarlo.",
                cmd: ["root@JorgeC:~/opt# touch menú.sh"],
                title: "Creación explícita del archivo menú.sh mediante touch",
                desc: "Aunque nano puede crear el archivo automáticamente al guardar, esta captura documenta la creación explícita mediante touch, práctica habitual quNe permite además establecer permisos o metadatos iniciales antes de comenzar la edición del contenido del script.",
                note: "Combinar touch y nano en secuencia constituye un patrón de trabajo ordenado que separa claramente la etapa de creación de la etapa de edición del archivo."
            },
            {
                lead: "Se desarrolla el contenido completo del script menú.sh, el cual construye un menú interactivo en la terminal con opciones para consultar espacio en disco, memoria y usuarios creados en el sistema.",
                cmd: [
                    "#!/bin/bash",
                    "clear",
                    "mostrar_menu() {",
                    '  echo "Menú de Opciones:"',
                    '  echo "1. Espacio en disco"',
                    '  echo "2. Memoria"',
                    '  echo "3. Usuarios creados"',
                    '  echo "4. Salir"',
                    "}",
                    "mostrar_espacio_disco() { df -h }",
                    "mostrar_memoria() { free -h }",
                    "mostrar_usuarios_creados() {",
                    "  cut -d: -f1,3,6 /etc/passwd | awk -F: '$2 >= 1000 && $2 <= 2000'",
                    "}",
                    'opciones="1 2 3 4"',
                    "while true; do",
                    "  mostrar_menu",
                    '  read -p "Seleccione una opción: " opcion',
                    "  case $opcion in",
                    "    1) mostrar_espacio_disco ;;",
                    "    2) mostrar_memoria ;;",
                    "    3) mostrar_usuarios_creados ;;",
                    '    4) echo "Saliendo del programa." ; break ;;',
                    '    *) echo "Opción inválida. Intente de nuevo." ;;',
                    "  esac",
                    '  read -p "Presione Enter para continuar..."',
                    "  clear",
                    "done"
                ],
                title: "Script completo del menú interactivo (menú.sh)",
                desc: "El script define una función mostrar_menu que despliega las opciones disponibles, tres funciones auxiliares que encapsulan los comandos df -h, free -h y la consulta filtrada de usuarios, y un bucle while true que mantiene el programa en ejecución hasta que el usuario seleccione la opción de salida; la instrucción case gestiona la lógica de selección según la opción ingresada por teclado mediante read -p.",
                note: "El uso de funciones dentro de un script de Bash mejora considerablemente la legibilidad y facilita el mantenimiento del código, siguiendo principios básicos de programación estructurada aplicados a la shell."
            },
            {
                lead: "Se ejecuta el script menú.sh y se selecciona la opción número 1 para visualizar el espacio disponible en disco.",
                cmd: [
                    "root@JorgeC:~/opt# ./menú.sh",
                    "Menú de Opciones:",
                    "1. Espacio en disco",
                    "2. Memoria",
                    "3. Usuarios creados",
                    "4. Salir",
                    "Seleccione una opción: 1"
                ],
                title: "Ejecución de menú.sh — opción 1: Espacio en disco",
                desc: "Al seleccionar la opción 1, el script invoca internamente la función mostrar_espacio_disco, la cual ejecuta df -h y despliega en pantalla el detalle de todos los sistemas de archivos montados, incluyendo la partición principal /dev/sda1 con un uso del 23%.",
                note: "El diseño modular del script permite que, si en el futuro se desea cambiar df -h por otra herramienta de diagnóstico, únicamente sea necesario modificar el cuerpo de la función correspondiente."
            },
            {
                lead: "Se selecciona la opción número 2 del menú interactivo para consultar el estado de la memoria del sistema.",
                cmd: ["Seleccione una opción: 2"],
                title: "Ejecución de menú.sh — opción 2: Memoria",
                desc: "La función mostrar_memoria ejecuta free -h y despliega en pantalla los valores de memoria total, usada, libre, compartida, en búfer/caché y disponible, tanto para la memoria RAM como para el espacio de intercambio (swap), en unidades legibles para el usuario.",
                note: "Observar simultáneamente los valores de RAM y swap resulta útil para diagnosticar si el sistema está recurriendo excesivamente al intercambio con disco, lo cual suele traducirse en una degradación notable del rendimiento."
            },
            {
                lead: "Finalmente, se selecciona la opción número 3 para consultar el listado de usuarios creados en el sistema con UID dentro del rango administrado.",
                cmd: [
                    "Seleccione una opción: 3",
                    "torvalds:1000:/home/torvalds",
                    "usr03:1002:/home/usr03",
                    "usr01:1001:/home/usr01",
                    "admin01:1007:/home/admin01",
                    "admin02:1008:/home/admin02"
                ],
                title: "Ejecución de menú.sh — opción 3: Usuarios creados",
                desc: "La función mostrar_usuarios_creados reutiliza la combinación previamente validada de cut y awk sobre /etc/passwd, mostrando en pantalla el nombre de usuario, el UID y el directorio home de cada cuenta cuyo identificador se encuentra entre 1000 y 2000, es decir, las cuentas de usuarios creadas manualmente durante prácticas anteriores.",
                note: "Este ejercicio demuestra cómo un mismo bloque de código validado previamente de forma manual puede reutilizarse sin modificaciones dentro de distintos scripts, favoreciendo la reutilización y el mantenimiento del código."
            }
        ]
    },
    {
        key: "b",
        title: "b) Navegación avanzada, comodines y gestión de directorios con ls, cp, mv y mkdir",
        intro: [
            "En esta sección del laboratorio se profundiza en el uso de comodines (wildcards) del intérprete de comandos Bash, aplicados principalmente al comando ls, así como en la manipulación de estructuras de directorios mediante mkdir, cp, mv y rm. Estos ejercicios permiten comprender de manera práctica los patrones de expansión de nombres de archivo (?, *, [ ], [^ ]) que resultan fundamentales para la administración eficiente del sistema de archivos en entornos Linux.",
            "Adicionalmente, se construye y manipula un árbol de directorios de prueba denominado PRUEBA, sobre el cual se ejecutan operaciones de copiado recursivo, movimiento de subárboles completos y ocultamiento de archivos, replicando escenarios habituales en tareas de administración de sistemas."
        ],
        figures: [
            { lead: "Se listan todos los archivos y subdirectorios contenidos en el directorio temporal del sistema.", cmd: ["$ ls /tmp"], title: "Listado del contenido del directorio /tmp", desc: "El directorio /tmp almacena archivos temporales generados por distintos servicios del sistema; en la captura se observan archivos de sockets privados creados por systemd para servicios como apache2, systemd-logind y systemd-timesyncd.", note: "El contenido de /tmp suele limpiarse automáticamente en cada reinicio del sistema o mediante tareas programadas de limpieza (tmpwatch o systemd-tmpfiles)." },
            { lead: "Se listan, en orden inverso alfabético, todos los archivos y directorios de /etc cuyo nombre comience con la letra t.", cmd: ["$ ls -dr /etc/t*"], title: "Listado inverso de elementos que inician con 't' en /etc", desc: "El parámetro -d evita que ls despliegue el contenido interno de los subdirectorios encontrados, mostrando únicamente el nombre del propio directorio o archivo que coincide con el patrón; el parámetro -r invierte el orden alfabético habitual, mostrando primero timezone y terminfo antes que tmpfiles.d.", note: "La combinación -d es especialmente útil cuando se desea inventariar únicamente los nombres de un conjunto de directorios sin recorrer su contenido interno." },
            { lead: "Se listan los archivos de dispositivo del directorio /dev cuyo nombre inicie con 'tty' seguido de exactamente dos caracteres adicionales.", cmd: ["$ ls /dev/tty??"], title: "Comodín de dos caracteres tras 'tty' en /dev", desc: "Cada símbolo ? representa exactamente un carácter comodín; por lo tanto, el patrón tty?? coincide con nombres como tty10, tty20, ttyS0 o tty51, pero excluye tty1 (un solo dígito) o ttyS10 (más de dos caracteres tras tty).", note: "Los archivos de dispositivo tty en Linux representan terminales virtuales y puertos seriales, siendo fundamentales para la gestión de sesiones de consola." },
            { lead: "Se listan los archivos de /dev/tty que finalizan en un dígito comprendido entre 1 y 4, sin importar la longitud del nombre.", cmd: ["$ ls /dev/tty*[1-4]"], title: "Comodín de rango numérico [1-4] como sufijo", desc: "El asterisco (*) representa cualquier cantidad de caracteres (incluyendo cero), mientras que el rango [1-4] exige que el último carácter del nombre sea uno de los dígitos 1, 2, 3 o 4; la combinación captura desde tty1 hasta tty64, siempre que el nombre termine exactamente en uno de esos dígitos.", note: "Los rangos entre corchetes ([a-z], [0-9], [1-4]) constituyen una de las herramientas de expansión de patrones más potentes de Bash para realizar selecciones precisas de archivos." },
            { lead: "Se listan todos los archivos y directorios del sistema de archivos raíz, incluyendo aquellos considerados ocultos por comenzar con un punto.", cmd: ["$ ls -a /"], title: "Listado de archivos ocultos y visibles en la raíz del sistema", desc: "El parámetro -a instruye a ls a mostrar también las entradas especiales . (directorio actual) y .. (directorio padre), así como cualquier otro archivo o carpeta cuyo nombre comience con un punto, las cuales normalmente se ocultan de la vista por convención en sistemas Unix/Linux.", note: "La convención de ocultar archivos que inician con punto no constituye una medida de seguridad, sino únicamente una preferencia estética para reducir el ruido visual en los listados de directorio." },
            { lead: "Se listan los elementos de /etc cuyo nombre NO comience con la letra t, utilizando negación de patrón entre corchetes.", cmd: ["$ ls -d /etc/[^t]*"], title: "Negación de patrón con [^t] en /etc", desc: "El símbolo de intercalación (^) dentro de los corchetes invierte la lógica de coincidencia, de modo que [^t] representa cualquier carácter que no sea la letra t; combinado con el asterisco final, el patrón selecciona todos los archivos y directorios de /etc cuyo primer carácter sea distinto de t.", note: "La negación de patrones es especialmente valiosa cuando se necesita excluir un subconjunto reducido de elementos de un listado extenso, evitando enumerar manualmente todos los casos restantes." },
            { lead: "Se listan de manera recursiva todos los archivos y subdirectorios contenidos dentro del directorio /usr.", cmd: ["$ ls -R /usr"], title: "Listado recursivo del árbol de directorios /usr", desc: "El parámetro -R (recursivo) hace que ls descienda automáticamente por cada subdirectorio encontrado, mostrando su contenido completo; en el caso de /usr, esto produce un listado extenso que abarca bin, games, include, lib, lib64, libexec, local, sbin, share y src, junto con todos sus archivos internos.", note: "Debido al volumen de información que puede generar, ls -R suele combinarse con herramientas de paginado como less o con redirección a un archivo para su posterior análisis." },
            { lead: "Se cambia el directorio de trabajo actual hacia /tmp utilizando el comando cd.", cmd: ["$ cd /tmp"], title: "Cambio de directorio de trabajo hacia /tmp", desc: "El comando cd (change directory) modifica el directorio actual de la sesión de terminal; al ejecutarse con la ruta absoluta /tmp, el prompt del sistema se actualiza reflejando la nueva ubicación (root@JorgeC:/tmp#).", note: "A diferencia de otros comandos de shell, cd es un comando interno (builtin) de Bash y no un programa externo, motivo por el cual afecta directamente el estado del proceso de la shell actual." },
            { lead: "Se verifica que el cambio de directorio se haya realizado correctamente utilizando el comando pwd.", cmd: ["$ pwd"], title: "Verificación del directorio actual con pwd", desc: "El comando pwd (print working directory) imprime en pantalla la ruta absoluta del directorio en el que se encuentra actualmente posicionada la sesión de terminal, confirmando en este caso que el usuario efectivamente se encuentra dentro de /tmp.", note: "pwd resulta especialmente útil dentro de scripts que necesitan construir rutas absolutas de manera dinámica a partir de la ubicación actual de ejecución." },
            { lead: "Se consulta el día y la hora actual del sistema utilizando un formato de salida específico mediante el parámetro + de date.", cmd: ['$ date +"%A %D - %r"  #ó simplemente DATE'], title: "Formato de fecha personalizado con date +", desc: "Los especificadores %A, %D y %r indican, respectivamente, el nombre completo del día de la semana, la fecha en formato mm/dd/aa y la hora en formato de 12 horas con indicador AM/PM; sin embargo, la captura muestra el resultado de ejecutar simplemente date, que produce un formato similar por defecto.", note: "El comando date admite decenas de especificadores de formato adicionales (%Y, %m, %d, %H, %M, %S, entre otros) que permiten construir cadenas de fecha completamente personalizadas para reportes o nombres de archivo." },
            { lead: "Con una única instrucción, se posiciona la sesión de terminal directamente en el directorio HOME del usuario actual.", cmd: ["$ cd /HOME"], title: "Posicionamiento directo en el directorio $HOME", desc: "Al tratarse de una máquina Linux donde el sistema de archivos no distingue mayúsculas de minúsculas de forma estricta en algunos contextos de shell, el comando cd /HOME resuelve efectivamente hacia /home; en la práctica, la forma recomendada sería utilizar simplemente cd o cd $HOME.", note: "La variable de entorno $HOME almacena la ruta del directorio personal del usuario actualmente autenticado y es ampliamente utilizada en scripts para referirse de forma portable a dicha ubicación." },
            { lead: "Se verifica que efectivamente la sesión se encuentra posicionada en el directorio /home.", cmd: ["$ pwd"], title: "Confirmación del directorio /home mediante pwd", desc: "La salida del comando pwd confirma que el cambio de directorio fue exitoso, mostrando la ruta /home como directorio de trabajo actual de la sesión de terminal.", note: "Verificar sistemáticamente el resultado de un cambio de directorio mediante pwd es una práctica recomendable, especialmente dentro de scripts donde un error de ruta podría desencadenar operaciones destructivas en ubicaciones no deseadas." },
            { lead: "Se listan todos los ficheros del directorio HOME mostrando adicionalmente su número de inodo mediante el parámetro -i.", cmd: ["$ ls -i"], title: "Listado de archivos con número de inodo (ls -i)", desc: "Cada archivo en un sistema de archivos tipo ext4 posee un número de inodo único que almacena sus metadatos (permisos, propietario, tamaño, ubicación física de los bloques de datos); la salida muestra, por ejemplo, el archivo archivos.tar.gz asociado al inodo 130278.", note: "El número de inodo es especialmente relevante para distinguir entre archivos con el mismo nombre en distintos sistemas de archivos o para detectar enlaces duros (hard links), los cuales comparten el mismo número de inodo." },
            {
                lead: "Se construye una estructura de prueba compuesta por archivos ocultos, archivos regulares y directorios dentro de una carpeta denominada PRUEBA, para posteriormente eliminar todo su contenido visible.", cmd: [
                    "mkdir PRUEBA",
                    "touch PRUEBA/{.f_hidden1,.f_hidden2,.f_hidden3}",
                    "touch PRUEBA/{file1,file2,file3}",
                    "mkdir PRUEBA/{dir1,dir2,dir3}",
                    "$ rm -rf PRUEBA/*"
                ], title: "Creación de estructura de prueba y eliminación de contenido visible", desc: "La expansión de llaves {a,b,c} permite crear múltiples archivos o directorios en una sola invocación de touch o mkdir; posteriormente, rm -rf PRUEBA/* elimina de forma recursiva y forzada todos los elementos visibles dentro de PRUEBA, pero respeta los archivos ocultos (.f_hidden1, .f_hidden2, .f_hidden3), ya que el asterisco no expande por defecto nombres que comienzan con punto.", note: "Este comportamiento del asterisco respecto a los archivos ocultos constituye una fuente frecuente de confusión y errores en scripts de limpieza automatizada, por lo que conviene combinarlo explícitamente con patrones que incluyan el punto inicial cuando se desea eliminar también los archivos ocultos."
            },
            {
                lead: "Se construye una jerarquía de directorios anidados dentro de PRUEBA utilizando expansión de llaves para crear múltiples niveles en una sola instrucción de mkdir.", cmd: [
                    "$ mkdir PRUEBA/{dir1,\\ dir1/dir11,\\ dir2,\\",
                    "dir3,\\ dir3/dir31,\\ dir3/dir31/dir311,\\ dir3/dir31/dir312}"
                ], title: "Creación de jerarquía anidada de directorios con mkdir y llaves", desc: "La instrucción crea de manera simultánea los directorios dir1, dir2 y dir3 en el primer nivel; dentro de dir1 se crea dir11; dentro de dir3 se crea dir31; y dentro de dir31 se crean, a su vez, dir311 y dir312, conformando así una estructura jerárquica de tres niveles de profundidad en una única línea de comando.", note: "mkdir por sí solo no crea directorios intermedios inexistentes salvo que se emplee el parámetro -p; en este caso, el orden de los elementos dentro de la expansión de llaves garantiza que cada directorio padre exista antes de intentar crear su respectivo subdirectorio."
            },
            { lead: "Se copia el archivo de mensaje del día del sistema (/etc/motd) hacia el directorio PRUEBA, renombrándolo como 'mensaje'.", cmd: ["$ cp /etc/motd ./PRUEBA"], title: "Copia de /etc/motd hacia el directorio PRUEBA", desc: "El comando cp duplica el contenido del archivo /etc/motd (message of the day, mostrado habitualmente al iniciar sesión) dentro del directorio PRUEBA, conservando su nombre original motd, el cual servirá como archivo base para las operaciones de copiado múltiple de las siguientes figuras.", note: "El archivo /etc/motd es ampliamente utilizado por los administradores de sistemas para comunicar avisos, mantenimientos programados o políticas de uso a todos los usuarios que inicien sesión en el servidor." },
            {
                lead: "Se copia el archivo mensaje hacia los tres subdirectorios dir1, dir2 y dir3, encadenando los tres comandos mediante el operador lógico &&.", cmd: [
                    "$ cd PRUEBA",
                    "$ cp -r mensaje dir1/mensaje && cp -r mensaje dir2/mensaje && cp -r mensaje dir3/mensaje"
                ], title: "Copia encadenada del archivo mensaje en dir1, dir2 y dir3", desc: "El operador && ejecuta el siguiente comando únicamente si el comando anterior finalizó exitosamente (código de salida 0), garantizando así que la copia hacia dir2 no se intente si la copia hacia dir1 falló, lo cual constituye una práctica defensiva habitual en la escritura de secuencias de comandos.", note: "Encadenar comandos con && es preferible a separarlos simplemente con punto y coma (;) cuando la ejecución de un paso depende del éxito del paso anterior, evitando así la propagación de errores silenciosos."
            },
            { lead: "Se verifica, mediante un único comando recursivo, que la copia del archivo mensaje se haya realizado correctamente en cada uno de los subdirectorios.", cmd: ["$ ls -R PRUEBA"], title: "Verificación recursiva de la estructura PRUEBA con ls -R", desc: "El listado recursivo confirma la presencia del archivo mensaje dentro de dir1, dir2 y dir3, además de mostrar la jerarquía completa de subdirectorios creados previamente (dir11, dir31, dir311 y dir312), permitiendo validar en una sola operación el resultado de múltiples comandos ejecutados con anterioridad.", note: "Verificar el resultado de operaciones de copiado masivo mediante un listado recursivo es una práctica recomendable antes de proceder con operaciones adicionales que dependan de dicha estructura." },
            { lead: "Se copian, de forma recursiva, todos los archivos del directorio de scripts de inicio rc0.d de /etc hacia el subdirectorio dir3.", cmd: ["$ cp -r /etc/rc0.d dir3"], title: "Copia recursiva del directorio /etc/rc0.d hacia dir3", desc: "El parámetro -r (recursivo) es indispensable cuando el origen de la copia es un directorio, ya que cp por defecto solo opera sobre archivos individuales; en este caso, se replica íntegramente el directorio rc0.d, históricamente utilizado para gestionar los scripts de apagado del sistema (runlevel 0), dentro de la estructura de prueba.", note: "En sistemas Debian modernos basados en systemd, los directorios rcX.d han sido en gran medida sustituidos funcionalmente por unidades de servicio (.service), aunque se mantienen por compatibilidad con scripts heredados de SysVinit." },
            { lead: "Se copian hacia dir311 aquellos archivos del directorio /bin cuyo nombre tenga exactamente cuatro caracteres y cuya segunda letra sea la letra 'a'.", cmd: ["$ cp -r /bin/?a?? PRUEBA/dir3/dir31/dir311"], title: "Copia selectiva con comodín posicional (segunda letra 'a', cuatro caracteres)", desc: "El patrón ?a?? combina cuatro comodines de un solo carácter cada uno, fijando explícitamente la letra 'a' en la segunda posición del nombre del archivo; esto permite realizar una selección extremadamente precisa dentro de un directorio con cientos de binarios del sistema, sin necesidad de enumerar manualmente cada nombre coincidente.", note: "La combinación de comodines posicionales (?) con letras fijas constituye una técnica avanzada de filtrado de nombres de archivo, particularmente útil quNe cuando se buscan patrones de nomenclatura muy específicos." },
            { lead: "Se mueve el directorio dir31, junto con todo su contenido y subdirectorios, para ubicarlo dentro de dir2.", cmd: ["$ mv PRUEBA/dir3/dir31 PRUEBA/dir2"], title: "Reubicación del subárbol dir31 dentro de dir2 con mv", desc: "A diferencia de cp, el comando mv no requiere el parámetro -r para operar sobre directorios completos, ya que mover un directorio no implica duplicar físicamente su contenido en el disco (en el mismo sistema de archivos), sino únicamente actualizar las referencias en la estructura de directorios.", note: "Cuando el origen y el destino de mv se encuentran en el mismo sistema de archivos, la operación es prácticamente instantánea; si se encuentran en sistemas de archivos distintos, mv realiza internamente una copia seguida de una eliminación del origen." },
            { lead: "Se listan, de manera recursiva, todos los archivos ordinarios contenidos dentro del directorio HOME y de todos sus subdirectorios.", cmd: ["$ ls -R $HOME"], title: "Listado recursivo completo del directorio $HOME", desc: "Esta captura ofrece una vista panorámica del árbol de directorios completo bajo /root, evidenciando la coexistencia de los archivos y carpetas generados durante las prácticas anteriores (Documentos4, examen.tar.gz, opt y plantillas descargadas) junto con la estructura de carpetas creada en ejercicios previos del curso.", note: "Un listado recursivo tan extenso resulta más útil cuando se redirige hacia un archivo (ls -R $HOME > inventario.txt) para su posterior análisis o comparación entre distintos momentos en el tiempo." },
            { lead: "Se oculta el archivo mensaje ubicado dentro de dir3, renombrándolo con un punto inicial mediante el comando mv.", cmd: ["$ mv PRUEBA/dir3/mensaje PRUEBA/dir3/.mensaje"], title: "Ocultamiento de un archivo mediante renombrado con punto inicial", desc: "En sistemas Unix/Linux, un archivo se considera oculto simplemente si su nombre comienza con un punto (.); mv se utiliza aquí no para reubicar el archivo, sino para renombrarlo, transformando 'mensaje' en '.mensaje', el cual dejará de aparecer en los listados de ls sin el parámetro -a.", note: "A diferencia de otros sistemas operativos, ocultar un archivo en Linux mediante esta convención no otorga ningún nivel adicional de protección o seguridad, únicamente evita su despliegue visual por defecto." },
            { lead: "Se eliminan de forma recursiva y forzada todos los archivos y subdirectorios contenidos en dir1, incluyendo el propio directorio dir1.", cmd: ["$ rm -rf PRUEBA/dir1"], title: "Eliminación recursiva y forzada del directorio dir1", desc: "El parámetro -r permite eliminar directorios junto con todo su contenido interno de forma recursiva, mientras que -f (force) suprime cualquier mensaje de confirmación y evita que el comando falle si alguno de los archivos no existe; la combinación rm -rf es una de las más potentes y, a la vez, más riesgosas del intérprete de comandos.", note: "El uso de rm -rf debe realizarse con extrema precaución, verificando siempre la ruta objetivo antes de la ejecución, dado que no existe una papelera de reciclaje que permita recuperar los archivos eliminados de esta manera." },
            { lead: "Se intenta copiar hacia dir312 los ficheros del directorio /dev cuyo nombre inicie con 't', tenga cinco caracteres en total y finalice en una letra comprendida entre la 'a' y la 'b'.", cmd: ["$ ls /dev/t???[a*b]"], title: "Patrón combinado de comodines sin coincidencias en /dev", desc: "El patrón solicitado no produce coincidencias en el sistema, arrojando el mensaje 'ls: cannot access 't???[ab]*': No such file or directory', lo cual evidencia que, si bien la sintaxis del patrón es válida, no existe en /dev ningún archivo cuyo nombre cumpla simultáneamente con todas las condiciones especificadas.", note: "Un resultado sin coincidencias no implica necesariamente un error de sintaxis; con frecuencia refleja simplemente la ausencia de archivos que satisfagan el patrón solicitado dentro del directorio consultado." },
            { lead: "Finalmente, se reubica el directorio dir312 (previamente ubicado dentro de dir2/dir31) directamente bajo dir3.", cmd: ["$ mv PRUEBA/dir2/dir31/dir312 PRUEBA/dir3"], title: "Reubicación final del subdirectorio dir312 bajo dir3", desc: "Esta operación de movimiento cierra el conjunto de ejercicios de manipulación de directorios, reorganizando nuevamente la estructura de PRUEBA y demostrando la flexibilidad del comando mv para reestructurar árboles de directorios completos sin necesidad de recrear su contenido desde cero.", note: "La capacidad de reorganizar estructuras de directorios de forma no destructiva mediante mv resulta fundamental en tareas de mantenimiento, migración de datos y reorganización de sistemas de archivos en producción." }
        ]
    },
    {
        key: "c",
        title: "c) Ejercicios con el comando echo y sus secuencias de escape",
        intro: [
            "El comando echo constituye una de las herramientas más versátiles del intérprete de comandos Bash, empleado tanto para desplegar mensajes en pantalla como para redirigir texto hacia archivos. En esta sección se exploran de manera sistemática las principales secuencias de escape admitidas por echo al invocarse con el parámetro -e, entre ellas \\n (salto de línea), \\t (tabulación horizontal), \\v (tabulación vertical), \\r (retorno de carro), \\b (retroceso), \\c (supresión de salida posterior) y \\a (alerta sonora).",
            "Asimismo, se examina el comportamiento de echo frente a la expansión de comodines del sistema de archivos y su empleo como mecanismo de creación y anexado de contenido en archivos de texto mediante los operadores de redirección > y >>."
        ],
        figures: [
            { lead: "Se ejecuta la sintaxis básica del comando echo, sin parámetros adicionales, para desplegar una cadena de texto simple en pantalla.", cmd: ["echo Hoy seguiremos en pandemia. A seguir cuidándonos"], title: "Sintaxis básica del comando echo sin comillas", desc: "Al no emplear comillas, Bash interpreta cada palabra separada por espacios como un argumento independiente que echo concatena nuevamente con un único espacio antes de imprimir la línea completa en la salida estándar.", note: "Aunque técnicamente funcional, omitir las comillas en echo puede provocar comportamientos inesperados cuando el texto contiene comodines o caracteres especiales que Bash intentaría expandir previamente." },
            { lead: "Se declara una variable de shell denominada Y con el valor 14, y posteriormente se despliega su contenido utilizando echo junto a la sintaxis de expansión $Y.", cmd: ["$ Y=14", '$ echo "El valor de la variable Y es: $Y"'], title: "Declaración de variables y expansión con echo", desc: "En Bash, la asignación de variables no admite espacios alrededor del signo igual (Y=14 es válido, Y = 14 generaría un error); al emplear el símbolo $ antes del nombre de la variable dentro de una cadena entre comillas dobles, Bash sustituye dicha referencia por su valor actual antes de imprimir el resultado.", note: "Las comillas dobles permiten la expansión de variables dentro de la cadena, mientras que las comillas simples tratan el contenido de manera completamente literal, sin realizar ningún tipo de sustitución." },
            { lead: "Se emplea el parámetro -e junto a la secuencia de escape \\b (backspace) tras cada palabra, con el objetivo de eliminar los espacios intermedios entre ellas.", cmd: ["echo -e \"Hoy \\bseguiremos \\ben \\bpandemia \\b. \\b aseguir \\b cuidándonos\""], title: "Eliminación de espacios intermedios mediante \\b (retroceso)", desc: "Cada secuencia \\b provoca que el cursor retroceda una posición antes de continuar imprimiendo, eliminando efectivamente el espacio en blanco previamente insertado entre palabras y produciendo como resultado una cadena de texto continua sin separaciones visibles.", note: "El comportamiento de \\b depende del terminal utilizado; en algunos emuladores de terminal modernos su efecto puede no visualizarse exactamente igual que en una consola de texto tradicional." },
            { lead: "Se utiliza la secuencia de escape \\n para insertar saltos de línea explícitos dentro de una única invocación de echo.", cmd: ["echo -e \"Hoy \\nseguiremos \\nen \\npandemia. \\n a seguir \\n cuidándonos\""], title: "Inserción de saltos de línea mediante \\n", desc: "Cada aparición de \\n dentro de la cadena provoca que el texto subsiguiente se imprima en una nueva línea, transformando una oración continua en un bloque de texto distribuido verticalmente, palabra por palabra o frase por frase según la posición de la secuencia de escape.", note: "La secuencia \\n es equivalente, en su efecto visual, a presionar la tecla Enter dentro del texto, y constituye una de las secuencias de escape más utilizadas en scripts de generación de reportes." },
            { lead: "Se utiliza la secuencia de escape \\t para introducir tabulaciones horizontales entre las distintas palabras de la cadena.", cmd: ["echo -e \"Hoy \\tseguiremos \\ten \\tpandemia. \\t a seguir \\t cuidándonos\""], title: "Inserción de tabulaciones horizontales mediante \\t", desc: "La secuencia \\t inserta un carácter de tabulación horizontal (equivalente a presionar la tecla Tab), desplazando el texto siguiente hacia la siguiente posición de tabulación predefinida por el terminal, generando una alineación en columnas del texto resultante.", note: "El ancho visual de una tabulación puede variar según la configuración del terminal, típicamente entre 4 y 8 espacios de ancho." },
            { lead: "Se combinan en una misma instrucción las secuencias de escape \\n (salto de línea) y \\t (tabulación), generando un bloque de texto estructurado con sangría.", cmd: ["echo -e \"Hoy \\n \\tseguiremos \\n \\ten \\n \\tpandemia. \\n \\t a seguir \\n\\t cuidándonos\""], title: "Combinación de \\n y \\t para texto estructurado con sangría", desc: "Al combinar ambas secuencias, cada nueva línea generada por \\n comienza además desplazada hacia la derecha gracias a la tabulación \\t, produciendo un efecto visual similar al de una lista jerárquica o un bloque de texto con formato de cita, útil para mejorar la legibilidad de salidas extensas en la terminal.", note: "Esta técnica de combinación de secuencias de escape es frecuentemente empleada en scripts que generan menús o reportes formateados directamente en la terminal, sin depender de herramientas externas de formato." },
            { lead: "Se emplea la secuencia de escape \\v (tabulación vertical) para desplazar el texto subsiguiente hacia abajo, manteniendo la posición horizontal.", cmd: ["echo -e \"Hoy \\vseguiremos \\ven \\vpandemia. \\v a \\vseguir \\v cuidándonos\""], title: "Inserción de tabulaciones verticales mediante \\v", desc: "A diferencia de \\n, que retorna el cursor al inicio de la nueva línea, \\v desplaza el cursor verticalmente hacia abajo conservando la posición horizontal previa, generando un efecto de texto escalonado o en diagonal descendente según el emulador de terminal empleado.", note: "El comportamiento exacto de \\v puede diferir considerablemente entre distintos emuladores de terminal, siendo una de las secuencias de escape menos utilizadas en scripts de uso cotidiano." },
            { lead: "Se combinan las secuencias \\n y \\v en una misma cadena para producir un efecto conjunto de salto de línea y tabulación vertical escalonada.", cmd: ["echo -e \"Hoy \\n \\vseguiremos \\n \\ven \\n \\vpandemia. \\n \\v a seguir \\n \\v cuidándonos\""], title: "Combinación de \\n y \\v para texto escalonado", desc: "Esta combinación demuestra cómo múltiples secuencias de escape pueden coexistir dentro de una misma invocación de echo -e, cada una aportando su propio comportamiento de desplazamiento del cursor y generando, en conjunto, patrones de texto visualmente más complejos que los obtenidos con una única secuencia de escape.", note: "Aunque estas combinaciones tienen fines principalmente didácticos, ilustran la flexibilidad del comando echo como herramienta de formateo básico de texto en entornos de terminal." },
            { lead: "Se emplea la secuencia de escape \\r (retorno de carro) al inicio de la cadena de texto.", cmd: ['echo -e "\\r Hoy seguiremos en pandemia. a seguir cuidándonos"'], title: "Uso de \\r (retorno de carro)", desc: "La secuencia \\r retorna el cursor al inicio de la línea actual sin avanzar a una nueva línea, comportamiento heredado de las máquinas de escribir e impresoras de línea; en terminales modernos de una sola línea de salida, su efecto suele ser prácticamente imperceptible al no existir texto previo que sobrescribir.", note: "La secuencia \\r se emplea comúnmente en scripts que muestran barras de progreso o contadores que se actualizan dinámicamente sobre la misma línea de la terminal." },
            { lead: "Se utiliza la secuencia de escape \\c para suprimir cualquier salida adicional posterior al punto donde esta aparece, incluyendo el salto de línea final.", cmd: ["echo -e \"Hoy seguiremos \\cen pandemia. a seguir cuidándonos\""], title: "Supresión de salida posterior mediante \\c", desc: "Al encontrar la secuencia \\c, echo detiene inmediatamente el procesamiento del resto de la cadena, de modo que únicamente se imprime 'Hoy seguiremos' seguido directamente del prompt de la siguiente línea de comandos, sin salto de línea final ni el resto del texto original.", note: "La secuencia \\c resulta útil cuando se desea construir una salida parcial que posteriormente será completada por otra instrucción echo dentro del mismo renglón de la terminal." },
            { lead: "Se utiliza el parámetro -n de echo, el cual omite el salto de línea final que normalmente se añade al término de la salida.", cmd: ['echo -n "Hoy seguiremos en pandemia. a seguir cuidándonos"'], title: "Omisión del salto de línea final con el parámetro -n", desc: "Por defecto, echo agrega automáticamente un carácter de nueva línea al final de su salida; el parámetro -n suprime dicho comportamiento, lo cual provoca que el prompt de la siguiente instrucción (root@JorgeC:/#) aparezca inmediatamente a continuación del texto impreso, sin salto de línea intermedio.", note: "El parámetro -n es ampliamente utilizado en scripts que necesitan concatenar múltiples salidas de echo en una misma línea, por ejemplo al construir barras de progreso o mensajes de estado incrementales." },
            { lead: "Se emplea la secuencia de escape \\a, correspondiente a una alerta sonora (bell character), combinada con el parámetro -e.", cmd: ["echo -e \"Hoy seguiremos en pandemia. a \\aseguir cuidándonos\""], title: "Emisión de una alerta sonora mediante \\a", desc: "El carácter de campana (\\a o BEL, ASCII 7) provoca, dependiendo de la configuración del terminal utilizado, la emisión de un sonido audible o un parpadeo visual como notificación; en la práctica documentada, la salida en pantalla muestra el texto de forma continua, evidenciando que el efecto sonoro no se refleja en la captura de pantalla.", note: "El uso del carácter de alerta sonora ha caído en desuso en entornos gráficos modernos, donde frecuentemente se encuentra deshabilitado o sustituido por una notificación visual silenciosa." },
            { lead: "Se emplea echo sin argumentos de texto, acompañado únicamente del comodín asterisco (*), replicando de forma rudimentaria el comportamiento del comando ls.", cmd: ["echo *"], title: "Expansión de comodines mediante echo * (similar a ls)", desc: "Antes de que echo reciba sus argumentos, Bash expande el asterisco (*) sustituyéndolo por la lista de todos los archivos y directorios presentes en el directorio actual; el resultado es, por tanto, indistinguible en esencia de un listado básico obtenido con ls, aunque desplegado en una sola línea separada por espacios.", note: "Esta demostración evidencia que la expansión de comodines es responsabilidad de la propia shell (Bash) y no del comando que recibe los argumentos ya expandidos, principio fundamental para comprender el comportamiento de cualquier comando frente a patrones como * o ?." },
            { lead: "Se combina echo con el comodín *.txt para filtrar y mostrar únicamente los archivos de texto plano del directorio actual.", cmd: ["echo *.txt"], title: "Filtrado de archivos por extensión mediante echo *.txt", desc: "Al no existir archivos con extensión .txt en el directorio de ejecución en ese momento, Bash no encuentra coincidencias para el patrón y, por comportamiento por defecto, lo deja sin expandir, imprimiendo literalmente la cadena '*.txt' tal como fue escrita.", note: "Este comportamiento por defecto (dejar el patrón sin expandir cuando no hay coincidencias) puede modificarse activando la opción nullglob de Bash, la cual haría que el patrón se sustituya por una cadena vacía en lugar de mostrarse literalmente." },
            { lead: "Se emplea echo junto al operador de redirección simple (>) para generar un nuevo archivo de texto a partir de la salida estándar del comando.", cmd: ['echo "prueba de fichero" > prueba.txt'], title: "Creación de un archivo mediante redirección estándar (>)", desc: "El operador > redirige la salida estándar de echo hacia el archivo prueba.txt; si el archivo no existe, se crea automáticamente, y si ya existe, su contenido previo es completamente sobrescrito por el nuevo texto proporcionado.", note: "A diferencia del operador >>, el operador > simple debe emplearse con precaución, dado que sobrescribe sin advertencia previa cualquier contenido existente en el archivo de destino." },
            { lead: "Se añade una segunda línea de texto al archivo prueba.txt utilizando el operador de anexado (>>), preservando el contenido previamente escrito.", cmd: ['echo "segunda prueba de fichero" >> prueba.txt'], title: "Anexado de contenido a un archivo existente mediante >>", desc: "A diferencia del operador > empleado en la figura anterior, el operador >> añade el nuevo texto al final del archivo sin eliminar su contenido previo, de modo que prueba.txt queda conformado finalmente por dos líneas: 'prueba de fichero' y 'segunda prueba de fichero'.", note: "Esta pareja de operadores (> para crear/sobrescribir y >> para anexar) constituye la base de prácticamente todos los mecanismos de generación de bitácoras y reportes automatizados estudiados a lo largo de este laboratorio." }
        ]
    },
    {
        key: "d",
        title: "d) Implementación de un sitio web mediante servidores Apache replicados",
        intro: [
            "En la etapa final del laboratorio se aborda un caso de uso integral de administración de sistemas: la implementación de un sitio web estático sobre dos servidores Linux Debian clonados a partir de una misma máquina virtual base, configurados en red bajo el modo No-NAT para permitir la comunicación directa entre ellos.",
            "El procedimiento abarca la actualización de repositorios e instalación del servidor web Apache2, la configuración del servicio SSH para permitir autenticación remota del usuario root, la instalación de utilidades de automatización como sshpass, la transferencia de archivos mediante WinSCP, la descarga de plantillas HTML/CSS gratuitas desde repositorios públicos y, finalmente, la verificación del correcto despliegue del sitio accediendo a la dirección IP del servidor desde un navegador web."
        ],
        figures: [
            { lead: "Tras clonar la máquina virtual en modo No-NAT, se actualiza el índice de paquetes disponibles y se instala el servidor web Apache2 en ambos servidores.", cmd: ["root@JorgeC:~# apt-get update", "root@JorgeC:~# apt-get install apache2"], title: "Actualización de repositorios e instalación de Apache2", desc: "El comando apt-get update sincroniza la lista local de paquetes disponibles con los repositorios oficiales de Debian (security y updates), mientras que apt-get install apache2 descarga e instala el servidor web junto con sus dependencias (apache2-bin, apache2-data, apache2-utils), solicitando confirmación previa mediante la pregunta '¿Desea continuar? [Y/n]'.", note: "Mantener actualizados los repositorios antes de instalar cualquier paquete garantiza que se obtenga la versión más reciente disponible y compatible con las dependencias del sistema, reduciendo el riesgo de conflictos de versiones." },
            { lead: "Se accede al archivo de configuración del servicio SSH en ambos servidores, verificando que la directiva PermitRootLogin se encuentre habilitada con el valor 'yes'.", cmd: ["nano /etc/ssh/sshd_config"], title: "Habilitación de PermitRootLogin en la configuración de SSH", desc: "La directiva PermitRootLogin controla si el usuario root puede autenticarse directamente mediante SSH; por defecto, muchas distribuciones restringen o deshabilitan esta opción por razones de seguridad, siendo necesario modificarla explícitamente a 'yes' para permitir la conexión remota entre ambos servidores utilizando dicha cuenta durante la práctica.", note: "En entornos de producción, habilitar el acceso root directo por SSH se considera una práctica de seguridad desaconsejada; se recomienda en su lugar emplear cuentas de usuario con privilegios elevados mediante sudo." },
            { lead: "Una vez modificado el archivo de configuración, se reinicia el servicio sshd en ambos servidores para que los cambios surtan efecto.", cmd: ["root@JorgeC:~# systemctl restart sshd"], title: "Reinicio del servicio SSH mediante systemctl", desc: "El comando systemctl restart sshd detiene y vuelve a iniciar el demonio del servicio SSH, aplicando de esta manera la nueva configuración de PermitRootLogin sin necesidad de reiniciar por completo el sistema operativo.", note: "Systemctl es la herramienta estándar de gestión de servicios en sistemas basados en systemd, sustituyendo a los antiguos scripts de inicio SysVinit para el control de arranque, parada y reinicio de servicios." },
            { lead: "Se instala la utilidad sshpass en ambos servidores, herramienta que permite proporcionar contraseñas de forma no interactiva a comandos SSH.", cmd: ["root@JorgeC:~# apt-get install sshpass"], title: "Instalación de sshpass para autenticación SSH no interactiva", desc: "sshpass es una pequeña utilidad que permite automatizar la introducción de la contraseña en conexiones SSH dentro de scripts, evitando la intervención manual del usuario en cada conexión; su instalación descarga un único paquete de reducido tamaño desde los repositorios oficiales de Debian.", note: "Aunque sshpass resulta extremadamente útil para la automatización, su uso en entornos de producción se desaconseja frente a alternativas más seguras basadas en autenticación por claves públicas/privadas (SSH keys)." },
            { lead: "Se elimina de forma recursiva y forzada un sitio web previamente configurado, ubicado en /var/www/html/sitio1, con el propósito de dejar el entorno limpio antes de desplegar la nueva plantilla.", cmd: ["root@JorgeC:~# rm -f -r /var/www/html/sitio1"], title: "Eliminación del sitio web previo /var/www/html/sitio1", desc: "El directorio /var/www/html constituye la raíz de documentos por defecto del servidor Apache en Debian; eliminar el subdirectorio sitio1 permite despejar el espacio necesario para la nueva plantilla que se descargará y desplegará en las siguientes figuras, evitando conflictos de nombres o contenido residual de prácticas anteriores.", note: "Es fundamental verificar la ruta exacta antes de ejecutar comandos rm -rf sobre el árbol de directorios del servidor web, dado que una eliminación accidental de /var/www/html completo dejaría el servidor sin contenido alguno que servir." },
            { lead: "Se reinician ambos servidores utilizando el comando shutdown con la opción de reinicio, aplicando así de forma definitiva todos los cambios de configuración realizados hasta el momento.", cmd: ["root@JorgeC:~# shutdown -r now"], title: "Reinicio completo de ambos servidores con shutdown -r now", desc: "El parámetro -r indica a shutdown que, tras completar el apagado ordenado del sistema, debe proceder inmediatamente a reiniciarlo; el argumento now especifica que la acción debe ejecutarse de inmediato, sin ningún período de espera adicional.", note: "Un reinicio completo del sistema, aunque no siempre estrictamente necesario tras cambios de configuración de servicios individuales, garantiza que absolutamente todos los procesos y servicios carguen limpiamente la nueva configuración aplicada." },
            {
                lead: "Desde el servidor01 se establece una conexión SSH hacia el servidor02 utilizando el usuario root, y viceversa, verificando así la comunicación bidireccional entre ambas máquinas virtuales.", cmd: [
                    "root@JorgeC:~# ssh -l root 172.30.107.107",
                    "root@172.30.107.107's password:",
                    "root@JorgeC:~# ssh -l root 172.30.107.57",
                    "The authenticity of host '172.30.107.57 (172.30.107.57)' can't be established.",
                    "Are you sure you want to continue connecting (yes/no/[fingerprint])? yes"
                ], title: "Conexión SSH bidireccional entre servidor01 y servidor02", desc: "La primera conexión hacia el servidor02 (172.30.107.107) se establece sin inconvenientes al ya contar con una clave de host conocida previamente; en la conexión inversa hacia el servidor01 (172.30.107.57), el sistema advierte que la autenticidad del host no puede establecerse y solicita confirmación explícita antes de añadir la huella digital (fingerprint) ED25519 al archivo known_hosts.", note: "La verificación de la huella digital del host en la primera conexión SSH constituye un mecanismo de seguridad que previene ataques de intermediario (man-in-the-middle), por lo que dicha huella nunca debe aceptarse sin antes verificar su origen legítimo."
            },
            {
                lead: "Tras completar la verificación mutua de conectividad, se cierra la sesión remota en ambos servidores utilizando el comando exit.", cmd: [
                    "root@JorgeC:~# exit",
                    "logout",
                    "Connection to 172.30.107.107 closed.",
                    "root@JorgeC:~# exit",
                    "logout",
                    "Connection to 172.30.107.57 closed."
                ], title: "Cierre de las sesiones remotas SSH mediante exit", desc: "El comando exit finaliza la sesión de shell remota establecida a través de SSH, devolviendo el control a la sesión local original; el sistema confirma explícitamente el cierre de cada conexión mediante el mensaje 'Connection to <IP> closed.'", note: "Cerrar adecuadamente las sesiones remotas al finalizar su uso constituye una buena práctica de seguridad que evita dejar sesiones abiertas innecesariamente expuestas en caso de acceso no autorizado a la terminal local."
            },
            { lead: "Se procede a descargar una plantilla HTML/CSS gratuita para cada servidor desde el sitio oficial de TemplateMo, repositorio ampliamente utilizado para obtener diseños web listos para usar.", cmd: ["# Descarga desde https://templatemo.com — 588+ plantillas HTML/CSS gratuitas"], title: "Descarga de plantillas HTML/CSS desde el sitio oficial TemplateMo", desc: "TemplateMo es un repositorio en línea que ofrece más de quinientas ochenta plantillas de sitio web responsivas, construidas sobre HTML5 y el framework CSS Bootstrap, todas ellas disponibles de forma completamente gratuita para su descarga y personalización.", note: "Al utilizar plantillas de terceros, incluso aquellas de uso gratuito, es recomendable revisar sus términos de licencia para conocer las condiciones bajo las cuales pueden modificarse, redistribuirse o emplearse en proyectos comerciales." },
            { lead: "Se seleccionan dos plantillas distintas, una para cada servidor: 'First Portfolio', un diseño de una sola página orientado a portafolios personales, y 'Lugx Gaming', una plantilla temática para sitios de videojuegos.", cmd: ["# First Portfolio — plantilla de portafolio personal", "# Lugx Gaming — plantilla temática de videojuegos (Bootstrap 5)"], title: "Selección de las plantillas First Portfolio y Lugx Gaming", desc: "First Portfolio corresponde a una plantilla de diseño minimalista en tonos verde, morado y blanco, pensada para presentar un perfil profesional individual; Lugx Gaming, en contraste, se construye sobre la biblioteca Bootstrap versión 5.3.0 alpha 1, con una paleta de colores predominantemente azul orientada a comunidades y tiendas de videojuegos.", note: "La elección de una plantilla adecuada al propósito del sitio web (portafolio personal, tienda en línea, blog corporativo, entre otros) es un paso de diseño que antecede a cualquier consideración puramente técnica de implementación." },
            { lead: "Una vez descargado el archivo comprimido de la plantilla, se transfiere hacia la máquina virtual mediante el cliente gráfico WinSCP, el cual utiliza el protocolo SFTP para la transferencia segura de archivos.", cmd: ["# WinSCP — Transferencia SFTP hacia root@172.30.107.57"], title: "Transferencia del archivo comprimido mediante WinSCP", desc: "WinSCP es un cliente gráfico para Windows que permite explorar y transferir archivos entre el equipo local y un servidor remoto mediante los protocolos SFTP, SCP o FTP; en la captura se observa el panel dividido característico de la aplicación, con el sistema de archivos local a la izquierda y el sistema de archivos remoto del servidor Linux a la derecha.", note: "El uso de un cliente gráfico como WinSCP simplifica considerablemente la transferencia de archivos en comparación con comandos de consola como scp, especialmente para usuarios que se inician en la administración de servidores Linux desde un entorno Windows." },
            { lead: "Con el archivo ya alojado en el servidor, se emplea el comando unzip para descomprimir el contenido de la plantilla Lugx Gaming dentro del directorio de trabajo.", cmd: ["root@JorgeC:/opt2# unzip templatemo_589_lugx_gaming.zip"], title: "Descompresión de la plantilla con unzip", desc: "El comando unzip extrae íntegramente la estructura de carpetas y archivos contenida en el paquete comprimido, incluyendo las carpetas assets, css, images, js y webfonts, así como los archivos de estilo (animate.css, flex-slider.css, fontawesome.css, owl.css) y las imágenes correspondientes a las distintas secciones de la plantilla.", note: "Es necesario que el paquete unzip se encuentre previamente instalado en el sistema (apt-get install unzip) para poder descomprimir archivos con extensión .zip en un entorno Debian." },
            { lead: "Finalmente, se verifica el correcto despliegue del sitio web accediendo desde un navegador web a la dirección IP pública o local del servidor configurado.", cmd: ["# Navegador → http://172.30.107.57"], title: "Verificación final del sitio desplegado mediante el navegador", desc: "Al introducir la dirección IP del servidor en la barra de direcciones del navegador, Apache2 responde sirviendo la página principal (index.html) de la plantilla First Portfolio previamente desplegada, confirmando así el correcto funcionamiento integral de todo el procedimiento: instalación del servidor, configuración de acceso remoto, transferencia de archivos y despliegue del contenido web.", note: "Este resultado final constituye la validación práctica de que todos los pasos anteriores del laboratorio, ejecutados de manera coordinada entre ambos servidores, culminan exitosamente en un servicio web funcional y accesible desde la red." }
        ]
    }
];

module.exports = { sections, FUENTE };