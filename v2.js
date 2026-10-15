const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
    ShadingType, VerticalAlign, PageNumber, LevelFormat, PageBreak
} = require('docx');
const fs = require('fs');

const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };
const border = { style: BorderStyle.SINGLE, size: 4, color: "2C5F8A" };
const borderThin = { style: BorderStyle.SINGLE, size: 2, color: "AAAAAA" };
const borders = { top: border, bottom: border, left: border, right: border };
const bordersThin = { top: borderThin, bottom: borderThin, left: borderThin, right: borderThin };

function heading1(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text, bold: true, size: 28, color: "1F3864" })],
        spacing: { before: 360, after: 200 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2C5F8A", space: 1 } }
    });
}

function heading2(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text, bold: true, size: 24, color: "2C5F8A" })],
        spacing: { before: 240, after: 120 }
    });
}

function heading3(text) {
    return new Paragraph({
        children: [new TextRun({ text, bold: true, size: 22, color: "1F3864" })],
        spacing: { before: 200, after: 100 }
    });
}

function para(text, opts = {}) {
    return new Paragraph({
        children: [new TextRun({ text, size: 22, ...opts })],
        spacing: { before: 60, after: 60 },
        alignment: AlignmentType.JUSTIFIED
    });
}

function bullet(text) {
    return new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text, size: 22 })],
        spacing: { before: 40, after: 40 }
    });
}

function makeHeaderRow(cells, widths, color = "1F3864") {
    return new TableRow({
        tableHeader: true,
        children: cells.map((text, i) =>
            new TableCell({
                borders,
                width: { size: widths[i], type: WidthType.DXA },
                margins: cellMargins,
                shading: { fill: color, type: ShadingType.CLEAR },
                verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: "FFFFFF", size: 20 })], alignment: AlignmentType.CENTER })]
            })
        )
    });
}

function makeRow(cells, widths, shade = false) {
    return new TableRow({
        children: cells.map((text, i) =>
            new TableCell({
                borders: bordersThin,
                width: { size: widths[i], type: WidthType.DXA },
                margins: cellMargins,
                shading: shade ? { fill: "EEF3FA", type: ShadingType.CLEAR } : undefined,
                children: [new Paragraph({ children: [new TextRun({ text: text || "", size: 20 })], alignment: AlignmentType.LEFT })]
            })
        )
    });
}

function sectionLabel(label) {
    return new Paragraph({
        children: [new TextRun({ text: label, bold: true, size: 22, color: "C55A11" })],
        spacing: { before: 160, after: 60 }
    });
}

function pageBreak() {
    return new Paragraph({ children: [new PageBreak()] });
}

// ============================================================
// PLANTILLA FORMS
// ============================================================

function plantilla1() {
    const W = [1500, 2800, 1500, 2560];
    const total = W.reduce((a, b) => a + b, 0); // 8360
    const W2 = [1500, 6860];
    const W3 = [8360];
    return [
        heading3("Plantilla 1 – Solicitud de Modificación (MR)"),
        new Table({
            width: { size: 8360, type: WidthType.DXA },
            columnWidths: W,
            rows: [
                makeHeaderRow(["SOLICITUD DE MODIFICACIÓN", "", "", ""], W, "2C5F8A"),
                new TableRow({
                    children: [
                        new TableCell({ borders, width: { size: W[0], type: WidthType.DXA }, margins: cellMargins, shading: { fill: "EEF3FA", type: ShadingType.CLEAR }, children: [new Paragraph({ children: [new TextRun({ text: "Sección I", bold: true, size: 20, color: "1F3864" })] })], columnSpan: 4 })
                    ]
                }),
                makeRow(["Nombre del Solicitante:", "Encargado de Sede / Usuario del Sistema", "Fecha de Recepción:", "DD/MM/AAAA"], W),
                makeRow(["Sistema:", "Sistema Web – Admisiones CBVP", "Creador:", "Área de Desarrollo – Dir. Sistemas CBVP"], W),
            ]
        }),
        new Table({
            width: { size: 8360, type: WidthType.DXA },
            columnWidths: [5560, 2800],
            rows: [
                makeRow(["Descripción del Problema:", "Prioridad: 1  2  3  4  5"], [5560, 2800]),
                makeRow(["", ""], [5560, 2800]),
                makeRow(["", ""], [5560, 2800]),
            ]
        }),
        new Table({
            width: { size: 8360, type: WidthType.DXA },
            columnWidths: W,
            rows: [
                new TableRow({
                    children: [
                        new TableCell({ borders, width: { size: 8360, type: WidthType.DXA }, margins: cellMargins, shading: { fill: "EEF3FA", type: ShadingType.CLEAR }, children: [new Paragraph({ children: [new TextRun({ text: "Sección II", bold: true, size: 20, color: "1F3864" })] })], columnSpan: 4 })
                    ]
                }),
                makeRow(["N° de MR:", "MR-CBVP-2026-001", "Mantenedor:", "Desarrollador Backend – Dir. Sistemas"], W),
                makeRow(["Prioridad:", "2 – Alta", "Tipo de Mantenimiento:", "Correctivo"], W),
                new TableRow({
                    children: [
                        new TableCell({ borders, width: { size: 8360, type: WidthType.DXA }, margins: cellMargins, shading: { fill: "EEF3FA", type: ShadingType.CLEAR }, children: [new Paragraph({ children: [new TextRun({ text: "Sección III", bold: true, size: 20, color: "1F3864" })] })], columnSpan: 4 })
                    ]
                }),
                makeRow(["ID Opción:", "2", "Estado del MR:", "DD/MM/AAAA"], W),
                makeRow(["Resultados del Análisis:", "Se identificó el componente afectado. La opción 2 fue seleccionada por bajo riesgo y validación previa en staging.", "", ""], W),
                makeRow(["Aprobado por:", "Director de Sistemas – CBVP", "", ""], W),
            ]
        }),
        new Paragraph({ spacing: { after: 120 } })
    ];
}

function plantilla2() {
    const W = [1200, 1400, 1400, 1360, 3000];
    return [
        heading3("Plantilla 2 – Registro del Historial del MR"),
        new Table({
            width: { size: 8360, type: WidthType.DXA },
            columnWidths: W,
            rows: [
                makeHeaderRow(["ID MR", "Fecha de Recepción", "Fecha de Análisis", "Estado del MR", "Documentos Afectados"], W),
                makeRow(["MR-CBVP-2026-001", "DD/MM/AAAA", "DD/MM/AAAA", "En Análisis", "Plantilla 1, Plantilla 3"], W),
                makeRow(["MR-CBVP-2026-002", "DD/MM/AAAA", "DD/MM/AAAA", "Aprobado", "Plantilla 1, Plantilla 6"], W, true),
                makeRow(["MR-CBVP-2026-003", "DD/MM/AAAA", "Pendiente", "Recibido", "Plantilla 1"], W),
            ]
        }),
        new Paragraph({ spacing: { after: 120 } })
    ];
}

function plantilla3() {
    const W = [1860, 1860, 4640];
    return [
        heading3("Plantilla 3 – Registro de Pruebas del MR"),
        new Table({
            width: { size: 8360, type: WidthType.DXA },
            columnWidths: [2787, 2787, 2786],
            rows: [
                makeHeaderRow(["ID Prueba", "ID MR", "Nombre del Sistema"], [2787, 2787, 2786]),
                makeRow(["PU-01A", "MR-CBVP-2026-001", "Sistema Web – Admisiones CBVP"], [2787, 2787, 2786]),
                makeHeaderRow(["Responsable de la Prueba", "", ""], [2787, 2787, 2786]),
                makeRow(["Desarrollador Backend – Dir. Sistemas CBVP", "", ""], [2787, 2787, 2786]),
                makeHeaderRow(["Objetivo", "", ""], [2787, 2787, 2786]),
                makeRow(["Verificar que la falla reportada en el MR sea reproducible y documentar el comportamiento incorrecto del sistema para orientar el análisis de opciones.", "", ""], [2787, 2787, 2786]),
                makeHeaderRow(["Resultado Obtenido", "", ""], [2787, 2787, 2786]),
                makeRow(["La falla fue reproducida exitosamente en el entorno de desarrollo local. El módulo afectado responde con error 500 al recibir datos de entrada con campos vacíos no validados.", "", ""], [2787, 2787, 2786]),
                makeRow(["Estado: Falla confirmada. Se procede con la Fase 2 – Análisis de Opciones.", "", ""], [2787, 2787, 2786]),
            ]
        }),
        new Paragraph({ spacing: { after: 120 } })
    ];
}

function plantilla_pruebas(idPrueba, idMR, nombreSistema, responsable, objetivo, tipo) {
    return new Table({
        width: { size: 8360, type: WidthType.DXA },
        columnWidths: [1200, 1200, 1400, 1680, 1680, 900, 1300],
        rows: [
            makeHeaderRow(["ID MR: " + idMR, "ID Prueba: " + idPrueba, "Fecha Ejec.:", "", "Sistema: " + nombreSistema, "", ""], [1200, 1200, 1400, 1680, 1680, 900, 1300]),
            makeRow(["Tipo de Prueba: " + tipo, "", "Responsable: " + responsable, "", "Objetivo: " + objetivo, "", ""], [1200, 1200, 1400, 1680, 1680, 900, 1300]),
            makeHeaderRow(["Elemento a Probar", "Precondición", "Datos de Entrada", "Resultados Esperados", "Resultados Obtenidos", "Estado", "Obs."], [1200, 1200, 1400, 1680, 1680, 900, 1300]),
            makeRow(["", "", "", "", "", "", ""], [1200, 1200, 1400, 1680, 1680, 900, 1300]),
            makeRow(["", "", "", "", "", "", ""], [1200, 1200, 1400, 1680, 1680, 900, 1300], true),
        ]
    });
}

// ============================================================
// TEST TABLES PER RF
// ============================================================

function rfTestTable(tests) {
    // tests: [{id, tipo, responsable, objetivo}]
    const W = [800, 800, 1000, 1000, 1200, 3560];
    return new Table({
        width: { size: 8360, type: WidthType.DXA },
        columnWidths: W,
        rows: [
            makeHeaderRow(["ID MR", "ID Prueba", "Tipo", "Etapa", "Responsable", "Objetivo"], W),
            ...tests.map((t, i) => makeRow([t.rf, t.id, t.tipo, t.etapa, t.resp, t.obj], W, i % 2 === 1))
        ]
    });
}

const allTests = [
    { rf: "RF-01", id: "PU-01A", tipo: "Unitaria", etapa: "Construcción", resp: "Desarrollador", obj: "Validar que Cls_Usuario::registrarUsuario() inserte correctamente datos y hash de contraseña en BD." },
    { rf: "RF-01", id: "PF-01A", tipo: "Funcional", etapa: "Construcción", resp: "Desarrollador", obj: "Validar que Ctrllr_Registro::registrarUsuario() rechace contraseñas inseguras y verifique duplicados (DNI/Email)." },
    { rf: "RF-01", id: "PA-01A", tipo: "Aceptación", etapa: "Transición", resp: "QA Tester", obj: "Validar que el formulario de UI notifique éxito al registrarse y no permita avanzar sin CAPTCHA." },
    { rf: "RF-02", id: "PU-02A", tipo: "Unitaria", etapa: "Construcción", resp: "Desarrollador", obj: "Verificar que Cls_Usuario::validarUsuario() retorne el array del usuario si la contraseña coincide con el hash." },
    { rf: "RF-02", id: "PF-02A", tipo: "Funcional", etapa: "Construcción", resp: "Desarrollador", obj: "Verificar que Ctrllr_Login::login() maneje correctamente el flujo de 2FA (TOTP) y rechace credenciales inválidas." },
    { rf: "RF-02", id: "PA-02A", tipo: "Aceptación", etapa: "Transición", resp: "QA Tester", obj: "Validar que postulantes, encargados y administradores sean redirigidos a su Dashboard tras login." },
    { rf: "RF-03", id: "PU-03A", tipo: "Unitaria", etapa: "Construcción", resp: "Desarrollador", obj: "Validar que Cls_Usuario::actualizarPerfil() actualice correctamente los campos seleccionados." },
    { rf: "RF-03", id: "PF-03A", tipo: "Funcional", etapa: "Construcción", resp: "Desarrollador", obj: "Verificar en Ctrllr_Usuario que no se permita usar un email ya perteneciente a otro ID." },
    { rf: "RF-03", id: "PA-03A", tipo: "Aceptación", etapa: "Transición", resp: "QA Tester", obj: "Validar que al modificar datos en 'Mi Perfil', los datos se refresquen y aparezca mensaje Toast de éxito." },
    { rf: "RF-04", id: "PU-04A", tipo: "Unitaria", etapa: "Construcción", resp: "Desarrollador", obj: "Validar que Cls_Usuario::crearEncargado() y crearAdministrador() asignen el id_grado y permisos correctos." },
    { rf: "RF-04", id: "PF-04A", tipo: "Funcional", etapa: "Construcción", resp: "Desarrollador", obj: "Verificar que el bloque action=crear_administrador bloquee la creación si la sesión no tiene superusuario == 1." },
    { rf: "RF-04", id: "PA-04A", tipo: "Aceptación", etapa: "Transición", resp: "QA Tester", obj: "Validar que el Administrador pueda ver tabla de usuarios filtrada por sede y registrar nuevos Encargados." },
    { rf: "RF-05", id: "PU-05A", tipo: "Unitaria", etapa: "Construcción", resp: "Desarrollador", obj: "Verificar que Cls_ProcesoAdmision::crearProceso() inserta el registro y retorna un ID válido." },
    { rf: "RF-05", id: "PF-05A", tipo: "Funcional", etapa: "Construcción", resp: "Desarrollador", obj: "Verificar que action=crear_proceso y actualizar_proceso devuelvan error JSON si cupos <= 0 o fecha_fin < fecha_inicio." },
    { rf: "RF-05", id: "PA-05A", tipo: "Aceptación", etapa: "Transición", resp: "QA Tester", obj: "Validar que el encargado llene el modal de nuevo proceso y este aparezca en la grilla de procesos activos." },
    { rf: "RF-06", id: "PU-06A", tipo: "Unitaria", etapa: "Construcción", resp: "Desarrollador", obj: "Validar que Cls_Resolucion::generarFirmaHtml() incluya la etiqueta <img> correctamente cuando la firma es URL." },
    { rf: "RF-06", id: "PF-06A", tipo: "Funcional", etapa: "Construcción", resp: "Desarrollador", obj: "Probar que Ctrllr_Resolucion::publicarResolucion() actualiza el estado del proceso a 'Activo' en la BD." },
    { rf: "RF-06", id: "PA-06A", tipo: "Aceptación", etapa: "Transición", resp: "QA Tester", obj: "Validar que el encargado pueda publicar una resolución y reciba el mensaje de éxito en la UI." },
    { rf: "RF-07", id: "PU-07A", tipo: "Unitaria", etapa: "Construcción", resp: "Desarrollador", obj: "Validar que Ctrllr_Inscripcion::procesarInscripcion() retorne success=false si el usuario no tiene documentos." },
    { rf: "RF-07", id: "PF-07A", tipo: "Funcional", etapa: "Construcción", resp: "Desarrollador", obj: "Verificar que al enviar documentos aprobados, se llame a Cls_InscripcionProceso::inscribirPostulante() y cree el registro." },
    { rf: "RF-07", id: "PA-07A", tipo: "Aceptación", etapa: "Transición", resp: "QA Tester", obj: "Validar que el postulante pueda inscribirse exitosamente desde el listado de procesos." },
    { rf: "RF-08", id: "PU-08A", tipo: "Unitaria", etapa: "Construcción", resp: "Desarrollador", obj: "Verificar que Cls_DocumentoPostulante::calcularHash() genera el hash SHA-256 correcto." },
    { rf: "RF-08", id: "PF-08A", tipo: "Funcional", etapa: "Construcción", resp: "Desarrollador", obj: "Verificar que el registro insertado en Cls_DocumentoPostulante contenga el hash y estado 'En Revisión'." },
    { rf: "RF-08", id: "PA-08A", tipo: "Aceptación", etapa: "Transición", resp: "QA Tester", obj: "Validar que el postulante adjunte un PDF y aparezca en la lista con estado 'En Revisión'." },
    { rf: "RF-09", id: "PU-09A", tipo: "Unitaria", etapa: "Construcción", resp: "Desarrollador", obj: "Validar que el bloque action=distribuir_automatico calcule y asigne correctamente la distribución de postulantes." },
    { rf: "RF-09", id: "PF-09A", tipo: "Funcional", etapa: "Construcción", resp: "Desarrollador", obj: "Validar que Cls_InscripcionProceso::asignarPostulantesAGrupo() actualice correctamente los IDs en BD." },
    { rf: "RF-09", id: "PA-09A", tipo: "Aceptación", etapa: "Transición", resp: "QA Tester", obj: "Validar que el encargado asigne postulantes a grupos y reciba confirmación de éxito en la UI." },
    { rf: "RF-10", id: "PU-10A", tipo: "Unitaria", etapa: "Construcción", resp: "Desarrollador", obj: "Verificar que Cls_HorarioExamen::guardarExamen() inserte la fecha, hora_inicio y lugar correctamente." },
    { rf: "RF-10", id: "PF-10A", tipo: "Funcional", etapa: "Construcción", resp: "Desarrollador", obj: "Verificar que al invocar action=guardarExamen, se retorne un error JSON si hay solapamiento de horarios." },
    { rf: "RF-10", id: "PA-10A", tipo: "Aceptación", etapa: "Transición", resp: "QA Tester", obj: "Validar que el encargado calendarice un examen físico y los postulantes del grupo lo vean en 'Mis Exámenes'." },
    { rf: "RF-11", id: "PU-11A", tipo: "Unitaria", etapa: "Construcción", resp: "Desarrollador", obj: "Validar que Cls_EvaluacionPostulante::registrarResultadoFisico() guarde correctamente resultados y veredicto." },
    { rf: "RF-11", id: "PF-11A", tipo: "Funcional", etapa: "Construcción", resp: "Desarrollador", obj: "Verificar que al enviar veredicto 'No Apto', el controlador Ctrllr_EvaluacionFisica.php asigne estado 'Rechazado'." },
    { rf: "RF-11", id: "PA-11A", tipo: "Aceptación", etapa: "Transición", resp: "QA Tester", obj: "Validar que el evaluador ingrese notas tipo Excel y al 'Guardar Todo' se notifique el éxito correctamente." },
    { rf: "RF-12", id: "PU-12A", tipo: "Unitaria", etapa: "Construcción", resp: "Desarrollador", obj: "Validar que las consultas JOIN agrupen correctamente los datos de postulantes, resultados e información del proceso." },
    { rf: "RF-12", id: "PF-12A", tipo: "Funcional", etapa: "Construcción", resp: "Desarrollador", obj: "Verificar que el sistema envíe la trama JSON correcta al servicio RPA al cerrarse el proceso de admisión." },
    { rf: "RF-12", id: "PA-12A", tipo: "Aceptación", etapa: "Transición", resp: "QA Tester", obj: "Validar que desde 'Procesos Finalizados', la Entidad Central pueda descargar el PDF oficial generado automáticamente." },
];

const doc = new Document({
    numbering: {
        config: [
            { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
            { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
        ]
    },
    styles: {
        default: { document: { run: { font: "Arial", size: 22 } } },
        paragraphStyles: [
            {
                id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 28, bold: true, font: "Arial", color: "1F3864" },
                paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 }
            },
            {
                id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 24, bold: true, font: "Arial", color: "2C5F8A" },
                paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 }
            },
        ]
    },
    sections: [{
        properties: {
            page: {
                size: { width: 12240, height: 15840 },
                margin: { top: 1440, right: 1260, bottom: 1440, left: 1440 }
            }
        },
        headers: {
            default: new Header({
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({ text: "Plan de Mantenimiento de Software – Sistema Web Admisiones CBVP", size: 18, color: "555555" }),
                        ],
                        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "2C5F8A", space: 1 } }
                    })
                ]
            })
        },
        footers: {
            default: new Footer({
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({ text: "Cuerpo de Bomberos Voluntarios del Perú – Confidencial  |  Página ", size: 18, color: "777777" }),
                            new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "777777" }),
                            new TextRun({ text: " de ", size: 18, color: "777777" }),
                            new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: "777777" }),
                        ],
                        alignment: AlignmentType.RIGHT,
                        border: { top: { style: BorderStyle.SINGLE, size: 4, color: "2C5F8A", space: 1 } }
                    })
                ]
            })
        },
        children: [
            // ====== PORTADA ======
            new Paragraph({ spacing: { before: 1440 } }),
            new Paragraph({
                children: [new TextRun({ text: "CUERPO DE BOMBEROS VOLUNTARIOS DEL PERÚ", bold: true, size: 36, color: "1F3864" })],
                alignment: AlignmentType.CENTER, spacing: { before: 240, after: 120 }
            }),
            new Paragraph({
                children: [new TextRun({ text: "JUNTA NACIONAL", bold: true, size: 26, color: "2C5F8A" })],
                alignment: AlignmentType.CENTER, spacing: { after: 480 }
            }),
            new Paragraph({
                children: [new TextRun({ text: "PLAN DE MANTENIMIENTO DE SOFTWARE", bold: true, size: 40, color: "C55A11" })],
                alignment: AlignmentType.CENTER, spacing: { before: 240, after: 240 }
            }),
            new Paragraph({
                children: [new TextRun({ text: "Sistema Web con Automatización Robótica de Procesos (RPA)", bold: true, size: 28, color: "1F3864" })],
                alignment: AlignmentType.CENTER, spacing: { after: 120 }
            }),
            new Paragraph({
                children: [new TextRun({ text: "para la Gestión de Admisiones", bold: true, size: 28, color: "1F3864" })],
                alignment: AlignmentType.CENTER, spacing: { after: 480 }
            }),
            new Paragraph({ children: [new TextRun({ text: "Versión: 1.0", size: 22, color: "444444" })], alignment: AlignmentType.CENTER }),
            new Paragraph({ children: [new TextRun({ text: "Elaborado por: Área de Desarrollo de Sistemas", size: 22, color: "444444" })], alignment: AlignmentType.CENTER }),
            new Paragraph({ children: [new TextRun({ text: "Fecha: Junio 2026", size: 22, color: "444444" })], alignment: AlignmentType.CENTER }),
            new Paragraph({ children: [new TextRun({ text: "Basado en: ISO/IEC 14764:2006", size: 22, color: "444444" })], alignment: AlignmentType.CENTER }),
            pageBreak(),

            // ====== CAPÍTULO 1: INTRODUCCIÓN ======
            heading1("CAPÍTULO 1: INTRODUCCIÓN"),
            heading2("1.1. Descripción del Sistema"),
            para("El sistema objeto de este plan es una aplicación web desarrollada con arquitectura MVC en PHP puro, que incorpora un módulo de Automatización Robótica de Procesos (RPA) para la generación automática de informes oficiales. El sistema gestiona de forma integral el proceso de admisión de voluntarios en el Cuerpo de Bomberos Voluntarios del Perú (CBVP), cubriendo desde el registro y autenticación de usuarios hasta la generación automatizada del informe final de resultados mediante tecnología RPA."),
            para("El sistema contempla tres perfiles de usuario: Postulante, Encargado de Sede y Entidad Central (Administrador). Cada perfil accede a funcionalidades específicas del flujo de admisión, asegurando la trazabilidad y seguridad de la información en cada etapa del proceso."),
            new Paragraph({ spacing: { after: 120 } }),
            heading2("1.2. Fases del Proceso de Mantenimiento"),
            para("El presente plan de mantenimiento se fundamenta en la norma ISO/IEC 14764:2006 y contempla las siguientes fases:"),
            ...[
                "Fase 1: Proceso de Implementación",
                "Fase 2: Análisis de Modificaciones",
                "Fase 3: Implementación de la Modificación",
                "Fase 4: Aceptación y Revisión del Mantenimiento",
                "Fase 5: Migración",
                "Fase 6: Retiro del Software"
            ].map(t => bullet(t)),
            new Paragraph({ spacing: { after: 120 } }),
            pageBreak(),

            // ====== CAPÍTULO 2: ALCANCE Y PROPÓSITO ======
            heading1("CAPÍTULO 2: ALCANCE Y PROPÓSITO"),
            heading2("2.1. Alcance"),
            para("El presente plan de mantenimiento aplica al Sistema Web de Gestión de Admisiones del Cuerpo de Bomberos Voluntarios del Perú, incluyendo todos sus módulos funcionales: gestión de usuarios y roles, administración de procesos de admisión, carga y validación de documentos, asignación de grupos, programación de exámenes físicos, registro de resultados y generación automatizada de informes mediante RPA."),
            heading2("2.2. Propósito"),
            para("Brindar los criterios, procedimientos y plantillas necesarios para ejecutar actividades de mantenimiento sobre el sistema de admisiones del CBVP de forma ordenada, trazable y alineada con los estándares de calidad de software internacionales. Este plan debe ser referenciado en el Plan de Gestión de Proyectos de la Dirección de Sistemas del CBVP."),
            pageBreak(),

            // ====== CAPÍTULO 3: DOCUMENTOS APLICABLES ======
            heading1("CAPÍTULO 3: DOCUMENTOS APLICABLES Y DEFINICIONES"),
            heading2("3.1. Documentos Aplicables"),
            ...[
                "Misión y Visión Institucional del CBVP (emitida por la Junta Nacional)",
                "Política de Mejora Continua de Procesos del CBVP",
                "Manual de Proceso de Organización de Mantenimiento de Software",
                "Manual de Operación del Sistema de Gestión de la Calidad",
                "Norma IEEE 610.12 – Glosario Estándar de Ingeniería de Software",
                "ISO/IEC 14764:2006 – Mantenimiento de Software",
                "ISO/IEC 12207:2008 – Procesos del Ciclo de Vida del Software",
                "Documentación técnica del sistema: MVC PHP, módulo RPA, base de datos MySQL"
            ].map(t => bullet(t)),
            heading2("3.2. Definiciones"),
            ...[
                ["Mantenimiento Correctivo:", "Modificación realizada tras la entrega para corregir defectos detectados en el sistema de admisiones."],
                ["Mantenimiento Adaptativo:", "Modificación para mantener el sistema operativo ante cambios en el entorno (nuevas versiones de PHP, cambios en la BD, integración con nuevos servicios)."],
                ["Mantenimiento Perfectivo:", "Modificación para detectar y corregir fallas latentes antes de que se manifiesten como errores operacionales."],
                ["Mantenimiento Preventivo:", "Modificación proactiva para mejorar la mantenibilidad o rendimiento sin alterar las funcionalidades existentes."],
                ["Solicitud de Modificación (MR):", "Término genérico para identificar cualquier cambio propuesto al sistema."],
                ["RPA (Robotic Process Automation):", "Tecnología utilizada en el módulo de generación automática de informes PDF, que automatiza la extracción y consolidación de datos de postulantes en plantillas oficiales del CBVP."],
                ["Mantenedor:", "Persona u organización responsable de ejecutar las actividades de mantenimiento del sistema."],
                ["Hash SHA-256:", "Función criptográfica utilizada para garantizar la integridad de los documentos subidos por los postulantes (DNI, certificados médicos, etc.)."],
                ["TOTP (Time-based One-Time Password):", "Mecanismo de autenticación de dos factores (2FA) implementado en el módulo de inicio de sesión del sistema."],
            ].map(([t, d]) => new Paragraph({
                children: [
                    new TextRun({ text: t + " ", bold: true, size: 22 }),
                    new TextRun({ text: d, size: 22 })
                ],
                spacing: { before: 80, after: 60 },
                numbering: { reference: "bullets", level: 0 }
            })),
            pageBreak(),

            // ====== CAPÍTULO 4: REQUERIMIENTOS FUNCIONALES ======
            heading1("CAPÍTULO 4: REQUERIMIENTOS FUNCIONALES DEL SISTEMA"),
            para("La siguiente tabla consolida los requerimientos funcionales del sistema agrupados por necesidad institucional. Estos requerimientos constituyen la línea base sobre la cual se ejecutarán las actividades de mantenimiento."),
            new Paragraph({ spacing: { after: 120 } }),
            new Table({
                width: { size: 8360, type: WidthType.DXA },
                columnWidths: [600, 1200, 1000, 5560],
                rows: [
                    makeHeaderRow(["Nec.", "Código RF", "Título", "Descripción del Requerimiento"], [600, 1200, 1000, 5560]),
                    makeRow(["N1", "RF-01", "Registrar usuario", "Permitir al postulante registrarse ingresando correo, DNI y nombre de usuario."], [600, 1200, 1000, 5560]),
                    makeRow(["N1", "RF-02", "Iniciar sesión", "Permitir acceso mediante usuario y contraseña con autenticación 2FA (TOTP)."], [600, 1200, 1000, 5560], true),
                    makeRow(["N1", "RF-03", "Modificar Perfil", "Permitir a todo tipo de usuario modificar sus datos personales."], [600, 1200, 1000, 5560]),
                    makeRow(["N2", "RF-04", "Gestionar usuarios y roles", "El superusuario registra administradores; los administradores registran encargados."], [600, 1200, 1000, 5560], true),
                    makeRow(["N3", "RF-05", "Gestionar procesos", "Administrador crea y configura procesos indicando sede, fechas, cupos y requisitos."], [600, 1200, 1000, 5560]),
                    makeRow(["N3", "RF-06", "Autorizar con firma digital", "La entidad central aprueba/rechaza apertura de proceso mediante firma digital."], [600, 1200, 1000, 5560], true),
                    makeRow(["N4", "RF-07", "Postular a proceso", "El postulante se inscribe en un proceso abierto y confirma su participación."], [600, 1200, 1000, 5560]),
                    makeRow(["N5", "RF-08", "Cargar documentos", "El postulante sube documentos (DNI, cert. médico, etc.) con validación por hash SHA-256."], [600, 1200, 1000, 5560], true),
                    makeRow(["N6", "RF-09", "Asignar grupo", "El sistema asigna automáticamente postulantes en grupos de examen."], [600, 1200, 1000, 5560]),
                    makeRow(["N6", "RF-10", "Programar Examen", "El encargado define cronograma de exámenes físicos con fecha, hora y lugar."], [600, 1200, 1000, 5560], true),
                    makeRow(["N6", "RF-11", "Registrar resultados", "El encargado ingresa resultados físicos (altura, peso, apto/no apto) en el sistema."], [600, 1200, 1000, 5560]),
                    makeRow(["N7", "RF-12", "Generar informe RPA", "El sistema genera automáticamente un PDF con resultados usando RPA y plantilla oficial."], [600, 1200, 1000, 5560], true),
                ]
            }),
            pageBreak(),

            // ====== CAPÍTULO 5: REGISTRO CONSOLIDADO DE PRUEBAS ======
            heading1("CAPÍTULO 5: REGISTRO CONSOLIDADO DE PRUEBAS"),
            para("Este capítulo consolida todas las pruebas planificadas para el sistema, clasificadas por tipo (Unitaria, Funcional, Aceptación) y organizadas por requerimiento funcional. Las pruebas fueron diseñadas y ejecutadas durante las fases de Construcción y Transición del proyecto."),
            new Paragraph({ spacing: { after: 120 } }),
            new Table({
                width: { size: 8360, type: WidthType.DXA },
                columnWidths: [600, 800, 900, 1000, 1000, 4060],
                rows: [
                    makeHeaderRow(["RF", "Prueba", "Tipo", "Prioridad", "Responsable", "Objetivo"], [600, 800, 900, 1000, 1000, 4060]),
                    ...allTests.map((t, i) => {
                        const prioMap = {
                            "PU-01A": "Bajo", "PF-01A": "Medio", "PA-01A": "Muy alto",
                            "PU-02A": "Bajo", "PF-02A": "Medio", "PA-02A": "Muy alto",
                            "PU-03A": "Bajo", "PF-03A": "Medio", "PA-03A": "Muy alto",
                            "PU-04A": "Bajo", "PF-04A": "Alto", "PA-04A": "Muy alto",
                            "PU-05A": "Bajo", "PF-05A": "Alto", "PA-05A": "Muy alto",
                            "PU-06A": "Bajo", "PF-06A": "Medio", "PA-06A": "Muy alto",
                            "PU-07A": "Bajo", "PF-07A": "Medio", "PA-07A": "Muy alto",
                            "PU-08A": "Bajo", "PF-08A": "Medio", "PA-08A": "Muy alto",
                            "PU-09A": "Bajo", "PF-09A": "Medio", "PA-09A": "Muy alto",
                            "PU-10A": "Bajo", "PF-10A": "Alto", "PA-10A": "Muy alto",
                            "PU-11A": "Bajo", "PF-11A": "Medio", "PA-11A": "Muy alto",
                            "PU-12A": "Bajo", "PF-12A": "Alto", "PA-12A": "Muy alto",
                        };
                        return makeRow([t.rf, t.id, t.tipo, prioMap[t.id] || "Medio", t.resp, t.obj], [600, 800, 900, 1000, 1000, 4060], i % 2 === 1);
                    })
                ]
            }),
            pageBreak(),

            // ====== CAPÍTULO 6: FASE 1 ======
            heading1("CAPÍTULO 6: FASE 1 – PROCESO DE IMPLEMENTACIÓN"),
            heading2("6.1. Entradas"),
            bullet("Una Solicitud de Modificación (MR) presentada por el solicitante (postulante, encargado, entidad central o área de TI)."),
            heading2("6.2. Actividad 1: Recepción del MR"),
            heading3("Tarea 1"),
            para("Recibir el MR (Plantilla 1) debidamente llenado por el solicitante y asignar un número de identificación único con el formato: MR-CBVP-[AÑO]-[SECUENCIAL] (ej.: MR-CBVP-2026-001)."),
            heading3("Tarea 2"),
            para("Asignar prioridad al MR según la tabla de prioridades definida para el Sistema de Admisiones CBVP:"),
            new Paragraph({ spacing: { after: 80 } }),
            new Table({
                width: { size: 8360, type: WidthType.DXA },
                columnWidths: [1000, 7360],
                rows: [
                    makeHeaderRow(["Prioridad", "Criterio de Aplicación"], [1000, 7360]),
                    makeRow(["1 – Crítica", "El fallo impide el proceso de admisión activo (ej.: fallo en login, inscripción o generación de informes RPA)."], [1000, 7360]),
                    makeRow(["2 – Alta", "Afecta gravemente una función esencial sin solución temporal (ej.: error en firma digital, validación de documentos)."], [1000, 7360], true),
                    makeRow(["3 – Media", "Afecta funcionalidad esencial pero existe workaround temporal (ej.: error en asignación de grupos)."], [1000, 7360]),
                    makeRow(["4 – Baja", "Presenta molestias al usuario sin afectar operaciones críticas (ej.: error visual en dashboard)."], [1000, 7360], true),
                    makeRow(["5 – Mínima", "Cualquier otro efecto menor (ej.: ajustes de texto, íconos, colores de interfaz)."], [1000, 7360]),
                ]
            }),
            new Paragraph({ spacing: { after: 120 } }),
            heading3("Plantilla 1 – Solicitud de Modificación"),
            ...[...plantilla1()],
            heading2("6.3. Controles y Salidas"),
            para("Las revisiones conjuntas entre el Mantenedor y el Director de Sistemas del CBVP se utilizarán para controlar las salidas de esta fase. La salida es el MR receptado y registrado con número único."),
            pageBreak(),

            // ====== CAPÍTULO 7: FASE 2 ======
            heading1("CAPÍTULO 7: FASE 2 – ANÁLISIS DE MODIFICACIONES"),
            heading2("7.1. Entradas"),
            bullet("MR receptado de la Fase 1"),
            bullet("Documentación técnica del sistema (código fuente PHP/MVC, esquema de base de datos MySQL, documentación del módulo RPA)"),
            heading2("7.2. Actividad 1: Análisis del MR"),
            heading3("Tarea 1"),
            para("Determinar el tipo de mantenimiento aplicable al MR recibido. Para el Sistema de Admisiones CBVP, los tipos más frecuentes son:"),
            ...[
                "Correctivo: errores en validaciones de documentos, cálculo de hash SHA-256, o fallo en el flujo 2FA.",
                "Adaptativo: actualizaciones por cambios en la API del servicio RPA externo o cambios en los requisitos documentales del CBVP.",
                "Perfectivo: mejoras al rendimiento del módulo de consultas JOIN para informes o mejoras en la UI del evaluador.",
                "Preventivo: refactorización de código PHP, optimización de índices en MySQL, revisión de seguridad del módulo CAPTCHA."
            ].map(t => bullet(t)),
            heading2("7.3. Actividad 2: Verificación"),
            heading3("Tarea 1 – Registro del Historial del MR (Plantilla 2)"),
            ...[...plantilla2()],
            heading3("Tarea 2"),
            para("Obtener el código fuente y configuración del módulo afectado. Para cambios en el módulo RPA, incluir la trama JSON de comunicación con el servicio externo."),
            heading3("Tarea 3 – Registro de Pruebas del MR (Plantilla 3)"),
            para("Realizar pruebas sobre el sistema para comprobar la falla reportada y documentar los resultados:"),
            ...[...plantilla3()],
            heading2("7.4. Actividad 3: Desarrollo de Opciones"),
            heading3("Tarea 4 – Registro de Opciones (Plantilla 4)"),
            new Table({
                width: { size: 8360, type: WidthType.DXA },
                columnWidths: [800, 2200, 1200, 1200, 1160, 1800],
                rows: [
                    makeHeaderRow(["ID MR", "Nombre del Sistema", "", "Responsable:", "", ""], [800, 2200, 1200, 1200, 1160, 1800]),
                    makeRow(["MR-CBVP-2026-001", "Sistema Web de Admisiones – CBVP", "", "Coordinador de Desarrollo – Dir. Sistemas", "", ""], [800, 2200, 1200, 1200, 1160, 1800]),
                    makeHeaderRow(["Requerimientos de la Organización:", "", "", "", "", ""], [800, 2200, 1200, 1200, 1160, 1800]),
                    makeRow(["Garantizar la continuidad operativa del proceso de admisión sin afectar datos de postulantes ni integridad del módulo RPA.", "", "", "", "", ""], [800, 2200, 1200, 1200, 1160, 1800]),
                    makeHeaderRow(["ID Opción", "Descripción", "Extensión", "Impacto", "Riesgos", "Recomendación"], [800, 2200, 1200, 1200, 1160, 1800]),
                    makeRow(["1", "Corrección directa en producción con respaldo previo de BD", "2 días", "Alto – afecta módulo activo", "Pérdida de datos si el respaldo falla", "No recomendada"], [800, 2200, 1200, 1200, 1160, 1800]),
                    makeRow(["2", "Implementación en entorno de pruebas, validación completa y pase a producción en ventana de mantenimiento", "5 días", "Medio – sin afectación al usuario", "Retraso si la ventana es corta", "Recomendada"], [800, 2200, 1200, 1200, 1160, 1800], true),
                    makeRow(["3", "Rollback de versión anterior mientras se desarrolla la corrección definitiva", "1 día (rollback) + 7 días (fix)", "Bajo en corto plazo, Alto en largo plazo", "Acumulación de deuda técnica", "Alternativa si opción 2 no es viable"], [800, 2200, 1200, 1200, 1160, 1800]),
                ]
            }),
            new Paragraph({ spacing: { after: 120 } }),
            heading3("Tarea 5 – Registro de Asignación de Recursos (Plantilla 5)"),
            new Table({
                width: { size: 8360, type: WidthType.DXA },
                columnWidths: [800, 1800, 1200, 1200, 1560, 1800],
                rows: [
                    makeHeaderRow(["ID Opción", "N° Personas", "N° Horas", "Costo (S/.)", "Herramientas", "Observación"], [800, 1800, 1200, 1200, 1560, 1800]),
                    makeRow(["1", "2 (1 Backend, 1 DBA)", "8", "0.00", "Git, MySQL Workbench", "Riesgo alto; no recomendada"], [800, 1800, 1200, 1200, 1560, 1800]),
                    makeRow(["2", "3 (1 Backend, 1 Frontend, 1 QA)", "40", "0.00", "Git, XAMPP local, Postman, Selenium", "Opción seleccionada; costo cubierto por área TI"], [800, 1800, 1200, 1200, 1560, 1800], true),
                    makeRow(["3", "2 (1 DevOps, 1 Backend)", "16", "0.00", "Git, servidor staging CBVP", "Solo si opción 2 no puede ejecutarse en el plazo"], [800, 1800, 1200, 1200, 1560, 1800]),
                ]
            }),
            new Paragraph({ spacing: { after: 120 } }),
            heading2("7.5. Actividad 4: Documentación"),
            para("Verificar que todas las pruebas y opciones estén documentadas en las plantillas correspondientes y actualizar el Registro del Historial del MR (Plantilla 2)."),
            heading2("7.6. Actividad 5: Aprobación"),
            heading3("Plantilla 6 – Registro del Análisis de la Solicitud de Modificación"),
            new Table({
                width: { size: 8360, type: WidthType.DXA },
                columnWidths: [1800, 6560],
                rows: [
                    makeHeaderRow(["REGISTRO DEL ANÁLISIS DE LA SOLICITUD DE MODIFICACIÓN", ""], [1800, 6560]),
                    makeRow(["Nombre del Sistema:", "Sistema Web de Admisiones – CBVP"], [1800, 6560]),
                    makeRow(["Responsable(s):", "Desarrollador Backend / QA Tester – Dirección de Sistemas CBVP"], [1800, 6560], true),
                    makeRow(["Fecha de recepción del MR:", "DD/MM/AAAA"], [1800, 6560]),
                    makeRow(["Fecha de aprobación/negado:", "DD/MM/AAAA"], [1800, 6560], true),
                    makeRow(["Fecha de cierre:", "DD/MM/AAAA"], [1800, 6560]),
                    makeHeaderRow(["ID MR", "Descripción de la Solución"], [1800, 6560]),
                    makeRow(["MR-CBVP-2026-001", "Se identificó el error en el controlador afectado. La solución implementada corrige la validación de campos requeridos y agrega manejo de excepciones. Se ejecutaron pruebas unitarias y funcionales con resultado satisfactorio en entorno staging antes del pase a producción."], [1800, 6560]),
                    makeHeaderRow(["Estado del MR", "Motivo del Rechazo (si aplica)"], [1800, 6560]),
                    makeRow(["Aprobado", "N/A"], [1800, 6560]),
                    makeHeaderRow(["Responsable del Análisis:", ""], [1800, 6560]),
                    makeRow(["Coordinador de Desarrollo – Dirección de Sistemas CBVP", ""], [1800, 6560]),
                ]
            }),
            new Paragraph({ spacing: { after: 120 } }),
            heading3("Plantilla 7 – Registro del Equipo de Trabajo"),
            new Table({
                width: { size: 8360, type: WidthType.DXA },
                columnWidths: [1200, 900, 2200, 1560, 2500],
                rows: [
                    makeHeaderRow(["Nombre del Sistema", "ID MR", "Mantenedor", "Rol", "Observación"], [1200, 900, 2200, 1560, 2500]),
                    makeRow(["Admisiones CBVP", "MR-CBVP-2026-001", "Ing. Carlos Quispe", "Desarrollador Backend", "Modificación de controladores PHP y validaciones de BD"], [1200, 900, 2200, 1560, 2500]),
                    makeRow(["Admisiones CBVP", "MR-CBVP-2026-001", "Ing. Ana Flores", "Desarrollador Frontend", "Actualización de vistas Vue/JS y mensajes de error en UI"], [1200, 900, 2200, 1560, 2500], true),
                    makeRow(["Admisiones CBVP", "MR-CBVP-2026-001", "Tec. Luis Mamani", "QA Tester", "Ejecución de pruebas unitarias, funcionales y de aceptación"], [1200, 900, 2200, 1560, 2500]),
                    makeRow(["Admisiones CBVP", "MR-CBVP-2026-001", "Ing. Rosa Condori", "Responsable RPA", "Verificación de integración con servicio RPA y trama JSON"], [1200, 900, 2200, 1560, 2500], true),
                ]
            }),
            new Paragraph({ spacing: { after: 120 } }),
            heading2("7.7. Salidas de la Fase 2"),
            ...[
                "Registro de Pruebas del MR (Plantilla 3)",
                "Registro de Opciones (Plantilla 4)",
                "Registro de Asignación de Recursos (Plantilla 5)",
                "Registro del Análisis de la Solicitud de Modificación (Plantilla 6)",
                "Registro del Equipo de Trabajo (Plantilla 7)",
                "MR actualizado (Plantilla 1 – Sección II)",
                "Registro del Historial del MR actualizado (Plantilla 2)"
            ].map(t => bullet(t)),
            pageBreak(),

            // ====== CAPÍTULO 8: FASE 3 ======
            heading1("CAPÍTULO 8: FASE 3 – IMPLEMENTACIÓN DE LA MODIFICACIÓN"),
            heading2("8.1. Entradas"),
            bullet("Salidas de la Fase 2 (análisis, opciones y equipo aprobados)"),
            bullet("Código fuente del sistema (repositorio Git)"),
            bullet("Acceso al servicio RPA en caso de modificaciones al módulo RF-12"),
            heading2("8.2. Actividad 1: Análisis Detallado"),
            heading3("Plantilla 8 – Listado de Elementos a Modificar"),
            new Table({
                width: { size: 8360, type: WidthType.DXA },
                columnWidths: [400, 1500, 1300, 2260, 1600, 1300],
                rows: [
                    makeHeaderRow(["N°", "Elemento a Modificar", "Capa Arquitectura", "Descripción de la Modificación", "Responsable", "Duración"], [400, 1500, 1300, 2260, 1600, 1300]),
                    makeRow(["1", "Cls_Usuario.php", "Modelo", "Agregar validación de campos obligatorios en registrarUsuario() y actualizarPerfil(); reforzar hash bcrypt.", "Desarrollador Backend", "4 h"], [400, 1500, 1300, 2260, 1600, 1300]),
                    makeRow(["2", "Ctrllr_Login.php", "Controlador", "Corregir manejo de excepciones en flujo 2FA (TOTP); agregar bloqueo por intentos fallidos consecutivos.", "Desarrollador Backend", "6 h"], [400, 1500, 1300, 2260, 1600, 1300], true),
                    makeRow(["3", "Vista de UI (Vue/JS)", "Vista", "Actualizar mensajes de error en formulario de login y registro; agregar feedback visual de validación en tiempo real.", "Desarrollador Frontend", "4 h"], [400, 1500, 1300, 2260, 1600, 1300]),
                    makeRow(["4", "Servicio RPA", "Servicio Externo", "Verificar y actualizar estructura de trama JSON enviada al servicio RPA para el informe PDF; ajustar campos de postulante según nueva versión de plantilla oficial.", "Responsable RPA", "3 h"], [400, 1500, 1300, 2260, 1600, 1300], true),
                    makeRow(["5", "BD MySQL", "Base de Datos", "Agregar índice en columna id_proceso de tabla inscripcion_proceso; revisar restricciones FK en tabla documento_postulante.", "DBA", "2 h"], [400, 1500, 1300, 2260, 1600, 1300]),
                ]
            }),
            new Paragraph({ spacing: { after: 120 } }),
            heading2("8.3. Actividad 2: Proceso de Desarrollo"),
            heading3("Tarea 2 – Registro de Pruebas sobre el Sistema Modificado (Plantilla 9)"),
            para("Para el sistema de admisiones CBVP, los tipos de prueba aplicables son:"),
            ...[
                "Pruebas Unitarias: validación de métodos de clases de modelo (Cls_Usuario, Cls_DocumentoPostulante, etc.)",
                "Pruebas Funcionales: validación de controladores y flujos de negocio completos.",
                "Pruebas de Aceptación: validación end-to-end desde la interfaz de usuario.",
                "Pruebas de Regresión: verificar que las modificaciones no alteren funcionalidades previas.",
                "Pruebas de Integración con RPA: verificar la trama JSON enviada al servicio RPA y el PDF resultante."
            ].map(t => bullet(t)),
            new Paragraph({ spacing: { after: 80 } }),
            new Table({
                width: { size: 8360, type: WidthType.DXA },
                columnWidths: [800, 900, 1600, 1200, 1260, 1200, 1400],
                rows: [
                    makeHeaderRow(["ID MR", "ID Prueba", "Fecha Ejecución", "Sistema", "Tipo", "Responsable", "Objetivo"], [800, 900, 1600, 1200, 1260, 1200, 1400]),
                    makeRow(["MR-CBVP-2026-001", "PU-01A", "DD/MM/AAAA", "CBVP-Admisiones", "Unitaria", "Desarrollador Backend", "Validar Cls_Usuario::registrarUsuario() con campos válidos e inválidos"], [800, 900, 1600, 1200, 1260, 1200, 1400]),
                    makeRow(["MR-CBVP-2026-001", "PF-02A", "DD/MM/AAAA", "CBVP-Admisiones", "Funcional", "Desarrollador Backend", "Verificar flujo 2FA en Ctrllr_Login con credenciales correctas e incorrectas"], [800, 900, 1600, 1200, 1260, 1200, 1400], true),
                    makeHeaderRow(["Elemento a Probar", "Precondición", "Datos de Entrada", "Resultados Esperados", "Resultados Obtenidos", "Estado", "Observaciones"], [800, 900, 1600, 1200, 1260, 1200, 1400]),
                    makeRow(["Cls_Usuario::registrarUsuario()", "BD activa; módulo RPA disponible", "DNI: 12345678, Email: test@cbvp.pe, Pass: Cbvp@2026", "Registro exitoso, hash bcrypt almacenado, id_usuario retornado", "Registro creado correctamente en BD con hash válido", "Aprobado", "Ninguna"], [800, 900, 1600, 1200, 1260, 1200, 1400]),
                    makeRow(["Ctrllr_Login::login() con 2FA", "Usuario registrado con TOTP activo", "Usuario: test@cbvp.pe, TOTP: código válido vigente", "Redirección al Dashboard según rol del usuario", "Redirección correcta al Dashboard de Postulante", "Aprobado", "Verificado en Chrome y Firefox"], [800, 900, 1600, 1200, 1260, 1200, 1400], true),
                ]
            }),
            new Paragraph({ spacing: { after: 120 } }),
            heading2("8.4. Salidas de la Fase 3"),
            ...[
                "Listado de elementos modificados (Plantilla 8)",
                "Código fuente modificado (commit en repositorio Git)",
                "Registro de Pruebas sobre el Sistema Modificado (Plantilla 9)",
                "Registro del Historial del MR actualizado"
            ].map(t => bullet(t)),
            pageBreak(),

            // ====== CAPÍTULO 9: FASE 4 ======
            heading1("CAPÍTULO 9: FASE 4 – ACEPTACIÓN Y REVISIÓN DEL MANTENIMIENTO"),
            heading2("9.1. Entradas"),
            bullet("Software modificado con los cambios aplicados"),
            bullet("Registro de Pruebas sobre el Sistema Modificado (Plantilla 9)"),
            heading2("9.2. Actividad 1: Revisiones y Aprobación"),
            heading3("Tarea 1"),
            para("El Mantenedor se reunirá con el Coordinador de Desarrollo de la Dirección de Sistemas del CBVP para ejecutar pruebas de aceptación sobre el sistema modificado, verificando que todos los requerimientos funcionales afectados operen correctamente."),
            heading3("Plantilla 10 – Registro de Pruebas de Aprobación"),
            new Table({
                width: { size: 8360, type: WidthType.DXA },
                columnWidths: [800, 900, 1600, 1800, 1060, 1000, 1200],
                rows: [
                    makeHeaderRow(["ID MR", "ID Prueba", "Fecha Ejecución", "Sistema", "Tipo", "Responsable", "Estado"], [800, 900, 1600, 1800, 1060, 1000, 1200]),
                    makeRow(["MR-CBVP-2026-001", "PA-01A", "DD/MM/AAAA", "Sistema Admisiones CBVP", "Aceptación", "QA Tester", "Aprobado"], [800, 900, 1600, 1800, 1060, 1000, 1200]),
                    makeHeaderRow(["Objetivo:", "", "", "", "", "", ""], [800, 900, 1600, 1800, 1060, 1000, 1200]),
                    makeRow(["Validar end-to-end que las modificaciones implementadas funcionan correctamente desde la interfaz de usuario, sin regresiones en funcionalidades previas, y que el sistema queda listo para producción.", "", "", "", "", "", ""], [800, 900, 1600, 1800, 1060, 1000, 1200]),
                    makeHeaderRow(["Elemento a Probar", "Precondición", "Datos de Entrada", "Resultados Esperados", "Resultados Obtenidos", "Estado", "Observaciones"], [800, 900, 1600, 1800, 1060, 1000, 1200]),
                    makeRow(["Flujo de registro de nuevo postulante vía UI", "Sistema en staging con BD limpia; CAPTCHA habilitado", "Datos completos: nombre, DNI, email y contraseña válidos", "Mensaje de éxito y correo de confirmación enviado al usuario", "Registro exitoso; mensaje Toast de éxito mostrado en pantalla", "Aprobado", "Validado en Chrome 125 y Firefox 126"], [800, 900, 1600, 1800, 1060, 1000, 1200]),
                    makeRow(["Generación de informe PDF por módulo RPA tras cierre de proceso", "Proceso de admisión cerrado con resultados registrados en BD", "Cierre de proceso activo con 10 postulantes evaluados", "PDF oficial generado automáticamente y disponible para descarga", "PDF generado con datos correctos; descarga funcional desde 'Procesos Finalizados'", "Aprobado", "Se verificó integridad del PDF con datos de las 10 inscripciones"], [800, 900, 1600, 1800, 1060, 1000, 1200], true),
                ]
            }),
            new Paragraph({ spacing: { after: 120 } }),
            heading2("9.3. Salidas de la Fase 4"),
            ...[
                "Sistema modificado y validado",
                "Registro de Pruebas de Aprobación (Plantilla 10)",
                "Registro del Análisis de la Solicitud de Modificación actualizado"
            ].map(t => bullet(t)),
            pageBreak(),

            // ====== CAPÍTULO 10: PRUEBAS POR RF ======
            heading1("CAPÍTULO 10: DETALLE DE PRUEBAS POR REQUERIMIENTO FUNCIONAL"),
            para("A continuación se presenta el detalle de cada prueba planificada, organizada por requerimiento funcional. Cada conjunto incluye prueba unitaria, funcional y de aceptación."),
            new Paragraph({ spacing: { after: 120 } }),

            ...allTests.reduce((acc, t, i) => {
                if (i === 0 || allTests[i - 1].rf !== t.rf) {
                    const rfTitles = {
                        "RF-01": "RF-01 – Registrar Usuario Postulante",
                        "RF-02": "RF-02 – Iniciar Sesión",
                        "RF-03": "RF-03 – Modificar Perfil",
                        "RF-04": "RF-04 – Gestionar Usuarios y Roles",
                        "RF-05": "RF-05 – Gestionar Procesos de Admisión",
                        "RF-06": "RF-06 – Autorizar Proceso con Firma Digital",
                        "RF-07": "RF-07 – Postular a un Proceso de Admisión",
                        "RF-08": "RF-08 – Cargar Documentos Requeridos",
                        "RF-09": "RF-09 – Asignar Grupo de Examen",
                        "RF-10": "RF-10 – Programar Examen",
                        "RF-11": "RF-11 – Registrar Resultados de Examen",
                        "RF-12": "RF-12 – Generar Informe Automático con RPA",
                    };
                    acc.push(heading2(rfTitles[t.rf] || t.rf));
                }
                const prioMap = {
                    "PU-01A": "Bajo", "PF-01A": "Medio", "PA-01A": "Muy alto",
                    "PU-02A": "Bajo", "PF-02A": "Medio", "PA-02A": "Muy alto",
                    "PU-03A": "Bajo", "PF-03A": "Medio", "PA-03A": "Muy alto",
                    "PU-04A": "Bajo", "PF-04A": "Alto", "PA-04A": "Muy alto",
                    "PU-05A": "Bajo", "PF-05A": "Alto", "PA-05A": "Muy alto",
                    "PU-06A": "Bajo", "PF-06A": "Medio", "PA-06A": "Muy alto",
                    "PU-07A": "Bajo", "PF-07A": "Medio", "PA-07A": "Muy alto",
                    "PU-08A": "Bajo", "PF-08A": "Medio", "PA-08A": "Muy alto",
                    "PU-09A": "Bajo", "PF-09A": "Medio", "PA-09A": "Muy alto",
                    "PU-10A": "Bajo", "PF-10A": "Alto", "PA-10A": "Muy alto",
                    "PU-11A": "Bajo", "PF-11A": "Medio", "PA-11A": "Muy alto",
                    "PU-12A": "Bajo", "PF-12A": "Alto", "PA-12A": "Muy alto",
                };
                const colorMap = { "Unitaria": "2C5F8A", "Funcional": "1F6B3A", "Aceptación": "7B3F00" };
                acc.push(
                    new Table({
                        width: { size: 8360, type: WidthType.DXA },
                        columnWidths: [1000, 1000, 1200, 1000, 4160],
                        rows: [
                            makeHeaderRow([t.id, t.tipo, "Prioridad: " + (prioMap[t.id] || "Medio"), t.etapa, t.resp], [1000, 1000, 1200, 1000, 4160], colorMap[t.tipo] || "1F3864"),
                            makeRow(["Objetivo:", t.obj, "", "", ""], [1000, 1000, 1200, 1000, 4160]),
                            makeHeaderRow(["Componente Evaluado", "Precondición", "Datos de Entrada", "Resultado Esperado", "Resultado Obtenido / Estado"], [1000, 1000, 1200, 1000, 4160]),
                            makeRow(["", "", "", "", ""], [1000, 1000, 1200, 1000, 4160]),
                        ]
                    }),
                    new Paragraph({ spacing: { after: 100 } })
                );
                return acc;
            }, []),

            pageBreak(),

            // ====== CAPÍTULO 11: FASE 5 – MIGRACIÓN ======
            heading1("CAPÍTULO 11: FASE 5 – MIGRACIÓN"),
            heading2("11.1. Descripción"),
            para("Esta fase aplica cuando el Sistema de Admisiones CBVP requiera ser trasladado a un nuevo entorno operativo, como migración de servidor, cambio de gestor de base de datos, actualización de versión de PHP, o integración con una nueva plataforma del servicio RPA."),
            heading2("11.2. Actividades"),
            ...[
                "Elaboración del Plan de Migración detallando entorno origen y destino.",
                "Notificación a todas las sedes afectadas y coordinación con la Junta Nacional.",
                "Exportación y respaldo completo de la base de datos MySQL (postulantes, procesos, documentos, resultados).",
                "Ejecución de la migración en entorno de pruebas antes del entorno productivo.",
                "Implementación de las operaciones de soporte y entrenamiento al equipo técnico.",
                "Notificación final a los usuarios y archivado de datos del sistema anterior."
            ].map(t => bullet(t)),
            heading2("11.3. Plantillas de Migración"),
            new Table({
                width: { size: 8360, type: WidthType.DXA },
                columnWidths: [1600, 1600, 1600, 1600, 1960],
                rows: [
                    makeHeaderRow(["Plantilla 11", "Plantilla 12", "Plantilla 13", "Plantilla 14", "Plantilla 15"], [1600, 1600, 1600, 1600, 1960]),
                    makeRow(["Listado de Elementos a Migrar", "Listado de Herramientas", "Conversión de Datos", "Verificación de Migración", "Registro Soporte Sistema Antiguo"], [1600, 1600, 1600, 1600, 1960]),
                ]
            }),
            new Paragraph({ spacing: { after: 200 } }),
            pageBreak(),

            // ====== CAPÍTULO 12: FASE 6 – RETIRO ======
            heading1("CAPÍTULO 12: FASE 6 – RETIRO DEL SOFTWARE"),
            heading2("12.1. Descripción"),
            para("Esta fase aplica cuando el Sistema de Admisiones CBVP llegue al final de su vida útil o sea reemplazado por una nueva plataforma institucional. El retiro debe realizarse de forma controlada garantizando la preservación de datos históricos de todos los procesos de admisión."),
            heading2("12.2. Actividades"),
            ...[
                "Elaboración del Plan de Retiro con cronograma de desactivación.",
                "Notificación formal a todas las sedes del CBVP y a los usuarios activos.",
                "Fijación de la fecha de finalización del servicio.",
                "Exportación y archivado seguro de todos los datos (postulantes, documentos, resultados de exámenes, informes RPA).",
                "Implementación de operaciones de entrenamiento para el sistema sucesor.",
                "Archivado final en repositorio institucional del CBVP."
            ].map(t => bullet(t)),
            heading3("Plantilla 16 – Registro del Retiro del Producto de Software"),
            new Table({
                width: { size: 8360, type: WidthType.DXA },
                columnWidths: [2000, 6360],
                rows: [
                    makeHeaderRow(["REGISTRO DEL RETIRO DEL PRODUCTO DE SOFTWARE", ""], [2000, 6360]),
                    makeRow(["Nombre del Sistema:", "Sistema Web de Gestión de Admisiones – CBVP"], [2000, 6360]),
                    makeRow(["Versión:", "1.0 (versión en retiro)"], [2000, 6360], true),
                    makeRow(["Fecha de inicio del retiro:", "DD/MM/AAAA"], [2000, 6360]),
                    makeRow(["Responsable:", "Director de Sistemas – Junta Nacional CBVP"], [2000, 6360], true),
                    makeRow(["Motivo del retiro:", "Reemplazo por nueva plataforma institucional con mayor escalabilidad; fin de vida útil del stack tecnológico (PHP puro / MVC sin framework)."], [2000, 6360]),
                    makeRow(["Datos archivados:", "Base de datos MySQL completa (postulantes, inscripciones, documentos, resultados, informes RPA); código fuente en repositorio Git institucional; copias de PDFs oficiales generados."], [2000, 6360], true),
                    makeRow(["Sistema sucesor:", "Nueva plataforma de admisiones CBVP (versión 2.0 – en desarrollo por la Dirección de Sistemas)"], [2000, 6360]),
                    makeRow(["Aprobado por:", "Director Nacional – Cuerpo de Bomberos Voluntarios del Perú"], [2000, 6360], true),
                    makeRow(["Fecha de cierre definitivo:", "DD/MM/AAAA"], [2000, 6360]),
                ]
            }),
            new Paragraph({ spacing: { after: 200 } }),
            pageBreak(),

            // ====== CAPÍTULO 13: RESUMEN ======
            heading1("CAPÍTULO 13: RESUMEN DEL PLAN – CORRESPONDENCIA RF / PRUEBAS"),
            para("La siguiente tabla presenta la trazabilidad completa entre requerimientos funcionales, tipos de prueba y fases del ciclo de mantenimiento del Sistema de Admisiones CBVP."),
            new Paragraph({ spacing: { after: 120 } }),
            new Table({
                width: { size: 8360, type: WidthType.DXA },
                columnWidths: [800, 2000, 800, 800, 800, 1000, 2160],
                rows: [
                    makeHeaderRow(["RF", "Título", "Unitaria", "Funcional", "Aceptación", "Prioridad", "Módulo/Componente Principal"], [800, 2000, 800, 800, 800, 1000, 2160]),
                    makeRow(["RF-01", "Registrar usuario", "PU-01A", "PF-01A", "PA-01A", "Medio-Alto", "Cls_Usuario / Ctrllr_Registro"], [800, 2000, 800, 800, 800, 1000, 2160]),
                    makeRow(["RF-02", "Iniciar sesión", "PU-02A", "PF-02A", "PA-02A", "Alto", "Cls_Usuario / Ctrllr_Login (2FA/TOTP)"], [800, 2000, 800, 800, 800, 1000, 2160], true),
                    makeRow(["RF-03", "Modificar Perfil", "PU-03A", "PF-03A", "PA-03A", "Medio", "Cls_Usuario / Ctrllr_Usuario"], [800, 2000, 800, 800, 800, 1000, 2160]),
                    makeRow(["RF-04", "Gestionar roles", "PU-04A", "PF-04A", "PA-04A", "Alto", "Cls_Usuario / superusuario"], [800, 2000, 800, 800, 800, 1000, 2160], true),
                    makeRow(["RF-05", "Gestionar procesos", "PU-05A", "PF-05A", "PA-05A", "Alto", "Cls_ProcesoAdmision / Ctrllr_ProcesoAdmision"], [800, 2000, 800, 800, 800, 1000, 2160]),
                    makeRow(["RF-06", "Firma digital", "PU-06A", "PF-06A", "PA-06A", "Medio", "Cls_Resolucion / Ctrllr_Resolucion"], [800, 2000, 800, 800, 800, 1000, 2160], true),
                    makeRow(["RF-07", "Postular", "PU-07A", "PF-07A", "PA-07A", "Medio", "Ctrllr_Inscripcion / Cls_InscripcionProceso"], [800, 2000, 800, 800, 800, 1000, 2160]),
                    makeRow(["RF-08", "Cargar documentos", "PU-08A", "PF-08A", "PA-08A", "Medio", "Cls_DocumentoPostulante (SHA-256)"], [800, 2000, 800, 800, 800, 1000, 2160], true),
                    makeRow(["RF-09", "Asignar grupo", "PU-09A", "PF-09A", "PA-09A", "Medio", "Ctrllr_GrupoProceso / Cls_InscripcionProceso"], [800, 2000, 800, 800, 800, 1000, 2160]),
                    makeRow(["RF-10", "Programar Examen", "PU-10A", "PF-10A", "PA-10A", "Alto", "Cls_HorarioExamen / Ctrllr_HorarioExamen"], [800, 2000, 800, 800, 800, 1000, 2160], true),
                    makeRow(["RF-11", "Registrar resultados", "PU-11A", "PF-11A", "PA-11A", "Medio", "Cls_EvaluacionPostulante / Ctrllr_EvaluacionFisica"], [800, 2000, 800, 800, 800, 1000, 2160]),
                    makeRow(["RF-12", "Generar informe RPA", "PU-12A", "PF-12A", "PA-12A", "Alto", "Módulo RPA / trama JSON / PDF oficial"], [800, 2000, 800, 800, 800, 1000, 2160], true),
                ]
            }),
            new Paragraph({ spacing: { after: 400 } }),
            new Paragraph({
                children: [new TextRun({ text: "— Fin del Plan de Mantenimiento de Software —", italic: true, size: 22, color: "777777" })],
                alignment: AlignmentType.CENTER
            })
        ]
    }]
});

Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync('Plan_Mantenimiento_CBVP_2026.docx', buffer);
    console.log('Document created successfully');
});