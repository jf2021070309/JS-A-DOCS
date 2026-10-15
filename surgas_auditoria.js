const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageBreak, LevelFormat, PageNumber, Footer, Header
} = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const headerBorder = { style: BorderStyle.SINGLE, size: 1, color: "2E4A8B" };
const headerBorders = { top: headerBorder, bottom: headerBorder, left: headerBorder, right: headerBorder };

function cell(text, w, isHeader = false, colspan = 1) {
  return new TableCell({
    borders: isHeader ? headerBorders : borders,
    width: { size: w, type: WidthType.DXA },
    columnSpan: colspan,
    shading: isHeader ? { fill: "2E4A8B", type: ShadingType.CLEAR } : { fill: "FFFFFF", type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: isHeader ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: [new TextRun({
        text,
        bold: isHeader,
        color: isHeader ? "FFFFFF" : "000000",
        size: isHeader ? 20 : 20,
        font: "Arial"
      })]
    })]
  });
}

function cellAlt(text, w, altColor = false) {
  return new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    shading: altColor ? { fill: "E8EEF7", type: ShadingType.CLEAR } : { fill: "FFFFFF", type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      children: [new TextRun({ text, size: 20, font: "Arial" })]
    })]
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: 28, font: "Arial", color: "2E4A8B" })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, bold: true, size: 24, font: "Arial", color: "2E4A8B" })]
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.JUSTIFIED,
    spacing: { before: 60, after: 60 },
    children: [new TextRun({
      text,
      size: opts.size || 22,
      font: "Arial",
      bold: opts.bold || false,
      italics: opts.italic || false,
      color: opts.color || "000000"
    })]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 22, font: "Arial" })]
  });
}

function space() {
  return new Paragraph({ children: [new TextRun({ text: "" })] });
}

// ---- DOCUMENT ----
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: "2E4A8B" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: "2E4A8B" },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 1 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [
      // ===== PORTADA =====
      space(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "UNIVERSIDAD PRIVADA DE TACNA", bold: true, size: 24, font: "Arial" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "FACULTAD DE INGENIERÍA", bold: true, size: 22, font: "Arial" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 400 },
        children: [new TextRun({ text: "ESCUELA DE INGENIERÍA DE SISTEMAS", bold: true, size: 22, font: "Arial" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 400, after: 200 },
        children: [new TextRun({ text: "Guía Práctica de Laboratorio", bold: true, size: 28, font: "Arial" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 400 },
        children: [new TextRun({ text: '"Taller Plan de Auditoría de Cumplimiento"', bold: true, size: 26, font: "Arial", italics: true })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 60 },
        children: [new TextRun({ text: "Que se presenta para el curso:", bold: true, size: 22, font: "Arial" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 200 },
        children: [new TextRun({ text: '"Auditoría de Sistemas"', size: 22, font: "Arial" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 60 },
        children: [new TextRun({ text: "Integrantes:", bold: true, size: 22, font: "Arial" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "Flores Quispe Jaime Elias  (2021070309)", size: 22, font: "Arial" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 400 },
        children: [new TextRun({ text: "Leyva Sardon Elvis Ronald  (2021072614)", size: 22, font: "Arial" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 60 },
        children: [new TextRun({ text: "Docente:", bold: true, size: 22, font: "Arial" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 400 },
        children: [new TextRun({ text: "Mtro. Omar Cueva Martínez", size: 22, font: "Arial" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 400 },
        children: [new TextRun({ text: "TACNA – PERÚ", bold: true, size: 22, font: "Arial" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 200 },
        children: [new TextRun({ text: "2025", bold: true, size: 22, font: "Arial" })]
      }),
      new Paragraph({ children: [new PageBreak()] }),

      // ===== PÁGINA 2: Título del plan =====
      space(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 100 },
        children: [new TextRun({ text: "AUDITORÍA DE CUMPLIMIENTO DE TECNOLOGÍAS DE LA INFORMACIÓN", bold: true, size: 26, font: "Arial" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 100 },
        children: [new TextRun({ text: "AL GRUPO SURGAS E.I.R.L.", bold: true, size: 26, font: "Arial" })]
      }),
      space(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 100, after: 100 },
        children: [new TextRun({ text: "Tacna, Tacna, Perú", size: 22, font: "Arial" })]
      }),
      space(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 100, after: 100 },
        children: [new TextRun({ text: '"Evaluación del Cumplimiento de los Sistemas de Información y Aplicaciones Tecnológicas del Grupo Surgas E.I.R.L."', bold: true, size: 24, font: "Arial", italics: true })]
      }),
      space(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 100, after: 100 },
        children: [new TextRun({ text: "Tacna, Septiembre de 2025", size: 22, font: "Arial" })]
      }),
      space(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: '"Decenio de la Igualdad de Oportunidades para Mujeres y Hombres"', size: 20, font: "Arial", italics: true })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: '"Año de la Esperanza y el Fortalecimiento de la Democracia"', size: 20, font: "Arial", italics: true })]
      }),
      new Paragraph({ children: [new PageBreak()] }),

      // ===== ÍNDICE =====
      h1("ÍNDICE"),
      p("I. ORIGEN ....................................................................................................... 4"),
      p("II. INFORMACIÓN DE LA ENTIDAD O DEPENDENCIA ........................................ 4"),
      p("III. DENOMINACIÓN DE LA MATERIA DE CONTROL .......................................... 5"),
      p("IV. ALCANCE ................................................................................................... 5"),
      p("V. OBJETIVOS .................................................................................................. 6"),
      p("VI. PROCEDIMIENTOS DE AUDITORÍA ............................................................... 7"),
      p("VII. PLAZO DE LA AUDITORÍA Y CRONOGRAMA ............................................. 10"),
      p("VIII. CRITERIOS DE AUDITORÍA ...................................................................... 11"),
      p("IX. INFORMACIÓN ADMINISTRATIVA ................................................................ 12"),
      p("X. DOCUMENTO A EMITIR ............................................................................... 12"),
      p("ANEXOS .......................................................................................................... 13"),
      new Paragraph({ children: [new PageBreak()] }),

      // ===== I. ORIGEN =====
      h1("I. ORIGEN"),
      p("La presente Auditoría de Cumplimiento de Tecnologías de la Información se origina en la necesidad de evaluar el estado actual y el grado de conformidad de los sistemas de información y aplicaciones tecnológicas implementadas por el Grupo Surgas E.I.R.L., empresa dedicada al envasado, distribución y comercialización de Gas Licuado de Petróleo (GLP) con presencia en las regiones de Tacna, Moquegua, Ilo y Arequipa. La empresa ha desarrollado e implementado un conjunto de soluciones tecnológicas para la gestión de pedidos, logística, atención al cliente y administración interna; sin embargo, no se ha realizado una evaluación formal que permita determinar si dichos sistemas operan con adecuados controles de calidad, seguridad, integridad de datos y continuidad del servicio."),
      space(),
      p("El origen específico de la auditoría radica en la ausencia de una revisión técnica sistemática sobre los sistemas de información en funcionamiento, incluyendo el sistema web de administración de pedidos, la aplicación móvil de clientes, la aplicación móvil de conductores, el software rastreador de vehículos, el sistema de puntos canjeables, el chatbot de WhatsApp para toma de pedidos, el sistema de reenvío de notificaciones de pagos y la API de registro de puntos de venta. La falta de documentación técnica actualizada, la ausencia de políticas formales de seguridad y la carencia de evaluaciones de riesgos tecnológicos constituyen señales de alerta que justifican la realización de este examen."),
      space(),
      p("Asimismo, la expansión geográfica de Surgas hacia nuevas regiones del sur del Perú y la mayor dependencia de canales digitales para la atención y distribución incrementan la exposición a riesgos tecnológicos. Por ello, la presente auditoría busca proporcionar una revisión estructurada, objetiva y técnicamente sustentada que permita identificar brechas, analizar causas y formular recomendaciones orientadas a fortalecer la seguridad, confiabilidad y eficiencia de los sistemas tecnológicos de la organización."),

      // ===== II. INFORMACIÓN DE LA ENTIDAD =====
      space(),
      h1("II. INFORMACIÓN DE LA ENTIDAD O DEPENDENCIA"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2500, 6526],
        rows: [
          new TableRow({ children: [cellAlt("Entidad / Dependencia", 2500, true), cellAlt("Grupo Surgas E.I.R.L.", 6526, false)] }),
          new TableRow({ children: [cellAlt("RUC", 2500, true), cellAlt("20601870224", 6526, false)] }),
          new TableRow({ children: [cellAlt("Tipo de empresa", 2500, true), cellAlt("Empresa Individual de Responsabilidad Limitada (E.I.R.L.)", 6526, false)] }),
          new TableRow({ children: [cellAlt("Nombre comercial", 2500, true), cellAlt("SURGAS", 6526, false)] }),
          new TableRow({ children: [cellAlt("Sector", 2500, true), cellAlt("Energía – Distribución de Gas Licuado de Petróleo (GLP)", 6526, false)] }),
          new TableRow({ children: [cellAlt("Nivel", 2500, true), cellAlt("Empresa privada – sector privado", 6526, false)] }),
          new TableRow({ children: [cellAlt("Domicilio fiscal", 2500, true), cellAlt("Av. Grau 292 Int. 601A, La Victoria – Lima", 6526, false)] }),
          new TableRow({ children: [cellAlt("Sede operativa Tacna", 2500, true), cellAlt("Av. Tarapacá 435, Urb. Los Delfines Mz. B Lt. 4, Tacna", 6526, false)] }),
          new TableRow({ children: [cellAlt("Fecha de inscripción", 2500, true), cellAlt("09 de febrero de 2017", 6526, false)] }),
          new TableRow({ children: [cellAlt("Ámbito de operaciones", 2500, true), cellAlt("Tacna, Moquegua, Ilo y Arequipa", 6526, false)] }),
          new TableRow({ children: [cellAlt("Actividades principales", 2500, true), cellAlt("Envasado, distribución y comercialización de GLP; gestión de red de distribuidores (Red Surgas); atención al cliente mediante canales digitales y presenciales", 6526, false)] }),
          new TableRow({ children: [cellAlt("Sistemas tecnológicos en operación", 2500, true), cellAlt("Sistema web de pedidos, app móvil de clientes, app móvil de conductores, rastreador GPS, sistema de puntos, chatbot WhatsApp, sistema de notificaciones de pagos, API de distribuidores, página web Radio Xero", 6526, false)] }),
        ]
      }),

      // ===== III. DENOMINACIÓN =====
      space(),
      h1("III. DENOMINACIÓN DE LA MATERIA DE CONTROL"),
      p('"Evaluación del Cumplimiento de los Sistemas de Información y Aplicaciones Tecnológicas del Grupo Surgas E.I.R.L."'),
      space(),
      p("La materia de control comprende los siguientes sistemas tecnológicos implementados en la organización:"),
      bullet("Sistema web de administración de pedidos"),
      bullet("Aplicación móvil de clientes"),
      bullet("Aplicación móvil de conductores"),
      bullet("Software rastreador de vehículos (GPS)"),
      bullet("Sistema web de tarjeta de puntos canjeables"),
      bullet("Sistema chatbot de WhatsApp para toma de pedidos"),
      bullet("Aplicación móvil de reenvío de notificaciones de pagos (administrador → empleados)"),
      bullet("API de puntos de venta (registro de distribuidores en base de datos)"),
      bullet("Página web de Radio Xero de Surgas"),

      // ===== IV. ALCANCE =====
      space(),
      h1("IV. ALCANCE"),
      p("La auditoría comprenderá la evaluación del cumplimiento, seguridad, integridad de datos y continuidad operativa de los sistemas de información y aplicaciones tecnológicas implementadas por el Grupo Surgas E.I.R.L. El examen abarcará la revisión del diseño y arquitectura de cada sistema, la documentación técnica disponible, los controles de acceso, los mecanismos de autenticación, el manejo de datos de clientes y operaciones, los procedimientos de respaldo, la disponibilidad de los servicios y la alineación con la normativa peruana aplicable en materia de protección de datos."),
      space(),
      p("En el sistema web de administración de pedidos se evaluará el control de accesos, la integridad de la información de pedidos, los registros de auditoría y la disponibilidad del servicio. Para las aplicaciones móviles (clientes y conductores) se revisarán los mecanismos de autenticación, la transmisión segura de datos, los permisos solicitados y la gestión de actualizaciones. En el software rastreador de vehículos se examinará la precisión de los datos, el control de acceso a la información de localización y la privacidad de los datos de los conductores. El sistema de puntos canjeables será evaluado respecto a la integridad de los saldos, la prevención de fraudes y los controles de canje. El chatbot de WhatsApp será revisado en cuanto a la seguridad del canal, el tratamiento de datos personales de los clientes y la continuidad del servicio. La API de puntos de venta será analizada en términos de autenticación, autorización y exposición de datos sensibles. Finalmente, la página web de Radio Xero será evaluada respecto a seguridad básica web y disponibilidad."),
      space(),
      p("La auditoría no incluirá una revisión del código fuente completo de los sistemas ni una auditoría financiera o contable, salvo en la medida en que los sistemas de información involucren transacciones económicas directas. El énfasis del examen estará en determinar si los sistemas implementados operan con controles adecuados y si cumplen con los estándares mínimos de seguridad y calidad esperados para una empresa de su giro y tamaño."),

      // ===== V. OBJETIVOS =====
      space(),
      h1("V. OBJETIVOS"),
      h2("Objetivo General"),
      p("Evaluar el grado de cumplimiento, seguridad y calidad de los sistemas de información y aplicaciones tecnológicas del Grupo Surgas E.I.R.L., con el fin de identificar brechas, riesgos tecnológicos y oportunidades de mejora que contribuyan a fortalecer la gestión digital de la empresa y la protección de los datos de sus clientes y operaciones."),
      space(),
      h2("Objetivos Específicos"),
      p("OE-01: Verificar si el sistema web de administración de pedidos y el chatbot de WhatsApp cuentan con controles de acceso adecuados, registros de auditoría y mecanismos que garanticen la integridad y disponibilidad de la información de pedidos.", { bold: false }),
      space(),
      p("OE-02: Evaluar si las aplicaciones móviles de clientes y conductores implementan mecanismos seguros de autenticación, transmisión cifrada de datos y gestión adecuada de permisos y actualizaciones.", { bold: false }),
      space(),
      p("OE-03: Determinar si el software rastreador de vehículos y la aplicación de notificaciones de pagos operan con controles de privacidad y acceso que protejan la información de conductores y empleados conforme a la normativa vigente.", { bold: false }),
      space(),
      p("OE-04: Comprobar si el sistema de puntos canjeables y la API de puntos de venta disponen de controles suficientes para garantizar la integridad de los datos, prevenir el fraude y proteger la información de distribuidores y clientes.", { bold: false }),
      space(),
      p("OE-05: Identificar las principales deficiencias y riesgos tecnológicos que afectan la seguridad, confiabilidad y continuidad de los sistemas de Surgas, formulando recomendaciones priorizadas para su corrección y mejora.", { bold: false }),

      // ===== VI. PROCEDIMIENTOS =====
      space(),
      h1("VI. PROCEDIMIENTOS DE AUDITORÍA"),
      p("Para el desarrollo de la auditoría de cumplimiento tecnológico del Grupo Surgas E.I.R.L. se aplicarán procedimientos orientados a obtener evidencia suficiente, competente, pertinente y útil para sustentar los hallazgos del examen. La metodología comprenderá la revisión documental de manuales técnicos, políticas de seguridad, contratos con proveedores de tecnología y registros operativos; entrevistas con los responsables de los sistemas; pruebas funcionales de los aplicativos; análisis de configuraciones de seguridad; y verificación del cumplimiento de la normativa aplicable."),
      space(),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [700, 400, 5126, 1600, 1200],
        rows: [
          new TableRow({
            children: [
              cell("Obj. Esp.", 700, true),
              cell("N°", 400, true),
              cell("Procedimiento de Auditoría", 5126, true),
              cell("Técnica", 1600, true),
              cell("Resp.", 1200, true),
            ]
          }),
          new TableRow({ children: [cellAlt("OE-01", 700, true), cellAlt("1", 400), cellAlt("Revisar la documentación técnica y funcional del sistema web de administración de pedidos, verificando la existencia de manuales de usuario, registros de incidencias, políticas de acceso y bitácoras del sistema.", 5126), cellAlt("Revisión documental", 1600), cellAlt("INT-01", 1200)] }),
          new TableRow({ children: [cellAlt("OE-01", 700, true), cellAlt("2", 400), cellAlt("Evaluar los mecanismos de autenticación y control de acceso del sistema web de pedidos y del chatbot de WhatsApp, verificando la existencia de credenciales diferenciadas, roles de usuario y registros de acceso.", 5126), cellAlt("Inspección técnica", 1600), cellAlt("INT-01", 1200)] }),
          new TableRow({ children: [cellAlt("OE-01", 700, true), cellAlt("3", 400), cellAlt("Verificar la disponibilidad y continuidad del sistema web de pedidos y del chatbot, revisando si existen procedimientos de respaldo, planes de contingencia y registros de interrupciones del servicio.", 5126), cellAlt("Revisión documental / Entrevista", 1600), cellAlt("INT-01", 1200)] }),
          new TableRow({ children: [cellAlt("OE-02", 700, true), cellAlt("4", 400), cellAlt("Revisar la configuración de las aplicaciones móviles de clientes y conductores, evaluando los mecanismos de autenticación (contraseñas, tokens, biometría), cifrado de datos en tránsito y en reposo, y gestión de sesiones.", 5126), cellAlt("Inspección técnica", 1600), cellAlt("INT-01", 1200)] }),
          new TableRow({ children: [cellAlt("OE-02", 700, true), cellAlt("5", 400), cellAlt("Verificar los permisos solicitados por las aplicaciones móviles de clientes y conductores en los dispositivos, determinando si son proporcionales a las funciones del aplicativo y si cumplen con las buenas prácticas de privacidad.", 5126), cellAlt("Inspección / Análisis funcional", 1600), cellAlt("INT-01", 1200)] }),
          new TableRow({ children: [cellAlt("OE-02", 700, true), cellAlt("6", 400), cellAlt("Entrevistar a los responsables del desarrollo y mantenimiento de las aplicaciones móviles para evaluar el proceso de gestión de actualizaciones, corrección de vulnerabilidades y comunicación de cambios a los usuarios.", 5126), cellAlt("Entrevista", 1600), cellAlt("INT-01", 1200)] }),
          new TableRow({ children: [cellAlt("OE-03", 700, true), cellAlt("7", 400), cellAlt("Evaluar el software rastreador de vehículos, verificando el control de acceso a la información de localización, la política de retención de datos de posicionamiento y la protección de la privacidad de los conductores.", 5126), cellAlt("Inspección técnica / Revisión documental", 1600), cellAlt("INT-02", 1200)] }),
          new TableRow({ children: [cellAlt("OE-03", 700, true), cellAlt("8", 400), cellAlt("Revisar la aplicación móvil de reenvío de notificaciones de pagos, evaluando los controles de acceso, la autenticación del administrador, el registro de notificaciones enviadas y la protección de la información financiera transmitida.", 5126), cellAlt("Inspección / Análisis documental", 1600), cellAlt("INT-02", 1200)] }),
          new TableRow({ children: [cellAlt("OE-04", 700, true), cellAlt("9", 400), cellAlt("Revisar el sistema web de tarjeta de puntos canjeables, verificando los controles de integridad del saldo, el registro de acumulación y canje, los mecanismos de prevención de manipulación y la validación de transacciones.", 5126), cellAlt("Análisis de datos / Inspección", 1600), cellAlt("INT-02", 1200)] }),
          new TableRow({ children: [cellAlt("OE-04", 700, true), cellAlt("10", 400), cellAlt("Evaluar la API de puntos de venta (distribuidores), revisando los mecanismos de autenticación y autorización de acceso, la validación de entradas, el manejo de errores, la exposición de datos sensibles y el registro de llamadas al servicio.", 5126), cellAlt("Inspección técnica / Análisis funcional", 1600), cellAlt("INT-02", 1200)] }),
          new TableRow({ children: [cellAlt("OE-04", 700, true), cellAlt("11", 400), cellAlt("Revisar la página web de Radio Xero de Surgas, evaluando la configuración de seguridad básica (HTTPS, cabeceras de seguridad, gestión de certificados), la disponibilidad del sitio y la protección frente a vulnerabilidades web comunes.", 5126), cellAlt("Inspección técnica", 1600), cellAlt("INT-02", 1200)] }),
          new TableRow({ children: [cellAlt("OE-05", 700, true), cellAlt("12", 400), cellAlt("Elaborar una matriz de brechas comparando la situación observada en cada sistema frente a los criterios de la norma ISO/IEC 27001:2022 y las buenas prácticas de OWASP aplicables, identificando incumplimientos, causas, riesgos e impactos.", 5126), cellAlt("Análisis comparativo", 1600), cellAlt("INT-01 / INT-02", 1200)] }),
          new TableRow({ children: [cellAlt("OE-05", 700, true), cellAlt("13", 400), cellAlt("Integrar los resultados obtenidos en cada procedimiento para formular observaciones de auditoría sustentadas, estableciendo la condición encontrada, el criterio aplicable, la causa del incumplimiento, el efecto generado y la recomendación correspondiente.", 5126), cellAlt("Análisis integral", 1600), cellAlt("INT-01 / INT-02", 1200)] }),
        ]
      }),
      space(),
      p("(*) Responsables: INT-01 = Auditor Principal; INT-02 = Auditor de Apoyo; TODOS = Equipo completo", { italic: true }),

      // ===== VII. CRONOGRAMA =====
      space(),
      h1("VII. PLAZO DE LA AUDITORÍA Y CRONOGRAMA"),
      p("El desarrollo de la Auditoría de Cumplimiento de TI del Grupo Surgas E.I.R.L. se llevará a cabo en un plazo total estimado de 42 días hábiles, distribuidos en las siguientes etapas:"),
      space(),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1600, 1400, 1400, 900, 3726],
        rows: [
          new TableRow({ children: [cell("Etapa", 1600, true), cell("Fecha Inicio", 1400, true), cell("Fecha Fin", 1400, true), cell("Duración", 900, true), cell("Descripción", 3726, true)] }),
          new TableRow({ children: [cellAlt("Planificación", 1600, true), cellAlt("15/09/2025", 1400), cellAlt("26/09/2025", 1400), cellAlt("10 días", 900), cellAlt("Elaboración del plan de auditoría, definición del alcance, criterios, objetivos, procedimientos, matriz de riesgos y cronograma detallado.", 3726)] }),
          new TableRow({ children: [cellAlt("Ejecución", 1600, true), cellAlt("29/09/2025", 1400), cellAlt("24/10/2025", 1400), cellAlt("20 días", 900), cellAlt("Aplicación de procedimientos, revisión técnica de sistemas, entrevistas, análisis de datos y elaboración de papeles de trabajo.", 3726)] }),
          new TableRow({ children: [cellAlt("Elaboración del Informe", 1600, true), cellAlt("27/10/2025", 1400), cellAlt("07/11/2025", 1400), cellAlt("10 días", 900), cellAlt("Redacción del informe de auditoría, conclusiones, recomendaciones y plan de acción correctivo.", 3726)] }),
          new TableRow({ children: [cellAlt("Presentación y cierre", 1600, true), cellAlt("10/11/2025", 1400), cellAlt("14/11/2025", 1400), cellAlt("5 días", 900), cellAlt("Comunicación formal de hallazgos a la Gerencia, levantamiento de descargos y firma del informe final.", 3726)] }),
        ]
      }),

      // ===== VIII. CRITERIOS =====
      space(),
      h1("VIII. CRITERIOS DE AUDITORÍA"),
      p("La auditoría se llevará a cabo con base en los siguientes criterios:"),
      space(),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2500, 6526],
        rows: [
          new TableRow({ children: [cell("Criterio / Norma", 2500, true), cell("Descripción", 6526, true)] }),
          new TableRow({ children: [cellAlt("ISO/IEC 27001:2022", 2500, true), cellAlt("Norma internacional para la gestión de la seguridad de la información. Constituye el criterio principal para evaluar los controles de seguridad implementados en los sistemas tecnológicos de Surgas.", 6526)] }),
          new TableRow({ children: [cellAlt("ISO/IEC 27002:2022", 2500, true), cellAlt("Guía de controles de seguridad de la información. Se utiliza como referencia complementaria para evaluar la aplicación de controles específicos en acceso, criptografía, operaciones y comunicaciones.", 6526)] }),
          new TableRow({ children: [cellAlt("OWASP Top 10 (2021)", 2500, true), cellAlt("Lista de los riesgos de seguridad más críticos en aplicaciones web y móviles. Se aplicará para evaluar vulnerabilidades en el sistema web de pedidos, chatbot, API de distribuidores y página web.", 6526)] }),
          new TableRow({ children: [cellAlt("Ley N.° 29733 (Perú)", 2500, true), cellAlt("Ley de Protección de Datos Personales. Aplicable a todos los sistemas que procesan datos personales de clientes, conductores y distribuidores de Surgas.", 6526)] }),
          new TableRow({ children: [cellAlt("D.S. N.° 003-2013-JUS", 2500, true), cellAlt("Reglamento de la Ley de Protección de Datos Personales. Establece los requisitos técnicos y organizativos para el tratamiento legítimo de datos.", 6526)] }),
          new TableRow({ children: [cellAlt("Políticas internas de la empresa", 2500, true), cellAlt("Procedimientos, manuales operativos y lineamientos internos del Grupo Surgas que regulen el uso, acceso y mantenimiento de los sistemas tecnológicos.", 6526)] }),
          new TableRow({ children: [cellAlt("Contratos con proveedores TI", 2500, true), cellAlt("Acuerdos suscritos con proveedores de desarrollo, hosting y mantenimiento de los sistemas, que establezcan compromisos de seguridad, disponibilidad y soporte.", 6526)] }),
        ]
      }),

      // ===== IX. INFORMACIÓN ADMINISTRATIVA =====
      space(),
      h1("IX. INFORMACIÓN ADMINISTRATIVA"),
      h2("IX.1 Comisión Auditora"),
      p("El equipo encargado de realizar la auditoría estará compuesto por profesionales con formación en tecnologías de la información, seguridad informática y auditoría de sistemas."),
      space(),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1400, 2200, 1600, 800, 800, 800, 1426],
        rows: [
          new TableRow({ children: [cell("Cargo", 1400, true), cell("Nombres y Apellidos", 2200, true), cell("Profesión", 1600, true), cell("Plan. (días)", 800, true), cell("Ejec. (días)", 800, true), cell("Inf. (días)", 800, true), cell("Total (días)", 1426, true)] }),
          new TableRow({ children: [cellAlt("Supervisor", 1400, true), cellAlt("Mtro. Omar Cueva Martínez", 2200), cellAlt("Mg. Auditoría TI", 1600), cellAlt("3", 800), cellAlt("8", 800), cellAlt("3", 800), cellAlt("14", 1426)] }),
          new TableRow({ children: [cellAlt("Auditor Principal", 1400, true), cellAlt("Flores Quispe Jaime Elias", 2200), cellAlt("Est. Ing. de Sistemas", 1600), cellAlt("5", 800), cellAlt("12", 800), cellAlt("5", 800), cellAlt("22", 1426)] }),
          new TableRow({ children: [cellAlt("Auditor de Apoyo", 1400, true), cellAlt("Leyva Sardon Elvis Ronald", 2200), cellAlt("Est. Ing. de Sistemas", 1600), cellAlt("3", 800), cellAlt("8", 800), cellAlt("3", 800), cellAlt("14", 1426)] }),
        ]
      }),
      space(),
      h2("IX.2 Costos Directos Estimados"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [400, 2000, 1000, 800, 1500, 1500, 1826],
        rows: [
          new TableRow({ children: [cell("N°", 400, true), cell("Miembro", 2000, true), cell("Nivel", 1000, true), cell("Días", 800, true), cell("Costo H/H (S/./día)", 1500, true), cell("Asignación (S/.)", 1500, true), cell("Costo Total (S/.)", 1826, true)] }),
          new TableRow({ children: [cellAlt("1", 400), cellAlt("Supervisor", 2000), cellAlt("Senior", 1000), cellAlt("14", 800), cellAlt("S/ 280.00", 1500), cellAlt("—", 1500), cellAlt("S/ 3,920.00", 1826)] }),
          new TableRow({ children: [cellAlt("2", 400), cellAlt("Auditor Principal", 2000), cellAlt("Junior", 1000), cellAlt("22", 800), cellAlt("S/ 150.00", 1500), cellAlt("—", 1500), cellAlt("S/ 3,300.00", 1826)] }),
          new TableRow({ children: [cellAlt("3", 400), cellAlt("Auditor de Apoyo", 2000), cellAlt("Medio", 1000), cellAlt("14", 800), cellAlt("S/ 200.00", 1500), cellAlt("—", 1500), cellAlt("S/ 2,800.00", 1826)] }),
          new TableRow({ children: [cellAlt("4", 400), cellAlt("Experto TI (*)", 2000), cellAlt("Experto", 1000), cellAlt("8", 800), cellAlt("S/ 350.00", 1500), cellAlt("S/ 400.00", 1500), cellAlt("S/ 3,200.00", 1826)] }),
          new TableRow({ children: [cell("TOTAL S/.", 9026, true, 6), cell("S/ 13,220.00", 1826, true)] }),
        ]
      }),
      space(),
      p("(*) Participación en fases donde se requiera experticia especializada en seguridad de aplicaciones y normativa de datos personales.", { italic: true }),

      // ===== X. DOCUMENTO A EMITIR =====
      space(),
      h1("X. DOCUMENTO A EMITIR"),
      p("Como resultado de la presente auditoría se emitirá un Informe de Auditoría de Cumplimiento de Tecnologías de la Información, el cual incluirá:"),
      bullet("Resumen ejecutivo con los principales hallazgos y conclusiones."),
      bullet("Análisis detallado del nivel de cumplimiento de cada sistema auditado respecto a los criterios establecidos."),
      bullet("Identificación de brechas (gap analysis) entre la situación actual y las buenas prácticas de seguridad y protección de datos."),
      bullet("Causas raíz de las deficiencias identificadas en cada sistema."),
      bullet("Recomendaciones específicas y priorizadas para cada sistema auditado."),
      bullet("Plan de acción correctivo con responsables, plazos y criterios de éxito."),
      bullet("Anexos técnicos con papeles de trabajo, matrices de riesgo y evidencias recopiladas."),
      space(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 60 },
        children: [new TextRun({ text: "Tacna, 15 de septiembre de 2025", size: 22, font: "Arial" })]
      }),
      space(), space(),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [3008, 3008, 3008],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
                width: { size: 3008, type: WidthType.DXA },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "____________________", size: 20, font: "Arial" })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Mtro. Omar Cueva Martínez", bold: true, size: 18, font: "Arial" })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Supervisor de Auditoría", size: 18, font: "Arial" })] }),
                ]
              }),
              new TableCell({
                borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
                width: { size: 3008, type: WidthType.DXA },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "____________________", size: 20, font: "Arial" })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Flores Quispe Jaime Elias", bold: true, size: 18, font: "Arial" })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Auditor Principal", size: 18, font: "Arial" })] }),
                ]
              }),
              new TableCell({
                borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
                width: { size: 3008, type: WidthType.DXA },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "____________________", size: 20, font: "Arial" })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Leyva Sardon Elvis Ronald", bold: true, size: 18, font: "Arial" })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Auditor de Apoyo", size: 18, font: "Arial" })] }),
                ]
              })
            ]
          })
        ]
      }),

      // ===== ANEXOS =====
      new Paragraph({ children: [new PageBreak()] }),
      h1("ANEXOS"),

      // Anexo 1: Matriz de Riesgos
      h2("Anexo 1: Matriz de Riesgos Tecnológicos"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [500, 1500, 1400, 1200, 700, 1400, 2326],
        rows: [
          new TableRow({ children: [cell("ID", 500, true), cell("Riesgo", 1500, true), cell("Causa", 1400, true), cell("Impacto", 1200, true), cell("Nivel", 700, true), cell("Control Existente", 1400, true), cell("Control Propuesto", 2326, true)] }),
          new TableRow({ children: [cellAlt("R-01", 500, true), cellAlt("Acceso no autorizado a sistemas web y API", 1500), cellAlt("Ausencia de políticas formales de control de acceso y gestión de credenciales", 1400), cellAlt("Robo de datos de clientes, manipulación de pedidos", 1200), cellAlt("ALTO", 700), cellAlt("Contraseñas básicas", 1400), cellAlt("Implementar autenticación multifactor y gestión de identidades centralizada", 2326)] }),
          new TableRow({ children: [cellAlt("R-02", 500, true), cellAlt("Exposición de datos personales de clientes y conductores", 1500), cellAlt("Falta de cifrado en transmisión/almacenamiento y ausencia de política de privacidad", 1400), cellAlt("Incumplimiento Ley N.° 29733, daño reputacional", 1200), cellAlt("ALTO", 700), cellAlt("No documentado", 1400), cellAlt("Implementar cifrado TLS en todas las comunicaciones y cifrado en reposo para datos sensibles", 2326)] }),
          new TableRow({ children: [cellAlt("R-03", 500, true), cellAlt("Vulnerabilidades en la API de puntos de venta", 1500), cellAlt("Falta de validación de entradas, ausencia de autenticación robusta", 1400), cellAlt("Manipulación del registro de distribuidores, fraude", 1200), cellAlt("ALTO", 700), cellAlt("Acceso por token básico", 1400), cellAlt("Aplicar estándar OAuth 2.0, validación de entradas y auditoría de llamadas a la API", 2326)] }),
          new TableRow({ children: [cellAlt("R-04", 500, true), cellAlt("Fraude en el sistema de puntos canjeables", 1500), cellAlt("Controles insuficientes en la lógica de acumulación y canje", 1400), cellAlt("Pérdidas económicas, desconfianza del cliente", 1200), cellAlt("ALTO", 700), cellAlt("Validación manual esporádica", 1400), cellAlt("Implementar auditoría automática de transacciones y alertas ante patrones anómalos", 2326)] }),
          new TableRow({ children: [cellAlt("R-05", 500, true), cellAlt("Interrupción del servicio de atención digital (chatbot, app)", 1500), cellAlt("Ausencia de planes de continuidad y respaldo de servicios cloud", 1400), cellAlt("Pérdida de ventas, insatisfacción de clientes", 1200), cellAlt("MEDIO", 700), cellAlt("Sin plan documentado", 1400), cellAlt("Elaborar plan de continuidad de TI y establecer acuerdos de nivel de servicio con proveedores", 2326)] }),
          new TableRow({ children: [cellAlt("R-06", 500, true), cellAlt("Privacidad de datos de geolocalización de conductores", 1500), cellAlt("Falta de política de retención y acceso a datos de rastreo GPS", 1400), cellAlt("Incumplimiento normativo, conflictos laborales", 1200), cellAlt("MEDIO", 700), cellAlt("Sin política documentada", 1400), cellAlt("Definir política de retención de datos GPS, control de acceso y consentimiento informado", 2326)] }),
        ]
      }),

      // Anexo 2: Mapeo de Sistemas
      space(),
      h2("Anexo 2: Mapeo de Sistemas Tecnológicos Auditados"),
      p("Los siguientes sistemas son objeto de la auditoría:"),
      bullet("Sistema web de pedidos: Recepción de pedido → Validación → Asignación a conductor → Despacho → Confirmación de entrega → Registro"),
      bullet("App móvil clientes: Registro → Autenticación → Pedido → Seguimiento en tiempo real → Confirmación → Calificación"),
      bullet("App móvil conductores: Autenticación → Recepción de ruta → Navegación → Confirmación de entrega → Reporte de incidencias"),
      bullet("Rastreador GPS: Activación → Transmisión de posición → Visualización en plataforma → Registro histórico → Alertas"),
      bullet("Sistema de puntos: Compra → Acumulación de puntos → Consulta de saldo → Solicitud de canje → Validación → Redención"),
      bullet("Chatbot WhatsApp: Recepción de mensaje → Interpretación → Respuesta automática → Escalado a operador (si aplica) → Registro"),
      bullet("App notificaciones pagos: Pago registrado → Notificación al administrador → Reenvío a empleados asignados → Confirmación de lectura"),
      bullet("API puntos de venta: Solicitud de registro → Autenticación → Validación de datos → Almacenamiento en BD → Respuesta al sistema solicitante"),
      bullet("Página web Radio Xero: Acceso público → Reproducción de contenido → Gestión de contenido (backend) → Publicación"),

      // Anexo 3: Normativa
      space(),
      h2("Anexo 3: Lista de Normativa Aplicable"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2500, 6526],
        rows: [
          new TableRow({ children: [cell("Norma / Estándar", 2500, true), cell("Descripción", 6526, true)] }),
          new TableRow({ children: [cellAlt("ISO/IEC 27001:2022", 2500, true), cellAlt("Sistemas de gestión de la seguridad de la información – Requisitos.", 6526)] }),
          new TableRow({ children: [cellAlt("ISO/IEC 27002:2022", 2500, true), cellAlt("Controles de seguridad de la información – Guía de buenas prácticas.", 6526)] }),
          new TableRow({ children: [cellAlt("OWASP Top 10 (2021)", 2500, true), cellAlt("Referencia internacional para los riesgos de seguridad más críticos en aplicaciones web y móviles.", 6526)] }),
          new TableRow({ children: [cellAlt("Ley N.° 29733 (Perú)", 2500, true), cellAlt("Ley de Protección de Datos Personales.", 6526)] }),
          new TableRow({ children: [cellAlt("D.S. N.° 003-2013-JUS", 2500, true), cellAlt("Reglamento de la Ley de Protección de Datos Personales.", 6526)] }),
          new TableRow({ children: [cellAlt("Políticas internas Surgas", 2500, true), cellAlt("Procedimientos y manuales operativos internos del Grupo Surgas E.I.R.L.", 6526)] }),
          new TableRow({ children: [cellAlt("Contratos con proveedores TI", 2500, true), cellAlt("Acuerdos de nivel de servicio y contratos de desarrollo/hosting.", 6526)] }),
        ]
      }),

      // Anexo 4: Instrumentos
      space(),
      h2("Anexo 4: Instrumentos de Recolección de Evidencia"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2800, 6226],
        rows: [
          new TableRow({ children: [cell("Instrumento", 2800, true), cell("Descripción", 6226, true)] }),
          new TableRow({ children: [cellAlt("Lista de verificación – Sistemas web y chatbot", 2800, true), cellAlt("Verifica controles de acceso, integridad de datos, disponibilidad y registros de auditoría del sistema de pedidos y chatbot.", 6226)] }),
          new TableRow({ children: [cellAlt("Lista de verificación – Apps móviles", 2800, true), cellAlt("Evalúa autenticación, cifrado, permisos y gestión de actualizaciones de las apps de clientes y conductores.", 6226)] }),
          new TableRow({ children: [cellAlt("Lista de verificación – Rastreador GPS y notificaciones", 2800, true), cellAlt("Revisa privacidad de datos de localización, control de acceso y manejo de información financiera.", 6226)] }),
          new TableRow({ children: [cellAlt("Lista de verificación – API y sistema de puntos", 2800, true), cellAlt("Verifica autenticación de API, integridad de saldos, prevención de fraude y exposición de datos.", 6226)] }),
          new TableRow({ children: [cellAlt("Guía de entrevistas", 2800, true), cellAlt("Cuestionario estructurado para responsables de cada sistema: desarrollador/proveedor TI, administrador de sistemas, responsable de operaciones.", 6226)] }),
          new TableRow({ children: [cellAlt("Plantilla de revisión documental", 2800, true), cellAlt("Formulario para registrar la revisión de manuales técnicos, contratos, políticas y registros operativos.", 6226)] }),
          new TableRow({ children: [cellAlt("Matriz de cumplimiento ISO 27001", 2800, true), cellAlt("Matriz con los controles relevantes de ISO/IEC 27001:2022, nivel de implementación (0-3) y observaciones del auditor por cada sistema.", 6226)] }),
        ]
      }),

      // Anexo 5: Trazabilidad
      space(),
      h2("Anexo 5: Mapeo de Objetivos, Procedimientos y Evidencias (Matriz de Trazabilidad)"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [700, 1400, 2800, 2426, 1700],
        rows: [
          new TableRow({ children: [cell("Obj. Esp.", 700, true), cell("Procedimientos Asociados", 1400, true), cell("Evidencias Requeridas", 2800, true), cell("Fuentes de Información", 2426, true), cell("Resp.", 1700, true)] }),
          new TableRow({ children: [cellAlt("OE-01", 700, true), cellAlt("Proc. 1, 2, 3", 1400), cellAlt("Documentación técnica del sistema web; registros de acceso; bitácoras del sistema; configuración del chatbot WhatsApp", 2800), cellAlt("Responsable TI de Surgas; proveedor del sistema", 2426), cellAlt("INT-01", 1700)] }),
          new TableRow({ children: [cellAlt("OE-02", 700, true), cellAlt("Proc. 4, 5, 6", 1400), cellAlt("Configuración de las apps móviles; permisos en tiendas (Play Store/App Store); actas de entrevistas", 2800), cellAlt("Desarrollador de apps; responsable TI", 2426), cellAlt("INT-01", 1700)] }),
          new TableRow({ children: [cellAlt("OE-03", 700, true), cellAlt("Proc. 7, 8", 1400), cellAlt("Configuración del rastreador GPS; registros de acceso a datos de posición; registros de notificaciones enviadas", 2800), cellAlt("Proveedor GPS; administrador del sistema", 2426), cellAlt("INT-02", 1700)] }),
          new TableRow({ children: [cellAlt("OE-04", 700, true), cellAlt("Proc. 9, 10, 11", 1400), cellAlt("Registros de acumulación/canje de puntos; documentación de la API; configuración de seguridad de la web Radio Xero", 2800), cellAlt("Responsable TI; desarrollador API; proveedor web", 2426), cellAlt("INT-02", 1700)] }),
          new TableRow({ children: [cellAlt("OE-05", 700, true), cellAlt("Proc. 12, 13", 1400), cellAlt("Resultados de todos los procedimientos anteriores; documentación de sistemas implementados", 2800), cellAlt("Todos los responsables de sistema", 2426), cellAlt("INT-01 / INT-02", 1700)] }),
        ]
      }),

      // Anexo 6: Cronograma detallado
      space(),
      h2("Anexo 6: Cronograma Detallado por Actividad y Responsable"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1400, 2800, 1500, 1726, 1600],
        rows: [
          new TableRow({ children: [cell("Semana", 1400, true), cell("Actividad", 2800, true), cell("Responsable", 1500, true), cell("Producto Intermedio", 1726, true), cell("Indicador de Avance", 1600, true)] }),
          new TableRow({ children: [cellAlt("Sem. 1 (15–19 sep)", 1400, true), cellAlt("Elaboración y aprobación del Plan de Auditoría", 2800), cellAlt("Supervisor / Auditor Principal", 1500), cellAlt("Plan de auditoría aprobado", 1726), cellAlt("Plan firmado", 1600)] }),
          new TableRow({ children: [cellAlt("Sem. 1 (15–19 sep)", 1400, true), cellAlt("Diseño de instrumentos: listas de verificación, guías de entrevista", 2800), cellAlt("Auditores", 1500), cellAlt("Instrumentos validados", 1726), cellAlt("Instrumentos listos", 1600)] }),
          new TableRow({ children: [cellAlt("Sem. 2 (22–26 sep)", 1400, true), cellAlt("Elaboración de matrices de riesgo y trazabilidad; solicitud formal de información a Surgas", 2800), cellAlt("Auditor Principal", 1500), cellAlt("Matrices y carta de solicitud", 1726), cellAlt("Acuse de recibo", 1600)] }),
          new TableRow({ children: [cellAlt("Sem. 3 (29 sep–03 oct)", 1400, true), cellAlt("Revisión técnica – Sistema web de pedidos, chatbot y app de clientes", 2800), cellAlt("INT-01", 1500), cellAlt("Papeles de trabajo", 1726), cellAlt("Evidencias registradas", 1600)] }),
          new TableRow({ children: [cellAlt("Sem. 4 (06–10 oct)", 1400, true), cellAlt("Revisión técnica – App conductores, rastreador GPS y app de notificaciones", 2800), cellAlt("INT-02", 1500), cellAlt("Papeles de trabajo", 1726), cellAlt("Evidencias registradas", 1600)] }),
          new TableRow({ children: [cellAlt("Sem. 5 (13–17 oct)", 1400, true), cellAlt("Revisión técnica – API de distribuidores, sistema de puntos y web Radio Xero", 2800), cellAlt("INT-01 / INT-02", 1500), cellAlt("Papeles de trabajo", 1726), cellAlt("Evidencias registradas", 1600)] }),
          new TableRow({ children: [cellAlt("Sem. 5–6 (13–24 oct)", 1400, true), cellAlt("Entrevistas a personal clave de TI y proveedores", 2800), cellAlt("Todos", 1500), cellAlt("Actas de entrevista", 1726), cellAlt("Actas firmadas", 1600)] }),
          new TableRow({ children: [cellAlt("Sem. 6 (20–24 oct)", 1400, true), cellAlt("Análisis de datos y elaboración de gap analysis; hallazgos preliminares", 2800), cellAlt("Auditor Principal + Experto", 1500), cellAlt("Borrador de hallazgos", 1726), cellAlt("Borrador revisado", 1600)] }),
          new TableRow({ children: [cellAlt("Sem. 7 (27–31 oct)", 1400, true), cellAlt("Comunicación de hallazgos y descargos de la empresa", 2800), cellAlt("Supervisor / Auditor Principal", 1500), cellAlt("Acta de comunicación", 1726), cellAlt("Respuestas recibidas", 1600)] }),
          new TableRow({ children: [cellAlt("Sem. 8–9 (03–14 nov)", 1400, true), cellAlt("Redacción del informe final y plan de acción correctivo", 2800), cellAlt("Todos", 1500), cellAlt("Informe final de auditoría", 1726), cellAlt("Informe aprobado", 1600)] }),
        ]
      }),
      space(),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('Plan_Auditoria_Surgas.docx', buffer);
  console.log('Documento creado exitosamente.');
});
