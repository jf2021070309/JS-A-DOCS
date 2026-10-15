const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, PageBreak, LevelFormat, Header, Footer,
  TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

// Colors
const BLUE_HEADER = "2E75B6";
const LIGHT_BLUE = "D5E8F0";
const DARK_BLUE = "1F4E79";
const ORANGE = "C55A11";
const GREEN_DARK = "375623";
const GREEN_LIGHT = "E2EFDA";
const GRAY_LIGHT = "F2F2F2";
const GRAY_MED = "D9D9D9";
const WHITE = "FFFFFF";
const YELLOW_LIGHT = "FFF2CC";
const ORANGE_LIGHT = "FCE4D6";
const RED_LIGHT = "FFCCCC";

const border = { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 160 },
    children: [new TextRun({ text, bold: true, size: 32, color: DARK_BLUE, font: "Arial" })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: 26, color: BLUE_HEADER, font: "Arial" })]
  });
}

function h3(text) {
  return new Paragraph({
    spacing: { before: 180, after: 80 },
    children: [new TextRun({ text, bold: true, size: 24, color: "333333", font: "Arial" })]
  });
}

function body(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    alignment: opts.justify ? AlignmentType.JUSTIFIED : AlignmentType.LEFT,
    children: [new TextRun({ text, size: 22, font: "Arial", ...opts })]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 22, font: "Arial" })]
  });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function emptyLine() {
  return new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun("")] });
}

function cell(text, opts = {}) {
  const fill = opts.fill || WHITE;
  const bold = opts.bold || false;
  const color = opts.color || "000000";
  const align = opts.align || AlignmentType.LEFT;
  const width = opts.width || 2340;
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text: String(text), bold, size: 20, font: "Arial", color })]
    })]
  });
}

function headerCell(text, width = 2340) {
  return cell(text, { fill: BLUE_HEADER, bold: true, color: WHITE, width });
}

function sectionTitle(text) {
  return new Paragraph({
    spacing: { before: 200, after: 100 },
    shading: { fill: BLUE_HEADER, type: ShadingType.CLEAR },
    children: [new TextRun({ text: `  ${text}`, bold: true, size: 26, color: WHITE, font: "Arial" })]
  });
}

// ===================== DOCUMENT =====================
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022",
          alignment: AlignmentType.LEFT,
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
        run: { size: 32, bold: true, font: "Arial", color: DARK_BLUE },
        paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: BLUE_HEADER },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE_HEADER, space: 1 } },
          children: [
            new TextRun({ text: "UNIVERSIDAD PRIVADA DE TACNA  |  Facultad de Ingeniería  |  Ing. de Sistemas", size: 18, font: "Arial", color: "666666" }),
            new TextRun({ text: "\t", size: 18 }),
            new TextRun({ text: "Caso Integrador Unidad 2 – Grupo SURGAS E.I.R.L.", size: 18, font: "Arial", color: BLUE_HEADER, bold: true })
          ],
          tabStops: [{ type: TabStopType.RIGHT, position: 9026 }]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          border: { top: { style: BorderStyle.SINGLE, size: 6, color: BLUE_HEADER, space: 1 } },
          alignment: AlignmentType.RIGHT,
          children: [
            new TextRun({ text: "Página ", size: 18, font: "Arial", color: "666666" }),
            new TextRun({ children: [PageNumber.CURRENT], size: 18, font: "Arial", color: "666666" }),
            new TextRun({ text: " / ", size: 18, font: "Arial", color: "666666" }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, font: "Arial", color: "666666" })
          ]
        })]
      })
    },
    children: [

      // ===================== PORTADA =====================
      emptyLine(), emptyLine(), emptyLine(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 400, after: 200 },
        children: [new TextRun({ text: "CASO INTEGRADOR UNIDAD 2", bold: true, size: 52, font: "Arial", color: DARK_BLUE })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 100, after: 80 },
        children: [new TextRun({ text: "GERENCIA DE TECNOLOGÍAS DE LA INFORMACIÓN", bold: true, size: 28, font: "Arial", color: BLUE_HEADER })]
      }),
      emptyLine(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 100 },
        children: [new TextRun({ text: "Empresa Analizada:", bold: true, size: 26, font: "Arial", color: "444444" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 200 },
        children: [new TextRun({ text: "GRUPO SURGAS E.I.R.L.", bold: true, size: 36, font: "Arial", color: ORANGE })]
      }),
      emptyLine(),
      new Table({
        width: { size: 7000, type: WidthType.DXA },
        columnWidths: [2500, 4500],
        rows: [
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 2500, type: WidthType.DXA }, shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "CURSO:", bold: true, size: 20, font: "Arial" })] })] }),
            new TableCell({ borders, width: { size: 4500, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Gerencia de Tecnologías de la Información", size: 20, font: "Arial" })] })] })
          ]}),
          new TableRow({ children: [
            new TableCell({ borders, shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR }, width: { size: 2500, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "DOCENTE:", bold: true, size: 20, font: "Arial" })] })] }),
            new TableCell({ borders, width: { size: 4500, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Mag. Ximena Gómez Valente", size: 20, font: "Arial" })] })] })
          ]}),
          new TableRow({ children: [
            new TableCell({ borders, shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR }, width: { size: 2500, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "EMPRESA:", bold: true, size: 20, font: "Arial" })] })] }),
            new TableCell({ borders, width: { size: 4500, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Grupo SURGAS E.I.R.L.", size: 20, font: "Arial" })] })] })
          ]}),
          new TableRow({ children: [
            new TableCell({ borders, shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR }, width: { size: 2500, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "FECHA:", bold: true, size: 20, font: "Arial" })] })] }),
            new TableCell({ borders, width: { size: 4500, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Junio 2026", size: 20, font: "Arial" })] })] })
          ]}),
        ]
      }),
      pageBreak(),

      // ===================== I. DATOS GENERALES =====================
      h1("I. DATOS GENERALES DE LA EMPRESA"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2800, 6226],
        rows: [
          new TableRow({ children: [headerCell("Razón Social", 2800), cell("Grupo SURGAS E.I.R.L.", { width: 6226 })] }),
          new TableRow({ children: [headerCell("RUC", 2800), cell("20601870224", { width: 6226 })] }),
          new TableRow({ children: [headerCell("Sector", 2800), cell("Energía / Distribución de Gas Licuado de Petróleo (GLP)", { width: 6226 })] }),
          new TableRow({ children: [headerCell("Actividad Principal", 2800), cell("Envasado, distribución y comercialización de Gas Licuado de Petróleo (GLP)", { width: 6226 })] }),
          new TableRow({ children: [headerCell("Dirección Fiscal", 2800), cell("Av. Grau 292 Int. 601A La Victoria – Lima - Lima", { width: 6226 })] }),
          new TableRow({ children: [headerCell("Dirección Operativa (Tacna)", 2800), cell("Av. Tarapacá 435, Urbanización Los Delfines, Tacna", { width: 6226 })] }),
          new TableRow({ children: [headerCell("Teléfono", 2800), cell("+51 952 XXX XXX (Call Center SURGAS)", { width: 6226 })] }),
          new TableRow({ children: [headerCell("Correo", 2800), cell("contacto@surgas.pe", { width: 6226 })] }),
          new TableRow({ children: [headerCell("N° Colaboradores", 2800), cell("Aprox. 30 colaboradores (distribución + administración + call center)", { width: 6226 })] }),
          new TableRow({ children: [headerCell("Inicio de Operaciones", 2800), cell("16 de octubre de 2014", { width: 6226 })] }),
        ]
      }),

      emptyLine(),

      h2("1. DESCRIPCIÓN DE LA EMPRESA"),
      body("Grupo SURGAS E.I.R.L. es una empresa tacneña fundada el 16 de octubre de 2014, dedicada al envasado, distribución y comercialización de Gas Licuado de Petróleo (GLP). Inició operaciones a pequeña escala en la ciudad de Tacna y fue ganando reconocimiento gracias a la calidad de su servicio y la confianza de los clientes locales.", { justify: true }),
      body("Con el paso del tiempo, la empresa implementó el modelo de emprendedores denominado 'Red SURGAS', que permite a los ciudadanos convertirse en distribuidores desde sus hogares, fortaleciendo la presencia de la marca y posibilitando la llegada a zonas más establecidas de la ciudad. En 2018 inauguró su planta envasadora en Calana, sumando más de 25 puntos de distribución en Tacna.", { justify: true }),
      body("En 2021 inició su expansión regional con operaciones en Ilo y Moquegua, y en 2025 ingresó a la región Arequipa. Actualmente SURGAS busca posicionarse como una empresa líder en distribución de GLP en el sur del Perú, reconocida por su innovación, seguridad y responsabilidad social.", { justify: true }),

      emptyLine(),

      h2("2. BIENES O SERVICIOS QUE BRINDA"),
      h3("Productos"),
      bullet("Gas Licuado de Petróleo (GLP) en balones de 5 kg, 10 kg y 45 kg"),
      bullet("Balones de GLP de uso doméstico e industrial"),
      bullet("GLP a granel para vehículos (GLP vehicular – en desarrollo)"),
      bullet("Accesorios de seguridad: mangueras, reguladores y válvulas certificadas"),

      h3("Servicios"),
      bullet("Distribución a domicilio de balones GLP mediante flota propia"),
      bullet("Distribución mediante Red SURGAS (distribuidores domiciliarios autorizados)"),
      bullet("Atención de pedidos a través de Call Center"),
      bullet("Asesoría en seguridad para el uso de GLP en hogares y comercios"),
      bullet("Envasado de GLP en planta propia ubicada en Calana, Tacna"),
      bullet("Venta en zonal principal y puntos de distribución estratégicos"),

      emptyLine(),

      h2("3. ORGANIGRAMA, FUNCIONES Y PERSONAL"),
      body("La estructura organizacional de SURGAS Tacna está compuesta por las siguientes áreas funcionales:", { justify: true }),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2500, 4026, 2500],
        rows: [
          new TableRow({ children: [headerCell("Área / Cargo", 2500), headerCell("Funciones Principales", 4026), headerCell("Personal Aprox.", 2500)] }),
          new TableRow({ children: [cell("Gerencia General", { width: 2500 }), cell("Dirección estratégica, toma de decisiones, relación con proveedores y reguladores.", { width: 4026 }), cell("1", { width: 2500, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Administración y Contabilidad", { width: 2500, fill: GRAY_LIGHT }), cell("Gestión financiera, planillas, compras, contabilidad y cumplimiento tributario (SUNAT).", { width: 4026, fill: GRAY_LIGHT }), cell("2", { width: 2500, align: AlignmentType.CENTER, fill: GRAY_LIGHT })] }),
          new TableRow({ children: [cell("Operaciones y Distribución", { width: 2500 }), cell("Recepción, envasado, despacho y control de balones GLP. Gestión de flota.", { width: 4026 }), cell("10–15", { width: 2500, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Call Center / Atención al Cliente", { width: 2500, fill: GRAY_LIGHT }), cell("Recepción de pedidos telefónicos y digitales, coordinación con reparto.", { width: 4026, fill: GRAY_LIGHT }), cell("3–5", { width: 2500, align: AlignmentType.CENTER, fill: GRAY_LIGHT })] }),
          new TableRow({ children: [cell("Red SURGAS (Distribuidores Externos)", { width: 2500 }), cell("Distribución domiciliaria desde hogares. Emprendedores autorizados en distritos.", { width: 4026 }), cell(">25 puntos", { width: 2500, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Sistemas y Tecnología (TI)", { width: 2500, fill: GRAY_LIGHT }), cell("Administración de sistemas de pedidos, facturación electrónica, redes y soporte.", { width: 4026, fill: GRAY_LIGHT }), cell("1–2", { width: 2500, align: AlignmentType.CENTER, fill: GRAY_LIGHT })] }),
        ]
      }),

      pageBreak(),

      // ===================== II. SÍNTESIS ESTRATÉGICA =====================
      h1("II. SÍNTESIS ESTRATÉGICA"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2000, 7026],
        rows: [
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, shading: { fill: BLUE_HEADER, type: ShadingType.CLEAR }, margins: { top: 120, bottom: 120, left: 120, right: 120 }, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "VISIÓN", bold: true, size: 24, font: "Arial", color: WHITE })] })] }),
            new TableCell({ borders, width: { size: 7026, type: WidthType.DXA }, margins: { top: 120, bottom: 120, left: 160, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Ser la empresa líder en distribución de gas en el sur del país, reconocida por su innovación, seguridad, responsabilidad social y compromiso con el bienestar de la comunidad.", size: 22, font: "Arial" })] })] })
          ]}),
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, shading: { fill: BLUE_HEADER, type: ShadingType.CLEAR }, margins: { top: 120, bottom: 120, left: 120, right: 120 }, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "MISIÓN", bold: true, size: 24, font: "Arial", color: WHITE })] })] }),
            new TableCell({ borders, width: { size: 7026, type: WidthType.DXA }, margins: { top: 120, bottom: 120, left: 160, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Brindar energía segura, accesible y confiable a los hogares, comercios e industrias del sur del Perú, contribuyendo al desarrollo regional mediante un servicio eficiente y de calidad.", size: 22, font: "Arial" })] })] })
          ]}),
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, shading: { fill: BLUE_HEADER, type: ShadingType.CLEAR }, margins: { top: 120, bottom: 120, left: 120, right: 120 }, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "VALORES", bold: true, size: 24, font: "Arial", color: WHITE })] })] }),
            new TableCell({ borders, width: { size: 7026, type: WidthType.DXA }, margins: { top: 120, bottom: 120, left: 160, right: 120 }, children: [
              new Paragraph({ children: [new TextRun({ text: "• Seguridad y prevención de riesgos en cada operación", size: 22, font: "Arial" })] }),
              new Paragraph({ children: [new TextRun({ text: "• Compromiso con la calidad del producto y el servicio", size: 22, font: "Arial" })] }),
              new Paragraph({ children: [new TextRun({ text: "• Cercanía y responsabilidad social con la comunidad", size: 22, font: "Arial" })] }),
              new Paragraph({ children: [new TextRun({ text: "• Honestidad y transparencia en cada transacción", size: 22, font: "Arial" })] }),
              new Paragraph({ children: [new TextRun({ text: "• Innovación y mejora continua en procesos", size: 22, font: "Arial" })] }),
            ] })
          ]}),
        ]
      }),

      pageBreak(),

      // ===================== III. BALANCED SCORECARD EMPRESARIAL =====================
      h1("III. BALANCED SCORECARD EMPRESARIAL"),
      h2("1. OBJETIVOS EMPRESARIALES"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1700, 3326, 2500, 1500],
        rows: [
          new TableRow({ children: [headerCell("Perspectiva", 1700), headerCell("Objetivo Estratégico (OKR)", 3326), headerCell("Indicadores (KPI)", 2500), headerCell("Meta 2025–2026", 1500)] }),
          new TableRow({ children: [cell("Financiera", { fill: LIGHT_BLUE, bold: true, width: 1700 }), cell("O1: Incrementar los ingresos por ventas de GLP en todas las regiones operativas.", { width: 3326 }), cell("Porcentaje de crecimiento de ventas anuales", { width: 2500 }), cell(">= 20%", { width: 1500, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Financiera", { fill: LIGHT_BLUE, bold: true, width: 1700 }), cell("O2: Reducir los costos operativos de distribución y envasado.", { width: 3326 }), cell("Porcentaje de costo operativo sobre ventas", { width: 2500 }), cell("<= 10%", { width: 1500, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Cliente", { fill: LIGHT_BLUE, bold: true, width: 1700 }), cell("O3: Incrementar la fidelidad y satisfacción de clientes domésticos e industriales.", { width: 3326 }), cell("Tasa de recompra mensual", { width: 2500 }), cell(">= 40%", { width: 1500, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Cliente", { fill: LIGHT_BLUE, bold: true, width: 1700 }), cell("O4: Consolidar una experiencia de atención ágil y confiable (call center + delivery).", { width: 3326 }), cell("Índice de Satisfacción del Cliente (CSAT)", { width: 2500 }), cell(">= 90%", { width: 1500, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Procesos Internos", { fill: LIGHT_BLUE, bold: true, width: 1700 }), cell("O5: Optimizar el control de stock de balones GLP y evitar quiebres de inventario.", { width: 3326 }), cell("Porcentaje de quiebre de stock", { width: 2500 }), cell("0%", { width: 1500, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Procesos Internos", { fill: LIGHT_BLUE, bold: true, width: 1700 }), cell("O6: Automatizar y agilizar los canales de atención y pedidos digitales.", { width: 3326 }), cell("Tiempo medio de respuesta al pedido", { width: 2500 }), cell("<= 5 min", { width: 1500, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Aprendizaje y Crecimiento", { fill: LIGHT_BLUE, bold: true, width: 1700 }), cell("O7: Elevar las competencias del personal en seguridad GLP y atención al cliente.", { width: 3326 }), cell("Horas de capacitación anual por colaborador", { width: 2500 }), cell(">= 40 h", { width: 1500, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Aprendizaje y Crecimiento", { fill: LIGHT_BLUE, bold: true, width: 1700 }), cell("O8: Fortalecer la infraestructura tecnológica para garantizar continuidad operativa.", { width: 3326 }), cell("Uptime del sistema de pedidos y facturación", { width: 2500 }), cell(">= 99%", { width: 1500, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Aprendizaje y Crecimiento", { fill: LIGHT_BLUE, bold: true, width: 1700 }), cell("O9: Fortalecer el clima laboral y el compromiso del equipo.", { width: 3326 }), cell("Índice de satisfacción laboral interno", { width: 2500 }), cell(">= 85%", { width: 1500, align: AlignmentType.CENTER })] }),
        ]
      }),

      emptyLine(),
      h2("2. MAPA ESTRATÉGICO"),
      body("El siguiente diagrama representa la cadena de valor estratégica de SURGAS, mostrando cómo las perspectivas de Aprendizaje y Crecimiento impulsan los Procesos Internos, que a su vez generan valor para el Cliente y se traducen en resultados Financieros.", { justify: true }),
      emptyLine(),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [9026],
        rows: [
          new TableRow({ children: [new TableCell({ borders, width: { size: 9026, type: WidthType.DXA }, shading: { fill: DARK_BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "PERSPECTIVA FINANCIERA", bold: true, size: 22, font: "Arial", color: WHITE })] })] })] }),
          new TableRow({ children: [new TableCell({ borders, width: { size: 9026, type: WidthType.DXA }, shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "O1: Incrementar ingresos por ventas GLP     |     O2: Reducir costos operativos de distribución", size: 20, font: "Arial" })] })] })] }),
          new TableRow({ children: [new TableCell({ borders, width: { size: 9026, type: WidthType.DXA }, shading: { fill: DARK_BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "PERSPECTIVA CLIENTE", bold: true, size: 22, font: "Arial", color: WHITE })] })] })] }),
          new TableRow({ children: [new TableCell({ borders, width: { size: 9026, type: WidthType.DXA }, shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "O3: Fidelizar clientes domésticos e industriales     |     O4: Atención ágil y confiable", size: 20, font: "Arial" })] })] })] }),
          new TableRow({ children: [new TableCell({ borders, width: { size: 9026, type: WidthType.DXA }, shading: { fill: DARK_BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "PERSPECTIVA PROCESOS INTERNOS", bold: true, size: 22, font: "Arial", color: WHITE })] })] })] }),
          new TableRow({ children: [new TableCell({ borders, width: { size: 9026, type: WidthType.DXA }, shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "O5: Control de stock de balones GLP     |     O6: Automatizar pedidos digitales", size: 20, font: "Arial" })] })] })] }),
          new TableRow({ children: [new TableCell({ borders, width: { size: 9026, type: WidthType.DXA }, shading: { fill: DARK_BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "PERSPECTIVA APRENDIZAJE Y CRECIMIENTO", bold: true, size: 22, font: "Arial", color: WHITE })] })] })] }),
          new TableRow({ children: [new TableCell({ borders, width: { size: 9026, type: WidthType.DXA }, shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "O7: Capacitar personal en seguridad GLP     |     O8: Fortalecer infraestructura TI     |     O9: Clima laboral óptimo", size: 20, font: "Arial" })] })] })] }),
        ]
      }),

      emptyLine(),
      h2("3. OBJETIVOS, INDICADORES E INICIATIVAS"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1400, 1800, 1600, 900, 1400, 1926],
        rows: [
          new TableRow({ children: [headerCell("Perspectiva", 1400), headerCell("Objetivo", 1800), headerCell("Indicador", 1600), headerCell("Unidad", 900), headerCell("Meta (Anual)", 1400), headerCell("Iniciativa Estratégica", 1926)] }),
          new TableRow({ children: [cell("Financiera", { fill: LIGHT_BLUE, width: 1400 }), cell("O1: Incrementar ingresos GLP en regiones.", { width: 1800 }), cell("Crecimiento de ventas", { width: 1600 }), cell("%", { width: 900, align: AlignmentType.CENTER }), cell(">= 20%", { width: 1400, align: AlignmentType.CENTER }), cell("Expansión de Red SURGAS en nuevos distritos y regiones sur.", { width: 1926 })] }),
          new TableRow({ children: [cell("Financiera", { fill: LIGHT_BLUE, width: 1400, fill: GRAY_LIGHT }), cell("O2: Reducir costos de distribución.", { width: 1800, fill: GRAY_LIGHT }), cell("Costo operativo s/ ventas", { width: 1600, fill: GRAY_LIGHT }), cell("%", { width: 900, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell("<= 10%", { width: 1400, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell("Optimizar rutas de reparto y negociar contratos de suministro.", { width: 1926, fill: GRAY_LIGHT })] }),
          new TableRow({ children: [cell("Cliente", { fill: LIGHT_BLUE, width: 1400 }), cell("O3: Fidelizar clientes domésticos.", { width: 1800 }), cell("Tasa de recompra mensual", { width: 1600 }), cell("%", { width: 900, align: AlignmentType.CENTER }), cell(">= 40%", { width: 1400, align: AlignmentType.CENTER }), cell("Programa de fidelización y beneficios para clientes frecuentes.", { width: 1926 })] }),
          new TableRow({ children: [cell("Cliente", { fill: LIGHT_BLUE, width: 1400, fill: GRAY_LIGHT }), cell("O4: Atención ágil call center + delivery.", { width: 1800, fill: GRAY_LIGHT }), cell("CSAT del cliente", { width: 1600, fill: GRAY_LIGHT }), cell("%", { width: 900, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell(">= 90%", { width: 1400, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell("Implementar sistema de gestión de pedidos integrado.", { width: 1926, fill: GRAY_LIGHT })] }),
          new TableRow({ children: [cell("Procesos Internos", { fill: LIGHT_BLUE, width: 1400 }), cell("O5: Control de stock de balones.", { width: 1800 }), cell("% quiebre de stock", { width: 1600 }), cell("%", { width: 900, align: AlignmentType.CENTER }), cell("0%", { width: 1400, align: AlignmentType.CENTER }), cell("Sistema ERP de inventarios en tiempo real para balones GLP.", { width: 1926 })] }),
          new TableRow({ children: [cell("Procesos Internos", { fill: LIGHT_BLUE, width: 1400, fill: GRAY_LIGHT }), cell("O6: Automatizar pedidos digitales.", { width: 1800, fill: GRAY_LIGHT }), cell("Tiempo medio respuesta pedido", { width: 1600, fill: GRAY_LIGHT }), cell("min", { width: 900, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell("<= 5 min", { width: 1400, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell("Chatbot de pedidos en WhatsApp y app móvil propia.", { width: 1926, fill: GRAY_LIGHT })] }),
          new TableRow({ children: [cell("Apr. y Crecim.", { fill: LIGHT_BLUE, width: 1400 }), cell("O7: Capacitar personal en seguridad.", { width: 1800 }), cell("Horas capacitación/colaborador", { width: 1600 }), cell("h", { width: 900, align: AlignmentType.CENTER }), cell(">= 40 h", { width: 1400, align: AlignmentType.CENTER }), cell("Plan continuo de capacitación en seguridad GLP y TI.", { width: 1926 })] }),
          new TableRow({ children: [cell("Apr. y Crecim.", { fill: LIGHT_BLUE, width: 1400, fill: GRAY_LIGHT }), cell("O8: Fortalecer infraestructura TI.", { width: 1800, fill: GRAY_LIGHT }), cell("Uptime sistemas críticos", { width: 1600, fill: GRAY_LIGHT }), cell("%", { width: 900, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell(">= 99%", { width: 1400, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell("Migrar sistemas a cloud y backups automáticos.", { width: 1926, fill: GRAY_LIGHT })] }),
          new TableRow({ children: [cell("Apr. y Crecim.", { fill: LIGHT_BLUE, width: 1400 }), cell("O9: Mejorar clima laboral.", { width: 1800 }), cell("Índice satisfacción laboral", { width: 1600 }), cell("%", { width: 900, align: AlignmentType.CENTER }), cell(">= 85%", { width: 1400, align: AlignmentType.CENTER }), cell("Encuestas trimestrales y programas de reconocimiento.", { width: 1926 })] }),
        ]
      }),

      pageBreak(),

      // ===================== IV. BSC DE TI =====================
      h1("IV. BALANCED SCORECARD DE TI"),
      h2("1. MAPA ESTRATÉGICO DE TI"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [9026],
        rows: [
          new TableRow({ children: [new TableCell({ borders, shading: { fill: DARK_BLUE, type: ShadingType.CLEAR }, width: { size: 9026, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "FINANCIERA TI", bold: true, size: 22, font: "Arial", color: WHITE })] })] })] }),
          new TableRow({ children: [new TableCell({ borders, shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR }, width: { size: 9026, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Maximizar ROI del canal digital de pedidos  |  Reducir costos TI con Cloud (SaaS/IaaS)", size: 20, font: "Arial" })] })] })] }),
          new TableRow({ children: [new TableCell({ borders, shading: { fill: DARK_BLUE, type: ShadingType.CLEAR }, width: { size: 9026, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "CLIENTE TI", bold: true, size: 22, font: "Arial", color: WHITE })] })] })] }),
          new TableRow({ children: [new TableCell({ borders, shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR }, width: { size: 9026, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "App/WhatsApp de pedidos en < 2 min  |  Elevar satisfacción usuarios internos con Helpdesk", size: 20, font: "Arial" })] })] })] }),
          new TableRow({ children: [new TableCell({ borders, shading: { fill: DARK_BLUE, type: ShadingType.CLEAR }, width: { size: 9026, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "PROCESOS INTERNOS TI", bold: true, size: 22, font: "Arial", color: WHITE })] })] })] }),
          new TableRow({ children: [new TableCell({ borders, shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR }, width: { size: 9026, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "ERP de inventario de balones en tiempo real  |  Chatbot de pedidos automatizado", size: 20, font: "Arial" })] })] })] }),
          new TableRow({ children: [new TableCell({ borders, shading: { fill: DARK_BLUE, type: ShadingType.CLEAR }, width: { size: 9026, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "APRENDIZAJE Y CRECIMIENTO TI", bold: true, size: 22, font: "Arial", color: WHITE })] })] })] }),
          new TableRow({ children: [new TableCell({ borders, shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR }, width: { size: 9026, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Certificar 100% personal en ciberseguridad  |  Alta disponibilidad 99%  |  Intranet colaborativa", size: 20, font: "Arial" })] })] })] }),
        ]
      }),

      emptyLine(),
      h2("2. OBJETIVOS, INDICADORES E INICIATIVAS DE TI"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1400, 3026, 2600, 2000],
        rows: [
          new TableRow({ children: [headerCell("Perspectiva", 1400), headerCell("Objetivo Estratégico TI (OKR)", 3026), headerCell("Indicadores (KPI)", 2600), headerCell("Meta 2025–2026", 2000)] }),
          new TableRow({ children: [cell("Financiera", { fill: LIGHT_BLUE, width: 1400 }), cell("TI-O1: Maximizar el ROI del canal digital de pedidos GLP (app + WhatsApp).", { width: 3026 }), cell("% incremento ventas procesadas digitalmente", { width: 2600 }), cell(">= 25% anual", { width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Financiera", { fill: LIGHT_BLUE, width: 1400, fill: GRAY_LIGHT }), cell("TI-O2: Reducir costos TI mediante cloud y automatización.", { width: 3026, fill: GRAY_LIGHT }), cell("% reducción costos mantenimiento correctivo", { width: 2600, fill: GRAY_LIGHT }), cell(">= 10% anual", { width: 2000, align: AlignmentType.CENTER, fill: GRAY_LIGHT })] }),
          new TableRow({ children: [cell("Cliente", { fill: LIGHT_BLUE, width: 1400 }), cell("TI-O3: Optimizar la velocidad y accesibilidad del canal digital de pedidos.", { width: 3026 }), cell("Tiempo medio respuesta automática (WhatsApp/App)", { width: 2600 }), cell("<= 2 minutos", { width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Cliente", { fill: LIGHT_BLUE, width: 1400, fill: GRAY_LIGHT }), cell("TI-O4: Proveer plataformas estables para la operación de call center y distribuidores.", { width: 3026, fill: GRAY_LIGHT }), cell("Índice satisfacción usuario interno con TI", { width: 2600, fill: GRAY_LIGHT }), cell(">= 90%", { width: 2000, align: AlignmentType.CENTER, fill: GRAY_LIGHT })] }),
          new TableRow({ children: [cell("Procesos", { fill: LIGHT_BLUE, width: 1400 }), cell("TI-O5: Automatizar control de inventario de balones en tiempo real.", { width: 3026 }), cell("% discrepancia stock sistema vs físico", { width: 2600 }), cell("0% en críticos", { width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Procesos", { fill: LIGHT_BLUE, width: 1400, fill: GRAY_LIGHT }), cell("TI-O6: Desplegar chatbot inteligente para automatizar pedidos y consultas.", { width: 3026, fill: GRAY_LIGHT }), cell("Tiempo medio respuesta inicial automatizada", { width: 2600, fill: GRAY_LIGHT }), cell("<= 2 minutos", { width: 2000, align: AlignmentType.CENTER, fill: GRAY_LIGHT })] }),
          new TableRow({ children: [cell("Apr. y Crec.", { fill: LIGHT_BLUE, width: 1400 }), cell("TI-O7: Capacitar al personal en uso seguro de sistemas y ciberseguridad.", { width: 3026 }), cell("% colaboradores certificados en TI/ciberseguridad", { width: 2600 }), cell("100%", { width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Apr. y Crec.", { fill: LIGHT_BLUE, width: 1400, fill: GRAY_LIGHT }), cell("TI-O8: Garantizar alta disponibilidad de sistemas críticos (pedidos, facturación).", { width: 3026, fill: GRAY_LIGHT }), cell("Tasa de disponibilidad (Uptime)", { width: 2600, fill: GRAY_LIGHT }), cell(">= 99%", { width: 2000, align: AlignmentType.CENTER, fill: GRAY_LIGHT })] }),
          new TableRow({ children: [cell("Apr. y Crec.", { fill: LIGHT_BLUE, width: 1400 }), cell("TI-O9: Facilitar comunicación interna digital entre áreas y distribuidores.", { width: 3026 }), cell("% adopción plataforma colaborativa interna", { width: 2600 }), cell(">= 85%", { width: 2000, align: AlignmentType.CENTER })] }),
        ]
      }),

      pageBreak(),

      // ===================== V. ANÁLISIS ORGANIZATIVO TI =====================
      h1("V. ANÁLISIS ORGANIZATIVO DEL ÁREA DE TI"),

      h2("1. ESTRUCTURA ORGANIZACIONAL: JERARQUÍA, ORGANIGRAMA, ROLES Y RESPONSABILIDADES"),
      body("El área de Sistemas y Tecnología de SURGAS es actualmente reducida y en proceso de formalización. La estructura refleja la etapa de crecimiento de la empresa:", { justify: true }),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2500, 4026, 2500],
        rows: [
          new TableRow({ children: [headerCell("Cargo", 2500), headerCell("Funciones Principales", 4026), headerCell("Cantidad", 2500)] }),
          new TableRow({ children: [cell("Jefe de TI / Responsable de Sistemas", { width: 2500 }), cell("Define estrategia TI, gestiona proyectos tecnológicos, coordina con Gerencia General, administra presupuesto y riesgos tecnológicos.", { width: 4026 }), cell("1", { width: 2500, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Administrador de Sistemas / Soporte", { fill: GRAY_LIGHT, width: 2500 }), cell("Administra servidores, redes, sistema de pedidos y facturación electrónica. Soporte técnico diario.", { fill: GRAY_LIGHT, width: 4026 }), cell("1", { fill: GRAY_LIGHT, width: 2500, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Especialista Digital / Canal Online", { width: 2500 }), cell("Gestiona redes sociales, WhatsApp Business, campañas digitales y analítica de pedidos en línea.", { width: 4026 }), cell("1 (parcial)", { width: 2500, align: AlignmentType.CENTER })] }),
        ]
      }),

      emptyLine(),
      h2("2. PROCESOS: SOPORTE, DESARROLLO, GESTIÓN DE PROYECTOS, SEGURIDAD"),

      h3("Soporte Tecnológico"),
      body("Asegura el correcto funcionamiento de los equipos y sistemas usados por call center, operaciones, administración y distribuidores Red SURGAS.", { justify: true }),
      bullet("Atención y resolución de incidencias de hardware y software"),
      bullet("Soporte al sistema de gestión de pedidos (call center y app)"),
      bullet("Soporte al sistema de facturación electrónica (SUNAT)"),
      bullet("Mantenimiento preventivo y correctivo de equipos"),
      bullet("Gestión de cuentas de usuario y accesos"),
      emptyLine(),

      h3("Desarrollo e Innovación Tecnológica"),
      body("Implementar mejoras tecnológicas que soporten el crecimiento hacia nuevas regiones y la digitalización de pedidos.", { justify: true }),
      bullet("Evaluación de nuevas tecnologías para distribución de GLP"),
      bullet("Integración entre sistema de pedidos, inventario y facturación"),
      bullet("Desarrollo de chatbot para pedidos automatizados"),
      bullet("Implementación de app móvil para clientes y distribuidores"),
      emptyLine(),

      h3("Gestión de Seguridad de la Información"),
      bullet("Administración de usuarios y permisos en sistemas críticos"),
      bullet("Implementación de políticas de seguridad informática"),
      bullet("Gestión de copias de seguridad automáticas"),
      bullet("Monitoreo de incidentes de seguridad"),
      bullet("Capacitación al personal en ciberseguridad y phishing"),
      emptyLine(),

      h2("3. RECURSOS HUMANOS: CANTIDAD, PERFILES, CAPACITACIÓN Y ROTACIÓN"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2500, 6526],
        rows: [
          new TableRow({ children: [new TableCell({ borders, shading: { fill: BLUE_HEADER, type: ShadingType.CLEAR }, colSpan: 2, width: { size: 9026, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Perfil del Personal de TI", bold: true, size: 22, font: "Arial", color: WHITE })] })] })] }),
          new TableRow({ children: [cell("Jefe de TI", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Profesional en Ingeniería de Sistemas o afines. Conocimientos en gestión de proyectos TI, infraestructura y seguridad de la información. Liderazgo y visión estratégica.", { width: 6526 })] }),
          new TableRow({ children: [cell("Administrador de Sistemas", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Técnico o universitario en Sistemas o Redes. Experiencia en administración de servidores, bases de datos, soporte técnico y mantenimiento.", { width: 6526 })] }),
          new TableRow({ children: [cell("Especialista Digital", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Conocimientos en marketing digital, gestión de redes sociales, WhatsApp Business API y análisis de métricas digitales.", { width: 6526 })] }),
        ]
      }),

      emptyLine(),
      h2("4. COMUNICACIÓN INTERNA Y CON OTRAS ÁREAS"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1800, 2426, 3000, 1800],
        rows: [
          new TableRow({ children: [headerCell("Área", 1800), headerCell("Objetivo de la Comunicación", 2426), headerCell("Actividades Principales", 3000), headerCell("Frecuencia", 1800)] }),
          new TableRow({ children: [cell("Call Center / Atención", { width: 1800 }), cell("Garantizar continuidad del sistema de pedidos.", { width: 2426 }), cell("Soporte sistema pedidos, incidencias, monitoreo de líneas.", { width: 3000 }), cell("Diaria", { width: 1800, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Operaciones / Distribución", { fill: GRAY_LIGHT, width: 1800 }), cell("Mantener integridad del inventario de balones.", { fill: GRAY_LIGHT, width: 2426 }), cell("Soporte ERP de inventario, actualización de stock, reportes logísticos.", { fill: GRAY_LIGHT, width: 3000 }), cell("Diaria", { fill: GRAY_LIGHT, width: 1800, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Administración y Contabilidad", { width: 1800 }), cell("Asegurar funcionamiento de facturación electrónica.", { width: 2426 }), cell("Soporte SUNAT, gestión de accesos, backups financieros.", { width: 3000 }), cell("Semanal", { width: 1800, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Gerencia General", { fill: GRAY_LIGHT, width: 1800 }), cell("Apoyar decisiones estratégicas con información TI.", { fill: GRAY_LIGHT, width: 2426 }), cell("Presentación KPIs TI, evaluación de riesgos, propuesta de proyectos.", { fill: GRAY_LIGHT, width: 3000 }), cell("Mensual", { fill: GRAY_LIGHT, width: 1800, align: AlignmentType.CENTER })] }),
        ]
      }),

      pageBreak(),

      // ===================== VI. SISTEMAS DE INFORMACIÓN =====================
      h1("VI. SISTEMAS DE INFORMACIÓN, NECESIDADES DE INFORMACIÓN Y CARTERA DE PROYECTOS"),

      h2("1. INVENTARIO DE SISTEMAS DE INFORMACIÓN"),
      body("El portafolio actual de SURGAS cubre los procesos básicos de gestión de pedidos y facturación, con limitada integración digital:", { justify: true }),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [4513, 1500],
        rows: [
          new TableRow({ children: [headerCell("Nombre del Sistema", 4513), headerCell("Código", 1500)] }),
          new TableRow({ children: [cell("Sistema web de pedidos", { width: 4513 }), cell("SI-001", { width: 1500, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Chatbot de WhatsApp", { fill: GRAY_LIGHT, width: 4513 }), cell("SI-002", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Sistema de puntos canjeables", { width: 4513 }), cell("SI-003", { width: 1500, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("API de puntos de venta", { fill: GRAY_LIGHT, width: 4513 }), cell("SI-004", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER })] }),
        ]
      }),

      emptyLine(),
      h3("Tabla de Evaluación de Sistemas de Información (Resumen Ejecutivo)"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2300, 2426, 2800, 1500],
        rows: [
          new TableRow({ children: [headerCell("Sistema", 2300), headerCell("Fortalezas", 2426), headerCell("Debilidades", 2800), headerCell("Puntaje", 1500)] }),
          new TableRow({ children: [cell("Sistema web de pedidos (SI-001)", { width: 2300 }), cell("Permite administrar los pedidos de manera centralizada.", { width: 2426 }), cell("Falta optimización móvil integral para los conductores.", { width: 2800 }), cell("3 / 5", { width: 1500, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Chatbot de WhatsApp (SI-002)", { fill: GRAY_LIGHT, width: 2300 }), cell("Automatización inicial de recepción, alta disponibilidad.", { fill: GRAY_LIGHT, width: 2426 }), cell("Interpretación limitada ante lenguaje natural muy informal.", { fill: GRAY_LIGHT, width: 2800 }), cell("4 / 5", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Sistema de puntos canjeables (SI-003)", { width: 2300 }), cell("Fideliza a clientes frecuentes con recompensas tangibles.", { width: 2426 }), cell("Riesgos de fraude si no se auditan los canjes correctamente.", { width: 2800 }), cell("3 / 5", { width: 1500, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("API de puntos de venta (SI-004)", { fill: GRAY_LIGHT, width: 2300 }), cell("Integra a los distribuidores a la base de datos principal.", { fill: GRAY_LIGHT, width: 2426 }), cell("Requiere mejorar los controles de seguridad y validación.", { fill: GRAY_LIGHT, width: 2800 }), cell("3 / 5", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER })] }),
        ]
      }),

      emptyLine(),
      h2("2. NECESIDADES DE INFORMACIÓN"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2200, 2826, 2500, 1500],
        rows: [
          new TableRow({ children: [headerCell("Gerencia / Área", 2200), headerCell("Necesidad de Información", 2826), headerCell("Objetivo Relacionado", 2500), headerCell("Uso", 1500)] }),
          new TableRow({ children: [cell("Operaciones y Distribución", { width: 2200 }), cell("Stock en tiempo real de balones GLP por tipo y punto de distribución.", { width: 2826 }), cell("O5: Optimizar control de inventario", { width: 2500 }), cell("Operativo", { width: 1500, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Call Center / Ventas", { fill: GRAY_LIGHT, width: 2200 }), cell("Tiempo de respuesta a pedidos, tasa de conversión y satisfacción del cliente.", { fill: GRAY_LIGHT, width: 2826 }), cell("O4/O6: Atención ágil omnicanal", { fill: GRAY_LIGHT, width: 2500 }), cell("Táctico", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Sistemas / TI", { width: 2200 }), cell("Disponibilidad y tiempo de actividad (Uptime) del sistema de pedidos y facturación.", { width: 2826 }), cell("O8: Continuidad TI (99%)", { width: 2500 }), cell("Operativo/Táctico", { width: 1500, align: AlignmentType.CENTER })] }),
        ]
      }),

      emptyLine(),
      h2("3. CARTERA DE PROYECTOS"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [3026, 1500, 4500],
        rows: [
          new TableRow({ children: [headerCell("Proyecto", 3026), headerCell("Tipo", 1500), headerCell("Origen / Sustento", 4500)] }),
          new TableRow({ children: [cell("Implementación de App Móvil de Pedidos SURGAS", { width: 3026 }), cell("Sistema", { width: 1500, align: AlignmentType.CENTER }), cell("Los pedidos se gestionan solo por teléfono o WhatsApp manual. No existe una app con carrito de pedidos, rastreo de entrega ni historial de compras.", { width: 4500 })] }),
          new TableRow({ children: [cell("Integración ERP Inventario GLP en Tiempo Real", { fill: GRAY_LIGHT, width: 3026 }), cell("Sistema", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("El inventario de balones no está sincronizado en tiempo real con el sistema de pedidos, generando riesgo de quiebres de stock.", { fill: GRAY_LIGHT, width: 4500 })] }),
          new TableRow({ children: [cell("Chatbot Inteligente para Pedidos (WhatsApp API)", { width: 3026 }), cell("Sistema", { width: 1500, align: AlignmentType.CENTER }), cell("La atención es 100% manual. SURGAS requiere automatizar pedidos frecuentes y dar respuesta inmediata a clientes 24/7.", { width: 4500 })] }),
          new TableRow({ children: [cell("Portal Web y E-commerce SURGAS", { fill: GRAY_LIGHT, width: 3026 }), cell("Sistema", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("SURGAS no cuenta con sitio web transaccional. Se requiere catálogo digital, opción de pedidos online y pasarela de pago.", { fill: GRAY_LIGHT, width: 4500 })] }),
          new TableRow({ children: [cell("Migración Infraestructura a Cloud Híbrida", { width: 3026 }), cell("Infraestructura", { width: 1500, align: AlignmentType.CENTER }), cell("Los sistemas críticos funcionan en servidores locales. Se requiere migrar a cloud para garantizar disponibilidad 99% y respaldo automático.", { width: 4500 })] }),
          new TableRow({ children: [cell("Sistema de Respaldo y Recuperación (BCP/DRP)", { fill: GRAY_LIGHT, width: 3026 }), cell("Infraestructura", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("Ausencia de políticas formales de respaldo. Riesgo alto de pérdida de información comercial y operativa ante fallas.", { fill: GRAY_LIGHT, width: 4500 })] }),
          new TableRow({ children: [cell("Fortalecimiento Ciberseguridad y Protección de Datos", { width: 3026 }), cell("Infraestructura", { width: 1500, align: AlignmentType.CENTER }), cell("Contraseñas débiles, accesos compartidos y sin políticas formales. Riesgo reputacional y legal ante fuga de datos de clientes.", { width: 4500 })] }),
          new TableRow({ children: [cell("Mesa de Ayuda (Help Desk) TI", { fill: GRAY_LIGHT, width: 3026 }), cell("Sistema", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("Las incidencias se reportan informalmente por WhatsApp o llamada. Se requiere centralizar el soporte con tickets.", { fill: GRAY_LIGHT, width: 4500 })] }),
          new TableRow({ children: [cell("Business Intelligence y Dashboard Gerencial", { width: 3026 }), cell("Sistema", { width: 1500, align: AlignmentType.CENTER }), cell("La gerencia carece de dashboards unificados para ventas, inventario y rentabilidad que apoyen decisiones estratégicas.", { width: 4500 })] }),
          new TableRow({ children: [cell("Sistema de Gestión de Red SURGAS (Distribuidores)", { fill: GRAY_LIGHT, width: 3026 }), cell("Sistema", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("Los +25 distribuidores de la Red SURGAS no cuentan con herramienta digital propia para gestionar sus pedidos y reportes.", { fill: GRAY_LIGHT, width: 4500 })] }),
        ]
      }),

      pageBreak(),

      // ===================== VII. HERRAMIENTAS DE ANÁLISIS =====================
      h1("VII. HERRAMIENTAS DE ANÁLISIS"),

      h2("1. FACTORES CRÍTICOS DE ÉXITO"),
      body("Para que la cartera de 10 proyectos priorizados llegue a buen puerto, se identifican los siguientes factores críticos:", { justify: true }),
      bullet("Compromiso y patrocinio de la Gerencia General: respaldo activo en inversión y decisiones, especialmente para app móvil, cloud y BI."),
      bullet("Alineación estratégica con objetivos de seguridad y expansión: cada iniciativa debe conectarse con metas comerciales (expansión sur, Red SURGAS)."),
      bullet("Disponibilidad presupuestal planificada: priorizar según índice B/C dado que varios proyectos compiten por el mismo presupuesto."),
      bullet("Capacidades técnicas y talento especializado: personal capacitado o proveedores certificados en desarrollo móvil, cloud e integración de sistemas."),
      bullet("Gestión del cambio organizacional: capacitar al personal de call center, distribuidores y operaciones para adoptar nuevas herramientas."),
      bullet("Calidad e integración de datos entre sistemas: la interoperabilidad entre inventario, pedidos y facturación depende de datos limpios."),
      bullet("Gobernanza de cartera (PMO): comité que revise periódicamente las matrices y reasigne recursos según el entorno operativo."),
      bullet("Gestión de riesgos de seguridad GLP y TI: especialmente crítico en proyectos que manejan datos de clientes y operaciones de distribución."),

      emptyLine(),
      h2("2. NIVEL DE MADUREZ EN GOBIERNO DE TI"),
      body("Aplicando el modelo de madurez COBIT (escala 0-5), se evalúa el estado actual de SURGAS:", { justify: true }),
      emptyLine(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1500, 3526, 4000],
        rows: [
          new TableRow({ children: [headerCell("Nivel", 1500), headerCell("Descripción", 3526), headerCell("Estado en SURGAS", 4000)] }),
          new TableRow({ children: [cell("0 – Inexistente", { width: 1500 }), cell("No hay procesos de gobierno TI reconocidos.", { width: 3526 }), cell("Superado parcialmente.", { width: 4000 })] }),
          new TableRow({ children: [cell("1 – Inicial/Ad hoc", { fill: YELLOW_LIGHT, width: 1500 }), cell("Iniciativas aisladas, sin estandarización.", { fill: YELLOW_LIGHT, width: 3526 }), cell("NIVEL ACTUAL: Proyectos reactivos sin priorización formal.", { fill: YELLOW_LIGHT, width: 4000, bold: true })] }),
          new TableRow({ children: [cell("2 – Repetible", { fill: YELLOW_LIGHT, width: 1500 }), cell("Patrones que se repiten, dependientes de personas clave.", { fill: YELLOW_LIGHT, width: 3526 }), cell("NIVEL ACTUAL: Algunos procesos se repiten pero sin documentación.", { fill: YELLOW_LIGHT, width: 4000, bold: true })] }),
          new TableRow({ children: [cell("3 – Definido", { fill: GREEN_LIGHT, width: 1500 }), cell("Procesos documentados con roles definidos.", { fill: GREEN_LIGHT, width: 3526 }), cell("META EN 12-18 MESES: Implementar PMO, Help Desk y políticas TI.", { fill: GREEN_LIGHT, width: 4000, bold: true })] }),
          new TableRow({ children: [cell("4 – Gestionado", { width: 1500 }), cell("Monitoreo de KPIs de TI.", { width: 3526 }), cell("Etapa futura con BI y dashboards gerenciales.", { width: 4000 })] }),
          new TableRow({ children: [cell("5 – Optimizado", { width: 1500 }), cell("Mejora continua con benchmarking.", { width: 3526 }), cell("Visión a 3-5 años.", { width: 4000 })] }),
        ]
      }),

      body("Diagnóstico: SURGAS se ubica actualmente entre Nivel 1 (Inicial) y Nivel 2 (Repetible), con procesos en gran parte reactivos. La cartera de proyectos (Help Desk, Backup, Ciberseguridad, BI) es el conjunto de iniciativas necesarias para escalar al Nivel 3 (Definido).", { justify: true }),

      emptyLine(),
      h2("3. MATRIZ DE McKINSEY"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [500, 2826, 1500, 1700, 1500, 1000],
        rows: [
          new TableRow({ children: [headerCell("N°", 500), headerCell("Proyecto", 2826), headerCell("Atractivo Mercado", 1500), headerCell("Fortaleza Competitiva", 1700), headerCell("Cuadrante", 1500), headerCell("Acción", 1000)] }),
          new TableRow({ children: [cell("1", { width: 500, align: AlignmentType.CENTER }), cell("App Móvil de Pedidos", { width: 2826 }), cell("5/5", { width: 1500, align: AlignmentType.CENTER }), cell("3/5", { width: 1700, align: AlignmentType.CENTER }), cell("Invertir/Crecer", { width: 1500, fill: GREEN_LIGHT }), cell("Invertir", { width: 1000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("2", { width: 500, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell("ERP Inventario GLP", { fill: GRAY_LIGHT, width: 2826 }), cell("4/5", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("4/5", { fill: GRAY_LIGHT, width: 1700, align: AlignmentType.CENTER }), cell("Invertir/Crecer", { fill: GREEN_LIGHT, width: 1500 }), cell("Invertir", { fill: GRAY_LIGHT, width: 1000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("3", { width: 500, align: AlignmentType.CENTER }), cell("Chatbot WhatsApp Pedidos", { width: 2826 }), cell("4/5", { width: 1500, align: AlignmentType.CENTER }), cell("3/5", { width: 1700, align: AlignmentType.CENTER }), cell("Selectividad", { fill: YELLOW_LIGHT, width: 1500 }), cell("Desarrollar", { width: 1000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("4", { width: 500, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell("Portal Web / E-commerce", { fill: GRAY_LIGHT, width: 2826 }), cell("5/5", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("3/5", { fill: GRAY_LIGHT, width: 1700, align: AlignmentType.CENTER }), cell("Invertir/Crecer", { fill: GREEN_LIGHT, width: 1500 }), cell("Invertir", { fill: GRAY_LIGHT, width: 1000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("5", { width: 500, align: AlignmentType.CENTER }), cell("Migración Cloud Híbrida", { width: 2826 }), cell("4/5", { width: 1500, align: AlignmentType.CENTER }), cell("2/5", { width: 1700, align: AlignmentType.CENTER }), cell("Selectividad", { fill: YELLOW_LIGHT, width: 1500 }), cell("Desarrollar", { width: 1000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("6", { width: 500, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell("BCP / Sistema Respaldo", { fill: GRAY_LIGHT, width: 2826 }), cell("3/5", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("4/5", { fill: GRAY_LIGHT, width: 1700, align: AlignmentType.CENTER }), cell("Invertir/Crecer", { fill: GREEN_LIGHT, width: 1500 }), cell("Invertir", { fill: GRAY_LIGHT, width: 1000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("7", { width: 500, align: AlignmentType.CENTER }), cell("Ciberseguridad y Datos", { width: 2826 }), cell("5/5", { width: 1500, align: AlignmentType.CENTER }), cell("3/5", { width: 1700, align: AlignmentType.CENTER }), cell("Invertir/Crecer", { fill: GREEN_LIGHT, width: 1500 }), cell("Invertir", { width: 1000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("8", { width: 500, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell("Help Desk / Mesa Ayuda", { fill: GRAY_LIGHT, width: 2826 }), cell("3/5", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("4/5", { fill: GRAY_LIGHT, width: 1700, align: AlignmentType.CENTER }), cell("Invertir/Crecer", { fill: GREEN_LIGHT, width: 1500 }), cell("Invertir", { fill: GRAY_LIGHT, width: 1000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("9", { width: 500, align: AlignmentType.CENTER }), cell("BI y Dashboard Gerencial", { width: 2826 }), cell("3/5", { width: 1500, align: AlignmentType.CENTER }), cell("2/5", { width: 1700, align: AlignmentType.CENTER }), cell("Cosechar/Replantear", { fill: RED_LIGHT, width: 1500 }), cell("Replantear", { width: 1000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("10", { width: 500, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell("Portal Distribuidores Red SURGAS", { fill: GRAY_LIGHT, width: 2826 }), cell("4/5", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("3/5", { fill: GRAY_LIGHT, width: 1700, align: AlignmentType.CENTER }), cell("Selectividad", { fill: YELLOW_LIGHT, width: 1500 }), cell("Desarrollar", { fill: GRAY_LIGHT, width: 1000, align: AlignmentType.CENTER })] }),
        ]
      }),

      emptyLine(),
      h2("4. MATRIZ DE McFARLAND"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [500, 2826, 1700, 2000, 2000],
        rows: [
          new TableRow({ children: [headerCell("N°", 500), headerCell("Proyecto", 2826), headerCell("Cuadrante", 1700), headerCell("Impacto Estratégico", 2000), headerCell("Riesgo Implementación", 2000)] }),
          new TableRow({ children: [cell("1", { width: 500, align: AlignmentType.CENTER }), cell("App Móvil de Pedidos", { width: 2826 }), cell("Estratégico", { fill: GREEN_LIGHT, width: 1700 }), cell("5/5", { width: 2000, align: AlignmentType.CENTER }), cell("3/5", { width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("2", { width: 500, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell("ERP Inventario GLP", { fill: GRAY_LIGHT, width: 2826 }), cell("Estratégico", { fill: GREEN_LIGHT, width: 1700 }), cell("4/5", { fill: GRAY_LIGHT, width: 2000, align: AlignmentType.CENTER }), cell("3/5", { fill: GRAY_LIGHT, width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("3", { width: 500, align: AlignmentType.CENTER }), cell("Chatbot WhatsApp Pedidos", { width: 2826 }), cell("Tornaround", { fill: YELLOW_LIGHT, width: 1700 }), cell("4/5", { width: 2000, align: AlignmentType.CENTER }), cell("4/5", { width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("4", { width: 500, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell("Portal Web / E-commerce", { fill: GRAY_LIGHT, width: 2826 }), cell("Estratégico", { fill: GREEN_LIGHT, width: 1700 }), cell("5/5", { fill: GRAY_LIGHT, width: 2000, align: AlignmentType.CENTER }), cell("2/5", { fill: GRAY_LIGHT, width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("5", { width: 500, align: AlignmentType.CENTER }), cell("Migración Cloud Híbrida", { width: 2826 }), cell("Tornaround", { fill: YELLOW_LIGHT, width: 1700 }), cell("4/5", { width: 2000, align: AlignmentType.CENTER }), cell("4/5", { width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("6", { width: 500, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell("BCP / Sistema Respaldo", { fill: GRAY_LIGHT, width: 2826 }), cell("Soporte", { fill: LIGHT_BLUE, width: 1700 }), cell("3/5", { fill: GRAY_LIGHT, width: 2000, align: AlignmentType.CENTER }), cell("2/5", { fill: GRAY_LIGHT, width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("7", { width: 500, align: AlignmentType.CENTER }), cell("Ciberseguridad y Datos", { width: 2826 }), cell("Estratégico", { fill: GREEN_LIGHT, width: 1700 }), cell("5/5", { width: 2000, align: AlignmentType.CENTER }), cell("2/5", { width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("8", { width: 500, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell("Help Desk / Mesa Ayuda", { fill: GRAY_LIGHT, width: 2826 }), cell("Soporte", { fill: LIGHT_BLUE, width: 1700 }), cell("3/5", { fill: GRAY_LIGHT, width: 2000, align: AlignmentType.CENTER }), cell("2/5", { fill: GRAY_LIGHT, width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("9", { width: 500, align: AlignmentType.CENTER }), cell("BI y Dashboard Gerencial", { width: 2826 }), cell("Tornaround", { fill: YELLOW_LIGHT, width: 1700 }), cell("4/5", { width: 2000, align: AlignmentType.CENTER }), cell("4/5", { width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("10", { width: 500, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell("Portal Red SURGAS Distribuidores", { fill: GRAY_LIGHT, width: 2826 }), cell("Estratégico", { fill: GREEN_LIGHT, width: 1700 }), cell("4/5", { fill: GRAY_LIGHT, width: 2000, align: AlignmentType.CENTER }), cell("3/5", { fill: GRAY_LIGHT, width: 2000, align: AlignmentType.CENTER })] }),
        ]
      }),

      emptyLine(),
      h2("5. MATRIZ GARTNER TIME"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [500, 2826, 2200, 1500, 2000],
        rows: [
          new TableRow({ children: [headerCell("N°", 500), headerCell("Proyecto", 2826), headerCell("Posición", 2200), headerCell("Valor Potencial", 1500), headerCell("Velocidad de Retorno", 2000)] }),
          new TableRow({ children: [cell("1", { width: 500, align: AlignmentType.CENTER }), cell("App Móvil de Pedidos", { width: 2826 }), cell("Alto Retorno (6–12 meses)", { fill: YELLOW_LIGHT, width: 2200 }), cell("5/5", { width: 1500, align: AlignmentType.CENTER }), cell("3/5", { width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("2", { width: 500, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell("ERP Inventario GLP", { fill: GRAY_LIGHT, width: 2826 }), cell("Victoria Rápida (< 6 meses)", { fill: GREEN_LIGHT, width: 2200 }), cell("4/5", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("5/5", { fill: GRAY_LIGHT, width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("3", { width: 500, align: AlignmentType.CENTER }), cell("Chatbot WhatsApp Pedidos", { width: 2826 }), cell("Alto Retorno (6–12 meses)", { fill: YELLOW_LIGHT, width: 2200 }), cell("4/5", { width: 1500, align: AlignmentType.CENTER }), cell("3/5", { width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("4", { width: 500, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell("Portal Web / E-commerce", { fill: GRAY_LIGHT, width: 2826 }), cell("Alto Retorno (6–12 meses)", { fill: YELLOW_LIGHT, width: 2200 }), cell("5/5", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("3/5", { fill: GRAY_LIGHT, width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("5", { width: 500, align: AlignmentType.CENTER }), cell("Migración Cloud Híbrida", { width: 2826 }), cell("Transformación (> 12 meses)", { fill: ORANGE_LIGHT, width: 2200 }), cell("4/5", { width: 1500, align: AlignmentType.CENTER }), cell("1/5", { width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("6", { width: 500, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell("BCP / Sistema Respaldo", { fill: GRAY_LIGHT, width: 2826 }), cell("Victoria Rápida (< 6 meses)", { fill: GREEN_LIGHT, width: 2200 }), cell("3/5", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("5/5", { fill: GRAY_LIGHT, width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("7", { width: 500, align: AlignmentType.CENTER }), cell("Ciberseguridad y Datos", { width: 2826 }), cell("Victoria Rápida (< 6 meses)", { fill: GREEN_LIGHT, width: 2200 }), cell("4/5", { width: 1500, align: AlignmentType.CENTER }), cell("5/5", { width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("8", { width: 500, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell("Help Desk / Mesa Ayuda", { fill: GRAY_LIGHT, width: 2826 }), cell("Victoria Rápida (< 6 meses)", { fill: GREEN_LIGHT, width: 2200 }), cell("3/5", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("5/5", { fill: GRAY_LIGHT, width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("9", { width: 500, align: AlignmentType.CENTER }), cell("BI y Dashboard Gerencial", { width: 2826 }), cell("Transformación (> 12 meses)", { fill: ORANGE_LIGHT, width: 2200 }), cell("4/5", { width: 1500, align: AlignmentType.CENTER }), cell("1/5", { width: 2000, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("10", { width: 500, align: AlignmentType.CENTER, fill: GRAY_LIGHT }), cell("Portal Red SURGAS", { fill: GRAY_LIGHT, width: 2826 }), cell("Alto Retorno (6–12 meses)", { fill: YELLOW_LIGHT, width: 2200 }), cell("4/5", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("3/5", { fill: GRAY_LIGHT, width: 2000, align: AlignmentType.CENTER })] }),
        ]
      }),

      emptyLine(),
      h2("6. MATRIZ DE UTILIDAD – FACTIBILIDAD"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [500, 2626, 1200, 1000, 1200, 1000, 1500],
        rows: [
          new TableRow({ children: [headerCell("N°", 500), headerCell("Proyecto", 2626), headerCell("Utilidad (1-5)", 1200), headerCell("Punt.", 1000), headerCell("Factibilidad (1-5)", 1200), headerCell("Punt.", 1000), headerCell("U×F / Cuadrante", 1500)] }),
          new TableRow({ children: [cell("1", { width: 500, align: AlignmentType.CENTER }), cell("App Móvil de Pedidos", { width: 2626 }), cell("Muy Alta", { width: 1200, align: AlignmentType.CENTER }), cell("5", { width: 1000, align: AlignmentType.CENTER }), cell("Media", { width: 1200, align: AlignmentType.CENTER }), cell("3", { width: 1000, align: AlignmentType.CENTER }), cell("15 – EJECUTAR", { fill: YELLOW_LIGHT, width: 1500 })] }),
          new TableRow({ children: [cell("2", { fill: GRAY_LIGHT, width: 500, align: AlignmentType.CENTER }), cell("ERP Inventario GLP", { fill: GRAY_LIGHT, width: 2626 }), cell("Muy Alta", { fill: GRAY_LIGHT, width: 1200, align: AlignmentType.CENTER }), cell("5", { fill: GRAY_LIGHT, width: 1000, align: AlignmentType.CENTER }), cell("Alta", { fill: GRAY_LIGHT, width: 1200, align: AlignmentType.CENTER }), cell("4", { fill: GRAY_LIGHT, width: 1000, align: AlignmentType.CENTER }), cell("20 – PRIORIZAR", { fill: GREEN_LIGHT, width: 1500 })] }),
          new TableRow({ children: [cell("3", { width: 500, align: AlignmentType.CENTER }), cell("Chatbot WhatsApp", { width: 2626 }), cell("Alta", { width: 1200, align: AlignmentType.CENTER }), cell("4", { width: 1000, align: AlignmentType.CENTER }), cell("Alta", { width: 1200, align: AlignmentType.CENTER }), cell("4", { width: 1000, align: AlignmentType.CENTER }), cell("16 – PRIORIZAR", { fill: GREEN_LIGHT, width: 1500 })] }),
          new TableRow({ children: [cell("4", { fill: GRAY_LIGHT, width: 500, align: AlignmentType.CENTER }), cell("Portal Web / E-commerce", { fill: GRAY_LIGHT, width: 2626 }), cell("Muy Alta", { fill: GRAY_LIGHT, width: 1200, align: AlignmentType.CENTER }), cell("5", { fill: GRAY_LIGHT, width: 1000, align: AlignmentType.CENTER }), cell("Media", { fill: GRAY_LIGHT, width: 1200, align: AlignmentType.CENTER }), cell("3", { fill: GRAY_LIGHT, width: 1000, align: AlignmentType.CENTER }), cell("15 – EJECUTAR", { fill: YELLOW_LIGHT, width: 1500 })] }),
          new TableRow({ children: [cell("5", { width: 500, align: AlignmentType.CENTER }), cell("Cloud Híbrida", { width: 2626 }), cell("Alta", { width: 1200, align: AlignmentType.CENTER }), cell("4", { width: 1000, align: AlignmentType.CENTER }), cell("Baja", { width: 1200, align: AlignmentType.CENTER }), cell("2", { width: 1000, align: AlignmentType.CENTER }), cell("8 – EVALUAR", { fill: ORANGE_LIGHT, width: 1500 })] }),
          new TableRow({ children: [cell("6", { fill: GRAY_LIGHT, width: 500, align: AlignmentType.CENTER }), cell("BCP / Respaldo", { fill: GRAY_LIGHT, width: 2626 }), cell("Media", { fill: GRAY_LIGHT, width: 1200, align: AlignmentType.CENTER }), cell("3", { fill: GRAY_LIGHT, width: 1000, align: AlignmentType.CENTER }), cell("Muy Alta", { fill: GRAY_LIGHT, width: 1200, align: AlignmentType.CENTER }), cell("5", { fill: GRAY_LIGHT, width: 1000, align: AlignmentType.CENTER }), cell("15 – EJECUTAR", { fill: YELLOW_LIGHT, width: 1500 })] }),
          new TableRow({ children: [cell("7", { width: 500, align: AlignmentType.CENTER }), cell("Ciberseguridad", { width: 2626 }), cell("Muy Alta", { width: 1200, align: AlignmentType.CENTER }), cell("5", { width: 1000, align: AlignmentType.CENTER }), cell("Alta", { width: 1200, align: AlignmentType.CENTER }), cell("4", { width: 1000, align: AlignmentType.CENTER }), cell("20 – PRIORIZAR", { fill: GREEN_LIGHT, width: 1500 })] }),
          new TableRow({ children: [cell("8", { fill: GRAY_LIGHT, width: 500, align: AlignmentType.CENTER }), cell("Help Desk", { fill: GRAY_LIGHT, width: 2626 }), cell("Media", { fill: GRAY_LIGHT, width: 1200, align: AlignmentType.CENTER }), cell("3", { fill: GRAY_LIGHT, width: 1000, align: AlignmentType.CENTER }), cell("Alta", { fill: GRAY_LIGHT, width: 1200, align: AlignmentType.CENTER }), cell("4", { fill: GRAY_LIGHT, width: 1000, align: AlignmentType.CENTER }), cell("12 – EJECUTAR", { fill: YELLOW_LIGHT, width: 1500 })] }),
          new TableRow({ children: [cell("9", { width: 500, align: AlignmentType.CENTER }), cell("BI y Dashboard", { width: 2626 }), cell("Alta", { width: 1200, align: AlignmentType.CENTER }), cell("4", { width: 1000, align: AlignmentType.CENTER }), cell("Baja", { width: 1200, align: AlignmentType.CENTER }), cell("2", { width: 1000, align: AlignmentType.CENTER }), cell("8 – EVALUAR", { fill: ORANGE_LIGHT, width: 1500 })] }),
          new TableRow({ children: [cell("10", { fill: GRAY_LIGHT, width: 500, align: AlignmentType.CENTER }), cell("Portal Red SURGAS", { fill: GRAY_LIGHT, width: 2626 }), cell("Alta", { fill: GRAY_LIGHT, width: 1200, align: AlignmentType.CENTER }), cell("4", { fill: GRAY_LIGHT, width: 1000, align: AlignmentType.CENTER }), cell("Media", { fill: GRAY_LIGHT, width: 1200, align: AlignmentType.CENTER }), cell("3", { fill: GRAY_LIGHT, width: 1000, align: AlignmentType.CENTER }), cell("12 – EJECUTAR", { fill: YELLOW_LIGHT, width: 1500 })] }),
        ]
      }),

      emptyLine(),
      h2("7. MATRIZ DE BENEFICIO – COSTO"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [500, 2226, 1300, 1500, 1000, 1200, 1300],
        rows: [
          new TableRow({ children: [headerCell("N°", 500), headerCell("Proyecto", 2226), headerCell("Beneficio Anual (S/)", 1300), headerCell("Costo Estimado (S/)", 1500), headerCell("Índice B/C", 1000), headerCell("Viabilidad", 1200), headerCell("Urgencia", 1300)] }),
          new TableRow({ children: [cell("1", { width: 500, align: AlignmentType.CENTER }), cell("App Móvil Pedidos", { width: 2226 }), cell("S/ 360,000", { width: 1300, align: AlignmentType.CENTER }), cell("S/ 120,000", { width: 1500, align: AlignmentType.CENTER }), cell("3.0", { fill: GREEN_LIGHT, width: 1000, align: AlignmentType.CENTER }), cell("Muy Viable", { width: 1200, align: AlignmentType.CENTER }), cell("5/5", { width: 1300, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("2", { fill: GRAY_LIGHT, width: 500, align: AlignmentType.CENTER }), cell("ERP Inventario GLP", { fill: GRAY_LIGHT, width: 2226 }), cell("S/ 280,000", { fill: GRAY_LIGHT, width: 1300, align: AlignmentType.CENTER }), cell("S/ 70,000", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("4.0", { fill: GREEN_LIGHT, width: 1000, align: AlignmentType.CENTER }), cell("Muy Viable", { fill: GRAY_LIGHT, width: 1200, align: AlignmentType.CENTER }), cell("5/5", { fill: GRAY_LIGHT, width: 1300, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("3", { width: 500, align: AlignmentType.CENTER }), cell("Chatbot WhatsApp", { width: 2226 }), cell("S/ 180,000", { width: 1300, align: AlignmentType.CENTER }), cell("S/ 60,000", { width: 1500, align: AlignmentType.CENTER }), cell("3.0", { fill: GREEN_LIGHT, width: 1000, align: AlignmentType.CENTER }), cell("Muy Viable", { width: 1200, align: AlignmentType.CENTER }), cell("4/5", { width: 1300, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("4", { fill: GRAY_LIGHT, width: 500, align: AlignmentType.CENTER }), cell("Portal Web / E-commerce", { fill: GRAY_LIGHT, width: 2226 }), cell("S/ 420,000", { fill: GRAY_LIGHT, width: 1300, align: AlignmentType.CENTER }), cell("S/ 140,000", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("3.0", { fill: GREEN_LIGHT, width: 1000, align: AlignmentType.CENTER }), cell("Muy Viable", { fill: GRAY_LIGHT, width: 1200, align: AlignmentType.CENTER }), cell("4/5", { fill: GRAY_LIGHT, width: 1300, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("5", { width: 500, align: AlignmentType.CENTER }), cell("Cloud Híbrida", { width: 2226 }), cell("S/ 390,000", { width: 1300, align: AlignmentType.CENTER }), cell("S/ 300,000", { width: 1500, align: AlignmentType.CENTER }), cell("1.3", { fill: YELLOW_LIGHT, width: 1000, align: AlignmentType.CENTER }), cell("Viable", { width: 1200, align: AlignmentType.CENTER }), cell("3/5", { width: 1300, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("6", { fill: GRAY_LIGHT, width: 500, align: AlignmentType.CENTER }), cell("BCP / Respaldo", { fill: GRAY_LIGHT, width: 2226 }), cell("S/ 160,000", { fill: GRAY_LIGHT, width: 1300, align: AlignmentType.CENTER }), cell("S/ 50,000", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("3.2", { fill: GREEN_LIGHT, width: 1000, align: AlignmentType.CENTER }), cell("Muy Viable", { fill: GRAY_LIGHT, width: 1200, align: AlignmentType.CENTER }), cell("4/5", { fill: GRAY_LIGHT, width: 1300, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("7", { width: 500, align: AlignmentType.CENTER }), cell("Ciberseguridad", { width: 2226 }), cell("S/ 240,000", { width: 1300, align: AlignmentType.CENTER }), cell("S/ 75,000", { width: 1500, align: AlignmentType.CENTER }), cell("3.2", { fill: GREEN_LIGHT, width: 1000, align: AlignmentType.CENTER }), cell("Muy Viable", { width: 1200, align: AlignmentType.CENTER }), cell("5/5", { width: 1300, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("8", { fill: GRAY_LIGHT, width: 500, align: AlignmentType.CENTER }), cell("Help Desk", { fill: GRAY_LIGHT, width: 2226 }), cell("S/ 140,000", { fill: GRAY_LIGHT, width: 1300, align: AlignmentType.CENTER }), cell("S/ 55,000", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("2.5", { fill: YELLOW_LIGHT, width: 1000, align: AlignmentType.CENTER }), cell("Muy Viable", { fill: GRAY_LIGHT, width: 1200, align: AlignmentType.CENTER }), cell("3/5", { fill: GRAY_LIGHT, width: 1300, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("9", { width: 500, align: AlignmentType.CENTER }), cell("BI y Dashboard", { width: 2226 }), cell("S/ 350,000", { width: 1300, align: AlignmentType.CENTER }), cell("S/ 260,000", { width: 1500, align: AlignmentType.CENTER }), cell("1.35", { fill: YELLOW_LIGHT, width: 1000, align: AlignmentType.CENTER }), cell("Viable", { width: 1200, align: AlignmentType.CENTER }), cell("2/5", { width: 1300, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("10", { fill: GRAY_LIGHT, width: 500, align: AlignmentType.CENTER }), cell("Portal Red SURGAS", { fill: GRAY_LIGHT, width: 2226 }), cell("S/ 200,000", { fill: GRAY_LIGHT, width: 1300, align: AlignmentType.CENTER }), cell("S/ 90,000", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("2.2", { fill: YELLOW_LIGHT, width: 1000, align: AlignmentType.CENTER }), cell("Muy Viable", { fill: GRAY_LIGHT, width: 1200, align: AlignmentType.CENTER }), cell("4/5", { fill: GRAY_LIGHT, width: 1300, align: AlignmentType.CENTER })] }),
        ]
      }),

      pageBreak(),

      // ===================== VIII. ESTRATEGIA DE CAMBIO =====================
      h1("VIII. ESTRATEGIA DE CAMBIO"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2000, 3513, 3513],
        rows: [
          new TableRow({ children: [headerCell("Cambio / Proyecto", 2000), headerCell("Retos", 3513), headerCell("Acciones", 3513)] }),
          new TableRow({ children: [cell("P1: App Móvil de Pedidos", { fill: LIGHT_BLUE, bold: true, width: 2000 }), cell("Resistencia de distribuidores Red SURGAS acostumbrados al pedido telefónico. Falta de experiencia interna en desarrollo móvil. Riesgo de baja adopción en zonas sin smartphone.", { width: 3513 }), cell("Implementar en fases: primero call center, luego distribuidores, luego clientes finales. Capacitar al personal. Soporte técnico intensivo primeros 60 días.", { width: 3513 })] }),
          new TableRow({ children: [cell("P2: ERP Inventario GLP", { fill: LIGHT_BLUE, bold: true, width: 2000, fill: GRAY_LIGHT }), cell("Dependencia de conteo manual de balones. Resistencia del personal operativo. Riesgo de discrepancias durante migración.", { fill: GRAY_LIGHT, width: 3513 }), cell("Prueba piloto con balones de 10 kg en zonal Tacna. Ambiente paralelo para validar sin afectar operaciones. Capacitar a logística y distribución.", { fill: GRAY_LIGHT, width: 3513 })] }),
          new TableRow({ children: [cell("P3: Chatbot WhatsApp", { fill: LIGHT_BLUE, bold: true, width: 2000 }), cell("Temor del personal a ser reemplazado. Falta de base de conocimiento para respuestas automáticas. Integración compleja con WhatsApp API.", { width: 3513 }), cell("Iniciar con FAQs y pedidos frecuentes. Modelo híbrido con derivación a asesor para casos complejos. Medir periódicamente satisfacción digital.", { width: 3513 })] }),
          new TableRow({ children: [cell("P4: Portal Web / E-commerce", { fill: LIGHT_BLUE, bold: true, width: 2000, fill: GRAY_LIGHT }), cell("Sin experiencia interna en e-commerce GLP. Riesgo de inconsistencias con stock físico. Regulación sobre venta de GLP en línea.", { fill: GRAY_LIGHT, width: 3513 }), cell("Catálogo digital primero, luego pedidos con pago en línea. Comité TI + Ventas + Legal para validar cumplimiento normativo.", { fill: GRAY_LIGHT, width: 3513 })] }),
          new TableRow({ children: [cell("P5–P7: Cloud, BCP y Ciberseguridad", { fill: LIGHT_BLUE, bold: true, width: 2000 }), cell("Baja experiencia en cloud y seguridad. Temor por costos. Riesgo de interrupción durante migración.", { width: 3513 }), cell("Migración gradual iniciando por backups. Contratar soporte especializado. Capacitar todo el personal en ciberseguridad y phishing.", { width: 3513 })] }),
        ]
      }),

      pageBreak(),

      // ===================== IX. CONCLUSIONES =====================
      h1("IX. CONCLUSIONES"),

      body("El análisis realizado evidencia que Grupo SURGAS E.I.R.L. presenta oportunidades significativas de mejora en sus capacidades tecnológicas, especialmente en la digitalización del canal de pedidos, el control de inventario en tiempo real, la automatización de la atención al cliente y el fortalecimiento de la infraestructura TI. La dependencia actual de procesos parcialmente manuales limita la eficiencia operativa y la capacidad de crecimiento sostenible hacia nuevas regiones del sur del Perú.", { justify: true }),
      emptyLine(),
      body("La aplicación del Balanced Scorecard de Tecnologías de Información permitió alinear los objetivos estratégicos del negocio con iniciativas tecnológicas específicas, estableciendo indicadores medibles relacionados con la continuidad operativa, satisfacción del cliente, eficiencia en distribución GLP y crecimiento de la Red SURGAS.", { justify: true }),
      emptyLine(),
      body("A partir del análisis de madurez tecnológica basado en COBIT, se identificó que SURGAS presenta un nivel inicial-intermedio de madurez (entre Nivel 1 y Nivel 2), caracterizado por procesos parcialmente definidos, limitada integración entre sistemas y una gestión tecnológica predominantemente reactiva, lo cual justifica la urgente necesidad de fortalecer la gobernanza de TI.", { justify: true }),
      emptyLine(),
      body("Las matrices de priorización aplicadas (McKinsey, McFarland y Gartner TIME) permitieron clasificar los 10 proyectos según su atractivo estratégico, fortaleza competitiva, impacto organizacional, riesgo de implementación y velocidad de retorno, facilitando una toma de decisiones objetiva sobre la inversión tecnológica.", { justify: true }),
      emptyLine(),
      body("Se concluye que proyectos como el ERP de inventario GLP, el sistema de ciberseguridad, el chatbot de pedidos y la mesa de ayuda representan iniciativas prioritarias por su impacto inmediato en la continuidad operativa y el control interno. La App Móvil y el Portal Web son los proyectos de mayor transformación digital a mediano plazo para consolidar la competitividad de SURGAS en el mercado energético del sur del Perú.", { justify: true }),

      emptyLine(),

      // ===================== X. RECOMENDACIONES =====================
      h1("X. RECOMENDACIONES"),

      bullet("Implementar los proyectos tecnológicos de manera progresiva y priorizada, iniciando por aquellos de victoria rápida e impacto operativo inmediato: ERP de inventario GLP, sistema de respaldo BCP, ciberseguridad y mesa de ayuda."),
      bullet("Establecer un plan formal de gestión del cambio organizacional que incluya capacitación continua al personal de call center, distribuidores Red SURGAS y operaciones, reduciendo la resistencia al cambio."),
      bullet("Fortalecer progresivamente el área de Tecnologías de Información mediante la contratación de al menos un Administrador de Sistemas dedicado y la formalización de roles y procesos TI."),
      bullet("Priorizar soluciones tecnológicas escalables y de bajo costo operativo (SaaS o servicios cloud administrados), considerando el tamaño actual de la organización y su proceso de expansión regional."),
      bullet("Implementar mecanismos de monitoreo y seguimiento de KPIs definidos en el Balanced Scorecard de TI, evaluando periódicamente el desempeño de los proyectos tecnológicos."),
      bullet("Establecer alianzas estratégicas con municipalidades, OSINERGMIN y bomberos para reforzar la imagen de seguridad de SURGAS, apoyadas en herramientas digitales de comunicación institucional."),
      bullet("Promover una cultura organizacional orientada al uso de datos para la toma de decisiones, preparando el terreno para la futura implementación de Business Intelligence y dashboards gerenciales."),
      bullet("Asegurar el cumplimiento normativo de SUNAT, INDECOPI y regulaciones del sector energético en todos los proyectos digitales, especialmente en el portal web y el e-commerce de GLP."),

      pageBreak(),

      // ===================== ANEXOS =====================
      h1("ANEXOS"),

      h2("ANEXO N° 01: OBJETIVOS, INDICADORES E INICIATIVAS DE TI"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1400, 1700, 1600, 1600, 1200, 1526],
        rows: [
          new TableRow({ children: [headerCell("Perspectiva", 1400), headerCell("Objetivo Empresa", 1700), headerCell("Objetivo TI", 1600), headerCell("Indicador TI (KPI)", 1600), headerCell("Meta TI", 1200), headerCell("Iniciativa TI", 1526)] }),
          new TableRow({ children: [cell("Financiera", { fill: LIGHT_BLUE, width: 1400 }), cell("O1: Incrementar ingresos GLP", { width: 1700 }), cell("TI-O1: Maximizar ROI canal digital pedidos", { width: 1600 }), cell("% ventas procesadas digitalmente", { width: 1600 }), cell(">= 25%", { width: 1200, align: AlignmentType.CENTER }), cell("Integrar pasarelas de pago y app móvil.", { width: 1526 })] }),
          new TableRow({ children: [cell("Financiera", { fill: LIGHT_BLUE, width: 1400, fill: GRAY_LIGHT }), cell("O2: Reducir costos distribución", { fill: GRAY_LIGHT, width: 1700 }), cell("TI-O2: Reducir costos TI con cloud", { fill: GRAY_LIGHT, width: 1600 }), cell("% reducción costos mantenimiento", { fill: GRAY_LIGHT, width: 1600 }), cell(">= 10%", { fill: GRAY_LIGHT, width: 1200, align: AlignmentType.CENTER }), cell("Migrar a SaaS/IaaS cloud.", { fill: GRAY_LIGHT, width: 1526 })] }),
          new TableRow({ children: [cell("Cliente", { fill: LIGHT_BLUE, width: 1400 }), cell("O3: Fidelizar clientes GLP", { width: 1700 }), cell("TI-O3: Optimizar velocidad canal digital", { width: 1600 }), cell("Tiempo respuesta pedido automatizado", { width: 1600 }), cell("<= 2 min", { width: 1200, align: AlignmentType.CENTER }), cell("Chatbot + App móvil SURGAS.", { width: 1526 })] }),
          new TableRow({ children: [cell("Cliente", { fill: LIGHT_BLUE, width: 1400, fill: GRAY_LIGHT }), cell("O4: Atención ágil call center", { fill: GRAY_LIGHT, width: 1700 }), cell("TI-O4: Plataformas estables para call center", { fill: GRAY_LIGHT, width: 1600 }), cell("Satisfacción usuario interno con TI", { fill: GRAY_LIGHT, width: 1600 }), cell(">= 90%", { fill: GRAY_LIGHT, width: 1200, align: AlignmentType.CENTER }), cell("Help Desk centralizado.", { fill: GRAY_LIGHT, width: 1526 })] }),
          new TableRow({ children: [cell("Procesos", { fill: LIGHT_BLUE, width: 1400 }), cell("O5: Control stock balones", { width: 1700 }), cell("TI-O5: ERP inventario GLP tiempo real", { width: 1600 }), cell("% discrepancia stock sistema vs físico", { width: 1600 }), cell("0% críticos", { width: 1200, align: AlignmentType.CENTER }), cell("API integración almacén-pedidos.", { width: 1526 })] }),
          new TableRow({ children: [cell("Procesos", { fill: LIGHT_BLUE, width: 1400, fill: GRAY_LIGHT }), cell("O6: Automatizar pedidos digitales", { fill: GRAY_LIGHT, width: 1700 }), cell("TI-O6: Chatbot pedidos automáticos", { fill: GRAY_LIGHT, width: 1600 }), cell("Tiempo respuesta inicial automatizada", { fill: GRAY_LIGHT, width: 1600 }), cell("<= 2 min", { fill: GRAY_LIGHT, width: 1200, align: AlignmentType.CENTER }), cell("WhatsApp Business API + chatbot IA.", { fill: GRAY_LIGHT, width: 1526 })] }),
          new TableRow({ children: [cell("Apr. y Crec.", { fill: LIGHT_BLUE, width: 1400 }), cell("O7: Capacitar personal seguridad", { width: 1700 }), cell("TI-O7: Capacitar en sistemas y ciberseg.", { width: 1600 }), cell("% colaboradores certificados", { width: 1600 }), cell("100%", { width: 1200, align: AlignmentType.CENTER }), cell("Talleres ciberseguridad y ERP.", { width: 1526 })] }),
          new TableRow({ children: [cell("Apr. y Crec.", { fill: LIGHT_BLUE, width: 1400, fill: GRAY_LIGHT }), cell("O8: Fortalecer infraestructura TI", { fill: GRAY_LIGHT, width: 1700 }), cell("TI-O8: Garantizar Uptime sistemas críticos", { fill: GRAY_LIGHT, width: 1600 }), cell("Tasa Uptime pedidos + facturación", { fill: GRAY_LIGHT, width: 1600 }), cell(">= 99%", { fill: GRAY_LIGHT, width: 1200, align: AlignmentType.CENTER }), cell("Cloud + backups encriptados diarios.", { fill: GRAY_LIGHT, width: 1526 })] }),
          new TableRow({ children: [cell("Apr. y Crec.", { fill: LIGHT_BLUE, width: 1400 }), cell("O9: Mejorar clima laboral", { width: 1700 }), cell("TI-O9: Intranet y comunicación digital", { width: 1600 }), cell("% adopción plataforma colaborativa", { width: 1600 }), cell(">= 85%", { width: 1200, align: AlignmentType.CENTER }), cell("Intranet y portal distribuidores Red SURGAS.", { width: 1526 })] }),
        ]
      }),

      emptyLine(),
      h2("ANEXO N° 02: FICHAS DE INDICADORES"),

      h3("FICHA 1 – INDICADOR TI: Porcentaje de discrepancia de stock de balones GLP"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2500, 6526],
        rows: [
          new TableRow({ children: [cell("Indicador de TI", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("% discrepancia de stock sistema vs inventario físico de balones GLP | Código: TI-SG-005", { width: 6526 })] }),
          new TableRow({ children: [cell("Objetivo Empresarial", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("O5: Optimizar el control de stock de balones GLP para erradicar quiebres de inventario.", { width: 6526 })] }),
          new TableRow({ children: [cell("Objetivo de TI", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("TI-O5: Automatizar la sincronización de inventarios de balones GLP en tiempo real.", { width: 6526 })] }),
          new TableRow({ children: [cell("Finalidad", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Medir el margen de error entre balones físicos en planta/zonal y lo que el sistema ERP registra, evitando sobreventas y quiebres.", { width: 6526 })] }),
          new TableRow({ children: [cell("Fórmula", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("(N° balones con diferencia / Total balones auditados) × 100", { width: 6526 })] }),
          new TableRow({ children: [cell("Unidad de Medida", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Porcentaje (%)", { width: 6526 })] }),
          new TableRow({ children: [cell("Frecuencia de Medición", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Mensual", { width: 6526 })] }),
          new TableRow({ children: [cell("Meta", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("0% de error al cierre del período (Mes 12)", { width: 6526 })] }),
          new TableRow({ children: [cell("Fuente de Datos", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("ERP de inventario / Conteo físico de almacén Calana / Reportes de auditoría logística.", { width: 6526 })] }),
          new TableRow({ children: [cell("Responsable", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Jefe de TI en coordinación con Jefe de Operaciones y Distribución.", { width: 6526 })] }),
        ]
      }),

      emptyLine(),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1500, 750, 750, 750, 750, 750, 750, 750, 750, 750, 750, 750, 750],
        rows: [
          new TableRow({ children: [headerCell("Resultado", 1500), headerCell("ENE", 750), headerCell("FEB", 750), headerCell("MAR", 750), headerCell("ABR", 750), headerCell("MAY", 750), headerCell("JUN", 750), headerCell("JUL", 750), headerCell("AGO", 750), headerCell("SET", 750), headerCell("OCT", 750), headerCell("NOV", 750), headerCell("DIC", 750)] }),
          new TableRow({ children: [cell("Verde: 0%", { fill: GREEN_LIGHT, width: 1500 }), cell("", { width: 750 }), cell("", { width: 750 }), cell("", { width: 750 }), cell("", { width: 750 }), cell("", { width: 750 }), cell("", { width: 750 }), cell("3%", { fill: GREEN_LIGHT, width: 750, align: AlignmentType.CENTER }), cell("2%", { fill: GREEN_LIGHT, width: 750, align: AlignmentType.CENTER }), cell("1%", { fill: GREEN_LIGHT, width: 750, align: AlignmentType.CENTER }), cell("0.5%", { fill: GREEN_LIGHT, width: 750, align: AlignmentType.CENTER }), cell("0.2%", { fill: GREEN_LIGHT, width: 750, align: AlignmentType.CENTER }), cell("0%", { fill: GREEN_LIGHT, width: 750, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Amarillo: 1%–4%", { fill: YELLOW_LIGHT, width: 1500 }), ...Array(12).fill(null).map(() => cell("", { width: 750 }))] }),
          new TableRow({ children: [cell("Rojo: > 4%", { fill: RED_LIGHT, width: 1500 }), ...Array(12).fill(null).map(() => cell("", { width: 750 }))] }),
        ]
      }),

      emptyLine(),
      h3("FICHA 2 – INDICADOR TI: Tasa de Disponibilidad (Uptime) del Sistema de Pedidos"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2500, 6526],
        rows: [
          new TableRow({ children: [cell("Indicador de TI", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Tasa de disponibilidad (Uptime) del Sistema de Pedidos y Facturación | Código: TI-SG-008", { width: 6526 })] }),
          new TableRow({ children: [cell("Objetivo Empresarial", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("O8: Fortalecer la infraestructura tecnológica para garantizar continuidad operativa.", { width: 6526 })] }),
          new TableRow({ children: [cell("Finalidad", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Monitorear el tiempo operativo del sistema de gestión de pedidos y facturación electrónica para evitar caídas del sistema, especialmente en días pico de demanda.", { width: 6526 })] }),
          new TableRow({ children: [cell("Fórmula", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("(Tiempo operativo programado – Tiempo de inactividad por fallas) / Tiempo operativo programado × 100", { width: 6526 })] }),
          new TableRow({ children: [cell("Meta", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell(">= 99% de disponibilidad mensual", { width: 6526 })] }),
          new TableRow({ children: [cell("Fuente de Datos", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Logs del servidor / Registros del sistema de pedidos / Monitor de red de zonales.", { width: 6526 })] }),
          new TableRow({ children: [cell("Responsable", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Jefe de TI / Administrador de Sistemas.", { width: 6526 })] }),
        ]
      }),

      emptyLine(),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1500, 750, 750, 750, 750, 750, 750, 750, 750, 750, 750, 750, 750],
        rows: [
          new TableRow({ children: [headerCell("Resultado", 1500), headerCell("ENE", 750), headerCell("FEB", 750), headerCell("MAR", 750), headerCell("ABR", 750), headerCell("MAY", 750), headerCell("JUN", 750), headerCell("JUL", 750), headerCell("AGO", 750), headerCell("SET", 750), headerCell("OCT", 750), headerCell("NOV", 750), headerCell("DIC", 750)] }),
          new TableRow({ children: [cell("Verde: >= 99%", { fill: GREEN_LIGHT, width: 1500 }), cell("", { width: 750 }), cell("", { width: 750 }), cell("", { width: 750 }), cell("", { width: 750 }), cell("", { width: 750 }), cell("", { width: 750 }), cell("97%", { fill: YELLOW_LIGHT, width: 750, align: AlignmentType.CENTER }), cell("97.5%", { fill: YELLOW_LIGHT, width: 750, align: AlignmentType.CENTER }), cell("98.5%", { fill: YELLOW_LIGHT, width: 750, align: AlignmentType.CENTER }), cell("99%", { fill: GREEN_LIGHT, width: 750, align: AlignmentType.CENTER }), cell("99.2%", { fill: GREEN_LIGHT, width: 750, align: AlignmentType.CENTER }), cell("99.5%", { fill: GREEN_LIGHT, width: 750, align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("Amarillo: 95%–98.9%", { fill: YELLOW_LIGHT, width: 1500 }), ...Array(12).fill(null).map(() => cell("", { width: 750 }))] }),
          new TableRow({ children: [cell("Rojo: < 95%", { fill: RED_LIGHT, width: 1500 }), ...Array(12).fill(null).map(() => cell("", { width: 750 }))] }),
        ]
      }),

      emptyLine(),
      h2("ANEXO N° 03: FICHAS DE INICIATIVAS"),

      h3("FICHA DE INICIATIVA 1 – PROYECTO GRUPO SURGAS E.I.R.L."),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2500, 6526],
        rows: [
          new TableRow({ children: [cell("Nombre de la Iniciativa", { fill: BLUE_HEADER, bold: true, color: WHITE, width: 2500 }), cell("Implementación de ERP de Inventario de Balones GLP en Tiempo Real", { width: 6526 })] }),
          new TableRow({ children: [cell("Objetivos Generales", { fill: LIGHT_BLUE, bold: true, width: 2500 }), new TableCell({ borders, width: { size: 6526, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [
            new Paragraph({ children: [new TextRun({ text: "1. Automatizar la actualización del stock de balones GLP en tiempo real tras cada despacho.", size: 20, font: "Arial" })] }),
            new Paragraph({ children: [new TextRun({ text: "2. Sincronizar los almacenes de Calana y zonales con el sistema de pedidos para erradicar quiebres.", size: 20, font: "Arial" })] }),
            new Paragraph({ children: [new TextRun({ text: "3. Optimizar los tiempos de auditoría de stock reduciendo procesos manuales del personal.", size: 20, font: "Arial" })] }),
          ] })] }),
        ]
      }),

      emptyLine(),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [4526, 1500, 3000],
        rows: [
          new TableRow({ children: [headerCell("Tarea", 4526), headerCell("Plazo (semanas)", 1500), headerCell("Responsable", 3000)] }),
          new TableRow({ children: [cell("1. Levantamiento de requerimientos del sistema de almacén y pedidos.", { width: 4526 }), cell("1", { width: 1500, align: AlignmentType.CENTER }), cell("Jefe de TI", { width: 3000 })] }),
          new TableRow({ children: [cell("2. Mapeo de campos críticos: tipo de balón, stock, precios, zonal.", { fill: GRAY_LIGHT, width: 4526 }), cell("1", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("Jefe de TI + Jefe de Operaciones", { fill: GRAY_LIGHT, width: 3000 })] }),
          new TableRow({ children: [cell("3. Programación de módulo ERP y API de sincronización con pedidos.", { width: 4526 }), cell("3", { width: 1500, align: AlignmentType.CENTER }), cell("Analista Desarrollador", { width: 3000 })] }),
          new TableRow({ children: [cell("4. Pruebas de carga y validación de consistencia de stock.", { fill: GRAY_LIGHT, width: 4526 }), cell("1", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("QA + Logística", { fill: GRAY_LIGHT, width: 3000 })] }),
          new TableRow({ children: [cell("5. Puesta en producción y capacitación del personal operativo.", { width: 4526 }), cell("1", { width: 1500, align: AlignmentType.CENTER }), cell("Jefe de TI + Jefe de Operaciones", { width: 3000 })] }),
          new TableRow({ children: [cell("TOTAL", { bold: true, width: 4526 }), cell("7 semanas", { bold: true, width: 1500, align: AlignmentType.CENTER }), cell("Inicio: 15/06/2026 | Fin: 04/08/2026 | Costo: S/ 70,000", { width: 3000 })] }),
        ]
      }),

      emptyLine(),
      h3("FICHA DE INICIATIVA 2 – PROYECTO GRUPO SURGAS E.I.R.L."),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2500, 6526],
        rows: [
          new TableRow({ children: [cell("Nombre de la Iniciativa", { fill: BLUE_HEADER, bold: true, color: WHITE, width: 2500 }), cell("Chatbot Inteligente para Pedidos GLP por WhatsApp Business API", { width: 6526 })] }),
          new TableRow({ children: [cell("Objetivos Generales", { fill: LIGHT_BLUE, bold: true, width: 2500 }), new TableCell({ borders, width: { size: 6526, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [
            new Paragraph({ children: [new TextRun({ text: "1. Automatizar la recepción de pedidos de GLP por WhatsApp, respondiendo consultas en menos de 2 minutos.", size: 20, font: "Arial" })] }),
            new Paragraph({ children: [new TextRun({ text: "2. Reducir la carga operativa del call center en preguntas frecuentes y pedidos estándar.", size: 20, font: "Arial" })] }),
            new Paragraph({ children: [new TextRun({ text: "3. Ampliar la atención al cliente a horario extendido 24/7 sin costo adicional de personal.", size: 20, font: "Arial" })] }),
          ] })] }),
        ]
      }),

      emptyLine(),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [4526, 1500, 3000],
        rows: [
          new TableRow({ children: [headerCell("Tarea", 4526), headerCell("Plazo (semanas)", 1500), headerCell("Responsable", 3000)] }),
          new TableRow({ children: [cell("1. Definir flujos de conversación para pedidos frecuentes y FAQs de GLP.", { width: 4526 }), cell("1", { width: 1500, align: AlignmentType.CENTER }), cell("Jefe de TI + Call Center", { width: 3000 })] }),
          new TableRow({ children: [cell("2. Configurar WhatsApp Business API y plataforma de chatbot.", { fill: GRAY_LIGHT, width: 4526 }), cell("2", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("Analista Desarrollador", { fill: GRAY_LIGHT, width: 3000 })] }),
          new TableRow({ children: [cell("3. Integrar chatbot con sistema web de pedidos (SI-001) y API de puntos de venta.", { width: 4526 }), cell("2", { width: 1500, align: AlignmentType.CENTER }), cell("Desarrollador + TI", { width: 3000 })] }),
          new TableRow({ children: [cell("4. Pruebas de usuario con clientes piloto y ajustes.", { fill: GRAY_LIGHT, width: 4526 }), cell("1", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("Especialista Digital + QA", { fill: GRAY_LIGHT, width: 3000 })] }),
          new TableRow({ children: [cell("5. Lanzamiento oficial y capacitación del equipo call center.", { width: 4526 }), cell("1", { width: 1500, align: AlignmentType.CENTER }), cell("Jefe de TI + Supervisor Call Center", { width: 3000 })] }),
          new TableRow({ children: [cell("TOTAL", { bold: true, width: 4526 }), cell("7 semanas", { bold: true, width: 1500, align: AlignmentType.CENTER }), cell("Inicio: 15/07/2026 | Fin: 02/09/2026 | Costo: S/ 60,000", { width: 3000 })] }),
        ]
      }),

      emptyLine(),
      h3("FICHA DE INICIATIVA 3 – PROYECTO GRUPO SURGAS E.I.R.L."),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2500, 6526],
        rows: [
          new TableRow({ children: [cell("Nombre de la Iniciativa", { fill: BLUE_HEADER, bold: true, color: WHITE, width: 2500 }), cell("Fortalecimiento de Ciberseguridad y Protección de Datos SURGAS", { width: 6526 })] }),
          new TableRow({ children: [cell("Objetivos Generales", { fill: LIGHT_BLUE, bold: true, width: 2500 }), new TableCell({ borders, width: { size: 6526, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [
            new Paragraph({ children: [new TextRun({ text: "1. Implementar políticas de seguridad informática, control de accesos y autenticación multifactor (MFA).", size: 20, font: "Arial" })] }),
            new Paragraph({ children: [new TextRun({ text: "2. Capacitar al 100% del personal en ciberseguridad, phishing y protección de datos de clientes.", size: 20, font: "Arial" })] }),
            new Paragraph({ children: [new TextRun({ text: "3. Reducir los riesgos de pérdida de información y accesos no autorizados a sistemas críticos.", size: 20, font: "Arial" })] }),
          ] })] }),
        ]
      }),

      emptyLine(),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [4526, 1500, 3000],
        rows: [
          new TableRow({ children: [headerCell("Tarea", 4526), headerCell("Plazo (semanas)", 1500), headerCell("Responsable", 3000)] }),
          new TableRow({ children: [cell("1. Auditoría de brechas de seguridad en todos los sistemas SURGAS.", { width: 4526 }), cell("1", { width: 1500, align: AlignmentType.CENTER }), cell("Jefe de TI + Auditor externo", { width: 3000 })] }),
          new TableRow({ children: [cell("2. Implementar MFA, políticas de contraseñas y control de accesos por roles.", { fill: GRAY_LIGHT, width: 4526 }), cell("2", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("Administrador de Sistemas", { fill: GRAY_LIGHT, width: 3000 })] }),
          new TableRow({ children: [cell("3. Instalar antivirus corporativo y firewall perimetral en red de zonales.", { width: 4526 }), cell("1", { width: 1500, align: AlignmentType.CENTER }), cell("Administrador de Sistemas", { width: 3000 })] }),
          new TableRow({ children: [cell("4. Capacitar al 100% del personal en taller de ciberseguridad y phishing.", { fill: GRAY_LIGHT, width: 4526 }), cell("1", { fill: GRAY_LIGHT, width: 1500, align: AlignmentType.CENTER }), cell("Jefe de TI + RRHH", { fill: GRAY_LIGHT, width: 3000 })] }),
          new TableRow({ children: [cell("5. Monitoreo continuo y simulacros trimestrales de incidentes de seguridad.", { width: 4526 }), cell("Permanente", { width: 1500, align: AlignmentType.CENTER }), cell("Jefe de TI", { width: 3000 })] }),
          new TableRow({ children: [cell("TOTAL", { bold: true, width: 4526 }), cell("5 semanas", { bold: true, width: 1500, align: AlignmentType.CENTER }), cell("Inicio: 15/08/2026 | Fin: 19/09/2026 | Costo: S/ 75,000", { width: 3000 })] }),
        ]
      }),

      emptyLine(),
      h2("ANEXO N° 04: NECESIDADES DE INFORMACIÓN – FICHAS DETALLADAS"),

      h3("Necesidad 1: Sincronización de Stock de Balones GLP en Tiempo Real"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2500, 6526],
        rows: [
          new TableRow({ children: [cell("Área / Unidad", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Operaciones y Distribución / Planta Calana", { width: 6526 })] }),
          new TableRow({ children: [cell("Responsable", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Jefe de Logística / Operaciones", { width: 6526 })] }),
          new TableRow({ children: [cell("Objetivo relacionado", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("O5: Optimizar el control de stock de balones GLP.", { width: 6526 })] }),
          new TableRow({ children: [cell("Descripción de la necesidad", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Visualizar en tiempo real la cantidad exacta de balones GLP disponibles (5 kg, 10 kg, 45 kg) en planta Calana, zonal Tacna y puntos de distribución Red SURGAS.", { width: 6526 })] }),
          new TableRow({ children: [cell("Uso de la información", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Operativo (consulta en cada despacho y pedido).", { width: 6526 })] }),
          new TableRow({ children: [cell("Indicadores requeridos", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("% discrepancia stock sistema vs físico (TI-O5).", { width: 6526 })] }),
          new TableRow({ children: [cell("Fuente actual", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Consulta manual con el encargado de almacén o llamada telefónica.", { width: 6526 })] }),
          new TableRow({ children: [cell("Periodicidad requerida", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Tiempo real.", { width: 6526 })] }),
          new TableRow({ children: [cell("Formato preferido", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("API consultable desde sistema de pedidos y app móvil.", { width: 6526 })] }),
          new TableRow({ children: [cell("Sistemas involucrados", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("SI-001 (Sistema web de pedidos) y SI-003 (Sistema de puntos canjeables).", { width: 6526 })] }),
          new TableRow({ children: [cell("Nivel de seguridad", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Alto – Datos internos comerciales y operativos.", { width: 6526 })] }),
          new TableRow({ children: [cell("Observaciones", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Fundamental implementar esto antes de lanzar la app móvil y el chatbot de pedidos, para evitar la sobreventa de balones.", { width: 6526 })] }),
        ]
      }),

      emptyLine(),
      h3("Necesidad 2: Métricas de Atención Digital y Tiempos de Pedido"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2500, 6526],
        rows: [
          new TableRow({ children: [cell("Área / Unidad", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Call Center / Atención al Cliente", { width: 6526 })] }),
          new TableRow({ children: [cell("Responsable", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Supervisor de Call Center / Especialista Digital", { width: 6526 })] }),
          new TableRow({ children: [cell("Objetivo relacionado", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("O4/O6: Atención ágil y automatizar pedidos digitales.", { width: 6526 })] }),
          new TableRow({ children: [cell("Descripción de la necesidad", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Monitorear cuánto demoran los asesores en responder pedidos por WhatsApp y llamada, y qué porcentaje de consultas se convierte en pedido confirmado.", { width: 6526 })] }),
          new TableRow({ children: [cell("Uso de la información", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Táctico / Estratégico.", { width: 6526 })] }),
          new TableRow({ children: [cell("Indicadores requeridos", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Tiempo medio de respuesta (<= 5 min) y tasa de conversión de pedidos.", { width: 6526 })] }),
          new TableRow({ children: [cell("Fuente actual", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Inexistente – cálculo humano estimativo.", { width: 6526 })] }),
          new TableRow({ children: [cell("Periodicidad requerida", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Diaria / Semanal.", { width: 6526 })] }),
          new TableRow({ children: [cell("Formato preferido", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Dashboard de call center y tablero de analítica digital.", { width: 6526 })] }),
          new TableRow({ children: [cell("Observaciones", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Reemplazar el WhatsApp manual por plataforma estructurada con chatbot que derive a asesor solo cuando sea necesario.", { width: 6526 })] }),
        ]
      }),

      emptyLine(),
      h3("Necesidad 3: Monitoreo de Disponibilidad de Infraestructura (Uptime)"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2500, 6526],
        rows: [
          new TableRow({ children: [cell("Área / Unidad", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Sistemas y Tecnología (TI)", { width: 6526 })] }),
          new TableRow({ children: [cell("Responsable", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Jefe de TI / Administrador de Sistemas", { width: 6526 })] }),
          new TableRow({ children: [cell("Objetivo relacionado", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("O8: Fortalecer infraestructura TI asegurando continuidad.", { width: 6526 })] }),
          new TableRow({ children: [cell("Descripción de la necesidad", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Visualizar en tiempo real caídas de red, latencia del sistema de pedidos y fallas del servidor de facturación, antes de que el call center quede sin sistema.", { width: 6526 })] }),
          new TableRow({ children: [cell("Uso de la información", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Táctico / Operativo.", { width: 6526 })] }),
          new TableRow({ children: [cell("Indicadores requeridos", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Tasa de disponibilidad Uptime (>= 99%).", { width: 6526 })] }),
          new TableRow({ children: [cell("Fuente actual", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Llamadas de auxilio del personal (reactivo).", { width: 6526 })] }),
          new TableRow({ children: [cell("Periodicidad requerida", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Tiempo real con alertas automatizadas.", { width: 6526 })] }),
          new TableRow({ children: [cell("Formato preferido", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Notificaciones push y tablero técnico TI (Zabbix/Grafana).", { width: 6526 })] }),
          new TableRow({ children: [cell("Observaciones", { fill: LIGHT_BLUE, bold: true, width: 2500 }), cell("Necesidad crítica para migrar a cloud y garantizar operaciones continuas en todos los zonales de Tacna, Moquegua, Ilo y Arequipa.", { width: 6526 })] }),

        ]
      }),

      emptyLine(),
      new Paragraph({
        spacing: { before: 400, after: 200 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "– Fin del Documento –", italic: true, size: 20, font: "Arial", color: "888888" })]
      }),

    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('Caso_Integrador_SURGAS.docx', buffer);
  console.log('Documento generado correctamente.');
}).catch(err => {
  console.error('Error:', err);
  process.exit(1);
});
