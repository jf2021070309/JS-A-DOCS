const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "CartaPro — Emprendimiento de Cartas para Restaurantes";

// Color palette (matching the HTML)
const C = {
    bg: "1a1208",
    bgDeep: "0d0a04",
    surface: "2a1f0f",
    text: "fdf6e3",
    textSec: "d4b896",
    textMut: "9b7f5e",
    acc1: "e8a020",  // gold
    acc2: "f5d87a",  // yellow
    acc3: "c0392b",  // red
    cardBg: "251c0e",
    cardBor: "3a2d1a",
    white: "FFFFFF",
    redLight: "3d1510",
    goldLight: "2e1f08",
};

const makeShadow = () => ({ type: "outer", color: "000000", blur: 8, offset: 3, angle: 45, opacity: 0.25 });

// Helper: dark card background shape
function addCard(slide, x, y, w, h, opts = {}) {
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y, w, h,
        fill: { color: opts.fillColor || C.cardBg },
        line: { color: opts.lineColor || C.cardBor, width: 1 },
        rectRadius: 0.12,
        shadow: makeShadow(),
    });
}

// Helper: section label (small caps label above title)
function addLabel(slide, text, y = 0.35) {
    slide.addText(text.toUpperCase(), {
        x: 0.5, y, w: 9, h: 0.28,
        fontFace: "Calibri", fontSize: 9, bold: true,
        color: C.acc1, charSpacing: 4, align: "left",
    });
}

// Helper: slide title
function addTitle(slide, text, y = 0.62, opts = {}) {
    slide.addText(text, {
        x: 0.5, y, w: 9, h: opts.h || 0.7,
        fontFace: "Cambria", fontSize: opts.size || 30, bold: true,
        color: opts.color || C.text, align: "left",
        margin: 0,
    });
}

// ─────────────────────────────────────────────
// SLIDE 1 — PORTADA
// ─────────────────────────────────────────────
{
    const s = pres.addSlide();
    s.background = { color: C.bg };

    // Decorative gold circle (top-right glow)
    s.addShape(pres.shapes.OVAL, {
        x: 7.5, y: -1.2, w: 4, h: 4,
        fill: { color: C.acc1, transparency: 85 },
        line: { color: C.acc1, transparency: 85 },
    });
    // Red blob bottom-left
    s.addShape(pres.shapes.OVAL, {
        x: -1, y: 3.5, w: 3.5, h: 3.5,
        fill: { color: C.acc3, transparency: 88 },
        line: { color: C.acc3, transparency: 88 },
    });

    // Emoji
    s.addText("🍽️", {
        x: 3.5, y: 0.55, w: 3, h: 1.1,
        fontFace: "Segoe UI Emoji", fontSize: 52,
        align: "center",
    });

    // Brand name
    s.addText("CartaPro", {
        x: 1.5, y: 1.6, w: 7, h: 1.3,
        fontFace: "Cambria", fontSize: 64, bold: true,
        color: C.acc1, align: "center",
    });

    // Tagline uppercase
    s.addText("Cartas para Restaurantes", {
        x: 1.5, y: 2.85, w: 7, h: 0.45,
        fontFace: "Calibri", fontSize: 13, bold: false,
        color: C.textMut, charSpacing: 5, align: "center",
    });

    // Sub tagline
    s.addText("Carta impresa · QR para mesas · Landing page digital", {
        x: 1.5, y: 3.35, w: 7, h: 0.5,
        fontFace: "Calibri", fontSize: 15,
        color: C.textSec, align: "center",
    });

    // Bottom pill
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 3.2, y: 4.25, w: 3.6, h: 0.55,
        fill: { color: C.acc1, transparency: 85 },
        line: { color: C.acc1, transparency: 60 },
        rectRadius: 0.28,
    });
    s.addText("Tacna, Perú · 2025", {
        x: 3.2, y: 4.27, w: 3.6, h: 0.5,
        fontFace: "Calibri", fontSize: 11,
        color: C.acc2, align: "center",
    });
}

// ─────────────────────────────────────────────
// SLIDE 2 — EL PROBLEMA
// ─────────────────────────────────────────────
{
    const s = pres.addSlide();
    s.background = { color: C.bg };

    // Blobs
    s.addShape(pres.shapes.OVAL, { x: 7, y: -1, w: 4, h: 4, fill: { color: C.acc3, transparency: 90 }, line: { color: C.acc3, transparency: 90 } });

    addLabel(s, "El Problema");
    addTitle(s, "¿Cómo muestran su carta hoy la mayoría de restaurantes?", 0.62, { h: 0.85, size: 26 });

    const cards = [
        { emoji: "📸", title: "Foto por WhatsApp", body: "Una imagen JPG que mandan por WhatsApp o postean en Facebook. Sin URL, sin calidad.", x: 0.45 },
        { emoji: "📄", title: "Carta solo en papel", body: "Diseño desactualizado, sin QR, sin versión digital. El cliente no la puede ver antes de llegar.", x: 3.58 },
        { emoji: "🔄", title: "Sin actualización", body: "Cambiar un precio significa rediseñar e imprimir todo de nuevo. Costoso y lento.", x: 6.72 },
    ];

    cards.forEach(card => {
        addCard(s, card.x, 1.55, 2.95, 2.15, { lineColor: "3d2a14" });
        s.addText(card.emoji, { x: card.x, y: 1.62, w: 2.95, h: 0.6, fontFace: "Segoe UI Emoji", fontSize: 28, align: "center" });
        s.addText(card.title, { x: card.x + 0.12, y: 2.22, w: 2.7, h: 0.38, fontFace: "Calibri", fontSize: 12, bold: true, color: C.acc2, align: "center" });
        s.addText(card.body, { x: card.x + 0.1, y: 2.6, w: 2.72, h: 0.95, fontFace: "Calibri", fontSize: 10.5, color: C.textMut, align: "center", valign: "top" });
    });

    // Case study callout
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.45, y: 3.88, w: 9.1, h: 0.9,
        fill: { color: C.redLight },
        line: { color: C.acc3, width: 1 },
        rectRadius: 0.1,
    });
    s.addText([
        { text: "Caso real: ", options: { color: C.textSec } },
        { text: "D'Brazza Dorada", options: { bold: true, color: C.acc2 } },
        { text: " (pollería en Tacna) usa una imagen JPG como carta de delivery — sin landing page, sin QR en mesas, sin URL para redes sociales.", options: { color: C.textSec } },
    ], { x: 0.65, y: 3.92, w: 8.7, h: 0.8, fontFace: "Calibri", fontSize: 12, valign: "middle" });
}

// ─────────────────────────────────────────────
// SLIDE 3 — LA SOLUCIÓN
// ─────────────────────────────────────────────
{
    const s = pres.addSlide();
    s.background = { color: C.bg };

    s.addShape(pres.shapes.OVAL, { x: -0.8, y: -0.8, w: 3.5, h: 3.5, fill: { color: C.acc1, transparency: 88 }, line: { color: C.acc1, transparency: 88 } });

    addLabel(s, "La Solución");
    addTitle(s, "Un solo paquete, todo resuelto", 0.62, { size: 28 });

    s.addText('"Diseñamos y entregamos la carta de tu restaurante, en papel y en internet, todo en menos de una semana."', {
        x: 0.5, y: 1.3, w: 9, h: 0.45,
        fontFace: "Calibri", fontSize: 12, italic: true,
        color: C.textMut, align: "left",
    });

    const cols = [
        { emoji: "🖨️", title: "Carta Impresa", body: "PDF profesional diseñado según la identidad del local. Listo para cualquier imprenta." },
        { emoji: "📱", title: "QR para Mesas", body: "Stickers o tarjetas con código QR único. El cliente lo escanea y ve la carta al instante." },
        { emoji: "🌐", title: "Landing Page", body: "Página web con la carta digital. URL propia para compartir en Instagram, Facebook y TikTok." },
    ];

    cols.forEach((col, i) => {
        const x = 0.45 + i * 3.13;
        addCard(s, x, 1.9, 2.95, 2.7, { lineColor: "3d2a14" });
        // Gold circle icon bg
        s.addShape(pres.shapes.OVAL, { x: x + 1.05, y: 2.0, w: 0.85, h: 0.85, fill: { color: C.goldLight }, line: { color: C.acc1, width: 1 } });
        s.addText(col.emoji, { x: x + 1.05, y: 2.0, w: 0.85, h: 0.85, fontFace: "Segoe UI Emoji", fontSize: 22, align: "center", valign: "middle" });
        s.addText(col.title, { x: x + 0.1, y: 2.93, w: 2.75, h: 0.38, fontFace: "Cambria", fontSize: 13, bold: true, color: C.acc1, align: "center" });
        s.addText(col.body, { x: x + 0.12, y: 3.32, w: 2.7, h: 1.15, fontFace: "Calibri", fontSize: 11, color: C.textMut, align: "center", valign: "top" });
    });
}

// ─────────────────────────────────────────────
// SLIDE 4 — ANTES vs DESPUÉS
// ─────────────────────────────────────────────
{
    const s = pres.addSlide();
    s.background = { color: C.bg };

    addLabel(s, "Caso: D'Brazza Dorada");
    addTitle(s, "Antes vs. Después", 0.62, { size: 28 });

    // ANTES card
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.45, y: 1.42, w: 4.45, h: 3.52,
        fill: { color: C.redLight }, line: { color: C.acc3, width: 1 }, rectRadius: 0.12,
    });
    s.addText("❌  ANTES", { x: 0.65, y: 1.55, w: 4, h: 0.36, fontFace: "Calibri", fontSize: 11, bold: true, color: "e74c3c", charSpacing: 2 });

    const before = [
        "Carta como imagen JPG por WhatsApp",
        "Sin código QR en las mesas",
        "Sin URL para Instagram / Facebook",
        "Actualizar precios = rediseñar la imagen",
        "El cliente no ve la carta antes de llegar",
    ];
    s.addText(before.map((t, i) => ({
        text: t,
        options: { bullet: true, breakLine: i < before.length - 1, color: C.textSec, fontSize: 12, fontFace: "Calibri", paraSpaceAfter: 6 },
    })), { x: 0.75, y: 1.96, w: 4.0, h: 2.8, fontFace: "Calibri", fontSize: 12, color: C.textSec });

    // DESPUÉS card
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 5.1, y: 1.42, w: 4.45, h: 3.52,
        fill: { color: C.goldLight }, line: { color: C.acc1, width: 1 }, rectRadius: 0.12,
    });
    s.addText("✅  CON CARTAPRO", { x: 5.3, y: 1.55, w: 4, h: 0.36, fontFace: "Calibri", fontSize: 11, bold: true, color: C.acc1, charSpacing: 2 });

    const after = [
        "PDF profesional + copias físicas impresas",
        "QR stickers en cada mesa del local",
        "URL propia en bio de IG, FB y TikTok",
        "Actualización de precios en línea (S/ 35)",
        "Carta visible antes de llegar al local",
    ];
    s.addText(after.map((t, i) => ({
        text: t,
        options: { bullet: true, breakLine: i < after.length - 1, color: C.textSec, fontSize: 12, fontFace: "Calibri", paraSpaceAfter: 6 },
    })), { x: 5.3, y: 1.96, w: 4.1, h: 2.8, fontFace: "Calibri", fontSize: 12, color: C.textSec });
}

// ─────────────────────────────────────────────
// SLIDE 5 — PAQUETES + PRECIOS
// ─────────────────────────────────────────────
{
    const s = pres.addSlide();
    s.background = { color: C.bg };

    s.addShape(pres.shapes.OVAL, { x: 7.5, y: -1, w: 3.5, h: 3.5, fill: { color: C.acc2, transparency: 90 }, line: { color: C.acc2, transparency: 90 } });

    addLabel(s, "Fuentes de Ingresos");
    addTitle(s, "Solo 2 paquetes + add-on", 0.62, { size: 28 });

    // Pack Básico
    addCard(s, 0.45, 1.42, 4.45, 2.85);
    s.addText("PACK BÁSICO", { x: 0.65, y: 1.55, w: 4, h: 0.3, fontFace: "Calibri", fontSize: 9, bold: true, color: C.textMut, charSpacing: 3 });
    s.addText("S/ 150 – 180", { x: 0.65, y: 1.85, w: 4, h: 0.65, fontFace: "Cambria", fontSize: 30, bold: true, color: C.acc1 });
    const basic = ["Diseño carta PDF", "10 cartas impresas", "QR stickers (5 unidades)", "— Sin landing page"];
    s.addText(basic.map((t, i) => ({
        text: t,
        options: { bullet: i < 3, breakLine: i < basic.length - 1, color: i === 3 ? C.textMut : C.textSec, fontSize: 12, fontFace: "Calibri", paraSpaceAfter: 5 },
    })), { x: 0.7, y: 2.55, w: 4.0, h: 1.55, fontFace: "Calibri", fontSize: 12 });

    // Pack Completo (highlighted)
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 5.1, y: 1.42, w: 4.45, h: 2.85,
        fill: { color: "261d09" }, line: { color: C.acc1, width: 2 }, rectRadius: 0.12,
        shadow: makeShadow(),
    });
    s.addText("⭐  PACK COMPLETO", { x: 5.3, y: 1.55, w: 4, h: 0.3, fontFace: "Calibri", fontSize: 9, bold: true, color: C.acc1, charSpacing: 3 });
    s.addText("S/ 350 – 400", { x: 5.3, y: 1.85, w: 4, h: 0.65, fontFace: "Cambria", fontSize: 30, bold: true, color: C.acc2 });
    const full = ["Todo el Básico incluido", "Landing page carta digital", "Dominio propio 1 año", "QR tarjetas (10 unidades)"];
    s.addText(full.map((t, i) => ({
        text: t,
        options: { bullet: true, breakLine: i < full.length - 1, color: C.textSec, fontSize: 12, fontFace: "Calibri", paraSpaceAfter: 5 },
    })), { x: 5.3, y: 2.55, w: 4.1, h: 1.55, fontFace: "Calibri", fontSize: 12 });

    // Add-on
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.45, y: 4.42, w: 9.1, h: 0.72,
        fill: { color: C.surface }, line: { color: C.cardBor, width: 1 }, rectRadius: 0.1,
    });
    s.addText("🔄  Add-on: Actualización de precios / platos en carta digital", {
        x: 0.7, y: 4.48, w: 7.2, h: 0.6,
        fontFace: "Calibri", fontSize: 12.5, color: C.textSec, valign: "middle",
    });
    s.addText("S/ 35", {
        x: 7.9, y: 4.48, w: 1.4, h: 0.6,
        fontFace: "Cambria", fontSize: 18, bold: true, color: C.acc1, align: "right", valign: "middle",
    });
}

// ─────────────────────────────────────────────
// SLIDE 6 — BUSINESS MODEL CANVAS
// ─────────────────────────────────────────────
{
    const s = pres.addSlide();
    s.background = { color: C.bg };

    addLabel(s, "Modelo de Negocio");
    addTitle(s, "Business Model Canvas", 0.62, { size: 26 });

    const blocks = [
        { title: "Socios Clave", body: "Imprentas locales Tacna · Hosting/dominio (Cloudflare, Railway) · Freelancers diseño/foto", col: C.acc1 },
        { title: "Actividades Clave", body: "Diseño carta PDF · Generación QR · Desarrollo landing page · Brief onboarding · Soporte WhatsApp", col: C.acc2 },
        { title: "Propuesta de Valor", body: "Carta impresa + QR en mesas + landing page digital, todo en un solo proveedor. Diseño profesional, entrega en 3–5 días, URL compartible en redes.", col: C.acc1 },
        { title: "Canales", body: "WhatsApp · Visitas directas · Instagram / TikTok · Facebook · Boca a boca", col: C.acc2 },
        { title: "Segmentos", body: "Pollerías, cafeterías, cevicherías, food trucks en Tacna · Pequeños y medianos locales", col: "e74c3c" },
        { title: "Recursos Clave", body: "Infraestructura propia · Figma / Canva Pro · Plantillas base · Portafolio real", col: C.acc2 },
        { title: "Ingresos", body: "Pack Básico S/150–180 · Pack Completo S/350–400 · Actualización S/35", col: C.acc1 },
        { title: "Estructura de Costos", body: "Impresión subcontratada · Hosting/dominio ~S/20–30/año · Canva Pro ~S/50/mes · Tiempo de diseño (costo principal)", col: "e74c3c" },
    ];

    // 4-column layout: [Socios|Actividades|Valor|Canales|Segmentos] × 2 rows + bottom
    const layout = [
        // row 1: 4 cols
        { i: 0, x: 0.3, y: 1.42, w: 2.15, h: 1.42 },
        { i: 1, x: 2.55, y: 1.42, w: 2.15, h: 1.42 },
        { i: 2, x: 4.8, y: 1.42, w: 2.15, h: 1.42 },
        { i: 3, x: 7.05, y: 1.42, w: 2.65, h: 1.42 },
        // row 2
        { i: 4, x: 0.3, y: 2.94, w: 2.15, h: 1.35 },
        { i: 5, x: 2.55, y: 2.94, w: 2.15, h: 1.35 },
        { i: 6, x: 4.8, y: 2.94, w: 2.15, h: 1.35 },
        { i: 7, x: 7.05, y: 2.94, w: 2.65, h: 1.35 },
    ];

    layout.forEach(pos => {
        const b = blocks[pos.i];
        addCard(s, pos.x, pos.y, pos.w, pos.h);
        s.addText(b.title, {
            x: pos.x + 0.1, y: pos.y + 0.08, w: pos.w - 0.2, h: 0.26,
            fontFace: "Calibri", fontSize: 8, bold: true, color: b.col, charSpacing: 1,
        });
        s.addText(b.body, {
            x: pos.x + 0.1, y: pos.y + 0.35, w: pos.w - 0.2, h: pos.h - 0.42,
            fontFace: "Calibri", fontSize: 9, color: C.textSec, valign: "top",
        });
    });

    // Relation / Customer relationship note
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.3, y: 4.38, w: 9.4, h: 0.58,
        fill: { color: C.surface }, line: { color: C.cardBor }, rectRadius: 0.08,
    });
    s.addText("💬  Relación con clientes: Soporte por WhatsApp · Seguimiento post-entrega · Venta de actualizaciones recurrentes", {
        x: 0.5, y: 4.41, w: 9.0, h: 0.52,
        fontFace: "Calibri", fontSize: 10.5, color: C.textMut, valign: "middle",
    });
}

// ─────────────────────────────────────────────
// SLIDE 7 — ARTEFACTOS ENTREGABLES
// ─────────────────────────────────────────────
{
    const s = pres.addSlide();
    s.background = { color: C.bg };

    s.addShape(pres.shapes.OVAL, { x: 7, y: -1, w: 4, h: 4, fill: { color: C.acc1, transparency: 90 }, line: { color: C.acc1, transparency: 90 } });

    addLabel(s, "Entregables");
    addTitle(s, "¿Qué recibe el cliente?", 0.62, { size: 28 });

    const items = [
        { num: "1", emoji: "📄", title: "Carta en PDF", body: "Diseño profesional con logo, colores y fotos del local. Listo para imprimir en A4, A5 o díptico." },
        { num: "2", emoji: "🖨️", title: "Copias físicas impresas", body: "10 o 20 copias plastificadas según el paquete, subcontratadas a imprenta local en Tacna." },
        { num: "3", emoji: "📲", title: "QR sticker/tarjeta para mesas", body: "Código QR único y permanente por local. El cliente lo escanea en la mesa y ve la carta digital." },
        { num: "4", emoji: "🌐", title: "Landing page con carta digital", body: "Página web responsiva con URL propia. Compartible por WhatsApp, Instagram, Facebook y TikTok." },
    ];

    items.forEach((item, i) => {
        const y = 1.42 + i * 0.98;
        addCard(s, 0.45, y, 9.1, 0.88);
        // Number circle
        s.addShape(pres.shapes.OVAL, { x: 0.62, y: y + 0.15, w: 0.55, h: 0.55, fill: { color: C.acc1, transparency: 60 }, line: { color: C.acc1 } });
        s.addText(item.num, { x: 0.62, y: y + 0.15, w: 0.55, h: 0.55, fontFace: "Cambria", fontSize: 14, bold: true, color: C.acc2, align: "center", valign: "middle" });
        // Emoji
        s.addText(item.emoji, { x: 1.28, y: y + 0.12, w: 0.65, h: 0.62, fontFace: "Segoe UI Emoji", fontSize: 24, align: "center", valign: "middle" });
        // Title
        s.addText(item.title, { x: 2.05, y: y + 0.1, w: 4.8, h: 0.32, fontFace: "Calibri", fontSize: 13, bold: true, color: C.acc2 });
        // Body
        s.addText(item.body, { x: 2.05, y: y + 0.42, w: 7.2, h: 0.38, fontFace: "Calibri", fontSize: 11, color: C.textMut });
    });
}

// ─────────────────────────────────────────────
// SLIDE 8 — CIERRE
// ─────────────────────────────────────────────
{
    const s = pres.addSlide();
    s.background = { color: C.bg };

    // Glow blobs
    s.addShape(pres.shapes.OVAL, { x: 6.5, y: -1.5, w: 5, h: 5, fill: { color: C.acc1, transparency: 85 }, line: { color: C.acc1, transparency: 85 } });
    s.addShape(pres.shapes.OVAL, { x: -1.5, y: 3, w: 4, h: 4, fill: { color: C.acc3, transparency: 88 }, line: { color: C.acc3, transparency: 88 } });

    // Rocket emoji
    s.addText("🚀", { x: 4.0, y: 0.28, w: 2, h: 0.9, fontFace: "Segoe UI Emoji", fontSize: 38, align: "center" });

    // Brand
    s.addText("CartaPro", { x: 1.5, y: 1.15, w: 7, h: 1.05, fontFace: "Cambria", fontSize: 54, bold: true, color: C.acc1, align: "center" });

    // Pitch
    s.addText('"Si tu carta es solo una foto de WhatsApp, nosotros te damos carta impresa, QR en tus mesas y página web — todo listo en menos de una semana."', {
        x: 1.0, y: 2.2, w: 8, h: 0.85,
        fontFace: "Calibri", fontSize: 13.5, italic: true,
        color: C.textSec, align: "center",
    });

    // 3 stat callouts
    const stats = [
        { val: "S/ 150", label: "Desde" },
        { val: "5 días", label: "Entrega" },
        { val: "3-en-1", label: "Paquete" },
    ];

    stats.forEach((st, i) => {
        const x = 1.8 + i * 2.45;
        s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
            x, y: 3.22, w: 2.1, h: 1.42,
            fill: { color: C.goldLight }, line: { color: C.acc1, width: 1 }, rectRadius: 0.12,
            shadow: makeShadow(),
        });
        s.addText(st.val, { x, y: 3.32, w: 2.1, h: 0.7, fontFace: "Cambria", fontSize: 26, bold: true, color: C.acc1, align: "center" });
        s.addText(st.label.toUpperCase(), { x, y: 4.02, w: 2.1, h: 0.4, fontFace: "Calibri", fontSize: 9, color: C.textMut, charSpacing: 2, align: "center" });
    });

    // CTA
    s.addText("cartapro.pe · WhatsApp · Instagram", {
        x: 2.5, y: 4.8, w: 5, h: 0.35,
        fontFace: "Calibri", fontSize: 10.5, color: C.textMut, charSpacing: 1, align: "center",
    });
}

// ─────────────────────────────────────────────
// WRITE FILE
// ─────────────────────────────────────────────
pres.writeFile({ fileName: "CartaPro.pptx" }).then(() => {
    console.log("¡PPTX Generado exitosamente como CartaPro.pptx!");
}).catch(e => console.error(e));