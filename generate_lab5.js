const { Document, Packer, Paragraph, TextRun, HeadingLevel } = require("docx");
const fs = require("fs");
const { sections, FUENTE } = require("./lab5.js");

const children = [];

sections.forEach(sec => {
    children.push(new Paragraph({ text: sec.title, heading: HeadingLevel.HEADING_1 }));
    if (sec.intro) {
        sec.intro.forEach(p => {
            children.push(new Paragraph({ text: p }));
        });
    }
    if (sec.figures) {
        sec.figures.forEach((fig, index) => {
            children.push(new Paragraph({ text: fig.lead }));
            
            if (fig.cmd && fig.cmd.length > 0) {
                const cmdRuns = [];
                fig.cmd.forEach((cmdLine, i) => {
                    cmdRuns.push(new TextRun({
                        text: cmdLine,
                        font: "Courier New",
                        break: i > 0 ? 1 : 0
                    }));
                });
                children.push(new Paragraph({
                    children: cmdRuns
                }));
            }
            
            children.push(new Paragraph({ text: `Figura ${index + 1}: ${fig.title}`, heading: HeadingLevel.HEADING_2 }));
            if (fig.desc) {
                children.push(new Paragraph({ text: fig.desc }));
            }
            if (fig.note) {
                children.push(new Paragraph({ 
                    children: [
                        new TextRun({ text: "Nota: ", bold: true }),
                        new TextRun({ text: fig.note })
                    ] 
                }));
            }
            children.push(new Paragraph({ text: "" })); // Empty line
        });
    }
});

children.push(new Paragraph({ text: FUENTE, alignment: "center" }));

const doc = new Document({
    sections: [{
        children: children
    }]
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("lab5.docx", buffer);
    console.log("Document created successfully");
});
