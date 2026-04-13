const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');

async function run() {
    try {
        const outPdf = await PDFDocument.create();
        for(let i=0; i<1; i++) {
            const templatePdfDoc = await PDFDocument.create();
            templatePdfDoc.addPage([595, 841]);
            
            const font = await templatePdfDoc.embedFont(StandardFonts.Helvetica);
            const pages = templatePdfDoc.getPages();
            
            pages[0].drawText('Test', {
                x: 100, y: 100, font: font, size: 11, color: rgb(1, 0, 0)
            });
            
            const copiedPages = await outPdf.copyPages(templatePdfDoc, templatePdfDoc.getPageIndices());
            copiedPages.forEach(p => outPdf.addPage(p));
        }
        await outPdf.save();
        console.log("SUCCESS");
    } catch (e) {
        console.error("ERROR", e);
    }
}

run();
