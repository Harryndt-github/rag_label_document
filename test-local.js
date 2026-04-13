const { PDFDocument } = require('pdf-lib');

(async () => {
    try {
        const outPdf = await PDFDocument.create();
        for(let i=0; i<3; i++) {
            const templatePdfDoc = await PDFDocument.create();
            templatePdfDoc.addPage([595, 841]);
            
            const copiedPages = await outPdf.copyPages(templatePdfDoc, templatePdfDoc.getPageIndices());
            copiedPages.forEach(p => outPdf.addPage(p));
        }
        await outPdf.save();
        console.log("Success");
    } catch(e) {
        console.error("Failed:", e.message);
    }
})();
