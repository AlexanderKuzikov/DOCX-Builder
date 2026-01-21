const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');

const ROOT_DIR = path.join(__dirname, 'IN');
const EMPTY_LINE_XML = '<w:p><w:pPr><w:keepNext/></w:pPr></w:p>'; 

function processBatches() {
    if (!fs.existsSync(ROOT_DIR)) {
        console.error(`Папка ${ROOT_DIR} не найдена!`);
        return;
    }
    const batches = fs.readdirSync(ROOT_DIR).filter(file => fs.statSync(path.join(ROOT_DIR, file)).isDirectory());
    if (batches.length === 0) {
        console.log('Нет папок в IN.');
        return;
    }
    console.log(`Найдено комплектов: ${batches.length}`);
    batches.forEach(batchName => processSingleBatch(path.join(ROOT_DIR, batchName), batchName));
}

function processSingleBatch(inputDir, folderName) {
    const files = fs.readdirSync(inputDir)
        .filter(file => file.endsWith('.docx') && /^\d/.test(file))
        .sort((a, b) => parseInt(a) - parseInt(b));

    if (files.length === 0) return;

    const masterPath = path.join(inputDir, files[0]);
    let masterZip;
    try { masterZip = new AdmZip(masterPath); } catch (e) { return; }
    let masterXml = masterZip.readAsText('word/document.xml');

    const bodyEndIndex = masterXml.lastIndexOf('</w:body>');
    if (bodyEndIndex === -1) return;

    const tail = masterXml.substring(Math.max(0, bodyEndIndex - 3000), bodyEndIndex);
    const sectPrMatch = tail.match(/<w:sectPr/);
    
    let insertIndex = bodyEndIndex;
    if (sectPrMatch) {
        insertIndex = (Math.max(0, bodyEndIndex - 3000)) + sectPrMatch.index;
    }

    let contentToAppend = '';

    if (files.length > 1) {
        console.log(`>>> Обработка: "${folderName}"`);
        for (let i = 1; i < files.length; i++) {
            const filePath = path.join(inputDir, files[i]);
            try {
                const zip = new AdmZip(filePath);
                const xml = zip.readAsText('word/document.xml');
                
                const start = xml.indexOf('<w:body');
                const end = xml.lastIndexOf('</w:body>');

                if (start !== -1 && end !== -1) {
                    const bodyTagClose = xml.indexOf('>', start);
                    if (bodyTagClose !== -1 && bodyTagClose < end) {
                        let content = xml.substring(bodyTagClose + 1, end);
                        
                        content = cleanContent(content);
                        contentToAppend += EMPTY_LINE_XML + content;
                    }
                }
            } catch (err) { }
        }
        
        const finalXml = masterXml.slice(0, insertIndex) + contentToAppend + masterXml.slice(insertIndex);
        masterZip.updateFile('word/document.xml', Buffer.from(finalXml, 'utf-8'));
    }

    const outputFilePath = path.join(ROOT_DIR, `${folderName}.docx`);
    masterZip.writeZip(outputFilePath);
    console.log(`[OK] Saved: ${folderName}.docx`);
}

function cleanContent(xml) {
    let c = xml;
    
    // 1. sectPr (безопасные регулярки без new RegExp)
    c = c.replace(/<w:sectPr[\s\S]*?<\/w:sectPr>/g, '');
    c = c.replace(/<w:sectPr[\s\S]*?\/>/g, '');

    // 2. Атрибуты ID
    c = c.replace(/w14:paraId=["'][^"']*["']/g, '');
    c = c.replace(/w14:textId=["'][^"']*["']/g, '');

    // 3. Версионность
    c = c.replace(/w:rsidR=["'][^"']*["']/g, '');
    c = c.replace(/w:rsidRDefault=["'][^"']*["']/g, '');
    c = c.replace(/w:rsidP=["'][^"']*["']/g, '');
    c = c.replace(/w:rsidRPr=["'][^"']*["']/g, '');

    // 4. w:id (на всякий случай убираем, раз с ним открывалось)
    // Если позже захотим стили чинить - уберем эту строку.
    c = c.replace(/w:id=["'][^"']*["']/g, '');

    return c;
}

processBatches();
