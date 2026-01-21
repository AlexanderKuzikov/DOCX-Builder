const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');

const ROOT_DIR = path.join(__dirname, 'IN');
const EMPTY_LINE_XML = '<w:p/>'; 

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

    // 1. Читаем Master
    const masterPath = path.join(inputDir, files[0]);
    let masterZip;
    try { masterZip = new AdmZip(masterPath); } catch (e) { return; }
    let masterXml = masterZip.readAsText('word/document.xml');

    // 2. Ищем точку вставки (самый надежный метод)
    // Находим </w:body>. Всё, что перед ним — это контент + секции.
    // Нам нужно вставить ПЕРЕД первой секцией, которая идет в конце.
    // Секции в конце выглядят как <w:sectPr>...</w:sectPr> (может быть несколько) или <w:sectPr/>
    // Мы ищем позицию первого <w:sectPr в "хвосте" body.
    
    const bodyEndIndex = masterXml.lastIndexOf('</w:body>');
    if (bodyEndIndex === -1) return;

    // Берем хвост (3000 символов)
    const tail = masterXml.substring(Math.max(0, bodyEndIndex - 3000), bodyEndIndex);
    
    // Ищем ПЕРВЫЙ w:sectPr в этом хвосте.
    // Регулярка ищет <w:sectPr с начала строки или после тега.
    const sectPrMatch = tail.match(/<w:sectPr/);
    
    let insertIndex = bodyEndIndex;
    if (sectPrMatch) {
        // Если нашли sectPr, то точка вставки = начало хвоста + индекс найденного sectPr
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
                
                // Парсим тело
                const start = xml.indexOf('<w:body');
                const end = xml.lastIndexOf('</w:body>');

                if (start !== -1 && end !== -1) {
                    const bodyTagClose = xml.indexOf('>', start);
                    if (bodyTagClose !== -1 && bodyTagClose < end) {
                        let content = xml.substring(bodyTagClose + 1, end);
                        
                        // Очистка
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
    
    // 1. Удаляем ВСЕ sectPr (и полные, и короткие, и с атрибутами)
    // Важно: жадность регулярки. Ищем от <w:sectPr до </w:sectPr> или />
    c = c.replace(/<w:sectPr[\s\S]*?<\/w:sectPr>/g, '');
    c = c.replace(/<w:sectPr[\s\S]*?\/>/g, '');

    // 2. Удаляем конфликтующие атрибуты (грубый метод без кавычек)
    // Удаляем w14:paraId="..." и w:rsid...="..."
    
    const attrs = [
        'w14:paraId', 'w14:textId',
        'w:rsidR', 'w:rsidRDefault', 'w:rsidP', 'w:rsidRPr',
        'w:id'
    ];
    
    attrs.forEach(attr => {
        // Ищем: (пробел или нет) + имя + ="..." 
        const regex = new RegExp(`${attr}="[^"]*"`, 'g');
        c = c.replace(regex, '');
        // И с одинарными кавычками
        const regexSingle = new RegExp(`${attr}='[^']*'`, 'g');
        c = c.replace(regexSingle, '');
    });

    return c;
}

processBatches();
