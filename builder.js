const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');

// --- ЧТЕНИЕ НАСТРОЕК ---
const SETTINGS_FILE = path.join(__dirname, 'settings.json');
let IN_DIR = path.join(__dirname, 'IN'); // Дефолтное значение

if (fs.existsSync(SETTINGS_FILE)) {
    try {
        const settings = JSON.parse(fs.readFileSync(SETTINGS_FILE, 'utf8'));
        if (settings.inDir) {
            IN_DIR = settings.inDir;
            // Убираем кавычки если вдруг попали при копировании
            IN_DIR = IN_DIR.replace(/^"|"$/g, ''); 
        }
    } catch (e) {
        console.error('Ошибка чтения settings.json, используется дефолтный путь.', e);
    }
}

// --- ЛОГИКА ---
const specificFolder = process.argv[2]; // node builder.js "Папка"

function start() {
    console.log(`Working directory: ${IN_DIR}`);
    
    if (!fs.existsSync(IN_DIR)) {
        console.error(`Папка ${IN_DIR} не найдена.`);
        return;
    }

    let foldersToProcess = [];

    if (specificFolder) {
        // Режим одного файла
        const targetPath = path.join(IN_DIR, specificFolder);
        if (fs.existsSync(targetPath) && fs.statSync(targetPath).isDirectory()) {
            foldersToProcess.push(specificFolder);
            console.log(`🎯 Целевая сборка: "${specificFolder}"`);
        } else {
            console.error(`❌ Папка "${specificFolder}" не найдена в IN/`);
            return;
        }
    } else {
        // Режим "Собрать всё"
        foldersToProcess = fs.readdirSync(IN_DIR).filter(file => 
            fs.statSync(path.join(IN_DIR, file)).isDirectory()
        );
        console.log(`📦 Пакетная сборка: найдено ${foldersToProcess.length} папок.`);
    }

    foldersToProcess.forEach(processFolder);
}

function processFolder(folderName) {
    console.log(`\nProcessing: ${folderName}...`);
    const folderPath = path.join(IN_DIR, folderName);
    const outputPath = path.join(IN_DIR, `${folderName}.docx`); 

    // 1. Собираем файлы docx в папке
    const files = fs.readdirSync(folderPath)
        .filter(f => f.endsWith('.docx') && !f.startsWith('~')) // Игнор временных файлов
        .sort((a, b) => parseFloat(a) - parseFloat(b));

    if (files.length === 0) {
        console.log(`  Skipped (пусто)`);
        return;
    }

    console.log(`  Files: ${files.join(', ')}`);

    // 2. Берем первый файл за основу (Master)
    const masterFile = files[0];
    const masterPath = path.join(folderPath, masterFile);
    
    try {
        const masterBuffer = fs.readFileSync(masterPath);
        const zip = new AdmZip(masterBuffer);
        let masterXml = zip.readAsText("word/document.xml");
        
        const bodyEndIndex = masterXml.lastIndexOf('</w:body>');
        if (bodyEndIndex === -1) {
            console.error('  Error: Invalid Master DOCX (no w:body)');
            return;
        }

        let contentToAppend = '';

        // Проходим по остальным файлам
        for (let i = 1; i < files.length; i++) {
            const partFile = files[i];
            const partPath = path.join(folderPath, partFile);
            
            try {
                const partZip = new AdmZip(partPath);
                let partXml = partZip.readAsText("word/document.xml");

                const startBody = partXml.indexOf('<w:body>') + 8;
                const endBody = partXml.lastIndexOf('</w:body>');
                let bodyContent = partXml.substring(startBody, endBody);

                // Чистка
                bodyContent = bodyContent.replace(/<w:sectPr[^>]*>[\s\S]*?<\/w:sectPr>/g, '');
                bodyContent = bodyContent.replace(/ w14:paraId="[^"]+"/g, '');
                bodyContent = bodyContent.replace(/ w14:textId="[^"]+"/g, '');

                contentToAppend += '<w:p/>' + bodyContent;

            } catch (err) {
                console.error(`  Error reading ${partFile}: ${err.message}`);
            }
        }

        const sectPrIndex = masterXml.lastIndexOf('<w:sectPr');
        let insertPosition = bodyEndIndex;

        if (sectPrIndex > -1 && sectPrIndex < bodyEndIndex) {
            insertPosition = sectPrIndex;
        }

        const finalXml = masterXml.slice(0, insertPosition) + contentToAppend + masterXml.slice(insertPosition);
        zip.updateFile("word/document.xml", Buffer.from(finalXml, 'utf8'));
        
        zip.writeZip(outputPath);
        console.log(`  ✅ Built: ${outputPath}`);
    } catch (e) {
        console.error(`  Fatal error processing folder: ${e.message}`);
    }
}

start();
