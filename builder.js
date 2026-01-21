const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');

// ПУТЬ К ПАПКЕ С КОМПЛЕКТАМИ
const ROOT_DIR = path.join(__dirname, 'IN');
// XML для пустой строки (пустого параграфа)
const EMPTY_LINE_XML = '<w:p/>'; 

function processBatches() {
    // 1. Проверяем существование папки IN
    if (!fs.existsSync(ROOT_DIR)) {
        console.error(`Папка ${ROOT_DIR} не найдена! Создайте её и положите туда папки с документами.`);
        return;
    }

    // 2. Получаем список всех папок внутри IN (каждая папка - один комплект)
    const batches = fs.readdirSync(ROOT_DIR).filter(file => {
        return fs.statSync(path.join(ROOT_DIR, file)).isDirectory();
    });

    if (batches.length === 0) {
        console.log('В папке IN нет подпапок для обработки.');
        return;
    }

    console.log(`Найдено комплектов: ${batches.length}`);

    // 3. Обрабатываем каждую папку отдельно
    batches.forEach(batchName => {
        const batchDir = path.join(ROOT_DIR, batchName);
        processSingleBatch(batchDir, batchName);
    });
}

function processSingleBatch(inputDir, folderName) {
    // Получаем список docx файлов, начинающихся с цифры
    const files = fs.readdirSync(inputDir)
        .filter(file => file.endsWith('.docx') && /^\d/.test(file))
        .sort((a, b) => parseInt(a) - parseInt(b));

    if (files.length === 0) {
        console.warn(`[SKIP] В папке "${folderName}" нет подходящих docx файлов.`);
        return;
    }

    // 1. Берем первый файл как MASTER (основу)
    const masterPath = path.join(inputDir, files[0]);
    let masterZip;
    
    try {
        masterZip = new AdmZip(masterPath);
    } catch (e) {
        console.error(`[ERROR] Не удалось открыть мастер-файл в "${folderName}": ${e.message}`);
        return;
    }

    let masterXml = masterZip.readAsText('word/document.xml');

    // Находим точку вставки в Master (перед последним sectPr или перед закрытием body)
    let insertIndex = masterXml.lastIndexOf('<w:sectPr');
    if (insertIndex === -1) {
        insertIndex = masterXml.lastIndexOf('</w:body>');
    }
    
    if (insertIndex === -1) {
        console.error(`[ERROR] Структура файла ${files[0]} повреждена (нет body).`);
        return;
    }

    let contentToAppend = '';

    // 2. Проходим по остальным файлам
    if (files.length > 1) {
        console.log(`>>> Обработка комплекта: "${folderName}" (${files.length} файлов)`);
        
        for (let i = 1; i < files.length; i++) {
            const filePath = path.join(inputDir, files[i]);
            try {
                const zip = new AdmZip(filePath);
                const xml = zip.readAsText('word/document.xml');

                // Вырезаем содержимое body
                const bodyMatch = xml.match(/<w:body[^>]*>([\s\S]*?)<\/w:body>/);

                if (bodyMatch && bodyMatch[1]) {
                    let content = bodyMatch[1];

                    // Удаляем настройки секций (sectPr), чтобы текст просто лился дальше
                    content = content.replace(/<w:sectPr[\s\S]*?(\/|<\/w:sectPr)>/g, '');

                    // Добавляем ПУСТУЮ СТРОКУ перед контентом и сам контент
                    contentToAppend += EMPTY_LINE_XML + content;
                }
            } catch (err) {
                console.warn(`[WARN] Ошибка чтения файла ${files[i]}: ${err.message}`);
            }
        }

        // 3. Вклеиваем собранный текст в Master
        const finalXml = masterXml.slice(0, insertIndex) + contentToAppend + masterXml.slice(insertIndex);
        
        // Обновляем XML в памяти архива
        masterZip.updateFile('word/document.xml', Buffer.from(finalXml, 'utf-8'));
    } else {
        console.log(`>>> Комплект "${folderName}" состоит из 1 файла. Просто копируем.`);
    }

    // 4. Сохраняем результат в корень папки IN
    const outputFilePath = path.join(ROOT_DIR, `${folderName}.docx`);
    masterZip.writeZip(outputFilePath);
    
    console.log(`[OK] Сохранен: IN/${folderName}.docx`);
}

// Запуск
processBatches();
