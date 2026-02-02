const fs = require('fs');
const path = require('path');

// НАСТРОЙКИ
const ROOT_DIR = path.join(__dirname, 'IN');
const CONFIG_FILE = path.join(__dirname, 'renamer_config.json');

function startRenaming() {
    // 1. Проверяем наличие папки IN
    if (!fs.existsSync(ROOT_DIR)) {
        console.error(`Папка ${ROOT_DIR} не найдена!`);
        return;
    }

    // 2. Загружаем конфиг с правилами имен
    let config = {};
    try {
        if (fs.existsSync(CONFIG_FILE)) {
            const rawConfig = fs.readFileSync(CONFIG_FILE, 'utf8');
            config = JSON.parse(rawConfig);
        } else {
            console.error(`Файл конфигурации ${CONFIG_FILE} не найден.`);
            return;
        }
    } catch (e) {
        console.error('Ошибка чтения JSON конфига:', e.message);
        return;
    }

    // 3. Получаем список папок
    const folders = fs.readdirSync(ROOT_DIR).filter(file => 
        fs.statSync(path.join(ROOT_DIR, file)).isDirectory()
    );

    if (folders.length === 0) {
        console.log('Нет папок для обработки.');
        return;
    }

    console.log(`Найдено папок: ${folders.length}`);

    // 4. Обрабатываем каждую папку
    folders.forEach(folderName => {
        const folderPath = path.join(ROOT_DIR, folderName);
        processFolder(folderPath, folderName, config);
    });
}

function processFolder(folderPath, folderName, config) {
    // Извлекаем текст из скобок папки
    // Логика: от ПЕРВОЙ '(' до ПОСЛЕДНЕЙ ')' — это поддерживает вложенные скобки
    // Пример: "Папка (Контекст (Уточнение))" -> "Контекст (Уточнение)"
    let contextText = '';
    const firstBracket = folderName.indexOf('(');
    const lastBracket = folderName.lastIndexOf(')');

    if (firstBracket !== -1 && lastBracket > firstBracket) {
        contextText = folderName.substring(firstBracket + 1, lastBracket);
    }

    // Читаем файлы
    const files = fs.readdirSync(folderPath).filter(f => f.endsWith('.docx'));

    files.forEach(file => {
        // Проверяем формат имени: "ЦИФРА_..."
        // Используем match, чтобы поймать префикс (цифры + возможно точки для 1.5)
        const match = file.match(/^([\d\.]+)_/);
        
        if (match) {
            const prefix = match[1]; // Например "1" или "1.5"
            
            // Ищем правило в конфиге для этого префикса
            // (или для целой части числа, если нужно, но пока ищем точное совпадение ключа)
            // Если у вас 1.5, а в конфиге только "1", то имя не поменяется, если не добавить логику округления.
            // Сейчас логика: ищем точное совпадение ключа. Если ключ "1", он сработает для "1_".
            
            // Если нужно, чтобы для "1.5_" бралось правило от "1", раскомментируйте строку ниже:
            // const configKey = prefix.split('.')[0]; 
            const configKey = prefix; 

            if (config[configKey]) {
                const mapName = config[configKey]; // "Шапка"
                
                // Формируем новое имя
                // База: "1_Шапка"
                let newNameBase = `${prefix}_${mapName}`;
                
                // Добавляем контекст в скобках, если он был в имени папки
                if (contextText) {
                    newNameBase += ` (${contextText})`;
                }

                const newFileName = `${newNameBase}.docx`;

                // Переименовываем только если имя отличается
                if (file !== newFileName) {
                    const oldPath = path.join(folderPath, file);
                    const newPath = path.join(folderPath, newFileName);
                    
                    try {
                        fs.renameSync(oldPath, newPath);
                        console.log(`[RENAME] ${folderName}\n   └─ ${file} -> ${newFileName}`);
                    } catch (err) {
                        console.error(`[ERROR] Не удалось переименовать ${file}: ${err.message}`);
                    }
                }
            }
        }
    });
}

startRenaming();
