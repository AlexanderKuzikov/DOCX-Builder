const fs = require('fs');
const path = require('path');

// --- ЧТЕНИЕ НАСТРОЕК ---
const SETTINGS_FILE = path.join(__dirname, 'settings.json');
const CONFIG_FILE = path.join(__dirname, 'renamer_config.json');
let ROOT_DIR = path.join(__dirname, 'IN'); // Дефолтное значение

if (fs.existsSync(SETTINGS_FILE)) {
    try {
        const settings = JSON.parse(fs.readFileSync(SETTINGS_FILE, 'utf8'));
        if (settings.inDir) {
            ROOT_DIR = settings.inDir;
            ROOT_DIR = ROOT_DIR.replace(/^"|"$/g, '');
        }
    } catch (e) {
        console.error('Ошибка чтения settings.json, используется дефолтный путь.', e);
    }
}

function startRenaming() {
    console.log(`Working directory: ${ROOT_DIR}`);

    // 1. Проверяем наличие папки
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
            // Можно использовать дефолтный конфиг, если файла нет
            config = {"1": "Шапка", "2": "Заголовок"}; 
        }
    } catch (e) {
        console.error('Ошибка чтения renamer_config.json:', e.message);
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

    // 4. Обрабатываем каждую папку
    folders.forEach(folderName => {
        const folderPath = path.join(ROOT_DIR, folderName);
        processFolder(folderPath, folderName, config);
    });
}

function processFolder(folderPath, folderName, config) {
    let contextText = '';
    const firstBracket = folderName.indexOf('(');
    const lastBracket = folderName.lastIndexOf(')');

    if (firstBracket !== -1 && lastBracket > firstBracket) {
        contextText = folderName.substring(firstBracket + 1, lastBracket);
    }

    const files = fs.readdirSync(folderPath).filter(f => f.endsWith('.docx') && !f.startsWith('~'));

    files.forEach(file => {
        const match = file.match(/^([\d\.]+)_/);
        
        if (match) {
            const prefix = match[1]; 
            const configKey = prefix; // Ищем точное совпадение ключа ("1", "1.5" и т.д.)

            if (config[configKey]) {
                const mapName = config[configKey];
                
                let newNameBase = `${prefix}_${mapName}`;
                
                if (contextText) {
                    newNameBase += ` (${contextText})`;
                }

                const newFileName = `${newNameBase}.docx`;

                if (file !== newFileName) {
                    const oldPath = path.join(folderPath, file);
                    const newPath = path.join(folderPath, newFileName);
                    
                    try {
                        // Проверка на случай, если файл с новым именем уже есть
                        if (fs.existsSync(newPath)) {
                            console.log(`[SKIP] Файл ${newFileName} уже существует.`);
                        } else {
                            fs.renameSync(oldPath, newPath);
                            console.log(`[RENAME] ${folderName}\n   └─ ${file} -> ${newFileName}`);
                        }
                    } catch (err) {
                        console.error(`[ERROR] Не удалось переименовать ${file}: ${err.message}`);
                    }
                }
            }
        }
    });
}

startRenaming();
