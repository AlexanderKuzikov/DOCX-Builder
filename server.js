const express = require('express');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');

// --- КОНФИГУРАЦИЯ ---
const PORT = 5555; // <-- Твой новый порт

// Динамический импорт open
const openBrowser = async (target) => {
    try {
        const open = (await import('open')).default;
        await open(target);
    } catch (e) {
        console.error('Ошибка при открытии браузера/файла:', e);
    }
};

const app = express();
const SETTINGS_FILE = path.join(__dirname, 'settings.json');
const CASES_FILE = path.join(__dirname, 'cases.txt');

// Хелпер для получения текущей рабочей папки
function getInDir() {
    let dir = path.join(__dirname, 'IN'); // Дефолт
    if (fs.existsSync(SETTINGS_FILE)) {
        try {
            const settings = JSON.parse(fs.readFileSync(SETTINGS_FILE, 'utf8'));
            if (settings.inDir) {
                dir = settings.inDir.trim().replace(/^"|"$/g, '');
            }
        } catch (e) {}
    }
    // Если папки нет физически, создаем
    if (!fs.existsSync(dir)) {
        try {
            fs.mkdirSync(dir, { recursive: true });
        } catch (e) {
            console.error(`Не удалось создать рабочую папку ${dir}:`, e);
            dir = path.join(__dirname, 'IN');
            if (!fs.existsSync(dir)) fs.mkdirSync(dir);
        }
    }
    return dir;
}

app.use(express.json());
app.use(express.static('public'));

// --- API НАСТРОЕК ---

app.get('/api/settings', (req, res) => {
    res.json({ inDir: getInDir() });
});

app.post('/api/settings', (req, res) => {
    let { inDir } = req.body;
    if (!inDir) return res.status(400).json({ error: 'Путь не указан' });
    
    inDir = inDir.trim().replace(/^"|"$/g, '');

    if (!fs.existsSync(inDir)) {
        try {
            fs.mkdirSync(inDir, { recursive: true });
        } catch (e) {
            return res.status(500).json({ error: `Не удалось создать папку: ${e.message}` });
        }
    }

    try {
        fs.writeFileSync(SETTINGS_FILE, JSON.stringify({ inDir }, null, 2));
        res.json({ success: true, inDir });
    } catch (e) {
        res.status(500).json({ error: 'Ошибка записи настроек' });
    }
});

// --- ОСНОВНОЕ API ---

app.get('/api/folders', (req, res) => {
    const IN_DIR = getInDir();
    if (!fs.existsSync(IN_DIR)) return res.json([]);

    let folders = [];
    try {
        folders = fs.readdirSync(IN_DIR).filter(file => 
            fs.statSync(path.join(IN_DIR, file)).isDirectory()
        );
    } catch (e) {
        return res.json([]);
    }

    const data = folders.map(name => {
        const folderPath = path.join(IN_DIR, name);
        let files = [];
        try { files = fs.readdirSync(folderPath); } catch (e) {}
        
        const resultFilePath = path.join(IN_DIR, `${name}.docx`);
        const resultExists = fs.existsSync(resultFilePath);

        const docxParts = files.filter(f => f.endsWith('.docx') && !f.startsWith('~'));
        const isEmpty = docxParts.length === 0;
        const isRenamed = docxParts.some(f => /^\d+_/.test(f));

        return {
            name,
            isEmpty,
            isRenamed,
            isBuilt: resultExists
        };
    });
    res.json(data);
});

app.get('/api/cases', (req, res) => {
    if (!fs.existsSync(CASES_FILE)) {
        fs.writeFileSync(CASES_FILE, "Аренда ТС\nБанкротство", 'utf8');
    }
    const fileContent = fs.readFileSync(CASES_FILE, 'utf8');
    const cases = fileContent.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
    res.json(cases);
});

app.post('/api/create', (req, res) => {
    const IN_DIR = getInDir();
    const { docType, selectedCases } = req.body;
    if (!docType || !selectedCases || !selectedCases.length) return res.status(400).json({ error: 'Нет данных' });

    const created = [];
    selectedCases.forEach(caseName => {
        const folderName = `${docType} (${caseName})`;
        const folderPath = path.join(IN_DIR, folderName);
        
        if (!fs.existsSync(folderPath)) {
            try {
                fs.mkdirSync(folderPath);
                created.push(folderName);
            } catch (e) {}
        }
    });
    res.json({ message: `Создано папок: ${created.length}`, created });
});

app.post('/api/open', async (req, res) => {
    const IN_DIR = getInDir();
    const { name, isFile } = req.body;
    
    let targetPath;
    if (isFile) {
        targetPath = path.join(IN_DIR, `${name}.docx`);
    } else {
        targetPath = path.join(IN_DIR, name);
    }

    if (fs.existsSync(targetPath)) {
        try {
            await openBrowser(targetPath);
            res.json({ success: true });
        } catch (e) {
            res.status(500).json({ error: 'Ошибка открытия' });
        }
    } else {
        res.status(404).json({ error: 'Объект не найден' });
    }
});

app.post('/api/run/renamer', (req, res) => {
    exec('node renamer.js', (error, stdout, stderr) => {
         if (error) return res.status(500).json({ error: stderr });
         res.json({ message: stdout });
    });
});

app.post('/api/run/builder', (req, res) => {
    const { folderName } = req.body;
    const command = folderName ? `node builder.js "${folderName}"` : 'node builder.js';
    console.log(`Run: ${command}`);
    exec(command, (error, stdout, stderr) => {
        if (error) return res.status(500).json({ error: stderr || error.message });
        res.json({ message: stdout });
    });
});

// SHUTDOWN
app.post('/api/shutdown', (req, res) => {
    res.json({ message: 'Server stopping...' });
    console.log('Shutdown requested.');
    setTimeout(() => process.exit(0), 500);
});

app.listen(PORT, async () => {
    const url = `http://localhost:${PORT}`;
    console.log(`Server running at ${url}`);
    try { await openBrowser(url); } catch (e) {}
});
