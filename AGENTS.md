# DOCX-Builder — Instructions for AI Agents

## Commands
- start: `npm start`
- rename: `npm run rename`
- test: `npm test`

## Conventions
- Node.js CommonJS
- adm-zip для работы с DOCX (XML fragments)
- Express для Web UI
- Юридические документы: папки, переименование, merge

## Structure
- `builder.js` — сборка DOCX из XML-фрагментов
- `renamer.js` — переименование файлов
- `public/` — Web UI

## Do NOT touch
- `node_modules/`

## Documentation rules
- После работы — обнови docs/CONTEXT.md
- НЕ создавай новых файлов документации без разрешения
