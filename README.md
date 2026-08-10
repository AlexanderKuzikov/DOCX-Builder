<p align="center">
  <a href="https://nodejs.org/"><img alt="Node" src="https://img.shields.io/badge/Node-18+-339933?logo=node.js&logoColor=white"></a>
  <a href="LICENSE"><img alt="License" src="https://img.shields.io/badge/License-Apache_2.0-blue.svg"></a>
</p>

<h1 align="center">DOCX-Builder</h1>
<p align="center">Пакетная сборка DOCX из XML-фрагментов с веб-интерфейсом</p>

---

Автоматизация юридических документов: сборка DOCX из XML-фрагментов через adm-zip, создание структуры папок, переименование файлов, one-click merge. Express Web UI для управления.

- **Сборка DOCX** — из XML-фрагментов (adm-zip)
- **Web UI** — Express, управление через браузер
- **Переименование** — массовое по шаблону
- **Структура папок** — автоматическое создание

## Быстрый старт

```bash
git clone https://github.com/AlexanderKuzikov/DOCX-Builder.git
cd DOCX-Builder
npm install
npm start              # builder.js
npm run rename         # renamer.js
```

## Документация

- [`docs/CONTEXT.md`](docs/CONTEXT.md) — состояние проекта
- [`docs/DECISIONS.md`](docs/DECISIONS.md) — архитектурные решения

## Статус

**Работает** — сборка DOCX, переименование, Web UI.

## Лицензия

[Apache-2.0](LICENSE) © Alexander Kuzikov
