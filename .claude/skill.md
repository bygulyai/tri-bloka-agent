---
name: tri-bloka-agent
description: "Пост-консультационный ИИ-агент: сырые заметки → CRM-отчёт + сообщение клиенту + PDF-презентация. Бизнес-агностичный, настраивается через config.json."
---

# Tri Bloka Agent — Skill

Этот скилл позволяет воспроизвести проект tri-bloka-agent на любом компьютере и адаптировать под любую компанию.

## Структура репозитория

```
tri-bloka-agent/
├── CLAUDE.md                          ← инструкции для ИИ-агента
├── config.example.json                ← шаблон конфигурации (скопировать в config.json)
├── README.md                          ← описание проекта
├── .gitignore                         ← исключения для git
├── .claude/skill.md                   ← этот файл
├── references/
│   ├── product-catalog.example.md     ← шаблон каталога продуктов
│   ├── client-routing.example.md      ← шаблон маршрутизации
│   ├── source-map.example.md          ← шаблон карты источников
│   ├── crm-message-SKILL.md           ← навык: CRM + сообщение
│   ├── tri-bloka-SKILL.md             ← навык: оркестратор
│   ├── presentation-SKILL.md          ← навык: презентация
│   ├── example-client-1.md            ← пример: три блока
│   ├── example-client-2.md            ← пример: CRM + сообщение
│   └── presentation_brief.example.json ← пример brief JSON
├── scripts/
│   ├── build_presentation.ps1         ← генерация презентации (PowerShell + PowerPoint COM)
│   ├── render_pdf_pages.py            ← рендер PDF-страниц в PNG (PyMuPDF)
│   └── validate_pdf.py                ← валидация PDF (PyMuPDF)
└── data/                              ← входные данные (заметки, CSV)
```

## Быстрый старт

1. Клонируйте репозиторий:
   ```bash
   git clone https://github.com/bygulyai/tri-bloka-agent.git
   cd tri-bloka-agent
   ```

2. Создайте конфигурацию:
   ```bash
   cp config.example.json config.json
   ```
   Заполните `config.json`: название компании, данные консультанта, пути, дефолтные цены.

3. Создайте справочники (скопируйте шаблоны и заполните):
   ```bash
   cd references/
   cp product-catalog.example.md product-catalog.md
   cp client-routing.example.md client-routing.md
   cp source-map.example.md source-map.md
   ```

4. Добавьте файлы компании:
   - `references/reference-presentation.pdf` — ваш шаблон презентации (слайды 5, 7-15)
   - `references/consultant-photo.png` — фото консультанта (слайд 6)

5. Установите зависимости Python:
   ```bash
   pip install PyMuPDF
   ```

6. Запустите Claude Code в папке проекта:
   ```bash
   claude
   ```

7. Отправьте сырые заметки с консультации. Агент выдаст три блока:
   - БЛОК 1: CRM-отчёт
   - БЛОК 2: сообщение клиенту
   - БЛОК 3: PDF-презентация

## Как внести изменения

### Изменить данные компании
Отредактируйте `config.json` — все настройки в одном месте:
- Название и сайт компании
- Имя и фото консультанта
- Компетенции консультанта
- Дефолтные цены (скидка, рассрочка)
- Заголовки слайдов

### Добавить продукты
Отредактируйте `references/product-catalog.md` — добавьте продукты по шаблону:
- Название, описание, URL, цены, тарифы

### Изменить логику маршрутизации
Отредактируйте `references/client-routing.md` — правила подбора продукта под запрос клиента.

### Изменить шаблон презентации
Замените `references/reference-presentation.pdf` на ваш шаблон. Статические слайды (5, 7-15) берутся из этого файла.

### Изменить тексты навыков
Все навыки — Markdown-файлы в `references/`:
- `crm-message-SKILL.md` — правила CRM-отчёта и сообщения
- `tri-bloka-SKILL.md` — оркестратор трёх блоков
- `presentation-SKILL.md` — правила генерации презентации

## Как проверить ошибки

1. Валидация PDF:
   ```bash
   python scripts/validate_pdf.py reports/ClientName.pdf 16
   ```
   Проверяет: файл существует, размер > 10KB, количество страниц = 16.

2. Рендер страниц для отладки:
   ```bash
   python scripts/render_pdf_pages.py references/reference-presentation.pdf output_dir 1 2 3
   ```
   Создаёт PNG-скриншоты указанных страниц.

3. Проверка конфигурации:
   - `config.json` должен быть валидным JSON
   - Все пути в `config.json` должны существовать (фото, референс PDF)
   - Справочники должны быть заполнены (не `.example.md`, а реальные файлы)

## Подготовка к публикации

1. Убедитесь, что `config.json` НЕ коммитится (в `.gitignore`)
2. Убедитесь, что `references/consultant-photo.png` и `references/reference-presentation.pdf` НЕ коммитятся (в `.gitignore`)
3. Убедитесь, что `reports/` НЕ коммитится (клиентские данные)
4. Шаблоны `.example.md` и `.example.json` ДОЛЖНЫ быть в репозитории
5. Обновите `README.md` при изменении структуры проекта

## Требования

- Windows 10/11 с PowerShell 5.1+
- Microsoft PowerPoint (для COM-автоматизации)
- Python 3.10+ с PyMuPDF
- Claude Code CLI
- Microsoft Edge (для скриншотов страниц продуктов)