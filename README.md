# job-application-pack

Скилл для Claude и ChatGPT, который собирает пакет для отклика на вакансию: адаптирует ваше мастер-резюме под вакансию, отдаёт его файлом .docx с сохранением оформления и пишет сопроводительное письмо.

Скилл не выдумывает опыт, цифры и навыки. Если в вакансии есть требования, которых нет в вашем резюме, он сначала спросит вас, а уже потом соберёт пакет.

## Скачать

**[job-application-pack.zip](../../releases/latest/download/job-application-pack.zip)** — готовый архив скилла. Его загружают в Claude или ChatGPT.

## Как пользоваться

1. Скачайте архив и распакуйте его.
2. Положите своё мастер-резюме (.docx) в папку `master`. Название файла — только латиницей и без пробелов, например `master-resume.docx`. Мастер-резюме ещё нет — заполните шаблон `assets/resume-template.docx`.
3. Запакуйте папку `job-application-pack` обратно в zip.
4. Загрузите архив:
   - **Claude:** Settings → Capabilities → включите Code execution and file creation, затем Customize → Skills → «+» → Create skill → Upload a skill.
   - **ChatGPT** (Business, Enterprise, Edu): Plugins → Skills → Create → Upload from your computer.
5. В новом чате вставьте вакансию и напишите: «Собери пакет для отклика на эту вакансию».

Подробная инструкция по использованию скилла: [гугл-док]((https://docs.google.com/document/d/1waDTUb4Wi54Edl7Sq9cyLeFJF55KF-9VEuK_7GopMcQ/edit?usp=sharing))

## Что внутри

| Файл | Зачем |
|---|---|
| `SKILL.md` | Главная инструкция: описание скилла и порядок шагов |
| `master/` | Сюда кладётся мастер-резюме |
| `assets/resume-template.docx` | Шаблон мастер-резюме |
| `references/vacancy-analysis.md` | Разбор вакансии и вопросы по пробелам |
| `references/resume-adaptation.md` | Правила адаптации резюме |
| `references/cover-letter.md` | Структура сопроводительного и образцы |
| `references/tone-of-voice.md` | Как писать по-человечески, без нейросетевых оборотов |
| `references/master-resume-rules.md` | Проверка мастер-резюме |
| `scripts/docx_tool.py` | Правка .docx без потери оформления |

## Автор

Елизавета Гусенкова · [t.me/liza_and_digital](https://t.me/liza_and_digital)
