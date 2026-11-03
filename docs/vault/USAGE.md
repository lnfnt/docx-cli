# Использование docx-cli

## Установка

```bash
pip install -e .
```

После установки доступна команда `docx-cli`.

## Формат spec.json

Spec — это единый JSON-контракт документа. Пример см. в `examples/spec_example.json`.

Основные секции:

- `profile` — `gost` (по умолчанию), `compact` или `court`.
- `output` — имя файла и автор.
- `header` — реквизитные строки. В профиле `court` шапка рендерится как блок с `left_indent = 7.0 см`, выравнивание строк LEFT.
- `title` — основной заголовок и подзаголовок.
- `body` — массив блоков.
- `checks` — флаги встроенных проверок.

## Типы блоков

### paragraph

```json
{"type": "paragraph", "text": "Текст с **жирным** и *курсивом*.", "inline_markdown": true}
```

### numbered_list

```json
{"type": "numbered_list", "items": ["первый пункт", "второй пункт"]}
```

Нумерация всегда 1, 2, 3… Валидатор проверяет отсутствие пропусков и дубликатов.

### table

```json
{
  "type": "table",
  "headers": ["Статья", "Сумма"],
  "rows": [["309 ГК РФ", "100 000"]],
  "widths_cm": [8.0, 5.0]
}
```

### pleading_marker

```json
{"type": "pleading_marker", "text": "ПРОШУ СУД:"}
```

### signature_block

```json
{"type": "signature_block", "left": "Истец", "right": "___________ /Иванов И.И./"}
```

По умолчанию в профиле `court` подпись рендерится текстом (без таблицы); в `gost` — таблицей 1×2 с прозрачными границами. Формат можно задать явно: `"format": "text"` или `"format": "table"`.

### attachments_list

```json
{"type": "attachments_list", "items": ["1. Доверенность на 1 л."]}
```

## Плейсхолдеры

Разрешены плейсхолдеры вида `[УКАЗАТЬ ...]`. Они не прерывают сборку, но валидатор помечает их предупреждением.

## Проверки

- `no_yo` — запрет буквы `ё` вне реквизитов.
- `no_en_dash` — запрет en-dash `–`.
- `require_norms_refs` — требовать ссылки на нормы права.

## Пример spec с профилем `court`

```json
{
  "profile": "court",
  "output": {"file_name": "Отзыв.docx", "author": "Адвокат Т.Т."},
  "header": {
    "lines": [
      "В Арбитражный суд города Москвы",
      "",
      "Истец: ООО \"Пример\", ИНН 7701234567",
      "Адрес: 123456, г. Москва, ул. Примерная, д. 1",
      "",
      "Ответчик: ООО \"Ответ\", ИНН 7707654321"
    ]
  },
  "title": {"main": "ОТЗЫВ ИСТЦА", "subtitle": "на возражения ответчика"},
  "body": [
    {"type": "paragraph", "text": "Прошу суд удовлетворить требования истца.", "inline_markdown": false},
    {"type": "pleading_marker", "text": "ПРОШУ СУД:"},
    {
      "type": "numbered_list",
      "items": [
        "Взыскать с ответчика задолженность в размере 100 000 руб.",
        "Взыскать судебные расходы."
      ]
    },
    {"type": "signature_block", "left": "Представитель истца", "right": "___________ /Иванов И.И./"},
    {"type": "attachments_list", "items": ["1. Доверенность на 1 л."]}
  ]
}
```

## CLI

```bash
docx-cli build spec.json output.docx
docx-cli validate output.docx [--json] [--strict]
docx-cli doctor
```
