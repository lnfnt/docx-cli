# docx-cli

**Русский** · [<img src="assets/flags/gb.svg" height="16" align="top" alt="EN"> English](README.md)

![license](https://img.shields.io/badge/license-Apache--2.0-blue) ![python](https://img.shields.io/badge/python-3.11%2B-green) ![release](https://img.shields.io/badge/release-v0.1.0-informational) ![MCP](https://img.shields.io/badge/MCP-8%20tools-brightgreen)

<p align="center"><img src="assets/poster.jpg" width="820" alt="docx-cli poster"></p>

*Основной README — английский ([README.md](README.md)); это русская версия.*

Детерминированный генератор русскоязычных юридических документов `.docx`. Сборка из
машиночитаемого JSON-спека (или Markdown для процессуальных актов), затем повторное открытие
записанного файла и обратная валидация вёрстки: поля, шрифт, интервалы, отступы, нумерация,
типографика, «ё». Расхождение — fail-closed: частичный файл удаляется, CLI возвращает
ошибочный JSON-конверт. Встроен сканер скрытых инъекций входящего текста и MCP-режим.

## Установка

```bash
pip install docx-cli
docx-cli doctor
```

На PyPI пакета нет. Checkout: `pip install -e .` или `pipx install .` (Python >= 3.11;
python-docx, pydantic, lxml). Сетевых вызовов нет.

## Быстрый старт

```bash
docx-cli build examples/spec_example.json output.docx
docx-cli validate output.docx
docx-cli doctor
```

Единый JSON-конверт; коды выхода: 0 ок, 1 бизнес-отказ, 2 usage/структурная ошибка.

## Стандарт оформления

A4; поля Л=3.0/П=1.5/В=2.0/Н=2.0 см; Times New Roman 12 pt (заголовки 14); интервал 1.5;
красная строка 1.25 см. Профиль `court`: 14 pt везде, интервал 1.0, шапка
`left_indent = 7.0 см`, висячие списки. Скелеты `vs_sud` и `soyu_rayon` — геометрия и
плейсхолдеры `{...}` без норм; шаблоны в `templates/`. Полный контракт — `SPEC.md`.

## MCP-режим

`docx-cli mcp` — JSON-RPC 2.0, восемь инструментов. Write-инструменты требуют `confirm:
true`, иначе возвращается `needs_confirmation` без записей.

## Конвертер md

Вложенные списки 1-4 (глубина >4 маркируется `list_depth_gt4`), GFM-таблицы с выравниваниями
и inline-форматированием в ячейках; блочные конструкции в ячейках помечаются `block_in_cell`.

## Ограничения

- Не юридический совет, факты дела не проверяются.
- Только генерация из спека/Markdown; чужие legacy .docx не редактирует.
- Профильные скелеты содержат вёрстку, нормы — в спеке юриста.

## Лицензия

Apache-2.0 (`LICENSE`). Проверка: `sha256sum -c SHA256SUMS`, `gpg --verify *.asc`.

Автор и правообладатель: [lnfnt](mailto:lnfnt@proton.me).
