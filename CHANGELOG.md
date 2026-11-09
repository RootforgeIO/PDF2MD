# Changelog

## 0.2.1 — 2026-11-09

### Русский

- Явное завершение конвертации: полоска исчезает, появляются «Открыть результат» и «Открыть папку».
- Имя результата в статусе, полный путь в подсказке и журнале; кнопки сохраняются при смене языка.
- Убрано повторное название с синей полосы. Внизу добавлена ссылка GitHub · RootforgeIO.
- Обновлены инструкции RU/EN и проверки интерфейса.

### English

- Clear completion state: the progress bar is replaced by “Open result” and “Open folder”.
- The status shows the output filename; the tooltip and log retain the full path. Result actions survive language changes.
- Removed the repeated title from the blue band and added a GitHub · RootforgeIO footer link.
- Updated RU/EN instructions and interface regression coverage.

## 0.2.0 — 2026-11-06

### Русский

- Русский и английский интерфейс; выбор языка через меню без перезапуска.
- Сохранение языка, движка, флажков и последней папки.
- «Авто» учитывает содержание всех страниц и выбирает Structured для текстовых PDF.
- Защита от записи в исходник, включая жёсткие ссылки; атомарная запись Markdown.
- Исправлен запуск без drag-and-drop; перетаскивание работает после смены языка.
- Прогресс Marker не откатывается между этапами; исправлено закрытие потока при отмене.
- Установщик автоматически создаёт ярлыки на рабочем столе и в меню «Пуск».
- Лёгкие сборки с необязательной установкой Marker CPU/CUDA.

### English

- Russian and English UI, with live language selection from the menu.
- Persistent language, engine, options and last folder.
- Auto examines all pages and prefers Structured for text PDFs.
- Source overwrite protection, including hard links; atomic Markdown writes.
- Startup without drag-and-drop is fixed; dropping files survives language changes.
- Marker progress no longer reverses between stages; cancellation closes the output stream.
- Installer automatically creates Desktop and Start menu shortcuts.
- Light builds with optional CPU/CUDA Marker installation.

## 0.1.0 — 2026-11-05

Windows installer; PDF, Word, Excel, CSV, TSV and TXT conversion; optional Marker installation; progress, cancellation and Windows XP Luna interface.
