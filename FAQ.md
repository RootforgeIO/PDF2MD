# PDF2MD FAQ

## Русский

**Нужно устанавливать Python?** Нет, он включён в готовые сборки.

**Чем portable отличается от установщика?** Portable запускается из распакованной папки и не регистрируется в Windows. Установщик создаёт ярлыки и запись удаления. Обе лёгкие сборки используют профиль Windows для настроек и отдельную папку для Marker.

**Где результат?** Рядом с документом: файл `.md` и, если есть изображения, папка `<имя>_images`. У пользователя должны быть права записи в эту папку.

**Можно выбрать другую папку результата?** В CLI: `python pdf2md.py document.pdf -o output/`. В окне результат сохраняется рядом с исходником. CLI перезаписывает существующий результат без вопроса, но запрещает путь к самому исходнику.

**Почему скан пустой?** Structured и PyMuPDF не выполняют OCR. Установите Marker и выберите его вручную, если «Авто» не распознало скан.

**Почему Marker медленный?** Он загружает модели и анализирует изображения. На CPU обработка может занимать около минуты на страницу. Вариант CUDA требует поддерживаемой видеокарты NVIDIA и драйвера.

**Почему таблица выглядит иначе?** Markdown ограничен; объединённые ячейки записываются как HTML. Таблицы без рамок сложнее распознавать. Проверьте другой движок и просмотрщик; цифры и расположение данных нужно сверить с исходником.

**Почему нет значений формул Excel?** Значения читаются из кэша книги. Откройте файл в Excel, пересчитайте, сохраните и повторите конвертацию. Экспортируется один выбранный лист.

**Как открыть старые Word и Excel?** Для `.doc` нужен Microsoft Word, для `.xls` — Microsoft Excel. Исходник открывается через временную копию. Можно заранее сохранить документ как `.docx` или `.xlsx`.

**Английский язык изменит числа и документ?** Нет. Меняется интерфейс. Числа Excel остаются в русском формате: `1 234,50`.

**Где настройки и журнал?** `%LOCALAPPDATA%\PDF2MD\settings.json` и `%LOCALAPPDATA%\PDF2MD\logs\pdf2md.log`. Удаление программы их сохраняет. Установленный Marker описан в отдельном `marker.json`.

**Работает без интернета?** Базовые конвертеры — да. Для установки Marker и первой загрузки моделей нужен интернет; после загрузки документы обрабатываются локально.

**Можно использовать в работе и изменять исходники?** Да, с соблюдением AGPL-3.0 и лицензий сторонних компонентов. Права на обрабатываемые документы остаются у их правообладателей. Полный текст условий — в `LICENSE.txt`.

**Что отправить при ошибке?** Версию программы и Windows, формат, шаги и обезличенный пример. Проверьте журнал: он может содержать пути и названия файлов/листов. Не отправляйте личные документы.

## English

**Do I need Python?** No. Ready-made builds include it.

**Portable or installer?** Portable runs from an extracted folder without registering an installation. The installer adds shortcuts and an uninstall entry. Both light builds use the Windows profile for settings and a separate folder for Marker.

**Where is the output?** Next to the source: a `.md` file and, when needed, `<name>_images`. You need write permission for that folder.

**Can I choose an output folder?** With the CLI: `python pdf2md.py document.pdf -o output/`. The GUI writes next to the source. The CLI replaces an existing output without asking, but rejects the source itself as output.

**Why is a scan empty?** Structured and PyMuPDF have no OCR. Install Marker and select it manually if Auto misses a scan.

**Why is Marker slow?** It loads models and analyses images. CPU conversion can take about a minute per page. CUDA requires a supported NVIDIA GPU and driver.

**Why does a table look different?** Markdown has limitations; merged cells use HTML. Borderless tables are harder to detect. Try another engine or viewer, and compare numbers and data positions with the source.

**Why are Excel formula values missing?** The converter reads cached values. Open, recalculate and save the workbook in Excel, then retry. One selected worksheet is exported.

**How do I open legacy Office files?** `.doc` needs Microsoft Word; `.xls` needs Microsoft Excel. Conversion uses a temporary copy. Alternatively, save as `.docx` or `.xlsx` first.

**Does switching to English change documents or numbers?** No. It changes the interface. Excel numbers retain Russian formatting, such as `1 234,50`.

**Where are settings and logs?** `%LOCALAPPDATA%\PDF2MD\settings.json` and `%LOCALAPPDATA%\PDF2MD\logs\pdf2md.log`. Uninstall keeps them. A separate `marker.json` describes the installed Marker engine.

**Does it work offline?** Basic converters do. Marker installation and initial model downloads need internet; conversion then runs locally.

**May I use it at work and modify the source?** Yes, subject to AGPL-3.0 and third-party licenses. Document rights remain with their respective owners. See `LICENSE.txt` for the full terms.

**What should a bug report include?** Application and Windows versions, format, steps and a sanitized sample. Review logs for paths and file/sheet names. Do not share private documents.
