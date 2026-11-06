# Обработка документов и данных / Document handling

## Русский

PDF2MD обрабатывает документы локально. Обычная конвертация не отправляет содержимое документов в облачные сервисы. Приложение не требует аккаунта или API-ключа.

Для установки Marker загружаются пакеты и модели с серверов поставщиков: PyPI, PyTorch и Datalab; отдельные зависимости могут загружать модельные файлы из своих источников. Эти сетевые запросы нужны для установки компонентов. Базовые конвертеры работают без загрузки моделей.

Результаты и изображения сохраняются рядом с исходным документом, если другой путь не задан в CLI. Word/Excel старых форматов используют временные копии; для XLS может сохраняться локальный кэш конвертации во временной папке Windows.

Настройки содержат выбранный язык, движок, флажки и последнюю папку. Журнал может содержать пути, названия файлов и листов, а также диагностику конвертеров. Перед отправкой отчёта проверьте журнал и удалите личные сведения. Не прикладывайте конфиденциальные документы к публичным обращениям.

Удаление приложения сохраняет документы, настройки и отдельно загруженный Marker. Файлы в папках пользователя управляются самим пользователем.

## English

PDF2MD processes documents locally. Normal conversion does not upload document contents to cloud services. No account or API key is required.

Installing Marker downloads packages and models from suppliers including PyPI, PyTorch and Datalab; individual dependencies may obtain model files from their own sources. These requests install components. Basic converters do not need model downloads.

Output and images are saved next to the source unless the CLI specifies another destination. Legacy Word/Excel conversion uses temporary copies; XLS conversion may retain a local cache in the Windows temporary folder.

Preferences store language, engine, options and the last folder. Logs may contain paths, file and worksheet names, and converter diagnostics. Review and redact logs before sharing. Do not attach confidential documents to public reports.

Uninstalling keeps documents, settings and separately downloaded Marker files. Files in user folders remain under the user's control.
