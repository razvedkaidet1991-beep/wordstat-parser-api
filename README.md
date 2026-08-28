# Парсер Яндекс Вордстат (Cloud API)

Сбор частотностей и популярных фраз через **Yandex Cloud Search API** (`topRequests`). Без браузера и Selenium.

## Возможности

1. **Частоты** — для каждого запроса из `input/queries.txt` считает выбранные типы:
   - базовую;
   - точную (`"фраза"`);
   - уточнённую (`"!слово !слово"`);
   - по умолчанию только **базовая** (1 API-вызов на фразу);
   - сохраняет Excel в `output/wordstat_report.xlsx`.
2. **Сбор фраз** — топ «Популярные» до 2000 фраз на seed, фильтр `input/stop_words.txt`, результат в `output/wordstat_words.txt`.

## Быстрый старт

### 1. API-ключ

Пошаговая инструкция: **[docs/API_KEY.md](docs/API_KEY.md)**.

Кратко: нужны `api_key` и `folder_id` в файле:

```text
data/credentials.txt
```

Шаблон: `data/credentials.example.txt`.

### 2. Установка

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate
# Linux/macOS
source .venv/bin/activate

pip install -r requirements.txt
```

### 3. Входные файлы

- `input/queries.txt` — запросы, по одному на строку
- `input/stop_words.txt` — минус-слова для режима 2 (можно пустой)
- `output/invalid_queries.txt` — невалидные seed-запросы (пустые после нормализации или `HTTP 400 Invalid query`)

### 4. Запуск

```bash
python wordstat_api.py
```

Меню:

1. частоты → `output/wordstat_report.xlsx`  
   затем мультивыбор типов: `b` / `e` / `p` (Enter = только базовая)
2. сбор фраз → `output/wordstat_words.txt`

Без меню:

```bash
python wordstat_api.py --mode 1
python wordstat_api.py --mode 1 --freq-types b
python wordstat_api.py --mode 1 --freq-types bep
python wordstat_api.py --mode 1 --freq-types base,exact
python wordstat_api.py --phrases
```

`--freq-types`: `b`/`e`/`p` или `base`/`exact`/`precise` (комбинации). Без флага — только базовая.
## Структура

```
wordstat api/
├── wordstat_api.py
├── requirements.txt
├── README.md
├── docs/
│   └── API_KEY.md          # как получить ключ
├── input/
│   ├── queries.txt
│   └── stop_words.txt
├── output/
│   ├── wordstat_report.xlsx
│   └── wordstat_words.txt
└── data/
    ├── credentials.example.txt
    └── credentials.txt     # не в git
```

## Стоп-слова

В режиме сбора фраз после выгрузки применяется фильтр по `input/stop_words.txt` (частичное вхождение без учёта регистра). Пустой или отсутствующий файл — фильтр не используется.

## Лимиты API и прогресс

- До **2000** фраз на один seed (`numPhrases`).
- Wordstat в Search API по умолчанию: **~100 запросов в час** (и до ~10 RPS). При превышении — HTTP **429**.
- После **10 ответов 429 подряд** скрипт **останавливается**.
- Успешно обработанный seed сразу пишется в `output/` и **удаляется** из `input/queries.txt`.
- Перед отправкой в API seed-нормализуется: убираются кавычки и `!`, а также `&`, `/`, `|`, `@`, `;`, `:` (операторы точной/уточнённой частоты навешиваются отдельно).
- Если после нормализации запрос пустой или API вернул `400 Invalid query`, seed уходит в `output/invalid_queries.txt` и удаляется из очереди.
- Повторный запуск продолжает с оставшихся строк в `queries.txt`; `wordstat_words.txt` / Excel **дополняются**.
- В консоли «Квоты» строка Wordstat может не отображаться; увеличение — через ТП.
- Подробности: [docs/API_KEY.md](docs/API_KEY.md).
- Между запросами пауза ~0.25 с.
