# Как получить API-ключ Яндекс Вордстат (Cloud Search API)

Этот проект использует **Yandex Cloud Search API v2** (`topRequests`), а не старый OAuth Wordstat и не браузер.

Нужны два значения:

| Параметр | Пример | Куда |
|----------|--------|------|
| `api_key` | `AQVN...` | секретный ключ |
| `folder_id` | `b1...` | ID каталога в Cloud |

Сохраните их в файл `data/credentials.txt` (шаблон — `data/credentials.example.txt`).

---

## 1. Аккаунт и биллинг

1. Зарегистрируйтесь в [Yandex Cloud Console](https://console.yandex.cloud/).
2. Привяжите платёжный аккаунт / карту (биллинг часто обязателен даже для тестовых вызовов).

## 2. Создайте каталог (folder)

1. В консоли Cloud откройте своё облако.
2. Создайте **каталог** (не «ресурс»).
3. Откройте каталог и скопируйте ID из URL:

```text
https://console.yandex.cloud/folders/b1xxxxxxxxxxxxxxxxxxxx/dashboard
                                      ↑ это folder_id
```

Или на странице каталога → **Обзор** → **Идентификатор**.

## 3. Сервисный аккаунт и роль

1. Каталог → **IAM** → **Сервисные аккаунты** → создать.
2. Назначьте роль на **каталог**: `search-api.webSearch.user`.

## 4. API-ключ (область действия)

1. Откройте сервисный аккаунт.
2. Блок **API-ключи** → **Создать API-ключ**.
3. В поле **Область действия (Scope)** выберите: `yc.search-api.execute`.
4. Сохраните **секретный ключ** (`AQVN...`) сразу — его показывают один раз.

> Это делается в [console.yandex.cloud](https://console.yandex.cloud/), не обязательно в AI Studio.

## 5. Файл credentials

Создайте `data/credentials.txt`:

```text
api_key=AQVN_ВАШ_СЕКРЕТНЫЙ_КЛЮЧ
folder_id=b1_ВАШ_FOLDER_ID
```

Не коммитьте этот файл в git.

## 6. Проверка через PowerShell

```powershell
chcp 65001 | Out-Null
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$apiKey   = "AQVN...."
$folderId = "b1...."
$uri = "https://searchapi.api.cloud.yandex.net/v2/wordstat/topRequests"

$body = @{ phrase = "seo оптимизация"; numPhrases = 5; folderId = $folderId } | ConvertTo-Json
$bytes = [System.Text.Encoding]::UTF8.GetBytes($body)

Invoke-RestMethod -Uri $uri -Method POST `
  -Headers @{ Authorization = "Api-Key $apiKey" } `
  -ContentType "application/json; charset=utf-8" `
  -Body $bytes
```

Успех: в ответе есть `totalCount` и `results`.

| Код | Обычно значит |
|-----|----------------|
| 401 | неверный `api_key` |
| 403 | нет роли / scope / биллинга |
| 400 | неверный или пустой `folder_id` |
| 429 | превышен лимит (часто **100 запросов Wordstat в час**) |

## 7. Лимиты Wordstat (Search API)

По ответу поддержки Яндекса для статистики Wordstat в Cloud Search API по умолчанию:

| Ограничение | Значение |
|-------------|----------|
| Запросов в час (Wordstat / статистика) | **100** |
| Запросов в секунду | обычно до **10** |

Один вызов `topRequests` = **1** единица квоты (независимо от `numPhrases`).

В разделе консоли **«Квоты»** отдельной строки Search API / Wordstat может **не быть** — лимит всё равно действует: при исчерпании API отвечает **HTTP 429**.

### Как увеличить квоту

1. Откройте поддержку в [консоли Yandex Cloud](https://console.yandex.cloud/).
2. Создайте обращение с просьбой увеличить квоту Wordstat.
3. Укажите облако, каталог, платёжный аккаунт и желаемое значение (например 1000 или 5000 запросов в час).

#### Пример обращения в ТП

```text
Здравствуйте!

Прошу увеличить квоту для сервиса Yandex Search API (Wordstat) в моём каталоге.

Данные организации/аккаунта:

Облако: <ID_ОБЛАКА>
Каталог (folderId): <ID_КАТАЛОГА>
Платёжный аккаунт: <ID_ПЛАТЁЖНОГО_АККАУНТА>

Авторизация: API-ключ сервисного аккаунта, scope yc.search-api.execute, роль search-api.webSearch.user

Запрашиваемая квота:

Название: Search API — количество запросов в час на получение статистики (Wordstat)
Текущее значение: 100 запросов в час
Желаемое значение: 5000 запросов в час

Обоснование:
Мы массово собираем популярные фразы через метод POST /v2/wordstat/topRequests по списку seed-запросов. Текущей квоты в 100 запросов в час недостаточно: при паузе ~0,25 с между запросами (~4 RPS, что укладывается в лимит 10 RPS) после ~100–200 успешных вызовов API начинает возвращать HTTP 429 Too Many Requests.

Прошу рассмотреть возможность увеличения квоты до указанного значения.
```

ID облака, каталога и платёжного аккаунта возьмите из консоли Cloud (URL и карточки ресурсов). **Не прикладывайте** секретный ключ `AQVN...` — достаточно описать способ авторизации.

Где найти ID:
- **Каталог** — из URL: `.../folders/b1...`
- **Облако** — в настройках облака / переключателе облаков (ID вида `b1...`)
- **Платёжный аккаунт** — Биллинг → платёжный аккаунт (ID вида `dn...`)

## Полезные ссылки

- [Консоль Yandex Cloud](https://console.yandex.cloud/)
- [Справка Wordstat API](https://yandex.ru/support2/wordstat/ru/content/api-wordstat)
- [Управление API-ключами](https://yandex.cloud/ru/docs/iam/operations/authentication/manage-api-keys)
- [Просмотр квот](https://yandex.cloud/ru/docs/quota-manager/operations/read-quotas)
