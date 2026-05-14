# FireProtect КП Telegram Bot — MVP

## Что делает бот

1. Принимает PDF спецификацию в Telegram.
2. Достаёт текст из PDF.
3. Ищет позиции по листу `Sinonime`.
4. Находит `Cod produs`.
5. Заполняет лист `Introducere_interna` в шаблоне `KP_Client_FireProtect.xlsx`.
6. Возвращает готовый Excel с КП.

## Что нужно подготовить

В эту же папку положи файл:

`KP_Client_FireProtect.xlsx`

Важно: внутри него должны быть листы:
- `KP_Client`
- `Introducere_interna`
- `Baza_preturi`
- `Sinonime`
- `Setari`

## Установка

Открой терминал в этой папке и выполни:

```bash
python -m venv .venv
```

Windows:

```bash
.venv\Scripts\activate
```

Mac/Linux:

```bash
source .venv/bin/activate
```

Потом:

```bash
pip install -r requirements.txt
```

## Telegram token

1. Открой Telegram.
2. Найди `@BotFather`.
3. Напиши `/newbot`.
4. Получи token.
5. Скопируй `.env.example` в `.env`.
6. Вставь token в строку `BOT_TOKEN=...`.

## Запуск

```bash
python bot.py
```

После запуска отправь боту PDF спецификацию.

## Важно

Это MVP. Он хорошо работает с PDF, где текст можно выделить.
Если PDF — скан/картинка, понадобится OCR. Это добавим отдельным этапом.

## Как обучать бота

Если бот пишет, что позиция не распознана:

1. Открой `KP_Client_FireProtect.xlsx`.
2. Покажи скрытый лист `Sinonime`.
3. Добавь новую строку:

`Название из PDF` → `Cod produs`

Пример:

`Труба стальная электросварная Ø89x3.5` → `TEV-BL-DN80`

4. Сохрани файл.
5. Перезапусти бота.
