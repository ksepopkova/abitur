# Vuzline Webhook Service

Отдельный сервис, принимающий HTTP-уведомления от ЮКассы о статусе платежа
и отправляющий письмо с результатами подбора вузов сразу после успешной оплаты —
независимо от того, вернулся ли пользователь на сайт.

## Переменные окружения (задаются в настройках Render)

| Переменная | Описание |
|---|---|
| `YUKASSA_SHOP_ID` | Тот же shopId, что используется в Streamlit-приложении |
| `YUKASSA_SECRET_KEY` | Тот же secret key, что используется в Streamlit-приложении |
| `SHEETS_ID` | ID Google Sheets таблицы (та же, что в Streamlit-приложении) |
| `GCP_SERVICE_ACCOUNT_JSON` | Полное содержимое JSON-ключа сервисного аккаунта Google, **одной строкой** |
| `EMAIL_FROM` | Адрес отправителя (Яндекс почта) |
| `EMAIL_PASSWORD` | Пароль приложения для SMTP |
| `TILDA_WEBHOOK_TOKEN` | Секретная строка для адреса вебхука методички (`/tilda/metodichka?token=...`). Без неё эндпоинт отвечает 403 |
| `METODICHKA_URL` | Ссылка на методичку в письме (по умолчанию `https://docs.google.com/document/d/1cLu6-d32t4fGg0OjdN-ncKRggb6LlvsetCsX5GlehYQ/edit?usp=sharing`) |

## Деплой на Render.com

1. Зарегистрируйтесь на render.com через GitHub
2. New + → Web Service
3. Подключите репозиторий `abitur`
4. Root Directory: `webhook`
5. Runtime: Python 3
6. Build Command: `pip install -r requirements.txt`
7. Start Command: `uvicorn main:app --host 0.0.0.0 --port $PORT`
8. Добавьте переменные окружения из таблицы выше во вкладке Environment
9. Deploy

После деплоя Render выдаст URL вида `https://vuzline-webhook.onrender.com`.

## Настройка в личном кабинете ЮКассы

Настройки → Уведомления (HTTP-уведомления) → URL:
`https://vuzline-webhook.onrender.com/webhook`

Включить уведомление о событии `payment.succeeded`.

## Локальный запуск для теста

```bash
pip install -r requirements.txt
export YUKASSA_SHOP_ID=...
export YUKASSA_SECRET_KEY=...
export SHEETS_ID=...
export GCP_SERVICE_ACCOUNT_JSON='{...}'
export EMAIL_FROM=...
export EMAIL_PASSWORD=...
uvicorn main:app --reload
```

Проверка: `curl http://localhost:8000/health`

## Методичка: письмо после оплаты на сайте (Тильда)

Эндпоинт `POST /tilda/metodichka?token=<TILDA_WEBHOOK_TOKEN>` принимает данные заказа из корзины
на странице методички vuzline.ru и отправляет покупателю письмо со ссылкой `METODICHKA_URL`.

Настройка в Тильде:
1. Настройки сайта → Формы → Webhook: URL `https://vuzline-webhook.onrender.com/tilda/metodichka?token=<TILDA_WEBHOOK_TOKEN>`.
2. В корзине на странице методички отметить этот Webhook в «Приём данных из формы».
3. В настройках ЮKassa в Тильде включить «Отправлять данные в сервисы приёма данных только после оплаты».

Письмо уходит, только если в заказе есть слово «методичка». Повторный запрос с тем же tranid не дублирует письмо (до перезапуска сервиса).
