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
