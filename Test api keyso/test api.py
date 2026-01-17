"""
Скрипт проверяет, что API Keys.so работает с указанным API‑ключом,
и возвращает базовую сводку по домену.

Что делает:
- Создаёт сессию requests и добавляет заголовок X-Keyso-TOKEN с API‑ключом.
- Отправляет запрос к эндпоинту /report/simple/domain_dashboard
  для домена kotofoto.ru и базы msk.
- Печатает в консоль HTTP‑код ответа и JSON‑данные, полученные от API.

Как подготовить:
1. Установить зависимость:
   pip install requests
2. В переменную API_KEY подставить свой рабочий ключ Keys.so.
3. При необходимости изменить домен и базу в параметрах:
   params={"base": "msk", "domain": "kotofoto.ru"}

Как пользоваться:
1. Запустить скрипт из консоли из папки со скриптом:
   py test-api.py
2. Убедиться, что статус 200 и в консоли появился JSON —
   значит, ключ и доступ к API работают корректно.
"""

import requests

API_KEY = "69424a9d42ab34.64613057787b24c68dcaacc0112242bc4bfe2090"
BASE_URL = "https://api.keys.so"

session = requests.Session()
session.headers.update({
    "X-Keyso-TOKEN": API_KEY,
    "Accept": "application/json",
})

resp = session.get(
    f"{BASE_URL}/report/simple/domain_dashboard",
    params={"base": "msk", "domain": "kotofoto.ru"},
)

print(resp.status_code)
print(resp.json())
