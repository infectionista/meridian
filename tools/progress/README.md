# Science Map progress backend

`Code.gs` — Google Apps Script web app, привязанный к Google-таблице
на аккаунте meridianlearning.world. Таблица = учительский дашборд:
Students (имя+код), Events (лог), State (матрица достижений, можно
править вручную).

API: GET ?name=&code= → {keys}; POST text/plain JSON {name,code,type,id};
везде работает &callback= (JSONP). Типы: trainer | trail | step | open.

Деплой: Deploy → Web app → Execute as Me / Anyone. Обновление кода:
Manage deployments → Edit → New version (URL сохраняется).
Client config: PROGRESS_URL в goals-map (tools/goals-map/template.html)
и в тренажёрах.
