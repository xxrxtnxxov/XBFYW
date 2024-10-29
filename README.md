# Экспорт внешних ссылок и популярных запросов из Яндекс.Вебмастера

Этот скрипт предназначен для использования на страницах Яндекс Вебмастера и позволяет экспортировать внешние ссылки и популярные поисковые запросы вашего сайта в формате XLSX (лимит строк равен 1 000 000). Скрипт автоматически интегрируется в интерфейс Яндекс Вебмастера, добавляя кнопки для скачивания данных на страницах «Внешние ссылки» и «Управление группами».

## Функции

- **Экспорт внешних ссылок:** Скачивание данных о внешних ссылках на сайт, включая URL источника, URL назначения, дату обнаружения и последний доступ.
- **Экспорт популярных запросов:** Скачивание популярных запросов, показов, кликов и средней позиции. Данные можно сортировать и фильтровать по типу устройства и периоду времени.
- **Настройки периода:** Возможность выбрать диапазон дат (Неделя, Месяц, Квартал, Год) или указать собственные значения.
- **Сортировка по показам и кликам:** Экспорт данных о поисковых запросах с возможностью сортировки по показам или кликам.

## Установка

1. Установите менеджер пользовательских скриптов, например, [Tampermonkey](https://www.tampermonkey.net/) или [Greasemonkey](https://www.greasespot.net/).
2. Добавьте скрипт в менеджер, создав новый пользовательский скрипт и вставив содержимое файла `xbdyw.user.js`, либо нажмите [сюда](https://www.tampermonkey.net/script_installation.php#url=https://github.com/xxrxtnxxov/XBFYW/raw/refs/heads/main/xbdyw.user.js), чтобы установить его в 1 клик.

## Использование

1. Перейдите на [страницу документации Яндекс.Вебмастера](https://yandex.ru/dev/webmaster/doc/ru/tasks/how-to-get-oauth) и выполните все шаги для создания приложения. В конце скопируйте полученный токен в формате `y0_*****`.
2. Откройте Панель управления в Tampermonkey или Greasemonkey, найдите установленный скрипт и отредактируйте его, изменив строку `const oauthToken = 'ВАШ_OAUTH_TOKEN'`. Вставьте значение токена в поле `ВАШ_OAUTH_TOKEN`, чтобы получилось, например: `const oauthToken = 'y0_*****'`.
3. Сохраните скрипт с помощью сочетания клавиш Ctrl+S или через меню «Файл - Сохранить».
4. На странице внешних ссылок появится кнопка «Скачать выгрузку XLSX» для экспорта данных. На странице управления запросами будут доступны кнопки выбора периода, фильтры по устройствам и кнопка «Скачать все запросы (XLSX)».
5. Выберите параметры и нажмите кнопку скачивания. Файл будет скачан в формате `XLSX`.

## Зависимости

- `xlsx.full.min.js`: для сохранения данных в формате XLSX.
- `GM_xmlhttpRequest`: для выполнения HTTP-запросов к API.

## Примечания
- Доступные типы устройств для выгрузки данных:
  - Все устройства
  - ПК
  - Мобильные и планшеты
  - Только мобильные
  - Только планшеты
- Лимит строк в XLSX-файле составляет 1 000 000 строк.

Убедитесь, что у вас есть права на доступ к данным веб-сайта в Яндекс Вебмастере и что ваш OAuth токен действителен (он действует 6 месяцев).
