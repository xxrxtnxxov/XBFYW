// ==UserScript==
// @name         Экспорт внешних ссылок и популярных запросов из Яндекс Вебмастера
// @namespace    https://github.com/xxrxtnxxov/XBFYW
// @icon         https://yastatic.net/iconostasis/_/ERuVE3RAnz0A7BgIzyypFYnR18A.svg
// @version      6.7.5
// @description  Скрипт с помощью API Я.Вебмастер может извлекать все внешние ссылки и популярные запросы в формате XLSX (лимит строк равен 1 000 000).
// @author       xxrxtnxxov
// @match        https://webmaster.yandex.ru/site/*/links/incoming*
// @match        https://webmaster.yandex.ru/site/*/efficiency/queries*
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.0/xlsx.full.min.js
// @grant        GM_xmlhttpRequest
// @connect      api.webmaster.yandex.net
// @license      MIT
// ==/UserScript==

(function () {
    'use strict';

    const oauthToken = 'ВАШ_OAUTH_TOKEN'; // Замените на ваш OAuth токен
    let userId = null;
    let selectedHostId = null;
    let allData = [];
    let numberOfDays = 1; // Количество выбранных дней для расчета средних значений

    // Функция для получения текущей даты в формате YYYY-MM-DD с учетом локального времени
    function getCurrentDateString() {
        const today = new Date();
        today.setMinutes(today.getMinutes() - today.getTimezoneOffset());
        return today.toISOString().split('T')[0];
    }

    // Функция для форматирования даты в формате YYYY-MM-DD с учетом локального времени
    function formatDate(date) {
        date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
        return date.toISOString().split('T')[0];
    }

    // Функция для получения user_id
    async function getUserId() {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                method: "GET",
                url: `https://api.webmaster.yandex.net/v4/user`,
                headers: { 'Authorization': `OAuth ${oauthToken}` },
                onload: function (response) {
                    const data = JSON.parse(response.responseText);
                    if (response.status === 200) {
                        resolve(data.user_id);
                    } else {
                        console.error("Ошибка при получении user_id:", data);
                        reject("Ошибка авторизации. Проверьте токен.");
                    }
                },
                onerror: function (error) {
                    reject("Ошибка сети при получении user_id");
                }
            });
        });
    }

    // Функция для получения списка подтверждённых сайтов пользователя
    async function getVerifiedHosts() {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                method: "GET",
                url: `https://api.webmaster.yandex.net/v4/user/${userId}/hosts`,
                headers: { 'Authorization': `OAuth ${oauthToken}` },
                onload: function (response) {
                    const data = JSON.parse(response.responseText);
                    if (response.status === 200) {
                        const verifiedHosts = data.hosts.filter(host => host.verified).map(host => ({
                            hostId: host.host_id,
                            asciiUrl: host.ascii_host_url,
                            unicodeUrl: host.unicode_host_url,
                            displayName: host.host_display_name || host.ascii_host_url
                        }));
                        resolve(verifiedHosts);
                    } else {
                        console.error("Ошибка при получении списка хостов:", data);
                        reject("Ошибка при получении списка хостов");
                    }
                },
                onerror: function (error) {
                    reject("Ошибка сети при получении списка хостов");
                }
            });
        });
    }

    // Функции для экспорта поисковых запросов
    const deviceTypeMapping = {
        'Все устройства': 'ALL',
        'ПК': 'DESKTOP',
        'Мобильные и планшеты': 'MOBILE_AND_TABLET',
        'Только мобильные': 'MOBILE',
        'Только планшеты': 'TABLET'
    };

    async function fetchSearchQueries(deviceType, orderBy, dateFrom, dateTo, updateProgress) {
        let offset = 0;
        const limit = 500;
        let hasMoreData = true;
        updateProgress("Начинаю загрузку данных...");

        // Вычисляем количество дней между датами для расчета средних значений
        const startDate = new Date(dateFrom);
        const endDate = new Date(dateTo);
        numberOfDays = Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;

        while (hasMoreData) {
            const queryIndicators = ['TOTAL_SHOWS', 'TOTAL_CLICKS', 'AVG_SHOW_POSITION', 'AVG_CLICK_POSITION'];
            const queryIndicatorParams = queryIndicators.map(indicator => `query_indicator=${indicator}`).join('&');

            const url = `https://api.webmaster.yandex.net/v4/user/${userId}/hosts/${selectedHostId}/search-queries/popular?order_by=${orderBy}&${queryIndicatorParams}&device_type_indicator=${deviceType}&date_from=${dateFrom}&date_to=${dateTo}&offset=${offset}&limit=${limit}`;

            const response = await new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method: "GET",
                    url: url,
                    headers: { 'Authorization': `OAuth ${oauthToken}` },
                    onload: function (response) {
                        if (response.status === 200) {
                            resolve(JSON.parse(response.responseText));
                        } else {
                            console.error("Ошибка при получении запросов:", response.responseText);
                            reject("Ошибка при получении запросов");
                        }
                    },
                    onerror: function (error) {
                        reject("Ошибка сети при получении запросов");
                    }
                });
            });

            const pageData = response.queries.map(query => {
                const indicators = query.indicators || {};
                const totalShows = indicators.TOTAL_SHOWS || 0;
                const totalClicks = indicators.TOTAL_CLICKS || 0;
                const avgShows = (totalShows / numberOfDays).toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
                const avgClicks = (totalClicks / numberOfDays).toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
                const avgShowPosition = indicators.AVG_SHOW_POSITION ? indicators.AVG_SHOW_POSITION.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '0,00';
                const avgClickPosition = indicators.AVG_CLICK_POSITION ? indicators.AVG_CLICK_POSITION.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '0,00';
                const ctr = totalShows > 0 ? ((totalClicks / totalShows) * 100).toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '0,00';

                return [
                    query.query_text,
                    totalShows,
                    avgShows,
                    totalClicks,
                    avgClicks,
                    avgShowPosition,
                    avgClickPosition,
                    ctr
                ];
            });

            allData = allData.concat(pageData);
            offset += limit;
            hasMoreData = pageData.length === limit;

            // Обновление прогресса для каждого шага загрузки
            updateProgress(`Загружено ${allData.length} записей...`);
        }

        updateProgress("Данные загружены. Подготовка к экспорту...");
    }

    function exportQueriesToXLSX(data, updateProgress) {
        const headers = [
            "Поисковый запрос",
            "Всего показов",
            "Среднее количество показов",
            "Всего кликов",
            "Среднее количество кликов",
            "Средняя позиция показа",
            "Средняя позиция клика",
            "CTR"
        ];
        data.unshift(headers);
        const worksheet = XLSX.utils.aoa_to_sheet(data);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Search Queries');
        updateProgress("Экспортирую данные в файл XLSX...");
        XLSX.writeFile(workbook, 'Yandex_Search_Queries.xlsx');
        updateProgress("Экспорт завершен!");
    }

    // Основная функция для экспорта поисковых запросов
    async function startExportQueriesProcess(deviceType, orderBy, dateFrom, dateTo, updateProgress) {
        try {
            // Очистка данных перед каждым запуском
            allData = [];

            userId = await getUserId();
            const hosts = await getVerifiedHosts();

            if (hosts.length === 0) {
                updateProgress("Нет подтвержденных сайтов для экспорта.");
                return;
            }

            const currentHostUrl = window.location.href.match(/https:\/\/webmaster\.yandex\.ru\/site\/([^\/]+)\/efficiency\/queries/)[1];
            const matchedHost = hosts.find(host => host.hostId.includes(currentHostUrl));

            if (!matchedHost) {
                updateProgress("Хост для текущего сайта не найден.");
                return;
            }

            selectedHostId = matchedHost.hostId;
            updateProgress(`Выгружаются запросы для хоста: ${matchedHost.displayName}`);
            await fetchSearchQueries(deviceType, orderBy, dateFrom, dateTo, updateProgress);
            exportQueriesToXLSX(allData, updateProgress);
        } catch (error) {
            console.error("Ошибка экспорта данных:", error);
            updateProgress("Ошибка экспорта данных.");
        }
    }

    // Функция для добавления элементов управления на страницу поисковых запросов
    function addExportQueriesControls() {
        // Используем предоставленный селектор для контейнера
        const container = document.querySelector('#root > div > div.AsideHeader.AsideHeader_withRadius > div > div > div.gn-aside-header__content > div > div > span > div > div.queries-table.content__chunk.checkgroup.i-bem.checkgroup_js_inited > div > div.queries-group-header.i-bem.queries-group-header_js_inited');

        if (!container) {
            console.warn("Не удалось найти контейнер для элементов управления.");
            return;
        }

        // Проверяем, есть ли уже элементы, чтобы не дублировать
        if (!document.getElementById('exportQueriesButton')) {
            const controlsWrapper = document.createElement('div');
            controlsWrapper.style.display = 'flex';
            controlsWrapper.style.alignItems = 'center';
            controlsWrapper.style.justifyContent = 'flex-start'; // Выравнивание по левому краю
            controlsWrapper.style.width = '100%';
            controlsWrapper.style.flexWrap = 'wrap';
            controlsWrapper.style.marginBottom = '10px'; // Отступ снизу для увеличения расстояния между строками

            // Создаем контейнер для выбора дат
            const dateControlsWrapper = document.createElement('div');
            dateControlsWrapper.style.display = 'flex';
            dateControlsWrapper.style.alignItems = 'center';
            dateControlsWrapper.style.marginRight = '10px';

            // Создаем быстрые кнопки выбора периода
            const periods = [
                { label: 'Неделя', days: 7 },
                { label: 'Месяц', days: 30 },
                { label: 'Квартал', days: 90 },
                { label: 'Год', days: 365 }
            ];

            periods.forEach(period => {
                const periodButton = document.createElement('button');
                periodButton.textContent = period.label;
                periodButton.style.marginRight = '5px';
                periodButton.style.padding = '5px 10px';
                periodButton.style.cursor = 'pointer';

                periodButton.onclick = () => {
                    const endDate = new Date();
                    const startDate = new Date();
                    startDate.setDate(endDate.getDate() - (period.days - 1));

                    dateFromInput.value = formatDate(startDate);
                    dateToInput.value = formatDate(endDate);
                };

                dateControlsWrapper.appendChild(periodButton);
            });

            // Создаем элементы для выбора даты
            const dateFromInput = document.createElement('input');
            dateFromInput.type = 'date';
            dateFromInput.id = 'dateFromInput';
            dateFromInput.style.marginRight = '5px';
            dateFromInput.style.padding = '5px';
            dateFromInput.style.fontSize = '14px';
            const initialStartDate = new Date();
            initialStartDate.setDate(initialStartDate.getDate() - 7);
            dateFromInput.value = formatDate(initialStartDate);
            dateFromInput.max = getCurrentDateString(); // Ограничение на выбор будущих дат

            const dateToInput = document.createElement('input');
            dateToInput.type = 'date';
            dateToInput.id = 'dateToInput';
            dateToInput.style.marginRight = '10px';
            dateToInput.style.padding = '5px';
            dateToInput.style.fontSize = '14px';
            dateToInput.value = getCurrentDateString();
            dateToInput.max = getCurrentDateString(); // Ограничение на выбор будущих дат

            dateControlsWrapper.appendChild(dateFromInput);
            dateControlsWrapper.appendChild(dateToInput);

            // Добавляем отступы между элементами
            dateControlsWrapper.style.marginRight = '10px';

            // Создаем выпадающий список для типа устройства
            const deviceSelect = document.createElement('select');
            deviceSelect.id = 'deviceTypeSelect';
            deviceSelect.style.marginRight = '10px';
            deviceSelect.style.padding = '5px';
            deviceSelect.style.fontSize = '14px';

            for (const [key, value] of Object.entries(deviceTypeMapping)) {
                const option = document.createElement('option');
                option.value = value;
                option.textContent = key;
                deviceSelect.appendChild(option);
            }

            // Создаем радиокнопки для выбора порядка сортировки
            const sortWrapper = document.createElement('div');
            sortWrapper.style.display = 'flex';
            sortWrapper.style.alignItems = 'center';
            sortWrapper.style.marginRight = '10px';

            const sortLabel = document.createElement('span');
            sortLabel.textContent = 'Сортировать по: ';
            sortLabel.style.marginRight = '5px';
            sortWrapper.appendChild(sortLabel);

            const sortByShowsLabel = document.createElement('label');
            sortByShowsLabel.style.marginRight = '10px';
            sortByShowsLabel.style.marginLeft = '5px';
            const sortByShows = document.createElement('input');
            sortByShows.type = 'radio';
            sortByShows.name = 'sortOrder';
            sortByShows.value = 'TOTAL_SHOWS';
            sortByShows.checked = true; // По умолчанию сортировка по показам
            sortByShowsLabel.appendChild(sortByShows);
            sortByShowsLabel.appendChild(document.createTextNode('Показам'));

            const sortByClicksLabel = document.createElement('label');
            sortByClicksLabel.style.marginLeft = '5px';
            const sortByClicks = document.createElement('input');
            sortByClicks.type = 'radio';
            sortByClicks.name = 'sortOrder';
            sortByClicks.value = 'TOTAL_CLICKS';
            sortByClicksLabel.appendChild(sortByClicks);
            sortByClicksLabel.appendChild(document.createTextNode('Кликам'));

            sortWrapper.appendChild(sortByShowsLabel);
            sortWrapper.appendChild(sortByClicksLabel);

            // Создаем кнопку
            const exportButton = document.createElement('button');
            exportButton.id = 'exportQueriesButton';
            exportButton.textContent = 'Скачать все запросы (XLSX)';
            exportButton.style.padding = '10px 15px';
            exportButton.style.fontWeight = 'bold';
            exportButton.style.color = '#ffffff';
            exportButton.style.backgroundColor = '#581845';
            exportButton.style.border = 'none';
            exportButton.style.cursor = 'pointer';

            exportButton.onclick = () => {
                const selectedDeviceType = deviceSelect.value;
                const selectedOrderBy = document.querySelector('input[name="sortOrder"]:checked').value;
                const dateFrom = dateFromInput.value;
                const dateTo = dateToInput.value;

                if (new Date(dateFrom) > new Date(dateTo)) {
                    alert("Дата начала не может быть позже даты окончания.");
                    return;
                }

                exportButton.disabled = true;
                deviceSelect.disabled = true;
                sortByShows.disabled = true;
                sortByClicks.disabled = true;
                dateFromInput.disabled = true;
                dateToInput.disabled = true;
                const originalButtonText = exportButton.textContent;

                startExportQueriesProcess(selectedDeviceType, selectedOrderBy, dateFrom, dateTo, (progressText) => {
                    exportButton.textContent = progressText;
                }).finally(() => {
                    exportButton.disabled = false;
                    deviceSelect.disabled = false;
                    sortByShows.disabled = false;
                    sortByClicks.disabled = false;
                    dateFromInput.disabled = false;
                    dateToInput.disabled = false;
                    exportButton.textContent = originalButtonText;
                });
            };

            // Добавляем элементы в обертку
            controlsWrapper.appendChild(dateControlsWrapper);
            controlsWrapper.appendChild(deviceSelect);
            controlsWrapper.appendChild(sortWrapper);
            controlsWrapper.appendChild(exportButton);

            // Добавляем обертку в контейнер
            container.appendChild(controlsWrapper);
        }
    }

    // Функция для добавления кнопки экспорта внешних ссылок (без изменений)
    function addExportLinksButton() {
        // Используем предоставленный селектор для контейнера
        const container = document.querySelector('#root > div > div.AsideHeader.AsideHeader_withRadius > div > div > div.gn-aside-header__content > div > div > span > div > div > div.external-links-table.external-links-table_state_current.i-bem.external-links-table_viewType_current-links.external-links-table_js_inited.external-links-table_loaded_yes > div.external-links-table__header');

        if (!container) {
            console.warn("Не удалось найти контейнер для кнопки.");
            return;
        }

        // Проверяем, есть ли уже кнопка, чтобы не дублировать
        if (!document.getElementById('exportLinksButton')) {
            const exportButton = document.createElement('button');
            exportButton.id = 'exportLinksButton';
            exportButton.textContent = 'Скачать выгрузку XLSX';
            exportButton.style.marginLeft = '10px';
            exportButton.style.padding = '10px 15px';
            exportButton.style.fontWeight = 'bold';
            exportButton.style.color = '#ffffff';
            exportButton.style.backgroundColor = '#581845';
            exportButton.style.border = 'none';
            exportButton.style.cursor = 'pointer';

            exportButton.onclick = () => {
                exportButton.disabled = true;
                const originalText = exportButton.textContent;
                startExportLinksProcess((progressText) => {
                    exportButton.textContent = progressText;
                }).finally(() => {
                    exportButton.disabled = false;
                    exportButton.textContent = originalText;
                });
            };

            container.appendChild(exportButton);
        }
    }

    // Основная функция для экспорта внешних ссылок (без изменений)
    async function startExportLinksProcess(updateProgress) {
        try {
            // Очистка данных перед каждым запуском
            allData = [];

            userId = await getUserId();
            const hosts = await getVerifiedHosts();

            if (hosts.length === 0) {
                updateProgress("Нет подтвержденных сайтов для экспорта.");
                return;
            }

            const currentHostUrl = window.location.href.match(/https:\/\/webmaster\.yandex\.ru\/site\/([^\/]+)\/links\/incoming/)[1];
            const matchedHost = hosts.find(host => host.hostId.includes(currentHostUrl));

            if (!matchedHost) {
                updateProgress("Хост для текущего сайта не найден.");
                return;
            }

            selectedHostId = matchedHost.hostId;
            updateProgress(`Выгружаются ссылки для хоста: ${matchedHost.displayName}`);
            await fetchAllLinks(updateProgress);
            exportLinksToXLSX(allData, updateProgress);
        } catch (error) {
            console.error("Ошибка экспорта данных:", error);
            updateProgress("Ошибка экспорта данных.");
        }
    }

    // Функции для экспорта внешних ссылок (без изменений)
    async function fetchAllLinks(updateProgress) {
        let offset = 0;
        const limit = 100;
        let hasMoreData = true;
        updateProgress("Начинаю загрузку данных...");

        while (hasMoreData) {
            const url = `https://api.webmaster.yandex.net/v4/user/${userId}/hosts/${selectedHostId}/links/external/samples?offset=${offset}&limit=${limit}`;

            const response = await new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method: "GET",
                    url: url,
                    headers: { 'Authorization': `OAuth ${oauthToken}` },
                    onload: function (response) {
                        if (response.status === 200) {
                            resolve(JSON.parse(response.responseText));
                        } else {
                            console.error("Ошибка при получении ссылок:", response.responseText);
                            reject("Ошибка при получении ссылок");
                        }
                    },
                    onerror: function (error) {
                        reject("Ошибка сети при получении ссылок");
                    }
                });
            });

            const pageData = response.links.map(link => [
                link.source_url,
                link.destination_url,
                link.discovery_date,
                link.source_last_access_date,
            ]);

            allData = allData.concat(pageData);
            offset += limit;
            hasMoreData = pageData.length === limit;

            // Обновление прогресса для каждого шага загрузки
            updateProgress(`Загружено ${allData.length} записей...`);
        }

        updateProgress("Данные загружены. Подготовка к экспорту...");
    }

    function exportLinksToXLSX(data, updateProgress) {
        const headers = ["Источник", "Назначение", "Дата обнаружения", "Последний доступ"];
        data.unshift(headers);
        const worksheet = XLSX.utils.aoa_to_sheet(data);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Links Data');
        updateProgress("Экспортирую данные в файл XLSX...");
        XLSX.writeFile(workbook, 'Yandex_Links_Data.xlsx');
        updateProgress("Экспорт завершен!");
    }

    // Определяем на какой странице мы находимся и добавляем соответствующие элементы управления
    function initialize() {
        const currentUrl = window.location.href;

        if (/https:\/\/webmaster\.yandex\.ru\/site\/[^\/]+\/links\/incoming/.test(currentUrl)) {
            // Страница внешних ссылок
            const observer = new MutationObserver(() => {
                addExportLinksButton();
            });

            observer.observe(document.body, { childList: true, subtree: true });
        } else if (/https:\/\/webmaster\.yandex\.ru\/site\/[^\/]+\/efficiency\/queries/.test(currentUrl)) {
            // Страница популярных запросов
            const observer = new MutationObserver(() => {
                addExportQueriesControls();
            });

            observer.observe(document.body, { childList: true, subtree: true });
        }
    }

    // Запускаем инициализацию
    initialize();

})();
