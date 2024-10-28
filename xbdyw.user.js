// ==UserScript==
// @name         Экспорт внешних ссылок из Яндекс Вебмастера
// @namespace    https://github.com/xxrxtnxxov/XBFYW
// @icon         https://yastatic.net/iconostasis/_/ERuVE3RAnz0A7BgIzyypFYnR18A.svg
// @version      4.4
// @description  Скрипт с помощью API извлекает все внешние ссылки сайта и сохраняет их в XLSX-файл (с ограничением в 1 000 000 строк).
// @author       ChatGPT-4o & xxrxtnxxov
// @match        https://webmaster.yandex.ru/site/*/links/incoming*
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.0/xlsx.full.min.js
// @grant        GM_xmlhttpRequest
// @connect      api.webmaster.yandex.net
// @license      MIT
// ==/UserScript==

(async function () {
    'use strict';

    const oauthToken = 'ВАШ_OAUTH_TOKEN'; // Замените на ваш OAuth токен
    let userId = null;
    let selectedHostId = null;
    let allData = [];

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

    // Получение всех внешних ссылок для выбранного host_id
    async function fetchAllLinks(updateProgress) {
        let offset = 0;
        const limit = 100;
        let hasMoreData = true;
        updateProgress("Начинаю загрузку данных...");

        while (hasMoreData) {
            const response = await new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method: "GET",
                    url: `https://api.webmaster.yandex.net/v4/user/${userId}/hosts/${selectedHostId}/links/external/samples?offset=${offset}&limit=${limit}`,
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

    // Экспорт данных в XLSX
    function exportToXLSX(data, updateProgress) {
        const headers = ["Источник", "Назначение", "Дата обнаружения", "Последний доступ"];
        data.unshift(headers);
        const worksheet = XLSX.utils.aoa_to_sheet(data);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Links Data');
        updateProgress("Экспортирую данные в файл XLSX...");
        XLSX.writeFile(workbook, 'Yandex_Links_Data.xlsx');
        updateProgress("Экспорт завершен!");
    }

    // Основная функция для начала процесса экспорта
    async function startExportProcess(updateProgress) {
        try {
            // Очистка данных перед каждым запуском, чтобы избежать дублирования
            allData = [];

            userId = await getUserId();
            const hosts = await getVerifiedHosts();

            if (hosts.length === 0) {
                updateProgress("Нет подтвержденных сайтов для экспорта.");
                return;
            }

            const currentHostUrl = window.location.href.match(/https:\/\/webmaster\.yandex\.ru\/site\/([^/]+)\/links\/incoming/)[1];
            const matchedHost = hosts.find(host => host.hostId.includes(currentHostUrl));

            if (!matchedHost) {
                updateProgress("Хост для текущего сайта не найден.");
                return;
            }

            selectedHostId = matchedHost.hostId;
            updateProgress(`Выгружаются ссылки для хоста: ${matchedHost.displayName}`);
            await fetchAllLinks(updateProgress);
            exportToXLSX(allData, updateProgress);
        } catch (error) {
            console.error("Ошибка экспорта данных:", error);
            updateProgress("Ошибка экспорта данных.");
        }
    }

    // Функция для добавления кнопки на страницу
    function addExportButton() {
        const container = document.evaluate(
            '//*[@id="root"]/div/div[2]/div/div/div[2]/div/div/span/div/div/div[2]/div[1]',
            document,
            null,
            XPathResult.FIRST_ORDERED_NODE_TYPE,
            null
        ).singleNodeValue;

        if (!container) {
            console.warn("Не удалось найти контейнер для кнопки.");
            return;
        }

        // Проверяем, есть ли уже кнопка, чтобы не дублировать
        if (!container.querySelector('#exportButton')) {
            const exportButton = document.createElement('button');
            exportButton.id = 'exportButton';
            exportButton.textContent = 'Скачать выгрузку XLSX';
            exportButton.style.marginLeft = '10px';
            exportButton.style.padding = '10px 15px';
            exportButton.style.fontWeight = 'bold';
            exportButton.style.color = '#ffffff';
            exportButton.style.backgroundColor = '#581845';
            exportButton.style.border = 'none';
            exportButton.style.cursor = 'pointer';

            exportButton.onclick = () => {
                startExportProcess((progressText) => {
                    exportButton.textContent = progressText;
                });
            };

            container.appendChild(exportButton);
        }
    }

    // Отслеживание изменений DOM и добавление кнопки при появлении нужного контейнера
    const observer = new MutationObserver(() => {
        addExportButton();
    });

    observer.observe(document.body, { childList: true, subtree: true });
})();