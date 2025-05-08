// static/script.js

document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('upload-form');
    const progressContainer = document.getElementById('progress-container');
    const progressBar = document.getElementById('progress-bar');
    const statusMessage = document.getElementById('status-message');
    const resultContainer = document.getElementById('result-container');
    const resultMessage = document.getElementById('result-message');
    const downloadLink = document.getElementById('download-link');
    const errorContainer = document.getElementById('error-container');

    // === Новые элементы ===
    // Убираем определения элементов отсюда:
    // const smetaTypeSelect = document.getElementById('smeta_type');
    // const fileInput = document.getElementById('file');
    // const turbosmetchikVersionGroup = document.getElementById('turbosmetchik-version-group');
    // const turbosmetchikVersionSelect = document.getElementById('turbosmetchik_version');
    // =====================

    let progressInterval = null; // Переменная для хранения ID интервала

    // Функция для генерации простого уникального ID сессии
    function generateClientSessionId() {
        return Date.now() + '-' + Math.random().toString(36).substring(2, 15);
    }

    // Функция для запроса прогресса
    async function pollProgress(sessionId) {
        try {
            const response = await fetch(`/progress/${sessionId}`);
            if (!response.ok) {
                // Если сервер не отвечает или ошибка, прекращаем поллинг
                console.error('Ошибка запроса прогресса:', response.status);
                stopPolling();
                return;
            }
            const data = await response.json();

            // Обновляем прогресс бар
            let percentage = 0;
            if (data.total > 0) {
                percentage = Math.round((data.processed / data.total) * 100);
            } else if (data.status && data.status !== "Ошибка" && data.status !== "Готово" && data.status !== "Не найдено") {
                 // Если total еще 0, но идет работа (например, распаковка), показываем небольшой прогресс
                 percentage = 10; // Или другое значение
            }
            // Проверяем progressBar перед использованием
            if (progressBar) progressBar.style.width = `${percentage}%`;

            // Обновляем текстовый статус
            // Проверяем statusMessage перед использованием
            if (statusMessage && data.status) {
                statusMessage.textContent = data.status;
            }

            // Если есть ошибка от сервера или статус "Готово" или "Ошибка", останавливаем поллинг
            if (data.error || data.status === "Готово" || data.status === "Ошибка") {
                stopPolling();
            }

        } catch (error) {
            console.error('Сетевая ошибка при запросе прогресса:', error);
            stopPolling(); // Останавливаем при сетевых ошибках
        }
    }

    // Функция для остановки поллинга
    function stopPolling() {
        if (progressInterval) {
            clearInterval(progressInterval);
            progressInterval = null;
            console.log("Поллинг прогресса остановлен.");
        }
    }

    // === Новая функция: Проверка валидности формы и управление UI ===
    function checkFormValidity() {
        // Получаем ВСЕ нужные элементы здесь:
        const smetaTypeSelect = document.getElementById('smeta_type');
        const fileInput = document.getElementById('file');
        const turbosmetchikVersionGroup = document.getElementById('turbosmetchik-version-group');
        const turbosmetchikVersionSelect = document.getElementById('turbosmetchik_version');
        // const submitButton = document.getElementById('submit-button'); // Кнопку можно не получать, если не меняем disabled

        // Проверяем, что основные элементы формы найдены
        if (!smetaTypeSelect || !fileInput || !turbosmetchikVersionGroup || !turbosmetchikVersionSelect) {
            console.error("Один или несколько элементов формы не найдены в checkFormValidity!");
            return; // Прерываем выполнение, если чего-то нет
        }

        // const fileSelected = fileInput.files.length > 0; // Можно убрать, если не используется
        const mainTypeSelected = smetaTypeSelect.value;

        // Логика показа/скрытия и required для версии Турбосметчика остается
        if (mainTypeSelected === 'Турбосметчик') {
            turbosmetchikVersionGroup.style.display = 'block';
            turbosmetchikVersionSelect.required = true;
            // versionSelected = turbosmetchikVersionSelect.value !== ''; // Не используется для disabled
        } else {
            turbosmetchikVersionGroup.style.display = 'none';
            turbosmetchikVersionSelect.required = false;
            turbosmetchikVersionSelect.value = '';
        }

        // --- УДАЛЯЕМ УПРАВЛЕНИЕ КНОПКОЙ И ЛОГ --- 
        /*
        const submitButton = document.getElementById('submit-button');
        if(submitButton) {
            // ... (код установки disabled) ... 
            console.log('[DEBUG] Состояние кнопки disabled:', submitButton.disabled);
        }
        */
        // ---------------------------------------
    }
    // ===================================================================

    // === Обработка выбора основного типа сметы ===
    const initialSmetaTypeSelect = document.getElementById('smeta_type');
    if (initialSmetaTypeSelect) {
        initialSmetaTypeSelect.addEventListener('change', checkFormValidity);
    } else {
        console.error("Не удалось найти smeta_type для добавления слушателя");
    }
    // ========================================================

    // === Обработка выбора версии Турбосметчика ===
    const initialTurbosmetchikVersionSelect = document.getElementById('turbosmetchik_version');
    if (initialTurbosmetchikVersionSelect) {
        initialTurbosmetchikVersionSelect.addEventListener('change', checkFormValidity);
    } else {
        console.error("Не удалось найти turbosmetchik_version для добавления слушателя");
    }
    // =========================================================

    // === Обработка выбора файла ===
    const initialFileInput = document.getElementById('file');
    if (initialFileInput) {
        initialFileInput.addEventListener('change', checkFormValidity);
    } else {
        console.error("Не удалось найти file input для добавления слушателя");
    }
    // ============================================

    if (form) { // Проверяем, найдена ли форма
        form.addEventListener('submit', async (event) => {
            event.preventDefault(); // Оставляем preventDefault, т.к. отправка асинхронная
            stopPolling();

            // --- УДАЛЯЕМ БЛОКИРОВКУ КНОПКИ ЗДЕСЬ --- 
            /*
            const submitButtonOnSubmit = document.getElementById('submit-button'); 
            if (submitButtonOnSubmit) submitButtonOnSubmit.disabled = true;
            */
           // ---------------------------------------

            // Показываем прогресс и статус
            if(progressContainer) progressContainer.style.display = 'block';
            if(progressBar) progressBar.style.width = '0%';
            // if(progressBar) progressBar.textContent = ''; // Очищаем текст на баре (если был)
            if(statusMessage) statusMessage.textContent = 'Загрузка файла...';
            if(resultContainer) resultContainer.style.display = 'none';
            if(errorContainer) errorContainer.style.display = 'none';
            if(errorContainer) errorContainer.textContent = '';

            // Получаем актуальные значения элементов формы ПЕРЕД отправкой
            const currentFileInput = document.getElementById('file');
            const currentSmetaTypeSelect = document.getElementById('smeta_type');
            const currentTurbosmetchikVersionSelect = document.getElementById('turbosmetchik_version');

            if (!currentFileInput || !currentSmetaTypeSelect || !currentTurbosmetchikVersionSelect) {
                 console.error("Ошибка: Не найдены элементы формы при отправке!");
                 if(progressContainer) progressContainer.style.display = 'none'; // Скрываем прогресс при ошибке
                 if(submitButtonOnSubmit) submitButtonOnSubmit.disabled = false; // Разблокируем кнопку
                 return;
            }
            if (currentFileInput.files.length === 0) {
                 console.error("Ошибка: Файл не выбран перед отправкой!");
                 if(progressContainer) progressContainer.style.display = 'none';
                 if(errorContainer) {
                      errorContainer.textContent = 'Ошибка: Файл не выбран.';
                      errorContainer.style.display = 'block';
                 }
                 if(submitButtonOnSubmit) submitButtonOnSubmit.disabled = false; // Разблокируем кнопку
                 return;
            }

            const clientSessionId = generateClientSessionId();
            console.log("Новая сессия:", clientSessionId);

            const formData = new FormData();
            formData.append('file', currentFileInput.files[0]);
            formData.append('client_session_id', clientSessionId);

            let finalSmetaType = currentSmetaTypeSelect.value;
            if (finalSmetaType === 'Турбосметчик') {
                 finalSmetaType += '-' + currentTurbosmetchikVersionSelect.value;
            }
            formData.append('smeta_type', finalSmetaType);
            console.log("Отправляемый тип сметы:", finalSmetaType);

            // --- Запускаем поллинг прогресса ---
            progressInterval = setInterval(() => {
                pollProgress(clientSessionId);
            }, 1500); // Запрашивать каждые 1.5 секунды
            // ---------------------------------

            try {
                const response = await fetch('/upload', {
                    method: 'POST',
                    body: formData,
                });

                stopPolling();

                const data = await response.json();

                 // Небольшая задержка перед скрытием прогресс-бара, если успешно
                if (response.ok && data.success) {
                    if(progressBar) progressBar.style.width = '100%';
                    if(statusMessage) statusMessage.textContent = data.message || 'Готово';
                    await new Promise(resolve => setTimeout(resolve, 500));
                }
                 if(progressContainer) progressContainer.style.display = 'none'; // Скрываем прогресс


                if (response.ok && data.success) {
                    if(resultMessage) resultMessage.textContent = data.message || 'Обработка успешно завершена!';
                    if(downloadLink) downloadLink.href = data.download_url;
                    let buttonText = "Скачать результат";
                    if (data.download_filename) {
                       if(downloadLink) downloadLink.setAttribute('download', data.download_filename);
                        if (data.download_filename.toLowerCase().endsWith('.zip')) buttonText = "Скачать результаты (ZIP)";
                        else if (data.download_filename.toLowerCase().endsWith('.xlsx') || data.download_filename.toLowerCase().endsWith('.xlsm')) buttonText = "Скачать результат (Excel)";
                        else { const parts = data.download_filename.split('.'); const extension = parts.length > 1 ? parts.pop() : 'файл'; buttonText = `Скачать (${extension.toUpperCase()})`; }
                    } else if(downloadLink) downloadLink.removeAttribute('download');
                    if(downloadLink) downloadLink.textContent = buttonText;
                    if(resultContainer) resultContainer.style.display = 'block';
                    if(errorContainer) errorContainer.style.display = 'none';
                } else {
                    throw new Error(data.error || `Ошибка сервера: ${response.status}`);
                }

            } catch (error) {
                stopPolling();
                console.error('Ошибка при отправке или обработке:', error);
                if(progressContainer) progressContainer.style.display = 'none';
                if(errorContainer) {
                     errorContainer.textContent = `Произошла ошибка: ${error.message}`;
                     errorContainer.style.display = 'block';
                }
                if(resultContainer) resultContainer.style.display = 'none';
            } finally {
                // Очистка полей остается, НО НЕ ФАЙЛА
                const finalFileInput = document.getElementById('file');
                // Удаляем строку ниже:
                // Строка ниже была удалена

                // --- УДАЛЯЕМ РАЗБЛОКИРОВКУ КНОПКИ ЗДЕСЬ ---
                /*
                const submitButtonFinally = document.getElementById('submit-button');
                if (submitButtonFinally) submitButtonFinally.disabled = false;
                */
                // ---------------------------------------

                // Вызов checkFormValidity в конце нужен, чтобы скрыть/показать поле версии
                checkFormValidity();
            }
        });
    } else {
        console.error("Форма с id 'upload-form' не найдена!");
    }

    // Вызов checkFormValidity при загрузке нужен только для установки видимости поля версии
    checkFormValidity(); 

});