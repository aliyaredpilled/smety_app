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

    // --- НОВЫЙ ЭЛЕМЕНТ ДЛЯ ДЕТАЛЕЙ ОТЧЕТА ---
    const reportDetailsContainer = document.getElementById('report-details');
    // --- КОНЕЦ НОВОГО ЭЛЕМЕНТА ---


    let progressInterval = null;

    function generateClientSessionId() {
        return Date.now() + '-' + Math.random().toString(36).substring(2, 15);
    }

    async function pollProgress(sessionId) {
        try {
            const response = await fetch(`/progress/${sessionId}`);
            if (!response.ok) {
                console.error('Ошибка запроса прогресса:', response.status);
                stopPolling();
                return;
            }
            const data = await response.json();

            // Обновляем прогресс бар на основе data.percent от сервера
            if (progressBar && typeof data.percent !== 'undefined') {
                progressBar.style.width = `${data.percent}%`;
                progressBar.textContent = `${data.percent}%`; // Показываем процент на баре
            }


            if (statusMessage && data.status) {
                let statusText = data.status;
                // Улучшенная проверка и форматирование статуса
                if (typeof statusText === 'string' && statusText.startsWith('Обработка')) {
                    const separatorIndex = statusText.indexOf(': ');
                    if (separatorIndex !== -1) {
                        // Нашли ": ", разделяем строку
                        const prefix = statusText.substring(0, separatorIndex);
                        let pathPart = statusText.substring(separatorIndex + 2);

                        // Удаляем возможное многоточие в конце пути
                        if (pathPart.endsWith('...')) {
                            pathPart = pathPart.slice(0, -3);
                        }

                        // Заменяем слеши только в части с путем
                        const formattedPath = pathPart.replace(/\//g, ' /\n');
                        
                        // Собираем обратно с переносом строки ПОСЛЕ двоеточия
                        statusText = prefix + ':\n' + formattedPath; 
                    } else if (statusText.includes('/')) {
                         // Если ": " нет, но есть слеши (старое поведение как fallback)
                         // Удаляем многоточие и здесь на всякий случай?
                         if (statusText.endsWith('...')) {
                             statusText = statusText.slice(0, -3);
                         }
                         statusText = statusText.replace(/\//g, ' /\n'); 
                    }
                    // Если нет ни ": ", ни '/', оставляем как есть
                }
                statusMessage.textContent = statusText;
            }

            if (data.error || data.status === "Готово" || data.status === "Ошибка" || data.status === "Сессия не найдена") {
                stopPolling();
                 // Если финальный статус, но progressBar не 100, доводим его
                if ((data.status === "Готово" || data.status === "Ошибка") && progressBar && progressBar.style.width !== '100%') {
                    progressBar.style.width = '100%';
                    progressBar.textContent = '100%';
                }
            }

        } catch (error) {
            console.error('Сетевая ошибка при запросе прогресса:', error);
            stopPolling();
        }
    }

    function stopPolling() {
        if (progressInterval) {
            clearInterval(progressInterval);
            progressInterval = null;
            console.log("Поллинг прогресса остановлен.");
        }
    }
    
    // Функция для экранирования HTML (важно для безопасности)
    function escapeHtml(unsafe) {
        if (typeof unsafe !== 'string') {
            if (unsafe === null || typeof unsafe === 'undefined') return '';
            unsafe = String(unsafe);
        }
        return unsafe
             .replace(/&/g, "&amp;")
             .replace(/</g, "&lt;")
             .replace(/>/g, "&gt;")
             .replace(/"/g, "&quot;")
             .replace(/'/g, "&#039;");
    }


    function checkFormValidity() {
        const smetaTypeSelect = document.getElementById('smeta_type');
        const fileInput = document.getElementById('file');
        const turbosmetchikVersionGroup = document.getElementById('turbosmetchik-version-group');
        const turbosmetchikVersionSelect = document.getElementById('turbosmetchik_version');

        if (!smetaTypeSelect || !fileInput || !turbosmetchikVersionGroup || !turbosmetchikVersionSelect) {
            console.error("Один или несколько элементов формы не найдены в checkFormValidity!");
            return;
        }

        const mainTypeSelected = smetaTypeSelect.value;

        if (mainTypeSelected === 'Турбосметчик') {
            turbosmetchikVersionGroup.style.display = 'block';
            turbosmetchikVersionSelect.required = true;
        } else {
            turbosmetchikVersionGroup.style.display = 'none';
            turbosmetchikVersionSelect.required = false;
            turbosmetchikVersionSelect.value = '';
        }
    }

    const initialSmetaTypeSelect = document.getElementById('smeta_type');
    if (initialSmetaTypeSelect) {
        initialSmetaTypeSelect.addEventListener('change', checkFormValidity);
    } else {
        console.error("Не удалось найти smeta_type для добавления слушателя");
    }

    const initialTurbosmetchikVersionSelect = document.getElementById('turbosmetchik_version');
    if (initialTurbosmetchikVersionSelect) {
        initialTurbosmetchikVersionSelect.addEventListener('change', checkFormValidity);
    } else {
        console.error("Не удалось найти turbosmetchik_version для добавления слушателя");
    }

    const initialFileInput = document.getElementById('file');
    if (initialFileInput) {
        initialFileInput.addEventListener('change', checkFormValidity);
    } else {
        console.error("Не удалось найти file input для добавления слушателя");
    }

    if (form) {
        form.addEventListener('submit', async (event) => {
            event.preventDefault();
            stopPolling();

            if(progressContainer) progressContainer.style.display = 'block';
            if(progressBar) {
                progressBar.style.width = '0%';
                progressBar.textContent = '0%'; // Обновляем текст на баре
            }
            if(statusMessage) statusMessage.textContent = 'Загрузка файла...';
            if(resultContainer) resultContainer.style.display = 'none';
            if(downloadLink) downloadLink.href = '#'; // Сброс ссылки
            if(downloadLink) downloadLink.textContent = ''; // Сброс текста кнопки/ссылки
            if(errorContainer) {
                errorContainer.style.display = 'none';
                errorContainer.textContent = '';
            }
            // --- СБРОС ДЕТАЛЕЙ ОТЧЕТА ---
            if(reportDetailsContainer) {
                reportDetailsContainer.innerHTML = '';
                reportDetailsContainer.style.display = 'none';
            }
            // --- КОНЕЦ СБРОСА ---


            const currentFileInput = document.getElementById('file');
            const currentSmetaTypeSelect = document.getElementById('smeta_type');
            const currentTurbosmetchikVersionSelect = document.getElementById('turbosmetchik_version');

            if (!currentFileInput || !currentSmetaTypeSelect || !currentTurbosmetchikVersionSelect) {
                 console.error("Ошибка: Не найдены элементы формы при отправке!");
                 if(progressContainer) progressContainer.style.display = 'none';
                 return;
            }
            if (currentFileInput.files.length === 0) {
                 console.error("Ошибка: Файл не выбран перед отправкой!");
                 if(progressContainer) progressContainer.style.display = 'none';
                 if(errorContainer) {
                      errorContainer.textContent = 'Ошибка: Файл не выбран.';
                      errorContainer.style.display = 'block';
                 }
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

            progressInterval = setInterval(() => {
                pollProgress(clientSessionId);
            }, 1500);

            try {
                const response = await fetch('/upload', {
                    method: 'POST',
                    body: formData,
                });

                stopPolling(); // Останавливаем поллинг сразу после получения ответа

                const responseBodyText = await response.text(); // Сначала читаем как текст

                let data;
                try {
                    data = JSON.parse(responseBodyText); // Пытаемся парсить JSON
                } catch (parseError) {
                    if(errorContainer) {
                         errorContainer.textContent = `Ошибка: Не удалось обработать ответ сервера (не JSON?). Статус: ${response.status}. Ответ: ${responseBodyText}`;
                         errorContainer.style.display = 'block';
                    }
                    if(progressContainer) progressContainer.style.display = 'none';
                    if(resultContainer) resultContainer.style.display = 'none';
                    return; // Прекращаем дальнейшую обработку
                }

                // Обновляем UI в зависимости от ответа
                if (response.ok) { // Код 200-299
                    if (data.success) {
                        if(progressBar) {
                            progressBar.style.width = '100%';
                            progressBar.textContent = '100%';
                        }
                        if(statusMessage) {
                            let statusText = data.message || 'Готово!';
                            if (typeof statusText === 'string' && statusText.includes('/') && statusText.startsWith('Обработка')) {
                                statusText = statusText.replace(/\//g, ' /\n'); 
                            }
                            statusMessage.textContent = statusText;
                        }
                        
                        await new Promise(resolve => setTimeout(resolve, 300)); 
                        
                        if(resultMessage) {
                            resultMessage.textContent = data.message || 'Обработка успешно завершена!';
                        }

                        if(downloadLink && data.download_url) {
                            downloadLink.href = data.download_url;
                            let buttonText = "Скачать результат";
                            if (data.download_filename) {
                               downloadLink.setAttribute('download', data.download_filename);
                               buttonText = `Скачать: ${escapeHtml(data.download_filename)}`;
                            }
                            downloadLink.textContent = buttonText;
                            downloadLink.style.display = 'inline-block'; // Показываем ссылку
                        } else if(downloadLink) { // Проверяем еще раз, что он есть, перед скрытием
                            downloadLink.style.display = 'none'; // Скрываем, если нет URL
                        }

                        if(resultContainer) {
                            resultContainer.style.display = 'block';
                        }

                        if(errorContainer) {
                            errorContainer.style.display = 'none';
                        }

                        if (reportDetailsContainer) {
                            let reportHtml = '';
                            if (data.empty_files && data.empty_files.length > 0) {
                                reportHtml += '<h4>⚠️ Пустые файлы (' + data.empty_files.length + '):</h4><ul>';
                                data.empty_files.forEach(filename => {
                                    reportHtml += '<li>' + escapeHtml(filename) + '</li>';
                                });
                                reportHtml += '</ul>';
                            }
                            if (data.failed_files && data.failed_files.length > 0) {
                                reportHtml += '<h4 style="margin-top:15px;">❌ Файлы с ошибками обработки (' + data.failed_files.length + '):</h4><ul>';
                                data.failed_files.forEach(filename => {
                                    reportHtml += '<li>' + escapeHtml(filename) + '</li>';
                                });
                                reportHtml += '</ul>';
                            }

                            if (reportHtml) {
                                reportDetailsContainer.innerHTML = reportHtml;
                                reportDetailsContainer.style.display = 'block';
                            } else {
                                reportDetailsContainer.style.display = 'none';
                            }
                        }
                    } else { // data.success === false (ошибка от бэкенда, но HTTP OK)
                        if(progressContainer) progressContainer.style.display = 'none';
                        if(errorContainer) {
                             errorContainer.innerHTML = `<b>Ошибка обработки:</b> ${escapeHtml(data.error || 'Неизвестная ошибка')}`;
                             errorContainer.style.display = 'block';
                        }
                        
                        if (reportDetailsContainer) {
                            let errorReportHtml = '';
                            if (data.empty_files && data.empty_files.length > 0) {
                                errorReportHtml += '<h4>⚠️ Обнаружены пустые файлы (' + data.empty_files.length + '):</h4><ul>';
                                data.empty_files.forEach(filename => {
                                    errorReportHtml += '<li>' + escapeHtml(filename) + '</li>';
                                });
                                errorReportHtml += '</ul>';
                            }
                            if (data.failed_files && data.failed_files.length > 0) {
                                errorReportHtml += '<h4 style="margin-top:15px;">❌ Файлы с ошибками обработки (' + data.failed_files.length + '):</h4><ul>';
                                data.failed_files.forEach(filename => {
                                    errorReportHtml += '<li>' + escapeHtml(filename) + '</li>';
                                });
                                errorReportHtml += '</ul>';
                            }

                            if (errorReportHtml) {
                                reportDetailsContainer.innerHTML = errorReportHtml;
                                reportDetailsContainer.style.display = 'block';
                            } else {
                                reportDetailsContainer.style.display = 'none';
                            }
                        }
                        if(resultContainer) resultContainer.style.display = 'none';
                    }
                } else { // HTTP ошибка (не 2xx)
                    if(progressContainer) progressContainer.style.display = 'none';
                    let errorMsg = `Ошибка сервера: ${response.status}`;
                    errorMsg += ` - ${responseBodyText}`;
                    if(errorContainer) {
                         errorContainer.textContent = errorMsg;
                         errorContainer.style.display = 'block';
                    }
                    if(resultContainer) resultContainer.style.display = 'none';
                }

            } catch (error) { // Сетевая ошибка при fetch('/upload') или другая ошибка до обработки ответа
                stopPolling();
                console.error('Ошибка при отправке/обработке ответа:', error); // Более общее сообщение
                if(progressContainer) progressContainer.style.display = 'none';
                if(errorContainer) {
                     errorContainer.textContent = `Произошла ошибка: ${error.message || error}`;
                     errorContainer.style.display = 'block';
                }
                if(resultContainer) resultContainer.style.display = 'none';
            } finally {
                const finalFileInput = document.getElementById('file');
                 if (finalFileInput) {
                    // Не очищаем файл, чтобы пользователь мог его скачать или отправить снова
                }
                checkFormValidity();
            }
        });
    } else {
        console.error("Форма с id 'upload-form' не найдена!");
    }
    checkFormValidity();
});