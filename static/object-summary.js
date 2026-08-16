(function () {
    'use strict';

    const openButton = document.getElementById('object-summary-open');
    const modal = document.getElementById('object-summary-modal');
    if (!openButton || !modal) return;

    const contractorSelect = document.getElementById('summary-contractor');
    const objectSelect = document.getElementById('summary-object');
    const systemSelect = document.getElementById('summary-system');
    const result = document.getElementById('object-summary-result');
    const message = document.getElementById('summary-message');
    let closingTimer = null;

    function setOptions(select, values, placeholder) {
        select.replaceChildren();
        const empty = document.createElement('option');
        empty.value = '';
        empty.textContent = placeholder;
        select.appendChild(empty);
        values.forEach(function (value) {
            const option = document.createElement('option');
            option.value = value;
            option.textContent = value;
            select.appendChild(option);
        });
        select.disabled = values.length === 0;
    }

    function showMessage(text, isError) {
        message.textContent = text;
        message.classList.toggle('is-error', Boolean(isError));
        message.hidden = false;
        result.hidden = true;
    }

    async function fetchJson(url) {
        const response = await fetch(url, {
            headers: {'Accept': 'application/json'},
            credentials: 'same-origin'
        });
        const payload = await response.json().catch(function () { return {}; });
        if (!response.ok) throw new Error(payload.error || 'Не удалось загрузить данные.');
        return payload;
    }

    async function loadOptions(contractor, objectAddress) {
        const params = new URLSearchParams();
        if (contractor) params.set('contractor', contractor);
        if (objectAddress) params.set('object', objectAddress);
        return fetchJson('/api/object-summary/options?' + params.toString());
    }

    async function openModal() {
        if (closingTimer) window.clearTimeout(closingTimer);
        modal.hidden = false;
        document.body.classList.add('modal-open');
        window.requestAnimationFrame(function () { modal.classList.add('is-open'); });
        showMessage('Загружаем актуальные данные…', false);
        try {
            const options = await loadOptions('', '');
            setOptions(contractorSelect, options.contractors, 'Выберите подрядчика');
            setOptions(objectSelect, [], 'Сначала выберите подрядчика');
            setOptions(systemSelect, [], 'Сначала выберите объект');
            showMessage('Выберите параметры, чтобы увидеть актуальную сводку.', false);
            contractorSelect.focus();
        } catch (error) {
            showMessage(error.message, true);
        }
    }

    function closeModal() {
        modal.classList.remove('is-open');
        document.body.classList.remove('modal-open');
        closingTimer = window.setTimeout(function () {
            modal.hidden = true;
            openButton.focus();
        }, 180);
    }

    contractorSelect.addEventListener('change', async function () {
        result.hidden = true;
        setOptions(objectSelect, [], 'Загружаем объекты…');
        setOptions(systemSelect, [], 'Сначала выберите объект');
        if (!contractorSelect.value) {
            showMessage('Выберите подрядчика.', false);
            return;
        }
        showMessage('Загружаем объекты…', false);
        try {
            const options = await loadOptions(contractorSelect.value, '');
            setOptions(objectSelect, options.objects, 'Выберите объект');
            showMessage(options.objects.length ? 'Теперь выберите объект.' : 'У подрядчика пока нет объектов.', false);
        } catch (error) {
            showMessage(error.message, true);
        }
    });

    objectSelect.addEventListener('change', async function () {
        result.hidden = true;
        setOptions(systemSelect, [], 'Загружаем системы…');
        if (!objectSelect.value) {
            showMessage('Выберите объект.', false);
            return;
        }
        showMessage('Загружаем инженерные системы…', false);
        try {
            const options = await loadOptions(contractorSelect.value, objectSelect.value);
            setOptions(systemSelect, options.systems, 'Выберите систему');
            showMessage(options.systems.length ? 'Выберите инженерную систему.' : 'Для объекта нет доступных систем.', false);
        } catch (error) {
            showMessage(error.message, true);
        }
    });

    systemSelect.addEventListener('change', async function () {
        if (!systemSelect.value) {
            showMessage('Выберите инженерную систему.', false);
            return;
        }
        showMessage('Формируем актуальную сводку…', false);
        const params = new URLSearchParams({
            contractor: contractorSelect.value,
            object: objectSelect.value,
            system: systemSelect.value
        });
        try {
            const summary = await fetchJson('/api/object-summary?' + params.toString());
            document.getElementById('summary-address').textContent = summary.object;
            document.getElementById('summary-context').textContent = summary.contractor + ' · ' + summary.system;
            document.getElementById('summary-progress').textContent = summary.progress + '%';
            document.getElementById('summary-progress-fill').style.width = summary.progress + '%';
            document.getElementById('summary-total').textContent = summary.total;
            document.getElementById('summary-done').textContent = summary.done;
            document.getElementById('summary-check').textContent = summary.check;
            document.getElementById('summary-denied').textContent = summary.denied;
            document.getElementById('summary-notdone').textContent = summary.notdone;

            const spill = document.getElementById('summary-spill');
            spill.replaceChildren();
            const title = document.createElement('span');
            title.textContent = 'Розлив';
            spill.appendChild(title);
            if (summary.system === 'ТС') {
                const upper = document.createElement('strong');
                upper.textContent = 'Верхний ' + summary.upper_rozliv + '%';
                const lower = document.createElement('strong');
                lower.textContent = 'Нижний ' + summary.lower_rozliv + '%';
                spill.append(upper, lower);
            } else {
                const single = document.createElement('strong');
                single.textContent = summary.rozliv + '%';
                spill.appendChild(single);
            }

            message.hidden = true;
            result.hidden = false;
        } catch (error) {
            showMessage(error.message, true);
        }
    });

    openButton.addEventListener('click', openModal);
    modal.querySelectorAll('[data-close-object-summary]').forEach(function (button) {
        button.addEventListener('click', closeModal);
    });
    document.addEventListener('keydown', function (event) {
        if (event.key === 'Escape' && !modal.hidden) closeModal();
    });
}());
