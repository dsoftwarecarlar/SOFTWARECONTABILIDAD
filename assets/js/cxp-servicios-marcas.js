(() => {
    const form = document.querySelector('[data-job-form]');
    const submitButton = document.querySelector('[data-submit-button]');
    if (form && submitButton) {
        form.addEventListener('submit', () => {
            submitButton.disabled = true;
            submitButton.textContent = submitButton.getAttribute('data-processing-label') || 'Procesando...';
        });
    }

    const stopForm = document.querySelector('[data-stop-form]');
    const stopButton = document.querySelector('[data-stop-button]');
    if (stopForm && stopButton && !stopButton.disabled) {
        stopForm.addEventListener('submit', (event) => {
            const confirmMessage = stopButton.getAttribute('data-confirm-message') || 'Se solicitara detener el proceso.';
            if (!window.confirm(confirmMessage)) {
                event.preventDefault();
                return;
            }

            stopButton.disabled = true;
            stopButton.textContent = stopButton.getAttribute('data-processing-label') || 'Procesando...';
        });
    }

    const jobBox = document.querySelector('[data-job-box]');
    if (!jobBox) {
        return;
    }

    const statusUrl = jobBox.getAttribute('data-status-url');
    const returnUrl = jobBox.getAttribute('data-return-url');
    const messageNode = jobBox.querySelector('[data-job-message]');
    const metaNode = jobBox.querySelector('[data-job-meta]');
    if (!statusUrl || !returnUrl || !messageNode || !metaNode) {
        return;
    }

    const intervalMs = Number(jobBox.getAttribute('data-poll-interval') || '4000');
    const initialDelayMs = Number(jobBox.getAttribute('data-poll-delay') || '2000');
    const reconnectingMessage = jobBox.getAttribute('data-reconnecting-message') || 'Estado: reconectando con el proceso...';

    const renderMeta = (job) => {
        const status = job.status || 'running';
        const stamp = job.started_at || job.created_at || job.updated_at || '';
        metaNode.textContent = stamp ? `Estado: ${status} - ${stamp}` : `Estado: ${status}`;
    };

    const poll = async () => {
        try {
            const response = await fetch(statusUrl, { cache: 'no-store' });
            if (!response.ok) {
                throw new Error(jobBox.getAttribute('data-status-error') || 'No se pudo consultar el estado del proceso.');
            }

            const job = await response.json();
            messageNode.textContent = job.message || 'Procesando en segundo plano.';
            renderMeta(job);

            if (job.status === 'complete' || job.status === 'error' || job.status === 'cancelled') {
                window.location.href = returnUrl;
                return;
            }
        } catch (_error) {
            metaNode.textContent = reconnectingMessage;
        }

        window.setTimeout(poll, intervalMs);
    };

    window.setTimeout(poll, initialDelayMs);
})();
