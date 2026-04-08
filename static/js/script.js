// Mostrar el loader cuando se envía un formulario
document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('resultForm');
    const loader = document.getElementById('loader');
    const overlay = document.getElementById('overlay');

    function mostrarLoader() {
        if (loader) {
            loader.style.display = 'block';
        }
        if (overlay) {
            overlay.style.display = 'block';
        }
    }

    function ocultarLoader() {
        if (loader) {
            loader.style.display = 'none';
        }
        if (overlay) {
            overlay.style.display = 'none';
        }
    }

    function obtenerNombreDescarga(response, form) {
        const disposition = response.headers.get('Content-Disposition') || '';
        const matchUtf = disposition.match(/filename\*=UTF-8''([^;]+)/i);
        if (matchUtf && matchUtf[1]) {
            return decodeURIComponent(matchUtf[1]);
        }

        const match = disposition.match(/filename=\"?([^\";]+)\"?/i);
        if (match && match[1]) {
            return match[1];
        }

        return form.dataset.downloadName || 'resultado.xlsx';
    }

    if (form) {
        form.addEventListener('submit', async function(e) {
            if (form.dataset.downloadDirect === 'true') {
                e.preventDefault();
                mostrarLoader();

                try {
                    const response = await fetch(form.action, {
                        method: 'POST',
                        body: new FormData(form),
                        credentials: 'same-origin',
                    });
                    const contentType = (response.headers.get('Content-Type') || '').toLowerCase();

                    if (contentType.includes('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')) {
                        const blob = await response.blob();
                        const url = window.URL.createObjectURL(blob);
                        const link = document.createElement('a');
                        link.href = url;
                        link.download = obtenerNombreDescarga(response, form);
                        document.body.appendChild(link);
                        link.click();
                        link.remove();
                        window.URL.revokeObjectURL(url);
                        return;
                    }

                    const html = await response.text();
                    document.open();
                    document.write(html);
                    document.close();
                } catch (error) {
                    ocultarLoader();
                    alert('No fue posible completar la descarga. Intenta nuevamente.');
                } finally {
                    ocultarLoader();
                }
                return;
            }

            mostrarLoader();
        });
    }
});
