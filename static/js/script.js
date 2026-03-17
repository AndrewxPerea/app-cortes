// Mostrar el loader cuando se envía un formulario
document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('resultForm');
    const loader = document.getElementById('loader');
    const overlay = document.getElementById('overlay');
    if (form) {
        form.addEventListener('submit', function(e) {
            if (loader) {
                loader.style.display = 'block';
            }
            if (overlay) {
                overlay.style.display = 'block';
            }
        });
    }
});
