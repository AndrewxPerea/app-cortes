// Mostrar el loader cuando se envía un formulario
document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('resultForm');
    if (form) {
        form.addEventListener('submit', function(e) {
            document.getElementById('loader').style.display = 'block';
            document.getElementById('overlay').style.display = 'block';
        });
    }
});
