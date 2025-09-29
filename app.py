from flask import Flask, render_template, request
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import io
import base64
from sklearn.cluster import KMeans

app = Flask(__name__)

@app.route('/auditoria_atenuacion', methods=['GET', 'POST'])
def index():
    plot_url = None
    tabla_html = None
    olt_name = ""

    if request.method == 'POST':
        file = request.files['file']
        olt_name = request.form.get('olt_name')
        db_threshold = -30
        df = pd.read_csv(file)

        # Filtrar por OLT
        df_olt = df[df['OLT'] == olt_name]

        # Agrupar por Board y Port: promedio señal y cantidad abonados críticos
        resumen = df_olt.groupby(['Board', 'Port']).agg(
            abonados_criticos = ('Signal 1310', lambda x: (x >= db_threshold).sum()),
            prom_signal = ('Signal 1310', 'mean'),
        ).reset_index()

        # Modelo KMeans para clasificar estados en 3 clusters
        X = resumen[['prom_signal', 'abonados_criticos']].values
        kmeans = KMeans(n_clusters=3, random_state=42)
        clusters = kmeans.fit_predict(X)
        resumen['cluster'] = clusters

        # Mapear números a etiquetas
        etiquetas = {0: 'Normal', 1: 'Alerta', 2: 'Crítico'}
        resumen['clasificacion'] = resumen['cluster'].map(etiquetas)

        # Crear heatmap de promedio señal
        tabla = resumen.pivot(index="Board", columns="Port", values="prom_signal")
        plt.figure(figsize=(10, 7))
        sns.heatmap(tabla, annot=True, cmap="coolwarm", fmt=".1f")
        plt.title(f"Mapa de Calor (Promedio Señal) OLT {olt_name}")

        img = io.BytesIO()
        plt.savefig(img, format='png')
        img.seek(0)
        plot_url = base64.b64encode(img.getvalue()).decode()
        plt.close()

        # Convertir resumen con clasificación a tabla HTML para mostrar
        tabla_html = resumen.to_html(classes='table table-striped', index=False)

    return render_template('auditoria_atenuacion.html', plot_url=plot_url, tabla_html=tabla_html, olt_name=olt_name)

if __name__ == "__main__":
    app.run(debug=True)
