from flask import Flask, request, send_file, jsonify
from generate_ppt import crear_informe
import tempfile

app = Flask(__name__)

@app.route('/generar', methods=['POST'])
def generar():
    try:
        data = request.get_json()

        nombre = data.get('nombre')
        mes = data.get('mes')
        mp_prog = int(data.get('mp_programados', 0))
        mp_ejec = int(data.get('mp_ejecutados', 0))
        ord_finalizadas = int(data.get('ordenes_finalizadas', 0))
        ord_pendientes = int(data.get('ordenes_pendientes', 0))
        tema = data.get('tema', 'Clásico')

        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
        ruta_salida = temp_file.name

        crear_informe(nombre, mes, mp_prog, mp_ejec, ord_finalizadas, ord_pendientes, ruta_salida, tema)

        return send_file(ruta_salida, as_attachment=True, download_name="InformeMensual.pptx")

    except Exception as e:
        return jsonify({"error": str(e)}), 400

# ❌ Render no necesita esto, pero puedes dejarlo si desarrollas localmente
if __name__ == '__main__':
    app.run()
