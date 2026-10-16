# Control remoto de la charla desde el celular

Pasa las láminas y muestra las **notas del orador de la lámina actual** en el celular de cada ponente.
Las notas se leen en vivo desde PowerPoint: si las editas, el celular se actualiza solo.

## El día de la charla (5 pasos)

1. **Hotspot:** un celular activa "Compartir internet" (datos). El portátil y el **otro celular** se conectan a esa red.
2. **PowerPoint:** abre tu presentación `.pptx` en el portátil, con las notas del orador ya escritas en cada lámina. No hace falta iniciar la presentación.
3. **Servidor:** doble clic en `iniciar.bat`. Muestra un PIN, los enlaces y un QR.
   - Si Windows pregunta por el firewall: **Permitir** (red privada y pública).
4. **Celulares:** abrir el enlace con la IP del hotspot (`172.20.10.x` en iPhone, `192.168.43.x` en Android) o escanear el QR.
   - Si hay dos IP, usa la que empieza igual que la del celular que comparte datos.
5. En el celular: **▶ Iniciar la presentación** (o inícienla en el portátil con F5). Ya se puede pasar con **Siguiente / Anterior**.

## Qué trae la pantalla del celular
- Lámina actual / total, título y **notas completas** (A− / A+ cambia el tamaño; desplazables).
- Cronómetro compartido (amarillo a los 25 min, rojo a los 30). ⟲ lo reinicia.
- Punto verde = conectado; rojo = sin conexión.
- Si los dos tocan a la vez, solo cuenta el primer toque (no se salta una lámina).
- En la última lámina, "Siguiente" no cierra la presentación.

## Respaldo
Lleven un **clicker USB/Bluetooth** o usen las flechas del teclado. El control remoto no sustituye al portátil: solo manda órdenes a PowerPoint.

## Si algo falla
| Síntoma | Qué hacer |
|---|---|
| El celular no abre la página | Revisar que esté en el mismo hotspot; probar la otra IP; permitir Python en el firewall |
| "PIN incorrecto" | Abrir el enlace completo (con `?k=`) o escanear el QR |
| "PowerPoint no está abierto" | Abrir el archivo y volver a tocar |
| Se queda "Sin conexión…" | El portátil cambió de red o se suspendió: reiniciar `iniciar.bat` |

Opciones: `python server.py --port 8765 --pin 1234` fija el PIN y el puerto.
