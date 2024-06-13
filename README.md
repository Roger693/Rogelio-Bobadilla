# Rogelio-Bobadilla
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Base de Datos de Trámites</title>
    <style>
        table {
            width: 100%;
            border-collapse: collapse;
        }
        th, td {
            padding: 10px;
            text-align: left;
            border: 1px solid #dddddd;
        }
        th {
            background-color: #f2f2f2;
        }
        .warning {
            color: red;
            font-weight: bold;
        }
        form {
            margin-bottom: 20px;
        }
        form div {
            margin-bottom: 10px;
        }
        label {
            display: inline-block;
            width: 150px;
        }
    </style>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.0/xlsx.full.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/axios/dist/axios.min.js"></script>
</head>
<body>
    <h1>Base de Datos de Trámites</h1>

    <form id="data-form">
        <div>
            <label for="curp">CURP:</label>
            <input type="text" id="curp" name="curp">
        </div>
        <div>
            <label for="nombre">Nombre:</label>
            <input type="text" id="nombre" name="nombre">
        </div>
        <div>
            <label for="tramite">Trámite:</label>
            <select id="tramite" name="tramite" required>
                <option value="RFC">RFC - $35.00</option>
                <option value="Acta de Nacimiento">Acta de Nacimiento - $120.00</option>
                <option value="ANP EdoMex">ANP EdoMex - $40.00</option>
                <option value="ANP Federal">ANP Federal - $350.00</option>
            </select>
        </div>
        <div>
            <label for="solicitadoPor">Solicitado por:</label>
            <input type="text" id="solicitadoPor" name="solicitadoPor" required>
        </div>
        <button type="submit">Agregar</button>
    </form>

    <button onclick="downloadExcel()">Descargar Informe en Excel</button>

    <table>
        <thead>
            <tr>
                <th>No.</th>
                <th>CURP</th>
                <th>Nombre</th>
                <th>Trámite</th>
                <th>Solicitado por</th>
                <th>Fecha</th>
                <th>Acciones</th>
            </tr>
        </thead>
        <tbody id="data-table">
            <!-- Las filas de la tabla serán agregadas aquí dinámicamente -->
        </tbody>
        <tfoot>
            <tr>
                <td colspan="3"></td>
                <td>Total:</td>
                <td id="total-amount">$0.00</td>
                <td colspan="2"></td>
            </tr>
        </tfoot>
    </table>

    <script>
        const data = [];
        const prices = {
            "RFC": 35.00,
            "Acta de Nacimiento": 120.00,
            "ANP EdoMex": 40.00,
            "ANP Federal": 350.00
        };

        const curpSet = new Set();
        const nombreSet = new Set();
        let totalAmount = 0;

        async function getToken() {
            // Aquí debes obtener el token de acceso utilizando OAuth2.0
            // Para simplificar, este es un ejemplo simplificado que asume que ya tienes el token
            const token = 'YOUR_ACCESS_TOKEN';
            return token;
        }

        async function uploadToOneDrive(data) {
            const token = await getToken();

            const ws = XLSX.utils.json_to_sheet(data);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Tramites");

            const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

            const blob = new Blob([wbout], { type: 'application/octet-stream' });

            const formData = new FormData();
            formData.append('file', blob, 'informe_tramites.xlsx');

            const url = 'https://graph.microsoft.com/v1.0/me/drive/items/root:/informe_tramites.xlsx:/content';

            try {
                const response = await axios.put(url, blob, {
                    headers: {
                        Authorization: `Bearer ${token}`,
                        'Content-Type': 'application/octet-stream'
                    }
                });
                console.log('Archivo subido:', response.data);
            } catch (error) {
                console.error('Error subiendo archivo:', error);
            }
        }

        function addRow(no, curp, nombre, tramite, solicitadoPor, fecha) {
            const row = document.createElement('tr');
            const cells = [no, curp, nombre, tramite, solicitadoPor, fecha];
            cells.forEach(value => {
                const cell = document.createElement('td');
                cell.textContent = value;
                row.appendChild(cell);
            });

            const accionesCell = document.createElement('td');
            const deleteButton = document.createElement('button');
            deleteButton.textContent = 'Eliminar';
            deleteButton.onclick = function() {
                eliminarEntrada(row, curp, nombre, tramite);
            };
            accionesCell.appendChild(deleteButton);
            row.appendChild(accionesCell);

            if (curpSet.has(curp) || nombreSet.has(nombre)) {
                const warningCell = document.createElement('td');
                warningCell.colSpan = 2;
                warningCell.className = 'warning';
                warningCell.textContent = 'Este servicio ya ha sido realizado';
                row.appendChild(warningCell);
            } else {
                curpSet.add(curp);
                nombreSet.add(nombre);
                totalAmount += prices[tramite];
                document.getElementById('total-amount').textContent = `$${totalAmount.toFixed(2)}`;
            }

            document.getElementById('data-table').appendChild(row);
            uploadToOneDrive(data); // Subir a OneDrive cada vez que se agrega un nuevo dato
        }

        function eliminarEntrada(row, curp, nombre, tramite) {
            row.remove();
            if (curpSet.has(curp)) {
                curpSet.delete(curp);
                totalAmount -= prices[tramite];
            }
            if (nombreSet.has(nombre)) {
                nombreSet.delete(nombre);
                totalAmount -= prices[tramite];
            }
            document.getElementById('total-amount').textContent = `$${totalAmount.toFixed(2)}`;
            uploadToOneDrive(data); // Subir a OneDrive cada vez que se elimina un dato
        }

        document.getElementById('data-form').addEventListener('submit', function(event) {
            event.preventDefault();

            const curp = event.target.curp.value;
            const nombre = event.target.nombre.value;
            const tramite = event.target.tramite.value;
            const solicitadoPor = event.target.solicitadoPor.value;
            const fecha = new Date().toLocaleDateString();

            const no = data.length + 1;

            data.push({ no, curp, nombre, tramite, solicitadoPor, fecha });
            addRow(no, curp, nombre, tramite, solicitadoPor, fecha);

            event.target.reset();
            document.getElementById('curp').disabled = false;
            document.getElementById('nombre').disabled = false;
        });

        document.getElementById('curp').addEventListener('input', function() {
            if (this.value.trim() !== '') {
                document.getElementById('nombre').disabled = true;
            } else {
                document.getElementById('nombre').disabled = false;
            }
        });

        document.getElementById('nombre').addEventListener('input', function() {
            if (this.value.trim() !== '') {
                document.getElementById('curp').disabled = true;
            }

