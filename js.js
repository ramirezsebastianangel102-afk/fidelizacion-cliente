var hojaSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

function inicializarHoja() {
    var headers = ['NumeroCliente', 'Nombre', 'Apellido', 'DNI', 'Nacimiento', 'Telefono', 'Mail', 'Sello1', 'Sello2', 'Sello3', 'Sello4', 'Sello5'];
    if (hojaSpreadsheet.getLastRow() === 0) {
        hojaSpreadsheet.appendRow(headers);
    }
}

function obtenerClienteDesdeFila(row) {
    return {
        numeroCliente: row[0],
        nombre: row[1],
        apellido: row[2],
        dni: row[3],
        nacimiento: row[4],
        telefono: row[5],
        mail: row[6],
        sellos: [
            convertirSello(row[7]),
            convertirSello(row[8]),
            convertirSello(row[9]),
            convertirSello(row[10]),
            convertirSello(row[11])
        ]
    };
}

function convertirSello(valor) {
    return valor === true || String(valor).toLowerCase() === 'true' || Number(valor) === 1 ? 1 : 0;
}

function buscarFilaPorDni(dni) {
    var rows = hojaSpreadsheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][3]).trim() === String(dni).trim()) {
            return i + 1;
        }
    }
    return null;
}

function buscarNumeroClienteUnico() {
    var numero = String(Math.floor(100000 + Math.random() * 900000));
    var rows = hojaSpreadsheet.getDataRange().getValues();

    for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === numero) {
            return buscarNumeroClienteUnico();
        }
    }

    return numero;
}

function doPost(e) {
    inicializarHoja();

    var data = {};
    try {
        data = JSON.parse(e.postData.contents);
    } catch (error) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'JSON inválido' })).setMimeType(ContentService.MimeType.JSON);
    }

    var accion = data.accion;

    if (accion === 'registrar') {
        return registrarCliente(data);
    }

    if (accion === 'actualizarSello') {
        return actualizarSelloCliente(data);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Acción no válida' })).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
    inicializarHoja();
    var accion = e.parameter.accion;

    if (accion === 'consultar') {
        return consultarClientePorDni(e.parameter.dni);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Acción no válida' })).setMimeType(ContentService.MimeType.JSON);
}

function registrarCliente(data) {
    var nombre = String(data.nombre || '').trim();
    var apellido = String(data.apellido || '').trim();
    var dni = String(data.dni || '').trim();

    if (!nombre || !apellido || !dni) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Faltan datos obligatorios.' })).setMimeType(ContentService.MimeType.JSON);
    }

    if (buscarFilaPorDni(dni)) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'El DNI ya se encuentra registrado.' })).setMimeType(ContentService.MimeType.JSON);
    }

    var numCliente = buscarNumeroClienteUnico();

    hojaSpreadsheet.appendRow([
        numCliente,
        nombre,
        apellido,
        dni,
        data.nacimiento || '',
        data.telefono || '',
        data.mail || '',
        0, 0, 0, 0, 0
    ]);

    var clienteObj = {
        numeroCliente: numCliente,
        nombre: nombre,
        apellido: apellido,
        dni: dni,
        sellos: [0, 0, 0, 0, 0]
    };

    return ContentService.createTextOutput(JSON.stringify({ success: true, cliente: clienteObj })).setMimeType(ContentService.MimeType.JSON);
}

function consultarClientePorDni(dni) {
    var fila = buscarFilaPorDni(dni);

    if (!fila) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Cliente no encontrado' })).setMimeType(ContentService.MimeType.JSON);
    }

    var row = hojaSpreadsheet.getRange(fila, 1, 1, 12).getValues()[0];
    var cliente = obtenerClienteDesdeFila(row);

    return ContentService.createTextOutput(JSON.stringify({ success: true, cliente: cliente })).setMimeType(ContentService.MimeType.JSON);
}

function actualizarSelloCliente(data) {
    var dni = String(data.dni || '').trim();
    var selloNum = Number(data.selloNum);

    if (!dni || !Number.isInteger(selloNum) || selloNum < 1 || selloNum > 5) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Sello inválido.' })).setMimeType(ContentService.MimeType.JSON);
    }

    var fila = buscarFilaPorDni(dni);

    if (!fila) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Cliente no encontrado' })).setMimeType(ContentService.MimeType.JSON);
    }

    var columnaSello = 8 + (selloNum - 1);
    hojaSpreadsheet.getRange(fila, columnaSello).setValue(1);

    var row = hojaSpreadsheet.getRange(fila, 1, 1, 12).getValues()[0];
    var cliente = obtenerClienteDesdeFila(row);

    return ContentService.createTextOutput(JSON.stringify({ success: true, cliente: cliente })).setMimeType(ContentService.MimeType.JSON);
}