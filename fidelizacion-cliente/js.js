var hojaSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
var TOTAL_SELLOS = 5;

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

function escaparHtml(valor) {
    return String(valor || '').replace(/[&<>"']/g, function (caracter) {
        return {
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#39;'
        }[caracter];
    });
}

function obtenerUrlTarjeta(urlTarjeta) {
    var url = String(urlTarjeta || '').trim();
    return /^https?:\/\/[^\s]+$/i.test(url) ? url : '';
}

function enviarCorreoBienvenida(cliente, urlTarjeta) {
    var nombre = String(cliente.nombre || '');
    var url = obtenerUrlTarjeta(urlTarjeta);
    var enlaceTexto = url
        ? 'Puedes acceder a tu tarjeta de fidelización digital en este enlace:\n' + url
        : 'El enlace público de la tarjeta aún no está configurado. Podrás acceder ingresando con tu DNI cuando More a la Moda te comparta la dirección de la aplicación.';
    var enlaceHtml = url
        ? '<p>👉 <a href="' + escaparHtml(url) + '">Ver mi tarjeta de fidelización e ingresar</a></p>'
        : '<p>El enlace público de la tarjeta aún no está configurado. Podrás acceder ingresando con tu DNI cuando More a la Moda te comparta la dirección de la aplicación.</p>';
    var asunto = '¡Bienvenida a More a la Moda! 💖 Tu tarjeta de fidelización ya está lista';
    var cuerpo = [
        '¡Hola, ' + nombre + '!',
        '',
        'Queremos agradecerte de corazón por registrarte en nuestra aplicación web y elegirnos. Sos parte de nuestra comunidad, y en More a la Moda nos encanta acompañarte en cada paso, por eso queremos premiar tu confianza y fidelidad.',
        '',
        'Desde ahora formas parte de nuestro exclusivo club de beneficios. Cada vez que realices tus compras, sumarás una estampa en tu tarjeta digital para estar más cerca de tu premio.',
        '',
        '¿Cómo funciona?',
        '- Colecciona 5 estampas: al completar tus primeros 5 sellos con tus compras, desbloquearás un 10% de descuento para utilizar en tu siguiente visita.',
        '- Beneficios especiales: accederás a sorpresas, descuentos y promociones exclusivas para el Día de la Madre y el Día de tu Cumpleaños.',
        '',
        enlaceTexto,
        '',
        'Gracias nuevamente por confiar en nosotras y por seguir eligiéndonos día a día. Si tienes alguna duda, estamos siempre a tu disposición.',
        '',
        'Con mucho cariño,',
        'El equipo de More a la Moda 👗✨'
    ].join('\n');
    var cuerpoHtml = '<p>¡Hola, ' + escaparHtml(nombre) + '!</p>' +
        '<p>Queremos agradecerte de corazón por registrarte en nuestra aplicación web y elegirnos. <strong>Sos parte de nuestra comunidad</strong>, y en <strong>More a la Moda</strong> nos encanta acompañarte en cada paso, por eso queremos premiar tu confianza y fidelidad.</p>' +
        '<p>Desde ahora formas parte de nuestro exclusivo club de beneficios. Cada vez que realices tus compras, sumarás una estampa en tu tarjeta digital para estar más cerca de tu premio.</p>' +
        '<p>✨ <strong>¿Cómo funciona?</strong></p>' +
        '<ul><li><strong>Colecciona 5 estampas:</strong> al completar tus primeros 5 sellos con tus compras, desbloquearás un <strong>10% de descuento</strong> para utilizar en tu siguiente visita.</li>' +
        '<li><strong>Beneficios especiales:</strong> accederás a sorpresas, descuentos y promociones exclusivas para el <strong>Día de la Madre</strong> y el <strong>Día de tu Cumpleaños</strong>.</li></ul>' +
        enlaceHtml +
        '<p>Gracias nuevamente por confiar en nosotras y por seguir eligiéndonos día a día. Si tienes alguna duda, estamos siempre a tu disposición.</p>' +
        '<p>Con mucho cariño,<br><strong>El equipo de More a la Moda</strong> 👗✨</p>';

    return enviarCorreo(cliente.mail, asunto, cuerpo, cuerpoHtml);
}

function enviarCorreoNuevaEstampa(cliente, totalSellos, urlTarjeta) {
    var nombre = String(cliente.nombre || '');
    var url = obtenerUrlTarjeta(urlTarjeta);
    var restantes = TOTAL_SELLOS - totalSellos;
    var estadoPremio = restantes > 0
        ? '¡Te faltan solo ' + restantes + ' para desbloquear tu 10% de descuento!'
        : '¡Completaste tus 5 estampas y desbloqueaste un 10% de descuento para tu siguiente visita!';
    var enlaceTexto = url
        ? 'Ingresa con tu DNI para consultar tu tarjeta: ' + url
        : 'El enlace público de la tarjeta aún no está configurado. Podrás consultar tu tarjeta cuando More a la Moda te comparta la dirección de la aplicación.';
    var enlaceHtml = url
        ? '<p>👉 <a href="' + escaparHtml(url) + '">Ver mi tarjeta de fidelización</a></p>'
        : '<p>El enlace público de la tarjeta aún no está configurado. Podrás consultar tu tarjeta cuando More a la Moda te comparta la dirección de la aplicación.</p>';
    var asunto = '¡Sumaste una nueva estampa! ✨ Gracias por seguir eligiéndonos, ' + nombre;
    var cuerpo = [
        '¡Hola, ' + nombre + '!',
        '',
        'Queremos darte las gracias por tu nueva visita y por seguir eligiéndonos. Para nosotras, sos parte fundamental de nuestra comunidad y nos encanta consentirte.',
        '',
        'Queremos contarte que acabamos de actualizar tu tarjeta digital. ¡Ya sumaste una nueva estampa! 🛍️💖',
        '',
        'Estado de tu tarjeta:',
        '- Tus estampas actuales: ' + totalSellos + ' de ' + TOTAL_SELLOS,
        '- Estampas restantes para tu premio: ' + estadoPremio,
        '',
        'Recuerda que puedes ingresar en cualquier momento para ver cómo va tu progreso y consultar tus beneficios exclusivos para el Día de la Madre y el Día de tu Cumpleaños.',
        enlaceTexto,
        '',
        'Gracias por confiar en More a la Moda y caminar junto a nosotras. ¡Te esperamos pronto!',
        '',
        'Con mucho cariño,',
        'El equipo de More a la Moda 👗✨'
    ].join('\n');
    var cuerpoHtml = '<p>¡Hola, ' + escaparHtml(nombre) + '!</p>' +
        '<p>Queremos darte las gracias por tu nueva visita y por seguir eligiéndonos. Para nosotras, <strong>sos parte fundamental de nuestra comunidad</strong> y nos encanta consentirte.</p>' +
        '<p>Queremos contarte que acabamos de actualizar tu tarjeta digital. ¡Ya sumaste una nueva estampa! 🛍️💖</p>' +
        '<p>✨ <strong>Estado de tu tarjeta:</strong></p>' +
        '<ul><li><strong>Tus estampas actuales:</strong> ' + totalSellos + ' de ' + TOTAL_SELLOS + '</li>' +
        '<li><strong>Estampas restantes para tu premio:</strong> ' + escaparHtml(estadoPremio) + '</li></ul>' +
        '<p>Recuerda que puedes ingresar en cualquier momento para ver cómo va tu progreso y consultar tus beneficios exclusivos para el Día de la Madre y el Día de tu Cumpleaños.</p>' +
        enlaceHtml +
        '<p>Gracias por confiar en <strong>More a la Moda</strong> y caminar junto a nosotras. ¡Te esperamos pronto!</p>' +
        '<p>Con mucho cariño,<br><strong>El equipo de More a la Moda</strong> 👗✨</p>';

    return enviarCorreo(cliente.mail, asunto, cuerpo, cuerpoHtml);
}

function enviarCorreo(destinatario, asunto, cuerpo, cuerpoHtml) {
    try {
        if (!destinatario) {
            throw new Error('El cliente no tiene un correo registrado.');
        }

        MailApp.sendEmail({
            to: destinatario,
            subject: asunto,
            body: cuerpo,
            htmlBody: cuerpoHtml
        });
        return true;
    } catch (error) {
        console.error('No se pudo enviar el correo:', error);
        return false;
    }
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
        mail: String(data.mail || '').trim(),
        sellos: [0, 0, 0, 0, 0]
    };
    var emailEnviado = enviarCorreoBienvenida(clienteObj, data.urlTarjeta);

    return ContentService.createTextOutput(JSON.stringify({
        success: true,
        cliente: clienteObj,
        emailEnviado: emailEnviado,
        enlaceIncluido: Boolean(obtenerUrlTarjeta(data.urlTarjeta))
    })).setMimeType(ContentService.MimeType.JSON);
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

    if (!dni) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'DNI inválido.' })).setMimeType(ContentService.MimeType.JSON);
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
        var fila = buscarFilaPorDni(dni);

        if (!fila) {
            return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Cliente no encontrado' })).setMimeType(ContentService.MimeType.JSON);
        }

        var row = hojaSpreadsheet.getRange(fila, 1, 1, 12).getValues()[0];
        var cliente = obtenerClienteDesdeFila(row);
        var siguienteSello = cliente.sellos.findIndex(function (sello) {
            return sello === 0;
        });

        if (siguienteSello === -1) {
            return ContentService.createTextOutput(JSON.stringify({ success: true, completo: true, estampaAgregada: false, cliente: cliente })).setMimeType(ContentService.MimeType.JSON);
        }

        var columnaSello = 8 + siguienteSello;
        hojaSpreadsheet.getRange(fila, columnaSello).setValue(1);
        cliente.sellos[siguienteSello] = 1;
        var totalSellos = cliente.sellos.reduce(function (total, sello) {
            return total + sello;
        }, 0);
    } finally {
        lock.releaseLock();
    }

    var emailEnviado = enviarCorreoNuevaEstampa(cliente, totalSellos, data.urlTarjeta);
    return ContentService.createTextOutput(JSON.stringify({
        success: true,
        completo: totalSellos === TOTAL_SELLOS,
        estampaAgregada: true,
        emailEnviado: emailEnviado,
        enlaceIncluido: Boolean(obtenerUrlTarjeta(data.urlTarjeta)),
        cliente: cliente
    })).setMimeType(ContentService.MimeType.JSON);
}