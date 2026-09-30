var ID_SPREADSHEET = 'REEMPLAZAR_CON_EL_ID_DEL_GOOGLE_SHEET_OFICIAL';
var libroSpreadsheet;
var hojaSpreadsheet;
var TOTAL_SELLOS = 5;
var ENCABEZADO_DESCUENTOS = 'Cantidad de veces que accedio a un descuento';
var ENCABEZADO_ULTIMA_COMPRA = 'Ultima fecha de compra';

function inicializarHoja() {
    if (ID_SPREADSHEET === 'REEMPLAZAR_CON_EL_ID_DEL_GOOGLE_SHEET_OFICIAL') {
        throw new Error('Configura el ID del Google Sheet oficial en ID_SPREADSHEET.');
    }

    libroSpreadsheet = SpreadsheetApp.openById(ID_SPREADSHEET);
    hojaSpreadsheet = libroSpreadsheet.getSheetByName('Clientes');
    if (!hojaSpreadsheet) {
        throw new Error('No existe la pestaña Clientes en el Google Sheet configurado.');
    }

    var headers = ['NumeroCliente', 'Nombre', 'Apellido', 'DNI', 'Nacimiento', 'Telefono', 'Mail', 'Sello1', 'Sello2', 'Sello3', 'Sello4', 'Sello5'];
    if (hojaSpreadsheet.getLastRow() === 0) {
        hojaSpreadsheet.appendRow(headers);
    }
    asegurarColumnasBeneficio();
}

function normalizarEncabezado(valor) {
    return String(valor || '').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function asegurarColumnasBeneficio() {
    var encabezados = hojaSpreadsheet.getRange(1, 1, 1, hojaSpreadsheet.getLastColumn()).getValues()[0];
    [ENCABEZADO_DESCUENTOS, ENCABEZADO_ULTIMA_COMPRA].forEach(function (encabezado) {
        var buscado = normalizarEncabezado(encabezado);
        var existe = encabezados.some(function (valor) {
            return normalizarEncabezado(valor) === buscado;
        });
        if (!existe) {
            var columna = encabezados.length + 1;
            hojaSpreadsheet.getRange(1, columna).setValue(encabezado);
            encabezados.push(encabezado);
        }
    });
}

function obtenerColumnasBeneficio() {
    var encabezados = hojaSpreadsheet.getRange(1, 1, 1, hojaSpreadsheet.getLastColumn()).getValues()[0];
    var columnas = {};
    encabezados.forEach(function (valor, indice) {
        var normalizado = normalizarEncabezado(valor);
        if (normalizado === normalizarEncabezado(ENCABEZADO_DESCUENTOS)) {
            columnas.descuentos = indice + 1;
        }
        if (normalizado === normalizarEncabezado(ENCABEZADO_ULTIMA_COMPRA)) {
            columnas.ultimaCompra = indice + 1;
        }
    });
    return columnas;
}

function obtenerClienteDesdeFila(row) {
    var columnas = obtenerColumnasBeneficio();
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
        ],
        descuentosObtenidos: Number(row[columnas.descuentos - 1]) || 0,
        ultimaCompra: row[columnas.ultimaCompra - 1] || ''
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

function enviarCorreoDescuentoDesbloqueado(cliente, urlTarjeta) {
    var nombre = String(cliente.nombre || '');
    var estampasCompletadas = Number(cliente.descuentosObtenidos) > 1 ? 'otras 5 estampas' : 'tus primeras 5 estampas';
    var url = obtenerUrlTarjeta(urlTarjeta);
    var enlaceTexto = url
        ? 'Ver mi tarjeta y usar mi beneficio: ' + url
        : 'El enlace público de la tarjeta todavía no está configurado. Puedes ingresar con tu DNI cuando More a la Moda te comparta la dirección de la aplicación.';
    var enlaceHtml = url
        ? '<p>👉 <a href="' + escaparHtml(url) + '"><strong>Ver mi tarjeta y usar mi beneficio</strong></a></p>'
        : '<p>El enlace público de la tarjeta todavía no está configurado. Puedes ingresar con tu DNI cuando More a la Moda te comparta la dirección de la aplicación.</p>';
    var asunto = '¡Felicitaciones! 🎉 Completaste tus 5 estampas en More a la Moda (¡Tienes un 10% de descuento esperando!)';
    var cuerpo = [
        '¡Hola, ' + nombre + '!',
        '',
        'Tenemos una gran noticia para darte: ¡completaste ' + estampasCompletadas + ' en tu tarjeta digital! 💖',
        '',
        'Queremos agradecerte de corazón por ser una parte tan importante de nuestra comunidad y por elegirnos en cada ocasión. Como premio a tu fidelidad, en tu próxima compra tienes un 10% de descuento para llevarte lo que más te guste.',
        '',
        'Detalles de tu beneficio:',
        '- Descuento: 10% en tu compra.',
        '- Tope de reintegro: aplica con un tope máximo de $10.000.',
        enlaceTexto,
        '',
        '¡Te esperamos para celebrar juntas y renovar tus looks favoritos! Gracias por confiar en nosotras.',
        '',
        'Con mucho cariño,',
        'El equipo de More a la Moda 👗✨'
    ].join('\n');
    var cuerpoHtml = '<p>¡Hola, ' + escaparHtml(nombre) + '!</p>' +
        '<p>Tenemos una gran noticia para darte: ¡completaste ' + escaparHtml(estampasCompletadas) + ' en tu tarjeta digital! 💖</p>' +
        '<p>Queremos agradecerte de corazón por ser una parte tan importante de nuestra comunidad y por elegirnos en cada ocasión. Como premio a tu fidelidad, <strong>en tu próxima compra tienes un 10% de descuento</strong> para llevarte lo que más te guste.</p>' +
        '<p>✨ <strong>Detalles de tu beneficio:</strong></p>' +
        '<ul><li><strong>Descuento:</strong> 10% en tu compra.</li>' +
        '<li><strong>Tope de reintegro:</strong> aplica con un tope máximo de $10.000.</li></ul>' +
        enlaceHtml +
        '<p>¡Te esperamos para celebrar juntas y renovar tus looks favoritos! Gracias por confiar en nosotras.</p>' +
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

function buscarFilaPorNumeroCliente(numeroCliente) {
    var rows = hojaSpreadsheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === String(numeroCliente).trim()) {
            return i + 1;
        }
    }
    return null;
}

function obtenerIndiceColumnaPorAlias(encabezados, aliases) {
    var aliasNormalizados = aliases.map(function (alias) {
        return normalizarEncabezado(alias);
    });

    for (var i = 0; i < encabezados.length; i++) {
        var encabezadoNormalizado = normalizarEncabezado(encabezados[i]);
        if (aliasNormalizados.indexOf(encabezadoNormalizado) !== -1) {
            return i;
        }
    }

    return -1;
}

function validarAdmin(dni, clave) {
    var hojaAdministradores = libroSpreadsheet.getSheetByName('Administradores');
    if (!hojaAdministradores || hojaAdministradores.getLastRow() < 2) {
        return false;
    }

    var filas = hojaAdministradores.getDataRange().getValues();
    var encabezados = filas[0].map(normalizarEncabezado);
    var columnaDni = obtenerIndiceColumnaPorAlias(encabezados, ['dni', 'documento', 'numero documento', 'nro documento']);
    if (columnaDni < 0) {
        columnaDni = 0;
    }
    var columnaRol = obtenerIndiceColumnaPorAlias(encabezados, ['rol', 'role', 'tipo', 'perfil']);
    var columnaClave = obtenerIndiceColumnaPorAlias(encabezados, ['contrasena', 'password', 'clave', 'pass', 'contraseña']);

    if (columnaRol < 0 || columnaClave < 0) {
        return false;
    }

    var dniBuscado = String(dni || '').trim();
    var claveBuscada = String(clave || '').trim();
    var rolesPermitidos = ['admin', 'administrador', 'administradora', 'administradores', 'administradoras', 'superadmin', 'superadministrador'];

    return filas.slice(1).some(function (fila) {
        var rol = normalizarEncabezado(fila[columnaRol]);
        return String(fila[columnaDni] || '').trim() === dniBuscado &&
            rolesPermitidos.indexOf(rol) !== -1 &&
            String(fila[columnaClave] || '').trim() === claveBuscada;
    });
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

    if (accion === 'listarClientesAdmin') {
        if (!validarAdmin(data.dniAdmin, data.claveAdmin)) {
            return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'DNI, contraseña o rol de administrador incorrectos.' })).setMimeType(ContentService.MimeType.JSON);
        }
        return listarClientesAdmin();
    }

    if (accion === 'actualizarSello') {
        if (!validarAdmin(data.dniAdmin, data.claveAdmin)) {
            return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'DNI, contraseña o rol de administrador incorrectos.' })).setMimeType(ContentService.MimeType.JSON);
        }
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

    var row = hojaSpreadsheet.getRange(fila, 1, 1, hojaSpreadsheet.getLastColumn()).getValues()[0];
    var cliente = obtenerClienteDesdeFila(row);

    return ContentService.createTextOutput(JSON.stringify({ success: true, cliente: cliente })).setMimeType(ContentService.MimeType.JSON);
}

function listarClientesAdmin() {
    var rows = hojaSpreadsheet.getDataRange().getValues();
    var columnasBeneficio = obtenerColumnasBeneficio();
    var clientes = rows.slice(1).filter(function (row) {
        return row[0] && row[1];
    }).map(function (row) {
        return {
            numeroCliente: String(row[0]),
            nombre: String(row[1]),
            apellido: String(row[2] || ''),
            dni: String(row[3] || ''),
            sellos: row.slice(7, 12).map(convertirSello),
            descuentosObtenidos: Number(row[columnasBeneficio.descuentos - 1]) || 0,
            ultimaCompra: row[columnasBeneficio.ultimaCompra - 1] || ''
        };
    });

    clientes.sort(function (a, b) {
        return (a.apellido + ' ' + a.nombre).localeCompare(b.apellido + ' ' + b.nombre);
    });

    return ContentService.createTextOutput(JSON.stringify({ success: true, clientes: clientes })).setMimeType(ContentService.MimeType.JSON);
}

function actualizarSelloCliente(data) {
    var numeroCliente = String(data.numeroCliente || '').trim();

    if (!numeroCliente) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Cliente inválido.' })).setMimeType(ContentService.MimeType.JSON);
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    var cliente;
    var totalSellos;
    var descuentosObtenidos;
    var ultimaCompra;
    var estampaAgregada = false;
    var descuentoAcreditado = false;

    try {
        var fila = buscarFilaPorNumeroCliente(numeroCliente);

        if (!fila) {
            return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Cliente no encontrado' })).setMimeType(ContentService.MimeType.JSON);
        }

        var row = hojaSpreadsheet.getRange(fila, 1, 1, hojaSpreadsheet.getLastColumn()).getValues()[0];
        var cliente = obtenerClienteDesdeFila(row);
        var siguienteSello = cliente.sellos.findIndex(function (sello) {
            return sello === 0;
        });

        if (siguienteSello !== -1) {
            var columnaSello = 8 + siguienteSello;
            hojaSpreadsheet.getRange(fila, columnaSello).setValue(1);
            cliente.sellos[siguienteSello] = 1;
            ultimaCompra = new Date();
            hojaSpreadsheet.getRange(fila, obtenerColumnasBeneficio().ultimaCompra).setValue(ultimaCompra);
            estampaAgregada = true;
        }

        totalSellos = cliente.sellos.reduce(function (total, sello) {
            return total + sello;
        }, 0);

        var columnasBeneficio = obtenerColumnasBeneficio();
        descuentosObtenidos = Number(hojaSpreadsheet.getRange(fila, columnasBeneficio.descuentos).getValue()) || 0;
        ultimaCompra = hojaSpreadsheet.getRange(fila, columnasBeneficio.ultimaCompra).getValue() || ultimaCompra || '';

        if (totalSellos >= TOTAL_SELLOS) {
            hojaSpreadsheet.getRange(fila, 8, 1, TOTAL_SELLOS).setValues([[0, 0, 0, 0, 0]]);
            descuentosObtenidos++;
            hojaSpreadsheet.getRange(fila, columnasBeneficio.descuentos).setValue(descuentosObtenidos);
            cliente.sellos = [0, 0, 0, 0, 0];
            totalSellos = 0;
            descuentoAcreditado = true;
        }

        cliente.descuentosObtenidos = descuentosObtenidos;
        cliente.ultimaCompra = ultimaCompra;
    } finally {
        lock.releaseLock();
    }

    var emailEnviado = descuentoAcreditado
        ? enviarCorreoDescuentoDesbloqueado(cliente, data.urlTarjeta)
        : enviarCorreoNuevaEstampa(cliente, totalSellos, data.urlTarjeta);
    return ContentService.createTextOutput(JSON.stringify({
        success: true,
        completo: descuentoAcreditado,
        descuentoAcreditado: descuentoAcreditado,
        estampaAgregada: estampaAgregada,
        emailEnviado: emailEnviado,
        enlaceIncluido: Boolean(obtenerUrlTarjeta(data.urlTarjeta)),
        cliente: {
            numeroCliente: String(cliente.numeroCliente),
            nombre: cliente.nombre,
            apellido: cliente.apellido,
            sellos: cliente.sellos,
            descuentosObtenidos: cliente.descuentosObtenidos,
            ultimaCompra: cliente.ultimaCompra
        }
    })).setMimeType(ContentService.MimeType.JSON);
}