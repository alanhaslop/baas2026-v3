// ============================================================
// BAAS 2026 — Gestor de vtas v3
// Google Apps Script — Backend completo
// ============================================================
// CONFIGURACIÓN: completar antes del congreso
// ============================================================

const CONFIG = {
  SHEET_ID:            '',           // ID de la Google Sheet nueva (vacío = crea en este spreadsheet)
  DRIVE_FOLDER_ID:     '',           // ID de carpeta en Drive para PDFs
  EMAIL_OPERADOR:      'alan.haslop@dermacells.com.ar',
  EMAIL_ADMINISTRATIVA:'',           // email del área de despacho
  PASSWORD_OPERADOR:   'agf2026op',
  PASSWORD_ADMIN:      'agf2026adm',
  CODIGO_SUPERVISOR:   'super2026',
  WA_NUMERO_COMERCIAL: '549',        // prefijo + número sin espacios
  CONGRESO_NOMBRE:     'BAAS 2026',
  PRECIO_CAJA_CERRADA: 750,
  PRECIO_CAJA_COMBINADA: 900,
  STOCK_INICIAL: {
    Dermal:   100,
    Capillary: 100,
    Pink:      100,
    Biomask:   100
  }
};

// ============================================================
// PUNTO DE ENTRADA — recibe POST del frontend
// ============================================================

function doPost(e) {
  const cors = ContentService.createTextOutput();
  cors.setMimeType(ContentService.MimeType.JSON);

  try {
    const data = JSON.parse(e.postData.contents);
    const accion = data.accion;
    let resultado;

    switch (accion) {
      case 'login':             resultado = login(data); break;
      case 'obtenerStock':      resultado = obtenerStock(); break;
      case 'confirmarVenta':    resultado = registrarTransaccion(data, 'Venta'); break;
      case 'generarReserva':    resultado = registrarTransaccion(data, 'Reserva'); break;
      case 'cerrarReserva':     resultado = cerrarReserva(data); break;
      case 'modificarReserva':  resultado = modificarReserva(data); break;
      case 'modificarAlRetirar':resultado = modificarAlRetirar(data); break;
      case 'cancelarVenta':     resultado = cancelarTransaccion(data); break;
      case 'avanzarEstado':     resultado = avanzarEstadoPedido(data); break;
      case 'obtenerPedidos':    resultado = obtenerPedidos(data); break;
      case 'obtenerReservas':   resultado = obtenerReservasPendientes(); break;
      case 'cierreDelDia':      resultado = generarCierreDelDia(); break;
      case 'reimprimir':        resultado = reimprimirComprobante(data); break;
      default:
        resultado = { ok: false, error: 'Acción desconocida: ' + accion };
    }

    cors.setContent(JSON.stringify(resultado));
  } catch (err) {
    cors.setContent(JSON.stringify({ ok: false, error: err.toString() }));
  }

  return cors;
}

function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({ ok: true, sistema: 'BAAS 2026 v3', estado: 'activo' })
  ).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// AUTENTICACIÓN
// ============================================================

function login(data) {
  const { password, rol } = data;
  if (rol === 'operador' && password === CONFIG.PASSWORD_OPERADOR) {
    return { ok: true, rol: 'operador' };
  }
  if (rol === 'administrativa' && password === CONFIG.PASSWORD_ADMIN) {
    return { ok: true, rol: 'administrativa' };
  }
  return { ok: false, error: 'Contraseña incorrecta' };
}

// ============================================================
// SHEETS — inicialización y acceso
// ============================================================

function getSheet(nombre) {
  const ss = CONFIG.SHEET_ID
    ? SpreadsheetApp.openById(CONFIG.SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  let hoja = ss.getSheetByName(nombre);
  if (!hoja) hoja = crearHoja(ss, nombre);
  return hoja;
}

function crearHoja(ss, nombre) {
  const hoja = ss.insertSheet(nombre);

  if (nombre === 'Ventas') {
    hoja.appendRow([
      'ID','Timestamp','Tipo','Estado','ID_Reserva_Origen',
      'Nombre','CUIT','Email','Telefono','Metodo_Pago','Moneda',
      'Tipo_Cambio','Dermal','Capillary','Pink','Biomask',
      'Detalle_Cajas','Descuento_Pct','Monto_USD','Monto_ARS',
      'Fecha_Vencimiento','Estado_Impresion','URL_PDF',
      'Email_Enviado','Operador','Notas'
    ]);
    hoja.getRange('1:1').setFontWeight('bold').setBackground('#1B3A52').setFontColor('#ffffff');
    hoja.setFrozenRows(1);
  }

  if (nombre === 'Stock') {
    hoja.appendRow(['Producto','Stock_Total','Vendido','Comprometido','Disponible']);
    hoja.getRange('1:1').setFontWeight('bold').setBackground('#1B3A52').setFontColor('#ffffff');
    Object.entries(CONFIG.STOCK_INICIAL).forEach(([prod, cant]) => {
      hoja.appendRow([prod, cant, 0, 0, cant]);
    });
  }

  if (nombre === 'CierresDeDia') {
    hoja.appendRow([
      'Fecha','Ventas_Directas_Qty','Ventas_Directas_USD',
      'Reservas_Cerradas_Qty','Reservas_Cerradas_USD',
      'Reservas_Activas_Qty','Anulaciones_Qty',
      'Efectivo_USD','Efectivo_ARS','Transferencia_USD',
      'Transferencia_ARS','Tarjeta_ARS','QR_ARS',
      'Stock_Dermal','Stock_Capillary','Stock_Pink','Stock_Biomask',
      'Total_Recaudado_USD','Total_Recaudado_ARS'
    ]);
    hoja.getRange('1:1').setFontWeight('bold').setBackground('#1B3A52').setFontColor('#ffffff');
  }

  return hoja;
}

// ============================================================
// STOCK
// ============================================================

function obtenerStock() {
  const hoja = getSheet('Stock');
  const datos = hoja.getDataRange().getValues();
  const stock = {};
  for (let i = 1; i < datos.length; i++) {
    stock[datos[i][0]] = {
      total:       datos[i][1],
      vendido:     datos[i][2],
      comprometido:datos[i][3],
      disponible:  datos[i][4]
    };
  }
  return { ok: true, stock };
}

function actualizarStock(producto, cantidad, tipo) {
  // tipo: 'vender' | 'comprometer' | 'liberar' | 'confirmar'
  const hoja = getSheet('Stock');
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === producto) {
      const fila = i + 1;
      const total       = datos[i][1];
      let vendido       = datos[i][2];
      let comprometido  = datos[i][3];

      if (tipo === 'vender') {
        vendido += cantidad;
      } else if (tipo === 'comprometer') {
        comprometido += cantidad;
      } else if (tipo === 'liberar') {
        comprometido = Math.max(0, comprometido - cantidad);
      } else if (tipo === 'confirmar') {
        // reserva → venta: pasa de comprometido a vendido
        comprometido = Math.max(0, comprometido - cantidad);
        vendido += cantidad;
      }

      const disponible = total - vendido - comprometido;
      hoja.getRange(fila, 3).setValue(vendido);
      hoja.getRange(fila, 4).setValue(comprometido);
      hoja.getRange(fila, 5).setValue(disponible);
      return disponible;
    }
  }
  return -1;
}

function verificarStockDisponible(unidades) {
  const stockActual = obtenerStock().stock;
  const errores = [];
  const productos = ['Dermal','Capillary','Pink','Biomask'];
  productos.forEach(prod => {
    const solicitado = unidades[prod] || 0;
    const disponible = stockActual[prod] ? stockActual[prod].disponible : 0;
    if (solicitado > disponible) {
      errores.push(`${prod}: solicitado ${solicitado}, disponible ${disponible}`);
    }
  });
  return errores;
}

function moverStock(unidades, tipo) {
  ['Dermal','Capillary','Pink','Biomask'].forEach(prod => {
    const cant = unidades[prod] || 0;
    if (cant > 0) actualizarStock(prod, cant, tipo);
  });
}

// ============================================================
// GENERACIÓN DE ID
// ============================================================

function proximoId() {
  const hoja = getSheet('Ventas');
  const ultima = hoja.getLastRow();
  if (ultima <= 1) return 1;
  const datos = hoja.getRange(2, 1, ultima - 1, 1).getValues();
  const ids = datos.map(r => parseInt(r[0]) || 0);
  return Math.max(...ids) + 1;
}

// ============================================================
// REGISTRAR TRANSACCIÓN (venta directa o reserva nueva)
// ============================================================

function registrarTransaccion(data, tipo) {
  const erroresStock = verificarStockDisponible(data.unidades);
  if (erroresStock.length > 0) {
    return { ok: false, error: 'Stock insuficiente', detalle: erroresStock };
  }

  const id = proximoId();
  const timestamp = new Date();
  const hoja = getSheet('Ventas');

  const fila = [
    id,
    Utilities.formatDate(timestamp, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm:ss'),
    tipo,
    tipo === 'Venta' ? 'Pendiente' : 'Pendiente',
    '',                                         // ID_Reserva_Origen
    data.nombre || '',
    data.cuit || 'Extranjero',
    data.email || '',
    data.telefono || '',
    data.metodoPago || '',
    data.moneda || 'USD',
    data.tipoCambio || '',
    data.unidades.Dermal    || 0,
    data.unidades.Capillary || 0,
    data.unidades.Pink      || 0,
    data.unidades.Biomask   || 0,
    JSON.stringify(data.detalleCajas || []),
    data.descuentoPct || 0,
    data.montoUSD || 0,
    data.montoARS || 0,
    tipo === 'Reserva' ? (data.fechaVencimiento || '') : '',
    'Pendiente',
    '',
    'No',
    data.operador || 'operador',
    data.notas || ''
  ];

  hoja.appendRow(fila);

  // Mover stock
  if (tipo === 'Venta') {
    moverStock(data.unidades, 'vender');
  } else {
    moverStock(data.unidades, 'comprometer');
  }

  // Generar PDF
  let urlPDF = '';
  try {
    urlPDF = generarPDF(id, tipo, data);
    const filaNum = hoja.getLastRow();
    hoja.getRange(filaNum, 23).setValue(urlPDF); // columna URL_PDF
    hoja.getRange(filaNum, 22).setValue('OK');   // Estado_Impresion
  } catch (e) {
    const filaNum = hoja.getLastRow();
    hoja.getRange(filaNum, 22).setValue('Error');
  }

  // Enviar email
  let emailOk = false;
  try {
    enviarEmailConfirmacion(id, tipo, data, urlPDF);
    const filaNum = hoja.getLastRow();
    hoja.getRange(filaNum, 24).setValue('Sí');
    emailOk = true;
  } catch (e) {
    // silencioso
  }

  return {
    ok: true,
    id,
    tipo,
    urlPDF,
    emailOk,
    estadoImpresion: urlPDF ? 'OK' : 'Error'
  };
}

// ============================================================
// CERRAR RESERVA
// ============================================================

function cerrarReserva(data) {
  const hoja = getSheet('Ventas');
  const fila = buscarFilaPorId(data.idReserva);
  if (!fila) return { ok: false, error: 'Reserva no encontrada' };

  const idCierre = proximoId();
  const timestamp = new Date();

  // Registrar el cierre como nueva fila
  const filaCierre = [
    idCierre,
    Utilities.formatDate(timestamp, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm:ss'),
    'Cierre reserva',
    'Pendiente',
    data.idReserva,
    data.nombre || fila[5],
    data.cuit   || fila[6],
    data.email  || fila[7],
    data.telefono || fila[8],
    data.metodoPago || fila[9],
    data.moneda || fila[10],
    data.tipoCambio || fila[11],
    data.unidades ? (data.unidades.Dermal    || fila[12]) : fila[12],
    data.unidades ? (data.unidades.Capillary || fila[13]) : fila[13],
    data.unidades ? (data.unidades.Pink      || fila[14]) : fila[14],
    data.unidades ? (data.unidades.Biomask   || fila[15]) : fila[15],
    data.detalleCajas ? JSON.stringify(data.detalleCajas) : fila[16],
    data.descuentoPct || fila[17],
    data.montoUSD || fila[18],
    data.montoARS || fila[19],
    '',
    'Pendiente',
    '',
    'No',
    data.operador || 'operador',
    ''
  ];

  hoja.appendRow(filaCierre);

  // Confirmar stock: de comprometido a vendido
  const unidades = {
    Dermal:    filaCierre[12],
    Capillary: filaCierre[13],
    Pink:      filaCierre[14],
    Biomask:   filaCierre[15]
  };
  moverStock(unidades, 'confirmar');

  // Marcar reserva original como cerrada
  const numFila = buscarNumFilaPorId(data.idReserva);
  if (numFila) hoja.getRange(numFila, 4).setValue('Cerrada');

  // PDF y email
  let urlPDF = '';
  try {
    const dataParaPDF = { ...data, idReservaOrigen: data.idReserva };
    urlPDF = generarPDF(idCierre, 'Cierre reserva', dataParaPDF);
    const filaNum = hoja.getLastRow();
    hoja.getRange(filaNum, 23).setValue(urlPDF);
    hoja.getRange(filaNum, 22).setValue('OK');
  } catch (e) {
    hoja.getRange(hoja.getLastRow(), 22).setValue('Error');
  }

  try {
    enviarEmailConfirmacion(idCierre, 'Cierre reserva', data, urlPDF);
    hoja.getRange(hoja.getLastRow(), 24).setValue('Sí');
  } catch (e) {}

  return { ok: true, id: idCierre, urlPDF };
}

// ============================================================
// MODIFICAR RESERVA (antes del cobro)
// ============================================================

function modificarReserva(data) {
  const hoja = getSheet('Ventas');
  const numFila = buscarNumFilaPorId(data.idReserva);
  if (!numFila) return { ok: false, error: 'Reserva no encontrada' };

  const filaActual = hoja.getRange(numFila, 1, 1, 26).getValues()[0];
  const unidadesViejas = {
    Dermal:    filaActual[12],
    Capillary: filaActual[13],
    Pink:      filaActual[14],
    Biomask:   filaActual[15]
  };

  // Liberar stock anterior
  moverStock(unidadesViejas, 'liberar');

  // Verificar nuevo stock
  const errores = verificarStockDisponible(data.unidades);
  if (errores.length > 0) {
    // Revertir
    moverStock(unidadesViejas, 'comprometer');
    return { ok: false, error: 'Stock insuficiente', detalle: errores };
  }

  // Actualizar fila
  hoja.getRange(numFila, 13).setValue(data.unidades.Dermal    || 0);
  hoja.getRange(numFila, 14).setValue(data.unidades.Capillary || 0);
  hoja.getRange(numFila, 15).setValue(data.unidades.Pink      || 0);
  hoja.getRange(numFila, 16).setValue(data.unidades.Biomask   || 0);
  hoja.getRange(numFila, 17).setValue(JSON.stringify(data.detalleCajas || []));
  hoja.getRange(numFila, 18).setValue(data.descuentoPct || 0);
  hoja.getRange(numFila, 19).setValue(data.montoUSD || 0);
  hoja.getRange(numFila, 20).setValue(data.montoARS || 0);
  hoja.getRange(numFila, 22).setValue('Pendiente');

  // Comprometer nuevo stock
  moverStock(data.unidades, 'comprometer');

  // Regenerar PDF
  let urlPDF = '';
  try {
    urlPDF = generarPDF(data.idReserva, 'Modificación reserva', data);
    hoja.getRange(numFila, 23).setValue(urlPDF);
    hoja.getRange(numFila, 22).setValue('OK');
  } catch(e) {}

  return { ok: true, id: data.idReserva, urlPDF };
}

// ============================================================
// MODIFICACIÓN AL RETIRAR
// ============================================================

function modificarAlRetirar(data) {
  const hoja = getSheet('Ventas');
  const numFila = buscarNumFilaPorId(data.idPedido);
  if (!numFila) return { ok: false, error: 'Pedido no encontrado' };

  const filaActual = hoja.getRange(numFila, 1, 1, 26).getValues()[0];
  const unidadesViejas = {
    Dermal:    filaActual[12],
    Capillary: filaActual[13],
    Pink:      filaActual[14],
    Biomask:   filaActual[15]
  };

  // Calcular diferencia de stock
  const productos = ['Dermal','Capillary','Pink','Biomask'];
  const diferencias = {};
  productos.forEach(p => {
    diferencias[p] = (data.unidades[p] || 0) - (unidadesViejas[p] || 0);
  });

  // Verificar stock para las unidades adicionales
  const adicionales = {};
  productos.forEach(p => { if (diferencias[p] > 0) adicionales[p] = diferencias[p]; });
  if (Object.keys(adicionales).length > 0) {
    const errores = verificarStockDisponible(adicionales);
    if (errores.length > 0) return { ok: false, error: 'Stock insuficiente', detalle: errores };
  }

  // Ajustar stock
  productos.forEach(p => {
    if (diferencias[p] > 0) actualizarStock(p, diferencias[p], 'vender');
    else if (diferencias[p] < 0) {
      // devuelve unidades al stock disponible
      const hStock = getSheet('Stock');
      const datosS = hStock.getDataRange().getValues();
      for (let i = 1; i < datosS.length; i++) {
        if (datosS[i][0] === p) {
          const fNum = i + 1;
          const vendido = Math.max(0, datosS[i][2] + diferencias[p]);
          const disponible = datosS[i][1] - vendido - datosS[i][3];
          hStock.getRange(fNum, 3).setValue(vendido);
          hStock.getRange(fNum, 5).setValue(disponible);
        }
      }
    }
  });

  // Actualizar fila
  hoja.getRange(numFila, 13).setValue(data.unidades.Dermal    || 0);
  hoja.getRange(numFila, 14).setValue(data.unidades.Capillary || 0);
  hoja.getRange(numFila, 15).setValue(data.unidades.Pink      || 0);
  hoja.getRange(numFila, 16).setValue(data.unidades.Biomask   || 0);
  hoja.getRange(numFila, 17).setValue(JSON.stringify(data.detalleCajas || []));
  hoja.getRange(numFila, 18).setValue(data.descuentoPct || 0);
  hoja.getRange(numFila, 19).setValue(data.montoUSD || 0);
  hoja.getRange(numFila, 20).setValue(data.montoARS || 0);
  hoja.getRange(numFila, 22).setValue('Pendiente');

  // Nuevo recibo
  let urlPDF = '';
  try {
    urlPDF = generarPDF(data.idPedido, 'Modificación al retirar', { ...data, idPedidoOrigen: data.idPedido });
    hoja.getRange(numFila, 23).setValue(urlPDF);
    hoja.getRange(numFila, 22).setValue('OK');
  } catch(e) {}

  return { ok: true, id: data.idPedido, urlPDF };
}

// ============================================================
// CANCELAR VENTA O RESERVA
// ============================================================

function cancelarTransaccion(data) {
  if (data.codigoSupervisor !== CONFIG.CODIGO_SUPERVISOR) {
    return { ok: false, error: 'Código de supervisor incorrecto' };
  }

  const hoja = getSheet('Ventas');
  const numFila = buscarNumFilaPorId(data.id);
  if (!numFila) return { ok: false, error: 'Transacción no encontrada' };

  const filaActual = hoja.getRange(numFila, 1, 1, 26).getValues()[0];
  const tipo = filaActual[2];
  const estado = filaActual[3];

  if (estado === 'Entregado') {
    return { ok: false, error: 'No se puede cancelar un pedido ya entregado' };
  }

  const unidades = {
    Dermal:    filaActual[12],
    Capillary: filaActual[13],
    Pink:      filaActual[14],
    Biomask:   filaActual[15]
  };

  // Liberar stock según tipo
  if (tipo === 'Venta' || tipo === 'Cierre reserva') {
    // Devolver a vendido → disponible
    const hStock = getSheet('Stock');
    const datosS = hStock.getDataRange().getValues();
    ['Dermal','Capillary','Pink','Biomask'].forEach(p => {
      const cant = unidades[p] || 0;
      if (cant > 0) {
        for (let i = 1; i < datosS.length; i++) {
          if (datosS[i][0] === p) {
            const fNum = i + 1;
            const vendido = Math.max(0, datosS[i][2] - cant);
            const disponible = datosS[i][1] - vendido - datosS[i][3];
            hStock.getRange(fNum, 3).setValue(vendido);
            hStock.getRange(fNum, 5).setValue(disponible);
          }
        }
      }
    });
  } else if (tipo === 'Reserva') {
    moverStock(unidades, 'liberar');
  }

  hoja.getRange(numFila, 4).setValue('Anulada');

  // Comprobante de anulación
  let urlPDF = '';
  try {
    urlPDF = generarPDF(data.id, 'Anulación', { ...filaActual, motivo: data.motivo || '' });
    hoja.getRange(numFila, 23).setValue(urlPDF);
  } catch(e) {}

  return { ok: true, id: data.id, urlPDF };
}

// ============================================================
// AVANZAR ESTADO DE PEDIDO (panel de despacho)
// ============================================================

function avanzarEstadoPedido(data) {
  const estados = ['Pendiente','En preparación','Listo','Entregado'];
  const hoja = getSheet('Ventas');
  const numFila = buscarNumFilaPorId(data.id);
  if (!numFila) return { ok: false, error: 'Pedido no encontrado' };

  const estadoActual = hoja.getRange(numFila, 4).getValue();
  const idx = estados.indexOf(estadoActual);
  if (idx === -1 || idx === estados.length - 1) {
    return { ok: false, error: 'Estado no puede avanzar' };
  }

  const nuevoEstado = estados[idx + 1];
  hoja.getRange(numFila, 4).setValue(nuevoEstado);

  // Si pasa a Entregado: registrar hora
  if (nuevoEstado === 'Entregado') {
    const horaEntrega = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'HH:mm');
    const notasActual = hoja.getRange(numFila, 26).getValue();
    hoja.getRange(numFila, 26).setValue(notasActual + ' | Entregado: ' + horaEntrega);
  }

  return { ok: true, nuevoEstado };
}

// ============================================================
// OBTENER PEDIDOS PARA EL PANEL DE DESPACHO
// ============================================================

function obtenerPedidos(data) {
  const hoja = getSheet('Ventas');
  const datos = hoja.getDataRange().getValues();
  const estadosActivos = ['Pendiente','En preparación','Listo'];
  // Incluir también entregados del día
  const hoy = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');

  const pedidos = [];
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const estado = fila[3];
    const timestamp = fila[1] ? String(fila[1]) : '';
    const esDeHoy = timestamp.startsWith(hoy);
    const tipo = fila[2];
    const esRelevante = estadosActivos.includes(estado) ||
      (estado === 'Entregado' && esDeHoy);
    const esAnulado = estado === 'Anulada' || estado === 'Cerrada';

    if (!esRelevante || esAnulado) continue;

    pedidos.push({
      id:          fila[0],
      timestamp:   fila[1],
      tipo:        tipo,
      estado:      estado,
      nombre:      fila[5],
      cuit:        fila[6],
      dermal:      fila[12],
      capillary:   fila[13],
      pink:        fila[14],
      biomask:     fila[15],
      detalleCajas:fila[16],
      montoUSD:    fila[18],
      montoARS:    fila[19],
      urlPDF:      fila[22],
      estadoImp:   fila[21],
      telefono:    fila[8]
    });
  }

  return { ok: true, pedidos };
}

// ============================================================
// OBTENER RESERVAS PENDIENTES (panel operador)
// ============================================================

function obtenerReservasPendientes() {
  const hoja = getSheet('Ventas');
  const datos = hoja.getDataRange().getValues();
  const hoy = new Date();
  hoy.setHours(0,0,0,0);

  const reservas = [];
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (fila[2] !== 'Reserva') continue;
    if (fila[3] === 'Cerrada' || fila[3] === 'Anulada') continue;

    const vencimiento = fila[20] ? new Date(fila[20]) : null;
    let estadoVenc = 'sin-fecha';
    if (vencimiento) {
      const diff = Math.floor((vencimiento - hoy) / 86400000);
      if (diff < 0)  estadoVenc = 'vencida';
      else if (diff === 0) estadoVenc = 'hoy';
      else if (diff === 1) estadoVenc = 'manana';
      else estadoVenc = 'futuro';
    }

    reservas.push({
      id:            fila[0],
      timestamp:     fila[1],
      nombre:        fila[5],
      cuit:          fila[6],
      telefono:      fila[8],
      email:         fila[7],
      dermal:        fila[12],
      capillary:     fila[13],
      pink:          fila[14],
      biomask:       fila[15],
      detalleCajas:  fila[16],
      montoUSD:      fila[18],
      montoARS:      fila[19],
      fechaVenc:     fila[20],
      estadoVenc,
      urlPDF:        fila[22]
    });
  }

  // Ordenar: vencidas primero, luego hoy, mañana, futuro
  const orden = { vencida:0, hoy:1, manana:2, futuro:3, 'sin-fecha':4 };
  reservas.sort((a,b) => orden[a.estadoVenc] - orden[b.estadoVenc]);

  return { ok: true, reservas };
}

// ============================================================
// CIERRE DEL DÍA
// ============================================================

function generarCierreDelDia() {
  const hoja = getSheet('Ventas');
  const datos = hoja.getDataRange().getValues();
  const hoy = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');

  const resumen = {
    fecha: hoy,
    ventasDirectas: { qty: 0, usd: 0 },
    reservasCerradas: { qty: 0, usd: 0 },
    reservasActivas: { qty: 0 },
    anulaciones: { qty: 0 },
    porMetodo: {
      Efectivo:     { usd: 0, ars: 0 },
      Transferencia:{ usd: 0, ars: 0 },
      Tarjeta:      { ars: 0 },
      QR:           { ars: 0 }
    }
  };

  for (let i = 1; i < datos.length; i++) {
    const f = datos[i];
    const ts = String(f[1]);
    if (!ts.startsWith(hoy)) continue;

    const tipo   = f[2];
    const estado = f[3];
    const metodo = f[9] || 'Efectivo';
    const moneda = f[10] || 'USD';
    const mUSD   = parseFloat(f[18]) || 0;
    const mARS   = parseFloat(f[19]) || 0;

    if (estado === 'Anulada') { resumen.anulaciones.qty++; continue; }

    if (tipo === 'Venta') {
      resumen.ventasDirectas.qty++;
      resumen.ventasDirectas.usd += mUSD;
    } else if (tipo === 'Cierre reserva') {
      resumen.reservasCerradas.qty++;
      resumen.reservasCerradas.usd += mUSD;
    } else if (tipo === 'Reserva' && estado !== 'Cerrada') {
      resumen.reservasActivas.qty++;
    }

    // Por método
    if (metodo === 'Efectivo') {
      if (moneda === 'USD') resumen.porMetodo.Efectivo.usd += mUSD;
      else resumen.porMetodo.Efectivo.ars += mARS;
    } else if (metodo === 'Transferencia') {
      if (moneda === 'USD') resumen.porMetodo.Transferencia.usd += mUSD;
      else resumen.porMetodo.Transferencia.ars += mARS;
    } else if (metodo === 'Tarjeta') {
      resumen.porMetodo.Tarjeta.ars += mARS;
    } else if (metodo === 'QR') {
      resumen.porMetodo.QR.ars += mARS;
    }
  }

  // Stock actual
  const stockActual = obtenerStock().stock;
  resumen.stock = stockActual;

  // Calcular totales
  resumen.totalRecaudadoUSD =
    resumen.porMetodo.Efectivo.usd + resumen.porMetodo.Transferencia.usd;
  resumen.totalRecaudadoARS =
    resumen.porMetodo.Efectivo.ars + resumen.porMetodo.Transferencia.ars +
    resumen.porMetodo.Tarjeta.ars + resumen.porMetodo.QR.ars;

  // Guardar en hoja CierresDeDia
  try {
    const hojaC = getSheet('CierresDeDia');
    hojaC.appendRow([
      hoy,
      resumen.ventasDirectas.qty,   resumen.ventasDirectas.usd,
      resumen.reservasCerradas.qty, resumen.reservasCerradas.usd,
      resumen.reservasActivas.qty,  resumen.anulaciones.qty,
      resumen.porMetodo.Efectivo.usd,    resumen.porMetodo.Efectivo.ars,
      resumen.porMetodo.Transferencia.usd, resumen.porMetodo.Transferencia.ars,
      resumen.porMetodo.Tarjeta.ars, resumen.porMetodo.QR.ars,
      stockActual.Dermal    ? stockActual.Dermal.disponible    : 0,
      stockActual.Capillary ? stockActual.Capillary.disponible : 0,
      stockActual.Pink      ? stockActual.Pink.disponible      : 0,
      stockActual.Biomask   ? stockActual.Biomask.disponible   : 0,
      resumen.totalRecaudadoUSD,
      resumen.totalRecaudadoARS
    ]);
  } catch(e) {}

  return { ok: true, resumen };
}

// ============================================================
// REIMPRIMIR COMPROBANTE
// ============================================================

function reimprimirComprobante(data) {
  const hoja = getSheet('Ventas');
  const numFila = buscarNumFilaPorId(data.id);
  if (!numFila) return { ok: false, error: 'Pedido no encontrado' };

  const filaActual = hoja.getRange(numFila, 1, 1, 26).getValues()[0];
  let urlPDF = filaActual[22];

  if (!urlPDF) {
    // Regenerar
    try {
      const dataRegen = {
        nombre: filaActual[5], cuit: filaActual[6], email: filaActual[7],
        telefono: filaActual[8], metodoPago: filaActual[9], moneda: filaActual[10],
        tipoCambio: filaActual[11],
        unidades: { Dermal: filaActual[12], Capillary: filaActual[13], Pink: filaActual[14], Biomask: filaActual[15] },
        detalleCajas: JSON.parse(filaActual[16] || '[]'),
        descuentoPct: filaActual[17], montoUSD: filaActual[18], montoARS: filaActual[19]
      };
      urlPDF = generarPDF(data.id, filaActual[2], dataRegen);
      hoja.getRange(numFila, 23).setValue(urlPDF);
      hoja.getRange(numFila, 22).setValue('OK');
    } catch(e) {
      return { ok: false, error: 'Error al regenerar PDF: ' + e.toString() };
    }
  }

  return { ok: true, urlPDF };
}

// ============================================================
// GENERACIÓN DE PDF
// ============================================================

function generarPDF(id, tipo, data) {
  const html = construirHTMLComprobante(id, tipo, data);
  const blob = Utilities.newBlob(html, 'text/html', `comprobante_${id}.html`);

  let folder;
  try {
    folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  } catch(e) {
    folder = DriveApp.getRootFolder();
  }

  const archivo = folder.createFile(blob);
  archivo.setName(`BAAS2026_${tipo.replace(/ /g,'_')}_${id}.html`);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return archivo.getUrl();
}

function construirHTMLComprobante(id, tipo, data) {
  const fecha = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');
  const esPendiente = tipo === 'Reserva';
  const esAnulacion = tipo === 'Anulación';
  const esModificacion = tipo.includes('Modificación') || tipo.includes('Modificacion');
  const esCierre = tipo === 'Cierre reserva';

  let watermark = '';
  if (esPendiente) watermark = `<div class="watermark">PENDIENTE DE PAGO</div>`;
  if (esAnulacion) watermark = `<div class="watermark anulado">ANULADO</div>`;

  let refOrigen = '';
  if (esCierre && data.idReservaOrigen)
    refOrigen = `<p class="ref">Cancela reserva #${data.idReservaOrigen}</p>`;
  if (esModificacion && (data.idPedidoOrigen || data.idReservaOrigen))
    refOrigen = `<p class="ref">Modifica pedido #${data.idPedidoOrigen || data.idReservaOrigen}</p>`;

  const unidades = data.unidades || {};
  const cajas = data.detalleCajas || [];
  let filasProductos = '';
  ['Dermal','Capillary','Pink','Biomask'].forEach(p => {
    if ((unidades[p] || 0) > 0) {
      filasProductos += `<tr><td>${p}</td><td>${unidades[p]}</td></tr>`;
    }
  });

  return `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<style>
  @import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@300;400;600&family=Inter:wght@300;400;500;600&display=swap');
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Inter',sans-serif;font-size:12px;color:#1a1a1a;background:#e8e6e0;padding:40px 20px}
  .page{max-width:760px;margin:0 auto;background:#fff;box-shadow:0 8px 40px rgba(0,0,0,.15);position:relative;overflow:hidden}
  .watermark{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%) rotate(-35deg);font-size:64px;font-weight:700;opacity:.06;white-space:nowrap;color:#1B3A52;pointer-events:none;z-index:0}
  .watermark.anulado{color:#e74c3c}
  .header{padding:20px 36px 16px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #e0ddd6;position:relative;z-index:1}
  .logo-text{font-family:'Cormorant Garamond',serif;font-size:22px;font-weight:600;color:#1B3A52}
  .header-right{text-align:right}
  .num{font-family:'Cormorant Garamond',serif;font-size:28px;font-weight:300;color:#1B3A52}
  .fecha{font-size:10px;color:#aaa;margin-top:4px}
  .ref{font-size:10px;color:#e67e22;margin-top:2px;font-style:italic}
  .banda{background:#1B3A52;padding:7px 36px;display:flex;justify-content:space-between}
  .banda span{font-size:9px;font-weight:500;letter-spacing:.2em;text-transform:uppercase;color:#9FD4C0}
  .content{padding:18px 36px 24px;position:relative;z-index:1}
  .section-title{font-size:8px;font-weight:600;letter-spacing:.2em;text-transform:uppercase;color:#1B3A52;border-bottom:1px solid #d8d4cc;padding-bottom:4px;margin-bottom:10px;margin-top:16px}
  .section-title:first-child{margin-top:0}
  .datos-grid{display:grid;grid-template-columns:1fr 1fr;gap:1px 24px}
  .dato-row{display:flex;gap:8px;padding:2px 0}
  .dato-label{color:#bbb;min-width:80px;font-size:10px}
  .dato-val{color:#1a1a1a;font-weight:500;font-size:11px}
  table.prod{width:100%;border-collapse:collapse;font-size:11px;margin-bottom:8px}
  table.prod th{text-align:left;font-size:8px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:#bbb;padding:4px 0;border-bottom:1px solid #e0ddd6}
  table.prod td{padding:5px 0;border-bottom:1px solid #f5f2ec}
  .cobro-total{display:flex;justify-content:space-between;align-items:flex-end;margin-top:16px;padding-top:12px;border-top:2px solid #1B3A52}
  .cobro-label{color:#bbb;min-width:110px;font-size:10px}
  .cobro-val{color:#1a1a1a;font-weight:500;font-size:11px}
  .total-bloque{text-align:right}
  .total-label{font-size:8px;letter-spacing:.2em;text-transform:uppercase;color:#bbb;margin-bottom:3px}
  .total-usd{font-family:'Cormorant Garamond',serif;font-size:38px;font-weight:600;color:#1B3A52;line-height:1}
  .total-ars{font-size:11px;color:#aaa;margin-top:3px}
  .firmas{display:flex;gap:20px;margin-top:32px}
  .firma-col{flex:1;display:flex;flex-direction:column;justify-content:flex-end}
  .firma-espacio{flex:1;min-height:44px}
  .firma-linea{border-top:1px solid #d0ccc4;padding-top:5px}
  .firma-label{font-size:9px;color:#ccc;letter-spacing:.06em}
  .footer{background:#f7f6f2;border-top:1px solid #e8e5de;padding:8px 36px;display:flex;justify-content:space-between}
  .footer span{font-size:9px;color:#ccc}
  @media print{body{background:#fff;padding:0}.page{box-shadow:none}}
</style>
</head>
<body>
<div class="page">
  ${watermark}
  <div class="header">
    <div class="logo-text">AGF Messenchymal</div>
    <div class="header-right">
      <div class="num">#${String(id).padStart(3,'0')}</div>
      <div class="fecha">${fecha}</div>
      ${refOrigen}
    </div>
  </div>
  <div class="banda">
    <span>${CONFIG.CONGRESO_NOMBRE}</span>
    <span>${tipo}</span>
  </div>
  <div class="content">
    <div class="section-title">Datos del cliente</div>
    <div class="datos-grid">
      <div class="dato-row"><span class="dato-label">Nombre</span><span class="dato-val">${data.nombre || ''}</span></div>
      <div class="dato-row"><span class="dato-label">CUIT</span><span class="dato-val">${data.cuit || 'Extranjero'}</span></div>
      <div class="dato-row"><span class="dato-label">Email</span><span class="dato-val">${data.email || ''}</span></div>
      <div class="dato-row"><span class="dato-label">Teléfono</span><span class="dato-val">${data.telefono || ''}</span></div>
    </div>
    <div class="section-title">Productos</div>
    <table class="prod">
      <thead><tr><th>Producto</th><th>Unidades</th></tr></thead>
      <tbody>${filasProductos}</tbody>
    </table>
    <div class="section-title">Condiciones de pago</div>
    <div class="cobro-total">
      <div>
        <div class="dato-row"><span class="cobro-label">Método</span><span class="cobro-val">${data.metodoPago || ''}</span></div>
        <div class="dato-row"><span class="cobro-label">Moneda</span><span class="cobro-val">${data.moneda || 'USD'}</span></div>
        ${data.tipoCambio ? `<div class="dato-row"><span class="cobro-label">Tipo de cambio</span><span class="cobro-val">AR$${data.tipoCambio} / U$D</span></div>` : ''}
        ${data.descuentoPct > 0 ? `<div class="dato-row"><span class="cobro-label">Descuento</span><span class="cobro-val" style="color:#e67e22">${data.descuentoPct}%</span></div>` : ''}
        ${esPendiente ? `<div class="dato-row"><span class="cobro-label">Vencimiento</span><span class="cobro-val" style="color:#e67e22">${data.fechaVencimiento || ''}</span></div>` : ''}
      </div>
      <div class="total-bloque">
        <div class="total-label">Total</div>
        <div class="total-usd">u$${Number(data.montoUSD || 0).toLocaleString('es-AR')}</div>
        ${data.montoARS ? `<div class="total-ars">AR$ ${Number(data.montoARS).toLocaleString('es-AR')}</div>` : ''}
      </div>
    </div>
    <div class="firmas">
      <div class="firma-col"><div class="firma-espacio"></div><div class="firma-linea"><div class="firma-label">Firma y aclaración del cliente</div></div></div>
      <div class="firma-col"><div class="firma-espacio"></div><div class="firma-linea"><div class="firma-label">Firma del responsable</div></div></div>
      <div class="firma-col"><div class="firma-espacio"></div><div class="firma-linea"><div class="firma-label">Fecha de entrega</div></div></div>
    </div>
  </div>
  <div class="footer">
    <span>Documento válido como comprobante de compra</span>
    <span>Dermacells S.A. · ${CONFIG.CONGRESO_NOMBRE}</span>
  </div>
</div>
</body>
</html>`;
}

// ============================================================
// EMAIL
// ============================================================

function enviarEmailConfirmacion(id, tipo, data, urlPDF) {
  const destinatario = data.email;
  if (!destinatario) return;

  const esPendiente = tipo === 'Reserva';
  const asunto = esPendiente
    ? `[${CONFIG.CONGRESO_NOMBRE}] Reserva #${String(id).padStart(3,'0')} confirmada`
    : `[${CONFIG.CONGRESO_NOMBRE}] Pedido #${String(id).padStart(3,'0')} — Comprobante`;

  const cuerpo = esPendiente
    ? `Hola ${data.nombre || ''},\n\nQuedó registrada tu reserva #${String(id).padStart(3,'0')} en ${CONFIG.CONGRESO_NOMBRE}.\n\nTe esperamos para completar el pago y retirar tu pedido.\n\nAdjuntamos tu comprobante de reserva:\n${urlPDF}\n\nAnte cualquier consulta, no dudes en contactarnos.\n\nEquipo AGF Messenchymal`
    : `Hola ${data.nombre || ''},\n\nGracias por tu compra en ${CONFIG.CONGRESO_NOMBRE}. Adjuntamos tu comprobante #${String(id).padStart(3,'0')}:\n${urlPDF}\n\nQuedamos a tu disposición para cualquier consulta.\n\nEquipo AGF Messenchymal`;

  MailApp.sendEmail({
    to: destinatario,
    cc: CONFIG.EMAIL_OPERADOR + (CONFIG.EMAIL_ADMINISTRATIVA ? ',' + CONFIG.EMAIL_ADMINISTRATIVA : ''),
    subject: asunto,
    body: cuerpo
  });
}

// ============================================================
// TRIGGER — alerta diaria de reservas vencidas (7:30 AM)
// ============================================================

function alertaReservasVencidas() {
  const resultado = obtenerReservasPendientes();
  const vencidas = resultado.reservas.filter(r => r.estadoVenc === 'vencida' || r.estadoVenc === 'hoy');
  if (vencidas.length === 0) return;

  let cuerpo = `Buenos días,\n\nHay ${vencidas.length} reserva(s) pendiente(s) para gestionar hoy en ${CONFIG.CONGRESO_NOMBRE}:\n\n`;
  vencidas.forEach(r => {
    cuerpo += `• Reserva #${String(r.id).padStart(3,'0')} — ${r.nombre} — ${r.estadoVenc === 'vencida' ? 'VENCIDA' : 'Vence hoy'}\n`;
    cuerpo += `  Dermal:${r.dermal} Capillary:${r.capillary} Pink:${r.pink} Biomask:${r.biomask}\n`;
    cuerpo += `  Monto: u$${r.montoUSD}\n\n`;
  });

  cuerpo += `Ingresar al sistema para gestionar cada reserva.`;

  MailApp.sendEmail({
    to: CONFIG.EMAIL_OPERADOR,
    cc: CONFIG.EMAIL_ADMINISTRATIVA || '',
    subject: `[${CONFIG.CONGRESO_NOMBRE}] ${vencidas.length} reserva(s) a gestionar hoy`,
    body: cuerpo
  });
}

// Para configurar el trigger: ejecutar esta función UNA VEZ desde el editor
function instalarTrigger() {
  // Eliminar triggers existentes
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  // Crear nuevo trigger diario a las 7:30 AM
  ScriptApp.newTrigger('alertaReservasVencidas')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .nearMinute(30)
    .inTimezone('America/Argentina/Buenos_Aires')
    .create();
}

// ============================================================
// UTILIDADES
// ============================================================

function buscarFilaPorId(id) {
  const hoja = getSheet('Ventas');
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]) === String(id)) return datos[i];
  }
  return null;
}

function buscarNumFilaPorId(id) {
  const hoja = getSheet('Ventas');
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]) === String(id)) return i + 1;
  }
  return null;
}
