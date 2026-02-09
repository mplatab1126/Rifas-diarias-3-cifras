// --- CONFIGURACIÓN ---
const MI_WHATSAPP = "573116675984"; // Tu número (sin el +)
const NOMBRE_HOJA = "Rifa";

function doGet(e) {
  // Si en la URL ponen ?page=abonos, mostramos la pantalla de pagos
  if (e.parameter.page === 'abonos') {
    return HtmlService.createTemplateFromFile('Abonos')
      .evaluate()
      .setTitle('Gestión de Pagos - Los Plata')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // Si no, mostramos la rifa normal
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Rifa Los Plata')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function obtenerDatos() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA);
  // CAMBIO: Ahora pedimos 4 columnas (antes era 2) para leer nombre y teléfono
  // El "4" al final indica la cantidad de columnas A, B, C, D
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
}

function reservarNumeros(listaNumeros, nombre, telefono, ipUsuario) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
  } catch (e) {
    return { exito: false, error: "Servidor ocupado, intenta de nuevo." };
  }

  // --- NUEVO: Validar Mínimo 3 ---
  if (listaNumeros.length < 3) {
    lock.releaseLock();
    return { exito: false, error: "⛔ Error: La compra mínima es de 3 boletas." };
  }
  // -------------------------------

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA);
  
  // CAMBIO 1: Leemos 7 columnas (A hasta G)
  const datos = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues(); 
  
  // --- 🛡️ BLOQUEO DE SEGURIDAD (ANTISPAM) ---
  if (ipUsuario && ipUsuario !== "IP_NO_DETECTADA") {
    let reservasDeEsteUsuario = 0;
    
    for (let i = 0; i < datos.length; i++) {
      const estado = String(datos[i][1]); // Columna B: Estado
      // CAMBIO 2: La IP está en el índice 6 (Columna G)
      const ipRegistrada = String(datos[i][6]); 
      
      if (estado === "Reservado" && ipRegistrada === ipUsuario) {
        reservasDeEsteUsuario++;
      }
    }

    if ((reservasDeEsteUsuario + listaNumeros.length) > 15) {
      lock.releaseLock();
      return { 
        exito: false, 
        error: "🚫 BLOQUEO DE SEGURIDAD 🚫\n\nTu dispositivo ya tiene el máximo de reservas permitidas (15).\nSi quieres más boletas, escribenos a WhatsApp." 
      };
    }
  }
  // ---------------------------------------------

  // 1. VERIFICACIÓN
  let filasEncontradas = [];
  for (let n of listaNumeros) {
    let encontrado = false;
    for (let i = 0; i < datos.length; i++) {
      if (String(datos[i][0]) == String(n)) {
        if (datos[i][1] !== "Disponible") {
          lock.releaseLock();
          return { exito: false, error: "⚠️ Lo sentimos, el número " + n + " ya fue ganado por otra persona." };
        }
        filasEncontradas.push(i + 2); 
        encontrado = true;
        break;
      }
    }
    if (!encontrado) {
      lock.releaseLock();
      return { exito: false, error: "Error: El número " + n + " no existe." };
    }
  }

  // 2. RESERVA
  filasEncontradas.forEach(fila => {
    sheet.getRange(fila, 2).setValue("Reservado");
    sheet.getRange(fila, 3).setValue(nombre);
    sheet.getRange(fila, 4).setValue("'" + telefono);
    
    // CAMBIO 3: Guardamos la IP en la Columna 7 (G)
    sheet.getRange(fila, 7).setValue(ipUsuario || "IP_NO_DETECTADA");
  });

  SpreadsheetApp.flush();
  lock.releaseLock();
  
  // 3. MENSAJE WHATSAPP
  let listaTexto = listaNumeros.join(", ");
  
  // Calculamos el total (asegúrate de que el precio sea 5000)
  let totalPagar = listaNumeros.length * 5000; 
  let totalFormato = totalPagar.toLocaleString('es-CO');

  // --- TU NUEVO MENSAJE EXACTO ---
  let mensaje = `Hola, soy *${nombre}*.\n\nAcabo de separar estos números en la web: *${listaTexto}*.\n\nEl total que debo transferir es: *$${totalFormato}*.\n\nPor favor envíame tu Nequi para realizar el pago.`;
  // -------------------------------
  
  let url = `https://wa.me/${MI_WHATSAPP}?text=${encodeURIComponent(mensaje)}`;
  
  return { exito: true, url: url };
}

// =========================================================
// === NUEVAS FUNCIONES PARA PAGOS Y TRANSFERENCIAS (OCR) ===
// =========================================================

// ¡IMPORTANTE! CAMBIA ESTO POR EL ID REAL DE TU CENTRAL DE TRANSFERENCIAS
const ID_CENTRAL_TRANSFERENCIAS = "1DtwLYhRE_3PN8Sl-5We6Qr9BBF54elBhGMQoGYwG28U"; // <--- PEGA AQUÍ EL ID

// Función auxiliar para normalizar texto (quitar espacios y raros)
function _normAlnum(s){ return String(s||"").toLowerCase().replace(/[^a-z0-9]/g,""); }

// Busca todas las hojas que se llamen "TRANSFERENCIAS..." en el archivo externo
function _getAllTransferSheets(){
  try {
    const ssExterna = SpreadsheetApp.openById(ID_CENTRAL_TRANSFERENCIAS);
    return ssExterna.getSheets().filter(s => /^TRANSFERENCIAS(\s*\d+)?$/i.test(s.getName()));
  } catch (e) {
    console.error("Error conectando a Central: " + e.message);
    return [];
  }
}

// Busca una transferencia por referencia en la Central
function buscarTransferenciaPorReferencia(ref){
  const needle = _normAlnum(ref);
  if(!needle) return {status: "error", mensaje: "Referencia vacía"};

  const sheets = _getAllTransferSheets();
  if (sheets.length === 0) return {status: "error", mensaje: "No se pudo conectar a la Central de Transferencias."};

  for (const sh of sheets){
    const last = sh.getLastRow();
    if (last < 2) continue;
    // Asumimos que la Central tiene estructura: 
    // A=Fecha, B=Plataforma, C=Monto, D=Referencia, ... G=Estado
    // Ajusta los índices si tu Central es diferente.
    // Aquí usamos indices base 0: Col D es indice 3.
    const vals = sh.getRange(2,1,last-1,7).getDisplayValues(); 
    
    for (let i=0; i<vals.length; i++){
      const referencia = vals[i][3]; // Columna D
      
      if (_normAlnum(referencia) === needle){
        const row = i+2;
        const plataforma = vals[i][1]; // Col B
        const montoStr   = vals[i][2]; // Col C
        const status     = String(vals[i][6]||"").trim(); // Col G (Estado)

        // Limpiar monto
        const monto = Number(String(montoStr||"").replace(/[^\d]/g,""))||0;

        return {
          status: "ok",
          encontrada: true,
          datos: {
            sheet: sh.getName(),
            row: row,
            referencia: referencia,
            plataforma: plataforma,
            monto: monto,
            estado: status,
            yaUsada: status.toLowerCase().startsWith("asignado")
          }
        };
      }
    }
  }
  return {status: "ok", encontrada: false};
}

// Marca la transferencia como ASIGNADA en la Central
function asignarTransferencia(sheetName, row, motivo){
  try {
    const ssExterna = SpreadsheetApp.openById(ID_CENTRAL_TRANSFERENCIAS);
    const sh = ssExterna.getSheetByName(sheetName);
    if (!sh) return {status: "error", mensaje: "Hoja no encontrada"};
    
    // Columna G (7) es el Estado
    sh.getRange(row, 7).setValue("Asignado - " + motivo);
    return {status: "ok"};
  } catch (e) {
    return {status: "error", mensaje: e.message};
  }
}

function obtenerListaNumeros() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  // Leemos 5 columnas (A, B, C, D, E)
  const datos = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

  return datos.map(fila => ({
    numero: fila[0],
    estado: fila[1], 
    nombre: fila[2] || "Sin nombre",
    telefono: String(fila[3] || ""), // <--- ¡AQUÍ ESTÁ LA CLAVE! Agregamos el teléfono
    pago:   fila[4]
  }));
}

// --- NUEVA FUNCIÓN: Registra pago en MÚLTIPLES boletas y divide el dinero ---
function registrarPagoMasivo(listaNumeros, referencia, montoTotal, metodo) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { exito: false, error: "Servidor ocupado." };
  }

  // --- NUEVA FUNCIÓN: Registra pago en MÚLTIPLES boletas y divide el dinero ---
function registrarPagoMasivo(listaNumeros, referencia, montoTotal, metodo) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { exito: false, error: "Servidor ocupado." };
  }

  // --- 🔒 VALIDACIÓN ESTRICTA DE PRECIO ($5.000) ---
  const cantidad = listaNumeros.length;
  if (cantidad === 0) {
     lock.releaseLock();
     return { exito: false, error: "No seleccionaste ninguna boleta." };
  }

  const valorCalculado = montoTotal / cantidad;

  if (valorCalculado !== 5000) {
     lock.releaseLock();
     // Devolvemos error si no da exacto
     return { 
       exito: false, 
       error: `⛔ ERROR CONTABLE:\n\nEl pago es de $${montoTotal} y seleccionaste ${cantidad} boletas.\nEso da $${valorCalculado} por boleta.\n\nSolo se permite abonar si el resultado es exactamente $5.000.` 
     };
  }
  // --------------------------------------------------

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA);
  const datos = sheet.getDataRange().getValues();
  
  // Como ya validamos que es 5000, usamos ese valor fijo
  const montoIndividual = 5000; 

  let boletasActualizadas = 0;

  // Recorremos los números que seleccionaste en la pantalla
  listaNumeros.forEach(numObjetivo => {
    // Buscamos la fila de cada número en el Excel
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][0]) == String(numObjetivo)) {
        const filaReal = i + 1;
        
        // --- ACTUALIZACIÓN EN EL EXCEL ---
        sheet.getRange(filaReal, 5).setValue("PAGADO");      // Columna E: Estado Pago
        sheet.getRange(filaReal, 6).setValue(referencia);    // Columna F: Referencia

        boletasActualizadas++;
        break; 
      }
    }
  });
  
  SpreadsheetApp.flush();
  lock.releaseLock();

  if (boletasActualizadas > 0) {
    return { exito: true, mensaje: `✅ ¡Perfecto! Se abonaron ${boletasActualizadas} boletas correctamente.` };
  } else {
    return { exito: false, error: "No se encontraron los números en el Excel." };
  }
}

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA);
  const datos = sheet.getDataRange().getValues(); 
  
  // DIVISIÓN DEL DINERO: Total / Cantidad de boletas
  const montoIndividual = listaNumeros.length > 0 ? (montoTotal / listaNumeros.length) : 0;

  let boletasActualizadas = 0;

  // Recorremos los números que seleccionaste en la pantalla
  listaNumeros.forEach(numObjetivo => {
    // Buscamos la fila de cada número en el Excel
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][0]) == String(numObjetivo)) {
        const filaReal = i + 1;
        
        // --- ACTUALIZACIÓN EN EL EXCEL ---
        sheet.getRange(filaReal, 5).setValue("PAGADO");      // Columna E: Estado
        sheet.getRange(filaReal, 6).setValue(referencia);    // Columna F: Referencia
        // sheet.getRange(filaReal, 7).setValue(montoIndividual); // Opcional: Columna G con valor dividido

        boletasActualizadas++;
        break; 
      }
    }
  });

  SpreadsheetApp.flush();
  lock.releaseLock();

  if (boletasActualizadas > 0) {
    return { exito: true, mensaje: `Se registraron ${boletasActualizadas} pagos. Valor por boleta: $${montoIndividual}` };
  } else {
    return { exito: false, error: "No se encontraron los números en el Excel." };
  }
}

// ======================================================
// === BLOQUE DE BÚSQUEDA POR FECHA (CORREGIDO) ===
// ======================================================

// 1. Función Principal: Busca en todas las hojas
function buscarTransferenciasPorFecha(fechaISO, hora12){
  try {
    const sheets = _getAllTransferSheets();
    if (sheets.length === 0) return { status:"error", mensaje: "No hay conexión con la Central (Revisa el ID)." };

    const out = [];
    
    for (const sh of sheets){
      const last = sh.getLastRow();
      if (last < 2) continue;
      
      // Traemos columnas A hasta G (Indices 0 a 6)
      const vals = sh.getRange(2,1,last-1,7).getDisplayValues();

      for (let i=0; i<vals.length; i++){
        // Datos crudos del Excel
        const fechaRaw = vals[i][4]; // Col E
        const horaRaw  = vals[i][5]; // Col F
        
        // Convertimos para comparar (con seguridad anti-errores)
        const iso = _fechaDispToISO(fechaRaw);
        const h12 = _normHora12(horaRaw);

        // COMPARACIÓN: Si coinciden fecha y hora
        if (iso === fechaISO && h12 === hora12){
          const status = String(vals[i][6]||"").trim(); // Col G
          const montoStr = vals[i][2]; // Col C
          const monto = Number(String(montoStr||"").replace(/[^\d]/g,""))||0;

          out.push({
            sheet: sh.getName(),
            row: i + 2, // Fila real
            plataforma: vals[i][1] || "Banco",
            monto: monto,
            referencia: vals[i][3],
            fecha: fechaRaw,
            hora: horaRaw,
            status: status,
            yaUsada: status.toLowerCase().startsWith("asignado")
          });
        }
      }
    }
    
    return { status:"ok", lista: out };

  } catch (e) {
    // Si falla, devolvemos el error exacto para que lo veas
    return { status:"error", mensaje: "Error interno: " + e.toString() };
  }
}

// 2. Ayudante: Convierte cualquier fecha loca a YYYY-MM-DD
function _fechaDispToISO(s){
  if (!s) return ""; // Si la celda está vacía, devuelve vacío
  if (s instanceof Date) { 
    const y=s.getFullYear(), m=("0"+(s.getMonth()+1)).slice(-2), d=("0"+s.getDate()).slice(-2); 
    return `${y}-${m}-${d}`; 
  }
  s = String(s).trim();
  // Intenta capturar dd/mm/yyyy o dd-mm-yyyy
  const m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/); 
  if(!m) return s; // Si ya es ISO o texto, lo devuelve tal cual
  
  const d=("0"+m[1]).slice(-2);
  const mo=("0"+m[2]).slice(-2);
  const y=m[3].length===2?("20"+m[3]):m[3];
  
  return `${y}-${mo}-${d}`;
}

// 3. Ayudante: Normaliza la hora a "09:30 AM"
function _normHora12(s){
  if (!s) return "";
  s = String(s).trim().toLowerCase().replace(/\./g,"").replace(/\s+/g," ");
  
  const ampm = s.includes("pm") ? "PM" : "AM";
  const m = s.match(/(\d{1,2})\s*:\s*(\d{2})/);
  
  if(!m) return "";
  
  let hh = ("0"+m[1]).slice(-2); 
  const mm = ("0"+m[2]).slice(-2);
  
  if(hh === "00") hh="12"; // Corregir medianoche
  
  return `${hh}:${mm} ${ampm}`;
}

function procesarVentas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const todasLasHojas = ss.getSheets();
  
  // Hojas que NO son de rifas (No las tocamos)
  const hojasIgnoradas = ["Index", "Abonos", "Config", "Instrucciones", "Central", "REPORTE_CLIENTES", "HISTORIAL"];

  let conteo = {};
  let nombres = {};
  let resumen = []; // Para el reporte chismoso

  todasLasHojas.forEach(hoja => {
    const nombreHoja = hoja.getName();
    
    // Si la hoja NO está en la lista de ignoradas, la revisamos
    if (!hojasIgnoradas.includes(nombreHoja)) {
      const lastRow = hoja.getLastRow();
      let ventasEnEstaHoja = 0;

      if (lastRow > 1) {
        // Leemos columnas A(0), B(1), C(2), D(3)
        // Asumimos que la columna D es TELÉFONO
        const datosHoja = hoja.getRange(2, 1, lastRow - 1, 4).getValues();

        datosHoja.forEach(fila => {
          // Limpiamos el teléfono (quitamos espacios y guiones)
          let telefono = String(fila[3]).replace(/\D/g, ''); 
          let nombreCliente = fila[2];

          // --- CAMBIO CLAVE AQUÍ ---
          // Ya no miramos si dice "Reservado".
          // Solo miramos si hay un teléfono válido (más de 6 dígitos).
          if (telefono.length > 6) {
            
            // Sumamos al cliente
            if (!conteo[telefono]) {
              conteo[telefono] = 0;
              nombres[telefono] = nombreCliente || "Sin Nombre";
            }
            conteo[telefono]++;
            
            // Sumamos al contador de la hoja
            ventasEnEstaHoja++;
          }
        });
      }
      // Guardamos cuántas encontró en esta hoja específica
      resumen.push({ nombreHoja: nombreHoja, cantidad: ventasEnEstaHoja });
    }
  });

  // Convertimos a lista
  let clientes = [];
  for (let tel in conteo) {
    clientes.push({ telefono: tel, nombre: nombres[tel], cantidad: conteo[tel] });
  }

  return { clientes: clientes, resumen: resumen };
}

// ==========================================
// 🖥️ INTERFAZ DE USUARIO (MENÚS Y ALERTAS)
// ==========================================

// 1. Crea el menú en la parte superior del Excel al abrirlo
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 REPORTES RIFA')
      .addItem('🏆 Ver Ganadores (Alerta Rápida)', 'auditarTodasLasHojas')
      .addItem('📑 Crear Tabla Completa de Clientes', 'generarTablaClientes')
      .addToUi();
}

// 2. Opción: Alerta Rápida (Solo muestra ganadores en pantalla)
function auditarTodasLasHojas() {
  const resultado = procesarVentas(); 
  const datos = resultado.clientes; 

  // Ordenamos de mayor a menor compras
  datos.sort((a, b) => b.cantidad - a.cantidad);
  
  let reporte = `🕵️ GANADORES (5+ Boletas)\n\n`;
  let hay = false;

  datos.forEach(c => {
    if (c.cantidad >= 5) {
      reporte += `👤 ${c.nombre}\n📱 ${c.telefono}\n🎫 ${c.cantidad} Boletas\n----------------\n`;
      hay = true;
    }
  });

  if (!hay) reporte += "Nadie tiene 5 o más compras todavía.";
  SpreadsheetApp.getUi().alert(reporte);
}

// 3. Opción: Crear Hoja de Excel (Reporte Completo con Resumen)
function generarTablaClientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultado = procesarVentas(); 
  const datos = resultado.clientes;   // La lista de gente
  const resumenHojas = resultado.resumen; // El chismoso de hojas

  // A. Preparar la hoja
  let hojaReporte = ss.getSheetByName("REPORTE_CLIENTES");
  if (!hojaReporte) {
    hojaReporte = ss.insertSheet("REPORTE_CLIENTES");
  } else {
    hojaReporte.clear();
  }

  // B. Encabezados Tabla Clientes
  hojaReporte.getRange("A1:E1").setValues([["POSICIÓN", "TELÉFONO", "NOMBRE CLIENTE", "TOTAL BOLETAS", "NIVEL"]]);
  hojaReporte.getRange("A1:E1").setFontWeight("bold").setBackground("#1b5e20").setFontColor("white");

  // Ordenar datos
  datos.sort((a, b) => b.cantidad - a.cantidad);

  let filasClientes = [];
  datos.forEach((c, index) => {
    let nivel = "Nuevo";
    if (c.cantidad >= 2) nivel = "Frecuente";
    if (c.cantidad >= 5) nivel = "🔥 SUPER VIP";

    filasClientes.push([index + 1, c.telefono, c.nombre, c.cantidad, nivel]);
  });

  if (filasClientes.length > 0) {
    hojaReporte.getRange(2, 1, filasClientes.length, 5).setValues(filasClientes);
  }

  // C. Encabezados Resumen (El Chismoso)
  hojaReporte.getRange("G1:H1").setValues([["HOJA ANALIZADA", "VENTAS ENCONTRADAS"]]);
  hojaReporte.getRange("G1:H1").setFontWeight("bold").setBackground("#f57c00").setFontColor("white");

  let filasResumen = [];
  let totalGeneral = 0;
  resumenHojas.forEach(h => {
    filasResumen.push([h.nombreHoja, h.cantidad]);
    totalGeneral += h.cantidad;
  });
  
  filasResumen.push(["TOTAL REAL:", totalGeneral]);

  if (filasResumen.length > 0) {
    hojaReporte.getRange(2, 7, filasResumen.length, 2).setValues(filasResumen);
  }

  hojaReporte.autoResizeColumns(1, 8);
  SpreadsheetApp.getUi().alert("✅ ¡Tabla Generada!\n\nRevisa la hoja 'REPORTE_CLIENTES'.\nEn las columnas G y H verás el resumen de hojas sumadas.");
}