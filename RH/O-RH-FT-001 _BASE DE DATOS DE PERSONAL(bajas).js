function pruebaBajas() {//para pruebas.
  bajas(169); // pon aquí una fila real
}

// al dar bajas se mandara a la hoja TOTAL BAJAS
function bajas(fila) {
//function bajas() {
  try {
    var ssOrigen = SpreadsheetApp.getActiveSpreadsheet();
    var hojaOrigen = ssOrigen.getSheetByName("RESTRUCTURACION");
    if (!hojaOrigen) {
      throw new Error("No se encontró la hoja 'RESTRUCTURACION' en el archivo baja.");
    }

    // Validar fila
    if (!fila || fila < 1) {
      throw new Error("Número de fila inválido: " + fila);
    }

    // Leer toda la fila del origen
    var numCols = hojaOrigen.getLastColumn();
    var valoresFila = hojaOrigen.getRange(fila, 1, 1, numCols).getValues()[0];
    var richValores = hojaOrigen.getRange(fila, 1, 1, numCols).getRichTextValues()[0];
    //var valoresFila = hojaOrigen.getRange(9, 1, 1, numCols).getValues()[0];

    // Mapear columnas (base 0 en el array)
    var valorNombre = valoresFila[3];   // Columna D //nombre//15/04/2026
    var valorDEPARTAMENTO = valoresFila[5];   // Columna I //estado civil
    var valorPUESTO = valoresFila[6];    // Columna A //area = departamento
    var valorFechaIN = valoresFila[9];   // Columna B //puesto

    //agegar fecha de hoy para la posicion de baja destino. FECHA DE BAJA

    var valorNivOficina = valoresFila[0];   // Columna N //nivel de estudios

    var valorNo = valoresFila[1];   // Columna S //FAMILIAR INTERNO 1
    var valorNivelEstudios = valoresFila[2];   // Columna T //PARENTESCO 1
    var valorEstadoCivil = valoresFila[4];   // Columna U //ÁREA 1
    var valorLlamadas  = valoresFila[11];   // Columna U //ÁREA 1
    var valorWHATSAPP  = valoresFila[12];   // Columna U //ÁREA 1
    var valorCorreo  = valoresFila[13];   // Columna U //ÁREA 1
    var valorComputadora  = valoresFila[14];   // Columna U //ÁREA 1
    var valorCelular  = valoresFila[15];   // Columna U //ÁREA 1
    var valorLAPTOP  = valoresFila[16];   // Columna U //ÁREA 1
    var valorCARRO   = valoresFila[17];   // Columna U //ÁREA 1
    var valorFOTOINST  = valoresFila[18];   // Columna U //ÁREA 1
    var valorFECHACUMPL  = valoresFila[19];   // Columna U //ÁREA 1
    var valorMES  = valoresFila[20];   // Columna U //ÁREA 1
    var valorHIJOS  = valoresFila[21];   // Columna U //ÁREA 1
    var valorHIJOS1  = valoresFila[22];   // Columna U //ÁREA 1
    var valorHIJOS2  = valoresFila[23];   // Columna U //ÁREA 1
    var valorHIJOS3  = valoresFila[24];   // Columna U //ÁREA 1
    var valorHIJOS4  = valoresFila[25];   // Columna U //ÁREA 1
    var valorHIJOS5  = valoresFila[26];   // Columna U //ÁREA 1
    var valorHIJOS6  = valoresFila[27];   // Columna U //ÁREA 1
    var valorRECOMENDADO  = valoresFila[28];   // Columna U //ÁREA 1
 
    var valorFaIn1 = valoresFila[29];   // Columna V //FAMILIAR INTERNO 2
    var valorPar1 = valoresFila[30];   // Columna W //PARENTESCO 2
    var valorArea1 = valoresFila[31];   // Columna X //ÁREA 2
    var valorFaIn2 = valoresFila[32];   // Columna V //FAMILIAR INTERNO 2
    var valorPar2 = valoresFila[33];   // Columna W //PARENTESCO 2
    var valorArea2 = valoresFila[34];   // Columna X //ÁREA 2
    var valorFaIn3 = valoresFila[35];   // Columna Y //FAMILIAR INTERNO 3
    var valorPar3 = valoresFila[36];   // Columna Z //PARENTESCO 3
    var valorArea3 = valoresFila[37];   // Columna AA //ÁREA 3
    
    var valorSOCIECONOMICO   = valoresFila[38];   // Columna U //ÁREA 1
    //var valorBATERIALIDERAZGO  = valoresFila[39];   // Columna U //ÁREA 1

    //var valorLinkPsic = valoresFila[40];   // Columna BC //LINK PSICOMETRICOS modv 18/12/2025
    var valorULTIMAFECHACTUALIZADA1  = valoresFila[41];   // Columna U //ÁREA 1
    //var valorLinkMED = valoresFila[42];   // Columna BD //LINK MEDICOS modv 18/12/2025
    var valorULTIMAFECHACTUALIZADA2  = valoresFila[43];   // Columna U //ÁREA 1
    //var valorLinkPOLI = valoresFila[44];   // Columna BE //LINK POLIGRAFOS modv 18/12/2025
    var valorULTIMAFECHACTUALIZADA3  = valoresFila[45];   // Columna U //ÁREA 1

    //nueva 18/12/2025
    var valor15Dias = valoresFila[46];   // Columna BL //PRUEBA 15 DÍAS 
    var valorLISTOWORKY  = valoresFila[47];   // Columna U //ÁREA 1
    var valorGRUPOS  = valoresFila[48];   // Columna U //ÁREA 1
    var valorNUMERO  = valoresFila[49];   // Columna U //ÁREA 1
    var valorCOMENTARIOS  = valoresFila[50];   // Columna U //ÁREA 1
    var valorFECHAALTAIMSS  = valoresFila[51];   // Columna U //ÁREA 1
    var valorEMPRESA  = valoresFila[52];   // Columna U //ÁREA 1
    var valorEMPRESARFCEMPRESA  = valoresFila[53];   // Columna U //ÁREA 1
    var valorALTAIMSS  = valoresFila[54];   // Columna U //ÁREA 1
    var valorCONTRATODEFINITIVO  = valoresFila[55];   // Columna U //ÁREA 1
    var valorINFONAVIT  = valoresFila[56];   // Columna U //ÁREA 1
    var valorBANCO  = valoresFila[57];   // Columna U //ÁREA 1
    var valorCUENTACLABETARJETA  = valoresFila[58];   // Columna U //ÁREA 1
    var valorNUMEROTARJETA  = valoresFila[59];   // Columna U //ÁREA 1
    var valorCLAVEINTERBANCARIA  = valoresFila[60];   // Columna U //ÁREA 1
    var valorDIASVACACIONES  = valoresFila[61];   // Columna U //ÁREA 1
    
    //var valorAlta = valoresFila[65];    // Columna BN (66) //onEdit modv 18/12/2025
    //var valorBaja = valoresFila[10];    // Columna BN (66) //onEdit modv 18/12/2025

    // Solo continuar si hay algo en la columna AZ
    //if (String(valorActivo).trim().toLocaleUpperCase() !== "SI") {
    /*if (String(valorBaja).trim().toLocaleUpperCase() !== "BAJA") {
      Logger.log("La columna BO está vacía en la fila " + fila + ". No se procesó.");
      return;
    }*/

    // Verificar que haya datos relevantes
    if ((valorNombre === "" || valorNombre === null) &&
        (valorDEPARTAMENTO=== "" || valorDEPARTAMENTO === null) &&
        (valorPUESTO === "" || valorPUESTO === null) &&
        (valorFechaIN === "" || valorFechaIN === null) 
        /*(valorNivOficina === "" || valorNivOficina === null) &&
        (valorNo === "" || valorNo === null) &&
        (valorNivelEstudios === "" || valorNivelEstudios === null) &&
        (valorEstadoCivil === "" || valorEstadoCivil === null) &&
        (valorLlamadas === "" || valorLlamadas === null) &&
        (valorWHATSAPP === "" || valorWHATSAPP === null) &&
        (valorCorreo === "" || valorCorreo === null) &&
        (valorComputadora === "" || valorComputadora === null) &&
        (valorCelular === "" || valorCelular === null) &&
        (valorLAPTOP === "" || valorLAPTOP === null) &&
        (valorCARRO === "" || valorCARRO === null) &&
        (valorFOTOINST === "" || valorFOTOINST === null) &&
        (valorFECHACUMPL === "" || valorFECHACUMPL === null) &&
        (valorMES === "" || valorMES === null) &&
        (valorHIJOS === "" || valorHIJOS === null) &&
        (valorHIJOS1 === "" || valorHIJOS1 === null) &&
        (valorHIJOS2 === "" || valorHIJOS2 === null) &&
        (valorHIJOS3 === "" || valorHIJOS3 === null) &&
        (valorHIJOS4 === "" || valorHIJOS4 === null) &&
        (valorHIJOS5 === "" || valorHIJOS5 === null) &&
        (valorHIJOS6 === "" || valorHIJOS6 === null) &&
        (valorRECOMENDADO === "" || valorRECOMENDADO === null) &&
        
        
        (valorFaIn1 === "" || valorFaIn1 === null) &&
        (valorPar1 === "" || valorPar1 === null) &&
        (valorArea1 === "" || valorArea1 === null) &&
        (valorFaIn2 === "" || valorFaIn2 === null) &&
        (valorPar2 === "" || valorPar2 === null) &&
        (valorArea2 === "" || valorArea2 === null) &&
        (valorFaIn3 === "" || valorFaIn3 === null) &&
        (valorPar3 === "" || valorPar3 === null) &&
        (valorArea3 === "" || valorArea3 === null) &&
        
        (valorSOCIECONOMICO === "" || valorSOCIECONOMICO === null) &&
        (valorBATERIALIDERAZGO === "" || valorBATERIALIDERAZGO === null) &&
        
        (valorLinkPsic === "" || valorLinkPsic === null) &&
        (valorULTIMAFECHACTUALIZADA1 === "" || valorULTIMAFECHACTUALIZADA1 === null) &&
        (valorLinkMED === "" || valorLinkMED === null) &&
        (valorULTIMAFECHACTUALIZADA2 === "" || valorULTIMAFECHACTUALIZADA2 === null) &&
        (valorLinkPOLI === "" || valorLinkPOLI === null) &&
        (valorULTIMAFECHACTUALIZADA3 === "" || valorULTIMAFECHACTUALIZADA3 === null) &&
        (valor15Dias === "" || valor15Dias === null)  &&
        
        (valorLISTOWORKY === "" || valorLISTOWORKY === null) &&
        (valorGRUPOS === "" || valorGRUPOS === null) &&
        (valorNUMERO === "" || valorNUMERO === null) &&
        (valorCOMENTARIOS === "" || valorCOMENTARIOS === null) &&
        (valorFECHAALTAIMSS === "" || valorFECHAALTAIMSS === null) &&
        (valorEMPRESARFCEMPRESA === "" || valorEMPRESARFCEMPRESA === null) &&
        (valorALTAIMSS === "" || valorALTAIMSS === null) &&
        (valorCONTRATODEFINITIVO === "" || valorCONTRATODEFINITIVO === null) &&
        (valorINFONAVIT === "" || valorINFONAVIT === null) &&
        (valorBANCO === "" || valorBANCO === null) &&
        (valorCUENTACLABETARJETA === "" || valorCUENTACLABETARJETA === null) &&
        (valorNUMEROTARJETA === "" || valorNUMEROTARJETA === null) &&
        (valorCLAVEINTERBANCARIA === "" || valorCLAVEINTERBANCARIA === null) &&
        (valorDIASVACACIONES === "" || valorDIASVACACIONES === null) &&*/
        //(valorBaja === "" || valorBaja === null)
        ) {
      Logger.log("No hay datos en Nombre/Puesto/Depto en la fila " + fila + ". No se procesó.");
      return;
    }

    // Preparar datos a enviar
    var fechaHoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yy"); //HojaDestino 9

    // Abrir libro destino por ID y obtener hoja destino
    var destinoId = '1wJwex0jcy6LHvuZVLQmAUi-B3oBM6qcm9eOSXFXzYsk'; // <- tu ID
    var libroDestino = SpreadsheetApp.openById(destinoId);
    var hojaDestino = libroDestino.getSheetByName("TOTAL BAJAS");
    if (!hojaDestino) {
      throw new Error("No se encontró la hoja 'Hoja 10' en el libro destino.");
    }

    // Calcular siguiente fila vacía y escribir datos (sin tocar columna E)
    var filaDestino = obtenerUltimaFila(hojaDestino, 1, 2);//A2
    //var filaDestino = obtenerUltimaFila(hojaDestino, 3, 4);
    var filaD = filaDestino + 1;
    hojaDestino.getRange(filaD, 1).setValue(valorNombre);//columna C
    hojaDestino.getRange(filaD, 2).setValue(valorDEPARTAMENTO);//columna C
    hojaDestino.getRange(filaD, 3).setValue(valorPUESTO);//columna D
    hojaDestino.getRange(filaD, 4).setValue(valorFechaIN);//columna E
    
    //reacciona cuando uno pone activo en la posicon 10
    //var celdaEstado = hojaDestino.getRange(filaD, 10)//activo
    hojaDestino.getRange(filaD, 5).setValue(fechaHoy);
    
    /*var celdaFecha = hojaDestino.getRange(filaD, 5)//fecha
    if (celdaFecha.getValue() === "") {
       celdaFecha.setValue(fechaHoy);
       Logger.log("celda poner: " + celdaFecha);
    }*/
    
    hojaDestino.getRange(filaD, 12).setValue(valorNivOficina);//columna F
    hojaDestino.getRange(filaD, 13).setValue(valorNo);//columna F
    hojaDestino.getRange(filaD, 14).setValue(valorNivelEstudios);//columna F
    hojaDestino.getRange(filaD, 15).setValue(valorEstadoCivil);//columna F
    hojaDestino.getRange(filaD, 16).setValue(valorLlamadas);//columna F
    hojaDestino.getRange(filaD, 17).setValue(valorWHATSAPP);//columna F
    hojaDestino.getRange(filaD, 18).setValue(valorCorreo);//columna F
    hojaDestino.getRange(filaD, 19).setValue(valorComputadora);//columna F
    hojaDestino.getRange(filaD, 20).setValue(valorCelular);//columna F
    hojaDestino.getRange(filaD, 21).setValue(valorLAPTOP);//columna F
    hojaDestino.getRange(filaD, 22).setValue(valorCARRO);//columna F
    hojaDestino.getRange(filaD, 23).setValue(valorFOTOINST);//columna F
    hojaDestino.getRange(filaD, 24).setValue(valorFECHACUMPL);//columna F
    hojaDestino.getRange(filaD, 25).setValue(valorMES);//columna F
    hojaDestino.getRange(filaD, 26).setValue(valorHIJOS);//columna F
    hojaDestino.getRange(filaD, 27).setValue(valorHIJOS1);//columna F
    hojaDestino.getRange(filaD, 28).setValue(valorHIJOS2);//columna F
    hojaDestino.getRange(filaD, 29).setValue(valorHIJOS3);//columna F
    hojaDestino.getRange(filaD, 30).setValue(valorHIJOS4);//columna F
    hojaDestino.getRange(filaD, 31).setValue(valorHIJOS5);//columna F
    hojaDestino.getRange(filaD, 32).setValue(valorHIJOS6);//columna F
    hojaDestino.getRange(filaD, 33).setValue(valorRECOMENDADO);//columna F

    hojaDestino.getRange(filaD, 34).setValue(valorFaIn1);//columna AC
    hojaDestino.getRange(filaD, 35).setValue(valorPar1);//columna AD
    hojaDestino.getRange(filaD, 36).setValue(valorArea1);//columna AE
    hojaDestino.getRange(filaD, 37).setValue(valorFaIn2);//columna AF
    hojaDestino.getRange(filaD, 38).setValue(valorPar2);//columna AG
    hojaDestino.getRange(filaD, 39).setValue(valorArea2);//columna AH
    hojaDestino.getRange(filaD, 40).setValue(valorFaIn3);//columna AI
    hojaDestino.getRange(filaD, 41).setValue(valorPar3);//columna AJ
    hojaDestino.getRange(filaD, 42).setValue(valorArea3);//columna AK
    
    hojaDestino.getRange(filaD, 43).setValue(valorSOCIECONOMICO);//columna AK
    //hojaDestino.getRange(filaD, 44).setValue(valorBATERIALIDERAZGO);//columna AK

    //hojaDestino.getRange(filaD, 45).setValue(valorLinkPsic);//columna AO
    hojaOrigen.getRange(fila, 40).copyTo(hojaDestino.getRange(filaD, 44),{contentsOnly: false});
    hojaOrigen.getRange(fila, 41).copyTo(hojaDestino.getRange(filaD, 45),{contentsOnly: false});
    hojaOrigen.getRange(fila, 43).copyTo(hojaDestino.getRange(filaD, 47),{contentsOnly: false});
    hojaOrigen.getRange(fila, 45).copyTo(hojaDestino.getRange(filaD, 49),{contentsOnly: false});
    //hojaDestino.getRange(filaD, 45).setRichTextValue(SpreadsheetApp.newRichTextValue().setText("PDF").setLinkUrl(valorLinkPsic).build());
    hojaDestino.getRange(filaD, 46).setValue(valorULTIMAFECHACTUALIZADA1);//columna AO
    //hojaDestino.getRange(filaD, 47).setValue(valorLinkMED);//columna AQ
    //hojaDestino.getRange(filaD, 45).setRichTextValue(SpreadsheetApp.newRichTextValue().setText("PDF").setLinkUrl(valorLinkMED).build());
    hojaDestino.getRange(filaD, 48).setValue(valorULTIMAFECHACTUALIZADA2);//columna AQ
    //hojaDestino.getRange(filaD, 49).setValue(valorLinkPOLI);//columna AS
    //hojaDestino.getRange(filaD, 45).setRichTextValue(SpreadsheetApp.newRichTextValue().setText("PDF").setLinkUrl(valorLinkPOLI).build());
    hojaDestino.getRange(filaD, 50).setValue(valorULTIMAFECHACTUALIZADA3);//columna AS
    hojaDestino.getRange(filaD, 51).setValue(valor15Dias);//NUEVO Col. AU

    hojaDestino.getRange(filaD, 52).setValue(valorLISTOWORKY);//columna AK
    hojaDestino.getRange(filaD, 53).setValue(valorGRUPOS);//columna AK
    hojaDestino.getRange(filaD, 54).setValue(valorNUMERO);//columna AK
    hojaDestino.getRange(filaD, 55).setValue(valorCOMENTARIOS);//columna AK
    hojaDestino.getRange(filaD, 56).setValue(valorFECHAALTAIMSS);//columna AK
    hojaDestino.getRange(filaD, 57).setValue(valorEMPRESA);//columna AK
    hojaDestino.getRange(filaD, 58).setValue(valorEMPRESARFCEMPRESA);//columna AK
    hojaDestino.getRange(filaD, 59).setValue(valorALTAIMSS);//columna AK
    hojaDestino.getRange(filaD, 60).setValue(valorCONTRATODEFINITIVO);//columna AK
    hojaDestino.getRange(filaD, 61).setValue(valorINFONAVIT);//columna AK
    hojaDestino.getRange(filaD, 62).setValue(valorBANCO);//columna AK
    hojaDestino.getRange(filaD, 63).setValue(valorCUENTACLABETARJETA);//columna AK
    hojaDestino.getRange(filaD, 64).setValue(valorNUMEROTARJETA);//columna AK
    hojaDestino.getRange(filaD, 65).setValue(valorCLAVEINTERBANCARIA);//columna AK
    hojaDestino.getRange(filaD, 66).setValue(valorDIASVACACIONES);//columna AK
    //hojaDestino.getRange(filaD, 1, 1, hojaDestino.getLastColumn()).setBorder(false, false, false, false, false, false);//para quitar las lineas
    /*
      Parámetro	Significado
          filaD...............	fila donde pegaste
          1...................	desde columna A
          1...................	solo una fila
          getLastColumn().....	TODAS las columnas
     */

    Logger.log("✅ Fila " + fila + " procesada. Datos enviados a fila destino " + filaD + ": " + valoresFila.join(", "));
  } catch (err) {
    Logger.log("❌ Error en altasFila: " + err);
    throw err; // opcional: lanzar para ver el error al ejecutar manualmente
  }
}

function activadorAltas(e) {
  try {
    if (!e) return;

    const hoja = e.range.getSheet();
    const fila = e.range.getRow();
    const columna = e.range.getColumn();
    const valor = String(e.value || "").trim().toUpperCase();

    // 1️⃣ SOLO hoja RESTRUCTURACION
    if (hoja.getName() !== "RESTRUCTURACION") return;

    // 2️⃣ SOLO columna K (10)
    if (columna !== 11) return;

    Logger.log("Columna detectada: " + columna);

    // 3️⃣ SOLO cuando sea "BAJA"
    if (valor !== "BAJA") return;

    // 4️⃣ Lock para evitar dobles ejecuciones
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(2000)) return;

    bajas(fila);

    lock.releaseLock();

  } catch (err) {
    Logger.log("❌ Error BAJA: " + err);
  }
}


// Función para obtener la última fila con datos en una columna específica
function obtenerUltimaFila(hoja, columna, filaInicio) {
  const datos = hoja.getRange(filaInicio, columna, hoja.getLastRow() - filaInicio + 1).getValues();

  for (let i = datos.length - 1; i >= 0; i--) {
    if (datos[i][0] !== "") { // Verifica si hay datos en la celda
      return filaInicio + i;
    }
  }

  return filaInicio - 1; // Retorna filaInicio - 1 si no hay datos en la columna
}
