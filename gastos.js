//im working in a google sheets document that i use to register my expenses and incomes.
//im using apps script to do functions.
//i have a sheet called "Movimientos" with the first column called "fecha", second column called "Descripción", third column called "importe", fourth column called "saldo"
//fifth column called "Categoría", sixth column called "Subcategoría" and seventh column called "Medio de pago".
//then i'll have another sheet called "Importar" with column 1 called "Fecha", column 2 called "Nro. transaccion", column 3 called "Descripción", 
//column 4 called "Importe" and column 4 called "Saldo"
//i want to compare the "Movimientos" sheet with the "Importar" sheet and see if there are any differences between them,
//if there are differences i want to add the new rows to the "Movimientos" sheet.
//to see if a row is new i'll compare the "Saldo" column of the "Movimientos" sheet with the "Saldo" column of the "Importar" sheet.
//also the "medio de pago" column will be "Macro" for all the new rows. But in "Movimientos" sheet exists registers with other payment methods.
//create function calld "ImportarMacro" that will do this.

const nombreHojaConfiguracion = 'Configuración';
const nombreHojaMovimientos = 'Movimientos';
const nombreHojaImportar = 'Importar';

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Importar datos')
   .addItem('Importar Macro', 'importarMacro')
  .addToUi();
}

function onEdit() {
  //FUNCION QUE CARGA LOS COMBOS DE LA HOJA MOVIMIENTOS

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var activeCell = ss.getActiveCell();
  
  if((activeCell.getColumn() == 5 && activeCell.getRow() > 1 && ss.getSheetName() == nombreHojaMovimientos)
  || (activeCell.getColumn() == 14 && activeCell.getRow() > 1 && ss.getSheetName() == nombreHojaConfiguracion)) {
    actualizarSubCategoria(activeCell);
  }
}

function importarMacro(){
    //FUNCION QUE IMPORTA EXCEL DE MACRO
    var hojaImportar = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaImportar);
    var hojaMovimientos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaMovimientos);
    var hConf = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaConfiguracion);
  
    var movimientos = hojaMovimientos.getRange("G1:G").getValues();   //COLUMNA "MEDIO DE PAGO" - HOJA MOVIMIENTOS.
    var ultMov = movimientos.filter(String).length;                   //CANTIDAD DE REGISTROS CON DATOS.
    var ultMovMacro = 0;
  
    //FORMATEO LAS COLUMNAS "Importe" Y "Saldos" DE LA HOJA "Importar" PARA QUE QUEDEN COMO NUMÉRICAS YA QUE VIENEN EN STRING.
    hojaImportar.getRange(1, 4, 203, 2).createTextFinder(" ").replaceAllWith("");
    hojaImportar.getRange(1, 4, 203, 2).createTextFinder(".").replaceAllWith("");
    hojaImportar.getRange(1, 4, 203, 2).createTextFinder("$").replaceAllWith("");
    
    var importar = hojaImportar.getRange("E1:E203").getValues(); //LISTA DE SALDOS DE LA HOJA IMPORTAR PARA DESPUES VER CUÁL ES EL ÚLTIMO QUE TENGO CARGADO EN MOVIMIENTOS Y CARGAR LOS QUE FALTAN.
  
    //DE LA HOJA MOVIMIENTOS, OBTENGO EL ÚLTIMO MOVIMIENTO DE MACRO
    for(var i = ultMov ; i > 0 ; i--) {
      if (movimientos[i][0] == 'Macro') {
        ultMovMacro = i + 1
        break;
      }
    }
  
    var ultimoSaldo = hojaMovimientos.getRange(ultMovMacro, 4).getValue(); //ÚLTIMO SALDO DE MACRO CARGADO EN MOVIMIENTOS.
    var ultMovImporte = 0;
  
    //DE LA HOJA "Importar" BUSCO CUÁL ES LA ÚLTIMA FILA QUE CARGUÉ A MOVIMIENTOS PARA DESCARTAR EL RESTO.
    for(var i = 1 ; i <= 203 ; i++) {
      if (importar[i][0] == ultimoSaldo) {
        ultMovImporte = i
        break;
      }
    }
  
    //RECORRO LA HOJA IMPORTAR DE ATRAS PARA ADELANTE Y VOY COPIANDO LA DATA QUE FALTA A LA HOJA MOVIMIENTOS
    for(var i = ultMovImporte ; i >= 4 ; i--) {
       ultMov++;    
       var registroImportar = hojaImportar.getRange(i, 1, 1, 5).getValues();  //OBTENGO LA FILTA DE LA HOJA IMPORTAR
       hojaMovimientos.getRange(ultMov, 1, 1, 1).setValue(registroImportar[0][0]); //FECHA
       hojaMovimientos.getRange(ultMov, 2, 1, 1).setValue(registroImportar[0][2]); //DESCRIPCIÓN
       hojaMovimientos.getRange(ultMov, 3, 1, 1).setValue(registroImportar[0][3]); //IMPORTE
       hojaMovimientos.getRange(ultMov, 4, 1, 1).setValue(registroImportar[0][4]); //SALDO
       //ACÁ AGREGAR LAS CATEGORÍAS
       cargarCategorias(hojaMovimientos, hConf, ultMov, registroImportar[0][2])
       hojaMovimientos.getRange(ultMov, 7, 1, 1).setValue("Macro");
    }
  }

  function cargarCategorias(i_h_moviomiento, i_h_config, i_n_movimiento, i_d_nuevo_mov,){
    var colNombre = i_h_config.getRange("K3:K36").getValues();      //COLUMNA "NOMBRE" - HOJA CONFIGURACIÓN.
    var kConf = colNombre.filter(String).length;                    //CANTIDAD DE REGISTROS CON DATOS EN LA HOJA CONFIGURACIÓN.
    var colCuil = i_h_config.getRange("L3:L36").getValues();        //COLUMNA "CUIL" - HOJA CONFIGURACIÓN.
    var colCatConf = i_h_config.getRange("N3:N36").getValues();     //COLUMNA "CATEGORÍA" - HOJA CONFIGURACIÓN.
    var colCatSecConf = i_h_config.getRange("O3:O36").getValues();  //COLUMNA "CATEGORÍA" - HOJA CONFIGURACIÓN.
    var existe = -1;
  
    for (var k = 0; k <= kConf - 1; k++){ //RECORRO LA LISTA DE LA HOJA CONFIGURACIÓN
      var nombreActual = colNombre[k][0];
      var cuilActual = colCuil[k][0];
  
      //PRIMERO COMPARO LA DESCRIPCION QUE VIENE POR PARÁMETRO CON LA COLUMNA "NOMBRE" DE LA HOJA CONFIGURACIÓN.
      if(i_d_nuevo_mov.indexOf(nombreActual) > -1) {
        existe = k;
        break;
      }
  
      //SEGUNDO VERIFICO SI ALGUNO DE LOS CUILS DE LA HOJA CONFIGURACIÓN ESTÁ EN LA DESCRIPCIÓN QUE VIENE POR PARÁMETRO.
      if (i_d_nuevo_mov.indexOf(cuilActual) > -1){
        existe = k;
        break;
      }
      
      //TERCERO VERIFICO SI ALGUNO DE LOS NOMBRES DE LA HOJA CONFIGURACIÓN ESTÁ EN LA DESCRIPCIÓN QUE VIENE POR PARÁMETRO.
      if(nombreActual.indexOf(i_d_nuevo_mov) > -1) {
        existe = k;
        break;
      }
    }
  
    if (existe >= 0) {
      var celdaActual = i_h_moviomiento.getRange(i_n_movimiento, 5);
      celdaActual.setValue(colCatConf[existe][0]) //CATEGORÍA
      actualizarSubCategoria(celdaActual);
      celdaActual.offset(0, 1).setValue(colCatSecConf[existe][0]); //SUBCATEGORÍA
    }
  }
  
  function actualizarSubCategoria(celdaActual){
    var hojaConfiguracion = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaConfiguracion);
    
    celdaActual.offset(0, 1).clearContent().clearDataValidations();
    var categorias = hojaConfiguracion.getRange(1, 1, 1, 7).getValues(); //getRange(fila inicio, columna inicio, número de filas, número de columnas)
    var categoriaID = categorias[0].indexOf(celdaActual.getValue()) + 1;
    
    if(categoriaID != 0) {
      var validationRange = hojaConfiguracion.getRange(2, categoriaID, 25);
      var validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
      celdaActual.offset(0, 1).setDataValidation(validationRule);
    }
  }