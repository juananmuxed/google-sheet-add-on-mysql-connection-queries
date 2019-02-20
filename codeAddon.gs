/* 

Created By MuXeD
Creative Commons

Attribution-NonCommercial 4.0 International (CC BY-NC 4.0) 
https://creativecommons.org/licenses/by-nc/4.0/

----------------------------------------------------------

Spanish version for menus, you can change as you wish

*/

//Connection Default Preferences
var connectionName = 'server.com'; //Ip or direction
var user = 'user'; // User
var userPwd = 'pass'; // Password
var db = 'database'; // Database
// This type of connection can change, this is configurated in MySQL, and change getConnection() function
var dbUrl = 'jdbc:mysql://' + connectionName + '/' + db;

// General Vars 
var ss = SpreadsheetApp.getActiveSpreadsheet();

// You can change the name of the Query Data Tab
var ssmelena = 'MELENA';
var choja = ss.getSheetByName(ssmelena);
var ui = SpreadsheetApp.getUi();

// If you like, onOpen() this function or active with a Triger it as you prefere 
function menuOpen(){
  
  ui.createMenu('La magia de mi melena')
  .addItem('Mírala','mira')
  .addItem('Probar Conexión por Defecto', 'testing')
  .addSeparator()
  .addItem('Crear Query','query')
  .addSeparator()
  .addItem('Consultas','consultas')
  .addToUi();
  
}

// Create Tab to archive Queries configs
function mira(){
  if (!choja) {
    ss.insertSheet(ssmelena);
    choja = ss.getSheetByName(ssmelena);
    choja.setFrozenRows(1);
    // Spanish
    choja.appendRow(["Función","Nombre","Menu","Query","Tabla","Última Actualización","Comentario","Tiempo Usado","Filas Máximas","Fecha Creación"]);
    choja.getRange('1:1').setBackground('#333333').setFontColor('#eeeeee').setFontWeight('bold');
    var titulo = "Hoja creada";
    var mensajeError = "Mira la magia de mi melena. Apenas he tenido contacto con los estupefacientes.";
    aviso(titulo,mensajeError);
    SpreadsheetApp.setActiveSheet(ss.getSheetByName(ssmelena));
  }
  else{
    var titulo = "La hoja ya existe";
    var mensajeError = "La hoja ya existe, pero la magia de mi melena me perturba. Ya te llevo yo, apenas he tenido contacto con los estupefacientes.";
    aviso(titulo,mensajeError);
    SpreadsheetApp.setActiveSheet(ss.getSheetByName(ssmelena));
  }
}

// Test the conection
function testing(){
  var pruebadefe = Jdbc.getConnection(dbUrl, user, userPwd).isValid(1000);
  if(pruebadefe){
    var titulo = "Conexión exitosa";
    var mensajeCorrecto = "La conexión está configurada correctamente.";
    aviso(titulo,mensajeCorrecto);
  }
}

// Generate a Modal Popup with content and AutoClose 5 Seconds
function aviso(titulo,aviso){
  var contenido = '<div class="block"><b>'+aviso+'</b></div></br>';
  var html = HtmlService.createHtmlOutputFromFile('TemplateAvisos').getContent();
  contenido+=html;
  var final = HtmlService.createHtmlOutputFromFile('TemplateAvisos').clear().append(contenido).setWidth(400).setHeight(200);
  ui.showModalDialog(final, titulo);
}

// Generate a Modal Popup
function modalrefrescar(titulo,aviso){  
  var final = HtmlService.createHtmlOutputFromFile('TemplateRefresh').append(aviso).setWidth(400).setHeight(200);
  ui.showModalDialog(final, titulo);
}

// Sidebar for create Query
function query(){
  var html = HtmlService.createHtmlOutputFromFile('TemplateSideb')
      .setTitle('Introduce datos de la Query')
      .setWidth(300);
  ui.showSidebar(html);
}

// Need to create a list to Refresh
function seleccion(){
  var ultichan = choja.getLastRow();
  var rangochan = choja.getRange(2, 1, ultichan-1, 1).getValues();
  var listadorango = [];
  for (var col = 0;col < ultichan-1;col++) {
       listadorango.push(rangochan[col]);
  }
  return listadorango;
}

// Sidebar for Refresh
function consultas(){
  var html = HtmlService.createHtmlOutputFromFile('TemplateConsultas')
      .setTitle('Consultas Guardadas')
      .setWidth(300);
  ui.showSidebar(html);
}

// Creating a Tab for a Query and the first Refresh
function creaciontablas(valores){
  var cdate = new Date();
  var nombretab = valores[1];
  var nombre = valores[0];
  var consulta = valores[2];
  var func = valores[0];
  var filasm = valores[3];
  var conn = Jdbc.getConnection(dbUrl, user, userPwd);
  var statm = conn.createStatement();
  statm.setMaxRows(filasm);
  var resultados = statm.executeQuery(consulta);
  var tabtemp = ss.getSheetByName(nombretab);
  if(!tabtemp && resultados != null){
    ss.getSheetByName(ssmelena).appendRow([func,nombre,nombre,consulta,nombretab,"","","",filasm,cdate]);
    ss.insertSheet(nombretab).getRange('1:1').setBackground('#333333').setFontColor('#eeeeee').setFontWeight('bold');
    var tabtemp = ss.getSheetByName(nombretab);
    var numCols = resultados.getMetaData().getColumnCount();
    var colNameArr=[[]],valArr=[];
    for(var col = 0; col < numCols; col++){
      colNameArr[0].push(resultados.getMetaData().getColumnName(col+1));      
    }
    tabtemp.getRange(1, 1, 1, numCols).setValues(colNameArr);
    
    while (resultados.next()){
      var tmpArr = [];
      var rowString = '';
      for (var col = 0; col < numCols; col++) {
        rowString += resultados.getString(col + 1) + '\t';
        tmpArr.push(resultados.getString(col + 1));
      }
      valArr.push(tmpArr);
    }
    resultados.close();
    statm.close();
    tabtemp.getRange(2, 1 , valArr.length, numCols).setValues(valArr);
  }
  else{
    SpreadsheetApp.setActiveSheet(ss.getSheetByName(tabtemp));
  }
}
// Refreshing All (work in progress)
function refrescarListado(){
  var last = choja.getLastRow();
  var listado = choja.getRange(2, 1, last, 1).getValues();
  for (var num = 0; num < listado.length-1 ; num++){
    var idEncontrada = listado[num];
    refrescar(idEncontrada);
    Utilities.sleep(1000);
  }
  aviso('¡Boom!','<i style="color:green" class="fas fa-2x fa-bomb"></i>  ¡Boom! ¡Refrescados todos!')
}

// Refreshing a Saved Query
function refrescar(idpulsado){
  // Connect to DB
  var contenido = '<div class="block"><i style="color:red" class="fas fa-2x fa-sync-alt"></i> Refrescando consulta '+idpulsado+'...</div>';
  modalrefrescar(idpulsado,contenido);
  var conn = Jdbc.getConnection(dbUrl, user, userPwd);
  var start = new Date();
  var statm = conn.createStatement();
  var last = choja.getLastRow();
  var busqueda = choja.getRange(1,1,last,12).getValues();
  for(var lol = 0 ; lol < busqueda.length ; ++lol){
    if (busqueda[lol][0] == idpulsado){break} ;
    }

  var maximrows = busqueda[lol][8];
  statm.setMaxRows(maximrows);
  var query = busqueda[lol][3];
  var update = busqueda[lol][5];
  var timeUpdate = busqueda[lol][7];
  var tab = ss.getSheetByName(busqueda[lol][4]);

  var resultados = statm.executeQuery(query);
  
  var numCols = resultados.getMetaData().getColumnCount();
  var numRows = tab.getLastRow();
  
  // Clearing the content
  var clear = '<div class="block"><i style="color:red" class="fas fa-2x fa-trash-alt"></i> Borrando contenido '+busqueda[lol][4]+'...</div>';
  modalrefrescar(idpulsado,clear);
  tab.getRange(1,1,numRows,numCols).clear({contentsOnly: true});
  var colNameArr=[[]],valArr=[];
  for(var col = 0; col < numCols; col++){
    colNameArr[0].push(resultados.getMetaData().getColumnName(col+1));      
  }
  tab.getRange(1, 1, 1, numCols).setValues(colNameArr);
  
  while (resultados.next()){
    var tmpArr = [];
    var rowString = '';
    for (var col = 0; col < numCols; col++) {
      rowString += resultados.getString(col + 1) + '\t';
      tmpArr.push(resultados.getString(col + 1));
    }
    valArr.push(tmpArr);
  }
  var clear = '<div class="block"><i style="color:orange" class="fas fa-2x fa-fill-drip"></i> Rellenando contenido en '+busqueda[lol][4]+'...</div>';
  modalrefrescar(idpulsado,clear);
  resultados.close();
  statm.close();
  
  tab.getRange(2, 1 , valArr.length, numCols).setValues(valArr);
  
  //Date and time of query
  var stop = new Date();
  var minstop = stop.getMinutes()*60000;
  var minstart = start.getMinutes()*60000;
  var segstop = stop.getSeconds()*1000;
  var segstart = start.getMinutes()*1000;
  var milstop = stop.getMilliseconds();
  var milstart = start.getMilliseconds();
  var stopmili = minstop + segstop + milstop;
  var startmili = minstart + segstart + milstart;
  var duracion = stopmili - startmili;
  choja.getRange(1,1,last,8).getCell(lol+1, 6).setValue(stop);
  choja.getRange(1,1,last,8).getCell(lol+1, 8).setValue(duracion);
  var clear = '<div class="block"><i style="color:green" class="fas fa-2x fa-check"></i> Terminado.</div>';
  aviso(idpulsado,clear);
  Utilities.sleep(1000);
}