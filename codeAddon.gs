/* 

Created By MuXeD
Creative Commons

Attribution-NonCommercial 4.0 International (CC BY-NC 4.0) 
https://creativecommons.org/licenses/by-nc/4.0/

Dependencies: 

cCryptoGS by brucemcpherson to encrypt the pass
http://ramblings.mcpher.com/Home/excelquirks/gassnips/cryptogs
https://github.com/brucemcpherson/cCryptoGS

----------------------------------------------------------

Spanish version for menus, you can change as you wish

-- V. 0.4 --

*/

//Connection Default Preferences
var connectionName = 'server.com'; //Ip or direction
var user = 'user'; // User
var userPwd = 'pass'; // Password
var db = 'database'; // Database
// This type of connection can change, this is configurated in MySQL, and change getConnection() function
// If you use other Subprotocol change here too jdbc:<subprotocol>
var dbUrl = 'jdbc:mysql://' + connectionName + '/' + db;

// General Vars 
var ss = SpreadsheetApp.getActiveSpreadsheet();

// Var for cypher with Bruce Script 
// Change Passphrase and codification (crypted in Rabbit, add/change other Algos to Cifrado.gs)
var cifrado = new Cipher('pues no es la magia, estupefacientes');

// You can change the name of the Query Data Tab
var ssmelena = 'MELENA';
var ssestupe = 'ESTUPEFACIENTES';
var estupehoja = ss.getSheetByName(ssestupe);
var choja = ss.getSheetByName(ssmelena);

// If you like, onOpen() this function or active it as you prefere
function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('La magia de mi melena')
  .addItem('Mírala','mira')
  .addSeparator()
  .addItem('Crear Query','query')
  .addItem('Crear Conexión','createConnection')
  .addSeparator()
  .addItem('Consultas','consultas')
  .addItem('Conexiones','conexiones')
  .addItem('Programar consultas', 'scheduling')
  .addItem('Activadores', 'verActivadores')
  .addToUi();
}

// Create Connection TAB
function crearestupehoja(){
if (!estupehoja){
    ss.insertSheet(ssestupe);
    estupehoja = ss.getSheetByName(ssestupe);
    estupehoja.setFrozenRows(1);
    // Spanish
    estupehoja.appendRow(["Nombre","Dirección","Base de Datos","Usuario","Contraseña Cifrada"]);
    estupehoja.getRange('1:1').setBackground('#333333').setFontColor('#eeeeee').setFontWeight('bold');
    var titulo = "Hoja creada";
    var mensajeError = "Mira la magia de mi melena. Apenas he tenido contacto con los estupefacientes.";
    aviso(titulo,mensajeError);
  }
}

// Create Queries TAB
function crearchoja(){
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
  }
}

// Create TABS to archive Queries configs
function mira(){
  crearestupehoja();
  crearchoja();
  if(choja && estupehoja){
    var titulo = "Las hojas ya existen";
    var mensajeError = "La hoja ya existe, pero la magia de mi melena me perturba. Apenas he tenido contacto con los estupefacientes.";
    aviso(titulo,mensajeError);
  }
}

// Test connections
function pruebaconexion(prueba){
  if(prueba == 'default'){
    var pruebadefe = Jdbc.getConnection(dbUrl, user, userPwd);
  }
  else{
    var lastserver = estupehoja.getLastRow();
    var busquedaserver = estupehoja.getRange(1,1,lastserver,6).getValues();
    for(var rofl = 0 ; rofl < busquedaserver.length ; ++rofl){
      if (busquedaserver[rofl][0] == prueba){break} ;
    }
    var nombre = busquedaserver[rofl][1];
    var usuario = busquedaserver[rofl][3];
    var pass = cifrado.decrypt(busquedaserver[rofl][4]);
    var db = busquedaserver[rofl][2];
    // If you use other Subprotocol change here too jdbc:<subprotocol>
    var urldb = 'jdbc:mysql://' + nombre + '/' + db;
    var pruebadefe = Jdbc.getConnection(urldb, usuario, pass).isValid(1000);
  }
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
  var ui = SpreadsheetApp.getUi();
  ui.showModalDialog(final, titulo);
}

// Schedule dialog
function scheduling(){
  var contenido = HtmlService.createHtmlOutputFromFile('TemplateProgramar').setWidth(600).setHeight(300);
  var ui = SpreadsheetApp.getUi();
  ui.showModelessDialog(contenido, 'Programar consultas');
}

// Activators dialog
function verActivadores(){
  var contenido = HtmlService.createHtmlOutputFromFile('TemplateActivadores').setWidth(600).setHeight(400);
  var ui = SpreadsheetApp.getUi();
  ui.showModelessDialog(contenido, 'Activadores');
}

// Generate a Modal Popup
function modalrefrescar(titulo,aviso){  
  var final = HtmlService.createHtmlOutputFromFile('TemplateRefresh').append(aviso).setWidth(400).setHeight(200);
  var ui = SpreadsheetApp.getUi();
  ui.showModalDialog(final, titulo);
}

// Sidebar for create Query
function query(){
  crearchoja();
  var html = HtmlService.createHtmlOutputFromFile('TemplateSideb')
      .setTitle('Introduce datos de la Query')
      .setWidth(300);
  var ui = SpreadsheetApp.getUi();
  ui.showSidebar(html);
}

// Sidebar for create connection
function createConnection(){
  crearestupehoja();
  var html = HtmlService.createHtmlOutputFromFile('TemplateSidebServer')
      .setTitle('Introduce datos de la Conexión')
      .setWidth(300);
  var ui = SpreadsheetApp.getUi();
  ui.showSidebar(html);
}

// Open connection tab
function conexiones(){
  crearestupehoja();
  var pruebaulti = estupehoja.getLastRow();
  if (pruebaulti < 2) {
    SpreadsheetApp.setActiveSheet(estupehoja);
    createConnection();
    var titulo = "Sin conexiones";
    var mensajeCorrecto = "Crea una conexión, solo tienes la configurada por defecto";
    aviso(titulo,mensajeCorrecto);
  } else {
    var html = HtmlService.createHtmlOutputFromFile('TemplateConexiones')
    .setTitle('Probar Conexiones')
    .setWidth(300);
    var ui = SpreadsheetApp.getUi();
    ui.showSidebar(html);
  }
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

// Pass the list of server
function selectserver(){
  var ultiest = estupehoja.getLastRow();
  var rangoest = estupehoja.getRange(2, 1, ultiest-1, 1).getValues();
  var listadoserver = [];
  for (var colt =0; colt < ultiest-1;colt++){
    listadoserver.push(rangoest[colt]);
  }
  return listadoserver;
}

// Sidebar for Refresh
function consultas(){
  crearchoja();
  var pruebaulti = choja.getLastRow();
  if (pruebaulti < 2){
    SpreadsheetApp.setActiveSheet(choja);
    query();
    var titulo = "Sin consultas";
    var mensajeCorrecto = "Crea una consulta.";
    aviso(titulo,mensajeCorrecto);
  } else {
    var html = HtmlService.createHtmlOutputFromFile('TemplateConsultas')
    .setTitle('Consultas Guardadas')
    .setWidth(300);
    var ui = SpreadsheetApp.getUi();
    ui.showSidebar(html);
  }  
}

// Creating a Data for connection 
function creaciontablasserver(valores){
  var nombre = valores[0];
  var direccion = valores[1];
  var db = valores[2];
  var usuario = valores[3];
  var password = valores[4];
  var passwordenc = cifrado.encrypt(password);
  // If you use other Subprotocol change here too jdbc:<subprotocol>
  var urldb = 'jdbc:mysql://' + direccion + '/' + db;;
  var conn = Jdbc.getConnection(urldb, usuario, password).isValid(1000);
  ss.getSheetByName(ssestupe).appendRow([nombre,direccion,db,usuario,passwordenc]);
  if(conn){
    var titulo = "Parece que está conectado";
    var mensajeCorrecto = "Pues parece que chispea";
    aviso(titulo,mensajeCorrecto);
  }
}

// Creating a Tab for a Query and the first Refresh
function creaciontablas(valores,server){
  var cdate = new Date();
  var nombretab = valores[1];
  var nombre = valores[0];
  var consulta = valores[2];
  var func = valores[0];
  var filasm = valores[3];
  if(server == 'Por defecto'){
    var conn = Jdbc.getConnection(dbUrl, user, userPwd);
  }
  else{
    var lastserver = estupehoja.getLastRow();
    var busquedaserver = estupehoja.getRange(1,1,lastserver,6).getValues();
    for(var rofl = 0 ; rofl < busquedaserver.length ; ++rofl){
      if (busquedaserver[rofl][0] == server){break} ;
    }
    var nombreserver = busquedaserver[rofl][1];
    var usuario = busquedaserver[rofl][3];
    var pass = cifrado.decrypt(busquedaserver[rofl][4]);
    var db = busquedaserver[rofl][2];
    // If you use other Subprotocol change here too jdbc:<subprotocol>
    var urldb = 'jdbc:mysql://' + nombreserver + '/' + db;
    var conn = Jdbc.getConnection(urldb, usuario, pass);
  }
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
// Refreshing All (detected a error in more than 6 queries)
function refrescarListado(server){
  var last = choja.getLastRow();
  var listado = choja.getRange(2, 1, last, 1).getValues();
  for (var num = 0; num < listado.length-1 ; num++){
    var idEncontrada = listado[num];
    refrescar(idEncontrada,server);
    Utilities.sleep(1000);
  }
  aviso('¡Boom!','<i style="color:green" class="fas fa-2x fa-bomb"></i>  ¡Boom! ¡Refrescados todos!')
}

// Refreshing a Saved Query
function refrescar(idpulsado,server){
  // Connect to DB
  var contenido = '<div class="block"><i style="color:red" class="fas fa-2x fa-sync-alt"></i> Refrescando consulta '+idpulsado+'...</div><div class="block"><i style="color:red" class="fas fa-2x fa-plug"></i> Conectando al server '+server+'...</div>';
  modalrefrescar(idpulsado,contenido);
  if(server == 'Por defecto'){
    var conn = Jdbc.getConnection(dbUrl, user, userPwd);
  }
  else{
    var lastserver = estupehoja.getLastRow();
    var busquedaserver = estupehoja.getRange(1,1,lastserver,6).getValues();
    for(var rofl = 0 ; rofl < busquedaserver.length ; ++rofl){
      if (busquedaserver[rofl][0] == server){break} ;
    }
    var nombre = busquedaserver[rofl][1];
    var usuario = busquedaserver[rofl][3];
    var pass = cifrado.decrypt(busquedaserver[rofl][4])
    var db = busquedaserver[rofl][2];
    // If you use other Subprotocol change here too jdbc:<subprotocol>
    var urldb = 'jdbc:mysql://' + nombre + '/' + db;
    var conn = Jdbc.getConnection(urldb, usuario, pass);
  }
  var start = new Date();
  var milstart = start.getTime();
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
  var milstop = stop.getTime();
  var duracion = (milstop - milstart)/1000;
  choja.getRange(1,1,last,8).getCell(lol+1, 6).setValue(stop);
  choja.getRange(1,1,last,8).getCell(lol+1, 8).setValue(duracion);
  choja.getRange(1,1,last,11).getCell(lol+1, 11).setValue("Refrescado Manual: "+stop).setBackground("orange").setFontColor("white");
  var clear = '<div class="block"><i style="color:green" class="fas fa-2x fa-check"></i> Terminado.</div>';
  aviso(idpulsado,clear);
  Utilities.sleep(1000);
}

// Take Triggers from Script Propierties
function recogerActivadores(){
  var triggers = ScriptApp.getProjectTriggers();
  var keys = PropertiesService.getScriptProperties().getKeys();
  var datos = PropertiesService.getScriptProperties().getProperties();
  var totalScript = [];
  for (var i = 0; i < triggers.length; i++){
    var ID = triggers[i].getUniqueId();
    for (var x = 0; x < keys.length; x++ ){
      if (ID == keys[x]){
        totalScript.push(ID);
        totalScript.push(datos[ID]);
      }
    }
  }
  return totalScript;
}

// Delete 1 Trigger
function eliminar(ID){
  var aviso = 'Con esto borrará el activador '+ID+'. ¿Está seguro que desea continuar?';
  var titulo = 'Eliminar Activador'+ID;
  var ui = SpreadsheetApp.getUi(); 
  var response = ui.alert(titulo,aviso, ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES){
    var allTriggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < allTriggers.length; i++){
      if(allTriggers[i].getUniqueId() == ID){
        ScriptApp.deleteTrigger(allTriggers[i]);
        var propiedadesScript = PropertiesService.getScriptProperties();
        propiedadesScript.deleteProperty(allTriggers[i].getUniqueId());
      }
    }
  }
  verActivadores();
}

// Clean All Trigers
function deleteAllTrigger() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var allTriggers = ScriptApp.getProjectTriggers();
  var aviso = 'Con esto borrará todas los activadores creados. ¿Está seguro que desea continuar?';
  var titulo = 'Eliminar todos los Activadores';
  var ui = SpreadsheetApp.getUi(); 
  var response = ui.alert(titulo,aviso, ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES){
    for (var i = 0; i < allTriggers.length; i++) {
      ScriptApp.deleteTrigger(allTriggers[i]);
      scriptProperties.deleteProperty(allTriggers[i].getUniqueId());
    }
  }
}

// Create All Triggers
function crearActivador(datos){
  var scriptProperties = PropertiesService.getScriptProperties();
  aviso('Activador guardado','<div class="block"><i style="color:green" class="fas fa-2x fa-grimace"></i> Activador Creado para verlo accede al menú "Activadores".</div>');
  var metodotemp = datos[1];
  var horadeldia = datos[2];
  var diasemana = datos[3];
  var horas = datos[4];
  switch(metodotemp){
    case "hora":
      var trigger = ScriptApp.newTrigger("manejadorA").timeBased().everyHours(1).create();
      var ID = trigger.getUniqueId();
      scriptProperties.setProperty(ID, JSON.stringify(datos));
      break;
    case "dia":
      var trigger = ScriptApp.newTrigger("manejadorA").timeBased().everyDays(1).atHour(horadeldia).create();
      var ID = trigger.getUniqueId();
      scriptProperties.setProperty(ID, JSON.stringify(datos));
      break;
    case "semana":
      switch (diasemana){
        case "Lunes":
          var trigger = ScriptApp.newTrigger("manejadorA").timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).create();
          var ID = trigger.getUniqueId();
          scriptProperties.setProperty(ID, JSON.stringify(datos));
          break;
        case "Martes":
          var trigger = ScriptApp.newTrigger("manejadorA").timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).create();
          var ID = trigger.getUniqueId();
          scriptProperties.setProperty(ID, JSON.stringify(datos));
          break;
        case "Miércoles":
          var trigger = ScriptApp.newTrigger("manejadorA").timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).create();
          var ID = trigger.getUniqueId();
          scriptProperties.setProperty(ID, JSON.stringify(datos));
          break;
        case "Jueves":
          var trigger = ScriptApp.newTrigger("manejadorA").timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).create();
          var ID = trigger.getUniqueId();
          scriptProperties.setProperty(ID, JSON.stringify(datos));
          break;
        case "Viernes":
          var trigger = ScriptApp.newTrigger("manejadorA").timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).create();
          var ID = trigger.getUniqueId();
          scriptProperties.setProperty(ID, JSON.stringify(datos));
          break;
        case "Sábado":
          var trigger = ScriptApp.newTrigger("manejadorA").timeBased().onWeekDay(ScriptApp.WeekDay.SATURDAY).create();
          var ID = trigger.getUniqueId();
          scriptProperties.setProperty(ID, JSON.stringify(datos));
          break;
        case "Domingo":
          var trigger = ScriptApp.newTrigger("manejadorA").timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).create();
          var ID = trigger.getUniqueId();
          scriptProperties.setProperty(ID, JSON.stringify(datos));
          break;
      }
      break;
    case "custom":
      var trigger = ScriptApp.newTrigger("manejadorA").timeBased().everyHours(horas).create();
      var ID = trigger.getUniqueId();
      scriptProperties.setProperty(ID, JSON.stringify(datos));
      break;
  }
}

// Used to interact with the Script Propierties archived when create the trigger 
function manejadorA(e){
  var scriptProperties = PropertiesService.getScriptProperties();
  var triggerId = e.triggerUid;
  var triggers = ScriptApp.getProjectTriggers();
  for(var i =0; i<triggers.length; i++){
    if(triggers[i].getUniqueId() == triggerId){
      var encontrado = scriptProperties.getProperty(triggerId);
      var lavueltita = JSON.parse(encontrado);
      break;
    }
  }
  for(var x = 5; x < lavueltita.length; x++){
    refrescarApa(lavueltita[x],lavueltita[0]);
  }
}

// Refreshing outprogram
function refrescarApa(idpulsado,server){
  if(server == 'Por defecto'){
    var conn = Jdbc.getConnection(dbUrl, user, userPwd);
  }
  else{
    var lastserver = estupehoja.getLastRow();
    var busquedaserver = estupehoja.getRange(1,1,lastserver,6).getValues();
    for(var rofl = 0 ; rofl < busquedaserver.length ; ++rofl){
      if (busquedaserver[rofl][0] == server){break} ;
    }
    var nombre = busquedaserver[rofl][1];
    var usuario = busquedaserver[rofl][3];
    var pass = cifrado.decrypt(busquedaserver[rofl][4])
    var db = busquedaserver[rofl][2];
    // If you use other Subprotocol change here too jdbc:<subprotocol>
    var urldb = 'jdbc:mysql://' + nombre + '/' + db;
    var conn = Jdbc.getConnection(urldb, usuario, pass);
  }
  var start = new Date();
  var milstart = start.getTime();
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
  resultados.close();
  statm.close();
  tab.getRange(2, 1 , valArr.length, numCols).setValues(valArr);
  var stop = new Date();
  var milstop = stop.getTime();
  var duracion = (milstop - milstart)/1000;
  choja.getRange(1,1,last,8).getCell(lol+1, 6).setValue(stop);
  choja.getRange(1,1,last,8).getCell(lol+1, 8).setValue(duracion);
  choja.getRange(1,1,last,11).getCell(lol+1, 11).setValue("Refrescado con Activador: "+stop).setBackground("green").setFontColor("white");
}