  function doGet() {
    return HtmlService.createTemplateFromFile('index').evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME); 
  }


  function include(filename){
    return HtmlService.createHtmlOutputFromFile(filename)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME).getContent();
  }
  var idPlanilha = "1ZsYLhznruW_S6N3rEdwwrc4HZLQFrcOxpWad5oP1nqo";
function obterDadosNaPlanilha(idPlanilha, nomeGuia, intervaloCelulas) {
  var sheet = SpreadsheetApp.openById(idPlanilha).getSheetByName(nomeGuia);
  var total =  sheet.getLastRow();
  if(intervaloCelulas[0] != 1){
    total = total - 1;
  }
  if(total == 0){
    return [];
  }
  var lista = sheet.getRange(intervaloCelulas[0], intervaloCelulas[1],
   total,
    intervaloCelulas[2]).getValues();
  // verifica se existe objeto do tipo data e formata no padrão adquado
//  for (i in lista) {
//    for (c in lista[i]) {
//      if (typeof lista[i][c] === "object") {
//        var data = formatarDataDaPlanilha(lista[i][c]);
//        lista[i][c] = data;
//      }
//    }
//  }
  
  return lista;
};

function incluirDadosNaPlanilha(idPlanilha, nomeGuia, dados) {
	var planilha = SpreadsheetApp.openById(idPlanilha);
	var guia = planilha.getSheetByName(nomeGuia);
	dados.unshift(guia.getLastRow()); // Incluir valor retornado pelo método
	guia.appendRow(dados);
};

function incluirDadosEmColunasSeparadas(idPlanilha, nomeGuia, numeroLinha,
		colunas, dados) {
	var idLinha = parseInt(numeroLinha);
	for (i in colunas) {
		SpreadsheetApp.openById(idPlanilha).getSheetByName(nomeGuia).getRange(
				idLinha + 1, colunas[i]).setValue(dados[i]);
	}
};

function obterEmail() {
	return Session.getActiveUser().getEmail();
};

function formatarDataDaPlanilha(data) {
	if (typeof (data) == "object") {
		var data = new Date(data);
		var dia = data.getDate();
		var mes = data.getMonth() + 1;
		var ano = data.getFullYear();
		// acrescenta o zero ao dia
		if (dia < 10)
			dia = "0" + dia;
		// acrescenta o zero ao mes
		if (mes < 10)
			mes = "0" + mes;
		data = dia + "/" + mes + "/" + ano;
	}
	return data;
};
function obterHoje(){
  var today = new Date();
  var dd = today.getDate();
  var mm = today.getMonth()+1; //January is 0!
  var yyyy = today.getFullYear();
  
  if(dd<10) {
      dd='0'+dd
  } 
  
  if(mm<10) {
      mm='0'+mm
  } 
  
  today = dd+'/'+mm+'/'+yyyy.toString().substr(2, 3);
  //Logger.log(today);
  return today;
}
function obterPontoFormatado(){
  var dados = obterDadosNaPlanilha(idPlanilha, "timesheet", [2, 2, 7]);
  for (i in dados) {
    var contador = 0;
    for (c in dados[i]) {
      if (typeof dados[i][c] === "object") {
//        if(contador == 0){
//          var data = formatarDataDaPlanilha(dados[i][c]);
//          dados[i][c] = data;
//        }
      }
      contador++;
    }
  }
  return dados;
}
function obterHora(){
  var d = new Date(); // for now
  var h = d.getHours(); // => 9
  var m = d.getMinutes(); // =>  30
  d.getSeconds();
  
  return h+":"+m;
}
function mostraTabelaPonto(){
  var dados = obterPontoFormatado();
  Logger.log(dados);
}
function registrarPonto(){
  var dados = obterPontoFormatado();
  Logger.log(dados);
  if(dados.length == 0){ //caso a planilha esteja vazia
    registraEntradaDoDia();
    return "Entrada do dia registrada com sucesso (primeira do mês)";
  }else{
    var id = parseInt(dados.length-1);
    var ultima = dados[id];
    if(ultima[2] == ""){
      registraHoraEmPosicao(id+1, 4);
      return "Saída para almoço registrada com sucesso";
    }
    if(ultima[3] == ""){
      registraHoraEmPosicao(id+1, 5);
      return "Retorno do almoço registrado com sucesso";
    }
    if(ultima[4] == ""){
      registraHoraEmPosicao(id+1, 6);
      calculaAlmoco(id+1);
      calculaTotal(id+1);
      return "Final do expediente registrado com sucesso";
    }
    registraEntradaDoDia();
    return "Entrada do dia registrada com sucesso";
  }
}
function registraEntradaDoDia(){
    incluirDadosNaPlanilha(idPlanilha, "timesheet", [obterHoje(), obterHora()]);
}
function registraHoraEmPosicao(linha, posicao){
    incluirDadosEmColunasSeparadas(idPlanilha, "timesheet", linha,
		[posicao], [new Date()]);
}
function calculaAlmoco(linha){
  incluirDadosEmColunasSeparadas(idPlanilha, "timesheet", linha,
		[7], ["=$E"+parseInt(linha+1)+"-$D"+parseInt(linha+1)+""]);
}
function calculaTotal(linha){
  incluirDadosEmColunasSeparadas(idPlanilha, "timesheet", linha,
		[8], ["=($F"+parseInt(linha+1)+"-$E"+parseInt(linha+1)+")+($D"+parseInt(linha+1)+"-$C"+parseInt(linha+1)+")"]);
}
