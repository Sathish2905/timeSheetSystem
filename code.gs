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
  return today;
}
function obterHora(){
  var d = new Date(); // for now
  return ((d.getHours() < 10)?"0":"") + d.getHours() +":"+ ((d.getMinutes() < 10)?"0":"") + d.getMinutes() +":"+ ((d.getSeconds() < 10)?"0":"") + d.getSeconds();
}
function registrarPonto(){
  var dados = obterDadosNaPlanilha(idPlanilha, "timesheet", [2, 2, 7]);
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
    var sheet = SpreadsheetApp.openById(idPlanilha).getSheetByName("timesheet");
    var linha =  sheet.getLastRow();
    sheet.getRange(linha, [3]).setNumberFormat([['HH:mm']]);
}
function registraHoraEmPosicao(linha, posicao){
    incluirDadosEmColunasSeparadas(idPlanilha, "timesheet", linha,
		[posicao], [obterHora()]);
    SpreadsheetApp.openById(idPlanilha).getSheetByName("timesheet").getRange(
				linha, [posicao]).setNumberFormat([['HH:mm']]);

}
function calculaAlmoco(linha){
  incluirDadosEmColunasSeparadas(idPlanilha, "timesheet", linha,
		[7], ["=$E"+parseInt(linha+1)+"-$D"+parseInt(linha+1)+""]);
  SpreadsheetApp.openById(idPlanilha).getSheetByName("timesheet").getRange(
				linha, [7]).setNumberFormat([['HH:mm']]);
}
function calculaTotal(linha){
  incluirDadosEmColunasSeparadas(idPlanilha, "timesheet", linha,
		[8], ["=($F"+parseInt(linha+1)+"-$E"+parseInt(linha+1)+")+($D"+parseInt(linha+1)+"-$C"+parseInt(linha+1)+")"]);
  SpreadsheetApp.openById(idPlanilha).getSheetByName("timesheet").getRange(
				linha, [8]).setNumberFormat([['HH:mm']]);
}
