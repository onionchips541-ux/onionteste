// Cole este código em: planilha > Extensões > Apps Script
var CHAVE = 'troque-esta-senha';   // tem que ser IGUAL à CHAVE do index.html
var MESES = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
var CAB = ['Data','Hora do envio','Responsável','Dia da contagem','Categoria','Produto','Campo','Quantidade','Unidade','Observação'];

function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// Aba "Config": B1 = SIM ou NAO (liga/desliga o app), B2 = mensagem da tela de bloqueio
function cfg(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(), s = ss.getSheetByName('Config');
  if(!s){ s = ss.insertSheet('Config');
    s.getRange('A1:B2').setValues([['ativo','SIM'],['mensagem','Acesso encerrado. Fale com a gerência.']]); }
  var v = s.getRange('B1:B2').getValues();
  return { ativo: String(v[0][0]).trim().toUpperCase() === 'SIM', msg: String(v[1][0]) };
}

function doGet(e){
  if((e.parameter||{}).chave !== CHAVE) return out({ativo:false, msg:'Chave inválida'});
  return out(cfg());
}

function doPost(e){
  var lock = LockService.getScriptLock();
  try{
    var d = JSON.parse(e.postData.contents);
    if(d.chave !== CHAVE) return out({ok:false, erro:'chave'});
    if(!cfg().ativo) return out({ok:false, erro:'bloqueado'});
    lock.waitLock(20000);
    var p = d.data.split('-'), dt = p[2]+'/'+p[1]+'/'+p[0];
    var nome = (d.lista==='frente' ? 'Frente ' : 'Cozinha ') + MESES[+p[1]-1] + ' ' + p[0];
    var ss = SpreadsheetApp.getActiveSpreadsheet(), s = ss.getSheetByName(nome);
    if(!s){ s = ss.insertSheet(nome); s.appendRow(CAB);
      s.getRange(1,1,1,CAB.length).setFontWeight('bold').setBackground('#1F6B4A').setFontColor('#ffffff');
      s.setFrozenRows(1); s.setColumnWidth(6,240); }
    // reenvio da mesma pessoa, data e dia substitui o anterior
    var n = s.getLastRow();
    if(n > 1){
      var v = s.getRange(2,1,n-1,4).getDisplayValues(), run = 0;
      for(var i=v.length-1; i>=-1; i--){
        var m = i>=0 && v[i][0]===dt && v[i][2]===d.resp && v[i][3]===d.dia;
        if(m) run++; else if(run){ s.deleteRows(i+3, run); run = 0; }
      }
    }
    var rows = d.linhas.map(function(l){ return [dt,d.hora,d.resp,d.dia,l[0],l[1],l[2],l[3],l[4],l[5]]; });
    var r = s.getRange(s.getLastRow()+1, 1, rows.length, CAB.length);
    s.getRange(s.getLastRow()+1, 1, rows.length, 4).setNumberFormat('@');
    r.setValues(rows);
    return out({ok:true, linhas:rows.length, aba:nome});
  }catch(err){ return out({ok:false, erro:String(err)}); }
  finally{ try{ lock.releaseLock(); }catch(x){} }
}
