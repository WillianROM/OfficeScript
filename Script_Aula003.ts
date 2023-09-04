function main(workbook: ExcelScript.Workbook) {
    // Declarar variáveis
    let w = workbook.getActiveWorksheet();
    let vTexto:string = '';
    let vNum: number = 0;

  // Selecionar alguma aba
    w.activate();
    workbook.getWorksheet('planilha1').activate();

  // Seleção de células
  vTexto = 'A1:A15';

  w.getRange(vTexto).select();
  w.getRange(vTexto).setValue('W')

  w.getRange('A1').getOffsetRange(2,2).select();

  w.getCell(1,1).setValue('Aqui'); // base zero, então a informação vai para a célula B2, não na A1

  // loop for

  let vUltCel = w.getCell(w.getRange().getRowCount() -1, 0).getRangeEdge(ExcelScript.KeyboardDirection.up).getRowIndex();

  for (let i = 0; i < vUltCel + 1; i++){
    w.getCell(i,2).setValue(i + 1)
  }


  // Escrever fórmulas nas células (Fórmulas em inglês)
  w.getRange('C16').setFormula('=SUM(C1:C15)')
  w.getRange('D16').setFormula('=COUNTA(C:C)')

  
  // Formatação
  w.getRange('A1:A15').getFormat().getFill().setColor('00ffff');
    // Limpar a formatação
  w.getRange('A1').getFormat().getFill().clear();
    // Deixar texto em negrito
  w.getRange('A1:A15').getFormat().getFont().setBold(true);
    // Deixar texto em itálico
  w.getRange('A1:A15').getFormat().getFont().setItalic(true);


  // Mostrar prints
      console.log(w.getName())
      console.log('Esse arquivo tem ' + workbook.getWorksheets().length + ' aba(s)')

  // Criar tabela de dados
    // Tabela de dados será feito na outra aba
  let w2 = workbook.getWorksheet('planilha2');
  w2.activate();

  let ln: number = 1;
  let col: number = 0;

  let vCabecalho = [['Item', 'Produto', 'Valor', 'Estoque']];
  let vLinha1 = [['1', 'Dell', 5000, 1]];
  let vLinha2 = [['2', 'HP', 5600, 3]];
  let vLinha3 = [['3', 'Acer', 4000, 10]];

  w2.getRange('A1').setValue('x');
  w2.getUsedRange().clear();

  w2.getRange('A1:D1').setValue(vCabecalho);
  w2.getRange('A2:D2').setValue(vLinha1);
  w2.getRange('A3:D3').setValue(vLinha2);
  w2.getRange('A4:D4').setValue(vLinha3);

  // Transformar os dados em uma tabela
  let x = w2.addTable(w2.getUsedRange(), true)

  // Renomear o nome da tabela
  x.setName('TB_EXEMPLO')

  // Colocar a linha de totais na tabela
  x.setShowTotals(true);

  // Converter a tabela em intervalo
  x.convertToRange();

  return; // Semelhante ao Exit Sub

  // loop While
  while(true){
    console.log('Passei por aqui');
    break; // Sair do loop
  }

}


