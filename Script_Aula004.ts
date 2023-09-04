function main(workbook: ExcelScript.Workbook) {
    // Atribuição de variábeis
    let wTeste = workbook.getActiveWorksheet();
    let vValor:string = 'Teste';
    let w = workbook.getWorksheet('Planilha1');

    // Seleção de abas
    w.activate();

    // Seleção de células
    w.getUsedRange().select();
    
    w.getRange('B2').select();

    w.getRange('B3').getOffsetRange(1,1).select();
    
    w.getCell(1,1).select();

    // Escrever nas células
    w.getRange('D1').setValue('Teste');

    w.getRange('A13').setFormula('=SUM(A1:A11)');
    w.getRange('A14').setFormula('=COUNTA(A1:A11)')

    // Formatação
    w.getRange('A1:A12').getFormat().getFill().setColor('000000');
    w.getRange('A1:A13').getFormat().getFont().setColor('d6f5c0');
    w.getRange('A1:A13').getFormat().getFont().setBold(true);
    w.getRange('A1:A13').getFormat().getFont().setItalic(true);

    // Console.log
    console.log(vValor);
    console.log('Teste');
    console.log(w.getName());
    console.log(workbook.getWorksheets().length);



}