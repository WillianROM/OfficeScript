function main(workbook: ExcelScript.Workbook) {
    // Movimentação na planilha

    let w = workbook.getActiveWorksheet();

    w.getRange('G5').setValue('Teste');

    w.getRange('A1:E5').select();

    w.getRange('G1:G3').setValue('Aqui há valor...');

    w.getRange('H1:J1')
        .setValues(
            [['Col1','Col2', 'Col3']]
            );

    console.log('Aqui termina o meu script...')

    console.log(w.getName())
}