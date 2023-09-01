function main(workbook: ExcelScript.Workbook) {
    // Your code here

    let vLn = 1; // linha 2 porque é base zero
    let vCl = 0; // coluna A porque é base zero
    let vSomatoria : number = 0;
    let vCiclo = 0;
    let vValor : number = 0;
    let w = workbook.getActiveWorksheet();

    w.getCell(vLn, vCl).select();

    while(w.getCell(vLn, vCl).getValue() > vCiclo){
        w.getCell(vLn, vCl).select();

        vValor = parseInt(w.getCell(vLn, vCl).getValue())

        vSomatoria = vSomatoria + vValor;

        vLn++;
    };

    console.log('O total da soma é: ' + vSomatoria);
    w.getRange('A9').setValue('O resultado é ' + vSomatoria)
}