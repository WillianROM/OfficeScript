function main(workbook: ExcelScript.Workbook) {
    let w = workbook.getActiveWorksheet();
    w.activate();
    console.log(w.getName())
    
    let vUltCel = w.getCell(w.getRange().getRowCount() -1, 0).getRangeEdge(ExcelScript.KeyboardDirection.up).getRowIndex();
    
    let ln: number;
    let col: number;
    let vLimite = w.getRange('C2').getValue();
    
    ln = 1; // Linha 2 porque é base zero
    col = 0; // Coluna A porque é base zero
    
    w.getCell(ln, col).select();
    
    // Estrutura While
    while(ln <= vUltCel){
        if (w.getCell(ln, col).getValue() >= vLimite){
            w.getCell(ln, col + 1).setValue('Maior que o limite')
            }
        ln++;
        
    }
}
