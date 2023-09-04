function main(workbook: ExcelScript.Workbook, teste:number) {
    let w1 = workbook.getActiveWorksheet();
    let w2 = workbook.getWorksheet('Dados1');
    let w3 = workbook.getWorksheet('Dados2');

    
    let vUltCel = w1.getCell(w1.getRange().getRowCount() -1, 0).getRangeEdge(ExcelScript.KeyboardDirection.up).getRowIndex();
    

    let lnW1:number = 1;
    let lnW2:number = 0;
    let lnW3:number = 0;
    let colW1:number = 0;
    let colW2:number = 0;
    let colW3:number = 0;

    w2.getRange('A:B').delete(ExcelScript.DeleteShiftDirection.left);
    w3.getRange('A:B').delete(ExcelScript.DeleteShiftDirection.left);

    w1.activate();
    w1.getCell(lnW1, colW1).select();
    w1.getRange('B2:B' + vUltCel).getFormat().getFont().setBold(false);

    while(lnW1 <= vUltCel){
        if(w1.getCell(lnW1, colW1 + 1).getValue() == 'Dados1'){
            w1.getCell(lnW1, colW1).getResizedRange(0, colW1 + 1).select();
            w1.getCell(lnW1, colW1 + 1).getFormat().getFont().setBold(true);

            // Copiar para a outra planilha
            w2.getCell(lnW2, colW2).copyFrom(w1.getCell(lnW1, colW1).
                getResizedRange(0, colW1 + 1), ExcelScript.RangeCopyType.all,false,false);
                lnW2++;
        }

        if (w1.getCell(lnW1, colW1 + 1).getValue() == 'Dados2') {
            w1.getCell(lnW1, colW1).getResizedRange(0, colW1 + 1).select();
            w1.getCell(lnW1, colW1 + 1).getFormat().getFont().setBold(true);

            // Copiar para a outra planilha
            w3.getCell(lnW3, colW3).copyFrom(w1.getCell(lnW1, colW1).
                getResizedRange(0, colW1 + 1), ExcelScript.RangeCopyType.all, false, false);
            lnW3++;
        }

        lnW1 ++;

    }

}
