function main(workbook: ExcelScript.Workbook) {
    let w = workbook.getActiveWorksheet();
    let ln: number = 1;
    let col: number = 0;
    let encontrou: boolean = false;
    let vTexto: string;

    //Código ASCII : a = 97 | z = 122

    while(ln <= 6){
        encontrou = false;
        vTexto = w.getCell(ln, col).getValue().toString();

        for(let i=0; i<10; i++){
            let x: string = vTexto.substr(i, i); // Semelhante ao MID do VBA
            
            if(x.charCodeAt() >=97 && x.charCodeAt() <= 122){
                //console.log(x.charCodeAt());
                encontrou = true;
                break;
            }
        }

        w.getCell(ln, col + 1).setValue(encontrou)
        ln++;
    }

}