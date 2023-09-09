function main(workbook: ExcelScript.Workbook) {
   let w = workbook.getActiveWorksheet()
   let ln: number = 1
   let col: number = 0
   let vContador: number = 0
   let vEncontrou: boolean
   let vTexto: string

   // Tabela ASCII: 97 e 122 (a...z)

   while(ln <= 6){

       vTexto = w.getCell(ln, col).getValue().toString()
   
       vEncontrou = false
       vContador = 0

       for(let i = 0; i < vTexto.length; i++){
           let x: string = vTexto.substr(i, 1) // Semelhante ao MID do VBA
           
           let xCharCodeAt:number = x.charCodeAt()

           if (xCharCodeAt >= 97 && xCharCodeAt <= 122) {
               vEncontrou = true;
               vContador++;
           }
       }

        if(vEncontrou){
            w.getCell(ln, col + 1).setValue(`Total: ${vContador}`)
        }else{
            w.getCell(ln, col + 1).setValue('Não encontrou')
        }

       ln++
   }

}