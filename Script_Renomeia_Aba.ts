function main(workbook: ExcelScript.Workbook) {
    // Percorre as planilhas e ignora a troca do nome
    // Se uma planilha já existe

    let wks = workbook.getWorksheets();
    let count = 0;

    wks.forEach(wks => {
        if(wks.getName() != "Testes de Script") {
            count++;
            if(wks.getName() == "Planilha2"){
                wks.setName('Troquei o nome')
            }
        }
    })
}