// https://www.youtube.com/watch?v=bk-75vAJZ6I&list=PL7iAT8C5wumoam5OqlBoA2QU4SmS3Kkun
function main(workbook: ExcelScript.Workbook) {
    // Declaração de variáveis
    // Em VBA seria: dim w as worksheet
    let w = workbook.getActiveWorksheet();
    
    // Nome da planilha
    console.log(w.getName())
    //return - Utilize return para não executar os códigos abaixo, por exemplo em testes
    
    console.log("Esse arquivo tem " + workbook.getWorksheets().length + " planilha(s).")
    
    // Em VBA seria: dim ln as integer
    let ln:number;
    // O numerador de linha e coluna é base zero
    // Em VBA seria:dim col as integer
    // Em VBA seria: let col = 0
    let col:number = 0;
    
    // Em VBA seria: dim vTexto as String
    // Em VBA seria: let vTexto = "Ola mundo"
    let vTexto:string = "Ola mundo";
    
    // Em VBA seria: dim vTeste as Boolean
    // let vTeste = false
    let vTeste:boolean = false;
    
    let w3 = workbook.getWorksheet("Exemplo3")
    w3.activate();
    
    // Seleção de células
    // Selecionar células com valores
    w.getUsedRange().select();
    
    // Selecionar uma célula específica
    w.getRange("B1").select();
    w.getRange("A1:F3").select();
    w.getCell(0, 0).select(); // Irá selecionar a célula A1 (lembrando que é base ZERO)
    
    
    // Última célula da minha coluna A
    // Em VBA seria: set UltCel = w.cells(w.rows.count, 1),end(xlup).row
    let vUltCel = w.getCell(w.getRange().getRowCount() - 1, 0).getRangeEdge(ExcelScript.KeyboardDirection.up).getRowIndex();
    
    // Última linha
    // No VBA seria: debug.print
    console.log("Última linha: " + (vUltCel + 1));
    console.log(1 + 1 + " texto")
    console.log("texto " + 1 + 1)
    console.log("texto " + (1 + 1))
    
    // ESCREVER NAS CÉLULAS
    w.getRange("G1").setValue('Meu primeiro código em Office Script');
    
    // Informar na célula A1 o número da última linha da planilha
    w.getCell(0, 0).setValue(w.getRange().getRowCount());
    
    // Fórmulas
    w.getRange('A14').setFormula('=SUM(A1:A12)')
    w.getRange('A15').setFormula('=COUNTA(A1:A12)')
    
    // FORMATAÇÃO
    // Preencher o background com alguma cor hexadecimal
    w.getRange('A1:a12').getFormat().getFill().setColor('00ff00')
    
    // Limpar preenchimento da cor de fundo
    w.getRange('A1:a10').getFormat().getFill().clear();
    
    // Deixar alguma célula na cor negrito
    w.getRange('A14').getFormat().getFont().setBold(true); // setBold precisa de um valor booleano
    
    // Deixar alguma célula em itálico
    w.getRange('A15').getFormat().getFont().setItalic(true);
    
}