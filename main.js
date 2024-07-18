function main(workbook: ExcelScript.Workbook) {
	let sheetName: string = "Dados do Sic-tos";

	let selectedSheet = workbook.getWorksheet(sheetName);
	if (!selectedSheet) {
    console.log(`Planilha com o nome ${sheetName} não foi encontrada.`);
    return;
  }
	
	let columnB = selectedSheet.getRange("B2:B" + selectedSheet.getUsedRange().getRowCount());
	
	let lastRow = columnB.getLastRow().getRowIndex();

	let hourFrequency: number[] = new Array(24).fill(0)

	for (let i = 1;i<= lastRow; i++){
		let cell = selectedSheet.getRange("B" + i)
		let cellValue = cell.getText();
		console.log("teste for")

		if(cellValue){

			let timeMatch = cellValue.match(/\b\d{1,2}:\d{2}:\d{2}\b/);
			
			// Se encontrar uma hora no formato esperado
			if (timeMatch) {
				let timePart = timeMatch[0];
				let hour = parseInt(timePart.split(":")[0]);
				console.log(`Linha ${i}: hora encontrada = '${hour}'`);
				hourFrequency[hour]++
				
			}	
		}
			
	}

	console.log(`Frequências das horas: ${hourFrequency}`);

	for(let h = 0; h < 24; h++){
		let hourCell = selectedSheet.getRange("H2").getOffsetRange(0, h);
		let freqCell = selectedSheet.getRange("H3").getOffsetRange(0, h);
		// Formata a hora para garantir que tenha dois dígitos
		hourCell.setValue(h.toString().padStart(2, '0'));  // Linha 2, coluna H e em diante
		freqCell.setValue(hourFrequency[h]);  // Linha 3, coluna H e em diante
	}
		
}