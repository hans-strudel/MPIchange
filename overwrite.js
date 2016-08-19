var hummus = require('hummus'),
	path = require('path')



module.exports = write = function(inputFile, outputFile, assyName, assyRev){
	var pdfWriter = hummus.createWriterToModify(inputFile, {
		modifiedFilePath: outputFile
	});
	
	var pathFillOptions = {color:0x00000000, colorspace:'cmyk', type:'fill'} // white
	var fontOptions = {font:pdfWriter.getFontForFile(__dirname + '\\calibri.ttf'),size:14,colorspace:'gray',color:0x00}
	
	var pdfLength = hummus.createReader(inputFile).getPagesCount()
	console.log("# of pages: " + pdfLength)
	for (var page = 0; page < pdfLength; page++){
		var pageModifier = new hummus.PDFPageModifier(pdfWriter, page)
		pageModifier.startContext()
		if (page == 0){ // first page has the table
			pageModifier.getContext().drawRectangle(415, 629, 110, 21, pathFillOptions)
			pageModifier.getContext().writeText(
				assyRev,
				420, 632,
				fontOptions
			);
			pageModifier.getContext().drawRectangle(415, 581, 110, 21, pathFillOptions)
			pageModifier.getContext().writeText(
				assyRev,
				420, 585,
				fontOptions
			);
		}
		pageModifier.getContext().drawRectangle(20, 20, 450,35, pathFillOptions)
		
		var botText = "MPI-" + assyName + ", Rev " + assyRev
		pageModifier.getContext().writeText(
			botText,
			30, 30,
			fontOptions
		);
		
		pageModifier.endContext().writePage();
	}
	pdfWriter.end();
}