# some file have a linear table

import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import os
from os import listdir
from docx import Document
from docx.shared import Pt
from docx.package import Package
import gc
import xml.etree.ElementTree as ET
from docx.oxml import parse_xml


xmlschema = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
path = os.path.abspath(sys.argv[1])
files = listdir(path)
assy = sys.argv[2]
outfolder = sys.argv[3]

def remove_row(table, row):
	tbl = table._tbl
	tr = row._tr
	tbl.remove(tr)

for i in range(100, len(files)):
	if not files[i].endswith('.docx'):
		continue
	in_file = path + '\\' + files[i]
	print files[i]
	out_file = os.path.abspath(outfolder + '//MPI-' + files[i])
	document_part = Package.open(in_file)
	doc = document_part.main_document_part.document
	
	
	
	table = doc.tables[0]
	
	if ("Normal" in doc.styles): # set styling
		doc.styles['Normal'].font.name = 'Batang'
		doc.styles['Normal'].font.bold = True
		doc.styles['Normal'].font.size = Pt(12)
	else: # use deprecated backup method
		doc.styles[0].font.name = 'Batang'
		doc.styles[0].font.bold = True
		doc.styles[0].font.size = Pt(12)
	
	if (len(table.rows) > 3):
		remove_row(table, table.rows[3])
	if (table._cells[0].text.find("CUSTOMER") < 0):
		continue
	for i in range(len(table._cells) - 1):
		if (table._cells[i].text.upper().find("ASSY") > -1):
			assyGuess = assy + table._cells[i+1].text.strip() + "-TK"
			assyPos = i+1
		if (table._cells[i].text.upper().find("REV #") > -1):
			assyRevGuess = (table._cells[i+1].text.strip() or "A") + "-A"
			assyRevPos = i+1
		if not table._cells[i].text.strip():
			table._cells[assyRevPos].text = ''
			table._cells[assyRevPos-1].text = ''
			table._cells[i].text = "Rev #"
			assyRevPos = i + 1
			break
	
	assyNum = raw_input(assyGuess + " : ") or assyGuess
	assyRev = raw_input(assyRevGuess + " : ") or assyRevGuess
	
	table._cells[assyPos].text = assyNum
	table._cells[assyRevPos].text = assyRev
	
	new_xml = doc._part._element.xml
	
	f = int(new_xml[new_xml.find('<w:headerReference r:id="rId')+len('<w:headerReference r:id="rId')]) + 1
	g = int(new_xml[new_xml.find('<w:footerReference r:id="rId')+len('<w:footerReference r:id="rId')]) + 1
	
	for i in range(f,70):
		new_xml = new_xml.replace('<w:headerReference r:id="rId' + str(i) + '" w:type="default"/>', '')
	for i in range(g,70):
		new_xml = new_xml.replace('<w:footerReference r:id="rId' + str(i) + '" w:type="default"/>', '')
	
	doc._part._element = parse_xml(new_xml)
	
	g = gc.get_referrers(document_part)
	length = len(g)
	for a in range(0, len(g)-1):
		#print "Length of g ", len(g)
		#print a
		if ("_partname" in g[a]):
			#print g[a]["_partname"]
			if (g[a]["_partname"].find("header1.xml") > -1):
				g[a]["_blob"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14"><w:p w:rsidR="008A5523" w:rsidRDefault="008A5523" w:rsidP="00AF4909"><w:pPr><w:pStyle w:val="Header"/><w:jc w:val="both"/></w:pPr><w:proofErr w:type="spellStart"/><w:r w:rsidRPr="008A5523"><w:rPr><w:i/></w:rPr><w:t>Bestronics</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r w:rsidRPr="008A5523"><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> Inc.</w:t></w:r><w:r><w:tab/></w:r><w:r w:rsidRPr="008A5523"><w:rPr><w:sz w:val="32"/></w:rPr><w:t>MANUFACTURING PROCESSING INSTRUCTIONS</w:t></w:r></w:p></w:hdr>'
			elif (g[a]["_partname"].find("header") > -1):
			#	print "header"
				g[a]["_blob"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14"><w:p w:rsidR="008A5523" w:rsidRDefault="008A5523" w:rsidP="00AF4909"><w:pPr><w:pStyle w:val="Header"/><w:jc w:val="both"/></w:pPr><w:proofErr w:type="spellStart"/><w:r w:rsidRPr="008A5523"><w:rPr><w:i/></w:rPr><w:t>Bestronics</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r w:rsidRPr="008A5523"><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> Inc.</w:t></w:r><w:r><w:tab/></w:r><w:r w:rsidRPr="008A5523"><w:rPr><w:sz w:val="32"/></w:rPr><w:t>MANUFACTURING PROCESSING INSTRUCTIONS</w:t></w:r></w:p></w:hdr>'
			if (g[a]["_partname"].find("footer1.xml") > -1):
				g[a]["_blob"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14"><w:p w:rsidR="00D6537D" w:rsidRDefault="00D6537D"><w:pPr><w:pStyle w:val="Footer"/></w:pPr><w:r><w:t>' + \
								assyNum + ' Rev ' + assyRev + '</w:t></w:r><w:r><w:tab/></w:r><w:r><w:tab/><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE   \* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType="end"/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t xml:space="preserve"> of </w:t></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:instrText xml:space="preserve"> NUMPAGES  \* Arabic  \* MERGEFORMAT </w:instrText></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>'
			elif (g[a]["_partname"].find("footer") > -1):
				g[a]["_blob"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14"><w:p w:rsidR="00D6537D" w:rsidRDefault="00D6537D"><w:pPr><w:pStyle w:val="Footer"/></w:pPr><w:r><w:t>' + \
								assyNum + ' Rev ' + assyRev + '</w:t></w:r><w:r><w:tab/></w:r><w:r><w:tab/><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE   \* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType="end"/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t xml:space="preserve"> of </w:t></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:instrText xml:space="preserve"> NUMPAGES  \* Arabic  \* MERGEFORMAT </w:instrText></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>'
		
	doc.save(out_file)