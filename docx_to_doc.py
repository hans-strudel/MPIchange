import sys
import os
from os import listdir
import comtypes.client

formatDocx = 12

path = os.path.abspath(sys.argv[1])
files = listdir(path)
#in_file = os.path.abspath(sys.argv[1])
#out_file = os.path.abspath(sys.argv[2])
print files
print path

word = comtypes.client.CreateObject('Word.Application')
for i in range(len(files)):
	if not files[i].endswith('.docx'):
		continue
	in_file = path + '\\' + files[i]
	print os.path.abspath('docx//' + files[i].replace('.doc', '.docx'))
	out_file = os.path.abspath('doc//' + files[i].replace('.docx', '.doc'))
	
	doc = word.Documents.Open(in_file)
	doc.SaveAs(out_file, FileFormat=0)
	doc.Close()

word.Quit()