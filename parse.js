var path = require('path'),
	write = require('./overwrite.js'),
	prompt = require('prompt-sync')(),
	fs = require('fs'),
	cp = require('child_process')
	
var outputDir = 'C:\\Users\\HansStrausl\\Desktop\\MPICHANGE\\outputs\\'
	
function handlePDF(filePath){
	var name = path.parse(filePath).name
	var name = name.replace(/M/g, function(m,ind){
		if (ind == 3) return ''
		return 'M'
	})
	var rev
	name.replace(/[0-9]/, function(m,ind){
		if (ind > 6){
			rev = name.substr(ind, name.length)
			name = name.substring(0, ind-1)
		}
		return ''
	})
	name = prompt('Name (' + name + '): ', name)
	rev = prompt('Rev (' + rev + '): ', rev)
	console.log(name, rev)
	write(filePath, outputDir + "MPI-" + name + ' Rev ' + rev + '.pdf', name, rev)
}
glob = 0

traverse(process.argv[2])


function traverse(dir){
	var files = fs.readdirSync(dir)
	files.forEach(function(elem,index,array){
		
		console.log(elem)
		try {
			var isDir = fs.statSync(dir + '\\' + elem).isDirectory()
			if (isDir){
				//console.log(elem)
				traverse(dir + '\\' + elem)
			}
			var info = path.parse(elem)
			//if (info.name.indexOf('MCEM') == 0){
				//console.log(elem)
				if (info.ext == '.pdf'){
					//outputDir = dir + '\\'
					//handlePDF(dir + '\\' + elem)
				}
				if (info.ext == '.hwp'){
					console.log('HWP FOUND')
					console.log('C:\\Users\\HansStrausl\\Desktop\\MPICHANGE\\allhwps\\' + glob + '__' + elem)
					//console.log('convert_to_docx.py "' + dir + '\\' +  elem + '" "' + outputDir + info.name + '.docx"')
					x = fs.readFileSync(dir + '\\' +  elem)
					fs.writeFileSync('C:\\Users\\HansStrausl\\Desktop\\MPICHANGE\\allhwps\\' + glob + '__' + elem, x)
					fs.appendFileSync('C:\\Users\\HansStrausl\\Desktop\\MPICHANGE\\locs.txt', dir + '\\' + glob++ + '__' + elem + '\r\n')
					//copyFile(dir + '\\' +  elem, outputDir + info.name + '.doc', ()=>{})
					//cp.exec('convert_to_docx.py "' + dir + '\\' +  elem + '" C:\\Users\\HansStrausl\\Desktop\\MPICHANGE\\outputs\\a\\MPI-' + info.name + '.docx"')
	
				}
				if (info.ext == '.docx'){
					//console.log('Editing DOCX')
					//console.log(info.base)
					//x = fs.readFileSync(dir + '\\' +  elem)
					//fs.writeFileSync('C:\\Users\\HansStrausl\\Desktop\\MPICHANGE\\docx\\' + elem, x)
					//name = prompt('Name : ')
					//rev = prompt('Rev : ')
					//cp.exec('editDocx.py "' + dir + '\\' +  elem + '" "' + 
					//'C:\\Users\\HansStrausl\\Desktop\\MPICHANGE\\outputs\\a\\MPI-' + info.name + '.docx" "' + name + '" "' + rev + '"')
					
				}
			//}
		} catch(e) {
			console.log(e)
		}
		
		
	})
}

function copyFile(source, target, cb) {
  var cbCalled = false;

  var rd = fs.createReadStream(source);
  rd.on("error", function(err) {
    done(err);
  });
  var wr = fs.createWriteStream(target);
  wr.on("error", function(err) {
    done(err);
  });
  wr.on("close", function(ex) {
    done();
  });
  rd.pipe(wr);

  function done(err) {
    if (!cbCalled) {
      cb(err);
      cbCalled = true;
    }
  }
}