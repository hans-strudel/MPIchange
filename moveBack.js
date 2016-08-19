var path = require('path'),
	fs = require('fs')

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
		console.log(target + ' COPIED')
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

dirs = fs.readFileSync('locs.txt', 'utf8').split('\r\n')
files = fs.readdirSync(process.argv[2])

files.forEach(function(file,ind){
	var f = path.parse(file)
	
	if (f.ext.toUpperCase() == '.DOC'){
		dirs.forEach(function(p,i,a){
			if (f.name == path.parse(p).name){
				console.log(path.parse(p).dir + '\\' + path.parse(p).base.replace('.hwp','.doc'))
				setTimeout(function(){
					copyFile(process.argv[2] + '\\' + file, path.parse(p).dir + '\\' + path.parse(p).base.replace('.hwp','.doc'), (e)=>{})	
				}, 1000+100*ind)
			}
		})
	}	
})

