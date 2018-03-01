var express = require('express');
var router = express.Router();
var xlsx = require('node-xlsx');
var fs = require('fs');

/* GET users listing. */
router.get('/', function (req, res, next){

	var fReadStream;
	var fileName = (req.query.fileName || 'excel') + '.xlsx';

	// this is the test data
	var data = [[1, 2, 3], [true, false, null, 'sheetjs'], ['你好', 'bars', new Date('2014-02-19T14:30Z'), '0.3'], ['baz', null, 'qux']];

	// returns a buffer
	var buffer = xlsx.build([{name: 'mySheet', data: data}]);

	// fs 生成文件
	fs.writeFile(fileName, buffer, function (err){

		if(err){
			return console.error(err);
		}

		res.set({
			"Content-type":"application/octet-stream",
            "Content-Disposition":"attachment;filename="+encodeURI(fileName)
		});

		fReadStream = fs.createReadStream(fileName);
		fReadStream.on("data", (chunk) => res.write(chunk, "binary"));
		fReadStream.on("end", function (){

			res.end();

			// fs 删除生成的文件
			fs.unlink(fileName, function (err){
				if(err){
					return console.error(err);
				}
			})
		})

	})
})

module.exports = router;
