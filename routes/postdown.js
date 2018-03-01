var express = require('express');
var router = express.Router();
var xlsx = require('node-xlsx');
var fs = require('fs');

var fileName, tableHead;

router.post('/', function (req, res, next){
	fileName = req.body.fileName || 'excel' + '.xlsx';
	tableHead = req.body.tableHead;

	res.send();
})


/* GET users listing. */
router.get('/excel', function (req, res, next){

	var fReadStream;

	// this is the test data
	var data = [[1, 2, 3], [true, false, null, 'sheetjs'], ['你好', 'bars', new Date('2014-02-19T14:30Z'), '0.3'], ['baz', null, 'qux']];
	data.unshift(tableHead.split(','));
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
