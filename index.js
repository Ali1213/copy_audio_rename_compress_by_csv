const fs = require('fs');
const path = require('path');
const Excel = require('exceljs');
var archiver = require('archiver');

let excelFilename = "1.csv";
let outputPath = path.join(__dirname, 'output');

fs.stat(outputPath, (err, stat) => {
    if (err) {
        fs.mkdirSync(outputPath);
    }
})

const readDir = function (filepath = __dirname) {
    return new Promise((rs, rj) => {
        fs.readdir(path.join(filepath), (err, fs) => {
            if (err) return rj(e);

            rs(fs.map(it => path.join(filepath, it)));
        });
    });
};

const checkFile = function (files) {
    let excelFileNum = 0,
        onOff = false,
        efile = '',
        audios = [];
    files.forEach(item => {
        let ext = path.extname(item);
        if (!ext) return;

        if (path.basename(item) === excelFilename) {
            onOff = true;
            efile = item;
        }

        if (/^\.(xlsx|csv|xls)$/.test(ext)) {
            efile = item;
            excelFileNum++;
        } else if (/^\.(mp3|wav)$/.test(ext)) {
            audios.push(item);
        }
    });

    if (excelFileNum > 1 && !onOff) {
        throw Error('请检查该文件夹，不允许该文件夹内存在多个excel文件');
    }

    return [efile, audios]
};

const readExcel = function ([efile, audios]) {
    const workbook = new Excel.Workbook();
    if (path.extname(efile) === '.csv') {
        return new Promise((rs, rj) => {
            fs.readFile(efile, 'utf-8', (err, data) => {
                if (err) return rj('讀取csv文件失敗');
                rs([data.match(/\d{8,}/g), audios]);
            })
        })
    } else {
        throw Error('暂时不支持非csv格式的文件');
    }
    // return workbook.xlsx.readFile(efile)
    //     .then(function (result) {
    //         console.log(workbook);
    //         console.log(result)
    //     }).catch(e=>console.log(e));
}

const copyFile = function (target, output) {
    return new Promise((rs, rj) => {
        let write = fs.createWriteStream(output);
        write.on('close', () => {
            console.log(`${output}已经生成`)
            rs();
        });
        write.on('error', () => {
            rj('复制文件出错');
        });
        fs.createReadStream(target).pipe(write);

        // fs.readFile(target,(e,data)=>{
        //     if(e)return rj(e);
        //     fs.writeFile(output,data,(e)=>{
        //         if(e)return rj(e);
        //         rs(data);
        //     })
        // })
    });
}

const copyFileRename = function ([tels, audios]) {
    let allFiles = [];
    let len = audios.length;
    let promiseAll = [];
    tels.forEach((item, i) => {
        let outputFilename = path.join(outputPath, path.basename(audios[i % len]).replace(/\d+(?=\.)/, ($$, $1) => {
            return item
        }));
        allFiles.push({ "target": audios[i % len], "output": outputFilename });
    });
    promiseAll = [...allFiles];
    return new Promise((rs, rj) => {
        (function copyAll(arr) {
            if (!arr.length) {
                return rs();
            }
            copyFile(arr[0].target, arr[0].output)
                .then((rs) => {
                    arr.shift();
                    copyAll(arr);
                }).catch(e => rj(e));
        })(allFiles)
    }).then(() => promiseAll);
}


const getZip = function (allFiles) {
    var output = fs.createWriteStream(__dirname + '/example.zip');
    var archive = archiver('zip', {
        zlib: { level: 9 } // Sets the compression level. 
    });
    output.on('close', function () {
        console.log('压缩包总体积'+ archive.pointer()  + ' total bytes');
        console.log('打成压缩吧成功！');
    });
    archive.on('error', function (err) {
        throw err;
    });

    archive.pipe(output);

    allFiles.forEach(item=>archive.append(fs.createReadStream(item.output), { name: path.basename(item.output) }))

    archive.finalize();
    
}

readDir()
    .then(checkFile)
    .then(readExcel)
    .then(copyFileRename)
    .then(getZip)
    .catch(e => console.log(e));