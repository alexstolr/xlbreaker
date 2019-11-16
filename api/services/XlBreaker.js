const XlsxPopulate = require("xlsx-populate");
const path = require('path');
const fs = require('fs');

module.exports = {
  breaklAllFiles: async function(){
    const directoryPath = path.join(__dirname, '../../lockedfiles');
    fs.readdir(directoryPath, function (err, files) {
      if (err) {
          return console.log('Unable to scan directory: ' + err);
      } 
      files.forEach(async function (fileName) {
        sails.log.info(`Trying to open ${fileName}`);
        openExcelFile(fileName);
      });
    });
  },
  
};

async function openExcelFile(fileName){
  var pass;
  var found = false;
  for(pass = 0; pass <= 9999; pass++){
    if(!found){
      var [err,ans] = await execute(foo(pass.toString(), fileName));
      if(err){
          sails.log.info("something went wrong!");
      }
      else if(ans){
            sails.log.info(`*********opened ${fileName} with password ${pass}*********`);
            found = true;
            
            break;
      }
    }
  }
}

function foo(num, fileName){
    return XlsxPopulate.fromFileAsync(`../xlbreaker/lockedfiles/${fileName}`, { password: num })
        .then(workbook => {
            return true;
        })
        .catch((error) => {
            if(parseInt(num, 10) % 20 === 0){
                sails.log.info(`${fileName} bad passwords: 0 - ${num}`);
            }
            return false;
        });
}

async function execute(promise){
  return promise.then(data => {
    return [null,data];
  })
  .catch( err => {
    sails.log.info(`Error: ${err}`);
    return [err];
  });
 }

