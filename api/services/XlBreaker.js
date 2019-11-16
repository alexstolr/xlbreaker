const XlsxPopulate = require("xlsx-populate");

module.exports = {
  openExcelFile: async function(){
    var pass;
    var found = false;
    for(pass = 0; pass <= 9999; pass++){
      if(!found){
        var [err,ans] = await execute(foo(pass.toString(), fileName));
        if(err){
            sails.log.info("something went wrong!");
        }
        else if(ans){
              sails.log.info(`*********opened ${fileName}.xlsx with password ${pass}*********`);
              found = true;
              break;
        }
      }
    }
    return res.json(pass);
  }
};

function foo(num, fileName){
    return XlsxPopulate.fromFileAsync(`./${fileName}.xlsx`, { password: num })
        .then(workbook => {
            return true;
        })
        .catch((error) => {
            if(parseInt(num, 10) % 20 === 0){
                sails.log.info(`bad passwords: 0 - ${num}`);
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

