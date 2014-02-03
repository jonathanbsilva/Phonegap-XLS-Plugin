//cordova.define("net.claudiomedeiros.xls.Xls", function(require, exports, module) {
    var xls =  {
        save: function(data, dirname, filename, sheetname, successCallback, errorCallback){
            /*  Modelo dos parâmetros
                var data = [
                    {id:"1", name:"claudio"} ,
                    {id:"2", name:"marta"} ,
                    {id:"3", name:"isabela"} 
                ];
                dirname = "ExcelAPI";
                filename = "file-example.xls";
                sheetname = "Plan1";
            */
            
            cordova.exec(
                successCallback,// callback de sucesso
                errorCallback,  // callback de erro
                'Xls',          // Classe java onde o código se encontra
                'saveXLS',      // Nome da action da classe que será chamada
                [{              // Array de parâmetros
                    "data"     : data,
                    "dirname"  : dirname,
                    "filename" : filename,
                    "sheetname": sheetname
                }]            
            );
        }
    }//fim do objeto

    module.exports = xls;
//});
