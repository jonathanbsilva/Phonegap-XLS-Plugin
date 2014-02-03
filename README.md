Phonegap-XLS-Plugin
====

A Phonegap plugin to save XLS files

## How to use ##

    var data = [
        {id:"1", name:"claudio"} ,
        {id:"2", name:"marta"} ,
        {id:"3", name:"isabela"} 
    ];
    dirname = "ExcelAPI";
    filename = "file-example.xls";
    sheetname = "Plan1";
    
    window.xls.save( data, dirname, filename, sheetname, 
        function(){ alert('success'); }, 
        function(err){ alert(err); }
        );
    }

## Instalation ##

#### From github ####
    
    cordova plugin add https://github.com/klawdyo/Phonegap-XLS-Plugin.git
    
#### From file ####
    
Copy full directory to project root

    cordova plugin add <Path-to-Plugin's-directory>
    
## Dependences ##

[Phonegap's users permissions to File plugin](http://docs.phonegap.com/en/3.3.0/cordova_file_file.md.html#File)

#### Install File plugin ####

    cordova plugin add org.apache.cordova.file
    
## Thanks To ##

* [JExcelAPI Project](http://sourceforge.net/projects/jexcelapi/)
* [Using JExcelAPI in Android](http://www.kylebeal.com/2011/10/using-jexcelapi-in-an-android-app/)
