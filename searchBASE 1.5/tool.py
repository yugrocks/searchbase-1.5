#This module is used to make a batch file to get metadata

import os

def createtool():
   try:
      os.chdir(r"C:\database1")
   except:
      os.mkdir(r"C:\database1")
      os.chdir(r"C:\database1")
   a=open("tool.bat",'w')
   a.write(r"""@if (@X)==(@Y) @end /* JScript comment
    @echo off

    rem :: the first argument is the script name as it will be used for proper help message
    cscript //E:JScript //nologo "%~f0" %*

    exit /b %errorlevel%

@if (@X)==(@Y) @end JScript comment */

////// 
FSOObj = new ActiveXObject("Scripting.FileSystemObject");
var ARGS = WScript.Arguments;
if (ARGS.Length < 1 ) {
 WScript.Echo("No file passed");
 WScript.Quit(1);
}
var filename=ARGS.Item(0);
var objShell=new ActiveXObject("Shell.Application");
/////


//fso
ExistsItem = function (path) {
    return FSOObj.FolderExists(path)||FSOObj.FileExists(path);
}

getFullPath = function (path) {
    return FSOObj.GetAbsolutePathName(path);
}
//

//paths
getParent = function(path){
    var splitted=path.split("\\");
    var result="";
    for (var s=0;s<splitted.length-1;s++){
        if (s==0) {
            result=splitted[s];
        } else {
            result=result+"\\"+splitted[s];
        }
    }
    return result;
}


getName = function(path){
    var splitted=path.split("\\");
    return splitted[splitted.length-1];
}
//

function main(){
    if (!ExistsItem(filename)) {
        WScript.Echo(filename + " does not exist");
        WScript.Quit(2);
    }
    var fullFilename=getFullPath(filename);
    var namespace=getParent(fullFilename);
    var name=getName(fullFilename);
    var objFolder=objShell.NameSpace(namespace);
    var objItem=objFolder.ParseName(name);
    //https://msdn.microsoft.com/en-us/library/windows/desktop/bb787870(v=vs.85).aspx
    WScript.Echo(fullFilename + " : ");
    WScript.Echo(objFolder.GetDetailsOf(objItem,-1));

}

main();""")



   a.close()


