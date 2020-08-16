function MetaToCSV() {
    this.requiredContext = "Needs to run in Bridge";
}

var menuID = "metatocsvMenuID";
var menuExportCommandID = "exportMetaMenuItemID";

//Reference to GUI Window.
var windowReference;

//CSV Export Versions
var exportShutterStock = false;
var exportPond5 = false;
var exportVideoblocks = false;
var exportAllMeta = false;

var genericObjMetaDesc = {
    filename: "",
    title: "",
    keywords: [],
    categories: [], //Shutterstock
    isEditorial: false,
    createDate: "",
    creatorTool: "",
    metadataDate: "",
    documentID: ""
};

var metadataObjectDescriptors = [];

//Debug Variables
var debugAllowAutoOverwrite = true;

MetaToCSV.prototype.run = function() {
    //Ensure we can run the script.
    var returnVal = true;
    if (!this.canRun()){
        returnVal =  false;
        return returnVal;
    }
    
    //Load the XMP external lib.
    LoadXMPLibrary();

    //Create our window layout and store reference to it.
    windowReference = BuildUIWindow();

    //Create new menu.
    var metaToCSVMenu = new MenuElement("menu", "Meta2CSV", "after Help", menuID);
    //Create export metadata menu item.
    var exportMetaMenuItem = new MenuElement("command", "Export Metadata", "at the end of " + menuID, menuExportCommandID);
    //When click show export ment.
    exportMetaMenuItem.onSelect = function (m) {
        windowReference.show();
        AggregateMetadataFromCurrentSelection();
    }

    return returnVal;
}

function AggregateMetadataFromCurrentSelection() {
    for (var i = 0; i < app.document.selections.length; i++) {
        
        var thumbnail = app.document.selections[i];
        if (thumbnail.hasMetadata) {
            var selectedFile = thumbnail.spec;
            var METAObjDesc;
            
            var xmpFile = new XMPFile(selectedFile.fsName, XMPConst.UNKNOWN, XMPConst.OPEN_FOR_READ);
            var xmpMeta = xmpFile.getXMP();
            
            //File Name
            var fileNameSplit = xmpFile.getFileInfo().filePath.split("/");
            METAObjDesc.filename = fileNameSplit[fileNameSplit.length - 1];
            
            METAObjDesc.createDate = xmpMeta.getProperty(XMPConst.NS_XMP, "CreateDate");
            METAObjDesc.creatorTool = xmpMeta.getProperty(XMPConst.NS_XMP, "CreatorTool", XMPConst.STRING);
            METAObjDesc.metadataDate = xmpMeta.getProperty(XMPConst.NS_XMP, "MetadataDate", XMPConst.XMPDATE);
            METAObjDesc.documentID = xmpMeta.getProperty(XMPConst.NS_XMP_MM, "DocumentID", XMPConst.STRING);
            METAObjDesc.title = ConvertXMPArray(XMPConst.NS_DC, "title");
            var keywordsRaw = ConvertXMPArray(XMPConst.NS_DC, "subject");
            
            
            for (var keyword in keywordsRaw) {
                if (keyword.charAt(0) == '_') { }   //Internal Keyword, ignore it
                else if (keyword.charAt(0) == "!") {//Shutterstock Keyword
                    var cleanKw = keyword.replace("!", "");
                    if (cleanKw == "EDITORIAL") {
                        METAObjDesc.isEditorial = true;
                    } else if (cleanKw.charAt(0) == "_") {
                        METAObjDesc.categories.push(cleanKw.substring(1, cleanKw.length));
                    }
                } else {
                    METAObjDesc.push(keyword);
                }
            }
        }
    }

    function ConvertXMPArray(ns, property){
        var array = [];
        var itemCount = xmpMeta.countArrayItems(ns, property)
        if (itemCount > 0){
            for (var index = 1; index <= itemCount; index++) {
                var item = String(xmpMeta.getArrayItem(XMPConst.NS_DC, property, index));
                array.push(item);
            }
        }
        return array;
    }
}

function DoExport() {
    if (exportShutterStock) {
        var outputPath = "~/Desktop/MetadataShutterstock" + Number(new Date());
        var file = File(outputPath);

        CheckFileForWriting(file, outputPath);

        var outputStream;
        if (file !== '') {
            outputStream = file.open('w', undefined, undefined);
            file.encoding = "UTF-8";
            file.linefeed = "Macintosh";
        }

        if (outputStream !== false) {
            file.writeln("Filename, Title / Description, Keywords, Categories, Editorial");

            for (var element in metadataObjectDescriptors) {
                file.writeln(CSV_Quote(element.filename));
                file.writeln(CSV_Quote(element.title));
                file.writeln(CSV_Quote(element.keywords));
                file.writeln(CSV_Quote(element.categories));
                if (element.isEditorial){
                    file.writeln(CSV_Quote("yes"));
                } else {
                    file.writeln(CSV_Quote("no"));
                }
            };

            file.close();
        }
    }
    if (exportPond5) {
        var outputPath = "~/Desktop/MetadataPond5" + Number(new Date());
        var file = File(outputPath);

        CheckFileForWriting(file, outputPath);

        var outputStream;
        if (file !== '') {
            outputStream = file.open('w', undefined, undefined);
            file.encoding = "UTF-8";
            file.linefeed = "Macintosh";
        }

        if (outputStream !== false) {

        }
    }
    if (exportVideoblocks) {
        var outputPath = "~/Desktop/MetadataVideoblocks" + Number(new Date());
        var file = File(outputPath);

        CheckFileForWriting(file, outputPath);

        var outputStream;
        if (file !== '') {
            outputStream = file.open('w', undefined, undefined);
            file.encoding = "UTF-8";
            file.linefeed = "Macintosh";
        }

        if (outputStream !== false) {

        }
    }
    if (exportAllMeta) {
        var outputPath = "~/Desktop/MetadataAll" + Number(new Date());
        var file = File(outputPath);

        CheckFileForWriting(file, outputPath);

        var outputStream;
        if (file !== '') {
            outputStream = file.open('w', undefined, undefined);
            file.encoding = "UTF-8";
            file.linefeed = "Macintosh";
        }

        if (outputStream !== false) {
            file.writeln("Filename, Title / Description, Keywords, CreateDate," +
                         "CreatorTool, MetadataDate, DocumentID");

            for (var element in metadataObjectDescriptors) {
                file.writeln(CSV_Quote(element.filename));
                file.writeln(CSV_Quote(element.title));
                file.writeln(CSV_Quote(element.keywords));
                file.writeln(CSV_Quote(element.createDate));
                file.writeln(CSV_Quote(element.creatorTool));
                file.writeln(CSV_Quote(element.metadataDate));
                file.writeln(CSV_Quote(element.documentID));
            };

            file.close();
        }
    } 
}

function BuildCSVLine(data) {
    $.writeln(data);
    var line = "";
    var length = data.length;
    for (var index = 0; index < length; index++) {
        if (index != 0 && index != length) {
            line += ", ";
         } 
         line += data;  
    }    
    return line;
}

function CSV_Quote(data) {
    return "\"" + data + "\"";
}

function BuildUIWindow() {
    //Create our window.
    var window = new Window("dialog", "MetaToCSV", [100,100, 480, 480]);
    windowReference = window;
    //Panel to hold controls.
    window.panel = window.add("panel", [5, 5, 255, 130], "MetaToCSV");
    //Ok button.
    window.panel.okBtn = window.panel.add("button", [15,65,105,86], "OK");
    window.panel.okBtn.onClick = function() {window.close();};
    //Checkboxes to handle which CSVs to export.
    window.panel.checkboxShutterstock = window.panel.add("checkbox", [10, 10, 120, 25], "Shutterstock");
    window.panel.checkboxShutterstock.value = false;
    window.panel.checkboxShutterstock.onClick = function() { exportShutterStock = window.panel.checkboxShutterstock.value;};

    window.panel.checkboxVideoblocks = window.panel.add("checkbox", [10, 30, 120, 45], "Videoblocks");
    window.panel.checkboxVideoblocks.value = false;
    window.panel.checkboxVideoblocks.onClick = function() { exportVideoblocks = window.panel.checkboxVideoblocks.value;};

    window.panel.checkboxPond5 = window.panel.add("checkbox", [10, 50, 120, 65], "Pond5");
    window.panel.checkboxPond5.value = false;
    window.panel.checkboxPond5.onClick = function() { exportPond5 = window.panel.checkboxPond5.value;};

    return window;
}

function LoadXMPLibrary(){
    // Load the XMP Script library for metadata retreival
    if( xmpLib == undefined ) 
    {
        //$.writeln("XMP Lib is undefined, loading...");
       if( Folder.fs == "Windows" )
        {
            var pathToLib = Folder.startup.fsName + "/AdobeXMPScript.dll";
        } 
        else 
        {
            var pathToLib = Folder.startup.fsName + "/AdobeXMPScript.framework";
        }

        var libfile = new File( pathToLib );
        var xmpLib = new ExternalObject("lib:" + pathToLib );
        $.writeln("XMP Lib loaded.");
    }
}

function CheckFileForWriting(file, outputPath) {
    if (!file.exists) {
        file = new File (outputPath);
    } else {
        //if it exists ask the user if it should be overwritten
        if (debugAllowAutoOverwrite) {
            var res = Window.prompt("The file already exists. Overwrite it");
            //if the user hits no stop the script
            if (res !== true) {
            return;
            }  
        }      
     }
}

MetaToCSV.prototype.canRun = function() {
    if ((BridgeTalk.appName == "bridge") && (app.document.selectionsLength > 0)) {
        if ((MenuElement.find(this.menuID)) && (MenuElement.find(this.menuExportCommandID))) {
            $.writeln("Error:Menu element already exists!\nRestart Bridge to run this snippet again.");
			return false;
        }

        return true;
    }

    if (app.document.selectionsLength == 0) {
        prompt("No files selected! Select one or more files to export metadata.");
    } else {
        $.writeln("ERROR: Cannot run MetaToCSV");
        $.writeln(this.requiredContext);
    }

    return false;
}

if (typeof(MetaToCSV_unittest)== "undefined") {
    new MetaToCSV().run();
}
