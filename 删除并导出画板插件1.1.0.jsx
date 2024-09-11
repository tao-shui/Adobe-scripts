#target illustrator
// app.bringToFront();
function deleteObjectsOnArtboard(artboardIndex) {
    var doc = app.activeDocument;
    var artboards = doc.artboards;

    if (artboardIndex < 1 || artboardIndex > artboards.length) {
        alert("画板索引超出范围");
        return;
    }

    var artboard = artboards[artboardIndex - 1];
    var bounds = artboard.artboardRect;
    var left = bounds[0];
    var top = bounds[1];
    var right = bounds[2];
    var bottom = bounds[3];
    if (doc.pageItems.length >= 0) {
        for (var i = doc.pageItems.length - 1; i >= 0; i--) {
            var item = doc.pageItems[i];
            var itemBounds = item.geometricBounds;
            if (itemBounds[0] >= left && itemBounds[2] <= right && itemBounds[1] <= top && itemBounds[3] >= bottom) {
                try {
                    if (item.textFrames) {
                        for (var j = 0; j < item.textFrames.length; j++) {
                            var textFrame = item.textFrames[j];
                            if (textFrame.textRange.characterAttributes.textFont == null) {
                                // 替换缺失字体
                                textFrame.textRange.characterAttributes.textFont = app.textFonts.getByName("Arial"); // 或其他默认字体
                            }
                        }
                    }
                    if (item) {
                        item.remove();
                    }
            } catch (e) {
                alert("错误: " + e.message);
            }
            }
        }
    }
}
res ="dialog { \
properties:{ closeButton:true, maximizeButton:true,			\
minimizeButton:false, resizeable:false },					\
text:'批处理-选择文件',\
group: Group{orientation: 'column',alignChildren:'left',\
folderO:Group{ orientation: 'row', \
b: Button {text:'待处理文件夹', properties:{name:'open'} ,helpTip:'选择您需要处理的文件所在的文件夹'},\
s: EditText { text:'', preferredSize: [360, 20] },\
},\
folderS:Group{ orientation: 'row', \
b: Button {text:'输出图像至', properties:{name:'save'} ,helpTip:'选择您处理好的文件要保存至的文件夹'},\
s: EditText { text:'', preferredSize: [360, 20] },\
},\
gg: Group{orientation: 'column',alignChildren:'left' },\
timeline:Progressbar{bounds:[0,0,400,10] , minvalue:0,maxvalue:100}\
aa: Button { text:'START'}, \
}\
}";
var mengPoint="";
// var mengColor =new SolidColor;
// mengColor.rgb.red =0;
// mengColor.rgb.green =0;
// mengColor.rgb.blue =0;

win = new Window (res);
// 打印输出窗口
win.myText = win.group.gg.add("edittext",[0,0,500,300],'~~~',{multiline:true, readonly:false});
// // 打开文件夹的操作
var inputfolder=win.group.folderO;
var outputfolder=win.group.folderS;
var folderOpen=win.group.folderO
var folderSave=win.group.folderS
folderOpen.b.onClick = function() {
var defaultFolder = folderOpen.s.text;
var testFolder = new Folder(defaultFolder);
if (!testFolder.exists) {
defaultFolder = "~";
}
var selFolder = Folder.selectDialog("选择待处理文件夹", defaultFolder);
if ( selFolder != null ) {
folderOpen.s.text = selFolder.fsName;
folderOpen.s.helpTip = selFolder.fsName.toString();
}
}
folderSave.b.onClick = function() {
var defaultFolder = folderSave.s.text;
var testFolder = new Folder(defaultFolder);
if (!testFolder.exists) {
defaultFolder = "~";
}
var selFolder = Folder.selectDialog("选择输出的文件夹", defaultFolder);
if ( selFolder != null ) {
folderSave.s.text = selFolder.fsName;
folderSave.s.helpTip = selFolder.fsName.toString();
}
}

win.group.aa.onClick=function(){
    var type = ".jpg"
    var outputname = ""
    var myText="";
    var inputfolder=Folder(win.group.folderO.s.text);
    var outputfolder=Folder(win.group.folderS.s.text);
    var files = inputfolder.getFiles("*.ai");
    win.group.timeline.value =0;
    var k=100/files.length;
    for (var i = 1; i <= files.length; i++) {
        try {
        file = files[i-1];
        app.open(file);
        // 删除画板上的所有对象
        var doc = app.activeDocument;
        var artboards = doc.artboards;
        var art1 = artboards.length;
        // 删除除了第一个画板以外的所有画板对象，逆向删除
        for (var j = 0; j < artboards.length; j++) {
                if (art1 > 1) {
                    deleteObjectsOnArtboard(art1);
                }
                art1 -=1;
        }
        // 删除除了第一个画板以外的所有画板，顺序删
        for (k = 0; k < app.activeDocument.artboards.length; k++) {
                if (app.activeDocument.artboards.length > 1) {
                    app.activeDocument.artboards[1].remove();
                }
        }
        // 对应第五个图层的油画底色
            if (app.activeDocument.layers.length > 5) {
                if (app.activeDocument.layers[5].name == "油画底色") {
                    app.activeDocument.layers[5].remove();
                    }
            }



        if (app.activeDocument.name.indexOf('.') < 0) {
            outputname = app.activeDocument.name.indexOf + type;
        } else {
            var dot = app.activeDocument.name.lastIndexOf('.');
            outputname = app.activeDocument.name.substring(0,dot)
            outputname += type
        }

        if (outputfolder) {
            var outputfile = new File(outputfolder + "/" + outputname);
            var imgo = new ImageCaptureOptions();
            imgo.resolution = 700;
            imgo.antiAliasing = true;
            var doc = app.activeDocument;
            var activeAB = doc.artboards[doc.artboards.getActiveArtboardIndex()];
            doc.imageCapture(outputfile,activeAB.artboardRect,imgo);
            var re = outputfolder + "/" + outputname + "已完成"
            win.myText.text=myText.replace(re, "");

        }
        // app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);


    } catch (e)  {
        alert("错误: " + e.message);
        app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);
        myText ="\r\n";
        win.group.timeline.value =win.group.timeline.value + k;

    } finally {
        app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);
        myText ="\r\n";
        win.group.timeline.value =win.group.timeline.value + k;
    }
    }
}
//////////////
win.center();
win.show();