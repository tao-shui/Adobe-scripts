// 选择画板1上的所有对象并删除



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

var exportOptions = new ExportOptionsJPEG();
exportOptions.qualitySetting = 100;

var inputfolder = Folder.selectDialog("选择处理文件夹","~\\desktop");
var outputfolder = Folder.selectDialog("选择保存文件夹","~\\desktop");
try {
    if (!outputfolder || !inputfolder) {
    throw new Error("未选择文件夹");
    }
} catch (e) {
    alert("错误: " + e.message);
}

var type = ".jpg"
var outputname = ""
var files = inputfolder.getFiles("*.ai");
for (var i = 1; i <= files.length; i++) {
    try {
    file = files[i-1];
    app.open(file);
    // 删除画板上的所有对象
    var doc = app.activeDocument;
    var artboards = doc.artboards;
    var art1 = artboards.length;
    // 删除除了第一个画板以外的所有画板对象，逆向删除
    for (var i = 0; i < artboards.length; i++) {
        try {
            if (art1 > 1) {
                deleteObjectsOnArtboard(art1);
            }
            art1 -=1;
        } catch (e) {
        alert("错误: " + e.message);
    }
    }
    // 删除除了第一个画板以外的所有画板，顺序删
    for (i = 0; i < app.activeDocument.artboards.length; i++) {
        try {
            if (app.activeDocument.artboards.length > 1) {
                app.activeDocument.artboards[1].remove();
            }
        } catch (e) {
            alert("错误: " + e.message);
        }
    }
    // 对应第五个图层的油画底色
    try {
        if (app.activeDocument.layers.length > 5) {
            if (app.activeDocument.layers[5].name == "油画底色") {
                app.activeDocument.layers[5].remove();
                }
        }
        } catch (e) {
            alert("错误: " + e.message);
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

    }
    $.writeln(app.activeDocument);
    $.writeln (outputfolder);
    app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);


} catch (e) {

            alert("错误: " + e.message);
}
}
// catch (e) {
//     alert("错误: " + e.message);
//     // app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);
// }