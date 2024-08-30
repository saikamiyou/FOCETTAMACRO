// 新規ファイルをデフォルト設定で作成
function createNewDocument() {
    var docWidth = 210; // 幅（mm単位）
    var docHeight = 297; // 高さ（mm単位）
    var docResolution = 300; // 解像度（dpi）
    var docName = "New Document";

    var doc = app.documents.add(
        UnitValue(docWidth, "mm"),
        UnitValue(docHeight, "mm"),
        docResolution,
        docName,
        NewDocumentMode.RGB,
        DocumentFill.WHITE
    );

    return doc;
}

// CSVファイルのパスを指定
var systemCsvFilePath = new File(File($.fileName).parent + "/system.csv");

// 相対パスを使用して基盤となるPSDファイルのパスを指定
var basePSDPath = "";

// 相対パスを使用してCSVファイルのパスを指定
var csvFilePath = "";

// 変数用のレイヤー数を設定
var numberOfLayers = 0; // ドキュメント内の変数用レイヤー数

var outputFolder = "";

// CSVを読み込む関数
function readCSV(file) {
    file.open('r'); // ファイルを読み取りモードで開く
    var content = file.read(); // ファイルの内容を読み込む
    file.close(); // ファイルを閉じる
    return content; // 読み込んだ内容を返す
}



// 相対パスを絶対パスに変換する関数
function convertRelativeToAbsolutePath(relativePath) {
    var scriptPath = new File($.fileName).path;
    var absolutePath = new File(scriptPath + "/" + relativePath).fsName;
    return absolutePath;
}

// レイヤーに画像を配置する関数
function replaceSmartObjectContent(layer, imagePath) {
    try {
        var file = new File(imagePath);
        if (!file.exists) {
            alert("File does not exist: " + imagePath);
            return;
        }

        // レイヤーがスマートオブジェクトであることを確認
        if (layer.kind != LayerKind.SMARTOBJECT) {
            alert("The selected layer is not a Smart Object.");
            return;
        }

        // 元のレイヤー名を保存
        var originalLayerName = layer.name;

        // 元のレイヤーのサイズをバックアップ
        var originalBounds = layer.bounds;
        var originalWidth = originalBounds[2] - originalBounds[0];
        var originalHeight = originalBounds[3] - originalBounds[1];

        // スマートオブジェクトの内容を置き換える
        var idplacedLayerReplaceContents = stringIDToTypeID("placedLayerReplaceContents");
        var desc = new ActionDescriptor();
        desc.putPath(charIDToTypeID("null"), file);
        executeAction(idplacedLayerReplaceContents, desc, DialogModes.NO);



		// 差し替えられたレイヤーのサイズを取得
		var newBounds = layer.bounds;
		var newWidth = newBounds[2] - newBounds[0];
		var newHeight = newBounds[3] - newBounds[1];

		// リサイズ処理: 新しいレイヤーの幅または高さが0でない場合のみ実行
		if (newWidth > 0 && newHeight > 0) {

		    // 縦横比を計算
		    var aspectRatio = newWidth / newHeight;

		    // 横幅と縦幅のどちらを基準にするかを判定
		    if (newWidth > newHeight) {  // 横長の場合
		        var targetWidth = originalWidth;
		        var targetHeight = originalWidth / aspectRatio; // 横幅に合わせた縦幅を計算
		    } else {  // 縦長の場合
		        var targetHeight = originalHeight;
		        var targetWidth = originalHeight * aspectRatio; // 縦幅に合わせた横幅を計算
		    }

		    // 元のサイズにリサイズ
		    layer.resize(
		        (targetWidth / newWidth) * 100,
		        (targetHeight / newHeight) * 100,
		        AnchorPosition.MIDDLECENTER
		    );
		}

        // 元のレイヤー名を再設定
        layer.name = originalLayerName;

    } catch (e) {
        throw new Error(
            "function: replaceSmartObjectContent" + "\n" +
            "Error executing action: " + e.message + "\n"
        );
    }
}

// データセットを適用する関数
function applyDataset(doc, dataset) {
// レイヤー名が `expectedLayerName` に一致するレイヤーを取得
        var max = 0;
        try {
            doc.artLayers.getByName("クッキー1");
            max =25;
        } catch (e) {
max =40;
        }
    for (var j = 1; j <= max ; j++) { // 25枚のレイヤーを処理する
        var csvColumnName = "image" + j;

        // レイヤー名が `expectedLayerName` に一致するレイヤーを取得
        var layer = null;
        try {
            layer = doc.artLayers.getByName("クッキー" + j);
        } catch (e) {
            // 一致するレイヤーが見つからない場合
			try {
	            layer = doc.artLayers.getByName("マカロン" + j);
	        } catch (e) {
	            // 一致するレイヤーが見つからない場合
	            continue;
	        }
        }

        // 一致するレイヤーが見つかった場合にのみ処理を行う
        if (layer && dataset.hasOwnProperty(csvColumnName)) {
            try {
                var relativeImagePath = dataset[csvColumnName];
                var absoluteImagePath = convertRelativeToAbsolutePath(relativeImagePath);

                // レイヤーをアクティブにする
                doc.activeLayer = layer;

                replaceSmartObjectContent(layer, absoluteImagePath); // スマートオブジェクトの内容を置き換える

            } catch (e) {
                throw new Error(
                    "function: applyDataset" + "\n\n" +
                    "Layer Name: " + (layer ? layer.name : "unknown") + "\n" +
                    "Error: " + e.message + "\n"
                );
            }
        }
    }
}

// CSVファイルを読み込む関数
function readCSV(file) {
    if (!file.exists) {
        // alert("CSVファイルが存在しません: " + file.fsName);
        return null;
    }
    file.open('r');
    var content = file.read();
    file.close();
    return content;
}

// CSVの内容をパースして配列に変換する関数
function parseSystemCSV(content) {
    var lines = content.split('\n');
    var result = [];
    for (var i = 1; i < lines.length; i++) { // ヘッダー行をスキップ
        var line = lines[i];
        if (line && line !== '') {
            var fields = line.split(',');
            result.push({
                printer: fields[0],
                shape: fields[1],
                type: fields[2],
                maxLayer: parseInt(fields[3], 10)
            });
        }
    }
    return result;
}

// CSVを解析する関数
function parseCSV(content) {
    var lines = content.split("\n"); // ファイルの内容を行ごとに分割する
    var headers = lines[0].split(","); // 最初の行をヘッダーとして取得する
    var datasets = [];

    for (var i = 1; i < lines.length; i++) {
        var line = lines[i].split(",");
        var dataset = {};
        for (var j = 0; j < headers.length; j++) {
            dataset[headers[j]] = line[j] ? line[j] : ""; // ヘッダーをキーとしてデータセットを作成する
        }
        datasets.push(dataset); // データセットをリストに追加する
    }

    return datasets; // 解析したデータセットを返す
}



// 文字列の前後の空白を削除する関数
function trim(str) {
    return str.replace(/^\s+|\s+$/g, '');
}

function main(){
	// CSVファイルを読み込み、データを取得
	var csvContent = readCSV(systemCsvFilePath);

	// CSVの内容が読み込まれた場合のみ処理を続ける
	if (csvContent) {
	    // CSVの内容をパース
	    var data = parseSystemCSV(csvContent);

	    // データをループして変数を設定
	    for (var i = 0; i < data.length; i++) {
	        var item = data[i];

	        // フィールドをトリム
	        item.printer = trim(item.printer);
	        item.shape = trim(item.shape);
	        item.type = trim(item.type);

	        // ファイル名を生成
	        if (item.type) {
	            basePSDPath = new File(File($.fileName).parent + "/" + item.printer + "_" + item.shape + ".psd");
	            csvFilePath = new File(File($.fileName).parent + "/" + item.type + "_Printer" + item.printer + "_" + item.shape + ".csv");
	        } else {
	            basePSDPath = new File(File($.fileName).parent + "/" + item.printer + "_" + item.shape + ".psd");
	            csvFilePath = new File(File($.fileName).parent + "/Printer" + item.printer + "_" + item.shape + ".csv");
	        }

	        // ベースPSDパスの存在を確認
	        if (!basePSDPath.exists) {
	            // alert("ベースPSDファイルが存在しません: " + basePSDPath.fsName);
	            continue; // 次のループに進む
	        }

	        // レイヤー数を設定
	        numberOfLayers = item.maxLayer;

	        // 出力フォルダを設定
	        outputFolder = new Folder(new File($.fileName).parent + "/../../printer" + item.printer);

	        // デバッグ用ログ
	        // alert("basePSDPath: " + basePSDPath + " csvFilePath: " + csvFilePath + " numberOfLayers: " + numberOfLayers + " outputFolder: " + outputFolder);

	        // CSVファイルの存在を確認
	        if (!csvFilePath.exists) {
	            // CSVファイルが存在しない場合は処理をスキップ
	            // alert("CSVファイルが存在しません: " + csvFilePath.fsName);
	            continue;
	        }

	        // ここからPSDを作る



	        app.open(basePSDPath); // 基盤となるPSDファイルを開く

	        if (csvFilePath.exists) {
	            var csvContent = readCSV(csvFilePath); // CSVファイルの内容を読み込む
	            
	            var datasets = parseCSV(csvContent); // CSVファイルを解析する
                if (datasets.length !== 1) {
                    for (var index_dataset = 0; index_dataset < datasets.length; index_dataset++) {
                        
                        if(datasets[index_dataset].filename == ""){
                            continue;
                        }
                        var doc = app.activeDocument.duplicate(); // 現在のドキュメントを複製する
                        
                        try{
                            applyDataset(doc, datasets[index_dataset]); // データセットを適用する
                        } catch (e) {
                            alert(
                                "function: applyDataset" + "\n" +
                                "fsName:" + basePSDPath.fsName + "\n\n" +
                                "Error executing action: " + e.message 
                            );
                            doc.close(SaveOptions.DONOTSAVECHANGES); // 保存せずにドキュメントを閉じる
                            return; 
                        }
    
                        // alert(datasets[index_dataset].filename );
                        
                        var saveFile = new File(File($.fileName).parent + "/output/" + datasets[index_dataset].filename + ".psd"); // 保存先パスに置き換える
                        //alert( datasets[index_dataset].filename);
                
                        if (!outputFolder.exists) outputFolder.create();
                
                        var saveFile = new File(outputFolder + "/" + datasets[index_dataset].filename + ".psd"); // 保存先パスに置き換える
                
                        // 保存オプションを設定
                        var psdOptions = new PhotoshopSaveOptions();
                        psdOptions.layers = true; // レイヤー情報を保持する
                
                        doc.saveAs(saveFile, psdOptions, true, Extension.LOWERCASE); // PSDファイルとして保存する
    
    
    
        
                        doc.close(SaveOptions.DONOTSAVECHANGES); // 保存せずにドキュメントを閉じる
    
                    }
                }
                    
                //実行CSVファイルを消す
                // Fileオブジェクトを作成
                var file = new File(csvFilePath);

                // ファイルが存在するかチェック
                if (file.exists) {
                    // ファイルを削除
                    var success = file.remove();
                }
	        } else {
	            // alert("CSV file not found."); // CSVファイルが見つからない場合にメッセージを表示する
	        }
	    }
	    
	} else {
	    alert("CSVファイルの読み込みに失敗しました。");
	}

	// Close all open documents without saving
	while (app.documents.length > 0) {
	    app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);
	}

	// Create a file to signal completion
	var file = new File(File($.fileName).parent + "/complete.txt");
	file.open("w");
	file.writeln("Completed");
	file.close();

	alert("Processing completed."); // 処理完了のメッセージを表示する

}

// メイン関数を実行
main();