// 複数の要素に特定の色があった場合、一括で別の色にするスクリプト

function changeColorFromObjectNames() {
    // ドキュメントを取得
    var doc = app.activeDocument;

    // オブジェクト "targetColor" と "newColor" を取得
    var targetObject = findObjectByName("targetColor");
    var newObject = findObjectByName("newColor");

    if (!targetObject || !newObject) {
        alert("オブジェクト 'targetColor' または 'newColor' が見つかりません。");
        return;
    }

    // オブジェクトから塗りの色を取得
    var targetColor = targetObject.fillColor;
    var newColor = newObject.fillColor;

    if (!targetColor || !newColor) {
        alert("オブジェクト 'targetColor' または 'newColor' に有効な塗りの色がありません。");
        return;
    }

    // 塗りの色をカラーコードに変換
    var targetHex = rgbToHex(targetColor);
    var newHex = rgbToHex(newColor);

    // 新しい色をカラーコードからRGBColorに変換
    var newRGBColor = hexToRGB(newHex);

    // すべてのオブジェクトを取得
    var items = doc.pageItems;

    // すべてのオブジェクトをループ
    for (var i = 0; i < items.length; i++) {
        var item = items[i];

        // "targetColor" や "newColor" オブジェクトは変更しない
        if (item === targetObject || item === newObject) {
            continue;  // スキップして次のオブジェクトへ
        }

        // 塗りの色を変更
        if (item.filled && compareColors(item.fillColor, hexToRGB(targetHex))) {
            item.fillColor = newRGBColor;
        }

        // 線の色を変更
        if (item.stroked && compareColors(item.strokeColor, hexToRGB(targetHex))) {
            item.strokeColor = newRGBColor;
        }

        // テキストオブジェクトの場合、テキストの塗りと線の色を変更
        if (item.typename === "TextFrame") {
            changeTextColors(item, targetHex, newRGBColor);
        }
    }
}

// オブジェクト名で検索して該当するオブジェクトを返す関数
function findObjectByName(name) {
    var items = app.activeDocument.pageItems;
    for (var i = 0; i < items.length; i++) {
        if (items[i].name === name) {
            return items[i];
        }
    }
    return null;
}

// テキストフレーム内のテキストの色を変更する関数
function changeTextColors(textFrame, targetHex, newRGBColor) {
    var characters = textFrame.textRange.characters;
    for (var i = 0; i < characters.length; i++) {
        var character = characters[i];
        if (compareColors(character.fillColor, hexToRGB(targetHex))) {
            character.fillColor = newRGBColor;
        }
        if (compareColors(character.strokeColor, hexToRGB(targetHex))) {
            character.strokeColor = newRGBColor;
        }
    }
}

// RGBColorオブジェクトをカラーコードに変換する関数
function rgbToHex(color) {
    function componentToHex(c) {
        var hex = Math.round(c).toString(16);
        return hex.length == 1 ? "0" + hex : hex;
    }
    return "#" + componentToHex(color.red) + componentToHex(color.green) + componentToHex(color.blue);
}

// カラーコードをRGBColorに変換する関数
function hexToRGB(hex) {
    var rgbColor = new RGBColor();
    hex = hex.replace("#", "");

    if (hex.length === 6) {
        rgbColor.red = parseInt(hex.substring(0, 2), 16);
        rgbColor.green = parseInt(hex.substring(2, 4), 16);
        rgbColor.blue = parseInt(hex.substring(4, 6), 16);
    }
    return rgbColor;
}

// RGBカラーオブジェクトを比較する関数
function compareColors(color1, color2) {
    return (color1.red == color2.red && color1.green == color2.green && color1.blue == color2.blue);
}

// スクリプトを実行
changeColorFromObjectNames();
