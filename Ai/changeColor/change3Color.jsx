function changeColorFromObjectNames() {
    // ドキュメントを取得
    var doc = app.activeDocument;

    // オブジェクト "targetColor1", "targetColor2", "targetColor3" と "newColor1", "newColor2", "newColor3" を取得
    var targetColors = [
        findObjectByName("targetColor1"),
        findObjectByName("targetColor2"),
        findObjectByName("targetColor3")
    ];

    var newColors = [
        findObjectByName("newColor1"),
        findObjectByName("newColor2"),
        findObjectByName("newColor3")
    ];

    var colorMap = {};

    for (var i = 0; i < targetColors.length; i++) {
        if (targetColors[i] && newColors[i]) {
            var targetColor = targetColors[i].fillColor;
            var newColor = newColors[i].fillColor;

            if (targetColor && newColor) {
                var targetHex = rgbToHex(targetColor);
                var newHex = rgbToHex(newColor);
                colorMap[targetHex] = hexToRGB(newHex);
            }
        }
    }

    // すべてのオブジェクトを取得
    var items = doc.pageItems;

    // すべてのオブジェクトをループ
    for (var j = 0; j < items.length; j++) {
        var item = items[j];

        // "targetColor" や "newColor" オブジェクトは変更しない
        var isTargetOrNewColor = false;
        for (var k = 0; k < targetColors.length; k++) {
            if (item === targetColors[k] || item === newColors[k]) {
                isTargetOrNewColor = true;
                break;
            }
        }

        if (isTargetOrNewColor) {
            continue;  // スキップして次のオブジェクトへ
        }

        // 通常の塗りと線の色を変更
        for (var targetHex in colorMap) {
            if (item.filled && compareColors(item.fillColor, hexToRGB(targetHex))) {
                item.fillColor = colorMap[targetHex];
            }

            if (item.stroked && compareColors(item.strokeColor, hexToRGB(targetHex))) {
                item.strokeColor = colorMap[targetHex];
            }
        }

        // アピアランスの項目に対して色を変更
        changePathAppearanceColors(item, colorMap);

        // テキストオブジェクトの場合、テキストの塗りと線の色を変更
        if (item.typename === "TextFrame") {
            changeTextColors(item, colorMap);
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
function changeTextColors(textFrame, colorMap) {
    var characters = textFrame.textRange.characters;
    for (var i = 0; i < characters.length; i++) {
        var character = characters[i];
        for (var targetHex in colorMap) {
            if (compareColors(character.fillColor, hexToRGB(targetHex))) {
                character.fillColor = colorMap[targetHex];
            }
            if (compareColors(character.strokeColor, hexToRGB(targetHex))) {
                character.strokeColor = colorMap[targetHex];
            }
        }
    }
}

// アピアランスの塗りや線の色を変更する関数
function changePathAppearanceColors(item, colorMap) {
    // アピアランスの効果を確認し、色を変更
    try {
        var appearanceItems = item.graphicStyles[0].attributes; // アピアランスの項目を取得
        for (var j = 0; j < appearanceItems.length; j++) {
            var appearanceItem = appearanceItems[j];
            for (var targetHex in colorMap) {
                if (appearanceItem.filled && compareColors(appearanceItem.fillColor, hexToRGB(targetHex))) {
                    appearanceItem.fillColor = colorMap[targetHex];
                }
                if (appearanceItem.stroked && compareColors(appearanceItem.strokeColor, hexToRGB(targetHex))) {
                    appearanceItem.strokeColor = colorMap[targetHex];
                }
            }
        }
    } catch (e) {
        // アピアランスの取得が失敗してもスクリプトを終了しない
        $.writeln("アピアランスの取得に失敗しました: " + e.message);
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
