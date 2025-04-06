 var colorCodeMap = {
    "255,0,0": "7",     // Red
    "255,255,0": "4",   // Yellow
    "0,255,255": "2",   // Cyan
    "128,0,128": "6",   // Purple
    "255,255,255": "1", // White
    "0,0,0": "0"        // Black
};

var doc = app.activeDocument;
var texts = doc.textFrames;

 function getRGBString(color) {
    if (color.typename === "RGBColor") {
        return color.red + "," + color.green + "," + color.blue;
    }
    return "";
}

for (var i = 0; i < texts.length; i++) {
    var text = texts[i];
    var bounds = text.geometricBounds; // [y1, x1, y2, x2]
    var centerX = (bounds[1] + bounds[3]) / 2;
    var centerY = (bounds[0] + bounds[2]) / 2;

for (var j = 0; j < doc.pageItems.length; j++) {
        var item = doc.pageItems[j];
         if (item == text || !item.filled) continue;
      
var itemBounds = item.geometricBounds;
        if (
            centerX >= itemBounds[1] &&
            centerX <= itemBounds[3] &&
            centerY <= itemBounds[0] &&
            centerY >= itemBounds[2]
        ) {
            var rgbString = getRGBString(item.fillColor);
            if (rgbString in colorCodeMap) {
                text.contents = colorCodeMap[rgbString];
            }
            break;
        }
    }
}

alert("Color codes applied to text!");
