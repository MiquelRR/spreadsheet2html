function NonOpen() {
  Browser.msgBox("Mensaje", "olakase", Browser.Buttons.OK);
  Utilities.sleep(100);
  mostrarPanelLateral();
  Browser.msgBox("Mensaje", "pepe", Browser.Buttons.OK);
}

function mostrarPanelLateral() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HTML");
  var rhtml = hoja.getRange("A2");
  var rcss = hoja.getRange("B2");
  var hh = hoja.getRange("A220");
  SpreadsheetApp.getActive().toast(
    "¡Atenció! la generació li costa una miqueta."
  );
  const r = interpretaFulla();
  const ht = r[0];
  const cs = r[1];
  SpreadsheetApp.getActive().toast(
    "Ya està fet, si li dona la gana, t'ho mostrarà al panell lateral, tranqui que s'ho pren amb calma"
  );
  //var cs = interpretaFulla().css
  rhtml.setValue(toHTML(ht));
  hh.setValue(ht);
  rcss.setValue(cs);
  hoja.activate();

  var html = HtmlService.createHtmlOutputFromFile("panell")
    .setTitle("Panell lateral")
    .setWidth(1000);
  SpreadsheetApp.getUi().showSidebar(html);
  //SpreadsheetApp.getActive().toast("Tienes el archivo completo pinchado el botón.");
}

function generaHtml() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HTML");
  // Obtiene el código HTML y CSS de la hoja de cálculo
  var html = hoja.getRange("A220").getValue();
  var css = hoja.getRange("B2").getValue();

  // Crea el código HTML con el estilo CSS
  var htmlWithStyle =
    "<html>\n  <head>\n   <title>Tabla Autogenerada</title>\n    <style>\n" +
    css +
    "    </style>\n  </head>\n    <body>\n   <table>" +
    html +
    "    </table>\n   </body>\n</html>";

  // Crea la página web
  var file = DriveApp.createFile("tablaAutogenerada.html", htmlWithStyle);
  file.setContent(htmlWithStyle);
  var url = file.getUrl();

  // Muestra la URL de la página web
  Browser.msgBox("Html amb estils ho tens a la URL:\n" + url);
}

function toHTML(texte) {
  let resultat =
    '<!DOCTYPE html>\n<html>\n  <head>\n     <link rel="stylesheet" href="estils.css">\n  </head>\n  <body>\n    <table>\n    ';
  resultat += texte;
  resultat += "    </table>\n  </body>\n</html>\n";
  return resultat;
}
function toCSS(llista) {
  var res = "/* estils.css */\n";
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HTML");
  var adcss = hoja.getRange("D4");
  res += adcss.getValue() + "\n";

  llista.forEach((item) => {
    descarrega(item, llista);
    res += item.eti + (item.clase ? "." + item.clase : "") + "{\n";
    if (item.color.length > 0) {
      res += "     background-color: " + item.color + ";\n";
    }
    if (item.fWeigth.length > 0) {
      res += "     font-weight: " + item.fWeigth.toLowerCase() + ";\n";
    }
    if (item.fStyle.length > 0) {
      res += "     font-style: " + item.fStyle.toLowerCase() + ";\n";
    }
    if (item.hAlign.length > 0) {
      res += "     text-align: " + item.hAlign.toLowerCase() + ";\n";
    }
    if (item.vAlign.length > 0) {
      res += "     vertical-align: " + item.vAlign.toLowerCase() + ";\n";
    }
    if (item.fSize.length > 0) {
      res += "     font-size: " + item.fSize + "rem;\n";
    }
    res += "}\n";
  });
  return res;
}

function buida(f, c) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ORIGEN");
  var rango = hoja.getRange(f, 2, 1, c - 1);
  var valores = rango.getValues();
  var estaVacia = valores.every(function (fila) {
    return fila.every(function (celda) {
      return celda === "" || celda === null;
    });
  });
  return estaVacia;
}

function totaNegreta(f, c) {
  //retorna true si la fila f esta tota en negreta fins la columna c
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ORIGEN");
  var rango = hoja.getRange(f, 1, 1, c);
  var estils = rango.getFontWeights();
  var resulta = estils[0].every(function (estils) {
    return estils === "bold";
  });
  return resulta;
}
function aplanaNomCSS(nombre) {
  if (nombre.length === 0) nombre = "sense_texte";
  // Utilizar una expresión regular para filtrar caracteres no admitidos
  return nombre.replace(/[^a-zA-Z0-9-_]/g, "x");
}

function getCSSBorder(range) {
  //he desistit de importar vores
  // Obtener los bordes de la celda
  var borders = range.getBorder();
  var sal = "";
  // Obtener los valores de los bordes
  if (borders !== null) {
    var values = [];
    values.push(borders.getTop());
    values.push(borders.getBottom());
    values.push(borders.getLeft());
    values.push(borders.getRight());
    // Convertir los valores de los bordes a cadenas de CSS
    values.forEach((value) => {
      var style = value.getBorderStyle();
      Logger.log(style);
      if (style == "SOLID") sal += "1px solid " + value.getColor() + " ";
      if (style == "DASHED") sal += "1px dashed " + value.getColor() + " ";
      if (style == "DOTTED") sal += "1px dotted " + value.getColor() + " ";
      if (style == "DOUBLE") sal += "2px solid " + value.getColor() + " ";
      if (style == "HAIRLINE") sal += "1px hairline " + value.getColor() + " ";
      if (style == "INSET") sal += "1px inset " + value.getColor() + " ";
      if (style == "OUTSET") sal += "1px outset " + value.getColor() + " ";
      if (style == "SHADOW") sal += "1px shadow " + value.getColor() + " ";
    });
  }
  return sal;
}
function descarrega(element, cssdata) {
  // Elimina les propietats que ja estan en la etiqueta principal
  //Logger.log("kase");
  if (element.clase.length > 0) {
    const etibase = cssdata.find(
      (item) => item.eti === element.eti && item.clase == ""
    );

    Logger.log(etibase.eti.length);

    if (etibase) {
      Logger.log(etibase.clase.length);
      //Logger.log("ola");
      for (const prop in element) {
        if (prop !== "clase" && prop !== "eti") {
          if (etibase[prop] === element[prop]) {
            element[prop] = "";
          }
        }
      }
    }
  }
}

function agregaCss(cssdata, f, c, etiqueta) {
  const tamanyLletra = 10;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ORIGEN");
  var range = sheet.getRange(f, c);
  var nm = range.getValue();
  var nom = nm.split(" ")[0];
  nom = aplanaNomCSS(nom);
  if (/^\d/.test(nom)) {
    // Si es numèric, afig 'N' per davant
    nom = "N" + nom;
  }

  var objecte = {
    eti: etiqueta,
    color: (color = range.getBackground()),
    fWeigth: range.getFontWeight(),
    fStyle: range.getFontStyle(),
    hAlign: range.getHorizontalAlignment(),
    vAlign: range.getVerticalAlignment(),
    fSize: range.getFontSize() / tamanyLletra,
    //borde  : getCSSBorder(range)
  };

  var igual = cssdata.find(function (item) {
    //busca si hi ha enregistrat ja un identic
    return Object.keys(objecte).every(function (key) {
      return item[key] === objecte[key];
    });
  });

  if (
    cssdata.some(function (item) {
      return item.eti === etiqueta;
    })
  ) {
    // si l'etiqueta existeix
    if (igual) {
      //Logger.log(igual);
      nom = igual.clase; // Si trobem un conjunt de caracteristiques iguals li asignem el mateix nom
    } else {
      var coincidencia = cssdata.some(function (objeto) {
        // comproba si ja tenim una parella etiqueta-clase igual
        return objeto.eti === etiqueta && objeto.clase === nom;
      });
      nom += coincidencia ? "-" : "";
      objecte.clase = nom; //afig - al nom si hi ha coincidència
      //descarrega(objecte,cssdata);
      cssdata.push(objecte);
    }
  } else {
    //no existeix l'etiqueta  MODIFICAT
    nom = ""; // AFEGIT
    objecte.clase = nom;
    cssdata.push(objecte);
  }
  var tag = etiqueta;
  tag += nom.length > 0 ? ' class="' + nom + '"' : "";
  //Logger.log('f'+f+'c'+c+'-'+tag);
  return { cls: tag };
}

function interpretaFulla() {
  var tg;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ORIGEN");
  var rangDades = sheet.getDataRange();
  var y2 = rangDades.getNumRows();
  var x2 = rangDades.getNumColumns();
  var y1 = 1,
    x1 = 1;
  var range = sheet.getRange(x1, y1);
  var isBold = range.getFontWeight() == "bold";
  var value = range.getValue();
  let resultat = "";
  var cssdata = []; //llista
  if (buida(1, x2)) {
    //Títol
    var tg = agregaCss(cssdata, 1, 1, "caption").cls;
    resultat +=
      " <" +
      tg +
      ">" +
      (isBold ? "<strong>" : "") +
      value +
      (isBold ? "</strong>" : "") +
      "</caption>\n";
    y1 = 2;
  }
  var posahead = totaNegreta(y1, x2);
  var posafoot = totaNegreta(y2, x2);
  for (var i = y1; i <= y2; i++) {
    if (posahead && i === y1) {
      tg = agregaCss(cssdata, y1, 1, "thead").cls;
      resultat += "  <" + tg + ">\n";
    }
    if (posafoot && i === y2) {
      tg = agregaCss(cssdata, y1, 1, "tfoot").cls;
      resultat += "  <" + tg + ">\n";
    }
    resultat += "  <tr>\n";
    for (var j = x1; j <= x2; j++) {
      range = sheet.getRange(i, j);
      var rp = 0;
      var f = i + 1;
      isBold = range.getFontWeight() == "bold";
      value = range.getValue();
      var thd = isBold ? "th" : "td";
      if (!agrupada(i, j).esta) {
        //si no esta agrupada
        tg = agregaCss(cssdata, i, j, thd).cls;
        resultat += "       <" + tg + ">" + value + "</" + thd + ">\n";
      } else {
        if (agrupada(i, j).inicial) {
          tg = agregaCss(cssdata, i, j, thd).cls;
          resultat +=
            "       <" +
            tg +
            agrupada(i, j).cadena +
            ">" +
            value +
            "</" +
            thd +
            ">\n";
        }
        j += agrupada(i, j).avanza;
      }
    }
    resultat += "  </tr>\n";
    resultat += posahead && i === y1 ? "  </thead>\n" : "";
    resultat += posafoot && i === y2 ? "  </tfoot>\n" : "";
  }
  return [resultat, toCSS(cssdata)];
}

function agrupada(f, c) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ORIGEN");
  var r = sheet.getRange(f, c);
  var mergedRange = r.getMergedRanges();
  var esta = mergedRange.length == 1; //true si es una celda soldada con otras
  var inicial = false,
    hor = false,
    ver = false;
  var ultF = 0,
    ultC = 0;
  if (esta) {
    inicial = mergedRange[0].getRow() == f && mergedRange[0].getColumn() == c;
    //Logger.log(inicial+" "+f+"vs"+mergedRange[0].getRow()+"-"+c+"vs"+mergedRange[0].getColumn())
    ultC = mergedRange[0].getLastColumn();
    ver = mergedRange[0].getHeight() > 1;
    hor = mergedRange[0].getWidth() > 1;
  }
  var cadena = "";
  if (inicial) {
    cadena = hor ? ' colspan="' + mergedRange[0].getWidth() + '"' : "";
    cadena += ver ? ' rowspan="' + mergedRange[0].getHeight() + '"' : "";
  }
  var avanza = ultC - c;
  return { esta: esta, inicial: inicial, cadena: cadena, avanza: avanza };
}
