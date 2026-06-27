console.log("Antigravity Power-Up Connector Cargado");

// Configuración global
var APP_KEY = (window.location.hostname.indexOf('github.io') > -1 || window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1')
  ? '30767998cbb7b57a692e8bb50cfa9a9c' // Power-Up Prepro (GitHub Pages / Local)
  : '352ae509bc7ffb5b7f5826a024353794'; // Power-Up Pro (Netlify / Custom Domain)
var APP_NAME = 'Antigravity';
var BASE_URL = window.location.href.substring(0, window.location.href.lastIndexOf('/') + 1);

// Función de utilidad inyectada para detectar el tipo de tablero de forma segura
var getBoardTypeSafe = function (t) {
  return t.lists('id', 'name').then(function (lists) {
    var paramList = lists.find(function (l) {
      var n = l.name.toUpperCase();
      return n === 'PARÁMETROS' || n === 'PARAMETROS';
    });
    if (!paramList) return 'CLIENTES';

    return t.cards('id', 'name', 'idList').then(function (openCards) {
      var typeCard = openCards.find(function (c) {
        return c.idList === paramList.id && c.name && c.name.toLowerCase().startsWith('tipo tablero');
      });
      if (!typeCard) return 'CLIENTES';

      var name = typeCard.name.toUpperCase();
      if (name.includes('PLANNING')) return 'PLANNING';
      if (name.includes('FACTURACIÓN') || name.includes('FACTURACION')) return 'FACTURACIÓN';
      return 'CLIENTES';
    });
  }).catch(function (err) {
    console.error("Error detectando tablero:", err);
    return 'CLIENTES';
  });
};

window.TrelloPowerUp.initialize({
  // Capability para añadir botones en el menú derecho de la tarjeta
  'card-buttons': function (t, options) {
    return getBoardTypeSafe(t).then(function (type) {
      if (type !== 'CLIENTES') return [];

      return [{
        // Icono que se mostrará junto al texto del botón
        icon: './icon.svg',
        // Texto que se mostrará en el botón
        text: 'Visualizar datos de interés',
        // Función que se ejecuta cuando el usuario hace clic en el botón
        callback: function (t) {
          return t.card('desc')
            .then(function (card) {
              var fields = {
                nombre: "No encontrado",
                dni: "No encontrado",
                telefono: "No encontrado"
              };

              if (card.desc) {
                var m;
                if (m = card.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*NOMBRE Y APELLIDOS[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.nombre = m[1].replace(/[*_~`]/g, '').trim(); }
                if (m = card.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*DNI[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.dni = m[1].replace(/[*_~`]/g, '').trim(); }
                if (m = card.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*(?:TELÉFONO|Tel[eé]fono|Tlf|Tel)[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.telefono = m[1].replace(/[*_~`]/g, '').trim(); }
              }

              return t.alert({
                message: 'El cliente se llama ' + fields.nombre + ', con DNI ' + fields.dni + ' y teléfono ' + fields.telefono,
                duration: 10,
                display: 'info'
              });
            });
        }
      }];
    });
  },
  // Capability para añadir una sección grande dentro del cuerpo de la tarjeta
  'card-back-section': function (t, options) {
    return {
      title: 'Powerup opciones 22',
      icon: './icon.svg', // Recomendado un icono gris, pero el SVG actual servirá
      content: {
        type: 'iframe',
        url: t.signUrl('./section.html?v=65'), // Apuntamos a un nuevo archivo HTML que vamos a crear
        height: 250 // Altura inicial en píxeles (aumentado para evitar scrollbar)
      }
    };
  },
  // Capability para añadir un botón a nivel de tablero
  'board-buttons': function (t, options) {
    return getBoardTypeSafe(t).then(function (type) {
      if (type !== 'CLIENTES' && type !== 'FACTURACIÓN') return [];

      var buttons = [];

      // Botón "Excel Tablero" común
      buttons.push({
        icon: {
          dark: BASE_URL + 'excel_icon.png',
          light: BASE_URL + 'excel_icon.png'
        },
        text: 'Excel Tablero',
        condition: 'always',
        callback: function (t) {
          t.alert({
            message: 'Recopilando tarjetas e iniciando exportación...',
            duration: 3,
            display: 'info'
          });

          return Promise.all([
            t.cards('id', 'name', 'desc', 'idList'),
            t.lists('id', 'name')
          ])
            .then(function (results) {
              var cards = results[0];
              var lists = results[1];
              var mapListas = {};
              lists.forEach(function (l) { mapListas[l.id] = l.name; });

              return new Promise((resolve, reject) => {
                if (window.ExcelJS) return resolve({ cards: cards, mapListas: mapListas });
                var script = document.createElement('script');
                script.src = "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js";
                script.onload = () => resolve({ cards: cards, mapListas: mapListas });
                script.onerror = reject;
                document.head.appendChild(script);
              });
            })
            .then(function (data) {
              var cards = data.cards;
              var mapListas = data.mapListas;
              var workbook = new ExcelJS.Workbook();
              var sheet = workbook.addWorksheet('Clientes');

              sheet.columns = [
                { header: 'Lista', key: 'lista', width: 20 },
                { header: 'Tarjeta Trello', key: 'tarjeta', width: 30 },
                { header: 'Nombre', key: 'nombre', width: 30 },
                { header: 'DNI', key: 'dni', width: 15 },
                { header: 'Lugar', key: 'lugar', width: 20 },
                { header: 'Ciudad', key: 'ciudad', width: 20 },
                { header: 'Matrícula 1', key: 'matricula1', width: 15 },
                { header: 'Aseguradora 1', key: 'aseguradora1', width: 20 },
                { header: 'Matrícula 2', key: 'matricula2', width: 15 },
                { header: 'Aseguradora 2', key: 'aseguradora2', width: 20 },
                { header: 'Fecha del siniestro', key: 'fecha', width: 20 },
                { header: 'Hora', key: 'hora', width: 10 },
                { header: 'Observaciones', key: 'observaciones', width: 40 }
              ];

              cards.forEach(function (c) {
                var fields = {
                  nombre: "No encontrado",
                  dni: "No encontrado",
                  domicilio: "No encontrado",
                  telefono: "No encontrado",
                  profesion: "No encontrado",
                  fecha: "No encontrado",
                  hora: "No encontrado",
                  lugar: "No encontrado",
                  ciudad: "No encontrado",
                  matricula1: "No encontrado",
                  aseguradora1: "No encontrado",
                  matricula2: "No encontrado",
                  aseguradora2: "No encontrado",
                  observaciones: "No encontrado"
                };

                if (c.desc) {
                  var m;
                  if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*NOMBRE Y APELLIDOS[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.nombre = m[1].replace(/[*_~`]/g, '').trim(); }
                  if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*DNI[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.dni = m[1].replace(/[*_~`]/g, '').trim(); }
                  if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*DOMICILIO[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.domicilio = m[1].replace(/[*_~`]/g, '').trim(); }
                  if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*(?:TELÉFONO|Tel[eé]fono|Tlf|Tel)[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.telefono = m[1].replace(/[*_~`]/g, '').trim(); }
                  if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*PROFESI[OÓ]N[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.profesion = m[1].replace(/[*_~`]/g, '').trim(); }
                  if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*OBSERVACIONES[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.observaciones = m[1].replace(/[*_~`]/g, '').trim(); }
                  if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*(?:FECHA DEL SINIESTRO|Fecha)[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.fecha = m[1].replace(/[*_~`]/g, '').replace(/\s*\(.*\)\s*$/, '').trim(); }
                  if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*(?:HORA)[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.hora = m[1].replace(/[*_~`]/g, '').trim(); }
                  if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*(?:LUGAR)[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.lugar = m[1].replace(/[*_~`]/g, '').trim(); }
                  if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*(?:CIUDAD)[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.ciudad = m[1].replace(/[*_~`]/g, '').trim(); }

                  var vLA = "(?=\\n\\s*(?:VEH[ÍI]CULO CONTRARIO|VEH[ÍI]CULO PROPIO|PARRAFADA|NOMBRE|DNI|DOMICILIO|PROFESI[OÓ]N|TEL[EÉ]FONO|EMAIL|OBSERVACIONES)|$)";

                  var propioMatch = c.desc.match(new RegExp("(?:^|\\n)[ \\t]*[*_~`#]*\s*VEH[ÍI]CULO PROPIO[\\s\\S]*?" + vLA, "i"));
                  if (propioMatch) {
                    var pT = propioMatch[0];
                    if (m = pT.match(/(?:^|\n)[ \t]*MATR[ÍI]CULA[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.matricula1 = m[1].trim(); }
                    if (m = pT.match(/(?:^|\n)[ \t]*ASEGURADORA[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.aseguradora1 = m[1].trim(); }
                  }
                  var contrarioMatch = c.desc.match(new RegExp("(?:^|\\n)[ \\t]*[*_~`#]*\s*VEH[ÍI]CULO CONTRARIO[\\s\\S]*?" + vLA, "i"));
                  if (contrarioMatch) {
                    var cT = contrarioMatch[0];
                    if (m = cT.match(/(?:^|\n)[ \t]*MATR[ÍI]CULA[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.matricula2 = m[1].trim(); }
                    if (m = cT.match(/(?:^|\n)[ \t]*ASEGURADORA[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.aseguradora2 = m[1].trim(); }
                  }
                }

                sheet.addRow({
                  lista: mapListas[c.idList] || "Desconocida",
                  tarjeta: c.name,
                  nombre: fields.nombre,
                  dni: fields.dni,
                  domicilio: fields.domicilio,
                  telefono: fields.telefono,
                  profesion: fields.profesion,
                  fecha: fields.fecha,
                  hora: fields.hora,
                  lugar: fields.lugar,
                  ciudad: fields.ciudad,
                  matricula1: fields.matricula1,
                  aseguradora1: fields.aseguradora1,
                  matricula2: fields.matricula2,
                  aseguradora2: fields.aseguradora2,
                  observaciones: fields.observaciones
                });
              });

              return workbook.xlsx.writeBuffer();
            })
            .then(function (buffer) {
              return new Promise((resolve, reject) => {
                if (window.saveAs) return resolve(buffer);
                var scriptFS = document.createElement('script');
                scriptFS.src = "https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js";
                scriptFS.onload = () => resolve(buffer);
                scriptFS.onerror = reject;
                document.head.appendChild(scriptFS);
              });
            })
            .then(function (buffer) {
              var blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
              window.saveAs(blob, "Reporte_Tablero.xlsx");
              return t.alert({ message: 'Descarga completada.', duration: 3, display: 'success' });
            })
            .catch(function (err) {
              console.error("Error exportando tablero:", err);
              return t.alert({ message: 'Error al exportar.', duration: 6, display: 'error' });
            });
        }
      });

      // Botón "Descargar clientes" sólo para tableros de CLIENTES
      if (type === 'CLIENTES') {
        buttons.push({
          icon: {
            dark: BASE_URL + 'excel_icon.png',
            light: BASE_URL + 'excel_icon.png'
          },
          text: 'Descargar clientes',
          condition: 'always',
          callback: function (t) {
            t.alert({
              message: 'Generando reporte de clientes...',
              duration: 3,
              display: 'info'
            });

            return Promise.all([
              t.cards('desc', 'idList'),
              t.lists('id', 'name')
            ])
              .then(function (results) {
                var cards = results[0];
                var lists = results[1];

                return new Promise((resolve, reject) => {
                  if (window.ExcelJS) return resolve({ cards: cards, lists: lists });
                  var script = document.createElement('script');
                  script.src = "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js";
                  script.onload = () => resolve({ cards: cards, lists: lists });
                  script.onerror = reject;
                  document.head.appendChild(script);
                });
              })
              .then(function (data) {
                var cards = data.cards;
                var lists = data.lists;
                var workbook = new ExcelJS.Workbook();
                var sheet = workbook.addWorksheet('Clientes');

                var infoClientePorLista = {};

                lists.forEach(function (l) {
                  infoClientePorLista[l.id] = { nombre: '', dni: '', domicilio: '', telefono: '', profesion: '', observaciones: '' };
                });

                cards.forEach(function (c) {
                  if (infoClientePorLista[c.idList]) {
                    if (infoClientePorLista[c.idList].nombre === '' && c.desc) {
                      var m;
                      if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*NOMBRE Y APELLIDOS[^:]*[:\s-]+\s*([^\n\r]*)/i)) { infoClientePorLista[c.idList].nombre = m[1].replace(/[*_~`]/g, '').trim(); }
                      if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*DNI[^:]*[:\s-]+\s*([^\n\r]*)/i)) { infoClientePorLista[c.idList].dni = m[1].replace(/[*_~`]/g, '').trim(); }
                      if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*DOMICILIO[^:]*[:\s-]+\s*([^\n\r]*)/i)) { infoClientePorLista[c.idList].domicilio = m[1].replace(/[*_~`]/g, '').trim(); }
                      if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*(?:TELÉFONO|Tel[eé]fono|Tlf|Tel)[^:]*[:\s-]+\s*([^\n\r]*)/i)) { infoClientePorLista[c.idList].telefono = m[1].replace(/[*_~`]/g, '').trim(); }
                      if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*PROFESI[OÓ]N[^:]*[:\s-]+\s*([^\n\r]*)/i)) { infoClientePorLista[c.idList].profesion = m[1].replace(/[*_~`]/g, '').trim(); }
                      if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*OBSERVACIONES[^:]*[:\s-]+\s*([^\n\r]*)/i)) { infoClientePorLista[c.idList].observaciones = m[1].replace(/[*_~`]/g, '').trim(); }
                    }
                  }
                });

                var encabezado = ['Nombre lista', 'Nombre', 'DNI', 'Domicilio', 'Teléfono', 'Profesión', 'Observaciones'];
                var headerRow = sheet.addRow(encabezado);
                headerRow.font = { bold: true };

                lists.forEach(function (l) {
                  var info = infoClientePorLista[l.id];
                  var fila = [l.name, info.nombre, info.dni, info.domicilio, info.telefono, info.profesion, info.observaciones];
                  sheet.addRow(fila);
                });

                sheet.getColumn(1).width = 25;
                sheet.getColumn(2).width = 30;
                sheet.getColumn(3).width = 15;
                sheet.getColumn(4).width = 25;
                sheet.getColumn(5).width = 15;
                sheet.getColumn(6).width = 20;
                sheet.getColumn(7).width = 40;

                return workbook.xlsx.writeBuffer();
              })
              .then(function (buffer) {
                return new Promise((resolve, reject) => {
                  if (window.saveAs) return resolve(buffer);
                  var scriptFS = document.createElement('script');
                  scriptFS.src = "https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js";
                  scriptFS.onload = () => resolve(buffer);
                  scriptFS.onerror = reject;
                  document.head.appendChild(scriptFS);
                });
              })
              .then(function (buffer) {
                var blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
                window.saveAs(blob, "Clientes.xlsx");
                return t.alert({ message: 'Descarga de clientes completada.', duration: 3, display: 'success' });
              })
              .catch(function (err) {
                console.error("Error exportando clientes:", err);
                return t.alert({ message: 'Error al exportar clientes.', duration: 6, display: 'error' });
              });
          }
        });
      }


      if (type === 'FACTURACIÓN') {
        buttons.push({
          icon: {
            dark: BASE_URL + 'facturas_icon.png',
            light: BASE_URL + 'facturas_icon.png'
          },
          text: 'Descargar Facturas',
          condition: 'always',
          callback: function (t) {
            return t.popup({
              title: 'Descargar Facturas',
              url: t.signUrl('./filtro_fecha.html'),
              height: 380
            });
          }
        });
      }

      return buttons;
    });
  }
}, {
  appKey: APP_KEY,
  appName: APP_NAME
});
