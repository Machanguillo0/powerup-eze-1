console.log("Antigravity Power-Up Connector Cargado");

// Configuración global
var APP_KEY = 'dcd7ea65b37ed1be20a5f1dca07f3382';
var APP_NAME = 'Antigravity';

window.TrelloPowerUp.initialize({

  // Capability para añadir botones en el menú derecho de la tarjeta
  'card-buttons': function (t, options) {
    return [{
      // Icono que se mostrará junto al texto del botón
      icon: './icon.svg',
      // Texto que se mostrará en el botón
      text: 'leer datos descripción',
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
              if (m = card.desc.match(/NOMBRE Y APELLIDOS[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.nombre = m[1].replace(/[*_~`]/g, '').trim(); }
              if (m = card.desc.match(/DNI[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.dni = m[1].replace(/[*_~`]/g, '').trim(); }
              if (m = card.desc.match(/(?:TELÉFONO|Tel[eé]fono|Tlf|Tel)[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.telefono = m[1].replace(/[*_~`]/g, '').trim(); }
            }

            return t.alert({
              message: 'El cliente se llama ' + fields.nombre + ', con DNI ' + fields.dni + ' y teléfono ' + fields.telefono,
              duration: 10,
              display: 'info'
            });
          });
      }
    }];
  },
  // Capability para añadir una sección grande dentro del cuerpo de la tarjeta
  'card-back-section': function (t, options) {
    return {
      title: 'Menú de acciones',
      icon: './icon.svg', // Recomendado un icono gris, pero el SVG actual servirá
      content: {
        type: 'iframe',
        url: t.signUrl('./section.html'), // Apuntamos a un nuevo archivo HTML que vamos a crear
        height: 250 // Altura inicial en píxeles (aumentado para evitar scrollbar)
      }
    };
  },
  // Capability para añadir un botón a nivel de tablero
  'board-buttons': function (t, options) {
    return [
      {
        icon: {
          dark: './export_tablero.png',
          light: './export_tablero.png'
        },
        text: 'Exportar Tablero a Excel',
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
                { header: 'Domicilio', key: 'domicilio', width: 30 },
                { header: 'Teléfono', key: 'telefono', width: 15 },
                { header: 'Profesión', key: 'profesion', width: 20 },
                { header: 'Observaciones', key: 'observaciones', width: 40 }
              ];

              cards.forEach(function (c) {
                var fields = {
                  nombre: "No encontrado",
                  dni: "No encontrado",
                  domicilio: "No encontrado",
                  telefono: "No encontrado",
                  profesion: "No encontrado",
                  observaciones: "No encontrado"
                };

                if (c.desc) {
                  var m;
                  if (m = c.desc.match(/NOMBRE Y APELLIDOS[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.nombre = m[1].replace(/[*_~`]/g, '').trim(); }
                  if (m = c.desc.match(/DNI[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.dni = m[1].replace(/[*_~`]/g, '').trim(); }
                  if (m = c.desc.match(/DOMICILIO[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.domicilio = m[1].replace(/[*_~`]/g, '').trim(); }
                  if (m = c.desc.match(/(?:TELÉFONO|Tel[eé]fono|Tlf|Tel)[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.telefono = m[1].replace(/[*_~`]/g, '').trim(); }
                  if (m = c.desc.match(/PROFESI[OÓ]N[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.profesion = m[1].replace(/[*_~`]/g, '').trim(); }
                  if (m = c.desc.match(/OBSERVACIONES[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.observaciones = m[1].replace(/[*_~`]/g, '').trim(); }
                }

                sheet.addRow({
                  lista: mapListas[c.idList] || "Desconocida",
                  tarjeta: c.name,
                  nombre: fields.nombre,
                  dni: fields.dni,
                  domicilio: fields.domicilio,
                  telefono: fields.telefono,
                  profesion: fields.profesion,
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
      },
      {
        icon: {
          dark: './export_listas.png',
          light: './export_listas.png'
        },
        text: 'Exportar Listas a Excel',
        condition: 'always',
        callback: function (t) {
          t.alert({
            message: 'Generando reporte por listas...',
            duration: 3,
            display: 'info'
          });

          return Promise.all([
            t.cards('name', 'desc', 'idList', 'labels'),
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
              var sheet = workbook.addWorksheet('Reporte por Listas');

              // Agrupamos tarjetas por lista, extraemos info de cliente y etiquetas
              var tarjetasPorLista = {};
              var infoClientePorLista = {};
              var etiquetasPorLista = {};
              var maxCards = 0;
              var maxLabels = 0;

              lists.forEach(function (l) {
                tarjetasPorLista[l.id] = [];
                infoClientePorLista[l.id] = { nombre: '', dni: '', domicilio: '', telefono: '', profesion: '', observaciones: '' };
                etiquetasPorLista[l.id] = new Set();
              });

              cards.forEach(function (c) {
                if (tarjetasPorLista[c.idList]) {
                  tarjetasPorLista[c.idList].push(c.name);

                  // Recopilamos todas las etiquetas de la lista
                  if (c.labels) {
                    c.labels.forEach(function (label) {
                      etiquetasPorLista[c.idList].add(label.name || label.color);
                    });
                  }

                  // Intentamos extraer info de cliente si aún no la tenemos para esta lista
                  if (infoClientePorLista[c.idList].nombre === '' && c.desc) {
                    var m;
                    if (m = c.desc.match(/NOMBRE Y APELLIDOS[^:]*[:\s-]+\s*([^\n\r]*)/i)) { infoClientePorLista[c.idList].nombre = m[1].replace(/[*_~`]/g, '').trim(); }
                    if (m = c.desc.match(/DNI[^:]*[:\s-]+\s*([^\n\r]*)/i)) { infoClientePorLista[c.idList].dni = m[1].replace(/[*_~`]/g, '').trim(); }
                    if (m = c.desc.match(/DOMICILIO[^:]*[:\s-]+\s*([^\n\r]*)/i)) { infoClientePorLista[c.idList].domicilio = m[1].replace(/[*_~`]/g, '').trim(); }
                    if (m = c.desc.match(/(?:TELÉFONO|Tel[eé]fono|Tlf|Tel)[^:]*[:\s-]+\s*([^\n\r]*)/i)) { infoClientePorLista[c.idList].telefono = m[1].replace(/[*_~`]/g, '').trim(); }
                    if (m = c.desc.match(/PROFESI[OÓ]N[^:]*[:\s-]+\s*([^\n\r]*)/i)) { infoClientePorLista[c.idList].profesion = m[1].replace(/[*_~`]/g, '').trim(); }
                    if (m = c.desc.match(/OBSERVACIONES[^:]*[:\s-]+\s*([^\n\r]*)/i)) { infoClientePorLista[c.idList].observaciones = m[1].replace(/[*_~`]/g, '').trim(); }
                  }

                  if (tarjetasPorLista[c.idList].length > maxCards) {
                    maxCards = tarjetasPorLista[c.idList].length;
                  }
                }
              });

              // Convertimos Sets a Arrays y calculamos el máximo de etiquetas
              lists.forEach(function (l) {
                var arrLabels = Array.from(etiquetasPorLista[l.id]);
                etiquetasPorLista[l.id] = arrLabels;
                if (arrLabels.length > maxLabels) maxLabels = arrLabels.length;
              });

              // Creamos y añadimos el encabezado dinámico
              var encabezado = ['Nombre lista', 'Nombre', 'DNI', 'Domicilio', 'Teléfono', 'Profesión', 'Observaciones'];
              for (var i = 1; i <= maxLabels; i++) {
                encabezado.push('Etiqueta ' + i);
              }
              for (var i = 1; i <= maxCards; i++) {
                encabezado.push('nombre tarjeta ' + i);
              }

              var headerRow = sheet.addRow(encabezado);
              headerRow.font = { bold: true }; // Ponemos el encabezado en negrita

              // Añadimos cada lista como una fila de datos
              lists.forEach(function (l) {
                var info = infoClientePorLista[l.id];
                var labels = etiquetasPorLista[l.id];
                var fila = [l.name, info.nombre, info.dni, info.domicilio, info.telefono, info.profesion, info.observaciones].concat(labels).concat(tarjetasPorLista[l.id]);
                sheet.addRow(fila);
              });

              // Ajustamos anchos de columna básicos
              sheet.getColumn(1).width = 25; // Nombre de la lista

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
              window.saveAs(blob, "Reporte_Listas.xlsx");
              return t.alert({ message: 'Descarga de listas completada.', duration: 3, display: 'success' });
            })
            .catch(function (err) {
              console.error("Error exportando listas:", err);
              return t.alert({ message: 'Error al exportar listas.', duration: 6, display: 'error' });
            });
        }
      }
    ];
  }
});
