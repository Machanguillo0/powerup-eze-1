console.log("Antigravity Power-Up Connector Cargado");

// Configuración global
var APP_KEY = '30767998cbb7b57a692e8bb50cfa9a9c';
var APP_NAME = 'Antigravity';
var BASE_URL = window.location.href.substring(0, window.location.href.lastIndexOf('/') + 1);

// Función de utilidad inyectada para detectar el tipo de tablero de forma segura
var getBoardTypeSafe = function(t) {
  return t.lists('all').then(function(lists) {
    var paramList = lists.find(function(l) { 
      var n = l.name.toUpperCase();
      return n === 'PARÁMETROS' || n === 'PARAMETROS'; 
    });
    if (!paramList) return 'CLIENTES';

    return t.cards('all').then(function(openCards) {
      return t.cards('closed').then(function(closedCards) {
        var allCards = openCards.concat(closedCards);
        var typeCard = allCards.find(function(c) {
          return c.idList === paramList.id && c.name.toLowerCase().startsWith('tipo tablero');
        });
        if (!typeCard) return 'CLIENTES';

        var name = typeCard.name.toUpperCase();
        if (name.includes('PLANNING')) return 'PLANNING';
        if (name.includes('FACTURACIÓN') || name.includes('FACTURACION')) return 'FACTURACIÓN';
        return 'CLIENTES';
      });
    });
  }).catch(function(err) {
    console.error("Error detectando tablero:", err);
    return 'CLIENTES';
  });
};

window.TrelloPowerUp.initialize({
  // Capability para añadir botones en el menú derecho de la tarjeta
  'card-buttons': function (t, options) {
    return getBoardTypeSafe(t).then(function(type) {
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
      title: 'Menú de acciones',
      icon: './icon.svg', // Recomendado un icono gris, pero el SVG actual servirá
      content: {
        type: 'iframe',
        url: t.signUrl('./section.html?v=47'), // Apuntamos a un nuevo archivo HTML que vamos a crear
        height: 250 // Altura inicial en píxeles (aumentado para evitar scrollbar)
      }
    };
  },
  // Capability para añadir un botón a nivel de tablero
  'board-buttons': function (t, options) {
    return getBoardTypeSafe(t).then(function(type) {
      if (type !== 'CLIENTES' && type !== 'FACTURACIÓN') return [];

      var buttons = [
        {
          icon: {
            dark: BASE_URL + 'export_tablero.svg',
            light: BASE_URL + 'export_tablero.svg'
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
                    
                    var propioMatch = c.desc.match(new RegExp("(?:^|\\n)[ \\t]*[*_~`#]*\\s*VEH[ÍI]CULO PROPIO[\\s\\S]*?" + vLA, "i"));
                    if (propioMatch) {
                      var pT = propioMatch[0];
                      if (m = pT.match(/(?:^|\n)[ \t]*MATR[ÍI]CULA[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.matricula1 = m[1].trim(); }
                      if (m = pT.match(/(?:^|\n)[ \t]*ASEGURADORA[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.aseguradora1 = m[1].trim(); }
                    }
                    var contrarioMatch = c.desc.match(new RegExp("(?:^|\\n)[ \\t]*[*_~`#]*\\s*VEH[ÍI]CULO CONTRARIO[\\s\\S]*?" + vLA, "i"));
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
        },
        {
          icon: {
            dark: BASE_URL + 'export_listas.svg',
            light: BASE_URL + 'export_listas.svg'
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

                    if (c.labels) {
                      c.labels.forEach(function (label) {
                        etiquetasPorLista[c.idList].add(label.name || label.color);
                      });
                    }

                    if (infoClientePorLista[c.idList].nombre === '' && c.desc) {
                      var m;
                      if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*NOMBRE Y APELLIDOS[^:]*[:\s-]+\s*([^\n\r]*)/i)) { infoClientePorLista[c.idList].nombre = m[1].replace(/[*_~`]/g, '').trim(); }
                      if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*DNI[^:]*[:\s-]+\s*([^\n\r]*)/i)) { infoClientePorLista[c.idList].dni = m[1].replace(/[*_~`]/g, '').trim(); }
                      if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*DOMICILIO[^:]*[:\s-]+\s*([^\n\r]*)/i)) { infoClientePorLista[c.idList].domicilio = m[1].replace(/[*_~`]/g, '').trim(); }
                      if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*(?:TELÉFONO|Tel[eé]fono|Tlf|Tel)[^:]*[:\s-]+\s*([^\n\r]*)/i)) { infoClientePorLista[c.idList].telefono = m[1].replace(/[*_~`]/g, '').trim(); }
                      if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*PROFESI[OÓ]N[^:]*[:\s-]+\s*([^\n\r]*)/i)) { infoClientePorLista[c.idList].profesion = m[1].replace(/[*_~`]/g, '').trim(); }
                      if (m = c.desc.match(/(?:^|\n)[ \t]*[*_~`#]*\s*OBSERVACIONES[^:]*[:\s-]+\s*([^\n\r]*)/i)) { infoClientePorLista[c.idList].observaciones = m[1].replace(/[*_~`]/g, '').trim(); }
                    }

                    if (tarjetasPorLista[c.idList].length > maxCards) {
                      maxCards = tarjetasPorLista[c.idList].length;
                    }
                  }
                });

                lists.forEach(function (l) {
                  var arrLabels = Array.from(etiquetasPorLista[l.id]);
                  etiquetasPorLista[l.id] = arrLabels;
                  if (arrLabels.length > maxLabels) maxLabels = arrLabels.length;
                });

                var encabezado = ['Nombre lista', 'Nombre', 'DNI', 'Domicilio', 'Teléfono', 'Profesión', 'Observaciones'];
                for (var i = 1; i <= maxLabels; i++) {
                  encabezado.push('Etiqueta ' + i);
                }
                for (var i = 1; i <= maxCards; i++) {
                  encabezado.push('nombre tarjeta ' + i);
                }

                var headerRow = sheet.addRow(encabezado);
                headerRow.font = { bold: true };

                lists.forEach(function (l) {
                  var info = infoClientePorLista[l.id];
                  var labels = etiquetasPorLista[l.id];
                  var fila = [l.name, info.nombre, info.dni, info.domicilio, info.telefono, info.profesion, info.observaciones].concat(labels).concat(tarjetasPorLista[l.id]);
                  sheet.addRow(fila);
                });

                sheet.getColumn(1).width = 25;

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

      if (type === 'FACTURACIÓN') {
        buttons.push({
          icon: {
            dark: BASE_URL + 'descarga_facturas.svg',
            light: BASE_URL + 'descarga_facturas.svg'
          },
          text: 'Descarga de facturas',
          condition: 'always',
          callback: function (t) {
            t.alert({
              message: 'Iniciando descarga de facturas...',
              duration: 3,
              display: 'info'
            });

            return Promise.all([
              t.cards('id', 'name', 'idList', 'pos'),
              t.lists('id', 'name')
            ])
              .then(function (results) {
                var cards = results[0];
                var lists = results[1];
                var mapListas = {};
                lists.forEach(function (l) { mapListas[l.id] = l.name; });

                var cardsByList = {};
                cards.forEach(function(c) {
                  if (!cardsByList[c.idList]) {
                    cardsByList[c.idList] = [];
                  }
                  cardsByList[c.idList].push(c);
                });

                var clientePorLista = {};
                lists.forEach(function(l) {
                  var listCards = cardsByList[l.id] || [];
                  listCards.sort(function(a, b) {
                    return a.pos - b.pos;
                  });

                  var firstCard = listCards.find(function(c) {
                    var nameLower = c.name.toLowerCase().trim();
                    return nameLower !== "facturas" && 
                           nameLower !== "facturación" && 
                           nameLower !== "facturacion" && 
                           nameLower !== "importes" && 
                           !nameLower.includes("historial de procesos") && 
                           !nameLower.includes("historial de avisos");
                  });

                  clientePorLista[l.id] = firstCard ? firstCard.name : "Sin Cliente";
                });

                var facturasCards = cards.filter(function(c) {
                  var nameLower = c.name.trim().toLowerCase();
                  return nameLower === "facturas" || nameLower === "facturación" || nameLower === "facturacion";
                });

                if (facturasCards.length === 0) {
                  return t.alert({
                    message: 'No se encontraron tarjetas de Facturas en este tablero.',
                    duration: 5,
                    display: 'warning'
                  });
                }

                return t.getRestApi().isAuthorized().then(function(auth) {
                  if (!auth) {
                    return t.alert({
                      message: 'Por favor, autoriza primero el Power-Up en el menú de acciones de cualquier tarjeta para poder descargar las facturas.',
                      duration: 8,
                      display: 'error'
                    });
                  }

                  return t.getRestApi().getToken().then(function(token) {
                    var appKey = t.getRestApi().appKey;

                    var fetchPromises = facturasCards.map(function(card) {
                      return fetch('https://api.trello.com/1/cards/' + card.id + '/checklists?key=' + appKey + '&token=' + token)
                        .then(function(res) {
                          if (!res.ok) throw new Error("Error fetching checklists");
                          return res.json();
                        })
                        .then(function(checklists) {
                          return { card: card, checklists: checklists };
                        })
                        .catch(function(err) {
                          console.error("Error al consultar checklists de tarjeta " + card.id + ":", err);
                          return { card: card, checklists: [] };
                        });
                    });

                    return Promise.all(fetchPromises).then(function(results) {
                      var invoiceRows = [];

                      results.forEach(function(res) {
                        var card = res.card;
                        var checklists = res.checklists;
                        var listName = mapListas[card.idList] || "Desconocida";
                        var clientName = clientePorLista[card.idList] || "Sin Cliente";

                        var ch = checklists.find(function(c) {
                          var n = c.name.trim().toLowerCase();
                          return n === "relación de facturas generadas" || 
                                 n === "relacion de facturas generadas" || 
                                 n === "relaación de facturas generadas" || 
                                 n === "relaacion de facturas generadas" ||
                                 n.indexOf("relación de facturas") > -1 ||
                                 n.indexOf("relacion de facturas") > -1 ||
                                 n.indexOf("relaación de facturas") > -1;
                        });

                        if (ch && ch.checkItems && ch.checkItems.length > 0) {
                          var checkItems = ch.checkItems.concat().sort(function(a, b) {
                            return a.pos - b.pos;
                          });

                          checkItems.forEach(function(item) {
                            var parts = item.name.split('|').map(function(s) { return s.trim(); });
                            var fDate = "";
                            var fNum = "";
                            var fAmount = "";
                            var fTotal = "";

                            if (parts.length >= 2) {
                              fDate = parts[0] || "";

                              var numPart = parts.find(function(p) { return /^(?:N[º°o]:|N[º°o]\s*:|Factura:|Num:)/i.test(p); });
                              if (numPart) {
                                fNum = numPart.replace(/^(?:N[º°o]:|N[º°o]\s*:|Factura:|Num:)\s*/i, '').trim();
                              } else {
                                fNum = parts[1] || "";
                              }

                              var basePart = parts.find(function(p) { return /^(?:Base:|Importe:)/i.test(p); });
                              if (basePart) {
                                fAmount = basePart.replace(/^(?:Base:|Importe:)\s*/i, '').trim();
                              } else {
                                fAmount = parts[parts.length - 1] || "";
                              }

                              var totalPart = parts.find(function(p) { return /^(?:Total:)/i.test(p); });
                              if (totalPart) {
                                fTotal = totalPart.replace(/^(?:Total:)\s*/i, '').trim();
                              } else {
                                fTotal = fAmount;
                              }
                            } else {
                              fDate = "";
                              fNum = item.name;
                              fAmount = "";
                              fTotal = "";
                            }

                            invoiceRows.push({
                              expediente: listName,
                              cliente: clientName,
                              fecha: fDate,
                              factura: fNum,
                              base: fAmount,
                              total: fTotal,
                              estado: item.state === "complete" ? "Cobrada" : "Pendiente"
                            });
                          });
                        }
                      });

                      if (invoiceRows.length === 0) {
                        return t.alert({
                          message: 'No se encontraron facturas registradas en los checklists de las tarjetas de Facturas.',
                          duration: 5,
                          display: 'warning'
                        });
                      }

                      return new Promise(function(resolve, reject) {
                        if (window.ExcelJS) return resolve();
                        var script = document.createElement('script');
                        script.src = "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js";
                        script.onload = function() { resolve(); };
                        script.onerror = reject;
                        document.head.appendChild(script);
                      })
                        .then(function() {
                          var workbook = new ExcelJS.Workbook();
                          var sheet = workbook.addWorksheet('Relación de Facturas');

                          sheet.columns = [
                            { header: 'Expediente / Lista', key: 'expediente', width: 25 },
                            { header: 'Cliente', key: 'cliente', width: 30 },
                            { header: 'Fecha', key: 'fecha', width: 15 },
                            { header: 'Nº Factura', key: 'factura', width: 15 },
                            { header: 'Importe Base', key: 'base', width: 15 },
                            { header: 'Importe Total', key: 'total', width: 15 },
                            { header: 'Estado', key: 'estado', width: 15 }
                          ];

                          var headerRow = sheet.getRow(1);
                          headerRow.font = { bold: true };
                          headerRow.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFE6F4EA' }
                          };

                          invoiceRows.forEach(function(row) {
                            sheet.addRow(row);
                          });

                          return workbook.xlsx.writeBuffer();
                        })
                        .then(function(buffer) {
                          return new Promise(function(resolve, reject) {
                            if (window.saveAs) return resolve(buffer);
                            var scriptFS = document.createElement('script');
                            scriptFS.src = "https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js";
                            scriptFS.onload = function() { resolve(buffer); };
                            scriptFS.onerror = reject;
                            document.head.appendChild(scriptFS);
                          });
                        })
                        .then(function(buffer) {
                          var blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
                          window.saveAs(blob, "Relacion_Facturas_Generadas.xlsx");
                          return t.alert({ message: 'Descarga de facturas completada.', duration: 3, display: 'success' });
                        });
                    });
                  });
                });
              })
              .catch(function (err) {
                console.error("Error al descargar facturas:", err);
                return t.alert({ message: 'Error al exportar facturas.', duration: 6, display: 'error' });
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
