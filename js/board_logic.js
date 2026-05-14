/**
 * Lógica compartida para detectar el tipo de tablero basado en la tarjeta de configuración.
 * Busca una lista llamada "PARÁMETROS" y una tarjeta que empiece por "Tipo tablero".
 */
window.getBoardType = function(t) {
  // Obtenemos todas las listas del tablero
  return t.lists('id', 'name')
    .then(function(lists) {
      var paramList = lists.find(function(l) { 
        var n = l.name.toUpperCase();
        return n === 'PARÁMETROS' || n === 'PARAMETROS'; 
      });
      
      if (!paramList) {
        console.log("No se encontró la lista PARÁMETROS");
        return 'UNKNOWN';
      }

      // Obtenemos las tarjetas del tablero para buscar la de configuración
      return t.cards('id', 'name', 'idList')
        .then(function(cards) {
          var typeCard = cards.find(function(c) {
            return c.idList === paramList.id && c.name.toLowerCase().startsWith('tipo tablero');
          });

          if (!typeCard) {
            console.log("No se encontró la tarjeta 'Tipo tablero' en la lista PARÁMETROS");
            return 'UNKNOWN';
          }

          // Extraemos el tipo del título de la tarjeta
          var name = typeCard.name.toUpperCase();
          if (name.includes('CLIENTE')) return 'CLIENTES';
          if (name.includes('PLANNING')) return 'PLANNING';
          if (name.includes('FACTURACIÓN') || name.includes('FACTURACION')) return 'FACTURACIÓN';

          return 'CLIENTES'; // Default a CLIENTES si hay lista pero no se reconoce el tipo
        });
    })
    .then(function(type) {
      // Si llegamos aquí con UNKNOWN (porque no había lista), devolvemos CLIENTES por defecto
      // para no romper la experiencia de usuario inicial.
      return type === 'UNKNOWN' ? 'CLIENTES' : type;
    })
    .catch(function(err) {
      console.error("Error al detectar el tipo de tablero:", err);
      return 'UNKNOWN';
    });
};
