/**
 * Lógica compartida para detectar el tipo de tablero basado en la tarjeta de configuración.
 * Busca una lista llamada "PARÁMETROS" y una tarjeta que empiece por "Tipo tablero".
 */
window.getBoardType = function(t) {
  // Obtenemos todas las listas del tablero
  return t.lists('id', 'name')
    .then(function(lists) {
      var paramList = lists.find(function(l) { 
        return l.name.toUpperCase() === 'PARÁMETROS'; 
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
          if (name.includes('CLIENTES')) return 'CLIENTES';
          if (name.includes('PLANNING')) return 'PLANNING';
          if (name.includes('FACTURACIÓN')) return 'FACTURACIÓN';

          return 'UNKNOWN';
        });
    })
    .catch(function(err) {
      console.error("Error al detectar el tipo de tablero:", err);
      return 'UNKNOWN';
    });
};
