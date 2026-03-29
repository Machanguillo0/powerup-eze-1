function extractFields(desc) {
    var fields = {
        nombre: "Cliente Desconocido",
        dni: "___________",
        domicilio: "___________",
        telefono: "___________",
        profesion: "___________",
        observaciones: "___________",
        parrafada: ""
    };

    if (!desc) return fields;

    var m;
    // Single line extractions
    if (m = desc.match(/[*_~`#]*\s*NOMBRE Y APELLIDOS[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.nombre = m[1].replace(/[*_~`]/g, '').trim(); }
    if (m = desc.match(/[*_~`#]*\s*DNI[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.dni = m[1].replace(/[*_~`]/g, '').trim(); }
    if (m = desc.match(/[*_~`#]*\s*DOMICILIO[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.domicilio = m[1].replace(/[*_~`]/g, '').trim(); }
    if (m = desc.match(/[*_~`#]*\s*(?:TELÉFONO|Tel[eé]fono|Tlf|Tel)[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.telefono = m[1].replace(/[*_~`]/g, '').trim(); }
    if (m = desc.match(/[*_~`#]*\s*PROFESI[OÓ]N[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.profesion = m[1].replace(/[*_~`]/g, '').trim(); }
    if (m = desc.match(/[*_~`#]*\s*OBSERVACIONES[^:]*[:\s-]+\s*([^\n\r]*)/i)) { fields.observaciones = m[1].replace(/[*_~`]/g, '').trim(); }

    // Multi-line extraction for PARRAFADA
    var regexParrafada = /PARRAFADA[^:]*[:\s-]+\s*([\s\S]*?)(?=\n\s*[*_~`#]*\s*(?:NOMBRE|DNI|DOMICILIO|TEL[Ée]FONO|TLF|PROFESI[ÓO]N|OBSERVACIONES|$)|\n\s*[*_~`#]*\s*[A-ZÁÉÍÓÚÑ]{3,}\s*:|$)/i;
    if (m = desc.match(regexParrafada)) {
        fields.parrafada = m[1].replace(/[*_~`]/g, '').trim();
    }

    return fields;
}

const testDesc = `NOMBRE Y APELLIDOS: Juan Pérez del río

DNI: 567845S

DOMICILIO: c/ general vives 56

TELÉFONO: 767587

PROFESIÓN:

PARRAFADA: Este es el párrafo 1.

Y este es el dos.

OBSERVACIONES: ninguna
------`;

const result = extractFields(testDesc);
console.log("--- RESULTS ---");
console.log("PARRAFADA:", JSON.stringify(result.parrafada));
console.log("OBSERVACIONES:", JSON.stringify(result.observaciones));
