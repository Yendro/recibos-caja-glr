function numeroALetras(numero) {
  const entero = Math.floor(numero);
  const decimales = Math.round((numero - entero) * 100);

  let texto = convertirNumero(entero);
  texto = ajustarMasculino(texto);

  const moneda = entero === 1 ? "PESO" : "PESOS";
  return `${texto} ${moneda} ${decimales.toString().padStart(2, "0")}/100 M.N.`;
}

function convertirNumero(numero) {
  numero = Math.floor(numero);
  switch (true) {
    case numero === 0:
      return "CERO";
    case numero <= 29:
      return unidadDecena(numero);
    case numero <= 99:
      return decenas(numero);
    case numero <= 999:
      return centenas(numero);
    case numero <= 999999: {
      const miles = Math.floor(numero / 1000);
      const restoMiles = numero % 1000;
      let textoMiles =
        miles === 1 ? "MIL" : `${ajustarMasculino(convertirNumero(miles))} MIL`;
      if (restoMiles > 0) {
        textoMiles += ` ${convertirNumero(restoMiles)}`;
      }
      return textoMiles;
    }
    case numero <= 999999999999: {
      const millones = Math.floor(numero / 1000000);
      const restoMillones = numero % 1000000;
      let textoMillones =
        millones === 1
          ? "UN MILLÓN"
          : `${ajustarMasculino(convertirNumero(millones))} MILLONES`;
      if (restoMillones > 0) {
        textoMillones += ` ${convertirNumero(restoMillones)}`;
      }
      return textoMillones;
    }
    default:
      return "CANTIDAD FUERA DE RANGO";
  }
}

function unidadDecena(numero) {
  const valores = {
    1: "UNO",
    2: "DOS",
    3: "TRES",
    4: "CUATRO",
    5: "CINCO",
    6: "SEIS",
    7: "SIETE",
    8: "OCHO",
    9: "NUEVE",
    10: "DIEZ",
    11: "ONCE",
    12: "DOCE",
    13: "TRECE",
    14: "CATORCE",
    15: "QUINCE",
    16: "DIECISÉIS",
    17: "DIECISIETE",
    18: "DIECIOCHO",
    19: "DIECINUEVE",
    20: "VEINTE",
    21: "VEINTIUNO",
    22: "VEINTIDÓS",
    23: "VEINTITRÉS",
    24: "VEINTICUATRO",
    25: "VEINTICINCO",
    26: "VEINTISÉIS",
    27: "VEINTISIETE",
    28: "VEINTIOCHO",
    29: "VEINTINUEVE",
  };
  return valores[numero] || "";
}

function decenas(numero) {
  const decenasTexto = {
    3: "TREINTA",
    4: "CUARENTA",
    5: "CINCUENTA",
    6: "SESENTA",
    7: "SETENTA",
    8: "OCHENTA",
    9: "NOVENTA",
  };
  const d = Math.floor(numero / 10);
  const u = numero % 10;
  let texto = decenasTexto[d];
  if (u > 0) {
    texto += ` Y ${unidadDecena(u)}`;
  }
  return texto;
}

function centenas(numero) {
  if (numero === 100) return "CIEN";
  const centenasTexto = {
    1: "CIENTO",
    2: "DOSCIENTOS",
    3: "TRESCIENTOS",
    4: "CUATROCIENTOS",
    5: "QUINIENTOS",
    6: "SEISCIENTOS",
    7: "SETECIENTOS",
    8: "OCHOCIENTOS",
    9: "NOVECIENTOS",
  };
  const c = Math.floor(numero / 100);
  const resto = numero % 100;
  let texto = centenasTexto[c];
  if (resto > 0) {
    texto += ` ${convertirNumero(resto)}`;
  }
  return texto;
}

function ajustarMasculino(texto) {
  if (texto.endsWith("UNO")) texto = texto.slice(0, -3) + "UN";
  if (texto.endsWith("VEINTIUNO")) texto = texto.slice(0, -9) + "VEINTIÚN";
  if (texto.endsWith(" Y UNO")) texto = texto.slice(0, -6) + " Y UN";
  return texto;
}
