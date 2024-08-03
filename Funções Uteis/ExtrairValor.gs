function EXTRAIR_VALOR(valor) {

  //Remove espaços e $
  valor = valor.replace(/\s|\$/g,'');

  if (valor.toLowerCase().endsWith('k')) {
    valor = valor.replace(/k/i, ''); //Remove o k do valor
    valor = parseFloat(valor) * 1000; //Converte o que sobrou para float e mutiplica por 1000, pois k = 1000
  } else if (valor.toLowerCase().endsWith('m')) {
    valor = valor.replace(/m/i, '');
    valor = parseFloat(valor) * 1000000;
  } else if (valor.toLowerCase().endsWith('b')) {
    valor = valor.replace(/b/i, '');
    valor = parseFloat(valor) * 1000000000;
  } else if (valor.toLowerCase().endsWith('t')) {
    valor = valor.replace(/t/i, '');
    valor = parseFloat(valor) * 1000000000000;
  }

  //Converte para número, mesmo quando não tiver sufixo
  return Number(valor);
}
