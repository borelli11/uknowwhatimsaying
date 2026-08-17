/**
 * ==============================================================================
 *  imageProcessor.js
 * ==============================================================================
 *
 *  Este arquivo é o "cérebro" do sistema — toda a lógica pesada mora aqui:
 *
 *    1) Ler o arquivo .xlsx (planilha Excel) no navegador
 *    2) Encontrar e extrair as imagens que estão dentro da planilha
 *    3) Calcular o "hash" de cada imagem (uma espécie de impressão digital)
 *    4) Comparar duas imagens e dizer o quão parecidas elas são (%)
 *    5) Calcular a nitidez de uma imagem (pra saber se está borrada)
 *    6) Descobrir se uma peça está vencida, vencendo ou dentro do prazo
 *
 *  IMPORTANTE: este arquivo NUNCA mexe em elementos da tela (DOM).
 *  Ele só recebe dados, processa e devolve o resultado. Quem decide o que
 *  fazer com esse resultado (mostrar na tela, salvar, etc.) é o `app.js`.
 *  Isso é uma boa prática chamada "separação de responsabilidades": cada
 *  arquivo tem um único trabalho, o que facilita muito a manutenção.
 *
 *  Todas as funções abaixo são exportadas (`export`) para que o `app.js`
 *  possa importá-las com:
 *      import { nomeDaFuncao } from './imageProcessor.js';
 *
 * ==============================================================================
 */

// ------------------------------------------------------------------------------
// 1) CARREGAMENTO DE IMAGENS
// ------------------------------------------------------------------------------

/**
 * Carrega uma imagem (a partir de um arquivo/Blob ou de uma URL) e devolve
 * um elemento <img> já pronto para ser usado em um <canvas>.
 *
 * Isso é necessário porque, no navegador, para "olhar" para os pixels de uma
 * imagem, primeiro precisamos desenhá-la em um <canvas> — e para desenhar,
 * precisamos de um objeto Image() carregado.
 */
export function carregarImagem(origem) {
  return new Promise((resolve, rejeitar) => {
    const imagem = new Image();
    imagem.crossOrigin = "anonymous";
    imagem.onload = () => resolve(imagem);
    imagem.onerror = () => rejeitar(new Error("Falha ao carregar a imagem."));
    imagem.src = typeof origem === "string" ? origem : URL.createObjectURL(origem);
  });
}

/**
 * Gera uma miniatura (thumbnail) em formato "data URL" (uma string que já
 * contém a imagem codificada) para ser usada em previews e no histórico.
 *
 * @param {Blob} blob      - a imagem original
 * @param {number} tamanho - tamanho máximo (em pixels) do lado maior
 */
export async function gerarMiniatura(blob, tamanho = 200) {
  try {
    const imagem = await carregarImagem(blob);
    const canvas = document.createElement("canvas");
    const escala = Math.min(1, tamanho / Math.max(imagem.naturalWidth, imagem.naturalHeight));
    canvas.width = Math.max(1, Math.round(imagem.naturalWidth * escala));
    canvas.height = Math.max(1, Math.round(imagem.naturalHeight * escala));
    canvas.getContext("2d").drawImage(imagem, 0, 0, canvas.width, canvas.height);
    return canvas.toDataURL("image/jpeg", 0.7);
  } catch {
    return null; // se der erro, simplesmente não teremos miniatura — não é crítico
  }
}

/** Devolve as dimensões (largura/altura em pixels) de uma imagem. */
export async function obterDimensoes(blob) {
  const imagem = await carregarImagem(blob);
  return { largura: imagem.naturalWidth, altura: imagem.naturalHeight };
}

/** Devolve o peso do arquivo em KB (arredondado em 2 casas decimais). */
export function obterPesoKb(arquivo) {
  return Math.round((arquivo.size / 1024) * 100) / 100;
}

// ------------------------------------------------------------------------------
// 2) PHASH (IMPRESSÃO DIGITAL DA IMAGEM) — usado para comparar semelhança
// ------------------------------------------------------------------------------
//
//  O que é um "pHash" (perceptual hash)?
//  É uma técnica que resume uma imagem em um código de 64 bits (0s e 1s) de
//  forma que IMAGENS PARECIDAS gerem CÓDIGOS PARECIDOS — mesmo que uma tenha
//  sido levemente recortada, comprimida ou redimensionada.
//
//  O algoritmo, resumidamente:
//    a) Reduz a imagem para 32x32 pixels e converte para tons de cinza
//    b) Aplica uma transformação matemática chamada DCT (Discrete Cosine
//       Transform) — a mesma usada em compressão de imagens tipo JPEG
//    c) Pega apenas o "canto" 8x8 de baixa frequência (as informações mais
//       importantes/gerais da imagem, ignorando detalhes finos)
//    d) Compara cada um dos 64 valores com a mediana: acima = 1, abaixo = 0
//    e) O resultado são 64 bits — o pHash da imagem
//
//  Esse é o MESMO algoritmo usado pela biblioteca Python `imagehash.phash`,
//  reimplementado aqui em JavaScript puro para rodar 100% no navegador.
// ------------------------------------------------------------------------------

const TAMANHO_DCT = 32; // reduzimos toda imagem para 32x32 antes de processar

// Tabela de cossenos pré-calculada — evita recalcular Math.cos() milhares
// de vezes a cada imagem (otimização de performance).
const TABELA_COSSENOS = (() => {
  const tabela = new Float64Array(TAMANHO_DCT * TAMANHO_DCT);
  for (let k = 0; k < TAMANHO_DCT; k++) {
    for (let n = 0; n < TAMANHO_DCT; n++) {
      tabela[k * TAMANHO_DCT + n] = Math.cos((Math.PI / TAMANHO_DCT) * (n + 0.5) * k);
    }
  }
  return tabela;
})();

/** Aplica a transformação DCT em uma linha (ou coluna) de 32 valores. */
function aplicarDct1d(entrada, saida) {
  const n = TAMANHO_DCT;
  for (let k = 0; k < n; k++) {
    let soma = 0;
    const deslocamento = k * n;
    for (let i = 0; i < n; i++) soma += entrada[i] * TABELA_COSSENOS[deslocamento + i];
    saida[k] = soma;
  }
}

/**
 * Calcula o pHash (64 bits) de uma imagem já carregada (elemento <img>).
 * Devolve um array de 64 posições, cada uma com 0 ou 1.
 */
function calcularPHashDeImagemCarregada(imagemCarregada) {
  const n = TAMANHO_DCT;

  // Desenha a imagem reduzida (32x32) em um canvas invisível
  const canvas = document.createElement("canvas");
  canvas.width = n;
  canvas.height = n;
  const contexto = canvas.getContext("2d");
  contexto.drawImage(imagemCarregada, 0, 0, n, n);
  const pixels = contexto.getImageData(0, 0, n, n).data;

  // Converte para tons de cinza (fórmula padrão de luminância)
  const cinza = new Float64Array(n * n);
  for (let i = 0; i < n * n; i++) {
    cinza[i] = 0.299 * pixels[i * 4] + 0.587 * pixels[i * 4 + 1] + 0.114 * pixels[i * 4 + 2];
  }

  // DCT aplicada primeiro nas linhas...
  const dctLinhas = new Float64Array(n * n);
  const entradaLinha = new Float64Array(n);
  const saidaLinha = new Float64Array(n);
  for (let y = 0; y < n; y++) {
    for (let x = 0; x < n; x++) entradaLinha[x] = cinza[y * n + x];
    aplicarDct1d(entradaLinha, saidaLinha);
    for (let x = 0; x < n; x++) dctLinhas[y * n + x] = saidaLinha[x];
  }

  // ...depois nas colunas (DCT 2D = DCT das linhas + DCT das colunas)
  const dctCompleta = new Float64Array(n * n);
  const entradaColuna = new Float64Array(n);
  const saidaColuna = new Float64Array(n);
  for (let x = 0; x < n; x++) {
    for (let y = 0; y < n; y++) entradaColuna[y] = dctLinhas[y * n + x];
    aplicarDct1d(entradaColuna, saidaColuna);
    for (let y = 0; y < n; y++) dctCompleta[y * n + x] = saidaColuna[y];
  }

  // Pega só o quadrante 8x8 superior-esquerdo (as frequências mais "importantes")
  const blocoBaixaFrequencia = new Float64Array(64);
  for (let y = 0; y < 8; y++) {
    for (let x = 0; x < 8; x++) blocoBaixaFrequencia[y * 8 + x] = dctCompleta[y * n + x];
  }

  // Calcula a mediana (ignorando a posição 0, que é só o "brilho médio" da imagem)
  const valoresOrdenados = Array.from(blocoBaixaFrequencia).slice(1).sort((a, b) => a - b);
  const mediana = valoresOrdenados[Math.floor(valoresOrdenados.length / 2)];

  // Gera o hash final: 1 se o valor está acima da mediana, 0 caso contrário
  const hash = new Uint8Array(64);
  for (let i = 0; i < 64; i++) hash[i] = blocoBaixaFrequencia[i] > mediana ? 1 : 0;
  return hash;
}

/** Calcula o pHash de uma imagem a partir de um Blob/File. */
export async function calcularPHash(blob) {
  const imagemCarregada = await carregarImagem(blob);
  return calcularPHashDeImagemCarregada(imagemCarregada);
}

/**
 * Distância de Hamming: conta em quantas das 64 posições dois hashes diferem.
 * 0 = hashes idênticos (imagens praticamente iguais).
 * Quanto maior o número, mais diferentes são as imagens.
 */
export function calcularDistanciaHamming(hashA, hashB) {
  let distancia = 0;
  for (let i = 0; i < 64; i++) {
    if (hashA[i] !== hashB[i]) distancia++;
  }
  return distancia;
}

/** Converte a distância de Hamming (0 a 64) em uma porcentagem de similaridade. */
export function calcularSimilaridade(distancia) {
  return Math.round(((64 - distancia) / 64) * 10000) / 100;
}

// ------------------------------------------------------------------------------
// 3) NITIDEZ DA IMAGEM
// ------------------------------------------------------------------------------
//
//  Heurística simples para detectar fotos borradas/baixa resolução: aplicamos
//  um filtro de detecção de bordas (o mesmo kernel usado pelo PIL.FIND_EDGES
//  do Python: [-1,-1,-1; -1,8,-1; -1,-1,-1]) e medimos a VARIÂNCIA do
//  resultado. Fotos nítidas têm bordas bem definidas → variância alta.
//  Fotos borradas têm transições suaves → variância baixa.
// ------------------------------------------------------------------------------

export async function calcularNitidez(blob) {
  const imagemCarregada = await carregarImagem(blob);

  // Reduz imagens muito grandes antes de processar, só para manter a
  // performance boa (não precisamos da resolução total pra essa heurística).
  const dimensaoMaxima = 512;
  let largura = imagemCarregada.naturalWidth;
  let altura = imagemCarregada.naturalHeight;
  if (largura > dimensaoMaxima || altura > dimensaoMaxima) {
    const escala = dimensaoMaxima / Math.max(largura, altura);
    largura = Math.max(1, Math.round(largura * escala));
    altura = Math.max(1, Math.round(altura * escala));
  }

  const canvas = document.createElement("canvas");
  canvas.width = largura;
  canvas.height = altura;
  const contexto = canvas.getContext("2d");
  contexto.drawImage(imagemCarregada, 0, 0, largura, altura);
  const pixels = contexto.getImageData(0, 0, largura, altura).data;

  const cinza = new Float64Array(largura * altura);
  for (let i = 0; i < largura * altura; i++) {
    cinza[i] = 0.299 * pixels[i * 4] + 0.587 * pixels[i * 4 + 1] + 0.114 * pixels[i * 4 + 2];
  }

  // Aplica o kernel de detecção de bordas pixel a pixel (ignorando a borda
  // externa da imagem, que não tem vizinhos completos dos 2 lados)
  let soma = 0;
  let somaQuadrados = 0;
  let contagem = 0;
  for (let y = 1; y < altura - 1; y++) {
    for (let x = 1; x < largura - 1; x++) {
      const centro = 8 * cinza[y * largura + x];
      const vizinhos =
        cinza[(y - 1) * largura + (x - 1)] + cinza[(y - 1) * largura + x] + cinza[(y - 1) * largura + (x + 1)] +
        cinza[y * largura + (x - 1)] + cinza[y * largura + (x + 1)] +
        cinza[(y + 1) * largura + (x - 1)] + cinza[(y + 1) * largura + x] + cinza[(y + 1) * largura + (x + 1)];
      const valorBorda = centro - vizinhos;
      soma += valorBorda;
      somaQuadrados += valorBorda * valorBorda;
      contagem++;
    }
  }

  if (contagem === 0) return 0;
  const media = soma / contagem;
  return somaQuadrados / contagem - media * media; // variância = E[x²] - (E[x])²
}

/** Diz se uma imagem tem "boa nitidez" com base no limiar configurado. */
export function avaliarNitidez(score, limiar) {
  return score >= limiar;
}

// ------------------------------------------------------------------------------
// 4) LEITURA DO ARQUIVO EXCEL (.xlsx)
// ------------------------------------------------------------------------------
//
//  Um arquivo .xlsx é, por baixo dos panos, um arquivo .zip contendo vários
//  XMLs. Usamos duas bibliotecas via CDN para lidar com isso:
//
//    - SheetJS (window.XLSX): lê os VALORES das células (texto, datas, etc.)
//    - JSZip   (window.JSZip): abre o .xlsx como um .zip para conseguirmos
//               extrair as IMAGENS embutidas nele (o SheetJS sozinho não
//               extrai imagens, só valores de células)
// ------------------------------------------------------------------------------

/** Lê o arquivo .xlsx tanto pelo SheetJS (valores) quanto pelo JSZip (bruto). */
export async function lerArquivoExcel(arquivo) {
  const bufferDoArquivo = await arquivo.arrayBuffer();
  const planilha = window.XLSX.read(bufferDoArquivo, { type: "array", cellDates: true });
  const zip = await window.JSZip.loadAsync(bufferDoArquivo);
  return { planilha, zip };
}

/**
 * Dentro do .xlsx, cada aba tem um arquivo XML próprio. Esta função descobre
 * o "caminho" desse XML a partir do NOME da aba (ex: "Desconsiderados").
 */
async function encontrarCaminhoXmlDaAba(zip, nomeDaAba) {
  const [xmlWorkbook, xmlRelacionamentos] = await Promise.all([
    zip.file("xl/workbook.xml").async("string"),
    zip.file("xl/_rels/workbook.xml.rels").async("string"),
  ]);

  const parser = new DOMParser();
  const docWorkbook = parser.parseFromString(xmlWorkbook, "text/xml");
  const docRelacionamentos = parser.parseFromString(xmlRelacionamentos, "text/xml");

  // Passo 1: encontrar o "id de relacionamento" (r:id) da aba pelo nome
  let idRelacionamento = null;
  const abas = docWorkbook.getElementsByTagName("sheet");
  for (let i = 0; i < abas.length; i++) {
    if (abas[i].getAttribute("name") === nomeDaAba) {
      idRelacionamento =
        abas[i].getAttribute("r:id") ||
        abas[i].getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
      break;
    }
  }
  if (!idRelacionamento) return null;

  // Passo 2: usar esse id para achar o caminho real do arquivo XML da aba
  const relacionamentos = docRelacionamentos.getElementsByTagName("Relationship");
  for (let i = 0; i < relacionamentos.length; i++) {
    if (relacionamentos[i].getAttribute("Id") === idRelacionamento) {
      let destino = relacionamentos[i].getAttribute("Target");
      destino = destino.startsWith("/") ? destino.substring(1) : "xl/" + destino;
      return destino;
    }
  }
  return null;
}

/** Resolve um caminho relativo (ex: "../media/imagem1.png") em um caminho absoluto dentro do zip. */
function resolverCaminhoRelativo(diretorioBase, destino) {
  if (destino.startsWith("/")) return destino.substring(1);
  if (destino.startsWith("../")) {
    const diretorioPai = diretorioBase.substring(0, diretorioBase.lastIndexOf("/"));
    return diretorioPai + "/" + destino.substring(3);
  }
  return diretorioBase + "/" + destino;
}

/** Encontra os arquivos "drawing" (desenhos/imagens) associados a uma aba. */
async function encontrarDrawingsDaAba(zip, caminhoXmlDaAba) {
  const diretorio = caminhoXmlDaAba.substring(0, caminhoXmlDaAba.lastIndexOf("/"));
  const nomeArquivo = caminhoXmlDaAba.substring(caminhoXmlDaAba.lastIndexOf("/") + 1);
  const caminhoRels = `${diretorio}/_rels/${nomeArquivo}.rels`;

  const arquivoRels = zip.file(caminhoRels);
  if (!arquivoRels) return [];

  const xmlRels = await arquivoRels.async("string");
  const doc = new DOMParser().parseFromString(xmlRels, "text/xml");
  const relacionamentos = doc.getElementsByTagName("Relationship");

  const drawings = [];
  for (let i = 0; i < relacionamentos.length; i++) {
    const tipo = relacionamentos[i].getAttribute("Type") || "";
    if (tipo.includes("/drawing")) {
      const destino = relacionamentos[i].getAttribute("Target");
      drawings.push(resolverCaminhoRelativo(diretorio, destino));
    }
  }
  return drawings;
}

/**
 * Lê um arquivo "drawing" do Excel e devolve a lista de imagens nele,
 * já com a linha/coluna da célula onde cada imagem está ancorada.
 */
async function extrairImagensDoDrawing(zip, caminhoDrawing) {
  const xmlDrawing = await zip.file(caminhoDrawing).async("string");
  const diretorio = caminhoDrawing.substring(0, caminhoDrawing.lastIndexOf("/"));
  const nomeArquivo = caminhoDrawing.substring(caminhoDrawing.lastIndexOf("/") + 1);
  const caminhoRels = `${diretorio}/_rels/${nomeArquivo}.rels`;

  // Mapa "id do relacionamento" -> "caminho real do arquivo de imagem"
  const mapaDeImagens = {};
  const arquivoRels = zip.file(caminhoRels);
  if (arquivoRels) {
    const xmlRels = await arquivoRels.async("string");
    const docRels = new DOMParser().parseFromString(xmlRels, "text/xml");
    const relacionamentos = docRels.getElementsByTagName("Relationship");
    for (let i = 0; i < relacionamentos.length; i++) {
      const id = relacionamentos[i].getAttribute("Id");
      const destino = relacionamentos[i].getAttribute("Target");
      mapaDeImagens[id] = resolverCaminhoRelativo(diretorio, destino);
    }
  }

  const docDrawing = new DOMParser().parseFromString(xmlDrawing, "text/xml");

  function coletarTags(...nomes) {
    return nomes.flatMap((nome) => Array.from(docDrawing.getElementsByTagName(nome)));
  }
  function primeiroTextoFilho(elemento, ...nomesPossiveis) {
    for (const nome of nomesPossiveis) {
      const encontrados = elemento.getElementsByTagName(nome);
      if (encontrados.length) return encontrados[0].textContent;
    }
    return null;
  }

  // Uma imagem pode estar "ancorada" de 3 formas diferentes no XML —
  // tratamos todas para não perder nenhuma imagem.
  const ancoras = coletarTags(
    "xdr:oneCellAnchor", "xdr:twoCellAnchor", "xdr:absoluteAnchor",
    "oneCellAnchor", "twoCellAnchor", "absoluteAnchor"
  );

  const imagensEncontradas = [];
  for (const ancora of ancoras) {
    // Descobre em qual célula (linha/coluna) a imagem começa
    const elementoOrigem =
      ancora.getElementsByTagName("xdr:from")[0] || ancora.getElementsByTagName("from")[0];
    let linha = 0;
    let coluna = 0;
    if (elementoOrigem) {
      coluna = parseInt(primeiroTextoFilho(elementoOrigem, "xdr:col", "col") || "0", 10);
      linha = parseInt(primeiroTextoFilho(elementoOrigem, "xdr:row", "row") || "0", 10);
    }

    // Encontra a referência da imagem em si (dentro de <pic><blipFill><blip>)
    const elementoImagem =
      ancora.getElementsByTagName("xdr:pic")[0] || ancora.getElementsByTagName("pic")[0];
    if (!elementoImagem) continue;

    const blip =
      elementoImagem.getElementsByTagName("a:blip")[0] ||
      elementoImagem.getElementsByTagName("blip")[0];
    if (!blip) continue;

    const idEmbed =
      blip.getAttribute("r:embed") ||
      blip.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed");
    if (!idEmbed) continue;

    const caminhoDoArquivo = mapaDeImagens[idEmbed];
    if (!caminhoDoArquivo) continue;

    // A planilha usa índice começando em 0; nós convertemos para base 1
    // (que é como o Excel mostra as linhas/colunas para o usuário)
    imagensEncontradas.push({ linha: linha + 1, coluna: coluna + 1, caminhoDoArquivo });
  }
  return imagensEncontradas;
}

/**
 * Função principal de extração: recebe o zip do .xlsx e o nome da aba, e
 * devolve a lista de TODAS as imagens embutidas naquela aba, cada uma já
 * com a linha/coluna onde está posicionada na planilha.
 */
export async function extrairImagensDaAba(zip, nomeDaAba) {
  const caminhoXmlDaAba = await encontrarCaminhoXmlDaAba(zip, nomeDaAba);
  if (!caminhoXmlDaAba) {
    throw new Error(`Não foi possível localizar o XML da aba '${nomeDaAba}'.`);
  }

  const drawings = await encontrarDrawingsDaAba(zip, caminhoXmlDaAba);
  if (drawings.length === 0) return [];

  const todasAsImagens = [];
  for (const caminhoDrawing of drawings) {
    const imagensDesteDrawing = await extrairImagensDoDrawing(zip, caminhoDrawing);
    todasAsImagens.push(...imagensDesteDrawing);
  }
  return todasAsImagens;
}

// ------------------------------------------------------------------------------
// 5) LEITURA DE CÉLULAS (considerando células mescladas)
// ------------------------------------------------------------------------------

/**
 * Lê o valor de uma célula, resolvendo automaticamente o caso de células
 * MESCLADAS: no Excel, quando várias células são mescladas em uma só, o
 * valor real só fica guardado na célula do canto superior-esquerdo. Esta
 * função encontra esse valor mesmo que a gente peça uma célula "do meio".
 */
export function obterValorCelula(aba, referenciaDaCelula) {
  const celula = aba[referenciaDaCelula];
  if (celula !== undefined) return celula;

  if (!aba["!merges"]) return null;

  const posicaoAlvo = window.XLSX.utils.decode_cell(referenciaDaCelula);
  for (const intervalo of aba["!merges"]) {
    const dentroDoIntervalo =
      posicaoAlvo.r >= intervalo.s.r && posicaoAlvo.r <= intervalo.e.r &&
      posicaoAlvo.c >= intervalo.s.c && posicaoAlvo.c <= intervalo.e.c;
    if (dentroDoIntervalo) {
      const referenciaMestre = window.XLSX.utils.encode_cell({ r: intervalo.s.r, c: intervalo.s.c });
      return aba[referenciaMestre] || null;
    }
  }
  return null;
}

/** Converte o valor "bruto" de uma célula para algo utilizável (texto ou Date). */
export function converterValorDaCelula(celula) {
  if (!celula) return null;
  if (celula.t === "d" || celula.v instanceof Date) {
    return celula.v instanceof Date ? celula.v : new Date(celula.v);
  }
  return celula.v;
}

// ------------------------------------------------------------------------------
// 6) REGRAS DE NEGÓCIO — VENCIMENTO
// ------------------------------------------------------------------------------

/**
 * Classifica o status de vencimento de uma peça com base na data e na lista
 * de "dias de alerta" configurada (ex: avisar quando faltam 15, 7, 3 ou 0 dias).
 */
export function classificarVencimento(dataVencimento, diasDeAlerta) {
  if (!(dataVencimento instanceof Date) || isNaN(dataVencimento.getTime())) {
    return "SEM DATA / OK";
  }

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  const dataAlvo = new Date(dataVencimento);
  dataAlvo.setHours(0, 0, 0, 0);

  const diasRestantes = Math.round((dataAlvo.getTime() - hoje.getTime()) / 86400000);

  if (diasRestantes < 0) return "VENCIDA";
  if (diasRestantes === 0) return "VENCE HOJE";
  if (diasDeAlerta.includes(diasRestantes)) return `VENCE EM ${diasRestantes} DIA(S)`;
  return "OK (fora da janela de alerta)";
}

/** Formata uma data no padrão brasileiro (DD/MM/AAAA). */
export function formatarDataBr(data) {
  if (!(data instanceof Date) || isNaN(data.getTime())) {
    return data != null ? String(data) : "-";
  }
  const dia = String(data.getDate()).padStart(2, "0");
  const mes = String(data.getMonth() + 1).padStart(2, "0");
  const ano = data.getFullYear();
  return `${dia}/${mes}/${ano}`;
}

// ------------------------------------------------------------------------------
// 7) UTILITÁRIO DE COLUNAS (ex: "B" <-> 2)
// ------------------------------------------------------------------------------

/** Converte letra(s) de coluna para número (A=1, B=2, ..., AA=27...). */
export function letraColunaParaIndice(letras) {
  let numero = 0;
  const texto = String(letras).toUpperCase();
  for (let i = 0; i < texto.length; i++) numero = numero * 26 + (texto.charCodeAt(i) - 64);
  return numero;
}

/** Converte número de coluna para letra(s) (1=A, 2=B, ..., 27=AA...). */
export function indiceParaLetraColuna(numero) {
  let letras = "";
  while (numero > 0) {
    const resto = (numero - 1) % 26;
    letras = String.fromCharCode(65 + resto) + letras;
    numero = Math.floor((numero - 1) / 26);
  }
  return letras;
}

// ------------------------------------------------------------------------------
// 8) VALIDAÇÃO DE LINKS (ex: "Link do Adobe")
// ------------------------------------------------------------------------------
//
//  OBJETIVO: dado o link de uma célula, descobrir se ele aponta para uma
//  imagem que existe e carrega com sucesso (equivalente a um "200 OK"), ou
//  se está quebrado (equivalente a um "404" / página de erro).
//
//  POR QUE NÃO USAMOS fetch()/HEAD AQUI:
//  A maioria desses links aponta para um domínio diferente do nosso app
//  (ex: domínios da Adobe) — isso é uma requisição "cross-origin". O
//  navegador tem uma proteção chamada CORS que, nesses casos:
//    - BLOQUEIA a leitura do status HTTP de um fetch() normal; e
//    - se usarmos fetch(url, { mode: "no-cors" }) pra "contornar" o bloqueio,
//      a resposta vem "opaca": o navegador some com o status code e o fetch
//      quase sempre parece "sucesso" mesmo quando o link está quebrado.
//  Ou seja: fetch NÃO é confiável aqui, mesmo tecnicamente "funcionando".
//
//  A ALTERNATIVA QUE REALMENTE FUNCIONA 100% NO NAVEGADOR (sem servidor,
//  sem back-end, sem burlar CORS) é deixar o próprio navegador tentar
//  carregar o link como se fosse uma imagem, usando um elemento Image():
//    - Se o link aponta pra uma imagem válida  -> dispara o evento "load".
//    - Se o link está quebrado (404, domínio errado, etc.) -> dispara "error".
//    - Se o servidor simplesmente não responde -> usamos um tempo-limite
//      (timeout) de segurança, pra não esperar pra sempre.
//  Essa é exatamente a checagem que os navegadores fazem "por baixo dos
//  panos" ao exibir uma <img>, então ela não esbarra em CORS.
// ------------------------------------------------------------------------------

/** Tempo padrão de espera antes de considerar um link "sem resposta". */
export const TIMEOUT_PADRAO_VERIFICACAO_LINK_MS = 6000;

/**
 * Verifica se `url` aponta para uma imagem que carrega com sucesso.
 *
 * @param {string} url       - o link a ser verificado (ex: valor da célula "Link do Adobe").
 * @param {number} timeoutMs - quanto tempo esperar antes de desistir (padrão 6s).
 * @returns {Promise<{valido: boolean, status: "sem_link"|"valido"|"quebrado"|"timeout"}>}
 */
export function verificarLinkImagem(url, timeoutMs = TIMEOUT_PADRAO_VERIFICACAO_LINK_MS) {
  return new Promise((resolve) => {
    const linkLimpo = (url ?? "").toString().trim();
    if (!linkLimpo) {
      resolve({ valido: false, status: "sem_link" });
      return;
    }

    const imagemDeTeste = new Image();
    let jaResolveu = false;

    const concluir = (valido, status) => {
      if (jaResolveu) return; // evita chamar resolve() duas vezes (ex: load + timeout na corrida)
      jaResolveu = true;
      clearTimeout(temporizador);
      imagemDeTeste.onload = null;
      imagemDeTeste.onerror = null;
      resolve({ valido, status });
    };

    const temporizador = setTimeout(() => concluir(false, "timeout"), timeoutMs);

    imagemDeTeste.onload = () => concluir(true, "valido");
    imagemDeTeste.onerror = () => concluir(false, "quebrado");
    imagemDeTeste.referrerPolicy = "no-referrer"; // evita bloqueios de hotlink baseados em referer
    imagemDeTeste.src = linkLimpo;
  });
}

/** Traduz o resultado técnico de `verificarLinkImagem` em um texto amigável (PT-BR). */
export function descreverResultadoLink(resultado) {
  if (!resultado) return "Não verificado";
  switch (resultado.status) {
    case "sem_link": return "Sem link";
    case "valido": return "Válido";
    case "quebrado": return "Inválido (link quebrado)";
    case "timeout": return "Inválido (tempo esgotado)";
    default: return "Inválido";
  }
}

