/**
 * ==============================================================================
 *  app.js
 * ==============================================================================
 *
 *  Este arquivo cuida da INTERFACE: cliques, arrastar-e-soltar, abas,
 *  atualização da tela, terminal de logs, histórico, exportação de CSV, etc.
 *
 *  Ele NÃO sabe como calcular um pHash ou ler um .xlsx — para isso, ele
 *  chama as funções prontas do `imageProcessor.js` (nosso "backend" local).
 *  Essa separação deixa o código muito mais fácil de manter: se um dia a
 *  lógica de comparação mudar, só mexemos em `imageProcessor.js`; se a
 *  interface mudar, só mexemos aqui.
 * ==============================================================================
 */

import {
  carregarImagem,
  gerarMiniatura,
  obterDimensoes,
  obterPesoKb,
  calcularPHash,
  calcularDistanciaHamming,
  calcularSimilaridade,
  calcularNitidez,
  avaliarNitidez,
  lerArquivoExcel,
  extrairImagensDaAba,
  obterValorCelula,
  converterValorDaCelula,
  classificarVencimento,
  formatarDataBr,
  indiceParaLetraColuna,
  verificarLinkImagem,
  descreverResultadoLink,
} from "./imageProcessor.js";

(function () {
  "use strict";

  // ============================================================
  //  ESTADO GLOBAL DA APLICAÇÃO
  // ============================================================
  // Guardamos aqui tudo que a aplicação precisa "lembrar" enquanto está
  // aberta: os arquivos selecionados, se está processando, os resultados
  // encontrados, etc. Centralizar isso em um único objeto facilita entender
  // "o que a aplicação sabe" a qualquer momento.
  const CHAVE_HISTORICO = "bradesco_modal_card_history_v3";

  // Tempo máximo de espera ao verificar cada "Link do Adobe" antes de
  // considerá-lo inválido — evita que um link fora do ar trave a execução.
  const TIMEOUT_VERIFICACAO_LINK_ADOBE_MS = 6000;

  const estado = {
    arquivoExcel: null,
    arquivoImagem: null,
    processando: false,
    interrompido: false,
    logs: [],
    matches: [],
    falhas: [],
    infoImagemPesquisa: null,
    idExecucao: null,
    ultimaExecucao: null,
    inicioExecucao: null,
    // Qual ordenação está ativa no preview: "similaridade" (idênticas
    // primeiro) ou "vencimento" (mais próximas de vencer primeiro).
    filtroPreview: "similaridade",
  };

  // ============================================================
  //  PEQUENOS UTILITÁRIOS DE APOIO
  // ============================================================
  const $ = (id) => document.getElementById(id);
  const on = (elemento, evento, fn) => elemento && elemento.addEventListener(evento, fn);
  const horaAtual = () => new Date().toTimeString().slice(0, 8);
  const gerarId = () =>
    crypto.randomUUID ? crypto.randomUUID() : Date.now() + "-" + Math.random().toString(16).slice(2);
  const cederControleAoNavegador = () => new Promise((resolve) => setTimeout(resolve, 0));

  function exibirToast(mensagem, tipo = "info") {
    const container = $("toast-container");
    const toast = document.createElement("div");
    toast.className = "toast" + (tipo ? " toast-" + tipo : "");
    toast.textContent = mensagem;
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 3500);
  }

  function escaparHtml(texto) {
    if (texto == null) return "";
    return String(texto)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#039;");
  }

  function formatarDataHora(iso) {
    if (!iso) return "-";
    try { return new Date(iso).toLocaleString("pt-BR"); } catch { return iso; }
  }

  // ============================================================
  //  ABAS (Execução / Histórico)
  // ============================================================
  document.querySelectorAll(".tab").forEach((botaoAba) => {
    on(botaoAba, "click", () => {
      document.querySelectorAll(".tab").forEach((b) => b.classList.toggle("active", b === botaoAba));
      const alvo = botaoAba.dataset.tab;
      $("panel-run").classList.toggle("hidden", alvo !== "run");
      $("panel-history").classList.toggle("hidden", alvo !== "history");
      if (alvo === "history") renderizarHistorico(carregarHistorico());
    });
  });

  // ============================================================
  //  ÁREAS DE ARRASTAR-E-SOLTAR (DROPZONES)
  // ============================================================
  function configurarDropzone(idZona, idInput, idIcone, idNome, idLimpar, idBotaoSelecionar, aoSelecionarArquivo) {
    const zona = $(idZona), input = $(idInput), icone = $(idIcone),
          nomeEl = $(idNome), botaoLimpar = $(idLimpar), botaoSelecionar = $(idBotaoSelecionar);

    const definirArquivo = (arquivo) => {
      aoSelecionarArquivo(arquivo);
      if (arquivo) {
        nomeEl.textContent = arquivo.name;
        nomeEl.classList.add("filled");
        icone.classList.add("filled");
        botaoLimpar.classList.remove("hidden");
      } else {
        nomeEl.textContent = "Arraste o arquivo aqui";
        nomeEl.classList.remove("filled");
        icone.classList.remove("filled");
        botaoLimpar.classList.add("hidden");
        input.value = "";
      }
      atualizarBotaoExecutar();
    };

    // Três formas de escolher o arquivo: clicar em qualquer ponto da
    // dropzone, clicar no botão clássico "Selecionar arquivo" (útil para
    // quem não pode/consegue arrastar), ou arrastar-e-soltar.
    on(zona, "click", (e) => { if (!e.target.closest(".dz-clear") && !e.target.closest(".dz-select-btn")) input.click(); });
    on(botaoSelecionar, "click", (e) => { e.stopPropagation(); input.click(); });
    on(input, "change", (e) => { if (e.target.files?.[0]) definirArquivo(e.target.files[0]); });
    on(botaoLimpar, "click", (e) => { e.stopPropagation(); definirArquivo(null); });
    on(zona, "dragover", (e) => { e.preventDefault(); zona.classList.add("dragging"); });
    on(zona, "dragleave", () => zona.classList.remove("dragging"));
    on(zona, "drop", (e) => {
      e.preventDefault();
      zona.classList.remove("dragging");
      if (e.dataTransfer.files?.[0]) definirArquivo(e.dataTransfer.files[0]);
    });

    return definirArquivo;
  }

  configurarDropzone("dz-xlsx", "input-xlsx", "dz-xlsx-icon", "dz-xlsx-name", "dz-xlsx-clear", "dz-xlsx-btn",
    (arquivo) => (estado.arquivoExcel = arquivo));
  configurarDropzone("dz-image", "input-image", "dz-image-icon", "dz-image-name", "dz-image-clear", "dz-image-btn",
    (arquivo) => (estado.arquivoImagem = arquivo));

  function atualizarBotaoExecutar() {
    $("btn-run").disabled = !(estado.arquivoExcel && estado.arquivoImagem) || estado.processando;
    const botaoVencimento = $("btn-run-vencimento");
    if (botaoVencimento) botaoVencimento.disabled = !estado.arquivoExcel || estado.processando;
  }

  // ============================================================
  //  CAMPOS ADICIONAIS (colunas extras mapeadas, configuradas antes da busca)
  // ============================================================
  // Além de Descrição e Vencimento (que já eram fixos), o usuário pode
  // configurar aqui outras colunas da planilha (ex: "C" -> "Categoria")
  // para que sejam lidas automaticamente em cada match encontrado.
  let contadorCampoExtra = 0;

  /** Cria (e insere na lista) uma linha de configuração "Coluna + Rótulo". */
  function criarLinhaCampoExtraParametro(coluna = "", rotulo = "") {
    const id = `ef-${++contadorCampoExtra}`;
    const linha = document.createElement("div");
    linha.className = "extra-field-row";
    linha.dataset.id = id;
    linha.innerHTML = `
      <input type="text" class="ef-col" placeholder="Coluna (ex: C)" maxlength="3" value="${escaparHtml(coluna)}" data-testid="${id}-col" />
      <input type="text" class="ef-label" placeholder="Rótulo (ex: Categoria)" value="${escaparHtml(rotulo)}" data-testid="${id}-label" />
      <button type="button" class="ef-remove" title="Remover campo" data-testid="${id}-remove">×</button>
    `;
    on(linha.querySelector(".ef-remove"), "click", () => linha.remove());
    $("extra-fields-list").appendChild(linha);
  }
  on($("btn-add-extra-field-param"), "click", () => criarLinhaCampoExtraParametro());

  /** Lê a configuração de campos extras montada no painel de Parâmetros. */
  function lerCamposExtrasConfigurados() {
    return Array.from(document.querySelectorAll("#extra-fields-list .extra-field-row"))
      .map((linha) => ({
        coluna: linha.querySelector(".ef-col").value.trim().toUpperCase(),
        rotulo: linha.querySelector(".ef-label").value.trim(),
      }))
      .filter((campo) => campo.coluna && campo.rotulo);
  }

  // ============================================================
  //  VALIDAÇÃO DO "LINK DO ADOBE" (coluna opcional)
  // ============================================================
  // Se o usuário configurar a "Coluna do Link do Adobe", cada linha lida
  // da planilha tem seu link verificado (a imagem carrega com sucesso?).
  // Toda a lógica de "como verificar" mora em imageProcessor.js — aqui a
  // gente só lê a célula da linha certa e decide o que fazer com o
  // resultado (isso é o que mantém essa função fácil de manter).
  async function verificarLinkAdobeDaLinha(aba, linha, parametros) {
    if (!parametros.colunaLinkAdobe) return null; // campo não configurado -> não verifica nada

    const referenciaLink = parametros.colunaLinkAdobe + linha;
    const celulaLink = obterValorCelula(aba, referenciaLink);
    const link = celulaLink?.v != null ? String(celulaLink.v).trim() : "";

    const resultado = await verificarLinkImagem(link, TIMEOUT_VERIFICACAO_LINK_ADOBE_MS);
    return {
      link: link || "-",
      valido: resultado.valido,
      statusTexto: descreverResultadoLink(resultado),
    };
  }

  // ============================================================
  //  PARÂMETROS AVANÇADOS (collapsible) + MODAIS
  // ============================================================
  on($("btn-advanced-toggle"), "click", () => {
    $("advanced-panel").classList.toggle("hidden");
    $("advanced-chevron").classList.toggle("rotated");
  });
  document.querySelectorAll("[data-close]").forEach((elemento) =>
    on(elemento, "click", (e) => {
      if (e.target === elemento || elemento.classList.contains("modal-close")) {
        elemento.closest(".modal").classList.add("hidden");
      }
    })
  );

  /** Lê os parâmetros configurados pelo usuário no painel "Parâmetros". */
  function lerParametros() {
    return {
      nomeDaAba: $("p-sheet").value.trim(),
      colunaDescricao: $("p-desc-col").value.trim().toUpperCase(),
      colunaVencimento: $("p-exp-col").value.trim().toUpperCase(),
      colunaLinkAdobe: $("p-link-col").value.trim().toUpperCase(),
      similaridadeMinima: parseFloat($("p-min-sim").value) || 90,
      diasDeAlerta: $("p-alert-days").value.split(",").map((s) => parseInt(s.trim(), 10)).filter((n) => !isNaN(n)),
      limiarNitidez: parseFloat($("p-sharp").value) || 500,
      camposExtras: lerCamposExtrasConfigurados(),
      nomeArquivoExcel: estado.arquivoExcel?.name,
      nomeArquivoImagem: estado.arquivoImagem?.name,
    };
  }

  // ============================================================
  //  TERMINAL (log em tempo real)
  // ============================================================
  const corpoTerminal = $("terminal-output");
  const contadorLinhas = $("term-lines-count");
  const statusTerminal = $("term-status");
  const textoStatusTerminal = $("term-status-text");
  const painelTerminal = $("terminal-wrap");
  const botaoAlternarTerminal = $("btn-terminal-toggle");

  /** Expande ou recolhe o corpo do terminal (os logs técnicos). */
  function alternarTerminal(forcarExpandido) {
    const estaExpandido = botaoAlternarTerminal.getAttribute("aria-expanded") === "true";
    const proximoEstado = forcarExpandido != null ? forcarExpandido : !estaExpandido;
    botaoAlternarTerminal.setAttribute("aria-expanded", String(proximoEstado));
    painelTerminal.classList.toggle("collapsed", !proximoEstado);
    corpoTerminal.classList.toggle("hidden", !proximoEstado);
  }
  on(botaoAlternarTerminal, "click", () => alternarTerminal());

  function adicionarLinhaLog(registro) {
    const cursorOcioso = $("idle-cursor");
    if (cursorOcioso) cursorOcioso.remove();

    const linha = document.createElement("div");
    linha.className = "log-line log-" + (registro.nivel || "info");
    if (registro.hora) {
      const spanHora = document.createElement("span");
      spanHora.className = "log-ts";
      spanHora.textContent = `[${registro.hora}]`;
      linha.appendChild(spanHora);
    }
    linha.appendChild(document.createTextNode(registro.mensagem ?? ""));
    corpoTerminal.appendChild(linha);
    corpoTerminal.scrollTop = corpoTerminal.scrollHeight;

    estado.logs.push(registro);
    contadorLinhas.textContent = estado.logs.length;
  }
  const registrarLog = (nivel, mensagem) => adicionarLinhaLog({ nivel, mensagem, hora: horaAtual() });

  function limparTerminal() {
    estado.logs = [];
    estado.matches = [];
    estado.falhas = [];
    estado.infoImagemPesquisa = null;
    estado.idExecucao = null;
    estado.ultimaExecucao = null;
    contadorLinhas.textContent = "0";
    corpoTerminal.innerHTML = `
      <div class="log-line log-dim">$ Bradesco Modal/Card Console v3.0 (100% offline)</div>
      <div class="log-line log-dim">$ Aguardando arquivos... selecione a planilha e a imagem, depois clique em EXECUTAR.</div>
      <div class="log-line" id="idle-cursor"><span class="log-success">➜</span> <span class="log-info">console</span><span class="cursor-blink"></span></div>
    `;
    $("summary-card").classList.add("hidden");
    $("matches-preview").classList.add("hidden");
    $("mp-grid").innerHTML = "";
    $("mp-count-badge").textContent = "0";
    $("preview-empty").classList.remove("hidden");
    alternarTerminal(false); // volta a esconder os logs por padrão
  }
  on($("btn-clear"), "click", () => { if (!estado.processando) limparTerminal(); });

  /** Rótulos amigáveis para cada fase do processamento — usados tanto no
   *  terminal quanto no indicador de status inline (sempre visível). */
  const ROTULOS_DE_FASE = {
    iniciando: "Iniciando...",
    "extraindo imagens": "Extraindo imagens da planilha...",
    comparando: "Comparando com a imagem de pesquisa...",
    "buscando por vencimento": "Lendo vencimentos da planilha...",
  };

  function definirProcessando(processando, textoStatus) {
    estado.processando = processando;
    $("btn-run").classList.toggle("hidden", processando);
    $("btn-run-vencimento").classList.toggle("hidden", processando);
    $("btn-stop").classList.toggle("hidden", !processando);
    statusTerminal.classList.toggle("hidden", !processando);
    $("run-status-line").classList.toggle("hidden", !processando);
    if (textoStatus) {
      textoStatusTerminal.textContent = textoStatus;
      $("run-status-line-text").textContent = ROTULOS_DE_FASE[textoStatus] || (textoStatus + "...");
    }
    atualizarBotaoExecutar();
  }

  // ============================================================
  //  FLUXO PRINCIPAL — executa a busca de ponta a ponta
  // ============================================================
  async function executarProcesso() {
    if (!estado.arquivoExcel) return exibirToast("Selecione o arquivo Excel", "error");
    if (!estado.arquivoImagem) return exibirToast("Selecione a imagem de pesquisa", "error");

    estado.interrompido = false;
    estado.logs = [];
    estado.matches = [];
    estado.falhas = [];
    estado.infoImagemPesquisa = null;
    estado.idExecucao = gerarId();
    corpoTerminal.innerHTML = "";
    contadorLinhas.textContent = "0";
    $("summary-card").classList.add("hidden");
    definirProcessando(true, "iniciando");
    alternarTerminal(true); // abre o terminal automaticamente durante a execução

    const parametros = lerParametros();

    // Reseta a área de preview de matches
    $("mp-grid").innerHTML = "";
    $("matches-preview").classList.add("hidden");
    $("mp-count-badge").textContent = "0";
    $("preview-empty").classList.add("hidden");

    try {
      // ---- Passo 1: informações da imagem de pesquisa ----
      registrarLog("header", "=".repeat(60));
      registrarLog("header", `INFORMAÇÕES DA IMAGEM — ${parametros.nomeArquivoImagem}`);
      registrarLog("header", "=".repeat(60));

      const dimensoes = await obterDimensoes(estado.arquivoImagem);
      const pesoKb = obterPesoKb(estado.arquivoImagem);
      const scoreNitidez = await calcularNitidez(estado.arquivoImagem);
      const nitidezBoa = avaliarNitidez(scoreNitidez, parametros.limiarNitidez);
      const legendaNitidez = nitidezBoa
        ? "Boa nitidez"
        : "Nitidez baixa (imagem pode estar borrada ou em baixa resolução)";

      registrarLog("info", `Dimensões : ${dimensoes.largura}x${dimensoes.altura} px`);
      registrarLog("info", `Peso      : ${pesoKb} KB`);
      registrarLog(nitidezBoa ? "success" : "warn", `Nitidez   : ${legendaNitidez} (score: ${Math.round(scoreNitidez * 10) / 10})`);

      estado.infoImagemPesquisa = {
        nomeArquivo: parametros.nomeArquivoImagem,
        largura: dimensoes.largura,
        altura: dimensoes.altura,
        pesoKb,
        scoreNitidez: Math.round(scoreNitidez * 10) / 10,
        nitidezBoa,
      };
      registrarLog("header", "=".repeat(60));
      await cederControleAoNavegador();
      if (estado.interrompido) return finalizarPorInterrupcao();

      // ---- Passo 2: abrir a planilha ----
      registrarLog("info", `Abrindo planilha: ${parametros.nomeArquivoExcel}`);
      const { planilha, zip } = await lerArquivoExcel(estado.arquivoExcel);
      if (!planilha.SheetNames.includes(parametros.nomeDaAba)) {
        registrarLog("error", `Aba '${parametros.nomeDaAba}' não encontrada. Abas disponíveis: ${planilha.SheetNames.join(", ")}`);
        return finalizarExecucao(false, `Aba ${parametros.nomeDaAba} não encontrada`);
      }
      const aba = planilha.Sheets[parametros.nomeDaAba];

      // ---- Passo 3: extrair imagens da aba ----
      definirProcessando(true, "extraindo imagens");
      let imagensDaPlanilha;
      try {
        imagensDaPlanilha = await extrairImagensDaAba(zip, parametros.nomeDaAba);
      } catch (erro) {
        registrarLog("error", erro.message);
        return finalizarExecucao(false, erro.message);
      }
      if (imagensDaPlanilha.length === 0) {
        registrarLog("error", "Nenhuma imagem encontrada na aba.");
        return finalizarExecucao(true, null, parametros);
      }

      registrarLog("info", `Imagens encontradas na planilha: ${imagensDaPlanilha.length}`);
      registrarLog("info", "Comparando com a imagem de pesquisa...");
      registrarLog("dim", "");
      await cederControleAoNavegador();

      // ---- Passo 4: calcula o hash da imagem de pesquisa ----
      const hashDePesquisa = await calcularPHash(estado.arquivoImagem);

      // ---- Passo 5: compara com cada imagem da planilha ----
      definirProcessando(true, "comparando");
      for (let i = 0; i < imagensDaPlanilha.length; i++) {
        if (estado.interrompido) return finalizarPorInterrupcao();

        const { linha, coluna, caminhoDoArquivo } = imagensDaPlanilha[i];
        const arquivoDeMidia = zip.file(caminhoDoArquivo);
        if (!arquivoDeMidia) {
          const mensagem = `Imagem #${i + 1} (linha ${linha}): mídia '${caminhoDoArquivo}' não encontrada`;
          estado.falhas.push(mensagem);
          registrarLog("warn", mensagem);
          continue;
        }

        let hashDaImagem;
        try {
          const blobDaImagem = await arquivoDeMidia.async("blob");
          hashDaImagem = await calcularPHash(blobDaImagem);
        } catch (erro) {
          const mensagem = `Imagem #${i + 1} (linha ${linha}): falha ao decodificar (${erro.message})`;
          estado.falhas.push(mensagem);
          registrarLog("warn", mensagem);
          continue;
        }

        const distancia = calcularDistanciaHamming(hashDePesquisa, hashDaImagem);
        const similaridade = calcularSimilaridade(distancia);

        // Não é um match suficiente -> ignora e segue para a próxima imagem
        if (distancia !== 0 && similaridade < parametros.similaridadeMinima) {
          if ((i + 1) % 50 === 0) await cederControleAoNavegador(); // evita travar a UI
          continue;
        }

        // Acima de 95% (ou distância 0) consideramos "idêntica"
        const ehIdentica = distancia === 0 || similaridade >= 95;
        const tipo = ehIdentica ? "identica" : "semelhante";

        const referenciaDescricao = parametros.colunaDescricao + linha;
        const referenciaVencimento = parametros.colunaVencimento + linha;
        const celulaDescricao = obterValorCelula(aba, referenciaDescricao);
        const celulaVencimento = obterValorCelula(aba, referenciaVencimento);

        const descricao = celulaDescricao?.v != null ? String(celulaDescricao.v) : "-";
        const valorVencimento = celulaVencimento ? converterValorDaCelula(celulaVencimento) : null;
        const statusVencimento = classificarVencimento(
          valorVencimento instanceof Date ? valorVencimento : null,
          parametros.diasDeAlerta
        );
        const vencimentoFormatado = valorVencimento instanceof Date
          ? formatarDataBr(valorVencimento)
          : (valorVencimento != null ? String(valorVencimento) : "-");
        const celula = `${indiceParaLetraColuna(coluna)}${linha}`;

        // Campos adicionais configurados pelo usuário (colunas extras além
        // de Descrição/Vencimento) — lidos na mesma linha da imagem.
        const camposExtras = (parametros.camposExtras || []).map((campoConfigurado) => {
          const referenciaExtra = campoConfigurado.coluna + linha;
          const celulaExtra = obterValorCelula(aba, referenciaExtra);
          const valorConvertidoExtra = celulaExtra ? converterValorDaCelula(celulaExtra) : null;
          const valorTextoExtra = valorConvertidoExtra instanceof Date
            ? formatarDataBr(valorConvertidoExtra)
            : (valorConvertidoExtra != null ? String(valorConvertidoExtra) : "-");
          return { coluna: campoConfigurado.coluna, rotulo: campoConfigurado.rotulo, valor: valorTextoExtra };
        });

        // Link do Adobe (opcional): só verifica se a coluna foi configurada.
        const linkAdobe = await verificarLinkAdobeDaLinha(aba, linha, parametros);

        // Gera miniaturas: uma "ao vivo" (blob URL) e outra em base64 (pro histórico)
        let urlMiniatura = null;
        let dadosMiniatura = null;
        try {
          const blobDaImagem = await arquivoDeMidia.async("blob");
          urlMiniatura = URL.createObjectURL(blobDaImagem);
          dadosMiniatura = await gerarMiniatura(blobDaImagem, 200);
        } catch { /* preview é "nice to have" — se falhar, seguimos sem ele */ }

        const match = {
          celula, linha, distancia, similaridade, descricao,
          vencimento: vencimentoFormatado, status_vencimento: statusVencimento, tipo,
          thumbnailUrl: urlMiniatura, thumbnailData: dadosMiniatura,
          camposExtras,
          linkAdobe: linkAdobe?.link ?? null,
          linkAdobeValido: linkAdobe?.valido ?? null,
          statusLinkAdobe: linkAdobe?.statusTexto ?? null,
        };
        estado.matches.push(match);
        renderizarCardDeMatch(match);

        const rotulo = distancia === 0
          ? "IDÊNTICA"
          : ehIdentica ? `IDÊNTICA (${similaridade}%)` : `SEMELHANTE (${similaridade}%)`;
        registrarLog(ehIdentica ? "success" : "warn", `>> ${rotulo} encontrada`);
        registrarLog("info", `   Célula da imagem : ${celula}`);
        registrarLog("info", `   Descrição        : ${descricao}`);
        registrarLog("info", `   Vencimento       : ${vencimentoFormatado}`);
        registrarLog("info", `   Status           : ${statusVencimento}`);
        if (linkAdobe) {
          registrarLog(linkAdobe.valido ? "success" : "warn", `   Link Adobe       : ${linkAdobe.statusTexto} (${linkAdobe.link})`);
        }
        registrarLog("dim", "-".repeat(60));
        await cederControleAoNavegador();
      }

      // ---- Passo 6: relatório final ----
      registrarLog("header", "=".repeat(60));
      registrarLog("header", "RESULTADO DA BUSCA");
      registrarLog("header", "=".repeat(60));
      if (estado.matches.length === 0) {
        registrarLog("warn", "Nenhuma ocorrência da imagem foi encontrada na planilha.");
      } else {
        registrarLog("success", `${estado.matches.length} ocorrência(s) encontrada(s).`);
      }
      if (estado.falhas.length > 0) {
        registrarLog("warn", `${estado.falhas.length} imagem(ns) não puderam ser processadas:`);
        estado.falhas.forEach((mensagem) => registrarLog("warn", `   - ${mensagem}`));
      }
      registrarLog("header", "=".repeat(60));

      finalizarExecucao(true, null, parametros);
    } catch (erro) {
      console.error(erro);
      registrarLog("error", `Erro fatal: ${erro.message}`);
      finalizarExecucao(false, erro.message);
    } finally {
      definirProcessando(false);
    }
  }

  function finalizarPorInterrupcao() {
    registrarLog("warn", "Execução interrompida pelo usuário.");
    finalizarExecucao(false, "interrompida");
    definirProcessando(false);
  }

  function finalizarExecucao(sucesso, erro, parametros) {
    const execucao = {
      id: estado.idExecucao,
      started_at: estado.inicioExecucao || new Date().toISOString(),
      finished_at: new Date().toISOString(),
      params: parametros || lerParametros(),
      matches: estado.matches,
      failures: estado.falhas,
      image_meta: estado.infoImagemPesquisa,
      success: sucesso,
      error: erro,
      modo: parametros?.modo || "comparacao",
    };
    estado.ultimaExecucao = execucao;

    if (sucesso && (estado.matches.length > 0 || estado.falhas.length > 0 || estado.infoImagemPesquisa)) {
      salvarExecucaoNoHistorico(execucao);
      exibirToast("Execução salva no histórico", "success");
    }
    if (!sucesso && erro && erro !== "interrompida") {
      exibirToast("A execução falhou — veja os detalhes no terminal", "error");
    }
    if (sucesso && estado.matches.length > 0) {
      renderizarResumo(estado.matches, execucao.modo);
      aplicarFiltroPreview(); // garante que o preview reflita o filtro ativo
    } else {
      // Nenhum match (ou execução falhou/foi interrompida): volta a mostrar
      // o placeholder no lugar do preview.
      const placeholder = $("preview-empty");
      placeholder.textContent = sucesso
        ? "Nenhuma ocorrência encontrada na planilha."
        : "Nenhum resultado ainda. Selecione a planilha e a imagem de pesquisa e clique em EXECUTAR.";
      placeholder.classList.remove("hidden");
    }
  }

  on($("btn-run"), "click", () => {
    estado.inicioExecucao = new Date().toISOString();
    executarProcesso();
  });
  on($("btn-stop"), "click", () => { estado.interrompido = true; });

  // ============================================================
  //  MODO "VENCIMENTO MAIS PRÓXIMO" — varre toda a planilha, sem precisar
  //  de imagem de pesquisa, e lista todas as linhas ordenadas pela
  //  urgência do vencimento (vencidas e vencendo em breve primeiro).
  // ============================================================
  async function executarBuscaPorVencimento() {
    if (!estado.arquivoExcel) return exibirToast("Selecione o arquivo Excel", "error");

    estado.inicioExecucao = new Date().toISOString();
    estado.interrompido = false;
    estado.logs = [];
    estado.matches = [];
    estado.falhas = [];
    estado.infoImagemPesquisa = null;
    estado.idExecucao = gerarId();
    corpoTerminal.innerHTML = "";
    contadorLinhas.textContent = "0";
    $("summary-card").classList.add("hidden");
    $("mp-grid").innerHTML = "";
    $("matches-preview").classList.add("hidden");
    $("mp-count-badge").textContent = "0";
    $("preview-empty").classList.add("hidden");
    definirProcessando(true, "buscando por vencimento");
    alternarTerminal(true);

    const parametros = { ...lerParametros(), modo: "vencimento" };

    try {
      registrarLog("header", "=".repeat(60));
      registrarLog("header", "MODO: BUSCA POR VENCIMENTO MAIS PRÓXIMO (sem imagem de pesquisa)");
      registrarLog("header", "=".repeat(60));

      registrarLog("info", `Abrindo planilha: ${parametros.nomeArquivoExcel}`);
      const { planilha, zip } = await lerArquivoExcel(estado.arquivoExcel);
      if (!planilha.SheetNames.includes(parametros.nomeDaAba)) {
        registrarLog("error", `Aba '${parametros.nomeDaAba}' não encontrada. Abas disponíveis: ${planilha.SheetNames.join(", ")}`);
        return finalizarExecucao(false, `Aba ${parametros.nomeDaAba} não encontrada`, parametros);
      }
      const aba = planilha.Sheets[parametros.nomeDaAba];

      let imagensDaPlanilha;
      try {
        imagensDaPlanilha = await extrairImagensDaAba(zip, parametros.nomeDaAba);
      } catch (erro) {
        registrarLog("error", erro.message);
        return finalizarExecucao(false, erro.message, parametros);
      }
      if (imagensDaPlanilha.length === 0) {
        registrarLog("error", "Nenhuma imagem encontrada na aba.");
        return finalizarExecucao(true, null, parametros);
      }

      registrarLog("info", `Imagens encontradas na planilha: ${imagensDaPlanilha.length}`);
      registrarLog("info", "Lendo Descrição/Vencimento de cada linha...");
      registrarLog("dim", "");
      await cederControleAoNavegador();

      for (let i = 0; i < imagensDaPlanilha.length; i++) {
        if (estado.interrompido) return finalizarPorInterrupcao();

        const { linha, coluna, caminhoDoArquivo } = imagensDaPlanilha[i];
        const arquivoDeMidia = zip.file(caminhoDoArquivo);
        if (!arquivoDeMidia) {
          const mensagem = `Imagem #${i + 1} (linha ${linha}): mídia '${caminhoDoArquivo}' não encontrada`;
          estado.falhas.push(mensagem);
          registrarLog("warn", mensagem);
          continue;
        }

        const referenciaDescricao = parametros.colunaDescricao + linha;
        const referenciaVencimento = parametros.colunaVencimento + linha;
        const celulaDescricao = obterValorCelula(aba, referenciaDescricao);
        const celulaVencimento = obterValorCelula(aba, referenciaVencimento);

        const descricao = celulaDescricao?.v != null ? String(celulaDescricao.v) : "-";
        const valorVencimento = celulaVencimento ? converterValorDaCelula(celulaVencimento) : null;
        const statusVencimento = classificarVencimento(
          valorVencimento instanceof Date ? valorVencimento : null,
          parametros.diasDeAlerta
        );
        const vencimentoFormatado = valorVencimento instanceof Date
          ? formatarDataBr(valorVencimento)
          : (valorVencimento != null ? String(valorVencimento) : "-");
        const celula = `${indiceParaLetraColuna(coluna)}${linha}`;

        const camposExtras = (parametros.camposExtras || []).map((campoConfigurado) => {
          const referenciaExtra = campoConfigurado.coluna + linha;
          const celulaExtra = obterValorCelula(aba, referenciaExtra);
          const valorConvertidoExtra = celulaExtra ? converterValorDaCelula(celulaExtra) : null;
          const valorTextoExtra = valorConvertidoExtra instanceof Date
            ? formatarDataBr(valorConvertidoExtra)
            : (valorConvertidoExtra != null ? String(valorConvertidoExtra) : "-");
          return { coluna: campoConfigurado.coluna, rotulo: campoConfigurado.rotulo, valor: valorTextoExtra };
        });

        // Link do Adobe (opcional): só verifica se a coluna foi configurada.
        const linkAdobe = await verificarLinkAdobeDaLinha(aba, linha, parametros);

        let urlMiniatura = null;
        let dadosMiniatura = null;
        try {
          const blobDaImagem = await arquivoDeMidia.async("blob");
          urlMiniatura = URL.createObjectURL(blobDaImagem);
          dadosMiniatura = await gerarMiniatura(blobDaImagem, 200);
        } catch { /* preview é "nice to have" */ }

        const match = {
          celula, linha, distancia: null, similaridade: null, descricao,
          vencimento: vencimentoFormatado, status_vencimento: statusVencimento, tipo: "vencimento",
          thumbnailUrl: urlMiniatura, thumbnailData: dadosMiniatura,
          camposExtras,
          linkAdobe: linkAdobe?.link ?? null,
          linkAdobeValido: linkAdobe?.valido ?? null,
          statusLinkAdobe: linkAdobe?.statusTexto ?? null,
        };
        estado.matches.push(match);

        if ((i + 1) % 25 === 0) await cederControleAoNavegador();
      }

      // Ordena tudo por urgência (vencidas e vencendo em breve primeiro)
      estado.matches.sort(ordenarPorVencimento);
      estado.matches.forEach((match) => renderizarCardDeMatch(match));
      estado.filtroPreview = "vencimento";
      document.querySelectorAll(".filter-btn").forEach((b) => b.classList.toggle("active", b.dataset.order === "vencimento"));

      registrarLog("header", "=".repeat(60));
      registrarLog("header", "RESULTADO DA BUSCA POR VENCIMENTO");
      registrarLog("header", "=".repeat(60));
      registrarLog("success", `${estado.matches.length} linha(s) lida(s) e ordenada(s) por vencimento.`);
      if (estado.falhas.length > 0) {
        registrarLog("warn", `${estado.falhas.length} imagem(ns) não puderam ser processadas:`);
        estado.falhas.forEach((mensagem) => registrarLog("warn", `   - ${mensagem}`));
      }
      registrarLog("header", "=".repeat(60));

      finalizarExecucao(true, null, parametros);
    } catch (erro) {
      console.error(erro);
      registrarLog("error", `Erro fatal: ${erro.message}`);
      finalizarExecucao(false, erro.message, parametros);
    } finally {
      definirProcessando(false);
    }
  }

  on($("btn-run-vencimento"), "click", () => executarBuscaPorVencimento());

  // ============================================================
  //  RESUMO (estatísticas da execução)
  // ============================================================
  function renderizarResumo(matches, modo = "comparacao") {
    const rotuloTotal = $("stat-label-total");
    const rotuloIdentical = $("stat-label-identical");
    const rotuloSimilar = $("stat-label-similar");

    if (modo === "vencimento") {
      $("stat-total").textContent = matches.length;
      $("stat-identical").textContent = matches.filter((m) => (m.status_vencimento || "").includes("VENCIDA")).length;
      $("stat-similar").textContent = matches.filter((m) => (m.status_vencimento || "").startsWith("VENCE")).length;
      if (rotuloTotal) rotuloTotal.textContent = "Linhas";
      if (rotuloIdentical) rotuloIdentical.textContent = "Vencidas";
      if (rotuloSimilar) rotuloSimilar.textContent = "Vencendo";
    } else {
      $("stat-total").textContent = matches.length;
      $("stat-identical").textContent = matches.filter((m) => m.tipo === "identica").length;
      $("stat-similar").textContent = matches.filter((m) => m.tipo === "semelhante").length;
      if (rotuloTotal) rotuloTotal.textContent = "Matches";
      if (rotuloIdentical) rotuloIdentical.textContent = "Idênticas";
      if (rotuloSimilar) rotuloSimilar.textContent = "Semelhantes";
    }
    $("summary-card").classList.remove("hidden");
    $("stat-identical").classList.add("stat-green");
  }

  // ============================================================
  //  CARDS DE PREVIEW + LIGHTBOX (ampliar imagem)
  // ============================================================
  function variantesDeStatus(status) {
    if (!status) return "st-green";
    if (status.includes("VENCIDA") || status.includes("VENCE HOJE")) return "st-red";
    if (status.startsWith("VENCE EM")) return "st-amber";
    return "st-green";
  }

  /** Monta (mas não insere na tela) o elemento DOM de um card de match. */
  function criarElementoCardDeMatch(match) {
    const ehIdentica = match.tipo === "identica";
    const ehVencimento = match.tipo === "vencimento";
    const ehManual = match.tipo === "manual";
    const classeTag = ehIdentica ? "tag-identical" : "tag-similar";
    const textoTag = ehVencimento ? "Vencimento" : (ehManual ? "Adicionada" : (ehIdentica ? "Idêntica" : "Semelhante"));
    const classeSimilaridade = match.similaridade >= 95 ? "high" : (match.similaridade >= 85 ? "mid" : "");
    const classeStatus = variantesDeStatus(match.status_vencimento);
    const origemImagem = match.thumbnailUrl || match.thumbnailData || "";
    const mostraSimilaridade = !ehVencimento && !ehManual && match.similaridade != null;

    const card = document.createElement("div");
    card.className = "mp-card";
    card.dataset.testid = `match-card-${match.celula}`;
    card.innerHTML = `
      <div class="mp-thumb" data-lightbox="${escaparHtml(origemImagem)}" data-caption="Célula ${escaparHtml(match.celula)} · ${escaparHtml(match.descricao)}">
        ${origemImagem
          ? `<img src="${origemImagem}" alt="preview ${escaparHtml(match.celula)}" loading="lazy" />`
          : `<span style="color:#999;font-size:12px">sem preview</span>`}
        <span class="mp-thumb-tag ${classeTag}">${textoTag}</span>
        ${mostraSimilaridade ? `<span class="mp-thumb-sim ${classeSimilaridade}">${match.similaridade}%</span>` : ""}
      </div>
      <div class="mp-info">
        <div class="mp-cell">CÉLULA ${escaparHtml(match.celula)} · LINHA ${match.linha}</div>
        <div class="mp-desc">${escaparHtml(match.descricao)}</div>
        <div class="mp-exp">
          <span class="mp-exp-date">${escaparHtml(match.vencimento)}</span>
          <span class="mp-exp-status ${classeStatus}">${escaparHtml(match.status_vencimento)}</span>
        </div>
        ${match.statusLinkAdobe ? `
          <div class="mp-exp">
            <span class="mp-exp-date">Link Adobe</span>
            <span class="mp-exp-status ${match.linkAdobeValido ? "st-green" : "st-red"}">${escaparHtml(match.statusLinkAdobe)}</span>
          </div>
        ` : ""}
        ${match.camposExtras && match.camposExtras.length ? `
          <div class="mp-extra">
            ${match.camposExtras.map((campo) => `<span class="mp-extra-item"><b>${escaparHtml(campo.rotulo)}:</b> ${escaparHtml(campo.valor)}</span>`).join("")}
          </div>
        ` : ""}
      </div>
    `;
    const miniatura = card.querySelector(".mp-thumb");
    on(miniatura, "click", () => abrirLightbox(miniatura.dataset.lightbox, miniatura.dataset.caption));
    return card;
  }

  /** Durante a busca, cada match encontrado é acrescentado ao vivo no grid
   *  (na ordem em que aparece na planilha). Ao final, aplicarFiltroPreview()
   *  reordena tudo de acordo com o filtro selecionado pelo usuário. */
  function renderizarCardDeMatch(match) {
    $("matches-preview").classList.remove("hidden");
    $("preview-empty").classList.add("hidden");
    $("mp-grid").appendChild(criarElementoCardDeMatch(match));
    $("mp-count-badge").textContent = estado.matches.length;
  }

  // ---------- Filtros de ordenação do preview ----------
  /** Idênticas primeiro; dentro de cada grupo, maior similaridade primeiro. */
  function ordenarPorSimilaridade(a, b) {
    const pesoA = a.tipo === "identica" ? 0 : 1;
    const pesoB = b.tipo === "identica" ? 0 : 1;
    if (pesoA !== pesoB) return pesoA - pesoB;
    return b.similaridade - a.similaridade;
  }

  /** Traduz o texto do status de vencimento em um número de "urgência":
   *  quanto menor, mais urgente (vencida < vence hoje < vence em N dias < ok). */
  function obterUrgenciaVencimento(match) {
    const status = match.status_vencimento || "";
    if (status.includes("VENCIDA")) return -1;
    if (status.includes("VENCE HOJE")) return 0;
    const capturado = status.match(/VENCE EM (\d+)/);
    if (capturado) return parseInt(capturado[1], 10);
    return Infinity; // status "OK" (ou desconhecido) vai para o final da lista
  }
  function ordenarPorVencimento(a, b) {
    return obterUrgenciaVencimento(a) - obterUrgenciaVencimento(b);
  }

  /** Limpa e redesenha o grid de preview inteiro, na ordem já definida. */
  function renderizarGradeCompleta(matchesOrdenados) {
    const grade = $("mp-grid");
    grade.innerHTML = "";
    matchesOrdenados.forEach((match) => grade.appendChild(criarElementoCardDeMatch(match)));
  }

  /** Reordena o preview de acordo com o filtro ativo (estado.filtroPreview). */
  function aplicarFiltroPreview() {
    const matches = estado.ultimaExecucao?.matches || estado.matches;
    if (!matches || matches.length === 0) return;
    const funcaoDeOrdenacao = estado.filtroPreview === "vencimento" ? ordenarPorVencimento : ordenarPorSimilaridade;
    renderizarGradeCompleta([...matches].sort(funcaoDeOrdenacao));
  }

  document.querySelectorAll(".filter-btn").forEach((botaoFiltro) => {
    on(botaoFiltro, "click", () => {
      document.querySelectorAll(".filter-btn").forEach((b) => b.classList.toggle("active", b === botaoFiltro));
      estado.filtroPreview = botaoFiltro.dataset.order;
      aplicarFiltroPreview();
    });
  });

  function abrirLightbox(origem, legenda) {
    if (!origem) return;
    $("lightbox-img").src = origem;
    $("lightbox-caption").textContent = legenda || "";
    $("modal-lightbox").classList.remove("hidden");
  }

  // ============================================================
  //  CAMPOS EXTRAS — utilitário compartilhado por CSV e Histórico
  // ============================================================
  /** Devolve a lista (sem repetição, na ordem em que aparecem) de rótulos
   *  de campos extras já presentes nos matches — usada para montar as
   *  colunas extras do CSV exportado e do modal de detalhe do histórico. */
  function obterRotulosExtrasUnicos(matches) {
    const vistos = new Set();
    const rotulos = [];
    (matches || []).forEach((match) => {
      (match.camposExtras || []).forEach((campo) => {
        if (!vistos.has(campo.rotulo)) {
          vistos.add(campo.rotulo);
          rotulos.push(campo.rotulo);
        }
      });
    });
    return rotulos;
  }

  // ============================================================
  //  EXPORTAÇÃO DE CSV (100% no navegador, sem servidor)
  // ============================================================
  function baixarCsv(matches, idExecucao) {
    if (!matches?.length) return exibirToast("Sem matches para exportar", "warn");

    // Campos extras (colunas mapeadas pelo usuário) viram colunas adicionais
    // no fim do CSV, na ordem em que apareceram pela primeira vez.
    const rotulosExtras = obterRotulosExtrasUnicos(matches);

    // A coluna "Link Adobe" só faz sentido se pelo menos um match tiver
    // esse dado preenchido (ou seja, se o campo foi configurado na execução).
    const temLinkAdobe = matches.some((m) => m.statusLinkAdobe != null);

    const cabecalho = [
      "Tipo", "Célula", "Linha", "Similaridade (%)", "Descrição", "Vencimento", "Status Vencimento",
      ...(temLinkAdobe ? ["Link Adobe", "Status do Link"] : []),
      ...rotulosExtras,
    ];
    const linhas = matches.map((m) => [
      m.tipo, m.celula, m.linha, m.similaridade, m.descricao, m.vencimento, m.status_vencimento,
      ...(temLinkAdobe ? [m.linkAdobe ?? "-", m.statusLinkAdobe ?? "-"] : []),
      ...rotulosExtras.map((rotulo) => (m.camposExtras || []).find((c) => c.rotulo === rotulo)?.valor ?? ""),
    ]);

    // "\ufeff" no início é o BOM UTF-8, necessário para o Excel BR abrir
    // acentos corretamente; usamos ";" como separador (padrão do Excel BR).
    const csv = "\ufeff" + [cabecalho, ...linhas]
      .map((linha) => linha.map((celula) => `"${String(celula ?? "").replace(/"/g, '""')}"`).join(";"))
      .join("\r\n");

    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `modal-card-${String(idExecucao).slice(0, 8)}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(link.href), 5000);
  }

  on($("btn-export-csv"), "click", () => {
    if (estado.ultimaExecucao) baixarCsv(estado.ultimaExecucao.matches, estado.ultimaExecucao.id);
    else if (estado.matches.length) baixarCsv(estado.matches, estado.idExecucao);
  });

  // ============================================================
  //  HISTÓRICO (persistido no localStorage do navegador)
  // ============================================================
  function carregarHistorico() {
    try { return JSON.parse(localStorage.getItem(CHAVE_HISTORICO) || "[]"); }
    catch { return []; }
  }

  const MAXIMO_EXECUCOES_NO_HISTORICO = 10;

  function persistirHistorico(lista) {
    const listaLimitada = lista.slice(0, MAXIMO_EXECUCOES_NO_HISTORICO);
    try {
      localStorage.setItem(CHAVE_HISTORICO, JSON.stringify(listaLimitada));
    } catch {
      // O localStorage tem um limite pequeno (geralmente ~5MB) e as
      // miniaturas em base64 são o que mais pesa. Se estourou o limite,
      // tentamos de novo sem as miniaturas em vez de simplesmente perder
      // o histórico inteiro — assim o usuário ainda mantém os dados
      // (célula, descrição, vencimento etc.), só sem o preview de imagem.
      try {
        const listaSemMiniaturas = listaLimitada.map((execucao) => ({
          ...execucao,
          matches: (execucao.matches || []).map((m) => ({ ...m, thumbnailData: null })),
        }));
        localStorage.setItem(CHAVE_HISTORICO, JSON.stringify(listaSemMiniaturas));
        exibirToast("Histórico salvo sem miniaturas (espaço do navegador estava cheio)", "warn");
      } catch {
        exibirToast("Não foi possível salvar no localStorage (espaço insuficiente)", "warn");
      }
    }
  }

  function salvarExecucaoNoHistorico(execucao) {
    // Guardamos só o necessário (a miniatura pequena em base64), descartando
    // as blob URLs, que são temporárias e não sobrevivem a um recarregamento.
    const copiaParaSalvar = {
      ...execucao,
      matches: (execucao.matches || []).map((m) => ({
        celula: m.celula, linha: m.linha, distancia: m.distancia,
        similaridade: m.similaridade, descricao: m.descricao,
        vencimento: m.vencimento, status_vencimento: m.status_vencimento,
        tipo: m.tipo, thumbnailData: m.thumbnailData || null,
        camposExtras: m.camposExtras || [],
        linkAdobe: m.linkAdobe ?? null,
        linkAdobeValido: m.linkAdobeValido ?? null,
        statusLinkAdobe: m.statusLinkAdobe ?? null,
      })),
    };
    const lista = carregarHistorico();
    lista.unshift(copiaParaSalvar);
    persistirHistorico(lista);
  }

  function excluirExecucao(id) {
    persistirHistorico(carregarHistorico().filter((r) => r.id !== id));
    renderizarHistorico(carregarHistorico());
    exibirToast("Execução excluída", "success");
  }

  function limparTodoHistorico() {
    if (!confirm("Excluir todas as execuções?")) return;
    localStorage.removeItem(CHAVE_HISTORICO);
    renderizarHistorico([]);
    exibirToast("Histórico limpo", "success");
  }

  function renderizarHistorico(execucoes) {
    const lista = $("history-list");
    const vazio = $("history-empty");
    lista.innerHTML = "";
    $("history-count").textContent = `${execucoes.length} registro(s)`;
    vazio.classList.toggle("hidden", execucoes.length > 0);

    execucoes.forEach((execucao) => {
      const quantidadeMatches = execucao.matches?.length || 0;
      const quantidadeFalhas = execucao.failures?.length || 0;
      const selo = execucao.success
        ? '<span class="badge badge-success">Concluído</span>'
        : '<span class="badge badge-fail">Falhou</span>';

      const card = document.createElement("div");
      card.className = "history-item";
      card.dataset.testid = `history-item-${execucao.id}`;
      card.innerHTML = `
        <div class="history-item-top">
          <div style="flex:1; min-width:0">
            <div class="hi-meta">
              ${selo}
              <span class="hi-time">${formatarDataHora(execucao.started_at)}</span>
            </div>
            <div class="hi-files">
              <div class="hi-file">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
                <span class="name">${escaparHtml(execucao.params?.xlsx_filename || execucao.params?.nomeArquivoExcel || "-")}</span>
                <span class="dim">· ${escaparHtml(execucao.params?.sheet_name || execucao.params?.nomeDaAba || "")}</span>
              </div>
              <div class="hi-file">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/></svg>
                <span class="name">${escaparHtml(execucao.params?.image_filename || execucao.params?.nomeArquivoImagem || "-")}</span>
              </div>
            </div>
            <div class="hi-stats">
              <span><b>${quantidadeMatches}</b> match(es)</span>
              ${quantidadeFalhas > 0 ? `<span style="color:#b45309"><b>${quantidadeFalhas}</b> falha(s)</span>` : ""}
            </div>
          </div>
          <div class="hi-actions">
            <button class="hi-btn" data-view="${execucao.id}" title="Visualizar">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>
            </button>
            <button class="hi-btn" data-csv="${execucao.id}" title="Exportar CSV">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            </button>
            <button class="hi-btn danger" data-del="${execucao.id}" title="Excluir">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-2 14a2 2 0 0 1-2 2H9a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6M14 11v6"/></svg>
            </button>
          </div>
        </div>
      `;
      lista.appendChild(card);
    });

    lista.querySelectorAll("[data-view]").forEach((botao) =>
      on(botao, "click", () => abrirDetalheExecucao(botao.dataset.view)));
    lista.querySelectorAll("[data-csv]").forEach((botao) =>
      on(botao, "click", () => {
        const execucao = carregarHistorico().find((r) => r.id === botao.dataset.csv);
        if (execucao) baixarCsv(execucao.matches, execucao.id);
      }));
    lista.querySelectorAll("[data-del]").forEach((botao) =>
      on(botao, "click", () => excluirExecucao(botao.dataset.del)));
  }

  function abrirDetalheExecucao(id) {
    const execucao = carregarHistorico().find((r) => r.id === id);
    if (!execucao) return;
    const parametros = execucao.params || {};
    const matches = execucao.matches || [];

    let matchesHtml = "";
    if (matches.length) {
      const rotulosExtras = obterRotulosExtrasUnicos(matches);
      const temLinkAdobe = matches.some((m) => m.statusLinkAdobe != null);
      matchesHtml = `
        <h4 style="margin:16px 0 8px;font-size:13px;font-weight:700">Matches encontrados</h4>
        <div style="overflow-x:auto">
          <table class="detail-table">
            <thead><tr><th></th><th>Tipo</th><th>Célula</th><th>Sim.</th><th>Descrição</th><th>Vencimento</th><th>Status</th>${temLinkAdobe ? `<th>Link Adobe</th>` : ""}${rotulosExtras.map((r) => `<th>${escaparHtml(r)}</th>`).join("")}</tr></thead>
            <tbody>
              ${matches.map((m) => `
                <tr>
                  <td>${m.thumbnailData ? `<img src="${m.thumbnailData}" alt="thumb" style="width:52px;height:40px;object-fit:contain;border-radius:4px;background:#f5f5f5;cursor:zoom-in" data-thumb="${escaparHtml(m.thumbnailData)}" data-caption="Célula ${escaparHtml(m.celula)} · ${escaparHtml(m.descricao)}" />` : ""}</td>
                  <td><span class="badge ${m.tipo === "identica" ? "badge-success" : "badge-warn"}">${escaparHtml(m.tipo)}</span></td>
                  <td class="mono">${escaparHtml(m.celula)}</td>
                  <td>${m.similaridade}%</td>
                  <td>${escaparHtml(m.descricao)}</td>
                  <td>${escaparHtml(m.vencimento)}</td>
                  <td style="color:var(--bradesco-red);font-weight:600">${escaparHtml(m.status_vencimento)}</td>
                  ${temLinkAdobe ? `<td style="color:${m.linkAdobeValido ? "#15803d" : "var(--bradesco-red)"};font-weight:600">${escaparHtml(m.statusLinkAdobe ?? "-")}</td>` : ""}
                  ${rotulosExtras.map((rotulo) => `<td>${escaparHtml((m.camposExtras || []).find((c) => c.rotulo === rotulo)?.valor ?? "-")}</td>`).join("")}
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
      `;
    }

    $("detail-body").innerHTML = `
      <div class="detail-meta">
        <div><b>Início:</b> ${formatarDataHora(execucao.started_at)}</div>
        <div><b>Fim:</b> ${formatarDataHora(execucao.finished_at)}</div>
        <div><b>Planilha:</b> ${escaparHtml(parametros.xlsx_filename || parametros.nomeArquivoExcel || "-")}</div>
        <div><b>Aba:</b> ${escaparHtml(parametros.sheet_name || parametros.nomeDaAba || "-")}</div>
        <div><b>Imagem:</b> ${escaparHtml(parametros.image_filename || parametros.nomeArquivoImagem || "-")}</div>
        <div><b>Similaridade mín:</b> ${parametros.min_similarity ?? parametros.similaridadeMinima ?? "-"}%</div>
      </div>
      ${matchesHtml}
      ${matches.length ? `<button class="btn-primary" style="margin-top:16px" id="detail-csv">Exportar CSV</button>` : ""}
    `;
    $("modal-detail").classList.remove("hidden");

    const botaoCsv = document.getElementById("detail-csv");
    if (botaoCsv) botaoCsv.onclick = () => baixarCsv(matches, execucao.id);

    document.querySelectorAll("#detail-body [data-thumb]").forEach((elemento) => {
      on(elemento, "click", () => abrirLightbox(elemento.dataset.thumb, elemento.dataset.caption));
    });
  }

  on($("btn-refresh-history"), "click", () => renderizarHistorico(carregarHistorico()));
  on($("btn-clear-history"), "click", limparTodoHistorico);

  // ============================================================
  //  INICIALIZAÇÃO
  // ============================================================
  atualizarBotaoExecutar();
  if (typeof window.XLSX === "undefined" || typeof window.JSZip === "undefined") {
    exibirToast("Falha ao carregar bibliotecas — verifique sua conexão", "error");
  }
})();
