const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

// ---------- Brand palette (TIM) ----------
const BLUE = "004691";
const BLUE_DARK = "003A78";
const WHITE = "FFFFFF";
const BG_LIGHT = "F4F6FA";
const RED = "D10A11";
const TEXT_DARK = "0F172A";
const TEXT_MUTED = "475569";
const TEXT_FAINT = "94A3B8";
const CRIT = "DC2626";
const ALTO = "D97706";
const MEDIO = "2563EB";
const BAIXO = "64748B";
const OK = "16A34A";
const BORDER = "E2E8F0";

const FONT = "Calibri";
const FONT_HEAD = "Calibri";

const MOCKUP_DIR = path.join(__dirname, "..", "mockups");
function mockup(name) {
  const p = path.join(MOCKUP_DIR, name);
  return fs.existsSync(p) ? p : null;
}

function newPres() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.3 x 7.5 in
  pres.author = "Planejamento & Controle";
  pres.company = "TIM Brasil";
  pres.title = "Quick Data — Diagnóstico e Proposta";
  return pres;
}

const W = 13.333;
const H = 7.5;
const MARGIN = 0.6;

function footer(slide, pageLabel) {
  slide.addText("Quick Data — Planejamento & Controle", {
    x: MARGIN, y: H - 0.42, w: 6, h: 0.3, fontFace: FONT, fontSize: 9,
    color: TEXT_FAINT, align: "left", margin: 0,
  });
  slide.addText(pageLabel || "", {
    x: W - MARGIN - 3, y: H - 0.42, w: 3, h: 0.3, fontFace: FONT, fontSize: 9,
    color: TEXT_FAINT, align: "right", margin: 0,
  });
}

function timWordmark(slide, opts) {
  const o = Object.assign({ x: MARGIN, y: 0.5, color: WHITE, size: 22 }, opts || {});
  slide.addShape("triangle", {
    x: o.x, y: o.y + 0.03, w: 0.16, h: 0.2, fill: { color: RED }, line: { type: "none" }, rotate: 0,
  });
  slide.addText("TIM", {
    x: o.x + 0.22, y: o.y - 0.06, w: 1.4, h: 0.36, fontFace: FONT_HEAD, fontSize: o.size,
    bold: true, color: o.color, margin: 0, align: "left", valign: "middle",
  });
}

function sectionTitle(slide, kicker, title, opts) {
  const o = Object.assign({ y: 0.55 }, opts || {});
  if (kicker) {
    slide.addText(kicker.toUpperCase(), {
      x: MARGIN, y: o.y, w: W - MARGIN * 2, h: 0.3, fontFace: FONT, fontSize: 12,
      bold: true, color: BLUE, charSpacing: 1, margin: 0,
    });
  }
  slide.addText(title, {
    x: MARGIN, y: o.y + (kicker ? 0.32 : 0), w: W - MARGIN * 2, h: 0.7, fontFace: FONT_HEAD,
    fontSize: 30, bold: true, color: TEXT_DARK, margin: 0,
  });
}

function statCard(slide, x, y, w, h, value, label, color) {
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.08, fill: { color: WHITE }, line: { color: BORDER, width: 1 },
    shadow: { type: "outer", color: "1E293B", opacity: 0.12, blur: 8, offset: 3, angle: 90 },
  });
  slide.addText(value, {
    x: x + 0.15, y: y + 0.18, w: w - 0.3, h: h * 0.55, fontFace: FONT_HEAD, fontSize: 34,
    bold: true, color: color || BLUE, align: "left", margin: 0, valign: "bottom",
  });
  slide.addText(label, {
    x: x + 0.15, y: y + h - 0.55, w: w - 0.3, h: 0.45, fontFace: FONT, fontSize: 12,
    color: TEXT_MUTED, align: "left", margin: 0, valign: "top",
  });
}

// ============================================================
const pres = newPres();

// ---------- S1: Capa ----------
{
  const s = pres.addSlide();
  s.background = { color: BLUE };
  timWordmark(s, { x: MARGIN, y: 0.55, color: WHITE, size: 22 });

  s.addText("Quick Data: da planilha\npara um aplicativo", {
    x: MARGIN, y: 2.5, w: 10.5, h: 1.9, fontFace: FONT_HEAD, fontSize: 44, bold: true,
    color: WHITE, margin: 0, lineSpacing: 48,
  });
  s.addText("Diagnóstico técnico completo do Quick Data 3.23 e proposta de modernização\npara o time de Planejamento & Controle", {
    x: MARGIN, y: 4.35, w: 9.5, h: 0.9, fontFace: FONT, fontSize: 16, color: "CADCFC", margin: 0, lineSpacing: 22,
  });

  s.addShape("line", { x: MARGIN, y: 6.55, w: 3.2, h: 0, line: { color: "3D6BA8", width: 1 } });
  s.addText("[PREENCHER: seu nome]  ·  [PREENCHER: data]", {
    x: MARGIN, y: 6.65, w: 6, h: 0.35, fontFace: FONT, fontSize: 12, color: "9FB8DD", margin: 0,
  });
}

// ---------- S2: O problema atual ----------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  sectionTitle(s, "O contexto", "Por que abrimos essa investigação");

  const items = [
    ["Cresceu por cima de si mesmo", "O Quick Data existe há anos e foi ganhando funções em cima das antigas, sem nunca ter sido documentado de verdade."],
    ["Ninguém sabe tudo o que ele faz", "O conhecimento de como ele funciona por dentro está espalhado entre poucas pessoas — e some quando elas saem do time."],
    ["Lentidão e erros no fechamento", "Processos que deveriam ser rápidos travam ou demoram, e alguns erros se repetem mês após mês sem causa raiz conhecida."],
  ];
  const colW = (W - MARGIN * 2 - 0.6) / 3;
  items.forEach((it, i) => {
    const x = MARGIN + i * (colW + 0.3);
    s.addShape("roundRect", {
      x, y: 1.75, w: colW, h: 3.0, rectRadius: 0.08, fill: { color: BG_LIGHT }, line: { type: "none" },
    });
    s.addText(String(i + 1).padStart(2, "0"), {
      x: x + 0.25, y: 2.0, w: colW - 0.5, h: 0.5, fontFace: FONT_HEAD, fontSize: 22, bold: true, color: BLUE, margin: 0,
    });
    s.addText(it[0], {
      x: x + 0.25, y: 2.55, w: colW - 0.5, h: 0.8, fontFace: FONT_HEAD, fontSize: 16, bold: true, color: TEXT_DARK, margin: 0, lineSpacing: 19,
    });
    s.addText(it[1], {
      x: x + 0.25, y: 3.35, w: colW - 0.5, h: 1.3, fontFace: FONT, fontSize: 12.5, color: TEXT_MUTED, margin: 0, lineSpacing: 17,
    });
  });

  s.addText("Um dado concreto do quanto o arquivo já envelheceu: 42 vínculos internos quebrados, alguns apontando para sistemas que não existem mais desde 2014.", {
    x: MARGIN, y: 5.15, w: W - MARGIN * 2, h: 0.6, fontFace: FONT, fontSize: 13, italic: true, color: TEXT_MUTED, margin: 0,
  });
  footer(s, "2");
}

// ---------- S3: O que fizemos ----------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  sectionTitle(s, "O trabalho realizado", "Uma auditoria técnica completa do arquivo");

  s.addText("Pela primeira vez, sabemos exatamente o que a planilha faz por dentro — cada botão, cada cálculo, cada regra de negócio, com evidência linha a linha.", {
    x: MARGIN, y: 1.75, w: W - MARGIN * 2, h: 0.6, fontFace: FONT, fontSize: 14, color: TEXT_MUTED, margin: 0,
  });

  const stats = [
    ["28", "abas mapeadas"],
    ["192", "rotinas de código documentadas"],
    ["112", "regras de negócio identificadas"],
    ["46", "riscos técnicos classificados"],
  ];
  const cw = (W - MARGIN * 2 - 0.3 * 3) / 4;
  stats.forEach((st, i) => {
    statCard(s, MARGIN + i * (cw + 0.3), 2.6, cw, 1.7, st[0], st[1], BLUE);
  });

  s.addText("Resultado: um documento técnico de referência — o primeiro que este sistema já teve.", {
    x: MARGIN, y: 4.75, w: W - MARGIN * 2, h: 0.5, fontFace: FONT, fontSize: 13, color: TEXT_MUTED, margin: 0,
  });
  footer(s, "3");
}

// ---------- S4 (era S5): Achado estrutural - duplicação de regra ----------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addShape("roundRect", {
    x: MARGIN, y: 0.55, w: 3.3, h: 0.4, rectRadius: 0.2, fill: { color: "FFEDD5" }, line: { type: "none" },
  });
  s.addText("O ACHADO QUE MAIS IMPORTA", {
    x: MARGIN, y: 0.55, w: 3.3, h: 0.4, fontFace: FONT, fontSize: 11, bold: true, color: ALTO,
    align: "center", valign: "middle", margin: 0, charSpacing: 1,
  });

  s.addText("O mesmo cálculo é feito de duas formas diferentes", {
    x: MARGIN, y: 1.15, w: 11.5, h: 0.9, fontFace: FONT_HEAD, fontSize: 28, bold: true, color: TEXT_DARK, margin: 0,
  });

  s.addText("A regra que classifica os números por empresa e por norma contábil (IFRS) foi implementada duas vezes, em pontos diferentes do sistema — e nada garante que as duas continuam produzindo o mesmo resultado.", {
    x: MARGIN, y: 2.15, w: 8.4, h: 1.1, fontFace: FONT, fontSize: 15, color: TEXT_MUTED, margin: 0, lineSpacing: 21,
  });

  // simple flow diagram: two paths converging on "mesmo dado"
  const boxY = 3.6, boxH = 0.85;
  s.addShape("roundRect", { x: MARGIN, y: boxY, w: 4.0, h: boxH, rectRadius: 0.06, fill: { color: BG_LIGHT }, line: { color: BORDER, width: 1 } });
  s.addText("Caminho A\nExtração completa da base", { x: MARGIN + 0.2, y: boxY, w: 3.6, h: boxH, fontFace: FONT, fontSize: 12.5, color: TEXT_DARK, valign: "middle", margin: 0, lineSpacing: 16 });

  s.addShape("roundRect", { x: MARGIN, y: boxY + 1.15, w: 4.0, h: boxH, rectRadius: 0.06, fill: { color: BG_LIGHT }, line: { color: BORDER, width: 1 } });
  s.addText("Caminho B\nAjuste pontual / botão de menu", { x: MARGIN + 0.2, y: boxY + 1.15, w: 3.6, h: boxH, fontFace: FONT, fontSize: 12.5, color: TEXT_DARK, valign: "middle", margin: 0, lineSpacing: 16 });

  s.addShape("roundRect", { x: 5.6, y: boxY + 0.55, w: 3.4, h: boxH, rectRadius: 0.06, fill: { color: "FEF3C7" }, line: { type: "none" } });
  s.addText("Mesmo número financeiro,\ndois resultados possíveis", { x: 5.8, y: boxY + 0.55, w: 3.0, h: boxH, fontFace: FONT, fontSize: 12.5, bold: true, color: "92400E", valign: "middle", margin: 0, lineSpacing: 16 });

  s.addShape("line", { x: 4.05, y: boxY + boxH / 2, w: 1.5, h: 0, line: { color: TEXT_FAINT, width: 1.5, endArrowType: "triangle" } });
  s.addShape("line", { x: 4.05, y: boxY + 1.15 + boxH / 2, w: 1.5, h: 0, line: { color: TEXT_FAINT, width: 1.5, endArrowType: "triangle" } });

  s.addText("Tradução direta: dependendo de qual botão gerou o número, o resultado reportado pode divergir — sem nenhum alerta do sistema.", {
    x: MARGIN, y: 5.85, w: 11.5, h: 0.5, fontFace: FONT, fontSize: 13.5, bold: true, italic: true, color: TEXT_DARK, margin: 0,
  });
  footer(s, "4");
}

// ---------- S6: Panorama de riscos ----------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  sectionTitle(s, "O panorama completo", "46 riscos técnicos identificados e classificados");

  s.addChart(
    pres.ChartType.bar,
    [
      {
        name: "Riscos",
        labels: ["Críticos", "Altos", "Médios", "Baixos"],
        values: [9, 17, 14, 6],
      },
    ],
    {
      x: MARGIN, y: 1.75, w: 7.2, h: 4.4,
      barDir: "col",
      chartColors: [CRIT, ALTO, MEDIO, BAIXO],
      valAxisHidden: false,
      showTitle: false,
      showLegend: false,
      showValue: true,
      dataLabelPosition: "outEnd",
      dataLabelColor: TEXT_DARK,
      dataLabelFontSize: 13,
      dataLabelFontBold: true,
      catAxisLabelColor: TEXT_MUTED,
      catAxisLabelFontSize: 12,
      valAxisLabelColor: TEXT_FAINT,
      valAxisLabelFontSize: 10,
      valGridLine: { color: BORDER, size: 0.75 },
      catGridLine: { style: "none" },
      barGapWidthPct: 40,
    }
  );

  const notes = [
    ["57%", "dos riscos são críticos ou altos — exigem atenção antes de qualquer mudança no sistema"],
    ["0", "riscos críticos são de “lentidão” — todos são de segurança ou confiabilidade do dado"],
  ];
  notes.forEach((n, i) => {
    const y = 1.9 + i * 2.0;
    s.addText(n[0], { x: 8.7, y, w: 3.9, h: 0.6, fontFace: FONT_HEAD, fontSize: 30, bold: true, color: BLUE, margin: 0 });
    s.addText(n[1], { x: 8.7, y: y + 0.6, w: 3.9, h: 1.0, fontFace: FONT, fontSize: 13, color: TEXT_MUTED, margin: 0, lineSpacing: 17 });
  });
  footer(s, "5");
}

// ---------- S7: Por que só consertar não resolve ----------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  sectionTitle(s, "A decisão", "Por que remendar a planilha não resolve");

  const rows = [
    ["", "Planilha atual", "Aplicativo novo"],
    ["Senha do banco protegida", "✗", "✓"],
    ["Mesmo dado sempre com o mesmo resultado", "✗", "✓"],
    ["Progresso visível, sem travar", "✗", "✓"],
    ["Depende de poucas pessoas para manter", "✗", "✓"],
    ["Mudanças testáveis antes de ir pro ar", "✗", "✓"],
  ];

  const tblRows = rows.map((r, ri) => {
    const isHeader = ri === 0;
    return r.map((cell, ci) => {
      let color = TEXT_DARK, bold = isHeader, fill = ri % 2 === 0 ? WHITE : BG_LIGHT;
      if (isHeader) fill = BLUE;
      if (isHeader && ci > 0) color = WHITE;
      if (isHeader && ci === 0) color = WHITE;
      if (!isHeader && ci === 1) color = CRIT;
      if (!isHeader && ci === 2) color = OK;
      if (!isHeader && ci > 0) bold = true;
      return {
        text: cell,
        options: {
          color, bold, fill: { color: fill },
          align: ci === 0 ? "left" : "center",
          fontFace: FONT, fontSize: 13, valign: "middle",
        },
      };
    });
  });

  s.addTable(tblRows, {
    x: MARGIN, y: 1.85, w: W - MARGIN * 2, h: 4.6,
    colW: [6.73, 2.7, 2.7],
    border: { type: "solid", color: BORDER, pt: 0.75 },
    autoPage: false,
    rowH: 0.73,
  });
  footer(s, "6");
}

// ---------- S8: A proposta ----------
{
  const s = pres.addSlide();
  s.background = { color: BLUE };
  timWordmark(s, { x: MARGIN, y: 0.5, color: WHITE, size: 16 });
  s.addText("A PROPOSTA", {
    x: MARGIN, y: 1.2, w: 8, h: 0.3, fontFace: FONT, fontSize: 12, bold: true, color: "9FB8DD", charSpacing: 1, margin: 0,
  });
  s.addText("Um novo Quick Data, como aplicativo desktop", {
    x: MARGIN, y: 1.55, w: 10.8, h: 1.0, fontFace: FONT_HEAD, fontSize: 32, bold: true, color: WHITE, margin: 0,
  });
  s.addText("Continua rodando no computador de cada analista, do jeito que já é hoje — só que fora do Excel, sem macro, sem senha exposta, e sem depender de reabrir a planilha inteira para cada operação.", {
    x: MARGIN, y: 2.65, w: 9.6, h: 0.9, fontFace: FONT, fontSize: 15, color: "CADCFC", margin: 0, lineSpacing: 21,
  });

  const pts = [
    "Mesma tecnologia já usada e validada em outro projeto interno da TIM — não é experimento novo, é o padrão que já funciona.",
    "Ninguém do time precisa aprender Python ou SQL para usar — a experiência é a de abrir um programa e clicar em botões, como hoje.",
    "As telas seguem o mesmo fluxo que vocês já conhecem — extrair, ajustar, consultar resultados — só que mais rápido e confiável.",
  ];
  pts.forEach((t, i) => {
    const y = 3.85 + i * 0.95;
    s.addShape("oval", { x: MARGIN, y: y + 0.03, w: 0.34, h: 0.34, fill: { color: RED }, line: { type: "none" } });
    s.addText(String(i + 1), { x: MARGIN, y: y + 0.03, w: 0.34, h: 0.34, fontFace: FONT_HEAD, fontSize: 14, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    s.addText(t, { x: MARGIN + 0.55, y: y - 0.06, w: 9.4, h: 0.85, fontFace: FONT, fontSize: 13.5, color: WHITE, margin: 0, lineSpacing: 18 });
  });
  footer(s, "7");
}

// ---------- Mockup slides ----------
function mockupSlide(title, kicker, imgFile, caption, opts) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  sectionTitle(s, kicker || "Como vai ficar", title, { y: 0.5 });

  const imgPath = mockup(imgFile);
  const areaX = MARGIN, areaY = 1.55, areaW = W - MARGIN * 2, areaH = 5.15;
  s.addShape("roundRect", {
    x: areaX, y: areaY, w: areaW, h: areaH, rectRadius: 0.05, fill: { color: BG_LIGHT }, line: { color: BORDER, width: 1 },
  });
  if (imgPath) {
    s.addImage({ path: imgPath, x: areaX + 0.15, y: areaY + 0.15, w: areaW - 0.3, h: areaH - 0.3, sizing: { type: "contain", w: areaW - 0.3, h: areaH - 0.3 } });
  } else {
    s.addText("[ IMAGEM PENDENTE: " + imgFile + " ]", {
      x: areaX, y: areaY, w: areaW, h: areaH, align: "center", valign: "middle",
      fontFace: FONT, fontSize: 16, color: TEXT_FAINT, margin: 0,
    });
  }
  if (caption) {
    s.addText(caption, {
      x: MARGIN, y: areaY + areaH + 0.12, w: areaW, h: 0.4, fontFace: FONT, fontSize: 12.5, italic: true, color: TEXT_MUTED, margin: 0,
    });
  }
  return s;
}

{
  const s = mockupSlide("Painel Principal — visão geral ao abrir o app", "Preview do novo Quick Data", "02_painel_principal.png",
    "O app sempre volta pra essa tela ao fim de cada operação — status das bases, atalhos e atividade recente em um só lugar.");
  footer(s, "8");
}
{
  const s = mockupSlide("Extração de Dados — a tela mais usada hoje", null, "03_extracao_de_dados.png",
    "Mesmos botões e mesma lógica da aba “Extração” atual — só que com status real de cada base, não mais um clique “no escuro”.");
  footer(s, "9");
}
{
  const s = mockupSlide("Ajustes Manuais — grade editável com validação", null, "04_ajustes_manuais.png",
    "Mesma função de sempre (lançamentos manuais), com validação de dropdown em tempo real em vez de fórmula escondida na célula.");
  footer(s, "10");
}
{
  const s = mockupSlide("Main Results — o relatório final, sem mudar o hábito", null, "05_main_results.png",
    "Filtros e visão dinâmica equivalentes aos slicers atuais — a leitura do resultado continua a mesma.");
  footer(s, "11");
}

// ---------- Grid slide of remaining screens ----------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  sectionTitle(s, "Preview do novo Quick Data", "E mais: histórico, log e configuração — tudo visível", { y: 0.5 });

  const files = [
    ["06_historico_de_extracoes.png", "Histórico de extrações"],
    ["07_log_console.png", "Log / Console"],
    ["08_configuracoes.png", "Configurações"],
    ["09_manual.png", "Manual"],
    ["10_atualizacoes.png", "Atualizações"],
    ["11_relatar_bug.png", "Relatar bug ou melhoria"],
  ];
  const cols = 3, rows = 2;
  const gap = 0.25;
  const cw = (W - MARGIN * 2 - gap * (cols - 1)) / cols;
  const chH = 1.95;
  files.forEach((f, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = MARGIN + col * (cw + gap);
    const y = 1.6 + row * (chH + 0.55);
    s.addShape("roundRect", { x, y, w: cw, h: chH, rectRadius: 0.05, fill: { color: BG_LIGHT }, line: { color: BORDER, width: 1 } });
    const imgPath = mockup(f[0]);
    if (imgPath) {
      s.addImage({ path: imgPath, x: x + 0.08, y: y + 0.08, w: cw - 0.16, h: chH - 0.16, sizing: { type: "contain", w: cw - 0.16, h: chH - 0.16 } });
    } else {
      s.addText("[ pendente ]", { x, y, w: cw, h: chH, align: "center", valign: "middle", fontFace: FONT, fontSize: 11, color: TEXT_FAINT, margin: 0 });
    }
    s.addText(f[1], { x, y: y + chH + 0.05, w: cw, h: 0.3, fontFace: FONT, fontSize: 11, bold: true, color: TEXT_MUTED, align: "center", margin: 0 });
  });
  footer(s, "12");
}

// ---------- Ganhos esperados ----------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  sectionTitle(s, "O que muda na prática", "Ganhos esperados com a reescrita");

  const gains = [
    ["Segurança", "Nenhuma credencial fica exposta no código — acesso ao banco passa a ser controlado e auditável."],
    ["Confiabilidade do número", "Uma única implementação de cada regra de negócio, testada — elimina o risco de dois cálculos divergentes para o mesmo dado."],
    ["Transparência do processo", "Progresso real e log de execução em cada operação — nunca mais “não sei se travou ou se ainda está rodando”."],
    ["Menos dependência de pessoas", "O conhecimento fica documentado e no código versionado, não só na cabeça de quem construiu."],
  ];
  const colW = (W - MARGIN * 2 - 0.3) / 2;
  gains.forEach((g, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = MARGIN + col * (colW + 0.3);
    const y = 1.8 + row * 2.15;
    s.addShape("roundRect", { x, y, w: colW, h: 1.9, rectRadius: 0.08, fill: { color: BG_LIGHT }, line: { type: "none" } });
    s.addShape("roundRect", { x: x + 0.25, y: y + 0.25, w: 0.5, h: 0.5, rectRadius: 0.08, fill: { color: BLUE }, line: { type: "none" } });
    s.addText(String(i + 1), { x: x + 0.25, y: y + 0.25, w: 0.5, h: 0.5, fontFace: FONT_HEAD, fontSize: 16, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    s.addText(g[0], { x: x + 0.95, y: y + 0.22, w: colW - 1.2, h: 0.4, fontFace: FONT_HEAD, fontSize: 15, bold: true, color: TEXT_DARK, margin: 0 });
    s.addText(g[1], { x: x + 0.25, y: y + 0.85, w: colW - 0.5, h: 0.95, fontFace: FONT, fontSize: 12, color: TEXT_MUTED, margin: 0, lineSpacing: 16 });
  });

  s.addText("Nota: tempo exato de ganho de velocidade será medido durante o período de operação em paralelo (fase de transição), não estamos prometendo um número antes de medir de verdade.", {
    x: MARGIN, y: 6.25, w: W - MARGIN * 2, h: 0.5, fontFace: FONT, fontSize: 10.5, italic: true, color: TEXT_FAINT, margin: 0,
  });
  footer(s, "13");
}

// ---------- Roadmap ----------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  sectionTitle(s, "Como chegamos lá", "Roadmap proposto");

  const phases = [
    ["Fase 0", "Mitigação imediata", "Trocar a senha do banco. Confirmar com o time algumas regras que só vocês sabem responder."],
    ["Fase 1", "Motor de dados", "Reconstruir a lógica de cálculo em Python, validando cada regra contra o resultado da planilha atual, linha a linha."],
    ["Fase 2", "Interface", "Construir as telas (já prototipadas — veja os slides anteriores), reaproveitando o fluxo que vocês já conhecem."],
    ["Fase 3", "Transição", "Rodar em paralelo com a planilha por 1-2 fechamentos, comparando resultado, antes de aposentar o Excel."],
  ];
  const cw = (W - MARGIN * 2 - 0.3 * 3) / 4;
  phases.forEach((p, i) => {
    const x = MARGIN + i * (cw + 0.3);
    s.addShape("roundRect", { x, y: 1.85, w: cw, h: 0.5, rectRadius: 0.25, fill: { color: i === 0 ? RED : BLUE }, line: { type: "none" } });
    s.addText(p[0], { x, y: 1.85, w: cw, h: 0.5, fontFace: FONT_HEAD, fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    s.addText(p[1], { x, y: 2.55, w: cw, h: 0.6, fontFace: FONT_HEAD, fontSize: 14, bold: true, color: TEXT_DARK, margin: 0, lineSpacing: 17 });
    s.addText(p[2], { x, y: 3.2, w: cw, h: 2.2, fontFace: FONT, fontSize: 11.5, color: TEXT_MUTED, margin: 0, lineSpacing: 15 });
    if (i < phases.length - 1) {
      s.addShape("line", { x: x + cw + 0.05, y: 2.1, w: 0.2, h: 0, line: { color: TEXT_FAINT, width: 1.5, endArrowType: "triangle" } });
    }
  });
  footer(s, "14");
}

// ---------- Ask / Closing ----------
{
  const s = pres.addSlide();
  s.background = { color: BLUE };
  timWordmark(s, { x: MARGIN, y: 0.55, color: WHITE, size: 16 });
  s.addText("O que precisamos", {
    x: MARGIN, y: 1.5, w: 10, h: 0.9, fontFace: FONT_HEAD, fontSize: 32, bold: true, color: WHITE, margin: 0,
  });

  const asks = [
    "Aprovar a troca imediata da senha do banco de dados (custo baixo, já é risco hoje).",
    "Aprovar seguir com a reescrita do Quick Data como aplicativo em Python.",
    "Alocar tempo do time técnico sênior para conduzir a Fase 1 (motor de dados).",
    "Uma rodada curta com os analistas para validar as poucas regras de negócio que só o time sabe responder.",
  ];
  asks.forEach((a, i) => {
    const y = 2.7 + i * 0.85;
    s.addShape("oval", { x: MARGIN, y: y + 0.02, w: 0.3, h: 0.3, fill: { color: RED }, line: { type: "none" } });
    s.addText(String(i + 1), { x: MARGIN, y: y + 0.02, w: 0.3, h: 0.3, fontFace: FONT_HEAD, fontSize: 12, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    s.addText(a, { x: MARGIN + 0.5, y: y - 0.08, w: 10.5, h: 0.7, fontFace: FONT, fontSize: 14.5, color: WHITE, margin: 0, lineSpacing: 19 });
  });

  s.addText("Obrigado.", {
    x: MARGIN, y: 6.5, w: 6, h: 0.5, fontFace: FONT_HEAD, fontSize: 16, bold: true, color: "9FB8DD", margin: 0,
  });
  footer(s, "15");
}

// ============================================================
const outPath = path.join(__dirname, "Quick_Data_Apresentacao.pptx");
pres.writeFile({ fileName: outPath }).then(() => {
  console.log("Wrote " + outPath);
});
