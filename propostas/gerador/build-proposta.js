const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const tb = require("react-icons/tb");
const { applyTheme } = require("/root/.claude/skills/synced/ae0912a2-8c90-44b1-85d6-372f2f397ece_1befaab7-63a1-4542-a496-c2e26716f5d9/pptx/scripts/apply_theme.js");

const REPO = "/home/user/Meov---MKT-Pessoal";
const BRAND = path.join(REPO, "brand/meov");
const OUT = process.argv[2] || "proposta.pptx";

const HEX = {
  ink: "202020", white: "FFFFFF", storm: "4F5F66", mist: "ECEEEF",
  pink: "FF19B1", lime: "DBFF00", cyan: "31F8FF", soft: "9AA5AA", line: "D3D8DA",
};

const THEME = {
  name: "Meov",
  headFontFace: "Chillax",
  bodyFontFace: "Chillax",
  colors: {
    dk1: HEX.ink, lt1: HEX.white, dk2: HEX.storm, lt2: HEX.mist,
    accent1: HEX.pink, accent2: HEX.lime, accent3: HEX.cyan, accent4: HEX.storm,
    accent5: HEX.soft, accent6: HEX.line, hlink: HEX.pink, folHlink: HEX.storm,
  },
};

// ---------- assets ----------
async function svgLogo(file, color) {
  let s = fs.readFileSync(path.join(BRAND, file), "utf8");
  s = s.replace(/<rect class="fil0"[^>]*\/>/, "");
  s = s.replace(/\.fil1 \{fill:[^}]*\}/, `.fil1 {fill:#${color}}`)
       .replace(/\.fil2 \{fill:[^;]*;/, `.fil2 {fill:#${color};`)
       .replace(/\.fil0 \{fill:[^}]*\}/, `.fil0 {fill:#${color}}`);
  const buf = await sharp(Buffer.from(s), { density: 300 }).trim().png().toBuffer();
  const meta = await sharp(buf).metadata();
  return { data: "image/png;base64," + buf.toString("base64"), ratio: meta.width / meta.height };
}
async function icon(name, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(tb[name], { color: "#" + color, size, strokeWidth: 1.6 }));
  const buf = await sharp(Buffer.from(svg)).resize(size, size).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
const tex = (f) => path.join(BRAND, "texturas", f);

(async () => {
  const L = {
    lime: await svgLogo("logo/logotipo/meov-logotipo-lima-fundo-grafite.svg", HEX.lime),
    white: await svgLogo("logo/logotipo/meov-logotipo-lima-fundo-grafite.svg", HEX.white),
    ink: await svgLogo("logo/logotipo/meov-logotipo-lima-fundo-grafite.svg", HEX.ink),
  };
  const S = {
    ink: await svgLogo("logo/simbolo/meov-simbolo-lima-fundo-grafite.svg", HEX.ink),
    pink: await svgLogo("logo/simbolo/meov-simbolo-lima-fundo-grafite.svg", HEX.pink),
  };
  const W = {
    pink: await svgLogo("pattern/meov-pattern-onda-preto.svg", HEX.pink),
    lime: await svgLogo("pattern/meov-pattern-onda-preto.svg", HEX.lime),
    line: await svgLogo("pattern/meov-pattern-onda-preto.svg", HEX.line),
  };

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.title = "Proposta Meov · Fernando Lohmann";
  pres.author = "Meov";
  pres.company = "Meov";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  const SW = 13.333, SH = 7.5, MX = 0.6;

  // ---------- layouts ----------
  const footerLight = [
    { image: { x: MX, y: 6.92, w: 0.2 * S.ink.ratio, h: 0.2, data: S.ink.data } },
    { text: { text: "proposta · fernando lohmann", options: { x: MX + 0.2 * S.ink.ratio + 0.15, y: 6.86, w: 5, h: 0.3, fontSize: 10, color: C.text2, margin: 0 } } },
  ];
  pres.defineSlideMaster({
    title: "Meov Claro",
    background: { color: C.background2 },
    objects: [
      ...footerLight,
      { placeholder: { options: { name: "kicker", type: "body", x: MX, y: 0.45, w: 8, h: 0.35, fontSize: 14, bold: true, color: C.accent1, margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.8, w: 12.1, h: 0.9, fontSize: 32, bold: true, color: C.text1, margin: 0, valign: "top", align: "left" }, text: "" } },
    ],
    slideNumber: { x: 12.25, y: 6.86, w: 0.5, h: 0.3, fontSize: 10, color: C.text2, align: "right" },
  });
  pres.defineSlideMaster({
    title: "Meov Escuro",
    background: { color: C.text1 },
    objects: [],
  });

  const names = new Set();
  const on = (o, n) => { o.objectName = n; return o; };
  const T = (slide, text, opts, name) => slide.addText(text, on({ isTextBox: true, margin: 0, ...opts }, name));
  const card = (slide, x, y, w, h, fill, name) => slide.addShape(pres.shapes.ROUNDED_RECTANGLE, on({ x, y, w, h, rectRadius: 0.14, fill: { color: fill }, line: { color: fill, width: 0 } }, name));
  const dot = (slide, x, y, d, fill, name) => slide.addShape(pres.shapes.OVAL, on({ x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 } }, name));
  const ico = (slide, data, x, y, d, name) => slide.addImage(on({ data, x, y, w: d, h: d }, name));
  const iconDot = async (slide, iname, x, y, d, fill, fg, name) => {
    dot(slide, x, y, d, fill, name + " fundo");
    ico(slide, await icon(iname, fg), x + d * 0.22, y + d * 0.22, d * 0.56, name + " ícone");
  };
  const light = (sec) => pres.addSlide({ masterName: "Meov Claro", sectionTitle: sec });
  const dark = (sec) => pres.addSlide({ masterName: "Meov Escuro", sectionTitle: sec });

  // =========================================================
  // 1. Capa
  pres.addSection({ title: "Abertura" });
  {
    const s = dark("Abertura");
    s.background = { path: tex("meov-textura-degrade-01.jpg") };
    s.addImage(on({ data: W.pink.data, x: 6.2, y: 1.2, w: 9.5, h: 9.5 / W.pink.ratio, transparency: 0 }, "Onda rosa"));
    s.addImage(on({ data: L.lime.data, x: MX, y: 0.6, w: 1.9, h: 1.9 / L.lime.ratio }, "Logotipo Meov"));
    T(s, "proposta de", { x: MX, y: 2.5, w: 7, h: 0.5, fontSize: 22, color: C.background1 }, "Sobretítulo");
    T(s, "Marca Pessoal", { x: MX, y: 2.95, w: 8, h: 1.3, fontSize: 66, bold: true, color: C.background1, fontFace: "Chillax" }, "Título");
    T(s, [
      { text: "aos cuidados de", options: { color: C.background2, fontSize: 16, breakLine: true } },
      { text: "Fernando Lohmann", options: { color: C.accent2, fontSize: 26, bold: true } },
    ], { x: MX, y: 4.55, w: 6, h: 1.0 }, "Destinatário");
    T(s, "outubro 2026  ·  ciclo de 6 meses", { x: MX, y: 6.75, w: 6, h: 0.35, fontSize: 12, color: C.background2 }, "Data");
    s.addNotes("Capa. Proposta de marca pessoal para Fernando Lohmann.");
  }

  // 2. O ponto de partida
  {
    const s = light("Abertura");
    s.addText("o ponto de partida", { placeholder: "kicker" });
    s.addText("Você já é referência. Falta o digital mostrar", { placeholder: "title" });
    const rows = [
      { label: "Reputação construída fora das redes", pct: 0.92, color: C.text1, tag: "anos de trabalho, clientes, indicações" },
      { label: "Presença no digital hoje", pct: 0.22, color: C.accent1, tag: "pouco do que você é aparece" },
    ];
    let y = 2.35;
    rows.forEach((r, i) => {
      T(s, r.label, { x: MX, y, w: 7, h: 0.4, fontSize: 18, bold: true, color: C.text1 }, `Barra ${i + 1} rótulo`);
      card(s, MX, y + 0.5, 7.6, 0.55, C.background1, `Barra ${i + 1} trilho`);
      card(s, MX, y + 0.5, 7.6 * r.pct, 0.55, r.color, `Barra ${i + 1} valor`);
      T(s, r.tag, { x: MX, y: y + 1.15, w: 7.6, h: 0.35, fontSize: 14, color: C.text2 }, `Barra ${i + 1} legenda`);
      y += 1.9;
    });
    card(s, 8.85, 2.35, 3.88, 3.75, C.text1, "Quadro distância");
    s.addImage(on({ data: W.lime.data, x: 9.15, y: 2.75, w: 3.3, h: 3.3 / W.lime.ratio }, "Onda lima"));
    T(s, [
      { text: "A distância entre as duas barras", options: { breakLine: true } },
      { text: "é exatamente onde a Meov trabalha.", options: { color: C.accent2, bold: true } },
    ], { x: 9.15, y: 4.2, w: 3.3, h: 1.5, fontSize: 18, color: C.background1, valign: "top" }, "Texto distância");
  }

  // 3. Quem somos
  pres.addSection({ title: "A Meov" });
  {
    const s = light("A Meov");
    s.addText("quem somos", { placeholder: "kicker" });
    s.addText("Um time ao seu lado, da estratégia à câmera", { placeholder: "title" });
    T(s, [
      { text: "Não somos uma agência de marketing.", options: { bold: true, breakLine: true } },
      { text: "Somos um time que apoia técnica e emocionalmente quem decide colocar a própria voz no mundo." },
    ], { x: MX, y: 2.3, w: 4.4, h: 2.4, fontSize: 22, color: C.text1, valign: "top", lineSpacingMultiple: 1.1 }, "Manifesto");
    T(s, "Seja nas redes quem você já é na vida.", { x: MX, y: 5.35, w: 4.4, h: 0.5, fontSize: 16, italic: true, color: C.accent1 }, "Frase");
    const fronts = [
      { n: "Estratégia", d: "O Método Meov: imersão, Brandbook e plano de 6 meses.", i: "TbCompass", c: C.accent2, fg: HEX.ink },
      { n: "Audiovisual", d: "Direção de imagem, foto, vídeo e edição com o nosso fino trato.", i: "TbMovie", c: C.accent1, fg: HEX.white },
      { n: "Gestão das redes", d: "Instagram, LinkedIn e Site/LP: calendário, roteiros, design, publicação e relatórios.", i: "TbLayoutGrid", c: C.accent3, fg: HEX.ink },
    ];
    for (let i = 0; i < fronts.length; i++) {
      const f = fronts[i], y = 2.3 + i * 1.45;
      card(s, 5.6, y, 7.13, 1.25, C.background1, `Frente ${i + 1} cartão`);
      await iconDot(s, f.i, 5.85, y + 0.25, 0.75, f.c, f.fg, `Frente ${i + 1}`);
      T(s, f.n, { x: 6.85, y: y + 0.2, w: 5.6, h: 0.4, fontSize: 18, bold: true, color: C.text1 }, `Frente ${i + 1} nome`);
      T(s, f.d, { x: 6.85, y: y + 0.6, w: 5.7, h: 0.5, fontSize: 14, color: C.text2 }, `Frente ${i + 1} descrição`);
    }
  }

  // 4. Método
  {
    const s = light("A Meov");
    s.addText("o método meov", { placeholder: "kicker" });
    s.addText("Da Raiz à Voz: estratégia antes de estética", { placeholder: "title" });
    const steps = [
      { n: "Raiz", q: "Quem você é e por que quer aparecer", c: HEX.ink, tc: C.background1 },
      { n: "Território", q: "Para quem você fala e onde se posiciona", c: HEX.pink, tc: C.background1 },
      { n: "Voz", q: "Como você fala, aparece e oferece", c: HEX.lime, tc: C.text1 },
      { n: "Movimento", q: "Presença constante por 6 meses", c: HEX.cyan, tc: C.text1 },
    ];
    const xs = [0.95, 4.0, 7.05, 10.1], ys = [2.0, 3.55, 2.0, 3.55], d = 1.25;
    for (let i = 0; i < 3; i++) {
      const x1 = xs[i] + d / 2, y1 = ys[i] + d / 2, x2 = xs[i + 1] + d / 2, y2 = ys[i + 1] + d / 2;
      s.addShape(pres.shapes.LINE, on({ x: x1, y: Math.min(y1, y2), w: x2 - x1, h: Math.abs(y2 - y1), flipV: y2 < y1, line: { color: C.text2, width: 2.5, dashType: "dash" } }, `Conexão ${i + 1}`));
    }
    steps.forEach((st, i) => {
      s.addShape(pres.shapes.OVAL, on({ x: xs[i], y: ys[i], w: d, h: d, fill: { color: st.c }, line: { color: st.c, width: 0 } }, `Movimento ${i + 1}`));
      T(s, String(i + 1).padStart(2, "0"), { x: xs[i], y: ys[i], w: d, h: d, fontSize: 26, bold: true, color: st.tc, align: "center", valign: "middle" }, `Movimento ${i + 1} número`);
      const ty = ys[i] < 3 ? ys[i] + d + 0.15 : ys[i] + d + 0.15;
      T(s, st.n, { x: xs[i] - 0.45, y: ty, w: d + 0.9, h: 0.4, fontSize: 20, bold: true, color: C.text1, align: "center" }, `Movimento ${i + 1} nome`);
      T(s, st.q, { x: xs[i] - 0.55, y: ty + 0.42, w: d + 1.1, h: 0.75, fontSize: 14, color: C.text2, align: "center", valign: "top" }, `Movimento ${i + 1} pergunta`);
    });
    T(s, "Tudo vira o seu Brandbook: o guia da marca e do trabalho dos próximos 6 meses.", { x: MX, y: 6.35, w: 12.1, h: 0.4, fontSize: 14, color: C.text1, bold: true }, "Rodapé método");
  }

  // 5. Jornada 6 meses
  {
    const s = light("A Meov");
    s.addText("sua jornada", { placeholder: "kicker" });
    s.addText("Seis meses do diagnóstico à colheita", { placeholder: "title" });
    const phases = [
      { m: "0", n: "Implantação", d: "Método, ensaio e canais prontos", c: C.text1, t: C.background1, s: C.background2 },
      { m: "1", n: "Fundação", d: "Primeiros conteúdos no ar", c: C.background1, t: C.text1, s: C.text2 },
      { m: "2", n: "Ritmo", d: "Constância e rituais", c: C.background1, t: C.text1, s: C.text2 },
      { m: "3", n: "Primeira leitura", d: "Revisão de rota com dados", c: C.accent2, t: C.text1, s: C.text1, flag: true },
      { m: "4", n: "Autoridade", d: "Casos, provas e novos palcos", c: C.background1, t: C.text1, s: C.text2 },
      { m: "5", n: "Autoridade", d: "Aprofundar e converter", c: C.background1, t: C.text1, s: C.text2 },
      { m: "6", n: "Colheita", d: "Resultado e próximo ciclo", c: C.accent1, t: C.background1, s: C.background1, flag: true },
    ];
    const gap = 0.15, w = (12.13 - gap * 6) / 7;
    phases.forEach((p, i) => {
      const x = MX + i * (w + gap);
      card(s, x, 2.45, w, 3.2, p.c, `Fase ${i} cartão`);
      T(s, i === 0 ? "início" : "mês", { x: x + 0.15, y: 2.65, w: w - 0.3, h: 0.3, fontSize: 12, color: p.s }, `Fase ${i} rótulo`);
      T(s, p.m, { x: x + 0.15, y: 2.9, w: w - 0.3, h: 1.0, fontSize: 54, bold: true, color: p.t }, `Fase ${i} número`);
      T(s, p.n, { x: x + 0.15, y: 4.0, w: w - 0.25, h: 0.65, fontSize: 15, bold: true, color: p.t, valign: "top" }, `Fase ${i} nome`);
      T(s, p.d, { x: x + 0.15, y: 4.65, w: w - 0.25, h: 0.85, fontSize: 12, color: p.s, valign: "top" }, `Fase ${i} descrição`);
    });
    T(s, "Revisões de rota no 3º e no 6º mês, sempre comparando com o seu ponto de partida.", { x: MX, y: 5.95, w: 12.1, h: 0.4, fontSize: 14, color: C.text2 }, "Nota revisões");
  }

  // 6. Implantação
  pres.addSection({ title: "Escopo" });
  {
    const s = light("Escopo");
    s.addText("parte 1 · implantação", { placeholder: "kicker" });
    s.addText("Sua marca construída, seus canais prontos", { placeholder: "title" });
    const items = [
      { n: "Método Meov", i: "TbCompass", c: C.accent2, fg: HEX.ink, d: "Diagnóstico digital e percepção externa · imersão presencial · Brandbook · plano de 6 meses · alinhamento com conselho e compliance" },
      { n: "Ensaio de marca", i: "TbCamera", c: C.accent1, fg: HEX.white, d: "1 sessão com direção de imagem · 12 fotos tratadas" },
      { n: "Instagram", i: "TbBrandInstagram", c: C.accent3, fg: HEX.ink, d: "Conta de criador · bio, link e destaques · 3 posts fixados · templates · duas etapas e Meta Business Suite" },
      { n: "LinkedIn", i: "TbBrandLinkedin", c: C.accent3, fg: HEX.ink, d: "Headline e Sobre · experiências como conquistas · banner e URL · Em destaque · modo criador" },
      { n: "Site/LP", i: "TbWorld", c: C.accent2, fg: HEX.ink, d: "Página única com sua história, como você ajuda, provas, WhatsApp e links para todas as redes" },
      { n: "Também inclui", i: "TbChecks", c: C.text1, fg: HEX.lime, d: "Fluxo de aprovação combinado · autorização de uso de imagem" },
    ];
    const cw = 3.9, ch = 2.15, gx = 0.21, gy = 0.25;
    for (let i = 0; i < items.length; i++) {
      const it = items[i], x = MX + (i % 3) * (cw + gx), y = 2.0 + Math.floor(i / 3) * (ch + gy);
      card(s, x, y, cw, ch, C.background1, `Entrega ${i + 1} cartão`);
      await iconDot(s, it.i, x + 0.25, y + 0.25, 0.62, it.c, it.fg, `Entrega ${i + 1}`);
      T(s, it.n, { x: x + 1.05, y: y + 0.3, w: cw - 1.2, h: 0.5, fontSize: 18, bold: true, color: C.text1, valign: "middle" }, `Entrega ${i + 1} nome`);
      T(s, it.d, { x: x + 0.25, y: y + 1.0, w: cw - 0.45, h: ch - 1.1, fontSize: 13, color: C.text2, valign: "top" }, `Entrega ${i + 1} descrição`);
    }
  }

  // 7. Acompanhamento mensal
  {
    const s = light("Escopo");
    s.addText("parte 2 · acompanhamento mensal", { placeholder: "kicker" });
    s.addText("Presença constante, todo mês, por 6 meses", { placeholder: "title" });
    const stats = [
      { v: "12", l: "reels editados", c: C.text1, t: C.accent2, s: C.background2 },
      { v: "8", l: "carrosséis com texto e design", c: C.background1, t: C.accent1, s: C.text2 },
      { v: "4", l: "posts de texto no LinkedIn", c: C.background1, t: C.text1, s: C.text2 },
      { v: "2", l: "gravações por mês", c: C.background1, t: C.text1, s: C.text2 },
    ];
    const w = 2.88, g = 0.2;
    stats.forEach((st, i) => {
      const x = MX + i * (w + g);
      card(s, x, 2.0, w, 2.55, st.c, `Número ${i + 1} cartão`);
      T(s, st.v, { x: x + 0.3, y: 2.2, w: w - 0.5, h: 1.4, fontSize: 80, bold: true, color: st.t }, `Número ${i + 1} valor`);
      T(s, st.l, { x: x + 0.3, y: 3.7, w: w - 0.5, h: 0.65, fontSize: 16, color: st.s, valign: "top" }, `Número ${i + 1} rótulo`);
    });
    T(s, "mínimos por mês: feriados e semanas difíceis não reduzem o volume", { x: MX, y: 4.7, w: 12, h: 0.35, fontSize: 12, italic: true, color: C.text2 }, "Nota mínimos");
    const chips = ["Pautas e roteiros", "Calendário editorial", "Legendas e publicação", "Adaptação para o LinkedIn", "Relatório e reunião mensal", "Revisões no 3º e 6º mês"];
    let cx = MX, cy = 5.3;
    chips.forEach((c, i) => {
      const w2 = 0.42 + c.length * 0.105;
      if (cx + w2 > SW - MX) { cx = MX; cy += 0.62; }
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, on({ x: cx, y: cy, w: w2, h: 0.48, rectRadius: 0.24, fill: { color: C.background1 }, line: { color: C.text2, width: 0.75 } }, `Inclui ${i + 1} fundo`));
      T(s, c, { x: cx, y: cy, w: w2, h: 0.48, fontSize: 13, color: C.text1, align: "center", valign: "middle" }, `Inclui ${i + 1}`);
      cx += w2 + 0.15;
    });
  }

  // 8. Ciclo do mês
  {
    const s = light("Escopo");
    s.addText("como funciona", { placeholder: "kicker" });
    s.addText("A gente cuida da técnica. Você aparece com verdade", { placeholder: "title" });
    T(s, [
      { text: "Todo mês roda um ciclo completo de produção.", options: { breakLine: true } },
      { text: "Seu tempo vai para o que só você pode fazer: estar diante da câmera e responder sua tribo.", options: { color: C.text2 } },
    ], { x: MX, y: 2.3, w: 4.6, h: 2.4, fontSize: 20, color: C.text1, valign: "top", paraSpaceAfter: 10 }, "Texto ciclo");
    card(s, MX, 5.05, 4.6, 1.2, C.background1, "Quadro você");
    await iconDot(s, "TbMicrophone", MX + 0.25, 5.27, 0.75, C.accent1, HEX.white, "Você");
    T(s, [{ text: "Seu papel", options: { bold: true, breakLine: true, color: C.text1 } }, { text: "gravar e aprovar", options: { color: C.text2 } }], { x: MX + 1.2, y: 5.2, w: 3.2, h: 0.9, fontSize: 15, valign: "middle" }, "Você texto");
    const cx = 8.95, cy = 4.3, r = 1.75;
    s.addShape(pres.shapes.OVAL, on({ x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, fill: { color: C.background2 }, line: { color: C.text2, width: 1.5, dashType: "dash" } }, "Anel do ciclo"));
    s.addImage(on({ data: S.pink.data, x: cx - 0.6, y: cy - 0.6 / S.pink.ratio / 2 - 0.05, w: 1.2, h: 1.2 / S.pink.ratio }, "Símbolo central"));
    const nodes = [
      { n: "Pauta e roteiro", i: "TbPencil", c: C.accent2, fg: HEX.ink },
      { n: "Gravação", i: "TbVideo", c: C.accent1, fg: HEX.white },
      { n: "Edição e design", i: "TbScissors", c: C.accent3, fg: HEX.ink },
      { n: "Publicação", i: "TbSend", c: C.text1, fg: HEX.lime },
      { n: "Relatório e reunião", i: "TbReportAnalytics", c: C.accent2, fg: HEX.ink },
    ];
    const d = 0.85;
    for (let i = 0; i < nodes.length; i++) {
      const a = -Math.PI / 2 + i * (2 * Math.PI / nodes.length);
      const nx = cx + r * Math.cos(a) - d / 2, ny = cy + r * Math.sin(a) - d / 2;
      await iconDot(s, nodes[i].i, nx, ny, d, nodes[i].c, nodes[i].fg, `Etapa ${i + 1}`);
      const right = Math.cos(a) > 0.2, left = Math.cos(a) < -0.2;
      const lw = 1.9;
      const lx = right ? nx + d + 0.12 : left ? nx - lw - 0.12 : nx + d / 2 - lw / 2;
      const ly = right || left ? ny + d / 2 - 0.2 : ny - 0.45;
      T(s, `${i + 1}. ${nodes[i].n}`, { x: lx, y: ly, w: lw, h: 0.4, fontSize: 14, bold: true, color: C.text1, align: right ? "left" : left ? "right" : "center" }, `Etapa ${i + 1} nome`);
    }
  }

  // 9. Boas-vindas
  {
    const s = light("Escopo");
    s.addText("condição especial de boas-vindas", { placeholder: "kicker" });
    s.addText("Nos 3 primeiros meses, gravação semanal", { placeholder: "title" });
    s.addChart(pres.charts.BAR, [
      { name: "Incluído no plano", labels: ["mês 1", "mês 2", "mês 3", "mês 4", "mês 5", "mês 6"], values: [2, 2, 2, 2, 2, 2] },
      { name: "Bônus Meov", labels: ["mês 1", "mês 2", "mês 3", "mês 4", "mês 5", "mês 6"], values: [2, 2, 2, 0, 0, 0] },
    ], on({
      x: MX, y: 1.95, w: 7.6, h: 4.6, barDir: "col", barGrouping: "stacked", barGapWidthPct: 55,
      chartColors: [HEX.storm, HEX.pink],
      showValue: true, dataLabelPosition: "ctr", dataLabelColor: HEX.white, dataLabelFontSize: 14, dataLabelFontFace: "+mn-lt", dataLabelFormatCode: "0;;;",
      showTitle: true, title: "gravações por mês", titleFontSize: 14, titleColor: HEX.ink, titleFontFace: "+mn-lt",
      showLegend: true, legendPos: "b", legendFontSize: 12, legendColor: HEX.storm, legendFontFace: "+mn-lt",
      catAxisLabelColor: HEX.storm, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", catAxisLineShow: false,
      valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      plotArea: { fill: { color: HEX.mist } }, chartArea: { fill: { color: HEX.mist } },
    }, "Gráfico de gravações"));
    card(s, 8.65, 1.95, 4.08, 4.6, C.accent1, "Quadro bônus");
    await iconDot(s, "TbGift", 8.95, 2.25, 0.8, C.background1, HEX.pink, "Presente");
    T(s, "+6", { x: 8.95, y: 3.2, w: 3.5, h: 1.2, fontSize: 80, bold: true, color: C.background1 }, "Bônus valor");
    T(s, "gravações de presente, sem custo adicional", { x: 8.95, y: 4.45, w: 3.5, h: 0.75, fontSize: 18, bold: true, color: C.background1, valign: "top" }, "Bônus rótulo");
    T(s, "Mais material no início, quando o ritmo da sua marca está nascendo.", { x: 8.95, y: 5.3, w: 3.5, h: 0.9, fontSize: 14, color: C.background1, valign: "top" }, "Bônus motivo");
  }

  // 10. Como medimos
  {
    const s = light("Escopo");
    s.addText("como medimos", { placeholder: "kicker" });
    s.addText("Não medimos só seguidores", { placeholder: "title" });
    const layers = [
      { n: "Presença", d: "você está aparecendo com constância?" },
      { n: "Alcance qualificado", d: "as pessoas certas estão te vendo?" },
      { n: "Conversa", d: "elas salvam, compartilham e te chamam?" },
      { n: "Negócio", d: "isso vira reuniões e clientes?" },
      { n: "Confiança", d: "você se sente mais seguro sendo quem é?" },
    ];
    const fills = [C.background1, C.background1, C.accent3, C.accent2, C.accent1];
    const tcol = [C.text1, C.text1, C.text1, C.text1, C.background1];
    layers.forEach((l, i) => {
      const w = 7.4 - i * 0.75, x = MX + (7.4 - w) / 2, y = 1.95 + i * 0.9;
      card(s, x, y, w, 0.78, fills[i], `Camada ${i + 1}`);
      T(s, l.n, { x, y, w, h: 0.78, fontSize: 17, bold: true, color: tcol[i], align: "center", valign: "middle" }, `Camada ${i + 1} nome`);
      T(s, l.d, { x: 8.4, y: y + 0.17, w: 4.4, h: 0.45, fontSize: 14, color: C.text2, valign: "middle" }, `Camada ${i + 1} pergunta`);
    });
    T(s, "Tudo comparado com o seu ponto de partida, levantado no diagnóstico.", { x: MX, y: 6.4, w: 12, h: 0.35, fontSize: 14, bold: true, color: C.text1 }, "Nota medição");
  }

  // 11. Investimento (escuro)
  pres.addSection({ title: "Investimento" });
  {
    const s = dark("Investimento");
    s.background = { path: tex("meov-textura-degrade-01.jpg") };
    T(s, "investimento", { x: MX, y: 0.45, w: 6, h: 0.35, fontSize: 14, bold: true, color: C.accent2 }, "Sobretítulo");
    T(s, "Um ciclo completo, do começo à colheita", { x: MX, y: 0.8, w: 12, h: 0.9, fontSize: 32, bold: true, color: C.background1 }, "Título");
    const boxes = [
      { k: "implantação", v: "R$ 12.000", d: "em 3x de R$ 4.000, junto às 3 primeiras mensalidades", c: HEX.lime },
      { k: "acompanhamento", v: "R$ 12.000", u: "/mês", d: "por 6 meses: produção, gravação e gestão das redes", c: HEX.pink },
    ];
    boxes.forEach((b, i) => {
      const x = MX + i * 6.15;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, on({ x, y: 2.15, w: 5.95, h: 3.0, rectRadius: 0.14, fill: { color: HEX.ink, transparency: 25 }, line: { color: b.c, width: 1.5 } }, `Valor ${i + 1} quadro`));
      T(s, b.k, { x: x + 0.4, y: 2.45, w: 5, h: 0.4, fontSize: 16, bold: true, color: b.c }, `Valor ${i + 1} rótulo`);
      T(s, [{ text: b.v, options: {} }, ...(b.u ? [{ text: b.u, options: { fontSize: 26 } }] : [])], { x: x + 0.4, y: 2.95, w: 5.3, h: 1.2, fontSize: 60, bold: true, color: C.background1 }, `Valor ${i + 1}`);
      T(s, b.d, { x: x + 0.4, y: 4.2, w: 5.2, h: 0.7, fontSize: 15, color: C.background2, valign: "top" }, `Valor ${i + 1} descrição`);
    });
    T(s, [
      { text: "investimento total do ciclo  ", options: { color: C.background2, fontSize: 16 } },
      { text: "R$ 84.000", options: { color: C.accent2, fontSize: 28, bold: true } },
    ], { x: MX, y: 5.6, w: 12, h: 0.6, valign: "middle" }, "Total do ciclo");
    s.addImage(on({ data: L.white.data, x: SW - MX - 1.1, y: 6.75, w: 1.1, h: 1.1 / L.white.ratio }, "Logotipo"));
  }

  // 12. Forma de pagamento
  {
    const s = light("Investimento");
    s.addText("forma de pagamento", { placeholder: "kicker" });
    s.addText("A implantação se dilui nos primeiros meses", { placeholder: "title" });
    s.addChart(pres.charts.BAR, [
      { name: "Acompanhamento", labels: ["mês 1", "mês 2", "mês 3", "mês 4", "mês 5", "mês 6"], values: [12, 12, 12, 12, 12, 12] },
      { name: "Implantação", labels: ["mês 1", "mês 2", "mês 3", "mês 4", "mês 5", "mês 6"], values: [4, 4, 4, 0, 0, 0] },
    ], on({
      x: MX, y: 1.95, w: 7.6, h: 4.6, barDir: "col", barGrouping: "stacked", barGapWidthPct: 55,
      chartColors: [HEX.ink, HEX.pink],
      showValue: true, dataLabelPosition: "ctr", dataLabelColor: HEX.white, dataLabelFontSize: 13, dataLabelFontFace: "+mn-lt", dataLabelFormatCode: '"R$ "0" mil";;;',
      showTitle: true, title: "valor mensal (R$ mil)", titleFontSize: 14, titleColor: HEX.ink, titleFontFace: "+mn-lt",
      showLegend: true, legendPos: "b", legendFontSize: 12, legendColor: HEX.storm, legendFontFace: "+mn-lt",
      catAxisLabelColor: HEX.storm, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", catAxisLineShow: false,
      valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      plotArea: { fill: { color: HEX.mist } }, chartArea: { fill: { color: HEX.mist } },
    }, "Gráfico de pagamento"));
    const rows = [
      { k: "meses 1 a 3", v: "R$ 16.000", d: "R$ 12.000 + R$ 4.000 de implantação" },
      { k: "meses 4 a 6", v: "R$ 12.000", d: "somente o acompanhamento" },
      { k: "total do ciclo", v: "R$ 84.000", d: "6 meses, implantação incluída", hl: true },
    ];
    rows.forEach((r, i) => {
      const y = 1.95 + i * 1.55;
      card(s, 8.65, y, 4.08, 1.38, r.hl ? C.text1 : C.background1, `Resumo ${i + 1} cartão`);
      T(s, r.k, { x: 8.95, y: y + 0.15, w: 3.5, h: 0.35, fontSize: 13, bold: true, color: r.hl ? C.accent2 : C.accent1 }, `Resumo ${i + 1} rótulo`);
      T(s, r.v, { x: 8.95, y: y + 0.45, w: 3.5, h: 0.55, fontSize: 28, bold: true, color: r.hl ? C.background1 : C.text1 }, `Resumo ${i + 1} valor`);
      T(s, r.d, { x: 8.95, y: y + 0.98, w: 3.6, h: 0.32, fontSize: 12, color: r.hl ? C.background2 : C.text2 }, `Resumo ${i + 1} descrição`);
    });
  }

  // 13. Fora do escopo
  pres.addSection({ title: "Combinados" });
  {
    const s = light("Combinados");
    s.addText("para ficar claro desde o começo", { placeholder: "kicker" });
    s.addText("O que não está incluso", { placeholder: "title" });
    const groups = [
      { n: "Produção", i: "TbMovie", d: ["Tráfego pago e anúncios", "Deslocamento fora da Grande SP", "Locação, maquiagem e figurino", "Modelos, drone e equipe extra", "Vídeos longos e eventos"] },
      { n: "Gestão e canais", i: "TbLayoutGrid", d: ["Responder comentários e DMs", "Stories do dia a dia", "Outras redes (TikTok, YouTube…)", "Página e redes da empresa", "E-mail marketing e newsletter"] },
      { n: "Comunicação", i: "TbMessageCircle", d: ["Assessoria de imprensa e crise", "Parcerias com influenciadores", "Tradução e outros idiomas", "Trilhas licenciadas pagas"] },
      { n: "Site", i: "TbWorld", d: ["Site com várias páginas, blog ou loja", "Domínio, hospedagem e manutenção"] },
      { n: "Marca e jurídico", i: "TbScale", d: ["Criação de logotipo", "Registro de marca no INPI", "Revisão jurídica e compliance", "Autorização de imagem de terceiros"] },
      { n: "Resultados", i: "TbTrendingUp", d: ["Garantia de seguidores ou vendas", "Selo de verificado", "Recuperação de conta invadida", "Compra de seguidores (nunca)"] },
    ];
    const cw = 3.9, ch = 2.2, gx = 0.21, gy = 0.2;
    for (let i = 0; i < groups.length; i++) {
      const g = groups[i], x = MX + (i % 3) * (cw + gx), y = 1.95 + Math.floor(i / 3) * (ch + gy);
      card(s, x, y, cw, ch, C.background1, `Fora ${i + 1} cartão`);
      await iconDot(s, g.i, x + 0.22, y + 0.2, 0.5, C.background2, HEX.storm, `Fora ${i + 1}`);
      T(s, g.n, { x: x + 0.85, y: y + 0.2, w: cw - 1, h: 0.5, fontSize: 16, bold: true, color: C.text1, valign: "middle" }, `Fora ${i + 1} nome`);
      T(s, g.d.map((t, k) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: k < g.d.length - 1 } })), { x: x + 0.22, y: y + 0.82, w: cw - 0.4, h: ch - 0.9, fontSize: 13, color: C.text2, valign: "top", paraSpaceAfter: 2 }, `Fora ${i + 1} itens`);
    }
  }

  // 14. Combinados + próximos passos
  {
    const s = light("Combinados");
    s.addText("combinados e próximos passos", { placeholder: "kicker" });
    s.addText("Simples de começar", { placeholder: "title" });
    const rules = [
      { i: "TbCalendarEvent", t: "Ciclo de 6 meses", d: "o tempo do método completo" },
      { i: "TbRepeat", t: "Até 2 rodadas de ajuste", d: "por conteúdo" },
      { i: "TbClockHour4", t: "Mínimos mensais", d: "o saldo do mês não se acumula" },
      { i: "TbSparkles", t: "Extras sob consulta", d: "valores sujeitos à disponibilidade" },
      { i: "TbPhoto", t: "Portfólio Meov", d: "podemos mostrar o trabalho, salvo pedido de sigilo" },
    ];
    for (let i = 0; i < rules.length; i++) {
      const y = 1.95 + i * 0.88;
      await iconDot(s, rules[i].i, MX, y, 0.62, C.background1, HEX.ink, `Combinado ${i + 1}`);
      T(s, [{ text: rules[i].t, options: { bold: true, color: C.text1, breakLine: true } }, { text: rules[i].d, options: { color: C.text2, fontSize: 13 } }], { x: MX + 0.85, y: y - 0.04, w: 5.2, h: 0.72, fontSize: 16, valign: "middle" }, `Combinado ${i + 1} texto`);
    }
    card(s, 7.0, 1.95, 5.73, 4.45, C.text1, "Quadro próximos passos");
    T(s, "próximos passos", { x: 7.4, y: 2.25, w: 5, h: 0.4, fontSize: 16, bold: true, color: C.accent2 }, "Próximos rótulo");
    const steps = ["Aprovação da proposta", "Assinatura do contrato", "Agendamento do diagnóstico e da imersão"];
    const cols = [C.accent2, C.accent1, C.accent3];
    steps.forEach((t, i) => {
      const y = 2.95 + i * 1.05;
      dot(s, 7.4, y, 0.72, cols[i], `Passo ${i + 1}`);
      T(s, String(i + 1), { x: 7.4, y, w: 0.72, h: 0.72, fontSize: 22, bold: true, color: i === 1 ? C.background1 : C.text1, align: "center", valign: "middle" }, `Passo ${i + 1} número`);
      T(s, t, { x: 8.35, y, w: 4.1, h: 0.72, fontSize: 18, color: C.background1, valign: "middle" }, `Passo ${i + 1} texto`);
    });
  }

  // 15. Encerramento
  pres.addSection({ title: "Encerramento" });
  {
    const s = dark("Encerramento");
    s.background = { path: tex("meov-textura-degrade-02-ondas.jpg") };
    s.addImage(on({ data: L.lime.data, x: MX, y: 0.6, w: 1.6, h: 1.6 / L.lime.ratio }, "Logotipo"));
    T(s, "Você já tem a voz.", { x: MX, y: 2.55, w: 12, h: 1.1, fontSize: 56, bold: true, color: C.background1 }, "Frase 1");
    T(s, "A gente ajuda você a colocá-la no mundo.", { x: MX, y: 3.65, w: 12, h: 0.8, fontSize: 30, color: C.accent2 }, "Frase 2");
    T(s, "Sua voz, no ritmo certo.", { x: MX, y: 4.6, w: 12, h: 0.5, fontSize: 18, italic: true, color: C.background2 }, "Frase 3");
    T(s, [
      { text: "@meov.creators", options: { bold: true, color: C.background1 } },
      { text: "   ·   meov.com.br", options: { color: C.background2 } },
    ], { x: MX, y: 6.55, w: 8, h: 0.4, fontSize: 16 }, "Contatos");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("ok", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
