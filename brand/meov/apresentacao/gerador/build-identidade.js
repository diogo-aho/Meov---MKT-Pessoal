// Gera a apresentação de identidade de marca da Meov (2026).
// Uso: NODE_PATH=<node_modules com pptxgenjs, sharp, react, react-icons> IMG_DIR=<pasta com malhas, thumbnails e fotos> node build-identidade.js saida.pptx
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const tb = require("react-icons/tb");
const { applyTheme } = require(process.env.APPLY_THEME || "/root/.claude/skills/synced/ae0912a2-8c90-44b1-85d6-372f2f397ece_1befaab7-63a1-4542-a496-c2e26716f5d9/pptx/scripts/apply_theme.js");

const BRAND = path.join(__dirname, "..", "..");
const IMG = process.env.IMG_DIR || path.join(__dirname, "img");
const OUT = process.argv[2] || "meov-identidade.pptx";

const HEX = { mineral: "FAF8F4", pedra: "D8DACD", musgo: "1C211E", ardosia: "59625A", rosa: "FF19B1", verde: "DBFF00", white: "FFFFFF" };
const MONO = "Geist Mono", LIGHT = "Manrope Light";
const THEME = {
  name: "Meov MP", headFontFace: LIGHT, bodyFontFace: "Manrope",
  colors: { dk1: HEX.musgo, lt1: HEX.mineral, dk2: HEX.ardosia, lt2: HEX.pedra, accent1: HEX.rosa, accent2: HEX.verde, accent3: HEX.pedra, accent4: HEX.ardosia, accent5: HEX.musgo, accent6: HEX.white, hlink: HEX.rosa, folHlink: HEX.ardosia },
};

async function logo(kind, color) {
  const buf = await sharp(path.join(BRAND, "logo", kind, `meov-${kind}-${color}-sem-fundo.svg`), { density: 400 }).png().toBuffer();
  const m = await sharp(buf).metadata();
  return { data: "image/png;base64," + buf.toString("base64"), ratio: m.width / m.height };
}
async function icon(name, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(tb[name], { color: "#" + color, size, strokeWidth: 1.4 }));
  return "image/png;base64," + (await sharp(Buffer.from(svg)).resize(size, size).png().toBuffer()).toString("base64");
}

(async () => {
  const LG = {}, SY = {};
  for (const c of ["musgo", "mineral", "rosa", "verde"]) { LG[c] = await logo("logotipo", c); SY[c] = await logo("simbolo", c); }

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.title = "Meov · Identidade de marca 2026";
  pres.author = "Diogo &amp; Rafael | Meov Marca Pessoal";
  pres.company = "Meov";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  const SW = 13.333, SH = 7.5, MX = 0.75, CW = SW - 2 * MX;

  const head = (dark) => [
    { image: { x: MX, y: 0.42, w: 0.18 * SY.musgo.ratio, h: 0.18, data: (dark ? SY.mineral : SY.musgo).data } },
    { text: { text: "MEOV · IDENTIDADE DE MARCA · 2026", options: { x: SW - MX - 5, y: 0.38, w: 5, h: 0.25, fontFace: MONO, fontSize: 9, charSpacing: 1.5, color: dark ? C.background2 : C.text2, align: "right", margin: 0 } } },
  ];
  const ph = (dark, kc) => [
    { placeholder: { options: { name: "kicker", type: "body", x: MX, y: 0.95, w: 9, h: 0.3, fontFace: MONO, fontSize: 11, charSpacing: 1.5, color: kc, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 1.25, w: CW, h: 0.8, fontFace: LIGHT, fontSize: 34, charSpacing: -0.6, color: dark ? C.background1 : C.text1, margin: 0, valign: "top", align: "left" }, text: "" } },
  ];
  pres.defineSlideMaster({ title: "Claro", background: { color: C.background1 }, objects: [...head(false), ...ph(false, C.accent1)], slideNumber: { x: SW - MX - 0.6, y: 6.95, w: 0.6, h: 0.25, fontFace: MONO, fontSize: 9, color: C.text2, align: "right", margin: 0 } });
  pres.defineSlideMaster({ title: "Pedra", background: { color: C.background2 }, objects: [...head(false), ...ph(false, C.text1)], slideNumber: { x: SW - MX - 0.6, y: 6.95, w: 0.6, h: 0.25, fontFace: MONO, fontSize: 9, color: C.text2, align: "right", margin: 0 } });
  pres.defineSlideMaster({ title: "Escuro", background: { color: C.text1 }, objects: [...head(true), ...ph(true, C.accent2)], slideNumber: { x: SW - MX - 0.6, y: 6.95, w: 0.6, h: 0.25, fontFace: MONO, fontSize: 9, color: C.background2, align: "right", margin: 0 } });
  pres.defineSlideMaster({ title: "Livre", background: { color: C.background1 }, objects: [] });

  const on = (o, n) => { o.objectName = n; return o; };
  const T = (s, text, opts, name) => s.addText(text, on({ isTextBox: true, margin: 0, fontFace: "Manrope", ...opts }, name));
  const label = (s, text, x, y, w, color, name, align = "left") => T(s, text.toUpperCase(), { x, y, w, h: 0.26, fontFace: MONO, fontSize: 10, charSpacing: 1.5, color, align }, name);
  const box = (s, x, y, w, h, fill, name, r = 0.16, line) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, on({ x, y, w, h, rectRadius: r, fill: { color: fill }, line: line || { color: fill, width: 0 } }, name));
  const rect = (s, x, y, w, h, fill, name) => s.addShape(pres.shapes.RECTANGLE, on({ x, y, w, h, fill: { color: fill }, line: { color: fill, width: 0 } }, name));
  const dot = (s, x, y, d, fill, name) => s.addShape(pres.shapes.OVAL, on({ x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 } }, name));
  const img = (s, art, x, y, w, name) => s.addImage(on({ data: art.data, x, y, w, h: w / art.ratio }, name));
  const imgH = (s, art, x, y, h, name) => s.addImage(on({ data: art.data, x, y, w: h * art.ratio, h }, name));
  const iconDot = async (s, iname, x, y, d, fill, fg, name) => { dot(s, x, y, d, fill, name + " fundo"); s.addImage(on({ data: await icon(iname, fg), x: x + d * 0.24, y: y + d * 0.24, w: d * 0.52, h: d * 0.52 }, name + " ícone")); };
  const body = (o = {}) => ({ fontSize: 14, color: C.text2, valign: "top", lineSpacingMultiple: 1.15, ...o });
  const slide = (m, sec) => pres.addSlide({ masterName: m, sectionTitle: sec });
  const ring = (s, x, y, d, name) => {
    s.addShape(pres.shapes.BLOCK_ARC, on({ x, y, w: d, h: d, angleRange: [90, 270], arcThicknessRatio: 0.07, fill: { color: C.accent1 }, line: { color: C.accent1, width: 0 } }, name + " rosa"));
    s.addShape(pres.shapes.BLOCK_ARC, on({ x, y, w: d, h: d, angleRange: [270, 90], arcThicknessRatio: 0.07, fill: { color: C.accent2 }, line: { color: C.accent2, width: 0 } }, name + " verde"));
  };

  // ================= 1 · Capa
  pres.addSection({ title: "Abertura" });
  {
    const s = slide("Livre", "Abertura");
    imgH(s, LG.musgo, MX, 0.75, 0.42, "Logotipo");
    T(s, [{ text: "Meov Marca Pessoal", options: { breakLine: true, color: C.text1 } }, { text: "Estratégia & Identidade Visual", options: { color: C.text2 } }], { x: SW - MX - 4.5, y: 0.72, w: 4.5, h: 0.6, fontSize: 13, align: "right" }, "Cliente e serviço");
    T(s, "Identidade de marca", { x: MX, y: 2.75, w: 9, h: 1.2, fontFace: LIGHT, fontSize: 64, charSpacing: -1.5, color: C.text1 }, "Título");
    T(s, "Sua voz, no ritmo certo.", { x: MX, y: 3.95, w: 9, h: 0.6, fontFace: LIGHT, fontSize: 26, color: C.accent1 }, "Assinatura");
    rect(s, MX, 4.75, 0.6, 0.08, C.accent1, "Traço rosa");
    rect(s, MX + 0.6, 4.75, 0.6, 0.08, C.accent2, "Traço verde");
    T(s, "MEOV.COM.BR\n@MEOV.MP", { x: MX, y: 6.45, w: 4, h: 0.5, fontFace: MONO, fontSize: 10, charSpacing: 1.5, color: C.text2 }, "Contatos");
    T(s, "2026", { x: SW / 2 - 1, y: 6.7, w: 2, h: 0.3, fontFace: MONO, fontSize: 10, color: C.text2, align: "center" }, "Ano");
    imgH(s, SY.musgo, SW - MX - 0.42 * SY.musgo.ratio, 6.55, 0.42, "Símbolo");
  }

  // ================= 2 · Frase de abertura
  {
    const s = slide("Claro", "Abertura");
    T(s, [{ text: "Vamos conhecer o time que ajuda especialistas a ", options: { color: C.text1 } }, { text: "colocar a própria voz no mundo.", options: { color: C.accent1 } }], { x: MX, y: 3.6, w: 9.5, h: 2.4, fontFace: LIGHT, fontSize: 44, charSpacing: -1, valign: "bottom", lineSpacingMultiple: 1.0 }, "Frase");
    rect(s, MX, 6.3, 0.6, 0.08, C.accent1, "Traço rosa");
    rect(s, MX + 0.6, 6.3, 0.6, 0.08, C.accent2, "Traço verde");
  }

  // ================= 3 · Nova fase
  pres.addSection({ title: "A marca" });
  {
    const s = slide("Claro", "A marca");
    s.addText("01 · uma nova fase".toUpperCase(), { placeholder: "kicker" });
    s.addText("De produtora de vídeos a marca pessoal.", { placeholder: "title" });
    T(s, "A Meov nasceu fazendo filmes que tocam na alma. Com o tempo, percebemos que o que mais nos movia era outra coisa: ver alguém que já é referência finalmente aparecer como é. Nesta nova fase, nós mesmos redesenhamos a nossa marca para contar essa história.", { x: MX, y: 2.2, w: CW, h: 0.95, ...body({ fontSize: 15 }) }, "Texto fase");
    const cols = [
      { y: "antes", n: "Produtora audiovisual", d: "Conexão através de experiências audiovisuais. O centro era o filme.", f: C.background2, t: C.text1, s: C.text2 },
      { y: "agora", n: "Meov Marca Pessoal", d: "Estratégia, audiovisual e gestão das redes para especialistas. O centro é a pessoa.", f: C.text1, t: C.background1, s: C.background2 },
    ];
    cols.forEach((c, i) => {
      const x = MX + i * 6.0, y = 3.45;
      box(s, x, y, 5.6, 2.2, c.f, `Fase ${i + 1}`);
      label(s, c.y, x + 0.4, y + 0.35, 4, i ? C.accent2 : C.accent1, `Fase ${i + 1} rótulo`);
      T(s, c.n, { x: x + 0.4, y: y + 0.7, w: 4.8, h: 0.6, fontFace: LIGHT, fontSize: 26, color: c.t }, `Fase ${i + 1} nome`);
      T(s, c.d, { x: x + 0.4, y: y + 1.35, w: 4.8, h: 0.7, fontSize: 14, color: c.s, valign: "top" }, `Fase ${i + 1} texto`);
    });
    T(s, "→", { x: MX + 5.55, y: 4.2, w: 0.5, h: 0.6, fontSize: 28, color: C.accent1, align: "center" }, "Seta");
    T(s, "O que fica: o desenho do logotipo e do símbolo, o fino trato audiovisual e a ideia de conexão. O que muda: o foco, as cores, a tipografia, a voz e os serviços.", { x: MX, y: 5.95, w: CW - 0.8, h: 0.7, ...body({ fontSize: 13 }) }, "Fica e muda");
  }

  // ================= 4 · Posicionamento
  {
    const s = slide("Claro", "A marca");
    s.addText("02 · posicionamento".toUpperCase(), { placeholder: "kicker" });
    s.addText("Não somos uma agência de marketing.", { placeholder: "title" });
    box(s, MX, 2.2, 6.3, 4.4, C.text1, "Quadro manifesto");
    T(s, "Somos um time que apoia técnica e emocionalmente quem decide colocar a própria voz no mundo: da estratégia à câmera e à gestão das redes.", { x: MX + 0.45, y: 2.6, w: 5.4, h: 2.2, fontFace: LIGHT, fontSize: 24, color: C.background1, valign: "top", lineSpacingMultiple: 1.05 }, "Manifesto");
    label(s, "mantra", MX + 0.45, 5.05, 4, C.accent2, "Rótulo mantra");
    T(s, "Ajudamos especialistas que já são referência a colocar sua voz no mundo digital, através de estratégia, audiovisual e gestão das redes, com acolhimento do começo ao fim.", { x: MX + 0.45, y: 5.35, w: 5.4, h: 1.1, fontSize: 13, color: C.background2, valign: "top" }, "Mantra");
    const rx = 7.35, rw = SW - MX - rx;
    box(s, rx, 2.2, rw, 2.1, C.background2, "Quadro público");
    label(s, "para quem", rx + 0.35, 2.5, 4, C.accent1, "Rótulo público");
    T(s, "Empresários e especialistas que já faturam acima de R$ 30 mil/mês, têm reputação construída fora das redes e ainda não conseguiram se posicionar no digital.", { x: rx + 0.35, y: 2.85, w: rw - 0.7, h: 1.3, ...body({ fontSize: 14, color: C.text1 }) }, "Público");
    box(s, rx, 4.5, rw, 2.1, C.accent2, "Quadro princípios");
    label(s, "princípios", rx + 0.35, 4.8, 4, C.text1, "Rótulo princípios");
    T(s, [{ text: "Raiz antes de vitrine.", options: { breakLine: true } }, { text: "Verdade acima de pose.", options: { breakLine: true } }, { text: "Constância no lugar de inspiração." }], { x: rx + 0.35, y: 5.15, w: rw - 0.7, h: 1.3, fontFace: LIGHT, fontSize: 20, color: C.text1, valign: "top" }, "Princípios");
  }

  // ================= 5 · Serviços
  {
    const s = slide("Claro", "A marca");
    s.addText("03 · o que fazemos".toUpperCase(), { placeholder: "kicker" });
    s.addText("Três frentes, um mesmo time.", { placeholder: "title" });
    const f = [
      { n: "Estratégia", i: "TbCompass", c: C.accent2, fg: HEX.musgo, d: "O Método Meov: diagnóstico, imersão presencial e um Brandbook que organiza quem você é, para quem fala e como se diferencia.", tags: ["Método Meov", "Brandbook", "plano de 6 meses"] },
      { n: "Audiovisual", i: "TbMovie", c: C.accent1, fg: HEX.white, d: "Direção de imagem, ensaio, gravação e edição. Dias de gravação semanais, quinzenais ou mensais, com o nosso fino trato.", tags: ["ensaio de marca", "Vlogcast", "reels"] },
      { n: "Gestão das redes", i: "TbLayoutGrid", c: C.text1, fg: HEX.verde, d: "Instagram, LinkedIn e Site/LP com calendário, roteiros, legendas, design, publicação e relatórios. Você grava e aprova.", tags: ["Instagram", "LinkedIn", "Site/LP"] },
    ];
    const w = (CW - 0.5) / 3;
    for (let i = 0; i < 3; i++) {
      const x = MX + i * (w + 0.25), y = 2.2;
      box(s, x, y, w, 3.6, C.background2, `Frente ${i + 1}`);
      await iconDot(s, f[i].i, x + 0.35, y + 0.35, 0.8, f[i].c, f[i].fg, `Frente ${i + 1}`);
      T(s, f[i].n, { x: x + 0.35, y: y + 1.35, w: w - 0.7, h: 0.5, fontFace: LIGHT, fontSize: 24, color: C.text1 }, `Frente ${i + 1} nome`);
      T(s, f[i].d, { x: x + 0.35, y: y + 1.9, w: w - 0.7, h: 1.1, ...body({ fontSize: 13 }) }, `Frente ${i + 1} texto`);
      T(s, f[i].tags.map((t) => t.toUpperCase()).join("  ·  "), { x: x + 0.35, y: y + 3.1, w: w - 0.6, h: 0.3, fontFace: MONO, fontSize: 9, charSpacing: 0.5, color: C.text1 }, `Frente ${i + 1} tags`);
    }
    box(s, MX, 6.0, CW, 0.7, C.text1, "Faixa ciclo");
    T(s, [{ text: "CICLO MEOV  ", options: { fontFace: MONO, fontSize: 10, charSpacing: 1.5, color: C.accent2 } }, { text: "implantação, 6 meses de presença constante e revisão no 3º e no 6º mês.", options: { fontSize: 14, color: C.background1 } }], { x: MX + 0.35, y: 6.0, w: CW - 0.7, h: 0.7, valign: "middle" }, "Texto ciclo");
  }

  // ================= 6 · Tom de voz
  {
    const s = slide("Pedra", "A marca");
    s.addText("04 · tom de voz".toUpperCase(), { placeholder: "kicker" });
    s.addText("Somos acolhedores, claros, verdadeiros. Nunca guru.", { placeholder: "title" });
    const p = [
      ["Acolhedores", "Falamos com quem tem medo de se expor como quem entende esse medo. Sem pressa, sem julgamento."],
      ["Claros", "Frases curtas, palavras simples, nada de jargão de marketing. Se precisa explicar, está complicado."],
      ["Verdadeiros", "Nada de fórmula, promessa de viral ou personagem. A história real é a estratégia."],
    ];
    const w = (CW - 0.5) / 3;
    p.forEach((it, i) => {
      const x = MX + i * (w + 0.25);
      dot(s, x, 2.3, 0.16, i === 1 ? C.accent2 : C.accent1, `Pilar ${i + 1} ponto`);
      T(s, it[0], { x: x + 0.3, y: 2.18, w: w - 0.3, h: 0.4, fontSize: 18, color: C.text1 }, `Pilar ${i + 1} nome`);
      T(s, it[1], { x, y: 2.65, w: w - 0.2, h: 0.9, ...body({ fontSize: 13 }) }, `Pilar ${i + 1} texto`);
    });
    const ex = [
      { k: "✓ assim", t: "Você já tem a voz. A gente ajuda você a colocá-la no mundo.", c: "4F7A45", f: "Manrope", fs: 17 },
      { k: "✕ não assim", t: "VIRALIZE EM 7 DIAS!", c: "B23A3A", f: "Manrope", fs: 22, b: true },
      { k: "✕ não assim", t: "O hack secreto para bombar o seu perfil.", c: "B23A3A", f: "Manrope", fs: 17 },
    ];
    ex.forEach((e, i) => {
      const x = MX + i * (w + 0.25), y = 3.75;
      box(s, x, y, w, 1.55, C.background1, `Exemplo ${i + 1}`);
      label(s, e.k, x + 0.3, y + 0.25, w - 0.6, e.c, `Exemplo ${i + 1} rótulo`);
      T(s, e.t, { x: x + 0.3, y: y + 0.6, w: w - 0.6, h: 0.85, fontFace: e.b ? "Manrope SemiBold" : LIGHT, fontSize: e.fs, color: C.text1, valign: "top" }, `Exemplo ${i + 1} texto`);
    });
    label(s, "palavras nossas", MX, 5.55, 5, C.text1, "Rótulo nossas");
    label(s, "palavras que não usamos", MX + CW / 2 + 0.1, 5.55, 5, C.text1, "Rótulo evitamos");
    const chips = (list, x0, fill, tc, line, nm) => {
      let x = x0;
      list.forEach((t, i) => {
        const w2 = 0.35 + t.length * 0.095;
        s.addShape(pres.shapes.ROUNDED_RECTANGLE, on({ x, y: 5.9, w: w2, h: 0.42, rectRadius: 0.21, fill: { color: fill }, line: line || { color: fill, width: 0 } }, `${nm} ${i + 1} fundo`));
        T(s, t, { x, y: 5.9, w: w2, h: 0.42, fontSize: 12, color: tc, align: "center", valign: "middle" }, `${nm} ${i + 1}`);
        x += w2 + 0.1;
      });
    };
    chips(["voz", "ritmo", "verdade", "raiz", "constância", "presença"], MX, C.text1, C.background1, null, "Nossa");
    chips(["viralizar", "hack", "fórmula", "bombar", "guru"], MX + CW / 2 + 0.1, C.background1, C.text2, { color: HEX.ardosia, width: 0.75 }, "Evitamos");
  }

  // ================= 7 · Divisor
  pres.addSection({ title: "Identidade visual" });
  {
    const s = slide("Escuro", "Identidade visual");
    T(s, "Identidade visual", { x: MX, y: 5.3, w: 9, h: 1.0, fontFace: LIGHT, fontSize: 54, charSpacing: -1.2, color: C.background1 }, "Título divisor");
    T(s, "Cores, tipografia, logotipo e símbolo.", { x: MX, y: 6.3, w: 9, h: 0.4, fontSize: 15, color: C.background2 }, "Subtítulo divisor");
    imgH(s, SY.rosa, SW - MX - 2.6 * SY.rosa.ratio * 0.55, 1.6, 1.45, "Símbolo rosa");
    imgH(s, SY.verde, SW - MX - 2.6 * SY.verde.ratio * 0.55, 3.25, 1.45, "Símbolo verde");
  }

  // ================= 8 · Cores
  {
    const s = slide("Livre", "Identidade visual");
    const blocks = [
      { n: "Branco mineral", h: "#FAF8F4", p: "50%", x: 0, y: 0, w: 8.33, hh: 4.2, f: HEX.mineral, t: HEX.musgo },
      { n: "Pedra líquen", h: "#D8DACD", p: "30%", x: 8.33, y: 0, w: 5.0, hh: 4.2, f: HEX.pedra, t: HEX.musgo },
      { n: "Preto musgo", h: "#1C211E", p: "9%", x: 0, y: 4.2, w: 4.4, hh: 3.3, f: HEX.musgo, t: HEX.mineral },
      { n: "Cinza ardósia verde", h: "#59625A", p: "8%", x: 4.4, y: 4.2, w: 4.0, hh: 3.3, f: HEX.ardosia, t: HEX.mineral },
      { n: "Rosa choque", h: "#FF19B1", p: "1,5%", x: 8.4, y: 4.2, w: 2.47, hh: 3.3, f: HEX.rosa, t: HEX.mineral },
      { n: "Verde vibrante", h: "#DBFF00", p: "1,5%", x: 10.87, y: 4.2, w: 2.47, hh: 3.3, f: HEX.verde, t: HEX.musgo },
    ];
    blocks.forEach((b, i) => {
      rect(s, b.x, b.y, b.w, b.hh, b.f, `Cor ${i + 1}`);
      T(s, b.n, { x: b.x + 0.4, y: b.y + 0.4, w: b.w - 0.6, h: 0.45, fontFace: LIGHT, fontSize: i < 2 ? 24 : 18, color: b.t }, `Cor ${i + 1} nome`);
      T(s, b.h, { x: b.x + 0.4, y: b.y + 0.85, w: b.w - 0.6, h: 0.3, fontFace: MONO, fontSize: 11, charSpacing: 1, color: b.t }, `Cor ${i + 1} hex`);
      T(s, b.p, { x: b.x + 0.4, y: b.y + b.hh - 1.25, w: b.w - 0.6, h: 0.95, fontFace: LIGHT, fontSize: i < 2 ? 60 : 44, color: b.t, valign: "bottom" }, `Cor ${i + 1} proporção`);
    });
    T(s, "05 · CORES", { x: 5.5, y: 0.45, w: 2.6, h: 0.26, fontFace: MONO, fontSize: 10, charSpacing: 1.5, color: C.accent1, align: "right" }, "Rótulo cores");
  }

  // ================= 9 · Por que as cores
  {
    const s = slide("Claro", "Identidade visual");
    s.addText("06 · o porquê das cores".toUpperCase(), { placeholder: "kicker" });
    s.addText("Muito respiro, dois pontos de energia.", { placeholder: "title" });
    const why = [
      { c: [HEX.mineral, HEX.pedra], n: "Branco mineral e pedra líquen", d: "A base clara e calma. Dá espaço para a pessoa aparecer: o destaque vem do respiro, não do excesso." },
      { c: [HEX.musgo, HEX.ardosia], n: "Preto musgo e ardósia", d: "Profundidade e seriedade. Usados nos textos e nos momentos que pedem elegância." },
      { c: [HEX.rosa], n: "Rosa choque", d: "A emoção. O coração que pulsa no símbolo: acolhimento, coragem e conexão." },
      { c: [HEX.verde], n: "Verde vibrante", d: "A energia do movimento. Constância, ação e o play de quem decide aparecer." },
    ];
    why.forEach((w, i) => {
      const y = 2.2 + i * 1.02;
      w.c.forEach((c, k) => s.addShape(pres.shapes.OVAL, on({ x: MX + k * 0.42, y, w: 0.7, h: 0.7, fill: { color: c }, line: { color: HEX.pedra, width: 0.75 } }, `Amostra ${i + 1}.${k + 1}`)));
      T(s, w.n, { x: MX + 1.35, y: y - 0.02, w: 5.6, h: 0.36, fontSize: 17, color: C.text1 }, `Porquê ${i + 1} nome`);
      T(s, w.d, { x: MX + 1.35, y: y + 0.35, w: 5.6, h: 0.6, ...body({ fontSize: 13 }) }, `Porquê ${i + 1} texto`);
    });
    const rx = 8.0, rw = SW - MX - rx;
    box(s, rx, 2.2, rw, 4.35, C.text1, "Quadro soma");
    dot(s, rx + 0.6, 2.75, 1.25, C.accent1, "Círculo rosa");
    T(s, "+", { x: rx + 1.9, y: 3.0, w: 0.5, h: 0.7, fontFace: LIGHT, fontSize: 36, color: C.background1, align: "center" }, "Mais");
    dot(s, rx + 2.45, 2.75, 1.25, C.accent2, "Círculo verde");
    T(s, "A Meov é a soma dos dois.", { x: rx + 0.5, y: 4.3, w: rw - 1.0, h: 0.5, fontFace: LIGHT, fontSize: 22, color: C.background1 }, "Soma título");
    T(s, "Rosa e verde aparecem sempre em par, na mesma medida, sobre o branco e os tons de cinza. Verde vibrante só sobre fundo escuro ou como superfície com texto escuro.", { x: rx + 0.5, y: 4.9, w: rw - 1.0, h: 1.4, fontSize: 13, color: C.background2, valign: "top", lineSpacingMultiple: 1.15 }, "Soma texto");
  }

  // ================= 10 · Tipografia
  {
    const s = slide("Claro", "Identidade visual");
    s.addText("07 · tipografia".toUpperCase(), { placeholder: "kicker" });
    s.addText([{ text: "Duas famílias. ", options: { color: C.text1 } }, { text: "Uma voz.", options: { color: C.text2 } }], { placeholder: "title" });
    const w = (CW - 0.3) / 2;
    box(s, MX, 2.2, w, 4.4, C.background2, "Cartão Manrope");
    label(s, "títulos e textos", MX + 0.45, 2.55, 5, C.text2, "Rótulo Manrope");
    T(s, "Aa", { x: MX + 0.45, y: 2.9, w: 3, h: 1.5, fontFace: LIGHT, fontSize: 96, color: C.text1 }, "Aa Manrope");
    T(s, "Manrope", { x: MX + 0.45, y: 4.45, w: 5, h: 0.6, fontFace: LIGHT, fontSize: 30, color: C.text1 }, "Nome Manrope");
    T(s, "Acolhedora e moderna. Leve nos títulos, porque o destaque vem do espaço, não do negrito. É a voz que acolhe.", { x: MX + 0.45, y: 5.1, w: w - 0.9, h: 0.8, ...body({ fontSize: 13, color: C.text1 }) }, "Porquê Manrope");
    T(s, "LIGHT 300 · REGULAR 400 · MEDIUM 500 · SEMIBOLD 600", { x: MX + 0.45, y: 6.15, w: w - 0.9, h: 0.26, fontFace: MONO, fontSize: 9, charSpacing: 1, color: C.text2 }, "Pesos Manrope");
    const x2 = MX + w + 0.3;
    box(s, x2, 2.2, w, 4.4, C.text1, "Cartão Geist");
    label(s, "dados e rótulos", x2 + 0.45, 2.55, 5, C.accent2, "Rótulo Geist");
    T(s, "Aa", { x: x2 + 0.45, y: 2.9, w: 3, h: 1.5, fontFace: MONO, fontSize: 88, color: C.background1 }, "Aa Geist");
    T(s, "Geist Mono", { x: x2 + 0.45, y: 4.45, w: 5, h: 0.6, fontFace: MONO, fontSize: 28, color: C.background1 }, "Nome Geist");
    T(s, "Precisa e técnica. Só em números, valores, datas, @, sites e selos. É o lado técnico da Meov.", { x: x2 + 0.45, y: 5.1, w: w - 0.9, h: 0.8, fontSize: 13, color: C.background2, valign: "top", lineSpacingMultiple: 1.15 }, "Porquê Geist");
    T(s, "REGULAR 400", { x: x2 + 0.45, y: 6.15, w: w - 0.9, h: 0.26, fontFace: MONO, fontSize: 9, charSpacing: 1, color: C.background2 }, "Pesos Geist");
  }

  // ================= 11 · Escala e regras
  {
    const s = slide("Claro", "Identidade visual");
    s.addText("08 · escala e regras".toUpperCase(), { placeholder: "kicker" });
    s.addText("Títulos leves, frases serenas.", { placeholder: "title" });
    const sc = [
      ["display", "64 / 300", "Sua voz, no ritmo certo", LIGHT, 30],
      ["título", "34 / 300", "Seja nas redes quem você já é", LIGHT, 20],
      ["texto", "14 / 400", "A gente cuida da técnica. Você aparece com verdade.", "Manrope", 14],
      ["dado", "11 / mono", "12 REELS · 8 CARROSSÉIS · 4 POSTS · @MEOV.MP", MONO, 11],
    ];
    sc.forEach((r, i) => {
      const y = 2.2 + [0, 0.95, 1.7, 2.3][i];
      s.addShape(pres.shapes.LINE, on({ x: MX, y, w: 7.3, h: 0, line: { color: HEX.pedra, width: 0.75 } }, `Escala linha ${i + 1}`));
      label(s, r[0], MX, y + 0.15, 1.3, C.text2, `Escala rótulo ${i + 1}`);
      T(s, r[1], { x: MX + 1.3, y: y + 0.15, w: 1.2, h: 0.26, fontFace: MONO, fontSize: 9, color: C.text2 }, `Escala tamanho ${i + 1}`);
      T(s, r[2], { x: MX + 2.5, y: y + 0.05, w: 4.9, h: [0.75, 0.6, 0.45, 0.4][i], fontFace: r[3], fontSize: r[4], color: C.text1, charSpacing: i < 2 ? -0.6 : 0, valign: "middle" }, `Escala exemplo ${i + 1}`);
    });
    const rules = [
      ["Títulos leves", "Light 300 nos títulos grandes."],
      ["Letras próximas", "−0,02 a −0,03 em nos títulos."],
      ["Mono só para dados", "Nunca em frases."],
      ["Dois pesos por peça", "300 no título, 400 no texto."],
      ["Cor do texto", "Musgo no principal, ardósia no apoio."],
      ["Frases serenas", "Títulos curtos e calmos."],
    ];
    const rx = 8.45, rw = SW - MX - rx;
    box(s, rx, 2.2, rw, 4.4, C.background2, "Quadro regras");
    rules.forEach((r, i) => {
      const y = 2.45 + i * 0.68;
      T(s, String(i + 1).padStart(2, "0"), { x: rx + 0.35, y, w: 0.45, h: 0.3, fontFace: MONO, fontSize: 10, color: i % 2 ? C.text1 : C.accent1 }, `Regra ${i + 1} número`);
      T(s, [{ text: r[0], options: { color: C.text1, breakLine: true, fontSize: 14 } }, { text: r[1], options: { color: C.text2, fontSize: 12 } }], { x: rx + 0.85, y: y - 0.04, w: rw - 1.1, h: 0.62, valign: "top" }, `Regra ${i + 1}`);
    });
  }

  // ================= 12-13 · Malhas
  for (const [k, n, d] of [["logotipo", "09 · logotipo", "Curvas contínuas, sem quebra: o m que vira onda."], ["simbolo", "10 · símbolo", "As diagonais do m e do v seguem o mesmo ângulo de 69,5°."]]) {
    const s = slide("Escuro", "Identidade visual");
    s.background = { path: path.join(IMG, `malha-${k}.png`) };
    label(s, n, MX, 0.95, 5, C.accent2, "Rótulo malha");
    T(s, d, { x: MX, y: 6.6, w: 8, h: 0.35, fontSize: 13, color: C.background2 }, "Legenda malha");
  }

  // ================= 14 · Porquê do símbolo
  {
    const s = slide("Claro", "Identidade visual");
    s.addText("11 · o porquê do símbolo".toUpperCase(), { placeholder: "kicker" });
    s.addText("Uma marca que pulsa.", { placeholder: "title" });
    imgH(s, SY.musgo, MX, 2.45, 1.7, "Símbolo");
    const eq = [["m", "Meov e marca", C.accent1], ["v", "voz", C.accent2], ["∿", "batimento e onda sonora", C.text1]];
    eq.forEach((e, i) => {
      const x = MX + i * 2.35;
      T(s, e[0], { x, y: 4.5, w: 1.0, h: 0.9, fontFace: LIGHT, fontSize: 54, color: C.text1 }, `Equação ${i + 1}`);
      rect(s, x, 5.45, 0.5, 0.07, e[2], `Equação ${i + 1} traço`);
      T(s, e[1], { x, y: 5.6, w: 2.0, h: 0.5, fontSize: 13, color: C.text2, valign: "top" }, `Equação ${i + 1} legenda`);
      if (i < 2) T(s, "+", { x: x + 1.15, y: 4.55, w: 0.6, h: 0.8, fontFace: LIGHT, fontSize: 40, color: C.text2, align: "center" }, `Mais ${i + 1}`);
    });
    const rx = 7.75, rw = SW - MX - rx;
    box(s, rx, 2.2, rw, 4.4, C.background2, "Quadro porquê");
    T(s, [
      { text: "O símbolo nasceu do m e do v de \"meov\", unidos por uma linha de batimento cardíaco: experiências que fazem o coração pulsar.", options: { breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 8 } },
      { text: "Na Meov Marca Pessoal, a mesma linha ganha um segundo sentido: é a onda sonora da voz de cada cliente e o ritmo da presença constante nas redes.", options: { breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 8 } },
      { text: "Sua voz, no ritmo certo.", options: { color: C.accent1 } },
    ], { x: rx + 0.4, y: 2.55, w: rw - 0.8, h: 3.8, fontSize: 15, color: C.text1, valign: "top", lineSpacingMultiple: 1.15 }, "Texto porquê");
  }

  // ================= 15 · Resultado
  {
    const s = slide("Livre", "Identidade visual");
    const t = [
      { f: HEX.mineral, l: LG.musgo }, { f: HEX.musgo, l: LG.verde },
      { f: HEX.rosa, l: LG.musgo }, { f: HEX.pedra, l: LG.musgo },
    ];
    t.forEach((q, i) => {
      const x = (i % 2) * SW / 2, y = Math.floor(i / 2) * SH / 2;
      rect(s, x, y, SW / 2, SH / 2, q.f, `Fundo ${i + 1}`);
      const w = 2.6;
      img(s, q.l, x + SW / 4 - w / 2, y + SH / 4 - w / q.l.ratio / 2, w, `Logotipo ${i + 1}`);
    });
    T(s, "12 · RESULTADO", { x: 0.5, y: 0.4, w: 4, h: 0.26, fontFace: MONO, fontSize: 10, charSpacing: 1.5, color: C.accent1 }, "Rótulo resultado");
  }

  // ================= 16 · Versões e redução
  {
    const s = slide("Claro", "Identidade visual");
    s.addText("13 · versões".toUpperCase(), { placeholder: "kicker" });
    s.addText("Logotipo para apresentar. Símbolo para assinar.", { placeholder: "title" });
    box(s, MX, 2.2, 7.4, 2.6, C.background2, "Quadro principal");
    img(s, LG.musgo, MX + 0.8, 2.95, 3.6, "Logotipo 100");
    imgH(s, SY.musgo, MX + 5.2, 3.0, 0.95, "Símbolo 25");
    s.addShape(pres.shapes.LINE, on({ x: MX + 0.8, y: 4.3, w: 3.6, h: 0, line: { color: HEX.ardosia, width: 0.75 } }, "Régua 100"));
    s.addShape(pres.shapes.LINE, on({ x: MX + 5.2, y: 4.3, w: 0.95 * SY.musgo.ratio, h: 0, line: { color: HEX.ardosia, width: 0.75 } }, "Régua 25"));
    T(s, "100%", { x: MX + 0.8, y: 4.38, w: 3.6, h: 0.25, fontFace: MONO, fontSize: 9, color: C.text2, align: "center" }, "Medida 100");
    T(s, "25%", { x: MX + 5.2, y: 4.38, w: 0.95 * SY.musgo.ratio, h: 0.25, fontFace: MONO, fontSize: 9, color: C.text2, align: "center" }, "Medida 25");
    const sm = [{ f: HEX.musgo, l: LG.mineral }, { f: HEX.verde, l: LG.musgo }, { f: HEX.rosa, l: LG.mineral }, { f: HEX.musgo, l: SY.rosa, s: true }];
    sm.forEach((q, i) => {
      const x = MX + i * 1.9, y = 5.0;
      box(s, x, y, 1.75, 1.6, q.f, `Versão ${i + 1}`, 0.12);
      const w = q.s ? 0.95 : 1.3;
      img(s, q.l, x + (1.75 - w) / 2, y + 0.8 - w / q.l.ratio / 2, w, `Versão ${i + 1} logo`);
    });
    const rx = 8.55, rw = SW - MX - rx;
    T(s, [
      { text: "Logotipo", options: { color: C.text1, fontSize: 17, breakLine: true } },
      { text: "Capas, site, assinaturas de e-mail e tudo que apresenta a Meov pela primeira vez.", options: { color: C.text2, fontSize: 13, breakLine: true } },
      { text: " ", options: { fontSize: 10, breakLine: true } },
      { text: "Símbolo", options: { color: C.text1, fontSize: 17, breakLine: true } },
      { text: "Avatar, favicon, selos, rodapés e aplicações pequenas, a partir de 16 px.", options: { color: C.text2, fontSize: 13, breakLine: true } },
      { text: " ", options: { fontSize: 10, breakLine: true } },
      { text: "Sem fundo", options: { color: C.text1, fontSize: 17, breakLine: true } },
      { text: "Musgo sobre claros; mineral ou verde sobre escuros. Os 26 arquivos estão em brand/meov/logo.", options: { color: C.text2, fontSize: 13 } },
    ], { x: rx, y: 2.2, w: rw, h: 4.4, valign: "top", lineSpacingMultiple: 1.12 }, "Usos");
  }

  // ================= 17 · Instagram
  pres.addSection({ title: "Universo visual" });
  {
    const s = slide("Pedra", "Universo visual");
    s.addText("14 · universo visual · instagram".toUpperCase(), { placeholder: "kicker" });
    s.addText("@meov.mp", { placeholder: "title" });
    const px = MX, py = 2.15, pw = 6.2;
    box(s, px, py, pw, 4.6, C.text1, "Perfil");
    dot(s, px + 0.4, py + 0.4, 1.25, C.background1, "Avatar fundo");
    img(s, SY.musgo, px + 0.58, py + 0.85, 0.9, "Avatar");
    T(s, "meov.mp", { x: px + 1.95, y: py + 0.45, w: 3.5, h: 0.4, fontSize: 18, color: C.background1 }, "Usuário");
    T(s, "Meov | Marca Pessoal", { x: px + 1.95, y: py + 0.85, w: 3.8, h: 0.3, fontSize: 13, color: C.background1 }, "Nome");
    T(s, "ajudamos especialistas a colocar sua voz no digital: estratégia, audiovisual e gestão das redes, com acolhimento do começo ao fim.\n↓ fala com a Meov", { x: px + 1.95, y: py + 1.18, w: 3.95, h: 1.0, fontSize: 11, color: C.background2, valign: "top" }, "Bio");
    T(s, "MEOV.COM.BR", { x: px + 1.95, y: py + 2.15, w: 3, h: 0.25, fontFace: MONO, fontSize: 9, color: C.accent2 }, "Link");
    const hl = [["nós", C.background1, SY.musgo], ["método", C.accent2, SY.musgo], ["clientes", C.accent1, SY.mineral], ["bastidores", C.background2, SY.musgo]];
    hl.forEach((h, i) => {
      const x = px + 0.4 + i * 1.4;
      dot(s, x, py + 2.65, 0.85, h[1], `Destaque ${i + 1}`);
      img(s, h[2], x + 0.17, py + 2.65 + 0.425 - 0.51 / h[2].ratio / 2, 0.51, `Destaque ${i + 1} símbolo`);
      T(s, h[0], { x: x - 0.2, y: py + 3.6, w: 1.25, h: 0.3, fontSize: 11, color: C.background1, align: "center" }, `Destaque ${i + 1} nome`);
    });
    const gs = 1.5, gx = SW - MX - (3 * gs + 0.2);
    const tiles = [
      [HEX.mineral, "Sua voz, no ritmo certo.", HEX.musgo], [HEX.rosa, "", null], [HEX.musgo, "Raiz antes de vitrine.", HEX.mineral],
      [HEX.verde, "3 sinais de que você precisa aparecer", HEX.musgo], [HEX.mineral, "", null], [HEX.ardosia, "Verdade acima de pose.", HEX.mineral],
      [HEX.musgo, "", null], [HEX.mineral, "A gente cuida da técnica.", HEX.musgo], [HEX.rosa, "Constância vence inspiração.", HEX.mineral],
    ];
    tiles.forEach((t, i) => {
      const x = gx + (i % 3) * (gs + 0.1), y = 2.15 + Math.floor(i / 3) * (gs + 0.1);
      rect(s, x, y, gs, gs, t[0], `Post ${i + 1}`);
      if (t[1]) T(s, t[1], { x: x + 0.15, y: y + 0.15, w: gs - 0.3, h: gs - 0.3, fontFace: LIGHT, fontSize: 13, color: t[2], valign: "bottom" }, `Post ${i + 1} texto`);
      else img(s, i === 1 ? SY.mineral : i === 4 ? SY.rosa : SY.verde, x + gs / 2 - 0.4, y + gs / 2 - 0.4 / SY.musgo.ratio, 0.8, `Post ${i + 1} símbolo`);
    });
  }

  // ================= 18 · Stories
  {
    const s = slide("Claro", "Universo visual");
    s.addText("15 · universo visual · stories".toUpperCase(), { placeholder: "kicker" });
    s.addText("Telas verticais, mesma voz.", { placeholder: "title" });
    const pw = 2.55, ph = 4.5, gap = 0.6, x0 = (SW - (3 * pw + 2 * gap)) / 2, y = 2.2;
    const phones = [
      { f: HEX.mineral, t: "Seja nas redes quem você já é na vida.", tc: HEX.musgo, l: LG.musgo },
      { photo: true },
      { f: HEX.musgo, t: "A gente cuida da técnica. Você aparece com verdade.", tc: HEX.mineral, l: LG.verde },
    ];
    phones.forEach((p, i) => {
      const x = x0 + i * (pw + gap);
      box(s, x - 0.08, y - 0.08, pw + 0.16, ph + 0.16, HEX.musgo, `Celular ${i + 1}`, 0.3);
      if (p.photo) {
        s.addImage(on({ path: path.join(IMG, "diogo-rafa-sq.png"), x, y, w: pw, h: ph, sizing: { type: "cover", w: pw, h: ph }, altText: "Diogo e Rafael" }, "Story foto"));
        rect(s, x, y + ph - 1.0, pw, 1.0, HEX.musgo, "Story faixa");
        img(s, LG.mineral, x + 0.25, y + ph - 0.65, 1.1, "Story logo");
        T(s, "@MEOV.MP", { x: x + 1.45, y: y + ph - 0.62, w: 0.9, h: 0.25, fontFace: MONO, fontSize: 8, color: C.accent2, align: "right" }, "Story arroba");
      } else {
        box(s, x, y, pw, ph, p.f, `Tela ${i + 1}`, 0.24);
        T(s, p.t, { x: x + 0.3, y: y + 1.2, w: pw - 0.6, h: 2.0, fontFace: LIGHT, fontSize: 20, color: p.tc, valign: "middle", align: "center" }, `Tela ${i + 1} texto`);
        rect(s, x + pw / 2 - 0.4, y + 3.4, 0.4, 0.06, C.accent1, `Tela ${i + 1} rosa`);
        rect(s, x + pw / 2, y + 3.4, 0.4, 0.06, C.accent2, `Tela ${i + 1} verde`);
        img(s, p.l, x + pw / 2 - 0.55, y + ph - 0.6, 1.1, `Tela ${i + 1} logo`);
      }
    });
  }

  // ================= 19 · Site + LinkedIn
  {
    const s = slide("Pedra", "Universo visual");
    s.addText("16 · universo visual · site e linkedin".toUpperCase(), { placeholder: "kicker" });
    s.addText("meov.com.br", { placeholder: "title" });
    const bx = MX, by = 2.15, bw = 7.4, bh = 4.55;
    box(s, bx, by, bw, bh, C.background1, "Navegador", 0.14);
    [HEX.rosa, HEX.verde, HEX.pedra].forEach((c, i) => dot(s, bx + 0.25 + i * 0.22, by + 0.2, 0.13, c, `Botão ${i + 1}`));
    box(s, bx + 1.1, by + 0.13, 3.0, 0.27, C.background2, "Endereço", 0.13);
    T(s, "meov.com.br", { x: bx + 1.25, y: by + 0.13, w: 2.8, h: 0.27, fontFace: MONO, fontSize: 9, color: C.text2, valign: "middle" }, "URL");
    imgH(s, LG.musgo, bx + 0.4, by + 0.75, 0.24, "Site logo");
    T(s, "MÉTODO   SERVIÇOS   CLIENTES   CONTATO", { x: bx + 3.4, y: by + 0.75, w: 3.7, h: 0.25, fontFace: MONO, fontSize: 8, charSpacing: 1, color: C.text2, align: "right" }, "Menu");
    T(s, "Sua voz, no ritmo certo.", { x: bx + 0.4, y: by + 1.5, w: 4.6, h: 1.4, fontFace: LIGHT, fontSize: 36, charSpacing: -0.8, color: C.text1, valign: "top" }, "Site título");
    T(s, "Estratégia, audiovisual e gestão das redes para especialistas que já são referência.", { x: bx + 0.4, y: by + 2.95, w: 4.2, h: 0.6, fontSize: 12, color: C.text2, valign: "top" }, "Site texto");
    box(s, bx + 0.4, by + 3.7, 1.75, 0.45, C.accent1, "Site botão", 0.22);
    T(s, "fala com a Meov", { x: bx + 0.4, y: by + 3.7, w: 1.75, h: 0.45, fontSize: 11, color: C.background1, align: "center", valign: "middle" }, "Site botão texto");
    s.addImage(on({ path: path.join(IMG, "diogo-rafa-sq.png"), x: bx + 5.2, y: by + 1.3, w: 1.85, h: 2.85, sizing: { type: "cover", w: 1.85, h: 2.85 }, altText: "Diogo e Rafael" }, "Site foto"));
    const lx = 8.45, lw = SW - MX - lx;
    box(s, lx, 2.15, lw, 4.55, C.background1, "LinkedIn", 0.14);
    rect(s, lx, 2.15, lw, 1.3, HEX.musgo, "LinkedIn banner");
    T(s, "Ajudamos especialistas a colocar sua voz no mundo digital.", { x: lx + 1.65, y: 2.35, w: lw - 1.85, h: 0.9, fontFace: LIGHT, fontSize: 13, color: C.background1, valign: "middle" }, "Banner texto");
    dot(s, lx + 0.3, 2.95, 1.15, C.accent2, "LinkedIn avatar");
    img(s, SY.musgo, lx + 0.47, 2.95 + 0.575 - 0.81 / SY.musgo.ratio / 2, 0.81, "LinkedIn símbolo");
    T(s, "Meov", { x: lx + 0.3, y: 4.25, w: lw - 0.6, h: 0.4, fontSize: 18, color: C.text1 }, "LinkedIn nome");
    T(s, "Ajudamos especialistas a colocar sua voz no mundo digital, com estratégia, audiovisual, gestão das redes e acolhimento.", { x: lx + 0.3, y: 4.65, w: lw - 0.6, h: 0.85, fontSize: 11, color: C.text2, valign: "top" }, "LinkedIn slogan");
    T(s, "MARKETING · GRANDE SÃO PAULO", { x: lx + 0.3, y: 5.55, w: lw - 0.6, h: 0.25, fontFace: MONO, fontSize: 8, charSpacing: 1, color: C.text2 }, "LinkedIn setor");
    box(s, lx + 0.3, 5.95, 1.2, 0.38, C.text1, "LinkedIn seguir", 0.19);
    T(s, "seguir", { x: lx + 0.3, y: 5.95, w: 1.2, h: 0.38, fontSize: 11, color: C.background1, align: "center", valign: "middle" }, "LinkedIn seguir texto");
  }

  // ================= 20 · Papelaria
  {
    const s = slide("Claro", "Universo visual");
    s.addText("17 · universo visual · papelaria".toUpperCase(), { placeholder: "kicker" });
    s.addText("Cartões que já dizem quem somos.", { placeholder: "title" });
    const cards = [
      { f: HEX.verde, l: LG.musgo, tc: HEX.musgo, back: false },
      { f: HEX.rosa, l: SY.musgo, tc: HEX.musgo, back: true },
      { f: HEX.musgo, l: LG.mineral, tc: HEX.mineral, back: false },
      { f: HEX.mineral, l: SY.rosa, tc: HEX.musgo, back: true, line: true },
    ];
    const cw = 3.5, ch = 2.0;
    cards.forEach((c, i) => {
      const x = MX + (i % 2) * (cw + 0.35) + (i >= 2 ? 0.6 : 0), y = 2.3 + Math.floor(i / 2) * (ch + 0.3);
      box(s, x, y, cw, ch, c.f, `Cartão ${i + 1}`, 0.1, c.line ? { color: HEX.pedra, width: 1 } : undefined);
      if (c.back) { const w = 1.3; img(s, c.l, x + cw / 2 - w / 2, y + ch / 2 - w / c.l.ratio / 2, w, `Cartão ${i + 1} símbolo`); }
      else {
        img(s, c.l, x + 0.3, y + 0.3, 1.5, `Cartão ${i + 1} logo`);
        T(s, "Diogo & Rafael\nMeov Marca Pessoal", { x: x + 0.3, y: y + 1.05, w: cw - 0.6, h: 0.5, fontSize: 10, color: c.tc, valign: "top" }, `Cartão ${i + 1} nomes`);
        T(s, "@MEOV.MP · MEOV.COM.BR", { x: x + 0.3, y: y + ch - 0.4, w: cw - 0.6, h: 0.22, fontFace: MONO, fontSize: 8, charSpacing: 0.5, color: c.tc }, `Cartão ${i + 1} contatos`);
      }
    });
    const rx = 8.85, rw = SW - MX - rx;
    T(s, [
      { text: "Frente e verso em par", options: { fontSize: 17, color: C.text1, breakLine: true } },
      { text: "Verde com rosa, musgo com mineral. A soma das duas cores aparece até no bolso.", options: { fontSize: 13, color: C.text2, breakLine: true } },
      { text: " ", options: { fontSize: 10, breakLine: true } },
      { text: "Pouco texto", options: { fontSize: 17, color: C.text1, breakLine: true } },
      { text: "Nome, marca e contatos em Geist Mono. O resto é respiro.", options: { fontSize: 13, color: C.text2 } },
    ], { x: rx, y: 2.3, w: rw, h: 4.3, valign: "top", lineSpacingMultiple: 1.12 }, "Texto papelaria");
  }

  // ================= 21 · Propostas
  {
    const s = slide("Pedra", "Universo visual");
    s.addText("18 · universo visual · propostas".toUpperCase(), { placeholder: "kicker" });
    s.addText("A identidade em uso real.", { placeholder: "title" });
    T(s, "A primeira proposta na nova identidade foi para Fernando Lohmann: diagnóstico anotado dos canais, Brandbook, plano de 6 meses e investimento, tudo na mesma voz.", { x: MX, y: 1.95, w: CW, h: 0.55, ...body({ fontSize: 14, color: C.text1 }) }, "Intro propostas");
    const th = ["p-01.png", "p-03.png", "p-09.png", "p-13.png"];
    const tw = (CW - 0.3) / 2, thh = tw * 9 / 16 * 0.5;
    const w2 = (CW - 0.3) / 2, h2 = (6.75 - 2.7 - 0.2) / 2, w3 = h2 * 16 / 9;
    th.forEach((f, i) => {
      const x = MX + (i % 2) * (w3 + 0.3), y = 2.7 + Math.floor(i / 2) * (h2 + 0.2);
      s.addImage(on({ path: path.join(IMG, f), x, y, w: w3, h: h2, altText: "Slide da proposta Fernando Lohmann" }, `Proposta ${i + 1}`));
    });
    const rx = MX + 2 * w3 + 0.6, rw = SW - MX - rx;
    if (rw > 1.5) T(s, [
      { text: "Capa com retrato", options: { fontSize: 14, color: C.text1, breakLine: true } },
      { text: "o cliente é a marca", options: { fontSize: 12, color: C.text2, breakLine: true } },
      { text: " ", options: { fontSize: 10, breakLine: true } },
      { text: "Prints anotados", options: { fontSize: 14, color: C.text1, breakLine: true } },
      { text: "onde vamos mexer", options: { fontSize: 12, color: C.text2, breakLine: true } },
      { text: " ", options: { fontSize: 10, breakLine: true } },
      { text: "Rosa e verde em par", options: { fontSize: 14, color: C.text1, breakLine: true } },
      { text: "em todos os slides", options: { fontSize: 12, color: C.text2 } },
    ], { x: rx, y: 2.7, w: rw, h: 3.9, valign: "top" }, "Notas propostas");
  }

  // ================= 22 · Encerramento
  pres.addSection({ title: "Encerramento" });
  {
    const s = slide("Escuro", "Encerramento");
    imgH(s, LG.mineral, MX, 0.95, 0.38, "Logotipo");
    T(s, "Obrigado.", { x: MX, y: 2.2, w: 6.5, h: 1.0, fontFace: LIGHT, fontSize: 54, charSpacing: -1.2, color: C.background1 }, "Obrigado");
    T(s, "Você já tem a voz. A gente ajuda você a colocá-la no mundo.", { x: MX, y: 3.25, w: 6.2, h: 1.0, fontFace: LIGHT, fontSize: 22, color: C.accent2, valign: "top" }, "Frase");
    T(s, "Marca criada e cuidada pela própria Meov.", { x: MX, y: 4.45, w: 6.2, h: 0.35, fontSize: 14, color: C.background2 }, "Autoria");
    T(s, "MEOV.COM.BR  ·  @MEOV.MP", { x: MX, y: 6.55, w: 5, h: 0.3, fontFace: MONO, fontSize: 10, charSpacing: 1.5, color: C.background2 }, "Contatos");
    const d = 3.4, x = 8.3, y = 1.25;
    ring(s, x, y, d, "Anel");
    s.addImage(on({ path: path.join(IMG, "diogo-rafa-sq.png"), x: x + 0.2, y: y + 0.2, w: d - 0.4, h: d - 0.4, rounding: true, altText: "Diogo e Rafael, da Meov" }, "Foto Diogo e Rafael"));
    T(s, "Diogo & Rafael", { x: x - 0.8, y: y + d + 0.3, w: d + 1.6, h: 0.45, fontSize: 20, color: C.background1, align: "center" }, "Nomes");
    T(s, "Meov Marca Pessoal", { x: x - 0.8, y: y + d + 0.75, w: d + 1.6, h: 0.35, fontSize: 14, color: C.background2, align: "center" }, "Marca");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("ok", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
