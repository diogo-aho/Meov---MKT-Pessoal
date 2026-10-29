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
    { text: { text: "Sua voz, no ritmo certo.", options: { x: SW - MX - 5, y: 0.36, w: 5, h: 0.28, fontSize: 11, color: dark ? C.background2 : C.text2, align: "right", margin: 0 } } },
  ];
  const ph = (dark, kc) => [
    { placeholder: { options: { name: "kicker", type: "body", x: MX, y: 0.95, w: 9, h: 0.3, fontFace: MONO, fontSize: 11, charSpacing: 1.5, color: kc, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 1.25, w: CW, h: 0.8, fontFace: LIGHT, fontSize: 34, charSpacing: -0.6, color: dark ? C.background1 : C.text1, margin: 0, valign: "top", align: "left" }, text: "" } },
  ];
  pres.defineSlideMaster({ title: "Claro", background: { color: C.background1 }, objects: [...head(false), ...ph(false, C.accent1)], slideNumber: { x: SW - MX - 0.6, y: 6.95, w: 0.6, h: 0.25, fontFace: MONO, fontSize: 9, color: C.text2, align: "right", margin: 0 } });
  pres.defineSlideMaster({ title: "Pedra", background: { color: C.background2 }, objects: [...head(false), ...ph(false, C.text1)], slideNumber: { x: SW - MX - 0.6, y: 6.95, w: 0.6, h: 0.25, fontFace: MONO, fontSize: 9, color: C.text2, align: "right", margin: 0 } });
  pres.defineSlideMaster({ title: "Escuro", background: { color: C.text1 }, objects: [...head(true), ...ph(true, C.accent2)], slideNumber: { x: SW - MX - 0.6, y: 6.95, w: 0.6, h: 0.25, fontFace: MONO, fontSize: 9, color: C.background2, align: "right", margin: 0 } });
  pres.defineSlideMaster({ title: "Livre", background: { color: C.background1 }, objects: [] });
  pres.defineSlideMaster({ title: "Marca", background: { color: C.background2 }, objects: [...head(false)] });

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


  const full = (sec, file, name, alt) => { const s = slide("Livre", sec); s.addImage(on({ path: path.join(IMG, file), x: 0, y: 0, w: SW, h: SH, altText: alt || name }, name)); return s; };
  const quote = (sec, t) => { const s = slide("Marca", sec); T(s, `"${t}"`, { x: 0, y: SH / 2 - 0.5, w: SW, h: 1.0, fontFace: LIGHT, fontSize: 40, charSpacing: -0.8, color: C.text1, align: "center", valign: "middle" }, "Título"); return s; };
  const bigTitle = (sec, t, sub) => { const s = slide("Marca", sec); T(s, t, { x: MX, y: 5.1, w: 10, h: 1.1, fontFace: LIGHT, fontSize: 54, charSpacing: -1.2, color: C.text1 }, "Título"); if (sub) T(s, sub, { x: MX, y: 6.2, w: 10, h: 0.4, fontSize: 15, color: C.text2 }, "Subtítulo"); return s; };

  // ================= 1 · Capa
  pres.addSection({ title: "Abertura" });
  {
    const s = slide("Livre", "Abertura");
    s.background = { color: HEX.pedra };
    T(s, "Meov", { x: MX, y: 0.7, w: 4, h: 0.8, fontFace: LIGHT, fontSize: 40, charSpacing: -0.8, color: C.text1 }, "Marca");
    T(s, [{ text: "Cliente: ", options: { fontFace: "Manrope" } }, { text: "Meov Marca Pessoal.", options: { breakLine: true } }, { text: "Serviço: ", options: { fontFace: "Manrope" } }, { text: "Estratégia & Identidade Visual." }], { x: SW - MX - 4.2, y: 0.75, w: 4.2, h: 0.6, fontFace: LIGHT, fontSize: 13, color: C.text1 }, "Cliente e serviço");
    T(s, '"Apresentação"', { x: 0, y: SH / 2 - 0.5, w: SW, h: 1.0, fontFace: LIGHT, fontSize: 40, color: C.text2, align: "center", valign: "middle" }, "Título");
    T(s, "meov.com.br\n@meov.mp", { x: MX, y: 6.5, w: 4, h: 0.55, fontSize: 12, color: C.text1 }, "Contatos");
    T(s, "2026", { x: SW / 2 - 1, y: 6.75, w: 2, h: 0.3, fontSize: 12, color: C.text1, align: "center" }, "Ano");
    imgH(s, SY.musgo, SW - MX - 0.42 * SY.musgo.ratio, 6.55, 0.42, "Símbolo");
  }
  {
    const s = slide("Marca", "Abertura");
    T(s, [{ text: "Vamos conhecer o time que ajuda especialistas a ", options: { color: C.text1 } }, { text: "colocar a própria voz no mundo.", options: { color: C.accent1 } }], { x: MX, y: 4.3, w: 9.5, h: 2.3, fontFace: LIGHT, fontSize: 40, charSpacing: -1, valign: "bottom" }, "Frase");
  }

  // ================= Estratégia
  pres.addSection({ title: "Estratégia" });
  bigTitle("Estratégia", "Estratégia", "Posicionamento, serviços e tom de voz.");
  {
    const s = full("Estratégia", "aplicacoes/pagina-06.jpg", "Textura");
    label(s, "posicionamento", MX, 0.95, 5, C.accent2, "Rótulo");
    T(s, "Não somos uma agência de marketing. Somos um time que apoia técnica e emocionalmente quem decide colocar a própria voz no mundo.", { x: MX, y: 1.5, w: 9.2, h: 3.0, fontFace: LIGHT, fontSize: 36, charSpacing: -0.8, color: C.background1, valign: "top", lineSpacingMultiple: 1.02 }, "Manifesto");
    T(s, "Da estratégia à câmera e à gestão das redes, com acolhimento do começo ao fim.", { x: MX, y: 6.2, w: 9, h: 0.5, fontSize: 15, color: C.accent2 }, "Complemento");
  }
  {
    const s = slide("Claro", "Estratégia");
    s.addText("o que fazemos".toUpperCase(), { placeholder: "kicker" });
    s.addText("Três frentes, um mesmo time.", { placeholder: "title" });
    const f = [
      { n: "Estratégia", c: C.accent1, d: "O Método Meov: diagnóstico, imersão presencial e um Brandbook que organiza quem você é, para quem fala e como se diferencia.", tags: "MÉTODO MEOV · BRANDBOOK · 6 MESES" },
      { n: "Audiovisual", c: C.accent2, d: "Direção de imagem, ensaio, gravação e edição, em dias de gravação semanais, quinzenais ou mensais. O fino trato que nasceu com a Meov.", tags: "ENSAIO · VLOGCAST · REELS" },
      { n: "Gestão das redes", c: C.text1, d: "Instagram, LinkedIn e Site/LP com calendário, roteiros, legendas, design, publicação e relatórios. Você grava e aprova.", tags: "INSTAGRAM · LINKEDIN · SITE/LP" },
    ];
    const w = (CW - 0.6) / 3;
    f.forEach((it, i) => {
      const x = MX + i * (w + 0.3);
      s.addShape(pres.shapes.LINE, on({ x, y: 2.35, w, h: 0, line: { color: HEX.musgo, width: 0.75 } }, `Linha ${i + 1}`));
      T(s, String(i + 1).padStart(2, "0"), { x, y: 2.55, w: 1.5, h: 1.1, fontFace: LIGHT, fontSize: 64, charSpacing: -1.5, color: C.text1 }, `Número ${i + 1}`);
      dot(s, x + 1.55, 2.85, 0.2, it.c, `Ponto ${i + 1}`);
      T(s, it.n, { x, y: 3.85, w, h: 0.55, fontFace: LIGHT, fontSize: 26, color: C.text1 }, `Frente ${i + 1}`);
      T(s, it.d, { x, y: 4.5, w: w - 0.1, h: 1.4, ...body({ fontSize: 14 }) }, `Frente ${i + 1} texto`);
      T(s, it.tags, { x, y: 6.15, w, h: 0.3, fontFace: MONO, fontSize: 9, charSpacing: 1, color: C.text1 }, `Frente ${i + 1} tags`);
    });
  }
  {
    const s = full("Estratégia", "aplicacoes/pagina-07.jpg", "Textura");
    label(s, "para quem", MX, 0.95, 5, C.accent2, "Rótulo público");
    T(s, "Empresários e especialistas que já são referência fora das redes e ainda não conseguiram se posicionar no digital.", { x: MX, y: 1.35, w: 6.0, h: 1.6, fontFace: LIGHT, fontSize: 22, color: C.background1, valign: "top" }, "Público");
    label(s, "princípios", MX, 4.6, 5, C.accent1, "Rótulo princípios");
    T(s, [{ text: "Raiz antes de vitrine.", options: { breakLine: true } }, { text: "Verdade acima de pose.", options: { breakLine: true } }, { text: "Constância no lugar de inspiração." }], { x: MX, y: 4.95, w: 8, h: 1.8, fontFace: LIGHT, fontSize: 30, charSpacing: -0.6, color: C.background1, valign: "top" }, "Princípios");
  }
  {
    const s = slide("Marca", "Estratégia");
    label(s, "tom de voz", MX, 0.95, 5, C.accent1, "Rótulo");
    T(s, [{ text: "Acolhedores.", options: { breakLine: true } }, { text: "Claros.", options: { breakLine: true } }, { text: "Verdadeiros.", options: { breakLine: true } }, { text: "Nunca guru.", options: { color: C.accent1 } }], { x: MX, y: 1.35, w: 6, h: 4.2, fontFace: LIGHT, fontSize: 52, charSpacing: -1.4, color: C.text1, valign: "top", lineSpacingMultiple: 0.95 }, "Tom");
    const rx = 7.2, rw = SW - MX - rx;
    const ex = [
      ["Acolhedores", "falamos com quem tem medo de se expor como quem entende esse medo."],
      ["Claros", "frases curtas, palavras simples, nada de jargão de marketing."],
      ["Verdadeiros", "nada de fórmula, promessa de viral ou personagem."],
    ];
    ex.forEach((e, i) => {
      const y = 1.45 + i * 0.95;
      s.addShape(pres.shapes.LINE, on({ x: rx, y, w: rw, h: 0, line: { color: HEX.ardosia, width: 0.75 } }, `Linha ${i + 1}`));
      T(s, [{ text: e[0] + "  ", options: { color: C.text1 } }, { text: e[1], options: { color: C.text2 } }], { x: rx, y: y + 0.12, w: rw, h: 0.75, fontSize: 14, valign: "top" }, `Pilar ${i + 1}`);
    });
    label(s, "assim", rx, 4.45, 3, "4F7A45", "Rótulo assim");
    T(s, "Você já tem a voz. A gente ajuda você a colocá-la no mundo.", { x: rx, y: 4.75, w: rw, h: 0.75, fontFace: LIGHT, fontSize: 18, color: C.text1, valign: "top" }, "Assim");
    label(s, "não assim", rx, 5.65, 3, "B23A3A", "Rótulo não assim");
    T(s, "VIRALIZE EM 7 DIAS!  ·  o hack secreto para bombar", { x: rx, y: 5.95, w: rw, h: 0.45, fontSize: 15, color: C.text2, strike: "sngStrike" }, "Não assim");
  }

  // ================= Identidade visual
  pres.addSection({ title: "Identidade visual" });
  bigTitle("Identidade visual", "Identidade Visual");
  quote("Identidade visual", "Cores");
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
  }
  full("Identidade visual", "aplicacoes/pagina-06.jpg", "Textura 1", "Textura Meov: degradê com granulado");
  full("Identidade visual", "aplicacoes/pagina-07.jpg", "Textura 2", "Textura Meov: degradê com onda");
  {
    const s = slide("Escuro", "Identidade visual");
    s.background = { path: path.join(IMG, "aplicacoes", "pagina-06.jpg") };
    s.addText("o porquê das cores".toUpperCase(), { placeholder: "kicker" });
    s.addText("Muito respiro, dois pontos de energia.", { placeholder: "title" });
    const why = [
      { c: [HEX.mineral, HEX.pedra], n: "Branco mineral e pedra líquen", d: "A base clara e calma, que dá espaço para a pessoa aparecer." },
      { c: [HEX.musgo, HEX.ardosia], n: "Preto musgo e ardósia", d: "Profundidade e elegância, nos textos e nas texturas." },
      { c: [HEX.rosa], n: "Rosa choque", d: "A emoção: o coração que pulsa no símbolo." },
      { c: [HEX.verde], n: "Verde vibrante", d: "A energia do movimento e da constância." },
    ];
    why.forEach((w, i) => {
      const y = 2.3 + i * 1.0;
      w.c.forEach((c, k) => s.addShape(pres.shapes.OVAL, on({ x: MX + k * 0.42, y, w: 0.7, h: 0.7, fill: { color: c }, line: { color: HEX.ardosia, width: 0.75 } }, `Amostra ${i + 1}.${k + 1}`)));
      T(s, w.n, { x: MX + 1.35, y: y - 0.02, w: 5.6, h: 0.36, fontSize: 17, color: C.background1 }, `Porquê ${i + 1} nome`);
      T(s, w.d, { x: MX + 1.35, y: y + 0.35, w: 5.6, h: 0.5, fontSize: 13, color: C.background2, valign: "top" }, `Porquê ${i + 1} texto`);
    });
    dot(s, 8.6, 2.6, 1.5, C.accent1, "Círculo rosa");
    T(s, "+", { x: 10.15, y: 2.95, w: 0.6, h: 0.8, fontFace: LIGHT, fontSize: 40, color: C.background1, align: "center" }, "Mais");
    dot(s, 10.8, 2.6, 1.5, C.accent2, "Círculo verde");
    T(s, "A Meov é a soma dos dois.", { x: 8.6, y: 4.45, w: 4.0, h: 0.5, fontFace: LIGHT, fontSize: 22, color: C.background1 }, "Soma");
    T(s, "Rosa e verde sempre em par, na mesma medida, sobre o branco e os tons de cinza.", { x: 8.6, y: 5.0, w: 4.0, h: 0.9, fontSize: 13, color: C.background2, valign: "top" }, "Soma texto");
  }
  quote("Identidade visual", "Tipografia");
  {
    const s = slide("Livre", "Identidade visual");
    s.background = { color: HEX.musgo };
    T(s, "Textos", { x: 0, y: 1.55, w: SW, h: 0.6, fontFace: LIGHT, fontSize: 26, color: C.accent2, align: "center" }, "Rótulo");
    T(s, [{ text: "Manrope", options: { fontFace: "Manrope" } }, { text: " + ", options: { fontFace: LIGHT } }, { text: "Geist Mono", options: { fontFace: MONO } }], { x: 0, y: 2.2, w: SW, h: 1.1, fontSize: 52, color: C.accent2, align: "center" }, "Fontes");
    T(s, "ABCD abcd", { x: 0, y: 3.4, w: SW, h: 1.1, fontFace: LIGHT, fontSize: 60, color: C.accent2, align: "center" }, "Alfabeto");
    T(s, "01234@#$%&", { x: 0, y: 4.6, w: SW, h: 1.1, fontFace: MONO, fontSize: 56, color: C.accent2, align: "center" }, "Números");
  }
  {
    const s = slide("Claro", "Identidade visual");
    s.addText("tipografia".toUpperCase(), { placeholder: "kicker" });
    s.addText([{ text: "Duas famílias. ", options: { color: C.text1 } }, { text: "Uma voz.", options: { color: C.text2 } }], { placeholder: "title" });
    const w = (CW - 0.3) / 2;
    box(s, MX, 2.2, w, 4.4, C.background2, "Cartão Manrope");
    label(s, "títulos e textos", MX + 0.45, 2.55, 5, C.text2, "Rótulo Manrope");
    T(s, "Aa", { x: MX + 0.45, y: 2.9, w: 3, h: 1.5, fontFace: LIGHT, fontSize: 96, color: C.text1 }, "Aa Manrope");
    T(s, "Manrope", { x: MX + 0.45, y: 4.45, w: 5, h: 0.6, fontFace: LIGHT, fontSize: 30, color: C.text1 }, "Nome Manrope");
    T(s, "Acolhedora e moderna. Leve nos títulos, porque o destaque vem do espaço, não do negrito.", { x: MX + 0.45, y: 5.1, w: w - 0.9, h: 0.8, ...body({ fontSize: 13, color: C.text1 }) }, "Porquê Manrope");
    T(s, "LIGHT 300 · REGULAR 400 · MEDIUM 500 · SEMIBOLD 600", { x: MX + 0.45, y: 6.15, w: w - 0.9, h: 0.26, fontFace: MONO, fontSize: 9, charSpacing: 1, color: C.text2 }, "Pesos Manrope");
    const x2 = MX + w + 0.3;
    box(s, x2, 2.2, w, 4.4, C.text1, "Cartão Geist");
    label(s, "dados e rótulos", x2 + 0.45, 2.55, 5, C.accent2, "Rótulo Geist");
    T(s, "Aa", { x: x2 + 0.45, y: 2.9, w: 3, h: 1.5, fontFace: MONO, fontSize: 88, color: C.background1 }, "Aa Geist");
    T(s, "Geist Mono", { x: x2 + 0.45, y: 4.45, w: 5, h: 0.6, fontFace: MONO, fontSize: 28, color: C.background1 }, "Nome Geist");
    T(s, "Precisa e técnica. Só em números, valores, datas, @, sites e selos.", { x: x2 + 0.45, y: 5.1, w: w - 0.9, h: 0.8, fontSize: 13, color: C.background2, valign: "top", lineSpacingMultiple: 1.15 }, "Porquê Geist");
    T(s, "REGULAR 400", { x: x2 + 0.45, y: 6.15, w: w - 0.9, h: 0.26, fontFace: MONO, fontSize: 9, charSpacing: 1, color: C.background2 }, "Pesos Geist");
  }

  // ================= Logotipo e símbolo
  quote("Identidade visual", "Logotipo");
  full("Identidade visual", "malha-logotipo.png", "Malha do logotipo", "Malha de construção do logotipo meov");
  quote("Identidade visual", "Símbolo");
  full("Identidade visual", "malha-simbolo.png", "Malha do símbolo", "Malha de construção do símbolo");
  full("Identidade visual", "aplicacoes/pagina-14.jpg", "Conceito do símbolo", "Símbolo sobre linha de batimento: m + v + onda");
  {
    const s = slide("Marca", "Identidade visual");
    label(s, "o porquê do símbolo", MX, 0.95, 5, C.accent1, "Rótulo");
    T(s, "Uma marca que pulsa.", { x: MX, y: 1.3, w: 8, h: 0.9, fontFace: LIGHT, fontSize: 40, charSpacing: -1, color: C.text1 }, "Título");
    T(s, [
      { text: "O símbolo une o m e o v de \"meov\" a uma linha de batimento cardíaco: experiências que fazem o coração pulsar.", options: { breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 8 } },
      { text: "Na Meov Marca Pessoal, a mesma linha é também a onda sonora da voz de cada cliente e o ritmo da presença constante nas redes." },
    ], { x: MX, y: 2.6, w: 6.2, h: 2.6, fontSize: 17, color: C.text1, valign: "top", lineSpacingMultiple: 1.15 }, "Texto");
    T(s, "Sua voz, no ritmo certo.", { x: MX, y: 5.6, w: 6, h: 0.6, fontFace: LIGHT, fontSize: 26, color: C.accent1 }, "Assinatura");
    imgH(s, SY.musgo, 7.9, 2.9, 1.9, "Símbolo");
  }

  // ================= Resultado
  quote("Identidade visual", "Resultado");
  full("Identidade visual", "aplicacoes/pagina-16.jpg", "Logotipo sobre textura", "Logotipo verde sobre a textura Meov");
  full("Identidade visual", "aplicacoes/pagina-17.jpg", "Logotipo com onda", "Logotipo branco sobre a onda rosa");
  for (const [bg, nm] of [[HEX.musgo, "musgo"], [HEX.ardosia, "ardósia"]]) {
    const s = slide("Livre", "Identidade visual");
    s.background = { color: bg };
    img(s, LG.mineral, SW / 2 - 1.6, SH / 2 - 1.6 / LG.mineral.ratio, 3.2, `Logotipo sobre ${nm}`);
  }
  {
    const s = slide("Livre", "Identidade visual");
    s.background = { color: HEX.musgo };
    const lw = 4.4, sh = lw / LG.mineral.ratio, sw = sh * SY.mineral.ratio;
    const x0 = (SW - (lw + 1.4 + sw)) / 2, y0 = SH / 2 - sh / 2 - 0.2;
    img(s, LG.mineral, x0, y0, lw, "Logotipo 100");
    imgH(s, SY.mineral, x0 + lw + 1.4, y0, sh, "Símbolo");
    s.addShape(pres.shapes.LINE, on({ x: x0, y: y0 + sh + 0.35, w: lw, h: 0, line: { color: HEX.ardosia, width: 0.75 } }, "Régua logo"));
    s.addShape(pres.shapes.LINE, on({ x: x0 + lw + 1.4, y: y0 + sh + 0.35, w: sw, h: 0, line: { color: HEX.ardosia, width: 0.75 } }, "Régua símbolo"));
    T(s, "100%", { x: x0, y: y0 + sh + 0.5, w: lw, h: 0.3, fontFace: MONO, fontSize: 10, color: C.background2, align: "center" }, "Medida logo");
    T(s, "25%", { x: x0 + lw + 1.4, y: y0 + sh + 0.5, w: sw, h: 0.3, fontFace: MONO, fontSize: 10, color: C.background2, align: "center" }, "Medida símbolo");
  }

  // ================= Universo visual
  pres.addSection({ title: "Universo visual" });
  quote("Universo visual", "Universo Visual");
  const alts = { 22: "Perfil do Instagram @meov.mp", 23: "Stories", 24: "Navegador com meov.com.br", 25: "Site no notebook", 26: "Página do LinkedIn", 27: "Página do LinkedIn", 28: "Cartaz", 29: "Fita com o pattern", 30: "Papelaria", 31: "Cartões de visita", 32: "Proposta no tablet", 33: "Slides de proposta", 34: "Ecobag verde", 35: "Ecobag rosa", 36: "Mídia indoor", 37: "Moletom, costas", 38: "Moletom, frente", 39: "Gorro com símbolo", 40: "Gorro com pattern", 41: "Boné", 42: "Óculos", 43: "Garrafa", 44: "Camiseta verde", 45: "Camisetas ardósia", 46: "Camisetas verdes", 47: "Camisetas rosa" };
  for (let p = 22; p <= 47; p++) {
    full("Universo visual", `aplicacoes/pagina-${p}.jpg`, `Aplicação ${p}`, alts[p]);
    if (p === 33) {
      const s = slide("Escuro", "Universo visual");
      s.addText("propostas".toUpperCase(), { placeholder: "kicker" });
      s.addText("A identidade em uma proposta real.", { placeholder: "title" });
      const th = ["p-01.png", "p-03.png", "p-09.png", "p-13.png"];
      const tw = (CW - 0.3) / 2, thh = tw * 9 / 16;
      const sc = Math.min(1, (SH - 2.4 - 0.75 - 0.25) / (2 * thh));
      th.forEach((f, i) => {
        const w = tw * sc, h = thh * sc;
        const x = (SW - (2 * w + 0.3)) / 2 + (i % 2) * (w + 0.3), y = 2.3 + Math.floor(i / 2) * (h + 0.25);
        s.addImage(on({ path: path.join(IMG, f), x, y, w, h, altText: "Slide da proposta Fernando Lohmann" }, `Proposta ${i + 1}`));
      });
    }
  }

  // ================= Obrigado
  pres.addSection({ title: "Encerramento" });
  {
    const s = slide("Livre", "Encerramento");
    s.background = { color: HEX.pedra };
    imgH(s, SY.musgo, MX, 0.7, 0.32, "Símbolo");
    const d = 2.5, x = SW / 2 - d / 2, y = 1.15;
    ring(s, x, y, d, "Anel");
    s.addImage(on({ path: path.join(IMG, "diogo-rafa-sq.png"), x: x + 0.15, y: y + 0.15, w: d - 0.3, h: d - 0.3, rounding: true, altText: "Diogo e Rafael, da Meov" }, "Foto Diogo e Rafael"));
    T(s, '"Obrigado"', { x: 0, y: 4.0, w: SW, h: 0.9, fontFace: LIGHT, fontSize: 40, color: C.text2, align: "center", valign: "middle" }, "Obrigado");
    T(s, "Diogo & Rafael  |  Meov Marca Pessoal", { x: 0, y: 4.9, w: SW, h: 0.4, fontSize: 15, color: C.text1, align: "center" }, "Assinatura");
    T(s, "meov.com.br\n@meov.mp", { x: MX, y: 6.5, w: 4, h: 0.55, fontSize: 12, color: C.text1 }, "Contatos");
    T(s, "2026", { x: SW / 2 - 1, y: 6.75, w: 2, h: 0.3, fontSize: 12, color: C.text1, align: "center" }, "Ano");
    T(s, "Sua voz, no ritmo certo.", { x: SW - MX - 4, y: 6.75, w: 4, h: 0.3, fontSize: 12, color: C.text1, align: "right" }, "Assinatura final");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("ok", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
