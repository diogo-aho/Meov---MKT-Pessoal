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


  // modelo de apresentação: 12 páginas em miniatura, no estilo e na paleta atuais
  async function modeloApresentacao() {
    const s = slide("Livre", "Universo visual");
    s.background = { color: HEX.musgo };
    const gap = 0.22, tw = (CW - 3 * gap) / 4, th = tw * 9 / 16;
    const y0 = (SH - (3 * th + 2 * gap)) / 2;
    const k = tw / SW; // escala da miniatura
    const pages = [
      { bg: HEX.musgo, f: (x, y) => { img(s, LG.mineral, x + 0.22, y + th / 2 - 0.25, 1.1, "m1 logo"); T(s, "APRESENTAÇÃO · 2026", { x: x + 0.22, y: y + th - 0.32, w: 2, h: 0.15, fontFace: MONO, fontSize: 5, charSpacing: 1, color: HEX.pedra }, "m1 rótulo"); } },
      { bg: HEX.pedra, f: (x, y) => { T(s, "POSICIONAMENTO", { x: x + 0.2, y: y + 0.2, w: 2, h: 0.15, fontFace: MONO, fontSize: 5, charSpacing: 1, color: HEX.ardosia }, "m2 rótulo"); T(s, "Não somos uma agência de marketing.", { x: x + 0.2, y: y + 0.45, w: tw - 0.5, h: 0.6, fontFace: LIGHT, fontSize: 12, color: HEX.musgo, valign: "top" }, "m2 frase"); } },
      { bg: HEX.mineral, f: (x, y) => { T(s, "Sua voz,\nno ritmo certo.", { x: x + 0.2, y: y + th - 0.75, w: tw - 0.4, h: 0.55, fontFace: LIGHT, fontSize: 12, color: HEX.musgo, valign: "bottom" }, "m3 frase"); dot(s, x + tw - 0.3, y + 0.2, 0.08, HEX.rosa, "m3 ponto"); } },
      { bg: HEX.pedra, f: (x, y) => { ["01", "02", "03"].forEach((n, i) => T(s, n, { x: x + 0.2 + i * 0.8, y: y + 0.3, w: 0.7, h: 0.4, fontFace: LIGHT, fontSize: 16, color: HEX.musgo }, `m4 ${n}`)); T(s, "ESTRATÉGIA · AUDIOVISUAL · REDES", { x: x + 0.2, y: y + th - 0.35, w: tw - 0.4, h: 0.15, fontFace: MONO, fontSize: 5, charSpacing: 0.5, color: HEX.ardosia }, "m4 rótulo"); } },
      { bg: HEX.musgo, f: (x, y) => { T(s, "Raiz antes de vitrine.", { x: x + 0.2, y: y + 0.3, w: tw - 0.4, h: 0.5, fontFace: LIGHT, fontSize: 12, color: HEX.mineral, valign: "top" }, "m5 frase"); rect(s, x + 0.2, y + th - 0.3, 0.4, 0.03, HEX.verde, "m5 linha"); } },
      { bg: HEX.mineral, f: (x, y) => { T(s, "PALETA", { x: x + 0.2, y: y + 0.2, w: 2, h: 0.15, fontFace: MONO, fontSize: 5, charSpacing: 1, color: HEX.ardosia }, "m6 rótulo"); let xx = x + 0.2; const bw = tw - 0.4; [[HEX.mineral, 34], [HEX.pedra, 32], [HEX.musgo, 27], [HEX.ardosia, 6], [HEX.verde, 0.8], [HEX.rosa, 0.2]].forEach(([c, p], i) => { const w = Math.max(bw * p / 100, 0.03); s.addShape(pres.shapes.RECTANGLE, on({ x: xx, y: y + th - 0.55, w, h: 0.32, fill: { color: c }, line: { color: i === 0 ? HEX.pedra : c, width: 0.5 } }, `m6 cor ${i + 1}`)); xx += w; }); } },
      { bg: HEX.pedra, f: (x, y) => { const h = 0.42; imgH(s, SY.musgo, x + (tw - h * SY.musgo.ratio) / 2, y + (th - h) / 2, h, "m7 símbolo"); } },
      { bg: HEX.mineral, f: (x, y) => { T(s, "Leve. Divertido.\nDo seu jeito.", { x: x + 0.2, y: y + 0.3, w: tw - 0.4, h: 0.6, fontFace: LIGHT, fontSize: 12, color: HEX.musgo, valign: "top" }, "m8 frase"); } },
      { bg: HEX.musgo, f: (x, y) => { T(s, [{ text: "Você já tem a voz.", options: { color: HEX.mineral, breakLine: true } }, { text: "A gente ajuda a colocá-la no mundo.", options: { color: HEX.verde, fontSize: 8 } }], { x: x + 0.2, y: y + 0.3, w: tw - 0.4, h: 0.7, fontFace: LIGHT, fontSize: 12, valign: "top" }, "m9 frase"); } },
      { bg: HEX.pedra, f: (x, y) => { T(s, "6 MESES", { x: x + 0.2, y: y + 0.3, w: tw - 0.4, h: 0.4, fontFace: MONO, fontSize: 16, color: HEX.musgo }, "m10 número"); T(s, "um ciclo para a sua voz ganhar ritmo", { x: x + 0.2, y: y + th - 0.4, w: tw - 0.4, h: 0.2, fontSize: 6, color: HEX.ardosia }, "m10 texto"); } },
      { bg: HEX.verde, f: (x, y) => { T(s, "Seja nas redes quem você já é na vida.", { x: x + 0.2, y: y + th - 0.75, w: tw - 0.4, h: 0.55, fontFace: LIGHT, fontSize: 12, color: HEX.musgo, valign: "bottom" }, "m11 frase"); } },
      { bg: HEX.musgo, f: (x, y) => { T(s, "obrigado.", { x, y: y + th / 2 - 0.2, w: tw, h: 0.3, fontFace: LIGHT, fontSize: 11, color: HEX.mineral, align: "center" }, "m12 frase"); T(s, "MARCA PESSOAL", { x, y: y + th / 2 + 0.12, w: tw, h: 0.15, fontFace: MONO, fontSize: 5, charSpacing: 1.5, color: HEX.pedra, align: "center" }, "m12 rótulo"); dot(s, x + tw / 2 - 0.035, y + th / 2 - 0.38, 0.07, HEX.rosa, "m12 ponto"); } },
    ];
    pages.forEach((pg, i) => {
      const x = MX + (i % 4) * (tw + gap), y = y0 + Math.floor(i / 4) * (th + gap);
      s.addShape(pres.shapes.RECTANGLE, on({ x, y, w: tw, h: th, fill: { color: pg.bg }, line: { color: HEX.ardosia, width: 0.5 } }, `Página ${i + 1}`));
      pg.f(x, y);
      if (i !== 6 && i !== 0) imgH(s, pg.bg === HEX.musgo ? SY.mineral : SY.musgo, x + tw - 0.12 - 0.07 * SY.musgo.ratio, y + th - 0.17, 0.07, `Página ${i + 1} símbolo`);
    });
    void k;
  }

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
      { n: "Branco mineral", h: "#FAF8F4", p: "34%", x: 0, y: 0, w: 6.87, hh: 4.2, f: HEX.mineral, t: HEX.musgo },
      { n: "Pedra líquen", h: "#D8DACD", p: "32%", x: 6.87, y: 0, w: 6.463, hh: 4.2, f: HEX.pedra, t: HEX.musgo },
      { n: "Preto musgo", h: "#1C211E", p: "27%", x: 0, y: 4.2, w: 5.2, hh: 3.3, f: HEX.musgo, t: HEX.mineral },
      { n: "Cinza ardósia verde", h: "#59625A", p: "6%", x: 5.2, y: 4.2, w: 3.4, hh: 3.3, f: HEX.ardosia, t: HEX.mineral },
      { n: "Verde vibrante", h: "#DBFF00", p: "0,8%", x: 8.6, y: 4.2, w: 2.55, hh: 3.3, f: HEX.verde, t: HEX.musgo },
      { n: "Rosa choque", h: "#FF19B1", p: "0,2%", x: 11.15, y: 4.2, w: 2.183, hh: 3.3, f: HEX.rosa, t: HEX.mineral },
    ];
    blocks.forEach((b, i) => {
      rect(s, b.x, b.y, b.w, b.hh, b.f, `Cor ${i + 1}`);
      T(s, b.n, { x: b.x + 0.4, y: b.y + 0.4, w: b.w - 0.6, h: 0.45, fontFace: LIGHT, fontSize: i < 2 ? 24 : 18, color: b.t }, `Cor ${i + 1} nome`);
      T(s, b.h, { x: b.x + 0.4, y: b.y + 0.85, w: b.w - 0.6, h: 0.3, fontFace: MONO, fontSize: 11, charSpacing: 1, color: b.t }, `Cor ${i + 1} hex`);
      T(s, b.p, { x: b.x + 0.4, y: b.y + b.hh - 1.25, w: b.w - 0.6, h: 0.95, fontFace: LIGHT, fontSize: i < 2 ? 60 : 44, color: b.t, valign: "bottom" }, `Cor ${i + 1} proporção`);
    });
  }
  full("Identidade visual", "aplicacoes/pagina-06.jpg", "Textura 1", "Textura Meov: degradê com granulado");
  {
    const s = slide("Escuro", "Identidade visual");
    s.background = { path: path.join(IMG, "aplicacoes", "pagina-06.jpg") };
    s.addText("o porquê das cores".toUpperCase(), { placeholder: "kicker" });
    s.addText("Muito respiro, dois pontos de energia.", { placeholder: "title" });
    const why = [
      { c: [HEX.mineral, HEX.pedra], n: "Branco mineral e pedra líquen", d: "A base clara e calma, que dá espaço para a pessoa aparecer." },
      { c: [HEX.musgo, HEX.ardosia], n: "Preto musgo e ardósia", d: "Profundidade e elegância, nos textos e nas texturas." },
      { c: [HEX.verde], n: "Verde vibrante", d: "A energia do movimento e da constância." },
      { c: [HEX.rosa], n: "Rosa choque", d: "A emoção e o acolhimento: o coração que pulsa no símbolo." },
    ];
    why.forEach((w, i) => {
      const y = 2.3 + i * 1.0;
      w.c.forEach((c, k) => s.addShape(pres.shapes.OVAL, on({ x: MX + k * 0.42, y, w: 0.7, h: 0.7, fill: { color: c }, line: { color: HEX.ardosia, width: 0.75 } }, `Amostra ${i + 1}.${k + 1}`)));
      T(s, w.n, { x: MX + 1.35, y: y - 0.02, w: 5.6, h: 0.36, fontSize: 17, color: C.background1 }, `Porquê ${i + 1} nome`);
      T(s, w.d, { x: MX + 1.35, y: y + 0.35, w: 5.6, h: 0.5, fontSize: 13, color: C.background2, valign: "top" }, `Porquê ${i + 1} texto`);
    });
    dot(s, 8.6, 2.6, 1.5, C.accent2, "Círculo verde");
    T(s, "+", { x: 10.15, y: 2.95, w: 0.6, h: 0.8, fontFace: LIGHT, fontSize: 40, color: C.background1, align: "center" }, "Mais");
    dot(s, 10.85, 3.0, 0.7, C.accent1, "Círculo rosa");
    T(s, "A Meov é a soma dos dois.", { x: 8.6, y: 4.45, w: 4.0, h: 0.5, fontFace: LIGHT, fontSize: 22, color: C.background1 }, "Soma");
    T(s, "O verde dá energia e ritmo. O rosa aparece só no detalhe que acolhe e emociona, sobre o branco e os tons de cinza.", { x: 8.6, y: 5.0, w: 4.0, h: 0.9, fontSize: 13, color: C.background2, valign: "top" }, "Soma texto");
  }
  quote("Identidade visual", "Tipografia");
  {
    const s = slide("Livre", "Identidade visual");
    s.background = { color: HEX.musgo };
    T(s, "TEXTOS", { x: 0, y: 2.0, w: SW, h: 0.35, fontFace: MONO, fontSize: 11, charSpacing: 3, color: C.background2, align: "center" }, "Rótulo");
    T(s, [{ text: "Manrope", options: { fontFace: "Manrope" } }, { text: " + ", options: { fontFace: LIGHT } }, { text: "Geist Mono", options: { fontFace: MONO } }], { x: 0, y: 2.55, w: SW, h: 0.8, fontSize: 34, color: C.accent2, align: "center" }, "Fontes");
    T(s, "ABCD abcd", { x: 0, y: 3.45, w: SW, h: 0.8, fontFace: LIGHT, fontSize: 38, color: C.background1, align: "center" }, "Alfabeto");
    T(s, "01234@#$%&", { x: 0, y: 4.3, w: SW, h: 0.8, fontFace: MONO, fontSize: 32, color: C.background2, align: "center" }, "Números");
  }
  {
    const s = slide("Claro", "Identidade visual");
    s.addText("tipografia".toUpperCase(), { placeholder: "kicker" });
    s.addText([{ text: "Duas famílias. ", options: { color: C.text1 } }, { text: "Uma voz.", options: { color: C.text2 } }], { placeholder: "title" });
    const w = (CW - 0.3) / 2;
    box(s, MX, 2.2, w, 4.4, C.background2, "Cartão Manrope");
    label(s, "títulos e textos", MX + 0.45, 2.55, 5, C.text2, "Rótulo Manrope");
    T(s, "Aa", { x: MX + 0.45, y: 2.9, w: 3, h: 1.5, fontFace: LIGHT, fontSize: 80, color: C.text1 }, "Aa Manrope");
    T(s, "Manrope", { x: MX + 0.45, y: 4.45, w: 5, h: 0.6, fontFace: LIGHT, fontSize: 26, color: C.text1 }, "Nome Manrope");
    T(s, "Acolhedora e moderna. Leve nos títulos, porque o destaque vem do espaço, não do negrito.", { x: MX + 0.45, y: 5.1, w: w - 0.9, h: 0.8, ...body({ fontSize: 13, color: C.text1 }) }, "Porquê Manrope");
    T(s, "LIGHT 300 · REGULAR 400 · MEDIUM 500 · SEMIBOLD 600", { x: MX + 0.45, y: 6.15, w: w - 0.9, h: 0.26, fontFace: MONO, fontSize: 9, charSpacing: 1, color: C.text2 }, "Pesos Manrope");
    const x2 = MX + w + 0.3;
    box(s, x2, 2.2, w, 4.4, C.text1, "Cartão Geist");
    label(s, "dados e rótulos", x2 + 0.45, 2.55, 5, C.accent2, "Rótulo Geist");
    T(s, "Aa", { x: x2 + 0.45, y: 2.9, w: 3, h: 1.5, fontFace: MONO, fontSize: 72, color: C.background1 }, "Aa Geist");
    T(s, "Geist Mono", { x: x2 + 0.45, y: 4.45, w: 5, h: 0.6, fontFace: MONO, fontSize: 24, color: C.background1 }, "Nome Geist");
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

  // grades de variações: logotipo (100%) e símbolo (25%)
  const TEX = path.join(IMG, "aplicacoes", "pagina-06.jpg");
  const combos = [
    [HEX.musgo, "mineral"], [HEX.musgo, "verde"], [HEX.musgo, "rosa"], ["tex", "mineral"],
    [HEX.pedra, "musgo"], [HEX.mineral, "musgo"], [HEX.mineral, "rosa"], [HEX.ardosia, "mineral"],
    [HEX.verde, "musgo"], [HEX.rosa, "musgo"], [HEX.ardosia, "verde"], ["tex", "verde"],
  ];
  for (const g of [{ set: LG, k: 0.68, n: "logotipo", m: "100%" }, { set: SY, k: 0.42, n: "símbolo", m: "25%" }]) {
    const s = slide("Livre", "Identidade visual");
    s.background = { color: HEX.pedra };
    const q = 2.25, gap = 0.1, gw = 4 * q + 3 * gap, gh = 3 * q + 2 * gap;
    const gx = SW - (SH - gh) / 2 - gw, gy = (SH - gh) / 2;
    label(s, "variações", MX, 0.95, 3, C.text2, "Rótulo");
    T(s, g.n[0].toUpperCase() + g.n.slice(1), { x: MX, y: 1.3, w: 2.6, h: 0.6, fontFace: LIGHT, fontSize: 30, charSpacing: -0.6, color: C.text1 }, "Título");
    T(s, g.m, { x: MX, y: 1.95, w: 2.6, h: 0.4, fontFace: MONO, fontSize: 14, color: C.text1 }, "Medida");
    T(s, "Sobre os fundos da paleta e a textura Meov. Verde e rosa como cor da marca ou do fundo, sempre com contraste.", { x: MX, y: 5.4, w: 2.4, h: 1.3, fontSize: 11, color: C.text2, valign: "bottom", lineSpacingMultiple: 1.15 }, "Nota");
    combos.forEach(([bg, c], i) => {
      const x = gx + (i % 4) * (q + gap), y = gy + Math.floor(i / 4) * (q + gap);
      if (bg === "tex") s.addImage(on({ path: TEX, x, y, w: q, h: q, sizing: { type: "cover", w: q, h: q } }, `Fundo ${i + 1}`));
      else if (bg === HEX.pedra) s.addShape(pres.shapes.RECTANGLE, on({ x, y, w: q, h: q, fill: { color: bg }, line: { color: HEX.ardosia, width: 0.5 } }, `Fundo ${i + 1}`));
      else rect(s, x, y, q, q, bg, `Fundo ${i + 1}`);
      const art = g.set[c], w = q * g.k;
      img(s, art, x + (q - w) / 2, y + (q - w / art.ratio) / 2, w, `${g.n} ${i + 1}`);
    });
  }

  full("Identidade visual", "aplicacoes/pagina-16.jpg", "Logotipo sobre textura", "Logotipo verde sobre a textura Meov");
  full("Identidade visual", "aplicacoes/pagina-17.jpg", "Logotipo com onda", "Logotipo branco sobre a onda rosa");
  // ================= Universo visual
  pres.addSection({ title: "Universo visual" });
  quote("Universo visual", "Universo Visual");
  const alts = { 22: "Perfil do Instagram @meov.mp", 23: "Stories", 24: "Navegador com meov.com.br", 27: "Página do LinkedIn", 30: "Papelaria", 31: "Cartões de visita", 32: "Proposta no tablet", 36: "Mídia indoor", 38: "Moletom", 41: "Boné", 43: "Garrafa" };
  const app = (p) => full("Universo visual", `aplicacoes/pagina-${p}.jpg`, `Aplicação ${p}`, alts[p]);
  [22, 23, 24].forEach(app);
  full("Universo visual", "aplicacoes/notebook.jpg", "Site no notebook", "Site meov.com.br no notebook");
  full("Universo visual", "aplicacoes/linkedin.jpg", "LinkedIn", "Página do LinkedIn com banner da voz ganhando o mundo");
  [27, 30, 31, 32].forEach(app);
  await modeloApresentacao();
  app(36);
  if (fs.existsSync(path.join(IMG, "aplicacoes", "estudio.jpg"))) full("Universo visual", "aplicacoes/estudio.jpg", "Estúdio Meov", "Gravação no estúdio Meov: entrevista, diretor e filmmaker");
  [38, 41, 43].forEach(app);

  // ================= Obrigado
  pres.addSection({ title: "Encerramento" });
  {
    const s = slide("Livre", "Encerramento");
    s.background = { color: HEX.musgo };
    const lw = 5.4, lh = lw / LG.mineral.ratio, ly = 2.2;
    img(s, LG.mineral, (SW - lw) / 2, ly, lw, "Logotipo");
    T(s, "Seja nas redes quem você já é na vida. Obrigado por chegar até aqui.", { x: 0, y: ly + lh + 0.65, w: SW, h: 0.4, fontFace: LIGHT, fontSize: 16, color: C.background2, align: "center" }, "Frase");
    rect(s, SW / 2 - 0.3, ly + lh + 1.35, 0.6, 0.025, HEX.verde, "Linha verde");
    T(s, "MARCA PESSOAL", { x: 0, y: ly + lh + 1.5, w: SW, h: 0.3, fontFace: MONO, fontSize: 11, charSpacing: 4, color: C.background1, align: "center" }, "Marca pessoal");
    dot(s, SW / 2 - 0.04, ly + lh + 1.95, 0.08, C.accent1, "Ponto rosa");
    T(s, "meov.com.br  ·  @meov.mp", { x: 0, y: 6.75, w: SW, h: 0.3, fontFace: MONO, fontSize: 9, charSpacing: 1, color: C.text2, align: "center" }, "Contatos");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("ok", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
