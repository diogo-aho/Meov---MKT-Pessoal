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
const IMG = process.env.IMG_DIR || path.join(REPO, "propostas/fernando-lohmann/imagens");
const OUT = process.argv[2] || "proposta.pptx";

const HEX = { mineral: "FAF8F4", pedra: "D8DACD", musgo: "1C211E", ardosia: "59625A", rosa: "FF19B1", lima: "DBFF00", white: "FFFFFF" };
const MONO = "Geist Mono";
const LIGHT = "Manrope Light";
const THEME = {
  name: "Meov MP", headFontFace: LIGHT, bodyFontFace: "Manrope",
  colors: { dk1: HEX.musgo, lt1: HEX.mineral, dk2: HEX.ardosia, lt2: HEX.pedra, accent1: HEX.rosa, accent2: HEX.lima, accent3: HEX.pedra, accent4: HEX.ardosia, accent5: HEX.musgo, accent6: HEX.white, hlink: HEX.rosa, folHlink: HEX.ardosia },
};

// Nome da função de Diogo e Rafael no set de gravação
const PAPEL_GRAVACAO = "Direção de Narrativa";

async function svgArt(file, color) {
  let s = fs.readFileSync(path.join(BRAND, file), "utf8");
  s = s.replace(/<rect class="fil0"[^>]*\/>/, "")
       .replace(/\.fil1 \{fill:[^}]*\}/, `.fil1 {fill:#${color}}`)
       .replace(/\.fil2 \{fill:[^;]*;/, `.fil2 {fill:#${color};`);
  const buf = await sharp(Buffer.from(s), { density: 300 }).trim().png().toBuffer();
  const m = await sharp(buf).metadata();
  return { data: "image/png;base64," + buf.toString("base64"), ratio: m.width / m.height };
}
async function icon(name, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(tb[name], { color: "#" + color, size, strokeWidth: 1.4 }));
  const buf = await sharp(Buffer.from(svg)).resize(size, size).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
const fmt = (n) => "R$ " + n.toLocaleString("pt-BR");

// ---------- preços ----------
const IMPL = [
  { n: "Método Meov e Brandbook", av: 4000, pc: 3000, es: 1650 },
  { n: "Ensaio de marca", av: 4000, pc: 3000, es: 1650 },
  { n: "Reformulação do Instagram", av: 2500, pc: 2000, es: 825 },
  { n: "Reformulação do LinkedIn", av: 2500, pc: 2000, es: 825 },
  { n: "Site/LP", av: 3000, pc: 2000, es: 1650 },
];
const MENS = [
  { n: "Gravações de Vlogcast", av: 4000, pc: 3000, es: 1600 },
  { n: "Edição de 12 reels", av: 4200, pc: 3000, es: 1600 },
  { n: "8 carrosséis", av: 2400, pc: 2000, es: 1070 },
  { n: "4 posts no LinkedIn", av: 1400, pc: 1000, es: 530 },
  { n: "Pautas e roteiros", av: 1400, pc: 1000, es: 530 },
  { n: "Gestão, relatório e reuniões", av: 2600, pc: 2000, es: 1070 },
];
const sum = (a, k) => a.reduce((t, x) => t + x[k], 0);
const T_IMPL = { av: sum(IMPL, "av"), pc: sum(IMPL, "pc"), es: sum(IMPL, "es") };
const T_MENS = { av: sum(MENS, "av"), pc: sum(MENS, "pc"), es: sum(MENS, "es") };
const CYCLE = { av: T_IMPL.av + 6 * T_MENS.av, pc: T_IMPL.pc + 6 * T_MENS.pc, es: T_IMPL.es + 6 * T_MENS.es };
if (T_IMPL.pc !== 12000 || T_MENS.pc !== 12000 || CYCLE.es !== 45000) throw new Error("preços não fecham: " + JSON.stringify({ T_IMPL, T_MENS, CYCLE }));

(async () => {
  const LOGO = {
    musgo: await svgArt("logo/logotipo/meov-logotipo-musgo-sem-fundo.svg", HEX.musgo),
    mineral: await svgArt("logo/logotipo/meov-logotipo-mineral-sem-fundo.svg", HEX.mineral),
  };

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.title = "A marca Fernando Lohmann · Proposta Meov";
  pres.author = "Diogo &amp; Rafael | Meov Marca Pessoal";
  pres.company = "Meov";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  const SW = 13.333, MX = 0.75, CW = SW - 2 * MX;

  // Fundos: pedra é o principal; branco mineral e preto musgo dividem o restante.
  // Rosa só em detalhes estratégicos; verde um pouco mais presente que o rosa.
  const header = (dark) => [{ text: { text: "PROPOSTA · FERNANDO LOHMANN · 2026", options: { x: SW - MX - 5, y: 0.38, w: 5, h: 0.25, fontFace: MONO, fontSize: 9, charSpacing: 1.5, color: dark ? C.background2 : C.text2, align: "right", margin: 0 } } }];
  const titlePh = (dark) => ({ placeholder: { options: { name: "title", type: "title", x: MX, y: 1.25, w: CW, h: 0.8, fontFace: LIGHT, fontSize: 34, charSpacing: -0.6, color: dark ? C.background1 : C.text1, margin: 0, valign: "top", align: "left" }, text: "" } });
  const pinkDot = await sharp(Buffer.from(`<svg xmlns="http://www.w3.org/2000/svg" width="64" height="64"><circle cx="32" cy="32" r="32" fill="#${HEX.rosa}"/></svg>`)).png().toBuffer();
  const kickerDot = { image: { x: MX, y: 1.05, w: 0.1, h: 0.1, data: "image/png;base64," + pinkDot.toString("base64") } };
  const kickerPh = (color) => ({ placeholder: { options: { name: "kicker", type: "body", x: MX + 0.24, y: 0.95, w: 9, h: 0.3, fontFace: MONO, fontSize: 11, charSpacing: 1.5, color, margin: 0 }, text: "" } });
  const slideNum = (color) => ({ x: MX, y: 0.38, w: 0.6, h: 0.25, fontFace: MONO, fontSize: 9, color, margin: 0 });
  const lightMaster = (title, bg) => pres.defineSlideMaster({
    title, background: { color: bg },
    objects: [...header(false), { image: { x: SW - MX - 0.95, y: 6.88, w: 0.95, h: 0.95 / LOGO.musgo.ratio, data: LOGO.musgo.data } }, kickerDot, kickerPh(C.text2), titlePh(false)],
    slideNumber: slideNum(C.text2),
  });
  lightMaster("Claro", C.background1);
  lightMaster("Pedra", C.background2);
  pres.defineSlideMaster({
    title: "Escuro", background: { color: C.text1 },
    objects: [...header(true), { image: { x: SW - MX - 0.95, y: 6.88, w: 0.95, h: 0.95 / LOGO.mineral.ratio, data: LOGO.mineral.data } }, kickerDot, kickerPh(C.background2), titlePh(true)],
    slideNumber: slideNum(C.background2),
  });
  pres.defineSlideMaster({
    title: "Capa", background: { color: C.text1 }, objects: [],
  });

  const on = (o, n) => { o.objectName = n; return o; };
  const T = (s, text, opts, name) => s.addText(text, on({ isTextBox: true, margin: 0, fontFace: "Manrope", ...opts }, name));
  const label = (s, text, x, y, w, color, name, align = "left") => T(s, text.toUpperCase(), { x, y, w, h: 0.26, fontFace: MONO, fontSize: 10, charSpacing: 1.5, color, align }, name);
  const card = (s, x, y, w, h, fill, name, line) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, on({ x, y, w, h, rectRadius: 0.16, fill: { color: fill }, line: line || { color: fill, width: 0 } }, name));
  const dot = (s, x, y, d, fill, name) => s.addShape(pres.shapes.OVAL, on({ x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 } }, name));
  const numDot = (s, n, x, y, d, fill, tc, name, fs = 11) => { dot(s, x, y, d, fill, name); T(s, String(n), { x, y, w: d, h: d, fontFace: MONO, fontSize: fs, color: tc, align: "center", valign: "middle" }, name + " número"); };
  const iconDot = async (s, iname, x, y, d, fill, fg, name) => {
    dot(s, x, y, d, fill, name + " fundo");
    s.addImage(on({ data: await icon(iname, fg), x: x + d * 0.24, y: y + d * 0.24, w: d * 0.52, h: d * 0.52 }, name + " ícone"));
  };
  // ícone padrão: círculo musgo com traço verde
  const mark = (s, iname, x, y, d, name) => iconDot(s, iname, x, y, d, C.text1, HEX.lima, name);
  // moldura branca das fotos
  const ring = (s, x, y, d, name) => {
    s.addShape(pres.shapes.BLOCK_ARC, on({ x, y, w: d, h: d, angleRange: [0, 359.9], arcThicknessRatio: 0.07, fill: { color: C.background1 }, line: { color: C.background1, width: 0 } }, name));
  };
  const slide = (master, sec) => pres.addSlide({ masterName: master, sectionTitle: sec });
  const light = (sec) => slide("Claro", sec);
  const stone = (sec) => slide("Pedra", sec);
  const dark = (sec) => slide("Escuro", sec);
  const body = (extra = {}) => ({ fontSize: 14, color: C.text2, valign: "top", lineSpacingMultiple: 1.15, ...extra });
  // anotações sobre prints: verde sobre prints escuros, musgo sobre prints claros
  const annotate = (s, img, marks, tone) => {
    const lc = tone === "lima" ? HEX.lima : HEX.musgo;
    const fill = tone === "lima" ? C.accent2 : C.text1, tc = tone === "lima" ? C.text1 : C.background1;
    marks.forEach((m, i) => {
      const x = img.x + m.r[0] * img.k, y = img.y + m.r[1] * img.k, w = (m.r[2] - m.r[0]) * img.k, h = (m.r[3] - m.r[1]) * img.k;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, on({ x, y, w, h, rectRadius: 0.06, fill: { color: HEX.white, transparency: 100 }, line: { color: lc, width: 2.25 } }, `Marca ${i + 1}`));
      numDot(s, i + 1, x - 0.16, y - 0.16, 0.34, fill, tc, `Marca ${i + 1} selo`, 10);
    });
  };
  const legend = (s, items, x, y, w, rowH, name) => {
    items.forEach((it, i) => {
      const yy = y + i * rowH;
      numDot(s, i + 1, x, yy, 0.34, C.text1, C.background1, `${name} ${i + 1}`, 10);
      T(s, [{ text: it.t, options: { color: C.text1, breakLine: true, fontSize: 14 } }, { text: it.d, options: { color: C.text2, fontSize: 12 } }], { x: x + 0.5, y: yy - 0.06, w: w - 0.5, h: rowH - 0.05, valign: "top", lineSpacingMultiple: 1.1 }, `${name} ${i + 1} texto`);
    });
  };

  // ============ 1 · Capa
  pres.addSection({ title: "Abertura" });
  {
    const s = pres.addSlide({ masterName: "Capa", sectionTitle: "Abertura" });
    s.addImage(on({ data: LOGO.mineral.data, x: MX, y: 0.75, w: 1.5, h: 1.5 / LOGO.mineral.ratio }, "Logotipo Meov"));
    label(s, "proposta de marca pessoal · outubro 2026", MX, 2.55, 8, C.accent2, "Rótulo");
    T(s, [{ text: "A marca", options: { color: C.background2, breakLine: true } }, { text: "Fernando Lohmann", options: { color: C.background1 } }], { x: MX, y: 2.95, w: 7.6, h: 2.1, fontFace: LIGHT, fontSize: 54, charSpacing: -1.2, valign: "top", lineSpacingMultiple: 0.95 }, "Título");
    T(s, "Cinco décadas de trajetória, agora com a presença digital à altura.", { x: MX, y: 5.2, w: 6.8, h: 0.5, fontSize: 16, color: C.background2 }, "Subtítulo");
    const d = 4.1, x = 8.45, y = 1.7;
    ring(s, x, y, d, "Moldura");
    const g = 0.22;
    s.addImage(on({ path: path.join(IMG, "fernando.png"), x: x + g, y: y + g, w: d - 2 * g, h: d - 2 * g, rounding: true, altText: "Foto de Fernando Lohmann" }, "Foto Fernando"));
  }

  // ============ 2 · Ponto de partida
  {
    const s = stone("Abertura");
    s.addText("01 · o ponto de partida".toUpperCase(), { placeholder: "kicker" });
    s.addText("Poucos sabem quem você é. Muita gente precisa saber.", { placeholder: "title" });
    T(s, [
      { text: "São cinco décadas entre indústria, energia, infraestrutura, executive search, governança e diplomacia econômica entre Brasil e Singapura. Quem trabalha com você sabe o valor disso. Fora desse círculo, pouca gente sabe quem você é e o que você faz.", options: { breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 8 } },
      { text: "Enquanto isso, há acionistas, conselhos, CEOs em transição e herdeiros precisando exatamente da sua experiência, sem saber que você existe. A marca pessoal é a ponte entre eles e você." },
    ], { x: MX, y: 2.25, w: 6.1, h: 2.55, ...body({ fontSize: 14 }) }, "Texto contexto");

    const gx = 7.25, gw = SW - MX - gx;
    card(s, gx, 2.25, gw, 2.55, C.background1, "Quadro barras");
    [{ l: "reputação no seu meio", p: 0.92, c: C.text1 }, { l: "presença no digital hoje", p: 0.15, c: C.accent1 }].forEach((b, i) => {
      const y = 2.5 + i * 0.82;
      label(s, b.l, gx + 0.35, y, gw - 0.7, C.text2, `Barra ${i + 1} rótulo`);
      card(s, gx + 0.35, y + 0.3, gw - 0.7, 0.34, C.background2, `Barra ${i + 1} trilho`);
      card(s, gx + 0.35, y + 0.3, (gw - 0.7) * b.p, 0.34, b.c, `Barra ${i + 1} valor`);
    });
    card(s, gx + 0.35, 4.15, gw - 0.7, 0.45, C.accent2, "Faixa legado");
    T(s, "Deixe sua marca. Registre seu legado.", { x: gx + 0.55, y: 4.15, w: gw - 1.1, h: 0.45, fontSize: 14, color: C.text1, valign: "middle" }, "Texto legado");

    const outs = [
      { n: "Clareza e verdade", d: "Saber o que dizer, para quem e por quê, sempre a partir da sua história real.", i: "TbCompass" },
      { n: "Constância", d: "Presença toda semana, com produção organizada e sem depender da agenda livre.", i: "TbCalendarEvent" },
      { n: "Segurança", d: "Aparecer sendo quem você é, com direção de câmera e um time ao lado.", i: "TbShieldCheck" },
      { n: "Legado", d: "Sua trajetória organizada e acessível, falando por você e para quem vem depois.", i: "TbSeedling" },
    ];
    const ow = (CW - 0.6) / 4;
    for (let i = 0; i < outs.length; i++) {
      const x = MX + i * (ow + 0.2), y = 5.1;
      await mark(s, outs[i].i, x, y, 0.5, `Ganho ${i + 1}`);
      T(s, outs[i].n, { x: x + 0.65, y: y + 0.07, w: ow - 0.7, h: 0.36, fontSize: 16, color: C.text1 }, `Ganho ${i + 1} nome`);
      T(s, outs[i].d, { x: x + 0.65, y: y + 0.45, w: ow - 0.7, h: 1.0, ...body({ fontSize: 12 }) }, `Ganho ${i + 1} texto`);
    }
  }

  // ============ 3 · Instagram hoje
  pres.addSection({ title: "Diagnóstico" });
  {
    const s = light("Diagnóstico");
    s.addText("02 · diagnóstico · instagram".toUpperCase(), { placeholder: "kicker" });
    s.addText("Um perfil fechado, pronto para começar.", { placeholder: "title" });
    const img = { x: MX + 0.15, y: 2.45, w: 6.9 };
    img.k = img.w / 1070; img.h = 326 * img.k;
    card(s, MX, 2.3, 7.2, img.h + 0.3, C.text1, "Moldura print");
    s.addImage(on({ path: path.join(IMG, "instagram.png"), x: img.x, y: img.y, w: img.w, h: img.h, altText: "Print do Instagram @fal50" }, "Print Instagram"));
    annotate(s, img, [
      { r: [20, 30, 262, 278] }, { r: [295, 30, 440, 85] }, { r: [295, 90, 530, 120] }, { r: [295, 132, 770, 172] }, { r: [295, 192, 400, 228] },
    ], "lima");
    T(s, "HOJE  ·  PERFIL PRIVADO  ·  9 POSTS  ·  204 SEGUIDORES  ·  2.040 SEGUINDO  ·  SEM BIO  ·  SEM LINK  ·  SEM DESTAQUES", { x: MX, y: img.y + img.h + 0.3, w: 7.2, h: 0.5, fontFace: MONO, fontSize: 9, charSpacing: 0.5, color: C.text2, valign: "top" }, "Dados hoje");
    T(s, "O perfil existe, mas está fechado e sem apresentação. Vamos abri-lo, transformá-lo em conta de criador e dar a partida com uma grade nova, pensada a partir do seu Brandbook.", { x: MX, y: 5.4, w: 7.2, h: 1.1, ...body({ fontSize: 14 }) }, "Texto Instagram");
    legend(s, [
      { t: "Foto do ensaio", d: "retrato novo, com direção de imagem" },
      { t: "@ e conta de criador", d: "perfil aberto, categoria e botões de contato" },
      { t: "Nome que aparece na busca", d: "nome + especialidade, para ser encontrado" },
      { t: "Proporção e grade", d: "conteúdo novo e seguidores qualificados" },
      { t: "Bio, link e destaques", d: "quem você é, o que faz e para onde ir" },
    ], 8.4, 2.3, SW - MX - 8.4, 0.82, "Legenda Instagram");
  }

  // ============ 4 · LinkedIn hoje
  {
    const s = light("Diagnóstico");
    s.addText("03 · diagnóstico · linkedin".toUpperCase(), { placeholder: "kicker" });
    s.addText("Um perfil sólido, que ainda fala por outra marca.", { placeholder: "title" });
    const img = { x: MX, y: 2.2, w: 6.6 };
    img.k = img.w / 1606; img.h = 980 * img.k;
    card(s, MX - 0.05, 2.15, img.w + 0.1, img.h + 0.1, C.background2, "Moldura print");
    s.addImage(on({ path: path.join(IMG, "linkedin.png"), x: img.x, y: img.y, w: img.w, h: img.h, altText: "Print do LinkedIn de Fernando Lohmann" }, "Print LinkedIn"));
    annotate(s, img, [
      { r: [560, 20, 1540, 340] }, { r: [185, 145, 532, 492] }, { r: [172, 760, 820, 810] }, { r: [172, 870, 510, 915] }, { r: [945, 565, 1495, 750] },
    ], "musgo");
    legend(s, [
      { t: "Banner pessoal", d: "hoje traz a frase da Weplace; passa a contar a sua história" },
      { t: "Foto do ensaio", d: "mesmo retrato nos três canais" },
      { t: "Headline em português", d: "do cargo para o valor que você entrega" },
      { t: "Modo criador e seguidores", d: "de conexões para audiência, com conteúdo semanal" },
      { t: "Experiências e Em destaque", d: "conquistas, empresas, Singapura e conselhos" },
    ], 7.75, 2.2, SW - MX - 7.75, 0.8, "Legenda LinkedIn");
    T(s, "É hoje o seu canal mais forte. A reformulação alinha o perfil ao Brandbook e prepara a chegada dos posts e vídeos.", { x: 7.75, y: 6.2, w: SW - MX - 7.75, h: 0.6, ...body({ fontSize: 12 }) }, "Texto LinkedIn");
  }

  // ============ 5 · Site hoje
  {
    const s = light("Diagnóstico");
    s.addText("04 · diagnóstico · site".toUpperCase(), { placeholder: "kicker" });
    s.addText("Do cartão de visitas à casa da sua marca.", { placeholder: "title" });
    const img = { x: MX, y: 2.2, w: 6.4 };
    img.k = img.w / 2000; img.h = 1039 * img.k;
    s.addImage(on({ path: path.join(IMG, "site.png"), x: img.x, y: img.y, w: img.w, h: img.h, altText: "Print do site prfal.com.br" }, "Print site"));
    annotate(s, img, [{ r: [60, 25, 262, 78] }, { r: [1225, 25, 1960, 78] }, { r: [50, 195, 1345, 470] }, { r: [50, 700, 480, 765] }], "lima");
    T(s, "PRFAL.COM.BR  ·  HOJE: SEM DEPOIMENTOS · SEM LINKS PARA AS EMPRESAS", { x: MX, y: img.y + img.h + 0.18, w: img.w, h: 0.3, fontFace: MONO, fontSize: 9, charSpacing: 0.5, color: C.text2 }, "Dados site");
    T(s, "A base está pronta: trajetória, Capital Humano, Mentoria e Ventures. Vamos redesenhar tudo para que o site converse com o Instagram e o LinkedIn e leve cada visitante ao próximo passo.", { x: MX, y: 6.0, w: img.w, h: 0.75, ...body({ fontSize: 13 }) }, "Texto site");
    const rx = 7.55, rw = SW - MX - rx;
    card(s, rx, 2.2, rw, 4.55, C.background2, "Quadro redefinição");
    label(s, "o que redefinimos", rx + 0.35, 2.45, rw - 0.7, C.text1, "Rótulo redefinição");
    const items = [
      ["Nova paleta e tipografia", "alinhadas ao Brandbook e aos outros canais"],
      ["Mensagem principal", "o que você faz e para quem, em uma frase"],
      ["Links clicáveis", "Capital Humano, Mentoria, Ventures e Trajetória"],
      ["Suas empresas", "Weplace, EnsureAtt, Board Intelligence e produtos"],
      ["Lado pessoal, com estratégia", "Singapura, bastidores e o que te move"],
      ["Depoimentos e pessoas", "clientes, conselhos e parceiros falando de você"],
      ["Chamada para conversa", "WhatsApp e links para todas as redes"],
    ];
    items.forEach((it, i) => {
      const y = 2.85 + i * 0.54;
      dot(s, rx + 0.35, y + 0.08, 0.16, C.text1, `Ponto ${i + 1}`);
      T(s, [{ text: it[0], options: { color: C.text1, breakLine: true, fontSize: 13 } }, { text: it[1], options: { color: C.text2, fontSize: 11 } }], { x: rx + 0.68, y: y - 0.02, w: rw - 1.0, h: 0.52, valign: "top" }, `Redefinição ${i + 1}`);
    });
  }

  // ============ 6 · Quem somos
  pres.addSection({ title: "A Meov" });
  {
    const s = dark("A Meov");
    s.addText("05 · quem somos".toUpperCase(), { placeholder: "kicker" });
    s.addText("Um time ao seu lado, da estratégia à câmera.", { placeholder: "title" });
    dot(s, MX, 2.45, 0.2, C.accent1, "Ponto rosa");
    dot(s, MX + 0.28, 2.45, 0.2, C.accent2, "Ponto verde");
    T(s, "Não somos uma agência de marketing. Somos um time que apoia técnica e emocionalmente quem decide colocar a própria voz no mundo.", { x: MX, y: 2.9, w: 4.8, h: 2.2, fontFace: LIGHT, fontSize: 22, color: C.background1, valign: "top", lineSpacingMultiple: 1.05 }, "Manifesto");
    T(s, "Nosso trabalho começa antes do primeiro post, entendendo quem você é, e continua depois dele, com produção e gestão todos os meses.", { x: MX, y: 5.35, w: 4.6, h: 1.1, fontSize: 13, color: C.background2, valign: "top", lineSpacingMultiple: 1.15 }, "Manifesto apoio");
    const fronts = [
      { n: "Estratégia", i: "TbCompass", d: "O Método Meov organiza quem você é, para quem você fala e como se diferencia. Tudo vira um Brandbook que guia os próximos 6 meses." },
      { n: "Audiovisual", i: "TbMovie", d: "Direção de imagem, gravação e edição com o cuidado de quem nasceu no audiovisual. Você aparece bem, com naturalidade e do seu jeito." },
      { n: "Gestão das redes", i: "TbLayoutGrid", d: "Instagram, LinkedIn e Site/LP com calendário, roteiros, legendas, design, publicação e relatórios. Você grava e aprova; a gente cuida do resto." },
    ];
    const fx = 5.85, fw = SW - MX - fx;
    for (let i = 0; i < 3; i++) {
      const y = 2.25 + i * 1.47;
      card(s, fx, y, fw, 1.4, C.background2, `Frente ${i + 1} cartão`);
      await mark(s, fronts[i].i, fx + 0.3, y + 0.32, 0.7, `Frente ${i + 1}`);
      T(s, fronts[i].n, { x: fx + 1.25, y: y + 0.2, w: fw - 1.5, h: 0.35, fontSize: 17, color: C.text1 }, `Frente ${i + 1} nome`);
      T(s, fronts[i].d, { x: fx + 1.25, y: y + 0.55, w: fw - 1.5, h: 0.8, ...body({ fontSize: 13 }) }, `Frente ${i + 1} texto`);
    }
  }

  // ============ 7 · O seu time
  {
    const s = stone("A Meov");
    s.addText("06 · o seu time".toUpperCase(), { placeholder: "kicker" });
    s.addText("Um time dedicado à sua marca.", { placeholder: "title" });
    T(s, "Você fala com quem pensa a sua marca e grava com quem conhece a sua história. O resto acontece nos bastidores, no Estúdio Meov.", { x: MX, y: 1.85, w: CW, h: 0.35, ...body({ fontSize: 14 }) }, "Intro time");

    const gap = 0.5, cw = (CW - 2 * gap) / 3, y0 = 2.5, ch = 4.2;
    const cols = [
      {
        step: "01 · pensa", n: "Estratégia e Direção", sub: "a cabeça da sua marca", f: C.background1, t: C.text1, st: C.text2, lb: C.text2,
        av: [{ l: "D", f: C.text1, t: C.background1 }, { l: "R", f: C.text1, t: C.background1 }],
        rows: [
          { n: "Diogo", r: "estrategista de conteúdo", d: "Conduz o Método Meov, o Brandbook e o plano de conteúdo dos 6 meses." },
          { n: "Rafael", r: "diretor de comunicação", d: "Cuida da mensagem, do tom de voz e da coerência entre os canais." },
        ],
      },
      {
        step: "02 · grava", n: "Time de Gravação", sub: "ao seu lado, no set", f: C.text1, t: C.background1, st: C.background2, lb: C.accent2,
        av: [{ l: "D", f: C.background1, t: C.text1 }, { l: "R", f: C.background1, t: C.text1 }, { l: "I", f: C.accent2, t: C.text1 }],
        rows: [
          { n: "Diogo e Rafael", r: PAPEL_GRAVACAO, d: "Entrevistam você, conduzem a conversa e lapidam cada resposta para funcionar nas redes." },
          { n: "Italo", r: "filmmaker e diretor de arte", d: "Luz, enquadramento, cenário e imagem com a cara da sua marca." },
        ],
      },
      {
        step: "03 · transforma", n: "Estúdio Meov", sub: "os bastidores, todos os meses", f: C.background1, t: C.text1, st: C.text2, lb: C.text2,
        roles: [
          { n: "Editor de vídeo", i: "TbScissors", d: "cortes, legendas e trilha dos reels" },
          { n: "Designer", i: "TbPalette", d: "carrosséis, templates e identidade" },
          { n: "Gestora de redes", i: "TbLayoutGrid", d: "calendário, publicação e relatórios" },
          { n: "Copywriter", i: "TbPencil", d: "roteiros, legendas e posts de texto" },
        ],
      },
    ];
    for (let i = 0; i < cols.length; i++) {
      const c = cols[i], x = MX + i * (cw + gap), y = y0;
      card(s, x, y, cw, ch, c.f, `Time ${i + 1} cartão`);
      label(s, c.step, x + 0.3, y + 0.28, cw - 0.6, c.lb, `Time ${i + 1} etapa`);
      T(s, c.n, { x: x + 0.3, y: y + 0.6, w: cw - 0.6, h: 0.42, fontFace: LIGHT, fontSize: 21, color: c.t }, `Time ${i + 1} nome`);
      T(s, c.sub, { x: x + 0.3, y: y + 1.02, w: cw - 0.6, h: 0.3, fontSize: 12, color: c.st }, `Time ${i + 1} apoio`);
      if (c.av) {
        c.av.forEach((a, k) => {
          const d = 0.6, ax = x + 0.3 + k * 0.48;
          s.addShape(pres.shapes.OVAL, on({ x: ax, y: y + 1.45, w: d, h: d, fill: { color: a.f }, line: { color: c.f, width: 2.5 } }, `Time ${i + 1} pessoa ${k + 1}`));
          T(s, a.l, { x: ax, y: y + 1.45, w: d, h: d, fontSize: 16, color: a.t, align: "center", valign: "middle" }, `Time ${i + 1} inicial ${k + 1}`);
        });
        c.rows.forEach((r, k) => {
          const ry = y + 2.25 + k * 1.02;
          T(s, [
            { text: r.n, options: { fontSize: 14, color: c.t, breakLine: true } },
            { text: r.r.toUpperCase(), options: { fontFace: MONO, fontSize: 9, charSpacing: 1, color: c.lb, breakLine: true } },
            { text: r.d, options: { fontSize: 11, color: c.st } },
          ], { x: x + 0.3, y: ry, w: cw - 0.55, h: 0.92, valign: "top", lineSpacingMultiple: 1.08 }, `Time ${i + 1} pessoa ${k + 1} texto`);
        });
      } else {
        for (let k = 0; k < c.roles.length; k++) {
          const r = c.roles[k], d = 0.6, ax = x + 0.3 + k * 0.48;
          s.addShape(pres.shapes.OVAL, on({ x: ax, y: y + 1.45, w: d, h: d, fill: { color: C.background2 }, line: { color: c.f, width: 2.5 } }, `Estúdio ${k + 1} fundo`));
          s.addImage(on({ data: await icon(r.i, HEX.musgo), x: ax + d * 0.26, y: y + 1.45 + d * 0.26, w: d * 0.48, h: d * 0.48 }, `Estúdio ${k + 1} ícone`));
          T(s, [
            { text: r.n, options: { fontSize: 13, color: c.t, breakLine: true } },
            { text: r.d, options: { fontSize: 11, color: c.st } },
          ], { x: x + 0.3, y: y + 2.3 + k * 0.47, w: cw - 0.55, h: 0.46, valign: "top" }, `Estúdio ${k + 1} texto`);
        }
      }
      if (i < cols.length - 1) {
        const ax = x + cw + gap / 2 - 0.17;
        s.addImage(on({ data: await icon("TbArrowRight", HEX.ardosia), x: ax, y: y + ch / 2 - 0.17, w: 0.34, h: 0.34 }, `Fluxo ${i + 1}`));
      }
    }
  }

  // ============ 8 · Método e jornada
  {
    const s = stone("A Meov");
    s.addText("07 · o método e a jornada".toUpperCase(), { placeholder: "kicker" });
    s.addText("Da raiz à marca consolidada.", { placeholder: "title" });
    const mov = [
      { n: "Raiz", d: "Quem você é e o que te torna referência." },
      { n: "Território", d: "Para quem você fala e que espaço só você ocupa." },
      { n: "Voz", d: "Seu tom, sua imagem e sua oferta." },
      { n: "Movimento", d: "Temas, canais e constância." },
      { n: "Comunidade", d: "Pessoas que acompanham, interagem e indicam." },
      { n: "Marca consolidada", d: "Referência reconhecida, que gera oportunidades e legado." },
    ];
    const mw = (CW - 5 * 0.15) / 6;
    s.addShape(pres.shapes.LINE, on({ x: MX + 0.3, y: 2.45, w: CW - mw, h: 0, line: { color: C.text2, width: 1, dashType: "dash" } }, "Linha movimentos"));
    mov.forEach((m, i) => {
      const x = MX + i * (mw + 0.15);
      numDot(s, String(i + 1).padStart(2, "0"), x, 2.15, 0.6, i === 5 ? C.accent2 : C.text1, i === 5 ? C.text1 : C.background1, `Movimento ${i + 1}`, 12);
      T(s, m.n, { x, y: 2.88, w: mw + 0.1, h: 0.36, fontSize: 15, color: C.text1 }, `Movimento ${i + 1} nome`);
      T(s, m.d, { x, y: 3.3, w: mw - 0.05, h: 0.85, ...body({ fontSize: 12 }) }, `Movimento ${i + 1} texto`);
    });
    T(s, "Os quatro primeiros movimentos viram o seu Brandbook. Os dois últimos são o que a constância constrói ao longo do ciclo.", { x: MX, y: 4.35, w: CW, h: 0.3, fontSize: 13, color: C.text1 }, "Nota método");
    const ph = [
      { m: "início", n: "Implantação", d: "método, ensaio e canais prontos", f: C.text1, t: C.background1, s: C.background2 },
      { m: "mês 1", n: "Fundação", d: "perfis abertos e primeiros conteúdos" },
      { m: "mês 2", n: "Ritmo", d: "Vlogcast e séries fixas no ar" },
      { m: "mês 3", n: "Leitura", d: "1ª revisão de rota" },
      { m: "mês 4", n: "Autoridade", d: "casos, provas e opinião" },
      { m: "mês 5", n: "Comunidade", d: "conversas, convites e indicações" },
      { m: "mês 6", n: "Colheita", d: "resultado e novo ciclo", f: C.accent2, t: C.text1, s: C.text1 },
    ];
    const pw = (CW - 6 * 0.12) / 7;
    ph.forEach((p, i) => {
      const x = MX + i * (pw + 0.12), y = 4.85;
      const tc = p.t || C.text1, sc = p.s || C.text2;
      card(s, x, y, pw, 1.75, p.f || C.background1, `Fase ${i} cartão`);
      label(s, p.m, x + 0.18, y + 0.2, pw - 0.3, sc, `Fase ${i} rótulo`);
      T(s, p.n, { x: x + 0.18, y: y + 0.55, w: pw - 0.3, h: 0.4, fontSize: 15, color: tc }, `Fase ${i} nome`);
      T(s, p.d, { x: x + 0.18, y: y + 0.95, w: pw - 0.3, h: 0.8, fontSize: 12, color: sc, valign: "top" }, `Fase ${i} texto`);
    });
  }

  // ============ 9 · Implantação
  pres.addSection({ title: "Escopo" });
  {
    const s = light("Escopo");
    s.addText("08 · parte 1 · implantação".toUpperCase(), { placeholder: "kicker" });
    s.addText("Primeiro, construímos a base.", { placeholder: "title" });
    T(s, "Antes do primeiro conteúdo, organizamos sua marca e deixamos cada canal pronto para receber quem vai te encontrar.", { x: MX, y: 1.85, w: CW, h: 0.35, ...body({ fontSize: 14 }) }, "Intro implantação");
    const items = [
      { n: "Método Meov", i: "TbCompass", d: "Diagnóstico da sua presença e de como você é percebido, imersão presencial e o seu Brandbook, detalhado no próximo slide. Regras dos conselhos e empresas entram desde o início." },
      { n: "Ensaio de marca", i: "TbCamera", d: "Uma sessão de fotos com direção de imagem e 12 fotos tratadas, pensadas para perfil, banner, site e conteúdos." },
      { n: "Instagram", i: "TbBrandInstagram", d: "Perfil aberto como conta de criador, bio, link, destaques, 3 posts fixados e templates próprios. Verificação em duas etapas e acesso pelo Meta Business Suite, sem troca de senha." },
      { n: "LinkedIn", i: "TbBrandLinkedin", d: "Headline em português, Sobre e experiências como conquistas, banner próprio, Em destaque, competências e modo criador." },
      { n: "Site/LP", i: "TbWorld", d: "Redesenho completo: apresentação, história, como você ajuda, depoimentos, WhatsApp, redes e links para os sites das suas empresas e produtos." },
      { n: "Também inclui", i: "TbChecks", d: "Fluxo de aprovação de conteúdos combinado entre você e a Meov, e autorização de uso de imagem, para trabalharmos com tranquilidade.", dk: true },
    ];
    const cw = (CW - 0.5) / 3, ch = 2.1;
    for (let i = 0; i < items.length; i++) {
      const it = items[i], x = MX + (i % 3) * (cw + 0.25), y = 2.35 + Math.floor(i / 3) * (ch + 0.18);
      card(s, x, y, cw, ch, it.dk ? C.text1 : C.background2, `Entrega ${i + 1} cartão`);
      if (it.dk) await iconDot(s, it.i, x + 0.28, y + 0.25, 0.5, C.accent2, HEX.musgo, `Entrega ${i + 1}`);
      else await mark(s, it.i, x + 0.28, y + 0.25, 0.5, `Entrega ${i + 1}`);
      T(s, it.n, { x: x + 0.92, y: y + 0.3, w: cw - 1.1, h: 0.4, fontSize: 16, color: it.dk ? C.background1 : C.text1 }, `Entrega ${i + 1} nome`);
      T(s, it.d, { x: x + 0.28, y: y + 0.85, w: cw - 0.5, h: ch - 0.95, fontSize: 12, color: it.dk ? C.background2 : C.text2, valign: "top", lineSpacingMultiple: 1.12 }, `Entrega ${i + 1} texto`);
    }
  }

  // ============ 10 · Brandbook
  {
    const s = stone("Escopo");
    s.addText("09 · o seu brandbook".toUpperCase(), { placeholder: "kicker" });
    s.addText("O mapa da sua marca, em um só documento.", { placeholder: "title" });
    T(s, "Organiza tudo o que descobrimos no diagnóstico e na imersão. É o guia de cada conteúdo, cada imagem e cada decisão dos próximos 6 meses, e continua seu depois do ciclo.", { x: MX, y: 1.85, w: CW, h: 0.55, ...body({ fontSize: 14 }) }, "Intro Brandbook");
    const lightCard = { f: C.background1, t: C.text1, s: C.text2, l: C.text2 };
    const parts = [
      { n: "Raiz", f: C.text1, t: C.background1, s: C.background2, l: C.accent2, items: ["Como você é visto hoje", "Por que aparecer: desejo e travas", "Propósito e frase-norte", "Valores inegociáveis", "Sua história, provas e o que mostrar ou guardar"] },
      { n: "Território", ...lightCard, items: ["Quem atrair e quem afastar", "Dores, desejos e linguagem do seu público", "Mantra de posicionamento", "Mapa do mercado e espaço em branco", "Frase de diferenciação"] },
      { n: "Voz", ...lightCard, items: ["Credo e vocabulário próprio", "Tom de voz", "Cores, fontes e direção de foto e vídeo", "Moodboard", "Oferta e chamada principal"] },
      { n: "Movimento", ...lightCard, items: ["3 territórios de conteúdo", "Banco de pautas inicial", "Linha do Vlogcast", "Papel de cada canal", "Plano de 6 meses, metas e revisões"] },
    ];
    const pw = (CW - 0.6) / 4;
    parts.forEach((p, i) => {
      const x = MX + i * (pw + 0.2), y = 2.6;
      card(s, x, y, pw, 3.6, p.f, `Parte ${i + 1} cartão`);
      label(s, `parte ${i + 1}`, x + 0.3, y + 0.28, pw - 0.6, p.l, `Parte ${i + 1} rótulo`);
      T(s, p.n, { x: x + 0.3, y: y + 0.58, w: pw - 0.6, h: 0.5, fontFace: LIGHT, fontSize: 24, color: p.t }, `Parte ${i + 1} nome`);
      T(s, p.items.map((t, k) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: k < p.items.length - 1 } })), { x: x + 0.3, y: y + 1.25, w: pw - 0.5, h: 2.2, fontSize: 13, color: p.s, valign: "top", paraSpaceAfter: 5 }, `Parte ${i + 1} itens`);
    });
    T(s, "Abre com \"a marca em uma página\": frase-norte, mantra, credo, tom de voz e chamada principal, para consulta rápida.", { x: MX, y: 6.35, w: CW - 1.3, h: 0.35, fontSize: 12, color: C.text2 }, "Nota Brandbook");
  }

  // ============ 11 · Acompanhamento
  {
    const s = light("Escopo");
    s.addText("10 · parte 2 · acompanhamento mensal".toUpperCase(), { placeholder: "kicker" });
    s.addText("Depois, presença constante todos os meses.", { placeholder: "title" });
    const stats = [
      { v: "12", l: "reels editados, com cortes, legendas e trilha" },
      { v: "8", l: "carrosséis com texto e design" },
      { v: "4", l: "posts de texto no LinkedIn, a partir das gravações" },
      { v: "2", l: "gravações de Vlogcast por mês, com direção" },
    ];
    const sw = 2.55, sh = 1.7;
    stats.forEach((st, i) => {
      const x = MX + (i % 2) * (sw + 0.2), y = 2.15 + Math.floor(i / 2) * (sh + 0.18);
      card(s, x, y, sw, sh, C.background2, `Número ${i + 1} cartão`);
      T(s, st.v, { x: x + 0.3, y: y + 0.1, w: sw - 0.5, h: 0.9, fontFace: LIGHT, fontSize: 54, color: C.text1 }, `Número ${i + 1} valor`);
      T(s, st.l, { x: x + 0.3, y: y + 0.98, w: sw - 0.5, h: 0.65, fontSize: 12, color: C.text1, valign: "top" }, `Número ${i + 1} rótulo`);
    });
    card(s, MX, 5.85, 2 * sw + 0.2, 0.85, C.background2, "Quadro Vlogcast");
    T(s, [{ text: "Vlogcast: ", options: { color: C.text1 } }, { text: "linha editorial inicial, escolhida pelo seu estilo de comunicação: conversa, experiência e opinião.", options: { color: C.text2 } }], { x: MX + 0.3, y: 5.85, w: 2 * sw - 0.4, h: 0.85, fontSize: 12, valign: "middle" }, "Texto Vlogcast");

    const rx = 6.4, rw = SW - MX - rx;
    card(s, rx, 2.15, rw, 1.15, C.text1, "Quadro condição especial");
    await iconDot(s, "TbGift", rx + 0.25, 2.43, 0.58, C.accent2, HEX.musgo, "Presente");
    T(s, [
      { text: "Condição especial para o Fernando", options: { color: C.accent2, breakLine: true, fontSize: 14 } },
      { text: "Nos 3 primeiros meses, a gravação é semanal (4 por mês), sem custo adicional.", options: { color: C.background1, fontSize: 12 } },
    ], { x: rx + 1.05, y: 2.2, w: rw - 1.25, h: 1.05, valign: "middle" }, "Texto condição especial");
    label(s, "como cada mês acontece", rx, 3.6, rw, C.text2, "Rótulo ciclo");
    const steps = [
      ["Pauta e roteiro", "a partir do Brandbook e do que acontece no seu mercado"],
      ["Gravação", "Vlogcast com direção de imagem, no seu ritmo"],
      ["Edição e design", "reels finalizados e carrosséis prontos para aprovar"],
      ["Publicação", "Instagram e LinkedIn, no formato certo para cada rede"],
      ["Relatório e reunião", "o que funcionou, o que ajustar e o plano seguinte"],
    ];
    steps.forEach((st, i) => {
      const y = 3.95 + i * 0.55;
      numDot(s, i + 1, rx, y, 0.36, C.text1, C.background1, `Etapa ${i + 1}`, 10);
      T(s, [{ text: st[0] + "  ", options: { color: C.text1 } }, { text: st[1], options: { color: C.text2 } }], { x: rx + 0.55, y: y - 0.04, w: rw - 0.55, h: 0.45, fontSize: 13, valign: "middle" }, `Etapa ${i + 1} texto`);
    });
  }

  // ============ 12 · Como medimos
  {
    const s = stone("Escopo");
    s.addText("11 · como medimos".toUpperCase(), { placeholder: "kicker" });
    s.addText("Não medimos só seguidores.", { placeholder: "title" });
    T(s, "Acompanhamos cinco camadas, sempre comparadas com o ponto de partida do diagnóstico, e revisamos a rota no 3º e no 6º mês.", { x: MX, y: 1.85, w: CW, h: 0.35, ...body({ fontSize: 14 }) }, "Intro medição");
    const layers = [
      { n: "Presença", d: "Você está aparecendo com constância? Calendário cumprido e séries mantidas.", f: C.background1 },
      { n: "Alcance qualificado", d: "As pessoas certas estão te vendo? Olhamos quem chega, não só quantos.", f: C.background1 },
      { n: "Conversa", d: "Salvam, compartilham, comentam e te chamam? É o sinal de vínculo.", f: C.accent2 },
      { n: "Negócio e ROI", d: "Viram reuniões, mandatos, convites e clientes? Comparamos esse retorno com o investimento.", f: C.text1, t: C.background1 },
      { n: "Confiança", d: "Você se sente mais seguro sendo quem é? Sem ela, nada se sustenta.", f: C.accent1, t: C.background1 },
    ];
    layers.forEach((l, i) => {
      const w = 5.6 - i * 0.55, x = MX + (5.6 - w) / 2, y = 2.45 + i * 0.66;
      card(s, x, y, w, 0.56, l.f, `Camada ${i + 1}`);
      T(s, l.n, { x, y, w, h: 0.56, fontSize: 14, color: l.t || C.text1, align: "center", valign: "middle" }, `Camada ${i + 1} nome`);
      T(s, l.d, { x: 6.85, y: y + 0.02, w: SW - MX - 6.85, h: 0.54, fontSize: 12, color: C.text2, valign: "middle" }, `Camada ${i + 1} texto`);
    });
    card(s, MX, 5.95, CW - 1.3, 0.75, C.text1, "Quadro ROI");
    T(s, [
      { text: "ROI DO CICLO = (NEGÓCIOS GERADOS PELA MARCA − INVESTIMENTO) ÷ INVESTIMENTO", options: { fontFace: MONO, fontSize: 11, color: C.accent2, charSpacing: 0.5, breakLine: true } },
      { text: "Registramos a origem de cada oportunidade e mostramos o retorno nos relatórios mensais.", options: { fontSize: 12, color: C.background2 } },
    ], { x: MX + 0.35, y: 5.95, w: CW - 2.0, h: 0.75, valign: "middle" }, "Texto ROI");
  }

  // ============ tabelas de investimento
  const priceTable = (s, x, y, w, title, rows, tot, cols, darkBg, unit) => {
    const tc = darkBg ? C.background1 : C.text1, sc = darkBg ? C.background2 : C.text2;
    const nameW = w - cols.length * 1.12;
    label(s, title, x, y, nameW, darkBg ? C.accent2 : C.text1, `${title} título`);
    cols.forEach((c, j) => label(s, c.h, x + nameW + j * 1.12, y, 1.12, sc, `${title} coluna ${j + 1}`, "right"));
    rows.forEach((r, i) => {
      const yy = y + 0.38 + i * 0.4;
      s.addShape(pres.shapes.LINE, on({ x, y: yy, w, h: 0, line: { color: HEX.ardosia, width: darkBg ? 0.75 : 0.5 } }, `${title} linha ${i + 1}`));
      T(s, r.n, { x, y: yy + 0.04, w: nameW, h: 0.34, fontSize: 13, color: tc, valign: "middle" }, `${title} item ${i + 1}`);
      cols.forEach((c, j) => T(s, fmt(r[c.k]), { x: x + nameW + j * 1.12, y: yy + 0.04, w: 1.12, h: 0.34, fontFace: MONO, fontSize: 11, color: c.strike ? sc : c.color, strike: c.strike ? "sngStrike" : undefined, align: "right", valign: "middle" }, `${title} valor ${i + 1}.${j + 1}`));
    });
    const yy = y + 0.38 + rows.length * 0.4;
    s.addShape(pres.shapes.LINE, on({ x, y: yy, w, h: 0, line: { color: tc === C.background1 ? HEX.mineral : HEX.musgo, width: 1 } }, `${title} linha total`));
    T(s, "Subtotal" + (unit || ""), { x, y: yy + 0.06, w: nameW, h: 0.36, fontSize: 13, color: tc, valign: "middle" }, `${title} subtotal`);
    cols.forEach((c, j) => T(s, fmt(tot[c.k]), { x: x + nameW + j * 1.12, y: yy + 0.06, w: 1.12, h: 0.36, fontFace: MONO, fontSize: 12, color: c.strike ? sc : c.color, strike: c.strike ? "sngStrike" : undefined, align: "right", valign: "middle" }, `${title} subtotal ${j + 1}`));
  };

  // ============ 13 · Investimento detalhado
  pres.addSection({ title: "Investimento" });
  {
    const s = stone("Investimento");
    s.addText("12 · investimento".toUpperCase(), { placeholder: "kicker" });
    s.addText("Cada entrega tem valor. Juntas, custam menos.", { placeholder: "title" });
    const cols = [{ h: "avulso", k: "av", strike: true }, { h: "no pacote", k: "pc", color: C.text1 }];
    const pw = (CW - 0.5) / 2;
    priceTable(s, MX, 2.15, pw, "implantação", IMPL, T_IMPL, cols, false);
    priceTable(s, MX + pw + 0.5, 2.15, pw, "mensal", MENS, T_MENS, cols, false, " por mês");
    card(s, MX, 5.75, CW - 1.3, 0.95, C.text1, "Faixa total");
    T(s, [
      { text: "CICLO DE 6 MESES  ", options: { fontFace: MONO, fontSize: 11, charSpacing: 1, color: C.background2 } },
      { text: fmt(CYCLE.av), options: { fontFace: MONO, fontSize: 14, color: C.background2, strike: "sngStrike" } },
      { text: "    no pacote  ", options: { fontSize: 14, color: C.background1 } },
      { text: fmt(CYCLE.pc), options: { fontFace: LIGHT, fontSize: 30, color: C.background1 } },
    ], { x: MX + 0.35, y: 5.75, w: CW - 2.0, h: 0.95, valign: "middle" }, "Total do ciclo");
  }

  // ============ 14 · Condição especial
  {
    const s = dark("Investimento");
    s.addText("13 · condição especial fernando lohmann".toUpperCase(), { placeholder: "kicker" });
    s.addText("Para começar juntos, uma condição única.", { placeholder: "title" });
    const cols = [{ h: "avulso", k: "av", strike: true }, { h: "pacote", k: "pc", strike: true }, { h: "para você", k: "es", color: C.accent2 }];
    const pw = (CW - 0.5) / 2;
    priceTable(s, MX, 2.15, pw, "implantação", IMPL, T_IMPL, cols, true);
    priceTable(s, MX + pw + 0.5, 2.15, pw, "mensal", MENS, T_MENS, cols, true, " por mês");
    card(s, MX, 5.6, CW - 1.3, 1.1, C.background1, "Faixa total especial");
    T(s, [
      { text: "CICLO DE 6 MESES  ", options: { fontFace: MONO, fontSize: 11, charSpacing: 1, color: C.text1 } },
      { text: fmt(CYCLE.av) + "   " + fmt(CYCLE.pc), options: { fontFace: MONO, fontSize: 13, color: C.text2, strike: "sngStrike" } },
      { text: "    para você  ", options: { fontSize: 14, color: C.text1 } },
      { text: fmt(CYCLE.es), options: { fontFace: LIGHT, fontSize: 34, color: C.accent1 } },
    ], { x: MX + 0.35, y: 5.6, w: CW - 2.0, h: 0.68, valign: "middle" }, "Total especial");
    T(s, `IMPLANTAÇÃO EM 3X DE ${fmt(T_IMPL.es / 3)}  ·  MESES 1–3: ${fmt(T_MENS.es + T_IMPL.es / 3)}/MÊS  ·  MESES 4–6: ${fmt(T_MENS.es)}/MÊS`, { x: MX + 0.35, y: 6.25, w: CW - 2.0, h: 0.35, fontFace: MONO, fontSize: 10, charSpacing: 0.5, color: C.text2 }, "Parcelas especial");
  }

  // ============ 15 · À parte
  pres.addSection({ title: "Combinados" });
  {
    const s = stone("Combinados");
    s.addText("14 · combinados".toUpperCase(), { placeholder: "kicker" });
    s.addText("O que fica de fora e pode ser contratado à parte.", { placeholder: "title" });
    const groups = [
      { n: "Produção", d: "Tráfego pago · deslocamento fora da Grande SP · locação, maquiagem e figurino · modelos, drone e equipe extra · vídeos longos e eventos" },
      { n: "Gestão e canais", d: "Responder comentários e DMs (fica com você, para manter sua voz real) · stories do dia a dia · outras redes · e-mail marketing" },
      { n: "Comunicação", d: "Assessoria de imprensa · parcerias com influenciadores · tradução e outros idiomas · trilhas licenciadas pagas" },
      { n: "Site", d: "Site com várias páginas, blog ou loja · domínio, hospedagem e manutenção" },
      { n: "Marca e jurídico", d: "Criação de logotipo · registro no INPI · revisão jurídica e compliance · autorização de imagem de terceiros" },
      { n: "Resultados", d: "Não garantimos seguidores, views ou vendas: garantimos método, qualidade e constância. Nunca compramos seguidores." },
    ];
    const cw = (CW - 0.5) / 3, ch = 1.8;
    groups.forEach((g, i) => {
      const x = MX + (i % 3) * (cw + 0.25), y = 2.15 + Math.floor(i / 3) * (ch + 0.18);
      card(s, x, y, cw, ch, C.background1, `Fora ${i + 1} cartão`);
      dot(s, x + 0.28, y + 0.29, 0.14, C.text1, `Fora ${i + 1} ponto`);
      label(s, g.n, x + 0.52, y + 0.22, cw - 0.75, C.text1, `Fora ${i + 1} nome`);
      T(s, g.d, { x: x + 0.28, y: y + 0.55, w: cw - 0.5, h: ch - 0.65, fontSize: 13, color: C.text2, valign: "top", lineSpacingMultiple: 1.12 }, `Fora ${i + 1} texto`);
    });
    label(s, "como trabalhamos", MX, 6.05, 4, C.text2, "Rótulo combinados");
    const rules = ["Ciclo de 6 meses", "Até 2 rodadas de ajuste", "Mínimos mensais sem acúmulo", "Extras sob consulta"];
    const ws = rules.map((r) => 0.6 + r.length * 0.085);
    const gapR = (CW - 1.3 - ws.reduce((a, b) => a + b, 0)) / (rules.length - 1);
    let x = MX;
    rules.forEach((r, i) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, on({ x, y: 6.35, w: ws[i], h: 0.42, rectRadius: 0.21, fill: { color: C.text1 }, line: { color: HEX.musgo, width: 0 } }, `Combinado ${i + 1} fundo`));
      T(s, r, { x, y: 6.35, w: ws[i], h: 0.42, fontSize: 11, color: C.background1, align: "center", valign: "middle" }, `Combinado ${i + 1}`);
      x += ws[i] + gapR;
    });
  }

  // ============ 16 · Próximos passos
  pres.addSection({ title: "Encerramento" });
  {
    const s = dark("Encerramento");
    s.addText("15 · próximos passos".toUpperCase(), { placeholder: "kicker" });
    T(s, "Você já tem a voz.", { x: MX, y: 1.4, w: 6.2, h: 0.9, fontFace: LIGHT, fontSize: 46, charSpacing: -1, color: C.background1 }, "Frase 1");
    T(s, "A gente ajuda você a colocá-la no mundo.", { x: MX, y: 2.3, w: 5.8, h: 1.0, fontFace: LIGHT, fontSize: 26, color: C.accent2, valign: "top" }, "Frase 2");
    const steps = [["Aprovação da proposta", "Você valida o escopo e a condição especial."], ["Assinatura do contrato", "Formalizamos o ciclo de 6 meses."], ["Diagnóstico e imersão", "Agendamos o primeiro encontro presencial."]];
    steps.forEach((st, i) => {
      const y = 3.65 + i * 0.85;
      numDot(s, i + 1, MX, y, 0.52, C.background1, C.text1, `Passo ${i + 1}`, 13);
      T(s, [{ text: st[0], options: { breakLine: true, color: C.background1, fontSize: 16 } }, { text: st[1], options: { color: C.background2, fontSize: 12 } }], { x: MX + 0.75, y: y - 0.06, w: 5.4, h: 0.7, valign: "middle" }, `Passo ${i + 1} texto`);
    });
    label(s, "diogo & rafael · meov marca pessoal · @meov.mp · meov.com.br", MX, 6.45, 6.6, C.text2, "Assinatura");

    // Ilustração: antes x depois do posicionamento
    const x0 = 7.15, w0 = SW - MX - x0, gp = 0.3, pw = (w0 - gp) / 2, py = 1.35, phH = 3.05;
    const person = (cx, base, opts, name) => {
      const tw = 1.25, th = 0.8, hd = 0.62;
      const ty = base - th, hy = ty - hd - 0.08;
      s.addShape(pres.shapes.ROUND_2_SAME_RECTANGLE, on({ x: cx - tw / 2, y: ty, w: tw, h: th, rectRadius: 0.5, ...opts }, name + " corpo"));
      s.addShape(pres.shapes.OVAL, on({ x: cx - hd / 2, y: hy, w: hd, h: hd, ...opts }, name + " cabeça"));
      return { hx: cx, hy: hy + hd / 2 };
    };
    // painel antes
    {
      const px = x0;
      card(s, px, py, pw, phH, HEX.ardosia, "Antes painel");
      const cx = px + pw / 2, base = py + phH - 0.02;
      const h = person(cx, base, { fill: { color: HEX.ardosia }, line: { color: HEX.pedra, width: 1.25, dashType: "dash" } }, "Antes pessoa");
      s.addShape(pres.shapes.OVAL, on({ x: cx - 0.95, y: h.hy - 0.95, w: 1.9, h: 1.9, fill: { color: HEX.ardosia, transparency: 100 }, line: { color: HEX.pedra, width: 0.75, dashType: "sysDot" } }, "Antes círculo"));
      [[0.25, 0.3], [0.55, 0.22], [2.05, 0.28], [2.25, 0.62], [0.2, 0.95], [2.2, 1.25]].forEach((p, i) =>
        dot(s, px + p[0], py + p[1], 0.14, HEX.musgo, `Antes público ${i + 1}`));
      label(s, "antes", px, py + phH + 0.2, pw, C.background2, "Antes rótulo");
      T(s, "Reputação forte, mas restrita ao seu círculo. Quem precisa da sua experiência ainda não sabe que você existe.", { x: px, y: py + phH + 0.5, w: pw, h: 1.2, fontSize: 12, color: C.background2, valign: "top", lineSpacingMultiple: 1.12 }, "Antes texto");
    }
    // painel depois
    {
      const px = x0 + pw + gp;
      card(s, px, py, pw, phH, HEX.ardosia, "Depois painel");
      const cx = px + pw / 2, base = py + phH - 0.02;
      const h = person(cx, base, { fill: { color: HEX.mineral }, line: { color: HEX.mineral, width: 0 } }, "Depois pessoa");
      [0.62, 0.86, 1.1].forEach((r, i) => {
        const col = i === 0 ? HEX.rosa : HEX.lima;
        [[300, 60], [120, 240]].forEach((ang, k) =>
          s.addShape(pres.shapes.BLOCK_ARC, on({ x: h.hx - r, y: h.hy - r, w: 2 * r, h: 2 * r, angleRange: ang, arcThicknessRatio: 0.05 / r, fill: { color: col }, line: { color: col, width: 0 } }, `Depois onda ${i + 1}.${k + 1}`)));
      });
      [[0.18, 0.25], [0.6, 0.15], [1.9, 0.15], [2.3, 0.3], [0.12, 0.75], [2.35, 0.8]].forEach((p, i) =>
        dot(s, px + p[0], py + p[1], 0.16, i === 2 ? HEX.lima : HEX.mineral, `Depois público ${i + 1}`));
      label(s, "depois", px, py + phH + 0.2, pw, C.accent2, "Depois rótulo");
      T(s, "Sua história clara e constante, com voz própria, chegando a quem precisa de você e gerando conversas e negócios.", { x: px, y: py + phH + 0.5, w: pw, h: 1.2, fontSize: 12, color: C.background1, valign: "top", lineSpacingMultiple: 1.12 }, "Depois texto");
    }
    // seta entre os painéis
    const ad = 0.46, ax = x0 + pw + gp / 2 - ad / 2, ay = py + phH / 2 - ad / 2;
    await iconDot(s, "TbArrowRight", ax, ay, ad, C.text1, HEX.lima, "Seta antes depois");
  }

  // ============ 17 · Capa final
  {
    const s = pres.addSlide({ masterName: "Capa", sectionTitle: "Encerramento" });
    const lw = 5.6, lh = lw / LOGO.mineral.ratio;
    const ly = (7.5 - lh) / 2 - 0.35;
    s.addImage(on({ data: LOGO.mineral.data, x: (SW - lw) / 2, y: ly, w: lw, h: lh, altText: "Logotipo Meov" }, "Logotipo Meov"));
    T(s, "obrigado.", { x: 0, y: ly + lh + 0.7, w: SW, h: 0.45, fontFace: LIGHT, fontSize: 18, color: C.background2, align: "center" }, "Agradecimento");
    label(s, "@meov.mp · meov.com.br", 0, 6.75, SW, C.text2, "Contatos", "center");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("ok", OUT, JSON.stringify({ T_IMPL, T_MENS, CYCLE }));
})().catch((e) => { console.error(e); process.exit(1); });
