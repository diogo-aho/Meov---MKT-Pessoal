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

// Paleta (base Aho, verde folha trocado pelo rosa Meov)
const HEX = {
  mineral: "FAF8F4", // 50%
  pedra: "D8DACD",   // 30%
  musgo: "1C211E",   // 9%
  ardosia: "59625A", // 8%
  rosa: "FF19B1",    // 2%
  lima: "DBFF00",    // 1%
  white: "FFFFFF",
};
const MONO = "Geist Mono";
const LIGHT = "Manrope Light";

const THEME = {
  name: "Meov MP",
  headFontFace: LIGHT,
  bodyFontFace: "Manrope",
  colors: {
    dk1: HEX.musgo, lt1: HEX.mineral, dk2: HEX.ardosia, lt2: HEX.pedra,
    accent1: HEX.rosa, accent2: HEX.lima, accent3: HEX.pedra, accent4: HEX.ardosia,
    accent5: HEX.musgo, accent6: HEX.white, hlink: HEX.rosa, folHlink: HEX.ardosia,
  },
};

async function svgArt(file, color) {
  let s = fs.readFileSync(path.join(BRAND, file), "utf8");
  s = s.replace(/<rect class="fil0"[^>]*\/>/, "");
  s = s.replace(/\.fil1 \{fill:[^}]*\}/, `.fil1 {fill:#${color}}`)
       .replace(/\.fil2 \{fill:[^;]*;/, `.fil2 {fill:#${color};`)
       .replace(/\.fil0 \{fill:[^}]*\}/, `.fil0 {fill:#${color}}`);
  const buf = await sharp(Buffer.from(s), { density: 300 }).trim().png().toBuffer();
  const m = await sharp(buf).metadata();
  return { data: "image/png;base64," + buf.toString("base64"), ratio: m.width / m.height };
}
async function icon(name, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(tb[name], { color: "#" + color, size, strokeWidth: 1.4 }));
  const buf = await sharp(Buffer.from(svg)).resize(size, size).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const LOGO = {
    musgo: await svgArt("logo/logotipo/meov-logotipo-lima-fundo-grafite.svg", HEX.musgo),
    mineral: await svgArt("logo/logotipo/meov-logotipo-lima-fundo-grafite.svg", HEX.mineral),
    lima: await svgArt("logo/logotipo/meov-logotipo-lima-fundo-grafite.svg", HEX.lima),
  };
  const SYM = { rosa: await svgArt("logo/simbolo/meov-simbolo-lima-fundo-grafite.svg", HEX.rosa) };
  const WAVE = {
    rosa: await svgArt("pattern/meov-pattern-onda-preto.svg", HEX.rosa),
    pedra: await svgArt("pattern/meov-pattern-onda-preto.svg", HEX.pedra),
    ardosia: await svgArt("pattern/meov-pattern-onda-preto.svg", HEX.ardosia),
  };

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.title = "Proposta Meov · Fernando Lohmann";
  pres.author = "Diogo & Rafa | Meov MP";
  pres.company = "Meov";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  const SW = 13.333, MX = 0.75, CW = SW - 2 * MX;

  const header = (dark) => [
    { text: { text: "PROPOSTA · FERNANDO LOHMANN · 2026", options: { x: SW - MX - 5, y: 0.38, w: 5, h: 0.25, fontFace: MONO, fontSize: 9, charSpacing: 1.5, color: dark ? C.background2 : C.text2, align: "right", margin: 0 } } },
  ];
  pres.defineSlideMaster({
    title: "Claro",
    background: { color: C.background1 },
    objects: [
      ...header(false),
      { image: { x: SW - MX - 0.95, y: 6.88, w: 0.95, h: 0.95 / LOGO.musgo.ratio, data: LOGO.musgo.data } },
      { placeholder: { options: { name: "kicker", type: "body", x: MX, y: 0.95, w: 8, h: 0.3, fontFace: MONO, fontSize: 11, charSpacing: 1.5, color: C.accent1, margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: MX, y: 1.25, w: CW, h: 0.8, fontFace: LIGHT, fontSize: 34, charSpacing: -0.6, color: C.text1, margin: 0, valign: "top", align: "left" }, text: "" } },
    ],
    slideNumber: { x: MX, y: 0.38, w: 0.6, h: 0.25, fontFace: MONO, fontSize: 9, color: C.text2, margin: 0 },
  });
  pres.defineSlideMaster({
    title: "Escuro",
    background: { color: C.text1 },
    objects: [...header(true)],
    slideNumber: { x: MX, y: 0.38, w: 0.6, h: 0.25, fontFace: MONO, fontSize: 9, color: C.background2, margin: 0 },
  });

  const on = (o, n) => { o.objectName = n; return o; };
  const T = (s, text, opts, name) => s.addText(text, on({ isTextBox: true, margin: 0, fontFace: "Manrope", ...opts }, name));
  const label = (s, text, x, y, w, color, name) => T(s, text.toUpperCase(), { x, y, w, h: 0.26, fontFace: MONO, fontSize: 10, charSpacing: 1.5, color }, name);
  const card = (s, x, y, w, h, fill, name) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, on({ x, y, w, h, rectRadius: 0.16, fill: { color: fill }, line: { color: fill, width: 0 } }, name));
  const dot = (s, x, y, d, fill, name) => s.addShape(pres.shapes.OVAL, on({ x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 } }, name));
  const iconDot = async (s, iname, x, y, d, fill, fg, name) => {
    dot(s, x, y, d, fill, name + " fundo");
    s.addImage(on({ data: await icon(iname, fg), x: x + d * 0.24, y: y + d * 0.24, w: d * 0.52, h: d * 0.52 }, name + " ícone"));
  };
  const light = (sec) => pres.addSlide({ masterName: "Claro", sectionTitle: sec });
  const dark = (sec) => pres.addSlide({ masterName: "Escuro", sectionTitle: sec });
  const body = (extra = {}) => ({ fontSize: 14, color: C.text2, valign: "top", lineSpacingMultiple: 1.15, ...extra });

  // ===================== 1 · Capa
  pres.addSection({ title: "Abertura" });
  {
    const s = dark("Abertura");
    s.addImage(on({ data: WAVE.rosa.data, x: 5.9, y: 5.6, w: 8.6, h: 8.6 / WAVE.rosa.ratio }, "Onda rosa"));
    s.addImage(on({ data: LOGO.mineral.data, x: MX, y: 1.0, w: 1.7, h: 1.7 / LOGO.mineral.ratio }, "Logotipo Meov"));
    label(s, "proposta comercial · outubro 2026", MX, 2.6, 8, C.accent2, "Rótulo");
    T(s, "Marca pessoal, do diagnóstico à presença constante.", { x: MX, y: 2.95, w: 8.6, h: 1.9, fontFace: LIGHT, fontSize: 46, charSpacing: -1, color: C.background1, valign: "top" }, "Título");
    T(s, [
      { text: "aos cuidados de ", options: { color: C.background2 } },
      { text: "Fernando Lohmann", options: { color: C.accent2 } },
    ], { x: MX, y: 5.0, w: 7, h: 0.45, fontSize: 18 }, "Destinatário");
  }

  // ===================== 2 · Ponto de partida
  {
    const s = light("Abertura");
    s.addText("01 · o ponto de partida".toUpperCase(), { placeholder: "kicker" });
    s.addText("Você já é referência. Falta o digital mostrar.", { placeholder: "title" });
    T(s, [
      { text: "Sua reputação foi construída com anos de entrega, clientes satisfeitos e indicações. Fora das redes, ninguém precisa ser convencido de quem você é.", options: { breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 8 } },
      { text: "No digital, a história é outra: quem te procura encontra pouco, ou encontra algo que não representa o profissional que você se tornou. Quase sempre não é falta de competência. É falta de clareza sobre o que mostrar, de tempo para produzir com constância e de segurança para se expor." },
    ], { x: MX, y: 2.25, w: 5.9, h: 2.4, ...body({ fontSize: 15 }) }, "Texto contexto");

    const gx = 7.15, gw = SW - MX - gx;
    card(s, gx, 2.25, gw, 2.4, C.background2, "Quadro barras");
    const bars = [
      { l: "reputação fora das redes", p: 0.9, c: C.text1 },
      { l: "presença no digital hoje", p: 0.2, c: C.accent1 },
    ];
    bars.forEach((b, i) => {
      const y = 2.5 + i * 0.85;
      label(s, b.l, gx + 0.35, y, gw - 0.7, C.text2, `Barra ${i + 1} rótulo`);
      card(s, gx + 0.35, y + 0.32, gw - 0.7, 0.36, C.background1, `Barra ${i + 1} trilho`);
      card(s, gx + 0.35, y + 0.32, (gw - 0.7) * b.p, 0.36, b.c, `Barra ${i + 1} valor`);
    });
    T(s, "A distância entre as duas barras é exatamente onde a Meov trabalha.", { x: gx + 0.35, y: 4.12, w: gw - 0.7, h: 0.4, fontSize: 13, color: C.text1 }, "Legenda barras");

    const outs = [
      { n: "Clareza", d: "Saber o que dizer, para quem e por quê. Cada conteúdo nasce de uma estratégia, não de um palpite.", i: "TbCompass" },
      { n: "Constância", d: "Presença toda semana, com produção organizada, sem depender da inspiração ou da agenda livre.", i: "TbWaveSine" },
      { n: "Segurança", d: "Aparecer sendo quem você é, sem personagem. Com direção de câmera e um time apoiando cada passo.", i: "TbShieldCheck" },
    ];
    const ow = (CW - 0.5) / 3;
    for (let i = 0; i < outs.length; i++) {
      const x = MX + i * (ow + 0.25), y = 4.95;
      await iconDot(s, outs[i].i, x, y, 0.55, C.background2, HEX.ardosia, `Ganho ${i + 1}`);
      T(s, outs[i].n, { x: x + 0.75, y: y + 0.08, w: ow - 0.8, h: 0.4, fontSize: 18, color: C.text1 }, `Ganho ${i + 1} nome`);
      T(s, outs[i].d, { x: x + 0.75, y: y + 0.5, w: ow - 0.8, h: 1.0, ...body({ fontSize: 13 }) }, `Ganho ${i + 1} texto`);
    }
  }

  // ===================== 3 · Quem somos
  pres.addSection({ title: "A Meov" });
  {
    const s = light("A Meov");
    s.addText("02 · quem somos".toUpperCase(), { placeholder: "kicker" });
    s.addText("Um time ao seu lado, da estratégia à câmera.", { placeholder: "title" });
    card(s, MX, 2.25, 4.75, 4.25, C.text1, "Quadro manifesto");
    s.addImage(on({ data: WAVE.rosa.data, x: MX + 0.4, y: 2.6, w: 2.2, h: 2.2 / WAVE.rosa.ratio }, "Onda manifesto"));
    T(s, "Não somos uma agência de marketing. Somos um time que apoia técnica e emocionalmente quem decide colocar a própria voz no mundo.", { x: MX + 0.4, y: 3.05, w: 3.95, h: 2.0, fontFace: LIGHT, fontSize: 19, color: C.background1, valign: "top", lineSpacingMultiple: 1.05 }, "Manifesto");
    T(s, "Por isso nosso trabalho começa antes do primeiro post, entendendo quem você é, e continua depois dele, com produção e gestão todos os meses.", { x: MX + 0.4, y: 5.3, w: 3.95, h: 1.05, fontSize: 13, color: C.background2, valign: "top", lineSpacingMultiple: 1.15 }, "Manifesto apoio");

    const fronts = [
      { n: "Estratégia", i: "TbCompass", d: "O Método Meov organiza quem você é, para quem você fala e como se diferencia no seu mercado. Tudo vira um Brandbook que guia os próximos 6 meses." },
      { n: "Audiovisual", i: "TbMovie", d: "Direção de imagem, gravação e edição com o cuidado de quem nasceu no audiovisual. Você aparece bem, com naturalidade, e do seu jeito." },
      { n: "Gestão das redes", i: "TbLayoutGrid", d: "Instagram, LinkedIn e Site/LP com calendário, roteiros, legendas, design, publicação e relatórios. Você grava e aprova; a gente cuida do resto." },
    ];
    const fx = 5.85, fw = SW - MX - fx;
    for (let i = 0; i < fronts.length; i++) {
      const y = 2.25 + i * 1.45;
      card(s, fx, y, fw, 1.38, C.background2, `Frente ${i + 1} cartão`);
      await iconDot(s, fronts[i].i, fx + 0.3, y + 0.3, 0.7, C.background1, i === 1 ? HEX.rosa : HEX.ardosia, `Frente ${i + 1}`);
      T(s, fronts[i].n, { x: fx + 1.25, y: y + 0.2, w: fw - 1.5, h: 0.35, fontSize: 17, color: C.text1 }, `Frente ${i + 1} nome`);
      T(s, fronts[i].d, { x: fx + 1.25, y: y + 0.52, w: fw - 1.5, h: 0.8, ...body({ fontSize: 13 }) }, `Frente ${i + 1} texto`);
    }
  }

  // ===================== 4 · Método e jornada
  {
    const s = light("A Meov");
    s.addText("03 · o método e a jornada".toUpperCase(), { placeholder: "kicker" });
    s.addText("Da raiz à colheita, em seis meses.", { placeholder: "title" });
    const mov = [
      { n: "Raiz", d: "Quem você é, por que quer aparecer e o que te torna referência." },
      { n: "Território", d: "Para quem você fala e qual espaço só você ocupa no mercado." },
      { n: "Voz", d: "Seu tom, sua imagem e o caminho de quem admira até quem contrata." },
      { n: "Movimento", d: "Temas, canais e plano para aparecer com constância." },
    ];
    const mw = (CW - 0.6) / 4;
    s.addShape(pres.shapes.LINE, on({ x: MX + 0.3, y: 2.5, w: CW - mw + 0.0, h: 0, line: { color: C.text2, width: 1, dashType: "dash" } }, "Linha movimentos"));
    mov.forEach((m, i) => {
      const x = MX + i * (mw + 0.2);
      dot(s, x, 2.2, 0.6, i === 3 ? C.accent1 : C.text1, `Movimento ${i + 1}`);
      T(s, String(i + 1).padStart(2, "0"), { x, y: 2.2, w: 0.6, h: 0.6, fontFace: MONO, fontSize: 13, color: C.background1, align: "center", valign: "middle" }, `Movimento ${i + 1} número`);
      T(s, m.n, { x, y: 2.92, w: mw, h: 0.4, fontSize: 18, color: C.text1 }, `Movimento ${i + 1} nome`);
      T(s, m.d, { x, y: 3.3, w: mw - 0.1, h: 0.7, ...body({ fontSize: 13 }) }, `Movimento ${i + 1} texto`);
    });
    T(s, "O resultado é o seu Brandbook: o guia da sua marca e do nosso trabalho conjunto.", { x: MX, y: 4.3, w: CW, h: 0.3, fontSize: 13, color: C.text1 }, "Nota Brandbook");

    const ph = [
      { m: "início", n: "Implantação", d: "método, ensaio e canais prontos" },
      { m: "mês 1", n: "Fundação", d: "primeiros conteúdos no ar" },
      { m: "mês 2", n: "Ritmo", d: "constância e séries fixas" },
      { m: "mês 3", n: "Leitura", d: "1ª revisão de rota", hl: true },
      { m: "mês 4", n: "Autoridade", d: "casos, provas e opinião" },
      { m: "mês 5", n: "Autoridade", d: "aprofundar e converter" },
      { m: "mês 6", n: "Colheita", d: "resultado e novo ciclo", hl: true },
    ];
    const pw = (CW - 6 * 0.12) / 7;
    ph.forEach((p, i) => {
      const x = MX + i * (pw + 0.12), y = 4.8;
      const fill = i === 0 ? C.text1 : p.hl ? C.accent1 : C.background2;
      const tc = i === 0 || p.hl ? C.background1 : C.text1;
      const sc = i === 0 || p.hl ? C.background1 : C.text2;
      card(s, x, y, pw, 1.75, fill, `Fase ${i} cartão`);
      label(s, p.m, x + 0.18, y + 0.2, pw - 0.3, sc, `Fase ${i} rótulo`);
      T(s, p.n, { x: x + 0.18, y: y + 0.55, w: pw - 0.3, h: 0.4, fontSize: 15, color: tc }, `Fase ${i} nome`);
      T(s, p.d, { x: x + 0.18, y: y + 0.95, w: pw - 0.3, h: 0.7, fontSize: 12, color: sc, valign: "top" }, `Fase ${i} texto`);
    });
  }

  // ===================== 5 · Implantação
  pres.addSection({ title: "Escopo" });
  {
    const s = light("Escopo");
    s.addText("04 · parte 1 · implantação".toUpperCase(), { placeholder: "kicker" });
    s.addText("Primeiro, construímos a base.", { placeholder: "title" });
    T(s, "Antes do primeiro conteúdo, organizamos sua marca e deixamos cada canal pronto para receber quem vai te encontrar.", { x: MX, y: 1.85, w: CW, h: 0.35, ...body({ fontSize: 14 }) }, "Intro implantação");
    const items = [
      { n: "Método Meov", i: "TbCompass", d: "Diagnóstico da sua presença atual e de como você é percebido, imersão presencial, Brandbook completo e plano de 6 meses. Regras do seu conselho ou da sua empresa entram desde o início." },
      { n: "Ensaio de marca", i: "TbCamera", d: "Uma sessão de fotos com direção de imagem e 12 fotos tratadas, pensadas para perfil, banner, site e conteúdos. Imagem que comunica antes de qualquer palavra." },
      { n: "Instagram", i: "TbBrandInstagram", d: "Conta de criador, bio, link, destaques, 3 posts fixados e templates próprios. Segurança com verificação em duas etapas e acesso pelo Meta Business Suite, sem troca de senha." },
      { n: "LinkedIn", i: "TbBrandLinkedin", d: "Headline, Sobre e experiências como conquistas, banner, URL, Em destaque, competências e modo criador. O perfil passa a contar a mesma história do conteúdo." },
      { n: "Site/LP", i: "TbWorld", d: "Uma página com sua apresentação, história, como você ajuda, provas, chamada para o WhatsApp e links para todas as suas redes. O lugar onde a decisão de te procurar se confirma." },
      { n: "Também inclui", i: "TbChecks", d: "Fluxo de aprovação de conteúdos combinado entre você e a Meov, e autorização de uso de imagem assinada, para trabalharmos com tranquilidade.", dk: true },
    ];
    const cw = (CW - 0.5) / 3, ch = 2.1;
    for (let i = 0; i < items.length; i++) {
      const it = items[i], x = MX + (i % 3) * (cw + 0.25), y = 2.35 + Math.floor(i / 3) * (ch + 0.18);
      card(s, x, y, cw, ch, it.dk ? C.text1 : C.background2, `Entrega ${i + 1} cartão`);
      await iconDot(s, it.i, x + 0.28, y + 0.25, 0.5, it.dk ? C.accent1 : C.background1, it.dk ? HEX.white : HEX.ardosia, `Entrega ${i + 1}`);
      T(s, it.n, { x: x + 0.92, y: y + 0.3, w: cw - 1.1, h: 0.4, fontSize: 16, color: it.dk ? C.background1 : C.text1 }, `Entrega ${i + 1} nome`);
      T(s, it.d, { x: x + 0.28, y: y + 0.85, w: cw - 0.5, h: ch - 1.0, fontSize: 12, color: it.dk ? C.background2 : C.text2, valign: "top", lineSpacingMultiple: 1.12 }, `Entrega ${i + 1} texto`);
    }
  }

  // ===================== 6 · Acompanhamento
  {
    const s = light("Escopo");
    s.addText("05 · parte 2 · acompanhamento mensal".toUpperCase(), { placeholder: "kicker" });
    s.addText("Depois, presença constante todos os meses.", { placeholder: "title" });
    const stats = [
      { v: "12", l: "reels editados, com cortes, legendas e trilha" },
      { v: "8", l: "carrosséis com texto e design" },
      { v: "4", l: "posts de texto no LinkedIn, a partir das gravações" },
      { v: "2", l: "gravações de meio período com direção" },
    ];
    const sw = 2.55, sh = 1.75;
    stats.forEach((st, i) => {
      const x = MX + (i % 2) * (sw + 0.2), y = 2.2 + Math.floor(i / 2) * (sh + 0.2);
      card(s, x, y, sw, sh, C.background2, `Número ${i + 1} cartão`);
      T(s, st.v, { x: x + 0.3, y: y + 0.12, w: sw - 0.5, h: 0.9, fontFace: LIGHT, fontSize: 54, color: i === 0 ? C.accent1 : C.text1 }, `Número ${i + 1} valor`);
      T(s, st.l, { x: x + 0.3, y: y + 1.0, w: sw - 0.5, h: 0.65, fontSize: 12, color: C.text2, valign: "top" }, `Número ${i + 1} rótulo`);
    });
    label(s, "mínimos por mês", MX, 6.17, 5, C.text2, "Nota mínimos");

    const rx = 6.4, rw = SW - MX - rx;
    card(s, rx, 2.2, rw, 1.0, C.accent1, "Quadro boas-vindas");
    await iconDot(s, "TbGift", rx + 0.25, 2.42, 0.56, C.background1, HEX.rosa, "Presente");
    T(s, [
      { text: "Boas-vindas: ", options: { color: C.background1 } },
      { text: "nos 3 primeiros meses, a gravação é semanal (4 por mês), sem custo adicional. Mais material quando a sua marca está nascendo.", options: { color: C.background1 } },
    ], { x: rx + 1.0, y: 2.3, w: rw - 1.2, h: 0.8, fontSize: 13, valign: "middle" }, "Texto boas-vindas");

    label(s, "como cada mês acontece", rx, 3.45, rw, C.text2, "Rótulo ciclo");
    const steps = [
      { n: "Pauta e roteiro", d: "a partir do seu Brandbook e do que acontece no seu mercado" },
      { n: "Gravação", d: "encontros com direção de imagem, no seu ritmo" },
      { n: "Edição e design", d: "reels finalizados e carrosséis prontos para aprovar" },
      { n: "Publicação", d: "Instagram e LinkedIn, com o formato certo para cada rede" },
      { n: "Relatório e reunião", d: "o que funcionou, o que ajustar e o plano seguinte" },
    ];
    steps.forEach((st, i) => {
      const y = 3.8 + i * 0.5;
      dot(s, rx, y, 0.36, i === 1 ? C.accent1 : C.text1, `Etapa ${i + 1}`);
      T(s, String(i + 1), { x: rx, y, w: 0.36, h: 0.36, fontFace: MONO, fontSize: 10, color: C.background1, align: "center", valign: "middle" }, `Etapa ${i + 1} número`);
      T(s, [{ text: st.n + "  ", options: { color: C.text1 } }, { text: st.d, options: { color: C.text2 } }], { x: rx + 0.55, y: y - 0.02, w: rw - 0.55, h: 0.4, fontSize: 13, valign: "middle" }, `Etapa ${i + 1} texto`);
    });
  }

  // ===================== 7 · Como medimos
  {
    const s = light("Escopo");
    s.addText("06 · como medimos".toUpperCase(), { placeholder: "kicker" });
    s.addText("Não medimos só seguidores.", { placeholder: "title" });
    T(s, "Número de seguidores é só uma parte da história. Acompanhamos cinco camadas, sempre comparadas com o ponto de partida levantado no diagnóstico, e revisamos a rota no 3º e no 6º mês.", { x: MX, y: 1.85, w: CW, h: 0.6, ...body({ fontSize: 14 }) }, "Intro medição");
    const layers = [
      { n: "Presença", d: "Você está aparecendo com constância? Medimos o calendário cumprido e as séries mantidas." },
      { n: "Alcance qualificado", d: "As pessoas certas estão te vendo? Olhamos quem chega, não só quantos." },
      { n: "Conversa", d: "Elas salvam, compartilham, comentam e te chamam? Sinal de vínculo." },
      { n: "Negócio", d: "Isso vira reuniões, convites e clientes? A camada que paga o investimento." },
      { n: "Confiança", d: "Você se sente mais seguro sendo quem é? Sem ela, nada se sustenta." },
    ];
    const fills = [C.background2, C.background2, C.background2, C.text1, C.accent1];
    layers.forEach((l, i) => {
      const w = 5.6 - i * 0.55, x = MX + (5.6 - w) / 2, y = 2.75 + i * 0.75;
      card(s, x, y, w, 0.62, fills[i], `Camada ${i + 1}`);
      T(s, l.n, { x, y, w, h: 0.62, fontSize: 15, color: i >= 3 ? C.background1 : C.text1, align: "center", valign: "middle" }, `Camada ${i + 1} nome`);
      T(s, l.d, { x: 6.85, y: y + 0.04, w: SW - MX - 6.85, h: 0.56, fontSize: 13, color: C.text2, valign: "middle" }, `Camada ${i + 1} texto`);
    });
  }

  // ===================== 8 · Investimento
  pres.addSection({ title: "Investimento" });
  {
    const s = dark("Investimento");
    label(s, "07 · investimento", MX, 0.95, 6, C.accent2, "Rótulo");
    T(s, "Um ciclo completo, do começo à colheita.", { x: MX, y: 1.25, w: CW, h: 0.8, fontFace: LIGHT, fontSize: 34, charSpacing: -0.6, color: C.background1 }, "Título");
    const boxes = [
      { k: "implantação", v: "R$ 12.000", d: "Método Meov, ensaio de marca, Instagram, LinkedIn e Site/LP. Em 3 parcelas de R$ 4.000, junto às 3 primeiras mensalidades.", c: C.accent2 },
      { k: "acompanhamento", v: "R$ 12.000/mês", d: "Por 6 meses: gravações com direção, reels, carrosséis, posts no LinkedIn e gestão completa das redes.", c: C.accent1 },
    ];
    boxes.forEach((b, i) => {
      const y = 2.3 + i * 2.0;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, on({ x: MX, y, w: 6.1, h: 1.8, rectRadius: 0.16, fill: { color: HEX.musgo }, line: { color: HEX.ardosia, width: 1 } }, `Valor ${i + 1} quadro`));
      label(s, b.k, MX + 0.35, y + 0.25, 5, b.c, `Valor ${i + 1} rótulo`);
      T(s, b.v, { x: MX + 0.35, y: y + 0.5, w: 5.5, h: 0.65, fontFace: LIGHT, fontSize: 36, color: C.background1 }, `Valor ${i + 1}`);
      T(s, b.d, { x: MX + 0.35, y: y + 1.15, w: 5.45, h: 0.6, fontSize: 12, color: C.background2, valign: "top" }, `Valor ${i + 1} texto`);
    });
    s.addChart(pres.charts.BAR, [
      { name: "Acompanhamento", labels: ["mês 1", "mês 2", "mês 3", "mês 4", "mês 5", "mês 6"], values: [12, 12, 12, 12, 12, 12] },
      { name: "Implantação", labels: ["mês 1", "mês 2", "mês 3", "mês 4", "mês 5", "mês 6"], values: [4, 4, 4, 0, 0, 0] },
    ], on({
      x: 7.25, y: 2.3, w: SW - MX - 7.25, h: 3.8, barDir: "col", barGrouping: "stacked", barGapWidthPct: 60,
      chartColors: [HEX.pedra, HEX.rosa],
      showValue: true, dataLabelPosition: "ctr", dataLabelColor: HEX.musgo, dataLabelFontSize: 11, dataLabelFontFace: "+mn-lt", dataLabelFormatCode: '0" mil";;;',
      showTitle: true, title: "valor por mês (R$)", titleFontSize: 11, titleColor: HEX.pedra, titleFontFace: "+mn-lt",
      showLegend: true, legendPos: "b", legendFontSize: 11, legendColor: HEX.pedra, legendFontFace: "+mn-lt",
      catAxisLabelColor: HEX.pedra, catAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt", catAxisLineShow: false,
      valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      plotArea: { fill: { color: HEX.musgo } }, chartArea: { fill: { color: HEX.musgo } },
    }, "Gráfico de pagamento"));
    T(s, [
      { text: "MESES 1–3  ", options: { color: C.background2 } }, { text: "R$ 16.000/mês", options: { color: C.background1 } },
      { text: "      MESES 4–6  ", options: { color: C.background2 } }, { text: "R$ 12.000/mês", options: { color: C.background1 } },
      { text: "      TOTAL DO CICLO  ", options: { color: C.background2 } }, { text: "R$ 84.000", options: { color: C.accent2 } },
    ], { x: MX, y: 6.45, w: CW, h: 0.35, fontFace: MONO, fontSize: 12, charSpacing: 0.5 }, "Resumo pagamento");
  }

  // ===================== 9 · Combinados
  pres.addSection({ title: "Combinados" });
  {
    const s = light("Combinados");
    s.addText("08 · combinados".toUpperCase(), { placeholder: "kicker" });
    s.addText("O que fica de fora, para ficar claro desde o começo.", { placeholder: "title" });
    const groups = [
      { n: "Produção", d: "Tráfego pago · deslocamento fora da Grande SP · locação, maquiagem e figurino · modelos, drone e equipe extra · vídeos longos e eventos" },
      { n: "Gestão e canais", d: "Responder comentários e DMs (fica com você, para manter sua voz real) · stories do dia a dia · outras redes · página da empresa · e-mail marketing" },
      { n: "Comunicação", d: "Assessoria de imprensa e gestão de crise · parcerias com influenciadores · tradução e outros idiomas · trilhas licenciadas pagas" },
      { n: "Site", d: "Site com várias páginas, blog ou loja · domínio, hospedagem e manutenção" },
      { n: "Marca e jurídico", d: "Criação de logotipo · registro no INPI · revisão jurídica e compliance · autorização de imagem de terceiros" },
      { n: "Resultados", d: "Não garantimos seguidores, views ou vendas: garantimos método, qualidade e constância. Nunca compramos seguidores." },
    ];
    const cw = (CW - 0.5) / 3, ch = 1.8;
    groups.forEach((g, i) => {
      const x = MX + (i % 3) * (cw + 0.25), y = 2.15 + Math.floor(i / 3) * (ch + 0.18);
      card(s, x, y, cw, ch, C.background2, `Fora ${i + 1} cartão`);
      label(s, g.n, x + 0.28, y + 0.22, cw - 0.5, C.text1, `Fora ${i + 1} nome`);
      T(s, g.d, { x: x + 0.28, y: y + 0.55, w: cw - 0.5, h: ch - 0.65, fontSize: 13, color: C.text2, valign: "top", lineSpacingMultiple: 1.12 }, `Fora ${i + 1} texto`);
    });
    label(s, "como trabalhamos", MX, 6.05, 4, C.text2, "Rótulo combinados");
    const rules = ["Ciclo de 6 meses", "Até 2 rodadas de ajuste", "Mínimos mensais sem acúmulo", "Extras sob consulta", "Portfólio, salvo sigilo"];
    let x = MX;
    rules.forEach((r, i) => {
      const w = 0.45 + r.length * 0.08;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, on({ x, y: 6.35, w, h: 0.42, rectRadius: 0.21, fill: { color: C.background1 }, line: { color: C.text2, width: 0.75 } }, `Combinado ${i + 1} fundo`));
      T(s, r, { x, y: 6.35, w, h: 0.42, fontSize: 11, color: C.text1, align: "center", valign: "middle" }, `Combinado ${i + 1}`);
      x += w + 0.12;
    });
  }

  // ===================== 10 · Próximos passos + assinatura
  pres.addSection({ title: "Encerramento" });
  {
    const s = dark("Encerramento");
    s.addImage(on({ data: WAVE.rosa.data, x: -1.2, y: 5.55, w: 9.0, h: 9.0 / WAVE.rosa.ratio }, "Onda rosa"));
    label(s, "09 · próximos passos", MX, 0.95, 6, C.accent2, "Rótulo");
    T(s, "Você já tem a voz.", { x: MX, y: 1.5, w: 6.5, h: 1.0, fontFace: LIGHT, fontSize: 48, charSpacing: -1, color: C.background1 }, "Frase 1");
    T(s, "A gente ajuda você a colocá-la no mundo.", { x: MX, y: 2.5, w: 6.2, h: 1.1, fontFace: LIGHT, fontSize: 28, color: C.accent2, valign: "top" }, "Frase 2");
    T(s, "Sua voz, no ritmo certo.", { x: MX, y: 3.75, w: 6, h: 0.4, fontSize: 15, color: C.background2 }, "Frase 3");

    const px = 7.6, pw = SW - MX - px;
    const steps = [
      { n: "Aprovação da proposta", d: "Você valida o escopo e o investimento." },
      { n: "Assinatura do contrato", d: "Formalizamos o ciclo de 6 meses." },
      { n: "Diagnóstico e imersão", d: "Agendamos o primeiro encontro presencial." },
    ];
    steps.forEach((st, i) => {
      const y = 1.55 + i * 1.05;
      dot(s, px, y, 0.6, i === 0 ? C.accent2 : C.background1, `Passo ${i + 1}`);
      T(s, String(i + 1), { x: px, y, w: 0.6, h: 0.6, fontFace: MONO, fontSize: 14, color: C.text1, align: "center", valign: "middle" }, `Passo ${i + 1} número`);
      T(s, [{ text: st.n, options: { breakLine: true, color: C.background1, fontSize: 17 } }, { text: st.d, options: { color: C.background2, fontSize: 13 } }], { x: px + 0.85, y: y - 0.05, w: pw - 0.85, h: 0.75, valign: "middle" }, `Passo ${i + 1} texto`);
    });

    T(s, "Diogo & Rafa | Meov MP", { x: px, y: 5.0, w: pw, h: 0.45, fontSize: 18, color: C.background1 }, "Assinatura");
    T(s, "@MEOV.CREATORS · MEOV.COM.BR", { x: px, y: 5.45, w: pw, h: 0.3, fontFace: MONO, fontSize: 10, charSpacing: 1.5, color: C.background2 }, "Contatos");
    s.addImage(on({ data: LOGO.mineral.data, x: SW - MX - 1.3, y: 6.55, w: 1.3, h: 1.3 / LOGO.mineral.ratio }, "Logotipo"));
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("ok", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
