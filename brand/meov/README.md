# Branding Meov

Materiais de identidade visual da Meov.

A referência oficial é a **apresentação de identidade de marca 2026**: `apresentacao/apresentacao-branding-meov-r05-2026.pdf` (43 slides; `.pptx` editável ao lado). Ela segue a estrutura da apresentação original: estratégia (posicionamento, serviços, público, princípios e tom de voz), cores e textura, tipografia, logotipo e símbolo com malhas e conceito, resultado com as grades de variações do logotipo e do símbolo, e o universo visual com as aplicações principais.

- Para regerar: `apresentacao/gerador/`. `atualizar_aplicacoes.py` atualiza os mockups de 2023 (troca @meovelab, "experiências audiovisuais", "produtora de vídeos" e afins, sem distorcer as imagens); `ajustes_aplicacoes.py` faz os ajustes da R05 (notebook sem a frase na tela e banner novo do LinkedIn); `malhas.py` desenha as malhas de construção; `build-identidade.js` monta a apresentação. A foto do estúdio entra sozinha quando existir `gerador/img/aplicacoes/estudio.jpg`.
- A identidade anterior, de 2023, fica em `apresentacao/meov-identidade-visual-2023.pdf`. Logotipo, símbolo e pattern vêm dela; paleta, tipografia, voz e serviços foram atualizados.

```
brand/meov/
├── posicionamento.md   posicionamento atual da Meov
├── apresentacao/   apresentação de identidade 2026 (PDF + PPTX), gerador e a versão de 2023
├── logo/
│   ├── simbolo/    o "M" da Meov, na paleta atual
│   ├── logotipo/   a palavra "meov", na paleta atual
│   └── 2023/       arquivos originais nas cores de 2023
├── pattern/        grafismo de ondas
├── texturas/       fundos com degradê e granulado
├── tipografia/     Manrope + Geist Mono (guia em PDF)
└── paleta/         paleta atual e paleta de 2023
```

## Resumo da apresentação

**Conceito do símbolo:** `M + V + ∿`. O símbolo une o M e o V de "meov" a uma linha de batimento cardíaco, a ideia de experiências que fazem o coração pulsar. Na apresentação, o símbolo aparece acompanhado da palavra "experiências".

**Frases que aparecem na identidade**
- "Um portal para experiências audiovisuais indescritíveis."
- "Conexão através de experiências audiovisuais." (bio e site)
- "Filmes que tocam na alma."
- "A chave para um mundo de sensações inexploradas."
- "Inspirando um filme de cada vez."

**Estrutura da apresentação**

| Páginas | Seção | Conteúdo |
|---|---|---|
| 1–3 | Abertura | Capa e frase de apresentação |
| 4–7 | Cores | Paleta e os dois fundos com degradê |
| 8–9 | Tipografia | Chillax para textos |
| 10–11 | Logotipo | Malha de construção do "meov" |
| 12–14 | Símbolo | Malha de construção e conceito M + V + batimento |
| 15–20 | Resultado | Logotipo sobre fundos e logotipo com símbolo (marcados 100% e 25%) |
| 21–47 | Universo visual | Instagram, stories, site, LinkedIn, cartaz, fita com pattern, papelaria, cartões, proposta comercial, ecobags, mídia indoor, moletons, gorros, boné, óculos, garrafa, camisetas |
| 48 | Encerramento | Agradecimento |

**Padrões de aplicação observados**
- Fundo predominante: preto uva passa, ou o degradê preto uva passa → cinza tempestade com granulado.
- Logotipo em verde vibrante ou branco sobre fundo escuro; em preto uva passa sobre verde vibrante ou rosa choque.
- Rosa choque e verde vibrante aparecem como cores de destaque e de superfícies inteiras (cartões, camisetas, ecobags).
- O pattern de ondas é usado em preto sobre cor (fitas, ecobags) ou como onda gigante recortada no fundo (cartaz, capa do LinkedIn).
- Textos curtos, em minúsculas, alinhados à esquerda.
- O "®" acompanha o logotipo e o símbolo.

> A apresentação posiciona a Meov como **produtora de vídeos**. Esse posicionamento foi atualizado: ver [`posicionamento.md`](posicionamento.md).
> O Instagram oficial é **@meov.mp** (os mockups mostram `meovelab` e `meovlab`, que não valem mais).

## Paleta de cores (`paleta/`)

Paleta atual, com a proporção de uso de cada cor (revisada em outubro de 2026). A base é feita de branco, pedra e musgo; o **verde** é o destaque de apoio e o **rosa** aparece só em detalhes estratégicos.

| Nome | HEX | Proporção | Uso |
|---|---|---|---|
| Branco mineral | `#FAF8F4` | 34% | fundo de slides e peças com imagens e prints |
| Pedra líquen | `#D8DACD` | 32% | fundo principal, cartões e superfícies |
| Preto musgo | `#1C211E` | 27% | textos, fundos de elegância e capas |
| Cinza ardósia verde | `#59625A` | 6% | textos de apoio e linhas |
| Verde vibrante | `#DBFF00` | 0,8% | destaque de apoio: ícones, um cartão ou faixa por peça |
| Rosa choque | `#FF19B1` | 0,2% | só detalhes estratégicos: o ponto que guia o olhar, o número que importa |

Referência visual: `paleta/meov-paleta-atual.png` (gerada por `paleta/gerador/paleta.py`, com Manrope e Geist Mono instaladas). A faixa fina na base da imagem mostra a proporção real. A paleta parte da base da Aho, com o verde folha trocado pelo rosa da Meov.

**Paleta da identidade de 2023 (anterior):** preto uva passa `#202020`, verde vibrante `#DBFF00`, rosa choque `#FF19B1`, azul fluorescente `#31F8FF`, cinza tempestade `#4F5F66` e branco `#FFFFFF`. Referência: `paleta/meov-paleta-2023.png`. O pattern, as texturas e os logos em `logo/2023/` ainda usam essas cores.

## Tipografia (`tipografia/`)

Duas famílias, uma voz.

| Fonte | Uso | Pesos |
|---|---|---|
| **Manrope** | títulos e textos | Light 300 nos títulos · Regular 400 nos textos (também Medium 500 e SemiBold 600) |
| **Geist Mono** | dados e rótulos: números, valores, datas, @, sites e selos, em maiúsculas | Regular 400 |

**Regras de uso**
1. Títulos leves: Light 300 em títulos grandes. O destaque vem do espaço, não do negrito.
2. Letras próximas: −0,02 a −0,03 em nos títulos; texto com espaçamento normal.
3. Mono só para dados e rótulos, nunca em frases.
4. Dois pesos por peça (ex.: 300 no título e 400 no texto).
5. Cor do texto: preto musgo no principal, ardósia no apoio, branco sobre escuro.
6. Frases serenas: títulos curtos e calmos.

As duas fontes são gratuitas no Google Fonts.
- Web: `https://fonts.googleapis.com/css2?family=Manrope:wght@300;400;500;600&family=Geist+Mono:wght@400&display=swap`
- Referência completa: `tipografia/meov-guia-tipografia.pdf` (escala, regras e aplicações).

**Tipografia da identidade de 2023 (anterior):** Chillax (Fontshare). Referência: `tipografia/meov-tipografia-chillax-2023.png`. O logotipo "meov" é desenhado e não depende de fonte.

## Logo (`logo/`)

SVGs na paleta atual. As versões **com fundo** têm 91,44 × 91,44 mm, como os originais. As versões **sem fundo** são recortadas no desenho, para aplicar sobre fotos, vídeos e peças.

Visão geral de todas as versões: `logo/meov-logos-visao-geral.png`.

Combinações recomendadas: preto musgo sobre fundos claros, branco mineral ou verde vibrante sobre fundos escuros; rosa só como detalhe. Evite verde vibrante sobre fundos claros, porque o contraste é baixo.

### Símbolo (`logo/simbolo/`)

| Arquivo | Cor | Fundo |
|---|---|---|
| `meov-simbolo-musgo-fundo-mineral.svg` | Preto musgo | Branco mineral |
| `meov-simbolo-rosa-fundo-mineral.svg` | Rosa choque | Branco mineral |
| `meov-simbolo-musgo-fundo-pedra.svg` | Preto musgo | Pedra líquen |
| `meov-simbolo-mineral-fundo-musgo.svg` | Branco mineral | Preto musgo |
| `meov-simbolo-verde-fundo-musgo.svg` | Verde vibrante | Preto musgo |
| `meov-simbolo-rosa-fundo-musgo.svg` | Rosa choque | Preto musgo |
| `meov-simbolo-musgo-fundo-verde.svg` | Preto musgo | Verde vibrante |
| `meov-simbolo-musgo-fundo-rosa.svg` | Preto musgo | Rosa choque |
| `meov-simbolo-mineral-fundo-ardosia.svg` | Branco mineral | Cinza ardósia verde |
| `meov-simbolo-musgo-sem-fundo.svg` | Preto musgo | sem fundo (transparente) |
| `meov-simbolo-mineral-sem-fundo.svg` | Branco mineral | sem fundo (transparente) |
| `meov-simbolo-rosa-sem-fundo.svg` | Rosa choque | sem fundo (transparente) |
| `meov-simbolo-verde-sem-fundo.svg` | Verde vibrante | sem fundo (transparente) |

### Logotipo (`logo/logotipo/`)

| Arquivo | Cor | Fundo |
|---|---|---|
| `meov-logotipo-musgo-fundo-mineral.svg` | Preto musgo | Branco mineral |
| `meov-logotipo-rosa-fundo-mineral.svg` | Rosa choque | Branco mineral |
| `meov-logotipo-musgo-fundo-pedra.svg` | Preto musgo | Pedra líquen |
| `meov-logotipo-mineral-fundo-musgo.svg` | Branco mineral | Preto musgo |
| `meov-logotipo-verde-fundo-musgo.svg` | Verde vibrante | Preto musgo |
| `meov-logotipo-rosa-fundo-musgo.svg` | Rosa choque | Preto musgo |
| `meov-logotipo-musgo-fundo-verde.svg` | Preto musgo | Verde vibrante |
| `meov-logotipo-musgo-fundo-rosa.svg` | Preto musgo | Rosa choque |
| `meov-logotipo-mineral-fundo-ardosia.svg` | Branco mineral | Cinza ardósia verde |
| `meov-logotipo-musgo-sem-fundo.svg` | Preto musgo | sem fundo (transparente) |
| `meov-logotipo-mineral-sem-fundo.svg` | Branco mineral | sem fundo (transparente) |
| `meov-logotipo-rosa-sem-fundo.svg` | Rosa choque | sem fundo (transparente) |
| `meov-logotipo-verde-sem-fundo.svg` | Verde vibrante | sem fundo (transparente) |

### Versões de 2023 (`logo/2023/`)

Os arquivos originais da identidade de 2023 (CorelDRAW), nas cores antigas: verde vibrante, branco, rosa e azul fluorescente sobre preto uva passa, e as versões em preto e branco puros.

## Pattern (`pattern/`)

| Arquivo | Descrição |
|---|---|
| `meov-pattern-onda-preto.svg` | Faixa de ondas contínuas derivada do símbolo, em preto, sem fundo. 162,56 × 91,44 mm (16:9). |

## Texturas (`texturas/`)

Fundos JPG 1920 × 1080 (16:9) com degradê de preto uva passa para cinza tempestade e granulado.

**Fundos base**

| Arquivo | Descrição |
|---|---|
| `meov-textura-degrade-01.jpg` | Fundo 01: degradê suave, sem elementos |
| `meov-textura-degrade-02-ondas.jpg` | Fundo 02: degradê com a silhueta das ondas desfocada |

**Onda do pattern aplicada sobre os fundos**

| Cor da onda | Sobre o fundo 01 | Sobre o fundo 02 |
|---|---|---|
| Rosa choque | `meov-textura-onda-rosa-choque-fundo-01.jpg` | `meov-textura-onda-rosa-choque-fundo-02.jpg` |
| Verde vibrante | `meov-textura-onda-verde-vibrante-fundo-01.jpg` | `meov-textura-onda-verde-vibrante-fundo-02.jpg` |
| Azul fluorescente | `meov-textura-onda-azul-fluorescente-fundo-01.jpg` | `meov-textura-onda-azul-fluorescente-fundo-02.jpg` |
| Branco | `meov-textura-onda-branco-fundo-01.jpg` | `meov-textura-onda-branco-fundo-02.jpg` |
