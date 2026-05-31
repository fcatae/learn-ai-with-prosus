# Prosus Presentation Project

## REGRA OBRIGATÓRIA — Apresentações

**Sempre que o usuário pedir uma apresentação, deck, slides, PowerPoint ou arquivo .pptx:**

1. Use o skill `/create-prosus-presentation` se disponível
2. Se não disponível, aplique diretamente as regras de brand abaixo
3. NUNCA gere uma apresentação sem identidade visual Prosus

---

## Brand System Prosus (fonte: prosus.com/stylesheets/global.css)

### Cores (constante `C` no script)

```js
const C = {
  cobaltDark:    "060F76",  // fundo hero, headers
  cobaltPrime:   "1136A8",  // painéis secundários
  cobaltLight:   "1D5EDC",  // acentos, bordas KPI
  graphiteDark:  "191919",
  graphitePrime: "3F3F3F",
  graphiteLight: "666666",
  chloridePrime: "00AC8E",  // métricas positivas, teal
  chlorideLight: "00D3AF",
  argonPrime:    "AF015F",
  argonLight:    "D90276",  // regra de acento, magenta
  oxidePrime:    "FCB828",  // destaques dourados
  oxideLight:    "FCDA28",
  lithiumPrime:  "5715B0",
  lithiumLight:  "7729E0",
  whitePrime:    "F6F7F9",
  white:         "FFFFFF",
  textColor:     "4D4D4D",
};
```

### Fonte
`"Gill Sans MT"` — substituto de sistema mais próximo de Neuzeit Grotesk (fonte oficial Prosus via Adobe Fonts)

### Logo
- Arquivo: `prosus_logo.png` na raiz do projeto (567×152 px)
- Posição: **canto superior direito** de **todos os slides**
- Largura: `1.4"`, altura proporcional (`1.4 / (567/152)`)

```js
function addLogo(slide) {
  const w = 1.4, h = w / (567 / 152);
  slide.addImage({ path: path.join(__dirname, "prosus_logo.png"), x: 8.35, y: 0.18, w, h });
}
```

### Layout padrão dos slides

**Slide de título:**
- Painel cobalt escuro cobrindo ~58% esquerdo
- Faixa dourada (oxide) no topo (`h: 0.07`)
- Título branco grande (`fontSize: 44, bold, charSpacing: 2`)
- Regra argon separando título do subtítulo (`h: 0.045`)
- Cards de KPI à direita com borda colorida no topo

**Slides de conteúdo:**
- Barra cobalt gradient no topo (`h: 0.72`)
- Título em branco sobre a barra (`fontSize: 22, bold`)
- Logo sempre no canto superior direito

**Card de métrica (fundo branco):**
```js
function metricCard(slide, x, y, w, h, label, value, sub, accent) {
  slide.addShape(pptx.ShapeType.rect, { x:x+0.04, y:y+0.04, w, h, fill:{color:"D8D8D8"}, line:{color:"D8D8D8"} });
  slide.addShape(pptx.ShapeType.rect, { x, y, w, h, fill:{color:"FFFFFF"}, line:{color:"E0E0E0",width:0.5} });
  slide.addShape(pptx.ShapeType.rect, { x, y, w, h:0.055, fill:{color:accent}, line:{color:accent} });
  slide.addText([
    { text:value+"\n", options:{ fontSize:22, bold:true, color:accent, breakLine:true } },
    { text:label+"\n", options:{ fontSize:8.5, bold:true, color:"3F3F3F", breakLine:true } },
    { text:sub,        options:{ fontSize:8, color:"666666" } },
  ], { x:x+0.15, y:y+0.15, w:w-0.3, h:h-0.2, valign:"top" });
}
```

**Card de segmento (fundo cobalt escuro):**
```js
function segmentCard(slide, x, y, w, h, title, lines, accent) {
  slide.addShape(pptx.ShapeType.rect, { x, y, w, h, fill:{color:"060F76"}, line:{color:"1136A8",width:0.5} });
  slide.addShape(pptx.ShapeType.rect, { x, y, w:0.045, h, fill:{color:accent}, line:{color:accent} });
  const rows = [
    { text:title+"\n", options:{ fontSize:9.5, bold:true, color:accent, breakLine:true } },
    ...lines.map(l => ({ text:l+"\n", options:{ fontSize:8.5, color:"FFFFFF", breakLine:true } })),
  ];
  slide.addText(rows, { x:x+0.12, y:y+0.1, w:w-0.2, h:h-0.2, valign:"top" });
}
```

### Dependências (npm bloqueado — instalar do GitHub)

Verificar antes de usar:
```bash
node -e "require('pptxgenjs')" 2>/dev/null || echo "INSTALAR"
```

Se precisar instalar:
```bash
mkdir -p node_modules/pptxgenjs/dist node_modules/jszip/dist node_modules/setimmediate

curl -sL "https://raw.githubusercontent.com/gitbrent/PptxGenJS/v4.0.1/dist/pptxgen.cjs.js" \
     -o node_modules/pptxgenjs/dist/pptxgen.cjs.js
printf '{"name":"pptxgenjs","version":"4.0.1","main":"dist/pptxgen.cjs.js"}' \
     > node_modules/pptxgenjs/package.json

curl -sL "https://raw.githubusercontent.com/Stuk/jszip/main/dist/jszip.js" \
     -o node_modules/jszip/dist/jszip.js
printf '{"name":"jszip","version":"3.10.1","main":"dist/jszip.js"}' \
     > node_modules/jszip/package.json

curl -sL "https://raw.githubusercontent.com/YuzuJS/setImmediate/master/setImmediate.js" \
     -o node_modules/setimmediate/setImmediate.js
curl -sL "https://raw.githubusercontent.com/YuzuJS/setImmediate/master/package.json" \
     -o node_modules/setimmediate/package.json
```

### Nome dos arquivos de saída
`Prosus_<Tema>_<Ano>.pptx`  — ex: `Prosus_FY2025_Overview.pptx`

---

## Verificação antes de entregar

- [ ] `prosus_logo.png` presente na raiz
- [ ] `node create_presentation.js` executa sem erro
- [ ] Arquivo `.pptx` gerado > 50 KB
- [ ] Logo no canto superior direito em todos os slides
- [ ] Paleta de cores Prosus aplicada (não usar temas genéricos)
