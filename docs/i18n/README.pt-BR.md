<p align="center">
  <a href="https://genoffice.ai/">
    <picture>
      <source srcset="../assets/readme/hero-dark.webp" media="(prefers-color-scheme: dark)">
      <img src="../assets/readme/hero.webp" alt="GenOffice — a suíte de escritório com IA de código aberto: Docs, Sheets, Slides, PDF, Markdown e HTML com um painel de IA integrado" width="100%">
    </picture>
  </a>
</p>

<h1 align="center">GenOffice</h1>

<p align="center"><b>A primeira suíte de escritório open source com IA completa do mundo.</b><br>
Arquivos do Word, Excel, PowerPoint e PDF, editados por você e sua IA, salvos de volta nos formatos reais.</p>

<p align="center">
  <a href="../../LICENSE"><img src="https://img.shields.io/github/license/genspark-ai/genoffice" alt="Licença: Apache-2.0"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases/latest"><img src="https://img.shields.io/github/v/release/genspark-ai/genoffice" alt="Versão mais recente"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases"><img src="https://img.shields.io/github/downloads/genspark-ai/genoffice/total" alt="Downloads"></a>
  <a href="https://github.com/genspark-ai/genoffice/stargazers"><img src="https://img.shields.io/github/stars/genspark-ai/genoffice?style=flat" alt="Estrelas no GitHub"></a>
</p>

<p align="center"><a href="../../README.md">English</a> · <a href="README.zh-CN.md">简体中文</a> · <a href="README.zh-TW.md">繁體中文</a> · <a href="README.ja.md">日本語</a> · <a href="README.ko.md">한국어</a> · <a href="README.de.md">Deutsch</a> · <a href="README.fr.md">Français</a> · <a href="README.es.md">Español</a> · <b>Português (Brasil)</b> · <a href="README.ru.md">Русский</a> · <a href="README.it.md">Italiano</a> · <a href="README.nl.md">Nederlands</a> · <a href="README.pl.md">Polski</a> · <a href="README.cs.md">Čeština</a> · <a href="README.id.md">Bahasa Indonesia</a> · <a href="README.ms.md">Bahasa Melayu</a> · <a href="README.th.md">ไทย</a> · <a href="README.hi.md">हिन्दी</a> · <a href="README.ar.md">العربية</a> · <a href="README.he.md">עברית</a></p>

<p align="center">
  <a href="#download"><b>Download</b></a> ·
  <a href="https://genoffice.ai/"><b>Site</b></a> ·
  <a href="https://genoffice.ai/join"><b>Comunidade</b></a> ·
  <a href="../../PRIVACY.md"><b>Privacidade</b></a>
</p>

O GenOffice é uma alternativa gratuita e de código aberto ao Microsoft Office
para macOS, Windows e Linux. Ele abre e salva arquivos nativos `.docx`,
`.xlsx` e `.pptx`, edita PDF, Markdown e HTML, e coloca um agente de IA ao
lado de cada documento — não um chat encaixado na lateral, mas um editor que
lê o arquivo, faz a alteração e mostra exatamente o que foi tocado.

- **Formatos reais, preservando cada byte.** Só o que você edita é reescrito.
  Todo o restante do arquivo permanece intacto byte a byte, então os
  documentos continuam funcionando no Word, no Excel e no PowerPoint.
- **IA que você pode revisar.** As edições chegam como alterações rastreadas e
  diffs, com reversão em um clique. As planilhas recebem fórmulas ao vivo, não
  números colados. Apresentações e páginas são geradas diretamente no canvas e
  permanecem totalmente editáveis.
- **Local por design.** Os arquivos são abertos, editados, salvos e
  convertidos na sua máquina. PDF → Word / Excel / PowerPoint, Markdown →
  Word e HTML → Word rodam localmente. Somente as chamadas de IA saem da
  máquina, para o provedor que você escolher.
- **Suas chaves, ou nenhuma.** Faça login com o Genspark e pule as chaves, ou
  use sua própria chave para Claude, OpenAI, Gemini, DeepSeek, Kimi, GLM,
  Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter, ou qualquer endpoint
  compatível com OpenAI, incluindo servidores locais.

**Baixe:** [macOS](https://github.com/genspark-ai/genoffice/releases/latest) (Apple Silicon e Intel) ·
[Windows](https://github.com/genspark-ai/genoffice/releases/latest) (x64 e Arm) ·
[Linux](https://github.com/genspark-ai/genoffice/releases/latest) (deb, rpm, AppImage) —
detalhes e requisitos em [Download](#download).

## Demo

Seis aplicativos, um painel de IA. Cada captura de tela é do aplicativo real
no macOS, com a IA acionada a partir do prompt que você pode ler no painel.

### 1 · Docs — abra e edite `.docx` com uma IA que você pode revisar

<table>
<tr>
<td width="50%"><img src="../assets/readme/docs-report.webp" alt="GenOffice Docs renderizando uma página de relatório anual em duas colunas, com imagem de capa em largura total, tabela de KPIs sombreada, cabeçalho e rodapé, em zoom de 80% com o painel de IA recolhido"></td>
<td width="50%"><img src="../assets/readme/docs-ai.webp" alt="GenOffice Docs: uma visão geral de empresa com uma imagem de banner; a IA condensou a seção Overview e inseriu uma nova seção com marcadores, e o painel oferece uma reversão em um clique"></td>
</tr>
<tr>
<td><b>Abre o arquivo como o Word o organiza</b> — seções em duas colunas, imagens de sangria total, tabelas sombreadas, cabeçalhos e rodapés, paginação seguindo as métricas de linha do Word. Estilos, comentários, alterações rastreadas, equações e tinta digital vão e voltam intactos.</td>
<td><b>Peça a edição</b> — a IA lê os blocos de que precisa, reescreve a seção Overview e insere uma nova seção com marcadores. Cada turno da IA é um snapshot que você pode reverter; com o <b>Controle de alterações</b> ativado, as edições chegam como revisões no estilo do Word.</td>
</tr>
</table>

### 2 · Sheets — `.xlsx` com fórmulas e gráficos ao vivo, não números colados

<table>
<tr>
<td width="50%"><img src="../assets/readme/sheets-ai.webp" alt="GenOffice Sheets: a IA adicionou uma planilha Summary com receita por região e categoria usando fórmulas SUMIF, além de um gráfico de colunas, e relata 43 alterações aplicadas com um botão Desfazer"></td>
<td width="50%"><img src="../assets/readme/sheets-qa.webp" alt="GenOffice Sheets: ao ser perguntada qual região liderou a receita do Q2, a IA responde Europa com o detalhamento por categoria e cita as células usadas como links, ao lado da planilha Orders"></td>
</tr>
<tr>
<td><b>Construa</b> — a partir de uma única frase, o agente adiciona uma planilha Summary com <code>SUMIF</code>s reais por região e categoria, insere um gráfico de colunas e aplica as 43 alterações como um único lote que pode ser desfeito.</td>
<td><b>Pergunte</b> — perguntas sobre a pasta de trabalho voltam com o raciocínio e as células exatas usadas como citações clicáveis. Por baixo do capô: um mecanismo `.xlsx` próprio em Rust, tabelas dinâmicas, segmentações de dados, formatação condicional e rastreamento de fórmulas.</td>
</tr>
</table>

### 3 · Slides — de um prompt a uma apresentação `.pptx`

<img src="../assets/readme/slides-generate.webp" alt="Timelapse do GenOffice Slides gerando o deck de investidores Aurora Home: a IA planeja a narrativa no painel, os slides aparecem no canvas um após o outro, e o deck finalizado termina no pedido de fechamento" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/slides-cover.webp" alt="GenOffice Slides: o slide de capa de uma apresentação para investidores da Aurora Home gerada por IA no canvas, com o prompt original de uma linha e o resumo da IA sobre o que construiu no painel"></td>
<td width="50%"><img src="../assets/readme/slides-ai.webp" alt="GenOffice Slides: o slide final projetado da mesma apresentação de 11 slides, com a faixa de miniaturas à esquerda e o painel de IA resumindo a narrativa"></td>
</tr>
<tr>
<td><b>Uma linha de entrada</b> — "Crie uma apresentação de 10 slides para investidores da Aurora Home…". O GenOffice planeja a narrativa, pesquisa os números e redige cada slide diretamente no canvas como um `.pptx` real.</td>
<td><b>Uma apresentação finalizada</b> — onze slides projetados com tipografia e imagens consistentes e uma chamada para ação de encerramento; continue editando com slides mestres, layouts, guias inteligentes e recorte não destrutivo, ou peça ao painel para restilizar, reescrever e reordenar.</td>
</tr>
</table>

### 4 · PDF — edite texto em PDF no lugar, converta PDF para Word no dispositivo

<table>
<tr>
<td width="50%"><img src="../assets/readme/pdf-edit.webp" alt="GenOffice PDF: o modo Editar texto contorna cada bloco de texto da página para edição no lugar, enquanto o painel de IA responde a uma pergunta sobre o relatório com citações de página"></td>
<td width="50%"><img src="../assets/readme/pdf-convert.webp" alt="GenOffice Docs mostrando um documento do Word convertido localmente a partir do PDF de revisão trimestral da Helios, aberto em uma segunda aba ao lado do PDF original"></td>
</tr>
<tr>
<td><b>Edite dentro da página</b> — o modo Editar texto contorna cada bloco de texto para redigitação no lugar; o content stream é reescrito via PDFium com as fontes originais, não uma anotação de encobrimento. Pergunte à IA sobre um relatório extenso e receba respostas com citações de página.</td>
<td><b>Converta no dispositivo</b> — <b>Conversor de PDF → PDF para Word</b> produz um `.docx` editável que abre no Docs ao lado do original, com títulos, linhas de estatísticas e parágrafos intactos. Os destinos Excel e PowerPoint funcionam da mesma forma; páginas escaneadas passam pelo OCR do sistema.</td>
</tr>
</table>

### 5 · HTML — um construtor de páginas e interfaces com IA, briefing de design primeiro

Diga para que serve a página e para quem ela é. A IA propõe primeiro um
**briefing de design** — gancho, paleta, tipografia e direções de estilo —
e então constrói um único arquivo `.html` autocontido a partir desses
tokens.

<img src="../assets/readme/html-restyle-motion.webp" alt="Timelapse do GenOffice HTML reestilizando a landing page Lumen: um pedido de Restyle no painel transforma a página escura Midnight Studio na versão quente Solar Daybreak, mantendo todas as seções e todo o texto no lugar" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/html-ai.webp" alt="GenOffice HTML: uma landing page gerada para uma luminária solar de mesa na direção escura Midnight Studio, mostrada na prévia em tempo real com o painel de IA resumindo a página que acabou de construir"></td>
<td width="50%"><img src="../assets/readme/html-restyle.webp" alt="A mesma landing page da Lumen restilizada pela IA na direção quente Solar Daybreak: fundo em tom de papel, títulos serifados e um destaque laranja, com todas as seções e todo o texto preservados"></td>
</tr>
<tr>
<td><b>Gerada a partir de um prompt</b> — um hero em destaque, cartões de recursos, preços e um formulário de lista de espera para a Lumen, construídos na direção Midnight Studio. Clique em qualquer elemento para restilizá-lo, dê duplo clique para editar o texto, ou alterne para a visualização de código no CodeMirror.</td>
<td><b>Mesmo design, nova direção</b> — um único pedido de <b>Restilizar</b> troca os tokens do briefing e a página acompanha: papel quente, serifa editorial, destaque laranja-sol, nada reescrito. Apresente em tela cheia, ou exporte como PDF ou um documento do Word nativo e editável.</td>
</tr>
</table>
<table>
<tr>
<td width="50%"><img src="../assets/readme/html-dashboard.webp" alt="GenOffice HTML: uma interface de painel pessoal gerada para uma designer freelancer em um estilo quente de linho, com uma barra lateral, saudação serifada e quatro cartões de métricas"></td>
<td width="50%"><img src="../assets/readme/html-report.webp" alt="GenOffice HTML: um relatório de dados sobre o mercado de veículos elétricos gerado em estilo jornal, com um cabeçalho serifado, um número de destaque de 17,3 milhões e uma linha de estatísticas"></td>
</tr>
<tr>
<td><b>Mockups de interface</b> — o modelo inicial "painel pessoal" transforma uma persona em um layout funcional: barra lateral, saudação, sparkline de horas faturáveis, cartões de faturas e utilização, tudo em HTML real que você pode entregar a um desenvolvedor.</td>
<td><b>Histórias com dados</b> — o modelo inicial "relatório de dados" constrói um jornal editorial: cabeçalho serifado, um número de destaque, uma linha de estatísticas separada por regra, gráficos SVG inline e uma nota metodológica.</td>
</tr>
</table>

### 6 · Markdown — um editor em blocos sobre `.md` puro, com Ask AI

<table>
<tr>
<td width="50%"><img src="../assets/readme/markdown-ai.webp" alt="GenOffice Markdown: um parágrafo selecionado mostra um popover Ask AI com uma instrução digitada e chips de sugestão como Aprimorar, Tornar mais conciso, Expandir e Corrigir gramática, além dos botões Enviar agora e Adicionar à fila"></td>
<td width="50%"><img src="../assets/readme/markdown-render.webp" alt="GenOffice Markdown renderizando um documento de notas de lançamento com uma tabela, um fluxograma Mermaid e uma lista de tarefas, com os prompts iniciais do painel de IA à esquerda"></td>
</tr>
<tr>
<td><b>Pergunte à IA sobre uma seleção</b> — selecione qualquer trecho e um chip <b>Ask AI</b> aparece: digite uma instrução ou escolha uma sugestão, envie agora, ou coloque em fila várias edições ancoradas e execute-as de uma vez. A mesma entrada existe em todos os aplicativos.</td>
<td><b>Renderizado, salvo como Markdown puro</b> — títulos, listas, tabelas, imagens, blocos de código e diagramas Mermaid em um editor de blocos Tiptap, gravados de volta como `.md` puro, com uma exportação <b>Markdown → Word</b> totalmente local.</td>
</tr>
</table>

## Por que o GenOffice

- **Código aberto**, Apache-2.0, desenvolvido abertamente no GitHub.
- **Seu para executar.** Apps nativos para macOS, Windows e Linux; os arquivos
  ficam no seu disco e cada edição, salvamento e conversão acontece na sua
  máquina.
- **Arquivos reais do Office.** `.docx`, `.xlsx` e `.pptx` nativos, preservando
  cada byte: as partes de um arquivo que você não tocou são copiadas
  exatamente como estavam.
- **Uma IA que edita o próprio documento.** Alterações rastreadas no Docs,
  fórmulas e gráficos ao vivo no Sheets, slides desenhados diretamente no
  canvas, cada interação da IA gera um snapshot que você pode desfazer.
- **Seu modelo, sua chave.** Entre com o Genspark ou traga uma chave para
  Claude, OpenAI, Gemini, DeepSeek e outros, com servidores locais e qualquer
  endpoint compatível com OpenAI incluídos.
- **PDF feito como deve ser.** Edite o texto dentro da página e converta PDF
  para Word, Excel ou PowerPoint no dispositivo, com OCR do sistema para
  digitalizações.
- **Markdown e HTML também**, com o mesmo painel de IA e exportação local para
  Word.
- **Gratuito**, para indivíduos e equipes.

## Backends de IA

**Faça login com o Genspark** e não há nada para configurar: as chamadas de
modelo passam pelo proxy do Genspark (famílias Claude, GPT e Gemini) e os
agentes ganham busca na web e de imagens, geração de imagens e análise de
imagem/áudio/vídeo.

**Ou traga sua própria chave.** Configurações → IA lista Claude, OpenAI,
Gemini, DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter
e OpenCode Zen/Go, além de um slot personalizado para qualquer endpoint
compatível com OpenAI (URL base + chave), incluindo servidores de modelo
locais. Busca e mídia têm seus próprios provedores por capacidade em **IA de
Mídia e Busca**: Serper ou Tavily para busca na web, e OpenAI, Gemini,
Doubao/Seedream, GLM, Grok, Qwen, MiniMax ou qualquer endpoint de imagens
compatível com OpenAI para geração de imagens e análise de imagem/vídeo.

A suíte completa vem com temas claro, escuro e do sistema. Os temas só
mudam o que aparece na tela: exportações, impressões e arquivos salvos
sempre mantêm as cores originais do documento.

<a id="download"></a>

## Download

| Plataforma                           | Requisitos                                            | Download                                                                                         |
| ------------------------------------ | ----------------------------------------------------- | ------------------------------------------------------------------------------------------------ |
| **macOS** — Apple Silicon (arm64)    | macOS 11+                                             | [`.dmg` mais recente (arm64)](https://github.com/genspark-ai/genoffice/releases/latest)          |
| **macOS** — Intel (x64)              | macOS 11+                                             | [`.dmg` mais recente (x64)](https://github.com/genspark-ai/genoffice/releases/latest)            |
| **Windows** (x64, maioria dos PCs)   | Windows 10+, Intel/AMD                                | [Instalador `-x64.exe` mais recente](https://github.com/genspark-ai/genoffice/releases/latest)   |
| **Windows** no Arm (ARM64)           | Windows 11 no Arm (Snapdragon X e similares)          | [Instalador `-arm64.exe` mais recente](https://github.com/genspark-ai/genoffice/releases/latest) |
| **Linux** — Debian / Ubuntu          | x86_64, glibc 2.34+ (Ubuntu 22.04 ou mais recente)    | [`.deb` mais recente](https://github.com/genspark-ai/genoffice/releases/latest)                  |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | [`.rpm` mais recente](https://github.com/genspark-ai/genoffice/releases/latest)                  |
| **Linux** — outras distribuições     | x86_64, glibc 2.34+, FUSE 2                           | [`.AppImage` mais recente](https://github.com/genspark-ai/genoffice/releases/latest)             |

Todas as builds vêm da `main`; os instaladores de macOS e Windows são
assinados. Versões mais antigas estão na página de
[Releases](https://github.com/genspark-ai/genoffice/releases).

<details>
<summary><b>Instalando no Linux</b></summary>

O deb instala com o apt — ele traz as dependências e adiciona o GenOffice
ao menu de aplicativos:

```bash
sudo apt install ./genoffice_<version>_amd64.deb
```

No Fedora / família RHEL / openSUSE, instale o rpm em vez disso:

```bash
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

O AppImage roda no lugar: instale o runtime FUSE 2
(`sudo apt install libfuse2`; no Ubuntu 24.04 o pacote é `libfuse2t64`),
torne o arquivo executável e então execute-o:

```bash
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

</details>

## Como funciona

Sete aplicativos Electron — Docs, Sheets, Slides, PDF, Markdown, HTML e o
shell com abas — compartilham uma camada de mecanismo de pacotes TypeScript
puro, além de um sidecar em Rust para `.xlsx`. O arquivo original é sempre a
fonte da verdade: as edições são aplicadas como patches estreitos, e tudo o
que o editor não tocou sobrevive intacto à ida e volta.

```
open docx ─► archive original by hash (never touched)
          ─► parse word/document.xml into a block tree, each block anchored to its original XML
          ─► Tiptap editor (manual + AI editing, dirty tracking)
save      ─► dirty blocks → OOXML fragments (referencing existing styles only)
          ─► splice into the original document.xml; untouched blocks keep their bytes
          ─► repack the zip; every other entry is copied byte-for-byte
```

O tour pacote a pacote (mecanismos de docx/pptx, `pdf2docx`, `html2docx`, o
núcleo do agente e os provedores) está em
[CONTRIBUTING.md](../../CONTRIBUTING.md#engine-packages).

## Desenvolvimento

```bash
npm install
npm run fixtures     # generate test .docx fixtures
npm test             # engine + app unit tests (docs/sheets/slides need no display)
npm run typecheck    # tsc --noEmit across every workspace
npm run dev          # all six editors + shell against Vite dev servers
npm run dev:docs     # a single app (same pattern works per workspace)
npm run dist:mac     # package macOS dmg (regenerates third-party notices)
npm run dist:win     # package Windows nsis installer
npm run dist:linux   # package Linux AppImage + deb + rpm
```

O aplicativo sheets também precisa de um toolchain Rust para seu sidecar
xlsx (`cargo` no PATH); `npm run build -w @genoffice/sheets` compila isso
automaticamente. Veja [CONTRIBUTING.md](../../CONTRIBUTING.md) para as
verificações que toda alteração precisa passar e como os pull requests são
integrados.

## Comunidade

O GenOffice está em desenvolvimento ativo e o seu feedback ajuda a moldá-lo.

- **Relate um bug ou peça um recurso** em
  [GitHub Issues](https://github.com/genspark-ai/genoffice/issues).
- **Participe do grupo de chat do GenOffice** no
  [GenTeam](https://genoffice.ai/join) para falar com a equipe e outros usuários.
- **Dê uma estrela no repositório** se o GenOffice for útil para você — é a
  melhor forma de apoiar o projeto.

## Perguntas frequentes

<details>
<summary><b>O GenOffice é gratuito?</b></summary>

Sim. O GenOffice é gratuito e de código aberto sob a licença Apache-2.0 —
sem teste gratuito, sem nível pago para os próprios aplicativos.

</details>

<details>
<summary><b>O GenOffice pode abrir arquivos do Microsoft Word, Excel e PowerPoint?</b></summary>

Sim. O GenOffice abre e salva arquivos nativos `.docx`, `.xlsx` e `.pptx`.
O salvamento preserva cada byte: as partes do arquivo que você não tocou são
gravadas de volta byte a byte, então os documentos continuam funcionando no
Microsoft Office.

</details>

<details>
<summary><b>O GenOffice funciona offline?</b></summary>

A edição de documentos é totalmente local — os arquivos nunca saem da sua
máquina para serem abertos, editados, salvos ou convertidos. Os recursos de
IA (agentes, busca, ferramentas de imagem) precisam de conexão de rede, com
login no Genspark ou sua própria chave de API de modelo.

</details>

<details>
<summary><b>O GenOffice pode editar arquivos PDF?</b></summary>

Sim — edição real de texto e imagem em PDF, que reescreve o content stream
da página preservando as fontes originais, em vez de anotações de
encobrimento.

</details>

<details>
<summary><b>O GenOffice pode converter PDF para Word, Excel ou PowerPoint?</b></summary>

Sim — totalmente no dispositivo: extração de caracteres via PDFium mais
análise de layout baseada em geometria, sem serviço na nuvem, sem upload.
Páginas escaneadas também são cobertas: no macOS e no Windows o OCR do
sistema as lê, então elas são convertidas em texto editável em vez de uma
imagem da página.

</details>

<details>
<summary><b>Posso usar meu próprio modelo de IA ou chave de API?</b></summary>

Sim. Além do login sem chave pelo Genspark, o GenOffice permite trazer sua
própria chave para Claude, OpenAI, Gemini, DeepSeek, Kimi, GLM, Qwen,
Doubao, MiniMax, Grok, Mistral, OpenRouter e OpenCode Zen/Go, além de
qualquer endpoint compatível com OpenAI — incluindo servidores de modelo
locais. Busca, geração de imagens e análise de imagem/vídeo usam suas
próprias chaves em Configurações → IA de Mídia e Busca.

</details>

<details>
<summary><b>O GenOffice pode converter HTML para Word?</b></summary>

Sim — Exportar como Word no aplicativo HTML produz um `.docx` nativo e
editável, totalmente no dispositivo. A página é renderizada no Chromium
integrado e reduzida a estruturas reais do Word: títulos, parágrafos,
listas, tabelas, cartões, linhas de KPIs, campos de formulário e fundos de
página; apenas elementos visuais sem equivalente no Word (gráficos, ícones,
caixas decoradas) são incorporados como imagens.

</details>

<details>
<summary><b>O GenOffice coleta algum dado?</b></summary>

As builds empacotadas oficiais enviam análises de uso limitadas por padrão,
e você pode desativar esse envio a qualquer momento em Configurações →
Geral. As análises nunca enviam conteúdo de documentos, nomes de arquivos,
caminhos de arquivos, identidade da conta ou endereços de e-mail. Veja
[GenOffice Privacy](../../PRIVACY.md) para a lista completa de eventos e
divulgações de dados.

</details>

## Segurança

Veja [SECURITY.md](../../SECURITY.md) para a postura de segurança do
processo (sandboxing do renderer, validação de IPC, controle de links
externos) e os modelos de ameaça para conteúdo gerado por IA.

## Agradecimentos

O GenOffice não seria possível sem estes projetos de código aberto:

- [Electron](https://www.electronjs.org/) — o runtime de desktop de todos os aplicativos.
- [Univer](https://github.com/dream-num/univer) (Apache-2.0) — o núcleo de
  interface de planilhas que o Sheets estende.
- [PDFium](https://pdfium.googlesource.com/pdfium/) (BSD-3-Clause,
  empacotado via [@embedpdf/pdfium](https://github.com/embedpdf/embed-pdf-viewer))
  — o mecanismo de content stream por trás da edição real de texto e imagem
  em PDF.
- [pdf.js](https://github.com/mozilla/pdf.js) (Apache-2.0) e
  [pdf-lib](https://github.com/Hopding/pdf-lib) (MIT) — renderização de PDF
  e montagem de documentos.
- [Tiptap](https://tiptap.dev/) / [ProseMirror](https://prosemirror.net/) —
  os editores em blocos do Docs e do Markdown.
- [CodeMirror](https://codemirror.net/) (MIT) — o editor de código-fonte do HTML.
- [Konva](https://konvajs.org/) — renderização em canvas para os gráficos do
  Slides e do Sheets.
- [HarfBuzz](https://github.com/harfbuzz/harfbuzz) (wasm) — métricas de
  text-shaping para escritas complexas.
- [calamine](https://github.com/tafia/calamine) e
  [IronCalc](https://github.com/ironcalc/IronCalc) — as camadas de leitura e
  cálculo do sidecar xlsx em Rust.
- [libeot](https://github.com/umanwizard/libeot) (MPL-2.0) — o decodificador
  MicroType Express para fontes do PowerPoint incorporadas, portado para
  TypeScript.
- [React](https://react.dev/) (MIT) — a camada de interface de todos os
  aplicativos.
- [Mermaid](https://mermaid.js.org/) (MIT) e [KaTeX](https://katex.org/)
  (MIT) — diagramas e fórmulas matemáticas no Markdown e no Docs.
- [opentype.js](https://opentype.js.org/) (MIT) — análise de fontes para
  métricas e busca de glifos.
- [JSZip](https://stuk.github.io/jszip/) (MIT) e
  [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser)
  (MIT) — as camadas de contêiner OOXML e XML.
- [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons)
  (MIT) — o conjunto de ícones usado nas faixas de opções.
- [electron-updater](https://www.electron.build/) (MIT) — atualizações
  dentro do próprio aplicativo.
- Fontes Liberation, Carlito, Caladea e Noto CJK (OFL/Apache-2.0) — fontes
  de documento incluídas.

`npm run notices` regenera o resumo de licenças de terceiros incluído
(`tools/gen-third-party-notices.mjs`); todas as dependências de runtime são
MIT/Apache-2.0/BSD-3-Clause/OFL.

## Licença

O GenOffice é licenciado sob a [Apache License 2.0](../../LICENSE), com uma
exceção: o diretório `ee/` é reservado para futuros módulos empresariais e
é coberto pela [GenOffice Enterprise License](../../ee/LICENSE).

Os nomes e logotipos GenOffice e Genspark são marcas registradas da
Mainfunc, Inc. A licença Apache-2.0 não concede permissão para usá-los (veja
a seção 6); forks devem usar sua própria marca.
