# Fontes: fontes do sistema e famílias para baixar

## As listas de fontes

Os controles de fonte ficam na guia Página Inicial de cada editor (a seguir: o grupo de fontes no Slides):

![O grupo de fontes na faixa de opções do Slides](img/fonts.png)

O seletor de fontes de cada editor junta: fontes instaladas localmente + candidatos habituais da plataforma (as famílias comuns do Windows e do macOS, ocidentais e CJK, incluindo os nomes localizados que os tipos CJK indicam nos seus próprios sistemas). Quando o sistema consegue enumerar os tipos locais, a lista é agrupada conforme as fontes disponíveis; caso contrário, degrada para a lista completa de candidatos.

## Fontes para baixar (Slides)

- O seletor de fontes do Slides abre um **catálogo de fontes**: uma seleção curada de OFL que abrange muitas escritas; e cada família traz os pesos normal e negrito.
- Ao escolher uma fonte, o aplicativo faz o download e a instala a partir da CDN (com soma de verificação fixada) na loja de fontes do aplicativo — fica disponível para todos os documentos a seguir, sem instalação no sistema.
- As instalações ficam em um diretório privado do aplicativo e desaparecem com ele.

## Tipografia CJK

- Os tipos CJK no Docs recebem a compressão de pontuação e as regras de quebra de linha kinsoku alinhadas com o Word (consulte o capítulo do Docs).
- As famílias CJK do catálogo cobrem Chinês simplificado e tradicional, japonês e coreano, em serifas e sem serifas.

## Gerenciar fontes locais

- A entrada install-local-font-files carrega arquivos .ttf/.otf do disco para a loja de fontes do aplicativo.
- A loja arquiva por família; vários pesos da mesma família podem coexistir.
