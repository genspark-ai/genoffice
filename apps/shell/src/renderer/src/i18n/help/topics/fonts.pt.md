# Tipos de letra: tipos do sistema e famílias transferíveis

## As listas de tipos de letra

Os controlos de tipo de letra ficam no separador Página Inicial de cada editor (a seguir: o grupo de tipos de letra no Slides):

![O grupo de tipos de letra no friso do Slides](img/fonts.png)

O seletor de tipos de letra de cada editor junta: tipos de letra instalados localmente + candidatos habituais da plataforma (as famílias comuns do Windows e do macOS, ocidentais e CJK, incluindo os nomes localizados que os tipos CJK indicam nos seus próprios sistemas). Quando o sistema consegue enumerar os tipos locais, a lista é agrupada pelo que existe de facto; caso contrário, degrada para a lista completa de candidatos.

## Tipos de letra transferíveis (Slides)

- O seletor de tipos de letra do Slides abre um **catálogo de tipos de letra**: uma seleção curada de OFL que abrange muitas escritas; e cada família traz os pesos normal e negrito.
- Escolher um transferes e instala-o a partir da CDN (com soma de verificação fixada) na loja de tipos de letra do aplicativo — fica disponível para todos os documentos a seguir, sem instalação ao nível do sistema.
- As instalações ficam num diretório privado do aplicativo e desaparecem com ele.

## Tipografia CJK

- Os tipos CJK no Docs recebem a compressão de pontuação e as regras de quebra de linha kinsoku alinhadas com o Word (consulte o capítulo do Docs).
- As famílias CJK do catálogo cobrem Chinês simplificado e tradicional, japonês e coreano, em serifas e sem serifas.

## Gerir tipos de letra locais

- A entrada install-local-font-files carrega ficheiros .ttf/.otf do disco para a loja de tipos de letra do aplicativo.
- A loja arquiva por família; vários pesos da mesma família podem coexistir.
