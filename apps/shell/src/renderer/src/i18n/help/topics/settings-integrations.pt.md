# Configurações, idioma, tema e integrações MCP

## Abrir as configurações

A linha da conta no canto inferior esquerdo do ecrã Início abre o painel de configurações (mostra Entrar quando não tem sessão iniciada). O painel tem seis secções: Conta, Modelo de IA, Mídia e busca de IA, Geral, Integrações e Sobre.

![Configurações ▸ Geral, onde vivem o idioma, o tema, a guardado automático e o interruptor das estatísticas de utilização](img/settings-general.png)

A configuração dos modelos tem o seu próprio artigo; em **Mídia e busca de IA** liga, por cada fornecedor, a geração de imagens, a análise de imagens, a análise de vídeo, a busca na web e a pesquisa de arquivos locais.

## Idioma

- As configurações oferecem **21 idiomas de interface**: inglês, chinês simplificado, japonês, coreano, francês, alemão, espanhol, tailandês, indonésio, russo, árabe, português, italiano, polaco, tcheco, neerlandês, malaio, hebraico, híndi, chinês tradicional e vietnamita.
- A troca aplica-se de imediato e persiste; e a barra de menus do sistema é reconstruída com o novo idioma.

## Tema

Claro / Escuro / Seguir o Sistema. O modo Seguir o Sistema acompanha o aspeto do sistema operativo, e os editores mudam de aspeto em sincronia sem cintilação.

## Geral

- **Enviar estatísticas de uso anônimas** — ativado por predefinição. Usa o Google Analytics 4 e envia o seu IP público e metadados de transporte; o conteúdo dos documentos e os nomes de ficheiros nunca são recolhidos, e cada evento transporta apenas um tipo como «aberto um .docx». Pode desativá-lo aqui a qualquer momento.
- **Posição da barra lateral de IA** (esquerda ou direita), **Tamanho do texto do painel de IA** e **Verificação ortográfica no chat de IA**.
- **Abrir o painel de IA em novos documentos** — desativado, um novo documento começa com o painel recolhido, a um clique de distância.
- **Salvar automaticamente todos os documentos** ativa a gravação automática por predefinição em todos os editores; continua a poder desativá-la para uma janela.
- **Local de salvamento** com o botão **Alterar** e **App padrão para documentos do Office** para reservar .docx / .xlsx / .pptx ao GenOffice.

## Mídia e busca de IA

Não são interruptores — cada capacidade escolhe o fornecedor que a serve, e a chave e o URL base de um fornecedor são introduzidos uma vez e partilhados:

- **Busca na web**, **Geração de imagens**, **Análise de imagens** e **Análise de vídeo**, cada uma com um fornecedor, um modelo, uma chave e um URL base.
- **Pesquisa de arquivos locais** é executada nesta máquina. Por baixo dela está a **Reordenação com Jev**, **desativada por predefinição**. Ative-a e os excertos dos 20 melhores resultados locais — até 1200 caracteres de cada documento, mais os nomes dos ficheiros e das pastas — são enviados para o modelo Jev da TypeSafe e reordenados por relevância. Com ela desativada, nada sai do dispositivo.

## Sobre

- **Versão**, a ligação do projeto ao GitHub e um botão **Dar uma estrela no GitHub**.
- **Canal de atualização**: Estável ou Beta. Alterá-lo tem efeito imediato e verifica se existe uma atualização; não reverte uma instalação Beta para Estável.

## Associações de aplicações predefinidas

As Configurações podem registar o GenOffice como a aplicação que abre ficheiros .docx / .xlsx / .pptx / .pdf e semelhantes (registo de aplicação predefinida ao nível do sistema; confirme quando for solicitado).

## Avisos de software de terceiros e atualizações

- Ajuda ▸ Avisos de software de terceiros: o inventário completo das licenças de código aberto que acompanha a aplicação.
- Ajuda ▸ Procurar atualizações…: inicia uma verificação manual; e se houver uma versão mais recente, o GenOffice propõe a instalação.

## Entrar na Genspark

- O acesso (nas Configurações ou na lista de projetos na nuvem) usa um fluxo de **código de dispositivo**: o GenOffice mostra um código e abre o início de sessão no navegador; e a partir daí o processo prossegue sozinho.
- O início de sessão serve apenas para: a lista de projetos na nuvem e os modelos alojados da Genspark. Sem ele, todas as funcionalidades locais e os modelos personalizados continuam a funcionar.
- Terminar sessão é um clique nas Configurações.

## Integração MCP (para utilizadores avançados / clientes de IA)

**Integrações** é o painel que liga o GenOffice a um agente de programação, e tem o seu próprio artigo: Ligar um agente de programação. A versão curta — escolha um caminho (a linha de comandos, ou MCP), siga essa secção e depois inicie uma conversa nova e pergunte.

![Configurações ▸ Integrações: os três passos, depois as linhas de competências e as opções de MCP](img/settings-integrations.png)

Em **Servidor HTTP local**, a aplicação também pode arrancar o servidor por si — um interruptor de ativação e uma porta —, enquanto **Avançado** acrescenta o URL de verificação de integridade e o arquivo de registro, em vez de o deixar ao assistente. Escuta apenas em localhost.

## Folha de referência da linha de comandos

| Comando            | O que faz                    |
| ------------------ | ---------------------------- |
| `genoffice <file>` | abrir um ficheiro            |
| `genoffice mcp`    | iniciar o servidor MCP local |
| `genoffice --help` | todos os comandos e opções   |
