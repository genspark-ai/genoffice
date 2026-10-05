# Modelos de IA e configurações

## Provedores e modelos

Os modelos e as chaves configuram-se nas Configurações (o botão da engrenagem no ecrã Início):

![A janela de configurações](img/settings-integrations.png)

- **Genspark alojado**: entre (com o fluxo de código de dispositivo) e use — sem configuração nenhuma.
- **Endpoints personalizados (BYOK)**: em Configurações ▸ Modelo de IA indicam-se a URL base e a chave de API de cada protocolo — o compatível com OpenAI, Anthropic, Gemini, DeepSeek, DashScope (qwen) e outros. As chaves vivem apenas nos cabeçalhos dos pedidos — nunca em disco, nos registos ou no ambiente dos processos filhos.
- Pode escolher um modelo diferente para cada capacidade: conversação/geração, geração de imagens, análise de imagens.
- **Testar conexão**: confirma que o endpoint está acessível e que o modelo está visível, antes de guardar.
- A URL base pode trazer um caminho e uma cadeia de consulta (ao estilo de gateway); e o caminho do endpoint é acrescentado corretamente.

## Integração com a CLI (estilo Codex)

- As Configurações aceitam o caminho de um programa CLI local (diretórios de casa não ASCII e o prefixo ~ funcionam; o ~ é expandido automaticamente); o Detectar automaticamente sonda os modelos que esse CLI disponibiliza.
- A validação verifica apenas a existência — não há restrições de conjunto de caracteres.

## Quando as alterações passam a vigorar

- As alterações de modelo e de endpoint aplicam-se de imediato; uma conversa em curso mantém a configuração antiga até ao turno seguinte.
