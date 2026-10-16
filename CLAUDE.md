## Hierarquia de prioridade

Instrução direta do usuário vence tudo. Depois, em caso de conflito:

1. *superpowers (processo) — sempre primeiro.*
 - Antes de qualquer feature nova ou mudança de comportamento: brainstorming, depois writing-plans.
 - Bug ou falha de teste: systematic-debugging antes de propor correção.
 - Implementação: test-driven-development.
 - Antes de dizer que terminou: verification-before-completion.
 2. *ponytail (código) — vale para lógica, backend, scripts, main.py e dados.*
 - Antes de escrever código, subir a escada: precisa existir? já existe no codebase, stdlib ou dependência instalada?
 - Nunca cortar: validação em fronteiras de confiança, tratamento de perda de dados, segurança, acessibilidade.
 - Ponytail também se aplica à quantidade de código de UI (sem componentes, wrappers ou abstrações desnecessárias), mas não decide o visual.
 3. *apple-design (visual) — só para camada visual de frontend*
 - apple-design decidem tipografia, cor, espaçamento e movimento, dentro do que o manual de marca permite. Não decidem estrutura de código nem arquitetura.