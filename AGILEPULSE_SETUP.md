# AgilePulse — configuração

## O que está incluído
- Dashboard de Throughput com uma única visualização de KPI.
- Dados demonstrativos explicitamente identificados como fictícios.
- Importação de Excel (.xlsx/.xls) e CSV, seleção de aba, prévia e mapeamento manual corrigível.
- Contagem por sprint em ordem numérica, deduplicação por identificador, status de conclusão configuráveis e alertas básicos de qualidade.
- Análise estatística local com média, mediana, extremos, variações e limitações.
- Chat embutido com respostas locais determinísticas e endpoint opcional para IA generativa.

## Configurar Claude (IA generativa)

O endpoint `POST /api/ai` usa diretamente a API oficial da Anthropic. Configure os segredos no ambiente do servidor/Lovable, nunca no código do navegador:

- `ANTHROPIC_API_KEY` — obrigatório.
- `ANTHROPIC_MODEL` — opcional; padrão `claude-sonnet-4-20250514`.

Sem `ANTHROPIC_API_KEY`, o dashboard mantém a análise estatística e as respostas locais de fallback, mas a análise generativa e o chat não serão executados pelo Claude.

Cada gráfico tem sua própria área de análise e chat logo abaixo dele. O conteúdo enviado ao modelo é limitado e contextualizado para a métrica selecionada.

## Limitações atuais
- A planilha é processada na sessão do navegador e não é persistida em banco compartilhado.
- Para reduzir a exposição e o consumo, o endpoint envia no máximo 300 linhas e limita o tamanho de valores.
- A análise estatística local não explica causalidade. Para investigar causas, colete contexto adicional, como bloqueios, capacidade e alterações de escopo.
- A validação de build e testes precisa ser executada no ambiente do projeto/Lovable após sincronização e instalação das dependências.
