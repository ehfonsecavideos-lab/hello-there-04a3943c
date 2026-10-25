# AgilePulse — configuração

## O que está incluído
- Dashboard de Throughput com uma única visualização de KPI.
- Dados demonstrativos explicitamente identificados como fictícios.
- Importação de Excel (.xlsx/.xls) e CSV, seleção de aba, prévia e mapeamento manual corrigível.
- Contagem por sprint em ordem numérica, deduplicação por identificador, status de conclusão configuráveis e alertas básicos de qualidade.
- Análise estatística local com média, mediana, extremos, variações e limitações.
- Chat embutido com respostas locais determinísticas e endpoint opcional para IA generativa.

## Configurar IA generativa
O endpoint `POST /api/ai` usa a API compatível com OpenAI. Configure os segredos no ambiente do servidor (não no código do navegador):

- `OPENAI_API_KEY` — obrigatório.
- `OPENAI_MODEL` — opcional; padrão `gpt-4o-mini`.
- `OPENAI_BASE_URL` — opcional; padrão `https://api.openai.com/v1`.

Sem `OPENAI_API_KEY`, o endpoint retorna `503 AI_NOT_CONFIGURED`. A interface apresenta cálculos locais e não afirma que a IA generativa está ativa.

## Limitações atuais
- A planilha é processada na sessão do navegador e não é persistida em banco compartilhado.
- Para reduzir a exposição e o consumo, o endpoint envia no máximo 300 linhas e limita o tamanho de valores.
- A análise estatística local não explica causalidade. Para investigar causas, colete contexto adicional, como bloqueios, capacidade e alterações de escopo.
- A validação de build e testes precisa ser executada no ambiente do projeto/Lovable após sincronização e instalação das dependências.
