# AgilePulse — configuração

## O que está incluído
- Dashboard de Throughput com uma única visualização de KPI.
- Dados demonstrativos explicitamente identificados como fictícios.
- Importação de Excel (.xlsx/.xls) e CSV, seleção de aba, prévia e mapeamento manual corrigível.
- Contagem por sprint em ordem numérica, deduplicação por identificador, status de conclusão configuráveis e alertas básicos de qualidade.
- Análise estatística local com média, mediana, extremos, variações e limitações.
- Chat embutido com respostas locais determinísticas e endpoint opcional para IA generativa.

## Configurar a IA nativa do Lovable

O endpoint `POST /api/ai` usa o gateway de IA do Lovable e, por padrão, o modelo `google/gemini-3.8-flash`, indicado pelo Lovable para experiências de chat e assistentes. O código não exige chave da Anthropic, OpenAI ou outro provedor externo.

- Ative o Lovable AI para o projeto e use o ambiente/backend do Lovable que disponibiliza `LOVABLE_API_KEY`.
- `LOVABLE_AI_MODEL` é opcional e permite selecionar outro modelo compatível com o gateway do Lovable.
- Não exponha chaves no navegador nem as inclua no repositório.
- Sem acesso ao gateway nativo, o dashboard mantém a análise estatística e as respostas locais de contingência.

Cada gráfico tem sua própria área de análise e chat logo abaixo dele. As chamadas enviam a métrica selecionada e um conjunto limitado de dados para evitar respostas genéricas ou valores inventados.

## Limitações atuais
- A planilha é processada na sessão do navegador e não é persistida em banco compartilhado.
- Para reduzir a exposição e o consumo, o endpoint envia no máximo 300 linhas e limita o tamanho de valores.
- A análise estatística local não explica causalidade. Para investigar causas, colete contexto adicional, como bloqueios, capacidade e alterações de escopo.
- A validação de build e testes precisa ser executada no ambiente do projeto/Lovable após sincronização e instalação das dependências.
