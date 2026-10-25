import { useEffect, useMemo, useRef, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { Bar, BarChart, CartesianGrid, Cell, ComposedChart, Line, ReferenceArea, ReferenceLine, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
import * as XLSX from "xlsx";
import { Activity, AlertCircle, Check, ChevronDown, Database, FileSpreadsheet, FileUp, MessageCircle, Send, Sparkles, X } from "lucide-react";

export const Route = createFileRoute("/")({ component: AgilePulse });
type Row = Record<string, unknown>;
type MapCols = { sprint: string; item: string; status: string; squad: string; created: string; started: string; completed: string; storyPoints: string };
type Msg = { role: "user" | "assistant"; content: string };
type MetricKey = "throughput" | "leadTime" | "cycleTime" | "velocity";
const metricNames: Record<MetricKey, string> = { throughput: "Throughput", leadTime: "Lead Time", cycleTime: "Cycle Time", velocity: "Velocity" };
const BASE_COLUMNS = ["Issue key","Resumo","Sprint","Sprint nº","Tipo","Papel","Story Points","Prioridade","Épico / EAP","Criado","Concluído","Lead time (dias corridos)","Cycle time (dias corridos)","In development","Waiting for QA","In QA review","In code review","Waiting for review","Dias bloqueado","Reaberturas"];
const OPTIONAL_COLUMNS: Record<string,string[]> = {
  "Calendario_Eventos":["Sprint","Início","Fim","Evento / contexto","Pessoas no time","Capacidade (h)","Fase"],
  "Defeitos":["ID","Sprint de detecção","Data de detecção","Item afetado (Issue key)","Módulo","Origem","Severidade","Horas de correção","Causou falha de deploy?","Horas para restaurar"],
  "Retrabalho":["ID","Item reaberto (Issue key)","Sprint do item","Nº da reabertura","Causa","Etapa em que foi detectada","Papel responsável","Horas de retrabalho"],
  "DORA_Sprint":["Sprint","Deploys na sprint"],
  "Backlog":["Issue key","Resumo","Módulo","Prioridade","Story Points","Depende de terceiros?"],
  "WIP_Aberto":[],
  "Premissas":["Parâmetro","Valor"],
  "Riscos_Gatilhos":[],
  "Stakeholders":[],
  "Dicionario_Metricas":[]
};
const IGNORED_SHEETS = new Set(["Indicadores_Sprint","Fluxo_Etapas","Forecast_Opcoes","Resumo_Executivo"]);
const OPTIONAL_SHEET_NAMES = Object.keys(OPTIONAL_COLUMNS);
const STORE_KEY = "agilepulse-workbook-v2";
const demo: Row[] = [
  { "Issue key": "AP-101", Summary: "Login", Sprint: "Sprint 1", "Issue Status": "Done", Squad: "Plataforma", "Created Date": "2026-01-05", "Development Start Date": "2026-01-08", "Resolution Date": "2026-01-10", "Story Points": 2 },
  { "Issue key": "AP-102", Summary: "Perfil", Sprint: "Sprint 1", "Issue Status": "Done", Squad: "Experiência", "Created Date": "2026-01-05", "Development Start Date": "2026-01-08", "Resolution Date": "2026-01-10", "Story Points": 3 },
  { "Issue key": "AP-103", Summary: "Busca", Sprint: "Sprint 1", "Issue Status": "In Progress", Squad: "Plataforma", "Created Date": "2026-01-05", "Development Start Date": "2026-01-08", "Resolution Date": "2026-01-10", "Story Points": 4 },
  { "Issue key": "AP-104", Summary: "Cadastro", Sprint: "Sprint 2", "Issue Status": "Done", Squad: "Experiência", "Created Date": "2026-01-08", "Development Start Date": "2026-01-11", "Resolution Date": "2026-01-13", "Story Points": 5 },
  { "Issue key": "AP-105", Summary: "Notificações", Sprint: "Sprint 2", "Issue Status": "Done", Squad: "Plataforma", "Created Date": "2026-01-08", "Development Start Date": "2026-01-11", "Resolution Date": "2026-01-13", "Story Points": 1 },
  { "Issue key": "AP-106", Summary: "Relatórios", Sprint: "Sprint 2", "Issue Status": "Done", Squad: "Dados", "Created Date": "2026-01-08", "Development Start Date": "2026-01-11", "Resolution Date": "2026-01-13", "Story Points": 2 },
  { "Issue key": "AP-107", Summary: "Permissões", Sprint: "Sprint 3", "Issue Status": "Done", Squad: "Plataforma", "Created Date": "2026-01-11", "Development Start Date": "2026-01-14", "Resolution Date": "2026-01-16", "Story Points": 3 },
  { "Issue key": "AP-108", Summary: "Exportação", Sprint: "Sprint 3", "Issue Status": "In Review", Squad: "Dados", "Created Date": "2026-01-11", "Development Start Date": "2026-01-14", "Resolution Date": "2026-01-16", "Story Points": 4 },
  { "Issue key": "AP-109", Summary: "Acessibilidade", Sprint: "Sprint 3", "Issue Status": "Done", Squad: "Experiência", "Created Date": "2026-01-11", "Development Start Date": "2026-01-14", "Resolution Date": "2026-01-16", "Story Points": 5 },
  { "Issue key": "AP-110", Summary: "Auditoria", Sprint: "Sprint 4", "Issue Status": "Done", Squad: "Dados", "Created Date": "2026-01-14", "Development Start Date": "2026-01-17", "Resolution Date": "2026-01-19", "Story Points": 1 },
  { "Issue key": "AP-111", Summary: "Filtros", Sprint: "Sprint 4", "Issue Status": "Done", Squad: "Experiência", "Created Date": "2026-01-14", "Development Start Date": "2026-01-17", "Resolution Date": "2026-01-19", "Story Points": 2 },
  { "Issue key": "AP-112", Summary: "Painel", Sprint: "Sprint 4", "Issue Status": "Done", Squad: "Dados", "Created Date": "2026-01-14", "Development Start Date": "2026-01-17", "Resolution Date": "2026-01-19", "Story Points": 3 },
  { "Issue key": "AP-113", Summary: "API", Sprint: "Sprint 4", "Issue Status": "Done", Squad: "Plataforma", "Created Date": "2026-01-14", "Development Start Date": "2026-01-17", "Resolution Date": "2026-01-19", "Story Points": 4 },
  { "Issue key": "AP-114", Summary: "Melhorias", Sprint: "Sprint 5", "Issue Status": "Done", Squad: "Experiência", "Created Date": "2026-01-17", "Development Start Date": "2026-01-20", "Resolution Date": "2026-01-22", "Story Points": 5 },
  { "Issue key": "AP-115", Summary: "API externa", Sprint: "Sprint 5", "Issue Status": "Blocked", Squad: "Plataforma", "Created Date": "2026-01-17", "Development Start Date": "2026-01-20", "Resolution Date": "2026-01-22", "Story Points": 1 },
  { "Issue key": "AP-116", Summary: "Logs", Sprint: "Sprint 6", "Issue Status": "Done", Squad: "Dados", "Created Date": "2026-01-20", "Development Start Date": "2026-01-23", "Resolution Date": "2026-01-25", "Story Points": 2 },
  { "Issue key": "AP-117", Summary: "Monitoramento", Sprint: "Sprint 6", "Issue Status": "Done", Squad: "Plataforma", "Created Date": "2026-01-20", "Development Start Date": "2026-01-23", "Resolution Date": "2026-01-25", "Story Points": 3 },
  { "Issue key": "AP-118", Summary: "Testes", Sprint: "Sprint 6", "Issue Status": "Done", Squad: "Experiência", "Created Date": "2026-01-20", "Development Start Date": "2026-01-23", "Resolution Date": "2026-01-25", "Story Points": 4 },
  { "Issue key": "AP-119", Summary: "Documentação", Sprint: "Sprint 6", "Issue Status": "Done", Squad: "Dados", "Created Date": "2026-01-20", "Development Start Date": "2026-01-23", "Resolution Date": "2026-01-25", "Story Points": 5 },
];
const normalize = (v: unknown) => String(v ?? "").trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
const str = (v: unknown) => String(v ?? "").trim();
const completedDefaults = ["done", "concluído", "concluido", "resolvido", "resolvida", "closed", "complete", "completed", "feito", "finalizado", "testado", "aprovado para produção"];
function guessMap(rows: Row[]): MapCols {
  const h = [...new Set(rows.flatMap(row => Object.keys(row)))];
  const find = (rx: RegExp[]) => h.find(k => rx.some(p => p.test(normalize(k)))) || "";
  return {
    sprint: find([/^sprint$/, /sprint.*(name|nome|number|numero)/, /iteracao/]),
    item: find([/^issue key$/, /^issue id$/, /^key$/, /^id$/, /identificador/]),
    status: find([/^issue status$/, /^status$/, /situacao/, /fase.*status/]),
    squad: find([/squad/, /equipe/, /team/]),
    created: find([/created date/, /data de criacao/, /created at/, /creation date/, /data criacao/, /issue created/]),
    started: find([/cycle start/, /start date/, /data de inicio/, /inicio do trabalho/, /phase start date/, /development start/]),
    completed: find([/resolution date/, /completed date/, /data de conclusao/, /data conclusao/, /resolved date/, /done date/, /completion date/]),
    storyPoints: find([/^story points$/, /story point/, /pontos da historia/, /pontos de historia/, /estimativa.*ponto/])
  };
}
function sprintNum(s: string) { const m = s.match(/\d+/); return m ? Number(m[0]) : Number.MAX_SAFE_INTEGER; }
function done(v: unknown, statuses: string[]) { return statuses.some(s => normalize(s) === normalize(v)); }
function dateValue(value: unknown): Date | null {
  if(value instanceof Date && !Number.isNaN(value.getTime())) return new Date(value.getFullYear(),value.getMonth(),value.getDate());
  if(typeof value === "number" && Number.isFinite(value)){const p=XLSX.SSF.parse_date_code(value);if(p)return new Date(p.y,p.m-1,p.d);}
  const raw = str(value); if (!raw) return null;
  const br = raw.match(/^(\\d{1,2})[/. -](\\d{1,2})[/. -](\\d{2,4})$/);
  let d: Date;
  if (br) { const year = Number(br[3].length === 2 ? "20" + br[3] : br[3]); d = new Date(year, Number(br[2])-1, Number(br[1])); }
  else d = new Date(raw);
  return Number.isNaN(d.getTime()) ? null : d;
}
function elapsedDays(from: unknown, to: unknown): number | null {
  const a=dateValue(from), b=dateValue(to); if(!a||!b||b<a)return null;
  return Math.max(0,(b.getTime()-a.getTime())/86400000);
}
function calculate(rows: Row[], m: MapCols, statuses: string[]) {
  const valid = rows.filter(r => str(r[m.sprint]));
  const labels = [...new Set(valid.map(r => str(r[m.sprint])))].sort((a,b) => sprintNum(a)-sprintNum(b) || a.localeCompare(b, "pt-BR", { numeric: true }));
  const seen = new Set<string>(); let duplicates = 0;
  const counts = labels.map(sprint => {
    let throughput = 0, velocity = 0, leadTotal = 0, leadCount = 0, cycleTotal = 0, cycleCount = 0;
    valid.filter(r => str(r[m.sprint]) === sprint).forEach(r => {
      const completed = str(r["Concluído"]) ? true : done(r[m.status], statuses);
      if (!completed) return;
      const id = m.item ? str(r[m.item]) : "";
      if (id && seen.has(id)) { duplicates++; return; }
      if (id) seen.add(id);
      throughput++;
      if (m.storyPoints) { const points = Number(str(r[m.storyPoints]).replace(",", ".")); if (Number.isFinite(points) && points > 0) velocity += points; }
      const directLead = Number(str(r["Lead time (dias corridos)"]).replace(",", "."));
      const lead = str(r["Lead time (dias corridos)"]) && Number.isFinite(directLead) ? directLead : (m.created && m.completed ? elapsedDays(r[m.created],r[m.completed]) : null);
      if (lead!==null && lead>=0) {leadTotal+=lead;leadCount++;}
      const directCycle = Number(str(r["Cycle time (dias corridos)"]).replace(",", "."));
      const cycle = str(r["Cycle time (dias corridos)"]) && Number.isFinite(directCycle) ? directCycle : (m.started && m.completed ? elapsedDays(r[m.started],r[m.completed]) : null);
      if (cycle!==null && cycle>=0) {cycleTotal+=cycle;cycleCount++;}
    });
    return { sprint, throughput, velocity, leadTime: leadCount ? leadTotal/leadCount : null, cycleTime: cycleCount ? cycleTotal/cycleCount : null, leadCount, cycleCount };
  });
  const values = counts.map(c => c.throughput), sorted = [...values].sort((a,b)=>a-b);
  const total = values.reduce((a,b)=>a+b,0), mean = values.length ? total/values.length : 0;
  const median = sorted.length ? (sorted.length%2 ? sorted[(sorted.length-1)/2] : (sorted[sorted.length/2-1]+sorted[sorted.length/2])/2) : 0;
  return { counts, total, mean, median, max: counts.reduce((a,b)=>!a||b.throughput>a.throughput?b:a,null as null|typeof counts[number]), min: counts.reduce((a,b)=>!a||b.throughput<a.throughput?b:a,null as null|typeof counts[number]), noSprint: rows.filter(r=>!str(r[m.sprint])).length, noStatus: rows.filter(r=>!str(r[m.status])).length, noItem: m.item?rows.filter(r=>!str(r[m.item])).length:0, duplicates, labels: labels.length };
}
function localAnswer(q: string, stats: ReturnType<typeof calculate>, rows: Row[], m: MapCols, statuses: string[]) {
  const n = normalize(q);
  if (/squad|equipe|team/.test(n)) {
    if (!m.squad) return "Não encontrei coluna de squad/equipe na planilha. Inclua esse campo para comparar squads.";
    const totals = new Map<string, Set<string>>();
    rows.forEach(r=>{ if (!done(r[m.status], statuses)) return; const s=str(r[m.squad])||"(sem squad)"; if(!totals.has(s))totals.set(s,new Set()); totals.get(s)!.add(m.item?str(r[m.item]):str(r.Summary)); });
    return [...totals].map(([s,ids])=>[s,ids.size] as const).sort((a,b)=>b[1]-a[1]).map(([s,c])=>s+": "+c+" itens concluídos").join("\n") || "Não há dados de squad para comparar.";
  }
  if (/média|media|average/.test(n)) return "Média: "+stats.mean.toFixed(2)+" itens por sprint ("+stats.total+" itens ÷ "+stats.labels+" sprints).";
  if (/mediana/.test(n)) return "A mediana é "+stats.median.toFixed(1)+" itens concluídos por sprint.";
  if (/maior|maximo|máximo|mais itens/.test(n)) return stats.max ? stats.max.sprint+" teve o maior Throughput: "+stats.max.throughput+" itens." : "Não há sprints suficientes.";
  if (/menor|minimo|mínimo/.test(n)) return stats.min ? stats.min.sprint+" teve o menor Throughput: "+stats.min.throughput+" itens." : "Não há sprints suficientes.";
  const match = q.match(/sprint\s*(\d+)/i);
  if (match) { const s=stats.counts.find(x=>sprintNum(x.sprint)===Number(match[1])); if(!s)return "Não encontrei Sprint "+match[1]+" nos dados carregados."; const i=stats.counts.indexOf(s), p=stats.counts[i-1]; return s.sprint+": "+s.throughput+" itens concluídos."+ (p?" Variação contra "+p.sprint+": "+(s.throughput-p.throughput>0?"+":"")+(s.throughput-p.throughput)+".":" Não há sprint anterior para comparar."); }
  if (/queda|caiu|diminu/.test(n)) { const drops=stats.counts.slice(1).map((s,i)=>({p:stats.counts[i],s,d:s.throughput-stats.counts[i].throughput})).filter(x=>x.d<0); return drops.length ? "Quedas observadas:\n"+drops.map(x=>x.p.sprint+" → "+x.s.sprint+": "+x.p.throughput+" para "+x.s.throughput+" ("+x.d+").").join("\n")+"\nThroughput não identifica a causa. Colete dados de bloqueios, capacidade e mudanças de escopo." : "Não há quedas entre sprints consecutivas nos dados carregados."; }
  return "Posso calcular maior/menor Throughput, média, mediana, variações, uma sprint específica ou comparação por squad. Esta resposta é calculada localmente. Para respostas generativas, ative a IA nativa do Lovable neste projeto.";
}
function metricAnalysis(stats: ReturnType<typeof calculate>, key: MetricKey): string[] {
  const series = stats.counts.map(s => ({sprint:s.sprint, value:s[key]})).filter((x): x is {sprint:string;value:number} => typeof x.value === "number" && Number.isFinite(x.value));
  if (!series.length) return ["Não há valores válidos para "+metricNames[key]+". Confira o mapeamento das colunas e os dados importados.", "A análise só será calculada quando houver datas válidas para Lead Time/Cycle Time ou Story Points numéricos para Velocity.", "Não é possível identificar tendências sem valores válidos.", "Nenhum resultado foi inventado; importe ou mapeie os campos necessários."];
  const mean=series.reduce((sum,x)=>sum+x.value,0)/series.length;
  const max=series.reduce((a,b)=>b.value>a.value?b:a), min=series.reduce((a,b)=>b.value<a.value?b:a);
  const unit=key==="velocity"?"pontos":key==="throughput"?"itens":"dias";
  const changes=series.slice(1).map((x,i)=>({from:series[i],to:x,delta:x.value-series[i].value}));
  const trend=changes.length?changes.map(x=>x.to.sprint+": "+(x.delta>0?"+":"")+x.delta.toFixed(1)+" "+unit).join(" · "):"Há apenas um ponto de dados; não é possível comparar sprints.";
  return ["Média de "+metricNames[key]+": "+mean.toFixed(1)+" "+unit+" em "+series.length+" sprints com valores válidos.", "Maior valor: "+max.sprint+" ("+max.value.toFixed(1)+" "+unit+"). Menor valor: "+min.sprint+" ("+min.value.toFixed(1)+" "+unit+").", "Variação entre sprints consecutivas: "+trend, key==="leadTime"||key==="cycleTime"?"Tempos mais altos indicam maior duração do fluxo, mas não explicam sozinhos a causa. Investigue bloqueios, filas e retrabalho.":key==="velocity"?"Velocity é a soma de Story Points concluídos por sprint; não deve ser interpretada isoladamente como produtividade ou qualidade.":"Throughput é a contagem de itens concluídos; a métrica descreve volume, não tamanho ou complexidade."];
}
function metricLocalAnswer(q: string, stats: ReturnType<typeof calculate>, key: MetricKey, rows: Row[], m: MapCols, statuses: string[]) {
  if (key === "throughput") return localAnswer(q,stats,rows,m,statuses);
  const n=normalize(q), series=stats.counts.map(s=>({sprint:s.sprint,value:s[key]})).filter((x): x is {sprint:string;value:number} => typeof x.value==="number" && Number.isFinite(x.value));
  if (!series.length) return "Ainda não há valores válidos de "+metricNames[key]+". Confira as colunas necessárias na importação.";
  const unit=key==="velocity"?"pontos":"dias", mean=series.reduce((sum,x)=>sum+x.value,0)/series.length;
  const max=series.reduce((a,b)=>b.value>a.value?b:a), min=series.reduce((a,b)=>b.value<a.value?b:a);
  if (/m[eé]dia|media|average/.test(n)) return "Média de "+metricNames[key]+": "+mean.toFixed(2)+" "+unit+" por sprint, considerando "+series.length+" sprints com dados válidos.";
  if (/maior|maximo|máximo|mais alto|mais longa|maior valor/.test(n)) return max.sprint+" teve o maior "+metricNames[key]+": "+max.value.toFixed(2)+" "+unit+".";
  if (/menor|minimo|mínimo|mais baixo|mais curta|menor valor/.test(n)) return min.sprint+" teve o menor "+metricNames[key]+": "+min.value.toFixed(2)+" "+unit+".";
  if (/queda|caiu|diminu|varia|diferen/.test(n)) { const changes=series.slice(1).map((x,i)=>({from:series[i],to:x,d:x.value-series[i].value})); return changes.length?changes.map(x=>x.from.sprint+" → "+x.to.sprint+": "+x.from.value.toFixed(1)+" → "+x.to.value.toFixed(1)+" ("+(x.d>0?"+":"")+x.d.toFixed(1)+" "+unit+")").join("\n"):"Não há sprints suficientes para comparar."; }
  const match=q.match(/sprint\s*(\d+)/i);
  if(match){const s=series.find(x=>sprintNum(x.sprint)===Number(match[1]));return s?s.sprint+": "+s.value.toFixed(2)+" "+unit+".":"Não encontrei essa sprint com valores válidos para "+metricNames[key]+".";}
  return metricAnalysis(stats,key).join("\n");
}
function InlineAssistant({ metricKey, stats, rows, mapping, statuses, sourceLabel, calendarRows, premises }: { metricKey:MetricKey; stats:ReturnType<typeof calculate>; rows:Row[]; mapping:MapCols; statuses:string[]; sourceLabel:string; calendarRows:Row[]; premises:Row[] }) {
  const [analysisOpen,setAnalysisOpen]=useState(false), [chatOpen,setChatOpen]=useState(false);
  const [analysisText,setAnalysisText]=useState(""), [analysisLoading,setAnalysisLoading]=useState(false), [analysisNote,setAnalysisNote]=useState("");
  const [messages,setMessages]=useState<Msg[]>([{role:"assistant",content:"Olá! Faça perguntas sobre esta métrica. As respostas generativas usam a IA nativa do Lovable quando ela estiver habilitada."}]);
  const [question,setQuestion]=useState(""), [sending,setSending]=useState(false), [error,setError]=useState("");
  async function askAI(q:string, history:Msg[], task:"analysis"|"chat") {
    const res=await fetch("/api/ai",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({task,question:q,metric:metricKey,metricName:metricNames[metricKey],messages:history.slice(-8),data:rows.slice(0,500),mapping,completedStatuses:statuses,stats,metricSeries:stats.counts.map(s=>({sprint:s.sprint,value:s[metricKey]})),calendarEvents:calendarRows.slice(0,100),premises:premises.slice(0,100)})});
    const body=await res.json();
    if(!res.ok||!body.answer) throw new Error(body.error||"A IA não conseguiu responder.");
    return String(body.answer);
  }
  async function toggleAnalysis() {
    if(analysisOpen){setAnalysisOpen(false);return;}
    setAnalysisOpen(true);setAnalysisLoading(true);setAnalysisNote("");setAnalysisText("");
    const prompt="Faça uma análise inteligente da métrica "+metricNames[metricKey]+" usando exclusivamente os dados da página, eventos da aba Calendario_Eventos e metas/limites da aba Premissas. Ao explicar melhora ou piora, cite o evento/contexto registrado para a sprint relevante; se não houver evento, diga isso e não atribua causalidade. Organize em: insights, riscos e próximas ações. Cite valores reais; diferencie fatos de hipóteses; nunca invente números, eventos, metas ou causas.";
    try { setAnalysisText(await askAI(prompt,[{role:"user",content:prompt}],"analysis")); }
    catch { setAnalysisText(metricAnalysis(stats,metricKey).join("\n\n"));setAnalysisNote("Análise estatística local exibida porque a IA do Lovable não está configurada ou está indisponível. Ative o Lovable AI/Cloud para habilitar a análise generativa."); }
    finally { setAnalysisLoading(false); }
  }
  async function send() {
    const q=question.trim();if(!q||sending)return;setQuestion("");setError("");
    const next=[...messages,{role:"user" as const,content:q}];setMessages(next);setSending(true);
    try { const answer=await askAI(q,next,"chat");setMessages(m=>[...m,{role:"assistant",content:answer}]); }
    catch { setMessages(m=>[...m,{role:"assistant",content:metricLocalAnswer(q,stats,metricKey,rows,mapping,statuses)}]);setError("A IA não respondeu. A resposta acima foi calculada localmente; confira se o Lovable AI está habilitado."); }
    finally { setSending(false); }
  }
  return <div className="ap-inline-assistant">
    <div className="ap-chart-actions ap-metric-actions"><button className="ap-button ap-secondary" onClick={()=>void toggleAnalysis()}><Sparkles size={15}/>{analysisOpen?"Ocultar análise":"Análise inteligente"}</button><button className="ap-button ap-primary" onClick={()=>setChatOpen(v=>!v)}><MessageCircle size={15}/>{chatOpen?"Ocultar chat":"Perguntar sobre os dados"}</button></div>
    {analysisOpen&&<section className="ap-inline-panel"><div className="ap-inline-heading"><div><Sparkles size={16}/><strong>Análise inteligente · {metricNames[metricKey]}</strong></div><span>{analysisLoading?"A IA está analisando…":"Fonte: "+sourceLabel}</span></div>{analysisLoading?<div className="ap-typing">Analisando a série deste gráfico…</div>:<div className="ap-inline-answer">{analysisText.split("\n").map((line,i)=><p key={i}>{line||" "}</p>)}</div>}{analysisNote&&<p className="ap-note"><AlertCircle size={14}/>{analysisNote}</p>}</section>}
    {chatOpen&&<section className="ap-inline-panel ap-inline-chat"><div className="ap-inline-heading"><div><MessageCircle size={16}/><strong>Pergunte à IA · {metricNames[metricKey]}</strong></div><span>{sourceLabel} · {stats.labels} sprints</span></div><div className="ap-chat-messages">{messages.map((m,i)=><div key={i} className={"ap-message "+m.role}><div className="ap-message-avatar">{m.role==="assistant"?<Sparkles size={14}/>:"Você"}</div><div>{m.content.split("\n").map((line,j)=><p key={j}>{line||" "}</p>)}</div></div>)}{sending&&<div className="ap-typing">A IA está consultando os dados…</div>}{error&&<div className="ap-error"><AlertCircle size={15}/>{error}<button onClick={()=>setError("")}>Fechar</button></div>}</div><div className="ap-suggestions">{[`Qual sprint teve maior ${metricNames[metricKey]}?`,`Qual foi a média de ${metricNames[metricKey]}?`,`Houve alguma queda em ${metricNames[metricKey]}?`].map(q=><button key={q} onClick={()=>setQuestion(q)}>{q}</button>)}</div><form className="ap-chat-form" onSubmit={e=>{e.preventDefault();void send();}}><input value={question} onChange={e=>setQuestion(e.target.value)} placeholder={"Pergunte sobre "+metricNames[metricKey]+"…"} /><button type="submit" disabled={!question.trim()||sending}><Send size={17}/></button></form><div className="ap-chat-footnote">A IA usa os dados carregados desta métrica. As respostas locais são apenas fallback.</div></section>}
  </div>;
}
function excelCell(value: unknown, header: string): unknown {
  if(value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0,10);
  if(typeof value==="number" && /(data|criado|conclu|in[ií]cio|fim|detec[cç][aã]o|snapshot)/i.test(header)){
    const parsed=XLSX.SSF.parse_date_code(value);
    if(parsed)return new Date(Date.UTC(parsed.y,parsed.m-1,parsed.d)).toISOString().slice(0,10);
  }
  return value ?? "";
}
function parseSheetRows(ws: XLSX.WorkSheet, headerIndex: number, startIndex: number, endIndex?: number): Row[] {
  const matrix=XLSX.utils.sheet_to_json<unknown[]>(ws,{header:1,defval:"",raw:true,blankrows:true}) as unknown[][];
  const headers=(matrix[headerIndex]||[]).map(v=>str(v));
  const rows:Row[]=[];
  const stop=Math.min(matrix.length,endIndex===undefined?matrix.length:endIndex);
  for(let i=startIndex;i<stop;i++){
    const values=matrix[i]||[];
    const row:Row={};
    headers.forEach((header,j)=>{if(header)row[header]=excelCell(values[j],header);});
    if(Object.values(row).some(v=>str(v)!==""))rows.push(row);
  }
  return rows;
}

const SPRINT_OPTIONS = Array.from({length:20},(_,i)=>"Sprint "+(i+1));
function inclusiveP85(values:number[]):number|null {
  if(!values.length)return null;
  const sorted=[...values].sort((a,b)=>a-b);
  const pos=(sorted.length-1)*0.85, lo=Math.floor(pos), hi=Math.ceil(pos);
  return sorted[lo]+(sorted[hi]-sorted[lo])*(pos-lo);
}
function numeric(v:unknown):number|null { const n=Number(str(v).replace(",","."));return str(v)!==""&&Number.isFinite(n)?n:null; }
function isCompletedRow(r:Row):boolean { return !!str(r["Concluído"]) || (str(r["Issue Status"]) ? completedDefaults.some(s=>normalize(s)===normalize(r["Issue Status"])) : !!str(r["Resolution Date"])); }
function buildFlowSeries(rows:Row[], calendar:Row[]) {
  const eventBySprint=new Map(calendar.map(r=>[str(r["Sprint"]),r]));
  const allLead=rows.filter(isCompletedRow).map(r=>numeric(r["Lead time (dias corridos)"])??elapsedDays(r["Criado"]??r["Created Date"],r["Concluído"]??r["Resolution Date"])).filter((v):v is number=>v!==null&&v>=0);
  const overallP85=inclusiveP85(allLead);
  const overallLeadMean=allLead.length?allLead.reduce((sum,v)=>sum+v,0)/allLead.length:null;
  return SPRINT_OPTIONS.map((sprint,i)=>{
    const items=rows.filter(r=>str(r["Sprint"])===sprint);
    const completed=items.filter(isCompletedRow);
    const leadValues=completed.map(r=>numeric(r["Lead time (dias corridos)"])??elapsedDays(r["Criado"]??r["Created Date"],r["Concluído"]??r["Resolution Date"])).filter((v):v is number=>v!==null&&v>=0);
    const cycleValues=completed.map(r=>numeric(r["Cycle time (dias corridos)"])??elapsedDays(r["In development"]??r["Development Start Date"],r["Concluído"]??r["Resolution Date"])).filter((v):v is number=>v!==null&&v>=0);
    const noReopen=completed.filter(r=>(numeric(r["Reaberturas"])??0)===0).length;
    const reopened=completed.filter(r=>(numeric(r["Reaberturas"])??0)>0).length;
    const lead=leadValues.length?leadValues.reduce((a,b)=>a+b,0)/leadValues.length:null;
    const cycle=cycleValues.length?cycleValues.reduce((a,b)=>a+b,0)/cycleValues.length:null;
    const velocity=completed.reduce((sum,r)=>sum+(numeric(r["Story Points"])??0),0);
    const ev=eventBySprint.get(sprint);
    return {sprint,sprintIndex:i+1,totalItems:items.length,noReopen,reopened,leadTime:lead,leadP85:inclusiveP85(leadValues),cycleTime:cycle,efficiency:lead!==null&&lead>0&&cycle!==null?(cycle/lead)*100:null,velocity,phase:str(ev?.["Fase"]),context:str(ev?.["Evento / contexto"]),leadP85General:overallP85,leadMeanGeneral:overallLeadMean,leadCount:leadValues.length,cycleCount:cycleValues.length,throughput:noReopen+reopened};
  });
}
function FlowTooltip({active,payload,label}:any) {
  if(!active||!payload?.length)return null;
  const row=payload[0]?.payload||{};
  return <div style={{background:"#101d33",border:"1px solid #304565",borderRadius:10,padding:"10px 12px",color:"#f1f5ff",fontSize:11,maxWidth:260}}>
    <strong>{label}</strong>{row.phase&&<div style={{marginTop:5,color:"#c4b5fd"}}>Fase: {row.phase}</div>}
    {row.context&&<div style={{marginTop:4,color:"#b5c5de",whiteSpace:"normal"}}>Evento / contexto: {row.context}</div>}
    {payload.map((p:any,i:number)=><div key={i} style={{marginTop:4,color:p.color||"#e5edfa"}}>{p.name}: {typeof p.value==="number"?p.value.toFixed(p.name==="Velocity"?0:1):p.value}{p.name?.includes("Time")?" dias corridos":p.name==="Eficiência de fluxo"?"%":""}</div>)}
  </div>;
}
function PhaseBands({data}:{data:any[]}) {
  const groups:any[]=[];
  data.forEach((d,i)=>{const phase=d.phase||"Sem fase";const last=groups[groups.length-1];if(last&&last.phase===phase)last.end=d.sprint;else groups.push({phase,start:d.sprint,end:d.sprint,fill:phase==="Sem fase"?"#8293ad":(["Planejamento","Descoberta"].includes(phase)?"#4f86c6":(["Desenvolvimento","Construção"].includes(phase)?"#7b6bd6":(["Validação","Testes"].includes(phase)?"#3d9f8b":"#c28b4b")))});});
  return <>{groups.map((g,i)=><ReferenceArea key={g.start+"-"+i} x1={g.start} x2={g.end} yAxisId="left" ifOverflow="hidden" fill={g.fill} fillOpacity={0.12} strokeOpacity={0}/>)}</>;
}
function FlowChart({title,description,data,fullData,kind,unavailable}:{title:string;description:string;data:any[];fullData:any[];kind:"throughput"|"leadTime"|"cycleTime"|"velocity";unavailable:string}) {
  const hasValues=data.some(d=>kind==="throughput"?d.totalItems>0:kind==="leadTime"?d.leadTime!==null:kind==="cycleTime"?d.cycleTime!==null:d.velocity>0);
  const total=data.reduce((sum,d)=>sum+d.noReopen+d.reopened,0);
  const latest3=fullData.slice(-3); const netLast3=latest3.reduce((sum,d)=>sum+d.noReopen,0)/Math.max(1,latest3.length);
  return <section className="ap-chart-card ap-metric-card">
    <div className="ap-chart-heading"><div><div className="ap-chart-title"><Activity size={19}/><h2>{title}</h2></div><p>{description}</p></div></div>
    {kind==="throughput"&&<div className="ap-summary ap-flow-cards"><div><span>TOTAL DE ITENS</span><strong>{total}<small> itens</small></strong></div><div><span>MÉDIA POR SPRINT</span><strong>{(total/Math.max(1,data.length)).toFixed(1)}<small> itens</small></strong></div><div><span>THROUGHPUT LÍQUIDO MÉDIO · ÚLTIMAS 3</span><strong>{netLast3.toFixed(1)}<small> itens</small></strong></div></div>}
    {kind==="leadTime"&&<div className="ap-summary ap-flow-cards"><div><span>MÉDIA GERAL</span><strong>{data[0]?.leadMeanGeneral==null?"—":data[0].leadMeanGeneral.toFixed(1)}<small> dias corridos</small></strong></div><div><span>P85 GERAL · INCLUSIVO</span><strong>{data[0]?.leadP85General==null?"—":data[0].leadP85General.toFixed(1)}<small> dias corridos</small></strong></div></div>}
    {hasValues?<div className="ap-chart-scroll"><div className="ap-chart-inner" style={{width:Math.max(980,data.length*58),height:340}}><ResponsiveContainer width="100%" height={330}><ComposedChart data={data} margin={{top:18,right:20,left:0,bottom:8}}>
      <CartesianGrid stroke="#263853" strokeDasharray="3 5" vertical={false}/>
      <XAxis dataKey="sprint" tick={{fill:"#91a4c4",fontSize:10}} axisLine={{stroke:"#31435f"}} tickLine={false} interval={0}/>
      <YAxis yAxisId="left" tick={{fill:"#91a4c4",fontSize:10}} axisLine={false} tickLine={false} width={48} />
      {(kind==="cycleTime")&&<YAxis yAxisId="right" orientation="right" tick={{fill:"#91a4c4",fontSize:10}} axisLine={false} tickLine={false} width={48} unit="%"/>}
      <PhaseBands data={data}/>
      <Tooltip content={<FlowTooltip/>}/>
      {kind==="throughput"&&<><Bar yAxisId="left" dataKey="noReopen" name="Itens sem reabertura" stackId="throughput" fill="#38bdf8" maxBarSize={32}/><Bar yAxisId="left" dataKey="reopened" name="Itens reabertos" stackId="throughput" fill="#c084fc" radius={[5,5,0,0]} maxBarSize={32}/></>}
      {kind==="leadTime"&&<><Line yAxisId="left" type="monotone" dataKey="leadTime" name="Lead Time médio (dias corridos)" stroke="#a78bfa" strokeWidth={2.5} connectNulls dot={{r:3}}/><Line yAxisId="left" type="monotone" dataKey="leadP85" name="P85 por sprint (dias corridos)" stroke="#f59e0b" strokeWidth={2} strokeDasharray="5 4" connectNulls dot={false}/></>}
      {kind==="cycleTime"&&<><Line yAxisId="left" type="monotone" dataKey="cycleTime" name="Cycle Time médio (dias corridos)" stroke="#38bdf8" strokeWidth={2.5} connectNulls dot={{r:3}}/><Line yAxisId="right" type="monotone" dataKey="efficiency" name="Eficiência de fluxo" stroke="#34d399" strokeWidth={2} connectNulls dot={false}/></>}
      {kind==="velocity"&&<Bar yAxisId="left" dataKey="velocity" name="Velocity" fill="#34d399" radius={[5,5,0,0]} maxBarSize={32}/>}
    </ComposedChart></ResponsiveContainer></div></div>:<div className="ap-empty ap-metric-empty">{unavailable}</div>}
    <div className="ap-phase-legend">{[...new Set(data.map(d=>d.phase||"Sem fase"))].map(phase=><span key={phase}><i/>{phase}</span>)}</div>
  </section>;
}


const WIP_STAGES=["In development","Waiting for QA","In QA review","In code review","Waiting for review"];
const WIP_STAGE_COLORS:Record<string,string>={"In development":"#60a5fa","Waiting for QA":"#f59e0b","In QA review":"#a78bfa","In code review":"#2dd4bf","Waiting for review":"#fb923c"};
function findField(row:Row,patterns:RegExp[]):string { return Object.keys(row).find(k=>patterns.some(p=>p.test(normalize(k))))||""; }
function getWipItems(rows:Row[],snapshotDate:string) {
  const sample=rows[0]||{};
  const keyField=findField(sample,[/^issue key$/,/issue.*key/,/^chave$/]);
  const stageField=findField(sample,[/^etapa atual$/,/etapa.*atual/,/^current stage$/,/^stage$/,/^etapa$/]);
  const startField=findField(sample,[/data.*inicio/,/inicio.*data/,/start date/,/started at/,/inicio da etapa/,/data de entrada/]);
  const roleField=findField(sample,[/^papel$/, /papel responsavel/,/^role$/]);
  const summaryField=findField(sample,[/^resumo$/, /^summary$/]);
  const snap=dateValue(snapshotDate);
  return rows.map((r,i)=>{
    const start=dateValue(r[startField]);
    const age=snap&&start&&snap>=start?Math.floor((snap.getTime()-start.getTime())/86400000):null;
    const stage=str(r[stageField]);
    return {key:str(r[keyField])||"Item "+(i+1),summary:str(r[summaryField]),stage,role:str(r[roleField])||"—",start:start?str(r[startField]):"",age,stageField,startField,keyField,roleField};
  }).filter(r=>!!r.key||!!r.stage||r.age!==null).sort((a,b)=>(b.age??-1)-(a.age??-1));
}
function wipReference(rows:Row[]) {
  const valid=rows.map(r=>({sprint:str(r["Sprint"]),n:numeric(r["Sprint nº"])??sprintNum(str(r["Sprint"])),cycle:numeric(r["Cycle time (dias corridos)"])})).filter(r=>r.sprint&&r.n!==Number.MAX_SAFE_INTEGER);
  const sprints=[...new Set(valid.map(r=>r.sprint))].sort((a,b)=>sprintNum(a)-sprintNum(b)||a.localeCompare(b,"pt-BR",{numeric:true}));
  const last3=sprints.slice(-3);
  const cycleValues=valid.filter(r=>last3.includes(r.sprint)&&r.cycle!==null&&r.cycle>=0).map(r=>r.cycle as number);
  const cycleMean=cycleValues.length?cycleValues.reduce((a,b)=>a+b,0)/cycleValues.length:null;
  const latest=sprints[sprints.length-1]||"";
  const latestRows=rows.filter(r=>str(r["Sprint"])===latest);
  const completed=latestRows.filter(isCompletedRow);
  const throughput=completed.length;
  const leads=completed.map(r=>numeric(r["Lead time (dias corridos)"])??elapsedDays(r["Criado"]??r["Created Date"],r["Concluído"]??r["Resolution Date"])).filter((v):v is number=>v!==null&&v>=0);
  return {cycleMean,alertLimit:cycleMean===null?null:cycleMean*2,latestSprint:latest,throughput,leadMean:leads.length?leads.reduce((a,b)=>a+b,0)/leads.length:null,expected:leads.length?(throughput/14)*(leads.reduce((a,b)=>a+b,0)/leads.length):null};
}
function WipAgeTooltip({active,payload}:{active?:boolean;payload?:any[]}) {
  if(!active||!payload?.length)return null;
  const r=payload[0].payload;
  return <div style={{background:"#101d33",border:"1px solid #304565",borderRadius:10,padding:"10px 12px",color:"#f1f5ff",fontSize:12}}>
    <strong>{r.key}</strong><div style={{marginTop:5}}>Etapa: {r.stage||"Não informada"}</div><div>Idade: {r.age===null?"Sem data":r.age+" dias corridos"}</div><div>Papel: {r.role}</div>
  </div>;
}
function WipAgingPage({rows,snapshotDate,baseRows,source}:{rows:Row[];snapshotDate:string;baseRows:Row[];source:"demo"|"imported"}) {
  const items=useMemo(()=>getWipItems(rows,snapshotDate),[rows,snapshotDate]);
  const ref=useMemo(()=>wipReference(baseRows),[baseRows]);
  const limit=ref.alertLimit;
  const alertCount=items.filter(i=>i.age!==null&&limit!==null&&i.age>limit).length;
  const latestWip=items.length;
  const difference=ref.expected===null?null:latestWip-ref.expected;
  const stageField=items[0]?.stageField||"";
  const startField=items[0]?.startField||"";
  const hasWipRows=rows.length>0;
  return <div className="ap-wip-page">
    <div className="ap-wip-top-grid">
      <section className="ap-panel ap-wip-panel">
        <div className="ap-panel-heading"><div><h2>WIP e Aging</h2><p>Itens em andamento no snapshot e tempo desde o início, em dias corridos.</p></div><span className="ap-wip-date">{snapshotDate||"Snapshot não informado"}</span></div>
        <div className="ap-wip-kpis"><div><span>DATA DO SNAPSHOT</span><strong>{snapshotDate||"—"}</strong></div><div><span>WIP ATUAL</span><strong>{latestWip}<small> itens</small></strong></div><div className={alertCount?"ap-wip-alert-kpi":""}><span>ITENS EM ALERTA</span><strong>{alertCount}<small> itens</small></strong></div></div>
        {!hasWipRows?<div className="ap-empty">Envie a aba <strong>WIP_Aberto</strong> para habilitar esta página.</div>:!stageField||!startField?<div className="ap-empty">Não encontrei as colunas de etapa atual e/ou data de início na aba WIP_Aberto. Confira os cabeçalhos da planilha.</div>:!snapshotDate?<div className="ap-empty">A célula B3 da aba WIP_Aberto deve conter a data do snapshot para calcular a idade dos itens.</div>:<>
          <div className="ap-wip-chart-heading"><div><strong>Idade dos itens em andamento</strong><span>Do mais velho ao mais novo · linha de alerta: {limit===null?"indisponível":limit.toFixed(1)+" dias corridos"}</span></div><div className="ap-wip-stage-legend">{WIP_STAGES.map(s=><span key={s}><i style={{background:WIP_STAGE_COLORS[s]}}/>{s}</span>)}</div></div>
          <div className="ap-wip-chart-scroll"><div style={{height:Math.max(180,items.length*42),minWidth:640}}><ResponsiveContainer width="100%" height="100%"><BarChart data={items} layout="vertical" margin={{top:10,right:35,left:22,bottom:10}} barCategoryGap="22%">
            <CartesianGrid stroke="#263853" strokeDasharray="3 5" horizontal={false}/>
            <XAxis type="number" tick={{fill:"#91a4c4",fontSize:10}} axisLine={{stroke:"#31435f"}} tickLine={false} label={{value:"Idade (dias corridos)",position:"insideBottom",offset:-4,fill:"#91a4c4",fontSize:10}}/>
            <YAxis type="category" dataKey="key" reversed width={100} tick={{fill:"#cbd8ed",fontSize:10}} axisLine={false} tickLine={false} />
            <Tooltip content={<WipAgeTooltip/>}/>
            {limit!==null&&<ReferenceLine x={limit} stroke="#f87171" strokeWidth={2} strokeDasharray="6 4" label={{value:"Limite de alerta",position:"top",fill:"#f87171",fontSize:10}}/>}
            <Bar dataKey="age" name="Idade (dias corridos)" maxBarSize={24} radius={[0,5,5,0]} isAnimationActive={false}>
              {items.map((item,i)=><Cell key={item.key+"-"+i} fill={item.age!==null&&limit!==null&&item.age>limit?"#ef4444":WIP_STAGE_COLORS[item.stage]||"#8293ad"}/>)}
            </Bar>
          </BarChart></ResponsiveContainer></div></div>
          <div className="ap-table-scroll ap-wip-table"><table><thead><tr><th>Issue key</th><th>Etapa</th><th>Idade</th><th>Papel</th></tr></thead><tbody>{items.map((item,i)=><tr key={item.key+"-"+i}><td>{item.key}</td><td><span className="ap-wip-stage"><i style={{background:WIP_STAGE_COLORS[item.stage]||"#8293ad"}}/>{item.stage||"—"}</span></td><td className={item.age!==null&&limit!==null&&item.age>limit?"ap-wip-age-alert":""}>{item.age===null?"Sem data":item.age+" dias corridos"}</td><td>{item.role}</td></tr>)}</tbody></table></div>
        </>}
      </section>
      <section className="ap-panel ap-little-panel">
        <div className="ap-panel-heading"><div><h2>Lei de Little</h2><p>WIP observado versus WIP estimado pelo fluxo recente.</p></div></div>
        <div className="ap-little-formula">WIP estimado = (Throughput da última sprint ÷ 14) × Lead Time médio da última sprint</div>
        <div className="ap-little-compare"><div><span>WIP ATUAL</span><strong>{latestWip}</strong><small>itens</small></div><div><span>WIP ESTIMADO</span><strong>{ref.expected===null?"—":ref.expected.toFixed(1)}</strong><small>itens</small></div></div>
        <div className={"ap-little-difference "+(difference!==null&&difference>0?"is-high":difference!==null&&difference<0?"is-low":"")}><span>Diferença (atual − estimado)</span><strong>{difference===null?"—":(difference>0?"+":"")+difference.toFixed(1)+" itens"}</strong></div>
        <div className="ap-little-details"><div><span>Última sprint</span><strong>{ref.latestSprint||"—"}</strong></div><div><span>Throughput</span><strong>{ref.throughput} itens</strong></div><div><span>Lead Time médio</span><strong>{ref.leadMean===null?"—":ref.leadMean.toFixed(1)+" dias corridos"}</strong></div><div><span>Cycle Time médio · últimas 3 sprints</span><strong>{ref.cycleMean===null?"—":ref.cycleMean.toFixed(1)+" dias corridos"}</strong></div><div><span>Limite de alerta · 2× Cycle Time</span><strong>{limit===null?"—":limit.toFixed(1)+" dias corridos"}</strong></div></div>
        {source==="demo"&&<p className="ap-note">Os dados demonstrativos atuais não incluem a aba WIP_Aberto. Importe a planilha para visualizar o aging real.</p>}
      </section>
    </div>
  </div>;
}


const DEFECT_ORIGINS=["QA interno","Code review","Produção (achado pelo time)","Escola-piloto"];
const QUALITY_COLORS:Record<string,string>={"QA interno":"#38bdf8","Code review":"#a78bfa","Produção (achado pelo time)":"#fb923c","Escola-piloto":"#f87171"};
const QUALITY_PHASES=[{label:"Todas",min:1,max:20},{label:"S1–S3",min:1,max:3},{label:"S4–S10",min:4,max:10},{label:"S11–S17",min:11,max:17},{label:"S18–S20",min:18,max:20}];
function premiseValue(rows:Row[],patterns:RegExp[]):number|null {
  const row=rows.find(r=>patterns.some(p=>p.test(normalize(r["Parâmetro"]??r["Parametro"]??r["Parâmetro / indicador"]??""))));
  if(!row)return null;
  const raw=row["Valor"]??row["valor"];
  if(typeof raw==="number"&&Number.isFinite(raw))return raw;
  const match=str(raw).replace(/\s/g,"").replace(/R\$/i,"").match(/-?\d+(?:[.,]\d+)?/);
  if(!match)return null;
  const value=Number(match[0].replace(",","."));
  return Number.isFinite(value)?value:null;
}
function qualitySprintNum(row:Row,field:string) { return sprintNum(str(row[field])); }
function QualityTooltip({active,payload,label}:{active?:boolean;payload?:any[];label?:string}) {
  if(!active||!payload?.length)return null;
  return <div style={{background:"#101d33",border:"1px solid #304565",borderRadius:10,padding:"10px 12px",color:"#f1f5ff",fontSize:11}}>
    <strong>{label}</strong>{payload.map((p,i)=><div key={String(p.dataKey)+i} style={{marginTop:5,color:p.color||"#dce7f8"}}>{p.name||p.dataKey}: {typeof p.value==="number"?p.value.toLocaleString("pt-BR",{maximumFractionDigits:2}):String(p.value??"—")}{String(p.dataKey).toLowerCase().includes("rate")||String(p.dataKey).toLowerCase().includes("percent")||String(p.dataKey).toLowerCase().includes("acumulado")?"%":""}</div>)}
  </div>;
}
function QualityLegend({items}:{items:{label:string;color:string}[]}) {
  return <div className="ap-quality-legend">{items.map(item=><span key={item.label}><i style={{background:item.color}}/>{item.label}</span>)}</div>;
}
function QualityReworkPage({defects,rework,calendar,premises,baseRows,source}:{defects:Row[];rework:Row[];calendar:Row[];premises:Row[];baseRows:Row[];source:"demo"|"imported"}) {
  const [phase,setPhase]=useState("Todas");
  const defectTarget=premiseValue(premises,[/meta.*taxa.*defeit/,/taxa.*defeit.*meta/]);
  const escapeTarget=premiseValue(premises,[/meta.*taxa.*escape/,/taxa.*escape.*meta/,/meta.*escape/]);
  const reworkTarget=premiseValue(premises,[/meta.*retrabalho.*capacidade/,/meta.*capacidade.*retrabalho/,/meta.*retrabalho/]);
  const hourlyCost=premiseValue(premises,[/custo.*medio.*por hora/,/custo.*hora/,/valor.*hora/]);
  const sprintLabels=useMemo(()=>{
    const all=[...baseRows,...defects,...rework,...calendar].map(r=>str(r["Sprint"]??r["Sprint de detecção"]??r["Sprint do item"])).filter(Boolean);
    const unique=[...new Set(all)];
    return unique.sort((a,b)=>sprintNum(a)-sprintNum(b)||a.localeCompare(b,"pt-BR",{numeric:true}));
  },[baseRows,defects,rework,calendar]);
  const data=useMemo(()=>sprintLabels.map(sprint=>{
    const d=defects.filter(r=>str(r["Sprint de detecção"])===sprint);
    const rw=rework.filter(r=>str(r["Sprint do item"])===sprint);
    const cal=calendar.find(r=>str(r["Sprint"])===sprint);
    const throughput=baseRows.filter(r=>str(r["Sprint"])===sprint&&isCompletedRow(r)).length;
    const originCounts=Object.fromEntries(DEFECT_ORIGINS.map(o=>[o,d.filter(r=>normalize(r["Origem"])===normalize(o)).length])) as Record<string,number>;
    const defectCount=d.length, school=d.filter(r=>normalize(r["Origem"])===normalize("Escola-piloto")).length;
    const correction=d.reduce((sum,r)=>sum+(numeric(r["Horas de correção"])??0),0);
    const restore=d.reduce((sum,r)=>sum+(numeric(r["Horas para restaurar"])??0),0);
    const reworkHours=rw.reduce((sum,r)=>sum+(numeric(r["Horas de retrabalho"])??0),0);
    const capacity=numeric(cal?.["Capacidade (h)"]);
    const totalHours=correction+reworkHours;
    return {sprint,throughput,defectCount,school,originCounts,defectRate:throughput>0?defectCount/throughput*100:null,escapeRate:defectCount>0?school/defectCount*100:0,correction,reworkHours,restore,totalHours,capacity,reworkCapacity:capacity&&capacity>0?totalHours/capacity*100:null,cost:hourlyCost===null?null:totalHours*hourlyCost};
  }),[sprintLabels,defects,rework,calendar,baseRows,hourlyCost]);
  const pareto=useMemo(()=>{
    const selected=QUALITY_PHASES.find(p=>p.label===phase)||QUALITY_PHASES[0];
    const filtered=rework.filter(r=>{const n=qualitySprintNum(r,"Sprint do item");return n>=selected.min&&n<=selected.max;});
    const counts=new Map<string,number>();
    filtered.forEach(r=>{const cause=str(r["Causa"])||"Causa não informada";counts.set(cause,(counts.get(cause)||0)+1);});
    const sorted=[...counts].map(([cause,count])=>({cause,count})).sort((a,b)=>b.count-a.count||a.cause.localeCompare(b.cause,"pt-BR"));
    const total=sorted.reduce((s,r)=>s+r.count,0);let running=0;
    return sorted.map(r=>{running+=r.count;return {...r,accumulated:total?running/total*100:0};});
  },[rework,phase]);
  const totalDefects=data.reduce((sum,r)=>sum+r.defectCount,0);
  const totalReworkHours=data.reduce((sum,r)=>sum+r.reworkHours,0);
  const totalCorrectionHours=data.reduce((sum,r)=>sum+r.correction,0);
  const totalCapacity=data.reduce((sum,r)=>sum+(r.capacity||0),0);
  const totalCost=hourlyCost===null?null:(totalReworkHours+totalCorrectionHours)*hourlyCost;
  const missing=!defects.length&&!rework.length&&!calendar.length&&!premises.length;
  return <div className="ap-quality-page">
    {source==="demo"&&<section className="ap-panel ap-quality-missing"><strong>Importe a planilha para visualizar os indicadores de qualidade.</strong><p>Esta página utiliza as abas Defeitos, Retrabalho, Calendario_Eventos e Premissas. Nenhum valor demonstrativo foi inventado.</p></section>}
    {source==="imported"&&missing&&<section className="ap-panel ap-quality-missing"><strong>As abas de qualidade não foram encontradas.</strong><p>Inclua Defeitos, Retrabalho, Calendario_Eventos e Premissas no arquivo .xlsx.</p></section>}
    <div className="ap-quality-kpis">
      <div><span>DEFEITOS REGISTRADOS</span><strong>{totalDefects}</strong><small>Origem registrada na aba Defeitos</small></div>
      <div><span>HORAS DE RETRABALHO</span><strong>{totalReworkHours.toLocaleString("pt-BR",{maximumFractionDigits:1})} h</strong><small>Retrabalho informado na aba Retrabalho</small></div>
      <div><span>HORAS DE CORREÇÃO</span><strong>{totalCorrectionHours.toLocaleString("pt-BR",{maximumFractionDigits:1})} h</strong><small>Correção informada na aba Defeitos</small></div>
      <div><span>CUSTO ESTIMADO DO RETRABALHO</span><strong>{totalCost===null?"—":totalCost.toLocaleString("pt-BR",{style:"currency",currency:"BRL"})}</strong><small>{hourlyCost===null?"Defina Custo médio por hora em Premissas":("Custo médio por hora: "+hourlyCost.toLocaleString("pt-BR",{style:"currency",currency:"BRL"}))}</small></div>
    </div>
    <div className="ap-quality-chart-grid">
      <section className="ap-panel ap-quality-chart-card">
        <div className="ap-panel-heading"><div><h2>Defeitos por sprint e taxa de defeitos</h2><p>Barras empilhadas por origem; taxa = defeitos ÷ throughput bruto da sprint.</p></div></div>
        <QualityLegend items={DEFECT_ORIGINS.map(label=>({label,color:QUALITY_COLORS[label]})).concat([{label:"Taxa de defeitos (%)",color:"#facc15"},{label:"Meta de taxa (%)",color:"#f87171"}])}/>
        {data.length?<div className="ap-quality-chart-scroll"><div style={{width:Math.max(650,data.length*62),height:310}}><ResponsiveContainer width="100%" height="100%"><ComposedChart data={data} margin={{top:22,right:28,left:0,bottom:14}}>
          <CartesianGrid stroke="#263853" strokeDasharray="3 5" vertical={false}/><XAxis dataKey="sprint" tick={{fill:"#91a4c4",fontSize:10}} interval={0} axisLine={{stroke:"#31435f"}} tickLine={false}/><YAxis yAxisId="count" tick={{fill:"#91a4c4",fontSize:10}} axisLine={false} tickLine={false} allowDecimals={false}/><YAxis yAxisId="rate" orientation="right" tick={{fill:"#91a4c4",fontSize:10}} axisLine={false} tickLine={false} unit="%"/><Tooltip content={<QualityTooltip/>}/>
          {DEFECT_ORIGINS.map(origin=><Bar key={origin} yAxisId="count" dataKey={"originCounts."+origin} name={origin} stackId="defects" fill={QUALITY_COLORS[origin]} maxBarSize={34}/>)}
          <Line yAxisId="rate" type="monotone" dataKey="defectRate" name="Taxa de defeitos (%)" stroke="#facc15" strokeWidth={2.5} connectNulls dot={{r:3}}/>
          {defectTarget!==null&&<ReferenceLine yAxisId="rate" y={defectTarget} stroke="#f87171" strokeDasharray="6 4" label={{value:"Meta: taxa de defeitos",fill:"#f87171",fontSize:9,position:"insideTopRight"}}/>}
        </ComposedChart></ResponsiveContainer></div></div>:<div className="ap-empty">Não há sprints para calcular os defeitos.</div>}
        <p className="ap-quality-note">Taxa expressa em percentual. A linha de meta usa Premissas: “Meta: taxa de defeitos”.</p>
      </section>
      <section className="ap-panel ap-quality-chart-card">
        <div className="ap-panel-heading"><div><h2>Taxa de escape por sprint</h2><p>Defeitos identificados na Escola-piloto ÷ total de defeitos da sprint.</p></div></div>
        <QualityLegend items={[{label:"Taxa de escape (%)",color:"#c084fc"},{label:"Meta de escape (%)",color:"#f87171"}]}/>
        {data.length?<div className="ap-quality-chart-scroll"><div style={{width:Math.max(650,data.length*62),height:310}}><ResponsiveContainer width="100%" height="100%"><ComposedChart data={data} margin={{top:22,right:22,left:0,bottom:14}}>
          <CartesianGrid stroke="#263853" strokeDasharray="3 5" vertical={false}/><XAxis dataKey="sprint" tick={{fill:"#91a4c4",fontSize:10}} interval={0} axisLine={{stroke:"#31435f"}} tickLine={false}/><YAxis tick={{fill:"#91a4c4",fontSize:10}} axisLine={false} tickLine={false} unit="%"/><Tooltip content={<QualityTooltip/>}/><Bar dataKey="escapeRate" name="Taxa de escape (%)" fill="#c084fc" radius={[5,5,0,0]} maxBarSize={34}/>
          {escapeTarget!==null&&<ReferenceLine y={escapeTarget} stroke="#f87171" strokeDasharray="6 4" label={{value:"Meta de escape",fill:"#f87171",fontSize:9,position:"insideTopRight"}}/>}
        </ComposedChart></ResponsiveContainer></div></div>:<div className="ap-empty">Não há dados de defeitos por sprint.</div>}
        <p className="ap-quality-note">A taxa é 0% quando a sprint não tem defeitos. Meta lida de Premissas, quando preenchida.</p>
      </section>
      <section className="ap-panel ap-quality-chart-card">
        <div className="ap-panel-heading"><div><h2>Retrabalho em % da capacidade</h2><p>(horas de retrabalho + horas de correção) ÷ capacidade da sprint.</p></div></div>
        <QualityLegend items={[{label:"Esforço de retrabalho e correção (%)",color:"#34d399"},{label:"Meta de retrabalho (%)",color:"#f87171"}]}/>
        {data.length?<div className="ap-quality-chart-scroll"><div style={{width:Math.max(650,data.length*62),height:310}}><ResponsiveContainer width="100%" height="100%"><ComposedChart data={data} margin={{top:22,right:22,left:0,bottom:14}}>
          <CartesianGrid stroke="#263853" strokeDasharray="3 5" vertical={false}/><XAxis dataKey="sprint" tick={{fill:"#91a4c4",fontSize:10}} interval={0} axisLine={{stroke:"#31435f"}} tickLine={false}/><YAxis tick={{fill:"#91a4c4",fontSize:10}} axisLine={false} tickLine={false} unit="%"/><Tooltip content={<QualityTooltip/>}/><Bar dataKey="reworkCapacity" name="Esforço sobre capacidade (%)" fill="#34d399" radius={[5,5,0,0]} maxBarSize={34}/>
          {reworkTarget!==null&&<ReferenceLine y={reworkTarget} stroke="#f87171" strokeDasharray="6 4" label={{value:"Meta de retrabalho",fill:"#f87171",fontSize:9,position:"insideTopRight"}}/>}
        </ComposedChart></ResponsiveContainer></div></div>:<div className="ap-empty">Não há sprints para calcular retrabalho.</div>}
        <p className="ap-quality-note">Sem capacidade válida na aba Calendario_Eventos, a taxa daquela sprint fica indisponível. Meta de retrabalho lida de Premissas, quando disponível.</p>
      </section>
      <section className="ap-panel ap-quality-chart-card">
        <div className="ap-panel-heading"><div><h2>Pareto das causas de retrabalho</h2><p>Ocorrências em ordem decrescente e percentual acumulado.</p></div><label className="ap-quality-select">Fase <select value={phase} onChange={e=>setPhase(e.target.value)}>{QUALITY_PHASES.map(p=><option key={p.label}>{p.label}</option>)}</select></label></div>
        <QualityLegend items={[{label:"Ocorrências por causa",color:"#60a5fa"},{label:"% acumulado",color:"#facc15"}]}/>
        {pareto.length?<div className="ap-quality-chart-scroll"><div style={{width:Math.max(650,pareto.length*72),height:320}}><ResponsiveContainer width="100%" height="100%"><ComposedChart data={pareto} margin={{top:20,right:24,left:0,bottom:62}}>
          <CartesianGrid stroke="#263853" strokeDasharray="3 5" vertical={false}/><XAxis dataKey="cause" tick={{fill:"#91a4c4",fontSize:9}} interval={0} angle={-30} textAnchor="end" height={76} axisLine={{stroke:"#31435f"}} tickLine={false}/><YAxis yAxisId="count" tick={{fill:"#91a4c4",fontSize:10}} axisLine={false} tickLine={false} allowDecimals={false}/><YAxis yAxisId="percent" orientation="right" domain={[0,100]} tick={{fill:"#91a4c4",fontSize:10}} axisLine={false} tickLine={false} unit="%"/><Tooltip content={<QualityTooltip/>}/>
          <Bar yAxisId="count" dataKey="count" name="Ocorrências" fill="#60a5fa" radius={[4,4,0,0]} maxBarSize={42}/><Line yAxisId="percent" type="monotone" dataKey="accumulated" name="% acumulado" stroke="#facc15" strokeWidth={2.5} dot={{r:3}}/>
        </ComposedChart></ResponsiveContainer></div></div>:<div className="ap-empty">Não há causas de retrabalho para a fase selecionada.</div>}
        <p className="ap-quality-note">Filtro aplicado à Sprint do item: {phase}. Causas vazias são agrupadas como “Causa não informada”.</p>
      </section>
    </div>
    <section className="ap-panel ap-quality-cost">
      <div className="ap-panel-heading"><div><h2>Custo do retrabalho por sprint</h2><p>(Horas de retrabalho + horas de correção de defeitos) × custo médio por hora.</p></div><strong className="ap-quality-cost-total">{totalCost===null?"Custo/h não informado":totalCost.toLocaleString("pt-BR",{style:"currency",currency:"BRL"})}</strong></div>
      {data.length?<div className="ap-table-scroll"><table><thead><tr><th>Sprint</th><th>Horas de retrabalho</th><th>Horas de correção</th><th>Capacidade (h)</th><th>Custo estimado</th></tr></thead><tbody>{data.map(r=><tr key={r.sprint}><td>{r.sprint}</td><td>{r.reworkHours.toLocaleString("pt-BR",{maximumFractionDigits:2})} h</td><td>{r.correction.toLocaleString("pt-BR",{maximumFractionDigits:2})} h</td><td>{r.capacity===null?"—":r.capacity.toLocaleString("pt-BR",{maximumFractionDigits:2})+" h"}</td><td>{r.cost===null?"—":r.cost.toLocaleString("pt-BR",{style:"currency",currency:"BRL"})}</td></tr>)}</tbody></table></div>:<div className="ap-empty">Importe dados para calcular o custo por sprint.</div>}
      <p className="ap-quality-note">Capacidade total informada: {totalCapacity.toLocaleString("pt-BR",{maximumFractionDigits:1})} h. Valores são estimativas com base nas horas registradas e em “Custo médio por hora” na aba Premissas.</p>
    </section>
  </div>;
}


const BOTTLENECK_STAGES=[
  {key:"dev",label:"In development",field:"In development",kind:"work",color:"#3b82f6"},
  {key:"waitQa",label:"Waiting for QA",field:"Waiting for QA",kind:"wait",color:"#94a3b8"},
  {key:"qa",label:"In QA review",field:"In QA review",kind:"work",color:"#60a5fa"},
  {key:"code",label:"In code review",field:"In code review",kind:"work",color:"#2563eb"},
  {key:"waitReview",label:"Waiting for review",field:"Waiting for review",kind:"wait",color:"#64748b"}
];
function bottleneckPremisePercent(rows:Row[]):number|null {
  const row=rows.find(r=>/meta.*(change failure|falha|failure)/.test(normalize(r["Parâmetro"]??r["Parametro"]??"")));
  if(!row)return null;
  const raw=str(row["Valor"]??row["valor"]);
  const value=numeric(raw.replace(/%/g,""));
  if(value===null)return null;
  return raw.includes("%")?value:(value>0&&value<=1?value*100:value);
}
function BottleneckDoraTooltip({active,payload,label}:{active?:boolean;payload?:any[];label?:string}) {
  if(!active||!payload?.length)return null;
  return <div style={{background:"#101d33",border:"1px solid #304565",borderRadius:10,padding:"10px 12px",color:"#f1f5ff",fontSize:11}}>
    <strong>{label}</strong>{payload.filter(p=>p.value!==null&&p.value!==undefined).map((p,i)=><div key={String(p.dataKey)+i} style={{marginTop:5,color:p.color||"#dce7f8"}}>{p.name||p.dataKey}: {typeof p.value==="number"?p.value.toLocaleString("pt-BR",{maximumFractionDigits:2}):String(p.value)}{String(p.dataKey).toLowerCase().includes("rate")||String(p.dataKey).toLowerCase().includes("percent")||String(p.dataKey).toLowerCase().includes("failure")?"%":""}</div>)}
  </div>;
}
function BottleneckDoraPage({baseRows,defects,doraSprints,premises,source}:{baseRows:Row[];defects:Row[];doraSprints:Row[];premises:Row[];source:"demo"|"imported"}) {
  const [period,setPeriod]=useState("Todas");
  const failureTarget=bottleneckPremisePercent(premises);
  const sprintLabels=useMemo(()=>{
    const all=[...baseRows,...defects,...doraSprints].map(r=>str(r["Sprint"]??r["Sprint de detecção"])).filter(Boolean);
    return [...new Set(all)].sort((a,b)=>sprintNum(a)-sprintNum(b)||a.localeCompare(b,"pt-BR",{numeric:true}));
  },[baseRows,defects,doraSprints]);
  const data=useMemo(()=>sprintLabels.map(sprint=>{
    const items=baseRows.filter(r=>str(r["Sprint"])===sprint);
    const stageMeans:Record<string,number|null>={};
    BOTTLENECK_STAGES.forEach(stage=>{
      const values=items.map(r=>numeric(r[stage.field])).filter((v):v is number=>v!==null&&v>=0);
      stageMeans[stage.key]=values.length?values.reduce((a,b)=>a+b,0)/values.length:null;
    });
    const queueValues=items.map(r=>{
      const lead=numeric(r["Lead time (dias corridos)"]),cycle=numeric(r["Cycle time (dias corridos)"]);
      return lead!==null&&cycle!==null?Math.max(0,lead-cycle):null;
    }).filter((v):v is number=>v!==null);
    const queue=queueValues.length?queueValues.reduce((a,b)=>a+b,0)/queueValues.length:null;
    const cycleValues=items.map(r=>numeric(r["Cycle time (dias corridos)"])).filter((v):v is number=>v!==null&&v>=0);
    const leadValues=items.map(r=>numeric(r["Lead time (dias corridos)"])).filter((v):v is number=>v!==null&&v>=0);
    const cycleMean=cycleValues.length?cycleValues.reduce((a,b)=>a+b,0)/cycleValues.length:null;
    const leadMean=leadValues.length?leadValues.reduce((a,b)=>a+b,0)/leadValues.length:null;
    const slowest=BOTTLENECK_STAGES.map(s=>({label:s.label,value:stageMeans[s.key]})).filter(s=>s.value!==null).sort((a,b)=>(b.value||0)-(a.value||0))[0]?.label||"Sem dados";
    const waitValues=items.map(r=>{
      const lead=numeric(r["Lead time (dias corridos)"]);
      const wqa=numeric(r["Waiting for QA"]),wreview=numeric(r["Waiting for review"]),cyc=numeric(r["Cycle time (dias corridos)"]);
      if(lead===null||cyc===null)return null;
      const q=Math.max(0,lead-cyc);
      return {lead,wait:(wqa===null?0:wqa)+(wreview===null?0:wreview)+q};
    }).filter((v):v is {lead:number;wait:number}=>v!==null&&v.lead>0);
    const waitPct=waitValues.length?waitValues.reduce((a,b)=>a+b.wait,0)/waitValues.reduce((a,b)=>a+b.lead,0)*100:null;
    const doraRow=doraSprints.find(r=>str(r["Sprint"])===sprint);
    const deploys=numeric(doraRow?.["Deploys na sprint"]);
    const failed=defects.filter(r=>str(r["Sprint de detecção"])===sprint&&/^(sim|yes|true|1)$/i.test(str(r["Causou falha de deploy?"])));
    const allFailures=defects.filter(r=>str(r["Sprint de detecção"])===sprint&&str(r["Causou falha de deploy?"])!=="");
    const restoration=failed.map(r=>numeric(r["Horas para restaurar"])).filter((v):v is number=>v!==null&&v>=0);
    const recovery=failed.length&&restoration.length?restoration.reduce((a,b)=>a+b,0)/restoration.length:null;
    const cfr=deploys!==null&&deploys>0?failed.length/deploys*100:null;
    return {sprint,...stageMeans,queue,cycleMean,leadMean,slowest,waitPct,deploys,deploysPerWeek:deploys===null?null:deploys/2,failedCount:failed.length,reportedFailures:allFailures.length,cfr,recovery};
  }),[sprintLabels,baseRows,defects,doraSprints]);
  const shown=period==="Todas"?data:data.filter(d=>d.sprint===period);
  const waitTotals=shown.reduce((totals,d)=>{
    const items=baseRows.filter(r=>str(r["Sprint"])===d.sprint);
    return items.reduce((acc,r)=>{
      const lead=numeric(r["Lead time (dias corridos)"]),cycle=numeric(r["Cycle time (dias corridos)"]);
      if(lead===null||cycle===null||lead<=0)return acc;
      const waiting=(numeric(r["Waiting for QA"])||0)+(numeric(r["Waiting for review"])||0)+Math.max(0,lead-cycle);
      return {wait:acc.wait+waiting,lead:acc.lead+lead};
    },totals);
  },{wait:0,lead:0});
  const totalWaitPct=waitTotals.lead>0?waitTotals.wait/waitTotals.lead*100:null;
  const hasData=baseRows.length>0||defects.length>0||doraSprints.length>0;
  const doraChart=(title:string,description:string,key:string,color:string,unit:string,meta?:number|null)=> <section className="ap-panel ap-bd-chart-card">
    <div className="ap-panel-heading"><div><h2>{title}</h2><p>{description}</p></div></div>
    <div className="ap-quality-legend"><span><i style={{background:color}}/>{title}</span>{meta!==undefined&&meta!==null&&<span><i style={{background:"#f87171"}}/>Meta de falha: {meta.toLocaleString("pt-BR",{maximumFractionDigits:2})}%</span>}</div>
    {shown.length?<div className="ap-quality-chart-scroll"><div style={{width:Math.max(520,shown.length*62),height:260}}><ResponsiveContainer width="100%" height="100%"><ComposedChart data={shown} margin={{top:18,right:25,left:0,bottom:12}}>
      <CartesianGrid stroke="#263853" strokeDasharray="3 5" vertical={false}/><XAxis dataKey="sprint" tick={{fill:"#91a4c4",fontSize:10}} interval={0} axisLine={{stroke:"#31435f"}} tickLine={false}/><YAxis tick={{fill:"#91a4c4",fontSize:10}} axisLine={false} tickLine={false} unit={unit}/><Tooltip content={<BottleneckDoraTooltip/>}/>
      {meta!==undefined&&meta!==null&&<ReferenceLine y={meta} stroke="#f87171" strokeDasharray="6 4" label={{value:"Meta",fill:"#f87171",fontSize:9,position:"insideTopRight"}}/>}
      {key==="recovery"?<Line type="monotone" dataKey={key} name={title} stroke={color} strokeWidth={2.5} connectNulls={false} dot={{r:4}} activeDot={{r:5}}/>:<Bar dataKey={key} name={title} fill={color} radius={[5,5,0,0]} maxBarSize={34}/>}
    </ComposedChart></ResponsiveContainer></div></div>:<div className="ap-empty">Não há dados no período selecionado.</div>}
  </section>;
  return <div className="ap-bd-page">
    <div className="ap-bd-controls"><div><strong>Período analisado</strong><span>O filtro atualiza os gráficos e o cartão de espera.</span></div><select aria-label="Selecionar período" value={period} onChange={e=>setPeriod(e.target.value)}><option>Todas</option>{sprintLabels.map(s=><option key={s}>{s}</option>)}</select></div>
    {!hasData&&<section className="ap-panel ap-quality-missing"><strong>Importe a planilha para visualizar Gargalo e DORA.</strong><p>Esta página utiliza Base_Itens, DORA_Sprint, Defeitos e Premissas.</p></section>}
    {source==="demo"&&<p className="ap-quality-note">Os dados demonstrativos atuais não incluem as abas necessárias para calcular os indicadores DORA e o detalhamento de gargalos. Importe a planilha para preencher os gráficos.</p>}
    <section className="ap-panel ap-bd-bottleneck">
      <div className="ap-panel-heading"><div><h2>Composição do Cycle Time por sprint</h2><p>Média das cinco etapas de Base_Itens e fila inicial (Lead Time − Cycle Time), em dias corridos.</p></div><div className="ap-bd-wait-card"><span>ESPERA TOTAL / LEAD TIME</span><strong>{totalWaitPct===null?"—":totalWaitPct.toLocaleString("pt-BR",{maximumFractionDigits:1})+"%"}</strong><small>{period==="Todas"?"Período completo":period}</small></div></div>
      <div className="ap-quality-legend ap-bd-stage-legend">{BOTTLENECK_STAGES.map(s=><span key={s.key}><i style={{background:s.color}}/>{s.label}</span>)}<span><i style={{background:"#475569"}}/>Fila antes de iniciar</span></div>
      {shown.length?<div className="ap-quality-chart-scroll"><div style={{width:Math.max(650,shown.length*68),height:330}}><ResponsiveContainer width="100%" height="100%"><BarChart data={shown} margin={{top:18,right:22,left:0,bottom:14}}>
        <CartesianGrid stroke="#263853" strokeDasharray="3 5" vertical={false}/><XAxis dataKey="sprint" tick={{fill:"#91a4c4",fontSize:10}} interval={0} axisLine={{stroke:"#31435f"}} tickLine={false}/><YAxis tick={{fill:"#91a4c4",fontSize:10}} axisLine={false} tickLine={false} unit=" d"/><Tooltip content={<BottleneckDoraTooltip/>}/>
        {BOTTLENECK_STAGES.map(s=><Bar key={s.key} dataKey={s.key} name={s.label+" (média, dias)"} stackId="time" fill={s.color} maxBarSize={38}/>)}
        <Bar dataKey="queue" name="Fila antes de iniciar (média, dias)" stackId="time" fill="#475569" maxBarSize={38}/>
      </BarChart></ResponsiveContainer></div></div>:<div className="ap-empty">Não há sprints para o período selecionado.</div>}
      <div className="ap-table-scroll ap-bd-table"><table><thead><tr><th>Sprint</th><th>Etapa mais lenta no cycle time</th><th>Espera / Lead Time</th><th>Fila antes de iniciar (dias)</th></tr></thead><tbody>{shown.map(d=><tr key={d.sprint}><td>{d.sprint}</td><td>{d.slowest}</td><td>{d.waitPct===null?"—":d.waitPct.toLocaleString("pt-BR",{maximumFractionDigits:1})+"%"}</td><td>{d.queue===null?"—":d.queue.toLocaleString("pt-BR",{maximumFractionDigits:2})}</td></tr>)}</tbody></table></div>
      <p className="ap-quality-note">A etapa mais lenta é a etapa de trabalho/espera com maior média entre as cinco etapas do Cycle Time; a fila inicial é mostrada separadamente. Percentual agregado de espera pondera as horas/dias de espera pela soma do Lead Time dos itens com dados válidos.</p>
    </section>
    <div className="ap-bd-dora-heading"><div><h2>Indicadores DORA por sprint</h2><p>Deploys, falhas de mudança, recuperação e lead time for changes.</p></div></div>
    <div className="ap-bd-dora-grid">
      {doraChart("Deploys por semana","Deploys na sprint ÷ 2.","deploysPerWeek","#60a5fa","")}
      {doraChart("Change failure rate","Falhas de deploy confirmadas ÷ deploys da sprint.","cfr","#f59e0b","%",failureTarget)}
      {doraChart("Tempo médio de recuperação","Média de Horas para restaurar dos defeitos que causaram falha; sem ponto quando não há falha.","recovery","#c084fc"," h")}
      {doraChart("Lead time for changes","Cycle Time médio dos itens na sprint, em dias corridos.","cycleMean","#34d399"," d")}
    </div>
    <p className="ap-quality-note">Change failure rate só é calculado quando há deploys maiores que zero. Falhas são identificadas por “Causou falha de deploy?” = Sim. O tempo de recuperação fica sem ponto quando a sprint não tem falhas registradas. Meta lida de Premissas (parâmetro relacionado a “meta de falha” ou “change failure”).</p>
  </div>;
}



function decisionPremise(rows:Row[],patterns:RegExp[]):{value:number|null;raw:string;label:string} {
  const row=rows.find(r=>patterns.some(p=>p.test(normalize(r["Parâmetro"]??r["Parametro"]??r["Parâmetro / indicador"]??r["Indicador"]??""))));
  if(!row)return {value:null,raw:"",label:""};
  const raw=str(row["Valor"]??row["valor"]);
  const parsed=premiseValue([row],[/.*/]);
  return {value:parsed,raw,label:str(row["Parâmetro"]??row["Parametro"]??row["Parâmetro / indicador"]??"")};
}
function decisionPercent(value:number|null,raw:string):number|null {
  if(value===null)return null;
  return /%/.test(raw)?value:(Math.abs(value)<=1?value*100:value);
}
function decisionMoney(value:number|null):string {
  return value===null?"—":value.toLocaleString("pt-BR",{style:"currency",currency:"BRL",maximumFractionDigits:2});
}
function decisionDate(value:Date|null):string {
  return value&&!Number.isNaN(value.getTime())?value.toLocaleDateString("pt-BR",{day:"2-digit",month:"2-digit",year:"numeric"}):"—";
}
function addSprintDays(value:Date|null,sprints:number):Date|null {
  if(!value||!Number.isFinite(sprints))return null;
  const d=new Date(value.getFullYear(),value.getMonth(),value.getDate());
  d.setDate(d.getDate()+Math.max(0,sprints)*14);
  return d;
}
function DecisionRisksPage({baseRows,backlog,premises,risks,defects,rework,calendar,doraSprints,source}:{baseRows:Row[];backlog:Row[];premises:Row[];risks:Row[];defects:Row[];rework:Row[];calendar:Row[];doraSprints:Row[];source:"demo"|"imported"}) {
  const teamCostPremise=decisionPremise(premises,[/custo.*time.*sprint/,/custo.*equipe.*sprint/,/custo.*sprint/,/custo.*por sprint/]);
  const reductionPremise=decisionPremise(premises,[/reduc.*reabert/,/reabert.*reduc/]);
  const actionCostPremise=decisionPremise(premises,[/custo.*acao/,/custo.*ação/,/investimento.*qualidade/,/reforco.*qualidade/,/reforço.*qualidade/]);
  const [reductionRaw,setReductionRaw]=useState("");
  const [actionCostRaw,setActionCostRaw]=useState("");
  const reductionDefault=decisionPercent(reductionPremise.value,reductionPremise.raw);
  const actionDefault=actionCostPremise.value;
  useEffect(()=>{setReductionRaw(reductionDefault===null?"":String(Number(reductionDefault.toFixed(2))));},[reductionPremise.raw,reductionPremise.value]);
  useEffect(()=>{setActionCostRaw(actionDefault===null?"":String(actionDefault));},[actionCostPremise.raw,actionCostPremise.value]);
  const reductionInput=Number(reductionRaw.replace(",",".")); 
  const reductionPct=reductionRaw.trim()!==""&&Number.isFinite(reductionInput)?Math.min(100,Math.max(0,reductionInput)):null;
  const actionInput=Number(actionCostRaw.replace(",",".")); 
  const actionCost=actionCostRaw.trim()!==""&&Number.isFinite(actionInput)&&actionInput>=0?actionInput:null;
  const allSprintNames=useMemo(()=>[...new Set(baseRows.map(r=>str(r["Sprint"])).filter(Boolean))].sort((a,b)=>sprintNum(a)-sprintNum(b)||a.localeCompare(b,"pt-BR",{numeric:true})),[baseRows]);
  const sprintFacts=useMemo(()=>allSprintNames.map(sprint=>{
    const items=baseRows.filter(r=>str(r["Sprint"])===sprint&&isCompletedRow(r));
    const net=items.filter(r=>(numeric(r["Reaberturas"])??0)===0).length;
    const reopened=items.filter(r=>(numeric(r["Reaberturas"])??0)>0).length;
    const lead=items.map(r=>numeric(r["Lead time (dias corridos)"])??elapsedDays(r["Criado"]??r["Created Date"],r["Concluído"]??r["Resolution Date"])).filter((v):v is number=>v!==null&&v>=0);
    const cycle=items.map(r=>numeric(r["Cycle time (dias corridos)"])).filter((v):v is number=>v!==null&&v>=0);
    return {sprint,n:sprintNum(sprint),gross:items.length,net,reopened,lead,cycle};
  }),[allSprintNames,baseRows]);
  const avg=(values:number[])=>values.length?values.reduce((a,b)=>a+b,0)/values.length:null;
  const bandMean=(min:number,max:number,key:"net"|"gross"|"reopened")=>avg(sprintFacts.filter(s=>s.n>=min&&s.n<=max).map(s=>s[key]));
  const optimistic=bandMean(4,10,"net"), probable=bandMean(18,20,"net"), pessimistic=bandMean(15,17,"net");
  const grossOptimistic=bandMean(4,10,"gross"),grossProbable=bandMean(18,20,"gross"),grossPessimistic=bandMean(15,17,"gross");
  const lastThree=sprintFacts.filter(s=>s.n>=Math.max(1,(Math.max(0,...sprintFacts.map(x=>x.n))-2))&&s.n<=Math.max(0,...sprintFacts.map(x=>x.n)));
  const reopenedLast3=avg(lastThree.map(s=>s.reopened));
  const backlogCount=backlog.length;
  const couldRows=backlog.filter(r=>normalize(r["Prioridade"])==="could");
  const backlogWithoutCould=Math.max(0,backlogCount-couldRows.length);
  const scenarios=[{name:"Otimista",range:"Sprints 4–10",net:optimistic,gross:grossOptimistic},{name:"Provável",range:"Sprints 18–20",net:probable,gross:grossProbable},{name:"Pessimista",range:"Sprints 15–17",net:pessimistic,gross:grossPessimistic}];
  const lastCalendar=calendar.filter(r=>str(r["Fim"])).sort((a,b)=>sprintNum(str(b["Sprint"]))-sprintNum(str(a["Sprint"])))[0];
  const lastBaseSprint=allSprintNames.reduce((best,s)=>sprintNum(s)>sprintNum(best)?s:best,"");
  const endOfLastSprint=lastCalendar?dateValue(lastCalendar["Fim"]):null;
  const calcOption=(name:string,scope:number,throughput:number|null,action:number|null)=>{
    const sprints=throughput!==null&&throughput>0?Math.ceil(scope/throughput):null;
    const team= sprints!==null&&teamCostPremise.value!==null?sprints*teamCostPremise.value:null;
    const total=team!==null&&action!==null?team+action:null;
    return {name,scope,throughput,sprints,finish:addSprintDays(endOfLastSprint,sprints??0),team,action,total};
  };
  const improvedThroughput=probable!==null&&reductionPct!==null&&reopenedLast3!==null?probable+(reductionPct/100)*reopenedLast3:null;
  const options=[
    calcOption("A · Manter",backlogCount,probable,0),
    calcOption("B · Reforço de qualidade",backlogCount,improvedThroughput,actionCost),
    calcOption("C · Cortar prioridade Could",backlogWithoutCould,probable,0),
    calcOption("D · Reforço + corte",backlogWithoutCould,improvedThroughput,actionCost)
  ];
  const baseline=options[0].total;
  const withSavings=options.map(o=>({...o,savings:baseline!==null&&o.total!==null?baseline-o.total:null}));
  const latest3=sprintFacts.slice(-3);
  const last3SprintNames=latest3.map(s=>s.sprint);
  const defectsInLast3=defects.filter(r=>last3SprintNames.includes(str(r["Sprint de detecção"])));
  const failures=defectsInLast3.filter(r=>/^(sim|yes|true|1)$/i.test(str(r["Causou falha de deploy?"])));
  const deployTotal=latest3.reduce((sum,s)=>sum+(numeric(doraSprints.find(r=>str(r["Sprint"])===s.sprint)?.["Deploys na sprint"])??0),0);
  const escapePerSprint=latest3.map(s=>{
    const ds=defects.filter(r=>str(r["Sprint de detecção"])===s.sprint);
    return ds.length?ds.filter(r=>normalize(r["Origem"])===normalize("Escola-piloto")).length/ds.length*100:null;
  }).filter((v):v is number=>v!==null);
  const cfr=deployTotal>0?failures.length/deployTotal*100:null;
  const restore=failures.map(r=>numeric(r["Horas para restaurar"])).filter((v):v is number=>v!==null&&v>=0);
  const recovery=avg(restore);
  const leadP85Values=latest3.flatMap(s=>s.lead);
  const sortedLeads=[...leadP85Values].sort((a,b)=>a-b);
  const leadP85=sortedLeads.length?sortedLeads[Math.min(sortedLeads.length-1,Math.ceil(.85*sortedLeads.length)-1)]:null;
  const reworkHours=rework.filter(r=>last3SprintNames.includes(str(r["Sprint do item"]))).reduce((sum,r)=>sum+(numeric(r["Horas de retrabalho"])??0),0);
  const correctionHours=defects.filter(r=>last3SprintNames.includes(str(r["Sprint de detecção"]))).reduce((sum,r)=>sum+(numeric(r["Horas de correção"])??0),0);
  const capacity=calendar.filter(r=>last3SprintNames.includes(str(r["Sprint"]))).reduce((sum,r)=>sum+(numeric(r["Capacidade (h)"])??0),0);
  const reworkPct=capacity>0?(reworkHours+correctionHours)/capacity*100:null;
  const currentValues:Record<string,number|null>={
    escape:avg(escapePerSprint),
    cfr,
    rework:reworkPct,
    recovery,
    lead:leadP85,
    throughput:avg(latest3.map(s=>s.net))
  };
  const metricAliases:Record<string,string[]>={
    escape:["escape","taxa de escape","escape rate"],
    cfr:["change failure","falha de mudanca","taxa de falha","cfr"],
    rework:["retrabalho","rework"],
    recovery:["recuperacao","recovery","restauracao"],
    lead:["lead time p85","p85","lead time"],
    throughput:["throughput liquido","throughput net","throughput"]
  };
  const currentForRisk=(row:Row):{value:number|null;metric:string;unit:string}=>{
    const metricField=Object.keys(row).find(k=>/metrica|indicador|kpi/.test(normalize(k)));
    const triggerField=Object.keys(row).find(k=>/gatilho|condicao|condição/.test(normalize(k)));
    const text=normalize([metricField?row[metricField]:"",triggerField?row[triggerField]:"",...Object.values(row)].join(" "));
    const key=Object.keys(metricAliases).find(k=>metricAliases[k].some(alias=>text.includes(normalize(alias))));
    if(!key)return {value:null,metric:"",unit:""};
    const unit=key==="escape"||key==="cfr"||key==="rework"?"%":key==="recovery"?" h":key==="lead"?" dias":" itens";
    return {value:currentValues[key],metric:key,unit};
  };
  const riskDetails=risks.map((row,index)=>{
    const keys=Object.keys(row);
    const current=currentForRisk(row);
    const triggerKey=keys.find(k=>/gatilho|limite|threshold|valor.*referencia|valor.*referência/.test(normalize(k)));
    const triggerText=str(triggerKey?row[triggerKey]:"");
    const triggerNum=numeric(triggerText.replace(/%/g,"").match(/-?\d+(?:[.,]\d+)?/)?.[0]);
    const senseKey=keys.find(k=>/sentido|comparador|operador|direcao|direção/.test(normalize(k)));
    const senseText=normalize(senseKey?row[senseKey]:"")+" "+normalize(triggerText);
    const greater=/(>|maior|acima|superior|excede|exceder)/.test(senseText);
    const less=/(<|menor|abaixo|inferior)/.test(senseText);
    const fired=current.value!==null&&triggerNum!==null&&(greater?current.value>triggerNum:less?current.value<triggerNum:false);
    const status=current.value===null||triggerNum===null?"SEM DADOS":fired?"DISPARADO":"OK";
    return {row,index,current,triggerKey,triggerText,triggerNum,status,sense:greater?">":less?"<":"—"};
  });
  const moneySavings=(v:number|null)=>v===null?"—":(v>0?"+":"")+decisionMoney(v);
  return <div className="ap-decision-page">
    {source==="demo"&&<section className="ap-panel ap-decision-note"><strong>Importe a planilha para calcular forecast e riscos reais.</strong><p>Os cálculos usam as abas Backlog, Premissas, Riscos_Gatilhos, Base_Itens, Defeitos, Retrabalho, Calendario_Eventos e DORA_Sprint. Não há estimativas inventadas quando faltam dados.</p></section>}
    <section className="ap-panel ap-decision-section">
      <div className="ap-panel-heading"><div><h2>Forecast do backlog</h2><p>Backlog restante = quantidade de linhas da aba Backlog. Throughput líquido conta itens concluídos sem reabertura.</p></div><span className="ap-decision-pill">{backlogCount} itens restantes</span></div>
      <div className="ap-decision-scenario-grid">{scenarios.map(sc=>{
        const netSprints=sc.net!==null&&sc.net>0?Math.ceil(backlogCount/sc.net):null;
        const grossSprints=sc.gross!==null&&sc.gross>0?Math.ceil(backlogCount/sc.gross):null;
        return <div className="ap-decision-scenario" key={sc.name}><span>{sc.name.toUpperCase()}</span><strong>{netSprints===null?"—":netSprints+" sprints"}</strong><small>{sc.range} · throughput líquido {sc.net===null?"—":sc.net.toFixed(2)} itens/sprint</small><div>Referência bruta: {grossSprints===null?"—":grossSprints+" sprints"} ({sc.gross===null?"—":sc.gross.toFixed(2)} itens/sprint)</div><div className="ap-decision-date">Conclusão estimada: {netSprints===null?"—":decisionDate(addSprintDays(endOfLastSprint,netSprints))}</div></div>;
      })}</div>
      <p className="ap-quality-note">Datas calculadas a partir do fim da última sprint registrada em Calendario_Eventos, adicionando 14 dias por sprint restante. {endOfLastSprint?("Última sprint com data de fim: "+str(lastCalendar?.["Sprint"])+", "+decisionDate(endOfLastSprint)+"."):"Informe a coluna Fim em Calendario_Eventos para calcular datas."}</p>
    </section>
    <section className="ap-panel ap-decision-section">
      <div className="ap-panel-heading"><div><h2>Opções de decisão</h2><p>Compare escopo, prazo e custo. As opções B e D usam a melhoria estimada com redução de reaberturas.</p></div></div>
      <div className="ap-decision-edit-grid">
        <label><span>Redução de reaberturas (%)</span><input type="number" min="0" max="100" step="1" value={reductionRaw} onChange={e=>setReductionRaw(e.target.value)} placeholder="Ex.: 25"/><small>Premissa: {reductionPremise.label||"redução de reaberturas"} · média de itens reabertos nas últimas 3 sprints: {reopenedLast3===null?"—":reopenedLast3.toFixed(2)}</small></label>
        <label><span>Custo da ação de qualidade (R$)</span><input type="number" min="0" step="0.01" value={actionCostRaw} onChange={e=>setActionCostRaw(e.target.value)} placeholder="Ex.: 1500"/><small>Premissa: {actionCostPremise.label||"custo da ação"} · editável para simulação</small></label>
      </div>
      <div className="ap-decision-meta">Custo do time por sprint: <strong>{decisionMoney(teamCostPremise.value)}</strong> · throughput provável líquido: <strong>{probable===null?"—":probable.toFixed(2)}</strong> · throughput B/D: <strong>{improvedThroughput===null?"—":improvedThroughput.toFixed(2)}</strong> itens/sprint</div>
      <div className="ap-table-scroll ap-decision-table"><table><thead><tr><th>Opção</th><th>Escopo</th><th>Throughput líquido</th><th>Sprints</th><th>Conclusão estimada</th><th>Custo do time</th><th>Custo da ação</th><th>Custo total</th><th>Economia vs. A</th></tr></thead><tbody>{withSavings.map(o=><tr key={o.name}><td><strong>{o.name}</strong></td><td>{o.scope} itens</td><td>{o.throughput===null?"—":o.throughput.toFixed(2)}</td><td>{o.sprints===null?"—":o.sprints}</td><td>{o.sprints===null?"—":decisionDate(addSprintDays(endOfLastSprint,o.sprints))}</td><td>{o.team===null?"—":decisionMoney(o.team)}</td><td>{o.action===null?"—":decisionMoney(o.action)}</td><td><strong>{o.total===null?"—":decisionMoney(o.total)}</strong></td><td className={o.savings!==null&&o.savings>0?"ap-decision-positive":o.savings!==null&&o.savings<0?"ap-decision-negative":""}>{moneySavings(o.savings)}</td></tr>)}</tbody></table></div>
      <p className="ap-quality-note">Custo total = (sprints arredondadas para cima × custo do time por sprint) + custo da ação. A economia compara o custo total da opção com A. Valores não calculados aparecem como “—” quando falta uma premissa ou throughput válido.</p>
    </section>
    <section className="ap-panel ap-decision-section">
      <div className="ap-panel-heading"><div><h2>Riscos e gatilhos</h2><p>Valor atual calculado com as últimas 3 sprints disponíveis. Status comparado ao sentido e ao limite cadastrados na aba Riscos_Gatilhos.</p></div><span className="ap-decision-pill">{risks.length} riscos</span></div>
      {!risks.length?<div className="ap-empty">A aba Riscos_Gatilhos não foi encontrada ou não contém linhas. Importe a aba para avaliar os gatilhos.</div>:<div className="ap-table-scroll ap-decision-table"><table><thead><tr>{Object.keys(risks[0]).map(k=><th key={k}>{k}</th>)}<th>Valor atual (últimas 3 sprints)</th><th>Comparação</th><th>Status</th></tr></thead><tbody>{riskDetails.map(({row,index,current,triggerText,triggerNum,status,sense})=><tr key={index}>{Object.keys(risks[0]).map(k=><td key={k}>{str(row[k])||"—"}</td>)}<td>{current.value===null?"—":current.value.toLocaleString("pt-BR",{maximumFractionDigits:2})+current.unit}</td><td>{sense!=="—"&&triggerNum!==null?sense+" "+triggerNum+(/%/.test(triggerText)?"%":""):"—"}</td><td><span className={"ap-risk-status "+(status==="DISPARADO"?"is-fired":status==="OK"?"is-ok":"is-missing")}>{status}</span></td></tr>)}</tbody></table></div>}
      <p className="ap-quality-note">Indicadores disponíveis: taxa de escape, change failure rate, retrabalho (% da capacidade), tempo de recuperação, Lead Time P85 e throughput líquido. Se um gatilho não identificar indicador, limite ou sentido, o status fica “SEM DADOS” em vez de presumir uma regra.</p>
    </section>
  </div>;
}



function avgNumbers(values:number[]):number|null { return values.length?values.reduce((a,b)=>a+b,0)/values.length:null; }
const SUMMARY_PHASES=[{label:"S1–S3",min:1,max:3},{label:"S4–S10",min:4,max:10},{label:"S11–S17",min:11,max:17},{label:"S18–S20",min:18,max:20}];
function summaryFmt(v:number|null,unit:string):string { if(v===null||!Number.isFinite(v))return "—";const n=v.toLocaleString("pt-BR",{maximumFractionDigits:2});return unit==="%"?n+"%":unit==="dias"?n+" dias":unit==="h"?n+" h":n; }
function ExecutiveSummaryPage({baseRows,defects,rework,calendar,doraSprints,premises,risks,backlog,source}:{baseRows:Row[];defects:Row[];rework:Row[];calendar:Row[];doraSprints:Row[];premises:Row[];risks:Row[];backlog:Row[];source:"demo"|"imported"}) {
 const sprints=[...new Set([...baseRows.map(r=>str(r["Sprint"])),...defects.map(r=>str(r["Sprint de detecção"])),...rework.map(r=>str(r["Sprint do item"])),...doraSprints.map(r=>str(r["Sprint"]))].filter(Boolean))].sort((a,b)=>sprintNum(a)-sprintNum(b));
 const data=sprints.map(sprint=>{
  const items=baseRows.filter(r=>str(r["Sprint"])===sprint&&isCompletedRow(r)), gross=items.length, net=items.filter(r=>(numeric(r["Reaberturas"])??0)===0).length;
  const lead=items.map(r=>numeric(r["Lead time (dias corridos)"])??elapsedDays(r["Criado"]??r["Created Date"],r["Concluído"]??r["Resolution Date"])).filter((v):v is number=>v!==null&&v>=0);
  const ds=defects.filter(r=>str(r["Sprint de detecção"])===sprint), school=ds.filter(r=>normalize(r["Origem"])===normalize("Escola-piloto")).length;
  const rw=rework.filter(r=>str(r["Sprint do item"])===sprint), cap=numeric(calendar.find(r=>str(r["Sprint"])===sprint)?.["Capacidade (h)"]);
  const hours=rw.reduce((a,r)=>a+(numeric(r["Horas de retrabalho"])??0),0)+ds.reduce((a,r)=>a+(numeric(r["Horas de correção"])??0),0);
  const deploy=numeric(doraSprints.find(r=>str(r["Sprint"])===sprint)?.["Deploys na sprint"]), failed=ds.filter(r=>/^(sim|yes|true|1)$/i.test(str(r["Causou falha de deploy?"])));
  const recovery=failed.map(r=>numeric(r["Horas para restaurar"])).filter((v):v is number=>v!==null&&v>=0);
  return {sprint,n:sprintNum(sprint),gross,net,lead:avgNumbers(lead),p85:inclusiveP85(lead),defect: gross?ds.length/gross*100:null,escape:ds.length?school/ds.length*100:null,rework:cap!==null&&cap>0?hours/cap*100:null,cfr:deploy!==null&&deploy>0?failed.length/deploy*100:null,recovery:avgNumbers(recovery),leadValues:lead,recoveryValues:recovery,defects:ds.length,school,deploy,failures:failed.length};
 });
 const metrics=[{key:"gross",label:"Throughput bruto",unit:"itens"},{key:"net",label:"Throughput líquido",unit:"itens"},{key:"lead",label:"Lead Time médio",unit:"dias"},{key:"p85",label:"Lead Time P85",unit:"dias"},{key:"defect",label:"Taxa de defeitos",unit:"%"},{key:"escape",label:"Taxa de escape",unit:"%"},{key:"rework",label:"Retrabalho / capacidade",unit:"%"},{key:"cfr",label:"Change failure rate",unit:"%"},{key:"recovery",label:"Tempo de recuperação",unit:"h"}];
 const phases=SUMMARY_PHASES.map(ph=>{
  const ss=data.filter(s=>s.n>=ph.min&&s.n<=ph.max), total=(k:"gross"|"net")=>ss.reduce((a,s)=>a+s[k],0);
  const mean=(k:string)=>avgNumbers(ss.map(s=>(s as any)[k] as number|null).filter((v):v is number=>v!==null&&Number.isFinite(v)));
  const totalGross=total("gross"),defectsTotal=ss.reduce((a,s)=>a+s.defects,0),school=ss.reduce((a,s)=>a+s.school,0),deploys=ss.reduce((a,s)=>a+(s.deploy||0),0),failures=ss.reduce((a,s)=>a+s.failures,0);
  const recoveryValues=ss.flatMap(s=>s.recoveryValues as number[]);
  const phaseLeadValues=ss.flatMap(s=>s.leadValues as number[]);
  return {label:ph.label,gross:total("gross"),net:total("net"),lead:avgNumbers(phaseLeadValues),p85:inclusiveP85(phaseLeadValues),defect:totalGross?defectsTotal/totalGross*100:null,escape:defectsTotal?school/defectsTotal*100:null,rework:mean("rework"),cfr:deploys?failures/deploys*100:null,recovery:avgNumbers(recoveryValues)};
 });
 const prev=phases[2],last=phases[3],meanRange=(min:number,max:number,key:"net"|"gross")=>{const ss=data.filter(s=>s.n>=min&&s.n<=max);return ss.length?avgNumbers(ss.map(s=>s[key])):null;};
 const probable=meanRange(18,20,"net"),optimistic=meanRange(4,10,"net"),pessimistic=meanRange(15,17,"net"),grossProbable=meanRange(18,20,"gross"),remaining=backlog.length;
 const sprintsLeft=probable!==null&&probable>0?Math.ceil(remaining/probable):null;
 const team=decisionPremise(premises,[/custo.*time.*sprint/,/custo.*equipe.*sprint/,/custo.*sprint/,/custo.*por sprint/]).value;
 const action=decisionPremise(premises,[/custo.*acao/,/custo.*ação/,/investimento.*qualidade/,/reforco.*qualidade/,/reforço.*qualidade/]).value;
 const redPrem=decisionPremise(premises,[/reduc.*reabert/,/reabert.*reduc/]),red=decisionPercent(redPrem.value,redPrem.raw);
 const latest=data.slice(-3),reopened=latest.length?avgNumbers(latest.map(s=>s.gross-s.net)):null,better=probable!==null&&red!==null&&reopened!==null?probable+red/100*reopened:null;
 const could=backlog.filter(r=>normalize(r["Prioridade"])==="could").length;
 const costA=probable&&team!==null?Math.ceil(remaining/probable)*team:null;
 const options=[{name:"A · Manter",cost:costA},{name:"B · Reforço de qualidade",cost:better&&team!==null&&action!==null?Math.ceil(remaining/better)*team+action:null},{name:"C · Cortar Could",cost:probable&&team!==null?Math.ceil(Math.max(0,remaining-could)/probable)*team:null},{name:"D · Reforço + corte",cost:better&&team!==null&&action!==null?Math.ceil(Math.max(0,remaining-could)/better)*team+action:null}];
 const viable=options.filter(o=>o.cost!==null) as {name:string;cost:number}[],best=viable.length?viable.reduce((a,b)=>b.cost<a.cost?b:a):null,savings=best&&costA!==null?costA-best.cost:null;
 const fired=risks.filter(r=>{const text=normalize(Object.values(r).join(" ")),raw=str(r[Object.keys(r).find(k=>/gatilho|limite|threshold/.test(normalize(k)))||""]),limit=numeric(raw.match(/-?\d+(?:[.,]\d+)?/)?.[0]);const key=text.includes("change failure")||text.includes("taxa de falha")?"cfr":text.includes("escape")?"escape":text.includes("retrabalho")?"rework":text.includes("recuper")?"recovery":text.includes("lead time")?"p85":text.includes("throughput")?"net":"";const current=key?(last as any)[key]:null;const gt=/(>|maior|acima|superior|excede)/.test(normalize(raw)+" "+text),lt=/(<|menor|abaixo|inferior)/.test(normalize(raw)+" "+text);return current!==null&&current!==undefined&&limit!==null&&(gt?current>limit:lt?current<limit:false);}).length;
 const statements=[
  "Forecast provável: "+(sprintsLeft===null?"não calculável com os dados atuais":sprintsLeft+" sprint(s) para "+remaining+" itens, com throughput líquido de "+probable!.toFixed(2)+" itens/sprint")+". Otimista: "+(optimistic===null?"sem dados":optimistic.toFixed(2)+" itens/sprint")+"; pessimista: "+(pessimistic===null?"sem dados":pessimistic.toFixed(2)+" itens/sprint")+". Referência bruta provável: "+(grossProbable===null?"sem dados":grossProbable.toFixed(2)+" itens/sprint")+".",
  "Opção recomendada pelo menor custo calculável: "+(best?.name||"não determinada")+". "+(savings===null?"Economia versus A não calculável com as premissas disponíveis.":"Economia estimada versus A: "+decisionMoney(savings)+"."),
  "Riscos disparados: "+fired+" de "+risks.length+" gatilhos identificados por indicador, sentido e limite; regras incompletas não são contadas como disparadas."
 ];
 return <div className="ap-exec-page">{source==="demo"&&<section className="ap-panel ap-decision-note"><strong>Resumo demonstrativo</strong><p>Importe a planilha para calcular os indicadores com dados reais.</p></section>}
 <section className="ap-panel ap-exec-section"><div className="ap-panel-heading"><div><h2>Indicadores por fase</h2><p>Variação da última fase (S18–S20) contra a anterior (S11–S17).</p></div></div><div className="ap-table-scroll ap-decision-table"><table><thead><tr><th>Indicador</th>{phases.map(p=><th key={p.label}>{p.label}</th>)}<th>Variação última fase</th></tr></thead><tbody>{metrics.map(m=>{const a=(prev as any)[m.key] as number|null,b=(last as any)[m.key] as number|null,d=a!==null&&b!==null?b-a:null,p=a!==null&&b!==null&&a!==0?(b-a)/Math.abs(a)*100:null;return <tr key={m.key}><td><strong>{m.label}</strong></td>{phases.map(ph=><td key={ph.label}>{summaryFmt((ph as any)[m.key] as number|null,m.unit)}</td>)}<td>{d===null?"—":(d>0?"+":"")+summaryFmt(d,m.unit)+(p===null?"":" ("+(p>0?"+":"")+p.toLocaleString("pt-BR",{maximumFractionDigits:1})+"%)")}</td></tr>;})}</tbody></table></div><p className="ap-quality-note">Bruto = itens concluídos; líquido = itens concluídos sem reabertura. Defeitos / throughput; escape = Escola-piloto / defeitos; retrabalho = horas de retrabalho + correção / capacidade; change failure = falhas de deploy / deploys.</p></section>
 <section className="ap-panel ap-exec-section"><div className="ap-panel-heading"><div><h2>Leitura executiva</h2><p>Três frases geradas a partir dos números e premissas disponíveis.</p></div></div><div className="ap-exec-sentences">{statements.map((line,i)=><div key={i}><span>{i+1}</span><p>{line}</p></div>)}</div></section></div>;
}
function CommunicationPage({stakeholders,metricDictionary,source}:{stakeholders:Row[];metricDictionary:Row[];source:"demo"|"imported"}) {
 const rows=stakeholders.map((r,i)=>{const entries=Object.entries(r),pick=(ps:RegExp[])=>{const x=entries.find(([k])=>ps.some(p=>p.test(normalize(k))));return x?str(x[1]):"";};
 const name=pick([/nome/,/stakeholder/,/parte interessada/,/^name$/])||"Stakeholder "+(i+1),power=pick([/poder/,/power/]),interest=pick([/interesse/,/interest/]);
 const score=(v:string)=>{const n=numeric(v);if(n!==null)return n;const x=normalize(v);return /muito alto|alto|high|elevado/.test(x)?3:/medio|moderado|moderate/.test(x)?2:/baixo|low|reduzido/.test(x)?1:null;};
 const p=score(power),it=score(interest),approach=p===null||it===null?"Não calculável":p>=2&&it>=2?"Gerenciar de perto":p>=2?"Manter satisfeito":it>=2?"Manter informado":"Monitorar";
 return {name,power:power||"—",interest:interest||"—",approach,objection:pick([/obje[cç][aã]o/,/resist[eê]ncia/,/objection/])||"Não informada",question:pick([/pergunta/,/question/,/d[uú]vida provável/])||"Não informada"};});
 return <div className="ap-communication-page">{source==="demo"&&<section className="ap-panel ap-decision-note"><strong>Comunicação e métricas</strong><p>Importe Stakeholders e Dicionario_Metricas para exibir as informações reais.</p></section>}
 <section className="ap-panel ap-exec-section"><div className="ap-panel-heading"><div><h2>Stakeholders</h2><p>Abordagem calculada pela matriz poder × interesse.</p></div><span className="ap-decision-pill">{stakeholders.length} stakeholders</span></div>
 {!rows.length?<div className="ap-empty">A aba Stakeholders não contém linhas disponíveis.</div>:<div className="ap-table-scroll ap-decision-table"><table><thead><tr><th>Nome</th><th>Poder</th><th>Interesse</th><th>Abordagem calculada</th><th>Objeção provável</th><th>Pergunta que fará</th></tr></thead><tbody>{rows.map((r,i)=><tr key={i}><td><strong>{r.name}</strong></td><td>{r.power}</td><td>{r.interest}</td><td><span className="ap-decision-pill">{r.approach}</span></td><td>{r.objection}</td><td>{r.question}</td></tr>)}</tbody></table></div>}</section>
 <section className="ap-panel ap-exec-section"><div className="ap-panel-heading"><div><h2>Dicionário de métricas</h2><p>Definições e regras da aba Dicionario_Metricas.</p></div><span className="ap-decision-pill">{metricDictionary.length} métricas</span></div>
 {!metricDictionary.length?<div className="ap-empty">A aba Dicionario_Metricas não contém linhas disponíveis.</div>:<div className="ap-table-scroll ap-decision-table"><table><thead><tr>{Object.keys(metricDictionary[0]).map(k=><th key={k}>{k}</th>)}</tr></thead><tbody>{metricDictionary.map((r,i)=><tr key={i}>{Object.keys(metricDictionary[0]).map(k=><td key={k}>{str(r[k])||"—"}</td>)}</tr>)}</tbody></table></div>}</section></div>;
}

function AgilePulse() {
  const [rows,setRows]=useState<Row[]>(demo), [source,setSource]=useState<"demo"|"imported">("demo"), [fileName,setFileName]=useState("");
  const [sheetsData,setSheetsData]=useState<Record<string,Row[]>>({});
  const [sheetNames,setSheetNames]=useState<string[]>([]), [snapshotDate,setSnapshotDate]=useState("");
  const [activePage,setActivePage]=useState("Dashboard");
  const [mapping,setMapping]=useState<MapCols>(guessMap(demo)), [statusText,setStatusText]=useState("Done, Concluído, Resolvido, Closed, Complete, Feito, Testado, Aprovado para Produção");
  const [sourceOpen,setSourceOpen]=useState(false), [editorOpen,setEditorOpen]=useState(false);
  const [uploadError,setUploadError]=useState(""), [uploadNotice,setUploadNotice]=useState("");
  const [filter,setFilter]=useState("Todas"); const fileRef=useRef<HTMLInputElement>(null);
  const [storageReady,setStorageReady]=useState(false);
  const headers=Object.keys(rows[0]||{}), statuses=statusText.split(",").map(s=>s.trim()).filter(Boolean);
  const stats=useMemo(()=>calculate(rows,mapping,statuses),[rows,mapping,statusText]);
  const calendarRows=sheetsData["Calendario_Eventos"]||[];
  const flowSeries=useMemo(()=>buildFlowSeries(rows,calendarRows),[rows,calendarRows]);
  const chartRows=filter==="Todas"?flowSeries:flowSeries.filter(s=>s.sprint===filter);
  const importedItems=rows.length;
  const importedSprints=new Set(rows.map(r=>str(r["Sprint"])).filter(Boolean)).size;
  const missingOptional=OPTIONAL_SHEET_NAMES.filter(name=>!sheetNames.includes(name));
  const qualityAlerts=useMemo(()=>{
    if(source!=="imported")return [] as string[];
    const alerts:string[]=[];
    const baseMissing=BASE_COLUMNS.filter(c=>!Object.prototype.hasOwnProperty.call(rows[0]||{},c));
    if(baseMissing.length)alerts.push("Colunas obrigatórias ausentes em Base_Itens: "+baseMissing.join(", "));
    for(const [name,required] of Object.entries(OPTIONAL_COLUMNS)){if(required.length&&sheetsData[name]){const missing=required.filter(c=>!Object.prototype.hasOwnProperty.call(sheetsData[name][0]||{},c));if(missing.length)alerts.push("Colunas obrigatórias ausentes em "+name+": "+missing.join(", "));}}
    const ids=rows.map(r=>str(r["Issue key"])).filter(Boolean);
    const dup=[...new Set(ids.filter((id,i)=>ids.indexOf(id)!==i))];
    if(dup.length)alerts.push("Issue key duplicada: "+dup.length+" chave(s) — "+dup.slice(0,8).join(", "));
    const stages=["In development","Waiting for QA","In QA review","In code review","Waiting for review"];
    const calendar=sheetsData["Calendario_Eventos"]||[];
    rows.forEach((r,i)=>{
      const key=str(r["Issue key"]||"linha "+(i+1));
      const lead=Number(str(r["Lead time (dias corridos)"]).replace(",","."));
      const cycle=Number(str(r["Cycle time (dias corridos)"]).replace(",","."));
      if(str(r["Lead time (dias corridos)"])&&str(r["Cycle time (dias corridos)"])&&Number.isFinite(lead)&&Number.isFinite(cycle)&&lead<cycle)alerts.push("Lead time menor que cycle time: "+key);
      if(str(r["Cycle time (dias corridos)"])){
        const sum=stages.reduce((total,c)=>total+(Number(str(r[c]).replace(",", "."))||0),0);
        if(Number.isFinite(cycle)&&Math.abs(sum-cycle)>0.01)alerts.push("Soma das 5 etapas difere do cycle time: "+key);
      }
      const sprint=str(r["Sprint"]), completed=dateValue(r["Concluído"]);
      const event=calendar.find(c=>str(c["Sprint"])===sprint);
      const begin=event?dateValue(event["Início"]):null, end=event?dateValue(event["Fim"]):null;
      if(completed&&begin&&end&&(completed<begin||completed>end))alerts.push("Data de conclusão fora da sprint: "+key);
    });
    return [...new Set(alerts)];
  },[rows,sheetsData,source]);
  useEffect(()=>{
    try{
      const saved=localStorage.getItem(STORE_KEY);
      if(saved){
        const data=JSON.parse(saved);
        if(data&&data.source==="imported"&&Array.isArray(data.rows)){
          setRows(data.rows);setSource("imported");setFileName(data.fileName||"Planilha salva");setSheetsData(data.sheetsData||{});setSheetNames(data.sheetNames||[]);setSnapshotDate(data.snapshotDate||"");setMapping(data.mapping||guessMap(data.rows));setUploadNotice("Dados importados restaurados do armazenamento local do navegador.");
        }
      }
    }catch{}
    setStorageReady(true);
  },[]);
  useEffect(()=>{
    if(!storageReady)return;
    try{
      if(source==="imported")localStorage.setItem(STORE_KEY,JSON.stringify({source,rows,fileName,sheetsData,sheetNames,snapshotDate,mapping}));
      else localStorage.removeItem(STORE_KEY);
    }catch{setUploadError("Não foi possível salvar no navegador. O armazenamento local pode estar cheio ou desabilitado.");}
  },[storageReady,source,rows,fileName,sheetsData,sheetNames,snapshotDate,mapping]);
  async function parseFile(file: File) {
    setUploadError("");setUploadNotice("");
    try{
      if(!/\.xlsx$/i.test(file.name))throw new Error("Envie um arquivo Excel .xlsx.");
      const wb=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:true});
      if(!wb.SheetNames.includes("Base_Itens"))throw new Error("A aba obrigatória Base_Itens não foi encontrada.");
      const importedNames=wb.SheetNames.filter(n=>!IGNORED_SHEETS.has(n)&&(n==="Base_Itens"||OPTIONAL_SHEET_NAMES.includes(n)));
      const baseRows=parseSheetRows(wb.Sheets["Base_Itens"],2,3);
      if(!baseRows.length)throw new Error("A aba Base_Itens não contém linhas de dados a partir da linha 4.");
      const missing=BASE_COLUMNS.filter(c=>!Object.keys(baseRows[0]).includes(c));
      const normalizedRows=baseRows;
      const dataBySheet:Record<string,Row[]>={"Base_Itens":normalizedRows};
      let snapshot="";
      for(const name of importedNames){
        if(name==="Base_Itens")continue;
        const ws=wb.Sheets[name];if(!ws)continue;
        if(name==="WIP_Aberto"){
          const grid=XLSX.utils.sheet_to_json<unknown[]>(ws,{header:1,defval:"",raw:true,blankrows:false}) as unknown[][];
          snapshot=str(excelCell(grid[2]?.[1],"Data do snapshot"));
          dataBySheet[name]=parseSheetRows(ws,5,6,20);
        }else if(name==="Premissas"){
          dataBySheet[name]=parseSheetRows(ws,2,3,15);
        }else{
          let parsed=parseSheetRows(ws,2,3);
          if(name==="DORA_Sprint")parsed=parsed.map(r=>({"Sprint":r["Sprint"]??"","Deploys na sprint":r["Deploys na sprint"]??""}));
          dataBySheet[name]=parsed;
        }
      }
      const map:MapCols={sprint:"Sprint",item:"Issue key",status:"",squad:"Papel",created:"Criado",started:"",completed:"Concluído",storyPoints:"Story Points"};
      setRows(normalizedRows);setMapping(map);setSheetsData(dataBySheet);setSheetNames(importedNames);setSnapshotDate(snapshot);setFileName(file.name);setSource("imported");setActivePage("Dashboard");setEditorOpen(false);setFilter("Todas");
      const optionalMissing=OPTIONAL_SHEET_NAMES.filter(n=>!importedNames.includes(n));
      setUploadNotice("Importação concluída: "+normalizedRows.length+" itens, "+new Set(normalizedRows.map(r=>str(r["Sprint"])).filter(Boolean)).size+" sprints e "+Object.keys(dataBySheet).length+" abas de dados. "+(missing.length?"Colunas obrigatórias ausentes: "+missing.join(", ")+". ":"")+(optionalMissing.length?"Abas opcionais ausentes: "+optionalMissing.join(", ")+".":""));
      setSourceOpen(true);
    }catch(e){setUploadError(e instanceof Error?e.message:"Não foi possível ler a planilha.");}
  }
  function updateCell(rowIndex:number, column:string, value:string) {
    setRows(current=>current.map((row,index)=>index===rowIndex?({...row,[column]:value}):row));
    setSheetsData(current=>({...current,Base_Itens:(current.Base_Itens||[]).map((row,index)=>index===rowIndex?({...row,[column]:value}):row)}));
  }
  return <main className="ap-shell">
    <aside className="ap-sidebar"><div className="ap-nav-label">NAVEGAÇÃO</div><button className={"ap-nav "+(activePage==="Dashboard"?"active":"")} onClick={()=>setActivePage("Dashboard")}><Activity size={18}/> Dashboard</button><button className={"ap-nav "+(activePage==="WIP e Aging"?"active":"")} onClick={()=>setActivePage("WIP e Aging")}><Activity size={18}/> WIP e Aging</button><button className={"ap-nav "+(activePage==="Qualidade e Retrabalho"?"active":"")} onClick={()=>setActivePage("Qualidade e Retrabalho")}><Activity size={18}/> Qualidade e Retrabalho</button><button className={"ap-nav "+(activePage==="Gargalo e DORA"?"active":"")} onClick={()=>setActivePage("Gargalo e DORA")}><Activity size={18}/> Gargalo e DORA</button><button className={"ap-nav "+(activePage==="Decisão e Riscos"?"active":"")} onClick={()=>setActivePage("Decisão e Riscos")}><Activity size={18}/> Decisão e Riscos</button><button className={"ap-nav "+(activePage==="Resumo Executivo"?"active":"")} onClick={()=>setActivePage("Resumo Executivo")}><Activity size={18}/> Resumo Executivo</button><button className={"ap-nav "+(activePage==="Comunicação"?"active":"")} onClick={()=>setActivePage("Comunicação")}><MessageCircle size={18}/> Comunicação</button><button className={"ap-nav "+(sourceOpen?"active":"") onClick={()=>setSourceOpen(v=>!v)}><Database size={18}/> Fontes de Dados <ChevronDown className="ap-nav-end" size={15}/></button>{source==="imported"&&<><div className="ap-nav-label">ABAS DA PLANILHA</div>{["Base_Itens",...OPTIONAL_SHEET_NAMES].map(name=><button key={name} className={"ap-nav ap-subnav "+(activePage===name?"active":"")} onClick={()=>setActivePage(name)}>{name}{!sheetNames.includes(name)&&<span className="ap-missing-dot">!</span>}</button>)}</>}<div className="ap-sidebar-bottom"><i/> <div><strong>Fonte de dados</strong><span>{source==="demo"?"Dados demonstrativos":"Arquivo importado"}</span></div></div></aside>
    <section className="ap-main">
    <div className="ap-content"><div className="ap-page-heading"><div><div className="ap-eyebrow">{activePage==="Resumo Executivo"?"VISÃO EXECUTIVA":activePage==="Comunicação"?"COMUNICAÇÃO E ALINHAMENTO":activePage==="Decisão e Riscos"?"PLANEJAMENTO E GOVERNANÇA":activePage==="Gargalo e DORA"?"FLUXO E CONFIABILIDADE":activePage==="Qualidade e Retrabalho"?"QUALIDADE E CONFIABILIDADE":activePage==="WIP e Aging"?"FLUXO DE TRABALHO":"MÉTRICAS DE FLUXO"}</div><h1>{activePage==="Resumo Executivo"?"Resumo Executivo":activePage==="Comunicação"?"Comunicação":activePage==="Decisão e Riscos"?"Decisão e Riscos":activePage==="Gargalo e DORA"?"Gargalo e DORA":activePage==="Qualidade e Retrabalho"?"Qualidade e Retrabalho":activePage==="WIP e Aging"?"WIP e Aging":"Throughput"} <span className="ap-live-pill"><i/>{source==="demo"?"Demonstração":"Dados importados"}</span></h1><p>{activePage==="Resumo Executivo"?"Acompanhe resultados por fase, variações e decisões recomendadas.":activePage==="Comunicação"?"Organize o engajamento dos stakeholders e as definições das métricas.":activePage==="Decisão e Riscos"?"Compare cenários de prazo e custo e acompanhe gatilhos de risco.":activePage==="Gargalo e DORA"?"Identifique filas, etapas lentas e estabilidade das entregas.":activePage==="Qualidade e Retrabalho"?"Acompanhe defeitos, escape, esforço de retrabalho e custo da qualidade.":activePage==="WIP e Aging"?"Monitore itens em andamento, envelhecimento e limite de alerta.":"Acompanhe o ritmo de entrega e explore os itens concluídos por sprint."}</p></div><button className="ap-button ap-secondary" onClick={()=>setSourceOpen(v=>!v)}><FileUp size={17}/> Importar dados</button></div>
    {sourceOpen&&<section className="ap-panel ap-import"><div className="ap-panel-heading"><div><h2>Fontes de Dados</h2><p>Excel .xlsx · leitura das linhas e abas conforme o modelo da planilha.</p></div><button className="ap-icon-button" onClick={()=>{setSourceOpen(false);}}><X size={18}/></button></div>
    <div className="ap-import-choice"><button className="ap-button ap-secondary" onClick={()=>{setRows(demo);setMapping(guessMap(demo));setSource("demo");setFileName("");setSheetsData({});setSheetNames([]);setSnapshotDate("");setActivePage("Dashboard");setUploadError("");setUploadNotice("Dados demonstrativos ativados.");setSourceOpen(false);}}>Usar dados demonstrativos</button><span>ou importe seu arquivo Excel</span></div><div className="ap-dropzone" onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();const f=e.dataTransfer.files?.[0];if(f)void parseFile(f);}}><FileUp size={26}/><strong>Selecione sua planilha</strong><span>Arraste o arquivo ou escolha no computador</span><small>Excel .xlsx · todas as abas de dados</small><input ref={fileRef} type="file" accept=".xlsx" onChange={e=>{const f=e.target.files?.[0];if(f)void parseFile(f);}}/><button className="ap-button ap-primary" onClick={()=>fileRef.current?.click()}>Escolher arquivo</button></div>
    {uploadError&&<div className="ap-error"><AlertCircle size={16}/>{uploadError}</div>}{uploadNotice&&<div className="ap-note">{uploadNotice}</div>}
    <p className="ap-note">A aba Base_Itens alimenta os gráficos. Abas opcionais habilitam páginas próprias; as abas calculadas são ignoradas. O arquivo e as edições ficam no navegador.</p></section>}
    {source==="imported"&&<div className="ap-imported-counts"><div><span>ITENS IMPORTADOS</span><strong>{importedItems}</strong></div><div><span>SPRINTS IMPORTADAS</span><strong>{importedSprints}</strong></div>{snapshotDate&&<div><span>SNAPSHOT WIP</span><strong>{snapshotDate}</strong></div>}</div>}
    {source==="imported"&&<section className="ap-panel ap-quality-panel"><div className="ap-panel-heading"><div><h2><AlertCircle size={17}/> Alertas de qualidade</h2><p>Validações calculadas a partir das abas de dados.</p></div><span className="ap-quality-count">{qualityAlerts.length} alerta(s)</span></div>{qualityAlerts.length?<ul>{qualityAlerts.map((alert,i)=><li key={i}>{alert}</li>)}</ul>:<p className="ap-note">Nenhum problema identificado nas validações disponíveis.</p>}</section>}
    {activePage==="WIP e Aging"&&<WipAgingPage rows={sheetsData["WIP_Aberto"]||[]} snapshotDate={snapshotDate} baseRows={rows} source={source}/>} 
     {activePage==="Gargalo e DORA"&&<BottleneckDoraPage baseRows={rows} defects={sheetsData["Defeitos"]||[]} doraSprints={sheetsData["DORA_Sprint"]||[]} premises={sheetsData["Premissas"]||[]} source={source}/>}
     {activePage==="Resumo Executivo"&&<ExecutiveSummaryPage baseRows={rows} defects={sheetsData["Defeitos"]||[]} rework={sheetsData["Retrabalho"]||[]} calendar={calendarRows} doraSprints={sheetsData["DORA_Sprint"]||[]} premises={sheetsData["Premissas"]||[]} risks={sheetsData["Riscos_Gatilhos"]||[]} backlog={sheetsData["Backlog"]||[]} source={source}/>}
     {activePage==="Comunicação"&&<CommunicationPage stakeholders={sheetsData["Stakeholders"]||[]} metricDictionary={sheetsData["Dicionario_Metricas"]||[]} source={source}/>}
     {activePage==="Decisão e Riscos"&&<DecisionRisksPage baseRows={rows} backlog={sheetsData["Backlog"]||[]} premises={sheetsData["Premissas"]||[]} risks={sheetsData["Riscos_Gatilhos"]||[]} defects={sheetsData["Defeitos"]||[]} rework={sheetsData["Retrabalho"]||[]} calendar={sheetsData["Calendario_Eventos"]||[]} doraSprints={sheetsData["DORA_Sprint"]||[]} source={source}/>}
     {activePage==="Qualidade e Retrabalho"&&<QualityReworkPage defects={sheetsData["Defeitos"]||[]} rework={sheetsData["Retrabalho"]||[]} calendar={sheetsData["Calendario_Eventos"]||[]} premises={sheetsData["Premissas"]||[]} baseRows={rows} source={source}/>}
    {activePage!=="Dashboard"&&activePage!=="WIP e Aging"&&activePage!=="Qualidade e Retrabalho"&&activePage!=="Gargalo e DORA"&&activePage!=="Decisão e Riscos"&&activePage!=="Resumo Executivo"&&activePage!=="Comunicação"&&<section className="ap-panel ap-data-page"><div className="ap-panel-heading"><div><h2>{activePage}</h2><p>{activePage==="WIP_Aberto"&&snapshotDate?"Snapshot em "+snapshotDate:"Dados lidos diretamente da aba "+activePage}</p></div><span>{(sheetsData[activePage]||[]).length} linhas</span></div>{!sheetNames.includes(activePage)?<div className="ap-empty">Esta página depende da aba <strong>{activePage}</strong>. Envie a aba {activePage} no arquivo Excel para habilitá-la.</div>:<div className="ap-table-scroll"><table><thead><tr>{Object.keys((sheetsData[activePage]||[])[0]||{}).map(h=><th key={h}>{h}</th>)}</tr></thead><tbody>{(sheetsData[activePage]||[]).map((row,i)=><tr key={i}>{Object.keys((sheetsData[activePage]||[])[0]||{}).map(h=><td key={h}>{str(row[h])||"—"}</td>)}</tr>)}</tbody></table></div>}</section>}
    {activePage==="Dashboard"&&source==="imported"&&missingOptional.length>0&&<section className="ap-panel ap-missing-sheets"><strong>Abas opcionais não encontradas</strong><p>Envie as seguintes abas para habilitar suas páginas: {missingOptional.join(", ")}.</p></section>}
    <div className="ap-source-strip"><div className="ap-source-icon"><Database size={18}/></div><div><strong>{source==="demo"?"Dados demonstrativos (fictícios)":fileName}</strong><span>{source==="demo"?"Exemplo ilustrativo · não representa dados reais":"Dados ativos nesta sessão do navegador"}</span></div><div className="ap-source-state"><Check size={13}/>{source==="demo"?"Exemplo ativo":"Importação confirmada"} <button onClick={()=>setSourceOpen(true)}>Gerenciar</button></div><button className="ap-button ap-secondary ap-edit-toggle" onClick={()=>setEditorOpen(v=>!v)}><FileSpreadsheet size={15}/>{editorOpen?"Ocultar dados":"Visualizar e editar dados"}</button></div>
    {editorOpen&&<section className="ap-panel ap-data-editor"><div className="ap-panel-heading"><div><div><h2>Dados carregados</h2><p>Edite qualquer célula. Os gráficos e indicadores são recalculados automaticamente.</p></div></div><button className="ap-icon-button" onClick={()=>setEditorOpen(false)} aria-label="Fechar editor"><X size={18}/></button></div><div className="ap-editor-meta">{rows.length} linhas · {Object.keys(rows[0]||{}).length} colunas · alterações aplicadas em tempo real</div><div className="ap-table-scroll ap-edit-table"><table><thead><tr>{Object.keys(rows[0]||{}).map(h=><th key={h}>{h}</th>)}</tr></thead><tbody>{rows.map((row,rowIndex)=><tr key={rowIndex}>{Object.keys(rows[0]||{}).map(column=><td key={column}><input aria-label={`Linha ${rowIndex+1}, ${column}`} title={`${column} · linha ${rowIndex+1}`} value={str(row[column])} onChange={e=>updateCell(rowIndex,column,e.target.value)}/></td>)}</tr>)}</tbody></table></div><p className="ap-note"><AlertCircle size={14}/>As edições ficam nesta sessão do navegador. Para manter alterações fora desta sessão, exporte ou salve os dados em um arquivo.</p></section>}
    {activePage==="Dashboard"&&<>
    <div className="ap-chart-controls"><span><i/> Histórico de 20 sprints</span><label>Sprint <select value={filter} onChange={e=>setFilter(e.target.value)}><option>Todas</option>{SPRINT_OPTIONS.map(s=><option key={s}>{s}</option>)}</select></label></div>
    <div className="ap-flow-grid">
      <FlowChart title="Throughput" description="Itens por sprint, separados entre itens sem reabertura e itens reabertos." data={chartRows} fullData={flowSeries} kind="throughput" unavailable="Não há itens nas sprints selecionadas."/>
      <FlowChart title="Lead Time" description="Média por sprint e P85 inclusivo, em dias corridos." data={chartRows} fullData={flowSeries} kind="leadTime" unavailable="Não há valores válidos na coluna Lead time (dias corridos)."/>
      <FlowChart title="Cycle Time" description="Cycle Time médio em dias corridos, mais eficiência de fluxo (Cycle Time ÷ Lead Time, em %)." data={chartRows} fullData={flowSeries} kind="cycleTime" unavailable="Não há valores válidos nas colunas Cycle time (dias corridos) e Lead time (dias corridos)."/>
      <FlowChart title="Velocity" description="Soma dos Story Points dos itens concluídos em cada sprint." data={chartRows} fullData={flowSeries} kind="velocity" unavailable="Não há Story Points numéricos válidos nos itens concluídos."/>
    </div>
    <div className="ap-chart-footer"><span>Fonte: {source==="demo"?"dados fictícios":fileName}</span><span>Faixas de fase e eventos usam Calendario_Eventos.</span></div>
    <InlineAssistant metricKey="throughput" stats={stats} rows={rows} mapping={mapping} statuses={statuses} sourceLabel={source==="demo"?"dados demonstrativos":fileName} calendarRows={calendarRows} premises={sheetsData["Premissas"]||[]}/>
    </>}
    </div></section>
  </main>;
}
