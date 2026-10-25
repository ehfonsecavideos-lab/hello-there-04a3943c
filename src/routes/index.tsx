import { useEffect, useMemo, useRef, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { Bar, BarChart, CartesianGrid, ComposedChart, Line, ReferenceArea, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
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
function InlineAssistant({ metricKey, stats, rows, mapping, statuses, sourceLabel }: { metricKey:MetricKey; stats:ReturnType<typeof calculate>; rows:Row[]; mapping:MapCols; statuses:string[]; sourceLabel:string }) {
  const [analysisOpen,setAnalysisOpen]=useState(false), [chatOpen,setChatOpen]=useState(false);
  const [analysisText,setAnalysisText]=useState(""), [analysisLoading,setAnalysisLoading]=useState(false), [analysisNote,setAnalysisNote]=useState("");
  const [messages,setMessages]=useState<Msg[]>([{role:"assistant",content:"Olá! Faça perguntas sobre esta métrica. As respostas generativas usam a IA nativa do Lovable quando ela estiver habilitada."}]);
  const [question,setQuestion]=useState(""), [sending,setSending]=useState(false), [error,setError]=useState("");
  async function askAI(q:string, history:Msg[], task:"analysis"|"chat") {
    const res=await fetch("/api/ai",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({task,question:q,metric:metricKey,metricName:metricNames[metricKey],messages:history.slice(-8),data:rows.slice(0,500),mapping,completedStatuses:statuses,stats,metricSeries:stats.counts.map(s=>({sprint:s.sprint,value:s[metricKey]}))})});
    const body=await res.json();
    if(!res.ok||!body.answer) throw new Error(body.error||"A IA não conseguiu responder.");
    return String(body.answer);
  }
  async function toggleAnalysis() {
    if(analysisOpen){setAnalysisOpen(false);return;}
    setAnalysisOpen(true);setAnalysisLoading(true);setAnalysisNote("");setAnalysisText("");
    const prompt="Faça uma análise inteligente da métrica "+metricNames[metricKey]+" com base exclusivamente na série e nos dados fornecidos. Organize em: principais insights, riscos ou pontos de atenção e próximas ações recomendadas. Cite os valores e sprints reais; diferencie fatos de hipóteses e não invente causas.";
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
  const matrix=XLSX.utils.sheet_to_json<unknown[]>(ws,{header:1,defval:"",raw:true,blankrows:false}) as unknown[][];
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
  data.forEach((d,i)=>{const phase=d.phase||"Sem fase";const last=groups[groups.length-1];if(last&&last.phase===phase)last.end=d.sprint;else groups.push({phase,start:d.sprint,end:d.sprint,fill:phase==="Sem fase"?"#8293ad":(["Planejamento","Descoberta"].includes(phase)?"#4f86c6":(["Desenvolvimento","Construção"].includes(phase)?"#7b6bd6":(["Validação","Testes"].includes(phase)?"#3d9f8b":"#c28b4b"))});});
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
    <aside className="ap-sidebar"><div className="ap-nav-label">NAVEGAÇÃO</div><button className={"ap-nav "+(activePage==="Dashboard"?"active":"")} onClick={()=>setActivePage("Dashboard")}><Activity size={18}/> Dashboard</button><button className={"ap-nav "+(sourceOpen?"active":"")} onClick={()=>setSourceOpen(v=>!v)}><Database size={18}/> Fontes de Dados <ChevronDown className="ap-nav-end" size={15}/></button>{source==="imported"&&<><div className="ap-nav-label">ABAS DA PLANILHA</div>{["Base_Itens",...OPTIONAL_SHEET_NAMES].map(name=><button key={name} className={"ap-nav ap-subnav "+(activePage===name?"active":"")} onClick={()=>setActivePage(name)}>{name}{!sheetNames.includes(name)&&<span className="ap-missing-dot">!</span>}</button>)}</>}<div className="ap-sidebar-bottom"><i/> <div><strong>Fonte de dados</strong><span>{source==="demo"?"Dados demonstrativos":"Arquivo importado"}</span></div></div></aside>
    <section className="ap-main">
    <div className="ap-content"><div className="ap-page-heading"><div><div className="ap-eyebrow">MÉTRICAS DE FLUXO</div><h1>Throughput <span className="ap-live-pill"><i/>{source==="demo"?"Demonstração":"Dados importados"}</span></h1><p>Acompanhe o ritmo de entrega e explore os itens concluídos por sprint.</p></div><button className="ap-button ap-secondary" onClick={()=>setSourceOpen(v=>!v)}><FileUp size={17}/> Importar dados</button></div>
    {sourceOpen&&<section className="ap-panel ap-import"><div className="ap-panel-heading"><div><h2>Fontes de Dados</h2><p>Excel .xlsx · leitura das linhas e abas conforme o modelo da planilha.</p></div><button className="ap-icon-button" onClick={()=>{setSourceOpen(false);}}><X size={18}/></button></div>
    <div className="ap-import-choice"><button className="ap-button ap-secondary" onClick={()=>{setRows(demo);setMapping(guessMap(demo));setSource("demo");setFileName("");setSheetsData({});setSheetNames([]);setSnapshotDate("");setActivePage("Dashboard");setUploadError("");setUploadNotice("Dados demonstrativos ativados.");setSourceOpen(false);}}>Usar dados demonstrativos</button><span>ou importe seu arquivo Excel</span></div><div className="ap-dropzone" onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();const f=e.dataTransfer.files?.[0];if(f)void parseFile(f);}}><FileUp size={26}/><strong>Selecione sua planilha</strong><span>Arraste o arquivo ou escolha no computador</span><small>Excel .xlsx · todas as abas de dados</small><input ref={fileRef} type="file" accept=".xlsx" onChange={e=>{const f=e.target.files?.[0];if(f)void parseFile(f);}}/><button className="ap-button ap-primary" onClick={()=>fileRef.current?.click()}>Escolher arquivo</button></div>
    {uploadError&&<div className="ap-error"><AlertCircle size={16}/>{uploadError}</div>}{uploadNotice&&<div className="ap-note">{uploadNotice}</div>}
    <p className="ap-note">A aba Base_Itens alimenta os gráficos. Abas opcionais habilitam páginas próprias; as abas calculadas são ignoradas. O arquivo e as edições ficam no navegador.</p></section>}
    {source==="imported"&&<div className="ap-imported-counts"><div><span>ITENS IMPORTADOS</span><strong>{importedItems}</strong></div><div><span>SPRINTS IMPORTADAS</span><strong>{importedSprints}</strong></div>{snapshotDate&&<div><span>SNAPSHOT WIP</span><strong>{snapshotDate}</strong></div>}</div>}
    {source==="imported"&&<section className="ap-panel ap-quality-panel"><div className="ap-panel-heading"><div><h2><AlertCircle size={17}/> Alertas de qualidade</h2><p>Validações calculadas a partir das abas de dados.</p></div><span className="ap-quality-count">{qualityAlerts.length} alerta(s)</span></div>{qualityAlerts.length?<ul>{qualityAlerts.map((alert,i)=><li key={i}>{alert}</li>)}</ul>:<p className="ap-note">Nenhum problema identificado nas validações disponíveis.</p>}</section>}
    {activePage!=="Dashboard"&&<section className="ap-panel ap-data-page"><div className="ap-panel-heading"><div><h2>{activePage}</h2><p>{activePage==="WIP_Aberto"&&snapshotDate?"Snapshot em "+snapshotDate:"Dados lidos diretamente da aba "+activePage}</p></div><span>{(sheetsData[activePage]||[]).length} linhas</span></div>{!sheetNames.includes(activePage)?<div className="ap-empty">Esta página depende da aba <strong>{activePage}</strong>. Envie a aba {activePage} no arquivo Excel para habilitá-la.</div>:<div className="ap-table-scroll"><table><thead><tr>{Object.keys((sheetsData[activePage]||[])[0]||{}).map(h=><th key={h}>{h}</th>)}</tr></thead><tbody>{(sheetsData[activePage]||[]).map((row,i)=><tr key={i}>{Object.keys((sheetsData[activePage]||[])[0]||{}).map(h=><td key={h}>{str(row[h])||"—"}</td>)}</tr>)}</tbody></table></div>}</section>}
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
    <InlineAssistant metricKey="throughput" stats={stats} rows={rows} mapping={mapping} statuses={statuses} sourceLabel={source==="demo"?"dados demonstrativos":fileName}/>
    </>}
    </div></section>
  </main>;
}
