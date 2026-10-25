import { useMemo, useRef, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { Bar, BarChart, CartesianGrid, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
import * as XLSX from "xlsx";
import { Activity, AlertCircle, Check, ChevronDown, Database, FileSpreadsheet, FileUp, MessageCircle, Send, Sparkles, X } from "lucide-react";

export const Route = createFileRoute("/")({ component: AgilePulse });
type Row = Record<string, unknown>;
type MapCols = { sprint: string; item: string; status: string; squad: string; created: string; started: string; completed: string; storyPoints: string };
type Msg = { role: "user" | "assistant"; content: string };
type MetricKey = "throughput" | "leadTime" | "cycleTime" | "velocity";
const metricNames: Record<MetricKey, string> = { throughput: "Throughput", leadTime: "Lead Time", cycleTime: "Cycle Time", velocity: "Velocity" };
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
      if (!done(r[m.status], statuses)) return;
      const id = m.item ? str(r[m.item]) : "";
      if (id && seen.has(id)) { duplicates++; return; }
      if (id) seen.add(id);
      throughput++;
      if (m.storyPoints) { const points = Number(str(r[m.storyPoints]).replace(",", ".")); if (Number.isFinite(points) && points > 0) velocity += points; }
      if (m.created && m.completed) { const days=elapsedDays(r[m.created],r[m.completed]); if(days!==null){leadTotal+=days;leadCount++;} }
      if (m.started && m.completed) { const days=elapsedDays(r[m.started],r[m.completed]); if(days!==null){cycleTotal+=days;cycleCount++;} }
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
function MetricChart({ title, description, data, dataKey, unit, color, unavailable, metricKey, stats, rows, mapping, statuses, sourceLabel }: { title:string; description:string; data: {sprint:string; [key:string]: any}[]; dataKey:string; unit:string; color:string; unavailable:string; metricKey:MetricKey; stats:ReturnType<typeof calculate>; rows:Row[]; mapping:MapCols; statuses:string[]; sourceLabel:string }) {
  const available = data.some(d => typeof d[dataKey] === "number" && Number.isFinite(d[dataKey]));
  const chartData = data.filter(d => typeof d[dataKey] === "number" && Number.isFinite(d[dataKey]));
  return <section className="ap-chart-card ap-metric-card">
    <div className="ap-chart-heading"><div><div className="ap-chart-title"><Activity size={19}/><h2>{title}</h2></div><p>{description}</p></div></div>
    {available ? <div className="ap-chart-scroll"><div className="ap-chart-inner" style={{width:Math.max(640,chartData.length*72)}}><ResponsiveContainer width="100%" height={260}><BarChart data={chartData} margin={{top:16,right:16,left:0,bottom:8}} barCategoryGap="36%"><CartesianGrid stroke="#263853" strokeDasharray="3 5" vertical={false}/><XAxis dataKey="sprint" tick={{fill:"#91a4c4",fontSize:11}} axisLine={{stroke:"#31435f"}} tickLine={false} interval={0}/><YAxis tick={{fill:"#91a4c4",fontSize:11}} axisLine={false} tickLine={false} width={45} domain={[0,(max:number)=>Math.max(1,Math.ceil(max*1.15))]}/><Tooltip contentStyle={{background:"#101d33",border:"1px solid #304565",borderRadius:12,color:"#f1f5ff"}} formatter={(v)=>[typeof v==="number"?v.toFixed(1)+" "+unit:String(v)+" "+unit,title]}/><Bar dataKey={dataKey} fill={color} radius={[7,7,2,2]} maxBarSize={36}/></BarChart></ResponsiveContainer></div></div> : <div className="ap-empty ap-metric-empty">{unavailable}</div>}
    <InlineAssistant metricKey={metricKey} stats={stats} rows={rows} mapping={mapping} statuses={statuses} sourceLabel={sourceLabel}/>
  </section>;
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
function AgilePulse() {
  const [rows,setRows]=useState<Row[]>(demo), [source,setSource]=useState<"demo"|"imported">("demo"), [fileName,setFileName]=useState("");
  const [pending,setPending]=useState<Row[]|null>(null), [sheetNames,setSheetNames]=useState<string[]>([]), [sheet,setSheet]=useState("");
  const [mapping,setMapping]=useState<MapCols>(guessMap(demo)), [statusText,setStatusText]=useState("Done, Concluído, Resolvido, Closed, Complete, Feito, Testado, Aprovado para Produção");
  const [sourceOpen,setSourceOpen]=useState(false), [editorOpen,setEditorOpen]=useState(false);
  const [uploadError,setUploadError]=useState(""), [uploadNotice,setUploadNotice]=useState("");
  const [filter,setFilter]=useState("Todas as sprints"); const fileRef=useRef<HTMLInputElement>(null);
  const headers=Object.keys((pending||rows)[0]||{}), statuses=statusText.split(",").map(s=>s.trim()).filter(Boolean);
  const stats=useMemo(()=>calculate(rows,mapping,statuses),[rows,mapping,statusText]);
  const chartRows=filter==="Todas as sprints"?stats.counts:stats.counts.filter(s=>s.sprint===filter);
  const quality=stats.noSprint+stats.noStatus+stats.noItem+stats.duplicates;
  async function parseFile(file: File) {
    setUploadError("");setUploadNotice("");
    setPending(null);
    try {
      const wb=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:true});
      if(!wb.SheetNames.length)throw new Error("O arquivo não contém abas legíveis.");
      const parsed: Row[]=[];
      const allColumns=new Set<string>();
      for(const sheetName of wb.SheetNames){
        const ws=wb.Sheets[sheetName];
        if(!ws)continue;
        const sheetRows=XLSX.utils.sheet_to_json<Row>(ws,{defval:"",raw:false});
        for(const row of sheetRows){
          if(!Object.values(row).some(value=>String(value??"").trim()!==""))continue;
          parsed.push({...row,"__Aba de origem":sheetName});
          Object.keys(row).forEach(key=>allColumns.add(key));
        }
      }
      if(!parsed.length)throw new Error("Nenhuma das abas contém linhas de dados com cabeçalhos reconhecíveis.");
      const columns=[...allColumns];
      const fallback=guessMap(parsed);
      let detected=fallback;
      let aiNote="";
      try{
        const response=await fetch("/api/ai",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({
          task:"organize",
          question:"Identifique automaticamente as colunas equivalentes às métricas ágeis e retorne somente JSON válido.",
          columns,
          sheetNames:wb.SheetNames,
          data:parsed.slice(0,120)
        })});
        const result=await response.json();
        if(response.ok&&result.answer){
          const cleaned=String(result.answer).replace(/^`{3}(?:json)?\s*/i,"").replace(/\s*`{3}$/,"");
          const inferred=JSON.parse(cleaned);
          const valid=(value:unknown)=>typeof value==="string"&&columns.includes(value);
          detected={
            sprint:valid(inferred.sprint)?inferred.sprint:fallback.sprint,
            item:valid(inferred.item)?inferred.item:fallback.item,
            status:valid(inferred.status)?inferred.status:fallback.status,
            squad:valid(inferred.squad)?inferred.squad:fallback.squad,
            created:valid(inferred.created)?inferred.created:fallback.created,
            started:valid(inferred.started)?inferred.started:fallback.started,
            completed:valid(inferred.completed)?inferred.completed:fallback.completed,
            storyPoints:valid(inferred.storyPoints)?inferred.storyPoints:fallback.storyPoints
          };
        }else aiNote=" A importação foi concluída com reconhecimento automático local porque a IA não estava disponível.";
      }catch{aiNote=" A importação foi concluída com reconhecimento automático local porque a IA não respondeu.";}
      setFileName(file.name);setSheetNames(wb.SheetNames);setSheet(wb.SheetNames[0]);
      setRows(parsed);setMapping(detected);setSource("imported");setPending(null);setEditorOpen(false);setFilter("Todas as sprints");
      setUploadNotice("Importação automática concluída: "+parsed.length+" linhas em "+wb.SheetNames.length+" abas e "+columns.length+" colunas reconhecidas."+aiNote);
      setSourceOpen(true);
    } catch(e) {setPending(null);setUploadError(e instanceof Error?e.message:"Não foi possível processar o arquivo.");}
  }
  async function changeSheet(name: string) {
    setSheet(name); const f=fileRef.current?.files?.[0]; if(!f)return;
    try {const wb=XLSX.read(await f.arrayBuffer(),{type:"array"});const data=XLSX.utils.sheet_to_json<Row>(wb.Sheets[name],{defval:"",raw:false});if(!data.length)throw new Error("A aba não contém linhas de dados.");setPending(data);setMapping(guessMap(data));setUploadError("");}
    catch(e){setPending(null);setUploadError(e instanceof Error?e.message:"Não foi possível ler a aba.");}
  }
  function confirmImport() {
    if(!pending)return;if(!mapping.sprint||!mapping.status){setUploadError("Mapeie Sprint e Status antes de confirmar.");return;}
    setRows(pending);setSource("imported");setPending(null);setSourceOpen(false);setEditorOpen(false);setFilter("Todas as sprints");setUploadError("");
  }
  function updateCell(rowIndex:number, column:string, value:string) {
    setRows(current=>current.map((row,index)=>index===rowIndex?({...row,[column]:value}):row));
  }
  return <main className="ap-shell">
    <aside className="ap-sidebar"><div className="ap-nav-label">NAVEGAÇÃO</div><button className="ap-nav active"><Activity size={18}/> Dashboard</button><button className={"ap-nav "+(sourceOpen?"active":"")} onClick={()=>setSourceOpen(v=>!v)}><Database size={18}/> Fontes de Dados <ChevronDown className="ap-nav-end" size={15}/></button><div className="ap-sidebar-bottom"><i/> <div><strong>Fonte de dados</strong><span>{source==="demo"?"Dados demonstrativos":"Arquivo importado"}</span></div></div></aside>
    <section className="ap-main">
    <div className="ap-content"><div className="ap-page-heading"><div><div className="ap-eyebrow">MÉTRICAS DE FLUXO</div><h1>Throughput <span className="ap-live-pill"><i/>{source==="demo"?"Demonstração":"Dados importados"}</span></h1><p>Acompanhe o ritmo de entrega e explore os itens concluídos por sprint.</p></div><button className="ap-button ap-secondary" onClick={()=>setSourceOpen(v=>!v)}><FileUp size={17}/> Importar dados</button></div>
    {sourceOpen&&<section className="ap-panel ap-import"><div className="ap-panel-heading"><div><h2>Fontes de Dados</h2><p>Excel ou CSV · a IA identifica os campos e importa automaticamente todas as abas.</p></div><button className="ap-icon-button" onClick={()=>{setSourceOpen(false);setPending(null);}}><X size={18}/></button></div>
    <div className="ap-dropzone" onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();const f=e.dataTransfer.files?.[0];if(f)void parseFile(f);}}><FileUp size={26}/><strong>Selecione sua planilha</strong><span>Arraste o arquivo ou escolha no computador</span><small>.xlsx · .xls · .csv</small><input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={e=>{const f=e.target.files?.[0];if(f)void parseFile(f);}}/><button className="ap-button ap-primary" onClick={()=>fileRef.current?.click()}>Escolher arquivo</button></div>
    {uploadError&&<div className="ap-error"><AlertCircle size={16}/>{uploadError}</div>}{uploadNotice&&<div className="ap-note">{uploadNotice}</div>}
    {pending&&<div className="ap-mapping"><div className="ap-file-info"><FileSpreadsheet size={20}/><div><strong>{fileName}</strong><span>{pending.length} linhas · {headers.length} colunas · prévia</span></div></div>{sheetNames.length>1&&<label className="ap-field">Aba do Excel<select value={sheet} onChange={e=>void changeSheet(e.target.value)}>{sheetNames.map(n=><option key={n}>{n}</option>)}</select></label>}
    <div className="ap-mapping-grid">{([{key:"sprint",label:"Sprint *"},{key:"item",label:"Identificador do item"},{key:"status",label:"Status *"},{key:"squad",label:"Squad (opcional)"},{key:"created",label:"Data de criação (Lead Time)"},{key:"started",label:"Data de início (Cycle Time)"},{key:"completed",label:"Data de conclusão (Lead/Cycle Time)"},{key:"storyPoints",label:"Story Points (Velocity)"}] as const).map(f=><label className="ap-field" key={f.key}>{f.label}<select value={mapping[f.key]} onChange={e=>setMapping(m=>({...m,[f.key]:e.target.value}))}><option value="">Não selecionada</option>{headers.map(h=><option key={h}>{h}</option>)}</select></label>)}</div>
    <label className="ap-field">Status que representam conclusão (separados por vírgula)<input value={statusText} onChange={e=>setStatusText(e.target.value)}/></label><div className="ap-preview"><strong>Prévia dos dados</strong><div className="ap-table-scroll"><table><thead><tr>{headers.slice(0,7).map(h=><th key={h}>{h}</th>)}</tr></thead><tbody>{pending.slice(0,5).map((r,i)=><tr key={i}>{headers.slice(0,7).map(h=><td key={h}>{str(r[h])||"—"}</td>)}</tr>)}</tbody></table></div></div>
    <div className="ap-import-actions"><button className="ap-button ap-secondary" onClick={()=>{setPending(null);setUploadError("");}}>Cancelar</button><button className="ap-button ap-primary" onClick={confirmImport}><Check size={16}/> Confirmar importação</button></div></div>}
    <p className="ap-note">Todas as abas e linhas são processadas no navegador; a IA ajuda a identificar as colunas. O arquivo não é salvo em banco compartilhado.</p></section>}
    <div className="ap-source-strip"><div className="ap-source-icon"><Database size={18}/></div><div><strong>{source==="demo"?"Dados demonstrativos (fictícios)":fileName}</strong><span>{source==="demo"?"Exemplo ilustrativo · não representa dados reais":"Dados ativos nesta sessão do navegador"}</span></div><div className="ap-source-state"><Check size={13}/>{source==="demo"?"Exemplo ativo":"Importação confirmada"} <button onClick={()=>setSourceOpen(true)}>Gerenciar</button></div><button className="ap-button ap-secondary ap-edit-toggle" onClick={()=>setEditorOpen(v=>!v)}><FileSpreadsheet size={15}/>{editorOpen?"Ocultar dados":"Visualizar e editar dados"}</button></div>
    {editorOpen&&<section className="ap-panel ap-data-editor"><div className="ap-panel-heading"><div><div><h2>Dados carregados</h2><p>Edite qualquer célula. Os gráficos e indicadores são recalculados automaticamente.</p></div></div><button className="ap-icon-button" onClick={()=>setEditorOpen(false)} aria-label="Fechar editor"><X size={18}/></button></div><div className="ap-editor-meta">{rows.length} linhas · {Object.keys(rows[0]||{}).length} colunas · alterações aplicadas em tempo real</div><div className="ap-table-scroll ap-edit-table"><table><thead><tr>{Object.keys(rows[0]||{}).map(h=><th key={h}>{h}</th>)}</tr></thead><tbody>{rows.map((row,rowIndex)=><tr key={rowIndex}>{Object.keys(rows[0]||{}).map(column=><td key={column}><input aria-label={`Linha ${rowIndex+1}, ${column}`} title={`${column} · linha ${rowIndex+1}`} value={str(row[column])} onChange={e=>updateCell(rowIndex,column,e.target.value)}/></td>)}</tr>)}</tbody></table></div><p className="ap-note"><AlertCircle size={14}/>As edições ficam nesta sessão do navegador. Para manter alterações fora desta sessão, exporte ou salve os dados em um arquivo.</p></section>}
    <section className="ap-chart-card"><div className="ap-chart-heading"><div><div className="ap-chart-title"><Activity size={20}/><h2>Itens concluídos por sprint</h2></div><p>Quantidade de itens que atingiram um status de conclusão em cada sprint.</p></div></div>
    <div className="ap-summary"><div><span>ITENS CONCLUÍDOS</span><strong>{stats.total}<small> itens</small></strong></div><div><span>MÉDIA POR SPRINT</span><strong>{stats.mean.toFixed(1)}<small> itens</small></strong></div><div><span>SPRINTS NO HISTÓRICO</span><strong>{stats.labels}</strong></div><div className="ap-quality"><i/>{quality?quality+" alerta(s) de qualidade":"Sem alertas de qualidade"}</div></div>
    <div className="ap-chart-controls"><span><i/> Itens concluídos</span><label>Visualização <select value={filter} onChange={e=>setFilter(e.target.value)}><option>Todas as sprints</option>{stats.counts.map(s=><option key={s.sprint}>{s.sprint}</option>)}</select></label></div>
    <div className="ap-chart-scroll"><div className="ap-chart-inner" style={{width:Math.max(640,chartRows.length*72)}}>{chartRows.length?<ResponsiveContainer width="100%" height={330}><BarChart data={chartRows} margin={{top:18,right:18,left:0,bottom:8}} barCategoryGap="36%"><defs><linearGradient id="throughputGradient" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#c084fc"/><stop offset="55%" stopColor="#7775f5"/><stop offset="100%" stopColor="#3787e8"/></linearGradient></defs><CartesianGrid stroke="#263853" strokeDasharray="3 5" vertical={false}/><XAxis dataKey="sprint" tick={{fill:"#91a4c4",fontSize:11}} axisLine={{stroke:"#31435f"}} tickLine={false} interval={0}/><YAxis allowDecimals={false} tick={{fill:"#91a4c4",fontSize:11}} axisLine={false} tickLine={false} width={35} domain={[0,(max:number)=>Math.max(4,max+1)]}/><Tooltip contentStyle={{background:"#101d33",border:"1px solid #304565",borderRadius:12,color:"#f1f5ff"}} cursor={{fill:"rgba(117,133,255,.08)"}} formatter={(v)=>[String(v)+" itens","Throughput"]}/><Bar dataKey="throughput" fill="url(#throughputGradient)" radius={[7,7,2,2]} maxBarSize={36}/></BarChart></ResponsiveContainer>:<div className="ap-empty">Nenhuma sprint disponível. Confira a coluna Sprint.</div>}</div></div><div className="ap-chart-footer"><span>Fonte: {source==="demo"?"dados fictícios":fileName}</span><span>Atualizado nesta sessão</span></div><InlineAssistant metricKey="throughput" stats={stats} rows={rows} mapping={mapping} statuses={statuses} sourceLabel={source==="demo"?"dados demonstrativos":fileName}/></section>
    <div className="ap-metric-grid">
      <MetricChart title="Lead Time" description="Tempo médio entre a criação da demanda e sua conclusão, agrupado por sprint." data={chartRows} dataKey="leadTime" unit="dias" color="#a78bfa" metricKey="leadTime" stats={stats} rows={rows} mapping={mapping} statuses={statuses} sourceLabel={source==="demo"?"dados demonstrativos":fileName} unavailable={mapping.created && mapping.completed ? "Não há datas válidas de criação e conclusão para calcular o Lead Time." : "Mapeie as colunas Data de criação e Data de conclusão na importação para calcular o Lead Time."}/>
      <MetricChart title="Cycle Time" description="Tempo médio entre o início efetivo do trabalho e a conclusão, agrupado por sprint." data={chartRows} dataKey="cycleTime" unit="dias" color="#38bdf8" metricKey="cycleTime" stats={stats} rows={rows} mapping={mapping} statuses={statuses} sourceLabel={source==="demo"?"dados demonstrativos":fileName} unavailable={mapping.started && mapping.completed ? "Não há datas válidas de início e conclusão para calcular o Cycle Time." : "Mapeie as colunas Data de início e Data de conclusão na importação para calcular o Cycle Time."}/>
      <MetricChart title="Velocity" description="Soma dos Story Points dos itens concluídos em cada sprint." data={chartRows} dataKey="velocity" unit="pontos" color="#34d399" metricKey="velocity" stats={stats} rows={rows} mapping={mapping} statuses={statuses} sourceLabel={source==="demo"?"dados demonstrativos":fileName} unavailable={mapping.storyPoints ? "Não há Story Points numéricos válidos nos itens concluídos." : "Mapeie a coluna Story Points na importação para calcular a Velocity."}/>
    </div>
    </div></section>
  </main>;
}
