"use client";

import type { AIVisualization } from "@/types/analytics";

const palette = ["#4f46e5", "#0ea5e9", "#10b981", "#f59e0b", "#ef4444", "#8b5cf6"];

function compact(value: number, format?: string, currency?: string | null) {
  if (format === "percent") return `${value.toFixed(2)}%`;
  if (format === "currency") return new Intl.NumberFormat("en-US", { style: "currency", currency: currency || "AUD", notation: "compact", maximumFractionDigits: 2 }).format(value);
  return new Intl.NumberFormat("en-US", { notation: "compact", maximumFractionDigits: 2 }).format(value);
}

export function InsightChart({ visualization }: { visualization?: AIVisualization | null }) {
  if (!visualization?.labels?.length || !visualization.series?.length) return null;
  return (
    <div className="mt-4 overflow-hidden rounded-2xl border bg-muted/25 p-4">
      <div className="flex items-start justify-between gap-3">
        <div><p className="text-sm font-semibold">{visualization.title}</p>{visualization.subtitle ? <p className="mt-1 text-xs text-muted-foreground">{visualization.subtitle}</p> : null}</div>
        <span className="rounded-full border px-2 py-1 text-[10px] uppercase tracking-wide text-muted-foreground">{visualization.type.replace("_", " ")}</span>
      </div>
      <div className="mt-4 h-56 w-full">
        {visualization.type === "donut" ? <DonutChart visualization={visualization}/> : visualization.type === "waterfall" ? <WaterfallChart visualization={visualization}/> : visualization.type === "stacked_bar" ? <StackedBarChart visualization={visualization}/> : visualization.type === "bar" ? <BarChart visualization={visualization}/> : <LineChart visualization={visualization} area={visualization.type === "area"}/>} 
      </div>
      {visualization.type !== "waterfall" ? <div className="mt-3 flex flex-wrap gap-3 text-[11px] text-muted-foreground">{visualization.series.map((series, index) => <span key={series.name} className="flex items-center gap-1.5"><span className="size-2 rounded-full" style={{ background: palette[index % palette.length] }}/>{series.name}</span>)}</div> : null}
    </div>
  );
}

function LineChart({ visualization, area }: { visualization: AIVisualization; area?: boolean }) {
  const all = visualization.series.flatMap((s) => s.data).map(Number).filter(Number.isFinite);
  const min = Math.min(...all, 0), max = Math.max(...all, 1), range = Math.max(max - min, 1);
  const width = 640, height = 200, padX = 28, padY = 18;
  const x = (i: number) => padX + (i * (width - padX * 2)) / Math.max(visualization.labels.length - 1, 1);
  const y = (v: number) => height - padY - ((v - min) / range) * (height - padY * 2 - 16);
  return <svg viewBox={`0 0 ${width} ${height}`} className="h-full w-full overflow-visible">
    {[0,1,2,3].map((i) => <line key={i} x1={padX} x2={width-padX} y1={padY+i*(height-padY*2-16)/3} y2={padY+i*(height-padY*2-16)/3} stroke="currentColor" opacity=".08"/>)}
    {visualization.series.map((series, si) => {
      const points = series.data.map((v,i) => `${x(i)},${y(Number(v)||0)}`).join(" ");
      const areaPoints = `${x(0)},${height-padY-16} ${points} ${x(series.data.length-1)},${height-padY-16}`;
      return <g key={series.name}>{area ? <polygon points={areaPoints} fill={palette[si%palette.length]} opacity=".12"/> : null}<polyline points={points} fill="none" stroke={palette[si%palette.length]} strokeWidth="3" strokeLinecap="round" strokeLinejoin="round"/>{series.data.map((v,i)=><circle key={i} cx={x(i)} cy={y(Number(v)||0)} r="3" fill={palette[si%palette.length]}><title>{visualization.labels[i]} · {series.name}: {compact(Number(v)||0,visualization.value_format,visualization.currency)}</title></circle>)}</g>;
    })}
    {visualization.labels.map((label,i) => i % Math.max(Math.ceil(visualization.labels.length/6),1)===0 ? <text key={label+i} x={x(i)} y={height-2} textAnchor="middle" fontSize="9" fill="currentColor" opacity=".55">{String(label).slice(0,10)}</text> : null)}
  </svg>;
}

function BarChart({ visualization }: { visualization: AIVisualization }) {
  const all = visualization.series.flatMap((s)=>s.data).map(Number).filter(Number.isFinite); const max=Math.max(...all.map(Math.abs),1);
  const width=640,height=200,pad=28; const groups=visualization.labels.length; const groupW=(width-pad*2)/Math.max(groups,1); const barW=Math.max(7,(groupW*.72)/Math.max(visualization.series.length,1));
  return <svg viewBox={`0 0 ${width} ${height}`} className="h-full w-full">{visualization.labels.map((label,i)=><g key={label+i}>{visualization.series.map((s,si)=>{const value=Number(s.data[i])||0;const h=Math.abs(value)/max*(height-60);const bx=pad+i*groupW+(groupW-barW*visualization.series.length)/2+si*barW;return <rect key={s.name} x={bx} y={height-30-h} width={Math.max(barW-2,2)} height={h} rx="4" fill={palette[si%palette.length]}><title>{label} · {s.name}: {compact(value,visualization.value_format,visualization.currency)}</title></rect>})}<text x={pad+i*groupW+groupW/2} y={height-8} textAnchor="middle" fontSize="9" fill="currentColor" opacity=".55">{String(label).slice(0,10)}</text></g>)}</svg>;
}

function StackedBarChart({ visualization }: { visualization: AIVisualization }) {
  const totals = visualization.labels.map((_, i) => visualization.series.reduce((sum, series) => sum + Math.max(Number(series.data[i]) || 0, 0), 0));
  const max = Math.max(...totals, 1), width = 640, height = 200, pad = 28, groupW = (width-pad*2)/Math.max(visualization.labels.length,1), barW = Math.min(46, groupW*.58);
  return <svg viewBox={`0 0 ${width} ${height}`} className="h-full w-full">{visualization.labels.map((label,i)=>{let cursor=height-30;return <g key={label+i}>{visualization.series.map((s,si)=>{const value=Math.max(Number(s.data[i])||0,0);const h=value/max*(height-60);cursor-=h;return <rect key={s.name} x={pad+i*groupW+(groupW-barW)/2} y={cursor} width={barW} height={h} fill={palette[si%palette.length]}><title>{label} · {s.name}: {compact(value,visualization.value_format,visualization.currency)}</title></rect>})}<text x={pad+i*groupW+groupW/2} y={height-8} textAnchor="middle" fontSize="9" fill="currentColor" opacity=".55">{String(label).slice(0,10)}</text></g>})}</svg>;
}

function DonutChart({ visualization }: { visualization: AIVisualization }) {
  const values=visualization.series[0]?.data.map(Number) || []; const total=values.reduce((a,b)=>a+Math.max(b,0),0)||1; let offset=0;
  return <div className="flex h-full items-center justify-center gap-6"><svg viewBox="0 0 120 120" className="h-40 w-40 -rotate-90">{values.map((v,i)=>{const pct=Math.max(v,0)/total;const dash=pct*251.2;const el=<circle key={i} cx="60" cy="60" r="40" fill="none" stroke={palette[i%palette.length]} strokeWidth="18" strokeDasharray={`${dash} ${251.2-dash}`} strokeDashoffset={-offset}/>;offset+=dash;return el;})}</svg><div className="space-y-2">{visualization.labels.map((label,i)=><div key={label+i} className="flex items-center gap-2 text-xs"><span className="size-2.5 rounded-full" style={{background:palette[i%palette.length]}}/><span className="text-muted-foreground">{label}</span><b>{compact(values[i]||0,visualization.value_format,visualization.currency)}</b></div>)}</div></div>;
}

function WaterfallChart({ visualization }: { visualization: AIVisualization }) {
  const values = visualization.series[0]?.data.map(Number) || [];
  if (!values.length) return null;
  const cumulative: number[] = [];
  let running = 0;
  values.forEach((value) => { running += value || 0; cumulative.push(running); });
  const points = [0, ...cumulative];
  const min = Math.min(...points), max = Math.max(...points), range = Math.max(max-min, 1);
  const width=640,height=200,pad=30,groupW=(width-pad*2)/values.length,barW=Math.min(56,groupW*.62);
  const y=(v:number)=>height-30-((v-min)/range)*(height-62);
  return <svg viewBox={`0 0 ${width} ${height}`} className="h-full w-full">{[0,1,2,3].map(i=><line key={i} x1={pad} x2={width-pad} y1={20+i*(height-55)/3} y2={20+i*(height-55)/3} stroke="currentColor" opacity=".07"/>)}{values.map((value,i)=>{const start=i===0?0:cumulative[i-1];const end=cumulative[i];const top=Math.min(y(start),y(end));const h=Math.max(Math.abs(y(end)-y(start)),2);const color=value>=0?"#10b981":"#ef4444";const x=pad+i*groupW+(groupW-barW)/2;return <g key={`${visualization.labels[i]}-${i}`}>{i>0?<line x1={pad+(i-1)*groupW+groupW/2+barW/2} x2={x} y1={y(start)} y2={y(start)} stroke="currentColor" opacity=".18" strokeDasharray="3 3"/>:null}<rect x={x} y={top} width={barW} height={h} rx="4" fill={color}><title>{visualization.labels[i]}: {compact(value,visualization.value_format,visualization.currency)} · running total {compact(end,visualization.value_format,visualization.currency)}</title></rect><text x={pad+i*groupW+groupW/2} y={height-8} textAnchor="middle" fontSize="9" fill="currentColor" opacity=".58">{String(visualization.labels[i]).slice(0,11)}</text></g>})}</svg>;
}
