import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import {
  Activity, AlertTriangle, BarChart3, Bell, CheckCircle2, ChevronDown,
  CircleDot, Clock3, Database, Download, Filter, LayoutDashboard,
  ListChecks, Menu, Plus, RefreshCw, Search, Settings2, ShieldCheck,
  SlidersHorizontal, Trash2, Upload, Users, X, Zap, ArrowUpRight,
  MoreHorizontal, Eye, Edit3, ChevronRight
} from "lucide-react";
import {
  ResponsiveContainer, AreaChart, Area, BarChart, Bar, XAxis, YAxis,
  CartesianGrid, Tooltip, PieChart, Pie, Cell
} from "recharts";
import "./styles.css";

const STATUS_MAP = {
  DELIVERED: "delivered",
  EXPIRED: "failed",
  UNDELIVERABLE: "failed",
  NO_DLR_RECEIVED: "waiting",
  REJECTED: "rejected",
  PENDING: "pending",
  WAITING: "waiting"
};

const seedRows = [
  ["Indonesia","XL","510","11","ACORN_DIR","Direct","Qsms","UNDELIVERABLE","NEGATIVE","15","6287742506345","03-10-2026 05:54:03","ef22c8-30891-8f31b-71139-393c2-82765"],
  ["Kuwait","Zain","419","02","MONTY_DIR","Direct","PKO BP","DELIVERED","POSITIVE","8","96599850653","03-10-2026 05:44:19","c7ca0d-35666-8c397-80936-ae7dd-37819"],
  ["Kuwait","Zain","419","02","MONTY_DIR","Direct","PKO BP","DELIVERED","POSITIVE","9","96598845437","03-10-2026 05:44:19","51ff99-11680-389f3-95820-da041-23558"],
  ["Kuwait","Zain","419","02","MONTY_DIR","Direct","PKO BP","DELIVERED","POSITIVE","9","96599028746","03-10-2026 05:44:19","9adf73-15952-c3c66-89295-b91dc-42575"],
  ["Kuwait","Zain","419","02","MONTY_DIR","Direct","PKO BP","DELIVERED","POSITIVE","9","96599838764","03-10-2026 05:44:19","4deb17-17778-ca4bc-35688-0b985-40691"],
  ["Kuwait","Zain","419","02","MONTY_DIR","Direct","PKO BP","DELIVERED","POSITIVE","9","96599461814","03-10-2026 05:44:19","b12d39-30265-59b17-68342-6fa17-36699"],
  ["Indonesia","XL","510","11","ACORN_DIR","Direct","Qsms","EXPIRED","NEGATIVE","310","6287798667304","03-10-2026 05:21:29","9cc84f-39560-36a5-c93011-cd21b-91103"],
  ["Indonesia","XL","510","11","ACORN_DIR","Direct","Qsms","DELIVERED","POSITIVE","12","6287767329787","03-10-2026 05:21:29","175af9-30722-0fdb5-33821-f7f7b-23246"],
  ["Indonesia","XL","510","11","ACORN_DIR","Direct","Qsms","DELIVERED","POSITIVE","10","6285921691579","03-10-2026 05:21:29","696790-32226-cc0d5-96557-943a8-22529"],
  ["Indonesia","XL","510","11","ACORN_DIR","Direct","Qsms","DELIVERED","TEST_NUMBER_OFFLINE","14","6281947970002","03-10-2026 05:21:29","d977fb-14852-13584-25886-1545a-23857"],
  ["Indonesia","XL","510","11","ACORN_DIR","Direct","WeTV","EXPIRED","NEGATIVE","309","6287798667303","03-10-2026 05:21:29","c79ec4-31831-e0ff-67547-65b66-13118"],
  ["Indonesia","XL","510","11","ACORN_DIR","Direct","WeTV","EXPIRED","NEGATIVE","304","6281927738343","03-10-2026 05:21:29","64a224-41648-a2832-68274-82df5-21113"],
  ["Indonesia","XL","510","11","ACORN_DIR","Direct","WeTV","DELIVERED","NEGATIVE","14","6283873907867","03-10-2026 05:21:29","0fc17d-34146-6f59-a13488-1f1ea-11402"],
  ["Indonesia","XL","510","11","ACORN_DIR","Direct","Shopee","DELIVERED","NEGATIVE","16","6281934245429","03-10-2026 05:21:29","41e49d-16955-af5c-7-53005-41b23-23247"],
  ["Indonesia","XL","510","11","ACORN_DIR","Direct","Shopee","DELIVERED","POSITIVE","11","6285921691579","03-10-2026 05:21:29","0172e8-37278-d48c6-27058-1f1bc-36881"],
  ["Indonesia","XL","510","11","ACORN_DIR","Direct","Shopee","DELIVERED","POSITIVE","12","6285940840479","03-10-2026 05:21:29","ed6b2c-27156-82b52-67073-f8f21-21295"],
  ["Indonesia","XL","510","11","SIGMA_DIR","Direct","StripeLink","DELIVERED","TEXT_REPLACED","15","62859175623714","03-10-2026 05:09:28","7f6717-37873-6b172-82536-f89b7-38444"],
  ["Indonesia","XL","510","11","SIGMA_DIR","Direct","StripeLink","DELIVERED","TEXT_REPLACED","10","6283143362653","03-10-2026 05:09:28","92003a-40683-4845e-02911-937be-62715"],
  ["Indonesia","XL","510","11","SIGMA_DIR","Direct","StripeLink","DELIVERED","TEXT_REPLACED","18","6287865579431","03-10-2026 05:09:28","9a3432-35538-7d04b-21018-e5f31-29790"],
  ["Indonesia","XL","510","11","SIGMA_DIR","Direct","StripeLink","NO_DLR_RECEIVED","NEGATIVE","","6287798667304","03-10-2026 05:09:28","024c46-25624-b884-2-66999-f17f4-28237"],
  ["Indonesia","XL","510","11","SIGMA_DIR","Direct","StripeLink","NO_DLR_RECEIVED","NEGATIVE","","6281927738339","03-10-2026 05:09:28","3b6e29-35316-231dd-47497-44d55-19307"],
  ["Indonesia","XL","510","11","SIGMA_DIR","Direct","Stripe","DELIVERED","TEXT_REPLACED","8","6287814116875","03-10-2026 05:08:33","163162-41475-82cc4-91336-70276-25338"],
  ["Indonesia","XL","510","11","SIGMA_DIR","Direct","Stripe","DELIVERED","TEXT_REPLACED","6","6283873330226","03-10-2026 05:08:33","1b12cb-14707-8b07d-54576-19a10-12796"],
  ["Indonesia","XL","510","11","SIGMA_DIR","Direct","Stripe","DELIVERED","TEXT_REPLACED","12","6287764798566","03-10-2026 05:08:33","c3b53b-11085-adce8-27162-03125-71057"],
  ["Indonesia","XL","510","11","SIGMA_DIR","Direct","Stripe","DELIVERED","TEXT_REPLACED","16","6283127921478","03-10-2026 05:08:33","06733e-71441-525d8-65821-5e4bf-56602"],
  ["Indonesia","XL","510","11","SIGMA_DIR","Direct","Stripe","DELIVERED","TEXT_REPLACED","7","6281906351539","03-10-2026 05:08:33","53d52b-21452-9027-c56293-4c8b0-15331"],
  ["Indonesia","XL","510","11","SIGMA_DIR","Direct","Stripe","DELIVERED","TEST_NUMBER_OFFLINE","6","6283819656567","03-10-2026 05:08:14","188dd4-15571-8f0f6-04639-17381-30520"],
  ["Indonesia","XL","510","11","SIGMA_DIR","Direct","Stripe","DELIVERED","TEXT_REPLACED","7","6283819656567","03-10-2026 05:08:14","d1b752-75895-719b7-61663-55c92-18294"],
  ["Indonesia","XL","510","11","SIGMA_DIR","Direct","Stripe","DELIVERED","TEXT_REPLACED","8","6283819656567","03-10-2026 05:08:14","e84f53-33551-a0afe-78478-51e75-11928"],
  ["Indonesia","XL","510","11","SIGMA_DIR","Direct","Stripe","DELIVERED","TEXT_REPLACED","7","6283819656567","03-10-2026 05:08:14","d026c6-90021-baf87-87772-cc858-75941"],
  ["Indonesia","XL","510","11","SIGMA_DIR","Direct","SECURITYBNK","DELIVERED","POSITIVE","12","6281805070691","03-10-2026 03:22:11","31a400-33755-ab57d-89412-2b36c-12251"],
  ["Indonesia","XL","510","11","SIGMA_DIR","Direct","SECURITYBNK","DELIVERED","POSITIVE","8","6287814116875","03-10-2026 03:22:11","5af6ad-21057-56cfb-27463-9d4fb-18556"],
  ["Indonesia","XL","510","11","SIGMA_DIR","Direct","SECURITYBNK","DELIVERED","POSITIVE","8","6283846703223","03-10-2026 03:22:11","24376f-10773-96fe-32485-db72a-16207"],
  ["Indonesia","XL","510","11","SIGMA_DIR","Direct","SECURITYBNK","DELIVERED","POSITIVE","7","6287814116875","03-10-2026 03:22:11","6869b2-28499-514d8-32012-7b05a-75246"],
  ["Indonesia","XL","510","11","SIGMA_DIR","Direct","SECURITYBNK","DELIVERED","POSITIVE","6","6287761312704","03-10-2026 03:22:11","0cde47-23866-e01a8-11699-36e67-22796"],
  ["Indonesia","XL","510","11","SIGMA_DIR","Direct","SECURITYBNK","DELIVERED","POSITIVE","7","6287814116875","03-10-2026 03:22:11","65c865-33239-0578-c38733-3472e-42651"],
  ["Indonesia","XL","510","11","SIGMA_DIR","Direct","SECURITYBNK","DELIVERED","POSITIVE","7","6283819656567","03-10-2026 03:21:46","427298-81509-47fc1-01502-8e4e3-62772"]
];

const makeRow = (r, i) => ({
  id: `seed-${i}-${r[12]}`,
  testId: r[1] || String(i + 1000),
  country: r[0], network: r[1], mcc: r[2], mnc: r[3],
  supplier: r[4], routeType: r[5], senderId: r[6],
  dlrStatus: r[7], receiptStatus: r[8], delay: r[9] === "" ? null : Number(r[9]),
  phone: r[10], timestamp: r[11], messageId: r[12],
  importedAt: Date.now(), shift: getShift(r[11])
});

function getShift(dateText) {
  const m = String(dateText || "").match(/(\d{2}):(\d{2}):/);
  const hour = m ? Number(m[1]) : 12;
  if (hour < 8) return "Shift 1";
  if (hour < 16) return "Shift 2";
  return "Shift 3";
}

const seedResults = seedRows.map(makeRow);

const defaultRequirements = [
  { id:"r1", country:"Indonesia", network:"XL", mcc:"510", mnc:"11", senderId:"Qsms", vendor:"ACORN_DIR", routeType:"Direct", frequency:"Every Shift", active:true },
  { id:"r2", country:"Indonesia", network:"XL", mcc:"510", mnc:"11", senderId:"Stripe", vendor:"SIGMA_DIR", routeType:"Direct", frequency:"Every Shift", active:true },
  { id:"r3", country:"Indonesia", network:"XL", mcc:"510", mnc:"11", senderId:"SECURITYBNK", vendor:"SIGMA_DIR", routeType:"Direct", frequency:"Every Shift", active:true },
  { id:"r4", country:"Kuwait", network:"Zain", mcc:"419", mnc:"02", senderId:"PKO BP", vendor:"MONTY_DIR", routeType:"Direct", frequency:"Every Shift", active:true },
  { id:"r5", country:"Indonesia", network:"XL", mcc:"510", mnc:"11", senderId:"StripeLink", vendor:"SIGMA_DIR", routeType:"Direct", frequency:"Every Shift", active:true }
];

const navItems = [
  {id:"dashboard", label:"Dashboard", icon:LayoutDashboard},
  {id:"testing", label:"Shift Testing", icon:ListChecks},
  {id:"results", label:"Results Explorer", icon:BarChart3},
  {id:"destinations", label:"Destinations", icon:SlidersHorizontal},
  {id:"sales", label:"Sales Management", icon:Users},
  {id:"activity", label:"Activity Log", icon:Activity}
];

function App() {
  const [page, setPage] = useState("dashboard");
  const [rows, setRows] = useState(() => JSON.parse(localStorage.getItem("sms_rows") || "null") || seedResults);
  const [requirements, setRequirements] = useState(() => JSON.parse(localStorage.getItem("sms_requirements") || "null") || defaultRequirements);
  const [activity, setActivity] = useState(() => JSON.parse(localStorage.getItem("sms_activity") || "null") || [
    {id:1, text:"System initialized with testing-tool sample data", actor:"System", time:"Just now"},
    {id:2, text:"Sales requirement set created for Indonesia / XL", actor:"Sales", time:"Today 10:14"},
    {id:3, text:"Kuwait / Zain HQ requirement activated", actor:"Sales", time:"Today 09:42"}
  ]);
  const [role, setRole] = useState("Support");
  const [search, setSearch] = useState("");
  const [showImport, setShowImport] = useState(false);
  const [showReq, setShowReq] = useState(false);
  const [editingReq, setEditingReq] = useState(null);
  const [selectedGroup, setSelectedGroup] = useState(null);

  useEffect(() => localStorage.setItem("sms_rows", JSON.stringify(rows)), [rows]);
  useEffect(() => localStorage.setItem("sms_requirements", JSON.stringify(requirements)), [requirements]);
  useEffect(() => localStorage.setItem("sms_activity", JSON.stringify(activity)), [activity]);

  const addActivity = (text, actor=role) => {
    setActivity(a => [{id:Date.now(), text, actor, time:"Just now"}, ...a].slice(0,100));
  };

  const importRows = async (file) => {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, {type:"array", cellDates:true});
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(sheet, {defval:""});
    const mapped = raw.map((x,i) => {
      const get = (...keys) => {
        const k = keys.find(k => Object.prototype.hasOwnProperty.call(x,k));
        return k ? x[k] : "";
      };
      const timestamp = String(get("Date (UTC)","Date","date")).replace("T"," ");
      return {
        id:`import-${Date.now()}-${i}`,
        testId:String(get("Test ID text","Test ID","ID") || i+1),
        country:String(get("Country")),
        network:String(get("Network")),
        mcc:String(get("MCC")),
        mnc:String(get("MNC")),
        supplier:String(get("Supplier")),
        routeType:String(get("Route Type")),
        senderId:String(get("Sender ID Sent","Sender ID","Sender")),
        dlrStatus:String(get("DLR status")).toUpperCase(),
        receiptStatus:String(get("Receipt Status")).toUpperCase(),
        delay:Number(get("DLR Delay (s)")) || null,
        phone:String(get("Phone")),
        timestamp,
        messageId:String(get("Message ID")),
        importedAt:Date.now(),
        shift:getShift(timestamp)
      };
    }).filter(x => x.country || x.supplier || x.dlrStatus);
    setRows(prev => [...mapped, ...prev]);
    addActivity(`Imported ${mapped.length} testing result${mapped.length===1?"":"s"} from ${file.name}`);
    setShowImport(false);
  };

  const clearData = () => {
    if (!confirm("Reset prototype data to the original sample?")) return;
    setRows(seedResults); setRequirements(defaultRequirements);
    addActivity("Prototype data reset to original sample", "Admin");
  };

  return <div className="app">
    <Sidebar page={page} setPage={setPage} role={role} setRole={setRole} />
    <main className="main">
      <Topbar page={page} search={search} setSearch={setSearch} onImport={()=>setShowImport(true)} role={role} setRole={setRole} />
      <div className="content">
        {page==="dashboard" && <Dashboard rows={rows} requirements={requirements} setPage={setPage} />}
        {page==="testing" && <ShiftTesting rows={rows} requirements={requirements} search={search} onSelect={setSelectedGroup} />}
        {page==="results" && <ResultsExplorer rows={rows} search={search} onSelect={setSelectedGroup} />}
        {page==="destinations" && <Destinations requirements={requirements} setRequirements={setRequirements} onAddActivity={addActivity} />}
        {page==="sales" && <SalesManagement requirements={requirements} setRequirements={setRequirements} onAddActivity={addActivity} />}
        {page==="activity" && <ActivityLog activity={activity} />}
      </div>
    </main>
    {showImport && <ImportModal onClose={()=>setShowImport(false)} onImport={importRows} />}
    {showReq && <RequirementModal req={editingReq} onClose={()=>{setShowReq(false);setEditingReq(null)}} onSave={(req)=>{
      setRequirements(old => editingReq ? old.map(x=>x.id===req.id?req:x) : [...old,{...req,id:`r-${Date.now()}`}]);
      addActivity(`${editingReq?"Updated":"Added"} testing requirement for ${req.country} / ${req.network}`);
      setShowReq(false); setEditingReq(null);
    }} />}
    {selectedGroup && <GroupDrawer group={selectedGroup} onClose={()=>setSelectedGroup(null)} />}
    <button className="floating-help" title="Reset prototype" onClick={clearData}><RefreshCw size={16}/></button>
  </div>;
}

function Sidebar({page,setPage,role,setRole}) {
  return <aside className="sidebar">
    <div className="brand">
      <div className="brand-mark"><Zap size={18}/></div>
      <div><strong>RouteControl</strong><span>SMS Operations</span></div>
    </div>
    <div className="env"><CircleDot size={11}/> INTERNAL PROTOTYPE</div>
    <div className="nav-section">OPERATIONS</div>
    <nav>{navItems.map(({id,label,icon:Icon})=><button key={id} className={page===id?"nav-item active":"nav-item"} onClick={()=>setPage(id)}><Icon size={18}/><span>{label}</span>{id==="testing"&&<span className="nav-count">12</span>}</button>)}</nav>
    <div className="sidebar-bottom">
      <div className="nav-section">ACCESS</div>
      <select value={role} onChange={e=>setRole(e.target.value)} className="role-select">
        <option>Support</option><option>Sales Manager</option><option>Management</option><option>Admin</option>
      </select>
      <div className="user-card"><div className="avatar">SK</div><div><strong>Operations User</strong><span>{role}</span></div><MoreHorizontal size={16}/></div>
    </div>
  </aside>
}

function Topbar({page,search,setSearch,onImport,role,setRole}) {
  const title = navItems.find(x=>x.id===page)?.label || "Dashboard";
  return <header className="topbar">
    <div className="breadcrumbs"><span>RouteControl</span><ChevronRight size={14}/><strong>{title}</strong></div>
    <div className="top-actions">
      <div className="global-search"><Search size={16}/><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search destinations, vendors..." /></div>
      <button className="icon-btn"><Bell size={18}/><i/></button>
      <button className="import-btn" onClick={onImport}><Upload size={16}/> Import Results</button>
    </div>
  </header>
}

function StatCard({label,value,sub,icon:Icon,tone="blue",trend}) {
  return <div className="stat-card">
    <div className={`stat-icon ${tone}`}><Icon size={19}/></div>
    <div className="stat-body"><span>{label}</span><strong>{value}</strong><small className={trend?.positive?"positive":""}>{trend?.text || sub}</small></div>
    {trend && <ArrowUpRight className="trend-icon" size={17}/>}
  </div>
}

function Dashboard({rows,requirements,setPage}) {
  const groups = useMemo(()=>groupRows(rows),[rows]);
  const total=groups.length, working=groups.filter(g=>g.delivered>=2).length, failed=groups.filter(g=>g.delivered<2).length;
  const delivery=rows.length?Math.round(rows.filter(r=>r.dlrStatus==="DELIVERED").length/rows.length*100):0;
  const chart=buildTrend(rows);
  const vendorData=Object.entries(rows.reduce((a,r)=>{a[r.supplier]=(a[r.supplier]||0)+1;return a},{})).map(([name,value])=>({name,value}));
  return <div>
    <PageIntro eyebrow="OPERATIONS OVERVIEW" title="SMS Route Control Center" desc="Live view of recurring HQ route testing, shift coverage and imported testing-tool results." action={<button className="primary-btn" onClick={()=>setPage("testing")}><ListChecks size={16}/> Open current shift</button>}/>
    <div className="stats-grid">
      <StatCard label="Test groups" value={total} sub={`${requirements.filter(r=>r.active).length} active requirements`} icon={ListChecks}/>
      <StatCard label="Working routes" value={working} sub={`${total?Math.round(working/total*100):0}% of groups`} icon={CheckCircle2} tone="green" trend={{text:"Operational",positive:true}}/>
      <StatCard label="Needs attention" value={failed} sub="Below 2 delivered" icon={AlertTriangle} tone="amber"/>
      <StatCard label="Delivery rate" value={`${delivery}%`} sub={`${rows.length} individual attempts`} icon={Activity} tone="purple"/>
    </div>
    <div className="dashboard-grid">
      <section className="panel chart-panel">
        <PanelHeader title="Testing activity" subtitle="Imported attempts by hour" action="Today"/>
        <div className="chart-wrap"><ResponsiveContainer width="100%" height={270}><AreaChart data={chart}><defs><linearGradient id="fillA" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#4f7cff" stopOpacity=".25"/><stop offset="100%" stopColor="#4f7cff" stopOpacity="0"/></linearGradient></defs><CartesianGrid vertical={false} stroke="#e7ebf2"/><XAxis dataKey="hour" tickLine={false} axisLine={false}/><YAxis tickLine={false} axisLine={false} allowDecimals={false}/><Tooltip/><Area type="monotone" dataKey="attempts" stroke="#4f7cff" fill="url(#fillA)" strokeWidth={2.5}/></AreaChart></ResponsiveContainer></div>
      </section>
      <section className="panel">
        <PanelHeader title="Route vendors" subtitle="Attempt volume by supplier" />
        <div className="mini-bars">{vendorData.slice(0,6).map((x,i)=><div className="mini-bar" key={x.name}><div><span>{x.name}</span><b>{x.value}</b></div><div className="bar-track"><i style={{width:`${Math.min(100,x.value/Math.max(...vendorData.map(v=>v.value))*100)}%`}}/></div></div>)}</div>
      </section>
    </div>
    <div className="dashboard-grid lower">
      <section className="panel">
        <PanelHeader title="Current shift snapshot" subtitle="Grouped by destination / network / supplier" action={<button className="text-btn" onClick={()=>setPage("testing")}>View all <ArrowUpRight size={14}/></button>}/>
        <GroupTable groups={groups.slice(0,6)} compact onSelect={()=>{}} />
      </section>
      <section className="panel">
        <PanelHeader title="Coverage health" subtitle="Sales-configured requirements"/>
        <div className="coverage">
          {requirements.filter(r=>r.active).slice(0,6).map(r=>{
            const g=groups.find(g=>sameRequirement(g,r));
            const pct=g?Math.min(100,Math.round(g.delivered/Math.max(3,g.total)*100)):0;
            return <div className="coverage-row" key={r.id}><div className="coverage-title"><span>{r.country} / {r.network}</span><b>{g?`${g.delivered}/${g.total}`:"Not tested"}</b></div><div className="bar-track"><i className={pct>=60?"good":pct>0?"warn":"bad"} style={{width:`${pct}%`}}/></div><small>{r.vendor} · {r.routeType}</small></div>
          })}
        </div>
      </section>
    </div>
  </div>
}

function PageIntro({eyebrow,title,desc,action}) {
  return <div className="page-intro"><div><div className="eyebrow">{eyebrow}</div><h1>{title}</h1><p>{desc}</p></div>{action}</div>
}
function PanelHeader({title,subtitle,action}) { return <div className="panel-header"><div><h2>{title}</h2>{subtitle&&<span>{subtitle}</span>}</div>{action&&<div>{typeof action==="string"?<button className="filter-btn">{action}<ChevronDown size={14}/></button>:action}</div>}</div> }

function ShiftTesting({rows,requirements,search,onSelect}) {
  const [shift,setShift]=useState("Shift 1");
  const filtered=rows.filter(r=>r.shift===shift && matchesSearch(r,search));
  const groups=groupRows(filtered);
  const completed=groups.filter(g=>g.delivered>=2).length;

  return (
    <div>
      <PageIntro
        eyebrow="SUPPORT OPERATIONS"
        title="Shift Testing"
        desc="One row represents one destination/network/supplier testing group for the selected shift."
        action={
          <div className="shift-tabs">
            {["Shift 1","Shift 2","Shift 3"].map(s=>(
              <button key={s} className={shift===s?"selected":""} onClick={()=>setShift(s)}>
                {s}
                <small>{s==="Shift 1"?"00:00–08:00":s==="Shift 2"?"08:00–16:00":"16:00–00:00"}</small>
              </button>
            ))}
          </div>
        }
      />
      <div className="progress-panel">
        <div>
          <span>Shift completion</span>
          <strong>{completed} / {groups.length || requirements.filter(r=>r.active).length} groups evaluated</strong>
        </div>
        <div className="big-progress">
          <i style={{width:`${groups.length?completed/groups.length*100:0}%`}}/>
        </div>
      </div>
      <div className="panel table-panel">
        <PanelHeader
          title={`${shift} testing queue`}
          subtitle={`${filtered.length} individual imported attempts · ${groups.length} grouped records`}
          action={<button className="filter-btn"><Filter size={14}/> Filters</button>}
        />
        <GroupTable groups={groups} onSelect={onSelect}/>
      </div>
    </div>
  );
}

function ResultsExplorer({rows,search,onSelect}) {
  const [status,setStatus] = useState("ALL");
  const filtered = rows.filter(
    r => (status === "ALL" || r.dlrStatus === status) && matchesSearch(r,search)
  );
  const groups = groupRows(filtered);

  return (
    <div>
      <PageIntro
        eyebrow="DATA EXPLORER"
        title="Results Explorer"
        desc="Inspect raw testing-tool attempts while keeping the operational view grouped."
        action={<button className="filter-btn"><Download size={15}/> Export view</button>}
      />
      <div className="filter-row">
        <div className="segmented">
          {["ALL","DELIVERED","EXPIRED","UNDELIVERABLE","NO_DLR_RECEIVED"].map(s=>(
            <button
              className={status===s?"on":""}
              key={s}
              onClick={()=>setStatus(s)}
            >
              {s === "ALL" ? "All" : pretty(s)}
            </button>
          ))}
        </div>
        <span>{filtered.length} attempts</span>
      </div>
      <div className="panel table-panel">
        <GroupTable groups={groups} onSelect={onSelect}/>
      </div>
    </div>
  );
}

function GroupTable({ groups, onSelect, compact = false }) {
  return (
    <div className="data-table-wrap">
      <table className="data-table">
        <thead>
          <tr>
            <th>Destination / Network</th>
            <th>MCC/MNC</th>
            <th>Supplier</th>
            <th>Route</th>
            <th>Attempts</th>
            <th>Outcome mix</th>
            <th>Avg DLR</th>
            <th>Evaluation</th>
            <th></th>
          </tr>
        </thead>

        <tbody>
          {groups.map((g) => (
            <tr key={g.key} onClick={() => onSelect(g)}>
              <td>
                <div className="route-cell">
                  <strong>{g.country}</strong>
                  <span>
                    {g.network} · Sender {g.senderId || "—"}
                  </span>
                </div>
              </td>

              <td>
                <code>
                  {g.mcc}/{g.mnc}
                </code>
              </td>

              <td>
                <span className="supplier">{g.supplier}</span>
              </td>

              <td>{g.routeType}</td>

              <td>
                <strong>{g.total}</strong>
              </td>

              <td>
                <div className="outcome-chips">
                  <span className="chip delivered">
                    ✓ {g.delivered}
                  </span>

                  {g.failed > 0 && (
                    <span className="chip failed">
                      × {g.failed}
                    </span>
                  )}

                  {g.waiting > 0 && (
                    <span className="chip waiting">
                      ◷ {g.waiting}
                    </span>
                  )}

                  {g.other > 0 && (
                    <span className="chip other">
                      • {g.other}
                    </span>
                  )}
                </div>
              </td>

              <td>
                {g.avgDelay ? `${g.avgDelay}s` : "—"}
              </td>

              <td>
                <StatusBadge
                  status={
                    g.delivered >= 2
                      ? "working"
                      : g.total === 0
                        ? "pending"
                        : "attention"
                  }
                />
              </td>

              <td>
                <ChevronRight size={16} className="row-chevron" />
              </td>
            </tr>
          ))}

          {!groups.length && (
            <tr>
              <td colSpan="9">
                <div className="empty">
                  No testing records match the current filters.
                </div>
              </td>
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );
}

function StatusBadge({status}) {
  const map={working:["Working","good"],attention:["Needs attention","bad"],pending:["Pending","warn"]};
  const [label,cls]=map[status]||map.pending;
  return <span className={`status ${cls}`}><i/>{label}</span>
}

function GroupDrawer({group,onClose}) {
  return <div className="drawer-overlay" onClick={onClose}><aside className="drawer" onClick={e=>e.stopPropagation()}>
    <div className="drawer-head"><div><div className="eyebrow">TESTING GROUP</div><h2>{group.country} / {group.network}</h2><span>{group.supplier} · {group.routeType} · {group.mcc}/{group.mnc}</span></div><button className="icon-btn" onClick={onClose}><X size={18}/></button></div>
    <div className="drawer-summary"><div><small>Total attempts</small><strong>{group.total}</strong></div><div><small>Delivered</small><strong className="green-text">{group.delivered}</strong></div><div><small>Avg DLR</small><strong>{group.avgDelay?`${group.avgDelay}s`:"—"}</strong></div></div>
    <div className="rule-callout"><ShieldCheck size={17}/><div><strong>Operational rule</strong><span>2 or more delivered attempts = Working. Extra attempts remain part of the same shift group.</span></div></div>
    <h3>Outcome breakdown</h3>
    <div className="outcome-grid">{Object.entries(group.statusCounts).map(([k,v])=><div key={k}><span>{pretty(k)}</span><strong>{v}</strong></div>)}</div>
    <h3>Individual attempts</h3>
    <div className="attempt-list">{group.rows.map(r=><div className="attempt" key={r.id}><div><strong>{r.testId}</strong><span>{r.phone || "No phone"} · {r.timestamp}</span></div><div><StatusBadge status={r.dlrStatus==="DELIVERED"?"working":"attention"}/><small>{r.delay?`${r.delay}s DLR`:"No DLR"}</small></div></div>)}</div>
  </aside></div>
}

function Destinations({requirements,setRequirements,onAddActivity}) {
  const toggle=(id)=>setRequirements(rs=>rs.map(r=>r.id===id?{...r,active:!r.active}:r));
  return <div><PageIntro eyebrow="CONFIGURATION" title="Destinations" desc="Master list of Sales-defined recurring testing requirements." action={<button className="primary-btn" onClick={()=>window.dispatchEvent(new Event("open-requirement"))}><Plus size={16}/> Add requirement</button>}/>
    <RequirementList requirements={requirements} onToggle={toggle} onEdit={(r)=>window.dispatchEvent(new CustomEvent("edit-requirement",{detail:r}))}/>
    <RequirementEvents /></div>
}

function SalesManagement({requirements,setRequirements,onAddActivity}) {
  const [modal,setModal]=useState(false), [edit,setEdit]=useState(null);
  const save=req=>{setRequirements(rs=>edit?rs.map(x=>x.id===req.id?req:x):[...rs,{...req,id:`r-${Date.now()}`}]);onAddActivity(`${edit?"Updated":"Added"} Sales testing requirement: ${req.country} / ${req.network} / ${req.vendor}`);setModal(false);setEdit(null)};
  return <div><PageIntro eyebrow="SALES CONTROL" title="Sales Management" desc="Sales owns the recurring testing requirements. Support sees changes automatically on its next shift." action={<button className="primary-btn" onClick={()=>setModal(true)}><Plus size={16}/> New requirement</button>}/>
    <div className="sales-banner"><div className="banner-icon"><Settings2 size={20}/></div><div><strong>Single source of truth</strong><span>Changes here update the Support testing queue without messages, spreadsheets or manual handover.</span></div><span className="live-pill"><CircleDot size={10}/> Live</span></div>
    <RequirementList requirements={requirements} onToggle={id=>setRequirements(rs=>rs.map(r=>r.id===id?{...r,active:!r.active}:r))} onEdit={r=>{setEdit(r);setModal(true)}} onDelete={r=>{if(confirm("Delete this requirement?")){setRequirements(rs=>rs.filter(x=>x.id!==r.id));onAddActivity(`Deleted testing requirement: ${r.country} / ${r.network} / ${r.vendor}`)}}}/>
    {modal&&<RequirementModal req={edit} onClose={()=>{setModal(false);setEdit(null)}} onSave={save}/>}
  </div>
}

function RequirementList({requirements,onToggle,onEdit,onDelete}) {
  return <div className="panel table-panel"><div className="panel-header"><div><h2>Testing requirements</h2><span>{requirements.filter(r=>r.active).length} active · {requirements.length} total</span></div><button className="filter-btn"><Filter size={14}/> Filter</button></div><div className="req-list">{requirements.map(r=><div className="req-row" key={r.id}><div className={`req-status ${r.active?"on":"off"}`}><i/></div><div className="req-main"><strong>{r.country} <span>·</span> {r.network}</strong><span>MCC {r.mcc} / MNC {r.mnc} · Sender ID {r.senderId || "—"}</span></div><div><strong>{r.vendor}</strong><span className="muted">{r.routeType}</span></div><div><span className="freq">{r.frequency}</span></div><div className="req-actions"><button onClick={()=>onToggle(r.id)}>{r.active?"Disable":"Enable"}</button>{onEdit&&<button onClick={()=>onEdit(r)}><Edit3 size={14}/></button>}{onDelete&&<button className="danger-icon" onClick={()=>onDelete(r)}><Trash2 size={14}/></button>}</div></div>)}</div></div>
}

function RequirementEvents() {
  useEffect(()=>{},[]);
  return null;
}

function RequirementModal({req,onClose,onSave}) {
  const [form,setForm]=useState(req||{country:"",network:"",mcc:"",mnc:"",senderId:"",vendor:"",routeType:"HQ",frequency:"Every Shift",active:true});
  const set=(k,v)=>setForm(f=>({...f,[k]:v}));
  return <div className="modal-overlay"><div className="modal"><div className="modal-head"><div><div className="eyebrow">SALES CONFIGURATION</div><h2>{req?"Edit":"Add"} testing requirement</h2></div><button className="icon-btn" onClick={onClose}><X size={18}/></button></div>
    <div className="form-grid"><label>Destination<input value={form.country} onChange={e=>set("country",e.target.value)} placeholder="e.g. United States"/></label><label>Network<input value={form.network} onChange={e=>set("network",e.target.value)} placeholder="e.g. AT&T"/></label><label>MCC<input value={form.mcc} onChange={e=>set("mcc",e.target.value)} placeholder="310"/></label><label>MNC<input value={form.mnc} onChange={e=>set("mnc",e.target.value)} placeholder="410"/></label><label>Sender ID / SID<input value={form.senderId} onChange={e=>set("senderId",e.target.value)} placeholder="Available export field"/></label><label>HQ Vendor / Supplier<input value={form.vendor} onChange={e=>set("vendor",e.target.value)} placeholder="Vendor name"/></label><label>Route Type<select value={form.routeType} onChange={e=>set("routeType",e.target.value)}><option>HQ</option><option>Direct</option><option>SS7</option><option>SIM</option><option>Local</option><option>Local Bypass</option></select></label><label>Frequency<select value={form.frequency} onChange={e=>set("frequency",e.target.value)}><option>Every Shift</option><option>Daily</option><option>On Demand</option></select></label></div>
    <label className="toggle-line"><input type="checkbox" checked={form.active} onChange={e=>set("active",e.target.checked)}/><span>Active requirement</span></label>
    <div className="modal-foot"><button className="secondary-btn" onClick={onClose}>Cancel</button><button className="primary-btn" onClick={()=>onSave(form)} disabled={!form.country||!form.network||!form.vendor}>Save requirement</button></div>
  </div></div>
}

function ImportModal({onClose,onImport}) {
  const [file,setFile]=useState(null);
  return <div className="modal-overlay"><div className="modal import-modal"><div className="modal-head"><div><div className="eyebrow">RESULT INGESTION</div><h2>Import testing-tool results</h2><p>Supports .xlsx, .xls and .csv exports.</p></div><button className="icon-btn" onClick={onClose}><X size={18}/></button></div>
    <label className="dropzone" onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();setFile(e.dataTransfer.files[0])}}><input type="file" accept=".xlsx,.xls,.csv" onChange={e=>setFile(e.target.files[0])}/><Upload size={28}/><strong>{file?file.name:"Drop the exported Excel here"}</strong><span>{file?"Ready to import":"or click to browse your testing-tool export"}</span></label>
    <div className="import-note"><Database size={16}/><span>Rows are kept as individual attempts and automatically grouped by shift, country, network, MCC/MNC, supplier, route type and Sender ID.</span></div>
    <div className="modal-foot"><button className="secondary-btn" onClick={onClose}>Cancel</button><button className="primary-btn" disabled={!file} onClick={()=>onImport(file)}><Upload size={15}/> Process import</button></div>
  </div></div>
}

function ActivityLog({activity}) {
  return <div><PageIntro eyebrow="AUDIT TRAIL" title="Activity Log" desc="A simple record of configuration changes and result imports in this prototype."/><div className="panel activity-panel">{activity.map(a=><div className="activity-row" key={a.id}><div className="activity-dot"><Activity size={14}/></div><div><strong>{a.text}</strong><span>{a.actor} · {a.time}</span></div></div>)}</div></div>
}

function matchesSearch(r,s){if(!s)return true;return [r.country,r.network,r.supplier,r.senderId,r.mcc,r.mnc,r.routeType,r.dlrStatus].join(" ").toLowerCase().includes(s.toLowerCase())}
function groupRows(rows){
  const map=new Map();
  rows.forEach(r=>{
    const key=[r.country,r.network,r.mcc,r.mnc,r.supplier,r.routeType,r.senderId].join("|");
    if(!map.has(key))map.set(key,{key,country:r.country,network:r.network,mcc:r.mcc,mnc:r.mnc,supplier:r.supplier,routeType:r.routeType,senderId:r.senderId,rows:[],total:0,delivered:0,failed:0,waiting:0,rejected:0,other:0,statusCounts:{},avgDelay:0});
    const g=map.get(key);g.rows.push(r);g.total++;
    const s=r.dlrStatus;g.statusCounts[s]=(g.statusCounts[s]||0)+1;
    if(s==="DELIVERED")g.delivered++; else if(["EXPIRED","UNDELIVERABLE"].includes(s))g.failed++; else if(["NO_DLR_RECEIVED","PENDING","WAITING"].includes(s))g.waiting++; else if(s==="REJECTED")g.rejected++; else g.other++;
  });
  return [...map.values()].map(g=>{const delays=g.rows.map(r=>r.delay).filter(n=>Number.isFinite(n));g.avgDelay=delays.length?Math.round(delays.reduce((a,b)=>a+b,0)/delays.length*10)/10:0;return g});
}
function sameRequirement(g,r){return g.country===r.country&&g.network===r.network&&g.mcc===r.mcc&&g.mnc===r.mnc&&g.supplier===r.vendor&&g.senderId===r.senderId}
function pretty(s){return String(s||"").replaceAll("_"," ").toLowerCase().replace(/\b\w/g,c=>c.toUpperCase())}
function buildTrend(rows){const a={};rows.forEach(r=>{const h=String(r.timestamp).match(/(\d{2}):/);const hour=h?`${h[1]}:00`:"—";a[hour]=(a[hour]||0)+1});return Object.entries(a).sort().map(([hour,attempts])=>({hour,attempts}))}

React.createElement;
createRoot(document.getElementById("root")).render(<App/>);
