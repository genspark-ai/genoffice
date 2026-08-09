import React, { useEffect, useMemo, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { App as CapacitorApp } from '@capacitor/app'
import { Keyboard } from '@capacitor/keyboard'
import { StatusBar, Style } from '@capacitor/status-bar'
import { AI_PROVIDERS, chatForProvider, defaultAiSettings } from '@genoffice/ai-provider'
import type { AiProviderId, AiSettings } from '@genoffice/ai-provider'
import { DocsEditorScreen } from './docs-editor'
import { SheetsEditorScreen } from './sheets-editor'
import { SlidesEditorScreen } from './slides-editor'
import { PdfEditorScreen } from './pdf-editor'
import './styles.css'
import './android-editors.css'
import './android-mobile-shell.css'

type Screen = 'home' | 'docs' | 'sheets' | 'slides' | 'pdf' | 'ai'
const SETTINGS_KEY = 'genoffice.android.ai.settings'

function loadSettings(): AiSettings { const defaults=defaultAiSettings(); try { const raw=localStorage.getItem(SETTINGS_KEY); if(!raw)return defaults; const saved=JSON.parse(raw) as Partial<AiSettings>; return {provider:saved.provider??defaults.provider,providers:{...defaults.providers,...(saved.providers??{})}} } catch { return defaults } }

function AiPanel(){
 const [settings,setSettings]=useState<AiSettings>(()=>loadSettings()); const [message,setMessage]=useState('Write a short professional leave letter.'); const [answer,setAnswer]=useState(''); const [busy,setBusy]=useState(false); const [error,setError]=useState('');
 const meta=useMemo(()=>AI_PROVIDERS.find(p=>p.id===settings.provider)??AI_PROVIDERS[0],[settings.provider]); const config=settings.providers[settings.provider];
 const persist=(next:AiSettings)=>{setSettings(next);localStorage.setItem(SETTINGS_KEY,JSON.stringify(next))}; const setProvider=(provider:AiProviderId)=>{const nextMeta=AI_PROVIDERS.find(p=>p.id===provider)!;const providers={...settings.providers};if(!providers[provider]?.model)providers[provider]={apiKey:'',model:nextMeta.defaultModel,baseUrl:''};persist({...settings,provider,providers})}; const updateConfig=(patch:Partial<typeof config>)=>persist({...settings,providers:{...settings.providers,[settings.provider]:{...config,...patch}}});
 const send=async()=>{setBusy(true);setError('');setAnswer('');try{if(!config.apiKey)throw new Error('Enter an API key first.');const result=await chatForProvider(settings.provider,config,'You are a helpful office assistant. Return only the requested result.',message);if(!result.ok)throw new Error(result.error??'AI request failed');setAnswer(result.content??'')}catch(e){setError(e instanceof Error?e.message:String(e))}finally{setBusy(false)}};
 return <div className="ai-panel"><h1>AI Assistant</h1><p className="muted">Shared GenOffice provider layer across Android editors.</p><label>Provider</label><select value={settings.provider} onChange={e=>setProvider(e.target.value as AiProviderId)}>{AI_PROVIDERS.filter(p=>p.id!=='genspark').map(p=><option key={p.id} value={p.id}>{p.label}</option>)}</select><label>Model</label>{meta.models.length?<select value={config.model} onChange={e=>updateConfig({model:e.target.value})}>{meta.models.map(model=><option key={model} value={model}>{model}</option>)}</select>:<input value={config.model} onChange={e=>updateConfig({model:e.target.value})} placeholder="Model name"/>}{meta.needsBaseUrl&&<><label>Base URL</label><input value={config.baseUrl??''} onChange={e=>updateConfig({baseUrl:e.target.value})} placeholder="https://your-provider.example/v1"/></>}<label>API key</label><input type="password" value={config.apiKey} onChange={e=>updateConfig({apiKey:e.target.value})} placeholder={meta.keyPlaceholder} autoComplete="off"/><label>Request</label><textarea value={message} onChange={e=>setMessage(e.target.value)} rows={5}/><button className="primary" disabled={busy} onClick={()=>void send()}>{busy?'Thinking…':'Send to AI'}</button>{error&&<div className="notice warning"><strong>Request failed</strong><span>{error}</span></div>}{answer&&<div className="answer"><strong>Response</strong><pre>{answer}</pre></div>}</div>
}

function EditorChrome({active,children}:{active:Exclude<Screen,'home'|'ai'>;children:React.ReactNode}){
 const [tools,setTools]=useState(false); const titles={docs:'Document',sheets:'Workbook',slides:'Presentation',pdf:'PDF'}; const subtitles={docs:'GenOffice Docs',sheets:'GenOffice Sheets',slides:'GenOffice Slides',pdf:'GenOffice PDF'}; const actions={docs:['Edit','Insert','Format','More'],sheets:['Edit','Format','Insert','More'],slides:['Text','Image','Shape','More'],pdf:['Annotate','Draw','Pages','More']}[active];
 const tool=(action:string,index:number)=>{if(action==='More')setTools(v=>!v);else window.dispatchEvent(new CustomEvent('genoffice-android-tool',{detail:{kind:active,action,index}}))};
 return <div className="android-mobile-editor"><header className="android-mobile-top"><button className="android-mobile-back" aria-label="Back" onClick={()=>window.dispatchEvent(new Event('genoffice-android-back'))}>‹</button><div className="android-mobile-title">{titles[active]}<span className="android-mobile-subtitle">{subtitles[active]}</span></div><button className="android-mobile-menu" aria-label="More options" onClick={()=>setTools(v=>!v)}>⋮</button></header><div className="android-mobile-canvas">{children}<div className={`android-mobile-tools ${tools?'open':''}`}>{actions.map((x,i)=><button key={x} onClick={()=>tool(x,i)}>{x}</button>)}</div></div><nav className="android-mobile-bottom">{actions.map((x,i)=><button key={x} className={i===0?'active':''} onClick={()=>tool(x,i)}><b>{['⌕','＋','✎','⋯'][i]}</b><span>{x}</span></button>)}</nav></div>
}

function MobileShell(){
 const [active,setActive]=useState<Screen>('home');
 useEffect(()=>{void StatusBar.setStyle({style:Style.Light}).catch(()=>{});void Keyboard.setAccessoryBarVisible({isVisible:false}).catch(()=>{});const back=CapacitorApp.addListener('backButton',({canGoBack})=>{if(active!=='home'){setActive('home');return}if(canGoBack)window.history.back()});const onBack=()=>setActive('home');window.addEventListener('genoffice-android-back',onBack);return()=>{void back.then(h=>h.remove());window.removeEventListener('genoffice-android-back',onBack)}},[active]);
 const open=(screen:Screen)=>setActive(screen); const isEditor=active==='docs'||active==='sheets'||active==='slides'||active==='pdf';
 const editor=active==='docs'?<DocsEditorScreen/>:active==='sheets'?<SheetsEditorScreen/>:active==='slides'?<SlidesEditorScreen/>:<PdfEditorScreen/>;
 return <main className={`mobile-shell ${isEditor?'editor-mode':''}`}>{!isEditor&&<header className="topbar"><div><div className="brand">GenOffice</div><div className="subtitle">Android</div></div><button className="settings" aria-label="AI settings" onClick={()=>open('ai')}>⚙</button></header>}{isEditor?<EditorChrome active={active}>{editor}</EditorChrome>:<section className="content">{active==='home'?<><h1>Office, wherever you are.</h1><p className="muted">One Android shell with the shared GenOffice Docs, Sheets, Slides and PDF engines.</p><div className="cards"><button className="app-card" onClick={()=>open('docs')}><span className="icon">DOCX</span><span><strong>Docs</strong><small>Word editor</small></span><span className="arrow">›</span></button><button className="app-card" onClick={()=>open('sheets')}><span className="icon">XLSX</span><span><strong>Sheets</strong><small>Spreadsheet editor</small></span><span className="arrow">›</span></button><button className="app-card" onClick={()=>open('slides')}><span className="icon">PPTX</span><span><strong>Slides</strong><small>Presentation editor</small></span><span className="arrow">›</span></button><button className="app-card" onClick={()=>open('pdf')}><span className="icon">PDF</span><span><strong>PDF</strong><small>Viewer and markup workspace</small></span><span className="arrow">›</span></button></div><button className="ai-card" onClick={()=>open('ai')}><strong>AI Assistant</strong><span>Use the shared provider configuration across the suite.</span><b>Open →</b></button></>:<AiPanel/>}</section>}{!isEditor&&<nav className="bottom-nav"><button className={active==='home'?'active':''} onClick={()=>open('home')}>Home</button><button onClick={()=>open('docs')}>Docs</button><button onClick={()=>open('sheets')}>Sheets</button><button onClick={()=>open('slides')}>Slides</button><button className={active==='ai'?'active':''} onClick={()=>open('ai')}>AI</button></nav>}</main>
}

createRoot(document.getElementById('root')!).render(<React.StrictMode><MobileShell/></React.StrictMode>)
