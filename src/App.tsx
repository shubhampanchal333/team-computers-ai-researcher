/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from 'react';
import { Search, Building2, Users, Rocket, Cpu, Briefcase, ExternalLink, Mail, Phone, Linkedin, CheckCircle2, AlertCircle, Loader2, Globe, Server, Laptop, Shield, Zap, Target, TrendingUp, Copy, Check, Info, Database, Cloud, Layout, Lock, AppWindow, MessageSquare, X, ChevronRight, BarChart3, ChevronDown, ChevronUp, Terminal, ListFilter, Sparkles, FileText, ArrowRight, ShieldAlert, AlertTriangle, BarChart, Download } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { researchCompany, CompanyResearch, discoverCompanies, bulkSummarize, CompanySummary, getChatResponse } from './services/geminiService';
import { cn } from './lib/utils';
import * as XLSX from 'xlsx';

const TECH_STACK_ICONS: Record<string, any> = {
  cloud: Cloud,
  workspace: Layout,
  security: Lock,
  businessApps: AppWindow,
  other: Database
};

const USD_TO_INR_DEFAULT = 83.5;

const parseCurrency = (str: string): number => {
  if (!str) return 0;
  const numericString = str.replace(/[^0-9.]/g, '');
  return parseFloat(numericString) || 0;
};

export default function App() {
  const [activeTab, setActiveTab] = useState<'single' | 'bulk' | 'discovery' | 'ranking'>('single');
  const [exchangeRate, setExchangeRate] = useState(USD_TO_INR_DEFAULT);

  // Fetch live exchange rate on mount
  React.useEffect(() => {
    fetch('https://open.er-api.com/v6/latest/USD')
      .then(res => res.json())
      .then(data => {
        if (data && data.rates && data.rates.INR) {
          setExchangeRate(data.rates.INR);
        }
      })
      .catch(() => console.log('Using default exchange rate'));
  }, []);

  const [query, setQuery] = useState('');
  const [bulkInput, setBulkInput] = useState('');
  const [discoveryPrompt, setDiscoveryPrompt] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [result, setResult] = useState<CompanyResearch | null>(null);
  const [summaries, setSummaries] = useState<CompanySummary[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [isCopied, setIsCopied] = useState(false);
  const [recentScans, setRecentScans] = useState<string[]>(() => {
    const saved = localStorage.getItem('recent_scans');
    return saved ? JSON.parse(saved) : [];
  });
  const [scanHistory, setScanHistory] = useState<CompanyResearch[]>(() => {
    const saved = localStorage.getItem('scan_history');
    return saved ? JSON.parse(saved) : [];
  });
  const [loadingStatus, setLoadingStatus] = useState('');
  const [showCallScript, setShowCallScript] = useState(false);
  const [openSections, setOpenSections] = useState<string[]>(['introduction']);
  const [copiedSection, setCopiedSection] = useState<string | null>(null);
  const [completedSteps, setCompletedSteps] = useState<Set<string>>(new Set());
  const [savedQueries, setSavedQueries] = useState<string[]>(() => {
    const saved = localStorage.getItem('saved_queries');
    return saved ? JSON.parse(saved) : [];
  });
  const [chatHistory, setChatHistory] = useState<Array<{ role: 'user' | 'model'; content: string }>>([]);
  const [chatInput, setChatInput] = useState('');
  const [isChatOpen, setIsChatOpen] = useState(false);
  const [isChatLoading, setIsChatLoading] = useState(false);
  const chatEndRef = React.useRef<HTMLDivElement>(null);

  const scrollToBottom = () => {
    chatEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  };

  React.useEffect(() => {
    if (isChatOpen) scrollToBottom();
  }, [chatHistory, isChatOpen]);

  const handleChat = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!chatInput.trim() || !result || isChatLoading) return;

    const userMessage = { role: 'user' as const, content: chatInput };
    setChatHistory(prev => [...prev, userMessage]);
    setChatInput('');
    setIsChatLoading(true);

    try {
      const response = await getChatResponse([...chatHistory, userMessage], result);
      setChatHistory(prev => [...prev, { role: 'model', content: response }]);
    } catch (err) {
      setChatHistory(prev => [...prev, { role: 'model', content: "Failed to connect to the intelligence core. Please try again." }]);
    } finally {
      setIsChatLoading(false);
    }
  };

  const toggleStep = (stepId: string) => {
    setCompletedSteps(prev => {
      const next = new Set(prev);
      if (next.has(stepId)) next.delete(stepId);
      else next.add(stepId);
      return next;
    });
  };

  const saveCurrentQuery = () => {
    const queryToSave = activeTab === 'discovery' ? discoveryPrompt : query;
    if (!queryToSave.trim() || savedQueries.includes(queryToSave)) return;
    
    const newSaved = [queryToSave, ...savedQueries].slice(0, 10);
    setSavedQueries(newSaved);
    localStorage.setItem('saved_queries', JSON.stringify(newSaved));
  };

  const removeSavedQuery = (q: string) => {
    const newSaved = savedQueries.filter(sq => sq !== q);
    setSavedQueries(newSaved);
    localStorage.setItem('saved_queries', JSON.stringify(newSaved));
  };

  const exportToJSON = (data: any, filename: string) => {
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${filename}.json`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const exportToExcel = (data: any | any[], filename: string, type: 'single' | 'summary' | 'ranking' = 'summary') => {
    const wb = XLSX.utils.book_new();

    if (type === 'single' && !Array.isArray(data)) {
      const research = data as CompanyResearch;
      
      // Sheet 1: Basic Info
      const basicInfo = [{
        Name: research.name,
        Domain: research.domain,
        ARR: research.arr,
        TeamSize: research.teamSize,
        CustomerBase: research.customerBase,
        FundingRound: research.funding.round,
        TotalFunding: research.funding.totalAmount,
        LastFundingDate: research.funding.lastFundingDate || 'N/A',
        StrategicAdvantage: research.teamComputersPitch.strategicAdvantage,
        Deadpooled: research.isDeadpooled ? 'Yes' : 'No'
      }];
      const wsBasic = XLSX.utils.json_to_sheet(basicInfo);
      XLSX.utils.book_append_sheet(wb, wsBasic, "Company Overview");

      // Sheet 2: Founders & Leadership
      const people: any[] = [];
      research.founders.forEach(f => {
        people.push({
          Type: 'Founder',
          Name: f.name,
          Role: f.role,
          LinkedIn: f.linkedin,
          Email: f.email,
          Phone: f.phone,
          ReportsTo: 'Board'
        });
      });
      research.leadership.forEach(l => {
        people.push({
          Type: 'Leadership',
          Name: l.name,
          Role: l.title,
          LinkedIn: l.linkedin,
          Email: 'N/A',
          Phone: 'N/A',
          ReportsTo: l.reportsTo || 'CEO'
        });
      });
      const wsPeople = XLSX.utils.json_to_sheet(people);
      XLSX.utils.book_append_sheet(wb, wsPeople, "Key Contacts");

      // Sheet 3: Tech Stack & Spend
      const techStack: any[] = [];
      Object.entries(research.techStackSpend).forEach(([cat, data]) => {
        data.items.forEach(item => {
          const usdMonthly = parseCurrency(item.spend);
          techStack.push({
            Category: cat.toUpperCase(),
            Technology: item.name,
            'Monthly Spend (USD)': `$${usdMonthly.toLocaleString()}`,
            'Monthly Spend (INR)': `₹${(usdMonthly * exchangeRate).toLocaleString()}`,
            'Yearly Spend (USD)': `$${(usdMonthly * 12).toLocaleString()}`,
            'Yearly Spend (INR)': `₹${(usdMonthly * 12 * exchangeRate).toLocaleString()}`,
            CategorySubtotal: data.subtotal
          });
        });
      });
      const wsTech = XLSX.utils.json_to_sheet(techStack);
      XLSX.utils.book_append_sheet(wb, wsTech, "Tech Stack & Spend");

      // Sheet 4: IT Infrastructure
      const infra = [
        { Category: 'Endpoints', Item: 'Laptops', Value: research.itInfrastructure.endpoints.laptops },
        { Category: 'Endpoints', Item: 'Desktops', Value: research.itInfrastructure.endpoints.desktops },
        { Category: 'Endpoints', Item: 'Mobile Devices', Value: research.itInfrastructure.endpoints.mobile },
        { Category: 'Compute', Item: 'Servers', Value: research.itInfrastructure.compute.servers },
        { Category: 'Compute', Item: 'Storage', Value: research.itInfrastructure.compute.storage },
        { Category: 'Compute', Item: 'Networking', Value: research.itInfrastructure.compute.networking },
        { Category: 'OEMs', Item: 'Current OEMs', Value: research.itInfrastructure.currentOEMs.join(', ') },
        { Category: 'OEMs', Item: 'Interested OEMs', Value: research.itInfrastructure.interestedOEMs.join(', ') },
        { Category: 'Analysis', Item: 'OEM Analysis', Value: research.itInfrastructure.oemAnalysis },
        { Category: 'Audit', Item: 'Full Stack Audit', Value: research.itInfrastructure.fullStackAudit }
      ];
      const wsInfra = XLSX.utils.json_to_sheet(infra);
      XLSX.utils.book_append_sheet(wb, wsInfra, "IT Infrastructure");

      // Sheet 5: Opportunities & Upsell
      const opportunities: any[] = [];
      research.teamComputersPitch.useCases.forEach(u => {
        opportunities.push({
          Type: 'Strategic Use Case',
          Title: u.title,
          Category: u.serviceCategory,
          Description: u.description,
          ValueBenefit: u.keyTakeaway,
          ROI: u.roiProjection,
          Timeline: u.implementationTimeline
        });
      });
      research.teamComputersPitch.upsellCrossSellOpportunities.forEach(u => {
        opportunities.push({
          Type: 'Upsell/Cross-sell',
          Title: u.title,
          Category: u.targetService,
          Description: u.description,
          ValueBenefit: 'High Interest',
          ROI: 'Immediate',
          Timeline: '0-3 Months'
        });
      });
      research.painPoints.forEach(p => {
        opportunities.push({
          Type: 'Pain Point to Address',
          Title: p.point,
          Category: 'Consulting',
          Description: p.impact,
          ValueBenefit: 'Efficiency Gain',
          ROI: 'Varies',
          Timeline: 'Immediate'
        });
      });
      const wsOpp = XLSX.utils.json_to_sheet(opportunities);
      XLSX.utils.book_append_sheet(wb, wsOpp, "Growth Opportunities");

      // Sheet 6: Competitive Landscape
      const competitive = research.competitiveLandscape.map(c => ({
        Competitor: c.competitorName,
        TheirStrength: c.serviceBenefit,
        TeamComputersEdge: c.teamComputersEdge,
        CompetitorTech: c.techStack.join(', '),
        CompetitorServices: c.services.join(', ')
      }));
      const wsComp = XLSX.utils.json_to_sheet(competitive);
      XLSX.utils.book_append_sheet(wb, wsComp, "Competitive Intel");

      // Sheet 7: Sales Strategy
      const strategy = [
        { Section: 'Introduction', Content: research.teamComputersPitch.callScript.introduction },
        { Section: 'Discovery Questions', Content: research.teamComputersPitch.callScript.discoveryQuestions.join(' | ') },
        { Section: 'Value Proposition', Content: research.teamComputersPitch.callScript.valueProposition },
        { Section: 'Handling Objections', Content: research.teamComputersPitch.callScript.handlingObjections.join(' | ') },
        { Section: 'Closing', Content: research.teamComputersPitch.callScript.closing },
        { Section: 'Pitch Slide 1', Content: research.pitchDeck.introduction },
        { Section: 'Pitch Slide 2', Content: research.pitchDeck.problemStatement },
        { Section: 'Pitch Slide 3', Content: research.pitchDeck.solution },
        { Section: 'Pitch Slide 4', Content: research.pitchDeck.marketAnalysis },
        { Section: 'Pitch Slide 5', Content: research.pitchDeck.competitiveAdvantage },
        { Section: 'Pitch Slide 6', Content: research.pitchDeck.proposedSolution },
        { Section: 'Pitch Slide 7', Content: research.pitchDeck.caseStudies.join(' | ') },
        { Section: 'Pitch Slide 8', Content: research.pitchDeck.callToAction }
      ];
      const wsStrat = XLSX.utils.json_to_sheet(strategy);
      XLSX.utils.book_append_sheet(wb, wsStrat, "Sales Playbook");

    } else if (Array.isArray(data)) {
      // For bulk, discovery, or ranking
      const ws = XLSX.utils.json_to_sheet(data);
      XLSX.utils.book_append_sheet(wb, ws, "Data");
    }

    XLSX.writeFile(wb, `${filename}.xlsx`);
  };

  const toggleSection = (section: string) => {
    setOpenSections(prev => 
      prev.includes(section) ? prev.filter(s => s !== section) : [...prev, section]
    );
  };

  const copySection = (text: string, section: string) => {
    navigator.clipboard.writeText(text);
    setCopiedSection(section);
    setTimeout(() => setCopiedSection(null), 2000);
  };

  const getMonthlySpend = (company: CompanyResearch | CompanySummary): number => {
    if ('techStackSpend' in company) {
      const spend = (company as CompanyResearch).techStackSpend;
      return [spend.cloud, spend.workspace, spend.security, spend.businessApps, spend.other]
        .reduce((acc, cat) => acc + parseCurrency(cat.subtotal), 0);
    }
    return parseCurrency((company as CompanySummary).estimatedMonthlyTechSpend || '0');
  };

  const getPropensityScore = (company: CompanyResearch | CompanySummary): number => {
    if ('serviceInterest' in company) {
      const ci = (company as CompanyResearch).serviceInterest;
      return Math.round(
        ((ci.infrastructure || 0) + 
         (ci.cloud || 0) + 
         (ci.managedServices || 0) + 
         (ci.digitalTransformation || 0) + 
         (ci.businessApps || 0)) / 5
      );
    }
    return 50; // Neutral score for summary objects
  };

  const getTopInterest = (company: CompanyResearch | CompanySummary): string => {
    if ('serviceInterest' in company) {
      const ci = (company as CompanyResearch).serviceInterest;
      return Object.entries(ci)
        .sort(([, a], [, b]) => b - a)[0][0]
        .replace(/([A-Z])/g, ' $1');
    }
    return 'Consulting'; // Default for summary objects
  };

  const statusMessages = [
    "Initializing High-Speed Intelligence Core...",
    "Scanning Global Registries & Financial Disclosures...",
    "Mapping Leadership Hierarchy & Reporting Lines...",
    "Analyzing Infrastructure Footprints & Tech Stack Spend...",
    "Evaluating Competitive Landscape & Market Edge...",
    "Synthesizing Strategic ROI & Use-Cases...",
    "Generating Slide-by-Slide Pitch Strategy...",
    "Finalizing Intelligence Report..."
  ];

  const handleSearch = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!query.trim()) return;

    setIsLoading(true);
    setError(null);
    setResult(null);
    
    let statusIdx = 0;
    const statusInterval = setInterval(() => {
      setLoadingStatus(statusMessages[statusIdx % statusMessages.length]);
      statusIdx++;
    }, 1500);

    try {
      const data = await researchCompany(query);
      setResult(data);
      
      const newScans = [data.name, ...recentScans.filter(s => s !== data.name)].slice(0, 5);
      setRecentScans(newScans);
      localStorage.setItem('recent_scans', JSON.stringify(newScans));

      const newHistory = [data, ...scanHistory.filter(s => s.domain !== data.domain)].slice(0, 30);
      setScanHistory(newHistory);
      localStorage.setItem('scan_history', JSON.stringify(newHistory));
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An unexpected error occurred');
    } finally {
      clearInterval(statusInterval);
      setIsLoading(false);
    }
  };

  const handleBulkProcess = async (e: React.FormEvent) => {
    e.preventDefault();
    const items = bulkInput.split(/[\n,]+/).map(d => d.trim()).filter(d => d.length > 0);
    if (items.length === 0) return;

    setIsLoading(true);
    setError(null);
    setResult(null);
    setSummaries([]);
    setLoadingStatus("Initializing Quantum Batch Engine...");

    try {
      // Optimized for up to 500 companies
      // Batch size of 20 with 3 concurrent batches balance speed and stability
      const batchSize = 20;
      const batches = [];
      for (let i = 0; i < items.length; i += batchSize) {
        batches.push(items.slice(i, i + batchSize));
      }

      const allSummaries: CompanySummary[] = [];
      const concurrencyLimit = 3;
      
      for (let i = 0; i < batches.length; i += concurrencyLimit) {
        const currentGroup = batches.slice(i, i + concurrencyLimit);
        setLoadingStatus(`Processing Group ${Math.floor(i / concurrencyLimit) + 1} of ${Math.ceil(batches.length / concurrencyLimit)} (${allSummaries.length}/${items.length})...`);
        
        const results = await Promise.all(currentGroup.map(batch => bulkSummarize(batch)));
        allSummaries.push(...results.flat());
        
        // Brief pause to stabilize rate limits if processing many batches
        if (batches.length > concurrencyLimit) {
          await new Promise(resolve => setTimeout(resolve, 500));
        }
      }
      
      setSummaries(allSummaries);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Bulk processing failed');
    } finally {
      setIsLoading(false);
    }
  };

  const handleDiscovery = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!discoveryPrompt.trim()) return;

    setIsLoading(true);
    setError(null);
    setResult(null);
    setSummaries([]);
    setLoadingStatus("Discovering Opportunities...");

    try {
      const data = await discoverCompanies(discoveryPrompt);
      setSummaries(data);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Discovery failed');
    } finally {
      setIsLoading(false);
    }
  };

  const handleDeepScan = async (company: string) => {
    setQuery(company);
    setActiveTab('single');
    // Trigger search manually
    const fakeEvent = { preventDefault: () => {} } as React.FormEvent;
    setTimeout(() => handleSearch(fakeEvent), 100);
  };

  const copyToClipboard = () => {
    if (!result) return;
    const text = JSON.stringify(result, null, 2);
    navigator.clipboard.writeText(text);
    setIsCopied(true);
    setTimeout(() => setIsCopied(false), 2000);
  };

  return (
    <div className="h-screen flex flex-col bg-[#0F172A] text-slate-200 font-sans selection:bg-blue-500/30 overflow-hidden relative">
      {/* Background Grid & Glows */}
      <div className="fixed inset-0 bg-grid pointer-events-none opacity-20" />
      <div className="fixed top-[-10%] right-[-10%] w-[50%] h-[50%] bg-blue-600/10 blur-[120px] rounded-full pointer-events-none" />
      <div className="fixed bottom-[-10%] left-[-10%] w-[50%] h-[50%] bg-red-600/5 blur-[120px] rounded-full pointer-events-none" />

      {/* Header */}
      <header className="glass z-50 px-6 sm:px-8 h-16 flex items-center justify-between border-b border-white/5 shrink-0">
        <div className="w-full flex items-center justify-between">
          <div className="flex items-center gap-4">
            <div className="flex-shrink-0 min-w-[160px] flex items-center gap-4">
              <div className="relative group">
                <motion.img 
                  src="https://www.teamcomputers.com/wp-content/uploads/2023/11/Team-Logo.png"
                  alt="Team Computers Logo"
                  className="h-10 w-auto object-contain"
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  referrerPolicy="no-referrer"
                  onError={(e) => {
                    e.currentTarget.style.display = 'none';
                    e.currentTarget.nextElementSibling?.classList.remove('hidden');
                  }}
                />
                <span className="hidden text-xl font-brand font-black tracking-tight text-white leading-none whitespace-nowrap">
                  Team Computers
                </span>
              </div>
              
              <div className="h-6 w-px bg-white/20" />
              
              <div className="relative group">
                <motion.img 
                  src="https://upload.wikimedia.org/wikipedia/commons/thumb/7/7d/Intel_logo_%282020%2C_white%29.svg/1024px-Intel_logo_%282020%2C_white%29.svg.png"
                  alt="Intel Logo"
                  className="h-6 w-auto object-contain opacity-90 hover:opacity-100 transition-opacity"
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  referrerPolicy="no-referrer"
                  onError={(e) => {
                    e.currentTarget.style.display = 'none';
                    e.currentTarget.nextElementSibling?.classList.remove('hidden');
                  }}
                />
                <span className="hidden text-lg font-brand font-bold tracking-tight text-white leading-none">
                  Intel
                </span>
              </div>
            </div>
            <div className="flex flex-col border-l border-white/10 pl-4 ml-2">
              <span className="text-[10px] font-mono tracking-[0.2em] text-blue-400 uppercase font-bold">AI RESEARCHER v3.0</span>
            </div>
          </div>
          <div className="hidden md:flex items-center gap-6 text-[11px] font-mono tracking-widest text-slate-400 uppercase">
            <span className="flex items-center gap-1.5"><Shield className="w-3 h-3 text-blue-400" /> SECURE SCAN</span>
            <span className="flex items-center gap-1.5"><Zap className="w-3 h-3 text-yellow-400" /> REAL-TIME DATA</span>
          </div>
        </div>
      </header>

      <div className="flex-1 overflow-y-auto custom-scrollbar relative z-10">
        <main className="w-full px-6 sm:px-8 py-12">
        {/* Search Section */}
        <section className="mb-12">
          <div className="max-w-3xl mx-auto text-center mb-8">
            <motion.div
              initial={{ opacity: 0, scale: 0.9 }}
              animate={{ opacity: 1, scale: 1 }}
              className="inline-block px-4 py-1.5 rounded-full glass border-white/10 text-[10px] font-mono font-bold tracking-widest text-blue-400 uppercase mb-6"
            >
              Intelligence Engine Active
            </motion.div>
            <h2 className="text-4xl font-black text-white mb-6 sm:text-6xl tracking-tight font-brand">
              Strategic <span className="text-transparent bg-clip-text bg-gradient-to-r from-blue-400 to-blue-600">Intelligence</span> Terminal
            </h2>
            
            {/* Tabs */}
            <div className="flex flex-wrap items-center justify-center gap-2 mb-8">
                {[
                  {id: 'single', label: 'Single Scan', icon: Search},
                  {id: 'bulk', label: 'Bulk Entry', icon: ListFilter},
                  {id: 'discovery', label: 'Discovery', icon: Sparkles},
                  {id: 'ranking', label: 'Ranking', icon: BarChart}
                ].map(tab => (
                  <button
                    key={`tab-${tab.id}`}
                    onClick={() => {
                    setActiveTab(tab.id as any);
                    setError(null);
                  }}
                  className={cn(
                    "px-6 py-2.5 rounded-xl text-xs font-bold uppercase tracking-widest transition-all flex items-center gap-2 border",
                    activeTab === tab.id 
                      ? "bg-blue-600 text-white border-blue-500 shadow-[0_0_20px_rgba(59,130,246,0.3)]" 
                      : "glass border-white/5 text-slate-500 hover:text-slate-300 hover:border-white/10"
                  )}
                >
                  <tab.icon className="w-4 h-4" />
                  {tab.label}
                </button>
              ))}
            </div>
          </div>

          <div className="max-w-4xl mx-auto relative group">
            <div className="absolute -inset-1 bg-gradient-to-r from-blue-600 to-red-600 rounded-2xl blur opacity-20 group-focus-within:opacity-40 transition duration-1000 group-hover:duration-200" />
            
            <div className="relative glass rounded-2xl border-white/10 p-2">
              {activeTab === 'single' && (
                <form onSubmit={handleSearch} className="flex items-center">
                  <Search className="ml-4 text-slate-500" />
                  <input
                    type="text"
                    value={query}
                    onChange={(e) => setQuery(e.target.value)}
                    placeholder="Enter company name or domain..."
                    className="w-full bg-transparent border-none outline-none px-4 py-4 text-white text-lg placeholder:text-slate-600"
                  />
                  <div className="flex items-center gap-2 pr-2">
                    <button
                      type="button"
                      onClick={saveCurrentQuery}
                      title="Save Search"
                      className="p-2 hover:bg-white/10 rounded-xl text-slate-400 hover:text-yellow-400 transition-all"
                    >
                      <Zap className={cn("w-5 h-5", savedQueries.includes(query) && "fill-yellow-400 text-yellow-400")} />
                    </button>
                    <button
                      type="submit"
                      disabled={isLoading || !query.trim()}
                      className="px-8 py-4 bg-blue-600 text-white font-bold rounded-xl hover:bg-blue-500 disabled:opacity-50 disabled:cursor-not-allowed transition-all shadow-lg shadow-blue-600/20 flex items-center gap-2"
                    >
                      {isLoading ? <Loader2 className="w-5 h-5 animate-spin" /> : 'SCAN'}
                    </button>
                  </div>
                </form>
              )}

              {activeTab === 'bulk' && (
                <form onSubmit={handleBulkProcess} className="space-y-4 p-4">
                  <div className="flex items-center gap-3 mb-2 px-2">
                    <FileText className="w-4 h-4 text-blue-400" />
                    <span className="text-[10px] font-mono text-slate-500 uppercase tracking-widest">Enter up to 500 company names or domains (one per line or comma separated)</span>
                  </div>
                  <textarea
                    value={bulkInput}
                    onChange={(e) => setBulkInput(e.target.value)}
                    placeholder="e.g. Google, microsoft.com, Apple Inc, Reliance Industries..."
                    className="w-full bg-white/5 border border-white/10 rounded-xl p-4 text-white font-mono text-sm h-64 outline-none focus:border-blue-500/50 transition-colors custom-scrollbar"
                  />
                  <div className="flex justify-end">
                    <button
                      type="submit"
                      disabled={isLoading || !bulkInput.trim()}
                      className="px-8 py-3 bg-blue-600 text-white font-bold rounded-xl hover:bg-blue-500 disabled:opacity-50 disabled:cursor-not-allowed transition-all shadow-lg shadow-blue-600/20 flex items-center gap-2"
                    >
                      {isLoading ? <Loader2 className="w-5 h-5 animate-spin" /> : 'BATCH PROCESS'}
                    </button>
                  </div>
                </form>
              )}

              {activeTab === 'discovery' && (
                <form onSubmit={handleDiscovery} className="flex items-center">
                  <Sparkles className="ml-4 text-blue-400" />
                  <input
                    type="text"
                    value={discoveryPrompt}
                    onChange={(e) => setDiscoveryPrompt(e.target.value)}
                    placeholder="e.g., Top 50 Fintech companies in India with >100 employees..."
                    className="w-full bg-transparent border-none outline-none px-4 py-4 text-white text-lg placeholder:text-slate-600"
                  />
                  <div className="flex items-center gap-2 pr-2">
                    <button
                      type="button"
                      onClick={saveCurrentQuery}
                      title="Save Discovery"
                      className="p-2 hover:bg-white/10 rounded-xl text-slate-400 hover:text-yellow-400 transition-all"
                    >
                      <Zap className={cn("w-5 h-5", savedQueries.includes(discoveryPrompt) && "fill-yellow-400 text-yellow-400")} />
                    </button>
                    <button
                      type="submit"
                      disabled={isLoading || !discoveryPrompt.trim()}
                      className="px-8 py-4 bg-blue-600 text-white font-bold rounded-xl hover:bg-blue-500 disabled:opacity-50 disabled:cursor-not-allowed transition-all shadow-lg shadow-blue-600/20 flex items-center gap-2"
                    >
                      {isLoading ? <Loader2 className="w-5 h-5 animate-spin" /> : 'DISCOVER'}
                    </button>
                  </div>
                </form>
              )}

              {activeTab === 'ranking' && (
                <div className="p-8 text-center">
                  <div className="inline-block p-4 bg-blue-600/10 rounded-2xl mb-4">
                    <BarChart className="w-8 h-8 text-blue-400" />
                  </div>
                  <h3 className="text-xl font-bold text-white mb-2">Propensity Intelligence</h3>
                  <p className="text-slate-500 text-sm max-w-md mx-auto">
                    A data-driven ranking of previously scanned accounts, sorted by their likelihood to partner based on detected infrastructure gaps and service interest alignment.
                  </p>
                </div>
              )}
            </div>
            
            {(activeTab === 'single' || activeTab === 'discovery') && savedQueries.length > 0 && !isLoading && !result && (
              <div className="mt-8">
                <div className="flex items-center justify-center gap-2 mb-4">
                  <span className="text-[10px] font-mono text-slate-600 uppercase tracking-widest font-bold">Saved Intelligence Queries</span>
                  <div className="h-px w-24 bg-white/5" />
                </div>
                <div className="flex flex-wrap justify-center gap-3">
                  {savedQueries.map((q, idx) => (
                    <div key={`saved-${idx}`} className="group relative">
                      <button
                        onClick={() => {
                          if (activeTab === 'single') setQuery(q);
                          else setDiscoveryPrompt(q);
                        }}
                        className="px-4 py-2 glass border-blue-500/10 text-xs font-mono text-blue-400 rounded-xl hover:bg-blue-500/10 hover:border-blue-500/30 transition-all flex items-center gap-2 pr-8"
                      >
                        <Terminal className="w-3 h-3" />
                        <span className="truncate max-w-[200px]">{q}</span>
                      </button>
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          removeSavedQuery(q);
                        }}
                        className="absolute right-2 top-1/2 -translate-y-1/2 p-1 text-slate-600 hover:text-red-400 opacity-0 group-hover:opacity-100 transition-all"
                      >
                        <X className="w-3 h-3" />
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            )}
            
            {activeTab === 'single' && recentScans.length > 0 && !isLoading && !result && (
              <div className="mt-6 flex flex-wrap justify-center gap-2">
                <span className="text-[10px] font-mono text-slate-600 uppercase tracking-widest mr-2 flex items-center">Recent:</span>
                {recentScans.map((scan, idx) => (
                  <button
                    key={`recent-${idx}`}
                    onClick={() => { setQuery(scan); }}
                    className="px-3 py-1 glass border-white/5 text-[10px] font-mono text-slate-400 rounded-full hover:text-blue-400 hover:border-blue-500/30 transition-all"
                  >
                    {scan}
                  </button>
                ))}
              </div>
            )}
          </div>
        </section>

        {/* Error State */}
        <AnimatePresence>
          {error && (
            <motion.div
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0 }}
              className="max-w-2xl mx-auto mb-8 p-4 glass border-red-500/20 rounded-xl flex items-start gap-3 text-red-400"
            >
              <AlertCircle className="w-5 h-5 shrink-0 mt-0.5" />
              <p className="text-sm font-medium">{error}</p>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Loading State */}
        {isLoading && (
          <div className="flex flex-col items-center justify-center py-20 space-y-8">
            <div className="relative w-32 h-32">
              <div className="absolute inset-0 border-2 border-blue-500/20 rounded-full" />
              <motion.div 
                animate={{ rotate: 360 }}
                transition={{ duration: 2, repeat: Infinity, ease: "linear" }}
                className="absolute inset-0 border-t-2 border-blue-500 rounded-full shadow-[0_0_15px_rgba(59,130,246,0.5)]"
              />
              <div className="absolute inset-4 overflow-hidden rounded-full glass border-white/5 flex items-center justify-center">
                <div className="absolute inset-0 bg-blue-500/10 animate-scan" />
                <Building2 className="w-8 h-8 text-blue-400 relative z-10" />
              </div>
            </div>
            <div className="text-center space-y-4">
              <div className="flex flex-col items-center">
                <p className="text-xl font-black text-white tracking-tight font-brand">SCALING INTELLIGENCE</p>
                <p className="text-xs font-mono text-blue-400/60 uppercase tracking-[0.3em] mt-1 animate-pulse">{loadingStatus}</p>
              </div>
              <div className="flex gap-1 justify-center">
                {[0, 1, 2].map(i => (
                  <motion.div
                    key={`dot-${i}`}
                    animate={{ opacity: [0.3, 1, 0.3] }}
                    transition={{ duration: 1, repeat: Infinity, delay: i * 0.2 }}
                    className="w-1.5 h-1.5 bg-blue-500 rounded-full"
                  />
                ))}
              </div>
            </div>
          </div>
        )}

        {/* Results Section */}
        <AnimatePresence>
          {activeTab === 'ranking' && (scanHistory.length > 0 || summaries.length > 0) && (
            <motion.div
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              className="space-y-6 max-w-5xl mx-auto"
            >
              <div className="flex flex-col md:flex-row items-start md:items-center justify-between px-2 gap-4">
                <div className="flex items-start sm:items-center gap-3">
                  <div className="w-8 h-8 bg-blue-600/20 rounded-lg flex items-center justify-center text-blue-400 shrink-0">
                    <BarChart className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-xl font-black text-white tracking-tight uppercase font-brand">Account Spend Ranking</h3>
                    <div className="flex flex-wrap items-center gap-2 mt-0.5">
                      <p className="text-[10px] font-mono text-slate-500 uppercase tracking-widest">Sorted by Monthly Tech Expenditure</p>
                      <div className="hidden sm:block h-1.5 w-1.5 bg-blue-500 rounded-full" />
                      <span className="px-2 py-0.5 bg-blue-500/10 border border-blue-500/20 text-blue-400 font-mono text-[9px] rounded font-bold">
                        1 USD = ₹{exchangeRate.toFixed(2)} INR
                      </span>
                    </div>
                  </div>
                </div>
                <div className="flex items-center gap-4 w-full md:w-auto">
                  <button
                    onClick={() => {
                      const allItems = [...scanHistory, ...summaries];
                      const uniqueItems = Array.from(new Map(allItems.map(item => [item.domain, item])).values());
                      const ranked = uniqueItems.map(item => {
                        const usdMonthly = getMonthlySpend(item);
                        const inrMonthly = usdMonthly * exchangeRate;
                        const usdYearly = usdMonthly * 12;
                        const inrYearly = inrMonthly * 12;

                        return {
                          Name: item.name,
                          Domain: item.domain,
                          'Monthly Spend (USD)': `$${usdMonthly.toLocaleString()}`,
                          'Monthly Spend (INR)': `₹${inrMonthly.toLocaleString()}`,
                          'Yearly Spend (USD)': `$${usdYearly.toLocaleString()}`,
                          'Yearly Spend (INR)': `₹${inrYearly.toLocaleString()}`,
                          PropensityScore: `${getPropensityScore(item)}%`,
                          TopInterest: getTopInterest(item),
                          TeamSize: item.teamSize,
                          Status: 'techStackSpend' in item ? 'Deep Scanned' : 'Bulk Analyzed'
                        };
                      }).sort((a, b) => parseCurrency(b['Monthly Spend (USD)']) - parseCurrency(a['Monthly Spend (USD)']));
                      exportToExcel(ranked, `spend_ranking_${new Date().toISOString().split('T')[0]}`, 'ranking');
                    }}
                    className="flex items-center gap-2 px-4 py-2 glass border-white/10 rounded-xl text-xs font-bold text-green-400 hover:text-white hover:border-green-500/50 transition-all font-mono"
                  >
                    <Download className="w-4 h-4" />
                    EXPORT EXCEL
                  </button>
                </div>
              </div>

              <div className="glass rounded-[2rem] border-white/5 overflow-hidden">
                <div className="overflow-x-auto custom-scrollbar">
                  <table className="w-full text-left border-collapse min-w-[800px]">
                  <thead>
                    <tr className="border-b border-white/10 bg-white/5">
                      <th className="py-4 px-6 text-[10px] font-mono text-slate-500 uppercase tracking-widest">Rank</th>
                      <th className="py-4 px-6 text-[10px] font-mono text-slate-500 uppercase tracking-widest">Account</th>
                      <th className="py-4 px-6 text-[10px] font-mono text-slate-500 uppercase tracking-widest">Monthly Spend (USD/INR)</th>
                      <th className="py-4 px-6 text-[10px] font-mono text-slate-500 uppercase tracking-widest">Annual Spend (USD/INR)</th>
                      <th className="py-4 px-6 text-[10px] font-mono text-slate-500 uppercase tracking-widest">Growth Potential</th>
                      <th className="py-4 px-6 text-[10px] font-mono text-slate-500 uppercase tracking-widest text-right">Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(() => {
                      const allItems = [...scanHistory, ...summaries];
                      const uniqueItems = Array.from(new Map(allItems.map(item => [item.domain, item])).values());
                      
                      return uniqueItems
                        .map(item => ({
                          ...item,
                          monthlySpendVal: getMonthlySpend(item),
                          propensityScore: getPropensityScore(item),
                          topInterest: getTopInterest(item)
                        }))
                        .sort((a, b) => b.monthlySpendVal - a.monthlySpendVal)
                        .map((company, idx) => (
                          <tr key={company.domain} className="border-b border-white/5 hover:bg-white/5 transition-colors group">
                            <td className="py-4 px-6">
                              <span className={cn(
                                "text-lg font-black font-mono",
                                idx === 0 ? "text-yellow-400" : idx === 1 ? "text-slate-300" : idx === 2 ? "text-orange-500" : "text-slate-600"
                              )}>
                                #{idx + 1}
                              </span>
                            </td>
                            <td className="py-4 px-6">
                              <div className="flex items-center gap-3">
                                <div className="w-10 h-10 glass border-white/10 rounded-xl flex items-center justify-center bg-white/5 group-hover:scale-110 transition-transform">
                                  {('logoUrl' in company && (company as any).logoUrl) ? (
                                    <img src={(company as any).logoUrl} alt="" className="w-full h-full object-contain p-1.5" referrerPolicy="no-referrer" />
                                  ) : (
                                    <Building2 className="w-5 h-5 text-blue-400" />
                                  )}
                                </div>
                                <div>
                                  <p className="text-sm font-bold text-white">{company.name}</p>
                                  <p className="text-[10px] font-mono text-slate-500">{company.domain}</p>
                                </div>
                              </div>
                            </td>
                            <td className="py-4 px-6">
                              <div className="flex flex-col">
                                <span className="text-sm font-black text-green-400 font-mono">
                                  ${company.monthlySpendVal.toLocaleString()}
                                </span>
                                <span className="text-[10px] font-black text-blue-400 font-mono">
                                  ₹{(company.monthlySpendVal * exchangeRate).toLocaleString()}
                                </span>
                              </div>
                            </td>
                            <td className="py-4 px-6">
                              <div className="flex flex-col">
                                <span className="text-sm font-black text-emerald-400 font-mono">
                                  ${(company.monthlySpendVal * 12).toLocaleString()}
                                </span>
                                <span className="text-[10px] font-black text-indigo-400 font-mono">
                                  ₹{(company.monthlySpendVal * 12 * exchangeRate).toLocaleString()}
                                </span>
                              </div>
                            </td>
                            <td className="py-4 px-6">
                              <div className="flex items-center gap-3">
                                <div className="w-24 h-1.5 bg-white/5 rounded-full overflow-hidden">
                                  <motion.div 
                                    initial={{ width: 0 }}
                                    animate={{ width: `${company.propensityScore}%` }}
                                    className={cn(
                                      "h-full rounded-full",
                                      company.propensityScore > 75 ? "bg-green-500" : company.propensityScore > 50 ? "bg-blue-500" : "bg-orange-500"
                                    )}
                                  />
                                </div>
                                <span className="text-xs font-mono font-bold text-white">{company.propensityScore}%</span>
                              </div>
                            </td>
                            <td className="py-4 px-6 text-right">
                              <button
                                onClick={() => {
                                  if ('techStackSpend' in company) {
                                    setResult(company as CompanyResearch);
                                    setActiveTab('single');
                                  } else {
                                    handleDeepScan(company.name);
                                  }
                                }}
                                className="p-2 glass border-white/10 rounded-xl text-slate-400 hover:text-white hover:border-blue-500/50 transition-all flex items-center gap-2 ml-auto"
                              >
                                <span className="text-[9px] font-bold uppercase tracking-tighter">View Intel</span>
                                <ExternalLink className="w-3 h-3" />
                              </button>
                            </td>
                          </tr>
                        ));
                    })()}
                  </tbody>
                </table>
                </div>
                {scanHistory.length === 0 && summaries.length === 0 && (
                  <div className="py-20 text-center">
                    <p className="text-slate-500 text-sm font-mono">NO SCAN DATA DETECTED</p>
                  </div>
                )}
              </div>
            </motion.div>
          )}

          {activeTab !== 'ranking' && summaries.length > 0 && (
            <motion.div
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              className="space-y-6"
            >
              <div className="flex items-center justify-between px-2">
                <div className="flex items-center gap-3">
                  <div className="w-8 h-8 bg-blue-600/20 rounded-lg flex items-center justify-center text-blue-400">
                    <ListFilter className="w-5 h-5" />
                  </div>
                  <h3 className="text-xl font-black text-white tracking-tight uppercase font-brand">Discovered Opportunities ({summaries.length})</h3>
                </div>
                <div className="flex items-center gap-4">
                  <button
                    onClick={() => {
                      const formattedSummaries = summaries.map(item => {
                        const usdMonthly = parseCurrency(item.estimatedMonthlyTechSpend);
                        const inrMonthly = usdMonthly * exchangeRate;
                        const usdYearly = usdMonthly * 12;
                        const inrYearly = inrMonthly * 12;
                        return {
                          Name: item.name,
                          Domain: item.domain,
                          Industry: item.industry,
                          Location: item.location,
                          TeamSize: item.teamSize,
                          Description: item.description,
                          'Monthly Spend (USD)': `$${usdMonthly.toLocaleString()}`,
                          'Monthly Spend (INR)': `₹${inrMonthly.toLocaleString()}`,
                          'Yearly Spend (USD)': `$${usdYearly.toLocaleString()}`,
                          'Yearly Spend (INR)': `₹${inrYearly.toLocaleString()}`
                        };
                      });
                      exportToExcel(formattedSummaries, `discovered_opportunities_${new Date().toISOString().split('T')[0]}`, 'summary');
                    }}
                    className="flex items-center gap-2 px-4 py-2 glass border-white/10 rounded-xl text-[10px] font-bold text-green-400 hover:text-white hover:border-green-500/50 transition-all font-mono"
                  >
                    <Download className="w-3 h-3" />
                    EXPORT EXCEL
                  </button>
                  <div className="h-4 w-px bg-white/10" />
                  <span className="text-[10px] font-mono text-slate-500 uppercase tracking-widest">Sort: High Spend First</span>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
                {[...summaries]
                  .sort((a, b) => parseCurrency(b.estimatedMonthlyTechSpend) - parseCurrency(a.estimatedMonthlyTechSpend))
                  .map((summary, idx) => (
                    <motion.div
                      key={`summary-${idx}`}
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      transition={{ delay: idx * 0.05 }}
                      className="glass p-6 rounded-3xl border-white/5 group hover:border-blue-500/30 transition-all relative overflow-hidden"
                    >
                      <div className="absolute top-0 right-0 w-32 h-32 bg-blue-600/5 blur-[40px] -mr-16 -mt-16 group-hover:bg-blue-600/10 transition-all" />
                      
                      <div className="flex items-start justify-between mb-4 relative z-10">
                        <div className="w-12 h-12 glass border-white/10 rounded-xl flex items-center justify-center bg-white/5">
                          <Building2 className="w-6 h-6 text-blue-400" />
                        </div>
                        <div className="flex flex-col items-end">
                          <span className="text-xs font-black text-green-400 font-mono">
                            ${parseCurrency(summary.estimatedMonthlyTechSpend).toLocaleString()}/mo
                          </span>
                          <span className="text-[8px] font-mono text-slate-500 uppercase tracking-tighter">Est. Spend</span>
                        </div>
                      </div>

                      <div className="relative z-10">
                        <div className="flex items-center justify-between gap-4 mb-1">
                          <h4 className="text-lg font-bold text-white font-brand truncate">{summary.name}</h4>
                          <button
                            onClick={() => handleDeepScan(summary.domain)}
                            className="shrink-0 p-1.5 glass border-white/10 text-blue-400 hover:bg-blue-600 hover:text-white rounded-lg transition-all"
                          >
                            <Target className="w-3.5 h-3.5" />
                          </button>
                        </div>
                        <div className="flex items-center gap-2 mb-3">
                          <Globe className="w-3 h-3 text-slate-500" />
                          <span className="text-[10px] font-mono text-slate-400 tracking-wider truncate">{summary.domain}</span>
                        </div>
                        
                        <div className="grid grid-cols-2 gap-2 mb-3">
                          <div className="px-2 py-1 bg-white/5 rounded-md border border-white/5">
                            <p className="text-[8px] font-mono text-slate-500 uppercase tracking-widest mb-0.5">Industry</p>
                            <p className="text-[10px] text-slate-300 font-bold truncate">{summary.industry}</p>
                          </div>
                          <div className="px-2 py-1 bg-white/5 rounded-md border border-white/5">
                            <p className="text-[8px] font-mono text-slate-500 uppercase tracking-widest mb-0.5">Size</p>
                            <p className="text-[10px] text-slate-300 font-bold truncate">{summary.teamSize}</p>
                          </div>
                        </div>

                        {/* Dual-Format 4-Dimension Financials */}
                        {(() => {
                          const val = parseCurrency(summary.estimatedMonthlyTechSpend);
                          return (
                            <div className="bg-white/5 rounded-2xl p-3 border border-white/5 mb-3 space-y-2">
                              <div className="flex items-center justify-between border-b border-white/5 pb-1">
                                <span className="text-[8px] font-mono text-slate-500 uppercase tracking-widest font-bold">Dual-Format Financials</span>
                                <span className="text-[8px] font-mono text-blue-400 uppercase tracking-widest">Est. Spend</span>
                              </div>
                              <div className="grid grid-cols-2 gap-x-3 gap-y-1.5">
                                <div>
                                  <p className="text-[7px] font-mono text-slate-500 uppercase tracking-wider mb-0.5">Monthly Spend</p>
                                  <p className="text-[10px] font-mono font-black text-green-400 leading-none">${val.toLocaleString()}</p>
                                  <p className="text-[9px] font-mono font-bold text-blue-400 leading-none mt-1">₹{Math.round(val * exchangeRate).toLocaleString()}</p>
                                </div>
                                <div>
                                  <p className="text-[7px] font-mono text-slate-500 uppercase tracking-wider mb-0.5">Yearly Spend</p>
                                  <p className="text-[10px] font-mono font-black text-emerald-400 leading-none">${Math.round(val * 12).toLocaleString()}</p>
                                  <p className="text-[9px] font-mono font-bold text-indigo-400 leading-none mt-1">₹{Math.round(val * 12 * exchangeRate).toLocaleString()}</p>
                                </div>
                              </div>
                            </div>
                          );
                        })()}

                      <p className="text-xs text-slate-400 leading-relaxed line-clamp-2 mb-4">
                        {summary.description}
                      </p>

                      <div className="flex items-center gap-2 text-[10px] font-mono text-blue-400/60 uppercase tracking-widest">
                        <Target className="w-3 h-3" />
                        <span>High Potential Target</span>
                      </div>
                    </div>
                  </motion.div>
                ))}
              </div>
            </motion.div>
          )}

          {result && (
            <motion.div
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              className="grid grid-cols-1 lg:grid-cols-12 gap-6"
            >
              {/* Left Column: Main Data */}
              <div className="lg:col-span-8 space-y-6">
                {/* Hero Card */}
                <div className={cn(
                  "glass rounded-[2.5rem] p-8 border-white/5 relative overflow-hidden group",
                  result.isDeadpooled && "border-red-500/30 bg-red-900/10"
                )}>
                  <div className="absolute top-0 right-0 w-64 h-64 bg-blue-600/10 blur-[80px] -mr-32 -mt-32 group-hover:bg-blue-600/20 transition-all duration-700" />
                  
                  <div className="relative z-10 flex flex-wrap items-center justify-between gap-6">
                    <div className="flex items-center gap-6">
                      <div className="w-24 h-24 glass border-white/10 rounded-3xl flex items-center justify-center overflow-hidden bg-white/5 shadow-inner">
                        {result.logoUrl ? (
                          <img src={result.logoUrl} alt={`${result.name} Logo`} className="w-full h-full object-contain p-2" referrerPolicy="no-referrer" />
                        ) : (
                          <span className="text-4xl font-black text-blue-400">{result.name.charAt(0)}</span>
                        )}
                      </div>
                      <div>
                        <div className="flex items-center gap-3 mb-1">
                          <h3 className="text-3xl font-black text-white tracking-tight font-brand">{result.name}</h3>
                          <button
                            onClick={() => exportToExcel(result, `${result.name.toLowerCase().replace(/\s+/g, '_')}_intelligence`, 'single')}
                            className="p-2 glass border-white/10 rounded-xl text-green-400 hover:text-white hover:border-green-500/50 transition-all ml-2"
                            title="Export to Excel"
                          >
                            <Download className="w-4 h-4" />
                          </button>
                          <button
                            onClick={() => exportToJSON(result, `${result.name.toLowerCase().replace(/\s+/g, '_')}_intelligence`)}
                            className="p-2 glass border-white/10 rounded-xl text-slate-400 hover:text-blue-400 hover:border-blue-500/50 transition-all"
                            title="Export to JSON"
                          >
                            <Info className="w-4 h-4" />
                          </button>
                          {result.isDeadpooled && (
                            <span className="px-3 py-1 bg-red-500 text-white text-[10px] font-black uppercase tracking-widest rounded-lg shadow-lg shadow-red-500/20">
                              DEADPOOLED
                            </span>
                          )}
                        </div>
                        <div className="flex flex-col gap-1">
                          <div className="flex flex-wrap items-center gap-x-4 gap-y-1">
                            <a href={`https://${result.domain}`} target="_blank" rel="noopener noreferrer" className="text-blue-400 hover:text-blue-300 transition-colors flex items-center gap-1.5 text-sm font-mono tracking-wider">
                              <Globe className="w-4 h-4" /> {result.domain}
                            </a>
                            {result.sources && result.sources.length > 0 && (
                              <button 
                                onClick={() => document.getElementById('sources')?.scrollIntoView({ behavior: 'smooth' })}
                                className="text-slate-500 hover:text-blue-400 transition-colors flex items-center gap-1.5 text-[10px] font-mono tracking-wider uppercase"
                              >
                                <Database className="w-3.5 h-3.5" /> Citations ({result.sources?.length || 0})
                              </button>
                            )}
                          </div>
                          <span className="text-slate-500 text-[10px] font-mono uppercase tracking-widest">{result.funding.round} • <span className="text-blue-400 font-bold underline decoration-blue-500/30 underline-offset-4">{result.arr} ARR</span></span>
                          {result.funding.lastFundingDate && (
                            <span className="text-slate-600 text-[9px] font-mono uppercase tracking-widest mt-1">Last Funding: {result.funding.lastFundingDate}</span>
                          )}
                        </div>
                      </div>
                    </div>
                    
                    <div className="flex gap-4">
                      <button 
                        onClick={copyToClipboard}
                        className="glass p-3 rounded-2xl border-white/10 text-slate-400 hover:text-blue-400 hover:bg-white/5 transition-all flex items-center gap-2"
                        title="Copy Report"
                      >
                        {isCopied ? <Check className="w-5 h-5 text-green-400" /> : <Copy className="w-5 h-5" />}
                      </button>
                      <div className="glass px-6 py-3 rounded-2xl border-white/5 text-center min-w-[120px]">
                        <p className="text-[10px] font-mono text-slate-500 uppercase tracking-widest mb-1">Valuation</p>
                        <p className="text-xl font-black text-blue-400 shadow-[0_0_10px_rgba(59,130,246,0.3)]">{result.funding.totalAmount || 'BOOTSTRAP'}</p>
                      </div>
                    </div>
                  </div>

                  {result.funding.investors && result.funding.investors.length > 0 && (
                    <div className="mt-8 pt-6 border-t border-white/5 relative z-10">
                      <p className="text-[10px] font-mono text-slate-500 uppercase tracking-widest mb-3">Key Investors</p>
                      <div className="flex flex-wrap gap-2">
                        {result.funding.investors.map((investor, idx) => (
                          <span key={`investor-${idx}`} className="px-3 py-1 bg-white/5 rounded-lg text-[10px] font-mono text-slate-400 border border-white/5">
                            {investor}
                          </span>
                        ))}
                      </div>
                    </div>
                  )}
                </div>

                {/* Employees & Customers Section with Dual-Format Financial Footprint Card */}
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <div className="glass p-6 rounded-3xl border-white/5 flex items-center gap-6">
                    <div className="w-14 h-14 bg-blue-600/20 rounded-2xl flex items-center justify-center text-blue-400 shrink-0">
                      <Users className="w-7 h-7" />
                    </div>
                    <div>
                      <p className="text-[10px] font-mono text-slate-500 uppercase tracking-widest mb-1">Current Employees</p>
                      <p className="text-2xl font-black text-white tracking-tight">{result.teamSize}</p>
                    </div>
                  </div>
                  <div className="glass p-6 rounded-3xl border-white/5 flex items-center gap-6">
                    <div className="w-14 h-14 bg-red-600/20 rounded-2xl flex items-center justify-center text-red-400 shrink-0">
                      <Briefcase className="w-7 h-7" />
                    </div>
                    <div>
                      <p className="text-[10px] font-mono text-slate-500 uppercase tracking-widest mb-1">Customer Base</p>
                      <p className="text-2xl font-black text-white tracking-tight">{result.customerBase}</p>
                    </div>
                  </div>
                  <div className="glass p-6 rounded-3xl border-white/5 flex flex-col justify-between">
                    <div className="flex items-center gap-4 mb-4">
                      <div className="w-10 h-10 bg-green-600/20 rounded-2xl flex items-center justify-center text-green-400 shrink-0">
                        <TrendingUp className="w-5 h-5" />
                      </div>
                      <div>
                        <p className="text-[9px] font-mono text-slate-500 uppercase tracking-widest leading-none">Est. Tech Spend</p>
                        <p className="text-sm font-black text-white font-brand tracking-tight mt-1">Dual-Format Budget</p>
                      </div>
                    </div>
                    <div className="grid grid-cols-2 gap-x-4 gap-y-2 pt-2 border-t border-white/5">
                      <div>
                        <p className="text-[7px] font-mono text-slate-500 uppercase tracking-wider mb-0.5">Monthly USD/INR</p>
                        <p className="text-xs font-mono font-black text-green-400 leading-none">${Math.round(getMonthlySpend(result)).toLocaleString()}</p>
                        <p className="text-[10px] font-mono font-bold text-blue-400 leading-none mt-1">₹{Math.round(getMonthlySpend(result) * exchangeRate).toLocaleString()}</p>
                      </div>
                      <div>
                        <p className="text-[7px] font-mono text-slate-500 uppercase tracking-wider mb-0.5">Yearly USD/INR</p>
                        <p className="text-xs font-mono font-black text-emerald-400 leading-none">${Math.round(getMonthlySpend(result) * 12).toLocaleString()}</p>
                        <p className="text-[10px] font-mono font-bold text-indigo-400 leading-none mt-1">₹{Math.round(getMonthlySpend(result) * 12 * exchangeRate).toLocaleString()}</p>
                      </div>
                    </div>
                  </div>
                </div>

                {/* Decision Maker Terminal */}
                <div className="space-y-4">
                  <div className="flex items-center justify-between px-2">
                    <h4 className="text-xs font-mono font-bold text-slate-500 uppercase tracking-[0.3em]">Decision Maker Terminal (Leadership)</h4>
                    <span className="text-[10px] font-mono text-blue-400/60 uppercase tracking-widest">{(result.leadership?.length || 0)} Profiles Identified</span>
                  </div>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {result.leadership?.map((leader, idx) => (
                      <motion.div 
                        key={`leader-${idx}`}
                        whileHover={{ y: -5 }}
                        className="glass p-6 rounded-3xl border-white/5 group relative overflow-hidden"
                      >
                        <div className="absolute top-0 left-0 w-1 h-full bg-blue-600 opacity-0 group-hover:opacity-100 transition-opacity" />
                        <div className="flex items-start justify-between mb-6">
                          <div className="w-12 h-12 bg-slate-800 rounded-xl flex items-center justify-center text-blue-400 font-bold relative overflow-hidden group-hover:scale-110 transition-transform">
                            <div className="absolute inset-0 bg-blue-500/10 animate-pulse" />
                            <span className="relative z-10">{leader.name.split(' ').map(n => n[0]).join('')}</span>
                          </div>
                          <a href={leader.linkedin} target="_blank" rel="noopener noreferrer" className="p-2.5 glass border-white/10 rounded-xl text-slate-400 hover:text-[#0077B5] hover:bg-white/5 transition-all">
                            <Linkedin className="w-5 h-5" />
                          </a>
                        </div>
                        <h5 className="text-lg font-bold text-white mb-1 font-brand">{leader.name}</h5>
                        <p className="text-xs font-mono text-blue-400 uppercase tracking-widest mb-1">{leader.title}</p>
                        {leader.reportsTo && (
                          <div className="flex items-center gap-1.5 mb-4 opacity-60">
                            <ArrowRight className="w-2.5 h-2.5 text-slate-500" />
                            <span className="text-[9px] font-mono text-slate-500 uppercase tracking-widest">Reports to: <span className="text-white font-bold">{leader.reportsTo}</span></span>
                          </div>
                        )}
                        {!leader.reportsTo && <div className="mb-4" />}
                        
                        <div className="bg-white/5 rounded-xl p-3 border border-white/5">
                          <p className="text-[9px] font-mono text-slate-500 uppercase tracking-widest font-bold mb-1.5 flex items-center gap-1.5">
                            <Target className="w-3 h-3 text-red-400" /> Focus Area
                          </p>
                          <p className="text-[11px] text-slate-300 leading-relaxed italic border-l border-red-500/30 pl-3">
                            {leader.focus}
                          </p>
                        </div>
                      </motion.div>
                    ))}
                  </div>
                </div>

                {/* Organizational Flow Chart Section */}
                <div className="space-y-6">
                  <div className="flex items-center justify-between px-2">
                    <div className="flex items-center gap-3">
                      <div className="w-8 h-8 bg-purple-600/20 rounded-lg flex items-center justify-center text-purple-400">
                        <Terminal className="w-5 h-5" />
                      </div>
                      <div>
                        <h4 className="text-xs font-mono font-bold text-slate-500 uppercase tracking-[0.3em]">Command Hierarchy (Flow Chart)</h4>
                        <p className="text-[9px] font-mono text-slate-600 uppercase tracking-widest mt-0.5">Reporting & Reporting Structure Mapping</p>
                      </div>
                    </div>
                  </div>

                  <div className="glass p-8 rounded-[2.5rem] border-white/5 bg-gradient-to-br from-purple-900/5 to-transparent overflow-x-auto">
                    <div className="min-w-[600px] flex flex-col items-center py-4">
                      {/* CEO / Root Node */}
                      <div className="flex flex-col items-center mb-12 relative">
                        <div className="glass p-5 rounded-2xl border-purple-500/30 bg-purple-600/10 w-64 text-center group hover:border-purple-400 transition-all cursor-default">
                          <p className="text-[10px] font-mono text-purple-400 uppercase tracking-widest font-bold mb-1">Chief Executive Officer</p>
                          <h6 className="text-lg font-black text-white font-brand tracking-tight">
                            {result.orgChart.ceo || (result.leadership?.find(l => l.title.toLowerCase().includes('ceo'))?.name || 'N/A')}
                          </h6>
                          <div className="h-0.5 w-12 bg-purple-500 mx-auto mt-3 rounded-full opacity-50" />
                        </div>
                        {/* Vertical line from CEO */}
                        <div className="absolute top-full left-1/2 -translate-x-1/2 w-0.5 h-12 bg-purple-500/20" />
                      </div>

                      {/* Second Tier / Direct Reports */}
                      <div className="relative pt-6">
                        {/* Horizontal Connector Line */}
                        {result.orgChart.structure.length > 1 && (
                          <div className="absolute top-0 left-0 right-0 h-0.5 bg-purple-500/20 mx-32 rounded-full" />
                        )}
                        
                        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-x-8 gap-y-12">
                          {result.orgChart.structure.map((link, idx) => {
                            const leader = result.leadership?.find(l => l.name === link.from);
                            return (
                              <div key={`chart-${idx}`} className="flex flex-col items-center relative">
                                {/* Vertical Connector to horizontal line */}
                                <div className="absolute -top-6 left-1/2 -translate-x-1/2 w-0.5 h-6 bg-purple-500/20" />
                                
                                <div className="glass p-4 rounded-2xl border-white/10 bg-white/5 w-56 text-center hover:bg-white/10 transition-all cursor-default group">
                                  <div className="flex items-center justify-center gap-2 mb-2">
                                    <span className="text-[8px] font-mono text-slate-500 uppercase tracking-widest bg-white/5 px-2 py-0.5 rounded border border-white/5">
                                      {link.relationship}
                                    </span>
                                  </div>
                                  <h6 className="text-sm font-black text-white font-brand truncate">{link.from}</h6>
                                  <p className="text-[10px] font-mono text-purple-400 uppercase tracking-widest mt-1 truncate">
                                    {leader?.name === link.from ? leader?.title : 'Executive Leader'}
                                  </p>
                                  <div className="mt-3 flex items-center justify-center gap-1.5 grayscale opacity-40 group-hover:grayscale-0 group-hover:opacity-100 transition-all">
                                    <div className="w-1.5 h-1.5 bg-green-500 rounded-full" />
                                    <span className="text-[8px] font-mono text-slate-500 uppercase">Active Role</span>
                                  </div>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    </div>

                    <div className="mt-12 pt-8 border-t border-white/5">
                      <div className="flex items-start gap-4 p-4 rounded-2xl bg-white/2 border border-white/5">
                        <Info className="w-5 h-5 text-purple-400 shrink-0 mt-0.5" />
                        <p className="text-[10px] font-mono text-slate-500 leading-relaxed">
                          <span className="text-purple-400 font-bold uppercase tracking-widest block mb-1">Hierarchy Logic Mapping</span>
                          Reporting structures are derived from public leadership statements, SEC filings, and professional network signals. In the absence of an explicit public org chart, reporting lines are inferred based on standard corporate governance patterns (C-Suite levels report to the CEO, functional heads report to respective C-Suite officers). This visual represents the primary decision-making core identified for current sales strategy.
                        </p>
                      </div>
                    </div>
                  </div>
                </div>

                {/* Legacy Founders (if different from leadership) */}
                {result.founders?.length > 0 && result.founders?.some(f => !(result.leadership || []).find(l => l.name === f.name)) && (
                  <div className="space-y-4">
                    <h4 className="text-xs font-mono font-bold text-slate-500 uppercase tracking-[0.3em] px-2">Founding Team</h4>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      {result.founders?.filter(f => !(result.leadership || []).find(l => l.name === f.name)).map((founder, idx) => (
                        <motion.div 
                          key={`founder-${idx}`}
                          whileHover={{ y: -5 }}
                          className="glass p-5 rounded-3xl border-white/5 group relative overflow-hidden opacity-80 hover:opacity-100"
                        >
                          <div className="flex items-center gap-4">
                            <div className="w-10 h-10 bg-slate-800 rounded-lg flex items-center justify-center text-slate-400 font-bold text-xs">
                              {founder.name.split(' ').map(n => n[0]).join('')}
                            </div>
                            <div className="flex-1">
                              <h5 className="text-sm font-bold text-white font-brand">{founder.name}</h5>
                              <p className="text-[10px] font-mono text-slate-500 uppercase tracking-widest">{founder.role}</p>
                            </div>
                            <a href={founder.linkedin} target="_blank" rel="noopener noreferrer" className="p-2 glass border-white/10 rounded-lg text-slate-500 hover:text-[#0077B5] transition-all">
                              <Linkedin className="w-4 h-4" />
                            </a>
                          </div>
                        </motion.div>
                      ))}
                    </div>
                  </div>
                )}

                {/* Detailed Tech Stack */}
                <div className="glass p-8 rounded-[2rem] border-white/5">
                  <div className="flex items-center gap-3 mb-8">
                    <div className="w-10 h-10 bg-blue-600/20 rounded-xl flex items-center justify-center text-blue-400">
                      <Cpu className="w-6 h-6" />
                    </div>
                    <h4 className="text-lg font-black text-white tracking-tight font-brand uppercase">Detailed Tech Stack</h4>
                  </div>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                    {Object.entries(result.techStack).map(([category, items]) => {
                      const Icon = TECH_STACK_ICONS[category] || Database;
                      const spendInfo = result.techStackSpend?.[category as keyof typeof result.techStackSpend];
                      const subtotal = spendInfo?.subtotal || 'N/A';
                      const spendItems = spendInfo?.items || [];
                      
                      return (
                        <div key={`tech-${category}`} className="space-y-4 group/stack">
                          <div className="flex items-center justify-between border-b border-white/5 pb-2">
                            <div className="flex items-center gap-2">
                              <Icon className="w-3.5 h-3.5 text-blue-400 group-hover/stack:scale-110 transition-transform" />
                              <p className="text-[10px] font-mono font-bold text-blue-400 uppercase tracking-widest">{category}</p>
                            </div>
                            <div className="flex items-center gap-1.5 px-2 py-1 bg-blue-500/10 border border-blue-500/20 rounded-md">
                              <TrendingUp className="w-2.5 h-2.5 text-blue-300" />
                              <span className="text-[9px] font-mono font-bold text-blue-200">TOTAL: <span className="text-white underline decoration-blue-400/50 underline-offset-2">{subtotal}</span></span>
                            </div>
                          </div>
                          
                          <div className="space-y-3">
                            {subtotal !== 'N/A' && (
                              <div className="bg-blue-950/20 rounded-xl p-3 border border-blue-500/10 space-y-2">
                                <div className="flex items-center justify-between border-b border-white/5 pb-1">
                                  <p className="text-[8px] font-mono text-blue-400 uppercase tracking-widest font-bold">Category Financials</p>
                                  <div className="flex items-center gap-1">
                                    <div className="w-1.5 h-1.5 bg-blue-500 rounded-full animate-pulse" />
                                    <span className="text-[8px] font-mono text-slate-500">4D BUDGET</span>
                                  </div>
                                </div>
                                <div className="grid grid-cols-2 gap-3">
                                  <div>
                                    <div className="text-[7px] font-mono text-slate-500 uppercase leading-none">Monthly USD/INR</div>
                                    <div className="text-[11px] font-mono font-black text-green-400 leading-tight mt-1">${parseCurrency(subtotal).toLocaleString()}</div>
                                    <div className="text-[9px] font-mono font-bold text-blue-400 leading-tight">₹{Math.round(parseCurrency(subtotal) * exchangeRate).toLocaleString()}</div>
                                  </div>
                                  <div>
                                    <div className="text-[7px] font-mono text-slate-500 uppercase leading-none">Yearly USD/INR</div>
                                    <div className="text-[11px] font-mono font-black text-emerald-400 leading-tight mt-1">${Math.round(parseCurrency(subtotal) * 12).toLocaleString()}</div>
                                    <div className="text-[9px] font-mono font-bold text-indigo-400 leading-tight">₹{Math.round(parseCurrency(subtotal) * 12 * exchangeRate).toLocaleString()}</div>
                                  </div>
                                </div>
                              </div>
                            )}

                            <div className="flex flex-wrap gap-2">
                              {items.length > 0 ? items.map((item, idx) => (
                                <span key={`item-${idx}`} className="px-3 py-1 glass border-white/5 text-slate-400 rounded-lg text-[10px] font-mono tracking-wider hover:text-white hover:border-blue-500/30 transition-colors">
                                  {item}
                                </span>
                              )) : <span className="text-slate-600 text-[10px] font-mono italic">No data detected</span>}
                            </div>
                            
                            {spendItems.length > 0 && (
                              <div className="bg-white/5 rounded-xl p-3 border border-white/5 space-y-2">
                                <p className="text-[8px] font-mono text-slate-500 uppercase tracking-widest font-bold">Estimated Cost Breakdown</p>
                                <div className="space-y-1.5">
                                  {spendItems.map((sItem, sIdx) => {
                                    const spendVal = parseCurrency(sItem.spend);
                                    return (
                                      <div key={`spend-${sIdx}`} className="flex flex-col gap-1 text-[9px] font-mono border-b border-white/5 pb-1.5 last:border-none last:pb-1">
                                        <div className="flex items-center justify-between">
                                          <span className="text-slate-200 font-bold">{sItem.name}</span>
                                          <span className="text-green-300 font-bold italic">${spendVal.toLocaleString()}/mo</span>
                                        </div>
                                        <div className="flex justify-between text-[8px] text-slate-500">
                                          <span>INR: ₹{Math.round(spendVal * exchangeRate).toLocaleString()}/mo</span>
                                          <span>Annual: ${Math.round(spendVal * 12).toLocaleString()} | ₹{Math.round(spendVal * 12 * exchangeRate).toLocaleString()}</span>
                                        </div>
                                      </div>
                                    );
                                  })}
                                </div>
                              </div>
                            )}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>

                {/* Detailed Infrastructure */}
                <div className="glass p-8 rounded-[2rem] border-white/5">
                  <div className="flex items-center gap-3 mb-8">
                    <div className="w-10 h-10 bg-red-600/20 rounded-xl flex items-center justify-center text-red-400">
                      <Server className="w-6 h-6" />
                    </div>
                    <h4 className="text-lg font-black text-white tracking-tight font-brand uppercase">Infrastructure Analysis</h4>
                  </div>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-12">
                    <div className="space-y-6">
                      <p className="text-[10px] font-mono font-bold text-red-400 uppercase tracking-widest border-b border-white/5 pb-2">Endpoints</p>
                      <div className="space-y-4">
                        <div className="flex items-start gap-4">
                          <Laptop className="w-5 h-5 text-slate-500 mt-1" />
                          <div>
                            <p className="text-xs font-bold text-white mb-1">Laptops</p>
                            <p className="text-xs text-slate-400 leading-relaxed">{result.itInfrastructure.endpoints?.laptops || '0'}</p>
                          </div>
                        </div>
                        <div className="flex items-start gap-4">
                          <Building2 className="w-5 h-5 text-slate-500 mt-1" />
                          <div>
                            <p className="text-xs font-bold text-white mb-1">Desktops</p>
                            <p className="text-xs text-slate-400 leading-relaxed">{result.itInfrastructure.endpoints?.desktops || '0'}</p>
                          </div>
                        </div>
                        <div className="flex items-start gap-4">
                          <Phone className="w-5 h-5 text-slate-500 mt-1" />
                          <div>
                            <p className="text-xs font-bold text-white mb-1">Mobile</p>
                            <p className="text-xs text-slate-400 leading-relaxed">{result.itInfrastructure.endpoints?.mobile || '0'}</p>
                          </div>
                        </div>
                      </div>
                    </div>
                    <div className="space-y-6">
                      <p className="text-[10px] font-mono font-bold text-red-400 uppercase tracking-widest border-b border-white/5 pb-2">Compute & Networking</p>
                      <div className="space-y-4">
                        <div className="flex items-start gap-4">
                          <Server className="w-5 h-5 text-slate-500 mt-1" />
                          <div>
                            <p className="text-xs font-bold text-white mb-1">Servers</p>
                            <p className="text-xs text-slate-400 leading-relaxed">{result.itInfrastructure.compute?.servers || '0'}</p>
                          </div>
                        </div>
                        <div className="flex items-start gap-4">
                          <Rocket className="w-5 h-5 text-slate-500 mt-1" />
                          <div>
                            <p className="text-xs font-bold text-white mb-1">Storage</p>
                            <p className="text-xs text-slate-400 leading-relaxed">{result.itInfrastructure.compute?.storage || '0'}</p>
                          </div>
                        </div>
                        <div className="flex items-start gap-4">
                          <Globe className="w-5 h-5 text-slate-500 mt-1" />
                          <div>
                            <p className="text-xs font-bold text-white mb-1">Networking</p>
                            <p className="text-xs text-slate-400 leading-relaxed">{result.itInfrastructure.compute?.networking || '0'}</p>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>

                  <div className="mt-12 grid grid-cols-1 md:grid-cols-2 gap-8 pt-8 border-t border-white/5">
                    <div className="space-y-4">
                      <div className="flex items-center gap-2">
                        <ShieldAlert className="w-4 h-4 text-red-400" />
                        <p className="text-[10px] font-mono font-bold text-red-400 uppercase tracking-widest">Full Stack Audit</p>
                      </div>
                      <p className="text-xs text-slate-400 leading-relaxed bg-white/5 p-4 rounded-2xl border border-white/5 italic">
                        "{result.itInfrastructure.fullStackAudit}"
                      </p>
                    </div>
                    <div className="space-y-4">
                      <div className="flex items-center gap-2">
                        <Zap className="w-4 h-4 text-yellow-400" />
                        <p className="text-[10px] font-mono font-bold text-yellow-400 uppercase tracking-widest">OEM Strategy Analysis</p>
                      </div>
                      <p className="text-xs text-slate-400 leading-relaxed bg-white/5 p-4 rounded-2xl border border-white/5 italic">
                        "{result.itInfrastructure.oemAnalysis}"
                      </p>
                    </div>
                  </div>
                </div>

                {/* Infrastructure Pain Points */}
                <div className="glass p-8 rounded-[2rem] border-white/5">
                  <div className="flex items-center gap-3 mb-8">
                    <div className="w-10 h-10 bg-orange-600/20 rounded-xl flex items-center justify-center text-orange-400">
                      <AlertTriangle className="w-6 h-6" />
                    </div>
                    <h4 className="text-lg font-black text-white tracking-tight font-brand uppercase">Infrastructure Pain Points</h4>
                  </div>
                  <div className="grid grid-cols-1 gap-4">
                    {result.painPoints.map((pain, idx) => (
                      <div key={`pain-${idx}`} className="glass p-5 rounded-2xl border-white/5 flex items-start gap-4 group hover:border-orange-500/30 transition-all">
                        <div className="w-8 h-8 bg-orange-500/10 rounded-lg flex items-center justify-center text-orange-400 shrink-0 mt-1">
                          <span className="text-xs font-bold">{idx + 1}</span>
                        </div>
                        <div>
                          <p className="text-sm font-bold text-white mb-1 uppercase tracking-tight">{pain.point}</p>
                          <p className="text-xs text-slate-400 leading-relaxed">{pain.impact}</p>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>

                {/* Service Interest Prediction */}
                <div className="glass p-8 rounded-[2rem] border-white/5">
                  <div className="flex items-center gap-3 mb-8">
                    <div className="w-10 h-10 bg-purple-600/20 rounded-xl flex items-center justify-center text-purple-400">
                      <BarChart className="w-6 h-6" />
                    </div>
                    <h4 className="text-lg font-black text-white tracking-tight font-brand uppercase">Service Interest Prediction</h4>
                  </div>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-x-12 gap-y-8">
                    {Object.entries(result.serviceInterest).map(([service, score]) => (
                      <div key={`interest-${service}`} className="space-y-3">
                        <div className="flex items-center justify-between">
                          <p className="text-[10px] font-mono font-bold text-slate-400 uppercase tracking-widest">{service.replace(/([A-Z])/g, ' $1')}</p>
                          <p className="text-[10px] font-mono font-bold text-purple-400">{score}%</p>
                        </div>
                        <div className="h-2 w-full bg-white/5 rounded-full overflow-hidden">
                          <motion.div 
                            initial={{ width: 0 }}
                            animate={{ width: `${score}%` }}
                            transition={{ duration: 1, delay: 0.5 }}
                            className="h-full bg-gradient-to-r from-purple-600 to-purple-400 shadow-[0_0_10px_rgba(168,85,247,0.5)]"
                          />
                        </div>
                      </div>
                    ))}
                  </div>
                </div>

                {/* Strategic Expansion Card */}
                {result.teamComputersPitch.upsellCrossSellOpportunities && result.teamComputersPitch.upsellCrossSellOpportunities.length > 0 && (
                  <div className="glass p-8 rounded-[2.5rem] border-white/5 bg-gradient-to-br from-blue-900/10 to-transparent">
                    <div className="flex items-center gap-3 mb-8">
                      <div className="w-10 h-10 bg-blue-600/20 rounded-xl flex items-center justify-center text-blue-400">
                        <Zap className="w-6 h-6" />
                      </div>
                      <h4 className="text-lg font-black text-white tracking-tight font-brand uppercase">Expansion Opportunities</h4>
                    </div>
                    <div className="space-y-4">
                      {result.teamComputersPitch.upsellCrossSellOpportunities.map((opp, idx) => (
                        <div key={`opp-${idx}`} className="glass p-5 rounded-2xl border-white/5 hover:border-blue-500/20 transition-all group">
                          <div className="flex items-center justify-between mb-3">
                            <p className="text-sm font-bold text-white group-hover:text-blue-400 transition-colors uppercase tracking-tight">{opp.title}</p>
                            <span className="px-2 py-0.5 bg-blue-500/10 border border-blue-500/20 rounded text-[8px] font-mono font-bold text-blue-300 uppercase">
                              {opp.targetService}
                            </span>
                          </div>
                          <p className="text-xs text-slate-400 leading-relaxed">
                            {opp.description}
                          </p>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {/* Strategic Pitch Deck Section */}
                <div className="glass p-8 rounded-[2.5rem] border-white/5 bg-gradient-to-br from-blue-900/10 to-transparent">
                  <div className="flex items-center justify-between mb-8">
                    <div className="flex items-center gap-3">
                      <div className="w-10 h-10 bg-blue-600/20 rounded-xl flex items-center justify-center text-blue-400">
                        <Layout className="w-6 h-6" />
                      </div>
                      <div>
                        <h4 className="text-lg font-black text-white tracking-tight font-brand uppercase">Strategic Pitch Deck Outline</h4>
                        <p className="text-[10px] font-mono text-slate-500 uppercase tracking-widest">AI Generated Slide-by-Slide Strategy</p>
                      </div>
                    </div>
                  </div>
                  
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    {[
                      { title: "I. Introduction", content: result.pitchDeck.introduction },
                      { title: "II. Problem Statement", content: result.pitchDeck.problemStatement },
                      { title: "III. Solution", content: result.pitchDeck.solution },
                      { title: "IV. Market Analysis", content: result.pitchDeck.marketAnalysis },
                      { title: "V. Competitive Advantage", content: result.pitchDeck.competitiveAdvantage },
                      { title: "VI. Proposed Solution", content: result.pitchDeck.proposedSolution },
                    ].map((slide, idx) => (
                      <div key={`slide-${idx}`} className="glass p-5 rounded-2xl border-white/5 bg-white/2 hover:border-blue-500/20 transition-all">
                        <h5 className="text-xs font-mono font-bold text-blue-400 uppercase tracking-[0.2em] mb-3">{slide.title}</h5>
                        <p className="text-[11px] text-slate-300 leading-relaxed font-medium">{slide.content}</p>
                      </div>
                    ))}
                    
                    <div className="glass p-5 rounded-2xl border-white/5 bg-blue-600/5 hover:border-blue-500/20 transition-all md:col-span-2">
                       <h5 className="text-xs font-mono font-bold text-blue-400 uppercase tracking-[0.2em] mb-3">VII. Relevant Case Studies</h5>
                       <ul className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                         {result.pitchDeck.caseStudies.map((cs, i) => (
                           <li key={`cs-${i}`} className="flex items-start gap-2 text-[11px] text-slate-300 font-medium bg-white/5 p-2 rounded-lg">
                             <CheckCircle2 className="w-3.5 h-3.5 text-blue-400 shrink-0 mt-0.5" />
                             <span>{cs}</span>
                           </li>
                         ))}
                       </ul>
                    </div>

                    <div className="glass p-6 rounded-2xl border-blue-500/30 bg-blue-600/10 hover:border-blue-500 transition-all md:col-span-2 text-center">
                       <h5 className="text-xs font-mono font-bold text-blue-400 uppercase tracking-[0.2em] mb-3">VIII. Final Call to Action</h5>
                       <p className="text-sm text-white font-black italic">"{result.pitchDeck.callToAction}"</p>
                    </div>
                  </div>
                </div>

                {/* Team Computers Impact Card */}
                <div className="glass p-8 rounded-[2.5rem] border-white/5 bg-gradient-to-br from-green-900/10 to-transparent">
                  <div className="flex items-center gap-3 mb-8">
                    <div className="w-10 h-10 bg-green-600/20 rounded-xl flex items-center justify-center text-green-400">
                      <TrendingUp className="w-6 h-6" />
                    </div>
                    <h4 className="text-lg font-black text-white tracking-tight font-brand uppercase">Team Computers Impact</h4>
                  </div>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    {result.teamComputersImpact.map((impact, idx) => (
                      <div key={`impact-${idx}`} className="space-y-2 relative pl-6 border-l border-green-500/30">
                        <div className="absolute top-0 left-0 w-1.5 h-1.5 bg-green-500 rounded-full -ml-[4px] mt-1.5 shadow-[0_0_8px_rgba(34,197,94,0.5)]" />
                        <p className="text-sm font-bold text-white uppercase tracking-tight">{impact.benefit}</p>
                        <p className="text-xs text-slate-400 leading-relaxed">{impact.description}</p>
                        <div className="mt-2 inline-block px-3 py-1 bg-green-500/10 border border-green-500/20 rounded-lg">
                          <p className="text-[10px] font-mono text-green-400 uppercase tracking-widest font-bold">
                            Revenue Impact: <span className="text-green-300 font-black underline decoration-green-500/30 underline-offset-2">{impact.revenueImpact}</span>
                          </p>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              {/* Right Column: Strategic Insights */}
              <div className="lg:col-span-4 space-y-6">
                {/* Pitch Card */}
                <div className="bg-gradient-to-br from-blue-700 to-blue-900 rounded-[2.5rem] p-8 shadow-2xl shadow-blue-900/50 relative overflow-hidden group">
                  <div className="absolute inset-0 bg-grid opacity-10" />
                  <div className="absolute bottom-0 right-0 w-48 h-48 bg-white/10 blur-[60px] -mb-24 -mr-24" />
                  
                  <div className="relative z-10">
                    <div className="flex items-center gap-3 mb-8">
                      <div className="w-12 h-12 bg-white rounded-2xl flex items-center justify-center shadow-lg">
                        <Target className="text-blue-700 w-7 h-7" />
                      </div>
                      <h4 className="text-2xl font-black text-white tracking-tight font-brand uppercase">Strategic Pitch</h4>
                    </div>

                    <div className="space-y-8">
                      <div>
                        <p className="text-blue-200 text-[10px] font-mono font-bold uppercase tracking-[0.2em] mb-4">OEM TARGETS</p>
                        <div className="flex flex-wrap gap-2">
                          {(result.itInfrastructure.interestedOEMs || []).map((oem, idx) => (
                            <span key={`oem-${idx}`} className="px-4 py-1.5 bg-white/10 backdrop-blur-md border border-white/10 rounded-xl text-xs font-bold">
                              {oem}
                            </span>
                          ))}
                        </div>
                      </div>

                      <div>
                        <p className="text-blue-200 text-[10px] font-mono font-bold uppercase tracking-[0.2em] mb-4">DETAILED OPPORTUNITIES</p>
                        <div className="space-y-6">
                          {(result.teamComputersPitch.useCases || []).map((useCase, idx) => (
                            <div key={`usecase-${idx}`} className="space-y-2">
                              <div className="flex items-center gap-2">
                                <div className="w-5 h-5 bg-white/20 rounded-lg flex items-center justify-center shrink-0 text-[10px] font-black">
                                  {idx + 1}
                                </div>
                                <div className="flex flex-col">
                                  <p className="text-sm font-bold text-white uppercase tracking-tight">{useCase.title}</p>
                                  <span className="text-[8px] font-mono text-blue-300 uppercase tracking-widest">{useCase.serviceCategory}</span>
                                </div>
                              </div>
                              <p className="text-xs text-blue-100 leading-relaxed ml-7">{useCase.description}</p>
                              <div className="ml-7 mt-2 p-2 bg-white/5 rounded-lg border border-white/5">
                                <p className="text-[10px] font-mono text-blue-300 uppercase tracking-widest font-bold mb-1">Key Takeaway</p>
                                <p className="text-[11px] text-white font-medium italic">"{useCase.keyTakeaway}"</p>
                              </div>

                              <div className="ml-7 grid grid-cols-2 gap-3 mt-3">
                                <div className="p-2.5 bg-green-500/10 border border-green-500/20 rounded-xl">
                                  <div className="flex items-center gap-1.5 mb-1">
                                    <BarChart3 className="w-2.5 h-2.5 text-green-400" />
                                    <span className="text-[8px] font-mono text-green-300 uppercase tracking-widest font-bold">ROI Projection</span>
                                  </div>
                                  <p className="text-[10px] text-white font-bold leading-tight">{useCase.roiProjection}</p>
                                </div>
                                <div className="p-2.5 bg-purple-500/10 border border-purple-500/20 rounded-xl">
                                  <div className="flex items-center gap-1.5 mb-1">
                                    <Zap className="w-2.5 h-2.5 text-purple-400" />
                                    <span className="text-[8px] font-mono text-purple-300 uppercase tracking-widest font-bold">Timeline</span>
                                  </div>
                                  <p className="text-[10px] text-white font-bold leading-tight">{useCase.implementationTimeline}</p>
                                </div>
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>

                      <div className="pt-6 border-t border-white/10">
                        <button 
                          onClick={() => setShowCallScript(true)}
                          className="w-full py-3 bg-white/10 hover:bg-white/20 border border-white/10 rounded-xl flex items-center justify-center gap-2 text-xs font-bold text-white transition-all group/btn"
                        >
                          <MessageSquare className="w-4 h-4 text-blue-300 group-hover/btn:scale-110 transition-transform" />
                          VIEW INTERACTIVE CALL SCRIPT
                        </button>
                      </div>

                      <div className="pt-8 border-t border-white/10">
                        <div className="flex items-center gap-2 mb-3">
                          <TrendingUp className="w-4 h-4 text-blue-300" />
                          <p className="text-blue-200 text-[10px] font-mono font-bold uppercase tracking-[0.2em]">ADVANTAGE</p>
                        </div>
                        <p className="text-sm leading-relaxed text-white italic font-medium">
                          "{result.teamComputersPitch.strategicAdvantage}"
                        </p>
                      </div>
                    </div>
                  </div>
                </div>

                {/* Competitive Landscape Card */}
                <div className="glass p-8 rounded-[2.5rem] border-white/5">
                  <div className="flex items-center gap-3 mb-8">
                    <div className="w-10 h-10 bg-red-600/20 rounded-xl flex items-center justify-center text-red-400">
                      <BarChart3 className="w-6 h-6" />
                    </div>
                    <h4 className="text-lg font-black text-white tracking-tight font-brand uppercase">Competitive Comparison</h4>
                  </div>
                  
                  <div className="overflow-x-auto -mx-4 px-4 sm:mx-0 sm:px-0 max-h-[600px] overflow-y-auto custom-scrollbar">
                    <table className="w-full border-collapse">
                      <thead className="sticky top-0 z-20">
                        <tr className="bg-[#0F172A] border-b border-white/10">
                          <th className="text-left py-4 px-4 text-[10px] font-mono text-slate-500 uppercase tracking-widest">Competitor</th>
                          <th className="text-left py-4 px-4 text-[10px] font-mono text-slate-500 uppercase tracking-widest">Involved Tech Stack</th>
                          <th className="text-left py-4 px-4 text-[10px] font-mono text-slate-500 uppercase tracking-widest">Service Portfolio</th>
                          <th className="text-left py-4 px-4 text-[10px] font-mono text-slate-500 uppercase tracking-widest">Market Benefit</th>
                          <th className="text-left py-4 px-4 text-[10px] font-mono text-blue-400 uppercase tracking-widest bg-blue-500/5">Team Computers Edge</th>
                        </tr>
                      </thead>
                      <tbody>
                        {(result.competitiveLandscape || []).map((item, idx) => (
                          <tr key={`comp-${idx}`} className="border-b border-white/5 group/row hover:bg-blue-900/20 transition-all duration-300">
                            <td className="py-6 px-4">
                              <div className="flex items-center gap-3">
                                <div className="w-2 h-2 bg-red-500 rounded-full shadow-[0_0_8px_rgba(239,68,68,0.5)]" />
                                <span className="text-sm font-black text-white uppercase tracking-widest leading-tight">{item.competitorName}</span>
                              </div>
                            </td>
                            <td className="py-6 px-4">
                              <div className="flex flex-wrap gap-1">
                                {item.techStack?.map((tech, tIdx) => (
                                  <span key={tIdx} className="px-2 py-0.5 bg-white/5 border border-white/10 rounded text-[9px] font-mono text-slate-400">
                                    {tech}
                                  </span>
                                ))}
                              </div>
                            </td>
                            <td className="py-6 px-4">
                              <div className="flex flex-wrap gap-1">
                                {item.services?.map((svc, sIdx) => (
                                  <span key={sIdx} className="px-2 py-0.5 bg-blue-500/5 border border-blue-500/10 rounded text-[9px] font-mono text-blue-300">
                                    {svc}
                                  </span>
                                ))}
                              </div>
                            </td>
                            <td className="py-6 px-4">
                              <p className="text-xs text-slate-400 leading-relaxed max-w-[200px] italic">"{item.serviceBenefit}"</p>
                            </td>
                            <td className="py-6 px-4 bg-blue-500/5">
                              <div className="flex flex-col gap-3">
                                <div className="flex items-start gap-2">
                                  <CheckCircle2 className="w-4 h-4 text-blue-400 shrink-0 mt-0.5" />
                                  <p className="text-xs text-blue-100 font-bold leading-relaxed">{item.teamComputersEdge}</p>
                                </div>
                                <div className="space-y-1.5">
                                  <div className="flex items-center justify-between">
                                    <span className="text-[8px] font-mono text-blue-400 uppercase tracking-widest">Strength Meter</span>
                                    <span className="text-[8px] font-mono text-blue-300">85%</span>
                                  </div>
                                  <div className="h-1 w-full bg-white/5 rounded-full overflow-hidden">
                                    <motion.div 
                                      initial={{ width: 0 }}
                                      animate={{ width: '85%' }}
                                      className="h-full bg-gradient-to-r from-blue-600 to-blue-400 shadow-[0_0_10px_rgba(59,130,246,0.5)]"
                                    />
                                  </div>
                                </div>
                              </div>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* Sources & Methodology */}
                {result.sources && result.sources.length > 0 && (
                  <div className="lg:col-span-12 mt-4 focus-section" id="sources">
                    <div className="glass p-8 rounded-[2.5rem] border-white/5 bg-white/2">
                      <div className="flex flex-col md:flex-row md:items-center justify-between gap-6 mb-8">
                        <div className="flex items-center gap-3">
                          <div className="w-10 h-10 bg-slate-600/20 rounded-xl flex items-center justify-center text-slate-400">
                            <Database className="w-6 h-6" />
                          </div>
                          <div>
                            <h4 className="text-lg font-black text-white tracking-tight font-brand uppercase">Intelligence Sources & References</h4>
                            <p className="text-[10px] font-mono text-slate-500 uppercase tracking-widest">Verification data & methodology citations</p>
                          </div>
                        </div>
                        <div className="flex items-center gap-4 px-4 py-2 bg-blue-500/5 rounded-xl border border-blue-500/10">
                          <CheckCircle2 className="w-4 h-4 text-blue-400" />
                          <span className="text-[10px] font-mono text-blue-300 uppercase tracking-widest font-bold">Verified by Intelligence Synthesis</span>
                        </div>
                      </div>

                      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                        {result.sources.map((source, idx) => (
                          <a 
                            key={`source-${idx}`}
                            href={source.url}
                            target="_blank"
                            rel="noopener noreferrer" 
                            className="glass p-4 rounded-2xl border-white/5 hover:bg-white/5 hover:border-blue-500/30 transition-all group"
                          >
                            <div className="flex items-start gap-4">
                              <div className="w-10 h-10 bg-white/5 rounded-xl flex items-center justify-center shrink-0 group-hover:scale-110 transition-transform">
                                <FileText className="w-5 h-5 text-slate-500 group-hover:text-blue-400 transition-colors" />
                              </div>
                              <div className="flex-1 min-w-0">
                                <div className="flex items-center justify-between mb-1">
                                  <span className={cn(
                                    "text-[9px] font-mono font-bold uppercase tracking-widest",
                                    source.type.toLowerCase().includes('official') ? "text-green-400" :
                                    source.type.toLowerCase().includes('news') ? "text-orange-400" :
                                    "text-blue-400"
                                  )}>
                                    {source.type}
                                  </span>
                                  <ExternalLink className="w-3 h-3 text-slate-600 group-hover:text-blue-400 transition-colors" />
                                </div>
                                <p className="text-sm font-bold text-white truncate mb-1 group-hover:text-blue-200 transition-colors">{source.title}</p>
                                <p className="text-[10px] font-mono text-slate-500 truncate">{source.url}</p>
                              </div>
                            </div>
                          </a>
                        ))}
                      </div>

                      <div className="mt-8 pt-8 border-t border-white/5">
                        <div className="bg-blue-500/5 p-6 rounded-3xl border border-blue-500/10">
                          <div className="flex items-start gap-4">
                            <div className="w-10 h-10 bg-blue-600/10 rounded-2xl flex items-center justify-center text-blue-400 shrink-0">
                              <Info className="w-5 h-5" />
                            </div>
                            <p className="text-[11px] font-mono text-slate-400 leading-relaxed">
                              <span className="text-blue-400 font-bold uppercase tracking-widest block mb-2">Methodology Note</span>
                              These intelligence reports are generated using multi-agent synthesis of real-time public data including corporate filings, professional networking signals, infrastructure footprinting, and technology stacks. Numerical estimates for spend and ARR are derived from proprietary benchmarking models tailored for the Indian enterprise IT ecosystem. Sources provided above represent the primary validation points used for this profile.
                            </p>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Empty State */}
        {!isLoading && !result && summaries.length === 0 && !error && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            className="flex flex-col items-center justify-center py-32 text-slate-500"
          >
            <div className="relative mb-8">
              <div className="absolute inset-0 bg-blue-500/20 blur-3xl rounded-full" />
              <div className="w-24 h-24 glass border-white/10 rounded-[2rem] flex items-center justify-center relative z-10">
                {activeTab === 'single' && <Search className="w-10 h-10 text-slate-600" />}
                {activeTab === 'bulk' && <ListFilter className="w-10 h-10 text-slate-600" />}
                {activeTab === 'discovery' && <Sparkles className="w-10 h-10 text-slate-600" />}
              </div>
            </div>
            <p className="text-xl font-bold text-slate-400 tracking-tight font-brand">
              {activeTab === 'single' && 'Awaiting Input Signal...'}
              {activeTab === 'bulk' && 'Ready for Batch Processing...'}
              {activeTab === 'discovery' && 'Awaiting Discovery Parameters...'}
            </p>
            <p className="text-sm text-slate-600 mt-2">
              {activeTab === 'single' && 'Enter a target company to begin deep-scan.'}
              {activeTab === 'bulk' && 'Paste a list of domains to generate strategic summaries.'}
              {activeTab === 'discovery' && 'Describe your target market to find high-potential leads.'}
            </p>
          </motion.div>
        )}
        </main>

        {/* Footer */}
        <footer className="glass border-t border-white/5 py-8 shrink-0">
          <div className="w-full px-8 text-center">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-6">
                <div className="flex items-center gap-4 flex-shrink-0">
                  <div className="relative">
                    <motion.img 
                      src="https://www.teamcomputers.com/wp-content/uploads/2023/11/Team-Logo.png"
                      alt="Team Computers Logo"
                      className="h-8 w-auto opacity-80 hover:opacity-100 transition-all grayscale hover:grayscale-0"
                      referrerPolicy="no-referrer"
                      onError={(e) => {
                        e.currentTarget.style.display = 'none';
                        e.currentTarget.nextElementSibling?.classList.remove('hidden');
                      }}
                    />
                    <span className="hidden font-brand font-black text-white tracking-tight text-sm">Team Computers</span>
                  </div>
                  
                  <div className="h-4 w-px bg-white/10" />
                  
                  <div className="relative">
                    <motion.img 
                      src="https://upload.wikimedia.org/wikipedia/commons/thumb/7/7d/Intel_logo_%282020%2C_white%29.svg/1024px-Intel_logo_%282020%2C_white%29.svg.png"
                      alt="Intel Logo"
                      className="h-5 w-auto opacity-50 hover:opacity-100 transition-all grayscale hover:grayscale-0"
                      referrerPolicy="no-referrer"
                      onError={(e) => {
                        e.currentTarget.style.display = 'none';
                        e.currentTarget.nextElementSibling?.classList.remove('hidden');
                      }}
                    />
                    <span className="hidden font-brand font-bold text-white tracking-tight text-xs">Intel</span>
                  </div>
                </div>
                <div className="h-4 w-px bg-white/5 hidden sm:block" />
                <span className="font-brand font-black text-white/40 tracking-tight text-xs hidden sm:block">Strategic Partnership</span>
              </div>
              <p className="text-[10px] font-mono text-slate-500 uppercase tracking-widest">
                © 2026 Intelligence Operations Center. All rights reserved.
              </p>
              <div className="flex gap-8">
                <a href="https://teamcomputers.com" target="_blank" rel="noopener noreferrer" className="text-slate-500 hover:text-blue-400 transition-colors text-[10px] font-mono uppercase tracking-[0.2em] flex items-center gap-2">
                  Main Terminal <ExternalLink className="w-3 h-3" />
                </a>
              </div>
            </div>
          </div>
        </footer>
      </div>

      {/* Floating Chat Interface */}
      {result && (
        <>
          <motion.button
            initial={{ scale: 0 }}
            animate={{ scale: 1 }}
            onClick={() => setIsChatOpen(true)}
            className="fixed bottom-8 right-8 w-16 h-16 bg-blue-600 rounded-full flex items-center justify-center text-white shadow-2xl shadow-blue-600/40 z-[90] hover:bg-blue-500 transition-all group"
          >
            <MessageSquare className="w-8 h-8 group-hover:scale-110 transition-transform" />
            <div className="absolute -top-1 -right-1 w-5 h-5 bg-red-500 rounded-full border-2 border-[#0F172A] flex items-center justify-center text-[10px] font-bold">1</div>
          </motion.button>

          <AnimatePresence>
            {isChatOpen && (
              <div className="fixed inset-0 z-[110] flex justify-end p-4 pointer-events-none">
                <motion.div
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  exit={{ opacity: 0 }}
                  onClick={() => setIsChatOpen(false)}
                  className="absolute inset-0 bg-[#0F172A]/40 backdrop-blur-sm pointer-events-auto"
                />
                
                <motion.div
                  initial={{ x: '100%' }}
                  animate={{ x: 0 }}
                  exit={{ x: '100%' }}
                  transition={{ type: 'spring', damping: 25, stiffness: 200 }}
                  className="relative w-full max-w-md glass border-white/10 rounded-[2.5rem] flex flex-col pointer-events-auto shadow-2xl overflow-hidden"
                >
                  {/* Chat Header */}
                  <div className="p-6 border-b border-white/5 flex items-center justify-between bg-blue-600/10">
                    <div className="flex items-center gap-3">
                      <div className="w-10 h-10 bg-blue-600/20 rounded-xl flex items-center justify-center text-blue-400">
                        <MessageSquare className="w-6 h-6" />
                      </div>
                      <div>
                        <h4 className="text-sm font-black text-white font-brand uppercase tracking-tight">Intelligence Chat</h4>
                        <p className="text-[10px] font-mono text-blue-400 uppercase tracking-widest leading-none mt-1">Context: {result.name}</p>
                      </div>
                    </div>
                    <button 
                      onClick={() => setIsChatOpen(false)}
                      className="p-2 hover:bg-white/10 rounded-xl text-slate-400 hover:text-white transition-colors"
                    >
                      <X className="w-5 h-5" />
                    </button>
                  </div>

                  {/* Chat History */}
                  <div className="flex-1 overflow-y-auto p-6 space-y-6 custom-scrollbar">
                    {chatHistory.length === 0 && (
                      <div className="h-full flex flex-col items-center justify-center text-center opacity-40">
                        <Terminal className="w-12 h-12 mb-4 text-slate-600" />
                        <p className="text-sm font-bold text-slate-400 uppercase tracking-widest font-brand">Secure Intelligence Channel</p>
                        <p className="text-[10px] font-mono text-slate-600 mt-2 max-w-[200px]">Ask follow-up questions about the account report, competitors, or strategic pitch.</p>
                      </div>
                    )}
                    
                    {chatHistory.map((msg, idx) => (
                      <div 
                        key={`msg-${idx}`}
                        className={cn(
                          "flex flex-col gap-2",
                          msg.role === 'user' ? "items-end" : "items-start"
                        )}
                      >
                        <div className={cn(
                          "max-w-[85%] p-4 rounded-2xl text-xs leading-relaxed",
                          msg.role === 'user' 
                            ? "bg-blue-600 text-white rounded-tr-none shadow-lg shadow-blue-600/10" 
                            : "glass border-white/5 text-slate-200 rounded-tl-none"
                        )}>
                          {msg.content}
                        </div>
                        <span className="text-[9px] font-mono text-slate-600 uppercase font-bold tracking-widest">
                          {msg.role === 'user' ? 'Sales Rep' : 'Intelligence Engine'}
                        </span>
                      </div>
                    ))}
                    
                    {isChatLoading && (
                      <div className="flex flex-col items-start gap-2">
                        <div className="glass p-4 rounded-2xl rounded-tl-none border-white/5">
                          <Loader2 className="w-4 h-4 animate-spin text-blue-400" />
                        </div>
                        <span className="text-[9px] font-mono text-slate-500 uppercase tracking-widest font-bold">Synthesizing...</span>
                      </div>
                    )}
                    <div ref={chatEndRef} />
                  </div>

                  {/* Chat Input */}
                  <form onSubmit={handleChat} className="p-6 border-t border-white/5 bg-white/2">
                    <div className="relative group">
                      <div className="absolute -inset-1 bg-gradient-to-r from-blue-600/20 to-blue-400/20 rounded-xl blur opacity-0 group-focus-within:opacity-100 transition-opacity" />
                      <div className="relative flex items-center gap-2">
                        <input
                          type="text"
                          value={chatInput}
                          onChange={(e) => setChatInput(e.target.value)}
                          placeholder="Ask a question..."
                          className="flex-1 bg-white/5 border border-white/10 rounded-xl py-3 px-4 text-xs text-white outline-none focus:border-blue-500/40 transition-colors"
                        />
                        <button
                          type="submit"
                          disabled={!chatInput.trim() || isChatLoading}
                          className="p-3 bg-blue-600 text-white rounded-xl hover:bg-blue-500 disabled:opacity-50 transition-all shadow-lg shadow-blue-600/20"
                        >
                          <ChevronRight className="w-4 h-4" />
                        </button>
                      </div>
                    </div>
                  </form>
                </motion.div>
              </div>
            )}
          </AnimatePresence>
        </>
      )}

      {/* Call Script Modal */}
      <AnimatePresence>
        {showCallScript && result && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 sm:p-6">
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowCallScript(false)}
              className="absolute inset-0 bg-[#0F172A]/80 backdrop-blur-md"
            />
            <motion.div
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="relative w-full max-w-4xl bg-[#0a0a0a] border border-blue-500/30 rounded-[2rem] overflow-hidden shadow-[0_0_50px_rgba(59,130,246,0.2)] font-mono"
            >
              <div className="flex items-center justify-between p-6 border-b border-white/10 bg-white/5">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 bg-blue-600/20 rounded-xl flex items-center justify-center text-blue-400 shadow-[0_0_15px_rgba(59,130,246,0.3)]">
                    <Terminal className="w-6 h-6" />
                  </div>
                  <div>
                    <h3 className="text-lg font-black text-white tracking-tight uppercase">Sales Intelligence Terminal</h3>
                    <p className="text-[10px] text-blue-400 uppercase tracking-widest">Target: {result.name} // Session Active</p>
                  </div>
                </div>
                <button 
                  onClick={() => setShowCallScript(false)}
                  className="p-2 hover:bg-white/10 rounded-full text-slate-400 hover:text-white transition-colors"
                >
                  <X className="w-6 h-6" />
                </button>
              </div>

              <div className="p-8 overflow-y-auto max-h-[75vh] custom-scrollbar space-y-4">
                {/* Introduction Section */}
                <div className="glass border-white/5 rounded-2xl overflow-hidden">
                  <div 
                    onClick={() => toggleSection('introduction')}
                    className="w-full flex items-center justify-between p-4 hover:bg-white/5 transition-colors cursor-pointer"
                  >
                    <div className="flex items-center gap-3">
                      <button 
                        onClick={(e) => {
                          e.stopPropagation();
                          toggleStep('intro');
                        }}
                        className={cn(
                          "w-5 h-5 rounded border transition-all flex items-center justify-center",
                          completedSteps.has('intro') ? "bg-blue-600 border-blue-500" : "border-white/20"
                        )}
                      >
                        {completedSteps.has('intro') && <Check className="w-3.5 h-3.5 text-white" />}
                      </button>
                      <span className="text-blue-500 font-bold">01.</span>
                      <span className={cn(
                        "text-xs font-bold uppercase tracking-widest transition-all",
                        completedSteps.has('intro') ? "text-slate-500 line-through" : "text-white"
                      )}>Mission Briefing (Intro)</span>
                    </div>
                    {openSections.includes('introduction') ? <ChevronUp className="w-4 h-4 text-slate-500" /> : <ChevronDown className="w-4 h-4 text-slate-500" />}
                  </div>
                  <AnimatePresence>
                    {openSections.includes('introduction') && (
                      <motion.div
                        initial={{ height: 0, opacity: 0 }}
                        animate={{ height: 'auto', opacity: 1 }}
                        exit={{ height: 0, opacity: 0 }}
                        className="overflow-hidden"
                      >
                        <div className="p-4 pt-0 space-y-4">
                          <div className="p-4 bg-blue-500/5 border border-blue-500/20 rounded-xl relative group">
                            <p className="text-sm text-blue-100 leading-relaxed italic">
                              "{result.teamComputersPitch.callScript?.introduction || 'No introduction available.'}"
                            </p>
                            <button 
                              onClick={() => copySection(result.teamComputersPitch.callScript?.introduction || '', 'intro')}
                              className="absolute top-2 right-2 p-2 bg-white/5 hover:bg-white/10 rounded-lg text-slate-400 hover:text-white transition-all"
                            >
                              {copiedSection === 'intro' ? <Check className="w-3 h-3 text-green-400" /> : <Copy className="w-3 h-3" />}
                            </button>
                          </div>
                        </div>
                      </motion.div>
                    )}
                  </AnimatePresence>
                </div>

                {/* Discovery Section */}
                <div className="glass border-white/5 rounded-2xl overflow-hidden">
                  <div 
                    onClick={() => toggleSection('discovery')}
                    className="w-full flex items-center justify-between p-4 hover:bg-white/5 transition-colors cursor-pointer"
                  >
                    <div className="flex items-center gap-3">
                      <button 
                        onClick={(e) => {
                          e.stopPropagation();
                          toggleStep('discovery');
                        }}
                        className={cn(
                          "w-5 h-5 rounded border transition-all flex items-center justify-center",
                          completedSteps.has('discovery') ? "bg-blue-600 border-blue-500" : "border-white/20"
                        )}
                      >
                        {completedSteps.has('discovery') && <Check className="w-3.5 h-3.5 text-white" />}
                      </button>
                      <span className="text-blue-500 font-bold">02.</span>
                      <span className={cn(
                        "text-xs font-bold uppercase tracking-widest transition-all",
                        completedSteps.has('discovery') ? "text-slate-500 line-through" : "text-white"
                      )}>Discovery Phase</span>
                    </div>
                    {openSections.includes('discovery') ? <ChevronUp className="w-4 h-4 text-slate-500" /> : <ChevronDown className="w-4 h-4 text-slate-500" />}
                  </div>
                  <AnimatePresence>
                    {openSections.includes('discovery') && (
                      <motion.div
                        initial={{ height: 0, opacity: 0 }}
                        animate={{ height: 'auto', opacity: 1 }}
                        exit={{ height: 0, opacity: 0 }}
                        className="overflow-hidden"
                      >
                        <div className="p-4 pt-0 space-y-3">
                          {(result.teamComputersPitch.callScript?.discoveryQuestions || []).map((q, i) => (
                            <div key={`q-${i}`} className="flex items-start gap-3 p-4 bg-white/5 border border-white/5 rounded-xl group relative">
                              <span className="text-blue-500 font-bold text-xs mt-0.5">{i + 1}.</span>
                              <p className="text-sm text-slate-300 leading-relaxed pr-8">{q}</p>
                              <button 
                                onClick={() => copySection(q, `q-${i}`)}
                                className="absolute top-2 right-2 p-2 opacity-0 group-hover:opacity-100 bg-white/5 hover:bg-white/10 rounded-lg text-slate-400 hover:text-white transition-all"
                              >
                                {copiedSection === `q-${i}` ? <Check className="w-3 h-3 text-green-400" /> : <Copy className="w-3 h-3" />}
                              </button>
                            </div>
                          ))}
                        </div>
                      </motion.div>
                    )}
                  </AnimatePresence>
                </div>

                {/* Value Prop Section */}
                <div className="glass border-white/5 rounded-2xl overflow-hidden">
                  <div 
                    onClick={() => toggleSection('value')}
                    className="w-full flex items-center justify-between p-4 hover:bg-white/5 transition-colors cursor-pointer"
                  >
                    <div className="flex items-center gap-3">
                      <button 
                        onClick={(e) => {
                          e.stopPropagation();
                          toggleStep('value');
                        }}
                        className={cn(
                          "w-5 h-5 rounded border transition-all flex items-center justify-center",
                          completedSteps.has('value') ? "bg-blue-600 border-blue-500" : "border-white/20"
                        )}
                      >
                        {completedSteps.has('value') && <Check className="w-3.5 h-3.5 text-white" />}
                      </button>
                      <span className="text-blue-500 font-bold">03.</span>
                      <span className={cn(
                        "text-xs font-bold uppercase tracking-widest transition-all",
                        completedSteps.has('value') ? "text-slate-500 line-through" : "text-white"
                      )}>Strategic Value Prop</span>
                    </div>
                    {openSections.includes('value') ? <ChevronUp className="w-4 h-4 text-slate-500" /> : <ChevronDown className="w-4 h-4 text-slate-500" />}
                  </div>
                  <AnimatePresence>
                    {openSections.includes('value') && (
                      <motion.div
                        initial={{ height: 0, opacity: 0 }}
                        animate={{ height: 'auto', opacity: 1 }}
                        exit={{ height: 0, opacity: 0 }}
                        className="overflow-hidden"
                      >
                        <div className="p-4 pt-0">
                          <div className="p-4 bg-blue-600/10 border border-blue-500/30 rounded-xl relative group">
                            <p className="text-sm text-blue-100 leading-relaxed">
                              {result.teamComputersPitch.callScript?.valueProposition || 'No value proposition available.'}
                            </p>
                            <button 
                              onClick={() => copySection(result.teamComputersPitch.callScript?.valueProposition || '', 'value')}
                              className="absolute top-2 right-2 p-2 bg-white/5 hover:bg-white/10 rounded-lg text-slate-400 hover:text-white transition-all"
                            >
                              {copiedSection === 'value' ? <Check className="w-3 h-3 text-green-400" /> : <Copy className="w-3 h-3" />}
                            </button>
                          </div>
                        </div>
                      </motion.div>
                    )}
                  </AnimatePresence>
                </div>

                {/* Objections Section */}
                <div className="glass border-white/5 rounded-2xl overflow-hidden">
                  <div 
                    onClick={() => toggleSection('objections')}
                    className="w-full flex items-center justify-between p-4 hover:bg-white/5 transition-colors cursor-pointer"
                  >
                    <div className="flex items-center gap-3">
                      <button 
                        onClick={(e) => {
                          e.stopPropagation();
                          toggleStep('objections');
                        }}
                        className={cn(
                          "w-5 h-5 rounded border transition-all flex items-center justify-center",
                          completedSteps.has('objections') ? "bg-red-600 border-red-500" : "border-white/20"
                        )}
                      >
                        {completedSteps.has('objections') && <Check className="w-3.5 h-3.5 text-white" />}
                      </button>
                      <span className="text-red-500 font-bold">04.</span>
                      <span className={cn(
                        "text-xs font-bold uppercase tracking-widest transition-all",
                        completedSteps.has('objections') ? "text-slate-500 line-through" : "text-white"
                      )}>Counter-Objections</span>
                    </div>
                    {openSections.includes('objections') ? <ChevronUp className="w-4 h-4 text-slate-500" /> : <ChevronDown className="w-4 h-4 text-slate-500" />}
                  </div>
                  <AnimatePresence>
                    {openSections.includes('objections') && (
                      <motion.div
                        initial={{ height: 0, opacity: 0 }}
                        animate={{ height: 'auto', opacity: 1 }}
                        exit={{ height: 0, opacity: 0 }}
                        className="overflow-hidden"
                      >
                        <div className="p-4 pt-0 space-y-3">
                          {(result.teamComputersPitch.callScript?.handlingObjections || []).map((obj, i) => (
                            <div key={`obj-${i}`} className="flex items-start gap-3 p-4 bg-red-500/5 border border-red-500/10 rounded-xl group relative">
                              <AlertCircle className="w-4 h-4 text-red-400 shrink-0 mt-0.5" />
                              <p className="text-sm text-slate-300 leading-relaxed pr-8">{obj}</p>
                              <button 
                                onClick={() => copySection(obj, `obj-${i}`)}
                                className="absolute top-2 right-2 p-2 opacity-0 group-hover:opacity-100 bg-white/5 hover:bg-white/10 rounded-lg text-slate-400 hover:text-white transition-all"
                              >
                                {copiedSection === `obj-${i}` ? <Check className="w-3 h-3 text-green-400" /> : <Copy className="w-3 h-3" />}
                              </button>
                            </div>
                          ))}
                        </div>
                      </motion.div>
                    )}
                  </AnimatePresence>
                </div>

                {/* Closing Section */}
                <div className="glass border-white/5 rounded-2xl overflow-hidden">
                  <div 
                    onClick={() => toggleSection('closing')}
                    className="w-full flex items-center justify-between p-4 hover:bg-white/5 transition-colors cursor-pointer"
                  >
                    <div className="flex items-center gap-3">
                      <button 
                        onClick={(e) => {
                          e.stopPropagation();
                          toggleStep('closing');
                        }}
                        className={cn(
                          "w-5 h-5 rounded border transition-all flex items-center justify-center",
                          completedSteps.has('closing') ? "bg-green-600 border-green-500" : "border-white/20"
                        )}
                      >
                        {completedSteps.has('closing') && <Check className="w-3.5 h-3.5 text-white" />}
                      </button>
                      <span className="text-green-500 font-bold">05.</span>
                      <span className={cn(
                        "text-xs font-bold uppercase tracking-widest transition-all",
                        completedSteps.has('closing') ? "text-slate-500 line-through" : "text-white"
                      )}>Mission Close</span>
                    </div>
                    {openSections.includes('closing') ? <ChevronUp className="w-4 h-4 text-slate-500" /> : <ChevronDown className="w-4 h-4 text-slate-500" />}
                  </div>
                  <AnimatePresence>
                    {openSections.includes('closing') && (
                      <motion.div
                        initial={{ height: 0, opacity: 0 }}
                        animate={{ height: 'auto', opacity: 1 }}
                        exit={{ height: 0, opacity: 0 }}
                        className="overflow-hidden"
                      >
                        <div className="p-4 pt-0">
                          <div className="p-4 bg-green-600/10 border border-green-500/30 rounded-xl relative group">
                            <p className="text-sm text-green-100 leading-relaxed font-bold">
                              "{result.teamComputersPitch.callScript?.closing || 'No closing script available.'}"
                            </p>
                            <button 
                              onClick={() => copySection(result.teamComputersPitch.callScript?.closing || '', 'closing')}
                              className="absolute top-2 right-2 p-2 bg-white/5 hover:bg-white/10 rounded-lg text-slate-400 hover:text-white transition-all"
                            >
                              {copiedSection === 'closing' ? <Check className="w-3 h-3 text-green-400" /> : <Copy className="w-3 h-3" />}
                            </button>
                          </div>
                        </div>
                      </motion.div>
                    )}
                  </AnimatePresence>
                </div>
              </div>

              <div className="p-6 bg-white/5 border-t border-white/10 flex justify-between items-center">
                <div className="flex items-center gap-2">
                  <div className="w-2 h-2 bg-green-500 rounded-full animate-pulse" />
                  <span className="text-[10px] text-slate-500 uppercase tracking-widest">Secure Connection Established</span>
                </div>
                <button 
                  onClick={() => setShowCallScript(false)}
                  className="px-8 py-2 bg-blue-600 hover:bg-blue-500 text-white font-bold rounded-xl transition-all shadow-[0_0_20px_rgba(59,130,246,0.3)]"
                >
                  TERMINATE SESSION
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

