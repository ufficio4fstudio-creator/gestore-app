
import * as XLSX from 'xlsx';
import React, { useState, useEffect, useMemo, useRef } from 'react';
import { 
  Users, Zap, FileText, Plus, Search, MapPin, Phone, Mail, ChevronRight,
  TrendingUp, Files, ShieldAlert, Activity, LogOut, Lock, User as UserIcon,
  ChevronDown, X, Download, HardDrive, AlertCircle, FileSpreadsheet, Paperclip,
  Edit2, Trash2, Info, Clock, Check, ExternalLink, Hash, BarChart3, Upload, Menu, Flame, MessageSquare, Send, Bell, BellOff,
  Wallet, FileSignature, Newspaper, Lightbulb
} from 'lucide-react';
import { io, Socket } from 'socket.io-client';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer 
} from 'recharts';
import { MapContainer, TileLayer, Marker, Popup, useMap } from 'react-leaflet';
import L from 'leaflet';
import 'leaflet/dist/leaflet.css';
import { 
  Portfolio, PDP, Contract, AppView, User, UserRole, 
  CabinaPrimaria, Caso, CaseCategory, Attachment, ContractStatus, Fornitore, CaseStatus,
  CreditCheckRequest, CreditCheckStatus, ContractType, PortfolioEntityType, Message, News, Associazione, AssociazioneStatus
} from './types';

// Fix Leaflet Icon issue
const DefaultIcon = L.icon({
  iconUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon.png',
  shadowUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-shadow.png',
  iconSize: [25, 41],
  iconAnchor: [12, 41],
});
L.Marker.prototype.options.icon = DefaultIcon;

const FILE_SIZE_LIMIT = 5 * 1024 * 1024; // 5MB
const FORNITORI: Fornitore[] = ['A2A1', 'A2A2', 'AXPO', 'AXP2', 'DOLO', 'DUFE', 'OPEN', 'SORG'];
const CONTRACT_TYPES: ContractType[] = ['SWITCH', 'ATTIVAZIONE', 'VOLT TIT III', 'VOLT TIT IV', 'ALLACCIO', 'RINNOVO'];

const MODULISTICA_URL = import.meta.env.VITE_MODULISTICA_URL || '#';
const OFFERTE_URL = import.meta.env.VITE_OFFERTE_URL || '#';

const calculateActivationDate = (dateStr: string) => {
  if (!dateStr) return '';
  const [y, m, d] = dateStr.split('-').map(Number);
  const date = new Date(Date.UTC(y, m - 1, d));
  if (isNaN(date.getTime())) return '';
  
  // Aggiunge 28 giorni alla data di creazione
  date.setUTCDate(date.getUTCDate() + 28);
  
  // Restituisce il primo giorno del mese successivo a quello della data risultante
  const nextMonth = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + 1, 1));
  return nextMonth.toISOString().split('T')[0];
};

const calculateEndDate = (startDateStr: string, durationMonths: number): string => {
  if (!startDateStr) return '';
  const [y, m, d] = startDateStr.split('-').map(Number);
  const start = new Date(Date.UTC(y, m - 1, d));
  if (isNaN(start.getTime())) return '';
  
  // Calcola l'ultimo giorno del mese aggiungendo i mesi indicati dalla durata
  const end = new Date(Date.UTC(start.getUTCFullYear(), start.getUTCMonth() + durationMonths, 0));
  return end.toISOString().split('T')[0];
};

const formatDate = (date: any) => {
  if (!date) return '';
  
  // Se è una stringa nel formato YYYY-MM-DD, la formattiamo direttamente per evitare shift di fuso orario
  if (typeof date === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(date)) {
    const [y, m, d] = date.split('-');
    return `${d}/${m}/${y}`;
  }

  let d = new Date(date);
  if (isNaN(d.getTime()) && typeof date === 'string') {
    // Handle DD/MM/YYYY or DD-MM-YYYY or DD.MM.YYYY
    const parts = date.split(/[ /.-]/).filter(Boolean);
    if (parts.length >= 3) {
      const day = parseInt(parts[0]);
      const month = parseInt(parts[1]) - 1;
      let year = parseInt(parts[2]);
      if (year < 100) year += 2000;
      d = new Date(year, month, day);
    }
  }
  if (isNaN(d.getTime())) return String(date);
  
  // Se la data è stata creata da una stringa ISO (es. YYYY-MM-DD), usiamo i metodi UTC
  // Altrimenti usiamo i metodi locali per date create localmente
  const isISO = typeof date === 'string' && date.includes('T');
  const day = String(isISO ? d.getUTCDate() : d.getDate()).padStart(2, '0');
  const month = String((isISO ? d.getUTCMonth() : d.getMonth()) + 1).padStart(2, '0');
  const year = String(isISO ? d.getUTCFullYear() : d.getFullYear());
  return `${day}/${month}/${year}`;
};

const formatDateTime = (date: any) => {
  if (!date) return '';
  let d = new Date(date);
  if (isNaN(d.getTime()) && typeof date === 'string') {
    const parts = date.split(/[ /.-]/).filter(Boolean);
    if (parts.length >= 3) {
      const day = parseInt(parts[0]);
      const month = parseInt(parts[1]) - 1;
      let year = parseInt(parts[2]);
      if (year < 100) year += 2000;
      const hours = parts[3] ? parseInt(parts[3]) : 0;
      const minutes = parts[4] ? parseInt(parts[4]) : 0;
      d = new Date(year, month, day, hours, minutes);
    }
  }
  if (isNaN(d.getTime())) return String(date);

  const isISO = typeof date === 'string' && date.includes('T');
  const day = String(isISO ? d.getUTCDate() : d.getDate()).padStart(2, '0');
  const month = String((isISO ? d.getUTCMonth() : d.getMonth()) + 1).padStart(2, '0');
  const year = String(isISO ? d.getUTCFullYear() : d.getFullYear());
  const hours = String(isISO ? d.getUTCHours() : d.getHours()).padStart(2, '0');
  const minutes = String(isISO ? d.getUTCMinutes() : d.getMinutes()).padStart(2, '0');
  return `${day}/${month}/${year} ${hours}:${minutes}`;
};

const formatVolume = (val: number | string) => {
  if (!val && val !== 0) return '';
  return Number(val).toLocaleString('it-IT');
};

const INITIAL_USERS: User[] = [
  { 
    id: '1', 
    username: 'admin', 
    password: 'admin2025', 
    name: 'Amministratore', 
    role: 'admin', 
    userCode: 'ADM01',
    email: 'direzione@progettoanticrisi.it',
    notificationSettings: {
      email: { contractStatusChange: true, caseStatusChange: true, newMessages: true, mentions: true },
      googleChat: { enabled: false }
    }
  },
  { 
    id: '2', 
    username: 'backoffice', 
    password: 'backoffice2024', 
    name: 'Operatore Backoffice', 
    role: 'backoffice', 
    userCode: 'BO01',
    email: 'backoffice@progettoanticrisi.it',
    notificationSettings: {
      email: { contractStatusChange: true, caseStatusChange: true, newMessages: true, mentions: true },
      googleChat: { enabled: false }
    }
  },
  { 
    id: '3', 
    username: 'utente1', 
    password: 'user1', 
    name: 'Agente Rossi', 
    role: 'utente', 
    userCode: 'AG01',
    email: 'agente.rossi@progettoanticrisi.it',
    notificationSettings: {
      email: { contractStatusChange: true, caseStatusChange: true, newMessages: true, mentions: true },
      googleChat: { enabled: false }
    }
  },
];

const CASE_CATEGORIES: CaseCategory[] = [
  'ANALISI', 'ANAGRAFICA', 'CONTENZIOSO', 'CONTRATTUALE', 'CREDITI', 'DISTRIBUZIONE', 'FATTURAZIONE'
].sort() as CaseCategory[];

const loadData = <T,>(key: string, initial: T[]): T[] => {
  const stored = localStorage.getItem(key);
  return stored ? JSON.parse(stored) : initial;
};

const getTimestamp = () => new Date().toISOString();

const splitAddress = (val: string) => {
  if (!val) return null;
  const parts = val.split(',').map(s => s.trim());
  if (parts.length >= 3) {
    const street = parts[0];
    const cityZip = parts[1].split(' ');
    const cap = cityZip[0].match(/\d{5}/) ? cityZip[0] : '';
    const city = cityZip.length > 1 ? cityZip.slice(1).join(' ') : cityZip[0];
    const prov = parts[parts.length - 1].split(' ')[0].substring(0, 2).toUpperCase();
    return { street, cap, city, province: prov };
  }
  return null;
};

const AddressLink: React.FC<{ address: string; label?: string; className?: string; showIconOnly?: boolean }> = ({ address, label, className, showIconOnly }) => (
  <a 
    href={`https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(address)}`} 
    target="_blank" 
    rel="noopener noreferrer"
    className={`inline-flex items-center gap-1 hover:text-emerald-600 transition-colors group ${className}`}
    title={address}
  >
    {!showIconOnly && (label || address)}
    <MapPin size={showIconOnly ? 16 : 12} className={showIconOnly ? "text-emerald-600" : "opacity-0 group-hover:opacity-100 transition-opacity"} />
  </a>
);

const App: React.FC = () => {
  const [currentUser, setCurrentUser] = useState<User | null>(() => {
    const stored = localStorage.getItem('charlie_current_user');
    return stored ? JSON.parse(stored) : null;
  });

  const [view, setView] = useState<AppView>('dashboard');
  const [users, setUsers] = useState<User[]>(() => {
    const data = loadData('users', INITIAL_USERS);
    // Reset admin password if it's the old one to ensure the reset takes effect
    return data.map(u => u.username === 'admin' && u.password === 'charlie2024' ? { ...u, password: 'admin2025' } : u);
  });
  const [portfolios, setPortfolios] = useState<Portfolio[]>(() => loadData('portfolios', []));
  const [pdps, setPdps] = useState<PDP[]>(() => loadData('pdps', []));
  const [contracts, setContracts] = useState<Contract[]>(() => loadData('contracts', []));
  const [cabine, setCabine] = useState<CabinaPrimaria[]>(() => loadData('cabine', []));
  const [casi, setCasi] = useState<Caso[]>(() => loadData('casi', []));
  const [creditChecks, setCreditChecks] = useState<CreditCheckRequest[]>(() => loadData('creditChecks', []));
  const [news, setNews] = useState<News[]>(() => loadData('news', []));
  const [associazioni, setAssociazioni] = useState<Associazione[]>(() => loadData('associazioni', []));
  const [selectedNews, setSelectedNews] = useState<News | null>(null);
  const [messages, setMessages] = useState<Message[]>([]);
  const [activeChat, setActiveChat] = useState<{ id: string; type: 'contract' | 'case' } | null>(null);
  const [showNotifications, setShowNotifications] = useState(false);
  const [showReportSubmenu, setShowReportSubmenu] = useState(false);
  const [prefilledPortfolioId, setPrefilledPortfolioId] = useState<string | null>(null);
  const socketRef = useRef<Socket | null>(null);

  // Migrazione data fine contratti per riflettere la nuova logica (ultimo giorno del mese)
  useEffect(() => {
    const migrated = localStorage.getItem('contracts_migrated_v3');
    if (!migrated && contracts.length > 0) {
      const updatedContracts = contracts.map(c => ({
        ...c,
        endDate: calculateEndDate(c.startDate, c.durationMonths || 12)
      }));
      setContracts(updatedContracts);
      localStorage.setItem('contracts_migrated_v3', 'true');
    }
  }, []);

  useEffect(() => {
    if (currentUser) {
      const socket = io();
      socketRef.current = socket;

      socket.on('init_messages', (msgs: Message[]) => {
        setMessages(msgs);
      });

      socket.on('new_message', (msg: Message) => {
        setMessages(prev => [...prev, msg]);
      });

      socket.on('message_read', (id: string) => {
        setMessages(prev => prev.map(m => m.id === id ? { ...m, isRead: true } : m));
      });

      socket.on('target_read', ({ targetId, userId }: { targetId: string, userId: string }) => {
        setMessages(prev => prev.map(m => 
          (m.targetId === targetId && m.senderId !== userId) ? { ...m, isRead: true } : m
        ));
      });

      return () => {
        socket.disconnect();
      };
    }
  }, [currentUser]);

  const sendMessage = (text: string, attachments?: Attachment[], asEmail: boolean = false) => {
    if (activeChat && currentUser && socketRef.current) {
      const msg = {
        targetId: activeChat.id,
        targetType: activeChat.type,
        senderId: currentUser.id,
        senderName: currentUser.name,
        senderRole: currentUser.role,
        text,
        attachments,
        asEmail,
      };
      socketRef.current.emit('send_message', msg);
    }
  };

  const markChatAsRead = (targetId: string) => {
    if (currentUser && socketRef.current) {
      socketRef.current.emit('mark_target_as_read', { targetId, userId: currentUser.id });
    }
  };

  const getUnreadCount = (targetId: string) => {
    return messages.filter(m => m.targetId === targetId && !m.isRead && m.senderId !== currentUser?.id).length;
  };

  const notifications = useMemo(() => {
    if (!currentUser) return [];
    
    return messages.filter(m => {
      if (m.isRead) return false;
      if (m.senderId === currentUser.id) return false;

      // Tagged (@name or @userCode)
      const isTagged = m.text.includes(`@${currentUser.name}`) || m.text.includes(`@${currentUser.userCode}`);
      if (isTagged) return true;

      // Backoffice sees everything
      if (currentUser.role === 'backoffice') return true;

      // Admin sees everything (optional, but user said "compreso admin hanno notifica" for tags, 
      // but let's assume admin also sees all unread for management)
      if (currentUser.role === 'admin') return true;

      // Agent sees their own
      if (currentUser.role === 'utente') {
        const contract = contracts.find(c => c.id === m.targetId);
        const caso = casi.find(c => c.id === m.targetId);
        if (contract && contract.assignedTo === currentUser.id) return true;
        if (caso && caso.assignedTo === currentUser.id) return true;
      }

      return false;
    });
  }, [messages, currentUser, contracts, casi]);

  const notificationGroups = useMemo(() => {
    const groups: { [key: string]: { id: string, type: 'contract' | 'case', count: number, lastMessage: Message } } = {};
    notifications.forEach(m => {
      if (!groups[m.targetId]) {
        groups[m.targetId] = { id: m.targetId, type: m.targetType, count: 0, lastMessage: m };
      }
      groups[m.targetId].count++;
      if (new Date(m.timestamp) > new Date(groups[m.targetId].lastMessage.timestamp)) {
        groups[m.targetId].lastMessage = m;
      }
    });
    return Object.values(groups).sort((a, b) => new Date(b.lastMessage.timestamp).getTime() - new Date(a.lastMessage.timestamp).getTime());
  }, [notifications]);

  useEffect(() => {
    localStorage.setItem('users', JSON.stringify(users));
    localStorage.setItem('portfolios', JSON.stringify(portfolios));
    localStorage.setItem('pdps', JSON.stringify(pdps));
    localStorage.setItem('contracts', JSON.stringify(contracts));
    localStorage.setItem('cabine', JSON.stringify(cabine));
    localStorage.setItem('casi', JSON.stringify(casi));
    localStorage.setItem('creditChecks', JSON.stringify(creditChecks));
    localStorage.setItem('news', JSON.stringify(news));
    localStorage.setItem('associazioni', JSON.stringify(associazioni));
  }, [users, portfolios, pdps, contracts, cabine, casi, creditChecks, news, associazioni]);

  const filteredData = useMemo(() => {
    const isAdminOrBO = currentUser?.role === 'admin' || currentUser?.role === 'backoffice';
    return {
      portfolios: isAdminOrBO ? portfolios : portfolios.filter(p => p.assignedTo === currentUser?.id),
      pdps: isAdminOrBO ? pdps : pdps.filter(p => p.assignedTo === currentUser?.id),
      contracts: isAdminOrBO ? contracts : contracts.filter(c => c.assignedTo === currentUser?.id),
      casi: isAdminOrBO ? casi : casi.filter(c => c.assignedTo === currentUser?.id),
      creditChecks: isAdminOrBO ? creditChecks : creditChecks.filter(s => s.assignedTo === currentUser?.id),
      associazioni: isAdminOrBO ? associazioni : associazioni.filter(a => a.assignedTo === currentUser?.id),
    };
  }, [currentUser, portfolios, pdps, contracts, casi, creditChecks, associazioni]);

  const handleLogin = (user: User) => {
    setCurrentUser(user);
    localStorage.setItem('charlie_current_user', JSON.stringify(user));
  };

  const handleLogout = (e?: React.MouseEvent) => {
    if(e) e.preventDefault();
    localStorage.removeItem('charlie_current_user');
    setCurrentUser(null);
    setView('dashboard'); // Reset della vista per il prossimo accesso
  };

  const exportToExcel = (data: any[], fileName: string) => {
    if (data.length === 0) return alert("Nessun dato da esportare");
    const cleanData = data.map(({ attachments, ...rest }) => rest);
    const ws = XLSX.utils.json_to_sheet(cleanData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, fileName);
    XLSX.writeFile(wb, `${fileName}_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  if (!currentUser) return <LoginScreen users={users} onLoginSuccess={handleLogin} />;

  const TPMRankingView: React.FC<{ stats: any, users: User[] }> = ({ stats, users }) => {
  const [selectedMonth, setSelectedMonth] = useState(new Date().getMonth());
  const [selectedYear, setSelectedYear] = useState(new Date().getFullYear());

  const months = [
    'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno',
    'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'
  ];

  const years = useMemo(() => {
    const currentYear = new Date().getFullYear();
    const startYear = 2024;
    const result = [];
    for (let y = currentYear + 1; y >= startYear; y--) result.push(y);
    return result;
  }, []);

  const ranking = useMemo(() => {
    const pdpFirstContractMap = new Map<string, string>();
    const allValidContracts = [...stats.contracts]
      .filter((c: Contract) => c.status !== 'KO')
      .sort((a, b) => new Date(a.creationDate || a.createdAt).getTime() - new Date(b.creationDate || b.createdAt).getTime());

    allValidContracts.forEach(c => {
      if (!pdpFirstContractMap.has(c.pdpId)) pdpFirstContractMap.set(c.pdpId, c.id);
    });

    const isJulyAugust = selectedMonth === 6 || selectedMonth === 7;
    const periodContracts = stats.contracts.filter((c: Contract) => {
      if (c.status === 'KO') return false;
      if (pdpFirstContractMap.get(c.pdpId) !== c.id) return false;
      const d = new Date(c.creationDate || c.createdAt);
      if (d.getFullYear() !== selectedYear) return false;
      return isJulyAugust ? (d.getMonth() === 6 || d.getMonth() === 7) : d.getMonth() === selectedMonth;
    });

    const periodAssociazioni = (stats.associazioni || []).filter((a: Associazione) => {
      if (a.status === 'nuovo' || a.isRenewal) return false;
      const d = new Date(a.startDate);
      if (d.getFullYear() !== selectedYear) return false;
      return isJulyAugust ? (d.getMonth() === 6 || d.getMonth() === 7) : d.getMonth() === selectedMonth;
    });

    const userStats = new Map<string, { tpm: number, volume: number }>();

    periodContracts.forEach((c: Contract) => {
      const portfolio = stats.portfolios.find((p: Portfolio) => p.id === c.portfolioId);
      if (!portfolio) return;

      const points = portfolio.entityType === 'Domestico' ? 1 : portfolio.entityType === 'Condominio' ? 2 : 3;
      const current = userStats.get(c.assignedTo) || { tpm: 0, volume: 0 };
      
      const weightedVolume = c.service === 'POWER' ? (c.volume || 0) : (c.volume || 0) * 3;
      
      userStats.set(c.assignedTo, {
        tpm: current.tpm + points,
        volume: current.volume + weightedVolume
      });
    });

    periodAssociazioni.forEach((a: Associazione) => {
      const portfolio = stats.portfolios.find((p: Portfolio) => p.id === a.portfolioId);
      if (!portfolio) return;

      const points = portfolio.entityType === 'Domestico' ? 3 : portfolio.entityType === 'Condominio' ? 5 : 8;
      const current = userStats.get(a.assignedTo) || { tpm: 0, volume: 0 };
      
      userStats.set(a.assignedTo, {
        tpm: current.tpm + points,
        volume: current.volume
      });
    });

    return Array.from(userStats.entries())
      .map(([userId, data]) => {
        const user = users.find(u => u.id === userId);
        return {
          userId,
          name: user?.name || 'Sconosciuto',
          userCode: user?.userCode || 'N/A',
          tpm: data.tpm,
          volume: Math.round(data.volume / 1000)
        };
      })
      .sort((a, b) => b.tpm - a.tpm);
  }, [stats.contracts, stats.portfolios, users, selectedMonth, selectedYear]);

  return (
    <div className="space-y-6 animate-in fade-in duration-500">
      <header className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h2 className="text-3xl font-black tracking-tighter text-emerald-900">Classifica TPM</h2>
          <p className="text-slate-500 font-medium italic">Performance e volumi per codice agente</p>
        </div>
        <div className="flex items-center gap-2 bg-white p-2 rounded-2xl border border-slate-100 shadow-sm">
          <select className="text-xs font-black uppercase bg-transparent outline-none px-2 cursor-pointer" value={selectedMonth} onChange={e => setSelectedMonth(Number(e.target.value))}>
            {months.map((m, i) => <option key={m} value={i}>{m}</option>)}
          </select>
          <select className="text-xs font-black uppercase bg-transparent outline-none px-2 cursor-pointer border-l" value={selectedYear} onChange={e => setSelectedYear(Number(e.target.value))}>
            {years.map(y => <option key={y} value={y}>{y}</option>)}
          </select>
        </div>
      </header>

      <div className="bg-white rounded-[2.5rem] border border-slate-100 shadow-xl overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse min-w-[600px]">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-100">
                <th className="p-6 text-[10px] font-black uppercase text-slate-400 tracking-widest">Pos.</th>
                <th className="p-6 text-[10px] font-black uppercase text-slate-400 tracking-widest">Codice</th>
                <th className="p-6 text-[10px] font-black uppercase text-slate-400 tracking-widest">Agente</th>
                <th className="p-6 text-[10px] font-black uppercase text-slate-400 tracking-widest text-right">Punti TPM</th>
                <th className="p-6 text-[10px] font-black uppercase text-slate-400 tracking-widest text-right">Volume Totale</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-50">
              {ranking.map((item, index) => (
                <tr key={item.userId} className="hover:bg-slate-50/50 transition-colors group">
                  <td className="p-6">
                    <div className={`w-8 h-8 rounded-lg flex items-center justify-center font-black text-sm ${index === 0 ? 'bg-yellow-100 text-yellow-700' : index === 1 ? 'bg-slate-200 text-slate-700' : index === 2 ? 'bg-orange-100 text-orange-700' : 'bg-slate-50 text-slate-400'}`}>
                      {index + 1}
                    </div>
                  </td>
                  <td className="p-6">
                    <span className="text-xs font-black text-emerald-600 bg-emerald-50 px-2 py-1 rounded-lg uppercase">{item.userCode}</span>
                  </td>
                  <td className="p-6">
                    <p className="font-bold text-slate-700">{item.name}</p>
                  </td>
                  <td className="p-6 text-right">
                    <p className="text-xl font-black text-slate-900">{item.tpm}</p>
                  </td>
                  <td className="p-6 text-right">
                    <p className="text-sm font-bold text-slate-500">{item.volume.toLocaleString('it-IT')}</p>
                  </td>
                </tr>
              ))}
              {ranking.length === 0 && (
                <tr>
                  <td colSpan={5} className="p-20 text-center">
                    <div className="flex flex-col items-center gap-4 opacity-20">
                      <BarChart3 size={64} />
                      <p className="font-black uppercase tracking-widest">Nessun dato per questo periodo</p>
                    </div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

const AssociazioneView: React.FC<any> = ({ data, setData, portfolios, user, exportFn }) => {
  const [isAdding, setIsAdding] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  
  const getFirstOfMonth = () => {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-01`;
  };

  const [form, setForm] = useState<Partial<Associazione>>({
    startDate: getFirstOfMonth(),
    quota: 0,
    status: 'nuovo',
    isRenewal: false
  });

  const isAdminOrBO = user.role === 'admin' || user.role === 'backoffice';

  useEffect(() => {
    if (form.startDate) {
      setForm(prev => ({ ...prev, endDate: calculateEndDate(form.startDate!, 12) }));
    }
  }, [form.startDate]);

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    if (!form.portfolioId || !form.quota) return;
    
    const newAssoc: Associazione = {
      ...form,
      id: Math.random().toString(36).substr(2, 9),
      assignedTo: user.id,
      status: 'nuovo', // Always nuovo on creation
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    } as Associazione;

    setData((prev: Associazione[]) => [...prev, newAssoc]);
    setIsAdding(false);
    setForm({
      startDate: getFirstOfMonth(),
      quota: 0,
      status: 'nuovo',
      isRenewal: false
    });
  };

  const updateStatus = (id: string, newStatus: AssociazioneStatus) => {
    if (!isAdminOrBO) return;
    setData((prev: Associazione[]) => prev.map(a => a.id === id ? { ...a, status: newStatus, updatedAt: new Date().toISOString() } : a));
  };

  const deleteAssoc = (id: string) => {
    if (confirm('Sei sicuro di voler eliminare questa associazione?')) {
      setData((prev: Associazione[]) => prev.filter(a => a.id !== id));
    }
  };

  const filteredAssoc = useMemo(() => {
    return data.filter((assoc: Associazione) => {
      const portfolio = portfolios.find((p: Portfolio) => p.id === assoc.portfolioId);
      const searchStr = `${portfolio?.businessName || ''} ${portfolio?.taxId || ''}`.toLowerCase();
      return searchStr.includes(searchTerm.toLowerCase());
    });
  }, [data, portfolios, searchTerm]);

  return (
    <div className="space-y-6 animate-in fade-in duration-500">
      <header className="flex justify-between items-center">
        <div>
          <h2 className="text-3xl font-black tracking-tighter text-emerald-900">Progetto ANTICRISI</h2>
          <p className="text-slate-500 font-medium italic">Gestione quote associative</p>
        </div>
        <div className="flex gap-3 items-center">
          <div className="relative w-64">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
            <input 
              className="w-full pl-10 pr-4 py-2.5 rounded-2xl border-2 border-slate-100 outline-none focus:border-emerald-500 font-bold text-sm" 
              placeholder="Cerca cliente..." 
              value={searchTerm}
              onChange={e => setSearchTerm(e.target.value)}
            />
          </div>
          <button onClick={() => exportFn(data, 'Associazioni')} className="bg-white text-slate-600 px-6 py-3 rounded-2xl font-black flex items-center gap-2 border border-slate-100 shadow-sm hover:bg-slate-50 transition-all text-sm">
            <Download size={18} /> Esporta
          </button>
          <button onClick={() => setIsAdding(true)} className="bg-emerald-700 text-white px-6 py-3 rounded-2xl font-black flex items-center gap-2 shadow-lg shadow-emerald-900/20 hover:scale-105 transition-all text-sm">
            <Plus size={18} /> Nuova Associazione
          </button>
        </div>
      </header>

      {isAdding && (
        <div className="fixed inset-0 z-[200] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-200">
          <form onSubmit={handleSave} className="bg-white w-full max-w-lg rounded-[2.5rem] shadow-2xl overflow-hidden animate-in zoom-in-95 duration-200">
            <div className="p-8 bg-emerald-900 text-white relative">
              <button type="button" onClick={() => setIsAdding(false)} className="absolute top-6 right-6 p-2 hover:bg-white/10 rounded-full transition-colors">
                <X size={24} />
              </button>
              <h3 className="text-2xl font-black tracking-tight">Nuova Associazione</h3>
              <p className="text-emerald-400 text-xs font-bold uppercase tracking-widest mt-1">Inserimento quota associativa</p>
            </div>
            <div className="p-8 space-y-6">
              <SearchableSelect 
                label="Anagrafica Portafoglio *" 
                placeholder="Scegli cliente..." 
                options={portfolios} 
                selectedId={form.portfolioId} 
                onSelect={(id:string) => setForm({...form, portfolioId: id})} 
                displayFn={(p:any)=>p.businessName} 
                searchFn={(p:any)=>`${p.businessName}`} 
                subTextFn={(p:any)=>p.taxId} 
                icon={<Users size={14}/>} 
              />

              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1">Data Inizio</label>
                  <input 
                    type="date"
                    required
                    disabled={user.role !== 'admin'}
                    className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none focus:border-emerald-500 font-bold disabled:opacity-50"
                    value={form.startDate || ''}
                    onChange={e => setForm({...form, startDate: e.target.value})}
                  />
                </div>
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1">Quota (€)</label>
                  <input 
                    type="number"
                    required
                    placeholder="0.00"
                    className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none focus:border-emerald-500 font-bold"
                    value={form.quota || ''}
                    onChange={e => setForm({...form, quota: Number(e.target.value)})}
                  />
                </div>
              </div>

              <div className="flex items-center gap-3 p-4 bg-slate-50 rounded-2xl border border-slate-100">
                <input 
                  type="checkbox" 
                  id="isRenewal" 
                  checked={form.isRenewal} 
                  onChange={e => setForm({...form, isRenewal: e.target.checked})} 
                  className="w-5 h-5 rounded border-slate-300 text-emerald-600 focus:ring-emerald-500" 
                />
                <label htmlFor="isRenewal" className="text-sm font-bold text-slate-700 cursor-pointer">Rinnovo (Secondo Anno)</label>
              </div>

              <div className="p-4 bg-emerald-50 rounded-2xl border border-emerald-100">
                <p className="text-[10px] font-black uppercase text-emerald-600 tracking-widest mb-1">Scadenza Automatica</p>
                <p className="text-lg font-black text-emerald-900">{form.endDate ? formatDate(form.endDate) : '--/--/----'}</p>
              </div>
            </div>
            <div className="p-6 bg-slate-50 border-t border-slate-100 flex justify-end gap-3">
              <button type="button" onClick={() => setIsAdding(false)} className="px-6 py-3 font-bold text-slate-400 hover:text-slate-600 transition-colors">Annulla</button>
              <button type="submit" className="bg-emerald-900 text-white px-10 py-3 rounded-xl font-black shadow-lg shadow-emerald-900/20">Salva Associazione</button>
            </div>
          </form>
        </div>
      )}

      <div className="bg-white rounded-[2.5rem] border border-slate-100 shadow-xl overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse min-w-[800px]">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-100">
                <th className="p-4 text-[10px] font-black uppercase text-slate-400 tracking-widest">Anagrafica</th>
                <th className="p-4 text-[10px] font-black uppercase text-slate-400 tracking-widest">Data Inizio</th>
                <th className="p-4 text-[10px] font-black uppercase text-slate-400 tracking-widest">Quota</th>
                <th className="p-4 text-[10px] font-black uppercase text-slate-400 tracking-widest">Tipo</th>
                <th className="p-4 text-[10px] font-black uppercase text-slate-400 tracking-widest">Stato</th>
                <th className="p-4 text-[10px] font-black uppercase text-slate-400 tracking-widest text-right">Azioni</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-50">
              {filteredAssoc.map((assoc: Associazione) => {
                const portfolio = portfolios.find((p: Portfolio) => p.id === assoc.portfolioId);
                const isExpired = new Date(assoc.endDate) < new Date();
                return (
                  <tr key={assoc.id} className="hover:bg-slate-50/50 transition-colors group">
                    <td className="p-4">
                      <p className="font-bold text-slate-700 text-sm">{portfolio?.businessName || 'N/A'}</p>
                      <p className="text-[9px] font-black text-slate-400 uppercase">{portfolio?.taxId}</p>
                    </td>
                    <td className="p-4">
                      <p className="text-xs font-bold text-slate-600">{formatDate(assoc.startDate)}</p>
                    </td>
                    <td className="p-4">
                      <p className="text-xs font-black text-emerald-700">€ {assoc.quota.toLocaleString('it-IT', { minimumFractionDigits: 2 })}</p>
                    </td>
                    <td className="p-4">
                      <span className={`px-2 py-1 rounded-lg text-[8px] font-black uppercase tracking-widest ${assoc.isRenewal ? 'bg-blue-100 text-blue-600' : 'bg-emerald-100 text-emerald-600'}`}>
                        {assoc.isRenewal ? 'Rinnovo' : 'Nuova'}
                      </span>
                    </td>
                    <td className="p-4">
                      {isAdminOrBO ? (
                        <select 
                          className="text-[9px] font-black uppercase bg-slate-100 p-1.5 rounded-lg outline-none focus:ring-2 focus:ring-emerald-500"
                          value={assoc.status}
                          onChange={e => updateStatus(assoc.id, e.target.value as AssociazioneStatus)}
                        >
                          <option value="nuovo">Nuovo</option>
                          <option value="incassato">Incassato</option>
                          <option value="registrato">Registrato</option>
                          <option value="pagato">Pagato</option>
                        </select>
                      ) : (
                        <span className={`px-3 py-1 rounded-full text-[9px] font-black uppercase tracking-widest ${assoc.status === 'nuovo' ? 'bg-slate-100 text-slate-500' : 'bg-emerald-100 text-emerald-600'}`}>
                          {assoc.status}
                        </span>
                      )}
                      {isExpired && <p className="text-[8px] text-red-500 font-bold uppercase mt-1">Scaduta</p>}
                    </td>
                    <td className="p-4 text-right">
                      <button onClick={() => deleteAssoc(assoc.id)} className="p-1.5 text-slate-300 hover:text-red-600 transition-colors">
                        <Trash2 size={16} />
                      </button>
                    </td>
                  </tr>
                );
              })}
              {filteredAssoc.length === 0 && (
                <tr>
                  <td colSpan={6} className="p-20 text-center">
                    <div className="flex flex-col items-center gap-4 opacity-20">
                      <Wallet size={64} />
                      <p className="font-black uppercase tracking-widest">Nessuna associazione trovata</p>
                    </div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

const RinnoviView: React.FC<any> = ({ contracts, portfolios, pdps, exportFn }) => {
  const [searchTerm, setSearchTerm] = useState('');

  const expiringContracts = useMemo(() => {
    return contracts.filter((c: Contract) => {
      if (c.status !== 'attivo' && c.status !== 'trasmesso') return false;

      // Is there a newer contract for the same PDP?
      const hasNewer = contracts.some(other => 
        other.pdpId === c.pdpId && 
        other.id !== c.id &&
        (other.status === 'attivo' || other.status === 'trasmesso') &&
        new Date(other.startDate) > new Date(c.startDate)
      );

      if (hasNewer) return false;

      // Search filter
      const port = portfolios.find(p => p.id === c.portfolioId);
      const pdp = pdps.find(p => p.id === c.pdpId);
      const searchStr = `${port?.businessName || ''} ${pdp?.pdpCode || ''} ${pdp?.address || ''}`.toLowerCase();
      return searchStr.includes(searchTerm.toLowerCase());
    }).sort((a, b) => new Date(a.endDate).getTime() - new Date(b.endDate).getTime());
  }, [contracts, portfolios, pdps, searchTerm]);

  return (
    <div className="space-y-6 animate-in fade-in duration-500">
      <header className="flex justify-between items-center">
        <div>
          <h2 className="text-3xl font-black tracking-tighter text-emerald-900">Rinnovi</h2>
          <p className="text-slate-500 font-medium italic">Contratti in scadenza senza rinnovo attivo</p>
        </div>
        <div className="flex items-center gap-3">
          <div className="relative w-64">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
            <input 
              className="w-full pl-10 pr-4 py-3 rounded-2xl border-2 border-slate-100 outline-none focus:border-emerald-500 font-bold text-sm" 
              placeholder="Cerca cliente, POD..." 
              value={searchTerm}
              onChange={e => setSearchTerm(e.target.value)}
            />
          </div>
          <button onClick={() => exportFn(expiringContracts, 'Rinnovi')} className="bg-white text-slate-600 px-6 py-3 rounded-2xl font-black flex items-center gap-2 border border-slate-100 shadow-sm hover:bg-slate-50 transition-all">
            <Download size={20} /> Esporta
          </button>
        </div>
      </header>

      <div className="bg-white rounded-[2.5rem] border border-slate-100 shadow-xl overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse min-w-[800px]">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-100">
                <th className="p-6 text-[10px] font-black uppercase text-slate-400 tracking-widest">Cliente</th>
                <th className="p-6 text-[10px] font-black uppercase text-slate-400 tracking-widest">POD / Indirizzo</th>
                <th className="p-6 text-[10px] font-black uppercase text-slate-400 tracking-widest">Servizio</th>
                <th className="p-6 text-[10px] font-black uppercase text-slate-400 tracking-widest">Scadenza</th>
                <th className="p-6 text-[10px] font-black uppercase text-slate-400 tracking-widest">Stato Attuale</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-50">
              {expiringContracts.map((c: Contract) => {
                const portfolio = portfolios.find(p => p.id === c.portfolioId);
                const pdp = pdps.find(p => p.id === c.pdpId);
                const isExpired = new Date(c.endDate) < new Date();
                const daysToExpiry = Math.ceil((new Date(c.endDate).getTime() - new Date().getTime()) / (1000 * 60 * 60 * 24));

                return (
                  <tr key={c.id} className="hover:bg-slate-50/50 transition-colors group">
                    <td className="p-6">
                      <p className="font-bold text-slate-700">{portfolio?.businessName || 'N/A'}</p>
                      <p className="text-[10px] font-black text-slate-400 uppercase">{portfolio?.taxId}</p>
                    </td>
                    <td className="p-6">
                      <p className="font-bold text-slate-600">{pdp?.pdpCode || 'N/A'}</p>
                      <p className="text-[10px] font-medium text-slate-400">{pdp?.address}</p>
                    </td>
                    <td className="p-6">
                      <div className="flex items-center gap-2">
                        {c.service === 'POWER' ? <Zap size={14} className="text-amber-500" /> : <Flame size={14} className="text-orange-500" />}
                        <span className="text-xs font-black uppercase tracking-widest">{c.service}</span>
                      </div>
                    </td>
                    <td className="p-6">
                      <p className={`font-black ${isExpired ? 'text-red-600' : daysToExpiry < 30 ? 'text-orange-500' : 'text-slate-700'}`}>
                        {formatDate(c.endDate)}
                      </p>
                      <p className="text-[10px] font-bold uppercase text-slate-400">
                        {isExpired ? 'Scaduto' : `Tra ${daysToExpiry} giorni`}
                      </p>
                    </td>
                    <td className="p-6">
                      <span className={`px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-widest ${c.status === 'attivo' ? 'bg-emerald-100 text-emerald-600' : 'bg-blue-100 text-blue-600'}`}>
                        {c.status}
                      </span>
                    </td>
                  </tr>
                );
              })}
              {expiringContracts.length === 0 && (
                <tr>
                  <td colSpan={5} className="p-20 text-center">
                    <div className="flex flex-col items-center gap-4 opacity-20">
                      <Activity size={64} />
                      <p className="font-black uppercase tracking-widest">Nessun contratto in scadenza senza rinnovo</p>
                    </div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

const renderContent = () => {
    switch (view) {
      case 'dashboard': return <Dashboard stats={filteredData} setView={setView} user={currentUser} news={news} onOpenNews={setSelectedNews} />;
      case 'portfolio': return <PortfolioView data={filteredData.portfolios} users={users} setData={setPortfolios} user={currentUser} exportFn={exportToExcel} setCreditChecks={setCreditChecks} onCreateContract={(id: string) => { setPrefilledPortfolioId(id); setView('contracts'); }} />;
      case 'pdp': return <PDPView data={filteredData.pdps} setData={setPdps} cabine={cabine} setCabine={setCabine} contracts={contracts} portfolios={portfolios} users={users} user={currentUser} exportFn={exportToExcel} />;
      case 'contracts': return <ContractView data={filteredData.contracts} setData={setContracts} portfolios={portfolios} setPortfolios={setPortfolios} pdps={pdps} cabine={cabine} setCabine={setCabine} setPdps={setPdps} users={users} user={currentUser} exportFn={exportToExcel} onOpenChat={(id: string) => { setActiveChat({ id, type: 'contract' }); markChatAsRead(id); }} getUnreadCount={getUnreadCount} prefilledPortfolioId={prefilledPortfolioId} onClearPrefill={() => setPrefilledPortfolioId(null)} />;
      case 'casi': return <CasiView data={filteredData.casi} setData={setCasi} portfolios={portfolios} pdps={pdps} user={currentUser} exportFn={exportToExcel} onOpenChat={(id: string) => { setActiveChat({ id, type: 'case' }); markChatAsRead(id); }} getUnreadCount={getUnreadCount} />;
      case 'associazioni': return <AssociazioneView data={filteredData.associazioni} setData={setAssociazioni} portfolios={portfolios} user={currentUser} exportFn={exportToExcel} />;
      case 'credit-check': return <CreditCheckView data={filteredData.creditChecks} setData={setCreditChecks} portfolios={portfolios} users={users} user={currentUser} exportFn={exportToExcel} />;
      case 'rinnovi': return <RinnoviView contracts={filteredData.contracts} portfolios={portfolios} pdps={pdps} exportFn={exportToExcel} />;
      case 'map': return <MapView pdps={filteredData.pdps} portfolios={portfolios} contracts={contracts} />;
      case 'statistics': return <StatisticsView contracts={filteredData.contracts} />;
      case 'bulk-upload': return <BulkUploadView setPortfolios={setPortfolios} setPdps={setPdps} setContracts={setContracts} setCabine={setCabine} setCasi={setCasi} setCreditChecks={setCreditChecks} />;
      case 'cabine': return <CabineView data={cabine} pdps={pdps} contracts={contracts} exportFn={exportToExcel} />;
      case 'users': return <UsersView users={users} setUsers={setUsers} />;
      case 'news-admin': return <NewsAdminView news={news} setNews={setNews} />;
      case 'tpm-ranking': return <TPMRankingView stats={filteredData} users={users} />;
      case 'profile': return <ProfileView user={currentUser} onUpdate={u => { setUsers(prev => prev.map(old => old.id === u.id ? u : old)); setCurrentUser(u); localStorage.setItem('charlie_current_user', JSON.stringify(u)); }} />;
      default: return <Dashboard stats={filteredData} setView={setView} user={currentUser} news={news} onOpenNews={setSelectedNews} />;
    }
  };

  return (
    <div className="flex flex-col md:flex-row min-h-screen bg-slate-50 text-slate-900">
      {/* Sidebar - Desktop */}
      <aside className="hidden md:flex w-64 bg-slate-900 text-white fixed h-full flex-col z-50">
        <div className="p-6 flex justify-between items-center">
          <div>
            <span className="text-[9px] uppercase tracking-widest text-emerald-400 font-black">PROGETTO ANTICRISI</span>
            <h1 className="text-4xl font-black tracking-tighter flex items-center gap-2 text-emerald-600">
              1335
            </h1>
          </div>
          <div className="relative">
            <button 
              onClick={() => setShowNotifications(!showNotifications)}
              className={`p-2 rounded-xl transition-all relative ${notificationGroups.length > 0 ? 'bg-red-500 text-white animate-pulse' : 'text-slate-400 hover:bg-slate-800'}`}
            >
              <Bell size={20} />
              {notificationGroups.length > 0 && (
                <span className="absolute -top-1 -right-1 bg-red-600 text-white text-[9px] font-black w-4 h-4 rounded-full flex items-center justify-center border border-slate-900">
                  {notificationGroups.length}
                </span>
              )}
            </button>
            {showNotifications && (
              <div className="absolute left-0 mt-4 w-64 bg-white border border-slate-100 rounded-2xl shadow-2xl z-[100] overflow-hidden animate-in fade-in slide-in-from-top-2 duration-200">
                <div className="p-4 bg-slate-900 text-white flex justify-between items-center">
                  <h4 className="font-black text-[10px] uppercase tracking-widest">Notifiche</h4>
                  <span className="text-[9px] font-bold bg-white/20 px-2 py-0.5 rounded-md">{notificationGroups.length}</span>
                </div>
                <div className="max-h-[300px] overflow-y-auto custom-scrollbar">
                  {notificationGroups.length === 0 ? (
                    <div className="p-8 text-center opacity-30">
                      <BellOff size={24} className="mx-auto mb-2" />
                      <p className="text-[10px] font-bold uppercase">Nessuna notifica</p>
                    </div>
                  ) : (
                    notificationGroups.map(group => (
                      <div 
                        key={group.id} 
                        onClick={() => {
                          setActiveChat({ id: group.id, type: group.type });
                          setShowNotifications(false);
                          markChatAsRead(group.id);
                        }}
                        className="p-3 border-b border-slate-50 hover:bg-slate-50 cursor-pointer transition-colors group"
                      >
                        <div className="flex justify-between items-start mb-1">
                          <span className={`text-[8px] font-black uppercase px-1.5 py-0.5 rounded ${group.type === 'contract' ? 'bg-emerald-100 text-emerald-700' : 'bg-blue-100 text-blue-700'}`}>
                            {group.type === 'contract' ? 'Contratto' : 'Caso'}
                          </span>
                        </div>
                        <p className="text-[10px] font-black text-slate-800 group-hover:text-emerald-600 truncate">ID: {group.id}</p>
                        <p className="text-[9px] text-slate-500 line-clamp-1 mt-0.5 italic">"{group.lastMessage.text}"</p>
                      </div>
                    ))
                  )}
                </div>
              </div>
            )}
          </div>
        </div>
        <nav className="flex-1 px-4 space-y-1 overflow-y-auto custom-scrollbar">
          <NavItem active={view === 'dashboard'} onClick={() => setView('dashboard')} icon={<TrendingUp size={18}/>} label="Dashboard" />
          <NavItem active={view === 'portfolio'} onClick={() => setView('portfolio')} icon={<Users size={18}/>} label="Portafoglio" />
          <NavItem active={view === 'pdp'} onClick={() => setView('pdp')} icon={<Zap size={18}/>} label="PDP" />
          <NavItem active={view === 'contracts'} onClick={() => setView('contracts')} icon={<FileText size={18}/>} label="Contratti" />
          <NavItem active={view === 'associazioni'} onClick={() => setView('associazioni')} icon={<Wallet size={18}/>} label="Associazioni" />
          <NavItem active={view === 'credit-check'} onClick={() => setView('credit-check')} icon={<ShieldAlert size={18}/>} label="Credit Check" />
          <NavItem active={view === 'casi'} onClick={() => setView('casi')} icon={<AlertCircle size={18}/>} label="Gestione Casi" />
          <NavItem active={view === 'map'} onClick={() => setView('map')} icon={<MapPin size={18}/>} label="Mappa" />
          
          <div className="pt-4 pb-2 px-4 text-[10px] font-black text-slate-500 uppercase tracking-widest">Report</div>
          <button 
            onClick={() => setShowReportSubmenu(!showReportSubmenu)}
            className={`w-full flex items-center justify-between px-4 py-3 rounded-xl transition-all font-bold text-sm outline-none ${['statistics', 'cabine', 'rinnovi'].includes(view) ? 'text-emerald-400' : 'text-slate-400 hover:text-emerald-400 hover:bg-emerald-400/5'}`}
          >
            <div className="flex items-center gap-3">
              <BarChart3 size={18}/>
              <span>Report</span>
            </div>
            <ChevronDown size={16} className={`transition-transform ${showReportSubmenu ? 'rotate-180' : ''}`} />
          </button>
          {showReportSubmenu && (
            <div className="space-y-1 animate-in slide-in-from-top-2 duration-200">
              <SubNavItem active={view === 'rinnovi'} onClick={() => setView('rinnovi')} label="Rinnovi" />
              <SubNavItem active={view === 'statistics'} onClick={() => setView('statistics')} label="Statistiche" />
              <SubNavItem active={view === 'cabine'} onClick={() => setView('cabine')} label="Cabine Primarie" />
            </div>
          )}

          <NavItem active={view === 'profile'} onClick={() => setView('profile')} icon={<UserIcon size={18}/>} label="Il Mio Profilo" />
          <div className="pt-4 pb-2 px-4 text-[10px] font-black text-slate-500 uppercase tracking-widest">Risorse Esterne</div>
          <a href={MODULISTICA_URL} target="_blank" rel="noreferrer" className="flex items-center gap-3 px-4 py-3 text-slate-400 hover:text-emerald-400 hover:bg-emerald-400/5 rounded-xl transition-all font-bold text-sm outline-none">
            <ExternalLink size={18} /> Modulistica
          </a>
          <a href={OFFERTE_URL} target="_blank" rel="noreferrer" className="flex items-center gap-3 px-4 py-3 text-slate-400 hover:text-emerald-400 hover:bg-emerald-400/5 rounded-xl transition-all font-bold text-sm outline-none">
            <ExternalLink size={18} /> Offerte
          </a>
          {currentUser.role === 'admin' && (
            <>
              <div className="pt-4 pb-2 px-4 text-[10px] font-black text-slate-500 uppercase tracking-widest">Admin</div>
              <NavItem active={view === 'news-admin'} onClick={() => setView('news-admin')} icon={<Newspaper size={18}/>} label="Gestione Notizie" />
              <NavItem active={view === 'tpm-ranking'} onClick={() => setView('tpm-ranking')} icon={<BarChart3 size={18}/>} label="Classifica TPM" />
              <NavItem active={view === 'bulk-upload'} onClick={() => setView('bulk-upload')} icon={<Upload size={18}/>} label="Caricamento Massivo" />
              <NavItem active={view === 'users'} onClick={() => setView('users')} icon={<UserIcon size={18}/>} label="Utenti" />
            </>
          )}
        </nav>
        <div className="p-4 border-t border-slate-800 space-y-2">
          <div className="flex items-center gap-3 px-3 py-2 bg-slate-800/50 rounded-xl">
            <div className="w-8 h-8 rounded-full bg-emerald-600 flex items-center justify-center font-bold text-xs uppercase">{currentUser.name.charAt(0)}</div>
            <div className="flex-1 overflow-hidden">
              <p className="text-[10px] text-slate-500 font-bold uppercase truncate">{currentUser.role}</p>
              <p className="text-xs font-bold truncate">{currentUser.name}</p>
              <p className="text-[9px] font-black text-emerald-400">{currentUser.userCode}</p>
            </div>
          </div>
          <button onClick={handleLogout} className="w-full flex items-center gap-3 px-4 py-2 text-slate-400 hover:text-red-400 hover:bg-red-400/10 rounded-xl transition-all font-bold text-sm outline-none">
            <LogOut size={18} /> Esci
          </button>
        </div>
      </aside>

      {/* Mobile Header */}
      <div className="md:hidden bg-slate-900 text-white p-4 flex items-center justify-between sticky top-0 z-50">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 bg-emerald-600 rounded-lg flex items-center justify-center text-white">
            <Zap size={18} />
          </div>
          <h1 className="font-black text-xl tracking-tighter text-emerald-600">1335</h1>
        </div>
        <div className="flex items-center gap-2">
          <div className="relative mr-2">
            <button 
              onClick={() => setShowNotifications(!showNotifications)}
              className={`p-2 rounded-xl transition-all relative ${notificationGroups.length > 0 ? 'bg-red-500 text-white' : 'text-slate-400'}`}
            >
              <Bell size={20} />
              {notificationGroups.length > 0 && (
                <span className="absolute -top-1 -right-1 bg-red-600 text-white text-[9px] font-black w-4 h-4 rounded-full flex items-center justify-center border border-slate-900">
                  {notificationGroups.length}
                </span>
              )}
            </button>
            {showNotifications && (
              <div className="absolute right-0 mt-4 w-64 bg-white border border-slate-100 rounded-2xl shadow-2xl z-[100] overflow-hidden animate-in fade-in slide-in-from-top-2 duration-200">
                <div className="p-4 bg-slate-900 text-white flex justify-between items-center">
                  <h4 className="font-black text-[10px] uppercase tracking-widest">Notifiche</h4>
                  <span className="text-[9px] font-bold bg-white/20 px-2 py-0.5 rounded-md">{notificationGroups.length}</span>
                </div>
                <div className="max-h-[300px] overflow-y-auto custom-scrollbar">
                  {notificationGroups.length === 0 ? (
                    <div className="p-8 text-center opacity-30">
                      <BellOff size={24} className="mx-auto mb-2" />
                      <p className="text-[10px] font-bold uppercase">Nessuna notifica</p>
                    </div>
                  ) : (
                    notificationGroups.map(group => (
                      <div 
                        key={group.id} 
                        onClick={() => {
                          setActiveChat({ id: group.id, type: group.type });
                          setShowNotifications(false);
                          markChatAsRead(group.id);
                        }}
                        className="p-3 border-b border-slate-50 hover:bg-slate-50 cursor-pointer transition-colors group text-slate-900"
                      >
                        <div className="flex justify-between items-start mb-1">
                          <span className={`text-[8px] font-black uppercase px-1.5 py-0.5 rounded ${group.type === 'contract' ? 'bg-emerald-100 text-emerald-700' : 'bg-blue-100 text-blue-700'}`}>
                            {group.type === 'contract' ? 'Contratto' : 'Caso'}
                          </span>
                        </div>
                        <p className="text-[10px] font-black group-hover:text-emerald-600 truncate">ID: {group.id}</p>
                        <p className="text-[9px] text-slate-500 line-clamp-1 mt-0.5 italic">"{group.lastMessage.text}"</p>
                      </div>
                    ))
                  )}
                </div>
              </div>
            )}
          </div>
          <div className="flex gap-1 mr-2">
            <a href={MODULISTICA_URL} target="_blank" rel="noreferrer" className="p-2 text-slate-400 hover:text-emerald-400" title="Modulistica"><ExternalLink size={18}/></a>
            <a href={OFFERTE_URL} target="_blank" rel="noreferrer" className="p-2 text-slate-400 hover:text-emerald-400" title="Offerte"><ExternalLink size={18}/></a>
          </div>
          <button onClick={handleLogout} className="p-2 text-slate-400 hover:text-red-400">
            <LogOut size={20} />
          </button>
        </div>
      </div>

      {/* Main Content */}
      <main className="flex-1 md:ml-64 p-4 md:p-8 overflow-x-hidden pb-24 md:pb-8">
        <div className="max-w-7xl mx-auto">
          {renderContent()}
        </div>
      </main>

      {/* Mobile Bottom Navigation */}
      <nav className="md:hidden fixed bottom-0 left-0 right-0 bg-white border-t border-slate-200 flex items-center justify-around p-2 z-50 shadow-2xl">
        <MobileNavItem active={view === 'dashboard'} onClick={() => setView('dashboard')} icon={<TrendingUp size={20}/>} />
        <MobileNavItem active={view === 'portfolio'} onClick={() => setView('portfolio')} icon={<Users size={20}/>} />
        <MobileNavItem active={view === 'pdp'} onClick={() => setView('pdp')} icon={<Zap size={20}/>} />
        <MobileNavItem active={view === 'contracts'} onClick={() => setView('contracts')} icon={<FileText size={20}/>} />
        <MobileNavItem active={view === 'credit-check'} onClick={() => setView('credit-check')} icon={<ShieldAlert size={20}/>} />
        <button 
          onClick={() => {
            const moreViews: AppView[] = ['cabine', 'casi', 'map', 'statistics', 'bulk-upload', 'users', 'rinnovi', 'associazioni'];
            const currentIndex = moreViews.indexOf(view);
            const nextIndex = (currentIndex + 1) % moreViews.length;
            setView(moreViews[nextIndex]);
          }}
          className={`p-3 rounded-2xl transition-all duration-200 ${['cabine', 'casi', 'map', 'statistics', 'bulk-upload', 'users', 'rinnovi', 'associazioni'].includes(view) ? 'bg-emerald-50 text-emerald-600' : 'text-slate-400'}`}
        >
          <Menu size={20} />
        </button>
      </nav>
      {activeChat && currentUser && (
        <ChatModal 
          isOpen={!!activeChat} 
          onClose={() => setActiveChat(null)} 
          targetId={activeChat.id} 
          targetType={activeChat.type} 
          messages={messages} 
          onSendMessage={sendMessage} 
          currentUser={currentUser} 
          users={users}
        />
      )}
      {selectedNews && (
        <NewsDetailModal news={selectedNews} onClose={() => setSelectedNews(null)} />
      )}
    </div>
  );
};

const MobileNavItem = ({ active, onClick, icon }: { active: boolean, onClick: () => void, icon: React.ReactNode }) => (
  <button 
    onClick={onClick}
    className={`p-3 rounded-2xl transition-all duration-200 ${active ? 'bg-emerald-50 text-emerald-600' : 'text-slate-400'}`}
  >
    {icon}
  </button>
);

// --- COMPONENTI UTILITY ---

const NavItem: React.FC<{ active: boolean, onClick: () => void, icon: React.ReactNode, label: string }> = ({ active, onClick, icon, label }) => (
  <button onClick={onClick} className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all ${active ? 'bg-emerald-700 text-white shadow-lg' : 'text-slate-400 hover:bg-slate-800/50 hover:text-white'}`}>
    {icon} <span className="text-sm font-semibold">{label}</span>
  </button>
);

const SubNavItem: React.FC<{ active: boolean, onClick: () => void, label: string }> = ({ active, onClick, label }) => (
  <button 
    onClick={onClick} 
    className={`w-full flex items-center gap-3 pl-12 pr-4 py-2 rounded-xl transition-all font-bold text-xs outline-none ${active ? 'text-emerald-400' : 'text-slate-500 hover:text-emerald-400'}`}
  >
    <div className={`w-1.5 h-1.5 rounded-full ${active ? 'bg-emerald-400' : 'bg-slate-700'}`} />
    <span>{label}</span>
  </button>
);

const SearchableSelect: React.FC<any> = ({ 
  options, onSelect, label, icon, displayFn, searchFn, subTextFn, selectedId, placeholder, allowCustom, customLabel, disabled, maxLength 
}) => {
  const [isOpen, setIsOpen] = useState(false);
  const [query, setQuery] = useState('');
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const close = (e: MouseEvent) => { if (ref.current && !ref.current.contains(e.target as Node)) setIsOpen(false); };
    document.addEventListener('mousedown', close);
    return () => document.removeEventListener('mousedown', close);
  }, []);

  const filtered = options.filter((o: any) => searchFn(o).toLowerCase().includes(query.toLowerCase()));
  const selected = options.find((o: any) => o.id === selectedId);

  return (
    <div className={`space-y-1.5 relative ${disabled ? 'opacity-50 pointer-events-none' : ''}`} ref={ref}>
      <label className="text-xs font-bold text-slate-500 uppercase flex items-center gap-1.5">{icon}{label}</label>
      <div onClick={() => !disabled && setIsOpen(!isOpen)} className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 flex items-center justify-between cursor-pointer focus:border-emerald-500 border-slate-100 min-h-[58px]">
        <span className={selected || (allowCustom && selectedId?.startsWith('NEW_')) ? "font-bold" : "text-slate-400"}>
          {selected ? displayFn(selected) : (selectedId?.startsWith('NEW_') ? selectedId.replace('NEW_', '') : placeholder)}
        </span>
        <ChevronDown size={16} className={isOpen ? 'rotate-180 transition-all' : ''} />
      </div>
      {isOpen && (
        <div className="absolute z-[100] w-full mt-2 bg-white border border-slate-200 rounded-2xl shadow-2xl overflow-hidden animate-in fade-in zoom-in duration-200">
          <div className="p-2 border-b bg-slate-50"><input autoFocus maxLength={maxLength} className="w-full p-2 text-sm outline-none bg-transparent" placeholder="Digita per cercare..." value={query} onChange={e => setQuery(e.target.value)} /></div>
          <div className="max-h-60 overflow-y-auto">
            {allowCustom && query && query.length === 14 && !options.some(o => displayFn(o).toLowerCase() === query.toLowerCase()) && (
              <div onClick={() => { onSelect(`NEW_${query.toUpperCase()}`); setIsOpen(false); setQuery(''); }} className="p-3 hover:bg-emerald-50 cursor-pointer border-b border-slate-50 bg-emerald-50/20">
                <p className="font-bold text-sm text-emerald-700">{customLabel || 'Nuovo record'}: {query.toUpperCase()}</p>
                <p className="text-[10px] text-emerald-600 font-bold uppercase">Crea record mancante</p>
              </div>
            )}
            {allowCustom && query && query.length > 0 && query.length < 14 && (
              <div className="p-3 bg-red-50/20 border-b border-slate-50">
                <p className="text-[10px] text-red-600 font-bold uppercase">Inserire 14 caratteri ({query.length}/14)</p>
              </div>
            )}
            {filtered.map((o: any) => (
              <div key={o.id} onClick={() => { onSelect(o.id); setIsOpen(false); setQuery(''); }} className="p-3 hover:bg-emerald-50 cursor-pointer border-b border-slate-50 last:border-0">
                <p className="font-bold text-sm">{displayFn(o)}</p>
                <p className="text-[10px] text-slate-400 font-bold uppercase">{subTextFn(o)}</p>
              </div>
            ))}
            {filtered.length === 0 && !allowCustom && <p className="p-4 text-center text-xs text-slate-400 font-bold">Nessun risultato</p>}
          </div>
        </div>
      )}
    </div>
  );
};

const FileUploader: React.FC<{ onUpload: (files: Attachment[]) => void, attachments: Attachment[], onRemove: (id: string) => void, canEdit: boolean }> = ({ onUpload, attachments, onRemove, canEdit }) => {
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    const validFiles: Attachment[] = [];
    files.forEach((f: File) => {
      if (f.size > FILE_SIZE_LIMIT) return alert(`Il file ${f.name} supera il limite di 5MB`);
      validFiles.push({
        id: Math.random().toString(36).substr(2, 9),
        fileName: f.name,
        fileSize: f.size,
        uploadDate: formatDate(new Date()),
        mimeType: f.type
      });
    });
    if (validFiles.length > 0) onUpload(validFiles);
  };

  return (
    <div className="space-y-4">
      <div className="p-4 border-2 border-dashed border-slate-200 rounded-2xl flex flex-col items-center justify-center gap-2 hover:border-emerald-400 hover:bg-emerald-50 transition-all group">
        <Paperclip className="text-slate-300 group-hover:text-emerald-500 transition-colors" />
        <p className="text-[10px] font-black uppercase text-slate-400 tracking-widest">PDF, XLS, XLSX (Max 5MB)</p>
        <input type="file" multiple className="hidden" id="file-up" onChange={handleFileChange} accept=".pdf,.xls,.xlsx" />
        <label htmlFor="file-up" className="cursor-pointer text-emerald-600 font-black text-xs hover:underline">Seleziona Allegati</label>
      </div>
      <div className="flex flex-wrap gap-2">
        {attachments.map(at => (
          <div key={at.id} className="bg-slate-100 px-3 py-2 rounded-xl flex items-center gap-2 max-w-[250px] border border-slate-200">
            <Files size={14} className="text-slate-400 shrink-0" />
            <div className="flex-1 min-w-0">
              <p className="text-[11px] font-bold text-slate-700 truncate">{at.fileName}</p>
              <p className="text-[9px] text-slate-400 font-bold uppercase">{at.uploadDate}</p>
            </div>
            {canEdit && <button type="button" onClick={() => onRemove(at.id)} className="text-red-400 hover:text-red-600 ml-auto"><X size={14} /></button>}
          </div>
        ))}
      </div>
    </div>
  );
};

const TimestampDetail: React.FC<{ data: any }> = ({ data }) => (
  <div className="mt-4 p-4 bg-slate-50 rounded-2xl border border-slate-100 flex flex-col gap-2">
    <div className="flex items-center gap-2 text-[10px] font-black text-slate-400 uppercase tracking-widest">
      <Clock size={12} /> Cronologia Record
    </div>
    <div className="flex justify-between">
      <div className="text-[11px] font-medium text-slate-500">Creato il: <span className="font-bold text-slate-700">{formatDateTime(data.createdAt) || 'N/D'}</span></div>
      <div className="text-[11px] font-medium text-slate-500">Ultimo aggiornamento: <span className="font-bold text-slate-700">{formatDateTime(data.updatedAt) || 'N/D'}</span> {data.updatedBy && <span className="text-emerald-600 font-black ml-1">({data.updatedBy})</span>}</div>
    </div>
  </div>
);

const StatCard: React.FC<{ label: string, value: number, icon: React.ReactNode, onClick: () => void }> = ({ label, value, icon, onClick }) => (
  <button onClick={onClick} className="bg-white p-6 rounded-[2rem] border border-slate-100 shadow-sm hover:shadow-xl transition-all text-left flex flex-col gap-4 group">
    <div className="flex justify-between items-start">
      <div className="p-3 bg-slate-50 rounded-2xl group-hover:bg-emerald-50 transition-colors">
        {icon}
      </div>
      <ChevronRight size={16} className="text-slate-300 group-hover:text-emerald-500 transition-colors" />
    </div>
    <div>
      <p className="text-[9px] font-bold uppercase text-slate-400 tracking-widest mb-0.5">{label}</p>
      <p className="text-lg font-black text-slate-900 tracking-tight">{value}</p>
    </div>
  </button>
);

// --- SCHERMATE PRINCIPALI ---

const ChatModal: React.FC<{ 
  isOpen: boolean; 
  onClose: () => void; 
  targetId: string; 
  targetType: 'contract' | 'case';
  messages: Message[];
  onSendMessage: (text: string, attachments?: Attachment[], asEmail?: boolean) => void;
  currentUser: User;
  users: User[];
}> = ({ isOpen, onClose, targetId, targetType, messages, onSendMessage, currentUser, users }) => {
  const [inputText, setInputText] = useState('');
  const [pendingAttachments, setPendingAttachments] = useState<Attachment[]>([]);
  const [mentionQuery, setMentionQuery] = useState('');
  const [showMentions, setShowMentions] = useState(false);
  const [sendAsEmail, setSendAsEmail] = useState(false);
  const scrollRef = useRef<HTMLDivElement>(null);

  const targetMessages = messages.filter(m => m.targetId === targetId);

  useEffect(() => {
    if (scrollRef.current) {
      scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
    }
  }, [targetMessages, isOpen]);

  if (!isOpen) return null;

  const handleSend = () => {
    if (inputText.trim() || pendingAttachments.length > 0) {
      onSendMessage(inputText, pendingAttachments.length > 0 ? pendingAttachments : undefined, sendAsEmail);
      setInputText('');
      setPendingAttachments([]);
      setShowMentions(false);
      setSendAsEmail(false);
    }
  };

  const handleInputChange = (val: string) => {
    setInputText(val);
    const lastAt = val.lastIndexOf('@');
    const isTrigger = lastAt !== -1 && (lastAt === 0 || val[lastAt - 1] === ' ');
    
    if (isTrigger) {
      const query = val.slice(lastAt + 1);
      if (!query.includes(' ')) {
        setMentionQuery(query);
        setShowMentions(true);
      } else {
        setShowMentions(false);
      }
    } else {
      setShowMentions(false);
    }
  };

  const insertMention = (user: User) => {
    const lastAt = inputText.lastIndexOf('@');
    const newVal = inputText.slice(0, lastAt) + '@' + user.userCode + ' ';
    setInputText(newVal);
    setShowMentions(false);
  };

  const filteredUsers = users.filter(u => 
    u.name.toLowerCase().includes(mentionQuery.toLowerCase()) || 
    u.userCode.toLowerCase().includes(mentionQuery.toLowerCase())
  );

  const handleFileUpload = (files: Attachment[]) => {
    setPendingAttachments(prev => [...prev, ...files]);
  };

  return (
    <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[200] flex items-center justify-center p-4 animate-in fade-in duration-300">
      <div className="bg-white w-full max-w-lg rounded-[2.5rem] shadow-2xl overflow-hidden flex flex-col h-[600px] animate-in zoom-in duration-300">
        <div className="p-6 bg-slate-900 text-white flex justify-between items-center">
          <div>
            <h3 className="text-xl font-black tracking-tight flex items-center gap-2">
              <MessageSquare size={20} className="text-emerald-400" /> Chat {targetType === 'contract' ? 'Contratto' : 'Caso'}
            </h3>
            <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mt-1">ID: {targetId}</p>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-white/10 rounded-xl transition-colors"><X size={24} /></button>
        </div>

        <div ref={scrollRef} className="flex-1 overflow-y-auto p-6 space-y-4 bg-slate-50 custom-scrollbar">
          {targetMessages.length === 0 ? (
            <div className="h-full flex flex-col items-center justify-center text-slate-400 gap-2 opacity-50">
              <MessageSquare size={48} />
              <p className="font-bold text-sm">Nessun messaggio. Inizia la conversazione!</p>
            </div>
          ) : (
            targetMessages.map((m) => (
              <div key={m.id} className={`flex flex-col ${m.senderId === currentUser.id ? 'items-end' : 'items-start'}`}>
                <div className={`max-w-[80%] p-4 rounded-2xl shadow-sm ${
                  m.senderId === currentUser.id 
                    ? 'bg-emerald-600 text-white rounded-tr-none' 
                    : 'bg-white text-slate-800 rounded-tl-none border border-slate-100'
                }`}>
                  <div className="flex justify-between items-center gap-4 mb-1">
                    <span className={`text-[9px] font-black uppercase tracking-widest ${m.senderId === currentUser.id ? 'text-emerald-200' : 'text-slate-400'}`}>
                      {m.senderName} ({m.senderRole})
                    </span>
                  </div>
                  {m.text && <p className="text-sm font-medium leading-relaxed">{m.text}</p>}
                  
                  {m.attachments && m.attachments.length > 0 && (
                    <div className="mt-2 space-y-1">
                      {m.attachments.map(at => (
                        <div key={at.id} className={`flex items-center gap-2 p-2 rounded-lg text-[10px] font-bold ${m.senderId === currentUser.id ? 'bg-emerald-700/50 text-emerald-100' : 'bg-slate-100 text-slate-600'}`}>
                          <Paperclip size={12} />
                          <span className="truncate flex-1">{at.fileName}</span>
                          <Download size={12} className="cursor-pointer hover:scale-110 transition-transform" />
                        </div>
                      ))}
                    </div>
                  )}

                  <p className={`text-[9px] mt-2 font-bold ${m.senderId === currentUser.id ? 'text-emerald-200/70' : 'text-slate-400'}`}>
                    {new Date(m.timestamp).toLocaleTimeString('it-IT', { hour: '2-digit', minute: '2-digit' })}
                  </p>
                </div>
              </div>
            ))
          )}
        </div>

          <div className="flex flex-col gap-4">
            {pendingAttachments.length > 0 && (
              <div className="flex flex-wrap gap-2">
                {pendingAttachments.map(at => (
                  <div key={at.id} className="bg-emerald-50 text-emerald-700 px-2 py-1 rounded-lg text-[10px] font-bold flex items-center gap-1 border border-emerald-100">
                    <Paperclip size={10} />
                    <span className="truncate max-w-[100px]">{at.fileName}</span>
                    <button onClick={() => setPendingAttachments(prev => prev.filter(x => x.id !== at.id))}><X size={10} /></button>
                  </div>
                ))}
              </div>
            )}
            
            <div className="flex items-center gap-2 px-2">
              <button 
                onClick={() => setSendAsEmail(!sendAsEmail)}
                className={`flex items-center gap-2 px-3 py-1.5 rounded-full text-[10px] font-black uppercase transition-all border ${sendAsEmail ? 'bg-emerald-600 text-white border-emerald-600 shadow-md' : 'bg-white text-slate-400 border-slate-200 hover:border-slate-300'}`}
              >
                <Mail size={12} />
                Invia anche via Email
              </button>
              {sendAsEmail && (
                <span className="text-[9px] font-bold text-emerald-600 animate-pulse">L'utente riceverà una notifica email</span>
              )}
            </div>

            <div className="flex gap-2">
            <div className="relative">
              <input 
                type="file" 
                multiple 
                className="hidden" 
                id="chat-file-up" 
                onChange={(e) => {
                  const files = Array.from(e.target.files || []);
                  const validFiles: Attachment[] = files.map((f: any) => ({
                    id: Math.random().toString(36).substr(2, 9),
                    fileName: f.name,
                    fileSize: f.size,
                    uploadDate: formatDate(new Date()),
                    mimeType: f.type
                  }));
                  handleFileUpload(validFiles);
                }}
              />
              <label htmlFor="chat-file-up" className="bg-slate-100 text-slate-500 p-4 rounded-2xl hover:bg-slate-200 transition-all cursor-pointer flex items-center justify-center">
                <Paperclip size={20} />
              </label>
            </div>
            <div className="relative flex-1">
              {showMentions && filteredUsers.length > 0 && (
                <div className="absolute bottom-full left-0 w-full bg-white border border-slate-200 rounded-2xl shadow-2xl mb-2 overflow-hidden z-[210] animate-in slide-in-from-bottom-2">
                  <div className="p-3 bg-slate-50 border-b border-slate-100 flex items-center justify-between">
                    <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Menziona Utente</span>
                    <button onClick={() => setShowMentions(false)}><X size={12} className="text-slate-400" /></button>
                  </div>
                  <div className="max-h-48 overflow-y-auto">
                    {filteredUsers.map(u => (
                      <button 
                        key={u.id} 
                        onClick={() => insertMention(u)}
                        className="w-full p-3 flex items-center gap-3 hover:bg-emerald-50 transition-colors text-left border-b border-slate-50 last:border-0"
                      >
                        <div className="w-8 h-8 rounded-full bg-emerald-100 flex items-center justify-center text-emerald-700 font-black text-[10px]">
                          {u.userCode.slice(0, 2)}
                        </div>
                        <div className="flex-1 overflow-hidden">
                          <p className="font-black text-emerald-600 uppercase tracking-widest text-xs">{u.userCode}</p>
                          <p className="text-[10px] font-bold text-slate-400 truncate">{u.name}</p>
                        </div>
                      </button>
                    ))}
                  </div>
                </div>
              )}
              <input 
                className="w-full bg-slate-50 border-2 border-slate-100 rounded-2xl px-4 py-3 outline-none focus:border-emerald-500 font-medium transition-all"
                placeholder="Scrivi un messaggio..."
                value={inputText}
                onChange={e => handleInputChange(e.target.value)}
                onKeyDown={e => e.key === 'Enter' && handleSend()}
              />
            </div>
            <button 
              onClick={handleSend}
              className="bg-emerald-600 text-white p-4 rounded-2xl hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-600/20 active:scale-95"
            >
              <Send size={20} />
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

const NewsBanner: React.FC<{ news: News[]; onOpenDetail: (n: News) => void }> = ({ news, onOpenDetail }) => {
  const activeNews = news.filter(n => n.isActive);
  const [currentIndex, setCurrentIndex] = useState(0);
  const [isPaused, setIsPaused] = useState(false);

  useEffect(() => {
    if (activeNews.length <= 1 || isPaused) return;
    const interval = setInterval(() => {
      setCurrentIndex(prev => (prev + 1) % activeNews.length);
    }, 5000);
    return () => clearInterval(interval);
  }, [activeNews.length, isPaused]);

  if (activeNews.length === 0) return null;

  const current = activeNews[currentIndex];

  const next = () => setCurrentIndex(prev => (prev + 1) % activeNews.length);
  const prev = () => setCurrentIndex(prev => (prev - 1 + activeNews.length) % activeNews.length);

  return (
    <div 
      className="relative bg-emerald-900 text-white p-6 rounded-[2rem] shadow-xl overflow-hidden group cursor-pointer"
      onMouseEnter={() => setIsPaused(true)}
      onMouseLeave={() => setIsPaused(false)}
      onClick={() => onOpenDetail(current)}
    >
      <div className="absolute top-0 right-0 p-4 opacity-10 group-hover:opacity-20 transition-opacity">
        <Bell size={80} />
      </div>
      
      <div className="relative z-10 flex items-center justify-between gap-4">
        <div className="flex-1 min-w-0">
          <p className="text-[10px] font-black uppercase tracking-[0.2em] text-emerald-400 mb-1">{current.occhiello}</p>
          <h3 className="text-xl md:text-2xl font-black tracking-tight truncate">{current.title}</h3>
          <p className="text-[10px] text-emerald-300/70 font-medium mt-2 flex items-center gap-1.5 uppercase">
            <Clock size={10} /> {formatDate(current.createdAt)} • Clicca per leggere
          </p>
        </div>

        {activeNews.length > 1 && (
          <div className="flex gap-2" onClick={e => e.stopPropagation()}>
            <button onClick={prev} className="p-2 rounded-full bg-white/10 hover:bg-white/20 transition-colors">
              <ChevronRight size={20} className="rotate-180" />
            </button>
            <button onClick={next} className="p-2 rounded-full bg-white/10 hover:bg-white/20 transition-colors">
              <ChevronRight size={20} />
            </button>
          </div>
        )}
      </div>

      {activeNews.length > 1 && (
        <div className="absolute bottom-2 left-1/2 -translate-x-1/2 flex gap-1.5">
          {activeNews.map((_, i) => (
            <div 
              key={i} 
              className={`h-1 rounded-full transition-all duration-500 ${i === currentIndex ? 'w-4 bg-emerald-400' : 'w-1 bg-white/20'}`}
            />
          ))}
        </div>
      )}
    </div>
  );
};

const NewsDetailModal: React.FC<{ news: News; onClose: () => void }> = ({ news, onClose }) => {
  return (
    <div className="fixed inset-0 z-[200] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-200">
      <div className="bg-white w-full max-w-2xl rounded-[2.5rem] shadow-2xl overflow-hidden animate-in zoom-in-95 duration-200 flex flex-col max-h-[90vh]">
        <div className="p-8 bg-emerald-900 text-white relative">
          <button onClick={onClose} className="absolute top-6 right-6 p-2 hover:bg-white/10 rounded-full transition-colors">
            <X size={24} />
          </button>
          <p className="text-xs font-black uppercase tracking-widest text-emerald-400 mb-2">{news.occhiello}</p>
          <h2 className="text-3xl font-black tracking-tight leading-tight">{news.title}</h2>
          <div className="flex items-center gap-4 mt-6 text-[10px] font-bold text-emerald-300/70 uppercase">
            <span className="flex items-center gap-1.5"><Clock size={12}/> {formatDate(news.createdAt)}</span>
            <span className="flex items-center gap-1.5"><Activity size={12}/> News Ufficiale</span>
          </div>
        </div>
        
        <div className="p-8 overflow-y-auto flex-1 space-y-8">
          <div className="prose prose-slate max-w-none">
            <p className="text-slate-600 leading-relaxed whitespace-pre-wrap text-lg">{news.content}</p>
          </div>

          {news.attachments.length > 0 && (
            <div className="space-y-4 pt-8 border-t border-slate-100">
              <h4 className="text-xs font-black uppercase text-slate-400 tracking-widest flex items-center gap-2">
                <Paperclip size={14} /> Allegati ({news.attachments.length})
              </h4>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                {news.attachments.map(a => (
                  <div key={a.id} className="flex items-center justify-between p-4 bg-slate-50 rounded-2xl border border-slate-100 group hover:border-emerald-200 transition-all">
                    <div className="flex items-center gap-3 min-w-0">
                      <div className="p-2 bg-white rounded-lg shadow-sm text-emerald-600">
                        <FileText size={18} />
                      </div>
                      <div className="min-w-0">
                        <p className="text-sm font-bold text-slate-700 truncate">{a.fileName}</p>
                        <p className="text-[10px] text-slate-400 font-medium">{(a.fileSize / 1024).toFixed(1)} KB</p>
                      </div>
                    </div>
                    <button className="p-2 text-slate-400 hover:text-emerald-600 transition-colors">
                      <Download size={18} />
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>

        <div className="p-6 bg-slate-50 border-t border-slate-100 flex justify-center">
          <button onClick={onClose} className="px-8 py-3 bg-emerald-900 text-white rounded-xl font-black shadow-lg hover:scale-105 transition-all">
            Ho letto tutto
          </button>
        </div>
      </div>
    </div>
  );
};

const NewsAdminView: React.FC<{ news: News[]; setNews: React.Dispatch<React.SetStateAction<News[]>> }> = ({ news, setNews }) => {
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [form, setForm] = useState<Partial<News>>({ title: '', occhiello: '', content: '', attachments: [], isActive: true });

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    const ts = getTimestamp();
    if (editingId) {
      setNews(prev => prev.map(n => n.id === editingId ? { ...n, ...form, updatedAt: ts } as News : n));
    } else {
      const newNews: News = {
        ...form,
        id: Math.random().toString(36).substr(2, 9),
        createdAt: ts,
        updatedAt: ts,
        attachments: form.attachments || [],
        isActive: form.isActive ?? true,
      } as News;
      setNews(prev => [newNews, ...prev]);
    }
    setIsFormOpen(false);
    setEditingId(null);
    setForm({ title: '', occhiello: '', content: '', attachments: [], isActive: true });
  };

  const deleteNews = (id: string) => {
    if (window.confirm('Sei sicuro di voler eliminare questa notizia?')) {
      setNews(prev => prev.filter(n => n.id !== id));
    }
  };

  return (
    <div className="space-y-8 animate-in fade-in duration-500">
      <div className="flex justify-between items-center">
        <div>
          <h2 className="text-3xl font-black tracking-tighter text-slate-900">Gestione Notizie</h2>
          <p className="text-sm text-slate-500 font-medium">Crea e gestisci i banner informativi per gli utenti</p>
        </div>
        <button 
          onClick={() => { setIsFormOpen(true); setEditingId(null); setForm({ title: '', occhiello: '', content: '', attachments: [], isActive: true }); }}
          className="bg-emerald-600 text-white px-6 py-3 rounded-xl font-black shadow-lg hover:scale-105 transition-all flex items-center gap-2"
        >
          <Plus /> Nuova Notizia
        </button>
      </div>

      {isFormOpen && (
        <form onSubmit={handleSave} className="bg-white p-8 rounded-[2.5rem] border-2 border-emerald-50 shadow-2xl space-y-6 animate-in slide-in-from-top-4">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1">Occhiello (Sottotitolo breve)</label>
              <input required className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold" value={form.occhiello || ''} onChange={e => setForm({...form, occhiello: e.target.value})} placeholder="Es: COMUNICAZIONE UFFICIALE" />
            </div>
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1">Titolo Notizia</label>
              <input required className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold" value={form.title || ''} onChange={e => setForm({...form, title: e.target.value})} placeholder="Titolo accattivante per il banner" />
            </div>
            <div className="col-span-full space-y-1.5">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1">Contenuto Integrale</label>
              <textarea required rows={6} className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-medium" value={form.content || ''} onChange={e => setForm({...form, content: e.target.value})} placeholder="Scrivi qui il testo completo della notizia..." />
            </div>
            <div className="col-span-full">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1">Allegati</label>
              <FileUploader attachments={form.attachments || []} canEdit={true} onUpload={files => setForm({...form, attachments: [...(form.attachments || []), ...files]})} onRemove={id => setForm({...form, attachments: form.attachments?.filter(a => a.id !== id)})} />
            </div>
            <div className="flex items-center gap-3 p-4 bg-slate-50 rounded-2xl border border-slate-100">
              <input type="checkbox" id="isActive" checked={form.isActive} onChange={e => setForm({...form, isActive: e.target.checked})} className="w-5 h-5 rounded border-slate-300 text-emerald-600 focus:ring-emerald-500" />
              <label htmlFor="isActive" className="text-sm font-bold text-slate-700 cursor-pointer">Notizia Attiva (visibile nel banner)</label>
            </div>
          </div>
          <div className="flex justify-end gap-4 border-t pt-6">
            <button type="button" onClick={() => setIsFormOpen(false)} className="font-bold text-slate-400">Annulla</button>
            <button type="submit" className="bg-emerald-900 text-white px-10 py-4 rounded-2xl font-black">
              {editingId ? 'Aggiorna Notizia' : 'Pubblica Notizia'}
            </button>
          </div>
        </form>
      )}

      <div className="grid grid-cols-1 gap-4">
        {news.map(n => (
          <div key={n.id} className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm flex items-center justify-between group hover:border-emerald-200 transition-all">
            <div className="flex items-center gap-4">
              <div className={`p-3 rounded-2xl ${n.isActive ? 'bg-emerald-100 text-emerald-600' : 'bg-slate-100 text-slate-400'}`}>
                {n.isActive ? <Bell size={24} /> : <BellOff size={24} />}
              </div>
              <div>
                <p className="text-[10px] font-black uppercase text-emerald-600 tracking-widest">{n.occhiello}</p>
                <h4 className="text-lg font-black text-slate-900">{n.title}</h4>
                <p className="text-xs text-slate-400 font-medium">Ultima modifica: {formatDateTime(n.updatedAt)} • {n.attachments.length} allegati</p>
              </div>
            </div>
            <div className="flex gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
              <button onClick={() => { setForm(n); setEditingId(n.id); setIsFormOpen(true); }} className="p-2 text-slate-400 hover:text-emerald-600 transition-colors"><Edit2 size={20} /></button>
              <button onClick={() => deleteNews(n.id)} className="p-2 text-slate-400 hover:text-red-600 transition-colors"><Trash2 size={20} /></button>
            </div>
          </div>
        ))}
        {news.length === 0 && (
          <div className="text-center py-20 bg-slate-50 rounded-[3rem] border-2 border-dashed border-slate-200">
            <BellOff size={48} className="mx-auto text-slate-300 mb-4" />
            <p className="text-slate-400 font-bold">Nessuna notizia presente</p>
          </div>
        )}
      </div>
    </div>
  );
};

const Dashboard: React.FC<any> = ({ stats, setView, user, news, onOpenNews }) => {
  const [selectedMonth, setSelectedMonth] = useState(new Date().getMonth());
  const [selectedYear, setSelectedYear] = useState(new Date().getFullYear());

  const months = [
    'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno',
    'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'
  ];

  const years = useMemo(() => {
    const currentYear = new Date().getFullYear();
    const startYear = 2024;
    const result = [];
    for (let y = currentYear + 1; y >= startYear; y--) result.push(y);
    return result;
  }, []);

  // Implementazione POD univoco: raggruppa contratti del mese per PDP e prendi il più recente per ogni servizio
  const monthlyContracts = useMemo(() => {
    const currentMonthContracts = stats.contracts.filter((c: Contract) => {
      const dateStr = c.creationDate || c.createdAt;
      if (!dateStr) return false;
      
      let d: Date;
      if (dateStr.includes('/')) {
        const [day, month, year] = dateStr.split('/').map(Number);
        d = new Date(year, month - 1, day);
      } else {
        d = new Date(dateStr);
      }
      
      return d.getMonth() === selectedMonth && d.getFullYear() === selectedYear && c.status !== 'KO';
    });

    const uniquePodMap = new Map<string, Contract>();
    currentMonthContracts.forEach((c: Contract) => {
      // Chiave univoca per PDP e Servizio
      const key = `${c.pdpId}_${c.service}`;
      const existing = uniquePodMap.get(key);
      if (!existing || new Date(c.updatedAt).getTime() > new Date(existing.updatedAt).getTime()) {
        uniquePodMap.set(key, c);
      }
    });

    return Array.from(uniquePodMap.values());
  }, [stats.contracts, selectedMonth, selectedYear]);

  const supplierStats = useMemo(() => {
    const stats: any = { POWER: {}, METANO: {} };
    monthlyContracts.forEach(c => {
      if (!stats[c.service][c.fornitore]) stats[c.service][c.fornitore] = 0;
      stats[c.service][c.fornitore] += c.volume;
    });
    return stats;
  }, [monthlyContracts]);

  const totals = monthlyContracts.reduce((acc: any, c: Contract) => {
    if (!acc[c.service]) return acc;
    const category = (c.status === 'trasmesso' || c.status === 'attivo') ? 'green' : 'yellow';
    acc[c.service][category].volume += (c.volume || 0);
    acc[c.service][category].count += 1;
    return acc;
  }, { 
    POWER: { green: { volume: 0, count: 0 }, yellow: { volume: 0, count: 0 } }, 
    METANO: { green: { volume: 0, count: 0 }, yellow: { volume: 0, count: 0 } } 
  });

  const SupplierBar = ({ data, total }: { data: any, total: number }) => {
    if (total === 0) return null;
    const colors: any = { A2A1: '#10b981', A2A2: '#059669', AXPO: '#3b82f6', AXP2: '#2563eb', DOLO: '#f59e0b', DUFE: '#d97706', OPEN: '#8b5cf6', SORG: '#ec4899' };
    return (
      <div className="w-full h-2 flex rounded-full overflow-hidden bg-slate-100 mt-4">
        {Object.entries(data).map(([supplier, volume]: [any, any]) => (
          <div 
            key={supplier} 
            style={{ width: `${(volume / total) * 100}%`, backgroundColor: colors[supplier] || '#cbd5e1' }} 
            title={`${supplier}: ${Math.round((volume / total) * 100)}%`}
          />
        ))}
      </div>
    );
  };

  const SupplierLegend = ({ data, total }: { data: any, total: number }) => {
    if (total === 0) return null;
    const colors: any = { A2A1: '#10b981', A2A2: '#059669', AXPO: '#3b82f6', AXP2: '#2563eb', DOLO: '#f59e0b', DUFE: '#d97706', OPEN: '#8b5cf6', SORG: '#ec4899' };
    return (
      <div className="flex flex-wrap gap-x-3 gap-y-1 mt-2">
        {Object.entries(data).map(([supplier, volume]: [any, any]) => (
          <div key={supplier} className="flex items-center gap-1">
            <div className="w-2 h-2 rounded-full" style={{ backgroundColor: colors[supplier] || '#cbd5e1' }} />
            <span className="text-[8px] font-black text-slate-400 uppercase">{supplier} {Math.round((volume / total) * 100)}%</span>
          </div>
        ))}
      </div>
    );
  };

  const totalPower = totals.POWER.green.volume + totals.POWER.yellow.volume;
  const totalMetano = totals.METANO.green.volume + totals.METANO.yellow.volume;

  const tpmValue = useMemo(() => {
    const pdpFirstContractMap = new Map<string, string>();
    const allValidContracts = [...stats.contracts]
      .filter((c: Contract) => c.status !== 'KO')
      .sort((a, b) => new Date(a.creationDate || a.createdAt).getTime() - new Date(b.creationDate || b.createdAt).getTime());

    allValidContracts.forEach(c => {
      if (!pdpFirstContractMap.has(c.pdpId)) pdpFirstContractMap.set(c.pdpId, c.id);
    });

    const isJulyAugust = selectedMonth === 6 || selectedMonth === 7;
    const periodContracts = stats.contracts.filter((c: Contract) => {
      if (c.status === 'KO') return false;
      if (pdpFirstContractMap.get(c.pdpId) !== c.id) return false;
      const d = new Date(c.creationDate || c.createdAt);
      if (d.getFullYear() !== selectedYear) return false;
      return isJulyAugust ? (d.getMonth() === 6 || d.getMonth() === 7) : d.getMonth() === selectedMonth;
    });

    const periodAssociazioni = (stats.associazioni || []).filter((a: Associazione) => {
      if (a.status === 'nuovo' || a.isRenewal) return false;
      const d = new Date(a.startDate);
      if (d.getFullYear() !== selectedYear) return false;
      return isJulyAugust ? (d.getMonth() === 6 || d.getMonth() === 7) : d.getMonth() === selectedMonth;
    });

    const tpmFromContracts = periodContracts
      .filter((c: Contract) => c.assignedTo === user.id)
      .reduce((acc: number, c: Contract) => {
        const p = stats.portfolios.find((p: Portfolio) => p.id === c.portfolioId);
        if (!p) return acc;
        return acc + (p.entityType === 'Domestico' ? 1 : p.entityType === 'Condominio' ? 2 : 3);
      }, 0);

    const tpmFromAssociazioni = periodAssociazioni
      .filter((a: Associazione) => a.assignedTo === user.id)
      .reduce((acc: number, a: Associazione) => {
        const p = stats.portfolios.find((p: Portfolio) => p.id === a.portfolioId);
        if (!p) return acc;
        return acc + (p.entityType === 'Domestico' ? 3 : p.entityType === 'Condominio' ? 5 : 8);
      }, 0);

    return tpmFromContracts + tpmFromAssociazioni;
  }, [stats.contracts, stats.portfolios, stats.associazioni, selectedMonth, selectedYear, user.id]);

  return (
    <div className="space-y-6 animate-in fade-in duration-500">
      <header className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h2 className="text-xl font-black tracking-tighter text-emerald-900">Dashboard</h2>
          <p className="text-[10px] text-slate-500 font-medium italic">Monitoraggio flussi Progetto ANTICRISI • {user.role}</p>
        </div>
        <div className="flex items-center gap-4">
          <div className="bg-emerald-900 text-white px-6 py-2.5 rounded-2xl shadow-lg shadow-emerald-900/20 flex items-center gap-3 border border-emerald-800">
            <div className="p-1.5 bg-emerald-400/20 rounded-lg">
              <TrendingUp size={16} className="text-emerald-400" />
            </div>
            <div>
              <p className="text-xl font-black leading-none">{tpmValue} <span className="text-[10px] opacity-50 ml-1 uppercase">TPM</span></p>
            </div>
          </div>
          <div className="flex items-center gap-2 bg-white p-1.5 rounded-xl border border-slate-100 shadow-sm">
            <select className="text-[10px] font-black uppercase bg-transparent outline-none px-2 cursor-pointer" value={selectedMonth} onChange={e => setSelectedMonth(Number(e.target.value))}>
              {months.map((m, i) => <option key={m} value={i}>{m}</option>)}
            </select>
            <select className="text-[10px] font-black uppercase bg-transparent outline-none px-2 cursor-pointer border-l" value={selectedYear} onChange={e => setSelectedYear(Number(e.target.value))}>
              {years.map(y => <option key={y} value={y}>{y}</option>)}
            </select>
          </div>
          <div className="bg-white px-3 py-2 rounded-lg shadow-sm border border-slate-100 flex items-center gap-2">
            <Activity className="text-emerald-500 animate-pulse" size={12} />
            <p className="text-[9px] font-bold text-slate-500 uppercase tracking-tight">Sistema protetto</p>
          </div>
        </div>
      </header>

      <NewsBanner news={news} onOpenDetail={onOpenNews} />
      
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div className="bg-white p-5 rounded-[1.5rem] border border-slate-100 shadow-sm space-y-3">
          <div className="flex items-center gap-2"><Zap className="text-emerald-600" size={16}/><h3 className="font-black text-sm text-slate-800 tracking-tight">Caricato Mese (POWER)</h3><p className="text-[8px] text-slate-400 ml-auto uppercase font-bold">Unico per POD</p></div>
          <div className="grid grid-cols-2 gap-3">
            <div className="bg-emerald-50/50 p-3 rounded-xl border border-emerald-100/50">
              <p className="text-[8px] font-black uppercase text-emerald-600 mb-1">Trasmesso/Attivo</p>
              <div className="flex items-baseline gap-2">
                <p className="text-lg font-black text-emerald-800 tracking-tight">{totals.POWER.green.volume.toLocaleString()} <span className="text-[10px] font-bold opacity-60">KWh</span></p>
                <span className="text-[10px] font-black text-emerald-600 bg-emerald-100/50 px-1.5 py-0.5 rounded-md">{totals.POWER.green.count}</span>
              </div>
            </div>
            <div className="bg-yellow-50/50 p-3 rounded-xl border border-yellow-100/50">
              <p className="text-[8px] font-black uppercase text-yellow-600 mb-1">Bozza/Altro</p>
              <div className="flex items-baseline gap-2">
                <p className="text-lg font-black text-yellow-800 tracking-tight">{totals.POWER.yellow.volume.toLocaleString()} <span className="text-[10px] font-bold opacity-60">KWh</span></p>
                <span className="text-[10px] font-black text-yellow-600 bg-yellow-100/50 px-1.5 py-0.5 rounded-md">{totals.POWER.yellow.count}</span>
              </div>
            </div>
          </div>
          <div className="pt-1">
            <p className="text-[7px] font-black uppercase text-slate-400 tracking-widest mb-1">Distribuzione Fornitori (%)</p>
            <SupplierBar data={supplierStats.POWER} total={totalPower} />
            <SupplierLegend data={supplierStats.POWER} total={totalPower} />
          </div>
        </div>
        <div className="bg-white p-5 rounded-[1.5rem] border border-slate-100 shadow-sm space-y-3">
          <div className="flex items-center gap-2"><Zap className="text-orange-600 rotate-180" size={16}/><h3 className="font-black text-sm text-slate-800 tracking-tight">Caricato Mese (METANO)</h3><p className="text-[8px] text-slate-400 ml-auto uppercase font-bold">Unico per PDR</p></div>
          <div className="grid grid-cols-2 gap-3">
            <div className="bg-emerald-50/50 p-3 rounded-xl border border-emerald-100/50">
              <p className="text-[8px] font-black uppercase text-emerald-600 mb-1">Trasmesso/Attivo</p>
              <div className="flex items-baseline gap-2">
                <p className="text-lg font-black text-emerald-800 tracking-tight">{totals.METANO.green.volume.toLocaleString()} <span className="text-[10px] font-bold opacity-60">smc</span></p>
                <span className="text-[10px] font-black text-emerald-600 bg-emerald-100/50 px-1.5 py-0.5 rounded-md">{totals.METANO.green.count}</span>
              </div>
            </div>
            <div className="bg-yellow-50/50 p-3 rounded-xl border border-yellow-100/50">
              <p className="text-[8px] font-black uppercase text-yellow-600 mb-1">Bozza/Altro</p>
              <div className="flex items-baseline gap-2">
                <p className="text-lg font-black text-yellow-800 tracking-tight">{totals.METANO.yellow.volume.toLocaleString()} <span className="text-[10px] font-bold opacity-60">smc</span></p>
                <span className="text-[10px] font-black text-yellow-600 bg-yellow-100/50 px-1.5 py-0.5 rounded-md">{totals.METANO.yellow.count}</span>
              </div>
            </div>
          </div>
          <div className="pt-1">
            <p className="text-[7px] font-black uppercase text-slate-400 tracking-widest mb-1">Distribuzione Fornitori (%)</p>
            <SupplierBar data={supplierStats.METANO} total={totalMetano} />
            <SupplierLegend data={supplierStats.METANO} total={totalMetano} />
          </div>
        </div>
      </div>

      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        <StatCard label="Portafoglio" value={stats.portfolios.length} icon={<Users className="text-emerald-500" />} onClick={() => setView('portfolio')} />
        <StatCard label="PDP Attivi" value={stats.pdps.length} icon={<Zap className="text-yellow-500" />} onClick={() => setView('pdp')} />
        <StatCard label="Contratti" value={stats.contracts.length} icon={<FileText className="text-emerald-500" />} onClick={() => setView('contracts')} />
        <StatCard label="Casi Aperti" value={stats.casi.length} icon={<AlertCircle className="text-red-500" />} onClick={() => setView('casi')} />
      </div>
    </div>
  );
};

const PortfolioView: React.FC<any> = ({ data, setData, user, users, exportFn, setCreditChecks, onCreateContract }) => {
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [form, setForm] = useState<Partial<Portfolio>>({ attachments: [], entityType: 'Domestico' });
  const [editingId, setEditingId] = useState<string | null>(null);
  const [creditCheckPortfolio, setCreditCheckPortfolio] = useState<Portfolio | null>(null);
  const [searchTerm, setSearchTerm] = useState('');

  const ENTITY_TYPES: PortfolioEntityType[] = ['Domestico', 'Impresa', 'Associazione', 'Condominio', 'PA'];

  const filteredPortfolios = useMemo(() => {
    if (!searchTerm) return data;
    const s = searchTerm.toLowerCase();
    return data.filter((p: Portfolio) => 
      p.businessName?.toLowerCase().includes(s) || 
      p.taxId?.toLowerCase().includes(s) ||
      p.city?.toLowerCase().includes(s) ||
      p.street?.toLowerCase().includes(s)
    );
  }, [data, searchTerm]);

  const handleAddressSync = (val: string) => {
    const s = splitAddress(val);
    if (s) setForm(prev => ({ ...prev, legalAddress: val, ...s }));
    else setForm(prev => ({ ...prev, legalAddress: val }));
  };

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    if (!form.taxId) return alert("C.F. / P.IVA obbligatorio");
    
    // Controllo univocità taxId
    const isDuplicate = data.some((p: Portfolio) => 
      p.taxId?.toUpperCase() === form.taxId?.toUpperCase() && p.id !== editingId
    );
    
    if (isDuplicate) {
      return alert(`Errore: Esiste già un'anagrafica con C.F./P.IVA ${form.taxId.toUpperCase()}`);
    }

    const ts = getTimestamp();
    const payload = { ...form, assignedTo: form.assignedTo || user.id, updatedBy: user.name };
    if (editingId) {
      setData((prev: Portfolio[]) => prev.map(p => p.id === editingId ? { ...p, ...payload, updatedAt: ts } : p));
    } else {
      setData((prev: Portfolio[]) => [...prev, { ...payload, id: Math.random().toString(36).substr(2, 9), createdAt: ts, updatedAt: ts } as Portfolio]);
    }
    setIsFormOpen(false);
    setEditingId(null);
    setForm({ attachments: [] });
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-300">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div><h2 className="text-3xl font-black tracking-tight text-emerald-900">Portafoglio 1335</h2><p className="text-slate-500 font-medium">Anagrafiche certificate</p></div>
        <div className="flex flex-wrap items-center gap-3 w-full md:w-auto">
          <div className="relative flex-1 md:w-64">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
            <input 
              className="w-full pl-10 pr-4 py-3 rounded-2xl border-2 border-slate-100 outline-none focus:border-emerald-500 font-bold text-sm" 
              placeholder="Cerca per nome, CF/PIVA, città..." 
              value={searchTerm}
              onChange={e => setSearchTerm(e.target.value)}
            />
          </div>
          <button onClick={() => exportFn(data, 'Portfolio')} className="p-3 text-slate-400 hover:text-slate-800 transition-colors"><FileSpreadsheet /></button>
          <button onClick={() => { setForm({attachments: [], assignedTo: user.id}); setEditingId(null); setIsFormOpen(true); }} className="bg-emerald-600 text-white px-6 py-3 rounded-2xl font-black flex items-center gap-2 shadow-lg hover:scale-105 transition-all"><Plus size={20}/> Nuovo Cliente</button>
        </div>
      </div>
      {isFormOpen && (
        <form onSubmit={handleSave} className="bg-white p-8 rounded-[2.5rem] border-2 border-emerald-50 shadow-2xl space-y-8 animate-in slide-in-from-top-4 duration-500">
           <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
             <div className="space-y-1.5">
               <label className="text-xs font-bold text-slate-500 uppercase ml-1">Tipologia Cliente</label>
               <select 
                 className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500"
                 value={form.entityType || 'Domestico'}
                 onChange={e => setForm({...form, entityType: e.target.value as PortfolioEntityType})}
               >
                 {ENTITY_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
               </select>
             </div>
             <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Ragione Sociale</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" required value={form.businessName || ''} onChange={e => setForm({...form, businessName: e.target.value})} /></div>
             <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Cellulare</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" value={form.phone || ''} onChange={e => setForm({...form, phone: e.target.value})} /></div>
             <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">E-mail</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" value={form.email || ''} onChange={e => setForm({...form, email: e.target.value})} /></div>
             <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">PEC</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" value={form.pec || ''} onChange={e => setForm({...form, pec: e.target.value})} /></div>

             {form.entityType !== 'Domestico' && (
               <>
                 <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Nome Rapp. Legale</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" value={form.legalRepresentativeName || ''} onChange={e => setForm({...form, legalRepresentativeName: e.target.value})} /></div>
                 <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">CF Rapp. Legale</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-mono font-bold focus:border-emerald-500" maxLength={16} value={form.legalRepresentativeTaxId || ''} onChange={e => setForm({...form, legalRepresentativeTaxId: e.target.value.toUpperCase()})} /></div>
               </>
             )}
             <div className="space-y-1.5">
               <label className="text-xs font-bold text-slate-500 uppercase ml-1 flex justify-between">
                 <span>CF o P.IVA</span>
                 <span className="text-[9px] text-slate-400">{form.entityType === 'Domestico' ? '16 caratteri' : '11 caratteri'}</span>
               </label>
               <input 
                 className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-mono font-bold focus:border-emerald-500" 
                 required 
                 maxLength={form.entityType === 'Domestico' ? 16 : 11}
                 value={form.taxId || ''} 
                 onChange={e => setForm({...form, taxId: e.target.value.toUpperCase()})} 
               />
             </div>
             {(user.role === 'admin' || user.role === 'backoffice') && (
               <div className="col-span-full">
                 <SearchableSelect label="Assegna a Utente" placeholder="Cerca utente..." options={users} selectedId={form.assignedTo} onSelect={(id:string) => setForm({...form, assignedTo: id})} displayFn={(u:any)=>u.name} searchFn={(u:any)=>u.name} subTextFn={(u:any)=>u.role} icon={<UserIcon size={14}/>} />
               </div>
             )}
             <div className="col-span-full space-y-1.5">
                <label className="text-xs font-bold text-slate-500 uppercase ml-1 flex items-center gap-2">
                    <MapPin size={12}/> Indirizzo (Via, CAP Città, PR) 
                    {form.legalAddress && <AddressLink address={form.legalAddress} label="(Vedi su Maps)" className="text-[10px] font-black uppercase text-blue-500 underline" />}
                </label>
                <input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none" placeholder="Es: Via Roma 1, 20121 Milano, MI" value={form.legalAddress || ''} onChange={e => handleAddressSync(e.target.value)} />
             </div>
             <div className="grid grid-cols-4 gap-3 col-span-full bg-slate-50 p-4 rounded-xl border border-slate-100">
               <div><label className="text-[10px] font-black uppercase text-slate-400">Via</label><input className="w-full bg-white p-2 rounded-lg text-xs font-bold border" value={form.street || ''} onChange={e => setForm({...form, street: e.target.value})} /></div>
               <div><label className="text-[10px] font-black uppercase text-slate-400">CAP</label><input className="w-full bg-white p-2 rounded-lg text-xs font-bold border" value={form.cap || ''} onChange={e => setForm({...form, cap: e.target.value})} /></div>
               <div><label className="text-[10px] font-black uppercase text-slate-400">Comune</label><input className="w-full bg-white p-2 rounded-lg text-xs font-bold border" value={form.city || ''} onChange={e => setForm({...form, city: e.target.value})} /></div>
               <div><label className="text-[10px] font-black uppercase text-slate-400">PR</label><input className="w-full bg-white p-2 rounded-lg text-xs font-bold border uppercase" value={form.province || ''} onChange={e => setForm({...form, province: e.target.value.toUpperCase()})} /></div>
             </div>
             <div className="col-span-full"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Allegati</label>
               <FileUploader attachments={form.attachments || []} canEdit={user.role !== 'utente'} onUpload={files => setForm({...form, attachments: [...(form.attachments || []), ...files]})} onRemove={id => setForm({...form, attachments: form.attachments?.filter(a => a.id !== id)})} />
             </div>
           </div>
           {editingId && <TimestampDetail data={form} />}
           <div className="flex justify-end gap-4 border-t pt-8"><button type="button" onClick={() => setIsFormOpen(false)} className="font-bold text-slate-400">Annulla</button><button type="submit" className="bg-emerald-900 text-white px-10 py-4 rounded-2xl font-black">{editingId ? 'Aggiorna Cliente' : 'Registra Anagrafica'}</button></div>
        </form>
      )}
      <div className="bg-white rounded-[2rem] border border-slate-100 shadow-sm overflow-hidden">
        <table className="w-full text-left">
          <thead className="bg-slate-50 border-b">
            <tr><th className="p-6 text-[10px] font-black uppercase text-slate-400">Tipologia / Ragione Sociale</th><th className="p-6 text-[10px] font-black uppercase text-slate-400">Assegnato (Codice)</th><th className="p-6"></th></tr>
          </thead>
          <tbody className="divide-y divide-slate-50">
            {filteredPortfolios.map((p: Portfolio) => {
              const assignedUser = users.find(u => u.id === p.assignedTo);
              return (
              <tr key={p.id} className="hover:bg-slate-50 transition-colors group">
                <td className="p-6">
                  <p className="text-[10px] font-black text-emerald-600 uppercase tracking-tighter mb-1">{p.entityType || '---'}</p>
                  <p className="font-bold">{p.businessName}</p>
                  <p className="text-[10px] text-slate-400 font-mono">{p.taxId}</p>
                  {p.updatedBy && (
                    <p className="text-[8px] text-slate-300 uppercase font-black mt-1 flex items-center gap-1">
                      <Clock size={8} /> Modificato da: {p.updatedBy}
                    </p>
                  )}
                </td>
                <td className="p-6 text-right">
                  <p className="text-sm font-black text-emerald-700 leading-none">{assignedUser?.userCode || '---'}</p>
                   <p className="text-[10px] font-bold text-slate-400 uppercase tracking-tighter mt-1">{assignedUser?.name || '---'}</p>
                </td>
                <td className="p-6 text-right space-x-2">
                  <div className="flex items-center justify-end gap-2">
                    <AddressLink address={p.legalAddress} showIconOnly className="text-blue-500 hover:scale-110 transition-transform" />
                     <button onClick={() => setCreditCheckPortfolio(p)} className="text-emerald-600 hover:bg-emerald-50 p-2 rounded-lg border border-emerald-100 transition-all" title="Credit Check">
                       <Wallet size={18} />
                     </button>
                     <button onClick={() => onCreateContract(p.id)} className="text-blue-600 hover:bg-blue-50 p-2 rounded-lg border border-blue-100 transition-all" title="Crea Contratto">
                       <FileSignature size={18} />
                     </button>
                    <button onClick={() => { setForm(p); setEditingId(p.id); setIsFormOpen(true); }} className="text-slate-400 hover:text-emerald-600 p-2"><Edit2 size={16}/></button>
                    <button onClick={() => setData(data.filter((i: any) => i.id !== p.id))} className="text-slate-400 hover:text-red-600 p-2"><X size={18}/></button>
                  </div>
                </td>
              </tr>
            )})}
          </tbody>
        </table>
      </div>

      {creditCheckPortfolio && (
        <CreditCheckModal 
          isOpen={!!creditCheckPortfolio} 
          onClose={() => setCreditCheckPortfolio(null)} 
          portfolio={creditCheckPortfolio} 
          onSave={(req) => {
            const ts = getTimestamp();
            setCreditChecks((prev: CreditCheckRequest[]) => [...prev, { ...req, id: Math.random().toString(36).substr(2, 9), assignedTo: user.id, createdAt: ts, updatedAt: ts } as CreditCheckRequest]);
          }} 
        />
      )}
    </div>
  );
};

const PDPView: React.FC<any> = ({ data, setData, cabine, setCabine, contracts, portfolios, users, user, exportFn }) => {
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [form, setForm] = useState<Partial<PDP>>({});
  const [editingId, setEditingId] = useState<string | null>(null);

  const handleAddressSync = (val: string) => {
    const s = splitAddress(val);
    if (s) setForm(prev => ({ ...prev, address: val, ...s }));
    else setForm(prev => ({ ...prev, address: val }));
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-300">
      <div className="flex justify-between items-center">
        <div><h2 className="text-3xl font-black text-emerald-900">Punti Di Prelievo 1335</h2><p className="text-slate-500 font-medium italic">Database tecnico utenze</p></div>
        <div className="flex gap-3">
          <button onClick={() => exportFn(data, 'PDP')} className="p-3 text-slate-400 hover:text-slate-800 transition-colors"><FileSpreadsheet /></button>
          <button onClick={() => { setForm({}); setEditingId(null); setIsFormOpen(true); }} className="bg-yellow-500 text-white px-6 py-3 rounded-2xl font-black flex items-center gap-2 shadow-lg hover:scale-105 transition-all"><Plus size={20}/> Nuovo PDP</button>
        </div>
      </div>
      {isFormOpen && (
        <form onSubmit={(e) => {
          e.preventDefault();
          const ts = getTimestamp();
          if (editingId) setData((prev: PDP[]) => prev.map(p => p.id === editingId ? { ...p, ...form, updatedAt: ts } : p));
          else setData((prev: PDP[]) => [...prev, { ...form, id: Math.random().toString(36).substr(2, 9), assignedTo: user.id, createdAt: ts, updatedAt: ts } as PDP]);
          setIsFormOpen(false); setEditingId(null); setForm({});
        }} className="bg-white p-8 rounded-[2.5rem] border-2 border-yellow-50 shadow-2xl space-y-6 animate-in slide-in-from-top-4">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Codice PDP</label><input maxLength={14} minLength={14} pattern=".{14}" title="Il codice PDP deve essere di 14 caratteri" className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-mono font-black uppercase" required value={form.pdpCode || ''} onChange={e => setForm({...form, pdpCode: e.target.value.toUpperCase()})} /></div>
            <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Indirizzo Fornitura</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none" value={form.address || ''} onChange={e => handleAddressSync(e.target.value)} /></div>
            <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Potenza Impegnata (kW)</label><input type="number" step="0.1" className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold" value={form.potenzaImpegnata || ''} onChange={e => setForm({...form, potenzaImpegnata: Number(e.target.value)})} /></div>
            <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Potenza Disponibile (kW)</label><input type="number" step="0.1" className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold" value={form.potenzaDisponibile || ''} onChange={e => setForm({...form, potenzaDisponibile: Number(e.target.value)})} /></div>
          </div>
          <div className="flex justify-end gap-4 border-t pt-6"><button type="button" onClick={() => setIsFormOpen(false)} className="font-bold text-slate-400">Annulla</button><button type="submit" className="bg-emerald-900 text-white px-10 py-4 rounded-2xl font-black">{editingId ? 'Aggiorna PDP' : 'Salva PDP'}</button></div>
        </form>
      )}
      <div className="bg-white rounded-3xl overflow-hidden shadow-sm border border-slate-100">
        <table className="w-full text-left">
          <thead className="bg-slate-50 border-b">
            <tr><th className="p-6 text-[10px] font-black uppercase text-slate-400">Codice</th><th className="p-6 text-[10px] font-black uppercase text-slate-400">Località (Maps)</th><th className="p-6 text-[10px] font-black uppercase text-slate-400">Pot Imp/Disp</th><th className="p-6 text-[10px] font-black uppercase text-slate-400">Utente (Codice)</th><th className="p-6"></th></tr>
          </thead>
          <tbody className="divide-y divide-slate-50">
            {data.map((p: PDP) => {
              const lastContract = [...contracts].filter(c => c.pdpId === p.id).sort((a,b) => new Date(b.startDate).getTime() - new Date(a.startDate).getTime())[0];
              const portfolio = lastContract ? portfolios.find(pf => pf.id === lastContract.portfolioId) : null;
              const assignedUser = portfolio ? users.find(u => u.id === portfolio.assignedTo) : null;
              
              return (
                <tr key={p.id} className="hover:bg-slate-50 transition-colors group">
                  <td className="p-6 font-mono font-black">{p.pdpCode}</td>
                  <td className="p-6 text-sm font-bold">
                     <div className="flex items-center gap-2">
                       <AddressLink address={p.address} showIconOnly className="text-blue-500" />
                       <div>
                         <p className="font-bold text-slate-800">{p.city || '---'} ({p.province || '--'})</p>
                         <p className="text-[10px] text-slate-400">{p.street || '---'}</p>
                       </div>
                     </div>
                  </td>
                  <td className="p-6 font-bold text-emerald-700 text-xs">{p.potenzaImpegnata || 0} / {p.potenzaDisponibile || 0} kW</td>
                  <td className="p-6 text-right">
                    <p className="text-lg font-black text-emerald-700 leading-none">{assignedUser?.userCode || '---'}</p>
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-tighter mt-1">{assignedUser?.name || '---'}</p>
                  </td>
                  <td className="p-6 text-right space-x-2">
                    <button onClick={() => { setForm(p); setEditingId(p.id); setIsFormOpen(true); }} className="text-slate-400 hover:text-blue-600 p-2"><Edit2 size={16}/></button>
                    <button onClick={() => setData(data.filter((i:any) => i.id !== p.id))} className="text-slate-400 hover:text-red-600 p-2"><X size={18}/></button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
};

const ContractView: React.FC<any> = ({ data, setData, portfolios, setPortfolios, pdps, cabine, setCabine, setPdps, users, user, exportFn, onOpenChat, getUnreadCount, prefilledPortfolioId, onClearPrefill }) => {
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [selectedPortfolio, setSelectedPortfolio] = useState<Portfolio | null>(null);
  const [selectedPDP, setSelectedPDP] = useState<PDP | null>(null);
  const [form, setForm] = useState<Partial<Contract & PDP>>({ 
    status: 'bozza', 
    service: 'POWER', 
    contractType: 'SWITCH',
    attachments: [], 
    fornitore: 'A2A1',
    creationDate: new Date().toISOString().split('T')[0],
    startDate: calculateActivationDate(new Date().toISOString().split('T')[0]),
    durationMonths: 12,
    endDate: calculateEndDate(calculateActivationDate(new Date().toISOString().split('T')[0]), 12)
  });
  const [editingId, setEditingId] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState('all');

  useEffect(() => {
    if (prefilledPortfolioId) {
      const today = new Date().toISOString().split('T')[0];
      const start = calculateActivationDate(today);
      setForm({
        status: 'bozza',
        service: 'POWER',
        contractType: 'SWITCH',
        attachments: [],
        fornitore: 'A2A1',
        creationDate: today,
        startDate: start,
        durationMonths: 12,
        endDate: calculateEndDate(start, 12),
        portfolioId: prefilledPortfolioId
      });
      setEditingId(null);
      setIsFormOpen(true);
      onClearPrefill();
    }
  }, [prefilledPortfolioId, onClearPrefill]);

  const filteredData = useMemo(() => {
    return data.filter((c: Contract) => {
      const port = portfolios.find(p => p.id === c.portfolioId);
      const pdp = pdps.find(p => p.id === c.pdpId);
      
      const matchesStatus = statusFilter === 'all' || c.status === statusFilter;
      
      const searchStr = `${port?.businessName || ''} ${pdp?.pdpCode || ''} ${pdp?.address || ''} ${c.fornitore || ''} ${c.tariffa || ''} ${c.contractType || ''} ${c.service || ''}`.toLowerCase();
      const matchesSearch = searchStr.includes(searchTerm.toLowerCase());
      
      return matchesStatus && matchesSearch;
    });
  }, [data, searchTerm, statusFilter, portfolios, pdps]);

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    if (!form.portfolioId || !form.pdpId) return alert("Associazione Cliente e PDP obbligatoria.");
    
    // Validazione PDP
    if (form.contractType === 'ALLACCIO') {
      // Per ALLACCIO può essere "da assegnare" o 13 caratteri
      const code = form.pdpId.startsWith('NEW_') ? form.pdpId.replace('NEW_', '') : form.pdpId;
      if (code.toLowerCase() !== 'da assegnare' && code.length !== 13) {
        return alert("Per i nuovi allacci il codice PDP deve essere 'da assegnare' o di 13 caratteri.");
      }
    } else {
      // Per gli altri casi deve essere 13 caratteri
      const code = form.pdpId.startsWith('NEW_') ? form.pdpId.replace('NEW_', '') : form.pdpId;
      // Se è un PDP esistente, cerchiamo il codice reale
      const existingPdp = pdps.find((p:any) => p.id === form.pdpId);
      const finalCode = existingPdp ? existingPdp.pdpCode : code;
      
      if (finalCode.length !== 13) {
        return alert("Il codice PDP deve essere obbligatoriamente di 13 caratteri.");
      }
    }

    const ts = getTimestamp();
    let finalPdpId = form.pdpId;

    if (finalPdpId && finalPdpId.startsWith('NEW_')) {
      const code = finalPdpId.replace('NEW_', '');
      let cId = undefined;
      if (form.technicalSpecs) {
        let ex = cabine.find((c:any)=>c.name.toLowerCase() === form.technicalSpecs?.toLowerCase());
        if(!ex){
          ex = { id: Math.random().toString(36).substr(2, 9), name: form.technicalSpecs };
          setCabine((prev:any)=>[...prev, ex]);
        }
        cId = ex.id;
      }
      const newPDP: PDP = { id: Math.random().toString(36).substr(2, 9), pdpCode: code, address: form.address || '', street: form.street || '', cap: form.cap || '', city: form.city || '', province: form.province || '', potenzaImpegnata: form.potenzaImpegnata || 0, potenzaDisponibile: form.potenzaDisponibile || 0, technicalSpecs: form.technicalSpecs || '', cabinaPrimariaId: cId, assignedTo: user.id, createdAt: ts, updatedAt: ts };
      setPdps((prev: PDP[]) => [...prev, newPDP]);
      finalPdpId = newPDP.id;
    }

    const payload = { ...form, pdpId: finalPdpId, updatedAt: ts };
    if (editingId) setData((prev: Contract[]) => prev.map(c => c.id === editingId ? { ...c, ...payload } : c));
    else setData((prev: Contract[]) => [...prev, { ...payload, id: Math.random().toString(36).substr(2, 9), assignedTo: user.id, createdAt: ts } as Contract]);
    
    setIsFormOpen(false); setEditingId(null); 
    const today = new Date().toISOString().split('T')[0];
    const start = calculateActivationDate(today);
    setForm({ 
      status: 'bozza', 
      service: 'POWER', 
      contractType: 'SWITCH', 
      attachments: [], 
      fornitore: 'A2A1', 
      creationDate: today, 
      startDate: start,
      durationMonths: 12,
      endDate: calculateEndDate(start, 12)
    });
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-300">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div><h2 className="text-3xl font-black text-emerald-900 tracking-tighter">Contratti 1335</h2><p className="text-slate-500 font-medium">Archivio contrattuale attivo</p></div>
        <div className="flex flex-wrap items-center gap-3 w-full md:w-auto">
          <div className="relative flex-1 md:w-64">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
            <input 
              className="w-full pl-10 pr-4 py-3 rounded-2xl border-2 border-slate-100 outline-none focus:border-emerald-500 font-bold text-sm" 
              placeholder="Cerca per nome, POD, indirizzo..." 
              value={searchTerm}
              onChange={e => setSearchTerm(e.target.value)}
            />
          </div>
          <select 
            className="p-3 rounded-2xl border-2 border-slate-100 outline-none font-bold text-sm bg-white"
            value={statusFilter}
            onChange={e => setStatusFilter(e.target.value)}
          >
            <option value="all">Tutti gli stati</option>
            <option value="bozza">Bozza</option>
            <option value="da firmare">Da Firmare</option>
            <option value="trasmesso">Trasmesso</option>
            <option value="attivo">Attivo</option>
            <option value="non conforme">Non Conforme</option>
            <option value="KO">KO</option>
          </select>
          <button onClick={() => exportFn(data, 'Contratti')} className="p-3 text-slate-400 hover:text-slate-800 transition-colors"><FileSpreadsheet /></button>
          <button onClick={() => { 
            const today = new Date().toISOString().split('T')[0];
            const start = calculateActivationDate(today);
            setForm({
              status: 'bozza', 
              service: 'POWER', 
              contractType: 'SWITCH',
              attachments: [], 
              fornitore: 'A2A1',
              creationDate: today,
              startDate: start,
              durationMonths: 12,
              endDate: calculateEndDate(start, 12)
            }); 
            setEditingId(null); 
            setIsFormOpen(true); 
          }} className="bg-emerald-600 text-white px-8 py-4 rounded-2xl font-black shadow-lg hover:scale-105 transition-all flex items-center gap-2"><Plus /> Nuovo Contratto</button>
        </div>
      </div>
      {isFormOpen && (
        <form onSubmit={handleSave} className="bg-white p-10 rounded-[3rem] border-2 border-emerald-50 shadow-2xl space-y-8 animate-in slide-in-from-top-4">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
            <SearchableSelect label="Cliente *" placeholder="Scegli cliente..." options={portfolios} selectedId={form.portfolioId} onSelect={(id:string) => setForm({...form, portfolioId: id})} displayFn={(p:any)=>p.businessName} searchFn={(p:any)=>`${p.businessName}`} subTextFn={(p:any)=>p.taxId} icon={<Users size={14}/>} />
            <SearchableSelect 
              label="PDP *" 
              placeholder="Scegli PDP o digita nuovo..." 
              options={pdps} 
              selectedId={form.pdpId} 
              onSelect={(id:string) => { 
                setForm({...form, pdpId: id}); 
                const p = pdps.find((x:any)=>x.id===id); 
                if(p) setForm(prev => ({...prev, address: p.address, potenzaImpegnata: p.potenzaImpegnata, potenzaDisponibile: p.potenzaDisponibile, technicalSpecs: p.technicalSpecs })); 
              }} 
              displayFn={(p:any)=>p.pdpCode} 
              searchFn={(p:any)=>p.pdpCode} 
              subTextFn={(p:any)=>p.address} 
              icon={<Zap size={14}/>} 
              allowCustom={true} 
              customLabel="Aggiungi nuovo PDP" 
              disabled={form.contractType === 'ALLACCIO'}
              maxLength={14}
            />
            
            <div className="col-span-full space-y-4 p-8 bg-slate-50 rounded-[2.5rem] border border-slate-100">
              <div className="flex gap-4 items-end">
                <div className="flex-1 space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1">Tipo Operazione</label>
                  <select className="w-full border-2 p-4 rounded-2xl bg-white font-black outline-none border-slate-100" value={form.contractType} onChange={e => {
                    const newType = e.target.value as any;
                    const updates: any = { contractType: newType };
                    if (newType === 'ALLACCIO') {
                      updates.pdpId = 'NEW_DA ASSEGNARE';
                    }
                    setForm({...form, ...updates});
                  }}>
                    {CONTRACT_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
                  </select>
                </div>
                <div className="flex-[2] flex gap-4">
                  {['POWER', 'METANO'].map(s => <button key={s} type="button" onClick={() => setForm({...form, service: s as any})} className={`flex-1 p-4 rounded-2xl font-black transition-all border-2 ${form.service === s ? 'bg-emerald-900 text-white border-emerald-700 shadow-xl' : 'bg-white text-slate-400 border-slate-100'}`}>{s}</button>)}
                </div>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="col-span-full">
                  <div className="flex justify-between items-center mb-1.5">
                    <label className="text-xs font-bold text-slate-500 uppercase ml-1"><MapPin size={12}/> Indirizzo Fornitura</label>
                    {form.portfolioId && (
                      <button 
                        type="button" 
                        onClick={() => {
                          const p = portfolios.find((x:any)=>x.id === form.portfolioId);
                          if (p) {
                            const s = splitAddress(p.legalAddress);
                            setForm(prev => ({ ...prev, address: p.legalAddress, ...(s || {}) }));
                          }
                        }}
                        className="text-[10px] font-black uppercase text-emerald-600 hover:text-emerald-700 flex items-center gap-1"
                      >
                        <Activity size={10} /> Usa Sede Legale
                      </button>
                    )}
                  </div>
                  <input className="w-full border-2 p-3.5 rounded-2xl bg-white outline-none focus:border-emerald-500" value={form.address || ''} onChange={e => { const s = splitAddress(e.target.value); setForm(prev => ({ ...prev, address: e.target.value, ...(s || {}) })); }} />
                </div>
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1">Volume Annuale</label>
                  <div className="relative">
                    <input type="number" className="w-full border-2 p-3.5 rounded-2xl bg-white outline-none font-bold pr-24" value={form.volume || ''} onChange={e => setForm({...form, volume: Number(e.target.value)})} />
                    <span className="absolute right-4 top-1/2 -translate-y-1/2 text-[10px] font-black text-emerald-600 bg-emerald-50 px-2 py-1 rounded-lg">{formatVolume(form.volume || 0)}</span>
                  </div>
                </div>
                {form.service === 'POWER' && (
                  <>
                    <div className="grid grid-cols-2 gap-4">
                      <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Pot. Impegnata</label><input type="number" step="0.1" className="w-full border-2 p-3.5 rounded-2xl bg-white outline-none font-bold" value={form.potenzaImpegnata || ''} onChange={e => setForm({...form, potenzaImpegnata: Number(e.target.value)})} /></div>
                      <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Pot. Disponibile</label><input type="number" step="0.1" className="w-full border-2 p-3.5 rounded-2xl bg-white outline-none font-bold" value={form.potenzaDisponibile || ''} onChange={e => setForm({...form, potenzaDisponibile: Number(e.target.value)})} /></div>
                    </div>
                    <div className="space-y-1.5">
                      <label className="text-xs font-bold text-slate-500 uppercase ml-1">Opz. Tar.</label>
                      <select 
                        className="w-full border-2 p-3.5 rounded-2xl bg-white font-bold outline-none border-slate-100" 
                        value={form.tariffOption || ''} 
                        onChange={e => setForm({...form, tariffOption: e.target.value})}
                      >
                        <option value="">Seleziona opzione...</option>
                        {['DOM2', 'DOM3', 'BTA', 'MTA', 'BTIP', 'BTVE', 'MTIP', 'MTVE'].map(opt => (
                          <option key={opt} value={opt}>{opt}</option>
                        ))}
                      </select>
                    </div>
                  </>
                )}
                <div className="col-span-full"><label className="text-xs font-bold text-slate-500 uppercase ml-1 text-emerald-600">Cabina Primaria</label><input disabled={form.service === 'METANO'} className="w-full border-2 p-3.5 rounded-2xl bg-white outline-none disabled:opacity-20 border-emerald-100" placeholder={form.service === 'POWER' ? "CP ..." : "N/A per Metano"} value={form.technicalSpecs || ''} onChange={e => setForm({...form, technicalSpecs: e.target.value})} /></div>
              </div>
            </div>

            <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Fornitore</label>
              <select className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 font-bold outline-none" value={form.fornitore} onChange={e => setForm({...form, fornitore: e.target.value as Fornitore})}>
                {FORNITORI.map(f => <option key={f} value={f}>{f}</option>)}
              </select>
            </div>
            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Data Creazione</label><input type="date" disabled={user.role === 'utente'} className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold disabled:opacity-50 text-xs" value={form.creationDate || ''} onChange={e => { const d = e.target.value; const start = calculateActivationDate(d); setForm({...form, creationDate: d, startDate: start, endDate: calculateEndDate(start, form.durationMonths || 12)}); }} /></div>
              <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Data Attivazione</label><input type="date" className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold text-xs" value={form.startDate || ''} onChange={e => { const d = e.target.value; setForm({...form, startDate: d, endDate: calculateEndDate(d, form.durationMonths || 12)}); }} /></div>
            </div>
            <div className="grid grid-cols-4 gap-4">
              <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Durata (Mesi)</label><input type="number" className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold text-xs" value={form.durationMonths || 12} onChange={e => { const m = Number(e.target.value); setForm({...form, durationMonths: m, endDate: calculateEndDate(form.startDate || '', m)}); }} /></div>
              <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Fine Contratto</label><input type="date" disabled className="w-full border-2 p-3.5 rounded-2xl bg-slate-200 outline-none font-bold opacity-70 text-xs" value={form.endDate || ''} /></div>
              <div className="col-span-2 space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Tariffa</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold text-xs" value={form.tariffa || ''} onChange={e => setForm({...form, tariffa: e.target.value})} /></div>
            </div>
            <div className="col-span-full"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Allegati Contratto</label>
              <FileUploader attachments={form.attachments || []} canEdit={true} onUpload={files => setForm({...form, attachments: [...(form.attachments || []), ...files]})} onRemove={id => setForm({...form, attachments: form.attachments?.filter(a => a.id !== id)})} />
            </div>
            <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Stato</label><select className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 font-bold outline-none" value={form.status} onChange={e => setForm({...form, status: e.target.value as any})}><option value="bozza">Bozza</option><option value="da firmare">Da Firmare</option><option value="trasmesso">Trasmesso</option><option value="attivo">Attivo</option><option value="non conforme">Non Conforme</option><option value="KO">KO</option></select></div>
          </div>
          <div className="flex justify-end gap-4 border-t pt-8"><button type="button" onClick={() => setIsFormOpen(false)} className="font-bold text-slate-400">Annulla</button><button type="submit" className="bg-emerald-900 text-white px-12 py-5 rounded-2xl font-black shadow-xl">Salva Contratto</button></div>
        </form>
      )}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        {filteredData.map((c: Contract) => {
          const port = portfolios.find(p => p.id === c.portfolioId);
          const pdp = pdps.find(p => p.id === c.pdpId);
          const assignedUser = port ? users.find(u => u.id === port.assignedTo) : null;
          
          return (
            <div key={c.id} className="bg-white p-6 rounded-[2.5rem] border border-slate-100 shadow-sm hover:shadow-2xl transition-all relative group overflow-hidden">
              <div className={`absolute top-0 right-0 px-6 py-2 rounded-bl-2xl text-[10px] font-black uppercase tracking-widest flex items-center gap-2 ${['attivo', 'trasmesso'].includes(c.status) ? 'bg-emerald-100 text-emerald-700' : c.status === 'KO' ? 'bg-red-100 text-red-700' : 'bg-yellow-100 text-yellow-700'}`}>
                {c.contractType && <span className="border-r border-current pr-2 mr-2 opacity-50">{c.contractType}</span>}
                {c.status}
              </div>
              <div className="flex gap-6">
                <div className="flex flex-col gap-2 shrink-0">
                  <div className={`p-4 rounded-2xl text-white h-fit ${c.service === 'POWER' ? 'bg-red-600' : 'bg-sky-500'}`}>{c.service === 'POWER' ? <Zap/> : <Flame/>}</div>
                </div>
                <div className="flex-1 space-y-0 overflow-hidden">
                   <div className="flex justify-between gap-4">
                     <div className="overflow-hidden flex-1">
                       <p className="text-[10px] font-black uppercase text-slate-400 truncate">Intestatario</p>
                       <h4 onClick={() => port && setSelectedPortfolio(port)} className="text-xl font-black tracking-tight text-emerald-950 truncate cursor-pointer hover:text-emerald-600 transition-colors">{port?.businessName || 'Anagrafica N/D'}</h4>
                       <div className="flex items-center gap-2 mt-1">
                          <p onClick={() => pdp && setSelectedPDP(pdp)} className="text-[11px] font-black text-emerald-600 font-mono bg-emerald-50 px-2 py-0.5 rounded-md cursor-pointer hover:bg-emerald-100 transition-colors">{pdp?.pdpCode || 'CODICE PDP N/D'}</p>
                          <p className="text-[10px] font-bold text-slate-400">| {c.fornitore} {c.tariffOption || c.tariffa ? `- ${c.tariffOption || c.tariffa}` : ''}</p>
                       </div>
                     </div>
                     <div className="text-right shrink-0 pt-4">
                        <p className="text-lg font-black text-emerald-700 leading-none">{assignedUser?.userCode || '---'}</p>
                        <p className="text-[10px] font-bold text-slate-400 uppercase tracking-tighter mt-1">{assignedUser?.name || '---'}</p>
                        <div className="opacity-0 group-hover:opacity-100 transition-opacity flex gap-1 justify-end mt-2">
                           <button onClick={() => onOpenChat(c.id)} className="p-2 text-slate-300 hover:text-blue-500 relative">
                              <MessageSquare size={16}/>
                              {getUnreadCount(c.id) > 0 && (
                                <span className="absolute top-0 right-0 w-2 h-2 bg-red-500 rounded-full border-2 border-white"></span>
                              )}
                           </button>
                           <button onClick={() => { setForm(c); setEditingId(c.id); setIsFormOpen(true); }} className="p-2 text-slate-300 hover:text-emerald-500"><Edit2 size={16}/></button>
                           <button onClick={() => setData(data.filter(i=>i.id !== c.id))} className="p-2 text-slate-300 hover:text-red-500"><X size={18}/></button>
                        </div>
                     </div>
                   </div>
                   <div className="flex justify-between items-center text-[10px] font-bold text-slate-400 px-1">
                      <div className="flex gap-3">
                         <span>DEC: {formatDate(c.startDate)}</span>
                         {c.endDate && <span>FINE: {formatDate(c.endDate)}</span>}
                         <span>Volume: {c.volume?.toLocaleString()} {c.service === 'POWER' ? 'KWh' : 'smc'}</span>
                      </div>
                   </div>
                </div>
              </div>
            </div>
          );
        })}
      </div>

      {selectedPortfolio && (
        <PortfolioModal 
          isOpen={!!selectedPortfolio} 
          onClose={() => setSelectedPortfolio(null)} 
          portfolio={selectedPortfolio} 
          setData={setPortfolios}
          user={user}
          users={users}
        />
      )}

      {selectedPDP && (
        <PDPModal 
          isOpen={!!selectedPDP} 
          onClose={() => setSelectedPDP(null)} 
          pdp={selectedPDP} 
          setData={setPdps}
          user={user}
        />
      )}
    </div>
  );
};

const CabineView: React.FC<any> = ({ data, pdps, contracts, exportFn }) => (
  <div className="space-y-6 animate-in fade-in duration-300">
    <div className="flex justify-between items-center"><div><h2 className="text-3xl font-black text-emerald-900">Cabine Primarie</h2><p className="text-slate-500 font-medium italic">Monitoraggio nodi di rete</p></div><button onClick={() => exportFn(data, 'Cabine')} className="p-3 text-slate-400 hover:text-slate-800"><FileSpreadsheet /></button></div>
    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">{data.map((c: any) => {
      const cabinaPdps = pdps.filter((p:any)=>p.cabinaPrimariaId === c.id);
      const totalKwh = cabinaPdps.reduce((acc: number, pdp: PDP) => {
        const pdpContracts = contracts.filter((ct: Contract) => ct.pdpId === pdp.id && ct.service === 'POWER');
        const last = pdpContracts.sort((a: any, b: any) => new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime())[0];
        return acc + (last ? last.volume : 0);
      }, 0);

      return (
        <div key={c.id} className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm flex flex-col gap-4">
          <div className="flex justify-between items-start"><div className="bg-emerald-50 p-3 rounded-2xl"><HardDrive className="text-emerald-600"/></div><span className="bg-emerald-100 text-emerald-700 px-3 py-1 rounded-full text-[10px] font-black uppercase">{cabinaPdps.length} PDP</span></div>
          <h4 className="text-xl font-black text-slate-800 tracking-tighter">{c.name}</h4>
          <div className="bg-emerald-900 p-4 rounded-xl text-white">
            <p className="text-[10px] font-black uppercase opacity-70">Potenza Caricata Totale (Recentest)</p>
            <p className="text-2xl font-black tracking-tighter">{totalKwh.toLocaleString()} KWh</p>
          </div>
        </div>
      );
    })}</div>
  </div>
);

const CasiView: React.FC<any> = ({ data, setData, portfolios, pdps, user, exportFn, onOpenChat, getUnreadCount }) => {
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [form, setForm] = useState<Partial<Caso>>({ status: 'nuovo', attachments: [] });
  const [editingId, setEditingId] = useState<string | null>(null);

  const availablePdps = useMemo(() => {
    if (!form.portfolioId) return pdps;
    return pdps; 
  }, [form.portfolioId, pdps]);

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    const ts = getTimestamp();
    if (editingId) setData((prev: Caso[]) => prev.map(c => c.id === editingId ? { ...c, ...form, updatedAt: ts } : c));
    else setData((prev: Caso[]) => [...prev, { ...form, id: Math.random().toString(36).substr(2, 9), assignedTo: user.id, createdAt: ts, updatedAt: ts } as Caso]);
    setIsFormOpen(false); setEditingId(null); setForm({ status: 'nuovo', attachments: [] });
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-300">
      <div className="flex justify-between items-center"><div><h2 className="text-3xl font-black text-emerald-900">Gestione Casi</h2><p className="text-slate-500 font-medium">Ticketing operativo</p></div><div className="flex gap-3"><button onClick={() => exportFn(data, 'Casi')} className="p-3 text-slate-400 hover:text-slate-800 transition-colors"><FileSpreadsheet /></button><button onClick={() => { setForm({status: 'nuovo', attachments: []}); setEditingId(null); setIsFormOpen(true); }} className="bg-red-500 text-white px-6 py-3 rounded-2xl font-black flex items-center gap-2 shadow-lg hover:scale-105 transition-all"><Plus /> Apri Caso</button></div></div>
      {isFormOpen && (
        <form onSubmit={handleSave} className="bg-white p-10 rounded-[3rem] border-2 border-red-50 shadow-2xl space-y-8 animate-in slide-in-from-top-4">
           <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
             <div className="col-span-full"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Titolo Problema</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold" required value={form.title || ''} onChange={e => setForm({...form, title: e.target.value})} /></div>
             <SearchableSelect label="Cliente *" placeholder="Associa cliente..." options={portfolios} selectedId={form.portfolioId} onSelect={(id:string) => setForm({...form, portfolioId: id})} displayFn={(p:any)=>p.businessName} searchFn={(p:any)=>p.businessName} subTextFn={(p:any)=>p.taxId} icon={<Users size={14}/>} />
             <SearchableSelect label="PDP (Opzionale)" placeholder="Cerca PDP..." options={availablePdps} selectedId={form.pdpId} onSelect={(id:string) => setForm({...form, pdpId: id})} displayFn={(p:any)=>p.pdpCode} searchFn={(p:any)=>p.pdpCode} subTextFn={(p:any)=>p.address} icon={<Zap size={14}/>} />
             <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Stato Caso</label>
               <select className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 font-bold outline-none" value={form.status} onChange={e => setForm({...form, status: e.target.value as CaseStatus})}>
                  <option value="nuovo">Nuovo</option>
                  <option value="in lavorazione">In Lavorazione</option>
                  <option value="risolto">Risolto</option>
                  <option value="KO">KO</option>
                  <option value="partner">Partner</option>
               </select>
             </div>
             <div className="col-span-full"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Note / Dettagli</label><textarea className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-medium" rows={4} value={form.notes || ''} onChange={e => setForm({...form, notes: e.target.value})} /></div>
            {editingId && <TimestampDetail data={form} />}
           </div>
           <div className="flex justify-end gap-4 border-t pt-8"><button type="button" onClick={() => setIsFormOpen(false)} className="font-bold text-slate-400">Annulla</button><button type="submit" className="bg-emerald-900 text-white px-10 py-4 rounded-2xl font-black">{editingId ? 'Aggiorna' : 'Apri'} Pratica</button></div>
        </form>
      )}
      <div className="space-y-4">{data.map((c: Caso) => {
        const port = portfolios.find(p => p.id === c.portfolioId);
        return (
        <div key={c.id} className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm hover:shadow-xl transition-all flex items-center justify-between group">
           <div className="flex items-center gap-6 flex-1">
             <div className={`p-4 rounded-2xl ${c.status === 'nuovo' ? 'bg-blue-50 text-blue-500' : c.status === 'KO' ? 'bg-red-50 text-red-500' : 'bg-emerald-50 text-emerald-500'}`}><AlertCircle/></div>
             <div className="overflow-hidden">
               <p className="text-[10px] font-black uppercase text-slate-400 truncate">{port?.businessName || 'Senza Cliente'}</p>
               <h4 className="text-lg font-black tracking-tight text-slate-800 truncate">{c.title}</h4>
               {c.notes && <p className="text-[10px] text-slate-400 truncate italic">"{c.notes}"</p>}
             </div>
           </div>
           <div className="flex items-center gap-8 shrink-0">
             <div className="text-right">
               <span className={`px-4 py-1 rounded-full text-[10px] font-black uppercase tracking-widest ${
                 c.status === 'nuovo' ? 'bg-blue-100 text-blue-700' : 
                 c.status === 'KO' ? 'bg-red-100 text-red-700' : 
                 'bg-emerald-100 text-emerald-700'
               }`}>{c.status}</span>
             </div>
             <div className="space-x-1 opacity-0 group-hover:opacity-100 transition-opacity flex items-center">
                <button onClick={() => onOpenChat(c.id)} className="p-3 text-slate-300 hover:text-blue-500 relative">
                  <MessageSquare size={16}/>
                  {getUnreadCount(c.id) > 0 && (
                    <span className="absolute top-0 right-0 w-2 h-2 bg-red-500 rounded-full border-2 border-white"></span>
                  )}
                </button>
                <button onClick={() => { setForm(c); setEditingId(c.id); setIsFormOpen(true); }} className="p-3 text-slate-300 hover:text-emerald-500"><Edit2 size={16}/></button>
                <button onClick={() => setData(data.filter(i=>i.id !== c.id))} className="p-3 text-slate-300 hover:text-red-400"><X size={18}/></button>
             </div>
           </div>
        </div>)})}
      </div>
    </div>
  );
};

const UsersView: React.FC<any> = ({ users, setUsers }) => {
  const [isAdding, setIsAdding] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [form, setForm] = useState<Partial<User>>({ role: 'utente' });

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    if (editingId) {
        setUsers((prev: User[]) => prev.map(u => u.id === editingId ? { ...u, ...form } as User : u));
    } else {
        setUsers((prev: User[]) => [...prev, { ...form, id: Math.random().toString(36).substr(2, 9) } as User]);
    }
    setIsAdding(false);
    setEditingId(null);
    setForm({ role: 'utente' });
  };

  const startEdit = (u: User) => {
    setForm(u);
    setEditingId(u.id);
    setIsAdding(true);
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-300">
       <div className="flex justify-between items-center"><div><h2 className="text-3xl font-black text-emerald-900 tracking-tighter">Gestione Profili</h2></div><button onClick={() => { setIsAdding(true); setEditingId(null); setForm({role:'utente'}); }} className="bg-emerald-700 text-white px-6 py-3 rounded-2xl font-black flex items-center gap-2 shadow-lg"><Plus/> Crea Profilo</button></div>
       {isAdding && (
         <form onSubmit={handleSave} className="bg-white p-8 rounded-[2rem] border shadow-2xl space-y-6 max-w-2xl animate-in slide-in-from-top-4">
           <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="space-y-1"><label className="text-xs font-bold text-slate-500 uppercase">Nome</label><input required className="w-full border p-3 rounded-xl bg-slate-50" value={form.name || ''} onChange={e => setForm({...form, name: e.target.value})} /></div>
              <div className="space-y-1"><label className="text-xs font-bold text-slate-500 uppercase">Codice Utente</label><input required className="w-full border p-3 rounded-xl bg-slate-50 font-black text-emerald-600" value={form.userCode || ''} onChange={e => setForm({...form, userCode: e.target.value.toUpperCase()})} /></div>
              <div className="space-y-1"><label className="text-xs font-bold text-slate-500 uppercase">Username</label><input required className="w-full border p-3 rounded-xl bg-slate-50 font-bold" value={form.username || ''} onChange={e => setForm({...form, username: e.target.value})} /></div>
              <div className="space-y-1"><label className="text-xs font-bold text-slate-500 uppercase">Password</label><input required type="password" className="w-full border p-3 rounded-xl bg-slate-50" value={form.password || ''} onChange={e => setForm({...form, password: e.target.value})} /></div>
              <div className="space-y-1"><label className="text-xs font-bold text-slate-500 uppercase">Grado</label><select className="w-full border p-3 rounded-xl bg-slate-50 font-bold outline-none" value={form.role} onChange={e => setForm({...form, role: e.target.value as any})}><option value="admin">Admin</option><option value="backoffice">Backoffice</option><option value="utente">Utente</option></select></div>
              <div className="space-y-1"><label className="text-xs font-bold text-slate-500 uppercase">Email</label><input type="email" className="w-full border p-3 rounded-xl bg-slate-50" value={form.email || ''} onChange={e => setForm({...form, email: e.target.value})} /></div>
              <div className="space-y-1"><label className="text-xs font-bold text-slate-500 uppercase">Cellulare</label><input type="tel" className="w-full border p-3 rounded-xl bg-slate-50" value={form.phone || ''} onChange={e => setForm({...form, phone: e.target.value})} /></div>
            </div>
           <div className="flex justify-end gap-3 pt-4 border-t"><button type="button" onClick={() => { setIsAdding(false); setEditingId(null); }} className="font-bold text-slate-400">Annulla</button><button type="submit" className="bg-emerald-600 text-white px-8 py-3 rounded-xl font-black">{editingId ? 'Aggiorna' : 'Crea'} Utente</button></div>
         </form>
       )}
       <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">{users.map((u: User) => (
         <div key={u.id} className={`p-6 rounded-3xl border shadow-sm flex items-center gap-4 group transition-all ${u.role === 'admin' ? 'bg-red-50 border-red-100' : u.role === 'backoffice' ? 'bg-yellow-50 border-yellow-100' : 'bg-white border-slate-100'}`}>
            <div className={`w-12 h-12 rounded-2xl flex items-center justify-center font-black text-lg ${u.role === 'admin' ? 'bg-red-100 text-red-600' : u.role === 'backoffice' ? 'bg-yellow-100 text-yellow-600' : 'bg-emerald-50 text-emerald-600'}`}>{u.name.charAt(0)}</div>
            <div className="flex-1 overflow-hidden">
              <p className="text-sm text-slate-800 font-medium truncate">{u.name}</p>
              <p className="text-[10px] font-black uppercase text-slate-400 tracking-wider">{u.userCode}</p>
            </div>
            <div className="flex gap-1 opacity-0 group-hover:opacity-100 transition-all">
                <button onClick={() => startEdit(u)} className="text-slate-400 hover:text-emerald-500 p-1"><Edit2 size={16}/></button>
                <button onClick={() => setUsers(users.filter(x=>x.id !== u.id))} className="text-slate-400 hover:text-red-500 p-1"><Trash2 size={16}/></button>
            </div>
         </div>))}
       </div>
    </div>
  );
};

const StatisticsView: React.FC<{ contracts: Contract[] }> = ({ contracts }) => {
  const chartData = useMemo(() => {
    const months: { [key: string]: { month: string, power: number, metano: number } } = {};
    
    contracts.forEach(c => {
      const dateStr = c.creationDate || c.createdAt;
      if (!dateStr) return;
      
      let d: Date;
      if (dateStr.includes('/')) {
        const [day, month, year] = dateStr.split('/').map(Number);
        d = new Date(year, month - 1, day);
      } else {
        d = new Date(dateStr);
      }
      
      const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
      if (!months[key]) {
        months[key] = { 
          month: d.toLocaleString('it-IT', { month: 'short', year: '2-digit' }), 
          power: 0, 
          metano: 0 
        };
      }
      
      if (c.service === 'POWER') months[key].power += c.volume;
      else months[key].metano += c.volume;
    });
    
    return Object.keys(months).sort().map(key => months[key]);
  }, [contracts]);

  return (
    <div className="space-y-8 animate-in fade-in duration-500">
      <header>
        <h2 className="text-4xl font-black tracking-tighter text-emerald-900">Statistiche</h2>
        <p className="text-slate-500 font-medium italic">Andamento mensile volumi caricati (Data Creazione)</p>
      </header>
      
      <div className="bg-white p-8 rounded-[2.5rem] border border-slate-100 shadow-sm h-[500px]">
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={chartData} margin={{ top: 20, right: 30, left: 20, bottom: 5 }}>
            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
            <XAxis dataKey="month" axisLine={false} tickLine={false} tick={{ fontSize: 12, fontWeight: 700, fill: '#64748b' }} />
            <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 12, fontWeight: 700, fill: '#64748b' }} />
            <Tooltip 
              contentStyle={{ borderRadius: '1rem', border: 'none', boxShadow: '0 20px 25px -5px rgb(0 0 0 / 0.1)' }}
              cursor={{ fill: '#f8fafc' }}
            />
            <Legend verticalAlign="top" align="right" iconType="circle" wrapperStyle={{ paddingBottom: '20px' }} />
            <Bar dataKey="power" name="Power (KWh)" fill="#991b1b" radius={[4, 4, 0, 0]} />
            <Bar dataKey="metano" name="Metano (smc)" fill="#2563eb" radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
};

const PortfolioModal: React.FC<{ 
  isOpen: boolean, 
  onClose: () => void, 
  portfolio: Portfolio,
  setData: React.Dispatch<React.SetStateAction<Portfolio[]>>,
  user: User,
  users: User[]
}> = ({ isOpen, onClose, portfolio, setData, user, users }) => {
  const [isEditing, setIsEditing] = useState(false);
  const [form, setForm] = useState<Partial<Portfolio>>(portfolio);
  const canEdit = user.role === 'admin' || user.role === 'backoffice';
  const ENTITY_TYPES: PortfolioEntityType[] = ['Domestico', 'Impresa', 'Associazione', 'Condominio', 'PA'];

  useEffect(() => {
    if (isOpen) {
      setForm(portfolio);
      setIsEditing(false);
    }
  }, [isOpen, portfolio]);

  if (!isOpen) return null;

  const handleAddressSync = (val: string) => {
    const s = splitAddress(val);
    if (s) setForm(prev => ({ ...prev, legalAddress: val, ...s }));
    else setForm(prev => ({ ...prev, legalAddress: val }));
  };

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    const ts = getTimestamp();
    setData((prev: Portfolio[]) => prev.map(p => p.id === portfolio.id ? { ...p, ...form, updatedAt: ts, updatedBy: user.name } : p));
    setIsEditing(false);
  };

  return (
    <div className="fixed inset-0 bg-emerald-950/40 backdrop-blur-sm z-[100] flex items-center justify-center p-4 animate-in fade-in duration-200">
      <div className="bg-white w-full max-w-2xl rounded-[2.5rem] shadow-2xl overflow-hidden animate-in zoom-in-95 duration-200">
        <div className="p-8 border-b border-slate-100 flex justify-between items-center bg-emerald-50/50">
          <div>
            <h3 className="text-2xl font-black text-emerald-900 tracking-tight">{isEditing ? 'Modifica Anagrafica' : portfolio.businessName}</h3>
            <p className="text-xs font-bold text-emerald-600 uppercase tracking-widest mt-1">Scheda Portafoglio</p>
          </div>
          <div className="flex items-center gap-2">
            {canEdit && !isEditing && (
              <button onClick={() => setIsEditing(true)} className="p-3 bg-white rounded-2xl text-emerald-600 hover:text-emerald-700 shadow-sm transition-all"><Edit2 size={20}/></button>
            )}
            <button onClick={onClose} className="p-3 hover:bg-white rounded-2xl transition-all text-slate-400 hover:text-slate-600 shadow-sm"><X size={24}/></button>
          </div>
        </div>
        <div className="p-8 space-y-6 max-h-[70vh] overflow-y-auto">
          {isEditing ? (
            <form id="modal-portfolio-form" onSubmit={handleSave} className="space-y-6">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1">Tipologia Cliente</label>
                  <select 
                    className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500"
                    value={form.entityType || 'Domestico'}
                    onChange={e => setForm({...form, entityType: e.target.value as PortfolioEntityType})}
                  >
                    {ENTITY_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
                  </select>
                </div>
                <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Ragione Sociale</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" required value={form.businessName || ''} onChange={e => setForm({...form, businessName: e.target.value})} /></div>
              <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Cellulare</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" value={form.phone || ''} onChange={e => setForm({...form, phone: e.target.value})} /></div>
              <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">E-mail</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" value={form.email || ''} onChange={e => setForm({...form, email: e.target.value})} /></div>
              <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">PEC</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" value={form.pec || ''} onChange={e => setForm({...form, pec: e.target.value})} /></div>

              {form.entityType !== 'Domestico' && (
                <>
                  <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Nome Rapp. Legale</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" value={form.legalRepresentativeName || ''} onChange={e => setForm({...form, legalRepresentativeName: e.target.value})} /></div>
                  <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">CF Rapp. Legale</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-mono font-bold focus:border-emerald-500" maxLength={16} value={form.legalRepresentativeTaxId || ''} onChange={e => setForm({...form, legalRepresentativeTaxId: e.target.value.toUpperCase()})} /></div>
                </>
              )}
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1 flex justify-between">
                    <span>CF o P.IVA</span>
                    <span className="text-[9px] text-slate-400">{form.entityType === 'Domestico' ? '16 caratteri' : '11 caratteri'}</span>
                  </label>
                  <input 
                    className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-mono font-bold focus:border-emerald-500" 
                    required 
                    maxLength={form.entityType === 'Domestico' ? 16 : 11}
                    value={form.taxId || ''} 
                    onChange={e => setForm({...form, taxId: e.target.value.toUpperCase()})} 
                  />
                </div>
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1">Email Cliente</label>
                  <input 
                    type="email"
                    className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" 
                    value={form.email || ''} 
                    onChange={e => setForm({...form, email: e.target.value})} 
                  />
                </div>
                <div className="col-span-full">
                  <SearchableSelect label="Assegna a Utente" placeholder="Cerca utente..." options={users} selectedId={form.assignedTo} onSelect={(id:string) => setForm({...form, assignedTo: id})} displayFn={(u:any)=>u.name} searchFn={(u:any)=>u.name} subTextFn={(u:any)=>u.role} icon={<UserIcon size={14}/>} />
                </div>
                <div className="col-span-full space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1 flex items-center gap-2"><MapPin size={12}/> Indirizzo</label>
                  <input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none" value={form.legalAddress || ''} onChange={e => handleAddressSync(e.target.value)} />
                </div>
                <div className="grid grid-cols-4 gap-3 col-span-full bg-slate-50 p-4 rounded-xl border border-slate-100">
                  <div><label className="text-[10px] font-black uppercase text-slate-400">Via</label><input className="w-full bg-white p-2 rounded-lg text-xs font-bold border" value={form.street || ''} onChange={e => setForm({...form, street: e.target.value})} /></div>
                  <div><label className="text-[10px] font-black uppercase text-slate-400">CAP</label><input className="w-full bg-white p-2 rounded-lg text-xs font-bold border" value={form.cap || ''} onChange={e => setForm({...form, cap: e.target.value})} /></div>
                  <div><label className="text-[10px] font-black uppercase text-slate-400">Comune</label><input className="w-full bg-white p-2 rounded-lg text-xs font-bold border" value={form.city || ''} onChange={e => setForm({...form, city: e.target.value})} /></div>
                  <div><label className="text-[10px] font-black uppercase text-slate-400">PR</label><input className="w-full bg-white p-2 rounded-lg text-xs font-bold border uppercase" value={form.province || ''} onChange={e => setForm({...form, province: e.target.value.toUpperCase()})} /></div>
                </div>
                <div className="col-span-full"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Allegati</label>
                  <FileUploader attachments={form.attachments || []} canEdit={true} onUpload={files => setForm({...form, attachments: [...(form.attachments || []), ...files]})} onRemove={id => setForm({...form, attachments: form.attachments?.filter(a => a.id !== id)})} />
                </div>
              </div>
            </form>
          ) : (
            <div className="grid grid-cols-2 gap-6">
              <div className="space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">Tipologia</p>
                <p className="font-bold text-emerald-600">{portfolio.entityType || '---'}</p>
              </div>
              <div className="space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">CF / P.IVA</p>
                <p className="font-mono font-bold text-slate-700">{portfolio.taxId}</p>
              </div>
              <div className="space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">Email</p>
                <p className="font-bold text-slate-700">{portfolio.email || '---'}</p>
              </div>
              <div className="space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">Cellulare</p>
                <p className="font-bold text-slate-700">{portfolio.phone || '---'}</p>
              </div>
              <div className="space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">PEC</p>
                <p className="font-bold text-slate-700">{portfolio.pec || '---'}</p>
              </div>
              {portfolio.entityType !== 'Domestico' && (
                <>
                  <div className="space-y-1">
                    <p className="text-[10px] font-black uppercase text-slate-400">Nome Rapp. Legale</p>
                    <p className="font-bold text-slate-700">{portfolio.legalRepresentativeName || '---'}</p>
                  </div>
                  <div className="space-y-1">
                    <p className="text-[10px] font-black uppercase text-slate-400">CF Rapp. Legale</p>
                    <p className="font-mono font-bold text-slate-700">{portfolio.legalRepresentativeTaxId || '---'}</p>
                  </div>
                </>
              )}
              <div className="col-span-full space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">Indirizzo Legale</p>
                <div className="flex items-center gap-2">
                  <p className="font-bold text-slate-700">{portfolio.legalAddress}</p>
                  <AddressLink address={portfolio.legalAddress} showIconOnly className="text-blue-500" />
                </div>
              </div>
              {portfolio.updatedBy && (
                <div className="col-span-full p-3 bg-emerald-50 rounded-xl border border-emerald-100 flex items-center gap-2">
                  <Clock size={14} className="text-emerald-600" />
                  <p className="text-[10px] font-black text-emerald-900 uppercase tracking-widest">Ultima modifica effettuata da: <span className="text-emerald-600">{portfolio.updatedBy}</span></p>
                </div>
              )}
              {portfolio.attachments && portfolio.attachments.length > 0 && (
                <div className="col-span-full space-y-3">
                  <p className="text-[10px] font-black uppercase text-slate-400">Documenti Allegati</p>
                  <div className="grid grid-cols-1 gap-2">
                    {portfolio.attachments.map(a => (
                      <div key={a.id} className="flex items-center justify-between p-3 bg-slate-50 rounded-xl border border-slate-100">
                        <div className="flex items-center gap-3">
                          <Paperclip size={14} className="text-slate-400" />
                          <span className="text-xs font-bold text-slate-600">{a.fileName}</span>
                        </div>
                        <span className="text-[10px] font-bold text-slate-400">{(a.fileSize / 1024).toFixed(1)} KB</span>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
        <div className="p-8 bg-slate-50 border-t border-slate-100 flex justify-end gap-3">
          {isEditing ? (
            <>
              <button onClick={() => setIsEditing(false)} className="font-bold text-slate-400 px-6">Annulla</button>
              <button form="modal-portfolio-form" type="submit" className="bg-emerald-900 text-white px-8 py-3 rounded-2xl font-black shadow-lg hover:scale-105 transition-all">Salva Modifiche</button>
            </>
          ) : (
            <button onClick={onClose} className="bg-emerald-900 text-white px-8 py-3 rounded-2xl font-black shadow-lg hover:scale-105 transition-all">Chiudi</button>
          )}
        </div>
      </div>
    </div>
  );
};

const PDPModal: React.FC<{ 
  isOpen: boolean, 
  onClose: () => void, 
  pdp: PDP,
  setData: React.Dispatch<React.SetStateAction<PDP[]>>,
  user: User
}> = ({ isOpen, onClose, pdp, setData, user }) => {
  const [isEditing, setIsEditing] = useState(false);
  const [form, setForm] = useState<Partial<PDP>>(pdp);
  const canEdit = user.role === 'admin' || user.role === 'backoffice';

  useEffect(() => {
    if (isOpen) {
      setForm(pdp);
      setIsEditing(false);
    }
  }, [isOpen, pdp]);

  if (!isOpen) return null;

  const handleAddressSync = (val: string) => {
    const s = splitAddress(val);
    if (s) setForm(prev => ({ ...prev, address: val, ...s }));
    else setForm(prev => ({ ...prev, address: val }));
  };

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    const ts = getTimestamp();
    setData((prev: PDP[]) => prev.map(p => p.id === pdp.id ? { ...p, ...form, updatedAt: ts } : p));
    setIsEditing(false);
  };

  return (
    <div className="fixed inset-0 bg-slate-950/40 backdrop-blur-sm z-[100] flex items-center justify-center p-4 animate-in fade-in duration-200">
      <div className="bg-white w-full max-w-2xl rounded-[2.5rem] shadow-2xl overflow-hidden animate-in zoom-in-95 duration-200">
        <div className="p-8 border-b border-slate-100 flex justify-between items-center bg-yellow-50/50">
          <div>
            <h3 className="text-2xl font-black text-slate-900 tracking-tight font-mono">{isEditing ? 'Modifica PDP' : pdp.pdpCode}</h3>
            <p className="text-xs font-bold text-yellow-600 uppercase tracking-widest mt-1">Scheda Tecnica PDP</p>
          </div>
          <div className="flex items-center gap-2">
            {canEdit && !isEditing && (
              <button onClick={() => setIsEditing(true)} className="p-3 bg-white rounded-2xl text-yellow-600 hover:text-yellow-700 shadow-sm transition-all"><Edit2 size={20}/></button>
            )}
            <button onClick={onClose} className="p-3 hover:bg-white rounded-2xl transition-all text-slate-400 hover:text-slate-600 shadow-sm"><X size={24}/></button>
          </div>
        </div>
        <div className="p-8 space-y-6 max-h-[70vh] overflow-y-auto">
          {isEditing ? (
            <form id="modal-pdp-form" onSubmit={handleSave} className="space-y-6">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Codice PDP</label><input maxLength={14} minLength={14} pattern=".{14}" title="Il codice PDP deve essere di 14 caratteri" className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-mono font-black uppercase" required value={form.pdpCode || ''} onChange={e => setForm({...form, pdpCode: e.target.value.toUpperCase()})} /></div>
                <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Indirizzo Fornitura</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none" value={form.address || ''} onChange={e => handleAddressSync(e.target.value)} /></div>
                <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Potenza Impegnata (kW)</label><input type="number" step="0.1" className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold" value={form.potenzaImpegnata || ''} onChange={e => setForm({...form, potenzaImpegnata: Number(e.target.value)})} /></div>
                <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Potenza Disponibile (kW)</label><input type="number" step="0.1" className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold" value={form.potenzaDisponibile || ''} onChange={e => setForm({...form, potenzaDisponibile: Number(e.target.value)})} /></div>
                <div className="col-span-full space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1 text-emerald-600">Cabina Primaria</label><input className="w-full border-2 p-3.5 rounded-2xl bg-white outline-none border-emerald-100" value={form.technicalSpecs || ''} onChange={e => setForm({...form, technicalSpecs: e.target.value})} /></div>
              </div>
            </form>
          ) : (
            <div className="grid grid-cols-2 gap-6">
              <div className="col-span-full space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">Indirizzo Fornitura</p>
                <div className="flex items-center gap-2">
                  <p className="font-bold text-slate-700">{pdp.address}</p>
                  <AddressLink address={pdp.address} className="text-blue-500" />
                </div>
              </div>
              <div className="space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">Potenza Impegnata</p>
                <p className="font-bold text-slate-700">{pdp.potenzaImpegnata} kW</p>
              </div>
              <div className="space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">Potenza Disponibile</p>
                <p className="font-bold text-slate-700">{pdp.potenzaDisponibile} kW</p>
              </div>
              <div className="col-span-full space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">Cabina Primaria</p>
                <p className="font-bold text-slate-700">{pdp.technicalSpecs || '---'}</p>
              </div>
            </div>
          )}
        </div>
        <div className="p-8 bg-slate-50 border-t border-slate-100 flex justify-end gap-3">
          {isEditing ? (
            <>
              <button onClick={() => setIsEditing(false)} className="font-bold text-slate-400 px-6">Annulla</button>
              <button form="modal-pdp-form" type="submit" className="bg-slate-900 text-white px-8 py-3 rounded-2xl font-black shadow-lg hover:scale-105 transition-all">Salva Modifiche</button>
            </>
          ) : (
            <button onClick={onClose} className="bg-slate-900 text-white px-8 py-3 rounded-2xl font-black shadow-lg hover:scale-105 transition-all">Chiudi</button>
          )}
        </div>
      </div>
    </div>
  );
};

const CreditCheckModal: React.FC<{ isOpen: boolean, onClose: () => void, portfolio: Portfolio, onSave: (req: Partial<CreditCheckRequest>) => void }> = ({ isOpen, onClose, portfolio, onSave }) => {
  const [power, setPower] = useState('');
  const [methane, setMethane] = useState('');
  const [selectedSuppliers, setSelectedSuppliers] = useState<Record<string, boolean>>({
    A2A1: false, AXPO: false, DOLO: false, SORG: false, OPEN: false
  });

  if (!isOpen) return null;

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const suppliers: any = {};
    Object.entries(selectedSuppliers).forEach(([key, val]) => {
      if (val) suppliers[key] = 'richiesto';
    });
    onSave({
      portfolioId: portfolio.id,
      powerVolume: Number(power),
      methaneVolume: Number(methane),
      suppliers
    });
    onClose();
  };

  return (
    <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-300">
      <div className="bg-white w-full max-w-lg rounded-[2.5rem] shadow-2xl overflow-hidden animate-in zoom-in-95 duration-300">
        <div className="p-8 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
          <div>
            <h3 className="text-2xl font-black tracking-tight text-slate-800">Richiesta Credit Check</h3>
            <p className="text-slate-500 text-sm font-medium">{portfolio.businessName}</p>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-slate-100 rounded-full transition-colors"><X size={24}/></button>
        </div>
        <form onSubmit={handleSubmit} className="p-8 space-y-6">
          <div className="grid grid-cols-2 gap-4">
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1">Vol Power (kWh)</label>
              <input type="number" required className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" value={power} onChange={e => setPower(e.target.value)} />
            </div>
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1">Vol Metano (smc)</label>
              <input type="number" required className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" value={methane} onChange={e => setMethane(e.target.value)} />
            </div>
          </div>
          <div className="space-y-3">
            <label className="text-xs font-bold text-slate-500 uppercase ml-1">Seleziona Fornitori</label>
            <div className="grid grid-cols-2 gap-3">
              {['A2A1', 'AXPO', 'DOLO', 'SORG', 'OPEN'].map(s => (
                <label key={s} className={`flex items-center gap-3 p-4 rounded-2xl border-2 cursor-pointer transition-all ${selectedSuppliers[s] ? 'border-emerald-500 bg-emerald-50' : 'border-slate-100 hover:border-slate-200'}`}>
                  <input type="checkbox" className="hidden" checked={selectedSuppliers[s]} onChange={() => setSelectedSuppliers(prev => ({ ...prev, [s]: !prev[s] }))} />
                  <div className={`w-5 h-5 rounded-md border-2 flex items-center justify-center transition-all ${selectedSuppliers[s] ? 'bg-emerald-500 border-emerald-500' : 'border-slate-300'}`}>
                    {selectedSuppliers[s] && <Check size={12} className="text-white" />}
                  </div>
                  <span className="font-black text-sm text-slate-700">{s}</span>
                </label>
              ))}
            </div>
          </div>
          <button type="submit" className="w-full py-4 bg-emerald-900 text-white rounded-2xl font-black text-lg shadow-xl hover:bg-black transition-all mt-4">
            Invia Richiesta
          </button>
        </form>
      </div>
    </div>
  );
};

const CreditCheckView: React.FC<{ data: CreditCheckRequest[], setData: React.Dispatch<React.SetStateAction<CreditCheckRequest[]>>, portfolios: Portfolio[], users: User[], user: User, exportFn: any }> = ({ data, setData, portfolios, users, user, exportFn }) => {
  const isAdminOrBO = user.role === 'admin' || user.role === 'backoffice';

  const handleStatusChange = (id: string, supplier: keyof CreditCheckRequest['suppliers'], status: CreditCheckStatus) => {
    setData(prev => prev.map(s => s.id === id ? { ...s, suppliers: { ...s.suppliers, [supplier]: status }, updatedAt: getTimestamp() } : s));
  };

  const getStatusColor = (status: CreditCheckStatus) => {
    switch (status) {
      case 'OK': return 'bg-emerald-100 text-emerald-700 border-emerald-200';
      case 'SDD': return 'bg-pink-100 text-pink-700 border-pink-200';
      case 'KO':
      case 'non affidabile': return 'bg-red-100 text-red-700 border-red-200';
      case 'cauzione': return 'bg-yellow-100 text-yellow-700 border-yellow-200';
      case 'richiesto': return 'bg-blue-100 text-blue-700 border-blue-200';
      default: return 'bg-slate-100 text-slate-500 border-slate-200';
    }
  };

  const statusOptions: CreditCheckStatus[] = ['inserito', 'OK', 'SDD', 'KO', 'cauzione', 'da affidare', 'non affidabile'];

  return (
    <div className="space-y-8 animate-in fade-in duration-500">
      <header className="flex justify-between items-center">
        <div>
          <h2 className="text-4xl font-black tracking-tighter text-emerald-900">Credit Check</h2>
          <p className="text-slate-500 font-medium italic">Gestione pareri fornitori</p>
        </div>
        <button onClick={() => exportFn(data, 'CreditCheck')} className="p-3 text-slate-400 hover:text-slate-800 transition-colors">
          <FileSpreadsheet />
        </button>
      </header>

      <div className="bg-white rounded-[2.5rem] border border-slate-100 shadow-sm overflow-hidden overflow-x-auto">
        <table className="w-full text-left min-w-[1000px]">
          <thead className="bg-slate-50 border-b">
            <tr>
              <th className="p-4 text-[10px] font-black uppercase text-slate-400">Cliente</th>
              <th className="p-4 text-[10px] font-black uppercase text-slate-400">Vol Power (kWh)</th>
              <th className="p-4 text-[10px] font-black uppercase text-slate-400">Vol Metano (smc)</th>
              <th className="p-4 text-[10px] font-black uppercase text-slate-400">A2A1</th>
              <th className="p-4 text-[10px] font-black uppercase text-slate-400">AXPO</th>
              <th className="p-4 text-[10px] font-black uppercase text-slate-400">DOLO</th>
              <th className="p-4 text-[10px] font-black uppercase text-slate-400">SORG</th>
              <th className="p-4 text-[10px] font-black uppercase text-slate-400">OPEN</th>
              {isAdminOrBO && <th className="p-4 text-[10px] font-black uppercase text-slate-400">Agente</th>}
              <th className="p-4"></th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-50">
            {data.map((s) => {
              const portfolio = portfolios.find(p => p.id === s.portfolioId);
              const agent = users.find(u => u.id === s.assignedTo);
              return (
                <tr key={s.id} className="hover:bg-slate-50/50 transition-colors">
                  <td className="p-4 font-bold text-slate-800 text-xs">{portfolio?.businessName || 'N/D'}</td>
                  <td className="p-4 font-mono text-xs">{s.powerVolume.toLocaleString()}</td>
                  <td className="p-4 font-mono text-xs">{s.methaneVolume.toLocaleString()}</td>
                  {(['A2A1', 'AXPO', 'DOLO', 'SORG', 'OPEN'] as const).map(supplier => (
                    <td key={supplier} className="p-4">
                      {isAdminOrBO ? (
                        <select 
                          value={s.suppliers[supplier] || ''} 
                          onChange={(e) => handleStatusChange(s.id, supplier, e.target.value as CreditCheckStatus)}
                          className={`text-[9px] font-black uppercase p-1.5 rounded-lg border outline-none transition-all ${getStatusColor(s.suppliers[supplier] as CreditCheckStatus)}`}
                        >
                          <option value="">-</option>
                          <option value="richiesto">Richiesto</option>
                          {statusOptions.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                        </select>
                      ) : (
                        <span className={`text-[9px] font-black uppercase px-2.5 py-1 rounded-full border ${getStatusColor(s.suppliers[supplier] as CreditCheckStatus)}`}>
                          {s.suppliers[supplier] || '-'}
                        </span>
                      )}
                    </td>
                  ))}
                  {isAdminOrBO && (
                    <td className="p-4">
                      <p className="text-[11px] font-bold text-slate-600">{agent?.name}</p>
                      <p className="text-[9px] font-black text-emerald-500">{agent?.userCode}</p>
                    </td>
                  )}
                  <td className="p-4 text-right">
                    <button onClick={() => setData(prev => prev.filter(item => item.id !== s.id))} className="text-slate-300 hover:text-red-500 transition-colors">
                      <Trash2 size={16} />
                    </button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
};

const MapView: React.FC<{ pdps: PDP[], portfolios: Portfolio[], contracts: Contract[] }> = ({ pdps, portfolios, contracts }) => {
  const [coords, setCoords] = useState<Record<string, [number, number]>>({});

  if (pdps.length === 0) {
    return (
      <div className="p-12 text-center bg-white rounded-[2.5rem] border border-slate-100 shadow-sm space-y-4">
        <div className="w-16 h-16 bg-slate-50 rounded-full flex items-center justify-center mx-auto">
          <MapPin className="text-slate-300" size={32} />
        </div>
        <div>
          <h3 className="text-lg font-black text-slate-800">Nessun PDP trovato</h3>
          <p className="text-slate-500 text-sm">Non ci sono punti di prelievo da visualizzare sulla mappa.</p>
        </div>
      </div>
    );
  }

  // Geocoding con delay per rispettare policy Nominatim (1 req/sec)
  useEffect(() => {
    const uniqueAddresses: string[] = Array.from(new Set(pdps.map(p => p.address).filter(Boolean)));
    
    const geocode = async () => {
      for (let i = 0; i < uniqueAddresses.length; i++) {
        const addr = uniqueAddresses[i];
        if (coords[addr]) continue;

        try {
          const res = await fetch(`https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(addr + ', Italy')}`, {
            headers: { 'Accept-Language': 'it' }
          });
          const data = await res.json() as any;
          if (data.length > 0) {
            const lat = parseFloat(data[0].lat);
            const lon = parseFloat(data[0].lon);
            if (!isNaN(lat) && !isNaN(lon)) {
              setCoords(prev => ({ ...prev, [addr]: [lat, lon] as [number, number] }));
            }
          }
          // Delay di 1 secondo tra le richieste
          await new Promise(resolve => setTimeout(resolve, 1000));
        } catch (e) {
          console.error("Geocoding error", e);
        }
      }
    };

    geocode();
  }, [pdps]);

  const markers = useMemo(() => {
    return pdps.map(p => {
      const pos = coords[p.address];
      if (!pos || isNaN(pos[0]) || isNaN(pos[1])) return null;

      const lastContract = [...contracts].filter(c => c.pdpId === p.id).sort((a,b) => new Date(b.startDate).getTime() - new Date(a.startDate).getTime())[0];
      const portfolio = lastContract ? portfolios.find(pf => pf.id === lastContract.portfolioId) : null;

      // Piccolo offset per PDP allo stesso indirizzo
      const offsetPos: [number, number] = [pos[0] + (Math.random() - 0.5) * 0.001, pos[1] + (Math.random() - 0.5) * 0.001];

      const isGas = lastContract?.service === 'METANO';
      const customIcon = L.divIcon({
        html: isGas 
          ? `<div class="flex items-center justify-center w-8 h-8 bg-white border-2 border-blue-600 rounded-full shadow-md text-blue-600">
              <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="lucide lucide-flame"><path d="M8.5 14.5A2.5 2.5 0 0 0 11 12c0-1.38-.5-2-1-3-1.072-2.143-.224-4.054 2-6 .5 2.5 2 4.9 4 6.5 2 1.6 3 3.5 3 5.5a7 7 0 1 1-14 0c0-1.153.433-2.21 1.146-3.027a1.242 1.242 0 0 1 1.354-.373c.586.214 1.001.739 1 1.375.001.028.001.056.001.084Z"/></svg>
             </div>`
          : `<div class="flex items-center justify-center w-8 h-8 bg-white border-2 border-red-600 rounded-full shadow-md text-red-600">
              <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="lucide lucide-lightbulb"><path d="M15 14c.2-1 .7-1.7 1.5-2.5 1-.9 1.5-2.2 1.5-3.5A5 5 0 0 0 8 8c0 1.3.5 2.6 1.5 3.5.8.8 1.3 1.5 1.5 2.5"/><path d="M9 18h6"/><path d="M10 22h4"/></svg>
             </div>`,
        className: '',
        iconSize: [32, 32],
        iconAnchor: [16, 32],
      });

      return (
        <Marker key={p.id} position={offsetPos} icon={customIcon}>
          <Popup>
            <div className="p-2 space-y-1">
              <p className="text-[10px] font-black text-emerald-600 uppercase tracking-widest">{p.pdpCode}</p>
              <p className="font-bold text-slate-800">{portfolio?.businessName || 'N/D'}</p>
              <p className="text-[10px] text-slate-500">{p.address}</p>
              <div className="pt-2 border-t mt-2">
                <AddressLink address={p.address} label="Apri in Google Maps" className="text-blue-600 text-[10px] font-bold" />
              </div>
            </div>
          </Popup>
        </Marker>
      );
    }).filter(Boolean);
  }, [pdps, coords, contracts, portfolios]);

  // Componente per fittare la mappa ai marker
  const ChangeView = ({ markers }: { markers: any[] }) => {
    const map = useMap();
    useEffect(() => {
      if (markers.length > 0) {
        try {
          const validPositions = markers
            .map(m => m.props.position)
            .filter(pos => Array.isArray(pos) && !isNaN(pos[0]) && !isNaN(pos[1]));
          
          if (validPositions.length > 0) {
            const bounds = L.latLngBounds(validPositions);
            map.fitBounds(bounds, { padding: [50, 50] });
          }
        } catch (err) {
          console.error("Error fitting map bounds:", err);
        }
      }
    }, [markers, map]);
    return null;
  };

  return (
    <div className="space-y-8 flex flex-col">
      <header>
        <h2 className="text-4xl font-black tracking-tighter text-emerald-900">Mappa PDP</h2>
        <p className="text-slate-500 font-medium italic">Visualizzazione geografica dei punti di prelievo</p>
      </header>

      <div className="bg-white rounded-[2.5rem] border border-slate-100 shadow-sm overflow-hidden relative">
        <MapContainer 
          center={[41.9028, 12.4964]} 
          zoom={6} 
          scrollWheelZoom={true} 
          style={{ height: '600px', width: '100%' }}
        >
          <TileLayer
            attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
            url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
          />
          <ChangeView markers={markers} />
          {markers}
        </MapContainer>
        
        <div className="absolute bottom-6 left-6 z-[1000] bg-white/90 backdrop-blur p-4 rounded-2xl border border-slate-200 shadow-xl max-w-xs">
          <p className="text-[10px] font-black uppercase text-slate-400 mb-2">Legenda</p>
          <div className="space-y-2">
            <div className="flex items-center gap-3">
              <div className="w-6 h-6 bg-white border border-red-600 rounded-full flex items-center justify-center text-red-600 shadow-sm">
                <Lightbulb size={12} />
              </div>
              <p className="text-xs font-bold text-slate-700">Energia Elettrica</p>
            </div>
            <div className="flex items-center gap-3">
              <div className="w-6 h-6 bg-white border border-blue-600 rounded-full flex items-center justify-center text-blue-600 shadow-sm">
                <Flame size={12} />
              </div>
              <p className="text-xs font-bold text-slate-700">Gas Metano</p>
            </div>
          </div>
          <p className="text-[9px] text-slate-400 mt-3 italic">I punti sono geolocalizzati in base all'indirizzo fornito. Clicca per i dettagli.</p>
        </div>
      </div>
    </div>
  );
};

const BulkUploadView: React.FC<{ 
  setPortfolios: React.Dispatch<React.SetStateAction<Portfolio[]>>,
  setPdps: React.Dispatch<React.SetStateAction<PDP[]>>,
  setContracts: React.Dispatch<React.SetStateAction<Contract[]>>,
  setCabine: React.Dispatch<React.SetStateAction<CabinaPrimaria[]>>,
  setCasi: React.Dispatch<React.SetStateAction<Caso[]>>,
  setCreditChecks: React.Dispatch<React.SetStateAction<CreditCheckRequest[]>>
}> = ({ setPortfolios, setPdps, setContracts, setCabine, setCasi, setCreditChecks }) => {
  const [jsonInput, setJsonInput] = useState('');
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleUpload = () => {
    try {
      const data = JSON.parse(jsonInput);
      if (data.portfolios) {
        const taxIds = new Set();
        const duplicates = [];
        for (const p of data.portfolios) {
          if (p.taxId) {
            if (taxIds.has(p.taxId.toUpperCase())) {
              duplicates.push(p.taxId.toUpperCase());
            }
            taxIds.add(p.taxId.toUpperCase());
          }
        }
        if (duplicates.length > 0) {
          return setError(`Errore: Il JSON contiene C.F./P.IVA duplicati: ${duplicates.join(', ')}`);
        }
        setPortfolios(data.portfolios);
      }
      if (data.pdps) setPdps(data.pdps);
      if (data.contracts) setContracts(data.contracts);
      if (data.cabine) setCabine(data.cabine);
      if (data.casi) setCasi(data.casi);
      if (data.creditChecks) setCreditChecks(data.creditChecks);
      
      setSuccess('Database aggiornato con successo!');
      setError('');
      setJsonInput('');
    } catch (e) {
      setError('Errore nel formato JSON. Controlla la sintassi.');
      setSuccess('');
    }
  };

  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        
        const portfoliosSheet = wb.Sheets['Portafoglio'];
        const pdpsSheet = wb.Sheets['PDP'];
        const contractsSheet = wb.Sheets['Contratti'];

        if (portfoliosSheet) {
          const data = XLSX.utils.sheet_to_json(portfoliosSheet) as Portfolio[];
          const taxIds = new Set();
          const duplicates = [];
          for (const p of data) {
            if (p.taxId) {
              if (taxIds.has(p.taxId.toUpperCase())) {
                duplicates.push(p.taxId.toUpperCase());
              }
              taxIds.add(p.taxId.toUpperCase());
            }
          }
          if (duplicates.length > 0) {
            setError(`Errore nell'Excel: C.F./P.IVA duplicati: ${duplicates.join(', ')}`);
            return;
          }
          setPortfolios(data);
        }
        if (pdpsSheet) {
          const data = XLSX.utils.sheet_to_json(pdpsSheet);
          setPdps(data as PDP[]);
        }
        if (contractsSheet) {
          const data = XLSX.utils.sheet_to_json(contractsSheet);
          setContracts(data as Contract[]);
        }

        setSuccess('Database aggiornato con successo da Excel!');
        setError('');
        if (fileInputRef.current) fileInputRef.current.value = '';
      } catch (err) {
        setError('Errore durante la lettura del file Excel.');
        setSuccess('');
      }
    };
    reader.readAsBinaryString(file);
  };

  const downloadTemplate = () => {
    const template = {
      portfolios: [],
      pdps: [],
      contracts: [],
      cabine: [],
      casi: [],
      creditChecks: []
    };
    const blob = new Blob([JSON.stringify(template, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'template_charlie.json';
    a.click();
  };

  const downloadExcelTemplate = () => {
    const wb = XLSX.utils.book_new();
    
    const portfoliosWS = XLSX.utils.json_to_sheet([
      { businessName: 'Esempio Azienda', taxId: '01234567890', legalAddress: 'Via Roma 1, Milano', city: 'Milano', province: 'MI', cap: '20100', email: 'info@esempio.com', phone: '021234567' }
    ], { header: ['businessName', 'taxId', 'legalAddress', 'city', 'province', 'cap', 'email', 'phone'] });
    XLSX.utils.book_append_sheet(wb, portfoliosWS, 'Portafoglio');

    const pdpsWS = XLSX.utils.json_to_sheet([
      { pdpCode: 'IT001E12345678', address: 'Via Milano 10, Roma', potenzaImpegnata: 10, potenzaDisponibile: 11 }
    ], { header: ['pdpCode', 'address', 'potenzaImpegnata', 'potenzaDisponibile'] });
    XLSX.utils.book_append_sheet(wb, pdpsWS, 'PDP');

    const contractsWS = XLSX.utils.json_to_sheet([
      { contractNumber: 'CTR-2024-001', fornitore: 'A2A1', service: 'POWER', volume: 5000, status: 'attivo' }
    ], { header: ['contractNumber', 'fornitore', 'service', 'volume', 'status'] });
    XLSX.utils.book_append_sheet(wb, contractsWS, 'Contratti');

    XLSX.writeFile(wb, 'template_charlie.xlsx');
  };

  return (
    <div className="space-y-8 animate-in fade-in duration-500">
      <header>
        <h2 className="text-4xl font-black tracking-tighter text-emerald-900">Caricamento Massivo</h2>
        <p className="text-slate-500 font-medium italic">Aggiorna l'intero database tramite file Excel o JSON</p>
      </header>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        <div className="bg-white p-8 rounded-[2.5rem] border border-slate-100 shadow-sm space-y-6">
          <div className="flex items-center justify-between">
            <h3 className="font-black text-xl text-slate-800 tracking-tight">Caricamento Excel</h3>
            <button 
              onClick={downloadExcelTemplate}
              className="flex items-center gap-2 text-emerald-600 hover:text-emerald-700 font-bold text-sm"
            >
              <Download size={18} />
              Template Excel
            </button>
          </div>
          
          <div className="border-2 border-dashed border-slate-200 rounded-3xl p-12 text-center space-y-4 hover:border-emerald-500 transition-colors group relative">
            <input 
              type="file" 
              accept=".xlsx, .xls" 
              onChange={handleExcelUpload}
              ref={fileInputRef}
              className="absolute inset-0 opacity-0 cursor-pointer"
            />
            <div className="w-16 h-16 bg-emerald-50 rounded-full flex items-center justify-center mx-auto group-hover:scale-110 transition-transform">
              <FileSpreadsheet className="text-emerald-600" size={32} />
            </div>
            <div>
              <p className="font-black text-slate-800">Trascina qui il file Excel</p>
              <p className="text-slate-500 text-sm">o clicca per selezionarlo dal computer</p>
            </div>
          </div>
          <p className="text-[10px] text-slate-400 uppercase font-bold tracking-wider text-center">Supporta fogli: Portafoglio, PDP, Contratti</p>
        </div>

        <div className="bg-white p-8 rounded-[2.5rem] border border-slate-100 shadow-sm space-y-6">
          <div className="flex items-center justify-between">
            <h3 className="font-black text-xl text-slate-800 tracking-tight">Incolla JSON</h3>
            <button 
              onClick={downloadTemplate}
              className="flex items-center gap-2 text-emerald-600 hover:text-emerald-700 font-bold text-sm"
            >
              <Download size={18} />
              Template JSON
            </button>
          </div>

          <textarea
            value={jsonInput}
            onChange={(e) => setJsonInput(e.target.value)}
            placeholder='{ "portfolios": [...], "pdps": [...], ... }'
            className="w-full h-48 p-6 bg-slate-50 border border-slate-200 rounded-3xl font-mono text-xs focus:ring-2 focus:ring-emerald-500 focus:border-transparent outline-none resize-none"
          />

          <button
            onClick={handleUpload}
            disabled={!jsonInput}
            className="w-full py-4 bg-emerald-600 text-white rounded-2xl font-black text-lg shadow-lg shadow-emerald-100 hover:bg-emerald-700 transition-all disabled:opacity-50 disabled:cursor-not-allowed"
          >
            Aggiorna da JSON
          </button>
        </div>
      </div>

      {(error || success) && (
        <div className={`p-6 rounded-[2rem] border flex items-center gap-4 animate-in slide-in-from-bottom-4 duration-300 ${error ? 'bg-red-50 border-red-100 text-red-600' : 'bg-emerald-50 border-emerald-100 text-emerald-600'}`}>
          {error ? <AlertCircle size={24} /> : <Check size={24} />}
          <p className="font-bold">{error || success}</p>
        </div>
      )}
    </div>
  );
};

const ProfileView: React.FC<{ user: User, onUpdate: (u: User) => void }> = ({ user, onUpdate }) => {
  const [form, setForm] = useState<User>(user);
  const [success, setSuccess] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleAvatarUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      setForm({ ...form, avatar: evt.target?.result as string });
    };
    reader.readAsDataURL(file);
  };

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    onUpdate(form);
    setSuccess('Profilo aggiornato con successo!');
    setTimeout(() => setSuccess(''), 3000);
  };

  return (
    <div className="max-w-4xl mx-auto space-y-8 animate-in fade-in duration-500">
      <header>
        <h2 className="text-4xl font-black tracking-tighter text-emerald-900">Il Mio Profilo</h2>
        <p className="text-slate-500 font-medium italic">Gestisci le tue informazioni e le preferenze di notifica</p>
      </header>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
        <div className="md:col-span-1 space-y-6">
          <div className="bg-white p-8 rounded-[2.5rem] border border-slate-100 shadow-sm text-center space-y-4">
            <div className="relative w-32 h-32 mx-auto group">
              <div className="w-full h-full rounded-full overflow-hidden border-4 border-emerald-50 bg-slate-100 flex items-center justify-center">
                {form.avatar ? (
                  <img src={form.avatar} alt="Avatar" className="w-full h-full object-cover" referrerPolicy="no-referrer" />
                ) : (
                  <UserIcon size={48} className="text-slate-300" />
                )}
              </div>
              <button 
                onClick={() => fileInputRef.current?.click()}
                className="absolute bottom-0 right-0 p-2 bg-emerald-600 text-white rounded-full shadow-lg hover:scale-110 transition-all"
              >
                <Edit2 size={16} />
              </button>
              <input type="file" ref={fileInputRef} className="hidden" accept="image/*" onChange={handleAvatarUpload} />
            </div>
            <div>
              <h3 className="font-black text-xl text-slate-800 tracking-tight">{form.name}</h3>
              <p className="text-[10px] font-black uppercase text-emerald-600 tracking-widest">{form.role} • {form.userCode}</p>
            </div>
          </div>

          <div className="bg-emerald-900 p-8 rounded-[2.5rem] text-white space-y-4 shadow-xl shadow-emerald-900/20">
            <h4 className="font-black text-sm uppercase tracking-widest opacity-60">Info Account</h4>
            <div className="space-y-3">
              <div className="flex items-center gap-3">
                <Mail size={16} className="text-emerald-400" />
                <span className="text-xs font-bold truncate">{form.email}</span>
              </div>
              {form.phone && (
                <div className="flex items-center gap-3">
                  <Phone size={16} className="text-emerald-400" />
                  <span className="text-xs font-bold truncate">{form.phone}</span>
                </div>
              )}
              <div className="flex items-center gap-3">
                <Hash size={16} className="text-emerald-400" />
                <span className="text-xs font-bold">{form.userCode}</span>
              </div>
            </div>
          </div>
        </div>

        <div className="md:col-span-2 space-y-6">
          <form onSubmit={handleSave} className="bg-white p-8 rounded-[2.5rem] border border-slate-100 shadow-sm space-y-8">
            <div className="space-y-6">
              <h3 className="font-black text-xl text-slate-800 tracking-tight flex items-center gap-2">
                <Edit2 size={20} className="text-emerald-600" /> Dati Personali
              </h3>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1">Nome Completo</label>
                  <input 
                    className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none focus:border-emerald-500 font-bold" 
                    value={form.name} 
                    onChange={e => setForm({...form, name: e.target.value})} 
                  />
                </div>
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1">Email</label>
                  <input 
                    type="email"
                    className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none focus:border-emerald-500 font-bold" 
                    value={form.email} 
                    onChange={e => setForm({...form, email: e.target.value})} 
                  />
                </div>
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1">Cellulare</label>
                  <input 
                    type="tel"
                    className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none focus:border-emerald-500 font-bold" 
                    value={form.phone || ''} 
                    onChange={e => setForm({...form, phone: e.target.value})} 
                  />
                </div>
              </div>
            </div>

            <div className="space-y-6 pt-4 border-t border-slate-50">
              <h3 className="font-black text-xl text-slate-800 tracking-tight flex items-center gap-2">
                <Bell size={20} className="text-emerald-600" /> Impostazioni Notifiche
              </h3>
              
              <div className="space-y-4">
                <div className="bg-slate-50 p-6 rounded-3xl space-y-4">
                  <h4 className="text-[10px] font-black uppercase text-slate-400 tracking-widest flex items-center gap-2">
                    <Mail size={14} /> Notifiche Email
                  </h4>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {[
                      { key: 'contractStatusChange', label: 'Cambi Stato Contratti' },
                      { key: 'caseStatusChange', label: 'Cambi Stato Casi' },
                      { key: 'newMessages', label: 'Nuovi Messaggi in Chat' },
                      { key: 'mentions', label: 'Menzioni (@codice)' }
                    ].map(item => (
                      <label key={item.key} className="flex items-center justify-between p-3 bg-white rounded-xl border border-slate-100 cursor-pointer hover:border-emerald-200 transition-all">
                        <span className="text-xs font-bold text-slate-600">{item.label}</span>
                        <input 
                          type="checkbox" 
                          className="w-4 h-4 rounded border-slate-300 text-emerald-600 focus:ring-emerald-500"
                          checked={(form.notificationSettings?.email as any)?.[item.key]}
                          onChange={e => setForm({
                            ...form, 
                            notificationSettings: {
                              ...form.notificationSettings!,
                              email: {
                                ...form.notificationSettings!.email,
                                [item.key]: e.target.checked
                              }
                            }
                          })}
                        />
                      </label>
                    ))}
                  </div>
                </div>

                <div className="bg-slate-50 p-6 rounded-3xl space-y-4">
                  <div className="flex items-center justify-between">
                    <h4 className="text-[10px] font-black uppercase text-slate-400 tracking-widest flex items-center gap-2">
                      <MessageSquare size={14} /> Google Chat (Webhook)
                    </h4>
                    <input 
                      type="checkbox" 
                      className="w-4 h-4 rounded border-slate-300 text-emerald-600 focus:ring-emerald-500"
                      checked={form.notificationSettings?.googleChat.enabled}
                      onChange={e => setForm({
                        ...form,
                        notificationSettings: {
                          ...form.notificationSettings!,
                          googleChat: {
                            ...form.notificationSettings!.googleChat,
                            enabled: e.target.checked
                          }
                        }
                      })}
                    />
                  </div>
                  {form.notificationSettings?.googleChat.enabled && (
                    <div className="space-y-1.5 animate-in slide-in-from-top-2 duration-200">
                      <label className="text-[10px] font-black uppercase text-slate-400 ml-1">Webhook URL</label>
                      <input 
                        placeholder="https://chat.googleapis.com/v1/spaces/..."
                        className="w-full border-2 p-3.5 rounded-2xl bg-white outline-none focus:border-emerald-500 font-mono text-[10px]" 
                        value={form.notificationSettings?.googleChat.webhookUrl || ''} 
                        onChange={e => setForm({
                          ...form,
                          notificationSettings: {
                            ...form.notificationSettings!,
                            googleChat: {
                              ...form.notificationSettings!.googleChat,
                              webhookUrl: e.target.value
                            }
                          }
                        })} 
                      />
                    </div>
                  )}
                </div>
              </div>
            </div>

            <div className="flex items-center justify-between pt-4">
              {success && (
                <p className="text-emerald-600 font-bold text-sm flex items-center gap-2 animate-in slide-in-from-left-2">
                  <Check size={18} /> {success}
                </p>
              )}
              <button 
                type="submit"
                className="ml-auto bg-emerald-900 text-white px-10 py-4 rounded-2xl font-black text-lg shadow-xl hover:scale-105 transition-all"
              >
                Salva Profilo
              </button>
            </div>
          </form>
        </div>
      </div>
    </div>
  );
};

const LoginScreen: React.FC<{ users: User[], onLoginSuccess: (u: User) => void }> = ({ users, onLoginSuccess }) => {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');

  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    const user = users.find(u => u.username === username && u.password === password);
    if (user) onLoginSuccess(user);
    else setError('Credenziali non valide.');
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-emerald-950 p-6 relative overflow-hidden">
      <div className="absolute top-0 left-0 w-full h-full opacity-30"><div className="absolute -top-32 -left-32 w-96 h-96 bg-emerald-600 rounded-full blur-[120px]"></div></div>
      <div className="w-full max-w-md relative z-10 animate-in fade-in zoom-in duration-500"><div className="bg-white rounded-[2.5rem] shadow-2xl p-10 space-y-8">
           <div className="text-center space-y-2"><p className="text-[10px] font-black uppercase tracking-[0.3em] text-emerald-500">PROGETTO ANTICRISI</p><h2 className="text-5xl font-black tracking-tighter text-emerald-950">1335</h2></div>
           <form onSubmit={handleLogin} className="space-y-6">
             <div className="space-y-1"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Username</label><input required className="w-full border-2 p-4 rounded-2xl bg-slate-50 outline-none focus:border-emerald-500 font-bold" onChange={e => setUsername(e.target.value)} /></div>
             <div className="space-y-1"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Password</label><input required type="password" className="w-full border-2 p-4 rounded-2xl bg-slate-50 outline-none focus:border-emerald-500 font-bold" onChange={e => setPassword(e.target.value)} /></div>
             {error && <p className="bg-red-50 text-red-600 p-4 rounded-2xl text-xs font-bold">{error}</p>}
             <button type="submit" className="w-full bg-emerald-900 text-white p-5 rounded-2xl font-black text-xl shadow-xl hover:bg-black transition-all">Accedi</button>
           </form>
        </div></div>
    </div>
  );
};

export default App;
