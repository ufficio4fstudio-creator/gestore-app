
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import React, { useState, useEffect, useMemo, useRef } from 'react';
import { 
  Users, Zap, FileText, Plus, Search, MapPin, Phone, Mail, ChevronRight,
  TrendingUp, Files, ShieldAlert, Activity, LogOut, Lock, User as UserIcon,
  ChevronDown, X, Download, HardDrive, AlertCircle, FileSpreadsheet, Paperclip,
  Edit2, Trash2, Info, Clock, Check, ExternalLink, Hash, BarChart3, Upload, Menu, Flame, MessageSquare, Send, Bell, BellOff,
  Wallet, FileSignature, Newspaper, Lightbulb, Eye, CheckCircle2, AlertTriangle, Building2, Megaphone, Tag, Handshake, Sparkles, Layers
} from 'lucide-react';
import { io, Socket } from 'socket.io-client';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer 
} from 'recharts';
import { MapContainer, TileLayer, Marker, Popup, useMap } from 'react-leaflet';
import L from 'leaflet';
import 'leaflet/dist/leaflet.css';
import { 
  Portfolio, PDP, Contract, AppView, User, UserRole, Partner, Appartenenza, Campagna,
  CabinaPrimaria, Caso, CaseCategory, CaseSector, Attachment, ContractStatus, Fornitore, CaseStatus,
  CreditCheckRequest, CreditCheckStatus, ContractType, PortfolioEntityType, Message, News, Associazione, AssociazioneStatus,
  ContrattoTelefonico, TelefonicoTipoOperazione, TelefonicoGestore
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
const TENSIONE_OPTIONS = [
  { value: '220', label: '220' },
  { value: '220 trifase', label: '220 trifase' },
  { value: '380', label: '380' },
  { value: '1.000', label: '1.000' },
  { value: '5.000', label: '5.000' },
  { value: '10.000', label: '10.000' },
  { value: '15.000', label: '15.000' },
  { value: '20.000', label: '20.000' }
];

const formatTensione = (val?: string | number): string => {
  if (!val) return '';
  const s = String(val).trim();
  if (s === '1000' || s === '1.000') return '1.000';
  if (s === '5000' || s === '5.000') return '5.000';
  if (s === '10000' || s === '10.000') return '10.000';
  if (s === '15000' || s === '15.000') return '15.000';
  if (s === '20000' || s === '20.000') return '20.000';
  return s;
};

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

const composeAddress = (street?: string, cap?: string, city?: string, province?: string): string => {
  const cleanStreet = (street || '').trim();
  const cleanCap = (cap || '').trim();
  const cleanCity = (city || '').trim();
  const cleanProv = (province || '').trim().toUpperCase();

  const cityZip = [cleanCap, cleanCity].filter(Boolean).join(' ');
  const parts: string[] = [];
  if (cleanStreet) parts.push(cleanStreet);
  if (cityZip) parts.push(cityZip);
  if (cleanProv) parts.push(cleanProv);
  
  return parts.join(', ');
};

const splitAddress = (val: string): { street: string; cap: string; city: string; province: string } | null => {
  if (!val || typeof val !== 'string') return null;
  const raw = val.trim();
  if (!raw) return null;

  let street = '';
  let cap = '';
  let city = '';
  let province = '';

  let working = raw;

  // 1. Estrai Provincia (es. "(MI)" o fine stringa dopo virgola/spazio es. ", MI" o " MI")
  const provParenMatch = working.match(/\(([A-Za-z]{2})\)/);
  if (provParenMatch) {
    province = provParenMatch[1].toUpperCase();
    working = working.replace(provParenMatch[0], ' ').trim();
  } else {
    const provEndMatch = working.match(/[, \t]+([A-Za-z]{2})\s*$/);
    if (provEndMatch) {
      province = provEndMatch[1].toUpperCase();
      working = working.substring(0, provEndMatch.index).trim();
    }
  }

  // 2. Estrai CAP (5 cifre)
  const capMatch = working.match(/\b(\d{5})\b/);
  if (capMatch) {
    cap = capMatch[1];
    const capIndex = capMatch.index!;
    const beforeCap = working.substring(0, capIndex).trim().replace(/,\s*$/, '');
    const afterCap = working.substring(capIndex + 5).trim().replace(/^,\s*/, '');
    
    if (beforeCap) street = beforeCap;
    if (afterCap) city = afterCap.replace(/^[,\s-]+|[,\s-]+$/g, '');
  } else {
    // Senza CAP: dividi per virgole se presenti
    const commaParts = working.split(',').map(s => s.trim()).filter(Boolean);
    if (commaParts.length >= 2) {
      street = commaParts[0];
      city = commaParts.slice(1).join(', ');
    } else {
      street = working;
    }
  }

  street = street.replace(/^[,\s]+|[,\s]+$/g, '');
  city = city.replace(/^[,\s]+|[,\s]+$/g, '');

  return { street, cap, city, province };
};

interface AddressInputGroupProps {
  label?: string;
  address?: string;
  street?: string;
  cap?: string;
  city?: string;
  province?: string;
  onChange: (fields: { address: string; street: string; cap: string; city: string; province: string }) => void;
  accentColor?: 'emerald' | 'yellow' | 'slate';
  bgBox?: string;
  placeholder?: string;
  extraHeaderAction?: React.ReactNode;
}

const AddressInputGroup: React.FC<AddressInputGroupProps> = ({
  label = 'Indirizzo (Via, CAP Città, PROV)',
  address = '',
  street = '',
  cap = '',
  city = '',
  province = '',
  onChange,
  accentColor = 'emerald',
  bgBox = 'bg-slate-50',
  placeholder = 'Es: Via Roma 1, 20121 Milano, MI',
  extraHeaderAction
}) => {
  // Quando l'utente scrive nel campo unico:
  const handleSingleAddressChange = (val: string) => {
    const s = splitAddress(val);
    onChange({
      address: val,
      street: s ? s.street : val,
      cap: s ? s.cap : (cap || ''),
      city: s ? s.city : (city || ''),
      province: s ? s.province : (province || '')
    });
  };

  // Quando l'utente scrive nei 4 campi distinti (via, cap, comune, PROV):
  const handleDistinctChange = (field: 'street' | 'cap' | 'city' | 'province', val: string) => {
    const updated = {
      street: field === 'street' ? val : (street || ''),
      cap: field === 'cap' ? val : (cap || ''),
      city: field === 'city' ? val : (city || ''),
      province: field === 'province' ? val.toUpperCase() : (province || '').toUpperCase()
    };
    const composed = composeAddress(updated.street, updated.cap, updated.city, updated.province);
    onChange({
      address: composed,
      ...updated
    });
  };

  const focusBorder = accentColor === 'yellow' ? 'focus:border-yellow-500' : 'focus:border-emerald-500';

  return (
    <div className="col-span-full space-y-2">
      <div className="flex justify-between items-center">
        <label className="text-xs font-bold text-slate-500 uppercase ml-1 flex items-center gap-2">
          <MapPin size={12}/> {label}
          {address && (
            <AddressLink address={address} label="(Vedi su Maps)" className="text-[10px] font-black uppercase text-blue-500 underline ml-1" />
          )}
        </label>
        {extraHeaderAction}
      </div>

      {/* Campo Singolo Unificato */}
      <input 
        className={`w-full border-2 p-3.5 rounded-2xl ${bgBox} outline-none font-bold text-sm ${focusBorder} transition-all`} 
        placeholder={placeholder}
        value={address} 
        onChange={e => handleSingleAddressChange(e.target.value)} 
      />

      {/* 4 Campi Distinti Confluenti (Via, CAP, Comune, PROV) */}
      <div className="grid grid-cols-12 gap-3 bg-slate-50/90 p-3.5 rounded-2xl border border-slate-100 shadow-inner">
        <div className="col-span-12 sm:col-span-5">
          <label className="text-[10px] font-black uppercase text-slate-400 block mb-1">Via</label>
          <input 
            className={`w-full bg-white p-2.5 rounded-xl text-xs font-bold border border-slate-200 outline-none ${focusBorder}`} 
            placeholder="Es: Via Roma 1"
            value={street} 
            onChange={e => handleDistinctChange('street', e.target.value)} 
          />
        </div>
        <div className="col-span-4 sm:col-span-2">
          <label className="text-[10px] font-black uppercase text-slate-400 block mb-1">CAP</label>
          <input 
            className={`w-full bg-white p-2.5 rounded-xl text-xs font-bold border border-slate-200 outline-none font-mono ${focusBorder}`} 
            placeholder="20121"
            maxLength={5}
            value={cap} 
            onChange={e => handleDistinctChange('cap', e.target.value)} 
          />
        </div>
        <div className="col-span-5 sm:col-span-3">
          <label className="text-[10px] font-black uppercase text-slate-400 block mb-1">Comune</label>
          <input 
            className={`w-full bg-white p-2.5 rounded-xl text-xs font-bold border border-slate-200 outline-none ${focusBorder}`} 
            placeholder="Milano"
            value={city} 
            onChange={e => handleDistinctChange('city', e.target.value)} 
          />
        </div>
        <div className="col-span-3 sm:col-span-2">
          <label className="text-[10px] font-black uppercase text-slate-400 block mb-1 flex items-center justify-between">
            <span>PROV</span>
            <span className="text-[8px] text-slate-300 font-normal">2 lett.</span>
          </label>
          <input 
            className={`w-full bg-white p-2.5 rounded-xl text-xs font-black uppercase border border-slate-200 outline-none font-mono text-center ${focusBorder}`} 
            placeholder="MI"
            maxLength={2}
            value={province} 
            onChange={e => handleDistinctChange('province', e.target.value.toUpperCase())} 
          />
        </div>
      </div>
    </div>
  );
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

interface ConfirmModalProps {
  isOpen: boolean;
  title: string;
  message: string;
  onConfirm: () => void;
  onCancel: () => void;
}

const ConfirmModal: React.FC<ConfirmModalProps> = ({ isOpen, title, message, onConfirm, onCancel }) => {
  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[10000] flex items-center justify-center p-4 animate-in fade-in duration-200">
      <div className="bg-white w-full max-w-md rounded-[2.5rem] shadow-2xl overflow-hidden p-8 space-y-6 animate-in zoom-in-95 duration-200">
        <div className="text-center space-y-2">
          <h3 className="text-xl font-black text-slate-900">{title}</h3>
          <p className="text-sm font-medium text-slate-500">{message}</p>
        </div>
        <div className="flex gap-4">
          <button
            type="button"
            onClick={onCancel}
            className="flex-1 py-3.5 px-6 rounded-2xl border-2 border-slate-100 font-bold text-slate-500 hover:bg-slate-50 transition-all text-xs outline-none"
          >
            Annulla
          </button>
          <button
            type="button"
            onClick={onConfirm}
            className="flex-1 py-3.5 px-6 rounded-2xl bg-red-500 hover:bg-red-600 text-white font-black shadow-lg hover:shadow-red-200 transition-all text-xs outline-none"
          >
            Conferma
          </button>
        </div>
      </div>
    </div>
  );
};

const App: React.FC = () => {
  const [currentUser, setCurrentUser] = useState<User | null>(() => {
    const stored = localStorage.getItem('charlie_current_user');
    return stored ? JSON.parse(stored) : null;
  });

  const [view, setView] = useState<AppView>('dashboard');
  const [users, setUsers] = useState<User[]>(() => {
    const data = loadData('users', INITIAL_USERS);
    const hasAdmin = data.some(u => u.role === 'admin' || u.username.toLowerCase() === 'admin');
    const normalized = hasAdmin ? data : [INITIAL_USERS[0], ...data];
    return normalized.map(u => 
      (u.role === 'admin' || u.username.toLowerCase() === 'admin')
        ? { ...u, username: 'admin', password: 'admin2025', role: 'admin' }
        : u
    );
  });
  const [portfolios, setPortfolios] = useState<Portfolio[]>(() => loadData('portfolios', []));
  const [pdps, setPdps] = useState<PDP[]>(() => loadData('pdps', []));
  const [contracts, setContracts] = useState<Contract[]>(() => loadData('contracts', []));
  const [cabine, setCabine] = useState<CabinaPrimaria[]>(() => loadData('cabine', []));
  const [casi, setCasi] = useState<Caso[]>(() => loadData('casi', []));
  const [creditChecks, setCreditChecks] = useState<CreditCheckRequest[]>(() => loadData('creditChecks', []));
  const [news, setNews] = useState<News[]>(() => loadData('news', []));
  const [associazioni, setAssociazioni] = useState<Associazione[]>(() => loadData('associazioni', []));
  const [contrattiTelefonici, setContrattiTelefonici] = useState<ContrattoTelefonico[]>(() => loadData('contrattiTelefonici', []));
  const [partners, setPartners] = useState<Partner[]>(() => loadData('partners', []));
  const [appartenenze, setAppartenenze] = useState<Appartenenza[]>(() => loadData('appartenenze', []));
  const [campagne, setCampagne] = useState<Campagna[]>(() => loadData('campagne', []));
  const [selectedNews, setSelectedNews] = useState<News | null>(null);
  const [selectedPdpDetail, setSelectedPdpDetail] = useState<PDP | null>(null);
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
    localStorage.setItem('contrattiTelefonici', JSON.stringify(contrattiTelefonici));
    localStorage.setItem('partners', JSON.stringify(partners));
    localStorage.setItem('appartenenze', JSON.stringify(appartenenze));
    localStorage.setItem('campagne', JSON.stringify(campagne));
  }, [users, portfolios, pdps, contracts, cabine, casi, creditChecks, news, associazioni, contrattiTelefonici, partners, appartenenze, campagne]);

  const filteredData = useMemo(() => {
    const isAdminOrBO = currentUser?.role === 'admin' || currentUser?.role === 'backoffice';
    const userPortfolios = isAdminOrBO
      ? portfolios
      : portfolios.filter(p => p.assignedTo === currentUser?.id || p.partnerId === currentUser?.id);
    const myPortfolioIds = new Set(userPortfolios.map(p => p.id));
    const userContracts = isAdminOrBO
      ? contracts
      : contracts.filter(c => c.assignedTo === currentUser?.id || myPortfolioIds.has(c.portfolioId));
    const myPdpIds = new Set(userContracts.map(c => c.pdpId));
    return {
      portfolios: userPortfolios,
      pdps: isAdminOrBO ? pdps : pdps.filter(p => p.assignedTo === currentUser?.id || myPdpIds.has(p.id)),
      contracts: userContracts,
      contrattiTelefonici: isAdminOrBO ? contrattiTelefonici : contrattiTelefonici.filter(c => c.assignedTo === currentUser?.id || myPortfolioIds.has(c.portfolioId)),
      casi: isAdminOrBO ? casi : casi.filter(c => c.assignedTo === currentUser?.id || (c.portfolioId && myPortfolioIds.has(c.portfolioId))),
      creditChecks: isAdminOrBO ? creditChecks : creditChecks.filter(s => s.assignedTo === currentUser?.id || myPortfolioIds.has(s.portfolioId)),
      associazioni: isAdminOrBO ? associazioni : associazioni.filter(a => a.assignedTo === currentUser?.id || myPortfolioIds.has(a.portfolioId)),
      appartenenze: isAdminOrBO ? appartenenze : appartenenze.filter(a => a.assignedTo === currentUser?.id),
    };
  }, [currentUser, portfolios, pdps, contracts, casi, creditChecks, associazioni, contrattiTelefonici, appartenenze]);

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
    const cleanData = data.map(({ attachments, ...rest }) => {
      const row: any = { ...rest };
      const rawAddr = row.address || row.legalAddress;
      if (rawAddr) {
        const s = splitAddress(rawAddr);
        if (!row.street && s?.street) row.street = s.street;
        if (!row.cap && s?.cap) row.cap = s.cap;
        if (!row.city && s?.city) row.city = s.city;
        if (!row.province && s?.province) row.province = s.province;
      } else if (row.street || row.cap || row.city || row.province) {
        const composed = composeAddress(row.street, row.cap, row.city, row.province);
        if ('legalAddress' in row || fileName.toLowerCase().includes('portafoglio')) {
          row.legalAddress = composed;
        } else {
          row.address = composed;
        }
      }
      return row;
    });
    const ws = XLSX.utils.json_to_sheet(cleanData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, fileName);
    XLSX.writeFile(wb, `${fileName}_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  if (!currentUser) return <LoginScreen users={[...users, ...partners]} onLoginSuccess={handleLogin} />;

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

const AssociazioneView: React.FC<any> = ({ data, setData, portfolios, campagne = [], users = [], user, exportFn }) => {
  const [isAdding, setIsAdding] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  
  const getFirstOfMonth = () => {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-01`;
  };

  const [form, setForm] = useState<Partial<Associazione>>({
    startDate: getFirstOfMonth(),
    quota: 0,
    status: 'nuovo',
    isRenewal: false,
    campagnaId: ''
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
    const now = new Date().toISOString();

    if (editingId) {
      setData((prev: Associazione[]) => prev.map(a => a.id === editingId ? {
        ...a,
        ...form,
        campagnaId: form.campagnaId || undefined,
        updatedAt: now
      } as Associazione : a));
    } else {
      const newAssoc: Associazione = {
        ...form,
        id: Math.random().toString(36).substr(2, 9),
        assignedTo: user.id,
        status: 'nuovo', // Always nuovo on creation
        campagnaId: form.campagnaId || undefined,
        createdAt: now,
        updatedAt: now
      } as Associazione;

      setData((prev: Associazione[]) => [...prev, newAssoc]);
    }

    setIsAdding(false);
    setEditingId(null);
    setForm({
      startDate: getFirstOfMonth(),
      quota: 0,
      status: 'nuovo',
      isRenewal: false,
      campagnaId: ''
    });
  };

  const updateStatus = (id: string, newStatus: AssociazioneStatus) => {
    if (!isAdminOrBO) return;
    setData((prev: Associazione[]) => prev.map(a => a.id === id ? { ...a, status: newStatus, updatedAt: new Date().toISOString() } : a));
  };

  const [deleteConfirmId, setDeleteConfirmId] = useState<string | null>(null);

  const deleteAssoc = (id: string) => {
    setDeleteConfirmId(id);
  };

  const filteredAssoc = useMemo(() => {
    return data.filter((assoc: Associazione) => {
      const portfolio = portfolios.find((p: Portfolio) => p.id === assoc.portfolioId);
      const camp = campagne.find((c: Campagna) => c.id === assoc.campagnaId);
      const assignedUser = users.find((u: User) => u.id === (assoc.assignedTo || portfolio?.assignedTo));
      const searchStr = `${portfolio?.businessName || ''} ${portfolio?.taxId || ''} ${camp?.code || ''} ${camp?.name || ''} ${assignedUser?.userCode || ''} ${assignedUser?.name || ''}`.toLowerCase();
      return searchStr.includes(searchTerm.toLowerCase());
    });
  }, [data, portfolios, campagne, users, searchTerm]);

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
              placeholder="Cerca cliente, campagna, codice..." 
              value={searchTerm}
              onChange={e => setSearchTerm(e.target.value)}
            />
          </div>
          <button onClick={() => exportFn(data.map((a: any) => {
            const portfolio = portfolios.find((p: Portfolio) => p.id === a.portfolioId);
            const camp = campagne.find((c: Campagna) => c.id === a.campagnaId);
            const assignedUser = users.find((u: User) => u.id === (a.assignedTo || portfolio?.assignedTo));
            return { 
              ...a, 
              codiceUtente: assignedUser?.userCode || '',
              nomeUtente: assignedUser?.name || '',
              campagna: camp ? `${camp.code} - ${camp.name}` : ''
            };
          }), 'Associazioni')} className="bg-white text-slate-600 px-6 py-3 rounded-2xl font-black flex items-center gap-2 border border-slate-100 shadow-sm hover:bg-slate-50 transition-all text-sm">
            <Download size={18} /> Esporta
          </button>
          <button onClick={() => {
            setEditingId(null);
            setForm({
              startDate: getFirstOfMonth(),
              quota: 0,
              status: 'nuovo',
              isRenewal: false,
              campagnaId: ''
            });
            setIsAdding(true);
          }} className="bg-emerald-700 text-white px-6 py-3 rounded-2xl font-black flex items-center gap-2 shadow-lg shadow-emerald-900/20 hover:scale-105 transition-all text-sm">
            <Plus size={18} /> Nuova Associazione
          </button>
        </div>
      </header>

      {isAdding && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-200 overflow-y-auto">
          <form onSubmit={handleSave} className="bg-white w-full max-w-lg rounded-[2.5rem] shadow-2xl overflow-hidden my-8 animate-in zoom-in-95 duration-200">
            <div className="p-8 bg-emerald-900 text-white relative">
              <button type="button" onClick={() => { setIsAdding(false); setEditingId(null); }} className="absolute top-6 right-6 p-2 hover:bg-white/10 rounded-full transition-colors">
                <X size={24} />
              </button>
              <h3 className="text-2xl font-black tracking-tight">{editingId ? 'Modifica Associazione' : 'Nuova Associazione'}</h3>
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

              <SearchableSelect 
                label="Campagna" 
                placeholder="Seleziona campagna (opzionale)..." 
                options={campagne} 
                selectedId={form.campagnaId} 
                onSelect={(id:string) => setForm({...form, campagnaId: id})} 
                displayFn={(c: Campagna) => `${c.code} - ${c.name}`} 
                searchFn={(c: Campagna) => `${c.code} ${c.name} ${c.description}`} 
                subTextFn={(c: Campagna) => c.exclusiveBenefits || c.description} 
                icon={<Megaphone size={14}/>}
                allowClear={true}
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
                  className="w-5 h-5 rounded border-slate-300 text-emerald-600 focus:ring-emerald-500 cursor-pointer" 
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
          <table className="w-full text-left border-collapse min-w-[850px]">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-100">
                <th className="p-4 text-[10px] font-black uppercase text-slate-400 tracking-widest">Anagrafica</th>
                <th className="p-4 text-[10px] font-black uppercase text-slate-400 tracking-widest">Campagna</th>
                <th className="p-4 text-[10px] font-black uppercase text-slate-400 tracking-widest">Codice Utente</th>
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
                const camp = campagne.find((c: Campagna) => c.id === assoc.campagnaId);
                const assignedUser = users.find((u: User) => u.id === (assoc.assignedTo || portfolio?.assignedTo));
                const isExpired = new Date(assoc.endDate) < new Date();
                return (
                  <tr key={assoc.id} className="hover:bg-slate-50/50 transition-colors group">
                    <td className="p-4">
                      <p className="font-bold text-slate-700 text-xs">{portfolio?.businessName || 'N/A'}</p>
                      <p className="text-[9px] font-black text-slate-400 uppercase">{portfolio?.taxId}</p>
                    </td>
                    <td className="p-4">
                      {camp ? (
                        <div className="inline-flex items-center gap-1.5 bg-purple-50 text-purple-800 border border-purple-200 px-2.5 py-1 rounded-lg">
                          <span className="text-[10px] font-mono font-black">{camp.code}</span>
                          <span className="text-[10px] font-bold truncate max-w-[120px]" title={camp.name}>{camp.name}</span>
                        </div>
                      ) : (
                        <span className="text-xs text-slate-300 font-bold">---</span>
                      )}
                    </td>
                    <td className="p-4">
                      <span className="text-xs font-black text-emerald-700 bg-emerald-50 border border-emerald-200 px-2.5 py-1 rounded-lg font-mono">
                        {assignedUser?.userCode || '---'}
                      </span>
                      {assignedUser?.name && (
                        <p className="text-[9px] font-bold text-slate-400 mt-0.5 truncate max-w-[120px]">{assignedUser.name}</p>
                      )}
                    </td>
                    <td className="p-4">
                      <p className="text-[11px] font-bold text-slate-600">{formatDate(assoc.startDate)}</p>
                    </td>
                    <td className="p-4">
                      <p className="text-[10px] font-black text-emerald-700">€ {assoc.quota.toLocaleString('it-IT', { minimumFractionDigits: 2 })}</p>
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
                    <td className="p-4 text-right space-x-1">
                      <button onClick={() => {
                        setForm({
                          ...assoc,
                          campagnaId: assoc.campagnaId || ''
                        });
                        setEditingId(assoc.id);
                        setIsAdding(true);
                      }} className="p-1.5 text-slate-300 hover:text-emerald-600 transition-colors" title="Modifica">
                        <Edit2 size={16} />
                      </button>
                      <button onClick={() => deleteAssoc(assoc.id)} className="p-1.5 text-slate-300 hover:text-red-600 transition-colors" title="Elimina">
                        <Trash2 size={16} />
                      </button>
                    </td>
                  </tr>
                );
              })}
              {filteredAssoc.length === 0 && (
                <tr>
                  <td colSpan={8} className="p-20 text-center">
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
      <ConfirmModal
        isOpen={!!deleteConfirmId}
        title="Conferma Eliminazione"
        message="Sei sicuro di voler eliminare questa quota associativa? Questa operazione non può essere annullata."
        onConfirm={() => {
          setData((prev: Associazione[]) => prev.filter(a => a.id !== deleteConfirmId));
          setDeleteConfirmId(null);
        }}
        onCancel={() => setDeleteConfirmId(null)}
      />
    </div>
  );
};

const RinnoviView: React.FC<any> = ({ contracts, portfolios, pdps, exportFn, onViewPdp }) => {
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
                      <p onClick={() => pdp && onViewPdp && onViewPdp(pdp)} className="font-bold text-emerald-600 font-mono bg-emerald-50 px-2 py-0.5 rounded cursor-pointer hover:bg-emerald-100 transition-colors inline-block">{pdp?.pdpCode || 'N/A'}</p>
                      <p className="text-[10px] font-medium text-slate-400 mt-1">{pdp?.address}</p>
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
      case 'portfolio': return <PortfolioView data={filteredData.portfolios} users={users} partners={partners} appartenenze={filteredData.appartenenze} contracts={filteredData.contracts} contrattiTelefonici={filteredData.contrattiTelefonici} pdps={filteredData.pdps} setData={setPortfolios} user={currentUser} exportFn={exportToExcel} setCreditChecks={setCreditChecks} onCreateContract={(id: string) => { setPrefilledPortfolioId(id); setView('contracts'); }} />;
      case 'pdp': return <PDPView data={filteredData.pdps} setData={setPdps} cabine={cabine} setCabine={setCabine} contracts={contracts} portfolios={portfolios} users={users} user={currentUser} exportFn={exportToExcel} casi={filteredData.casi} onViewPdp={setSelectedPdpDetail} />;
      case 'contracts': return <ContractView data={filteredData.contracts} setData={setContracts} portfolios={portfolios} setPortfolios={setPortfolios} pdps={pdps} cabine={cabine} setCabine={setCabine} setPdps={setPdps} users={users} partners={partners} appartenenze={filteredData.appartenenze} campagne={campagne} user={currentUser} exportFn={exportToExcel} onOpenChat={(id: string) => { setActiveChat({ id, type: 'contract' }); markChatAsRead(id); }} getUnreadCount={getUnreadCount} prefilledPortfolioId={prefilledPortfolioId} onClearPrefill={() => setPrefilledPortfolioId(null)} onViewPdp={setSelectedPdpDetail} />;
      case 'telefonico': return <TelefonicoView data={filteredData.contrattiTelefonici} setData={setContrattiTelefonici} portfolios={portfolios} contracts={contracts} pdps={pdps} users={users} campagne={campagne} user={currentUser} exportFn={exportToExcel} />;
      case 'casi': return <CasiView data={filteredData.casi} setData={setCasi} portfolios={portfolios} pdps={pdps} contracts={contracts} users={[...users, ...partners]} user={currentUser} exportFn={exportToExcel} onOpenChat={(id: string) => { setActiveChat({ id, type: 'case' }); markChatAsRead(id); }} getUnreadCount={getUnreadCount} onViewPdp={setSelectedPdpDetail} />;
      case 'associazioni': return <AssociazioneView data={filteredData.associazioni} setData={setAssociazioni} portfolios={filteredData.portfolios} campagne={campagne} users={[...users, ...partners]} user={currentUser} exportFn={exportToExcel} />;
      case 'appartenenze': return <AppartenenzeView data={filteredData.appartenenze} setData={setAppartenenze} portfolios={filteredData.portfolios} setPortfolios={setPortfolios} users={users} partners={partners} contracts={filteredData.contracts} contrattiTelefonici={filteredData.contrattiTelefonici} pdps={filteredData.pdps} user={currentUser} exportFn={exportToExcel} />;
      case 'campagne': return <CampagneView data={campagne} setData={setCampagne} user={currentUser} exportFn={exportToExcel} />;
      case 'credit-check': return <CreditCheckView data={filteredData.creditChecks} setData={setCreditChecks} portfolios={portfolios} users={users} user={currentUser} exportFn={exportToExcel} />;
      case 'rinnovi': return <RinnoviView contracts={filteredData.contracts} portfolios={portfolios} pdps={pdps} exportFn={exportToExcel} onViewPdp={setSelectedPdpDetail} />;
      case 'map': return <MapView pdps={filteredData.pdps} portfolios={portfolios} contracts={contracts} onViewPdp={setSelectedPdpDetail} />;
      case 'statistics': return <StatisticsView contracts={filteredData.contracts} associazioni={filteredData.associazioni} />;
      case 'bulk-upload': return <BulkUploadView setPortfolios={setPortfolios} setPdps={setPdps} setContracts={setContracts} setCabine={setCabine} setCasi={setCasi} setCreditChecks={setCreditChecks} users={users} currentUser={currentUser} />;
      case 'cabine': return <CabineView data={cabine} pdps={pdps} contracts={contracts} exportFn={exportToExcel} />;
      case 'users': return <UsersView users={users} setUsers={setUsers} />;
      case 'partners': return currentUser.role === 'admin' ? <PartnersView partners={partners} setPartners={setPartners} /> : <Dashboard stats={filteredData} setView={setView} user={currentUser} news={news} onOpenNews={setSelectedNews} />;
      case 'news-admin': return <NewsAdminView news={news} setNews={setNews} />;
      case 'tpm-ranking': return <TPMRankingView stats={filteredData} users={[...users, ...partners]} />;
      case 'profile': return <ProfileView user={currentUser} onUpdate={u => { 
        setUsers(prev => prev.map(old => old.id === u.id ? u : old)); 
        setPartners(prev => prev.map(old => old.id === u.id ? ({ ...u, role: 'utente' } as Partner) : old));
        setCurrentUser(u); 
        localStorage.setItem('charlie_current_user', JSON.stringify(u)); 
      }} />;
      default: return <Dashboard stats={filteredData} setView={setView} user={currentUser} news={news} onOpenNews={setSelectedNews} />;
    }
  };

  return (
    <div className="flex flex-col md:flex-row min-h-screen bg-slate-50 text-slate-900">
      {/* Sidebar - Desktop */}
      <aside className="hidden md:flex w-64 bg-slate-900 text-white fixed h-full flex-col z-50">
        <div className="p-6 flex justify-between items-center">
          <div>
            <span className="text-[10px] uppercase tracking-widest text-emerald-400 font-black block">PROGETTO</span>
            <h1 className="text-2xl font-black tracking-tight text-white flex items-center gap-2">
              ANTICRISI
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
          <NavItem active={view === 'associazioni'} onClick={() => setView('associazioni')} icon={<Wallet size={18} className="text-emerald-700"/>} label="Associazioni" />
          <NavItem active={view === 'casi'} onClick={() => setView('casi')} icon={<AlertCircle size={18} className="text-red-500"/>} label="Casi" />
          <NavItem 
            active={view === 'contracts'} 
            onClick={() => setView('contracts')} 
            icon={
              <span className="flex items-center -space-x-0.5 text-amber-500">
                <Flame size={15} className="fill-amber-500/20" />
                <Lightbulb size={15} className="text-yellow-400 fill-yellow-400/20" />
              </span>
            } 
            label="Forniture" 
          />
          <NavItem active={view === 'telefonico'} onClick={() => setView('telefonico')} icon={<Phone size={18} className="text-sky-400"/>} label="Telefonico" />
          <NavItem active={view === 'credit-check'} onClick={() => setView('credit-check')} icon={<ShieldAlert size={18}/>} label="Credit Check" />
          <NavItem active={view === 'map'} onClick={() => setView('map')} icon={<MapPin size={18}/>} label="Mappa" />
          <NavItem active={view === 'pdp'} onClick={() => setView('pdp')} icon={<Zap size={18}/>} label="PDP" />
          <NavItem active={view === 'appartenenze'} onClick={() => setView('appartenenze')} icon={<Layers size={18} className="text-indigo-400"/>} label="Appartenenza" />
          <NavItem active={view === 'campagne'} onClick={() => setView('campagne')} icon={<Megaphone size={18} className="text-purple-400"/>} label="Campagne" />
          
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
              <NavItem active={view === 'partners'} onClick={() => setView('partners')} icon={<Handshake size={18} className="text-teal-400"/>} label="Partner" />
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
        <MobileNavItem active={view === 'associazioni'} onClick={() => setView('associazioni')} icon={<Wallet size={20} className="text-emerald-700"/>} />
        <MobileNavItem active={view === 'casi'} onClick={() => setView('casi')} icon={<AlertCircle size={20} className="text-red-500"/>} />
        <MobileNavItem 
          active={view === 'contracts'} 
          onClick={() => setView('contracts')} 
          icon={
            <span className="flex items-center -space-x-1">
              <Flame size={17} className="text-amber-500 fill-amber-500/20" />
              <Lightbulb size={17} className="text-yellow-400 fill-yellow-400/20" />
            </span>
          } 
        />
        <button 
          onClick={() => {
            const moreViews: AppView[] = currentUser.role === 'admin'
              ? ['telefonico', 'credit-check', 'map', 'pdp', 'appartenenze', 'campagne', 'statistics', 'cabine', 'rinnovi', 'bulk-upload', 'users', 'partners']
              : ['telefonico', 'credit-check', 'map', 'pdp', 'appartenenze', 'campagne', 'statistics', 'cabine', 'rinnovi'];
            const currentIndex = moreViews.indexOf(view);
            const nextIndex = (currentIndex + 1) % moreViews.length;
            setView(moreViews[nextIndex]);
          }}
          className={`p-3 rounded-2xl transition-all duration-200 ${['telefonico', 'credit-check', 'map', 'pdp', 'appartenenze', 'campagne', 'statistics', 'cabine', 'rinnovi', 'bulk-upload', 'users', 'partners'].includes(view) ? 'bg-emerald-50 text-emerald-600' : 'text-slate-400'}`}
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
      {selectedPdpDetail && (
        <PDPDetailModal
          isOpen={!!selectedPdpDetail}
          onClose={() => setSelectedPdpDetail(null)}
          pdp={selectedPdpDetail}
          contracts={contracts}
          portfolios={portfolios}
          casi={casi}
          cabine={cabine}
        />
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
  options, onSelect, label, icon, displayFn, searchFn, subTextFn, selectedId, placeholder, allowCustom, customLabel, disabled, maxLength, allowClear 
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
        <div className="flex items-center gap-2">
          {allowClear && selectedId && !disabled && (
            <button
              type="button"
              onClick={(e) => {
                e.stopPropagation();
                onSelect('');
                setIsOpen(false);
              }}
              className="p-1 text-slate-400 hover:text-red-500 rounded-full hover:bg-slate-200/60 transition-colors"
              title="Rimuovi selezione"
            >
              <X size={14} />
            </button>
          )}
          <ChevronDown size={16} className={isOpen ? 'rotate-180 transition-all' : ''} />
        </div>
      </div>
      {isOpen && (
        <div className="absolute z-[100] w-full mt-2 bg-white border border-slate-200 rounded-2xl shadow-2xl overflow-hidden animate-in fade-in zoom-in duration-200">
          <div className="p-2 border-b bg-slate-50"><input autoFocus maxLength={maxLength} className="w-full p-2 text-sm outline-none bg-transparent" placeholder="Digita per cercare..." value={query} onChange={e => setQuery(e.target.value)} /></div>
          <div className="max-h-60 overflow-y-auto">
            {allowClear && selectedId && (
              <div onClick={() => { onSelect(''); setIsOpen(false); setQuery(''); }} className="p-3 hover:bg-slate-50 cursor-pointer border-b border-slate-100 text-xs font-bold text-slate-400 italic">
                — Nessuno —
              </div>
            )}
            {allowCustom && query && query.length === 14 && !options.some((o: any) => displayFn(o).toLowerCase() === query.toLowerCase()) && (
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

const formatYYMMDD = (dateStr?: string | Date): string => {
  const d = dateStr ? new Date(dateStr) : new Date();
  if (isNaN(d.getTime())) {
    const now = new Date();
    const yy = String(now.getFullYear()).slice(-2);
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const dd = String(now.getDate()).padStart(2, '0');
    return `${yy}${mm}${dd}`;
  }
  if (typeof dateStr === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    const [yyyy, mm, dd] = dateStr.split('-');
    return `${yyyy.slice(-2)}${mm}${dd}`;
  }
  const yy = String(d.getFullYear()).slice(-2);
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${yy}${mm}${dd}`;
};

const readFileAsDataUrl = (file: File): Promise<string> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ''));
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
};

const loadImageDimensions = (dataUrl: string): Promise<{ width: number; height: number; normalizedDataUrl: string }> => {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      const maxDim = 1600;
      let w = img.width || 800;
      let h = img.height || 600;
      if (w > maxDim || h > maxDim) {
        if (w > h) {
          h = Math.round((h * maxDim) / w);
          w = maxDim;
        } else {
          w = Math.round((w * maxDim) / h);
          h = maxDim;
        }
      }
      canvas.width = w;
      canvas.height = h;
      const ctx = canvas.getContext('2d');
      if (ctx) {
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(0, 0, w, h);
        ctx.drawImage(img, 0, 0, w, h);
      }
      resolve({
        width: w,
        height: h,
        normalizedDataUrl: canvas.toDataURL('image/jpeg', 0.85)
      });
    };
    img.onerror = reject;
    img.src = dataUrl;
  });
};

const createPdfFromPhotos = async (
  files: File[],
  pdfFileName: string,
  headerTitle?: string,
  subHeaderTitle?: string
): Promise<Attachment> => {
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 15;

  const imagesData: { width: number; height: number; normalizedDataUrl: string; name: string }[] = [];
  for (const f of files) {
    const rawUrl = await readFileAsDataUrl(f);
    const dims = await loadImageDimensions(rawUrl);
    imagesData.push({ ...dims, name: f.name });
  }

  // Se sono 2 foto (es. Fronte/Retro documento d'identità), le impagina insieme sulla stessa pagina A4
  if (imagesData.length === 2) {
    let currentY = margin;
    if (headerTitle) {
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(13);
      doc.setTextColor(15, 23, 42);
      doc.text(headerTitle, margin, currentY + 5);
      if (subHeaderTitle) {
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(9);
        doc.setTextColor(100, 116, 139);
        doc.text(subHeaderTitle, margin, currentY + 11);
      }
      currentY += 18;
    }

    const availableHeightPerImg = (pageHeight - currentY - margin - 10) / 2;
    const availableWidth = pageWidth - margin * 2;

    imagesData.forEach((img, idx) => {
      const ratio = Math.min(availableWidth / img.width, availableHeightPerImg / img.height);
      const drawW = img.width * ratio;
      const drawH = img.height * ratio;
      const drawX = margin + (availableWidth - drawW) / 2;
      const slotY = currentY + idx * (availableHeightPerImg + 8);
      const drawY = slotY + (availableHeightPerImg - drawH) / 2;

      doc.setDrawColor(226, 232, 240);
      doc.roundedRect(drawX - 2, drawY - 2, drawW + 4, drawH + 4, 2, 2);
      doc.addImage(img.normalizedDataUrl, 'JPEG', drawX, drawY, drawW, drawH);
    });
  } else {
    imagesData.forEach((img, index) => {
      if (index > 0) doc.addPage();
      let currentY = margin;
      if (headerTitle) {
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(13);
        doc.setTextColor(15, 23, 42);
        doc.text(headerTitle, margin, currentY + 5);
        if (subHeaderTitle) {
          doc.setFont('helvetica', 'normal');
          doc.setFontSize(9);
          doc.setTextColor(100, 116, 139);
          doc.text(`${subHeaderTitle} - Pag. ${index + 1}/${imagesData.length}`, margin, currentY + 11);
        }
        currentY += 18;
      }

      const availableWidth = pageWidth - margin * 2;
      const availableHeight = pageHeight - currentY - margin;
      const ratio = Math.min(availableWidth / img.width, availableHeight / img.height);
      const drawW = img.width * ratio;
      const drawH = img.height * ratio;
      const drawX = margin + (availableWidth - drawW) / 2;
      const drawY = currentY + (availableHeight - drawH) / 2;

      doc.setDrawColor(226, 232, 240);
      doc.roundedRect(drawX - 2, drawY - 2, drawW + 4, drawH + 4, 2, 2);
      doc.addImage(img.normalizedDataUrl, 'JPEG', drawX, drawY, drawW, drawH);
    });
  }

  const pdfDataUrl = doc.output('datauristring');
  const pdfBlob = doc.output('blob');
  const finalName = pdfFileName.toLowerCase().endsWith('.pdf') ? pdfFileName : `${pdfFileName}.pdf`;

  return {
    id: Math.random().toString(36).substr(2, 9),
    fileName: finalName,
    fileSize: pdfBlob.size,
    uploadDate: formatDate(new Date()),
    mimeType: 'application/pdf',
    dataUrl: pdfDataUrl
  };
};

const triggerDownloadAttachment = (at: Attachment) => {
  if (!at.dataUrl) {
    alert(`Anteprima/Download non disponibile per "${at.fileName}".`);
    return;
  }
  const link = document.createElement('a');
  link.href = at.dataUrl;
  link.download = at.fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};

const parseFlexibleDateToISO = (val: string): string => {
  const s = (val || '').trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const parts = s.split(/[ /.-]/).filter(Boolean);
  if (parts.length === 3) {
    let dd = parts[0];
    let mm = parts[1];
    let yyyy = parts[2];
    if (dd.length === 4) {
      yyyy = parts[0];
      mm = parts[1];
      dd = parts[2];
    } else if (yyyy.length === 2) {
      yyyy = `20${yyyy}`;
    }
    return `${yyyy}-${mm.padStart(2, '0')}-${dd.padStart(2, '0')}`;
  }
  return '';
};

const FileUploader: React.FC<{ 
  onUpload: (files: Attachment[]) => void, 
  attachments: Attachment[], 
  onRemove: (id: string) => void, 
  canEdit: boolean,
  photoToPdfConfig?: {
    enabled: boolean;
    mode: 'case' | 'portfolio';
    caseTitle?: string;
    defaultPersonName?: string;
  }
}> = ({ onUpload, attachments, onRemove, canEdit, photoToPdfConfig }) => {
  const inputId = useMemo(() => `file-up-${Math.random().toString(36).substr(2, 7)}`, []);
  const idDocInputId = useMemo(() => `id-doc-up-${Math.random().toString(36).substr(2, 7)}`, []);
  const [isGeneratingPdf, setIsGeneratingPdf] = useState(false);
  const [isScanningOcr, setIsScanningOcr] = useState(false);
  const [ocrStatusMessage, setOcrStatusMessage] = useState<string | null>(null);
  const [isIdDocMode, setIsIdDocMode] = useState(false);
  const [idExpiryDate, setIdExpiryDate] = useState('');
  const [idExpiryText, setIdExpiryText] = useState('');
  const [idPersonName, setIdPersonName] = useState(photoToPdfConfig?.defaultPersonName || '');
  const [pendingIdImages, setPendingIdImages] = useState<{ file: File; previewUrl: string }[]>([]);

  useEffect(() => {
    if (photoToPdfConfig?.defaultPersonName && !idPersonName) {
      setIdPersonName(photoToPdfConfig.defaultPersonName);
    }
  }, [photoToPdfConfig?.defaultPersonName]);

  const resolvedExpiryISO = useMemo(() => {
    return idExpiryDate || parseFlexibleDateToISO(idExpiryText);
  }, [idExpiryDate, idExpiryText]);

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files: File[] = Array.from(e.target.files || []);
    if (files.length === 0) return;

    const imageFiles = files.filter((f: File) => f.type.startsWith('image/') || /\.(jpe?g|png|webp|heic)$/i.test(f.name));
    const otherFiles = files.filter((f: File) => !(f.type.startsWith('image/') || /\.(jpe?g|png|webp|heic)$/i.test(f.name)));

    const validFiles: Attachment[] = [];

    for (const f of otherFiles) {
      if (f.size > FILE_SIZE_LIMIT) {
        alert(`Il file ${f.name} supera il limite di 5MB`);
        continue;
      }
      const dataUrl = await readFileAsDataUrl(f);
      validFiles.push({
        id: Math.random().toString(36).substr(2, 9),
        fileName: f.name,
        fileSize: f.size,
        uploadDate: formatDate(new Date()),
        mimeType: f.type || 'application/octet-stream',
        dataUrl
      });
    }

    if (imageFiles.length > 0) {
      if (photoToPdfConfig?.enabled) {
        if (photoToPdfConfig.mode === 'case') {
          setIsGeneratingPdf(true);
          try {
            const cleanCaseName = (photoToPdfConfig.caseTitle || 'CASO').trim().replace(/[\\/:*?"<>|]/g, '') || 'CASO';
            const yymmdd = formatYYMMDD(new Date());
            const pdfName = `${cleanCaseName}_${yymmdd}.pdf`;
            const pdfAttachment = await createPdfFromPhotos(
              imageFiles,
              pdfName,
              `Caso: ${cleanCaseName}`,
              `Data: ${formatDate(new Date())}`
            );
            validFiles.push(pdfAttachment);
          } catch (err) {
            console.error('Errore creazione PDF da foto:', err);
            alert('Errore durante l\'impaginazione del PDF dalle foto.');
          } finally {
            setIsGeneratingPdf(false);
          }
        } else if (photoToPdfConfig.mode === 'portfolio') {
          // Su anagrafica: avvia il riconoscimento automatico o manuale del documento
          await processIdDocumentImages(imageFiles);
        }
      } else {
        for (const f of imageFiles) {
          if (f.size > FILE_SIZE_LIMIT) {
            alert(`Il file ${f.name} supera il limite di 5MB`);
            continue;
          }
          const dataUrl = await readFileAsDataUrl(f);
          validFiles.push({
            id: Math.random().toString(36).substr(2, 9),
            fileName: f.name,
            fileSize: f.size,
            uploadDate: formatDate(new Date()),
            mimeType: f.type,
            dataUrl
          });
        }
      }
    }

    if (validFiles.length > 0) onUpload(validFiles);
    e.target.value = '';
  };

  const processIdDocumentImages = async (imageFiles: File[]) => {
    setIsIdDocMode(true);
    const loadedPreviews: { file: File; previewUrl: string }[] = [];
    for (const f of imageFiles) {
      const rawUrl = await readFileAsDataUrl(f);
      const dims = await loadImageDimensions(rawUrl);
      loadedPreviews.push({ file: f, previewUrl: dims.normalizedDataUrl });
    }
    const nextImages = [...pendingIdImages, ...loadedPreviews];
    setPendingIdImages(nextImages);

    // Esegui autoriconoscimento (OCR AI) sulle immagini caricate
    setIsScanningOcr(true);
    setOcrStatusMessage("Autoriconoscimento dati documento in corso...");
    try {
      const response = await fetch("/api/ocr-id-document", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          images: nextImages.map(item => ({ data: item.previewUrl, mimeType: "image/jpeg" }))
        })
      });
      if (response.ok) {
        const result = await response.json();
        if (result.fullName && result.fullName.trim()) {
          setIdPersonName(result.fullName.trim().toUpperCase());
        }
        if (result.expiryDate && result.expiryDate.trim()) {
          const iso = parseFlexibleDateToISO(result.expiryDate.trim());
          if (iso) {
            setIdExpiryDate(iso);
            setIdExpiryText(formatDate(iso));
          }
        }
        if (result.expiryDate || result.fullName) {
          setOcrStatusMessage(`Dati riconosciuti automaticamente${result.documentType ? ` (${result.documentType})` : ''}. Verifica o modifica i campi qui sotto e clicca su "Genera e Carica PDF in Anagrafica".`);
        } else {
          setOcrStatusMessage("Impossibile leggere automaticamente tutti i dati dalla foto. Puoi inserirli o completarli manualmente qui sotto.");
        }
      } else {
        setOcrStatusMessage("Autoriconoscimento non disponibile per questa immagine. Inserisci o verifica i dati manualmente qui sotto.");
      }
    } catch (err) {
      console.error("Errore OCR:", err);
      setOcrStatusMessage("Inserisci o verifica manualmente Nome, Cognome e Data di Scadenza qui sotto.");
    } finally {
      setIsScanningOcr(false);
    }
  };

  const handleIdDocPhotosChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files: File[] = Array.from(e.target.files || []);
    if (files.length === 0) return;
    const imageFiles = files.filter((f: File) => f.type.startsWith('image/') || /\.(jpe?g|png|webp|heic)$/i.test(f.name));
    if (imageFiles.length === 0) {
      alert("Seleziona una o più foto (es. fronte e retro) del documento d'identità.");
      e.target.value = '';
      return;
    }
    await processIdDocumentImages(imageFiles);
    e.target.value = '';
  };

  const handleConfirmCreateIdPdf = async () => {
    if (pendingIdImages.length === 0) {
      alert("Seleziona almeno una foto del documento prima di generare il PDF.");
      return;
    }
    if (!resolvedExpiryISO) {
      alert("Inserisci la data di scadenza del documento (es. GG/MM/AAAA o dal calendario).");
      return;
    }
    setIsGeneratingPdf(true);
    try {
      const cleanPerson = (idPersonName || photoToPdfConfig?.defaultPersonName || 'NOME COGNOME').trim().toUpperCase().replace(/[\\/:*?"<>|]/g, '');
      const yymmdd = formatYYMMDD(resolvedExpiryISO);
      const pdfName = `DOC${yymmdd} ${cleanPerson}.pdf`;
      const pdfAttachment = await createPdfFromPhotos(
        pendingIdImages.map(x => x.file),
        pdfName,
        `DOCUMENTO D'IDENTITA' - ${cleanPerson}`,
        `Scadenza: ${formatDate(resolvedExpiryISO)} (Cod. DOC${yymmdd})`
      );
      onUpload([pdfAttachment]);
      setPendingIdImages([]);
      setOcrStatusMessage(null);
      setIsIdDocMode(false);
    } catch (err) {
      console.error(err);
      alert("Errore durante la creazione del PDF del documento d'identità.");
    } finally {
      setIsGeneratingPdf(false);
    }
  };

  const handleConfirmCreateGenericPortfolioPdf = async () => {
    if (pendingIdImages.length === 0) return;
    setIsGeneratingPdf(true);
    try {
      const cleanClient = (idPersonName || photoToPdfConfig?.defaultPersonName || 'ANAGRAFICA').trim().replace(/[\\/:*?"<>|]/g, '') || 'ANAGRAFICA';
      const yymmdd = formatYYMMDD(new Date());
      const pdfName = `${cleanClient}_${yymmdd}.pdf`;
      const pdfAttachment = await createPdfFromPhotos(
        pendingIdImages.map(x => x.file),
        pdfName,
        `Documentazione Anagrafica: ${cleanClient}`,
        `Data: ${formatDate(new Date())}`
      );
      onUpload([pdfAttachment]);
      setPendingIdImages([]);
      setOcrStatusMessage(null);
      setIsIdDocMode(false);
    } catch (err) {
      console.error(err);
      alert("Errore durante la creazione del PDF.");
    } finally {
      setIsGeneratingPdf(false);
    }
  };

  return (
    <div className="space-y-4">
      {canEdit && (
        <div className="space-y-3">
          <div className="p-4 border-2 border-dashed border-slate-200 rounded-2xl flex flex-col items-center justify-center gap-2 hover:border-emerald-400 hover:bg-emerald-50/40 transition-all group">
            <Paperclip className="text-slate-300 group-hover:text-emerald-500 transition-colors" />
            <p className="text-[10px] font-black uppercase text-slate-400 tracking-widest text-center">
              {photoToPdfConfig?.enabled
                ? 'PDF, XLS, XLSX o FOTO (le foto vengono impaginate in PDF e caricate direttamente nella scheda)'
                : 'PDF, XLS, XLSX (Max 5MB)'}
            </p>
            <input
              type="file"
              multiple
              className="hidden"
              id={inputId}
              onChange={handleFileChange}
              accept={photoToPdfConfig?.enabled ? '.pdf,.xls,.xlsx,image/*' : '.pdf,.xls,.xlsx'}
            />
            <div className="flex flex-wrap items-center justify-center gap-3 mt-1">
              <label htmlFor={inputId} className="cursor-pointer bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-xl font-black text-xs shadow-sm transition-all">
                {isGeneratingPdf ? 'Impaginazione PDF in corso...' : (photoToPdfConfig?.enabled ? 'Carica File o Foto' : 'Seleziona Allegati')}
              </label>
              {photoToPdfConfig?.mode === 'portfolio' && (
                <button
                  type="button"
                  onClick={() => setIsIdDocMode(prev => !prev)}
                  className={`px-4 py-2 rounded-xl font-black text-xs border transition-all flex items-center gap-1.5 ${
                    isIdDocMode
                      ? 'bg-indigo-600 text-white border-indigo-600 shadow-sm'
                      : 'bg-indigo-50 text-indigo-700 border-indigo-200 hover:bg-indigo-100'
                  }`}
                >
                  <Sparkles size={14} /> {isIdDocMode ? 'Chiudi Pannello Documento Identità' : 'Documento d\'Identità (Autoriconoscimento & PDF)'}
                </button>
              )}
            </div>
          </div>

          {photoToPdfConfig?.mode === 'portfolio' && isIdDocMode && (
            <div className="p-5 bg-indigo-50/70 border-2 border-indigo-100 rounded-2xl space-y-4 animate-in fade-in duration-200">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                <div>
                  <p className="text-xs font-black uppercase text-indigo-950 tracking-wider flex items-center gap-1.5">
                    <Sparkles size={14} className="text-indigo-600" /> Documento d'Identità (Autoriconoscimento + Inserimento Manuale)
                  </p>
                  <p className="text-[10px] font-bold text-indigo-600 mt-0.5">
                    Nome file generato: <span className="font-mono bg-white px-2 py-0.5 rounded border border-indigo-200">DOC{resolvedExpiryISO ? formatYYMMDD(resolvedExpiryISO) : 'aammgg'} {(idPersonName || photoToPdfConfig?.defaultPersonName || 'NOME COGNOME').toUpperCase()}.pdf</span>
                  </p>
                </div>
                <div>
                  <input
                    type="file"
                    multiple
                    accept="image/*"
                    id={idDocInputId}
                    className="hidden"
                    onChange={handleIdDocPhotosChange}
                  />
                  <label
                    htmlFor={idDocInputId}
                    className="px-4 py-2.5 rounded-xl font-black text-xs inline-flex items-center gap-2 bg-indigo-600 hover:bg-indigo-700 text-white cursor-pointer shadow-md transition-all"
                  >
                    <Upload size={14} /> {pendingIdImages.length > 0 ? 'Aggiungi Altra Foto (es. Retro)' : '1. Carica Foto Documento (Fronte / Retro)'}
                  </label>
                </div>
              </div>

              {isScanningOcr && (
                <div className="p-3 bg-indigo-100/80 border border-indigo-200 rounded-xl flex items-center gap-2 text-xs font-black text-indigo-900 animate-pulse">
                  <Sparkles size={16} className="text-indigo-600" /> Lettura automatica di Nome, Cognome e Data di Scadenza dal documento in corso...
                </div>
              )}

              {ocrStatusMessage && !isScanningOcr && (
                <div className="p-3 bg-white border border-indigo-200 rounded-xl flex items-center gap-2 text-xs font-bold text-indigo-900">
                  <CheckCircle2 size={15} className="text-indigo-600 shrink-0" />
                  <span>{ocrStatusMessage}</span>
                </div>
              )}

              {pendingIdImages.length > 0 && (
                <div className="flex flex-wrap gap-3 pt-1">
                  {pendingIdImages.map((item, idx) => (
                    <div key={idx} className="relative w-28 h-20 rounded-xl overflow-hidden border-2 border-indigo-200 bg-white shadow-sm group">
                      <img src={item.previewUrl} alt={`Pagina ${idx + 1}`} className="w-full h-full object-cover" />
                      <span className="absolute bottom-1 left-1 bg-indigo-950/80 text-white text-[9px] font-black px-1.5 py-0.5 rounded">
                        Foto {idx + 1}
                      </span>
                      <button
                        type="button"
                        onClick={() => setPendingIdImages(prev => prev.filter((_, i) => i !== idx))}
                        className="absolute top-1 right-1 bg-red-600 text-white rounded-full p-0.5 hover:bg-red-700"
                        title="Rimuovi foto"
                      >
                        <X size={11} />
                      </button>
                    </div>
                  ))}
                </div>
              )}

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 pt-1">
                <div className="space-y-1">
                  <label className="text-[10px] font-black uppercase text-indigo-900">2. Nome e Cognome Titolare (modificabile a mano) *</label>
                  <input
                    type="text"
                    placeholder="Es. MARIO ROSSI"
                    value={idPersonName}
                    onChange={e => setIdPersonName(e.target.value)}
                    className="w-full border-2 border-indigo-200 p-2.5 rounded-xl bg-white font-bold text-xs outline-none focus:border-indigo-600 uppercase"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-[10px] font-black uppercase text-indigo-900">3. Data Scadenza Documento (a mano GG/MM/AAAA o calendario) *</label>
                  <div className="flex items-center gap-2">
                    <input
                      type="text"
                      placeholder="GG/MM/AAAA (es. 15/08/2030)"
                      value={idExpiryText}
                      onChange={e => {
                        const val = e.target.value;
                        setIdExpiryText(val);
                        const iso = parseFlexibleDateToISO(val);
                        if (iso) setIdExpiryDate(iso);
                      }}
                      className="flex-1 border-2 border-indigo-200 p-2.5 rounded-xl bg-white font-mono font-bold text-xs outline-none focus:border-indigo-600"
                    />
                    <input
                      type="date"
                      value={idExpiryDate}
                      onChange={e => {
                        setIdExpiryDate(e.target.value);
                        setIdExpiryText(e.target.value ? formatDate(e.target.value) : '');
                      }}
                      className="border-2 border-indigo-200 p-2.5 rounded-xl bg-white font-bold text-xs outline-none focus:border-indigo-600"
                      title="Seleziona da calendario"
                    />
                  </div>
                </div>
              </div>

              <div className="flex flex-wrap justify-end gap-2 pt-2 border-t border-indigo-100">
                {pendingIdImages.length > 0 && (
                  <button
                    type="button"
                    onClick={handleConfirmCreateGenericPortfolioPdf}
                    disabled={isGeneratingPdf || isScanningOcr}
                    className="px-4 py-2.5 rounded-xl font-bold text-xs bg-white hover:bg-slate-100 text-slate-700 border border-slate-200 transition-all"
                  >
                    Salva come PDF Generico Anagrafica
                  </button>
                )}
                <button
                  type="button"
                  onClick={handleConfirmCreateIdPdf}
                  disabled={isGeneratingPdf || isScanningOcr || pendingIdImages.length === 0 || !resolvedExpiryISO}
                  className="px-5 py-2.5 rounded-xl font-black text-xs flex items-center gap-2 bg-emerald-700 hover:bg-emerald-800 disabled:bg-slate-200 disabled:text-slate-400 text-white shadow-md transition-all"
                >
                  <Check size={14} /> 4. Genera e Carica PDF in Anagrafica (`DOC{resolvedExpiryISO ? formatYYMMDD(resolvedExpiryISO) : 'aammgg'} {(idPersonName || photoToPdfConfig?.defaultPersonName || 'NOME COGNOME').toUpperCase()}.pdf`)
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      <div className="flex flex-wrap gap-2">
        {attachments.map(at => (
          <div key={at.id} className="bg-slate-100 px-3 py-2 rounded-xl flex items-center gap-2 max-w-[320px] border border-slate-200">
            <Files size={14} className="text-emerald-600 shrink-0" />
            <div className="flex-1 min-w-0">
              <p className="text-[11px] font-bold text-slate-700 truncate" title={at.fileName}>{at.fileName}</p>
              <p className="text-[9px] text-slate-400 font-bold uppercase">{at.uploadDate} • {(at.fileSize / 1024).toFixed(0)} KB</p>
            </div>
            {at.dataUrl && (
              <button
                type="button"
                onClick={() => triggerDownloadAttachment(at)}
                className="text-emerald-600 hover:text-emerald-800 p-1 rounded-lg hover:bg-emerald-50 transition-colors"
                title="Scarica file / PDF"
              >
                <Download size={14} />
              </button>
            )}
            {canEdit && (
              <button type="button" onClick={() => onRemove(at.id)} className="text-red-400 hover:text-red-600 p-1">
                <X size={14} />
              </button>
            )}
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
    <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[9999] flex items-center justify-center p-4 animate-in fade-in duration-300">
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
    <div className="fixed inset-0 z-[9999] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-200">
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

  const [deleteConfirmId, setDeleteConfirmId] = useState<string | null>(null);

  const deleteNews = (id: string) => {
    setDeleteConfirmId(id);
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
      <ConfirmModal
        isOpen={!!deleteConfirmId}
        title="Conferma Eliminazione"
        message="Sei sicuro di voler eliminare questa notizia? Questa operazione non può essere annullata."
        onConfirm={() => {
          setNews(prev => prev.filter(n => n.id !== deleteConfirmId));
          setDeleteConfirmId(null);
        }}
        onCancel={() => setDeleteConfirmId(null)}
      />
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

      <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
        <StatCard label="Portafoglio" value={stats.portfolios.length} icon={<Users className="text-emerald-500" />} onClick={() => setView('portfolio')} />
        <StatCard label="PDP Attivi" value={stats.pdps.length} icon={<Zap className="text-yellow-500" />} onClick={() => setView('pdp')} />
        <StatCard 
          label="Forniture" 
          value={stats.contracts.length} 
          icon={
            <span className="flex items-center -space-x-1">
              <Flame className="text-amber-500 fill-amber-500/20" size={20} />
              <Lightbulb className="text-yellow-400 fill-yellow-400/20" size={20} />
            </span>
          } 
          onClick={() => setView('contracts')} 
        />
        <StatCard label="Telefonico" value={stats.contrattiTelefonici?.length || 0} icon={<Phone className="text-sky-400" />} onClick={() => setView('telefonico')} />
        <StatCard label="Casi Aperti" value={stats.casi.length} icon={<AlertCircle className="text-red-500" />} onClick={() => setView('casi')} />
      </div>
    </div>
  );
};

const PortfolioView: React.FC<any> = ({ data, setData, user, users, partners = [], appartenenze = [], contracts = [], contrattiTelefonici = [], pdps = [], exportFn, setCreditChecks, onCreateContract }) => {
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [form, setForm] = useState<Partial<Portfolio>>({ attachments: [], entityType: 'Domestico' });
  const [editingId, setEditingId] = useState<string | null>(null);
  const [creditCheckPortfolio, setCreditCheckPortfolio] = useState<Portfolio | null>(null);
  const [selectedPortfolio, setSelectedPortfolio] = useState<Portfolio | null>(null);
  const [deleteConfirmId, setDeleteConfirmId] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');

  const ENTITY_TYPES: PortfolioEntityType[] = ['Domestico', 'Impresa', 'Associazione', 'Condominio', 'PA'];

  const filteredPortfolios = useMemo(() => {
    if (!searchTerm) return data;
    const s = searchTerm.toLowerCase();
    return data.filter((p: Portfolio) => {
      const partner = partners.find((pt: Partner) => pt.id === p.partnerId);
      const appartenenza = appartenenze.find((a: Appartenenza) => a.id === p.appartenenzaId);
      return (
        p.businessName?.toLowerCase().includes(s) || 
        p.taxId?.toLowerCase().includes(s) ||
        p.city?.toLowerCase().includes(s) ||
        p.street?.toLowerCase().includes(s) ||
        p.province?.toLowerCase().includes(s) ||
        p.cap?.toLowerCase().includes(s) ||
        p.legalAddress?.toLowerCase().includes(s) ||
        partner?.name?.toLowerCase().includes(s) ||
        partner?.userCode?.toLowerCase().includes(s) ||
        appartenenza?.name?.toLowerCase().includes(s)
      );
    });
  }, [data, searchTerm, partners, appartenenze]);

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
    const finalLegalAddress = form.legalAddress || composeAddress(form.street, form.cap, form.city, form.province);
    const parsed = splitAddress(finalLegalAddress);
    const payload = { 
      ...form, 
      legalAddress: finalLegalAddress,
      street: form.street || parsed?.street || '',
      cap: form.cap || parsed?.cap || '',
      city: form.city || parsed?.city || '',
      province: (form.province || parsed?.province || '').toUpperCase(),
      assignedTo: form.assignedTo || user.id, 
      partnerId: form.partnerId || undefined,
      appartenenzaId: form.appartenenzaId || undefined,
      updatedBy: user.name 
    };
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
             <div className="space-y-1.5">
               <label className="text-xs font-bold text-slate-500 uppercase ml-1">{form.entityType === 'Domestico' ? 'Nome' : 'Ragione Sociale'}</label>
               <input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" required value={form.businessName || ''} onChange={e => setForm({...form, businessName: e.target.value})} />
             </div>
             <div className="space-y-1.5">
               <label className="text-xs font-bold text-slate-500 uppercase ml-1 flex justify-between">
                 <span>{form.entityType === 'Domestico' ? 'Codice Fiscale' : 'C.F. o P.IVA'}</span>
                 <span className="text-[9px] text-slate-400">{form.entityType === 'Domestico' ? '16 caratteri' : '11 o 16 caratteri'}</span>
               </label>
               <input 
                 className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-mono font-bold focus:border-emerald-500" 
                 required 
                 maxLength={16}
                 value={form.taxId || ''} 
                 onChange={e => setForm({...form, taxId: e.target.value.toUpperCase()})} 
               />
             </div>
             <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">PEC</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" value={form.pec || ''} onChange={e => setForm({...form, pec: e.target.value})} /></div>
             <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Cellulare</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" value={form.phone || ''} onChange={e => setForm({...form, phone: e.target.value})} /></div>
             <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">E-mail</label><input type="email" className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" value={form.email || ''} onChange={e => setForm({...form, email: e.target.value})} /></div>

             {form.entityType !== 'Domestico' && (
               <>
                 <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">CF Rapp. Legale</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-mono font-bold focus:border-emerald-500" maxLength={16} value={form.legalRepresentativeTaxId || ''} onChange={e => setForm({...form, legalRepresentativeTaxId: e.target.value.toUpperCase()})} /></div>
                 <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Nome Rapp. Legale</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" value={form.legalRepresentativeName || ''} onChange={e => setForm({...form, legalRepresentativeName: e.target.value})} /></div>
               </>
             )}
             {(user.role === 'admin' || user.role === 'backoffice') && (
               <div className="col-span-full">
                 <SearchableSelect label="Assegna a Utente" placeholder="Cerca utente..." options={users} selectedId={form.assignedTo} onSelect={(id:string) => setForm({...form, assignedTo: id})} displayFn={(u:any)=>u.name} searchFn={(u:any)=>u.name} subTextFn={(u:any)=>u.role} icon={<UserIcon size={14}/>} />
               </div>
             )}
             <div>
               <SearchableSelect 
                 label="Partner" 
                 placeholder="Seleziona Partner (opzionale)..." 
                 options={partners} 
                 selectedId={form.partnerId} 
                 onSelect={(id:string) => setForm({...form, partnerId: id})} 
                 displayFn={(p:Partner)=>p.name} 
                 searchFn={(p:Partner)=>`${p.name} ${p.userCode}`} 
                 subTextFn={(p:Partner)=>`Codice: ${p.userCode}`} 
                 icon={<Handshake size={14}/>} 
                 allowClear={true}
               />
             </div>
             <div>
               <SearchableSelect 
                 label="Appartenenza (Marchio / Gruppo)" 
                 placeholder="Seleziona Appartenenza (opzionale)..." 
                 options={appartenenze} 
                 selectedId={form.appartenenzaId} 
                 onSelect={(id:string) => setForm({...form, appartenenzaId: id})} 
                 displayFn={(a:Appartenenza)=>a.name} 
                 searchFn={(a:Appartenenza)=>`${a.name} ${a.description || ''}`} 
                 subTextFn={(a:Appartenenza)=>a.description || 'Gruppo / Marchio'} 
                 icon={<Layers size={14}/>} 
                 allowClear={true}
               />
             </div>
             <AddressInputGroup 
               label="Indirizzo (Via, CAP Città, PROV)"
               address={form.legalAddress || ''}
               street={form.street || ''}
               cap={form.cap || ''}
               city={form.city || ''}
               province={form.province || ''}
               accentColor="emerald"
               onChange={({ address, street, cap, city, province }) => {
                 setForm(prev => ({
                   ...prev,
                   legalAddress: address,
                   street,
                   cap,
                   city,
                   province
                 }));
               }}
             />
             <div className="col-span-full"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Allegati & Documenti (PDF / Foto)</label>
                <FileUploader 
                  attachments={form.attachments || []} 
                  canEdit={true} 
                  onUpload={files => setForm({...form, attachments: [...(form.attachments || []), ...files]})} 
                  onRemove={id => setForm({...form, attachments: form.attachments?.filter(a => a.id !== id)})}
                  photoToPdfConfig={{
                    enabled: true,
                    mode: 'portfolio',
                    defaultPersonName: form.entityType === 'Domestico' ? (form.businessName || '') : (form.legalRepresentativeName || form.businessName || '')
                  }}
                />
             </div>
           </div>
           {editingId && <TimestampDetail data={form} />}
           <div className="flex justify-end gap-4 border-t pt-8"><button type="button" onClick={() => setIsFormOpen(false)} className="font-bold text-slate-400">Annulla</button><button type="submit" className="bg-emerald-900 text-white px-10 py-4 rounded-2xl font-black">{editingId ? 'Aggiorna Cliente' : 'Registra Anagrafica'}</button></div>
        </form>
      )}
      <div className="bg-white rounded-[2rem] border border-slate-100 shadow-sm overflow-hidden">
        <table className="w-full text-left">
          <thead className="bg-slate-50 border-b">
            <tr>
              <th className="p-6 text-[10px] font-black uppercase text-slate-400">Tipologia / Ragione Sociale</th>
              <th className="p-6 text-[10px] font-black uppercase text-slate-400">Appartenenza / Partner</th>
              {(user.role === 'admin' || user.role === 'backoffice') && (
                <th className="p-6 text-[10px] font-black uppercase text-slate-400 text-right">Assegnato (Codice)</th>
              )}
              <th className="p-6"></th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-50">
            {filteredPortfolios.map((p: Portfolio) => {
              const assignedUser = [...users, ...partners].find(u => u.id === p.assignedTo);
              const assignedPartner = partners.find((pt: Partner) => pt.id === p.partnerId);
              const assignedAppartenenza = appartenenze.find((a: Appartenenza) => a.id === p.appartenenzaId);
              const isAdminOrBO = user.role === 'admin' || user.role === 'backoffice';
              return (
              <tr key={p.id} className="hover:bg-slate-50 transition-colors group">
                <td className="p-6">
                  <p className="text-[10px] font-black text-emerald-600 uppercase tracking-tighter mb-1">{p.entityType || '---'}</p>
                  <p onClick={() => setSelectedPortfolio(p)} className="font-bold text-emerald-950 hover:text-emerald-600 cursor-pointer transition-colors hover:underline decoration-dotted">{p.businessName}</p>
                  <p className="text-[10px] text-slate-400 font-mono">{p.taxId}</p>
                  {p.updatedBy && (
                    <p className="text-[8px] text-slate-300 uppercase font-black mt-1 flex items-center gap-1">
                      <Clock size={8} /> Modificato da: {p.updatedBy}
                    </p>
                  )}
                </td>
                <td className="p-6">
                  <div className="flex flex-col items-start gap-1.5">
                    {assignedAppartenenza ? (
                      <span className="inline-flex items-center gap-1 bg-indigo-50 text-indigo-800 border border-indigo-200 px-2.5 py-0.5 rounded-lg text-[10px] font-black">
                        <Layers size={11} className="text-indigo-500" /> {assignedAppartenenza.name}
                      </span>
                    ) : (
                      <span className="text-[10px] text-slate-300 font-bold">Nessuna appartenenza</span>
                    )}
                    {assignedPartner && (
                      <span className="inline-flex items-center gap-1 bg-teal-50 text-teal-800 border border-teal-200 px-2.5 py-0.5 rounded-lg text-[10px] font-black">
                        <Handshake size={11} className="text-teal-600" /> Partner: {assignedPartner.name}{isAdminOrBO ? ` (${assignedPartner.userCode})` : ''}
                      </span>
                    )}
                  </div>
                </td>
                {isAdminOrBO && (
                  <td className="p-6 text-right">
                    <p className="text-sm font-black text-emerald-700 leading-none">{assignedUser?.userCode || '---'}</p>
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-tighter mt-1">{assignedUser?.name || '---'}</p>
                  </td>
                )}
                <td className="p-6 text-right space-x-2">
                  <div className="flex items-center justify-end gap-2">
                    <AddressLink address={p.legalAddress} showIconOnly className="text-blue-500 hover:scale-110 transition-transform" />
                    {(() => {
                      const canCreditCheck = p.entityType === 'Impresa' || p.entityType === 'Condominio';
                      return (
                        <button 
                          type="button"
                          disabled={!canCreditCheck}
                          onClick={canCreditCheck ? () => setCreditCheckPortfolio(p) : undefined} 
                          className={`p-2 rounded-lg border transition-all ${
                            canCreditCheck 
                              ? 'text-emerald-600 hover:bg-emerald-50 border-emerald-100 cursor-pointer' 
                              : 'text-slate-300 bg-slate-50 border-slate-100 cursor-not-allowed opacity-40'
                          }`} 
                          title={canCreditCheck ? "Richiedi Credit Check" : "Credit Check disattivato: abilitato solo per IMPRESA e CONDOMINIO"}
                        >
                          <Wallet size={18} />
                        </button>
                      );
                    })()}
                    <button onClick={() => onCreateContract(p.id)} className="text-blue-600 hover:bg-blue-50 p-2 rounded-lg border border-blue-100 transition-all" title="Crea Contratto">
                       <FileSignature size={18} />
                     </button>
                    <button onClick={() => { 
                      const parsed = splitAddress(p.legalAddress);
                      setForm({
                        ...p,
                        street: p.street || parsed?.street || '',
                        cap: p.cap || parsed?.cap || '',
                        city: p.city || parsed?.city || '',
                        province: p.province || parsed?.province || ''
                      }); 
                      setEditingId(p.id); 
                      setIsFormOpen(true); 
                    }} className="text-slate-400 hover:text-emerald-600 p-2"><Edit2 size={16}/></button>
                    <button onClick={() => setDeleteConfirmId(p.id)} className="text-slate-400 hover:text-red-600 p-2"><X size={18}/></button>
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

      {selectedPortfolio && (
        <PortfolioModal 
          isOpen={!!selectedPortfolio} 
          onClose={() => setSelectedPortfolio(null)} 
          portfolio={selectedPortfolio} 
          setData={setData}
          user={user}
          users={users}
          partners={partners}
          appartenenze={appartenenze}
          portfolios={data}
          contracts={contracts}
          contrattiTelefonici={contrattiTelefonici}
          pdps={pdps}
          onSelectPortfolio={(p) => setSelectedPortfolio(p)}
        />
      )}
      <ConfirmModal
        isOpen={!!deleteConfirmId}
        title="Conferma Eliminazione"
        message="Sei sicuro di voler eliminare questa anagrafica cliente? Questa operazione eliminerà anche tutti i record ad essa collegati."
        onConfirm={() => {
          setData((prev: Portfolio[]) => prev.filter(i => i.id !== deleteConfirmId));
          setDeleteConfirmId(null);
        }}
        onCancel={() => setDeleteConfirmId(null)}
      />
    </div>
  );
};

const PDPView: React.FC<any> = ({ data, setData, cabine, setCabine, contracts, portfolios, users, user, exportFn, casi = [], onViewPdp }) => {
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [form, setForm] = useState<Partial<PDP>>({ service: 'POWER', potenzaImpegnata: 0, potenzaDisponibile: 0 });
  const [editingId, setEditingId] = useState<string | null>(null);
  const [deleteConfirmId, setDeleteConfirmId] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');

  const filteredPdps = useMemo(() => {
    return data.filter((p: PDP) => {
      // Find associated clients for this PDP based on contracts
      const associatedContracts = contracts.filter((c: any) => c.pdpId === p.id);
      const clientNames = associatedContracts
        .map((c: any) => {
          const pf = portfolios.find((portfolio: any) => portfolio.id === c.portfolioId);
          return pf ? pf.businessName : '';
        })
        .filter(Boolean)
        .join(' ');

      const searchStr = `${p.pdpCode} ${p.matricola || ''} ${p.address || ''} ${p.street || ''} ${p.cap || ''} ${p.city || ''} ${p.province || ''} ${p.service || ''} ${p.cabinaPrimaria || p.technicalSpecs || ''} ${p.remi || ''} ${p.portata || ''} ${clientNames}`.toLowerCase();
      return searchStr.includes(searchTerm.toLowerCase());
    });
  }, [data, contracts, portfolios, searchTerm]);

  return (
    <div className="space-y-6 animate-in fade-in duration-300">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h2 className="text-3xl font-black text-emerald-900">Punti Di Prelievo</h2>
          <p className="text-slate-500 font-medium italic">Database tecnico utenze</p>
        </div>
        <div className="flex flex-wrap items-center gap-3 w-full md:w-auto">
          {/* Campo di ricerca */}
          <div className="relative flex-1 md:flex-none md:w-64">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
            <input 
              className="w-full pl-10 pr-4 py-2.5 rounded-2xl border-2 border-slate-100 outline-none focus:border-emerald-500 font-bold text-sm bg-white" 
              placeholder="Cerca PDP, matricola, REMI..." 
              value={searchTerm} 
              onChange={e => setSearchTerm(e.target.value)}
            />
          </div>
          <button onClick={() => exportFn(filteredPdps, 'PDP')} className="p-3 bg-white border border-slate-100 rounded-2xl text-slate-400 hover:text-slate-800 transition-colors shadow-sm" title="Esporta"><FileSpreadsheet size={20} /></button>
          <button onClick={() => { setForm({ service: 'POWER', potenzaImpegnata: 0, potenzaDisponibile: 0 }); setEditingId(null); setIsFormOpen(true); }} className="bg-yellow-500 text-white px-6 py-3 rounded-2xl font-black flex items-center gap-2 shadow-lg hover:scale-105 transition-all text-sm"><Plus size={18}/> Nuovo PDP</button>
        </div>
      </div>
      {isFormOpen && (
        <form onSubmit={(e) => {
          e.preventDefault();
          const svc = form.service || 'POWER';
          if (svc === 'METANO' && form.portata) {
            if (!form.portata.toUpperCase().startsWith('G')) {
              return alert("Errore: La portata per il servizio METANO deve iniziare obbligatoriamente con la lettera 'G' maiuscola (es. G4, G6, G10).");
            }
          }
          const ts = getTimestamp();
          let cId = form.cabinaPrimariaId;
          const cpName = form.cabinaPrimaria || form.technicalSpecs;
          if (cpName && svc === 'POWER') {
            let ex = cabine.find((c: any) => c.name.toLowerCase() === cpName.toLowerCase());
            if (!ex) {
              ex = { id: Math.random().toString(36).substr(2, 9), name: cpName };
              setCabine((prev: any) => [...prev, ex]);
            }
            cId = ex.id;
          }

          const finalAddress = form.address || composeAddress(form.street, form.cap, form.city, form.province);
          const parsed = splitAddress(finalAddress);

          const pdpPayload: Partial<PDP> = {
            ...form,
            service: svc,
            pdpCode: (form.pdpCode || '').toUpperCase().trim(),
            address: finalAddress,
            street: form.street || parsed?.street || '',
            cap: form.cap || parsed?.cap || '',
            city: form.city || parsed?.city || '',
            province: (form.province || parsed?.province || '').toUpperCase(),
            matricola: form.matricola ? form.matricola.toUpperCase().trim() : '',
            cabinaPrimaria: svc === 'POWER' ? cpName : '',
            technicalSpecs: svc === 'POWER' ? cpName : '',
            cabinaPrimariaId: svc === 'POWER' ? cId : undefined,
            potenzaImpegnata: svc === 'POWER' ? Number(form.potenzaImpegnata || 0) : 0,
            potenzaDisponibile: svc === 'POWER' ? Number(form.potenzaDisponibile || 0) : 0,
            tensione: svc === 'POWER' ? formatTensione(form.tensione) : '',
            remi: svc === 'METANO' ? (form.remi || '').toUpperCase().trim() : '',
            portata: svc === 'METANO' ? (form.portata || '').toUpperCase().trim() : '',
            updatedAt: ts
          };

          if (editingId) {
            setData((prev: PDP[]) => prev.map(p => p.id === editingId ? { ...p, ...pdpPayload } : p));
          } else {
            setData((prev: PDP[]) => [...prev, { 
              ...pdpPayload, 
              id: Math.random().toString(36).substr(2, 9), 
              assignedTo: user.id, 
              createdAt: ts 
            } as PDP]);
          }
          setIsFormOpen(false); 
          setEditingId(null); 
          setForm({ service: 'POWER', potenzaImpegnata: 0, potenzaDisponibile: 0 });
        }} className="bg-white p-8 rounded-[2.5rem] border-2 border-yellow-50 shadow-2xl space-y-6 animate-in slide-in-from-top-4">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            {/* Scelta POWER / METANO come su Contratti */}
            <div className="col-span-full space-y-1.5">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1">Tipo Servizio</label>
              <div className="flex gap-4">
                {(['POWER', 'METANO'] as const).map(s => (
                  <button 
                    key={s} 
                    type="button" 
                    onClick={() => setForm(prev => ({ ...prev, service: s }))} 
                    className={`flex-1 p-4 rounded-2xl font-black transition-all border-2 flex items-center justify-center gap-2 ${
                      (form.service || 'POWER') === s 
                        ? 'bg-yellow-500 text-white border-yellow-600 shadow-lg scale-[1.01]' 
                        : 'bg-slate-50 text-slate-400 border-slate-100 hover:bg-slate-100'
                    }`}
                  >
                    {s === 'POWER' ? <Zap size={18} /> : <Flame size={18} />}
                    {s}
                  </button>
                ))}
              </div>
            </div>

            {/* anziché CODICE chiamalo PDP */}
            <div className="col-span-full md:col-span-1 space-y-1.5">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1 flex justify-between">
                <span>PDP</span>
                <span className="text-[10px] text-slate-400">14 caratteri</span>
              </label>
              <input 
                maxLength={14} 
                minLength={14} 
                pattern=".{14}" 
                title="Il PDP deve essere di 14 caratteri" 
                className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-mono font-black uppercase focus:border-yellow-500" 
                required 
                placeholder="Es: IT001E12345678"
                value={form.pdpCode || ''} 
                onChange={e => setForm({...form, pdpCode: e.target.value.toUpperCase()})} 
              />
            </div>

            {/* Indirizzo Fornitura Confluente */}
            <AddressInputGroup 
              label="Indirizzo Fornitura (Via, CAP Città, PROV)"
              address={form.address || ''}
              street={form.street || ''}
              cap={form.cap || ''}
              city={form.city || ''}
              province={form.province || ''}
              accentColor="yellow"
              onChange={({ address, street, cap, city, province }) => {
                setForm(prev => ({
                  ...prev,
                  address,
                  street,
                  cap,
                  city,
                  province
                }));
              }}
            />

            {/* Matricola */}
            <div className="col-span-full md:col-span-1 space-y-1.5">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1">Matricola</label>
              <input 
                className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold uppercase focus:border-yellow-500" 
                placeholder="Matricola contatore"
                value={form.matricola || ''} 
                onChange={e => setForm({...form, matricola: e.target.value.toUpperCase()})} 
              />
            </div>

            {/* Campi condizionali POWER */}
            {(form.service || 'POWER') === 'POWER' && (
              <>
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1">Cabina Primaria</label>
                  <input 
                    className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-yellow-500" 
                    placeholder="Es: CP MILANO NORD"
                    value={form.cabinaPrimaria || form.technicalSpecs || ''} 
                    onChange={e => setForm({...form, cabinaPrimaria: e.target.value, technicalSpecs: e.target.value})} 
                  />
                </div>
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1">Pot. Disp. (KWp)</label>
                  <input 
                    type="number" 
                    step="0.1" 
                    className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-yellow-500" 
                    value={form.potenzaDisponibile ?? ''} 
                    onChange={e => setForm({...form, potenzaDisponibile: Number(e.target.value)})} 
                  />
                </div>
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1">Pot. Imp. (KWp)</label>
                  <input 
                    type="number" 
                    step="0.1" 
                    className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-yellow-500" 
                    value={form.potenzaImpegnata ?? ''} 
                    onChange={e => setForm({...form, potenzaImpegnata: Number(e.target.value)})} 
                  />
                </div>
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1">Tensione</label>
                  <select 
                    className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-yellow-500" 
                    value={formatTensione(form.tensione)} 
                    onChange={e => setForm({...form, tensione: e.target.value})}
                  >
                    <option value="">Seleziona Tensione...</option>
                    {TENSIONE_OPTIONS.map(opt => (
                      <option key={opt.value} value={opt.value}>{opt.label} V</option>
                    ))}
                  </select>
                </div>
              </>
            )}

            {/* Campi condizionali METANO */}
            {form.service === 'METANO' && (
              <>
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1">REMI</label>
                  <input 
                    className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold uppercase focus:border-yellow-500" 
                    placeholder="Codice punto REMI"
                    value={form.remi || ''} 
                    onChange={e => setForm({...form, remi: e.target.value.toUpperCase()})} 
                  />
                </div>
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1 flex justify-between">
                    <span>Portata</span>
                    <span className="text-[10px] text-amber-600 font-bold">Inizia obbligatoriamente con 'G' (es. G4, G6)</span>
                  </label>
                  <input 
                    className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold uppercase focus:border-yellow-500" 
                    placeholder="Es: G4, G6, G10..."
                    value={form.portata || ''} 
                    onChange={e => setForm({...form, portata: e.target.value.toUpperCase()})} 
                  />
                </div>
              </>
            )}
          </div>
          <div className="flex justify-end gap-4 border-t pt-6">
            <button type="button" onClick={() => setIsFormOpen(false)} className="font-bold text-slate-400">Annulla</button>
            <button type="submit" className="bg-emerald-900 text-white px-10 py-4 rounded-2xl font-black">{editingId ? 'Aggiorna PDP' : 'Salva PDP'}</button>
          </div>
        </form>
      )}
      <div className="bg-white rounded-3xl overflow-hidden shadow-sm border border-slate-100">
        <table className="w-full text-left">
          <thead className="bg-slate-50 border-b">
            <tr>
              <th className="p-6 text-[10px] font-black uppercase text-slate-400">PDP</th>
              <th className="p-6 text-[10px] font-black uppercase text-slate-400">Matricola</th>
              <th className="p-6 text-[10px] font-black uppercase text-slate-400">Località (Maps)</th>
              <th className="p-6 text-[10px] font-black uppercase text-slate-400">Dati Tecnici</th>
              {(user.role === 'admin' || user.role === 'backoffice') && (
                <th className="p-6 text-[10px] font-black uppercase text-slate-400">Utente (Codice)</th>
              )}
              <th className="p-6"></th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-50">
            {filteredPdps.map((p: PDP) => {
              const lastContract = [...contracts].filter(c => c.pdpId === p.id).sort((a,b) => new Date(b.startDate).getTime() - new Date(a.startDate).getTime())[0];
              const portfolio = lastContract ? portfolios.find(pf => pf.id === lastContract.portfolioId) : null;
              const assignedUser = portfolio ? users.find(u => u.id === portfolio.assignedTo) : null;
              const isMetano = p.service === 'METANO' || (!p.service && !!p.remi);
              const isAdminOrBO = user.role === 'admin' || user.role === 'backoffice';
              
              return (
                <tr key={p.id} onClick={() => onViewPdp && onViewPdp(p)} className="hover:bg-emerald-50/20 transition-colors group cursor-pointer">
                  <td className="p-6">
                    <div className="flex items-center gap-2">
                      <span className={`px-2 py-0.5 rounded-md text-[9px] font-black uppercase tracking-wider flex items-center gap-1 ${
                        isMetano ? 'bg-blue-100 text-blue-800' : 'bg-yellow-100 text-yellow-800'
                      }`}>
                        {isMetano ? <Flame size={10} /> : <Zap size={10} />}
                        {isMetano ? 'METANO' : 'POWER'}
                      </span>
                      <span className="font-mono font-black text-emerald-800">{p.pdpCode}</span>
                    </div>
                  </td>
                  <td className="p-6 font-mono font-bold text-slate-700 text-xs">
                    {p.matricola || '---'}
                  </td>
                  <td className="p-6 text-sm font-bold">
                     <div className="flex items-center gap-2" onClick={(e) => e.stopPropagation()}>
                       <AddressLink address={p.address} showIconOnly className="text-blue-500" />
                       <div>
                         <p className="font-bold text-slate-800">{p.city || '---'} ({p.province || '--'})</p>
                         <p className="text-[10px] text-slate-400">{p.street || p.address || '---'}</p>
                       </div>
                     </div>
                  </td>
                  <td className="p-6 text-xs">
                    {isMetano ? (
                      <div className="space-y-0.5">
                        <p className="font-bold text-slate-700"><span className="text-slate-400 font-medium">REMI:</span> {p.remi || '---'}</p>
                        <p className="font-bold text-blue-700"><span className="text-slate-400 font-medium">Portata:</span> {p.portata || '---'}</p>
                      </div>
                    ) : (
                      <div className="space-y-0.5">
                        <p className="font-bold text-emerald-700">
                          <span className="text-slate-400 font-medium">Pot Imp/Disp:</span> {p.potenzaImpegnata || 0} / {p.potenzaDisponibile || 0} kWp
                        </p>
                        <div className="flex flex-wrap items-center gap-2 text-[10px] text-slate-500">
                          {p.tensione && <span><strong>Tensione:</strong> {formatTensione(p.tensione)} V</span>}
                          {(p.cabinaPrimaria || p.technicalSpecs) && <span><strong>CP:</strong> {p.cabinaPrimaria || p.technicalSpecs}</span>}
                        </div>
                      </div>
                    )}
                  </td>
                  {isAdminOrBO && (
                    <td className="p-6 text-right">
                      <p className="text-lg font-black text-emerald-700 leading-none">{assignedUser?.userCode || '---'}</p>
                      <p className="text-[10px] font-bold text-slate-400 uppercase tracking-tighter mt-1">{assignedUser?.name || '---'}</p>
                    </td>
                  )}
                  <td className="p-6 text-right space-x-2" onClick={(e) => e.stopPropagation()}>
                    <button onClick={() => onViewPdp && onViewPdp(p)} className="text-slate-400 hover:text-emerald-600 p-2" title="Visualizza Scheda"><Eye size={16}/></button>
                    <button onClick={() => { 
                      const parsed = splitAddress(p.address);
                      setForm({
                        ...p,
                        street: p.street || parsed?.street || '',
                        cap: p.cap || parsed?.cap || '',
                        city: p.city || parsed?.city || '',
                        province: p.province || parsed?.province || ''
                      }); 
                      setEditingId(p.id); 
                      setIsFormOpen(true); 
                    }} className="text-slate-400 hover:text-blue-600 p-2" title="Modifica"><Edit2 size={16}/></button>
                    <button onClick={() => setDeleteConfirmId(p.id)} className="text-slate-400 hover:text-red-600 p-2" title="Elimina"><X size={18}/></button>
                  </td>
                </tr>
              );
            })}
            {filteredPdps.length === 0 && (
              <tr>
                <td colSpan={6} className="p-12 text-center text-slate-400 font-bold uppercase tracking-wider">
                  Nessun PDP trovato corrispondente alla ricerca
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      <ConfirmModal
        isOpen={!!deleteConfirmId}
        title="Conferma Eliminazione"
        message="Sei sicuro di voler eliminare questo punto di prelievo (PDP)? Questa operazione non può essere annullata."
        onConfirm={() => {
          setData((prev: PDP[]) => prev.filter(i => i.id !== deleteConfirmId));
          setDeleteConfirmId(null);
        }}
        onCancel={() => setDeleteConfirmId(null)}
      />
    </div>
  );
};

const ContractView: React.FC<any> = ({ data, setData, portfolios, setPortfolios, pdps, cabine, setCabine, setPdps, users, partners = [], appartenenze = [], campagne = [], user, exportFn, onOpenChat, getUnreadCount, prefilledPortfolioId, onClearPrefill, onViewPdp }) => {
  const isAdminOrBO = user.role === 'admin' || user.role === 'backoffice';
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [selectedPortfolio, setSelectedPortfolio] = useState<Portfolio | null>(null);
  const [selectedPDP, setSelectedPDP] = useState<PDP | null>(null);
  const [deleteConfirmId, setDeleteConfirmId] = useState<string | null>(null);
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

  const getDayBefore = (dateStr: string): string => {
    if (!dateStr) return '';
    const [y, m, d] = dateStr.split('-').map(Number);
    const date = new Date(Date.UTC(y, m - 1, d));
    date.setUTCDate(date.getUTCDate() - 1);
    return date.toISOString().split('T')[0];
  };

  const mostRecentContractForSelectedPdp = useMemo(() => {
    if (!form.pdpId || form.pdpId.startsWith('NEW_')) return null;
    const pdpContracts = data.filter((c: any) => c.pdpId === form.pdpId && c.id !== editingId);
    if (pdpContracts.length === 0) return null;
    return pdpContracts.sort((a, b) => {
      const dateA = a.endDate || '';
      const dateB = b.endDate || '';
      return dateB.localeCompare(dateA);
    })[0];
  }, [form.pdpId, data, editingId]);

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
      const camp = campagne.find((cp: Campagna) => cp.id === c.campagnaId);
      
      const matchesStatus = statusFilter === 'all' || c.status === statusFilter;
      
      const searchStr = `${port?.businessName || ''} ${pdp?.pdpCode || ''} ${pdp?.address || ''} ${pdp?.street || ''} ${pdp?.city || ''} ${pdp?.province || ''} ${c.address || ''} ${c.street || ''} ${c.city || ''} ${c.province || ''} ${c.fornitore || ''} ${c.tariffa || ''} ${c.contractType || ''} ${c.service || ''} ${camp?.code || ''} ${camp?.name || ''}`.toLowerCase();
      const matchesSearch = searchStr.includes(searchTerm.toLowerCase());
      
      return matchesStatus && matchesSearch;
    });
  }, [data, searchTerm, statusFilter, portfolios, pdps, campagne]);

  const handleSave = (e?: React.FormEvent, targetStatus?: ContractStatus) => {
    if (e && e.preventDefault) e.preventDefault();
    if (!form.portfolioId || !form.pdpId) return alert("Associazione Cliente e PDP obbligatoria.");
    
    // Validazione PDP
    if (form.contractType === 'ALLACCIO') {
      // Per ALLACCIO può essere "da assegnare" o 14 caratteri
      const code = form.pdpId.startsWith('NEW_') ? form.pdpId.replace('NEW_', '') : form.pdpId;
      if (code.toLowerCase() !== 'da assegnare' && code.length !== 14) {
        return alert("Per i nuovi allacci il codice PDP deve essere 'da assegnare' o di 14 caratteri.");
      }
    } else {
      // Per gli altri casi deve essere 14 caratteri
      const code = form.pdpId.startsWith('NEW_') ? form.pdpId.replace('NEW_', '') : form.pdpId;
      // Se è un PDP esistente, cerchiamo il codice reale
      const existingPdp = pdps.find((p:any) => p.id === form.pdpId);
      const finalCode = existingPdp ? existingPdp.pdpCode : code;
      
      if (finalCode.length !== 14) {
        return alert("Il codice PDP deve essere obbligatoriamente di 14 caratteri.");
      }
    }

    if (form.service === 'METANO' && form.portata) {
      if (!form.portata.toUpperCase().startsWith('G')) {
        return alert("Errore: La portata per il servizio METANO deve iniziare obbligatoriamente con la lettera 'G' maiuscola (es. G4, G6, G10).");
      }
    }

    const ts = getTimestamp();
    let finalPdpId = form.pdpId;

    const finalAddress = form.address || composeAddress(form.street, form.cap, form.city, form.province);
    const parsed = splitAddress(finalAddress);
    const street = form.street || parsed?.street || '';
    const cap = form.cap || parsed?.cap || '';
    const city = form.city || parsed?.city || '';
    const province = (form.province || parsed?.province || '').toUpperCase();

    const pdpUpdates: Partial<PDP> = {
      service: form.service,
      address: finalAddress,
      street,
      cap,
      city,
      province,
      matricola: form.matricola ? form.matricola.toUpperCase().trim() : '',
      updatedAt: ts
    };

    if (form.service === 'POWER') {
      const cpName = form.cabinaPrimaria || form.technicalSpecs || '';
      let cId = undefined;
      if (cpName) {
        let ex = cabine.find((c: any) => c.name.toLowerCase() === cpName.toLowerCase());
        if (!ex) {
          ex = { id: Math.random().toString(36).substr(2, 9), name: cpName };
          setCabine((prev: any) => [...prev, ex]);
        }
        cId = ex.id;
      }
      pdpUpdates.potenzaImpegnata = form.potenzaImpegnata || 0;
      pdpUpdates.potenzaDisponibile = form.potenzaDisponibile || 0;
      pdpUpdates.technicalSpecs = cpName;
      pdpUpdates.cabinaPrimaria = cpName;
      pdpUpdates.cabinaPrimariaId = cId;
      pdpUpdates.tensione = formatTensione(form.tensione);
    } else {
      pdpUpdates.remi = form.remi ? form.remi.toUpperCase().trim() : '';
      pdpUpdates.portata = form.portata ? form.portata.toUpperCase().trim() : '';
    }

    if (finalPdpId && finalPdpId.startsWith('NEW_')) {
      const code = finalPdpId.replace('NEW_', '');
      const newPDP: PDP = {
        id: Math.random().toString(36).substr(2, 9),
        pdpCode: code,
        potenzaImpegnata: 0,
        potenzaDisponibile: 0,
        assignedTo: user.id,
        createdAt: ts,
        ...pdpUpdates
      } as PDP;
      setPdps((prev: PDP[]) => [...prev, newPDP]);
      finalPdpId = newPDP.id;
    } else if (finalPdpId) {
      setPdps((prev: PDP[]) => prev.map(p => p.id === finalPdpId ? { ...p, ...pdpUpdates } : p));
    }

    // Regole stato per Agenti:
    // Possono salvare in bozza o trasmettere (diventa 'trasmesso')
    // Se correggono un non conforme, diventa 'corretto'
    let finalStatus: ContractStatus;
    if (isAdminOrBO) {
      finalStatus = targetStatus || form.status || 'bozza';
    } else {
      if (editingId) {
        const existing = data.find((c: any) => c.id === editingId);
        if (existing?.status === 'non conforme') {
          finalStatus = 'corretto';
        } else {
          finalStatus = targetStatus === 'trasmesso' ? 'trasmesso' : 'bozza';
        }
      } else {
        finalStatus = targetStatus === 'trasmesso' ? 'trasmesso' : 'bozza';
      }
    }

    const payload = { 
      ...form, 
      status: finalStatus,
      address: finalAddress,
      street,
      cap,
      city,
      province,
      pdpId: finalPdpId, 
      updatedAt: ts 
    };

    const adjustPreviousContracts = (contractsList: Contract[]) => {
      if (payload.status !== 'trasmesso' && payload.status !== 'attivo') return contractsList;
      if (!payload.startDate || !payload.pdpId) return contractsList;
      return contractsList.map(c => {
        if (c.pdpId === payload.pdpId && c.id !== editingId) {
          if (c.endDate && payload.startDate < c.endDate) {
            const newEndDate = getDayBefore(payload.startDate);
            return {
              ...c,
              endDate: newEndDate,
              updatedAt: ts
            };
          }
        }
        return c;
      });
    };

    if (editingId) {
      setData((prev: Contract[]) => {
        const updated = prev.map(c => c.id === editingId ? { ...c, ...payload } : c);
        return adjustPreviousContracts(updated);
      });
    } else {
      setData((prev: Contract[]) => {
        const newContract = { ...payload, id: Math.random().toString(36).substr(2, 9), assignedTo: user.id, createdAt: ts } as Contract;
        const updated = [...prev, newContract];
        return adjustPreviousContracts(updated);
      });
    }
    
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
        <div><h2 className="text-3xl font-black text-emerald-900 tracking-tighter">Forniture</h2><p className="text-slate-500 font-medium">Archivio forniture energetiche attivo</p></div>
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
            <option value="corretto">Corretto</option>
            <option value="KO">KO</option>
          </select>
          <button onClick={() => exportFn(data, 'Forniture')} className="p-3 text-slate-400 hover:text-slate-800 transition-colors"><FileSpreadsheet /></button>
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
          }} className="bg-emerald-600 text-white px-8 py-4 rounded-2xl font-black shadow-lg hover:scale-105 transition-all flex items-center gap-2"><Plus /> Nuova Fornitura</button>
        </div>
      </div>
      {isFormOpen && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm overflow-y-auto">
          <div className="bg-white w-full max-w-4xl rounded-[2.5rem] shadow-2xl overflow-hidden my-8 animate-in zoom-in-95 duration-300 flex flex-col max-h-[90vh]">
            <div className="p-6 md:p-8 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
              <h3 className="text-2xl font-black tracking-tight text-slate-800">
                {editingId ? 'Modifica Fornitura' : 'Nuova Fornitura'}
              </h3>
              <button type="button" onClick={() => setIsFormOpen(false)} className="text-slate-400 hover:text-slate-600 p-2 rounded-full hover:bg-slate-100 transition-all">
                <X size={20} />
              </button>
            </div>
            <form onSubmit={(e) => handleSave(e)} className="p-6 md:p-8 space-y-8 overflow-y-auto flex-1 text-left">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
            <SearchableSelect label="Cliente *" placeholder="Scegli cliente..." options={portfolios} selectedId={form.portfolioId} onSelect={(id:string) => setForm({...form, portfolioId: id})} displayFn={(p:any)=>p.businessName} searchFn={(p:any)=>`${p.businessName}`} subTextFn={(p:any)=>p.taxId} icon={<Users size={14}/>} />
            <SearchableSelect 
              label="PDP *" 
              placeholder="Scegli PDP o digita nuovo..." 
              options={pdps} 
              selectedId={form.pdpId} 
              onSelect={(id:string) => { 
                const p = pdps.find((x:any)=>x.id===id); 
                if(p) {
                  const s = splitAddress(p.address);
                  const street = p.street || s?.street || '';
                  const cap = p.cap || s?.cap || '';
                  const city = p.city || s?.city || '';
                  const province = p.province || s?.province || '';
                  const address = p.address || composeAddress(street, cap, city, province);
                  setForm(prev => ({
                    ...prev,
                    pdpId: id,
                    address,
                    street,
                    cap,
                    city,
                    province,
                    service: p.service || prev.service || 'POWER',
                    potenzaImpegnata: p.potenzaImpegnata ?? prev.potenzaImpegnata,
                    potenzaDisponibile: p.potenzaDisponibile ?? prev.potenzaDisponibile,
                    technicalSpecs: p.cabinaPrimaria || p.technicalSpecs || prev.technicalSpecs,
                    cabinaPrimaria: p.cabinaPrimaria || p.technicalSpecs || prev.cabinaPrimaria,
                    matricola: p.matricola || prev.matricola,
                    tensione: p.tensione || prev.tensione,
                    remi: p.remi || prev.remi,
                    portata: p.portata || prev.portata
                  }));
                } else {
                  setForm(prev => ({ ...prev, pdpId: id }));
                }
              }} 
              displayFn={(p:any)=>`${p.pdpCode} (${p.service || 'POWER'})`} 
              searchFn={(p:any)=>`${p.pdpCode} ${p.matricola || ''} ${p.remi || ''}`} 
              subTextFn={(p:any)=>`${p.address || ''} ${p.matricola ? '• Matr: ' + p.matricola : ''}`} 
              icon={<Zap size={14}/>} 
              allowCustom={true} 
              customLabel="Aggiungi nuovo PDP" 
              disabled={form.contractType === 'ALLACCIO'}
              maxLength={14}
            />
            
            {mostRecentContractForSelectedPdp && (
              <div className="col-span-full bg-amber-50 border border-amber-200 rounded-2xl p-4 flex flex-col gap-2 animate-in fade-in slide-in-from-top-1 duration-200">
                <div className="flex items-center gap-2 text-amber-800 font-bold text-sm">
                  <Clock size={16} className="text-amber-600 animate-pulse" />
                  <span>Info Contratto Precedente:</span>
                </div>
                <div className="text-xs text-amber-700 font-semibold space-y-1 pl-6">
                  <p>Il contratto più recente su questo PDP scade il: <strong className="text-sm font-black text-amber-900">{formatDate(mostRecentContractForSelectedPdp.endDate)}</strong>.</p>
                  {form.startDate && mostRecentContractForSelectedPdp.endDate && form.startDate < mostRecentContractForSelectedPdp.endDate && (
                    <p className="text-red-700 bg-red-50 border border-red-100 p-2.5 rounded-xl font-bold flex items-center gap-1.5 mt-1">
                      <AlertCircle size={14} className="shrink-0" />
                      Attenzione: la data di decorrenza inserita ({formatDate(form.startDate)}) è precedente alla scadenza del contratto precedente. Al salvataggio, la scadenza del precedente contratto sarà modificata automaticamente al <strong className="font-black underline">{formatDate(getDayBefore(form.startDate))}</strong>.
                    </p>
                  )}
                </div>
              </div>
            )}
            
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
                <AddressInputGroup 
                  label="Indirizzo Fornitura (Via, CAP Città, PROV)"
                  address={form.address || ''}
                  street={form.street || ''}
                  cap={form.cap || ''}
                  city={form.city || ''}
                  province={form.province || ''}
                  accentColor="emerald"
                  bgBox="bg-white"
                  extraHeaderAction={
                    form.portfolioId && (
                      <button 
                        type="button" 
                        onClick={() => {
                          const p = portfolios.find((x:any)=>x.id === form.portfolioId);
                          if (p) {
                            const parsed = splitAddress(p.legalAddress);
                            const street = p.street || parsed?.street || '';
                            const cap = p.cap || parsed?.cap || '';
                            const city = p.city || parsed?.city || '';
                            const province = p.province || parsed?.province || '';
                            const address = p.legalAddress || composeAddress(street, cap, city, province);
                            setForm(prev => ({ ...prev, address, street, cap, city, province }));
                          }
                        }}
                        className="text-[10px] font-black uppercase text-emerald-600 hover:text-emerald-700 flex items-center gap-1 bg-emerald-50 px-2.5 py-1 rounded-xl border border-emerald-100 hover:bg-emerald-100 transition-colors"
                      >
                        <Activity size={10} /> Usa Sede Legale
                      </button>
                    )
                  }
                  onChange={({ address, street, cap, city, province }) => {
                    setForm(prev => ({
                      ...prev,
                      address,
                      street,
                      cap,
                      city,
                      province
                    }));
                  }}
                />
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1">Matricola</label>
                  <input 
                    className="w-full border-2 p-3.5 rounded-2xl bg-white outline-none font-bold uppercase focus:border-emerald-500" 
                    placeholder="Matricola contatore"
                    value={form.matricola || ''} 
                    onChange={e => setForm({...form, matricola: e.target.value.toUpperCase()})} 
                  />
                </div>
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1">Volume Annuale</label>
                  <div className="relative">
                    <input type="number" className="w-full border-2 p-3.5 rounded-2xl bg-white outline-none font-bold pr-24" value={form.volume || ''} onChange={e => setForm({...form, volume: Number(e.target.value)})} />
                    <span className="absolute right-4 top-1/2 -translate-y-1/2 text-[10px] font-black text-emerald-600 bg-emerald-50 px-2 py-1 rounded-lg">{formatVolume(form.volume || 0)}</span>
                  </div>
                </div>
                {form.service === 'POWER' ? (
                  <>
                    <div className="grid grid-cols-2 gap-4">
                      <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Pot. Impegnata (KWp)</label><input type="number" step="0.1" className="w-full border-2 p-3.5 rounded-2xl bg-white outline-none font-bold" value={form.potenzaImpegnata || ''} onChange={e => setForm({...form, potenzaImpegnata: Number(e.target.value)})} /></div>
                      <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Pot. Disponibile (KWp)</label><input type="number" step="0.1" className="w-full border-2 p-3.5 rounded-2xl bg-white outline-none font-bold" value={form.potenzaDisponibile || ''} onChange={e => setForm({...form, potenzaDisponibile: Number(e.target.value)})} /></div>
                    </div>
                    <div className="grid grid-cols-2 gap-4">
                      <div className="space-y-1.5">
                        <label className="text-xs font-bold text-slate-500 uppercase ml-1">Tensione</label>
                        <select 
                          className="w-full border-2 p-3.5 rounded-2xl bg-white font-bold outline-none border-slate-100" 
                          value={formatTensione(form.tensione)} 
                          onChange={e => setForm({...form, tensione: e.target.value})}
                        >
                          <option value="">Seleziona Tensione...</option>
                          {TENSIONE_OPTIONS.map(opt => (
                            <option key={opt.value} value={opt.value}>{opt.label} V</option>
                          ))}
                        </select>
                      </div>
                      <div className="space-y-1.5">
                        <label className="text-xs font-bold text-slate-500 uppercase ml-1">Cabina Primaria</label>
                        <input className="w-full border-2 p-3.5 rounded-2xl bg-white outline-none font-bold border-slate-100" placeholder="CP ..." value={form.cabinaPrimaria || form.technicalSpecs || ''} onChange={e => setForm({...form, cabinaPrimaria: e.target.value, technicalSpecs: e.target.value})} />
                      </div>
                    </div>
                    <div className="grid grid-cols-2 gap-4">
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
                      <div className="space-y-1.5">
                        <label className="text-xs font-bold text-slate-500 uppercase ml-1">Stato</label>
                        {isAdminOrBO ? (
                          <select className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 font-bold outline-none" value={form.status} onChange={e => setForm({...form, status: e.target.value as any})}>
                            <option value="bozza">Bozza</option>
                            <option value="da firmare">Da Firmare</option>
                            <option value="trasmesso">Trasmesso</option>
                            <option value="attivo">Attivo</option>
                            <option value="non conforme">Non Conforme</option>
                            <option value="corretto">Corretto</option>
                            <option value="KO">KO</option>
                          </select>
                        ) : (
                          <div className="p-3.5 rounded-2xl bg-slate-100 border border-slate-200 flex items-center text-xs">
                            <span className="font-black uppercase text-slate-700">
                              {editingId && form.status === 'non conforme' ? 'In Correzione → Corretto' : (form.status || 'Bozza').toUpperCase()}
                            </span>
                          </div>
                        )}
                      </div>
                    </div>
                  </>
                ) : (
                  <>
                    <div className="grid grid-cols-2 gap-4">
                      <div className="space-y-1.5">
                        <label className="text-xs font-bold text-slate-500 uppercase ml-1">REMI</label>
                        <input className="w-full border-2 p-3.5 rounded-2xl bg-white font-bold uppercase outline-none" placeholder="Codice REMI" value={form.remi || ''} onChange={e => setForm({...form, remi: e.target.value.toUpperCase()})} />
                      </div>
                      <div className="space-y-1.5">
                        <label className="text-xs font-bold text-slate-500 uppercase ml-1 flex justify-between">
                          <span>Portata</span>
                          <span className="text-[10px] text-amber-600 font-bold">Inizia con 'G'</span>
                        </label>
                        <input className="w-full border-2 p-3.5 rounded-2xl bg-white font-bold uppercase outline-none" placeholder="Es: G4, G6..." value={form.portata || ''} onChange={e => setForm({...form, portata: e.target.value.toUpperCase()})} />
                      </div>
                    </div>
                    <div className="space-y-1.5">
                      <label className="text-xs font-bold text-slate-500 uppercase ml-1">Stato</label>
                      {isAdminOrBO ? (
                        <select className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 font-bold outline-none" value={form.status} onChange={e => setForm({...form, status: e.target.value as any})}>
                          <option value="bozza">Bozza</option>
                          <option value="da firmare">Da Firmare</option>
                          <option value="trasmesso">Trasmesso</option>
                          <option value="attivo">Attivo</option>
                          <option value="non conforme">Non Conforme</option>
                          <option value="corretto">Corretto</option>
                          <option value="KO">KO</option>
                        </select>
                      ) : (
                        <div className="p-3.5 rounded-2xl bg-slate-100 border border-slate-200 flex items-center text-xs">
                          <span className="font-black uppercase text-slate-700">
                            {editingId && form.status === 'non conforme' ? 'In Correzione → Corretto' : (form.status || 'Bozza').toUpperCase()}
                          </span>
                        </div>
                      )}
                    </div>
                  </>
                )}
              </div>
            </div>

            <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Fornitore</label>
              <select className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 font-bold outline-none" value={form.fornitore} onChange={e => setForm({...form, fornitore: e.target.value as Fornitore})}>
                {FORNITORI.map(f => <option key={f} value={f}>{f}</option>)}
              </select>
            </div>
            <div className="space-y-1.5">
              <SearchableSelect 
                label="Campagna" 
                placeholder="Seleziona campagna (opzionale)..." 
                options={campagne} 
                selectedId={form.campagnaId} 
                onSelect={(id: string) => setForm({...form, campagnaId: id || undefined})} 
                displayFn={(c: Campagna) => `${c.code} - ${c.name}`} 
                searchFn={(c: Campagna) => `${c.code} ${c.name} ${c.description}`} 
                subTextFn={(c: Campagna) => c.exclusiveBenefits || c.description} 
                icon={<Megaphone size={14}/>} 
                allowClear={true}
              />
            </div>
            <div className="col-span-full grid grid-cols-2 gap-4">
              <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Data Creazione</label><input type="date" disabled={user.role === 'utente'} className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold disabled:opacity-50 text-xs" value={form.creationDate || ''} onChange={e => { const d = e.target.value; const start = calculateActivationDate(d); setForm({...form, creationDate: d, startDate: start, endDate: calculateEndDate(start, form.durationMonths || 12)}); }} /></div>
              <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Data Attivazione</label><input type="date" className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold text-xs" value={form.startDate || ''} onChange={e => { const d = e.target.value; setForm({...form, startDate: d, endDate: calculateEndDate(d, form.durationMonths || 12)}); }} /></div>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Durata (Mesi)</label><input type="number" className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold text-xs" value={form.durationMonths || 12} onChange={e => { const m = Number(e.target.value); setForm({...form, durationMonths: m, endDate: calculateEndDate(form.startDate || '', m)}); }} /></div>
              <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Fine Contratto</label><input type="date" disabled className="w-full border-2 p-3.5 rounded-2xl bg-slate-200 outline-none font-bold opacity-70 text-xs" value={form.endDate || ''} /></div>
              <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Tariffa</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold text-xs" value={form.tariffa || ''} onChange={e => setForm({...form, tariffa: e.target.value})} /></div>
            </div>
            <div className="col-span-full"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Allegati Contratto</label>
              <FileUploader attachments={form.attachments || []} canEdit={true} onUpload={files => setForm({...form, attachments: [...(form.attachments || []), ...files]})} onRemove={id => setForm({...form, attachments: form.attachments?.filter(a => a.id !== id)})} />
            </div>
          </div>
          {isAdminOrBO ? (
            <div className="flex justify-end gap-4 border-t pt-8">
              <button type="button" onClick={() => setIsFormOpen(false)} className="font-bold text-slate-400">Annulla</button>
              <button type="submit" className="bg-emerald-900 text-white px-12 py-5 rounded-2xl font-black shadow-xl cursor-pointer hover:bg-emerald-950 transition-all">
                {editingId ? 'Aggiorna Fornitura' : 'Salva Fornitura'}
              </button>
            </div>
          ) : (
            editingId && form.status === 'non conforme' ? (
              <div className="flex justify-end items-center gap-4 border-t pt-8">
                <button type="button" onClick={() => setIsFormOpen(false)} className="font-bold text-slate-400 mr-2">Annulla</button>
                <button 
                  type="button" 
                  onClick={(e) => handleSave(e, 'corretto')} 
                  className="bg-teal-600 hover:bg-teal-700 text-white px-10 py-4 rounded-2xl font-black shadow-xl flex items-center gap-2 hover:scale-105 transition-all text-sm cursor-pointer"
                >
                  <CheckCircle2 size={18} /> Salva Correzione (Corretto)
                </button>
              </div>
            ) : (
              <div className="flex justify-end items-center gap-3 border-t pt-8">
                <button type="button" onClick={() => setIsFormOpen(false)} className="font-bold text-slate-400 mr-2">Annulla</button>
                <button 
                  type="button" 
                  onClick={(e) => handleSave(e, 'bozza')} 
                  className="bg-slate-200 hover:bg-slate-300 text-slate-800 px-7 py-4 rounded-2xl font-bold transition-all text-sm cursor-pointer"
                >
                  Salva in Bozza
                </button>
                <button 
                  type="button" 
                  onClick={(e) => handleSave(e, 'trasmesso')} 
                  className="bg-emerald-900 hover:bg-emerald-950 text-white px-9 py-4 rounded-2xl font-black shadow-xl flex items-center gap-2 hover:scale-105 transition-all text-sm cursor-pointer"
                >
                  <Send size={18} /> Trasmetti
                </button>
              </div>
            )
          )}
        </form>
          </div>
        </div>
      )}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        {filteredData.map((c: Contract) => {
          const port = portfolios.find(p => p.id === c.portfolioId);
          const pdp = pdps.find(p => p.id === c.pdpId);
          const assignedUser = port ? [...users, ...partners].find(u => u.id === port.assignedTo) : null;
          const camp = campagne.find((cp: Campagna) => cp.id === c.campagnaId);
          const canEdit = isAdminOrBO || c.status === 'bozza' || c.status === 'non conforme';
          const canDelete = isAdminOrBO || c.status === 'bozza';

          const getStatusBadgeStyle = (st: ContractStatus) => {
            if (['attivo', 'trasmesso'].includes(st)) return 'bg-emerald-100 text-emerald-700';
            if (st === 'corretto') return 'bg-teal-100 text-teal-800 border border-teal-300';
            if (st === 'KO') return 'bg-red-100 text-red-700';
            if (st === 'non conforme') return 'bg-purple-100 text-purple-700';
            return 'bg-yellow-100 text-yellow-700';
          };
          
          return (
            <div key={c.id} className="bg-white p-6 rounded-[2.5rem] border border-slate-100 shadow-sm hover:shadow-2xl transition-all relative group overflow-hidden">
              <div className={`absolute top-0 right-0 px-6 py-2 rounded-bl-2xl text-[10px] font-black uppercase tracking-widest flex items-center gap-2 ${getStatusBadgeStyle(c.status)}`}>
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
                       <div className="flex flex-wrap items-center gap-2 mt-1">
                          <p onClick={() => pdp && onViewPdp && onViewPdp(pdp)} className="text-[11px] font-black text-emerald-600 font-mono bg-emerald-50 px-2 py-0.5 rounded-md cursor-pointer hover:bg-emerald-100 transition-colors">{pdp?.pdpCode || 'CODICE PDP N/D'}</p>
                          <p className="text-[10px] font-bold text-slate-400">| {c.fornitore} {c.tariffOption || c.tariffa ? `- ${c.tariffOption || c.tariffa}` : ''}</p>
                          {camp && (
                            <span className="inline-flex items-center gap-1 bg-purple-50 text-purple-800 border border-purple-200 px-2 py-0.5 rounded-md text-[10px] font-black" title={camp.name}>
                              <Megaphone size={10} className="text-purple-600" /> {camp.code} - {camp.name}
                            </span>
                          )}
                       </div>
                     </div>
                     <div className="text-right shrink-0 pt-4">
                        {isAdminOrBO && (
                          <>
                            <p className="text-lg font-black text-emerald-700 leading-none">{assignedUser?.userCode || '---'}</p>
                            <p className="text-[10px] font-bold text-slate-400 uppercase tracking-tighter mt-1">{assignedUser?.name || '---'}</p>
                          </>
                        )}
                        <div className="opacity-0 group-hover:opacity-100 transition-opacity flex gap-1 justify-end mt-2">
                           <button onClick={() => onOpenChat(c.id)} className="p-2 text-slate-300 hover:text-blue-500 relative">
                              <MessageSquare size={16}/>
                              {getUnreadCount(c.id) > 0 && (
                                <span className="absolute top-0 right-0 w-2 h-2 bg-red-500 rounded-full border-2 border-white"></span>
                              )}
                           </button>
                            {canEdit && (
                              <button onClick={() => { 
                                const p = pdps.find((x: any) => x.id === c.pdpId);
                                const rawAddr = c.address || p?.address || '';
                                const s = splitAddress(rawAddr);
                                const street = c.street || p?.street || s?.street || '';
                                const cap = c.cap || p?.cap || s?.cap || '';
                                const city = c.city || p?.city || s?.city || '';
                                const province = c.province || p?.province || s?.province || '';
                                const address = rawAddr || composeAddress(street, cap, city, province);
                                setForm({
                                  ...c,
                                  address,
                                  street,
                                  cap,
                                  city,
                                  province,
                                  matricola: c.matricola || p?.matricola || '',
                                  cabinaPrimaria: c.cabinaPrimaria || p?.cabinaPrimaria || p?.technicalSpecs || '',
                                  technicalSpecs: c.technicalSpecs || p?.cabinaPrimaria || p?.technicalSpecs || '',
                                  tensione: c.tensione || p?.tensione || '',
                                  remi: c.remi || p?.remi || '',
                                  portata: c.portata || p?.portata || ''
                                });
                                setEditingId(c.id); 
                                setIsFormOpen(true); 
                              }} className="p-2 text-slate-300 hover:text-emerald-500" title={c.status === 'non conforme' ? "Correggi contratto non conforme" : "Modifica bozza"}><Edit2 size={16}/></button>
                            )}
                            {canDelete && (
                              <button onClick={() => setDeleteConfirmId(c.id)} className="p-2 text-slate-300 hover:text-red-500" title="Elimina bozza"><X size={18}/></button>
                            )}
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
          partners={partners}
          appartenenze={appartenenze}
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

      <ConfirmModal
        isOpen={!!deleteConfirmId}
        title="Conferma Eliminazione"
        message="Sei sicuro di voler eliminare questo contratto? Questa operazione non può essere annullata."
        onConfirm={() => {
          setData((prev: Contract[]) => prev.filter(i => i.id !== deleteConfirmId));
          setDeleteConfirmId(null);
        }}
        onCancel={() => setDeleteConfirmId(null)}
      />
    </div>
  );
};

const TelefonicoView: React.FC<{
  data: ContrattoTelefonico[];
  setData: React.Dispatch<React.SetStateAction<ContrattoTelefonico[]>>;
  portfolios: Portfolio[];
  contracts: Contract[];
  pdps: PDP[];
  users: User[];
  campagne?: Campagna[];
  user: User;
  exportFn: (data: any[], fileName: string) => void;
}> = ({ data, setData, portfolios, contracts, pdps, users, campagne = [], user, exportFn }) => {
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [deleteConfirmId, setDeleteConfirmId] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState<string>('all');
  const [gestoreFilter, setGestoreFilter] = useState<string>('all');

  // GESTORE in strictly alphabetical order: FASTWEB, TIM, VERY, VODAFONE, WIND
  const GESTORI: TelefonicoGestore[] = ['FASTWEB', 'TIM', 'VERY', 'VODAFONE', 'WIND'];
  const TIPI_OPERAZIONE: TelefonicoTipoOperazione[] = ['Cambio Gestore', 'Nuovo contratto', 'Voltura'];
  const STATI: ContractStatus[] = ['bozza', 'da firmare', 'trasmesso', 'attivo', 'non conforme', 'corretto', 'KO'];
  const isAdminOrBO = user.role === 'admin' || user.role === 'backoffice';

  const initialForm: Partial<ContrattoTelefonico> = {
    tipoOperazione: 'Nuovo contratto',
    gestore: 'FASTWEB',
    status: 'bozza',
    lineeFisse: 0,
    lineeMobili: 0,
    address: '',
    street: '',
    cap: '',
    city: '',
    province: '',
    notes: '',
    campagnaId: '',
    attachments: []
  };

  const [form, setForm] = useState<Partial<ContrattoTelefonico>>(initialForm);

  // Sedi che trovi impegnate nei contratti e nel portafoglio legate al Cliente selezionato
  const availableSedi = useMemo(() => {
    if (!form.portfolioId) return [];
    const client = portfolios.find(p => p.id === form.portfolioId);
    const list: { label: string; address: string }[] = [];
    const seen = new Set<string>();

    if (client?.legalAddress && client.legalAddress.trim()) {
      const addr = client.legalAddress.trim();
      list.push({ label: 'Sede Legale Portafoglio', address: addr });
      seen.add(addr.toLowerCase());
    }

    contracts.filter(c => c.portfolioId === form.portfolioId).forEach(c => {
      if (c.address && c.address.trim()) {
        const addr = c.address.trim();
        if (!seen.has(addr.toLowerCase())) {
          list.push({ label: `Contratto ${c.contractNumber || c.service || 'Energia'}`, address: addr });
          seen.add(addr.toLowerCase());
        }
      }
    });

    const clientPdpIds = new Set(contracts.filter(c => c.portfolioId === form.portfolioId).map(c => c.pdpId));
    pdps.filter(p => clientPdpIds.has(p.id)).forEach(p => {
      if (p.address && p.address.trim()) {
        const addr = p.address.trim();
        if (!seen.has(addr.toLowerCase())) {
          list.push({ label: `Sede PDP (${p.pdpCode})`, address: addr });
          seen.add(addr.toLowerCase());
        }
      }
    });

    return list;
  }, [form.portfolioId, portfolios, contracts, pdps]);

  const handleSave = (e?: React.FormEvent, targetStatus?: ContractStatus) => {
    if (e && e.preventDefault) e.preventDefault();
    if (!form.portfolioId) {
      alert("Seleziona obbligatoriamente un Cliente dal menu.");
      return;
    }

    const ts = getTimestamp();
    const finalAddress = form.address || composeAddress(form.street, form.cap, form.city, form.province);
    const spl = splitAddress(finalAddress);

    let finalStatus: ContractStatus;
    if (isAdminOrBO) {
      finalStatus = targetStatus || form.status || 'bozza';
    } else {
      if (editingId) {
        const existing = data.find(c => c.id === editingId);
        if (existing?.status === 'non conforme') {
          finalStatus = 'corretto';
        } else {
          finalStatus = targetStatus === 'trasmesso' ? 'trasmesso' : 'bozza';
        }
      } else {
        finalStatus = targetStatus === 'trasmesso' ? 'trasmesso' : 'bozza';
      }
    }

    const payload: ContrattoTelefonico = {
      id: editingId || Math.random().toString(36).substr(2, 9),
      portfolioId: form.portfolioId,
      tipoOperazione: form.tipoOperazione || 'Nuovo contratto',
      gestore: form.gestore || 'FASTWEB',
      address: finalAddress,
      street: form.street || spl?.street || '',
      cap: form.cap || spl?.cap || '',
      city: form.city || spl?.city || '',
      province: (form.province || spl?.province || '').toUpperCase(),
      lineeFisse: Math.max(0, parseInt(String(form.lineeFisse || 0), 10) || 0),
      lineeMobili: Math.max(0, parseInt(String(form.lineeMobili || 0), 10) || 0),
      status: finalStatus,
      notes: form.notes || '',
      campagnaId: form.campagnaId || undefined,
      attachments: form.attachments || [],
      assignedTo: form.assignedTo || user.id,
      createdAt: form.createdAt || ts,
      updatedAt: ts
    };

    if (editingId) {
      setData(prev => prev.map(c => c.id === editingId ? payload : c));
    } else {
      setData(prev => [payload, ...prev]);
    }

    setIsFormOpen(false);
    setEditingId(null);
    setForm(initialForm);
  };

  const startEdit = (item: ContrattoTelefonico) => {
    const spl = splitAddress(item.address || '');
    setForm({
      ...item,
      street: item.street || spl?.street || '',
      cap: item.cap || spl?.cap || '',
      city: item.city || spl?.city || '',
      province: item.province || spl?.province || ''
    });
    setEditingId(item.id);
    setIsFormOpen(true);
  };

  const filteredData = useMemo(() => {
    return data.filter(c => {
      const port = portfolios.find(p => p.id === c.portfolioId);
      const camp = campagne.find(cp => cp.id === c.campagnaId);
      const matchesStatus = statusFilter === 'all' || c.status === statusFilter;
      const matchesGestore = gestoreFilter === 'all' || c.gestore === gestoreFilter;
      const s = searchTerm.toLowerCase();
      const matchesSearch = !searchTerm || 
        port?.businessName?.toLowerCase().includes(s) ||
        port?.taxId?.toLowerCase().includes(s) ||
        c.gestore?.toLowerCase().includes(s) ||
        c.tipoOperazione?.toLowerCase().includes(s) ||
        c.address?.toLowerCase().includes(s) ||
        c.street?.toLowerCase().includes(s) ||
        c.city?.toLowerCase().includes(s) ||
        c.province?.toLowerCase().includes(s) ||
        c.notes?.toLowerCase().includes(s) ||
        camp?.code?.toLowerCase().includes(s) ||
        camp?.name?.toLowerCase().includes(s);
      return matchesStatus && matchesGestore && matchesSearch;
    });
  }, [data, portfolios, campagne, statusFilter, gestoreFilter, searchTerm]);

  const kpis = useMemo(() => {
    const total = data.length;
    const totalFisse = data.reduce((acc, c) => acc + (c.lineeFisse || 0), 0);
    const totalMobili = data.reduce((acc, c) => acc + (c.lineeMobili || 0), 0);
    const attivi = data.filter(c => c.status === 'attivo').length;
    return { total, totalFisse, totalMobili, attivi };
  }, [data]);

  const getGestoreBadge = (gestore: TelefonicoGestore) => {
    switch (gestore) {
      case 'FASTWEB':
        return 'bg-amber-100 text-amber-900 border-amber-300';
      case 'TIM':
        return 'bg-blue-100 text-blue-900 border-blue-300';
      case 'VERY':
        return 'bg-emerald-100 text-emerald-900 border-emerald-300';
      case 'VODAFONE':
        return 'bg-red-100 text-red-900 border-red-300';
      case 'WIND':
        return 'bg-orange-100 text-orange-900 border-orange-300';
      default:
        return 'bg-slate-100 text-slate-800 border-slate-200';
    }
  };

  const getStatusBadge = (status: ContractStatus) => {
    switch (status) {
      case 'attivo':
        return 'bg-emerald-100 text-emerald-800 border-emerald-200';
      case 'trasmesso':
        return 'bg-blue-100 text-blue-800 border-blue-200';
      case 'corretto':
        return 'bg-teal-100 text-teal-800 border-teal-300';
      case 'da firmare':
        return 'bg-yellow-100 text-yellow-800 border-yellow-200';
      case 'bozza':
        return 'bg-slate-100 text-slate-700 border-slate-200';
      case 'non conforme':
        return 'bg-purple-100 text-purple-800 border-purple-200';
      case 'KO':
        return 'bg-red-100 text-red-800 border-red-200';
      default:
        return 'bg-slate-100 text-slate-700 border-slate-200';
    }
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-300">
      <header className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <div className="flex items-center gap-2">
            <span className="bg-blue-100 text-blue-800 text-[10px] font-black uppercase px-2.5 py-0.5 rounded-full flex items-center gap-1">
              <Phone size={10} /> Telefonia
            </span>
          </div>
          <h2 className="text-3xl font-black text-slate-900 tracking-tight mt-1">Telefonico</h2>
          <p className="text-slate-500 font-medium italic text-xs">Gestione contratti di telefonia fissa e mobile per i clienti</p>
        </div>
        <div className="flex gap-3">
          <button 
            onClick={() => exportFn(data, 'Contratti_Telefonici')} 
            className="p-3 text-slate-400 hover:text-slate-800 bg-white border border-slate-100 rounded-2xl shadow-sm transition-all"
            title="Esporta in Excel"
          >
            <FileSpreadsheet size={20} />
          </button>
          <button 
            onClick={() => {
              setForm(initialForm);
              setEditingId(null);
              setIsFormOpen(true);
            }} 
            className="bg-blue-600 hover:bg-blue-700 text-white px-6 py-3 rounded-2xl font-black flex items-center gap-2 shadow-lg shadow-blue-500/20 hover:scale-105 transition-all text-sm"
          >
            <Plus size={18} /> Nuovo Contratto Telefonico
          </button>
        </div>
      </header>

      {/* KPI Cards */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        <div className="bg-white p-5 rounded-[1.75rem] border border-slate-100 shadow-sm flex items-center gap-4">
          <div className="p-3.5 bg-sky-50 rounded-2xl text-sky-500">
            <Phone size={22} />
          </div>
          <div>
            <p className="text-[10px] font-black uppercase tracking-wider text-slate-400">Totale Contratti</p>
            <p className="text-2xl font-black text-slate-900">{kpis.total}</p>
          </div>
        </div>

        <div className="bg-white p-5 rounded-[1.75rem] border border-slate-100 shadow-sm flex items-center gap-4">
          <div className="p-3.5 bg-emerald-50 rounded-2xl text-emerald-600">
            <CheckCircle2 size={22} />
          </div>
          <div>
            <p className="text-[10px] font-black uppercase tracking-wider text-slate-400">Contratti Attivi</p>
            <p className="text-2xl font-black text-emerald-700">{kpis.attivi}</p>
          </div>
        </div>

        <div className="bg-white p-5 rounded-[1.75rem] border border-slate-100 shadow-sm flex items-center gap-4">
          <div className="p-3.5 bg-indigo-50 rounded-2xl text-indigo-600">
            <Hash size={22} />
          </div>
          <div>
            <p className="text-[10px] font-black uppercase tracking-wider text-slate-400">Linee Fisse Totali</p>
            <p className="text-2xl font-black text-indigo-900">{kpis.totalFisse}</p>
          </div>
        </div>

        <div className="bg-white p-5 rounded-[1.75rem] border border-slate-100 shadow-sm flex items-center gap-4">
          <div className="p-3.5 bg-purple-50 rounded-2xl text-purple-600">
            <Activity size={22} />
          </div>
          <div>
            <p className="text-[10px] font-black uppercase tracking-wider text-slate-400">Linee Mobili Totali</p>
            <p className="text-2xl font-black text-purple-900">{kpis.totalMobili}</p>
          </div>
        </div>
      </div>

      {/* Form Inserimento/Modifica */}
      {isFormOpen && (
        <form onSubmit={handleSave} className="bg-white p-8 md:p-10 rounded-[2.5rem] border-2 border-blue-100 shadow-2xl space-y-6 animate-in slide-in-from-top-4 duration-300">
          <div className="flex justify-between items-center border-b border-slate-100 pb-4">
            <div>
              <h3 className="text-2xl font-black text-slate-900 tracking-tight">
                {editingId ? 'Modifica Contratto Telefonico' : 'Nuovo Contratto Telefonico'}
              </h3>
              <p className="text-xs text-slate-400 font-medium">Compila i dati dell'offerta telefonica</p>
            </div>
            <button 
              type="button" 
              onClick={() => { setIsFormOpen(false); setEditingId(null); }}
              className="p-2 text-slate-400 hover:text-slate-600 rounded-full hover:bg-slate-100 transition-colors"
            >
              <X size={22} />
            </button>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            {/* Menu Cliente */}
            <div className="space-y-1.5">
              <SearchableSelect 
                label="Cliente *" 
                placeholder="Seleziona cliente dal portafoglio..." 
                options={portfolios} 
                selectedId={form.portfolioId} 
                onSelect={(id: string) => {
                  setForm(prev => ({ ...prev, portfolioId: id }));
                }} 
                displayFn={(p: any) => p.businessName} 
                searchFn={(p: any) => p.businessName} 
                subTextFn={(p: any) => p.taxId} 
                icon={<Users size={14} />} 
              />
            </div>

            {/* Tipo Operazione */}
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1">Tipo Operazione *</label>
              <select 
                className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-blue-500 cursor-pointer"
                required
                value={form.tipoOperazione || 'Nuovo contratto'}
                onChange={e => setForm({ ...form, tipoOperazione: e.target.value as TelefonicoTipoOperazione })}
              >
                {TIPI_OPERAZIONE.map(t => (
                  <option key={t} value={t}>{t}</option>
                ))}
              </select>
            </div>

            {/* Sedi che trovi impegnate nei contratti e nel portafoglio legate al Cliente selezionato */}
            {form.portfolioId && availableSedi.length > 0 && (
              <div className="col-span-full space-y-2 p-5 bg-blue-50/80 border border-blue-200 rounded-2xl">
                <div className="flex items-center justify-between">
                  <label className="text-xs font-black text-blue-900 uppercase flex items-center gap-1.5">
                    <MapPin size={16} className="text-blue-600" />
                    Menu Sedi Note del Cliente (Copia Indirizzo)
                  </label>
                  <span className="text-[10px] font-bold text-blue-700 bg-white/70 px-2 py-0.5 rounded-md border border-blue-200">
                    {availableSedi.length} sedi disponibili
                  </span>
                </div>
                <select 
                  className="w-full border-2 border-blue-300 p-3 rounded-xl bg-white outline-none font-bold text-xs text-blue-950 focus:border-blue-600 cursor-pointer shadow-sm"
                  value=""
                  onChange={(e) => {
                    const chosen = e.target.value;
                    if (!chosen) return;
                    const spl = splitAddress(chosen);
                    setForm(prev => ({
                      ...prev,
                      address: chosen,
                      street: spl?.street || '',
                      cap: spl?.cap || '',
                      city: spl?.city || '',
                      province: spl?.province || ''
                    }));
                  }}
                >
                  <option value="">-- Seleziona una sede per copiare istantaneamente l'indirizzo --</option>
                  {availableSedi.map((s, idx) => (
                    <option key={idx} value={s.address}>
                      {s.label}: {s.address}
                    </option>
                  ))}
                </select>
                <p className="text-[11px] text-blue-700 font-medium">
                  💡 Scegliendo una sede, l'indirizzo viene copiato nei campi sottostanti. Puoi modificarlo liberamente per nuovi riferimenti.
                </p>
              </div>
            )}

            {/* Campi Indirizzo Confluenti (Via, CAP, Comune, PROV) */}
            <div className="col-span-full">
              <AddressInputGroup 
                label="Indirizzo Sede / Impianto (Via, CAP Città, PROV)"
                address={form.address || ''}
                street={form.street || ''}
                cap={form.cap || ''}
                city={form.city || ''}
                province={form.province || ''}
                accentColor="blue"
                onChange={({ address, street, cap, city, province }) => {
                  setForm(prev => ({
                    ...prev,
                    address,
                    street,
                    cap,
                    city,
                    province
                  }));
                }}
              />
            </div>

            {/* Menu Linee Fisse */}
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1">Linee Fisse (Numero)</label>
              <input 
                type="number" 
                min={0}
                step={1}
                className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-blue-500" 
                value={form.lineeFisse ?? 0}
                onChange={e => setForm({ ...form, lineeFisse: Math.max(0, parseInt(e.target.value, 10) || 0) })}
              />
            </div>

            {/* Menu Linee Mobili */}
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1">Linee Mobili (Numero)</label>
              <input 
                type="number" 
                min={0}
                step={1}
                className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-blue-500" 
                value={form.lineeMobili ?? 0}
                onChange={e => setForm({ ...form, lineeMobili: Math.max(0, parseInt(e.target.value, 10) || 0) })}
              />
            </div>

            {/* GESTORE (in ordine alfabetico: FASTWEB, TIM, VERY, VODAFONE, WIND) */}
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1">Gestore *</label>
              <select 
                className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-blue-500 cursor-pointer"
                required
                value={form.gestore || 'FASTWEB'}
                onChange={e => setForm({ ...form, gestore: e.target.value as TelefonicoGestore })}
              >
                {GESTORI.map(g => (
                  <option key={g} value={g}>{g}</option>
                ))}
              </select>
            </div>

            {/* Stato (uguale ai contratti) */}
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1">Stato</label>
              {isAdminOrBO ? (
                <select 
                  className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-blue-500 cursor-pointer"
                  value={form.status || 'bozza'}
                  onChange={e => setForm({ ...form, status: e.target.value as ContractStatus })}
                >
                  {STATI.map(s => (
                    <option key={s} value={s}>{s.toUpperCase()}</option>
                  ))}
                </select>
              ) : (
                <div className="p-3.5 rounded-2xl bg-slate-100 border border-slate-200 flex items-center text-xs">
                  <span className="font-black uppercase text-slate-700">
                    {editingId && form.status === 'non conforme' ? 'In Correzione → Corretto' : (form.status || 'Bozza').toUpperCase()}
                  </span>
                </div>
              )}
            </div>

            {/* Campagna e Assegnazione Utente */}
            <div>
              <SearchableSelect 
                label="Campagna" 
                placeholder="Seleziona campagna (opzionale)..." 
                options={campagne} 
                selectedId={form.campagnaId} 
                onSelect={(id: string) => setForm({ ...form, campagnaId: id })} 
                displayFn={(c: Campagna) => `${c.code} - ${c.name}`} 
                searchFn={(c: Campagna) => `${c.code} ${c.name} ${c.description}`} 
                subTextFn={(c: Campagna) => c.exclusiveBenefits || c.description} 
                icon={<Megaphone size={14} />} 
                allowClear={true}
              />
            </div>

            <div>
              <SearchableSelect 
                label="Assegna a Profilo Operatore" 
                placeholder="Seleziona agente responsabile..." 
                options={users} 
                selectedId={form.assignedTo || user.id} 
                onSelect={(id: string) => setForm({ ...form, assignedTo: id })} 
                displayFn={(u: any) => u.name} 
                searchFn={(u: any) => u.name} 
                subTextFn={(u: any) => `${u.userCode} - ${u.role}`} 
                icon={<UserIcon size={14} />} 
              />
            </div>

            {/* Note */}
            <div className="col-span-full space-y-1.5">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1">Note Operative</label>
              <textarea 
                className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-medium focus:border-blue-500" 
                rows={3}
                placeholder="Dettagli piano tariffario, numeri telefonici da portare, ecc..."
                value={form.notes || ''}
                onChange={e => setForm({ ...form, notes: e.target.value })}
              />
            </div>

            {/* Campo Allegati */}
            <div className="col-span-full">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1 mb-1 block">Documenti Allegati</label>
              <FileUploader 
                attachments={form.attachments || []} 
                canEdit={true} 
                onUpload={files => setForm(prev => ({ ...prev, attachments: [...(prev.attachments || []), ...files] }))} 
                onRemove={id => setForm(prev => ({ ...prev, attachments: prev.attachments?.filter(a => a.id !== id) }))} 
              />
            </div>

            {editingId && <div className="col-span-full"><TimestampDetail data={form} /></div>}
          </div>

          {isAdminOrBO ? (
            <div className="flex justify-end gap-4 border-t pt-6">
              <button 
                type="button" 
                onClick={() => { setIsFormOpen(false); setEditingId(null); }} 
                className="font-bold text-slate-400 px-6 py-3 hover:text-slate-600 transition-colors"
              >
                Annulla
              </button>
              <button 
                type="submit" 
                className="bg-blue-600 hover:bg-blue-700 text-white px-10 py-3.5 rounded-2xl font-black shadow-lg shadow-blue-500/20 hover:scale-105 transition-all text-sm cursor-pointer"
              >
                {editingId ? 'Aggiorna Contratto Telefonico' : 'Salva Contratto Telefonico'}
              </button>
            </div>
          ) : (
            editingId && form.status === 'non conforme' ? (
              <div className="flex justify-end items-center gap-4 border-t pt-6">
                <button 
                  type="button" 
                  onClick={() => { setIsFormOpen(false); setEditingId(null); }} 
                  className="font-bold text-slate-400 px-6 py-3 hover:text-slate-600 transition-colors"
                >
                  Annulla
                </button>
                <button 
                  type="button" 
                  onClick={(e) => handleSave(e, 'corretto')} 
                  className="bg-teal-600 hover:bg-teal-700 text-white px-10 py-3.5 rounded-2xl font-black shadow-lg shadow-teal-500/20 hover:scale-105 transition-all text-sm cursor-pointer flex items-center gap-2"
                >
                  <CheckCircle2 size={18} /> Salva Correzione (Corretto)
                </button>
              </div>
            ) : (
              <div className="flex justify-end items-center gap-3 border-t pt-6">
                <button 
                  type="button" 
                  onClick={() => { setIsFormOpen(false); setEditingId(null); }} 
                  className="font-bold text-slate-400 px-6 py-3 hover:text-slate-600 transition-colors mr-2"
                >
                  Annulla
                </button>
                <button 
                  type="button" 
                  onClick={(e) => handleSave(e, 'bozza')} 
                  className="bg-slate-200 hover:bg-slate-300 text-slate-800 px-7 py-3.5 rounded-2xl font-bold transition-all text-sm cursor-pointer"
                >
                  Salva in Bozza
                </button>
                <button 
                  type="button" 
                  onClick={(e) => handleSave(e, 'trasmesso')} 
                  className="bg-blue-600 hover:bg-blue-700 text-white px-9 py-3.5 rounded-2xl font-black shadow-lg shadow-blue-500/20 hover:scale-105 transition-all text-sm cursor-pointer flex items-center gap-2"
                >
                  <Send size={18} /> Trasmetti
                </button>
              </div>
            )
          )}
        </form>
      )}

      {/* Barra di ricerca e filtri */}
      <div className="flex flex-col md:flex-row gap-4 justify-between items-center bg-white p-4 rounded-2xl border border-slate-100 shadow-sm">
        <div className="relative flex-1 w-full">
          <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400" size={18} />
          <input 
            className="w-full pl-11 pr-4 py-2.5 rounded-xl border border-slate-200 outline-none focus:border-blue-500 font-bold text-sm" 
            placeholder="Cerca per cliente, gestore, sede o note..." 
            value={searchTerm}
            onChange={e => setSearchTerm(e.target.value)}
          />
        </div>
        <div className="flex items-center gap-3 w-full md:w-auto">
          <select 
            className="p-2.5 rounded-xl border border-slate-200 text-xs font-bold outline-none bg-white cursor-pointer"
            value={gestoreFilter}
            onChange={e => setGestoreFilter(e.target.value)}
          >
            <option value="all">Tutti i Gestori</option>
            {GESTORI.map(g => (
              <option key={g} value={g}>{g}</option>
            ))}
          </select>
          <select 
            className="p-2.5 rounded-xl border border-slate-200 text-xs font-bold outline-none bg-white cursor-pointer"
            value={statusFilter}
            onChange={e => setStatusFilter(e.target.value)}
          >
            <option value="all">Tutti gli Stati</option>
            {STATI.map(s => (
              <option key={s} value={s}>{s.toUpperCase()}</option>
            ))}
          </select>
        </div>
      </div>

      {/* Tabella Contratti Telefonici */}
      <div className="bg-white rounded-[2.5rem] border border-slate-100 shadow-xl overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse min-w-[900px]">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-100">
                <th className="p-4 text-[10px] font-black uppercase text-slate-400 tracking-widest">Cliente</th>
                <th className="p-4 text-[10px] font-black uppercase text-slate-400 tracking-widest">Gestore</th>
                <th className="p-4 text-[10px] font-black uppercase text-slate-400 tracking-widest">Operazione</th>
                <th className="p-4 text-[10px] font-black uppercase text-slate-400 tracking-widest">Campagna</th>
                <th className="p-4 text-[10px] font-black uppercase text-slate-400 tracking-widest">Sede / Indirizzo</th>
                <th className="p-4 text-[10px] font-black uppercase text-slate-400 tracking-widest text-center">Fisse</th>
                <th className="p-4 text-[10px] font-black uppercase text-slate-400 tracking-widest text-center">Mobili</th>
                <th className="p-4 text-[10px] font-black uppercase text-slate-400 tracking-widest text-center">Stato</th>
                {isAdminOrBO && (
                  <th className="p-4 text-[10px] font-black uppercase text-slate-400 tracking-widest">Operatore</th>
                )}
                <th className="p-4 text-[10px] font-black uppercase text-slate-400 tracking-widest text-right">Azioni</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-50">
              {filteredData.map(c => {
                const port = portfolios.find(p => p.id === c.portfolioId);
                const assigned = users.find(u => u.id === c.assignedTo);
                const camp = campagne.find(cp => cp.id === c.campagnaId);
                const canEdit = isAdminOrBO || c.status === 'bozza' || c.status === 'non conforme';
                const canDelete = isAdminOrBO || c.status === 'bozza';

                return (
                  <tr key={c.id} className="hover:bg-blue-50/20 transition-colors group">
                    <td className="p-4">
                      <p className="font-bold text-slate-800 text-sm">{port?.businessName || 'Cliente non specificato'}</p>
                      <p className="text-[10px] text-slate-400 font-mono">{port?.taxId || '---'}</p>
                    </td>
                    <td className="p-4">
                      <span className={`inline-block px-3 py-1 rounded-xl text-xs font-black uppercase tracking-wider border shadow-sm ${getGestoreBadge(c.gestore)}`}>
                        {c.gestore}
                      </span>
                    </td>
                    <td className="p-4">
                      <span className="text-xs font-bold text-slate-700 bg-slate-100 px-2.5 py-1 rounded-lg">
                        {c.tipoOperazione}
                      </span>
                    </td>
                    <td className="p-4">
                      {camp ? (
                        <div className="inline-flex items-center gap-1.5 bg-purple-50 text-purple-800 border border-purple-200 px-2.5 py-1 rounded-lg">
                          <span className="text-[10px] font-mono font-black">{camp.code}</span>
                          <span className="text-[10px] font-bold truncate max-w-[110px]" title={camp.name}>{camp.name}</span>
                        </div>
                      ) : (
                        <span className="text-xs text-slate-300 font-bold">---</span>
                      )}
                    </td>
                    <td className="p-4 max-w-xs">
                      {c.address ? (
                        <div className="flex items-center gap-1.5 text-xs text-slate-700">
                          <AddressLink address={c.address} showIconOnly className="text-blue-500 shrink-0" />
                          <span className="truncate" title={c.address}>{c.address}</span>
                        </div>
                      ) : (
                        <span className="text-xs text-slate-400 italic">Nessun indirizzo</span>
                      )}
                    </td>
                    <td className="p-4 text-center">
                      <span className="inline-flex items-center gap-1 font-mono font-black text-indigo-700 bg-indigo-50 border border-indigo-200 px-2.5 py-0.5 rounded-lg text-xs">
                        ☎️ {c.lineeFisse ?? 0}
                      </span>
                    </td>
                    <td className="p-4 text-center">
                      <span className="inline-flex items-center gap-1 font-mono font-black text-purple-700 bg-purple-50 border border-purple-200 px-2.5 py-0.5 rounded-lg text-xs">
                        📱 {c.lineeMobili ?? 0}
                      </span>
                    </td>
                    <td className="p-4 text-center">
                      <span className={`inline-block px-2.5 py-1 rounded-full text-[10px] font-black uppercase tracking-wider border ${getStatusBadge(c.status)}`}>
                        {c.status}
                      </span>
                    </td>
                    {isAdminOrBO && (
                      <td className="p-4">
                        <p className="text-xs font-bold text-slate-700 leading-tight">{assigned?.name || '---'}</p>
                        <p className="text-[10px] font-black text-emerald-600">{assigned?.userCode || '---'}</p>
                      </td>
                    )}
                    <td className="p-4 text-right">
                      <div className="flex items-center justify-end gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                        {canEdit && (
                          <button 
                            onClick={() => startEdit(c)} 
                            className="p-2 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded-xl transition-all"
                            title={c.status === 'non conforme' ? "Correggi contratto non conforme" : "Modifica bozza"}
                          >
                            <Edit2 size={16} />
                          </button>
                        )}
                        {canDelete && (
                          <button 
                            onClick={() => setDeleteConfirmId(c.id)} 
                            className="p-2 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded-xl transition-all"
                            title="Elimina"
                          >
                            <Trash2 size={16} />
                          </button>
                        )}
                      </div>
                    </td>
                  </tr>
                );
              })}
              {filteredData.length === 0 && (
                <tr>
                  <td colSpan={9} className="p-16 text-center">
                    <div className="flex flex-col items-center gap-3 opacity-30">
                      <Phone size={56} />
                      <p className="font-black uppercase tracking-widest text-sm">Nessun contratto telefonico trovato</p>
                    </div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      <ConfirmModal
        isOpen={!!deleteConfirmId}
        title="Conferma Eliminazione"
        message="Sei sicuro di voler eliminare questo contratto telefonico? Questa operazione non può essere annullata."
        onConfirm={() => {
          setData(prev => prev.filter(c => c.id !== deleteConfirmId));
          setDeleteConfirmId(null);
        }}
        onCancel={() => setDeleteConfirmId(null)}
      />
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

const CasiView: React.FC<any> = ({ data, setData, portfolios, pdps, contracts = [], users = [], user, exportFn, onOpenChat, getUnreadCount, onViewPdp }) => {
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [form, setForm] = useState<Partial<Caso>>({ status: 'nuovo', category: 'Anagrafica', sector: undefined, attachments: [] });
  const [editingId, setEditingId] = useState<string | null>(null);
  const [deleteConfirmId, setDeleteConfirmId] = useState<string | null>(null);
  const [statusFilter, setStatusFilter] = useState<string>('');
  const [agentCodeFilter, setAgentCodeFilter] = useState<string>('');

  const isAdminOrBO = user?.role === 'admin' || user?.role === 'backoffice';
  const SETTORI: CaseSector[] = ['Energetico', 'Telefonico', 'Efficienza', 'CER', 'TARI', 'Acqua'];

  const availablePdps = useMemo(() => {
    if (!form.portfolioId) return pdps;
    const clientPdpIds = new Set(contracts.filter((c: any) => c.portfolioId === form.portfolioId).map((c: any) => c.pdpId));
    if (clientPdpIds.size === 0) return pdps;
    const clientPdps = pdps.filter((p: any) => clientPdpIds.has(p.id));
    const otherPdps = pdps.filter((p: any) => !clientPdpIds.has(p.id));
    return [...clientPdps, ...otherPdps];
  }, [form.portfolioId, pdps, contracts]);

  const isClientRequired = ['Contenzioso', 'Commerciale', 'Crediti', 'Fatture', 'Fiscale', 'Distribuzione'].includes(form.category || '');
  const isPdpRequired = form.category === 'Distribuzione';
  const isSectorModifiable = form.category === 'Commerciale' || form.category === 'Contenzioso' || form.category === 'Mandati';
  const isSectorRequired = form.category !== 'Anagrafica' && form.category !== 'Interni';
  const effectiveSector: CaseSector | undefined = isSectorRequired && !isSectorModifiable ? 'Energetico' : form.sector;
  const showPdpField = effectiveSector === 'Energetico';

  const filteredCasi = useMemo(() => {
    return data.filter((c: Caso) => {
      if (statusFilter && c.status !== statusFilter) return false;
      if (agentCodeFilter.trim()) {
        const port = portfolios.find((p: Portfolio) => p.id === c.portfolioId);
        const assignedUser = users.find((u: User) => u.id === (c.assignedTo || port?.assignedTo));
        const codeStr = `${assignedUser?.userCode || ''} ${assignedUser?.name || ''}`.toLowerCase();
        if (!codeStr.includes(agentCodeFilter.trim().toLowerCase())) return false;
      }
      return true;
    });
  }, [data, statusFilter, agentCodeFilter, portfolios, users]);

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    if (isClientRequired && !form.portfolioId) {
      alert(`Per la categoria "${form.category}" il campo Cliente è obbligatorio.`);
      return;
    }

    let finalSector: CaseSector | undefined = form.sector;
    if (!isSectorModifiable) {
      if (isSectorRequired) {
        finalSector = 'Energetico';
      } else {
        finalSector = form.sector === 'Energetico' ? 'Energetico' : undefined;
      }
    } else {
      finalSector = form.sector || 'Energetico';
    }

    if (isSectorRequired && !finalSector) {
      alert("Il campo Settore è obbligatorio per questa categoria.");
      return;
    }

    if (finalSector === 'Energetico' && isPdpRequired && !form.pdpId) {
      alert('Per la categoria "Distribuzione" il campo PDP è obbligatorio.');
      return;
    }

    let finalStatus: CaseStatus = 'nuovo';
    if (isAdminOrBO) {
      finalStatus = (form.status as CaseStatus) || 'nuovo';
    } else {
      if (editingId) {
        const existing = data.find((c: any) => c.id === editingId);
        finalStatus = existing?.status || 'nuovo';
      } else {
        finalStatus = 'nuovo';
      }
    }

    const ts = getTimestamp();
    const finalCategory = form.category || 'Anagrafica';
    const finalPdpId = finalSector === 'Energetico' ? form.pdpId : undefined;
    if (editingId) {
      setData((prev: Caso[]) => prev.map(c => c.id === editingId ? { ...c, ...form, pdpId: finalPdpId, category: finalCategory, sector: finalSector, status: finalStatus, attachments: form.attachments || [], updatedAt: ts } : c));
    } else {
      setData((prev: Caso[]) => [...prev, { ...form, pdpId: finalPdpId, category: finalCategory, sector: finalSector, status: finalStatus, attachments: form.attachments || [], id: Math.random().toString(36).substr(2, 9), assignedTo: user.id, createdAt: ts, updatedAt: ts } as Caso]);
    }
    setIsFormOpen(false); 
    setEditingId(null); 
    setForm({ status: 'nuovo', category: 'Anagrafica', sector: undefined, attachments: [] });
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-300">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h2 className="text-3xl font-black text-emerald-900">Gestione Casi</h2>
          <p className="text-slate-500 font-medium">Ticketing operativo</p>
        </div>
        <div className="flex flex-wrap items-center gap-3 w-full md:w-auto">
          <select
            className="border-2 border-slate-100 rounded-2xl px-4 py-2.5 bg-white font-bold text-xs outline-none focus:border-red-500 cursor-pointer"
            value={statusFilter}
            onChange={e => setStatusFilter(e.target.value)}
          >
            <option value="">Tutti gli stati</option>
            <option value="nuovo">Nuovo</option>
            <option value="in lavorazione">In Lavorazione</option>
            <option value="risolto">Risolto</option>
            <option value="KO">KO</option>
            <option value="partner">Partner</option>
          </select>

          <div className="relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={14} />
            <input
              className="pl-9 pr-4 py-2.5 rounded-2xl border-2 border-slate-100 bg-white outline-none focus:border-red-500 font-bold text-xs w-48"
              placeholder="Filtra per codice agente..."
              value={agentCodeFilter}
              onChange={e => setAgentCodeFilter(e.target.value)}
            />
          </div>

          <button 
            onClick={() => exportFn(filteredCasi.map((c: Caso) => {
              const port = portfolios.find(p => p.id === c.portfolioId);
              const assignedUser = users.find((u: User) => u.id === (c.assignedTo || port?.assignedTo));
              return {
                id: c.id,
                titolo: c.title,
                categoria: c.category,
                settore: c.sector || 'N/D',
                codice_agente: assignedUser?.userCode || '',
                cliente: port?.businessName || '',
                cf_iva: port?.taxId || '',
                pdp: pdps.find(p => p.id === c.pdpId)?.pdpCode || '',
                stato: c.status,
                note: c.notes || '',
                creato_il: c.createdAt
              };
            }), 'Casi')} 
            className="p-3 text-slate-400 hover:text-slate-800 transition-colors"
          >
            <FileSpreadsheet />
          </button>
          <button 
            onClick={() => { 
              setForm({ status: 'nuovo', category: 'Anagrafica', sector: undefined, attachments: [] }); 
              setEditingId(null); 
              setIsFormOpen(true); 
            }} 
            className="bg-red-500 text-white px-6 py-3 rounded-2xl font-black flex items-center gap-2 shadow-lg hover:scale-105 transition-all"
          >
            <Plus /> Apri Caso
          </button>
        </div>
      </div>
      {isFormOpen && (
        <form onSubmit={handleSave} className="bg-white p-10 rounded-[3rem] border-2 border-red-50 shadow-2xl space-y-8 animate-in slide-in-from-top-4">
           <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
             <div className="col-span-full grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-6">
               <div className="space-y-1.5">
                 <label className="text-xs font-bold text-slate-500 uppercase ml-1">Categoria *</label>
                 <select 
                   className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-red-500 cursor-pointer" 
                   required 
                   value={form.category || 'Anagrafica'} 
                   onChange={e => {
                     const newCat = e.target.value as CaseCategory;
                     const modifiable = newCat === 'Commerciale' || newCat === 'Contenzioso' || newCat === 'Mandati';
                     const required = newCat !== 'Anagrafica' && newCat !== 'Interni';
                     setForm(prev => {
                       const nextSector = !required ? (prev.sector === 'Energetico' ? 'Energetico' : undefined) : (!modifiable ? 'Energetico' : (prev.sector || 'Energetico'));
                       return {
                         ...prev,
                         category: newCat,
                         sector: nextSector,
                         pdpId: nextSector === 'Energetico' ? prev.pdpId : undefined
                       };
                     });
                   }}
                 >
                   <option value="Anagrafica">Anagrafica</option>
                   <option value="Interni">Interni</option>
                   <option value="Mandati">Mandati</option>
                   <option value="Contenzioso">Contenzioso</option>
                   <option value="Commerciale">Commerciale</option>
                   <option value="Crediti">Crediti</option>
                   <option value="Fatture">Fatture</option>
                   <option value="Fiscale">Fiscale</option>
                   <option value="Distribuzione">Distribuzione</option>
                 </select>
               </div>

               <div className="space-y-1.5">
                 <label className="text-xs font-bold text-slate-500 uppercase ml-1">
                   Settore {isSectorRequired ? '*' : '(Opzionale)'}
                 </label>
                 <select 
                   className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-red-500 disabled:opacity-60 disabled:cursor-not-allowed cursor-pointer" 
                   disabled={!isSectorModifiable && isSectorRequired}
                   value={isSectorRequired && !isSectorModifiable ? 'Energetico' : (form.sector || '')} 
                   onChange={e => {
                     const nextSec = (e.target.value as CaseSector) || undefined;
                     setForm({...form, sector: nextSec, pdpId: nextSec === 'Energetico' ? form.pdpId : undefined});
                   }}
                 >
                   {!isSectorRequired && (
                     <option value="">Nessuno</option>
                   )}
                   {isSectorModifiable ? (
                     SETTORI.map(s => (
                       <option key={s} value={s}>{s}</option>
                     ))
                   ) : (
                     <option value="Energetico">Energetico</option>
                   )}
                 </select>
                 {!isSectorModifiable && isSectorRequired && (
                   <p className="text-[10px] text-slate-400 font-medium ml-1">Fisso su Energetico</p>
                 )}
                 {!isSectorRequired && (
                   <p className="text-[10px] text-slate-400 font-medium ml-1">Opzionale per {form.category}</p>
                 )}
               </div>

               {/* Invertiti Stato Caso e Titolo */}
               <div className="space-y-1.5 sm:col-span-2">
                 <label className="text-xs font-bold text-slate-500 uppercase ml-1">Stato Caso</label>
                 {isAdminOrBO ? (
                   <select 
                     className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 font-bold outline-none cursor-pointer focus:border-red-500" 
                     value={form.status || 'nuovo'} 
                     onChange={e => setForm({...form, status: e.target.value as CaseStatus})}
                   >
                      <option value="nuovo">Nuovo</option>
                      <option value="in lavorazione">In Lavorazione</option>
                      <option value="risolto">Risolto</option>
                      <option value="KO">KO</option>
                      <option value="partner">Partner</option>
                   </select>
                 ) : (
                   <div className="p-3.5 rounded-2xl bg-slate-100 border border-slate-200 flex items-center text-xs h-[52px]">
                     <span className="font-black uppercase text-slate-700">
                       {(form.status || 'nuovo').toUpperCase()}
                     </span>
                   </div>
                 )}
               </div>
             </div>

             <SearchableSelect 
               label={isClientRequired ? "Cliente *" : "Cliente (Opzionale)"} 
               placeholder="Associa cliente..." 
               options={portfolios} 
               selectedId={form.portfolioId} 
               onSelect={(id:string) => setForm({...form, portfolioId: id})} 
               displayFn={(p:any)=>p.businessName} 
               searchFn={(p:any)=>p.businessName} 
               subTextFn={(p:any)=>p.taxId} 
               icon={<Users size={14}/>} 
             />

             {showPdpField && (
               <SearchableSelect 
                 label={isPdpRequired ? "PDP *" : "PDP (Opzionale)"} 
                 placeholder="Cerca PDP..." 
                 options={availablePdps} 
                 selectedId={form.pdpId} 
                 onSelect={(id:string) => setForm({...form, pdpId: id})} 
                 displayFn={(p:any)=>p.pdpCode} 
                 searchFn={(p:any)=>p.pdpCode} 
                 subTextFn={(p:any)=>p.address} 
                 icon={<Zap size={14}/>} 
               />
             )}

             <div className={`space-y-1.5 ${!showPdpField ? '' : 'col-span-full'}`}>
               <label className="text-xs font-bold text-slate-500 uppercase ml-1">Titolo *</label>
               <input 
                 className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-red-500" 
                 required 
                 value={form.title || ''} 
                 onChange={e => setForm({...form, title: e.target.value})} 
                 placeholder="Titolo Caso"
               />
             </div>

             <div className="col-span-full"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Note / Dettagli</label><textarea className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-medium" rows={4} value={form.notes || ''} onChange={e => setForm({...form, notes: e.target.value})} /></div>

             <div className="col-span-full">
               <label className="text-xs font-bold text-slate-500 uppercase ml-1">Allegati (le foto caricate generano automaticamente un PDF scaricabile)</label>
               <FileUploader
                 attachments={form.attachments || []}
                 canEdit={true}
                 onUpload={files => setForm({ ...form, attachments: [...(form.attachments || []), ...files] })}
                 onRemove={id => setForm({ ...form, attachments: form.attachments?.filter(a => a.id !== id) })}
                 photoToPdfConfig={{
                   enabled: true,
                   mode: 'case',
                   caseTitle: form.title || 'CASO'
                 }}
               />
             </div>

             {editingId && <TimestampDetail data={form} />}
           </div>
           <div className="flex justify-end gap-4 border-t pt-8"><button type="button" onClick={() => setIsFormOpen(false)} className="font-bold text-slate-400">Annulla</button><button type="submit" className="bg-emerald-900 text-white px-10 py-4 rounded-2xl font-black">{editingId ? 'Aggiorna' : 'Apri'} Pratica</button></div>
        </form>
      )}
      <div className="space-y-4">{filteredCasi.map((c: Caso) => {
        const port = portfolios.find(p => p.id === c.portfolioId);
        const pdp = pdps.find((p: any) => p.id === c.pdpId);
        const assignedUser = users.find((u: User) => u.id === (c.assignedTo || port?.assignedTo));
        return (
        <div key={c.id} className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm hover:shadow-xl transition-all flex items-center justify-between group">
           <div className="flex items-center gap-6 flex-1">
             <div className={`p-4 rounded-2xl ${c.status === 'nuovo' ? 'bg-blue-50 text-blue-500' : c.status === 'KO' ? 'bg-red-50 text-red-500' : 'bg-emerald-50 text-emerald-500'}`}><AlertCircle/></div>
             <div className="overflow-hidden">
               <div className="flex flex-wrap items-center gap-2 mb-0.5">
                 <span className="text-[9px] font-black uppercase tracking-wider px-2 py-0.5 rounded-md bg-red-100 text-red-700 border border-red-200">
                   {c.category || 'Anagrafica'}
                 </span>
                 <span className={`text-[9px] font-black uppercase tracking-wider px-2 py-0.5 rounded-md border ${
                   c.sector === 'Telefonico' ? 'bg-blue-100 text-blue-700 border-blue-200' :
                   c.sector === 'Efficienza' ? 'bg-amber-100 text-amber-700 border-amber-200' :
                   c.sector === 'CER' ? 'bg-purple-100 text-purple-700 border-purple-200' :
                   c.sector === 'TARI' ? 'bg-orange-100 text-orange-700 border-orange-200' :
                   c.sector === 'Acqua' ? 'bg-cyan-100 text-cyan-700 border-cyan-200' :
                   c.sector === 'Energetico' ? 'bg-emerald-100 text-emerald-700 border-emerald-200' :
                   'bg-slate-100 text-slate-600 border-slate-200'
                 }`}>
                   Settore: {c.sector || 'Generale'}
                 </span>
                 {assignedUser?.userCode && (
                   <span className="text-[9px] font-mono font-black uppercase tracking-wider px-2 py-0.5 rounded-md bg-slate-100 text-slate-700 border border-slate-200">
                     Agente: {assignedUser.userCode}
                   </span>
                 )}
                 <p className="text-[10px] font-black uppercase text-slate-400 truncate">{port?.businessName || 'Senza Cliente'}</p>
               </div>
               <h4 className="text-lg font-black tracking-tight text-slate-800 truncate">{c.title}</h4>
               <div className="flex flex-wrap items-center gap-2 mt-1">
                 {pdp && c.sector === 'Energetico' && (
                   <span onClick={(e) => { e.stopPropagation(); onViewPdp && onViewPdp(pdp); }} className="text-[10px] font-black text-emerald-600 font-mono bg-emerald-50 px-2 py-0.5 rounded cursor-pointer hover:bg-emerald-100 transition-colors">
                     {pdp.pdpCode}
                   </span>
                 )}
                 {c.attachments && c.attachments.length > 0 && (
                   <div className="flex flex-wrap items-center gap-1.5">
                     {c.attachments.map(att => (
                       <button
                         key={att.id}
                         type="button"
                         onClick={(e) => {
                           e.stopPropagation();
                           triggerDownloadAttachment(att);
                         }}
                         className="inline-flex items-center gap-1 text-[10px] font-black text-emerald-700 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 px-2 py-0.5 rounded transition-colors cursor-pointer"
                         title={`Scarica ${att.fileName}`}
                       >
                         <Download size={10} /> {att.fileName}
                       </button>
                     ))}
                   </div>
                 )}
                 {c.notes && <span className="text-[10px] text-slate-400 truncate italic">"{c.notes}"</span>}
               </div>
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
                <button onClick={() => setDeleteConfirmId(c.id)} className="p-3 text-slate-300 hover:text-red-400"><X size={18}/></button>
             </div>
           </div>
        </div>)})}
      </div>
      <ConfirmModal
        isOpen={!!deleteConfirmId}
        title="Conferma Eliminazione"
        message="Sei sicuro di voler eliminare questo caso di supporto? Questa operazione non può essere annullata."
        onConfirm={() => {
          setData((prev: Caso[]) => prev.filter(i => i.id !== deleteConfirmId));
          setDeleteConfirmId(null);
        }}
        onCancel={() => setDeleteConfirmId(null)}
      />
    </div>
  );
};

const UsersView: React.FC<any> = ({ users, setUsers }) => {
  const [isAdding, setIsAdding] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [deleteConfirmId, setDeleteConfirmId] = useState<string | null>(null);
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
                <button onClick={() => setDeleteConfirmId(u.id)} className="text-slate-400 hover:text-red-500 p-1"><Trash2 size={16}/></button>
            </div>
         </div>))}
       </div>
      <ConfirmModal
        isOpen={!!deleteConfirmId}
        title="Conferma Eliminazione"
        message="Sei sicuro di voler eliminare questo profilo utente? Questa operazione non può essere annullata."
        onConfirm={() => {
          setUsers((prev: User[]) => prev.filter(x => x.id !== deleteConfirmId));
          setDeleteConfirmId(null);
        }}
        onCancel={() => setDeleteConfirmId(null)}
      />
    </div>
  );
};

const PartnersView: React.FC<{
  partners: Partner[];
  setPartners: React.Dispatch<React.SetStateAction<Partner[]>>;
}> = ({ partners, setPartners }) => {
  const [isAdding, setIsAdding] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [deleteConfirmId, setDeleteConfirmId] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [form, setForm] = useState<Partial<Partner>>({ role: 'utente' });

  const filteredPartners = useMemo(() => {
    if (!searchTerm) return partners;
    const s = searchTerm.toLowerCase();
    return partners.filter(p =>
      p.name?.toLowerCase().includes(s) ||
      p.userCode?.toLowerCase().includes(s) ||
      p.username?.toLowerCase().includes(s) ||
      p.email?.toLowerCase().includes(s)
    );
  }, [partners, searchTerm]);

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    const payload: Partner = {
      id: editingId || Math.random().toString(36).substr(2, 9),
      name: (form.name || '').trim(),
      userCode: (form.userCode || '').toUpperCase().trim(),
      username: (form.username || '').trim(),
      password: form.password || '',
      role: 'utente',
      email: (form.email || '').trim(),
      phone: (form.phone || '').trim(),
      avatar: form.avatar,
      notificationSettings: form.notificationSettings
    };
    if (editingId) {
      setPartners(prev => prev.map(p => p.id === editingId ? payload : p));
    } else {
      setPartners(prev => [...prev, payload]);
    }
    setIsAdding(false);
    setEditingId(null);
    setForm({ role: 'utente' });
  };

  const startEdit = (p: Partner) => {
    setForm(p);
    setEditingId(p.id);
    setIsAdding(true);
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-300">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h2 className="text-3xl font-black text-teal-900 tracking-tighter flex items-center gap-2">
            <Handshake className="text-teal-600" size={28} /> Gestione Partner
          </h2>
          <p className="text-slate-500 font-medium text-xs">Profili Partner singoli con accesso dedicato (senza distinzione di grado)</p>
        </div>
        <div className="flex items-center gap-3 w-full md:w-auto">
          <div className="relative flex-1 md:w-64">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
            <input
              className="w-full pl-10 pr-4 py-3 rounded-2xl border-2 border-slate-100 outline-none focus:border-teal-500 font-bold text-sm bg-white"
              placeholder="Cerca partner o codice..."
              value={searchTerm}
              onChange={e => setSearchTerm(e.target.value)}
            />
          </div>
          <button
            onClick={() => { setIsAdding(true); setEditingId(null); setForm({ role: 'utente' }); }}
            className="bg-teal-700 hover:bg-teal-800 text-white px-6 py-3 rounded-2xl font-black flex items-center gap-2 shadow-lg transition-all"
          >
            <Plus size={18} /> Nuovo Partner
          </button>
        </div>
      </div>

      {isAdding && (
        <form onSubmit={handleSave} className="bg-white p-8 rounded-[2rem] border-2 border-teal-50 shadow-2xl space-y-6 max-w-2xl animate-in slide-in-from-top-4">
          <div className="flex justify-between items-center border-b border-slate-100 pb-3">
            <h3 className="text-xl font-black text-teal-950">{editingId ? 'Modifica Partner' : 'Crea Nuovo Partner'}</h3>
            <span className="bg-teal-50 text-teal-700 border border-teal-200 px-3 py-1 rounded-full text-[10px] font-black uppercase">
              Utente Singolo (Senza Grado)
            </span>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="space-y-1">
              <label className="text-xs font-bold text-slate-500 uppercase">Nome / Ragione Sociale *</label>
              <input required className="w-full border p-3 rounded-xl bg-slate-50 font-bold" value={form.name || ''} onChange={e => setForm({ ...form, name: e.target.value })} />
            </div>
            <div className="space-y-1">
              <label className="text-xs font-bold text-slate-500 uppercase">Codice Partner *</label>
              <input required className="w-full border p-3 rounded-xl bg-slate-50 font-black text-teal-600" value={form.userCode || ''} onChange={e => setForm({ ...form, userCode: e.target.value.toUpperCase() })} />
            </div>
            <div className="space-y-1">
              <label className="text-xs font-bold text-slate-500 uppercase">Username *</label>
              <input required className="w-full border p-3 rounded-xl bg-slate-50 font-bold" value={form.username || ''} onChange={e => setForm({ ...form, username: e.target.value })} />
            </div>
            <div className="space-y-1">
              <label className="text-xs font-bold text-slate-500 uppercase">Password *</label>
              <input required type="password" className="w-full border p-3 rounded-xl bg-slate-50" value={form.password || ''} onChange={e => setForm({ ...form, password: e.target.value })} />
            </div>
            <div className="space-y-1">
              <label className="text-xs font-bold text-slate-500 uppercase">Email</label>
              <input type="email" className="w-full border p-3 rounded-xl bg-slate-50" value={form.email || ''} onChange={e => setForm({ ...form, email: e.target.value })} />
            </div>
            <div className="space-y-1">
              <label className="text-xs font-bold text-slate-500 uppercase">Cellulare</label>
              <input type="tel" className="w-full border p-3 rounded-xl bg-slate-50" value={form.phone || ''} onChange={e => setForm({ ...form, phone: e.target.value })} />
            </div>
          </div>
          <div className="flex justify-end gap-3 pt-4 border-t">
            <button type="button" onClick={() => { setIsAdding(false); setEditingId(null); }} className="font-bold text-slate-400 px-4">Annulla</button>
            <button type="submit" className="bg-teal-600 hover:bg-teal-700 text-white px-8 py-3 rounded-xl font-black">
              {editingId ? 'Aggiorna Partner' : 'Crea Partner'}
            </button>
          </div>
        </form>
      )}

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        {filteredPartners.map(p => (
          <div key={p.id} className="p-6 rounded-3xl border border-teal-100 bg-white shadow-sm flex items-center gap-4 group transition-all hover:shadow-md">
            <div className="w-12 h-12 rounded-2xl bg-teal-50 text-teal-600 flex items-center justify-center font-black text-lg">
              {p.name?.charAt(0) || 'P'}
            </div>
            <div className="flex-1 overflow-hidden">
              <p className="text-sm text-slate-800 font-bold truncate">{p.name}</p>
              <p className="text-[10px] font-black uppercase text-teal-600 tracking-wider">{p.userCode}</p>
              <p className="text-[10px] text-slate-400 truncate">@{p.username} {p.email ? `• ${p.email}` : ''}</p>
            </div>
            <div className="flex gap-1 opacity-0 group-hover:opacity-100 transition-all">
              <button onClick={() => startEdit(p)} className="text-slate-400 hover:text-teal-600 p-1"><Edit2 size={16} /></button>
              <button onClick={() => setDeleteConfirmId(p.id)} className="text-slate-400 hover:text-red-500 p-1"><Trash2 size={16} /></button>
            </div>
          </div>
        ))}
        {filteredPartners.length === 0 && (
          <div className="col-span-full bg-white p-14 rounded-3xl border border-slate-100 text-center text-slate-400">
            <Handshake size={40} className="mx-auto mb-2 opacity-30" />
            <p className="font-black uppercase text-xs tracking-widest">Nessun Partner registrato</p>
          </div>
        )}
      </div>

      <ConfirmModal
        isOpen={!!deleteConfirmId}
        title="Conferma Eliminazione Partner"
        message="Sei sicuro di voler eliminare questo profilo Partner? Questa operazione non può essere annullata."
        onConfirm={() => {
          setPartners(prev => prev.filter(x => x.id !== deleteConfirmId));
          setDeleteConfirmId(null);
        }}
        onCancel={() => setDeleteConfirmId(null)}
      />
    </div>
  );
};

const AppartenenzeView: React.FC<{
  data: Appartenenza[];
  setData: React.Dispatch<React.SetStateAction<Appartenenza[]>>;
  portfolios: Portfolio[];
  setPortfolios: React.Dispatch<React.SetStateAction<Portfolio[]>>;
  users: User[];
  partners?: Partner[];
  contracts?: Contract[];
  contrattiTelefonici?: ContrattoTelefonico[];
  pdps?: PDP[];
  user: User;
  exportFn: (data: any[], fileName: string) => void;
}> = ({ data, setData, portfolios, setPortfolios, users, partners = [], contracts = [], contrattiTelefonici = [], pdps = [], user, exportFn }) => {
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [deleteConfirmId, setDeleteConfirmId] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedPortfolioToAdd, setSelectedPortfolioToAdd] = useState<{ [groupId: string]: string }>({});
  const [selectedPortfolio, setSelectedPortfolio] = useState<Portfolio | null>(null);
  const [form, setForm] = useState<Partial<Appartenenza>>({ name: '', description: '', assignedTo: user.id });

  const allUsersAndPartners = useMemo(() => [...users, ...partners], [users, partners]);

  const isAdminOrBO = user.role === 'admin' || user.role === 'backoffice';

  const filteredAppartenenze = useMemo(() => {
    if (!searchTerm) return data;
    const s = searchTerm.toLowerCase();
    return data.filter(a => {
      const memberPortfolios = portfolios.filter(p => p.appartenenzaId === a.id);
      const membersStr = memberPortfolios.map(p => `${p.businessName} ${p.taxId}`).join(' ').toLowerCase();
      return (
        a.name?.toLowerCase().includes(s) ||
        a.description?.toLowerCase().includes(s) ||
        membersStr.includes(s)
      );
    });
  }, [data, portfolios, searchTerm]);

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    if (!form.name?.trim()) return alert("Inserisci il nome del marchio o gruppo di appartenenza.");
    const ts = getTimestamp();
    const payload: Appartenenza = {
      id: editingId || Math.random().toString(36).substr(2, 9),
      name: form.name.trim(),
      description: (form.description || '').trim(),
      assignedTo: isAdminOrBO ? (form.assignedTo || user.id) : user.id,
      createdAt: form.createdAt || ts,
      updatedAt: ts
    };

    if (editingId) {
      setData(prev => prev.map(a => a.id === editingId ? payload : a));
    } else {
      setData(prev => [...prev, payload]);
    }
    setIsFormOpen(false);
    setEditingId(null);
    setForm({ name: '', description: '', assignedTo: user.id });
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-300">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h2 className="text-3xl font-black text-indigo-950 tracking-tighter flex items-center gap-2">
            <Layers className="text-indigo-600" size={28} /> Appartenenza (Marchi & Gruppi)
          </h2>
          <p className="text-slate-500 font-medium text-xs">
            Raggruppa più anagrafiche di Portafoglio sotto lo stesso marchio o gruppo societario
          </p>
        </div>
        <div className="flex flex-wrap items-center gap-3 w-full md:w-auto">
          <div className="relative flex-1 md:w-64">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
            <input
              className="w-full pl-10 pr-4 py-3 rounded-2xl border-2 border-slate-100 outline-none focus:border-indigo-500 font-bold text-sm bg-white"
              placeholder="Cerca marchio, gruppo, cliente..."
              value={searchTerm}
              onChange={e => setSearchTerm(e.target.value)}
            />
          </div>
          <button onClick={() => exportFn(data, 'Appartenenze')} className="p-3 text-slate-400 hover:text-slate-800 bg-white border border-slate-100 rounded-2xl shadow-sm transition-colors">
            <FileSpreadsheet size={20} />
          </button>
          <button
            onClick={() => {
              setForm({ name: '', description: '', assignedTo: user.id });
              setEditingId(null);
              setIsFormOpen(true);
            }}
            className="bg-indigo-600 hover:bg-indigo-700 text-white px-6 py-3 rounded-2xl font-black flex items-center gap-2 shadow-lg transition-all"
          >
            <Plus size={18} /> Nuova Appartenenza
          </button>
        </div>
      </div>

      {isFormOpen && (
        <form onSubmit={handleSave} className="bg-white p-8 rounded-[2.5rem] border-2 border-indigo-50 shadow-2xl space-y-6 animate-in slide-in-from-top-4">
          <div className="flex justify-between items-center border-b border-slate-100 pb-4">
            <h3 className="text-xl font-black text-indigo-950">{editingId ? 'Modifica Appartenenza' : 'Nuovo Marchio / Gruppo di Appartenenza'}</h3>
            <button type="button" onClick={() => { setIsFormOpen(false); setEditingId(null); }} className="p-2 text-slate-400 hover:text-slate-600 rounded-full hover:bg-slate-100">
              <X size={20} />
            </button>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1">Nome Marchio / Gruppo *</label>
              <input
                required
                placeholder="Es. Gruppo Rossi, Holding Energia, Marchio XYZ..."
                className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-indigo-500"
                value={form.name || ''}
                onChange={e => setForm({ ...form, name: e.target.value })}
              />
            </div>
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1">Descrizione / Riferimenti Gruppo</label>
              <input
                placeholder="Note sul gruppo o sulle società collegate..."
                className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-indigo-500"
                value={form.description || ''}
                onChange={e => setForm({ ...form, description: e.target.value })}
              />
            </div>
            {isAdminOrBO && (
              <div className="col-span-full">
                <SearchableSelect
                  label="Utente Titolare"
                  placeholder="Seleziona utente..."
                  options={allUsersAndPartners}
                  selectedId={form.assignedTo || user.id}
                  onSelect={(id: string) => setForm({ ...form, assignedTo: id })}
                  displayFn={(u: any) => u.name}
                  searchFn={(u: any) => `${u.name} ${u.userCode || ''}`}
                  subTextFn={(u: any) => `${u.userCode || ''} - ${u.role}`}
                  icon={<UserIcon size={14} />}
                />
              </div>
            )}
          </div>
          <div className="flex justify-end gap-3 pt-4 border-t">
            <button type="button" onClick={() => { setIsFormOpen(false); setEditingId(null); }} className="font-bold text-slate-400 px-4">Annulla</button>
            <button type="submit" className="bg-indigo-600 hover:bg-indigo-700 text-white px-8 py-3.5 rounded-2xl font-black shadow-lg">
              {editingId ? 'Aggiorna Appartenenza' : 'Salva Appartenenza'}
            </button>
          </div>
        </form>
      )}

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {filteredAppartenenze.map(a => {
          const owner = allUsersAndPartners.find(u => u.id === a.assignedTo);
          const groupPortfolios = portfolios.filter(p => p.appartenenzaId === a.id);
          const availablePortfolios = portfolios.filter(p => p.appartenenzaId !== a.id);

          return (
            <div key={a.id} className="bg-white p-6 rounded-[2rem] border border-slate-100 shadow-sm hover:shadow-md transition-all flex flex-col justify-between space-y-5">
              <div className="space-y-3">
                <div className="flex justify-between items-start gap-4">
                  <div className="flex items-center gap-3">
                    <div className="p-3.5 rounded-2xl bg-indigo-50 text-indigo-600">
                      <Layers size={22} />
                    </div>
                    <div>
                      <h3 className="text-xl font-black text-slate-900 tracking-tight">{a.name}</h3>
                      {a.description && <p className="text-xs text-slate-500 font-medium">{a.description}</p>}
                    </div>
                  </div>
                  <div className="flex items-center gap-1">
                    <button
                      onClick={() => {
                        setForm(a);
                        setEditingId(a.id);
                        setIsFormOpen(true);
                      }}
                      className="p-2 text-slate-400 hover:text-indigo-600 rounded-xl hover:bg-indigo-50 transition-colors"
                      title="Modifica Appartenenza"
                    >
                      <Edit2 size={16} />
                    </button>
                    <button
                      onClick={() => setDeleteConfirmId(a.id)}
                      className="p-2 text-slate-400 hover:text-red-600 rounded-xl hover:bg-red-50 transition-colors"
                      title="Elimina Appartenenza"
                    >
                      <Trash2 size={16} />
                    </button>
                  </div>
                </div>

                <div className="flex items-center justify-between text-[10px] font-black uppercase tracking-wider text-slate-400 pt-2 border-t border-slate-50">
                  <span>Anagrafiche nel Gruppo ({groupPortfolios.length})</span>
                  {isAdminOrBO && owner && <span className="text-indigo-600">Gestito da: {owner.name} ({owner.userCode})</span>}
                </div>

                {groupPortfolios.length > 0 ? (
                  <div className="space-y-2 max-h-48 overflow-y-auto pr-1">
                    {groupPortfolios.map(p => (
                      <div key={p.id} className="flex items-center justify-between p-2.5 rounded-xl bg-slate-50 hover:bg-indigo-50/60 border border-slate-100 hover:border-indigo-200 text-xs transition-colors group">
                        <div
                          onClick={() => setSelectedPortfolio(p)}
                          className="cursor-pointer flex-1"
                          title="Clicca per aprire la scheda anagrafica"
                        >
                          <p className="font-bold text-slate-800 group-hover:text-indigo-700 group-hover:underline decoration-dotted transition-colors">{p.businessName}</p>
                          <p className="text-[10px] font-mono text-slate-400">{p.taxId} • {p.entityType}</p>
                        </div>
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            setPortfolios(prev => prev.map(item => item.id === p.id ? { ...item, appartenenzaId: undefined } : item));
                          }}
                          className="text-slate-400 hover:text-red-500 p-1 rounded-lg hover:bg-white transition-colors"
                          title="Rimuovi anagrafica dal gruppo"
                        >
                          <X size={14} />
                        </button>
                      </div>
                    ))}
                  </div>
                ) : (
                  <p className="text-xs text-slate-400 italic py-2">Nessuna anagrafica associata a questo marchio/gruppo.</p>
                )}
              </div>

              {availablePortfolios.length > 0 && (
                <div className="pt-3 border-t border-slate-100 flex items-center gap-2">
                  <select
                    className="flex-1 border border-slate-200 rounded-xl p-2.5 text-xs font-bold bg-slate-50 outline-none focus:border-indigo-500"
                    value={selectedPortfolioToAdd[a.id] || ''}
                    onChange={e => setSelectedPortfolioToAdd(prev => ({ ...prev, [a.id]: e.target.value }))}
                  >
                    <option value="">+ Associa un'anagrafica a {a.name}...</option>
                    {availablePortfolios.map(p => (
                      <option key={p.id} value={p.id}>{p.businessName} ({p.taxId})</option>
                    ))}
                  </select>
                  <button
                    type="button"
                    disabled={!selectedPortfolioToAdd[a.id]}
                    onClick={() => {
                      const pid = selectedPortfolioToAdd[a.id];
                      if (!pid) return;
                      setPortfolios(prev => prev.map(item => item.id === pid ? { ...item, appartenenzaId: a.id } : item));
                      setSelectedPortfolioToAdd(prev => ({ ...prev, [a.id]: '' }));
                    }}
                    className="bg-indigo-600 disabled:opacity-40 text-white px-4 py-2.5 rounded-xl text-xs font-black hover:bg-indigo-700 transition-colors"
                  >
                    Aggiungi
                  </button>
                </div>
              )}
            </div>
          );
        })}

        {filteredAppartenenze.length === 0 && (
          <div className="col-span-full bg-white p-16 rounded-[2.5rem] border border-slate-100 text-center text-slate-400">
            <Layers size={48} className="mx-auto mb-3 opacity-25" />
            <p className="font-black uppercase tracking-widest text-sm">Nessuna Appartenenza (Marchio / Gruppo) presente</p>
          </div>
        )}
      </div>

      {selectedPortfolio && (
        <PortfolioModal
          isOpen={!!selectedPortfolio}
          onClose={() => setSelectedPortfolio(null)}
          portfolio={portfolios.find(p => p.id === selectedPortfolio.id) || selectedPortfolio}
          setData={setPortfolios}
          user={user}
          users={users}
          partners={partners}
          appartenenze={data}
          portfolios={portfolios}
          contracts={contracts}
          contrattiTelefonici={contrattiTelefonici}
          pdps={pdps}
          onSelectPortfolio={(p) => setSelectedPortfolio(p)}
        />
      )}

      <ConfirmModal
        isOpen={!!deleteConfirmId}
        title="Conferma Eliminazione Appartenenza"
        message="Sei sicuro di voler eliminare questo gruppo di appartenenza? Le anagrafiche collegate non verranno eliminate ma perderanno il riferimento al gruppo."
        onConfirm={() => {
          const targetId = deleteConfirmId;
          setData(prev => prev.filter(a => a.id !== targetId));
          setPortfolios(prev => prev.map(p => p.appartenenzaId === targetId ? { ...p, appartenenzaId: undefined } : p));
          setDeleteConfirmId(null);
        }}
        onCancel={() => setDeleteConfirmId(null)}
      />
    </div>
  );
};

const CampagneView: React.FC<{
  data: Campagna[];
  setData: React.Dispatch<React.SetStateAction<Campagna[]>>;
  user: User;
  exportFn: (data: any[], fileName: string) => void;
}> = ({ data, setData, user, exportFn }) => {
  const isAdmin = user.role === 'admin';
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [deleteConfirmId, setDeleteConfirmId] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [form, setForm] = useState<Partial<Campagna>>({ code: '', name: '', description: '', exclusiveBenefits: '' });

  const filteredCampagne = useMemo(() => {
    if (!searchTerm) return data;
    const s = searchTerm.toLowerCase();
    return data.filter(c =>
      c.code?.toLowerCase().includes(s) ||
      c.name?.toLowerCase().includes(s) ||
      c.description?.toLowerCase().includes(s) ||
      c.exclusiveBenefits?.toLowerCase().includes(s)
    );
  }, [data, searchTerm]);

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    if (!isAdmin) return;
    const cleanCode = (form.code || '').toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 4);
    if (!cleanCode || cleanCode.length > 4) {
      return alert("Il codice identificativo deve contenere da 1 a massimo 4 caratteri alfanumerici.");
    }
    if (!form.name?.trim()) {
      return alert("Il nome della campagna è obbligatorio.");
    }
    const isDuplicateCode = data.some(c => c.code.toUpperCase() === cleanCode && c.id !== editingId);
    if (isDuplicateCode) {
      return alert(`Esiste già una campagna con codice ${cleanCode}.`);
    }

    const ts = getTimestamp();
    const payload: Campagna = {
      id: editingId || Math.random().toString(36).substr(2, 9),
      code: cleanCode,
      name: form.name.trim(),
      description: (form.description || '').trim(),
      exclusiveBenefits: (form.exclusiveBenefits || '').trim(),
      createdAt: form.createdAt || ts,
      updatedAt: ts
    };

    if (editingId) {
      setData(prev => prev.map(c => c.id === editingId ? payload : c));
    } else {
      setData(prev => [...prev, payload]);
    }
    setIsFormOpen(false);
    setEditingId(null);
    setForm({ code: '', name: '', description: '', exclusiveBenefits: '' });
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-300">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h2 className="text-3xl font-black text-purple-950 tracking-tighter flex items-center gap-2">
            <Megaphone className="text-purple-600" size={28} /> Campagne Promozionali
          </h2>
          <p className="text-slate-500 font-medium text-xs">
            Catalogo campagne associabili a Contratti (Forniture e Telefonico) e Associazioni {isAdmin ? '• Gestione Admin' : '• Consultazione'}
          </p>
        </div>
        <div className="flex flex-wrap items-center gap-3 w-full md:w-auto">
          <div className="relative flex-1 md:w-64">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
            <input
              className="w-full pl-10 pr-4 py-3 rounded-2xl border-2 border-slate-100 outline-none focus:border-purple-500 font-bold text-sm bg-white"
              placeholder="Cerca codice, nome, vantaggi..."
              value={searchTerm}
              onChange={e => setSearchTerm(e.target.value)}
            />
          </div>
          <button onClick={() => exportFn(data, 'Campagne')} className="p-3 text-slate-400 hover:text-slate-800 bg-white border border-slate-100 rounded-2xl shadow-sm transition-colors">
            <FileSpreadsheet size={20} />
          </button>
          {isAdmin && (
            <button
              onClick={() => {
                setForm({ code: '', name: '', description: '', exclusiveBenefits: '' });
                setEditingId(null);
                setIsFormOpen(true);
              }}
              className="bg-purple-600 hover:bg-purple-700 text-white px-6 py-3 rounded-2xl font-black flex items-center gap-2 shadow-lg transition-all"
            >
              <Plus size={18} /> Nuova Campagna
            </button>
          )}
        </div>
      </div>

      {isAdmin && isFormOpen && (
        <form onSubmit={handleSave} className="bg-white p-8 rounded-[2.5rem] border-2 border-purple-100 shadow-2xl space-y-6 animate-in slide-in-from-top-4">
          <div className="flex justify-between items-center border-b border-slate-100 pb-4">
            <h3 className="text-xl font-black text-purple-950">{editingId ? 'Modifica Campagna' : 'Nuova Campagna'}</h3>
            <button type="button" onClick={() => { setIsFormOpen(false); setEditingId(null); }} className="p-2 text-slate-400 hover:text-slate-600 rounded-full hover:bg-slate-100">
              <X size={20} />
            </button>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1 flex justify-between">
                <span>Codice Identificativo *</span>
                <span className="text-[10px] text-purple-600 font-black">(4 caratteri)</span>
              </label>
              <input
                required
                maxLength={4}
                placeholder="Es. C001"
                className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-mono font-black uppercase text-purple-700 focus:border-purple-500 tracking-widest"
                value={form.code || ''}
                onChange={e => setForm({ ...form, code: e.target.value.toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 4) })}
              />
            </div>
            <div className="md:col-span-2 space-y-1.5">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1">Nome Campagna *</label>
              <input
                required
                placeholder="Nome della campagna..."
                className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-purple-500"
                value={form.name || ''}
                onChange={e => setForm({ ...form, name: e.target.value })}
              />
            </div>
            <div className="col-span-full space-y-1.5">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1">Descrizione</label>
              <textarea
                rows={2}
                placeholder="Descrizione dettagliata della campagna..."
                className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-medium focus:border-purple-500"
                value={form.description || ''}
                onChange={e => setForm({ ...form, description: e.target.value })}
              />
            </div>
            <div className="col-span-full space-y-1.5">
              <label className="text-xs font-bold text-slate-500 uppercase ml-1">Vantaggi Esclusivi</label>
              <textarea
                rows={3}
                placeholder="Elenca i vantaggi esclusivi riservati agli aderenti alla campagna..."
                className="w-full border-2 p-3.5 rounded-2xl bg-purple-50/40 border-purple-100 outline-none font-bold text-purple-950 focus:border-purple-500"
                value={form.exclusiveBenefits || ''}
                onChange={e => setForm({ ...form, exclusiveBenefits: e.target.value })}
              />
            </div>
          </div>
          <div className="flex justify-end gap-3 pt-4 border-t">
            <button type="button" onClick={() => { setIsFormOpen(false); setEditingId(null); }} className="font-bold text-slate-400 px-4">Annulla</button>
            <button type="submit" className="bg-purple-600 hover:bg-purple-700 text-white px-8 py-3.5 rounded-2xl font-black shadow-lg">
              {editingId ? 'Aggiorna Campagna' : 'Salva Campagna'}
            </button>
          </div>
        </form>
      )}

      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        {filteredCampagne.map(c => (
          <div key={c.id} className="bg-white p-6 rounded-[2rem] border border-slate-100 shadow-sm hover:shadow-md transition-all space-y-4">
            <div className="flex justify-between items-start gap-4">
              <div className="flex items-center gap-3">
                <span className="px-3.5 py-2 rounded-2xl bg-purple-100 text-purple-900 font-mono font-black text-base tracking-widest border border-purple-200">
                  {c.code}
                </span>
                <div>
                  <h3 className="text-lg font-black text-slate-900">{c.name}</h3>
                  <p className="text-[10px] font-bold uppercase text-slate-400">Codice Campagna: {c.code}</p>
                </div>
              </div>
              {isAdmin && (
                <div className="flex items-center gap-1">
                  <button
                    onClick={() => {
                      setForm(c);
                      setEditingId(c.id);
                      setIsFormOpen(true);
                    }}
                    className="p-2 text-slate-400 hover:text-purple-600 rounded-xl hover:bg-purple-50 transition-colors"
                    title="Modifica Campagna"
                  >
                    <Edit2 size={16} />
                  </button>
                  <button
                    onClick={() => setDeleteConfirmId(c.id)}
                    className="p-2 text-slate-400 hover:text-red-600 rounded-xl hover:bg-red-50 transition-colors"
                    title="Elimina Campagna"
                  >
                    <Trash2 size={16} />
                  </button>
                </div>
              )}
            </div>

            {c.description && (
              <div className="space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400 tracking-wider">Descrizione</p>
                <p className="text-xs text-slate-600 font-medium leading-relaxed">{c.description}</p>
              </div>
            )}

            <div className="p-4 rounded-2xl bg-purple-50/70 border border-purple-100 space-y-1">
              <p className="text-[10px] font-black uppercase text-purple-700 tracking-wider flex items-center gap-1.5">
                <Megaphone size={12} /> Vantaggi Esclusivi
              </p>
              <p className="text-xs font-bold text-purple-950 leading-relaxed whitespace-pre-line">
                {c.exclusiveBenefits || 'Nessun vantaggio esclusivo specificato.'}
              </p>
            </div>
          </div>
        ))}

        {filteredCampagne.length === 0 && (
          <div className="col-span-full bg-white p-16 rounded-[2.5rem] border border-slate-100 text-center text-slate-400">
            <Megaphone size={48} className="mx-auto mb-3 opacity-25" />
            <p className="font-black uppercase tracking-widest text-sm">Nessuna Campagna presente</p>
          </div>
        )}
      </div>

      <ConfirmModal
        isOpen={!!deleteConfirmId}
        title="Conferma Eliminazione Campagna"
        message="Sei sicuro di voler eliminare questa campagna? Questa operazione non può essere annullata."
        onConfirm={() => {
          setData(prev => prev.filter(c => c.id !== deleteConfirmId));
          setDeleteConfirmId(null);
        }}
        onCancel={() => setDeleteConfirmId(null)}
      />
    </div>
  );
};

const StatisticsView: React.FC<{ contracts: Contract[], associazioni?: Associazione[] }> = ({ contracts, associazioni = [] }) => {
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

  const assocStats = useMemo(() => {
    const totalCount = associazioni.length;
    const totalQuota = associazioni.reduce((acc, a) => acc + (a.quota || 0), 0);
    const incassati = associazioni.filter(a => a.status === 'incassato' || a.status === 'registrato' || a.status === 'pagato' || !!a.incassato);
    const incassatiCount = incassati.length;
    const incassatoQuota = incassati.reduce((acc, a) => acc + (a.quota || 0), 0);
    const daIncassare = associazioni.filter(a => !(a.status === 'incassato' || a.status === 'registrato' || a.status === 'pagato' || !!a.incassato));
    const daIncassareCount = daIncassare.length;
    const daIncassareQuota = daIncassare.reduce((acc, a) => acc + (a.quota || 0), 0);
    const percentualeIncassato = totalQuota > 0 ? Math.round((incassatoQuota / totalQuota) * 100) : 0;
    return { totalCount, totalQuota, incassatiCount, incassatoQuota, daIncassareCount, daIncassareQuota, percentualeIncassato };
  }, [associazioni]);

  return (
    <div className="space-y-8 animate-in fade-in duration-500">
      <header>
        <h2 className="text-4xl font-black tracking-tighter text-emerald-900">Report & Statistiche</h2>
        <p className="text-slate-500 font-medium italic">Andamento mensile volumi caricati e monitoraggio incassi quote associative</p>
      </header>

      {/* Report Incassi Quote Associative */}
      <div className="bg-white p-8 rounded-[2.5rem] border border-slate-100 shadow-sm space-y-6">
        <div className="flex flex-col md:flex-row justify-between md:items-center gap-2">
          <div>
            <div className="flex items-center gap-2">
              <span className="bg-emerald-100 text-emerald-800 text-[10px] font-black uppercase px-2.5 py-0.5 rounded-full">
                Report Quote ANTICRISI
              </span>
            </div>
            <h3 className="text-2xl font-black text-slate-900 tracking-tight mt-1">Stato Incasso Quote Associative</h3>
            <p className="text-xs text-slate-500">Riepilogo quote con stato "Incassato" (baffato)</p>
          </div>
          <div className="flex items-center gap-3">
            <div className="bg-emerald-900 text-white px-5 py-2.5 rounded-2xl flex items-center gap-3">
              <span className="text-[10px] uppercase font-bold text-emerald-300">Tasso Incasso</span>
              <span className="text-2xl font-black">{assocStats.percentualeIncassato}%</span>
            </div>
          </div>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          <div className="bg-emerald-50/70 p-5 rounded-2xl border border-emerald-200">
            <p className="text-[10px] font-black text-emerald-800 uppercase tracking-widest mb-1 flex items-center gap-1.5">
              <CheckCircle2 size={14} className="text-emerald-600" /> Totale Incassato
            </p>
            <p className="text-2xl font-black text-emerald-950">€ {assocStats.incassatoQuota.toLocaleString('it-IT', { minimumFractionDigits: 2 })}</p>
            <p className="text-[10px] text-emerald-700 font-bold mt-1">{assocStats.incassatiCount} quote incassate</p>
          </div>

          <div className="bg-amber-50/70 p-5 rounded-2xl border border-amber-200">
            <p className="text-[10px] font-black text-amber-800 uppercase tracking-widest mb-1 flex items-center gap-1.5">
              <Clock size={14} className="text-amber-600" /> Da Incassare
            </p>
            <p className="text-2xl font-black text-amber-950">€ {assocStats.daIncassareQuota.toLocaleString('it-IT', { minimumFractionDigits: 2 })}</p>
            <p className="text-[10px] text-amber-700 font-bold mt-1">{assocStats.daIncassareCount} quote in attesa</p>
          </div>

          <div className="bg-slate-50 p-5 rounded-2xl border border-slate-200">
            <p className="text-[10px] font-black text-slate-600 uppercase tracking-widest mb-1">Totale Quote Censite</p>
            <p className="text-2xl font-black text-slate-800">€ {assocStats.totalQuota.toLocaleString('it-IT', { minimumFractionDigits: 2 })}</p>
            <p className="text-[10px] text-slate-500 font-bold mt-1">{assocStats.totalCount} quote complessive</p>
          </div>

          <div className="bg-blue-50/70 p-5 rounded-2xl border border-blue-200">
            <p className="text-[10px] font-black text-blue-800 uppercase tracking-widest mb-1">Avanzamento Incassi</p>
            <div className="w-full bg-blue-100 rounded-full h-3 mt-2 overflow-hidden">
              <div 
                className="bg-emerald-600 h-full rounded-full transition-all duration-500" 
                style={{ width: `${assocStats.percentualeIncassato}%` }}
              />
            </div>
            <p className="text-[10px] text-blue-800 font-bold mt-2">{assocStats.incassatiCount} su {assocStats.totalCount} quote incassate</p>
          </div>
        </div>
      </div>
      
      <div className="bg-white p-8 rounded-[2.5rem] border border-slate-100 shadow-sm h-[500px]">
        <h3 className="text-lg font-black text-slate-900 tracking-tight mb-4">Volumi Mensili Contratti Energia</h3>
        <ResponsiveContainer width="100%" height="85%">
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
  users: User[],
  partners?: Partner[],
  appartenenze?: Appartenenza[],
  portfolios?: Portfolio[],
  contracts?: Contract[],
  contrattiTelefonici?: ContrattoTelefonico[],
  pdps?: PDP[],
  onSelectPortfolio?: (p: Portfolio) => void
}> = ({ isOpen, onClose, portfolio, setData, user, users, partners = [], appartenenze = [], portfolios = [], contracts = [], contrattiTelefonici = [], pdps = [], onSelectPortfolio }) => {
  const [isEditing, setIsEditing] = useState(false);
  const [form, setForm] = useState<Partial<Portfolio>>(portfolio);
  const [openMenu, setOpenMenu] = useState<'trasmesso' | 'attivo' | null>(null);
  const [showGroupList, setShowGroupList] = useState(false);
  const canEdit = user.role === 'admin' || user.role === 'backoffice' || portfolio.assignedTo === user.id;
  const ENTITY_TYPES: PortfolioEntityType[] = ['Domestico', 'Impresa', 'Associazione', 'Condominio', 'PA'];

  const portfolioContracts = useMemo(() => {
    const energy = contracts.filter(c => c.portfolioId === portfolio.id);
    const tel = contrattiTelefonici.filter(t => t.portfolioId === portfolio.id);
    return {
      trasmesso: {
        energy: energy.filter(c => c.status === 'trasmesso'),
        tel: tel.filter(t => t.status === 'trasmesso')
      },
      attivo: {
        energy: energy.filter(c => c.status === 'attivo'),
        tel: tel.filter(t => t.status === 'attivo')
      }
    };
  }, [contracts, contrattiTelefonici, portfolio.id]);

  const groupPortfolios = useMemo(() => {
    if (!portfolio.appartenenzaId) return [];
    return portfolios.filter(p => p.appartenenzaId === portfolio.appartenenzaId);
  }, [portfolios, portfolio.appartenenzaId]);

  useEffect(() => {
    if (isOpen) {
      const parsed = splitAddress(portfolio.legalAddress);
      setForm({
        ...portfolio,
        street: portfolio.street || parsed?.street || '',
        cap: portfolio.cap || parsed?.cap || '',
        city: portfolio.city || parsed?.city || '',
        province: portfolio.province || parsed?.province || ''
      });
      setIsEditing(false);
      setOpenMenu(null);
      setShowGroupList(false);
    }
  }, [isOpen, portfolio]);

  if (!isOpen) return null;

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    const ts = getTimestamp();
    const finalLegalAddress = form.legalAddress || composeAddress(form.street, form.cap, form.city, form.province);
    const parsed = splitAddress(finalLegalAddress);
    const updatedPayload = { 
      ...form, 
      legalAddress: finalLegalAddress,
      street: form.street || parsed?.street || '',
      cap: form.cap || parsed?.cap || '',
      city: form.city || parsed?.city || '',
      province: (form.province || parsed?.province || '').toUpperCase(),
      partnerId: form.partnerId || undefined,
      appartenenzaId: form.appartenenzaId || undefined,
      updatedAt: ts, 
      updatedBy: user.name 
    };
    setData((prev: Portfolio[]) => prev.map(p => p.id === portfolio.id ? { ...p, ...updatedPayload } : p));
    setIsEditing(false);
  };

  return (
    <div className="fixed inset-0 bg-emerald-950/40 backdrop-blur-sm z-[9999] flex items-center justify-center p-4 animate-in fade-in duration-200">
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
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1">{form.entityType === 'Domestico' ? 'Nome' : 'Ragione Sociale'}</label>
                  <input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" required value={form.businessName || ''} onChange={e => setForm({...form, businessName: e.target.value})} />
                </div>
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1 flex justify-between">
                    <span>{form.entityType === 'Domestico' ? 'Codice Fiscale' : 'C.F. o P.IVA'}</span>
                    <span className="text-[9px] text-slate-400">{form.entityType === 'Domestico' ? '16 caratteri' : '11 o 16 caratteri'}</span>
                  </label>
                  <input 
                    className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-mono font-bold focus:border-emerald-500" 
                    required 
                    maxLength={16}
                    value={form.taxId || ''} 
                    onChange={e => setForm({...form, taxId: e.target.value.toUpperCase()})} 
                  />
                </div>
                <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">PEC</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" value={form.pec || ''} onChange={e => setForm({...form, pec: e.target.value})} /></div>
                <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Cellulare</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" value={form.phone || ''} onChange={e => setForm({...form, phone: e.target.value})} /></div>
                <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">E-mail</label><input type="email" className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" value={form.email || ''} onChange={e => setForm({...form, email: e.target.value})} /></div>

                {form.entityType !== 'Domestico' && (
                  <>
                    <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">CF Rapp. Legale</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-mono font-bold focus:border-emerald-500" maxLength={16} value={form.legalRepresentativeTaxId || ''} onChange={e => setForm({...form, legalRepresentativeTaxId: e.target.value.toUpperCase()})} /></div>
                    <div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Nome Rapp. Legale</label><input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold focus:border-emerald-500" value={form.legalRepresentativeName || ''} onChange={e => setForm({...form, legalRepresentativeName: e.target.value})} /></div>
                  </>
                )}
                {(user.role === 'admin' || user.role === 'backoffice') && (
                  <div className="col-span-full">
                    <SearchableSelect label="Assegna a Utente" placeholder="Cerca utente..." options={users} selectedId={form.assignedTo} onSelect={(id:string) => setForm({...form, assignedTo: id})} displayFn={(u:any)=>u.name} searchFn={(u:any)=>u.name} subTextFn={(u:any)=>u.role} icon={<UserIcon size={14}/>} />
                  </div>
                )}
                <div>
                  <SearchableSelect 
                    label="Partner" 
                    placeholder="Seleziona Partner (opzionale)..." 
                    options={partners} 
                    selectedId={form.partnerId} 
                    onSelect={(id:string) => setForm({...form, partnerId: id})} 
                    displayFn={(p:Partner)=>p.name} 
                    searchFn={(p:Partner)=>`${p.name} ${p.userCode}`} 
                    subTextFn={(p:Partner)=>`Codice: ${p.userCode}`} 
                    icon={<Handshake size={14}/>} 
                    allowClear={true}
                  />
                </div>
                <div>
                  <SearchableSelect 
                    label="Appartenenza (Marchio / Gruppo)" 
                    placeholder="Seleziona Appartenenza (opzionale)..." 
                    options={appartenenze} 
                    selectedId={form.appartenenzaId} 
                    onSelect={(id:string) => setForm({...form, appartenenzaId: id})} 
                    displayFn={(a:Appartenenza)=>a.name} 
                    searchFn={(a:Appartenenza)=>`${a.name} ${a.description || ''}`} 
                    subTextFn={(a:Appartenenza)=>a.description || 'Gruppo / Marchio'} 
                    icon={<Layers size={14}/>} 
                    allowClear={true}
                  />
                </div>
                <AddressInputGroup 
                  label="Indirizzo (Via, CAP Città, PROV)"
                  address={form.legalAddress || ''}
                  street={form.street || ''}
                  cap={form.cap || ''}
                  city={form.city || ''}
                  province={form.province || ''}
                  accentColor="emerald"
                  onChange={({ address, street, cap, city, province }) => {
                    setForm(prev => ({
                      ...prev,
                      legalAddress: address,
                      street,
                      cap,
                      city,
                      province
                    }));
                  }}
                />
                <div className="col-span-full"><label className="text-xs font-bold text-slate-500 uppercase ml-1">Allegati & Documenti (PDF / Foto)</label>
                  <FileUploader 
                    attachments={form.attachments || []} 
                    canEdit={true} 
                    onUpload={files => setForm({...form, attachments: [...(form.attachments || []), ...files]})} 
                    onRemove={id => setForm({...form, attachments: form.attachments?.filter(a => a.id !== id)})}
                    photoToPdfConfig={{
                      enabled: true,
                      mode: 'portfolio',
                      defaultPersonName: form.entityType === 'Domestico' ? (form.businessName || '') : (form.legalRepresentativeName || form.businessName || '')
                    }}
                  />
                </div>
              </div>
            </form>
          ) : (
            <div className="grid grid-cols-2 gap-6">
              <div className="space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">Tipologia Cliente</p>
                <p className="font-bold text-emerald-600">{portfolio.entityType || '---'}</p>
              </div>
              <div className="space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">{portfolio.entityType === 'Domestico' ? 'Nome' : 'Ragione Sociale'}</p>
                <p className="font-bold text-slate-700">{portfolio.businessName || '---'}</p>
              </div>
              <div className="space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">{portfolio.entityType === 'Domestico' ? 'Codice Fiscale' : 'C.F. o P.IVA'}</p>
                <p className="font-mono font-bold text-slate-700">{portfolio.taxId || '---'}</p>
              </div>
              <div className="space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">PEC</p>
                <p className="font-bold text-slate-700">{portfolio.pec || '---'}</p>
              </div>
              <div className="space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">Cellulare</p>
                <p className="font-bold text-slate-700">{portfolio.phone || '---'}</p>
              </div>
              <div className="space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">E-mail</p>
                <p className="font-bold text-slate-700">{portfolio.email || '---'}</p>
              </div>
              {portfolio.entityType !== 'Domestico' && (
                <>
                  <div className="space-y-1">
                    <p className="text-[10px] font-black uppercase text-slate-400">CF Rapp. Legale</p>
                    <p className="font-mono font-bold text-slate-700">{portfolio.legalRepresentativeTaxId || '---'}</p>
                  </div>
                  <div className="space-y-1">
                    <p className="text-[10px] font-black uppercase text-slate-400">Nome Rapp. Legale</p>
                    <p className="font-bold text-slate-700">{portfolio.legalRepresentativeName || '---'}</p>
                  </div>
                </>
              )}
              <div className="space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">Partner</p>
                {(() => {
                  const pt = partners.find(x => x.id === portfolio.partnerId);
                  const isAdminOrBO = user.role === 'admin' || user.role === 'backoffice';
                  return pt ? (
                    <span className="inline-flex items-center gap-1 bg-teal-50 text-teal-800 border border-teal-200 px-2.5 py-1 rounded-lg text-xs font-black">
                      <Handshake size={12} className="text-teal-600" /> {pt.name}{isAdminOrBO ? ` (${pt.userCode})` : ''}
                    </span>
                  ) : (
                    <p className="font-bold text-slate-400">---</p>
                  );
                })()}
              </div>
              <div className="space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">Appartenenza (Marchio / Gruppo)</p>
                {(() => {
                  const ap = appartenenze.find(x => x.id === portfolio.appartenenzaId);
                  return ap ? (
                    <div className="space-y-2">
                      <button
                        type="button"
                        onClick={() => setShowGroupList(prev => !prev)}
                        className="inline-flex items-center gap-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-800 border border-indigo-200 px-2.5 py-1 rounded-lg text-xs font-black transition-colors"
                        title="Clicca per vedere le anagrafiche del gruppo"
                      >
                        <Layers size={12} className="text-indigo-600" /> {ap.name} ({groupPortfolios.length})
                      </button>
                      {showGroupList && groupPortfolios.length > 0 && (
                        <div className="p-2.5 bg-indigo-50/40 border border-indigo-100 rounded-xl space-y-1.5 max-h-40 overflow-y-auto">
                          <p className="text-[9px] font-black uppercase text-indigo-600 tracking-wider">Anagrafiche nel gruppo (clicca per aprire):</p>
                          {groupPortfolios.map(gp => (
                            <div
                              key={gp.id}
                              onClick={() => {
                                if (onSelectPortfolio) onSelectPortfolio(gp);
                              }}
                              className={`p-2 rounded-lg text-xs flex items-center justify-between cursor-pointer transition-colors ${
                                gp.id === portfolio.id
                                  ? 'bg-indigo-600 text-white font-black'
                                  : 'bg-white hover:bg-indigo-100 text-slate-700 font-bold border border-indigo-100/60'
                              }`}
                            >
                              <span className="truncate">{gp.businessName}</span>
                              <span className={`text-[10px] font-mono ml-2 ${gp.id === portfolio.id ? 'text-indigo-100' : 'text-slate-400'}`}>{gp.taxId}</span>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  ) : (
                    <p className="font-bold text-slate-400">---</p>
                  );
                })()}
              </div>
              <div className="col-span-full space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">Indirizzo Legale</p>
                <div className="flex items-center gap-2">
                  <p className="font-bold text-slate-700">{portfolio.legalAddress}</p>
                  <AddressLink address={portfolio.legalAddress} showIconOnly className="text-blue-500" />
                </div>
              </div>

              {/* Menu Contratti (Trasmesso / Attivo) */}
              <div className="col-span-full pt-2 border-t border-slate-100 space-y-3">
                <p className="text-[10px] font-black uppercase text-slate-400 tracking-wider">Elenco Contratti Anagrafica</p>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  {(() => {
                    const countTrasmesso = portfolioContracts.trasmesso.energy.length + portfolioContracts.trasmesso.tel.length;
                    const isOpenTrasmesso = openMenu === 'trasmesso';
                    return (
                      <button
                        type="button"
                        onClick={() => setOpenMenu(isOpenTrasmesso ? null : 'trasmesso')}
                        className={`flex items-center justify-between p-3.5 rounded-2xl border text-left transition-all ${
                          isOpenTrasmesso
                            ? 'bg-blue-600 text-white border-blue-600 shadow-md'
                            : 'bg-blue-50/60 hover:bg-blue-100/70 text-blue-950 border-blue-200'
                        }`}
                      >
                        <div>
                          <p className={`text-[10px] font-black uppercase tracking-widest ${isOpenTrasmesso ? 'text-blue-100' : 'text-blue-600'}`}>Menu Stato</p>
                          <p className="text-sm font-black">Contratti Trasmessi</p>
                        </div>
                        <span className={`px-2.5 py-1 rounded-xl text-xs font-black ${isOpenTrasmesso ? 'bg-white text-blue-700' : 'bg-blue-600 text-white'}`}>
                          {countTrasmesso}
                        </span>
                      </button>
                    );
                  })()}

                  {(() => {
                    const countAttivo = portfolioContracts.attivo.energy.length + portfolioContracts.attivo.tel.length;
                    const isOpenAttivo = openMenu === 'attivo';
                    return (
                      <button
                        type="button"
                        onClick={() => setOpenMenu(isOpenAttivo ? null : 'attivo')}
                        className={`flex items-center justify-between p-3.5 rounded-2xl border text-left transition-all ${
                          isOpenAttivo
                            ? 'bg-emerald-600 text-white border-emerald-600 shadow-md'
                            : 'bg-emerald-50/60 hover:bg-emerald-100/70 text-emerald-950 border-emerald-200'
                        }`}
                      >
                        <div>
                          <p className={`text-[10px] font-black uppercase tracking-widest ${isOpenAttivo ? 'text-emerald-100' : 'text-emerald-600'}`}>Menu Stato</p>
                          <p className="text-sm font-black">Contratti Attivi</p>
                        </div>
                        <span className={`px-2.5 py-1 rounded-xl text-xs font-black ${isOpenAttivo ? 'bg-white text-emerald-700' : 'bg-emerald-600 text-white'}`}>
                          {countAttivo}
                        </span>
                      </button>
                    );
                  })()}
                </div>

                {openMenu && (
                  <div className="p-4 rounded-2xl bg-slate-50 border border-slate-200/80 space-y-3 animate-in fade-in duration-200">
                    <div className="flex items-center justify-between">
                      <p className="text-xs font-black uppercase tracking-wider text-slate-700">
                        Elenco Contratti: <span className={openMenu === 'attivo' ? 'text-emerald-600' : 'text-blue-600'}>{openMenu.toUpperCase()}</span>
                      </p>
                      <button type="button" onClick={() => setOpenMenu(null)} className="text-slate-400 hover:text-slate-600">
                        <X size={14} />
                      </button>
                    </div>

                    {portfolioContracts[openMenu].energy.length === 0 && portfolioContracts[openMenu].tel.length === 0 ? (
                      <p className="text-xs text-slate-400 italic py-2">Nessun contratto nello stato "{openMenu}" per questa anagrafica.</p>
                    ) : (
                      <div className="space-y-2 max-h-60 overflow-y-auto pr-1">
                        {portfolioContracts[openMenu].energy.map(c => {
                          const pdp = pdps.find(p => p.id === c.pdpId);
                          return (
                            <div key={c.id} className="p-3 bg-white rounded-xl border border-slate-100 flex items-center justify-between gap-3 text-xs shadow-sm">
                              <div className="space-y-0.5">
                                <div className="flex items-center gap-1.5">
                                  {c.service === 'POWER' ? <Zap size={13} className="text-amber-500" /> : <Flame size={13} className="text-orange-500" />}
                                  <span className="font-black text-slate-800">{c.service} • {c.supplier}</span>
                                  <span className="text-[10px] font-mono bg-slate-100 px-1.5 py-0.5 rounded text-slate-600">{pdp?.pdpCode || 'N/D'}</span>
                                </div>
                                <p className="text-[10px] text-slate-400">
                                  {pdp?.address ? `${pdp.address} • ` : ''}Volume: {c.volume?.toLocaleString()} {c.service === 'POWER' ? 'KWh' : 'smc'}
                                </p>
                              </div>
                              <div className="text-right shrink-0">
                                <span className={`px-2 py-0.5 rounded-full text-[9px] font-black uppercase ${c.status === 'attivo' ? 'bg-emerald-100 text-emerald-700' : 'bg-blue-100 text-blue-700'}`}>
                                  {c.status}
                                </span>
                                {c.startDate && <p className="text-[9px] text-slate-400 mt-1">Dal {formatDate(c.startDate)}</p>}
                              </div>
                            </div>
                          );
                        })}

                        {portfolioContracts[openMenu].tel.map(t => (
                          <div key={t.id} className="p-3 bg-white rounded-xl border border-slate-100 flex items-center justify-between gap-3 text-xs shadow-sm">
                            <div className="space-y-0.5">
                              <div className="flex items-center gap-1.5">
                                <Phone size={13} className="text-sky-500" />
                                <span className="font-black text-slate-800">TELEFONICO • {t.operatore}</span>
                                <span className="text-[10px] font-bold bg-sky-50 text-sky-700 px-1.5 py-0.5 rounded">{t.tipoOperazione}</span>
                              </div>
                              <p className="text-[10px] text-slate-400">
                                {t.offerta || 'Offerta standard'} {t.codiceMigrazione ? `• Cod. Migrazione: ${t.codiceMigrazione}` : ''}
                              </p>
                            </div>
                            <div className="text-right shrink-0">
                              <span className={`px-2 py-0.5 rounded-full text-[9px] font-black uppercase ${t.status === 'attivo' ? 'bg-emerald-100 text-emerald-700' : 'bg-blue-100 text-blue-700'}`}>
                                {t.status}
                              </span>
                              {t.createdAt && <p className="text-[9px] text-slate-400 mt-1">{formatDate(t.createdAt)}</p>}
                            </div>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                )}
              </div>
              {portfolio.updatedBy && (
                <div className="col-span-full p-3 bg-emerald-50 rounded-xl border border-emerald-100 flex items-center gap-2">
                  <Clock size={14} className="text-emerald-600" />
                  <p className="text-[10px] font-black text-emerald-900 uppercase tracking-widest">Ultima modifica effettuata da: <span className="text-emerald-600">{portfolio.updatedBy}</span></p>
                </div>
              )}
              <div className="col-span-full space-y-3 pt-2 border-t border-slate-100">
                <p className="text-[10px] font-black uppercase text-slate-400">Documenti Allegati & Caricamento Rapido</p>
                <FileUploader
                  attachments={portfolio.attachments || []}
                  canEdit={canEdit}
                  onUpload={files => {
                    const updated = [...(portfolio.attachments || []), ...files];
                    setData((prev: Portfolio[]) => prev.map(p => p.id === portfolio.id ? { ...p, attachments: updated, updatedAt: getTimestamp(), updatedBy: user.name } : p));
                  }}
                  onRemove={id => {
                    const updated = (portfolio.attachments || []).filter(a => a.id !== id);
                    setData((prev: Portfolio[]) => prev.map(p => p.id === portfolio.id ? { ...p, attachments: updated, updatedAt: getTimestamp(), updatedBy: user.name } : p));
                  }}
                  photoToPdfConfig={{
                    enabled: true,
                    mode: 'portfolio',
                    defaultPersonName: portfolio.entityType === 'Domestico' ? (portfolio.businessName || '') : (portfolio.legalRepresentativeName || portfolio.businessName || '')
                  }}
                />
              </div>
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
      const parsed = splitAddress(pdp.address);
      setForm({
        ...pdp,
        street: pdp.street || parsed?.street || '',
        cap: pdp.cap || parsed?.cap || '',
        city: pdp.city || parsed?.city || '',
        province: pdp.province || parsed?.province || ''
      });
      setIsEditing(false);
    }
  }, [isOpen, pdp]);

  if (!isOpen) return null;

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    const svc = form.service || 'POWER';
    if (svc === 'METANO' && form.portata) {
      if (!form.portata.toUpperCase().startsWith('G')) {
        return alert("Errore: La portata per il servizio METANO deve iniziare obbligatoriamente con la lettera 'G' maiuscola (es. G4, G6, G10).");
      }
    }
    const ts = getTimestamp();
    const cpName = form.cabinaPrimaria || form.technicalSpecs;
    const finalAddress = form.address || composeAddress(form.street, form.cap, form.city, form.province);
    const parsed = splitAddress(finalAddress);

    const updatedPayload: Partial<PDP> = {
      ...form,
      service: svc,
      pdpCode: (form.pdpCode || '').toUpperCase().trim(),
      address: finalAddress,
      street: form.street || parsed?.street || '',
      cap: form.cap || parsed?.cap || '',
      city: form.city || parsed?.city || '',
      province: (form.province || parsed?.province || '').toUpperCase(),
      matricola: form.matricola ? form.matricola.toUpperCase().trim() : '',
      cabinaPrimaria: svc === 'POWER' ? cpName : '',
      technicalSpecs: svc === 'POWER' ? cpName : '',
      potenzaDisponibile: svc === 'POWER' ? Number(form.potenzaDisponibile || 0) : 0,
      potenzaImpegnata: svc === 'POWER' ? Number(form.potenzaImpegnata || 0) : 0,
      tensione: svc === 'POWER' ? formatTensione(form.tensione) : '',
      remi: svc === 'METANO' ? (form.remi || '').toUpperCase().trim() : '',
      portata: svc === 'METANO' ? (form.portata || '').toUpperCase().trim() : '',
      updatedAt: ts
    };
    setData((prev: PDP[]) => prev.map(p => p.id === pdp.id ? { ...p, ...updatedPayload } : p));
    setIsEditing(false);
  };

  const isMetano = (form.service === 'METANO') || (!form.service && pdp.service === 'METANO');

  return (
    <div className="fixed inset-0 bg-slate-950/40 backdrop-blur-sm z-[9999] flex items-center justify-center p-4 animate-in fade-in duration-200">
      <div className="bg-white w-full max-w-2xl rounded-[2.5rem] shadow-2xl overflow-hidden animate-in zoom-in-95 duration-200">
        <div className="p-8 border-b border-slate-100 flex justify-between items-center bg-yellow-50/50">
          <div>
            <div className="flex items-center gap-2">
              <span className={`px-2 py-0.5 rounded-md text-[9px] font-black uppercase tracking-wider flex items-center gap-1 ${
                isMetano ? 'bg-blue-100 text-blue-800' : 'bg-yellow-100 text-yellow-800'
              }`}>
                {isMetano ? <Flame size={10} /> : <Zap size={10} />}
                {isMetano ? 'METANO' : 'POWER'}
              </span>
              <h3 className="text-2xl font-black text-slate-900 tracking-tight font-mono">{isEditing ? 'Modifica PDP' : pdp.pdpCode}</h3>
            </div>
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
                {/* Tipo Servizio */}
                <div className="col-span-full space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1">Tipo Servizio</label>
                  <div className="flex gap-4">
                    {(['POWER', 'METANO'] as const).map(s => (
                      <button 
                        key={s} 
                        type="button" 
                        onClick={() => setForm(prev => ({ ...prev, service: s }))} 
                        className={`flex-1 p-3 rounded-2xl font-black transition-all border-2 flex items-center justify-center gap-2 ${
                          (form.service || 'POWER') === s 
                            ? 'bg-yellow-500 text-white border-yellow-600 shadow-md' 
                            : 'bg-slate-50 text-slate-400 border-slate-100 hover:bg-slate-100'
                        }`}
                      >
                        {s === 'POWER' ? <Zap size={16} /> : <Flame size={16} />}
                        {s}
                      </button>
                    ))}
                  </div>
                </div>

                {/* PDP */}
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1 flex justify-between">
                    <span>PDP</span>
                    <span className="text-[10px] text-slate-400">14 caratteri</span>
                  </label>
                  <input maxLength={14} minLength={14} pattern=".{14}" title="Il codice PDP deve essere di 14 caratteri" className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-mono font-black uppercase" required value={form.pdpCode || ''} onChange={e => setForm({...form, pdpCode: e.target.value.toUpperCase()})} />
                </div>

                {/* Matricola */}
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase ml-1">Matricola</label>
                  <input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold uppercase" placeholder="Matricola contatore" value={form.matricola || ''} onChange={e => setForm({...form, matricola: e.target.value.toUpperCase()})} />
                </div>

                {/* Indirizzo Fornitura Confluente */}
                <AddressInputGroup 
                  label="Indirizzo Fornitura (Via, CAP Città, PROV)"
                  address={form.address || ''}
                  street={form.street || ''}
                  cap={form.cap || ''}
                  city={form.city || ''}
                  province={form.province || ''}
                  accentColor="yellow"
                  onChange={({ address, street, cap, city, province }) => {
                    setForm(prev => ({
                      ...prev,
                      address,
                      street,
                      cap,
                      city,
                      province
                    }));
                  }}
                />

                {/* Se POWER */}
                {(form.service || 'POWER') === 'POWER' ? (
                  <>
                    <div className="space-y-1.5">
                      <label className="text-xs font-bold text-slate-500 uppercase ml-1">Cabina Primaria</label>
                      <input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold" placeholder="CP ..." value={form.cabinaPrimaria || form.technicalSpecs || ''} onChange={e => setForm({...form, cabinaPrimaria: e.target.value, technicalSpecs: e.target.value})} />
                    </div>
                    <div className="space-y-1.5">
                      <label className="text-xs font-bold text-slate-500 uppercase ml-1">Pot. Disp. (KWp)</label>
                      <input type="number" step="0.1" className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold" value={form.potenzaDisponibile ?? ''} onChange={e => setForm({...form, potenzaDisponibile: Number(e.target.value)})} />
                    </div>
                    <div className="space-y-1.5">
                      <label className="text-xs font-bold text-slate-500 uppercase ml-1">Pot. Imp. (KWp)</label>
                      <input type="number" step="0.1" className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold" value={form.potenzaImpegnata ?? ''} onChange={e => setForm({...form, potenzaImpegnata: Number(e.target.value)})} />
                    </div>
                    <div className="space-y-1.5">
                      <label className="text-xs font-bold text-slate-500 uppercase ml-1">Tensione</label>
                      <select 
                        className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold"
                        value={formatTensione(form.tensione)}
                        onChange={e => setForm({...form, tensione: e.target.value})}
                      >
                        <option value="">Seleziona Tensione...</option>
                        {TENSIONE_OPTIONS.map(opt => (
                          <option key={opt.value} value={opt.value}>{opt.label} V</option>
                        ))}
                      </select>
                    </div>
                  </>
                ) : (
                  <>
                    <div className="space-y-1.5">
                      <label className="text-xs font-bold text-slate-500 uppercase ml-1">REMI</label>
                      <input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold uppercase" placeholder="Codice punto REMI" value={form.remi || ''} onChange={e => setForm({...form, remi: e.target.value.toUpperCase()})} />
                    </div>
                    <div className="space-y-1.5">
                      <label className="text-xs font-bold text-slate-500 uppercase ml-1 flex justify-between">
                        <span>Portata</span>
                        <span className="text-[10px] text-amber-600 font-bold">Inizia con 'G'</span>
                      </label>
                      <input className="w-full border-2 p-3.5 rounded-2xl bg-slate-50 outline-none font-bold uppercase" placeholder="Es: G4, G6..." value={form.portata || ''} onChange={e => setForm({...form, portata: e.target.value.toUpperCase()})} />
                    </div>
                  </>
                )}
              </div>
            </form>
          ) : (
            <div className="grid grid-cols-2 gap-6">
              <div className="col-span-full space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">Indirizzo Fornitura</p>
                <div className="flex items-center gap-2">
                  <p className="font-bold text-slate-700">{pdp.address || '---'}</p>
                  {pdp.address && <AddressLink address={pdp.address} className="text-blue-500" />}
                </div>
              </div>
              <div className="space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">Matricola</p>
                <p className="font-bold text-slate-700">{pdp.matricola || '---'}</p>
              </div>
              <div className="space-y-1">
                <p className="text-[10px] font-black uppercase text-slate-400">Servizio</p>
                <p className="font-bold text-slate-700">{pdp.service || 'POWER'}</p>
              </div>
              {pdp.service === 'METANO' ? (
                <>
                  <div className="space-y-1">
                    <p className="text-[10px] font-black uppercase text-slate-400">REMI</p>
                    <p className="font-mono font-bold text-slate-700">{pdp.remi || '---'}</p>
                  </div>
                  <div className="space-y-1">
                    <p className="text-[10px] font-black uppercase text-slate-400">Portata</p>
                    <p className="font-bold text-blue-700">{pdp.portata || '---'}</p>
                  </div>
                </>
              ) : (
                <>
                  <div className="space-y-1">
                    <p className="text-[10px] font-black uppercase text-slate-400">Cabina Primaria</p>
                    <p className="font-bold text-slate-700">{pdp.cabinaPrimaria || pdp.technicalSpecs || '---'}</p>
                  </div>
                  <div className="space-y-1">
                    <p className="text-[10px] font-black uppercase text-slate-400">Tensione</p>
                    <p className="font-bold text-slate-700">{pdp.tensione ? `${formatTensione(pdp.tensione)} V` : '---'}</p>
                  </div>
                  <div className="space-y-1">
                    <p className="text-[10px] font-black uppercase text-slate-400">Potenza Disponibile</p>
                    <p className="font-bold text-slate-700">{pdp.potenzaDisponibile || 0} kWp</p>
                  </div>
                  <div className="space-y-1">
                    <p className="text-[10px] font-black uppercase text-slate-400">Potenza Impegnata</p>
                    <p className="font-bold text-slate-700">{pdp.potenzaImpegnata || 0} kWp</p>
                  </div>
                </>
              )}
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

const PDPDetailModal: React.FC<{
  isOpen: boolean;
  onClose: () => void;
  pdp: PDP;
  contracts: Contract[];
  portfolios: Portfolio[];
  casi: Caso[];
  cabine: CabinaPrimaria[];
}> = ({ isOpen, onClose, pdp, contracts, portfolios, casi, cabine }) => {
  if (!isOpen || !pdp) return null;

  // i casi mettili in ordine di stato dando priorità ai casi aperti
  const statusWeight: Record<string, number> = {
    'nuovo': 1,
    'in lavorazione': 2,
    'partner': 3,
    'KO': 4,
    'risolto': 5
  };

  const associatedContracts = contracts
    .filter((c: any) => c.pdpId === pdp.id)
    .sort((a, b) => {
      // nel popup i contratti mettili in ordine di data di creazione (most recent first)
      const dateA = a.createdAt || a.updatedAt || '';
      const dateB = b.createdAt || b.updatedAt || '';
      return dateB.localeCompare(dateA);
    });

  const associatedCases = casi
    .filter((c: any) => c.pdpId === pdp.id)
    .sort((a, b) => {
      const weightA = statusWeight[a.status] || 99;
      const weightB = statusWeight[b.status] || 99;
      if (weightA !== weightB) {
        return weightA - weightB;
      }
      // if same status, sort by date descending (most recent first)
      const dateA = a.createdAt || a.updatedAt || '';
      const dateB = b.createdAt || b.updatedAt || '';
      return dateB.localeCompare(dateA);
    });

  return (
    <div className="fixed inset-0 z-[9999] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-200">
      <div className="bg-white w-full max-w-4xl rounded-[2.5rem] shadow-2xl overflow-hidden animate-in zoom-in-95 duration-200 flex flex-col max-h-[90vh]">
        {/* Header */}
        <div className="p-8 bg-emerald-900 text-white relative flex-shrink-0">
          <button 
            onClick={onClose} 
            className="absolute top-6 right-6 p-2 hover:bg-white/10 rounded-full transition-colors"
          >
            <X size={24} />
          </button>
          <h3 className="text-2xl font-black tracking-tight flex items-center gap-2">
            {pdp.service === 'METANO' ? <Flame className="text-blue-400" size={24} /> : <Zap className="text-yellow-400" size={24} />} 
            Scheda Tecnica Punto di Prelievo
          </h3>
          <div className="flex items-center gap-3 mt-2">
            <span className={`px-2.5 py-0.5 rounded-full text-xs font-black uppercase tracking-wider ${
              pdp.service === 'METANO' ? 'bg-blue-500/30 text-blue-200 border border-blue-400/40' : 'bg-yellow-500/30 text-yellow-200 border border-yellow-400/40'
            }`}>
              {pdp.service || 'POWER'}
            </span>
            <p className="text-emerald-300 text-xs font-bold uppercase tracking-widest">
              PDP: <span className="font-mono font-black text-white bg-emerald-800/80 px-2 py-0.5 rounded ml-1 text-sm">{pdp.pdpCode}</span>
            </p>
            {pdp.matricola && (
              <p className="text-emerald-300 text-xs font-bold uppercase tracking-widest">
                Matricola: <span className="font-mono font-black text-white bg-emerald-800/80 px-2 py-0.5 rounded ml-1 text-sm">{pdp.matricola}</span>
              </p>
            )}
          </div>
        </div>

        {/* Body */}
        <div className="p-8 overflow-y-auto space-y-8 flex-1">
          {/* Sezione 1: Dati Anagrafici e Tecnici */}
          <div className="space-y-4">
            <h4 className="text-sm font-black uppercase text-emerald-800 tracking-wider border-b pb-2 flex items-center gap-2">
              <Info size={16} /> Dati Generali e Specifiche Tecniche
            </h4>
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
              <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                <p className="text-[10px] font-bold text-slate-400 uppercase">PDP</p>
                <p className="font-mono font-black text-slate-800 text-base mt-1">{pdp.pdpCode}</p>
              </div>

              <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                <p className="text-[10px] font-bold text-slate-400 uppercase">Matricola</p>
                <p className="font-mono font-bold text-slate-800 text-sm mt-1">{pdp.matricola || 'N/D'}</p>
              </div>

              <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                <p className="text-[10px] font-bold text-slate-400 uppercase">Servizio</p>
                <span className={`inline-flex items-center gap-1.5 px-3 py-1 rounded-xl text-xs font-black uppercase tracking-wider mt-1 ${
                  pdp.service === 'METANO' ? 'bg-blue-100 text-blue-800' : 'bg-yellow-100 text-yellow-800'
                }`}>
                  {pdp.service === 'METANO' ? <Flame size={14} /> : <Zap size={14} />}
                  {pdp.service || 'POWER'}
                </span>
              </div>

              <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                <p className="text-[10px] font-bold text-slate-400 uppercase">Indirizzo Fornitura</p>
                <div className="flex items-start gap-2 mt-1">
                  <p className="font-bold text-slate-800 text-sm">{pdp.address || 'N/D'}</p>
                  {pdp.address && (
                    <AddressLink address={pdp.address} showIconOnly className="text-emerald-600 hover:text-emerald-800 mt-0.5" />
                  )}
                </div>
              </div>

              <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                <p className="text-[10px] font-bold text-slate-400 uppercase">Località</p>
                <p className="font-bold text-slate-800 text-sm mt-1">
                  {pdp.city || 'N/D'} {pdp.province ? `(${pdp.province})` : ''}
                </p>
                <p className="text-xs text-slate-500">{pdp.street || ''} {pdp.cap ? `CAP ${pdp.cap}` : ''}</p>
              </div>

              {pdp.service === 'METANO' ? (
                <>
                  <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                    <p className="text-[10px] font-bold text-slate-400 uppercase">REMI</p>
                    <p className="font-mono font-black text-slate-800 text-base mt-1">{pdp.remi || 'N/D'}</p>
                  </div>

                  <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                    <p className="text-[10px] font-bold text-slate-400 uppercase">Portata</p>
                    <p className="font-bold text-blue-700 text-base mt-1">{pdp.portata || 'N/D'}</p>
                  </div>
                </>
              ) : (
                <>
                  <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                    <p className="text-[10px] font-bold text-slate-400 uppercase">Cabina Primaria</p>
                    <p className="font-bold text-slate-800 text-sm mt-1">
                      {pdp.cabinaPrimaria || (pdp.cabinaPrimariaId 
                        ? cabine.find((c: any) => c.id === pdp.cabinaPrimariaId)?.name || 'N/D' 
                        : pdp.technicalSpecs || 'Nessuna cabina associata')}
                    </p>
                  </div>

                  <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                    <p className="text-[10px] font-bold text-slate-400 uppercase">Tensione</p>
                    <p className="font-bold text-slate-800 text-sm mt-1">
                      {pdp.tensione ? `${formatTensione(pdp.tensione)} V` : 'N/D'}
                    </p>
                  </div>

                  <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                    <p className="text-[10px] font-bold text-slate-400 uppercase">Pot. Disp. (KWp)</p>
                    <p className="font-black text-emerald-700 text-lg mt-1">{pdp.potenzaDisponibile || 0} kWp</p>
                  </div>

                  <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                    <p className="text-[10px] font-bold text-slate-400 uppercase">Pot. Imp. (KWp)</p>
                    <p className="font-black text-emerald-700 text-lg mt-1">{pdp.potenzaImpegnata || 0} kWp</p>
                  </div>
                </>
              )}
            </div>
          </div>

          {/* Sezione 2: Contratti Collegati */}
          <div className="space-y-4">
            <h4 className="text-sm font-black uppercase text-emerald-800 tracking-wider border-b pb-2 flex items-center gap-2">
              <FileText size={16} /> Contratti Collegati
            </h4>
            {associatedContracts.length === 0 ? (
              <div className="p-6 bg-slate-50 border border-slate-100 rounded-3xl text-center">
                <p className="text-sm text-slate-400 font-bold uppercase tracking-wider">Nessun contratto associato a questo PDP</p>
              </div>
            ) : (
              <div className="bg-white rounded-3xl overflow-hidden border border-slate-100 shadow-sm overflow-x-auto">
                <table className="w-full text-left min-w-[600px]">
                  <thead className="bg-slate-50 border-b">
                    <tr>
                      <th className="p-4 text-[9px] font-black uppercase text-slate-400">N. Contratto</th>
                      <th className="p-4 text-[9px] font-black uppercase text-slate-400">Cliente (Business)</th>
                      <th className="p-4 text-[9px] font-black uppercase text-slate-400">Servizio / Fornitore</th>
                      <th className="p-4 text-[9px] font-black uppercase text-slate-400">Validità</th>
                      <th className="p-4 text-[9px] font-black uppercase text-slate-400">Stato</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-50">
                    {associatedContracts.map((c: any) => {
                      const port = portfolios.find((pf: any) => pf.id === c.portfolioId);
                      const statusColor = 
                        c.status === 'attivo' ? 'bg-emerald-100 text-emerald-600' :
                        c.status === 'trasmesso' ? 'bg-blue-100 text-blue-600' :
                        c.status === 'bozza' ? 'bg-slate-100 text-slate-500' :
                        c.status === 'da firmare' ? 'bg-amber-100 text-amber-600' :
                        'bg-red-100 text-red-600';
                      
                      return (
                        <tr key={c.id} className="hover:bg-slate-50/50 transition-colors">
                          <td className="p-4 font-bold text-xs">{c.contractNumber || '---'}</td>
                          <td className="p-4">
                            <p className="font-bold text-slate-800 text-xs">{port?.businessName || '---'}</p>
                            <p className="text-[9px] font-black text-slate-400 uppercase">{port?.taxId}</p>
                          </td>
                          <td className="p-4 text-xs font-bold">
                            <span className="flex items-center gap-1.5">
                              {c.service === 'POWER' ? (
                                <Lightbulb size={12} className="text-red-500" />
                              ) : (
                                <Flame size={12} className="text-blue-500" />
                              )}
                              {c.service} - <span className="text-emerald-700 font-black">{c.fornitore}</span>
                            </span>
                          </td>
                          <td className="p-4 text-xs">
                            <p className="font-bold text-slate-600">{formatDate(c.startDate)} al {formatDate(c.endDate)}</p>
                            <p className="text-[9px] text-slate-400">{c.durationMonths} mesi</p>
                          </td>
                          <td className="p-4">
                            <span className={`px-2 py-0.5 rounded-full text-[9px] font-black uppercase tracking-widest ${statusColor}`}>
                              {c.status}
                            </span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>

          {/* Sezione 3: Ticket / Casi Collegati */}
          <div className="space-y-4">
            <h4 className="text-sm font-black uppercase text-emerald-800 tracking-wider border-b pb-2 flex items-center gap-2">
              <ShieldAlert size={16} /> Casi / Assistenza Collegati
            </h4>
            {associatedCases.length === 0 ? (
              <div className="p-6 bg-slate-50 border border-slate-100 rounded-3xl text-center">
                <p className="text-sm text-slate-400 font-bold uppercase tracking-wider">Nessun caso aperto associato a questo PDP</p>
              </div>
            ) : (
              <div className="bg-white rounded-3xl overflow-hidden border border-slate-100 shadow-sm overflow-x-auto">
                <table className="w-full text-left min-w-[600px]">
                  <thead className="bg-slate-50 border-b">
                    <tr>
                      <th className="p-4 text-[9px] font-black uppercase text-slate-400">Categoria</th>
                      <th className="p-4 text-[9px] font-black uppercase text-slate-400">Titolo / Descrizione</th>
                      <th className="p-4 text-[9px] font-black uppercase text-slate-400">Stato</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-50">
                    {associatedCases.map((c: any) => {
                      const caseStatusColor = 
                        c.status === 'risolto' ? 'bg-emerald-100 text-emerald-600' :
                        c.status === 'in lavorazione' ? 'bg-amber-100 text-amber-600' :
                        c.status === 'nuovo' ? 'bg-blue-100 text-blue-600' :
                        'bg-red-100 text-red-600';
                      
                      return (
                        <tr key={c.id} className="hover:bg-slate-50/50 transition-colors">
                          <td className="p-4">
                            <span className="bg-slate-100 text-slate-600 px-2.5 py-1 rounded-xl text-[9px] font-black uppercase">
                              {c.category}
                            </span>
                          </td>
                          <td className="p-4">
                            <p className="font-bold text-slate-800 text-xs">{c.title}</p>
                            <p className="text-[10px] text-slate-400 mt-0.5 line-clamp-1">{c.description}</p>
                          </td>
                          <td className="p-4">
                            <span className={`px-2 py-0.5 rounded-full text-[9px] font-black uppercase tracking-widest ${caseStatusColor}`}>
                              {c.status}
                            </span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>

        {/* Footer */}
        <div className="p-6 bg-slate-50 border-t border-slate-100 flex justify-end gap-3 flex-shrink-0">
          <button 
            onClick={onClose} 
            className="bg-emerald-900 hover:bg-emerald-800 text-white px-8 py-3 rounded-xl font-black text-sm transition-colors"
          >
            Chiudi
          </button>
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
    <div className="fixed inset-0 z-[9999] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-300">
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
  const [deleteConfirmId, setDeleteConfirmId] = useState<string | null>(null);

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
                  <td className="p-4 font-bold text-slate-800 text-[11px]">{portfolio?.businessName || 'N/D'}</td>
                  <td className="p-4 font-mono text-[11px]">{s.powerVolume.toLocaleString()}</td>
                  <td className="p-4 font-mono text-[11px]">{s.methaneVolume.toLocaleString()}</td>
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
                    <button onClick={() => setDeleteConfirmId(s.id)} className="text-slate-300 hover:text-red-500 transition-colors">
                      <Trash2 size={16} />
                    </button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      <ConfirmModal
        isOpen={!!deleteConfirmId}
        title="Conferma Eliminazione"
        message="Sei sicuro di voler eliminare questa richiesta di Credit Check? Questa operazione non può essere annullata."
        onConfirm={() => {
          setData(prev => prev.filter(item => item.id !== deleteConfirmId));
          setDeleteConfirmId(null);
        }}
        onCancel={() => setDeleteConfirmId(null)}
      />
    </div>
  );
};

const MapView: React.FC<{ pdps: PDP[], portfolios: Portfolio[], contracts: Contract[], onViewPdp?: (p: PDP) => void }> = ({ pdps, portfolios, contracts, onViewPdp }) => {
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
              <div className="pt-2 border-t mt-2 flex justify-between items-center gap-4">
                <AddressLink address={p.address} label="Google Maps" className="text-blue-600 text-[10px] font-bold shrink-0" />
                <button onClick={() => onViewPdp && onViewPdp(p)} className="text-emerald-700 hover:text-emerald-900 text-[10px] font-black uppercase tracking-tight cursor-pointer shrink-0">
                  Dettagli
                </button>
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
        
        <div className="absolute bottom-6 left-6 z-[500] bg-white/90 backdrop-blur p-4 rounded-2xl border border-slate-200 shadow-xl max-w-xs">
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
  setCreditChecks: React.Dispatch<React.SetStateAction<CreditCheckRequest[]>>,
  users: User[],
  currentUser: User | null
}> = ({ setPortfolios, setPdps, setContracts, setCabine, setCasi, setCreditChecks, users, currentUser }) => {
  const [jsonInput, setJsonInput] = useState('');
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');
  const [assignmentMode, setAssignmentMode] = useState<'auto_with_default' | 'force_user'>('auto_with_default');
  const [selectedDefaultUser, setSelectedDefaultUser] = useState<string>(() => currentUser?.id || users[0]?.id || '1');
  const [uploadSummary, setUploadSummary] = useState<{
    portfolios: number;
    pdps: number;
    contracts: number;
    userCounts: Record<string, number>;
    unmatchedCodes: string[];
  } | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Helper per la risoluzione dell'agente / codice utente
  const createResolver = () => {
    const userCounts: Record<string, number> = {};
    const unmatchedCodes = new Set<string>();

    const resolveUser = (rawVal: any, inheritedUserId?: string): string => {
      // 1. Modalità Forzata: assegna tutto all'agente selezionato
      if (assignmentMode === 'force_user' && selectedDefaultUser) {
        userCounts[selectedDefaultUser] = (userCounts[selectedDefaultUser] || 0) + 1;
        return selectedDefaultUser;
      }

      // 2. Se è specificato un codice / matricola nel file
      if (rawVal !== undefined && rawVal !== null && String(rawVal).trim() !== '') {
        const val = String(rawVal).trim();
        const valUpper = val.toUpperCase();
        const valLower = val.toLowerCase();

        // Cerca per userCode (es. AG01, ADM01, BO01 o matricola numerica)
        const byCode = users.find(u => (u.userCode || '').trim().toUpperCase() === valUpper);
        if (byCode) {
          userCounts[byCode.id] = (userCounts[byCode.id] || 0) + 1;
          return byCode.id;
        }

        // Cerca per username
        const byUsername = users.find(u => (u.username || '').trim().toLowerCase() === valLower);
        if (byUsername) {
          userCounts[byUsername.id] = (userCounts[byUsername.id] || 0) + 1;
          return byUsername.id;
        }

        // Cerca per id
        const byId = users.find(u => u.id === val);
        if (byId) {
          userCounts[byId.id] = (userCounts[byId.id] || 0) + 1;
          return byId.id;
        }

        // Cerca per nome
        const byName = users.find(u => (u.name || '').trim().toLowerCase() === valLower);
        if (byName) {
          userCounts[byName.id] = (userCounts[byName.id] || 0) + 1;
          return byName.id;
        }

        // Codice non riconosciuto: memorizzalo per notificarlo all'utente
        unmatchedCodes.add(val);
      }

      // 3. Ereditarietà: se presente utente del Cliente collegato (es. contratto o PDP)
      if (inheritedUserId) {
        userCounts[inheritedUserId] = (userCounts[inheritedUserId] || 0) + 1;
        return inheritedUserId;
      }

      // 4. Fallback all'utente predefinito impostato da interfaccia
      const fallbackId = selectedDefaultUser || currentUser?.id || users[0]?.id || '1';
      userCounts[fallbackId] = (userCounts[fallbackId] || 0) + 1;
      return fallbackId;
    };

    return { resolveUser, userCounts, unmatchedCodes };
  };

  const handleUpload = () => {
    try {
      const data = JSON.parse(jsonInput);
      const { resolveUser, userCounts, unmatchedCodes } = createResolver();
      let currentPortfolios: Portfolio[] = [];
      if (data.portfolios) {
        const taxIds = new Set();
        const duplicates = [];
        for (const p of data.portfolios) {
          if (p.taxId) {
            const cleanTax = String(p.taxId).toUpperCase().trim();
            if (taxIds.has(cleanTax)) {
              duplicates.push(cleanTax);
            }
            taxIds.add(cleanTax);
          }
        }
        if (duplicates.length > 0) {
          setError(`Errore: Il JSON contiene C.F./P.IVA duplicati: ${duplicates.join(', ')}`);
          setUploadSummary(null);
          return;
        }
        currentPortfolios = data.portfolios.map((p: any) => {
          const rawAddr = p.legalAddress || p.address || p.indirizzo || '';
          const spl = splitAddress(rawAddr);
          const street = p.street || p.via || spl?.street || '';
          const cap = String(p.cap || spl?.cap || '');
          const city = p.city || p.comune || spl?.city || '';
          const province = String(p.province || p.provincia || p.pr || p.prov || spl?.province || '').toUpperCase();
          const legalAddress = rawAddr || composeAddress(street, cap, city, province);

          const rawUser = p.userCode || p.codiceUtente || p.matricolaAgente || p.codiceAgente || p.agente || p.assegnatoA || p.assignedTo;
          return {
            ...p,
            id: p.id || Math.random().toString(36).substr(2, 9),
            taxId: p.taxId ? String(p.taxId).toUpperCase().trim() : '',
            legalAddress,
            street,
            cap,
            city,
            province,
            assignedTo: resolveUser(rawUser),
            createdAt: p.createdAt || getTimestamp(),
            updatedAt: p.updatedAt || getTimestamp()
          };
        });
        setPortfolios(currentPortfolios);
      }

      let currentPdps: PDP[] = [];
      if (data.pdps) {
        currentPdps = data.pdps.map((p: any) => {
          const svc = p.service ? (String(p.service).toUpperCase() === 'METANO' ? 'METANO' : 'POWER') : (p.remi || p.portata ? 'METANO' : 'POWER');
          const rawAddr = p.address || p.indirizzo || '';
          const spl = splitAddress(rawAddr);
          const street = p.street || p.via || spl?.street || '';
          const cap = String(p.cap || spl?.cap || '');
          const city = p.city || p.comune || spl?.city || '';
          const province = String(p.province || p.provincia || p.pr || p.prov || spl?.province || '').toUpperCase();
          const address = rawAddr || composeAddress(street, cap, city, province);

          const rawUser = p.userCode || p.codiceUtente || p.matricolaAgente || p.codiceAgente || p.agente || p.assegnatoA || p.assignedTo;
          return {
            ...p,
            id: p.id || Math.random().toString(36).substr(2, 9),
            service: svc,
            pdpCode: String(p.pdpCode || p.pdp || p.codice || '').toUpperCase().trim(),
            address,
            street,
            cap,
            city,
            province,
            matricola: p.matricola ? String(p.matricola).toUpperCase().trim() : '',
            cabinaPrimaria: svc === 'POWER' ? (p.cabinaPrimaria || p.technicalSpecs || p.cabina || '') : '',
            technicalSpecs: svc === 'POWER' ? (p.technicalSpecs || p.cabinaPrimaria || p.cabina || '') : '',
            potenzaDisponibile: svc === 'POWER' ? Number(p.potenzaDisponibile || p.potDisp || 0) : 0,
            potenzaImpegnata: svc === 'POWER' ? Number(p.potenzaImpegnata || p.potImp || 0) : 0,
            tensione: svc === 'POWER' ? formatTensione(p.tensione) : '',
            remi: svc === 'METANO' ? String(p.remi || '').toUpperCase().trim() : '',
            portata: svc === 'METANO' ? String(p.portata || '').toUpperCase().trim() : '',
            assignedTo: resolveUser(rawUser),
            createdAt: p.createdAt || getTimestamp(),
            updatedAt: getTimestamp()
          };
        });
      }

      let mappedContractsCount = 0;
      if (data.contracts) {
        mappedContractsCount = data.contracts.length;
        const mappedContracts = data.contracts.map((c: any) => {
          const contractCode = c.pdpCode || c.pdp || '';
          let matchedPdp = currentPdps.find(p => p.id === c.pdpId || (contractCode && p.pdpCode === String(contractCode).toUpperCase().trim()));
          
          const contractService = c.service ? (String(c.service).toUpperCase() === 'METANO' ? 'METANO' : 'POWER') : (c.remi || c.portata ? 'METANO' : 'POWER');
          
          // Collega portfolioId se taxId presente
          let portId = c.portfolioId;
          let matchedPort: Portfolio | undefined = undefined;
          if (!portId && (c.taxId || c.cf || c.partitaIva)) {
            const searchTax = String(c.taxId || c.cf || c.partitaIva).toUpperCase().trim();
            matchedPort = currentPortfolios.find(p => p.taxId.toUpperCase() === searchTax);
            if (matchedPort) portId = matchedPort.id;
          } else if (portId) {
            matchedPort = currentPortfolios.find(p => p.id === portId);
          }

          const rawUser = c.userCode || c.codiceUtente || c.matricolaAgente || c.codiceAgente || c.agente || c.assegnatoA || c.assignedTo;
          const contractAssignedId = resolveUser(rawUser, matchedPort?.assignedTo);

          const rawAddr = c.address || c.indirizzo || '';
          const spl = splitAddress(rawAddr);
          const street = c.street || c.via || spl?.street || matchedPdp?.street || '';
          const cap = String(c.cap || spl?.cap || matchedPdp?.cap || '');
          const city = c.city || c.comune || spl?.city || matchedPdp?.city || '';
          const province = String(c.province || c.provincia || c.pr || c.prov || spl?.province || matchedPdp?.province || '').toUpperCase();
          const address = rawAddr || composeAddress(street, cap, city, province) || matchedPdp?.address || '';

          if (!matchedPdp && contractCode) {
            matchedPdp = {
              id: Math.random().toString(36).substr(2, 9),
              pdpCode: String(contractCode).toUpperCase().trim(),
              service: contractService,
              address,
              street,
              cap,
              city,
              province,
              matricola: c.matricola ? String(c.matricola).toUpperCase().trim() : '',
              cabinaPrimaria: contractService === 'POWER' ? (c.cabinaPrimaria || c.technicalSpecs || c.cabina || '') : '',
              technicalSpecs: contractService === 'POWER' ? (c.technicalSpecs || c.cabinaPrimaria || c.cabina || '') : '',
              potenzaDisponibile: contractService === 'POWER' ? Number(c.potenzaDisponibile || c.potDisp || 0) : 0,
              potenzaImpegnata: contractService === 'POWER' ? Number(c.potenzaImpegnata || c.potImp || 0) : 0,
              tensione: contractService === 'POWER' ? formatTensione(c.tensione) : '',
              remi: contractService === 'METANO' ? String(c.remi || '').toUpperCase().trim() : '',
              portata: contractService === 'METANO' ? String(c.portata || '').toUpperCase().trim() : '',
              assignedTo: contractAssignedId,
              createdAt: getTimestamp(),
              updatedAt: getTimestamp()
            };
            currentPdps.push(matchedPdp);
          } else if (matchedPdp) {
            // Aggiorna le info PDP se specificate nel contratto
            if (c.matricola) matchedPdp.matricola = String(c.matricola).toUpperCase().trim();
            if (c.cabinaPrimaria) {
              matchedPdp.cabinaPrimaria = c.cabinaPrimaria;
              matchedPdp.technicalSpecs = c.cabinaPrimaria;
            }
            if (c.potenzaDisponibile !== undefined) matchedPdp.potenzaDisponibile = Number(c.potenzaDisponibile);
            if (c.potenzaImpegnata !== undefined) matchedPdp.potenzaImpegnata = Number(c.potenzaImpegnata);
            if (c.tensione) matchedPdp.tensione = formatTensione(c.tensione);
            if (c.remi) matchedPdp.remi = String(c.remi).toUpperCase().trim();
            if (c.portata) matchedPdp.portata = String(c.portata).toUpperCase().trim();
            if (address) {
              matchedPdp.address = address;
              matchedPdp.street = street;
              matchedPdp.cap = cap;
              matchedPdp.city = city;
              matchedPdp.province = province;
            }
            matchedPdp.updatedAt = getTimestamp();
          }

          return {
            ...c,
            id: c.id || Math.random().toString(36).substr(2, 9),
            portfolioId: portId || '',
            pdpId: matchedPdp ? matchedPdp.id : (c.pdpId || ''),
            service: contractService,
            address,
            street,
            cap,
            city,
            province,
            matricola: c.matricola || matchedPdp?.matricola || '',
            cabinaPrimaria: c.cabinaPrimaria || matchedPdp?.cabinaPrimaria || '',
            tensione: c.tensione ? formatTensione(c.tensione) : (matchedPdp?.tensione || ''),
            remi: c.remi || matchedPdp?.remi || '',
            portata: c.portata || matchedPdp?.portata || '',
            assignedTo: contractAssignedId,
            updatedAt: getTimestamp()
          };
        });
        setContracts(mappedContracts);
      }

      if (currentPdps.length > 0) {
        setPdps(currentPdps);
      }

      if (data.cabine) setCabine(data.cabine);
      if (data.casi) setCasi(data.casi);
      if (data.creditChecks) setCreditChecks(data.creditChecks);
      
      setUploadSummary({
        portfolios: currentPortfolios.length,
        pdps: currentPdps.length,
        contracts: mappedContractsCount,
        userCounts,
        unmatchedCodes: Array.from(unmatchedCodes)
      });
      setSuccess('Database aggiornato con successo da JSON!');
      setError('');
      setJsonInput('');
    } catch (e) {
      setError('Errore nel formato JSON. Controlla la sintassi.');
      setSuccess('');
      setUploadSummary(null);
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
        const { resolveUser, userCounts, unmatchedCodes } = createResolver();
        
        const portfoliosSheet = wb.Sheets['Portafoglio'] || wb.Sheets['portafoglio'] || wb.Sheets['Clienti'];
        const pdpsSheet = wb.Sheets['PDP'] || wb.Sheets['pdp'] || wb.Sheets['Punti di Prelievo'] || wb.Sheets['Punti Prelievo'];
        const contractsSheet = wb.Sheets['Contratti'] || wb.Sheets['contratti'] || wb.Sheets['Contracts'];

        let uploadedPortfolios: Portfolio[] = [];
        if (portfoliosSheet) {
          const data = XLSX.utils.sheet_to_json(portfoliosSheet) as any[];
          const taxIds = new Set();
          const duplicates = [];
          for (const p of data) {
            if (p.taxId) {
              const cleanTax = String(p.taxId).toUpperCase().trim();
              if (taxIds.has(cleanTax)) {
                duplicates.push(cleanTax);
              }
              taxIds.add(cleanTax);
            }
          }
          if (duplicates.length > 0) {
            setError(`Errore nell'Excel: C.F./P.IVA duplicati: ${duplicates.join(', ')}`);
            setUploadSummary(null);
            return;
          }
          uploadedPortfolios = data.map(p => {
            const rawAddr = p.legalAddress || p.address || p.indirizzo || '';
            const spl = splitAddress(rawAddr);
            const street = p.street || p.via || spl?.street || '';
            const cap = String(p.cap || spl?.cap || '');
            const city = p.city || p.comune || spl?.city || '';
            const province = String(p.province || p.provincia || p.pr || p.prov || spl?.province || '').toUpperCase();
            const legalAddress = rawAddr || composeAddress(street, cap, city, province);

            const rawUser = p.userCode || p.codiceUtente || p.matricolaAgente || p.codiceAgente || p.agente || p.assegnatoA || p.assignedTo;
            return {
              ...p,
              id: p.id || Math.random().toString(36).substr(2, 9),
              taxId: p.taxId ? String(p.taxId).toUpperCase().trim() : '',
              legalAddress,
              street,
              cap,
              city,
              province,
              assignedTo: resolveUser(rawUser),
              createdAt: p.createdAt || getTimestamp(),
              updatedAt: p.updatedAt || getTimestamp()
            };
          });
          setPortfolios(uploadedPortfolios);
        }

        let uploadedPdps: PDP[] = [];
        if (pdpsSheet) {
          const rawPdps = XLSX.utils.sheet_to_json(pdpsSheet) as any[];
          uploadedPdps = rawPdps.map((p) => {
            const svc = p.service ? (String(p.service).toUpperCase() === 'METANO' ? 'METANO' : 'POWER') : (p.remi || p.portata ? 'METANO' : 'POWER');
            const rawAddr = p.address || p.indirizzo || '';
            const spl = splitAddress(rawAddr);
            const street = p.street || p.via || spl?.street || '';
            const cap = String(p.cap || spl?.cap || '');
            const city = p.city || p.comune || spl?.city || '';
            const province = String(p.province || p.provincia || p.pr || p.prov || spl?.province || '').toUpperCase();
            const address = rawAddr || composeAddress(street, cap, city, province);

            const rawUser = p.userCode || p.codiceUtente || p.matricolaAgente || p.codiceAgente || p.agente || p.assegnatoA || p.assignedTo;
            return {
              id: p.id || Math.random().toString(36).substr(2, 9),
              pdpCode: String(p.pdpCode || p.pdp || p.codice || '').trim().toUpperCase(),
              service: svc,
              address,
              street,
              cap,
              city,
              province,
              matricola: p.matricola ? String(p.matricola).toUpperCase().trim() : '',
              cabinaPrimaria: svc === 'POWER' ? (p.cabinaPrimaria || p.cabina || p.technicalSpecs || '') : '',
              technicalSpecs: svc === 'POWER' ? (p.technicalSpecs || p.cabinaPrimaria || p.cabina || '') : '',
              potenzaDisponibile: svc === 'POWER' ? Number(p.potenzaDisponibile || p.potDisp || 0) : 0,
              potenzaImpegnata: svc === 'POWER' ? Number(p.potenzaImpegnata || p.potImp || 0) : 0,
              tensione: svc === 'POWER' ? formatTensione(p.tensione) : '',
              remi: svc === 'METANO' ? String(p.remi || '').toUpperCase().trim() : '',
              portata: svc === 'METANO' ? String(p.portata || '').toUpperCase().trim() : '',
              cabinaPrimariaId: p.cabinaPrimariaId,
              assignedTo: resolveUser(rawUser),
              createdAt: p.createdAt || getTimestamp(),
              updatedAt: p.updatedAt || getTimestamp()
            };
          });
        }

        let mappedContractsCount = 0;
        if (contractsSheet) {
          const rawContracts = XLSX.utils.sheet_to_json(contractsSheet) as any[];
          mappedContractsCount = rawContracts.length;
          const mappedContracts: Contract[] = rawContracts.map((c) => {
            const contractCode = c.pdpCode || c.pdp || c.codicePdp || c.pdpId || '';
            const contractService = c.service ? (String(c.service).toUpperCase() === 'METANO' ? 'METANO' : 'POWER') : (c.remi || c.portata ? 'METANO' : 'POWER');
            
            let matchedPdp = uploadedPdps.find(p => p.id === c.pdpId || (contractCode && p.pdpCode === String(contractCode).toUpperCase().trim()));

            // Associa portfolioId se taxId presente
            let portId = c.portfolioId;
            let matchedPort: Portfolio | undefined = undefined;
            if (!portId && (c.taxId || c.cf || c.partitaIva)) {
              const searchTax = String(c.taxId || c.cf || c.partitaIva).toUpperCase().trim();
              matchedPort = uploadedPortfolios.find(p => p.taxId.toUpperCase() === searchTax);
              if (matchedPort) portId = matchedPort.id;
            } else if (portId) {
              matchedPort = uploadedPortfolios.find(p => p.id === portId);
            }

            const rawUser = c.userCode || c.codiceUtente || c.matricolaAgente || c.codiceAgente || c.agente || c.assegnatoA || c.assignedTo;
            const contractAssignedId = resolveUser(rawUser, matchedPort?.assignedTo);

            const rawAddr = c.address || c.indirizzo || '';
            const spl = splitAddress(rawAddr);
            const street = c.street || c.via || spl?.street || matchedPdp?.street || '';
            const cap = String(c.cap || spl?.cap || matchedPdp?.cap || '');
            const city = c.city || c.comune || spl?.city || matchedPdp?.city || '';
            const province = String(c.province || c.provincia || c.pr || c.prov || spl?.province || matchedPdp?.province || '').toUpperCase();
            const address = rawAddr || composeAddress(street, cap, city, province) || matchedPdp?.address || '';

            if (!matchedPdp && contractCode && String(contractCode).length === 14) {
              matchedPdp = {
                id: Math.random().toString(36).substr(2, 9),
                pdpCode: String(contractCode).toUpperCase().trim(),
                service: contractService,
                address,
                street,
                cap,
                city,
                province,
                matricola: c.matricola ? String(c.matricola).toUpperCase().trim() : '',
                cabinaPrimaria: contractService === 'POWER' ? (c.cabinaPrimaria || c.cabina || c.technicalSpecs || '') : '',
                technicalSpecs: contractService === 'POWER' ? (c.technicalSpecs || c.cabinaPrimaria || c.cabina || '') : '',
                potenzaDisponibile: contractService === 'POWER' ? Number(c.potenzaDisponibile || c.potDisp || 0) : 0,
                potenzaImpegnata: contractService === 'POWER' ? Number(c.potenzaImpegnata || c.potImp || 0) : 0,
                tensione: contractService === 'POWER' ? formatTensione(c.tensione) : '',
                remi: contractService === 'METANO' ? String(c.remi || '').toUpperCase().trim() : '',
                portata: contractService === 'METANO' ? String(c.portata || '').toUpperCase().trim() : '',
                assignedTo: contractAssignedId,
                createdAt: getTimestamp(),
                updatedAt: getTimestamp()
              };
              uploadedPdps.push(matchedPdp);
            } else if (matchedPdp) {
              // Aggiorna PDP con i dati del contratto
              if (c.matricola) matchedPdp.matricola = String(c.matricola).toUpperCase().trim();
              if (c.cabinaPrimaria || c.cabina) {
                matchedPdp.cabinaPrimaria = c.cabinaPrimaria || c.cabina;
                matchedPdp.technicalSpecs = c.cabinaPrimaria || c.cabina;
              }
              if (c.potenzaDisponibile !== undefined || c.potDisp !== undefined) matchedPdp.potenzaDisponibile = Number(c.potenzaDisponibile || c.potDisp || 0);
              if (c.potenzaImpegnata !== undefined || c.potImp !== undefined) matchedPdp.potenzaImpegnata = Number(c.potenzaImpegnata || c.potImp || 0);
              if (c.tensione) matchedPdp.tensione = formatTensione(c.tensione);
              if (c.remi) matchedPdp.remi = String(c.remi).toUpperCase().trim();
              if (c.portata) matchedPdp.portata = String(c.portata).toUpperCase().trim();
              if (address) {
                matchedPdp.address = address;
                matchedPdp.street = street;
                matchedPdp.cap = cap;
                matchedPdp.city = city;
                matchedPdp.province = province;
              }
              matchedPdp.updatedAt = getTimestamp();
            }

            return {
              id: c.id || Math.random().toString(36).substr(2, 9),
              contractNumber: c.contractNumber || c.numero || `CTR-${Math.floor(Math.random()*10000)}`,
              portfolioId: portId || '',
              pdpId: matchedPdp ? matchedPdp.id : (c.pdpId || ''),
              fornitore: c.fornitore || 'A2A1',
              contractType: c.contractType || 'SWITCH',
              status: c.status || c.stato || 'attivo',
              service: contractService,
              volume: Number(c.volume || 0),
              startDate: c.startDate || new Date().toISOString().split('T')[0],
              durationMonths: Number(c.durationMonths || 12),
              endDate: c.endDate || '',
              tariffa: c.tariffa || '',
              tariffOption: c.tariffOption || '',
              address,
              street,
              cap,
              city,
              province,
              matricola: c.matricola || matchedPdp?.matricola || '',
              cabinaPrimaria: c.cabinaPrimaria || matchedPdp?.cabinaPrimaria || '',
              tensione: c.tensione ? formatTensione(c.tensione) : (matchedPdp?.tensione || ''),
              remi: c.remi || matchedPdp?.remi || '',
              portata: c.portata || matchedPdp?.portata || '',
              potenzaDisponibile: Number(c.potenzaDisponibile || matchedPdp?.potenzaDisponibile || 0),
              potenzaImpegnata: Number(c.potenzaImpegnata || matchedPdp?.potenzaImpegnata || 0),
              notes: c.notes || '',
              assignedTo: contractAssignedId,
              attachments: [],
              creationDate: c.creationDate || new Date().toISOString().split('T')[0],
              createdAt: c.createdAt || getTimestamp(),
              updatedAt: getTimestamp()
            };
          });
          setContracts(mappedContracts);
        }

        if (uploadedPdps.length > 0) {
          setPdps(uploadedPdps);
        }

        setUploadSummary({
          portfolios: uploadedPortfolios.length,
          pdps: uploadedPdps.length,
          contracts: mappedContractsCount,
          userCounts,
          unmatchedCodes: Array.from(unmatchedCodes)
        });
        setSuccess('Database aggiornato con successo da Excel!');
        setError('');
        if (fileInputRef.current) fileInputRef.current.value = '';
      } catch (err) {
        setError('Errore durante la lettura del file Excel.');
        setSuccess('');
        setUploadSummary(null);
      }
    };
    reader.readAsBinaryString(file);
  };

  const downloadTemplate = () => {
    const template = {
      portfolios: [
        {
          businessName: "Esempio Azienda",
          taxId: "01234567890",
          legalAddress: "Via Roma 1, 20100 Milano MI",
          via: "Via Roma 1",
          cap: "20100",
          comune: "Milano",
          prov: "MI",
          email: "info@esempio.com",
          phone: "021234567",
          userCode: "AG01"
        }
      ],
      pdps: [
        {
          pdpCode: "IT001E12345678",
          service: "POWER",
          address: "Via Milano 10, 00100 Roma RM",
          via: "Via Milano 10",
          cap: "00100",
          comune: "Roma",
          prov: "RM",
          matricola: "MAT123456",
          cabinaPrimaria: "CP ROMA CENTRO",
          potenzaDisponibile: 11,
          potenzaImpegnata: 10,
          tensione: "380",
          userCode: "AG01"
        },
        {
          pdpCode: "00881234567890",
          service: "METANO",
          address: "Corso Italia 5, 20121 Milano MI",
          via: "Corso Italia 5",
          cap: "20121",
          comune: "Milano",
          prov: "MI",
          matricola: "GAS987654",
          remi: "34880001",
          portata: "G6",
          userCode: "AG01"
        }
      ],
      contracts: [
        {
          contractNumber: "CTR-2024-001",
          taxId: "01234567890",
          pdpCode: "IT001E12345678",
          service: "POWER",
          fornitore: "A2A1",
          contractType: "SWITCH",
          status: "attivo",
          address: "Via Milano 10, 00100 Roma RM",
          via: "Via Milano 10",
          cap: "00100",
          comune: "Roma",
          prov: "RM",
          matricola: "MAT123456",
          cabinaPrimaria: "CP ROMA CENTRO",
          potenzaDisponibile: 11,
          potenzaImpegnata: 10,
          tensione: "380",
          volume: 5000,
          startDate: "2024-01-01",
          durationMonths: 12,
          tariffa: "Fissa 0.12",
          tariffOption: "BTA",
          userCode: "AG01"
        },
        {
          contractNumber: "CTR-2024-002",
          taxId: "01234567890",
          pdpCode: "00881234567890",
          service: "METANO",
          fornitore: "A2A1",
          contractType: "SWITCH",
          status: "attivo",
          address: "Corso Italia 5, 20121 Milano MI",
          via: "Corso Italia 5",
          cap: "20121",
          comune: "Milano",
          prov: "MI",
          matricola: "GAS987654",
          remi: "34880001",
          portata: "G6",
          volume: 1500,
          startDate: "2024-01-01",
          durationMonths: 12,
          tariffa: "Fissa 0.45",
          userCode: "AG01"
        }
      ],
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
      { businessName: 'Esempio Azienda', taxId: '01234567890', legalAddress: 'Via Roma 1, 20100 Milano MI', via: 'Via Roma 1', cap: '20100', comune: 'Milano', prov: 'MI', email: 'info@esempio.com', phone: '021234567', userCode: 'AG01' }
    ], { header: ['businessName', 'taxId', 'legalAddress', 'via', 'cap', 'comune', 'prov', 'email', 'phone', 'userCode'] });
    XLSX.utils.book_append_sheet(wb, portfoliosWS, 'Portafoglio');

    const pdpsWS = XLSX.utils.json_to_sheet([
      { 
        pdpCode: 'IT001E12345678', 
        service: 'POWER', 
        address: 'Via Milano 10, 00100 Roma RM', 
        via: 'Via Milano 10',
        cap: '00100',
        comune: 'Roma',
        prov: 'RM',
        matricola: 'MAT123456', 
        cabinaPrimaria: 'CP ROMA CENTRO', 
        potenzaDisponibile: 11, 
        potenzaImpegnata: 10, 
        tensione: '380', 
        remi: '', 
        portata: '',
        userCode: 'AG01'
      },
      { 
        pdpCode: '00881234567890', 
        service: 'METANO', 
        address: 'Corso Italia 5, 20121 Milano MI', 
        via: 'Corso Italia 5',
        cap: '20121',
        comune: 'Milano',
        prov: 'MI',
        matricola: 'GAS987654', 
        cabinaPrimaria: '', 
        potenzaDisponibile: 0, 
        potenzaImpegnata: 0, 
        tensione: '', 
        remi: '34880001', 
        portata: 'G6',
        userCode: 'AG01'
      }
    ], { header: ['pdpCode', 'service', 'address', 'via', 'cap', 'comune', 'prov', 'matricola', 'cabinaPrimaria', 'potenzaDisponibile', 'potenzaImpegnata', 'tensione', 'remi', 'portata', 'userCode'] });
    XLSX.utils.book_append_sheet(wb, pdpsWS, 'PDP');

    const contractsWS = XLSX.utils.json_to_sheet([
      { 
        contractNumber: 'CTR-2024-001', 
        taxId: '01234567890', 
        pdpCode: 'IT001E12345678', 
        service: 'POWER', 
        fornitore: 'A2A1', 
        contractType: 'SWITCH', 
        status: 'attivo', 
        address: 'Via Milano 10, 00100 Roma RM', 
        via: 'Via Milano 10',
        cap: '00100',
        comune: 'Roma',
        prov: 'RM',
        matricola: 'MAT123456', 
        cabinaPrimaria: 'CP ROMA CENTRO', 
        potenzaDisponibile: 11, 
        potenzaImpegnata: 10, 
        tensione: '380', 
        remi: '', 
        portata: '', 
        volume: 5000, 
        startDate: '2024-01-01', 
        durationMonths: 12, 
        tariffa: 'Fissa 0.12', 
        tariffOption: 'BTA',
        userCode: 'AG01'
      },
      { 
        contractNumber: 'CTR-2024-002', 
        taxId: '01234567890', 
        pdpCode: '00881234567890', 
        service: 'METANO', 
        fornitore: 'A2A1', 
        contractType: 'SWITCH', 
        status: 'attivo', 
        address: 'Corso Italia 5, 20121 Milano MI', 
        via: 'Corso Italia 5',
        cap: '20121',
        comune: 'Milano',
        prov: 'MI',
        matricola: 'GAS987654', 
        cabinaPrimaria: '', 
        potenzaDisponibile: 0, 
        potenzaImpegnata: 0, 
        tensione: '', 
        remi: '34880001', 
        portata: 'G6', 
        volume: 1500, 
        startDate: '2024-01-01', 
        durationMonths: 12, 
        tariffa: 'Fissa 0.45', 
        tariffOption: '',
        userCode: 'AG01'
      }
    ], { header: ['contractNumber', 'taxId', 'pdpCode', 'service', 'fornitore', 'contractType', 'status', 'address', 'via', 'cap', 'comune', 'prov', 'matricola', 'cabinaPrimaria', 'potenzaDisponibile', 'potenzaImpegnata', 'tensione', 'remi', 'portata', 'volume', 'startDate', 'durationMonths', 'tariffa', 'tariffOption', 'userCode'] });
    XLSX.utils.book_append_sheet(wb, contractsWS, 'Contratti');

    XLSX.writeFile(wb, 'template_charlie.xlsx');
  };

  return (
    <div className="space-y-8 animate-in fade-in duration-500">
      <header>
        <h2 className="text-4xl font-black tracking-tighter text-emerald-900">Caricamento Massivo</h2>
        <p className="text-slate-500 font-medium italic">Aggiorna l'intero database tramite file Excel o JSON con abbinamento profili</p>
      </header>

      {/* Sezione Configurazione Abbinamento Agenti */}
      <div className="bg-gradient-to-r from-emerald-950 via-slate-900 to-emerald-900 text-white p-8 rounded-[2.5rem] shadow-xl space-y-6">
        <div className="flex flex-col md:flex-row justify-between md:items-center gap-4">
          <div>
            <div className="flex items-center gap-2">
              <span className="bg-emerald-500/20 text-emerald-300 border border-emerald-400/30 text-[10px] font-black uppercase px-2.5 py-0.5 rounded-full">
                Abbinamento Agenti & Profili
              </span>
            </div>
            <h3 className="text-2xl font-black tracking-tight mt-1">Gestione Assegnazione Utente (Matricola)</h3>
            <p className="text-xs text-slate-300 mt-0.5">
              Configura come abbinare le anagrafiche, i PDP e i contratti ai rispettivi profili operatore
            </p>
          </div>
          
          <div className="flex items-center gap-2 bg-white/10 backdrop-blur p-1.5 rounded-2xl border border-white/10 shrink-0">
            <button
              type="button"
              onClick={() => setAssignmentMode('auto_with_default')}
              className={`px-4 py-2 rounded-xl text-xs font-black transition-all ${
                assignmentMode === 'auto_with_default' 
                  ? 'bg-emerald-500 text-white shadow-lg' 
                  : 'text-slate-300 hover:text-white'
              }`}
            >
              Riconosci da Matricola nel File
            </button>
            <button
              type="button"
              onClick={() => setAssignmentMode('force_user')}
              className={`px-4 py-2 rounded-xl text-xs font-black transition-all ${
                assignmentMode === 'force_user' 
                  ? 'bg-emerald-500 text-white shadow-lg' 
                  : 'text-slate-300 hover:text-white'
              }`}
            >
              Forza Singolo Agente
            </button>
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 pt-4 border-t border-white/10 text-xs">
          {assignmentMode === 'auto_with_default' ? (
            <div className="space-y-3 bg-white/5 backdrop-blur p-5 rounded-2xl border border-white/10">
              <p className="font-bold text-emerald-300 flex items-center gap-1.5">
                <CheckCircle2 size={16} /> Lettura automatica della colonna <code className="bg-black/40 px-1.5 py-0.5 rounded font-mono text-white text-[11px]">userCode</code> o <code className="bg-black/40 px-1.5 py-0.5 rounded font-mono text-white text-[11px]">codiceUtente</code>
              </p>
              <p className="text-slate-300 text-[11px] leading-relaxed">
                Il sistema riconosce la <strong>matricola dell'agente</strong> (es. <code className="text-white font-mono font-bold bg-white/10 px-1.5 py-0.5 rounded">AG01</code>) inserita nel foglio Excel e abbina ogni riga al relativo profilo.
                Se nei Contratti o PDP la colonna è vuota, il sistema eredita automaticamente l'agente del Cliente (Portafoglio).
              </p>
              <div className="pt-2">
                <label className="text-[10px] font-black uppercase tracking-wider text-slate-300 block mb-1">
                  Se una riga non ha matricola e non ha cliente collegato, assegna a:
                </label>
                <select 
                  className="w-full bg-slate-800 text-white p-3 rounded-xl font-bold outline-none border border-slate-700 focus:border-emerald-500 cursor-pointer"
                  value={selectedDefaultUser}
                  onChange={e => setSelectedDefaultUser(e.target.value)}
                >
                  {users.map(u => (
                    <option key={u.id} value={u.id}>
                      {u.userCode} - {u.name} ({u.role})
                    </option>
                  ))}
                </select>
              </div>
            </div>
          ) : (
            <div className="space-y-3 bg-white/5 backdrop-blur p-5 rounded-2xl border border-white/10">
              <p className="font-bold text-yellow-300 flex items-center gap-1.5">
                <Info size={16} /> Assegnazione forzata di tutto il file
              </p>
              <p className="text-slate-300 text-[11px] leading-relaxed">
                Ignora eventuali codici nel file e assegna <strong>tutti i record caricati</strong> direttamente al collaboratore scelto qui sotto. Ideale quando si carica un file Excel intestato a un solo agente senza dover modificare le colonne del foglio.
              </p>
              <div className="pt-2">
                <label className="text-[10px] font-black uppercase tracking-wider text-slate-300 block mb-1">
                  Assegna tutti i record importati a:
                </label>
                <select 
                  className="w-full bg-slate-800 text-white p-3 rounded-xl font-bold outline-none border border-slate-700 focus:border-emerald-500 cursor-pointer"
                  value={selectedDefaultUser}
                  onChange={e => setSelectedDefaultUser(e.target.value)}
                >
                  {users.map(u => (
                    <option key={u.id} value={u.id}>
                      {u.userCode} - {u.name} ({u.role})
                    </option>
                  ))}
                </select>
              </div>
            </div>
          )}

          {/* Badges Matricole Registrate */}
          <div className="bg-white/5 backdrop-blur p-5 rounded-2xl border border-white/10 space-y-3 flex flex-col justify-between">
            <div>
              <span className="text-[10px] font-black uppercase tracking-widest text-slate-400 block mb-2">
                Matricole Agenti Registrate nel CRM ({users.length})
              </span>
              <div className="flex flex-wrap gap-2 max-h-36 overflow-y-auto">
                {users.map(u => (
                  <div 
                    key={u.id}
                    className={`border px-2.5 py-1.5 rounded-xl flex items-center gap-2 transition-all cursor-pointer ${
                      selectedDefaultUser === u.id 
                        ? 'bg-emerald-500/30 border-emerald-400 text-white ring-1 ring-emerald-400' 
                        : 'bg-white/10 hover:bg-white/20 border-white/10 text-slate-200'
                    }`}
                    title={`Clicca per impostare come predefinito: ${u.name}`}
                    onClick={() => setSelectedDefaultUser(u.id)}
                  >
                    <span className="font-mono font-black text-emerald-400 text-xs">{u.userCode}</span>
                    <span className="font-bold text-[11px] truncate max-w-[120px]">{u.name}</span>
                    <span className="text-[8px] uppercase px-1 rounded bg-black/40 text-slate-400">{u.role}</span>
                  </div>
                ))}
              </div>
            </div>
            <p className="text-[10px] text-slate-400 italic pt-2 border-t border-white/10">
              💡 Clicca su un agente per selezionarlo come profilo predefinito. Nel file Excel inserisci il codice matricola (es. <span className="text-emerald-400 font-bold font-mono">{users[0]?.userCode || 'AG01'}</span>) nella colonna <strong>userCode</strong>.
            </p>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        <div className="bg-white p-8 rounded-[2.5rem] border border-slate-100 shadow-sm space-y-6">
          <div className="flex items-center justify-between">
            <h3 className="font-black text-xl text-slate-800 tracking-tight">Caricamento Excel</h3>
            <button 
              onClick={downloadExcelTemplate}
              className="flex items-center gap-2 text-emerald-600 hover:text-emerald-700 font-bold text-sm bg-emerald-50 hover:bg-emerald-100 px-4 py-2 rounded-xl transition-colors"
            >
              <Download size={18} />
              Scarica Template Excel
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
          <p className="text-[10px] text-slate-400 uppercase font-bold tracking-wider text-center">
            Supporta fogli: Portafoglio, PDP, Contratti • Colonne Indirizzo: via, cap, comune, prov o campo unico • Matricola: userCode
          </p>
        </div>

        <div className="bg-white p-8 rounded-[2.5rem] border border-slate-100 shadow-sm space-y-6">
          <div className="flex items-center justify-between">
            <h3 className="font-black text-xl text-slate-800 tracking-tight">Incolla JSON</h3>
            <button 
              onClick={downloadTemplate}
              className="flex items-center gap-2 text-emerald-600 hover:text-emerald-700 font-bold text-sm bg-emerald-50 hover:bg-emerald-100 px-4 py-2 rounded-xl transition-colors"
            >
              <Download size={18} />
              Scarica Template JSON
            </button>
          </div>

          <textarea
            value={jsonInput}
            onChange={(e) => setJsonInput(e.target.value)}
            placeholder='{ "portfolios": [...], "pdps": [...], "contracts": [...] }'
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

      {error && (
        <div className="p-6 rounded-[2rem] border flex items-center gap-4 animate-in slide-in-from-bottom-4 duration-300 bg-red-50 border-red-100 text-red-600">
          <AlertCircle size={24} className="shrink-0" />
          <p className="font-bold">{error}</p>
        </div>
      )}

      {success && (
        <div className="p-6 rounded-[2rem] border space-y-4 animate-in slide-in-from-bottom-4 duration-300 bg-emerald-50 border-emerald-200 text-emerald-900">
          <div className="flex items-center gap-3">
            <CheckCircle2 size={24} className="text-emerald-600 shrink-0" />
            <p className="font-black text-base">{success}</p>
          </div>
          
          {uploadSummary && (
            <div className="bg-white/80 backdrop-blur rounded-2xl p-4 border border-emerald-100 space-y-3 text-xs">
              <div className="flex flex-wrap gap-6 font-bold text-slate-700">
                <span>Anagrafiche Clienti: <strong className="text-emerald-700 text-sm ml-1">{uploadSummary.portfolios}</strong></span>
                <span>Punti di Prelievo (PDP): <strong className="text-emerald-700 text-sm ml-1">{uploadSummary.pdps}</strong></span>
                <span>Contratti Caricati: <strong className="text-emerald-700 text-sm ml-1">{uploadSummary.contracts}</strong></span>
              </div>
              
              <div className="pt-3 border-t border-emerald-100">
                <p className="font-black uppercase tracking-wider text-[10px] text-slate-400 mb-2">
                  Riepilogo Abbinamenti Effettuati per Profilo Utente:
                </p>
                <div className="flex flex-wrap gap-2">
                  {Object.entries(uploadSummary.userCounts).map(([uId, count]) => {
                    const u = users.find(usr => usr.id === uId);
                    return (
                      <span key={uId} className="inline-flex items-center gap-2 bg-emerald-100/80 text-emerald-900 px-3 py-1.5 rounded-xl font-bold shadow-sm">
                        <span className="font-mono font-black text-white bg-emerald-700 px-1.5 py-0.5 rounded text-[10px]">
                          {u?.userCode || 'N/D'}
                        </span>
                        <span>{u?.name || 'Utente ' + uId}:</span>
                        <span className="text-emerald-800 font-black">{count} record</span>
                      </span>
                    );
                  })}
                </div>
              </div>

              {uploadSummary.unmatchedCodes.length > 0 && (
                <div className="pt-3 border-t border-amber-200 text-amber-800 flex items-start gap-2 bg-amber-50/50 p-3 rounded-xl">
                  <AlertTriangle size={18} className="text-amber-600 shrink-0 mt-0.5" />
                  <p className="text-[11px] leading-relaxed">
                    <strong>Avviso codici non riconosciuti:</strong> Nel file erano presenti matricole non registrate nel sistema (<code className="font-mono font-bold">{uploadSummary.unmatchedCodes.join(', ')}</code>). I relativi record sono stati associati al profilo predefinito. Puoi creare questi profili nella sezione <strong>Utenti</strong> per i prossimi caricamenti.
                  </p>
                </div>
              )}
            </div>
          )}
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
    const cleanUser = username.trim().toLowerCase();
    const cleanPass = password.trim();
    const user = users.find(u => 
      (u.username.trim().toLowerCase() === cleanUser || (u.userCode && u.userCode.trim().toLowerCase() === cleanUser)) && 
      u.password === cleanPass
    );
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
