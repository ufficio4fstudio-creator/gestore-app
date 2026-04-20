
export type UserRole = 'admin' | 'backoffice' | 'utente';

export interface User {
  id: string;
  username: string;
  password?: string;
  name: string;
  role: UserRole;
  userCode: string; // Codice Agente/Operatore
}

export interface Attachment {
  id: string;
  fileName: string;
  fileSize: number;
  uploadDate: string;
  mimeType: string;
}

export interface Timestamps {
  createdAt: string;
  updatedAt: string;
}

export type PortfolioEntityType = 'Domestico' | 'Impresa' | 'Associazione' | 'Condominio' | 'PA';

export interface Portfolio extends Timestamps {
  id: string;
  taxId: string;
  legalAddress: string;
  street: string;
  cap: string;
  city: string;
  province: string;
  phone: string;
  email: string;
  businessName: string;
  entityType: PortfolioEntityType;
  assignedTo: string; // User ID
  attachments: Attachment[];
}

export interface CabinaPrimaria {
  id: string;
  name: string;
}

export interface PDP extends Timestamps {
  id: string;
  pdpCode: string;
  address: string;
  street: string;
  cap: string;
  city: string;
  province: string;
  technicalSpecs: string;
  potenzaImpegnata: number;
  potenzaDisponibile: number;
  cabinaPrimariaId?: string;
  assignedTo: string;
}

export type ContractStatus = 'bozza' | 'da firmare' | 'trasmesso' | 'attivo' | 'non conforme' | 'KO';
export type Fornitore = 'A2A1' | 'A2A2' | 'AXPO' | 'AXP2' | 'DOLO' | 'DUFE' | 'OPEN' | 'SORG';
export type ContractType = 'SWITCH' | 'ATTIVAZIONE' | 'VOLT TIT III' | 'VOLT TIT IV' | 'ALLACCIO';

export interface Contract extends Timestamps {
  id: string;
  portfolioId: string;
  pdpId: string;
  contractNumber: string;
  fornitore: Fornitore;
  contractType: ContractType;
  tariffa: string;
  startDate: string;
  status: ContractStatus;
  notes: string;
  service: 'POWER' | 'METANO';
  tariffOption?: string; // DOM2, DOM3, BTA, MTA, BTIP, BTVE, MTIP, MTVE
  volume: number; // KWh o smc
  potenzaImpegnata: number;
  potenzaDisponibile: number;
  assignedTo: string;
  attachments: Attachment[];
  creationDate: string; // YYYY-MM-DD
  durationMonths: number;
  endDate: string; // YYYY-MM-DD
}

export type CaseCategory = 'ANALISI' | 'ANAGRAFICA' | 'CONTENZIOSO' | 'CONTRATTUALE' | 'CREDITI' | 'DISTRIBUZIONE' | 'FATTURAZIONE';
export type CaseStatus = 'nuovo' | 'in lavorazione' | 'risolto' | 'KO' | 'partner';

export interface Caso extends Timestamps {
  id: string;
  title: string;
  category: CaseCategory;
  portfolioId: string;
  pdpId?: string;
  description: string;
  notes?: string;
  status: CaseStatus;
  attachments: Attachment[];
  assignedTo: string;
}

export type ScoringStatus = 'richiesto' | 'inserito' | 'OK' | 'SDD' | 'KO' | 'cauzione' | 'da affidare' | 'non affidabile' | '';

export interface ScoringRequest extends Timestamps {
  id: string;
  portfolioId: string;
  powerVolume: number;
  methaneVolume: number;
  suppliers: {
    A2A1?: ScoringStatus;
    AXPO?: ScoringStatus;
    DOLO?: ScoringStatus;
    SORG?: ScoringStatus;
    OPEN?: ScoringStatus;
  };
  assignedTo: string; // User ID
}

export interface Message {
  id: string;
  targetId: string; // ID del contratto o del caso
  targetType: 'contract' | 'case';
  senderId: string;
  senderName: string;
  senderRole: UserRole;
  text: string;
  timestamp: string;
  isRead: boolean;
  attachments?: Attachment[];
}

export type AppView = 'portfolio' | 'pdp' | 'contracts' | 'dashboard' | 'cabine' | 'casi' | 'users' | 'statistics' | 'map' | 'bulk-upload' | 'scoring';
