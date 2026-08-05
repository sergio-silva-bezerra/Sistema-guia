import React, { useState, useMemo, useEffect, useRef } from 'react';
import { firestoreService } from '../lib/firestoreService';
import { useAuth } from '../lib/FirebaseProvider';
import { setTreLocationsForCoordinator, clearTreLocationsCache, getAllTreLocations } from '../lib/treDataService';
import { eleitoralStorage } from '../lib/eleitoralStorage';
import { SupabaseConfigModal } from './SupabaseConfigModal';
import { supabaseService } from '../lib/supabaseService';
import { isSupabaseConfigured } from '../lib/supabase';
import { motion, AnimatePresence } from 'motion/react';
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
  Legend
} from 'recharts';
import { 
  LayoutDashboard, 
  Building2, 
  Map as MapIcon, 
  Hash, 
  TrendingUp, 
  Filter, 
  RefreshCw, 
  ArrowUpDown, 
  Award, 
  MapPin,
  FileSpreadsheet,
  UploadCloud,
  Download,
  Trash2,
  Database,
  CheckCircle2,
  AlertCircle,
  Target,
  Users,
  Percent,
  Compass,
  Briefcase,
  Search,
  AlertTriangle,
  X,
  ChevronLeft,
  ChevronRight,
  ChevronsLeft,
  ChevronsRight,
  HelpCircle,
  Info,
  ExternalLink
} from 'lucide-react';
import { VotingLocation, TSE_COLUMNS, TseColumnDef } from '../data/eleitoralData';
import * as XLSX from 'xlsx';
import { parseExcelOrCSVBuffer } from '../lib/excelParser';

// Constants for theme colors (Navy & Royal Blue)
const COLORS = [
  '#0578d3', // Brand Royal Blue (Primary)
  '#0b122f', // Dark Navy Blue
  '#0284c7', // Sky Royal Blue
  '#0f172a', // Slate Dark Navy
  '#3b82f6', // Bright Blue
  '#2563eb', // Indigo Blue
  '#0284c7', // Dark Sky
  '#1d4ed8'  // Royal Blue
];

// Header cell component with hover speech bubble tooltip (balão explicativo)
function TseHeaderCell({ 
  variable, 
  description, 
  align = 'left',
  isSortable = false,
  onSort
}: { 
  key?: string;
  variable: string; 
  description: string; 
  align?: 'left' | 'center' | 'right';
  isSortable?: boolean;
  onSort?: () => void;
}) {
  const [hovered, setHovered] = useState(false);

  return (
    <th 
      className={`py-3.5 px-3 relative font-black text-[9px] uppercase tracking-wider text-zinc-700 dark:text-zinc-200 border-b border-zinc-200 dark:border-zinc-800 bg-zinc-100 dark:bg-zinc-850 select-none whitespace-nowrap transition-colors ${
        isSortable ? 'cursor-pointer hover:bg-zinc-200 dark:hover:bg-zinc-800' : ''
      } ${align === 'right' ? 'text-right' : align === 'center' ? 'text-center' : 'text-left'}`}
      onClick={() => isSortable && onSort && onSort()}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
    >
      <div className={`inline-flex items-center gap-1.5 ${align === 'right' ? 'justify-end' : align === 'center' ? 'justify-center' : 'justify-start'}`}>
        <span className="text-zinc-900 dark:text-zinc-100 font-extrabold">{variable}</span>
        <HelpCircle className="w-3 h-3 text-blue-500 opacity-70 hover:opacity-100 shrink-0 cursor-help" />
        {isSortable && <ArrowUpDown className="w-3 h-3 text-zinc-400 shrink-0" />}
      </div>

      {/* BALÃO EXPLICATIVO (HOVER TOOLTIP) */}
      <AnimatePresence>
        {hovered && (
          <motion.div 
            initial={{ opacity: 0, y: 6, scale: 0.95 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, y: 4, scale: 0.95 }}
            transition={{ duration: 0.12 }}
            className={`absolute bottom-full mb-2 ${align === 'right' ? 'right-0' : 'left-0'} w-72 p-3 bg-zinc-950 text-white rounded-md shadow-2xl border border-blue-500/50 z-50 pointer-events-none normal-case font-normal text-left text-xs leading-relaxed`}
          >
            <div className="flex items-center gap-1.5 font-bold text-blue-400 uppercase text-[10px] tracking-wider mb-1 border-b border-zinc-800 pb-1">
              <Info className="w-3.5 h-3.5 text-blue-400 shrink-0" />
              <span>Variável TSE: {variable}</span>
            </div>
            <p className="text-zinc-200 text-[11px] whitespace-pre-line leading-snug">
              {description}
            </p>
            <div className={`absolute top-full ${align === 'right' ? 'right-4' : 'left-4'} border-4 border-transparent border-t-zinc-950`}></div>
          </motion.div>
        )}
      </AnimatePresence>
    </th>
  );
}

// Empty default array - data must be loaded by Coordenador Geral
const SAMPLE_TRE_DATA: VotingLocation[] = [];

export default function EleitoralDashboard({ 
  isCoordinator = false, 
  canEditTreData = false,
  campaignVoters = [],
  coordinatorId: propCoordinatorId
}: { 
  isCoordinator?: boolean; 
  canEditTreData?: boolean;
  campaignVoters?: any[]; 
  coordinatorId?: string;
}) {
  const { user, coordinatorId: authCoordId } = useAuth();
  const activeCoordId = propCoordinatorId || authCoordId || user?.uid || 'default';

  const [subTab, setSubTab] = useState<'tre_oficial' | 'cruzamento'>('tre_oficial');

  // Load data from memory cache or IndexedDB
  const [votingLocations, setVotingLocations] = useState<VotingLocation[]>(() => {
    if (activeCoordId) {
      const cached = getAllTreLocations(activeCoordId);
      if (cached && cached.length > 0) {
        return cached as any;
      }
      try {
        const key = `sistema_urna360_eleitoral_data_${activeCoordId}`;
        const saved = localStorage.getItem(key);
        if (saved) {
          const parsed = JSON.parse(saved);
          if (Array.isArray(parsed) && parsed.length > 0) {
            setTreLocationsForCoordinator(activeCoordId, parsed);
            return parsed;
          }
        }
      } catch (e) {}
    }
    return [];
  });

  // Dynamic list of municipalities extracted from loaded TRE data or campaign voters
  const MUNICIPALITIES = useMemo(() => {
    const setMuni = new Set<string>();
    votingLocations.forEach(loc => {
      if (loc.municipio) setMuni.add(loc.municipio.trim());
    });
    campaignVoters.forEach((v: any) => {
      if (v.municipio) setMuni.add(v.municipio.trim());
      if (v.mappedMunicipio) setMuni.add(v.mappedMunicipio.trim());
    });
    const DEFAULT_RR_MUNICIPALITIES = [
      "Alto Alegre", "Amajari", "Boa Vista", "Bonfim", "Cantá", 
      "Caracaraí", "Caroebe", "Iracema", "Mucajaí", "Normandia", 
      "Pacaraima", "Rorainópolis", "São João da Baliza", "São Luiz", "Uiramutã"
    ];
    DEFAULT_RR_MUNICIPALITIES.forEach(m => setMuni.add(m));
    return Array.from(setMuni).sort((a, b) => a.localeCompare(b, 'pt-BR'));
  }, [votingLocations, campaignVoters]);

  const isSavingRef = React.useRef(false);

  // Helper to sanitize and fix swapped/misplaced fields in VotingLocation records
  const sanitizeLocation = (item: VotingLocation): VotingLocation => {
    if (!item) return item;
    let nmMuni = String(item.nmMunicipio || item.municipio || '').trim();
    let nmLoc = String(item.nmLocalVotacao || item.local || '').trim();
    let nrLoc = String(item.nrLocalVotacao || '').trim();
    let nmLocOrig = String(item.nmLocalVotacaoOriginal || '').trim();
    let dsEnd = String(item.dsEndereco || item.endereco || '').trim();
    let dsEndOrig = String(item.dsEnderecoLocvtOriginal || '').trim();
    let nmB = String(item.nmBairro || item.bairro || '').trim();
    let nrZ = String(item.nrZona || item.zona || '').trim();
    let nrS = String(item.nrSecao || item.secoes || '').trim();
    let cdTipo = item.cdTipoSecaoAgregada != null ? Number(item.cdTipoSecaoAgregada) : -1;
    if (isNaN(cdTipo)) cdTipo = -1;
    let dsTipo = String(item.dsTipoSecaoAgregada || '').trim();
    let nrSecPrin = item.nrSecaoPrincipal != null ? Number(item.nrSecaoPrincipal) : -1;
    if (isNaN(nrSecPrin)) nrSecPrin = -1;
    let qtEleit = Number(item.qtEleitorSecao ?? item.eleitores) || 0;

    // Detect if nmLoc contains address string (AV, RUA, TRAVESSA) and dsEnd contains school/place name
    const isNmLocAddress = nmLoc.toUpperCase().startsWith("AV") || 
                           nmLoc.toUpperCase().startsWith("RUA") || 
                           nmLoc.toUpperCase().startsWith("TRAVESSA") || 
                           nmLoc.toUpperCase().startsWith("RODOVIA") || 
                           nmLoc.toUpperCase().startsWith("PRAÇA");
    const isDsEndPlace = dsEnd.toUpperCase().includes("ESCOLA") || 
                         dsEnd.toUpperCase().includes("COL") || 
                         dsEnd.toUpperCase().includes("CENTRO") || 
                         dsEnd.toUpperCase().includes("FACULDADE") || 
                         dsEnd.toUpperCase().includes("EMEF") || 
                         dsEnd.toUpperCase().includes("GINASIO") || 
                         dsEnd.toUpperCase().includes("INSTITUTO") || 
                         dsEnd.toUpperCase().includes("CRECHE");

    if (isNmLocAddress && isDsEndPlace) {
      const temp = nmLoc;
      nmLoc = dsEnd;
      dsEnd = temp;
    }

    // Auto-fix if nmLoc was mistakenly set to a purely numeric local number like "1015"
    if (/^\d+$/.test(nmLoc) && nmLoc.length <= 6) {
      if (!nrLoc) {
        nrLoc = nmLoc;
      }
      if (nmLocOrig && !/^\d+$/.test(nmLocOrig)) {
        nmLoc = nmLocOrig;
      } else {
        nmLoc = nrLoc ? `Local de Votação ${nrLoc}` : `Local (${nmB || nmMuni})`;
      }
    }

    // Auto-fix if nrLoc and nmLoc were swapped (e.g. nrLoc has "ESCOLA..." and nmLoc has "1015")
    if (nrLoc && !/^\d+$/.test(nrLoc) && /^\d+$/.test(nmLoc)) {
      const temp = nmLoc;
      nmLoc = nrLoc;
      nrLoc = temp;
    }

    // Fix DS_TIPO_SECAO_AGREGADA display if missing or #NULO
    if (!dsTipo || dsTipo === '#NULO') {
      if (cdTipo === 1) dsTipo = 'Principal';
      else if (cdTipo === 2) dsTipo = 'Agregada';
      else if (cdTipo === 3) dsTipo = 'Distribuída de ofício';
      else if (cdTipo === -1) dsTipo = 'Principal';
      else dsTipo = '#NULO';
    }

    const zonaFormatted = nrZ ? (nrZ.toLowerCase().includes('ze') ? nrZ : `${nrZ}ª ZE`) : '';

    return {
      ...item,
      nmMunicipio: nmMuni || "MUNICÍPIO / BASE LOCAL",
      nrZona: nrZ,
      nrSecao: nrS,
      cdTipoSecaoAgregada: cdTipo,
      dsTipoSecaoAgregada: dsTipo,
      nrSecaoPrincipal: nrSecPrin,
      nrLocalVotacao: nrLoc,
      nmLocalVotacao: nmLoc,
      dsEndereco: dsEnd,
      nmBairro: nmB,
      qtEleitorSecao: qtEleit,
      nmLocalVotacaoOriginal: nmLocOrig || nmLoc,
      dsEnderecoLocvtOriginal: dsEndOrig || dsEnd,

      // Compatibility fields
      municipio: nmMuni || "MUNICÍPIO / BASE LOCAL",
      zona: zonaFormatted || nrZ,
      secoes: nrS,
      secoesCount: item.secoesCount || 1,
      local: nmLoc,
      endereco: dsEnd,
      bairro: nmB,
      eleitores: qtEleit
    };
  };

  // Helper to format/clean items for compact Firestore storage
  const cleanLocationForFirestore = (loc: VotingLocation): any => {
    const s = sanitizeLocation(loc);
    return {
      nmMunicipio: s.nmMunicipio,
      nrZona: s.nrZona,
      nrSecao: s.nrSecao,
      nmLocalVotacao: s.nmLocalVotacao,
      dsEndereco: s.dsEndereco,
      nmBairro: s.nmBairro,
      qtEleitorSecao: s.qtEleitorSecao,
      cdTipoSecaoAgregada: s.cdTipoSecaoAgregada,
      dsTipoSecaoAgregada: s.dsTipoSecaoAgregada,
      nrSecaoPrincipal: s.nrSecaoPrincipal,
      nrLocalVotacao: s.nrLocalVotacao,
      nmLocalVotacaoOriginal: s.nmLocalVotacaoOriginal,
      dsEnderecoLocvtOriginal: s.dsEnderecoLocvtOriginal,

      // Compatibility fields
      municipio: s.municipio,
      zona: s.zona,
      secoes: s.secoes,
      local: s.local,
      endereco: s.endereco,
      bairro: s.bairro,
      eleitores: s.eleitores
    };
  };

  // Helper to aggregate individual section rows into unique Voting Locations for extreme speed & lightweight memory footprint
  const aggregateLocationsIfSections = (rows: VotingLocation[]): VotingLocation[] => {
    if (!rows || rows.length === 0) return [];

    const map = new Map<string, VotingLocation>();

    for (let i = 0; i < rows.length; i++) {
      const sanitized = sanitizeLocation(rows[i]);
      const groupKey = `${sanitized.nmMunicipio.toLowerCase()}|${sanitized.nrZona.toLowerCase()}|${sanitized.nmLocalVotacao.toLowerCase()}`;

      if (!map.has(groupKey)) {
        map.set(groupKey, { ...sanitized });
      } else {
        const existing = map.get(groupKey)!;
        existing.qtEleitorSecao += sanitized.qtEleitorSecao;
        existing.eleitores += sanitized.qtEleitorSecao;
        existing.secoesCount = (existing.secoesCount || 1) + (sanitized.secoesCount || 1);
        if (sanitized.nrSecao && !existing.secoes.includes(sanitized.nrSecao)) {
          existing.secoes = existing.secoes ? `${existing.secoes}, ${sanitized.nrSecao}` : sanitized.nrSecao;
          existing.nrSecao = existing.secoes;
        }
        if (!existing.nrLocalVotacao && sanitized.nrLocalVotacao) {
          existing.nrLocalVotacao = sanitized.nrLocalVotacao;
        }
      }
    }

    return Array.from(map.values());
  };

  // Save parsed data locally and to Firestore database for cross-browser synchronization
  const saveVotingLocations = async (rawNewData: VotingLocation[]) => {
    isSavingRef.current = true;
    const newData = aggregateLocationsIfSections(rawNewData);

    // 1. Update React state, memory cache, localStorage, and IndexedDB immediately
    setVotingLocations(newData);
    setTreLocationsForCoordinator(activeCoordId, newData);

    if (!newData || newData.length === 0) {
      clearTreLocationsCache(activeCoordId);
      await eleitoralStorage.clearLocations(activeCoordId);
      try {
        localStorage.removeItem(`sistema_urna360_eleitoral_data_${activeCoordId}`);
      } catch (e) {}

      try {
        const docId = `coord_${activeCoordId}`;
        await firestoreService.setDocument('eleitoral_data', docId, {
          locations: [],
          cleared: true,
          updatedAt: Date.now(),
          coordinatorId: activeCoordId,
          chunksCount: 0,
          isChunked: false
        }, true);
        for (let i = 0; i < 30; i++) {
          try {
            await firestoreService.deleteDocument('eleitoral_data', `${docId}_${i}`);
          } catch (e) {}
        }
      } catch (err) {
        console.error("Erro ao zerar dados no banco:", err);
      } finally {
        isSavingRef.current = false;
      }
      return;
    }

    try {
      localStorage.setItem(`sistema_urna360_eleitoral_data_${activeCoordId}`, JSON.stringify(newData));
    } catch (e) {}

    // Save to high-capacity IndexedDB locally
    await eleitoralStorage.saveLocations(activeCoordId, newData);

    // Save to Supabase Cloud Database if configured
    if (isSupabaseConfigured()) {
      try {
        await supabaseService.saveTreLocations(activeCoordId, newData as any);
      } catch (err) {
        console.warn("Supabase save warning:", err);
      }
    }

    try {
      const docId = `coord_${activeCoordId}`;
      const cleanData = newData.map(cleanLocationForFirestore);

      await firestoreService.setDocument('eleitoral_data', docId, {
        locations: cleanData,
        cleared: false,
        updatedAt: Date.now(),
        coordinatorId: activeCoordId,
        chunksCount: 1,
        isChunked: false
      }, true);
      setSuccessMsg(`✅ ${newData.length} locais de votação salvos e sincronizados no seu banco de dados com sucesso!`);
    } catch (err: any) {
      console.error("Aviso ao salvar no banco (mantido localmente):", err);
      setSuccessMsg(`✅ ${newData.length} locais de votação salvos no seu navegador com sucesso!`);
    } finally {
      isSavingRef.current = false;
    }
  };

  // Real-time synchronization with Firestore across all browsers/devices
  useEffect(() => {
    let isSubscribed = true;

    // Load from Supabase / IndexedDB / localStorage
    const loadLocalData = async () => {
      try {
        if (isSupabaseConfigured()) {
          const supData = await supabaseService.loadTreLocations(activeCoordId);
          if (isSubscribed && supData && Array.isArray(supData) && supData.length > 0) {
            const aggregated = aggregateLocationsIfSections(supData as any);
            setVotingLocations(aggregated);
            setTreLocationsForCoordinator(activeCoordId, aggregated);
            return;
          }
        }

        const cached = await eleitoralStorage.loadLocations(activeCoordId);
        if (isSubscribed && cached && Array.isArray(cached) && cached.length > 0) {
          const aggregated = aggregateLocationsIfSections(cached);
          setVotingLocations(prev => (prev.length === 0 ? aggregated : prev));
          setTreLocationsForCoordinator(activeCoordId, aggregated);
          return;
        }

        // Fallback to localStorage
        const key = `sistema_urna360_eleitoral_data_${activeCoordId}`;
        const saved = localStorage.getItem(key);
        if (saved) {
          const parsed = JSON.parse(saved);
          if (isSubscribed && Array.isArray(parsed) && parsed.length > 0) {
            const aggregated = aggregateLocationsIfSections(parsed);
            setVotingLocations(prev => (prev.length === 0 ? aggregated : prev));
            setTreLocationsForCoordinator(activeCoordId, aggregated);
          }
        }
      } catch (e) {
        console.warn("Erro ao carregar dados locais do navegador:", e);
      }
    };

    loadLocalData();

    if (!activeCoordId) return;

    const unsub = firestoreService.subscribeToCollectionFiltered<any>('eleitoral_data', activeCoordId, (dataList) => {
      if (isSavingRef.current) return;
      const data = dataList.find(item => item.id === `coord_${activeCoordId}`);
      if (!data) return;

      if (data?.cleared) {
        setVotingLocations([]);
        setTreLocationsForCoordinator(activeCoordId, []);
        clearTreLocationsCache(activeCoordId);
        eleitoralStorage.clearLocations(activeCoordId);
        try {
          localStorage.removeItem(`sistema_urna360_eleitoral_data_${activeCoordId}`);
        } catch (e) {}
        return;
      }

      if (Array.isArray(data?.locations) && data.locations.length > 0) {
        if (isSubscribed) {
          const aggregated = aggregateLocationsIfSections(data.locations);
          setVotingLocations(aggregated);
          setTreLocationsForCoordinator(activeCoordId, aggregated);
          eleitoralStorage.saveLocations(activeCoordId, aggregated);
        }
      }
    });

    return () => {
      isSubscribed = false;
      unsub();
    };
  }, [activeCoordId]);

  // Drag-and-drop state & Modal
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [isLoadingFile, setIsLoadingFile] = useState<boolean>(false);
  const [isImportModalOpen, setIsImportModalOpen] = useState<boolean>(false);
  const [isSupabaseModalOpen, setIsSupabaseModalOpen] = useState<boolean>(false);
  const [dragActive, setDragActive] = useState<boolean>(false);
  const [successMsg, setSuccessMsg] = useState<string | null>(null);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  const triggerFilePicker = () => {
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
      fileInputRef.current.click();
    }
  };

  // Filters State
  const [selectedMunicipio, setSelectedMunicipio] = useState<string>('Todos');
  const [selectedZona, setSelectedZona] = useState<string>('Todos');
  const [selectedBairro, setSelectedBairro] = useState<string>('Todos');
  const [selectedLocal, setSelectedLocal] = useState<string>('Todos');

  // Strategic Tab States
  const [selectedStrategyMun, setSelectedStrategyMun] = useState<string | null>(null);
  const [strategicSort, setStrategicSort] = useState<'coverage' | 'missing'>('coverage');
  const [strategicStatusFilter, setStrategicStatusFilter] = useState<'Todos' | 'critical' | 'low' | 'medium' | 'good'>('Todos');

  // Sorting State for Table
  const [sortField, setSortField] = useState<'local' | 'eleitores'>('eleitores');
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>('desc');

  // Pagination State for Table
  const [currentPage, setCurrentPage] = useState<number>(1);
  const ITEMS_PER_PAGE = 10;

  // Reset page to 1 whenever filters or sort change
  useEffect(() => {
    setCurrentPage(1);
  }, [selectedMunicipio, selectedZona, selectedBairro, selectedLocal, sortField, sortOrder]);

  // Total State Metrics
  const totalStateEleitores = useMemo(() => {
    return votingLocations.reduce((sum, item) => sum + item.eleitores, 0);
  }, [votingLocations]);

  const totalStateMunicipios = useMemo(() => {
    return new Set(votingLocations.map(item => item.municipio)).size;
  }, [votingLocations]);

  const totalStateLocais = votingLocations.length;
  
  const totalStateSecoes = useMemo(() => {
    return votingLocations.reduce((sum, item) => sum + item.secoesCount, 0);
  }, [votingLocations]);

  // Mapping of campaign voters to a municipality
  const mappedCampaignVoters = useMemo(() => {
    return campaignVoters.map((voter: any) => {
      let location: string | null = null;
      
      // 1. Direct match if voter has a municipio field (e.g. from team or direct selection)
      if (voter.municipio) {
        const directMun = voter.municipio.trim();
        const found = MUNICIPALITIES.find(m => m.toLowerCase() === directMun.toLowerCase());
        if (found) location = found;
      }

      // 2. Address string search
      if (!location && voter.address) {
        const addrLower = voter.address.toLowerCase();
        for (const mun of MUNICIPALITIES) {
          if (addrLower.includes(mun.toLowerCase())) {
            location = mun;
            break;
          }
        }
        if (!location) {
          if (addrLower.includes("baliza") || addrLower.includes("são joão") || addrLower.includes("sao joao")) {
            location = "São João da Baliza";
          } else if (addrLower.includes("sao luiz") || addrLower.includes("luiz do anau") || addrLower.includes("são luiz")) {
            location = "São Luiz";
          }
        }
      }

      // 3. Polling place (localVotacao) search
      if (!location && voter.localVotacao) {
        const lvLower = voter.localVotacao.toLowerCase();
        for (const mun of MUNICIPALITIES) {
          if (lvLower.includes(mun.toLowerCase())) {
            location = mun;
            break;
          }
        }
        if (!location) {
          if (lvLower.includes("baliza") || lvLower.includes("são joão") || lvLower.includes("sao joao")) {
            location = "São João da Baliza";
          } else if (lvLower.includes("sao luiz") || lvLower.includes("luiz do anau") || lvLower.includes("são luiz")) {
            location = "São Luiz";
          }
        }
      }

      // 4. Electoral zone match (same as Map)
      if (!location && voter.zona) {
        const z = voter.zona.toString().replace(/\D/g, '');
        const lvText = `${voter.address || ""} ${voter.localVotacao || ""}`.toLowerCase();
        
        if (z === '1') location = "Boa Vista";
        else if (z === '2') location = "Caracaraí";
        else if (z === '3') location = "Alto Alegre";
        else if (z === '4') {
          if (lvText.includes("caroebe")) location = "Caroebe";
          else if (lvText.includes("baliza") || lvText.includes("joão") || lvText.includes("joao")) location = "São João da Baliza";
          else location = "São Luiz";
        }
        else if (z === '5') {
          if (lvText.includes("cantá") || lvText.includes("canta")) location = "Cantá";
          else location = "Boa Vista";
        }
        else if (z === '6') {
          if (lvText.includes("iracema")) location = "Iracema";
          else location = "Mucajaí";
        }
        else if (z === '7') {
          if (lvText.includes("pacaraima")) location = "Pacaraima";
          else if (lvText.includes("uiramutã") || lvText.includes("uiramuta")) location = "Uiramutã";
          else location = "Amajari";
        }
        else if (z === '8') location = "Rorainópolis";
        else if (z === '9') {
          if (lvText.includes("normandia")) location = "Normandia";
          else location = "Bonfim";
        }
      }

      // 5. Team name matching (optional fallback)
      const teamValue = voter.teamName || voter.team;
      if (!location && teamValue) {
        const tLower = teamValue.toLowerCase();
        for (const mun of MUNICIPALITIES) {
          if (tLower.includes(mun.toLowerCase())) {
            location = mun;
            break;
          }
        }
      }

      // Default fallback
      if (!location) {
        location = "Boa Vista";
      }

      return {
        ...voter,
        mappedMunicipio: location
      };
    });
  }, [campaignVoters, MUNICIPALITIES]);

  // Aggregate stats by municipality for BOTH TRE and Campaign data
  const crossReferencedData = useMemo(() => {
    // 1. Group TRE Electors by Municipality
    const treCounts: Record<string, number> = {};
    MUNICIPALITIES.forEach(m => {
      treCounts[m] = 0;
    });
    
    votingLocations.forEach(vl => {
      const mun = vl.municipio;
      if (treCounts[mun] !== undefined) {
        treCounts[mun] += vl.eleitores;
      } else {
        // Try matching with MUNICIPALITIES case-insensitively
        const matchedMun = MUNICIPALITIES.find(m => m.toLowerCase() === mun.toLowerCase().trim());
        if (matchedMun) {
          treCounts[matchedMun] += vl.eleitores;
        }
      }
    });

    // 2. Group Campaign Voters by Municipality
    const campaignCounts: Record<string, number> = {};
    MUNICIPALITIES.forEach(m => {
      campaignCounts[m] = 0;
    });

    mappedCampaignVoters.forEach(mv => {
      const mun = mv.mappedMunicipio;
      if (campaignCounts[mun] !== undefined) {
        campaignCounts[mun]++;
      }
    });

    // 3. Build array of stats
    return MUNICIPALITIES.map(m => {
      const treElectors = treCounts[m];
      const registeredVoters = campaignCounts[m];
      const coverageRate = treElectors > 0 ? (registeredVoters / treElectors) * 100 : 0;
      const missingElectors = Math.max(0, treElectors - registeredVoters);
      
      const treSharePercent = totalStateEleitores > 0 ? (treElectors / totalStateEleitores) * 100 : 0;

      // Status classification
      let status: 'critical' | 'low' | 'medium' | 'good' = 'critical';
      let recommendation = "";

      if (coverageRate < 0.5) {
        status = 'critical';
        recommendation = "Alerta crítico! Cobertura abaixo de 0.5%. Enviar equipes volantes e focar em mutirões de cadastro urgente nesta região.";
      } else if (coverageRate < 1.5) {
        status = 'low';
        recommendation = "Baixa penetração. Necessário intensificar visitas presenciais de lideranças e designar um líder de equipe exclusivo.";
      } else if (coverageRate < 4.0) {
        status = 'medium';
        recommendation = "Presença moderada. Organizar caminhadas e panfletagens focadas nos principais locais de votação de maior eleitorado.";
      } else {
        status = 'good';
        recommendation = "Excelente engajamento. Consolidar rede de apoiadores e focar no monitoramento no dia da eleição para garantir presença.";
      }

      return {
        municipio: m,
        treElectors,
        campaignVoters: registeredVoters,
        coverageRate,
        missingElectors,
        treSharePercent,
        status,
        recommendation
      };
    });
  }, [votingLocations, mappedCampaignVoters, MUNICIPALITIES, totalStateEleitores]);

  // Ranked municipalities to easily see where we have the lowest registered count relative to TRE (least reached)
  const rankedByCoverage = useMemo(() => {
    return [...crossReferencedData].sort((a, b) => {
      if (a.treElectors === 0 && b.treElectors > 0) return 1;
      if (b.treElectors === 0 && a.treElectors > 0) return -1;
      return a.coverageRate - b.coverageRate;
    });
  }, [crossReferencedData]);

  // Ranked by absolute missing voters to target high-yield areas with low coverage
  const rankedByStrategicPriority = useMemo(() => {
    return [...crossReferencedData].sort((a, b) => {
      return b.missingElectors - a.missingElectors;
    });
  }, [crossReferencedData]);

  // Cascading lists for filters based on selected Municipio
  const municipiosList = useMemo(() => {
    const set = new Set<string>();
    for (let i = 0; i < votingLocations.length; i++) {
      if (votingLocations[i].municipio) set.add(votingLocations[i].municipio);
    }
    return ['Todos', ...Array.from(set).sort((a, b) => a.localeCompare(b, 'pt-BR'))];
  }, [votingLocations]);

  const zonasList = useMemo(() => {
    const set = new Set<string>();
    for (let i = 0; i < votingLocations.length; i++) {
      const item = votingLocations[i];
      if (selectedMunicipio === 'Todos' || item.municipio === selectedMunicipio) {
        if (item.zona) set.add(item.zona);
      }
    }
    return ['Todos', ...Array.from(set).sort((a, b) => a.localeCompare(b, 'pt-BR'))];
  }, [selectedMunicipio, votingLocations]);

  const bairrosList = useMemo(() => {
    const set = new Set<string>();
    for (let i = 0; i < votingLocations.length; i++) {
      const item = votingLocations[i];
      const matchMun = selectedMunicipio === 'Todos' || item.municipio === selectedMunicipio;
      const matchZona = selectedZona === 'Todos' || item.zona === selectedZona;
      if (matchMun && matchZona) {
        if (item.bairro) set.add(item.bairro);
      }
    }
    return ['Todos', ...Array.from(set).sort((a, b) => a.localeCompare(b, 'pt-BR'))];
  }, [selectedMunicipio, selectedZona, votingLocations]);

  const locaisList = useMemo(() => {
    const set = new Set<string>();
    for (let i = 0; i < votingLocations.length; i++) {
      const item = votingLocations[i];
      const matchMun = selectedMunicipio === 'Todos' || item.municipio === selectedMunicipio;
      const matchZona = selectedZona === 'Todos' || item.zona === selectedZona;
      const matchBairro = selectedBairro === 'Todos' || item.bairro === selectedBairro;
      if (matchMun && matchZona && matchBairro) {
        if (item.local) set.add(item.local);
      }
    }
    return ['Todos', ...Array.from(set).sort((a, b) => a.localeCompare(b, 'pt-BR'))];
  }, [selectedMunicipio, selectedZona, selectedBairro, votingLocations]);

  // Handle cascading reset on parent filter change
  const handleMunicipioChange = (val: string) => {
    setSelectedMunicipio(val);
    setSelectedZona('Todos');
    setSelectedBairro('Todos');
    setSelectedLocal('Todos');
  };

  const handleZonaChange = (val: string) => {
    setSelectedZona(val);
    setSelectedBairro('Todos');
    setSelectedLocal('Todos');
  };

  const handleBairroChange = (val: string) => {
    setSelectedBairro(val);
    setSelectedLocal('Todos');
  };

  const resetFilters = () => {
    setSelectedMunicipio('Todos');
    setSelectedZona('Todos');
    setSelectedBairro('Todos');
    setSelectedLocal('Todos');
  };

  // Filtered dataset for computations and UI display
  const filteredData = useMemo(() => {
    if (selectedMunicipio === 'Todos' && selectedZona === 'Todos' && selectedBairro === 'Todos' && selectedLocal === 'Todos') {
      return votingLocations;
    }
    return votingLocations.filter(item => {
      const matchMun = selectedMunicipio === 'Todos' || item.municipio === selectedMunicipio;
      const matchZona = selectedZona === 'Todos' || item.zona === selectedZona;
      const matchBairro = selectedBairro === 'Todos' || item.bairro === selectedBairro;
      const matchLocal = selectedLocal === 'Todos' || item.local === selectedLocal;
      return matchMun && matchZona && matchBairro && matchLocal;
    });
  }, [selectedMunicipio, selectedZona, selectedBairro, selectedLocal, votingLocations]);

  // Current scope totals
  const totalScopeEleitores = useMemo(() => {
    let sum = 0;
    for (let i = 0; i < filteredData.length; i++) sum += filteredData[i].eleitores;
    return sum;
  }, [filteredData]);

  const totalScopeMunicipios = useMemo(() => {
    const set = new Set<string>();
    for (let i = 0; i < filteredData.length; i++) set.add(filteredData[i].municipio);
    return set.size;
  }, [filteredData]);

  const totalScopeLocais = filteredData.length;
  const totalScopeSecoes = useMemo(() => {
    let sum = 0;
    for (let i = 0; i < filteredData.length; i++) sum += (filteredData[i].secoesCount || 1);
    return sum;
  }, [filteredData]);

  // Calculation of KPIs
  const kpiMetrics = useMemo(() => {
    if (filteredData.length === 0) {
      return {
        maiorLocal: { local: 'Nenhum', eleitores: 0, percentMun: 0 },
        menorLocal: { local: 'Nenhum', eleitores: 0, percentMun: 0 },
        mediaEleitores: 0,
        top5AccumulatedPercent: 0
      };
    }

    // Sort locations in current scope by electors descending
    const sorted = [...filteredData].sort((a, b) => b.eleitores - a.eleitores);
    
    const maior = sorted[0];
    const menor = sorted[sorted.length - 1];
    
    const media = Math.round(totalScopeEleitores / sorted.length);

    // Sum of top 5 locations
    const top5Sum = sorted.slice(0, 5).reduce((sum, item) => sum + item.eleitores, 0);
    const top5Percent = totalScopeEleitores > 0 ? (top5Sum / totalScopeEleitores) * 100 : 0;

    // Percent representativeness of the biggest/smallest in the municipality (or scope total)
    const maiorPercent = totalScopeEleitores > 0 ? (maior.eleitores / totalScopeEleitores) * 100 : 0;
    const menorPercent = totalScopeEleitores > 0 ? (menor.eleitores / totalScopeEleitores) * 100 : 0;

    return {
      maiorLocal: { local: maior.local, eleitores: maior.eleitores, percentMun: maiorPercent },
      menorLocal: { local: menor.local, eleitores: menor.eleitores, percentMun: menorPercent },
      mediaEleitores: media,
      top5AccumulatedPercent: top5Percent
    };
  }, [filteredData, totalScopeEleitores]);

  // Chart 1 data: Horizonal Bar Chart of Voting Locations in the selected municipality/scope (Top 10)
  const barChartData = useMemo(() => {
    const sorted = [...filteredData].sort((a, b) => b.eleitores - a.eleitores);
    // Return top 10 to keep the layout neat and high quality
    return sorted.slice(0, 10).map(item => ({
      name: item.local.length > 28 ? item.local.substring(0, 25) + '...' : item.local,
      fullName: item.local,
      eleitores: item.eleitores
    })).reverse(); // Reverse so Recharts vertical renders highest at the top
  }, [filteredData]);

  // Chart 2 data: Pie/Donut Chart for local distribution in selected municipality/scope (Top 5 + Others)
  const pieChartData = useMemo(() => {
    const sorted = [...filteredData].sort((a, b) => b.eleitores - a.eleitores);
    if (sorted.length <= 5) {
      return sorted.map(item => ({
        name: item.local.length > 20 ? item.local.substring(0, 18) + '...' : item.local,
        value: item.eleitores
      }));
    } else {
      const top5 = sorted.slice(0, 5).map(item => ({
        name: item.local.length > 20 ? item.local.substring(0, 18) + '...' : item.local,
        value: item.eleitores
      }));
      const othersSum = sorted.slice(5).reduce((sum, item) => sum + item.eleitores, 0);
      return [
        ...top5,
        { name: 'Outros Locais', value: othersSum }
      ];
    }
  }, [filteredData]);

  // Chart 3 data: Ranking of Municipalities by Voter Count (Statewide context)
  const muniRankingChartData = useMemo(() => {
    const grouped: Record<string, number> = {};
    for (let i = 0; i < votingLocations.length; i++) {
      const item = votingLocations[i];
      if (item.municipio) {
        grouped[item.municipio] = (grouped[item.municipio] || 0) + item.eleitores;
      }
    }
    return Object.entries(grouped)
      .map(([name, val]) => ({ name, eleitores: val }))
      .sort((a, b) => b.eleitores - a.eleitores);
  }, [votingLocations]);

  // Compute Município-level details (for selected Municipio panel)
  const municipioPanelData = useMemo(() => {
    if (selectedMunicipio === 'Todos') return null;

    let muniTotalEleitores = 0;
    let muniLocaisCount = 0;
    let muniSecoesCount = 0;

    for (let i = 0; i < votingLocations.length; i++) {
      const item = votingLocations[i];
      if (item.municipio === selectedMunicipio) {
        muniTotalEleitores += item.eleitores;
        muniLocaisCount++;
        muniSecoesCount += (item.secoesCount || 1);
      }
    }
    const representativeness = totalStateEleitores > 0 ? (muniTotalEleitores / totalStateEleitores) * 100 : 0;

    return {
      nome: selectedMunicipio,
      eleitores: muniTotalEleitores,
      locais: muniLocaisCount,
      secoes: muniSecoesCount,
      representatividade: representativeness
    };
  }, [selectedMunicipio, totalStateEleitores, votingLocations]);

  // Pre-calculate totals per municipality in O(N)
  const muniTotalsMap = useMemo(() => {
    const map: Record<string, number> = {};
    for (let i = 0; i < votingLocations.length; i++) {
      const m = votingLocations[i].municipio;
      if (m) map[m] = (map[m] || 0) + (votingLocations[i].eleitores || 0);
    }
    return map;
  }, [votingLocations]);

  // Compute detailed voting locations sorted for the dyn table in O(N)
  const processedTableData = useMemo(() => {
    return filteredData.map(item => {
      const muniTotal = muniTotalsMap[item.municipio] || 0;
      const percentMuni = muniTotal > 0 ? (item.eleitores / muniTotal) * 100 : 0;
      const percentTotal = totalStateEleitores > 0 ? (item.eleitores / totalStateEleitores) * 100 : 0;

      return {
        ...item,
        percentMuni,
        percentTotal
      };
    }).sort((a, b) => {
      if (sortField === 'eleitores') {
        return sortOrder === 'desc' ? b.eleitores - a.eleitores : a.eleitores - b.eleitores;
      } else {
        return sortOrder === 'desc' 
          ? b.local.localeCompare(a.local) 
          : a.local.localeCompare(b.local);
      }
    });
  }, [filteredData, sortField, sortOrder, totalStateEleitores, muniTotalsMap]);

  const totalPages = Math.max(1, Math.ceil(processedTableData.length / ITEMS_PER_PAGE));

  const paginatedTableData = useMemo(() => {
    const startIndex = (currentPage - 1) * ITEMS_PER_PAGE;
    return processedTableData.slice(startIndex, startIndex + ITEMS_PER_PAGE);
  }, [processedTableData, currentPage]);

  const toggleSort = (field: 'local' | 'eleitores') => {
    if (sortField === field) {
      setSortOrder(prev => prev === 'asc' ? 'desc' : 'asc');
    } else {
      setSortField(field);
      setSortOrder('desc');
    }
  };

  // EXCEL IMPORTER & TEMPLATE GENERATOR
  const downloadTemplate = () => {
    const headers = [
      TSE_COLUMNS.map(col => col.variable)
    ];
    
    const ws = XLSX.utils.aoa_to_sheet(headers);
    ws['!cols'] = TSE_COLUMNS.map(col => ({ wch: Math.max(col.variable.length + 4, 18) }));

    // Example rows in official TSE format
    XLSX.utils.sheet_add_aoa(ws, [
      [
        "SÃO PAULO", "001", "0001", 1, "Principal", -1, 1015,
        "ESCOLA ESTADUAL PADRE ANCHIETA", "RUA DOS TRÊS IRMÃOS, 100", "MORUMBI",
        380, "ESCOLA ESTADUAL PADRE ANCHIETA", "RUA DOS TRÊS IRMÃOS, 100"
      ],
      [
        "RIO DE JANEIRO", "004", "0012", 2, "Agregada", 11, 1088,
        "COLÉGIO PEDRO II", "AVENIDA MARECHAL FLORIANO, 80", "CENTRO",
        295, "COLÉGIO PEDRO II", "AVENIDA MARECHAL FLORIANO, 80"
      ]
    ], { origin: -1 });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Tabela_Variaveis_TSE");
    XLSX.writeFile(wb, "modelo_planilha_oficial_tse.xlsx");
    
    setSuccessMsg("Modelo oficial de planilha do TSE (.xlsx) baixado com sucesso!");
    setTimeout(() => setSuccessMsg(null), 5000);
  };

  const normalizeHeaderKey = (str: any): string => {
    return String(str || '')
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "") // removes accents (e.g. Í -> I, Ç -> C)
      .toLowerCase()
      .replace(/[^a-z0-9]/g, ''); // removes spaces, underscores, punctuation
  };

  const parseEleitoresCount = (val: any): number => {
    if (val === undefined || val === null || val === '') return 0;
    if (typeof val === 'number') return isNaN(val) ? 0 : val;
    const str = String(val).trim();
    const cleanStr = str.replace(/\./g, '').replace(/,/g, '.').replace(/[^0-9.]/g, '');
    const num = parseFloat(cleanStr);
    return isNaN(num) ? 0 : Math.round(num);
  };

  const parseCSVText = (text: string): string[][] => {
    const cleanText = text.replace(/^\uFEFF/, '');
    const rawLines = cleanText.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
    if (rawLines.length === 0) return [];

    const sampleLines = rawLines.slice(0, Math.min(10, rawLines.length));
    let countSemicolon = 0;
    let countTab = 0;
    let countComma = 0;

    for (const line of sampleLines) {
      countSemicolon += (line.match(/;/g) || []).length;
      countTab += (line.match(/\t/g) || []).length;
      countComma += (line.match(/,/g) || []).length;
    }

    let delimiter = ';';
    if (countTab > countSemicolon && countTab > countComma) {
      delimiter = '\t';
    } else if (countComma > countSemicolon && countComma > countTab) {
      delimiter = ',';
    }

    return rawLines.map(line => {
      if (line.includes('"')) {
        const cells: string[] = [];
        let current = '';
        let inQuotes = false;
        for (let i = 0; i < line.length; i++) {
          const char = line[i];
          if (char === '"') {
            inQuotes = !inQuotes;
          } else if (char === delimiter && !inQuotes) {
            cells.push(current.trim().replace(/^["']|["']$/g, ''));
            current = '';
          } else {
            current += char;
          }
        }
        cells.push(current.trim().replace(/^["']|["']$/g, ''));
        return cells;
      } else {
        return line.split(delimiter).map(cell => cell.trim());
      }
    });
  };

  const processFile = (file: File) => {
    setSuccessMsg(null);
    setErrorMsg(null);
    setIsLoadingFile(true);

    const reader = new FileReader();
    reader.onload = async (event) => {
      try {
        const buffer = event.target?.result as ArrayBuffer;
        if (!buffer) {
          setErrorMsg("Não foi possível ler os dados do arquivo selecionado.");
          setIsLoadingFile(false);
          return;
        }

        const data = new Uint8Array(buffer);
        const fileName = file.name.toLowerCase();

        let matrix: any[][] = parseExcelOrCSVBuffer(buffer, file.name);

        if (!matrix || matrix.length === 0) {
          setErrorMsg("A planilha selecionada está vazia ou não pôde ser lida. Certifique-se de carregar um arquivo .xlsx, .xls ou .csv válido.");
          setIsLoadingFile(false);
          return;
        }

        matrix = matrix.map(row => {
          if (Array.isArray(row) && row.length === 1 && typeof row[0] === 'string') {
            const line = row[0];
            const delim = line.includes(';') ? ';' : (line.includes('\t') ? '\t' : ',');
            return line.split(delim).map(cell => String(cell || '').replace(/^["']|["']$/g, '').trim());
          }
          return Array.isArray(row) ? row.map(cell => String(cell ?? '').trim()) : [];
        }).filter(r => r.length > 0 && r.some(cell => cell.length > 0));

        if (matrix.length === 0) {
          setErrorMsg("Nenhum dado encontrado no arquivo.");
          setIsLoadingFile(false);
          return;
        }

        const headerKeywords = [
          "municipio", "cidade", "nmmunicipio", "cdmunicipio", "nomemunicipio", "mun", "ds_municipio", "nm_municipio",
          "local", "escola", "nmlocalvotacao", "localvotacao", "colegio", "estabelecimento", "nm_local_votacao", "ds_local_votacao", "nm_estabelecimento",
          "zona", "nrzona", "ze", "nr_zona", "zonaeleitoral",
          "secao", "nrsecao", "secoes", "nr_secao", "numsecao",
          "bairro", "nmbairro", "nm_bairro",
          "endereco", "dsendereco", "logradouro", "ds_endereco",
          "eleitor", "eleitores", "aptos", "qteleitorsecao", "qteleitor", "qt_aptos", "qt_eleitores"
        ];

        let bestHeaderRowIndex = 0;
        let maxMatches = -1;

        const scanLimit = Math.min(30, matrix.length);
        for (let r = 0; r < scanLimit; r++) {
          const row = matrix[r];
          if (!Array.isArray(row)) continue;
          let matches = 0;
          for (let c = 0; c < row.length; c++) {
            const normCell = normalizeHeaderKey(row[c]);
            if (normCell && headerKeywords.some(kw => normCell.includes(kw))) {
              matches++;
            }
          }
          if (matches > maxMatches) {
            maxMatches = matches;
            bestHeaderRowIndex = r;
          }
        }

        const headerRow = matrix[bestHeaderRowIndex] || [];
        const detectedHeadersRaw = headerRow.map(cell => String(cell || '').trim());
        const normHeaders = detectedHeadersRaw.map(cell => normalizeHeaderKey(cell));

        const munTargets = [
          "nmmunicipio", "nm_municipio", "municipio", "cidade", "nomemunicipio", "cdmunicipio", "cd_municipio", "dsmunicipio", "ds_municipio", "mun"
        ];
        const nrLocalTargets = [
          "nrlocalvotacao", "nr_local_votacao", "cdlocalvotacao", "cd_local_votacao", "numlocalvotacao", "num_local_votacao", "nrlocal", "nr_local", "cdlocal", "cd_local", "codlocal", "nr_locvt"
        ];
        const localTargets = [
          "nmlocalvotacao", "nm_local_votacao", "localdevotacao", "local_de_votacao", "localvotacao", "local_votacao", 
          "nomelocal", "nome_local", "nmlocal", "nm_local", "escola", "colegio", "nmestabelecimento", "nm_estabelecimento", 
          "estabelecimento", "locdevotacao", "local"
        ];
        const zonaTargets = ["nrzona", "nr_zona", "zona", "ze", "zonaeleitoral", "zona_eleitoral", "numzona", "cdzona", "cd_zona"];
        const secaoTargets = ["nrsecao", "nr_secao", "secao", "secoes", "numsecao", "cdsecao", "cd_secao"];
        const enderecoTargets = [
          "dsendereco", "ds_endereco", "dslocalvotacao", "ds_local_votacao", "ds_local", "dslocal", 
          "endereco", "logradouro", "rua", "locvtendereco", "ds_localizacao", "localizacao"
        ];
        const bairroTargets = ["nmbairro", "nm_bairro", "bairro", "regiao", "distrito", "dsbairro", "ds_bairro"];
        const eleitorTargets = [
          "qteleitorsecao", "qt_eleitor_secao", "qteleitoressecao", "qt_eleitores_secao", "qteleitor", "qt_eleitor", 
          "eleitores", "aptos", "quantidadedeeleitoresaptos", "quantidadedeeleitores", "totaleleitores", "qteleitores", 
          "qt_eleitores", "numeleitores", "qtaptos", "qt_aptos", "qteleitoresperfil", "totaleleitor"
        ];
        const tipoAgregadaTargets = ["cdtiposecaoagregada", "cd_tipo_secao_agregada", "cdtiposecao", "cd_tipo_secao", "tiposecao", "tipo_secao"];
        const dsTipoAgregadaTargets = ["dstiposecaoagregada", "ds_tipo_secao_agregada", "dstiposecao", "ds_tipo_secao", "descripcaotiposecao", "desc_tipo_secao"];
        const secaoPrincipalTargets = ["nrsecaoprincipal", "nr_secao_principal", "secaoprincipal", "secao_principal", "numsecaoprincipal"];
        const localOriginalTargets = ["nmlocalvotacaooriginal", "nm_local_votacao_original", "localoriginal", "local_original", "nm_local_original"];
        const enderecoOriginalTargets = ["dsenderecolocvtoriginal", "ds_endereco_locvt_original", "dsenderecooriginal", "ds_endereco_original", "enderecooriginal"];

        const findColIdx = (targets: string[], excludeTargets: string[] = []): number => {
          const normExclude = excludeTargets.map(t => normalizeHeaderKey(t));

          // 1. Priority pass: Exact match
          for (const target of targets) {
            const normTarget = normalizeHeaderKey(target);
            for (let i = 0; i < normHeaders.length; i++) {
              const h = normHeaders[i];
              if (!h) continue;
              if (normExclude.some(ex => h === ex)) continue;
              if (h === normTarget) return i;
            }
          }

          // 2. Second pass: Substring/Includes match
          for (const target of targets) {
            const normTarget = normalizeHeaderKey(target);
            for (let i = 0; i < normHeaders.length; i++) {
              const h = normHeaders[i];
              if (!h) continue;
              if (normExclude.some(ex => h.includes(ex) || ex.includes(h))) continue;
              if (h.includes(normTarget) || normTarget.includes(h)) return i;
            }
          }

          return -1;
        };

        const colMun = findColIdx(munTargets);
        const colNrLocal = findColIdx(nrLocalTargets, ["nmlocal", "nm_local", "ds_local", "localoriginal"]);
        const colLocal = findColIdx(localTargets, [
          "nrlocal", "nr_local", "cdlocal", "cd_local", "numlocal", 
          "dsendereco", "ds_endereco", "dslocal", "ds_local", "dslocalvotacao", "ds_local_votacao", 
          "nmlocalvotacaooriginal", "nm_local_votacao_original", "localoriginal", "local_original"
        ]);
        const colZona = findColIdx(zonaTargets);
        const colSecao = findColIdx(secaoTargets, ["nrsecaoprincipal", "nr_secao_principal", "secaoprincipal"]);
        const colEndereco = findColIdx(enderecoTargets, ["dsenderecolocvtoriginal", "ds_endereco_locvt_original", "enderecooriginal"]);
        const colBairro = findColIdx(bairroTargets);
        const colEleitores = findColIdx(eleitorTargets);
        const colTipoAgregada = findColIdx(tipoAgregadaTargets, ["dstiposecaoagregada", "ds_tipo_secao_agregada"]);
        const colDsTipoAgregada = findColIdx(dsTipoAgregadaTargets, ["cdtiposecaoagregada", "cd_tipo_secao_agregada"]);
        const colSecaoPrincipal = findColIdx(secaoPrincipalTargets);
        const colLocalOriginal = findColIdx(localOriginalTargets);
        const colEnderecoOriginal = findColIdx(enderecoOriginalTargets);

        const dataRows = matrix.slice(bestHeaderRowIndex + 1);
        const parsedRows: VotingLocation[] = [];

        for (let i = 0; i < dataRows.length; i++) {
          const row = dataRows[i];
          if (!Array.isArray(row) || row.length === 0) continue;

          let nmMunicipio = colMun !== -1 && row[colMun] !== undefined && row[colMun] !== null ? String(row[colMun]).trim() : '';
          let nrLocalVotacao = colNrLocal !== -1 && row[colNrLocal] !== undefined && row[colNrLocal] !== null ? String(row[colNrLocal]).trim() : '';
          let nmLocalVotacao = colLocal !== -1 && row[colLocal] !== undefined && row[colLocal] !== null ? String(row[colLocal]).trim() : '';
          let nmLocalVotacaoOriginal = colLocalOriginal !== -1 && row[colLocalOriginal] !== undefined && row[colLocalOriginal] !== null ? String(row[colLocalOriginal]).trim() : '';
          let dsEnderecoLocvtOriginal = colEnderecoOriginal !== -1 && row[colEnderecoOriginal] !== undefined && row[colEnderecoOriginal] !== null ? String(row[colEnderecoOriginal]).trim() : '';
          let dsTipoSecaoAgregada = colDsTipoAgregada !== -1 && row[colDsTipoAgregada] !== undefined && row[colDsTipoAgregada] !== null ? String(row[colDsTipoAgregada]).trim() : '';

          // Smart auto-correction if nrLocalVotacao and nmLocalVotacao were misplaced/swapped
          if (/^\d+$/.test(nmLocalVotacao) && nmLocalVotacao.length <= 6) {
            if (!nrLocalVotacao) {
              nrLocalVotacao = nmLocalVotacao;
            }
            if (nmLocalVotacaoOriginal && !/^\d+$/.test(nmLocalVotacaoOriginal)) {
              nmLocalVotacao = nmLocalVotacaoOriginal;
            } else {
              nmLocalVotacao = '';
            }
          }

          if (!/^\d+$/.test(nrLocalVotacao) && /^\d+$/.test(nmLocalVotacao)) {
            const temp = nmLocalVotacao;
            nmLocalVotacao = nrLocalVotacao;
            nrLocalVotacao = temp;
          }

          if (!nmLocalVotacao) {
            for (let c = 0; c < row.length; c++) {
              if (c === colMun || c === colZona || c === colSecao || c === colEleitores || c === colNrLocal) continue;
              const cellVal = String(row[c] || '').trim();
              if (cellVal.length > 3 && (
                cellVal.toUpperCase().includes("ESCOLA") || 
                cellVal.toUpperCase().includes("COL") || 
                cellVal.toUpperCase().includes("CENTRO") || 
                cellVal.toUpperCase().includes("FACULDADE") || 
                cellVal.toUpperCase().includes("EMEF") || 
                cellVal.toUpperCase().includes("CRECHE") ||
                cellVal.toUpperCase().includes("E.E") ||
                cellVal.toUpperCase().includes("E.M") ||
                cellVal.toUpperCase().includes("GINASIO") ||
                cellVal.toUpperCase().includes("UNIDADE") ||
                cellVal.toUpperCase().includes("INSTITUTO") ||
                cellVal.toUpperCase().includes("POSTO") ||
                cellVal.toUpperCase().includes("CAMARA") ||
                cellVal.toUpperCase().includes("PREFEITURA") ||
                cellVal.toUpperCase().includes("SECRETARIA") ||
                cellVal.toUpperCase().includes("ASSOCIACAO") ||
                cellVal.toUpperCase().includes("TRIBUNAL") ||
                cellVal.toUpperCase().includes("IFAC") ||
                cellVal.toUpperCase().includes("UFAC") ||
                cellVal.toUpperCase().includes("SESC") ||
                cellVal.toUpperCase().includes("SENAI") ||
                cellVal.toUpperCase().includes("SENAC") ||
                cellVal.toUpperCase().includes("INCRA") ||
                cellVal.toUpperCase().includes("IDAF") ||
                cellVal.toUpperCase().includes("CRAS")
              )) {
                nmLocalVotacao = cellVal;
                break;
              }
            }
          }

          if (!nmLocalVotacao && row.some(cell => String(cell || '').trim().length > 0)) {
            for (let c = 0; c < row.length; c++) {
              if (c === colMun || c === colZona || c === colSecao || c === colEleitores || c === colNrLocal) continue;
              const val = String(row[c] || '').trim();
              if (val.length > 3 && isNaN(Number(val))) {
                nmLocalVotacao = val;
                break;
              }
            }
          }

          if (!nmLocalVotacao && nrLocalVotacao) {
            nmLocalVotacao = `Local de Votação ${nrLocalVotacao}`;
          }

          if (!nmLocalVotacao) continue;

          if (!nmMunicipio) {
            nmMunicipio = "MUNICÍPIO ÚNICO";
          }

          const nrZona = colZona !== -1 && row[colZona] !== undefined && row[colZona] !== null ? String(row[colZona]).trim() : '';
          const nrSecao = colSecao !== -1 && row[colSecao] !== undefined && row[colSecao] !== null ? String(row[colSecao]).trim() : '';
          const dsEndereco = colEndereco !== -1 && row[colEndereco] !== undefined && row[colEndereco] !== null ? String(row[colEndereco]).trim() : '';
          const nmBairro = colBairro !== -1 && row[colBairro] !== undefined && row[colBairro] !== null ? String(row[colBairro]).trim() : '';
          const qtEleitorSecao = colEleitores !== -1 ? parseEleitoresCount(row[colEleitores]) : 0;
          const cdTipoSecaoAgregada = colTipoAgregada !== -1 ? Number(row[colTipoAgregada]) || -1 : -1;
          const nrSecaoPrincipal = colSecaoPrincipal !== -1 ? Number(row[colSecaoPrincipal]) || -1 : -1;

          if (!dsTipoSecaoAgregada || dsTipoSecaoAgregada === '#NULO') {
            if (cdTipoSecaoAgregada === 1) dsTipoSecaoAgregada = 'Principal';
            else if (cdTipoSecaoAgregada === 2) dsTipoSecaoAgregada = 'Agregada';
            else if (cdTipoSecaoAgregada === 3) dsTipoSecaoAgregada = 'Distribuída de ofício';
            else if (cdTipoSecaoAgregada === -1) dsTipoSecaoAgregada = 'Principal';
            else dsTipoSecaoAgregada = '#NULO';
          }

          const zonaFormatted = nrZona ? (nrZona.toLowerCase().includes('ze') ? nrZona : `${nrZona}ª ZE`) : '';

          parsedRows.push({
            nmMunicipio,
            nrZona,
            nrSecao,
            cdTipoSecaoAgregada,
            dsTipoSecaoAgregada,
            nrSecaoPrincipal,
            nrLocalVotacao,
            nmLocalVotacao,
            dsEndereco,
            nmBairro,
            qtEleitorSecao,
            nmLocalVotacaoOriginal: nmLocalVotacaoOriginal || nmLocalVotacao,
            dsEnderecoLocvtOriginal: dsEnderecoLocvtOriginal || dsEndereco,

            municipio: nmMunicipio,
            zona: zonaFormatted || nrZona,
            secoes: nrSecao,
            secoesCount: 1,
            local: nmLocalVotacao,
            endereco: dsEndereco,
            bairro: nmBairro,
            eleitores: qtEleitorSecao
          });
        }

        if (parsedRows.length === 0 && dataRows.length > 0) {
          for (let i = 0; i < dataRows.length; i++) {
            const row = dataRows[i];
            if (!Array.isArray(row) || row.length === 0) continue;

            let bestLocal = "";
            let bestMuni = "MUNICÍPIO ÚNICO";
            let bestZona = "";
            let bestSecao = "";
            let maxEleit = 0;

            for (let c = 0; c < row.length; c++) {
              const val = row[c];
              if (val === undefined || val === null || val === "") continue;
              const strVal = String(val).trim();
              const numVal = parseEleitoresCount(val);

              if (numVal > maxEleit && numVal < 1000000) {
                maxEleit = numVal;
              } else if (strVal.length > 3 && isNaN(Number(strVal))) {
                if (!bestLocal) {
                  bestLocal = strVal;
                } else if (bestMuni === "MUNICÍPIO ÚNICO") {
                  bestMuni = strVal;
                }
              }
            }

            if (!bestLocal && maxEleit === 0) continue;
            if (!bestLocal) bestLocal = `Local de Votação ${i + 1}`;

            parsedRows.push({
              nmMunicipio: bestMuni,
              nrZona: bestZona,
              nrSecao: bestSecao,
              cdTipoSecaoAgregada: -1,
              dsTipoSecaoAgregada: '#NULO',
              nrSecaoPrincipal: -1,
              nrLocalVotacao: '',
              nmLocalVotacao: bestLocal,
              dsEndereco: '',
              nmBairro: '',
              qtEleitorSecao: maxEleit,
              nmLocalVotacaoOriginal: bestLocal,
              dsEnderecoLocvtOriginal: '',

              municipio: bestMuni,
              zona: bestZona,
              secoes: bestSecao,
              secoesCount: 1,
              local: bestLocal,
              endereco: '',
              bairro: '',
              eleitores: maxEleit
            });
          }
        }

        if (parsedRows.length === 0) {
          setErrorMsg("Não foi possível reconhecer as colunas de locais de votação nesta planilha. Verifique se o arquivo possui colunas com nomes de município, local de votação e eleitores.");
          setIsLoadingFile(false);
          return;
        }

        await saveVotingLocations(parsedRows);

        const totalEleitoresCalc = parsedRows.reduce((acc, r) => acc + (r.eleitores || 0), 0);
        const uniqueLocsCount = new Set(parsedRows.map(r => r.local)).size;

        setSuccessMsg(`✅ Sucesso! Foram importados e salvos no banco de dados ${uniqueLocsCount} locais de votação com um total de ${totalEleitoresCalc.toLocaleString('pt-BR')} eleitores.`);
        setTimeout(() => setSuccessMsg(null), 8000);
      } catch (err: any) {
        console.error("Erro detalhado ao ler planilha:", err);
        setErrorMsg(`Erro ao processar o arquivo (${err?.message || 'formato ou dados inválidos'}). Certifique-se de carregar uma planilha Excel ou CSV válida.`);
      } finally {
        setIsLoadingFile(false);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      processFile(file);
      e.target.value = '';
    }
  };

  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      processFile(e.dataTransfer.files[0]);
    }
  };

  const clearData = () => {
    if (window.confirm("Aviso: Deseja realmente zerar todos os dados eleitorais oficiais salvos no sistema?")) {
      saveVotingLocations([]);
      setSuccessMsg("O banco de dados de locais do TRE foi limpo!");
      setTimeout(() => setSuccessMsg(null), 3000);
    }
  };

  const loadDemoData = () => {
    saveVotingLocations(SAMPLE_TRE_DATA);
    setSuccessMsg("Dados demonstrativos carregados com sucesso!");
    setTimeout(() => setSuccessMsg(null), 3000);
  };

  return (
    <div id="eleitoral_bi_dashboard" className="w-full bg-zinc-50 dark:bg-zinc-950 text-zinc-800 dark:text-zinc-100 p-4 md:p-6 space-y-6">
      
      {/* SINGLE RELIABLE HIDDEN FILE INPUT */}
      <input 
        ref={fileInputRef}
        type="file" 
        id="excel-file-upload-input"
        className="hidden" 
        accept=".xlsx, .xls, .csv, .txt" 
        onChange={handleFileUpload} 
      />

      {/* LOADING OVERLAY WHEN PROCESSING PLANILHA */}
      {isLoadingFile && (
        <div className="fixed inset-0 bg-black/70 backdrop-blur-sm z-50 flex flex-col items-center justify-center p-4">
          <div className="bg-zinc-900 border border-zinc-800 rounded-lg p-8 max-w-sm w-full text-center space-y-4 shadow-2xl animate-in fade-in zoom-in duration-150">
            <RefreshCw className="w-12 h-12 text-blue-500 animate-spin mx-auto" />
            <div>
              <h3 className="text-base font-black text-white uppercase tracking-wider">Processando Planilha TRE</h3>
              <p className="text-xs text-zinc-400 mt-1">Lendo linhas, agregando locais de votação e sincronizando banco de dados...</p>
            </div>
          </div>
        </div>
      )}
      
      {/* HEADER SECTION */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-zinc-200 dark:border-zinc-800 pb-5">
        <div>
          <h1 className="text-2xl md:text-3xl font-black text-zinc-900 dark:text-white tracking-tight uppercase">
            Análise Eleitoral
          </h1>
          <p className="text-xs md:text-sm text-zinc-500 dark:text-zinc-400 mt-1">
            Monitoramento analítico de eleitores aptos, locais de votação e representatividade estatística.
          </p>
        </div>
        
        {/* RESET FILTERS BUTTON */}
        <div className="flex items-center gap-2 flex-wrap">
          {votingLocations.length > 0 && (
            <button 
              onClick={resetFilters}
              className="flex items-center justify-center gap-2 px-3 py-1.5 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-sm text-xs font-black uppercase tracking-tight text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 hover:text-blue-600 dark:hover:text-blue-600 transition-colors shadow-sm"
            >
              <RefreshCw className="w-3.5 h-3.5" />
              <span>Limpar Filtros</span>
            </button>
          )}
        </div>
      </div>

      {/* SUB-NAVIGATION TABS & ACTION BUTTON */}
      <div className="flex items-center justify-between border-b border-zinc-200 dark:border-zinc-800 gap-2 pb-px flex-wrap">
        <div className="flex gap-1 overflow-x-auto">
          <button
            onClick={() => setSubTab('tre_oficial')}
            className={`px-5 py-3 text-xs font-black uppercase tracking-wider border-b-2 transition-all shrink-0 flex items-center gap-2 ${
              subTab === 'tre_oficial'
                ? 'border-blue-600 text-blue-600 dark:text-blue-500 bg-blue-600/5'
                : 'border-transparent text-zinc-500 hover:text-zinc-800 dark:hover:text-zinc-200 hover:bg-zinc-100/50 dark:hover:bg-zinc-900/50'
            }`}
          >
            <Database className="w-4 h-4 text-blue-600" />
            Dados Oficiais do TRE
          </button>
          <button
            onClick={() => setSubTab('cruzamento')}
            className={`px-5 py-3 text-xs font-black uppercase tracking-wider border-b-2 transition-all shrink-0 flex items-center gap-2 ${
              subTab === 'cruzamento'
                ? 'border-blue-600 text-blue-600 dark:text-blue-500 bg-blue-600/5'
                : 'border-transparent text-zinc-500 hover:text-zinc-800 dark:hover:text-zinc-200 hover:bg-zinc-100/50 dark:hover:bg-zinc-900/50'
            }`}
          >
            <Target className="w-4 h-4 text-blue-600" />
            Cruzamento de Dados & Estratégia
          </button>
        </div>

        {isCoordinator && canEditTreData && (
          <div className="flex items-center gap-2 flex-wrap my-1">
            <a
              href="https://drive.google.com/drive/folders/1D3mvu08C-fxvk9CQrAynEmeol7mMr32s?usp=sharing"
              target="_blank"
              rel="noopener noreferrer"
              className="flex items-center gap-2 px-3.5 py-1.5 bg-blue-600 hover:bg-blue-500 text-white rounded text-xs font-bold transition-all shadow-sm active:scale-95 cursor-pointer uppercase tracking-tight"
              title="Acessar pasta no Google Drive com as planilhas oficiais preenchidas por Estado"
            >
              <Download className="w-4 h-4 text-blue-100" />
              <span>BAIXE A PLANILHA PREENCHIDA DO SEU ESTADO</span>
              <ExternalLink className="w-3 h-3 text-blue-200 ml-0.5" />
            </a>

            {votingLocations.length > 0 && (
              <>
                <button
                  onClick={triggerFilePicker}
                  className="flex items-center gap-2 px-3.5 py-1.5 bg-zinc-900 hover:bg-zinc-800 dark:bg-zinc-800 dark:hover:bg-zinc-700 text-white rounded text-xs font-bold transition-all shadow-sm active:scale-95 cursor-pointer"
                  title="Substituir ou importar nova planilha de dados oficiais do TRE"
                >
                  <UploadCloud className="w-4 h-4 text-blue-400" />
                  <span>Substituir Planilha TRE</span>
                </button>

                <button
                  onClick={clearData}
                  className="flex items-center gap-2 px-3.5 py-1.5 bg-red-600 hover:bg-red-500 text-white rounded text-xs font-bold transition-all shadow-sm active:scale-95 cursor-pointer uppercase tracking-tight"
                  title="Zerar e apagar permanentemente os dados desatualizados do TRE"
                >
                  <Trash2 className="w-4 h-4 text-white" />
                  <span>ZERAR DADOS TRE</span>
                </button>
              </>
            )}

            {!isSupabaseConfigured() && (
              <button
                onClick={() => setIsSupabaseModalOpen(true)}
                className="flex items-center gap-2 px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded text-xs font-bold transition-all shadow-sm active:scale-95 cursor-pointer"
                title="Configurar banco relacional Supabase PostgreSQL (Apenas Administrador)"
              >
                <Database className="w-4 h-4 text-emerald-200" />
                <span>Conectar Supabase (Admin)</span>
              </button>
            )}
          </div>
        )}
      </div>

      {/* FEEDBACK BANNERS */}
      {successMsg && (
        <div className="flex items-center gap-3 bg-emerald-50 dark:bg-emerald-950/40 border border-emerald-200 dark:border-emerald-800/60 text-emerald-800 dark:text-emerald-300 p-4 rounded-sm shadow-sm">
          <CheckCircle2 className="w-5 h-5 text-emerald-600 dark:text-emerald-400 shrink-0" />
          <span className="text-xs font-bold uppercase tracking-wide">{successMsg}</span>
        </div>
      )}

      {errorMsg && (
        <div className="flex items-center gap-3 bg-rose-50 dark:bg-rose-950/40 border border-rose-200 dark:border-rose-800/60 text-rose-800 dark:text-rose-300 p-4 rounded-sm shadow-sm">
          <AlertCircle className="w-5 h-5 text-rose-600 dark:text-rose-400 shrink-0" />
          <span className="text-xs font-bold uppercase tracking-wide">{errorMsg}</span>
        </div>
      )}

      {/* MODAL: EXCEL IMPORT FOR COORDENADOR GERAL */}
      {isImportModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4">
          <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-lg shadow-2xl max-w-xl w-full p-6 text-zinc-900 dark:text-white relative animate-in fade-in zoom-in duration-150">
            <button 
              onClick={() => setIsImportModalOpen(false)}
              className="absolute top-4 right-4 p-1.5 text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200 rounded-lg hover:bg-zinc-100 dark:hover:bg-zinc-800 transition-colors"
            >
              <X className="w-5 h-5" />
            </button>

            <div className="flex items-center gap-3 mb-4">
              <div className="p-2.5 bg-blue-50 dark:bg-blue-950/50 text-blue-600 dark:text-blue-400 rounded-lg">
                <Database className="w-5 h-5" />
              </div>
              <div>
                <h3 className="text-base font-bold text-zinc-900 dark:text-white">
                  Importador de Dados Oficiais do TRE
                </h3>
                <p className="text-xs text-zinc-500 dark:text-zinc-400">
                  Gerencie a planilha de dados eleitorais oficiais de votação.
                </p>
              </div>
            </div>

            <div className="my-4 border-t border-b border-zinc-100 dark:border-zinc-800 py-4 space-y-4">
              {/* Actions bar */}
              <div className="flex items-center justify-between gap-2 flex-wrap">
                <button
                  onClick={downloadTemplate}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-zinc-100 dark:bg-zinc-800 hover:bg-zinc-200 dark:hover:bg-zinc-700 text-zinc-800 dark:text-zinc-200 rounded text-xs font-medium transition-colors"
                >
                  <Download className="w-3.5 h-3.5 text-blue-600 dark:text-blue-400" />
                  <span>Baixar Planilha Modelo</span>
                </button>

                <div className="flex items-center gap-2">
                  <button
                    onClick={() => {
                      clearData();
                      setIsImportModalOpen(false);
                    }}
                    className="flex items-center gap-1.5 px-3 py-1.5 bg-red-50 dark:bg-red-950/30 text-red-600 dark:text-red-400 hover:bg-red-100 dark:hover:bg-red-900/40 rounded text-xs font-medium transition-colors"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                    <span>Zerar Banco TRE</span>
                  </button>
                </div>
              </div>

              {/* Drive Folder Link Option */}
              <div className="p-3 bg-blue-50/70 dark:bg-blue-950/30 border border-blue-200 dark:border-blue-900/50 rounded-lg flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3">
                <div className="flex items-center gap-2.5">
                  <div className="p-2 bg-blue-600 text-white rounded-md shrink-0">
                    <Database className="w-4 h-4" />
                  </div>
                  <div>
                    <h4 className="text-xs font-bold text-zinc-900 dark:text-white uppercase tracking-wider">
                      Repositório Oficial TSE no Google Drive
                    </h4>
                    <p className="text-[11px] text-zinc-600 dark:text-zinc-400">
                      Acesse as tabelas oficiais por estado prontas no Google Drive ou cole o link do seu estado.
                    </p>
                  </div>
                </div>

                <a 
                  href="https://drive.google.com/drive/folders/1D3mvu08C-fxvk9CQrAynEmeol7mMr32s?usp=sharing" 
                  target="_blank" 
                  rel="noopener noreferrer"
                  className="px-3 py-1.5 bg-blue-600 hover:bg-blue-500 text-white rounded text-xs font-bold transition-colors shrink-0 flex items-center gap-1.5 shadow-sm"
                >
                  <span>Acessar Pasta no Drive</span>
                  <ExternalLink className="w-3 h-3" />
                </a>
              </div>

              {/* Dropzone */}
              <div 
                onDragEnter={handleDrag}
                onDragOver={handleDrag}
                onDragLeave={handleDrag}
                onDrop={(e) => {
                  handleDrop(e);
                  setIsImportModalOpen(false);
                }}
                onClick={triggerFilePicker}
                className={`border-2 border-dashed rounded-lg p-6 text-center transition-all cursor-pointer ${
                  dragActive 
                    ? "border-blue-500 bg-blue-50/50 dark:bg-blue-950/30" 
                    : "border-zinc-300 dark:border-zinc-700 hover:border-blue-500 bg-zinc-50 dark:bg-zinc-800/50"
                }`}
              >
                <div className="flex flex-col items-center justify-center gap-2">
                  <UploadCloud className="w-10 h-10 text-blue-600 dark:text-blue-400 mb-1" />
                  <span className="text-sm font-bold text-zinc-800 dark:text-zinc-100">
                    Arraste a planilha do TSE aqui ou clique para selecionar
                  </span>
                  <span className="text-xs text-zinc-500 dark:text-zinc-400">
                    Suporta arquivos .xlsx, .xls e .csv baixados do Drive ou do site oficial do TSE
                  </span>
                </div>
              </div>
            </div>

            <div className="flex justify-end">
              <button
                onClick={() => setIsImportModalOpen(false)}
                className="px-4 py-2 bg-zinc-100 dark:bg-zinc-800 hover:bg-zinc-200 dark:hover:bg-zinc-700 text-zinc-800 dark:text-zinc-200 text-xs font-medium rounded transition-colors"
              >
                Fechar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MAIN SCREEN HANDLING - EMPTY STATE VS GRAPH COMPONENT */}
      {votingLocations.length === 0 ? (
        <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-sm p-12 text-center max-w-2xl mx-auto shadow-sm">
          <FileSpreadsheet className="w-16 h-16 text-blue-600 mx-auto mb-4 animate-pulse" />
          <h2 className="text-lg font-black text-zinc-900 dark:text-white uppercase tracking-wide">
            Nenhum Dado Eleitoral Carregado
          </h2>
          <p className="text-xs text-zinc-500 dark:text-zinc-400 mt-2 max-w-md mx-auto leading-relaxed">
            Iremos trabalhar exclusivamente com dados oficiais do TRE. Os dados simulados de teste foram zerados de acordo com as diretrizes de segurança da campanha.
          </p>

          {isCoordinator && canEditTreData ? (
            <div className="mt-8 space-y-4">
              <p className="text-[11px] font-black text-blue-600 dark:text-blue-600 uppercase tracking-widest">
                Você possui privilégios de Coordenador Geral para enviar dados do TRE.
              </p>
              <div className="flex flex-col sm:flex-row items-center justify-center gap-3">
                <button
                  onClick={triggerFilePicker}
                  className="flex items-center justify-center gap-2 px-6 py-3 bg-blue-600 text-white hover:bg-blue-500 font-black text-xs uppercase tracking-wider rounded transition-all shadow-md hover:shadow-lg active:scale-95 cursor-pointer w-full sm:w-auto"
                >
                  <UploadCloud className="w-5 h-5 text-white" />
                  <span>Carregar Planilha Oficial (.xlsx / .csv)</span>
                </button>

                <button
                  onClick={downloadTemplate}
                  className="flex items-center justify-center gap-2 px-5 py-3 bg-zinc-900 text-white hover:bg-zinc-800 font-black text-xs uppercase tracking-wider rounded transition-all shadow-sm border border-zinc-800 w-full sm:w-auto"
                >
                  <Download className="w-4 h-4 text-blue-400" />
                  <span>Baixar Modelo Excel</span>
                </button>
              </div>
            </div>
          ) : (
            <div className="mt-6 p-4 bg-zinc-50 dark:bg-zinc-950 border border-zinc-200 dark:border-zinc-800 rounded-sm space-y-1">
              <p className="text-[10px] text-zinc-500 dark:text-zinc-400 font-black uppercase tracking-widest">
                Aguardando carregamento da planilha oficial de locais de votação do TRE pelo Coordenador Geral no painel administrativo.
              </p>
              <p className="text-[9px] text-amber-600 dark:text-amber-400 font-bold uppercase tracking-wider">
                Como Coordenador Regional, seu acesso é para visualização e análise de dados assim que a base for carregada.
              </p>
            </div>
          )}
        </div>
      ) : (
        <>
          {subTab === 'tre_oficial' ? (
            <>
              {/* FILTER CONTROLS */}
          <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-sm p-4 shadow-sm">
            <div className="flex items-center gap-2 text-zinc-900 dark:text-white font-black text-sm uppercase tracking-wider border-b border-zinc-100 dark:border-zinc-800 pb-2 mb-4">
              <Filter className="w-4 h-4 text-blue-600" />
              <span>Segmentação e Filtros Interativos</span>
            </div>
            
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              {/* Municipio Filter */}
              <div className="space-y-1">
                <label className="text-[10px] font-black uppercase tracking-wider text-zinc-500 dark:text-zinc-400">
                  Município
                </label>
                <select
                  value={selectedMunicipio}
                  onChange={(e) => handleMunicipioChange(e.target.value)}
                  className="w-full text-xs font-black bg-zinc-50 dark:bg-zinc-950 border border-zinc-200 dark:border-zinc-800 rounded-sm p-2 focus:ring-1 focus:ring-blue-600 focus:outline-none transition-all uppercase"
                >
                  {municipiosList.map(mun => (
                    <option key={mun} value={mun}>{mun === 'Todos' ? '✦ Todos os Municípios' : mun}</option>
                  ))}
                </select>
              </div>

              {/* Zona Eleitoral Filter */}
              <div className="space-y-1">
                <label className="text-[10px] font-black uppercase tracking-wider text-zinc-500 dark:text-zinc-400">
                  Zona Eleitoral
                </label>
                <select
                  value={selectedZona}
                  onChange={(e) => handleZonaChange(e.target.value)}
                  className="w-full text-xs font-black bg-zinc-50 dark:bg-zinc-950 border border-zinc-200 dark:border-zinc-800 rounded-sm p-2 focus:ring-1 focus:ring-blue-600 focus:outline-none transition-all uppercase"
                >
                  {zonasList.map(z => (
                    <option key={z} value={z}>{z === 'Todos' ? '✦ Todas as Zonas' : z}</option>
                  ))}
                </select>
              </div>

              {/* Bairro Filter */}
              <div className="space-y-1">
                <label className="text-[10px] font-black uppercase tracking-wider text-zinc-500 dark:text-zinc-400">
                  Bairro
                </label>
                <select
                  value={selectedBairro}
                  onChange={(e) => handleBairroChange(e.target.value)}
                  className="w-full text-xs font-black bg-zinc-50 dark:bg-zinc-950 border border-zinc-200 dark:border-zinc-800 rounded-sm p-2 focus:ring-1 focus:ring-blue-600 focus:outline-none transition-all uppercase"
                >
                  {bairrosList.map(b => (
                    <option key={b} value={b}>{b === 'Todos' ? '✦ Todos os Bairros' : b}</option>
                  ))}
                </select>
              </div>

              {/* Local Votacao Filter */}
              <div className="space-y-1">
                <label className="text-[10px] font-black uppercase tracking-wider text-zinc-500 dark:text-zinc-400">
                  Local de Votação
                </label>
                <select
                  value={selectedLocal}
                  onChange={(e) => setSelectedLocal(e.target.value)}
                  className="w-full text-xs font-black bg-zinc-50 dark:bg-zinc-950 border border-zinc-200 dark:border-zinc-800 rounded-sm p-2 focus:ring-1 focus:ring-blue-600 focus:outline-none transition-all uppercase"
                >
                  {locaisList.map(l => (
                    <option key={l} value={l}>
                      {l === 'Todos' ? '✦ Todos os Locais' : l.length > 35 ? l.substring(0, 32) + '...' : l}
                    </option>
                  ))}
                </select>
              </div>
            </div>
          </div>

          {/* PAINEL GERAL (RESUMO EXECUTIVO) */}
          <div className="space-y-3">
            <h2 className="text-sm font-black text-zinc-400 uppercase tracking-widest pl-1">
              Painel Geral (Resumo Executivo)
            </h2>
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
              
              {/* Total Eleitores */}
              <div className="bg-gradient-to-br from-zinc-900 to-black text-white rounded-sm p-4 shadow-sm border border-zinc-800 flex flex-col justify-between relative overflow-hidden group border-t-4 border-t-blue-600">
                <div className="absolute right-2 top-2 text-white/5 group-hover:scale-110 transition-transform">
                  <TrendingUp className="w-16 h-16 text-blue-600" />
                </div>
                <p className="text-[10px] font-bold uppercase tracking-wider text-zinc-400">
                  Eleitores Aptos Geral
                </p>
                <div className="mt-3">
                  <h3 className="text-2xl md:text-3xl font-black tracking-tight leading-none text-blue-600">
                    {totalStateEleitores.toLocaleString()}
                  </h3>
                  <p className="text-[9px] text-zinc-500 mt-1 uppercase font-semibold">
                    Consolidado Carregado
                  </p>
                </div>
              </div>

              {/* Total Municipios */}
              <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-sm p-4 shadow-sm flex flex-col justify-between relative overflow-hidden group border-t-4 border-t-zinc-300 dark:border-t-zinc-700">
                <div className="absolute right-2 top-2 text-zinc-100 dark:text-zinc-800 group-hover:scale-110 transition-transform">
                  <Building2 className="w-14 h-14" />
                </div>
                <p className="text-[10px] font-black uppercase tracking-wider text-zinc-400">
                  Municípios Analisados
                </p>
                <div className="mt-3">
                  <h3 className="text-2xl md:text-3xl font-black tracking-tight text-zinc-900 dark:text-white leading-none">
                    {totalStateMunicipios}
                  </h3>
                  <p className="text-[9px] text-zinc-400 mt-1 uppercase font-semibold">
                    Cidades Mapeadas
                  </p>
                </div>
              </div>

              {/* Total Locais */}
              <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-sm p-4 shadow-sm flex flex-col justify-between relative overflow-hidden group border-t-4 border-t-zinc-300 dark:border-t-zinc-700">
                <div className="absolute right-2 top-2 text-zinc-100 dark:text-zinc-800 group-hover:scale-110 transition-transform">
                  <MapIcon className="w-14 h-14" />
                </div>
                <p className="text-[10px] font-black uppercase tracking-wider text-zinc-400">
                  Locais de Votação
                </p>
                <div className="mt-3">
                  <h3 className="text-2xl md:text-3xl font-black tracking-tight text-zinc-900 dark:text-white leading-none">
                    {totalStateLocais}
                  </h3>
                  <p className="text-[9px] text-zinc-400 mt-1 uppercase font-semibold">
                    Pontos Estatísticos
                  </p>
                </div>
              </div>

              {/* Total Secoes */}
              <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-sm p-4 shadow-sm flex flex-col justify-between relative overflow-hidden group border-t-4 border-t-zinc-300 dark:border-t-zinc-700">
                <div className="absolute right-2 top-2 text-zinc-100 dark:text-zinc-800 group-hover:scale-110 transition-transform">
                  <Hash className="w-14 h-14" />
                </div>
                <p className="text-[10px] font-black uppercase tracking-wider text-zinc-400">
                  Seções Eleitorais
                </p>
                <div className="mt-3">
                  <h3 className="text-2xl md:text-3xl font-black tracking-tight text-zinc-900 dark:text-white leading-none">
                    {totalStateSecoes}
                  </h3>
                  <p className="text-[9px] text-zinc-400 mt-1 uppercase font-semibold">
                    Urnas Registradas
                  </p>
                </div>
              </div>

            </div>
          </div>

          {/* FILTER METRICS PANEL */}
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            
            {/* CURRENT FILTERS SUMMARY */}
            <div className="lg:col-span-1 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded p-4 shadow-sm flex flex-col justify-between">
              <div>
                <div className="flex items-center gap-2 text-zinc-900 dark:text-white font-bold text-xs uppercase tracking-wider border-b border-zinc-100 dark:border-zinc-800 pb-2 mb-3">
                  <MapPin className="w-4 h-4 text-blue-600" />
                  <span>Escopo Atual Filtrado</span>
                </div>
                
                <div className="space-y-3">
                  <div className="flex justify-between items-center text-xs">
                    <span className="text-zinc-400 uppercase font-black tracking-wider text-[9px]">Município:</span>
                    <span className="font-bold text-zinc-800 dark:text-zinc-200 bg-zinc-100 dark:bg-zinc-800 px-2 py-0.5 rounded">
                      {selectedMunicipio}
                    </span>
                  </div>
                  <div className="flex justify-between items-center text-xs">
                    <span className="text-zinc-400 uppercase font-black tracking-wider text-[9px]">Zona Eleitoral:</span>
                    <span className="font-bold text-zinc-800 dark:text-zinc-200 bg-zinc-100 dark:bg-zinc-800 px-2 py-0.5 rounded">
                      {selectedZona}
                    </span>
                  </div>
                  <div className="flex justify-between items-center text-xs">
                    <span className="text-zinc-400 uppercase font-black tracking-wider text-[9px]">Bairro:</span>
                    <span className="font-bold text-zinc-800 dark:text-zinc-200 bg-zinc-100 dark:bg-zinc-800 px-2 py-0.5 rounded text-right max-w-[200px] truncate" title={selectedBairro}>
                      {selectedBairro}
                    </span>
                  </div>
                  <div className="flex justify-between items-center text-xs">
                    <span className="text-zinc-400 uppercase font-black tracking-wider text-[9px]">Local:</span>
                    <span className="font-bold text-zinc-800 dark:text-zinc-200 bg-zinc-100 dark:bg-zinc-800 px-2 py-0.5 rounded text-right max-w-[200px] truncate" title={selectedLocal}>
                      {selectedLocal}
                    </span>
                  </div>
                </div>
              </div>

              <div className="border-t border-zinc-100 dark:border-zinc-800 pt-4 mt-4 grid grid-cols-2 gap-2 text-center">
                <div className="bg-zinc-50 dark:bg-zinc-950 p-2 rounded border border-zinc-100 dark:border-zinc-850">
                  <p className="text-[9px] font-black uppercase text-zinc-400 tracking-wider">Eleitores Escopo</p>
                  <h4 className="text-lg font-black text-blue-600 dark:text-blue-600 mt-1">
                    {totalScopeEleitores.toLocaleString()}
                  </h4>
                  <p className="text-[8px] text-zinc-400">
                    {totalStateEleitores > 0 ? ((totalScopeEleitores / totalStateEleitores) * 100).toFixed(2) : 0}% do Total
                  </p>
                </div>
                <div className="bg-zinc-50 dark:bg-zinc-950 p-2 rounded border border-zinc-100 dark:border-zinc-850">
                  <p className="text-[9px] font-black uppercase text-zinc-400 tracking-wider">Seções Escopo</p>
                  <h4 className="text-lg font-black text-amber-500 dark:text-amber-400 mt-1">
                    {totalScopeSecoes}
                  </h4>
                  <p className="text-[8px] text-zinc-400">
                    {totalStateSecoes > 0 ? ((totalScopeSecoes / totalStateSecoes) * 100).toFixed(2) : 0}% das Urnas
                  </p>
                </div>
              </div>
            </div>

            {/* INDICADORES (KPIS) */}
            <div className="lg:col-span-2 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded p-4 shadow-sm">
              <div className="flex items-center gap-2 text-zinc-900 dark:text-white font-bold text-xs uppercase tracking-wider border-b border-zinc-100 dark:border-zinc-800 pb-2 mb-4">
                <Award className="w-4 h-4 text-blue-600" />
                <span>Indicadores de Desempenho e Extremos (KPIs)</span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                
                {/* Maior Local de Votação */}
                <div className="bg-zinc-50 dark:bg-zinc-950 border border-zinc-100 dark:border-zinc-850 p-3 rounded flex flex-col justify-between">
                  <div>
                    <span className="text-[9px] font-black uppercase tracking-widest text-zinc-400">
                      Maior Local de Votação {selectedMunicipio !== 'Todos' ? `de ${selectedMunicipio}` : ''}
                    </span>
                    <h4 className="text-xs font-bold text-zinc-900 dark:text-white mt-1.5 truncate" title={kpiMetrics.maiorLocal.local}>
                      {kpiMetrics.maiorLocal.local}
                    </h4>
                  </div>
                  <div className="flex justify-between items-baseline mt-2">
                    <span className="text-lg font-black text-blue-600 dark:text-blue-600">
                      {kpiMetrics.maiorLocal.eleitores.toLocaleString()} eleitores
                    </span>
                    <span className="text-[10px] font-bold text-blue-600 bg-blue-600/10 px-1.5 py-0.5 rounded-sm">
                      {kpiMetrics.maiorLocal.percentMun.toFixed(2)}%
                    </span>
                  </div>
                </div>

                {/* Menor Local de Votação */}
                <div className="bg-zinc-50 dark:bg-zinc-950 border border-zinc-100 dark:border-zinc-850 p-3 rounded flex flex-col justify-between">
                  <div>
                    <span className="text-[9px] font-black uppercase tracking-widest text-zinc-400">
                      Menor Local de Votação {selectedMunicipio !== 'Todos' ? `de ${selectedMunicipio}` : ''}
                    </span>
                    <h4 className="text-xs font-bold text-zinc-900 dark:text-white mt-1.5 truncate" title={kpiMetrics.menorLocal.local}>
                      {kpiMetrics.menorLocal.local}
                    </h4>
                  </div>
                  <div className="flex justify-between items-baseline mt-2">
                    <span className="text-lg font-black text-zinc-700 dark:text-zinc-300">
                      {kpiMetrics.menorLocal.eleitores.toLocaleString()} eleitores
                    </span>
                    <span className="text-[10px] font-bold text-zinc-500 bg-zinc-100 dark:bg-zinc-800 px-1.5 py-0.5 rounded">
                      {kpiMetrics.menorLocal.percentMun.toFixed(2)}%
                    </span>
                  </div>
                </div>

                {/* Média de Eleitores por Local */}
                <div className="bg-zinc-50 dark:bg-zinc-950 border border-zinc-100 dark:border-zinc-850 p-3 rounded flex flex-col justify-between">
                  <div>
                    <span className="text-[9px] font-black uppercase tracking-widest text-zinc-400">
                      Média de Eleitores por Local
                    </span>
                    <h4 className="text-xs font-bold text-zinc-500 mt-1">
                      Mapeado no escopo selecionado
                    </h4>
                  </div>
                  <div className="flex justify-between items-baseline mt-2">
                    <span className="text-lg font-black text-amber-500 dark:text-amber-400">
                      {kpiMetrics.mediaEleitores.toLocaleString()}
                    </span>
                    <span className="text-[9px] text-zinc-400 uppercase font-semibold">
                      Média / Local
                    </span>
                  </div>
                </div>

                {/* Percentual Acumulado dos 5 Maiores */}
                <div className="bg-zinc-50 dark:bg-zinc-950 border border-zinc-100 dark:border-zinc-850 p-3 rounded flex flex-col justify-between">
                  <div>
                    <span className="text-[9px] font-black uppercase tracking-widest text-zinc-400">
                      Concentração nos 5 Maiores Locais
                    </span>
                    <h4 className="text-xs font-bold text-zinc-500 mt-1">
                      Representatividade somada top 5
                    </h4>
                  </div>
                  <div className="flex justify-between items-baseline mt-2">
                    <span className="text-lg font-black text-blue-600 dark:text-blue-600">
                      {kpiMetrics.top5AccumulatedPercent.toFixed(1)}%
                    </span>
                    <span className="text-[9px] text-zinc-400 uppercase font-semibold">
                      do Escopo
                    </span>
                  </div>
                </div>

              </div>
            </div>

          </div>

          {/* MUNICIPIO LEVEL DRILL DOWN (PAINEL POR MUNICIPIO) */}
          {municipioPanelData && (
            <div className="bg-gradient-to-r from-zinc-950 via-zinc-900 to-zinc-950 text-white rounded-sm p-5 shadow-sm border border-zinc-800">
              <div className="flex items-center gap-2 text-blue-600 font-bold text-xs uppercase tracking-widest mb-2">
                <Building2 className="w-4 h-4 text-blue-600" />
                <span>Painel Detalhado por Município</span>
              </div>
              
              <div className="grid grid-cols-1 md:grid-cols-5 gap-4 items-center">
                <div className="md:col-span-2">
                  <h3 className="text-2xl font-black uppercase text-white tracking-tight text-blue-600">
                    {municipioPanelData.nome}
                  </h3>
                  <p className="text-xs text-zinc-400 mt-1">
                    Análise específica de representatividade e capilaridade urbana.
                  </p>
                </div>

                <div className="grid grid-cols-3 md:col-span-3 gap-2">
                  <div className="bg-white/10 p-3 rounded text-center border border-white/5">
                    <p className="text-[8px] font-bold text-blue-200 uppercase tracking-widest">Eleitores Aptos</p>
                    <h4 className="text-lg font-black text-white mt-1">
                      {municipioPanelData.eleitores.toLocaleString()}
                    </h4>
                  </div>
                  <div className="bg-white/10 p-3 rounded text-center border border-white/5">
                    <p className="text-[8px] font-bold text-blue-200 uppercase tracking-widest">Locais</p>
                    <h4 className="text-lg font-black text-white mt-1">
                      {municipioPanelData.locais}
                    </h4>
                  </div>
                  <div className="bg-white/10 p-3 rounded-sm text-center border border-white/5">
                    <p className="text-[8px] font-black text-zinc-300 uppercase tracking-widest">% do Estado</p>
                    <h4 className="text-lg font-black text-blue-600 mt-1">
                      {municipioPanelData.representatividade.toFixed(2)}%
                    </h4>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* CHARTS CONTAINER */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            
            {/* CHART 1: LOCAL VOTING RANKING (HORIZONTAL BARS) */}
            <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded p-4 shadow-sm flex flex-col justify-between">
              <div>
                <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-2 mb-4">
                  <div className="text-zinc-900 dark:text-white font-bold text-xs uppercase tracking-wider">
                    Gráfico 1: Concentração por Local de Votação (Top 10)
                  </div>
                  <span className="text-[8px] font-black text-blue-600 bg-blue-600/10 px-1.5 py-0.5 rounded-sm uppercase">
                    Eleitores Aptos
                  </span>
                </div>
                <p className="text-[10px] text-zinc-500 dark:text-zinc-400 mb-4">
                  Eixo X = Quantidade de Eleitores Aptos | Eixo Y = Local de Votação
                </p>
              </div>

              <div className="h-80 w-full text-xs font-medium">
                {barChartData.length === 0 ? (
                  <div className="h-full flex items-center justify-center text-zinc-400">Nenhum dado filtrado</div>
                ) : (
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart
                      data={barChartData}
                      layout="vertical"
                      margin={{ top: 5, right: 10, left: 15, bottom: 5 }}
                    >
                      <CartesianGrid strokeDasharray="3 3" stroke="#e4e4e7" opacity={0.3} />
                      <XAxis type="number" stroke="#888888" fontSize={9} />
                      <YAxis 
                        type="category" 
                        dataKey="name" 
                        stroke="#888888" 
                        fontSize={8} 
                        width={110}
                      />
                      <Tooltip 
                        contentStyle={{ backgroundColor: '#18181b', color: '#fff', fontSize: 10, borderRadius: 4 }}
                        formatter={(value: any) => [`${value.toLocaleString()} Eleitores`, 'Total']}
                      />
                      <Bar dataKey="eleitores" radius={[0, 4, 4, 0]}>
                        {barChartData.map((entry, index) => (
                          <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                        ))}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                )}
              </div>
            </div>

            {/* CHART 2: PIE/DONUT CHART (PARTICIPATION OF LOCALS) */}
            <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded p-4 shadow-sm flex flex-col justify-between">
              <div>
                <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-2 mb-4">
                  <div className="text-zinc-900 dark:text-white font-bold text-xs uppercase tracking-wider">
                    Gráfico 2: Participação de cada Local de Votação
                  </div>
                  <span className="text-[8px] font-black text-blue-600 bg-blue-600/10 px-1.5 py-0.5 rounded-sm uppercase">
                    Participação %
                  </span>
                </div>
                <p className="text-[10px] text-zinc-500 dark:text-zinc-400 mb-4">
                  Divisão percentual dos maiores locais e aglutinação de remanescentes em &quot;Outros&quot;
                </p>
              </div>

              <div className="h-80 w-full text-xs">
                {pieChartData.length === 0 ? (
                  <div className="h-full flex items-center justify-center text-zinc-400">Nenhum dado filtrado</div>
                ) : (
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie
                        data={pieChartData}
                        cx="50%"
                        cy="45%"
                        innerRadius={60}
                        outerRadius={85}
                        paddingAngle={3}
                        dataKey="value"
                      >
                        {pieChartData.map((entry, index) => (
                          <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                        ))}
                      </Pie>
                      <Tooltip 
                        contentStyle={{ backgroundColor: '#18181b', color: '#fff', fontSize: 10, borderRadius: 4 }}
                        formatter={(value: any) => [`${value.toLocaleString()} Eleitores`, 'Aptos']}
                      />
                      <Legend 
                        layout="horizontal" 
                        verticalAlign="bottom" 
                        align="center"
                        iconSize={8}
                        wrapperStyle={{ fontSize: 9 }}
                      />
                    </PieChart>
                  </ResponsiveContainer>
                )}
              </div>
            </div>

            {/* CHART 3: STATEWIDE MUNICIPAL RANKING */}
            <div className="lg:col-span-2 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded p-4 shadow-sm">
              <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-2 mb-4">
                <div className="text-zinc-900 dark:text-white font-bold text-xs uppercase tracking-wider">
                  Gráfico 3: Ranking Estadual Geral dos Municípios
                </div>
                <span className="text-[8px] font-black text-blue-600 bg-blue-600/10 px-1.5 py-0.5 rounded-sm uppercase">
                  Eleitorado do maior para o menor
                </span>
              </div>

              <div className="h-64 w-full text-xs font-medium">
                {muniRankingChartData.length === 0 ? (
                  <div className="h-full flex items-center justify-center text-zinc-400">Nenhum dado carregado</div>
                ) : (
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart
                      data={muniRankingChartData}
                      margin={{ top: 10, right: 10, left: 10, bottom: 20 }}
                    >
                      <CartesianGrid strokeDasharray="3 3" stroke="#e4e4e7" opacity={0.3} />
                      <XAxis 
                        dataKey="name" 
                        stroke="#888888" 
                        fontSize={8} 
                        angle={-45} 
                        textAnchor="end"
                        height={50}
                      />
                      <YAxis stroke="#888888" fontSize={9} />
                      <Tooltip 
                        contentStyle={{ backgroundColor: '#18181b', color: '#fff', fontSize: 10, borderRadius: 4 }}
                        formatter={(value: any) => [`${value.toLocaleString()} Eleitores Aptos`, 'Eleitores']}
                      />
                      <Bar dataKey="eleitores" fill="#0578d3" radius={[4, 4, 0, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                )}
              </div>
            </div>

          </div>

          {/* DETAILED DYNAMIC TABLE (TABELA DETALHADA NA PARTE INFERIOR) */}
          <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded shadow-sm overflow-hidden">
            
            {/* Table Header Controls */}
            <div className="p-4 border-b border-zinc-200 dark:border-zinc-800 flex flex-col sm:flex-row sm:items-center justify-between gap-2 bg-zinc-50 dark:bg-zinc-900">
              <div>
                <h3 className="text-sm font-black text-zinc-900 dark:text-white uppercase tracking-wider">
                  Painel de Locais de Votação (Tabela Dinâmica)
                </h3>
                <p className="text-[11px] text-zinc-400 mt-0.5">
                  Exibindo {processedTableData.length} locais filtrados em {totalPages} {totalPages === 1 ? 'página' : 'páginas'} (10 por página).
                </p>
              </div>
              <div className="text-[10px] font-bold text-zinc-500 uppercase bg-zinc-200 dark:bg-zinc-800 px-2.5 py-1 rounded">
                Fórmulas de BI Aplicadas
              </div>
            </div>

            {/* Dynamic Table */}
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs border-collapse">
                <thead>
                  <tr className="bg-zinc-100 dark:bg-zinc-850 text-zinc-500 uppercase font-black text-[9px] border-b border-zinc-200 dark:border-zinc-800 tracking-wider">
                    {TSE_COLUMNS.map((col) => (
                      <TseHeaderCell 
                        key={col.key}
                        variable={col.variable}
                        description={col.description}
                        align={col.key === 'qtEleitorSecao' || col.key === 'cdTipoSecaoAgregada' || col.key === 'nrSecaoPrincipal' || col.key === 'nrLocalVotacao' ? 'center' : 'left'}
                        isSortable={col.key === 'nmLocalVotacao' || col.key === 'qtEleitorSecao'}
                        onSort={() => {
                          if (col.key === 'nmLocalVotacao') toggleSort('local');
                          if (col.key === 'qtEleitorSecao') toggleSort('eleitores');
                        }}
                      />
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-zinc-200 dark:divide-zinc-800 font-medium">
                  {processedTableData.length === 0 ? (
                    <tr>
                      <td colSpan={TSE_COLUMNS.length} className="py-12 text-center text-zinc-400">
                        Nenhum registro correspondente aos filtros selecionados.
                      </td>
                    </tr>
                  ) : (
                    paginatedTableData.map((row, index) => (
                      <tr 
                        key={`${row.nmMunicipio || row.municipio}-${row.nrSecao || row.secoes}-${index}`}
                        className="hover:bg-blue-50/30 dark:hover:bg-zinc-800/60 transition-colors text-[11px]"
                      >
                        {/* 1. NM MUNICIPIO */}
                        <td className="py-3 px-3 font-bold text-zinc-900 dark:text-white uppercase whitespace-nowrap">
                          {row.nmMunicipio || row.municipio}
                        </td>

                        {/* 2. NR ZONA */}
                        <td className="py-3 px-3 text-zinc-700 dark:text-zinc-300 font-semibold whitespace-nowrap">
                          {row.nrZona || row.zona}
                        </td>

                        {/* 3. NR SECAO */}
                        <td className="py-3 px-3 font-mono text-blue-600 dark:text-blue-400 font-bold whitespace-nowrap">
                          {row.nrSecao || row.secoes}
                        </td>

                        {/* 4. CD TIPO SECAO AGREGADA */}
                        <td className="py-3 px-3 text-center text-zinc-500 font-mono">
                          {row.cdTipoSecaoAgregada ?? -1}
                        </td>

                        {/* 5. DS_TIPO_SECAO_AGREGADA */}
                        <td className="py-3 px-3 text-zinc-600 dark:text-zinc-400 whitespace-nowrap">
                          <span className={`px-2 py-0.5 rounded text-[9px] font-bold uppercase ${
                            row.dsTipoSecaoAgregada === 'Agregada' 
                              ? 'bg-amber-100 dark:bg-amber-950/60 text-amber-700 dark:text-amber-400' 
                              : row.dsTipoSecaoAgregada === 'Principal'
                              ? 'bg-emerald-100 dark:bg-emerald-950/60 text-emerald-700 dark:text-emerald-400'
                              : 'bg-zinc-100 dark:bg-zinc-800 text-zinc-500'
                          }`}>
                            {row.dsTipoSecaoAgregada || '#NULO'}
                          </span>
                        </td>

                        {/* 6. NR SECAO PRINCIPAL */}
                        <td className="py-3 px-3 text-center text-zinc-500 font-mono">
                          {row.nrSecaoPrincipal ?? -1}
                        </td>

                        {/* 7. NR LOCAL VOTACAO */}
                        <td className="py-3 px-3 text-center font-mono font-semibold text-zinc-700 dark:text-zinc-300">
                          {row.nrLocalVotacao || '---'}
                        </td>

                        {/* 8. NM LOCAL VOTACAO */}
                        <td className="py-3 px-3 font-bold text-zinc-900 dark:text-zinc-100 max-w-[220px] truncate" title={row.nmLocalVotacao || row.local}>
                          {row.nmLocalVotacao || row.local}
                        </td>

                        {/* 9. DS ENDERECO */}
                        <td className="py-3 px-3 text-zinc-500 dark:text-zinc-400 max-w-[200px] truncate" title={row.dsEndereco || row.endereco}>
                          {row.dsEndereco || row.endereco || '---'}
                        </td>

                        {/* 10. NM BAIRRO */}
                        <td className="py-3 px-3 text-zinc-700 dark:text-zinc-300 font-semibold uppercase whitespace-nowrap">
                          {row.nmBairro || row.bairro || '---'}
                        </td>

                        {/* 11. QT_ELEITOR_SECAO */}
                        <td className="py-3 px-3 text-center font-black text-blue-600 dark:text-blue-400 text-xs whitespace-nowrap">
                          {(row.qtEleitorSecao || row.eleitores || 0).toLocaleString()}
                        </td>

                        {/* 12. NM LOCAL VOTACAO ORIGINAL */}
                        <td className="py-3 px-3 text-zinc-500 dark:text-zinc-400 max-w-[180px] truncate" title={row.nmLocalVotacaoOriginal || row.nmLocalVotacao || row.local}>
                          {row.nmLocalVotacaoOriginal || row.nmLocalVotacao || row.local || '---'}
                        </td>

                        {/* 13. DS ENDERECO_LOCVT_ORIGINAL */}
                        <td className="py-3 px-3 text-zinc-500 dark:text-zinc-400 max-w-[180px] truncate" title={row.dsEnderecoLocvtOriginal || row.dsEndereco || row.endereco}>
                          {row.dsEnderecoLocvtOriginal || row.dsEndereco || row.endereco || '---'}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>

            {/* Pagination Controls */}
            {processedTableData.length > 0 && (
              <div className="p-4 border-t border-zinc-200 dark:border-zinc-800 bg-zinc-50 dark:bg-zinc-900/50 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs">
                <div className="text-zinc-500 dark:text-zinc-400 text-center sm:text-left">
                  Mostrando <span className="font-bold text-zinc-900 dark:text-zinc-100">{Math.min((currentPage - 1) * ITEMS_PER_PAGE + 1, processedTableData.length)}</span> a <span className="font-bold text-zinc-900 dark:text-zinc-100">{Math.min(currentPage * ITEMS_PER_PAGE, processedTableData.length)}</span> de <span className="font-bold text-zinc-900 dark:text-zinc-100">{processedTableData.length}</span> locais
                </div>

                <div className="flex items-center gap-2">
                  <span className="text-zinc-500 dark:text-zinc-400 font-medium mr-1">
                    Página <span className="font-bold text-zinc-900 dark:text-zinc-100">{currentPage}</span> de <span className="font-bold text-zinc-900 dark:text-zinc-100">{totalPages}</span>
                  </span>

                  <div className="flex items-center gap-1">
                    <button
                      onClick={() => setCurrentPage(1)}
                      disabled={currentPage === 1}
                      className="p-1.5 rounded border border-zinc-200 dark:border-zinc-700 bg-white dark:bg-zinc-800 text-zinc-700 dark:text-zinc-300 disabled:opacity-40 disabled:cursor-not-allowed hover:bg-zinc-100 dark:hover:bg-zinc-700 transition-colors cursor-pointer"
                      title="Primeira página"
                    >
                      <ChevronsLeft className="w-4 h-4" />
                    </button>
                    
                    <button
                      onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
                      disabled={currentPage === 1}
                      className="p-1.5 rounded border border-zinc-200 dark:border-zinc-700 bg-white dark:bg-zinc-800 text-zinc-700 dark:text-zinc-300 disabled:opacity-40 disabled:cursor-not-allowed hover:bg-zinc-100 dark:hover:bg-zinc-700 transition-colors cursor-pointer"
                      title="Página anterior"
                    >
                      <ChevronLeft className="w-4 h-4" />
                    </button>

                    <button
                      onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
                      disabled={currentPage === totalPages}
                      className="p-1.5 rounded border border-zinc-200 dark:border-zinc-700 bg-white dark:bg-zinc-800 text-zinc-700 dark:text-zinc-300 disabled:opacity-40 disabled:cursor-not-allowed hover:bg-zinc-100 dark:hover:bg-zinc-700 transition-colors cursor-pointer"
                      title="Próxima página"
                    >
                      <ChevronRight className="w-4 h-4" />
                    </button>

                    <button
                      onClick={() => setCurrentPage(totalPages)}
                      disabled={currentPage === totalPages}
                      className="p-1.5 rounded border border-zinc-200 dark:border-zinc-700 bg-white dark:bg-zinc-800 text-zinc-700 dark:text-zinc-300 disabled:opacity-40 disabled:cursor-not-allowed hover:bg-zinc-100 dark:hover:bg-zinc-700 transition-colors cursor-pointer"
                      title="Última página"
                    >
                      <ChevronsRight className="w-4 h-4" />
                    </button>
                  </div>
                </div>
              </div>
            )}

          </div>
        </>
        ) : (
            /* THE BRAND NEW CRUZAMENTO DE DADOS VIEW */
            <div className="space-y-6 animate-fadeIn">
              {/* INTRO AND FILTERS */}
              <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 p-5 rounded-sm shadow-sm">
                <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
                  <div>
                    <h2 className="text-sm font-black text-zinc-950 dark:text-white uppercase tracking-wider flex items-center gap-2">
                      <Compass className="w-5 h-5 text-blue-600 animate-pulse" />
                      Diretrizes Estratégicas de Cobertura Eleitoral
                    </h2>
                    <p className="text-xs text-zinc-500 dark:text-zinc-400 mt-1">
                      Cruzamento em tempo real dos eleitores oficiais do TRE com os cadastros captados pelas equipes da campanha. Use este painel para direcionar visitas, organizar panfletagens e priorizar municípios com baixo engajamento.
                    </p>
                  </div>

                  <div className="flex flex-wrap items-center gap-2">
                    <div className="flex items-center gap-1.5 bg-zinc-100 dark:bg-zinc-800 p-1 rounded-sm border border-zinc-200 dark:border-zinc-700">
                      <button
                        onClick={() => setStrategicSort('coverage')}
                        className={`px-3 py-1.5 text-[10px] font-black uppercase tracking-wider rounded-sm transition-all ${
                          strategicSort === 'coverage'
                            ? 'bg-blue-600 text-white shadow-sm'
                            : 'text-zinc-500 hover:text-zinc-800 dark:hover:text-zinc-200'
                        }`}
                      >
                        Menor Cobertura
                      </button>
                      <button
                        onClick={() => setStrategicSort('missing')}
                        className={`px-3 py-1.5 text-[10px] font-black uppercase tracking-wider rounded-sm transition-all ${
                          strategicSort === 'missing'
                            ? 'bg-blue-600 text-white shadow-sm'
                            : 'text-zinc-500 hover:text-zinc-800 dark:hover:text-zinc-200'
                        }`}
                      >
                        Mais Faltantes (Volume)
                      </button>
                    </div>

                    <select
                      value={strategicStatusFilter}
                      onChange={(e: any) => setStrategicStatusFilter(e.target.value)}
                      className="text-[10px] font-black uppercase tracking-wider bg-zinc-100 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-sm p-1.5 focus:outline-none focus:ring-1 focus:ring-blue-600"
                    >
                      <option value="Todos">✦ Filtro: Todos os Status</option>
                      <option value="critical">🔴 Crítico (&lt; 0.5%)</option>
                      <option value="low">🟠 Baixo (0.5% - 1.5%)</option>
                      <option value="medium">🟡 Médio (1.5% - 4.0%)</option>
                      <option value="good">🟢 Bom (&gt;= 4.0%)</option>
                    </select>
                  </div>
                </div>
              </div>

              {/* BENTO STATS GRID */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                {/* Card 1: Total TRE */}
                <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 p-4 rounded-sm shadow-sm flex items-center justify-between">
                  <div>
                    <p className="text-[10px] font-black uppercase text-zinc-400 tracking-wider">Eleitores TRE (Estado)</p>
                    <p className="text-2xl font-black text-zinc-950 dark:text-white mt-1">
                      {totalStateEleitores.toLocaleString()}
                    </p>
                    <p className="text-[8px] text-zinc-400 dark:text-zinc-500 uppercase tracking-widest mt-0.5">Base Oficial TRE Carregada</p>
                  </div>
                  <div className="p-3 bg-zinc-100 dark:bg-zinc-800 rounded-full text-zinc-500 dark:text-zinc-400">
                    <Database className="w-5 h-5" />
                  </div>
                </div>

                {/* Card 2: Total Campaign */}
                <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 p-4 rounded-sm shadow-sm flex items-center justify-between">
                  <div>
                    <p className="text-[10px] font-black uppercase text-zinc-400 tracking-wider">Cadastros na Campanha</p>
                    <p className="text-2xl font-black text-blue-600 dark:text-blue-600 mt-1">
                      {campaignVoters.length.toLocaleString()}
                    </p>
                    <p className="text-[8px] text-zinc-400 dark:text-zinc-500 uppercase tracking-widest mt-0.5">Eleitores Mapeados Ativos</p>
                  </div>
                  <div className="p-3 bg-blue-600/10 rounded-full text-blue-600">
                    <Users className="w-5 h-5" />
                  </div>
                </div>

                {/* Card 3: Average Coverage */}
                <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 p-4 rounded-sm shadow-sm flex items-center justify-between">
                  <div>
                    <p className="text-[10px] font-black uppercase text-zinc-400 tracking-wider">Índice Médio de Penetração</p>
                    <p className="text-2xl font-black text-zinc-950 dark:text-white mt-1">
                      {totalStateEleitores > 0 ? ((campaignVoters.length / totalStateEleitores) * 100).toFixed(2) : '0.00'}%
                    </p>
                    <p className="text-[8px] text-zinc-400 dark:text-zinc-500 uppercase tracking-widest mt-0.5">Média Geral de Cobertura</p>
                  </div>
                  <div className="p-3 bg-zinc-100 dark:bg-zinc-800 rounded-full text-zinc-500 dark:text-zinc-400">
                    <Percent className="w-5 h-5" />
                  </div>
                </div>

                {/* Card 4: Critical Municipalities */}
                <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 p-4 rounded-sm shadow-sm flex items-center justify-between">
                  <div>
                    <p className="text-[10px] font-black uppercase text-zinc-400 tracking-wider">Municípios Críticos</p>
                    <p className="text-2xl font-black text-rose-600 dark:text-rose-400 mt-1">
                      {crossReferencedData.filter(d => d.status === 'critical' || d.status === 'low').length}
                    </p>
                    <p className="text-[8px] text-zinc-400 dark:text-zinc-500 uppercase tracking-widest mt-0.5">Cobertura Abaixo de 1.5%</p>
                  </div>
                  <div className="p-3 bg-rose-500/10 rounded-full text-rose-500">
                    <AlertTriangle className="w-5 h-5" />
                  </div>
                </div>
              </div>

              {/* VISUAL CHART AND DETAILED LIST SIDE BY SIDE */}
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
                
                {/* CHART: VISUALIZATION OF COVERAGE INDEX */}
                <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 p-4 rounded-sm shadow-sm lg:col-span-7">
                  <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-2 mb-4">
                    <div className="text-zinc-950 dark:text-white font-black text-xs uppercase tracking-wider flex items-center gap-1.5">
                      <TrendingUp className="w-4 h-4 text-blue-600" />
                      Gráfico Estratégico: Índice de Cobertura por Município (%)
                    </div>
                    <span className="text-[8px] font-black text-blue-600 bg-blue-600/10 px-1.5 py-0.5 rounded-sm uppercase">
                      Menor Cobertura = Maior Prioridade
                    </span>
                  </div>
                  
                  <div className="h-[420px] w-full text-xs">
                    {crossReferencedData.length === 0 ? (
                      <div className="h-full flex items-center justify-center text-zinc-400 uppercase tracking-wider">Carregue dados do TRE para gerar o gráfico</div>
                    ) : (
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart
                          data={[...crossReferencedData].sort((a, b) => a.coverageRate - b.coverageRate)}
                          layout="vertical"
                          margin={{ top: 5, right: 20, left: 20, bottom: 5 }}
                        >
                          <CartesianGrid strokeDasharray="3 3" stroke="#e4e4e7" opacity={0.3} />
                          <XAxis type="number" stroke="#888888" fontSize={9} tickFormatter={(tick) => `${tick}%`} />
                          <YAxis 
                            type="category" 
                            dataKey="municipio" 
                            stroke="#888888" 
                            fontSize={9} 
                            width={95}
                          />
                          <Tooltip 
                            contentStyle={{ backgroundColor: '#18181b', color: '#fff', fontSize: 10, borderRadius: 4 }}
                            formatter={(value: any) => [`${parseFloat(value).toFixed(3)}% de Penetração`, 'Índice de Cobertura']}
                          />
                          <Bar dataKey="coverageRate" radius={[0, 4, 4, 0]}>
                            {[...crossReferencedData].sort((a, b) => a.coverageRate - b.coverageRate).map((entry, index) => {
                              // Style bars color based on status
                              let color = '#ef4444'; // Red for critical
                              if (entry.coverageRate >= 4.0) color = '#10b981'; // Green
                              else if (entry.coverageRate >= 1.5) color = '#0578d3'; // Yellow
                              else if (entry.coverageRate >= 0.5) color = '#f97316'; // Orange
                              
                              return <Cell key={`cell-${index}`} fill={color} opacity={0.85} />;
                            })}
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    )}
                  </div>
                  <div className="mt-3 flex items-center justify-center gap-4 flex-wrap text-[9px] font-black uppercase tracking-wider border-t border-zinc-100 dark:border-zinc-800 pt-3">
                    <div className="flex items-center gap-1.5"><span className="w-3 h-3 bg-red-500 rounded-sm inline-block"></span> Crítico (&lt; 0.5%)</div>
                    <div className="flex items-center gap-1.5"><span className="w-3 h-3 bg-orange-500 rounded-sm inline-block"></span> Baixo (0.5% - 1.5%)</div>
                    <div className="flex items-center gap-1.5"><span className="w-3 h-3 bg-blue-600 rounded-sm inline-block"></span> Médio (1.5% - 4.0%)</div>
                    <div className="flex items-center gap-1.5"><span className="w-3 h-3 bg-emerald-500 rounded-sm inline-block"></span> Bom (&gt;= 4.0%)</div>
                  </div>
                </div>

                {/* DETAILED SIDE DRILLDOWN INTERACTIVE CARD */}
                <div className="lg:col-span-5 flex flex-col justify-between bg-zinc-950 text-white border border-zinc-800 p-5 rounded-sm shadow-lg">
                  {selectedStrategyMun ? (
                    (() => {
                      const data = crossReferencedData.find(d => d.municipio === selectedStrategyMun);
                      if (!data) return <p className="text-xs text-zinc-400">Município não encontrado</p>;
                      return (
                        <div className="space-y-5 h-full flex flex-col justify-between">
                          <div>
                            <div className="flex items-center justify-between">
                              <span className="text-[9px] font-black tracking-widest text-blue-600 uppercase">Diagnóstico Regional</span>
                              <button 
                                onClick={() => setSelectedStrategyMun(null)}
                                className="text-[10px] text-zinc-400 hover:text-white transition-colors uppercase"
                              >
                                [Fechar]
                              </button>
                            </div>
                            <h3 className="text-xl font-black text-white uppercase tracking-tight mt-1 border-b border-zinc-800 pb-2">
                              {data.municipio}
                            </h3>

                            {/* Stat items */}
                            <div className="grid grid-cols-2 gap-4 mt-4">
                              <div className="bg-zinc-900 border border-zinc-850 p-2.5 rounded-sm">
                                <p className="text-[8px] font-black text-zinc-400 uppercase tracking-widest">Aptos (TRE)</p>
                                <p className="text-lg font-black text-white mt-1">{data.treElectors.toLocaleString()}</p>
                                <p className="text-[7px] text-zinc-500 uppercase mt-0.5">({data.treSharePercent.toFixed(1)}% do Estado)</p>
                              </div>
                              <div className="bg-zinc-900 border border-zinc-850 p-2.5 rounded-sm">
                                <p className="text-[8px] font-black text-zinc-400 uppercase tracking-widest">Cadastrados</p>
                                <p className="text-lg font-black text-blue-600 mt-1">{data.campaignVoters.toLocaleString()}</p>
                                <p className="text-[7px] text-zinc-500 uppercase mt-0.5">Eleitores Cooptados</p>
                              </div>
                            </div>

                            {/* Progress Meter */}
                            <div className="mt-4 bg-zinc-900 border border-zinc-850 p-3 rounded-sm space-y-2">
                              <div className="flex items-center justify-between text-[9px] font-black uppercase">
                                <span className="text-zinc-400">Taxa de Cobertura</span>
                                <span className={
                                  data.status === 'critical' ? 'text-red-500' :
                                  data.status === 'low' ? 'text-orange-500' :
                                  data.status === 'medium' ? 'text-blue-600' : 'text-emerald-500'
                                }>
                                  {data.coverageRate.toFixed(3)}%
                                </span>
                              </div>
                              <div className="w-full bg-zinc-800 h-2 rounded-sm overflow-hidden">
                                <div 
                                  className={`h-full ${
                                    data.status === 'critical' ? 'bg-red-500' :
                                    data.status === 'low' ? 'bg-orange-500' :
                                    data.status === 'medium' ? 'bg-blue-600' : 'bg-emerald-500'
                                  }`}
                                  style={{ width: `${Math.min(100, data.coverageRate * 15)}%` }} // Scaling multiplier for visibility
                                ></div>
                              </div>
                              <p className="text-[8px] text-zinc-500 leading-normal uppercase">
                                Potencial restante de captação: <strong className="text-zinc-300">{(data.treElectors - data.campaignVoters).toLocaleString()}</strong> eleitores sem contato registrado.
                              </p>
                            </div>

                            {/* Actions list */}
                            <div className="mt-5 space-y-3">
                              <div className="text-[10px] font-black text-blue-600 uppercase tracking-wider flex items-center gap-1">
                                <Briefcase className="w-3.5 h-3.5" />
                                Plano de Ação Tático Sugerido
                              </div>
                              <div className="p-3 bg-blue-600/5 border border-blue-600/20 text-blue-200 text-[11px] leading-relaxed rounded-sm">
                                {data.recommendation}
                              </div>

                              <ul className="text-[10px] space-y-1.5 text-zinc-300 uppercase font-black tracking-wide pl-1 mt-2">
                                <li className="flex items-start gap-2">
                                  <span className="text-blue-600">•</span>
                                  <span>Alocar articulador local na Zona {data.treElectors > 10000 ? "de alto quociente" : "eleitoral da região"}</span>
                                </li>
                                <li className="flex items-start gap-2">
                                  <span className="text-blue-600">•</span>
                                  <span>Mapear bairros com menor número de fichas entregues</span>
                                </li>
                                <li className="flex items-start gap-2">
                                  <span className="text-blue-600">•</span>
                                  <span>Visita oficial do candidato / Coordenadores na região</span>
                                </li>
                              </ul>
                            </div>
                          </div>

                          <div className="pt-4 border-t border-zinc-800 mt-4">
                            <button
                              onClick={() => {
                                setSelectedMunicipio(data.municipio);
                                setSubTab('tre_oficial');
                              }}
                              className="w-full py-2 bg-blue-600 text-white font-black text-[10px] uppercase tracking-wider rounded-sm hover:bg-blue-500 transition-all text-center flex items-center justify-center gap-1"
                            >
                              <Search className="w-3.5 h-3.5" />
                              Ver Locais de Votação no TRE
                            </button>
                          </div>
                        </div>
                      );
                    })()
                  ) : (
                    <div className="h-full flex flex-col items-center justify-center text-center p-6 space-y-4">
                      <Target className="w-12 h-12 text-blue-600/40 animate-pulse" />
                      <div>
                        <h4 className="text-xs font-black uppercase text-white tracking-widest">Selecione um Município</h4>
                        <p className="text-[10px] text-zinc-400 max-w-xs mt-1.5 leading-relaxed uppercase">
                          Clique em qualquer município na tabela ou gráfico ao lado para abrir o diagnóstico completo e receber as diretrizes de mobilização das equipes.
                        </p>
                      </div>
                    </div>
                  )}
                </div>

              </div>

              {/* STRATEGIC RANKED TABLE */}
              <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-sm p-4 shadow-sm overflow-hidden">
                <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3 mb-4 flex-wrap gap-2">
                  <div className="flex items-center gap-1.5 text-zinc-950 dark:text-white font-black text-xs uppercase tracking-wider">
                    <Award className="w-4 h-4 text-blue-600" />
                    Lista de Priorização por Cobertura Eleitoral ({strategicStatusFilter === 'Todos' ? 'Todos os Municípios' : `Status: ${strategicStatusFilter.toUpperCase()}`})
                  </div>
                  <div className="text-[9px] text-zinc-400 font-bold uppercase">
                    Exibindo {
                      (strategicSort === 'coverage' ? rankedByCoverage : rankedByStrategicPriority)
                        .filter(d => strategicStatusFilter === 'Todos' || d.status === strategicStatusFilter)
                        .length
                    } municípios ordenados por {strategicSort === 'coverage' ? 'Menor Cobertura' : 'Volume de Eleitores Faltantes'}
                  </div>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="border-b border-zinc-200 dark:border-zinc-800 text-[10px] font-black uppercase tracking-wider text-zinc-500 dark:text-zinc-400">
                        <th className="py-2.5 px-4">Município</th>
                        <th className="py-2.5 px-4 text-right">Eleitores (TRE)</th>
                        <th className="py-2.5 px-4 text-right">Cadastrados (Nossos)</th>
                        <th className="py-2.5 px-4 text-center">Índice de Cobertura</th>
                        <th className="py-2.5 px-4 text-right">Faltam Cadastrar</th>
                        <th className="py-2.5 px-4">Prioridade / Status</th>
                        <th className="py-2.5 px-4 text-center">Ações</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-zinc-100 dark:divide-zinc-850 text-xs">
                      {(strategicSort === 'coverage' ? rankedByCoverage : rankedByStrategicPriority)
                        .filter(d => strategicStatusFilter === 'Todos' || d.status === strategicStatusFilter)
                        .map((row) => {
                          return (
                            <tr 
                              key={row.municipio}
                              className={`hover:bg-zinc-50 dark:hover:bg-zinc-800/40 transition-colors cursor-pointer ${selectedStrategyMun === row.municipio ? 'bg-blue-600/5 dark:bg-blue-600/5 border-l-2 border-l-blue-600' : ''}`}
                              onClick={() => setSelectedStrategyMun(row.municipio)}
                            >
                              {/* Municipio */}
                              <td className="py-3 px-4 font-black text-zinc-950 dark:text-white uppercase text-[11px]">
                                {row.municipio}
                              </td>

                              {/* TRE Electors */}
                              <td className="py-3 px-4 text-right font-semibold text-zinc-600 dark:text-zinc-300">
                                {row.treElectors > 0 ? row.treElectors.toLocaleString() : '---'}
                              </td>

                              {/* Campaign Voters */}
                              <td className="py-3 px-4 text-right font-black text-blue-600 dark:text-blue-600 text-[12px]">
                                {row.campaignVoters.toLocaleString()}
                              </td>

                              {/* Coverage Indicator */}
                              <td className="py-3 px-4">
                                <div className="flex flex-col items-center justify-center max-w-[140px] mx-auto space-y-1">
                                  <span className="font-bold text-[10px] text-zinc-700 dark:text-zinc-300">
                                    {row.coverageRate.toFixed(3)}%
                                  </span>
                                  <div className="w-full bg-zinc-100 dark:bg-zinc-800 h-1.5 rounded-sm overflow-hidden">
                                    <div 
                                      className={`h-full ${
                                        row.status === 'critical' ? 'bg-red-500' :
                                        row.status === 'low' ? 'bg-orange-500' :
                                        row.status === 'medium' ? 'bg-blue-600' : 'bg-emerald-500'
                                      }`}
                                      style={{ width: `${Math.min(100, row.coverageRate * 15)}%` }}
                                    ></div>
                                  </div>
                                </div>
                              </td>

                              {/* Missing Voters */}
                              <td className="py-3 px-4 text-right font-bold text-zinc-500 dark:text-zinc-400">
                                {row.missingElectors > 0 ? row.missingElectors.toLocaleString() : '---'}
                              </td>

                              {/* Priority Status Badge */}
                              <td className="py-3 px-4">
                                {row.status === 'critical' && (
                                  <span className="px-2 py-0.5 text-[8px] font-black uppercase rounded-sm bg-red-500/10 border border-red-500/30 text-red-600 dark:text-red-400">🔴 Crítico</span>
                                )}
                                {row.status === 'low' && (
                                  <span className="px-2 py-0.5 text-[8px] font-black uppercase rounded-sm bg-orange-500/10 border border-orange-500/30 text-orange-600 dark:text-orange-400">🟠 Baixo</span>
                                )}
                                {row.status === 'medium' && (
                                  <span className="px-2 py-0.5 text-[8px] font-black uppercase rounded-sm bg-blue-600/10 border border-blue-600/30 text-blue-600 dark:text-blue-600">🟡 Médio</span>
                                )}
                                {row.status === 'good' && (
                                  <span className="px-2 py-0.5 text-[8px] font-black uppercase rounded-sm bg-emerald-500/10 border border-emerald-500/30 text-emerald-600 dark:text-emerald-400">🟢 Consolidado</span>
                                )}
                              </td>

                              {/* Action Buttons */}
                              <td className="py-3 px-4 text-center">
                                <button 
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    setSelectedStrategyMun(row.municipio);
                                  }}
                                  className="px-2.5 py-1 bg-zinc-100 dark:bg-zinc-800 text-zinc-700 dark:text-zinc-300 hover:bg-blue-600 hover:text-white font-black text-[9px] uppercase tracking-wider rounded-sm transition-all"
                                >
                                  Diagnóstico
                                </button>
                              </td>
                            </tr>
                          );
                        })}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}
        </>
      )}

      {/* SUPABASE DATABASE CONFIGURATION MODAL */}
      <SupabaseConfigModal
        isOpen={isSupabaseModalOpen}
        onClose={() => setIsSupabaseModalOpen(false)}
      />

    </div>
  );
}
