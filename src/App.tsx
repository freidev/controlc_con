import { useState, useMemo } from 'react';
import { 
  LayoutDashboard, Database, Wallet, Scale, Wrench, History, 
  FileInput, FileOutput, Fuel, Users, LogOut, Bell, Check, 
  X, Search, Filter, Plus, Upload, Download, Eye, EyeOff,
  ChevronDown, Calendar, TrendingUp, DollarSign, Droplets,
  AlertCircle, CheckCircle, Clock, User as UserIcon, Settings,
  Menu, Trash2, Pencil, Camera, Mail, Briefcase
} from 'lucide-react';
import { 
  BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  AreaChart, Area
} from 'recharts';
import * as XLSX from 'xlsx';
import { clsx } from 'clsx';
import { twMerge } from 'tailwind-merge';
import type { 
  User, UserRole, UserStatus, Equipment, Rateio, RateioItem,
  Budget, Abastecimento, DieselPrice, FilterState, Notification 
} from './types';

function cn(...classes: (string | undefined | null | false)[]) {
  return twMerge(clsx(classes));
}

function stripEmojis(value: string) {
  return value.replace(/[\p{Extended_Pictographic}\uFE0F]/gu, '').replace(/\s{2,}/g, ' ').trim();
}

function cleanText(value: string) {
  return stripEmojis(String(value || ''));
}

function padTwoDigits(value: number | string) {
  return String(value).padStart(2, '0');
}

function normalizeDateValue(value: string | number) {
  const raw = cleanText(String(value ?? ''));
  if (!raw) return '';

  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(raw)) {
    const [dia, mes, ano] = raw.split('/');
    return `${ano}-${mes}-${dia}`;
  }

  if (/^\d{2}-\d{2}-\d{4}$/.test(raw)) {
    const [dia, mes, ano] = raw.split('-');
    return `${ano}-${mes}-${dia}`;
  }

  if (/^\d+(\.\d+)?$/.test(raw)) {
    const excelDate = XLSX.SSF.parse_date_code(Number(raw));
    if (excelDate) {
      return `${excelDate.y}-${padTwoDigits(excelDate.m)}-${padTwoDigits(excelDate.d)}`;
    }
  }

  const parsedDate = new Date(raw);
  if (!Number.isNaN(parsedDate.getTime())) {
    return `${parsedDate.getFullYear()}-${padTwoDigits(parsedDate.getMonth() + 1)}-${padTwoDigits(parsedDate.getDate())}`;
  }

  return raw;
}

function getDateParts(value: string) {
  const normalized = normalizeDateValue(value);
  if (/^\d{4}-\d{2}-\d{2}$/.test(normalized)) {
    const [ano, mes, dia] = normalized.split('-');
    return { dia, mes, ano };
  }

  return { dia: '', mes: '', ano: '' };
}

function deriveWeekLabel(value: string) {
  const { dia } = getDateParts(value);
  return dia ? `Sem ${Math.ceil(Number(dia) / 7)}` : '';
}

function sortTextValues(values: string[]) {
  return [...new Set(values.filter(Boolean))].sort((a, b) =>
    a.localeCompare(b, 'pt-BR', { numeric: true, sensitivity: 'base' }),
  );
}

function sortNumericTextValues(values: string[]) {
  return [...new Set(values.filter(Boolean))].sort((a, b) => Number(a) - Number(b));
}

function formatNumber(value: number | undefined | null, decimals = 2) {
  const safe = Number.isFinite(Number(value)) ? Number(value) : 0;
  return safe.toLocaleString('pt-BR', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
}

function formatCurrency(value: number | undefined | null) {
  return `R$ ${formatNumber(value, 2)}`;
}

function formatLiters(value: number | undefined | null, decimals = 2) {
  return `${formatNumber(value, decimals)} L`;
}

function formatCompactNumber(value: number | undefined | null) {
  const safe = Number.isFinite(Number(value)) ? Number(value) : 0;
  if (Math.abs(safe) >= 1_000_000) {
    return `${(safe / 1_000_000).toLocaleString('pt-BR', { maximumFractionDigits: 2 })}M`;
  }
  if (Math.abs(safe) >= 1_000) {
    return `${(safe / 1_000).toLocaleString('pt-BR', { maximumFractionDigits: 2 })}k`;
  }
  return safe.toLocaleString('pt-BR', { maximumFractionDigits: 2 });
}

function getRoleLabel(role: string) {
  if (role === 'admin') return 'Administrador';
  if (role === 'supervisor') return 'Supervisor';
  return 'Operador';
}

function getMonthLabel(value: string) {
  const monthNames = [
    'Janeiro',
    'Fevereiro',
    'Março',
    'Abril',
    'Maio',
    'Junho',
    'Julho',
    'Agosto',
    'Setembro',
    'Outubro',
    'Novembro',
    'Dezembro',
  ];

  const monthIndex = Number(value) - 1;
  return monthNames[monthIndex] ? `${value} - ${monthNames[monthIndex]}` : value;
}

function splitChartLabel(value: string, maxChars = 24, maxLines = 3) {
  const words = cleanText(value).split(' ').filter(Boolean);
  const lines: string[] = [];
  let currentLine = '';

  for (const word of words) {
    const nextLine = currentLine ? `${currentLine} ${word}` : word;
    if (nextLine.length <= maxChars) {
      currentLine = nextLine;
      continue;
    }

    if (currentLine) lines.push(currentLine);
    currentLine = word;

    if (lines.length === maxLines - 1) break;
  }

  if (currentLine && lines.length < maxLines) {
    lines.push(currentLine);
  }

  const remainingWords = words.slice(lines.join(' ').split(' ').filter(Boolean).length);
  if (remainingWords.length > 0 && lines.length > 0) {
    lines[lines.length - 1] = `${lines[lines.length - 1].slice(0, Math.max(0, maxChars - 3))}...`;
  }

  return lines.length > 0 ? lines : [value];
}

function CategoryAxisTick({ x = 0, y = 0, payload, anchorX = -12, textAnchor = 'end' as 'start' | 'end' | 'middle' }: any) {
  const lines = splitChartLabel(String(payload?.value || ''));

  return (
    <g transform={`translate(${x},${y})`}>
      <text x={anchorX} y={0} textAnchor={textAnchor} fill="#475569" fontSize={11}>
        {lines.map((line, index) => (
          <tspan key={`${line}-${index}`} x={anchorX} dy={index === 0 ? -6 : 13}>
            {line}
          </tspan>
        ))}
      </text>
    </g>
  );
}

function getMultiSelectLabel(values: string[], options: string[], formatOption?: (value: string) => string) {
  const selected = values.filter((value) => options.includes(value));

  if (selected.length === 0) return 'Todos';
  if (selected.length === 1) return formatOption ? formatOption(selected[0]) : selected[0];
  return `${selected.length} selecionados`;
}

function DashboardMultiSelect({
  label,
  options,
  selectedValues,
  onChange,
  formatOption,
}: {
  label: string;
  options: string[];
  selectedValues: string[];
  onChange: (values: string[]) => void;
  formatOption?: (value: string) => string;
}) {
  const toggleValue = (value: string) => {
    onChange(
      selectedValues.includes(value)
        ? selectedValues.filter((item) => item !== value)
        : [...selectedValues, value],
    );
  };

  const clearSelection = () => onChange([]);

  return (
    <div>
      <label className="mb-1 block text-xs font-medium text-slate-600">{label}</label>
      <details className="group relative">
        <summary className="flex cursor-pointer list-none items-center justify-between rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm text-slate-700 outline-none transition hover:border-slate-300 group-open:border-red-800 group-open:ring-2 group-open:ring-red-800/20">
          <span className="truncate pr-3">{getMultiSelectLabel(selectedValues, options, formatOption)}</span>
          <ChevronDown className="h-4 w-4 shrink-0 text-slate-400 transition group-open:rotate-180" />
        </summary>
        <div className="absolute left-0 right-0 z-30 mt-2 rounded-xl border border-slate-200 bg-white p-2 shadow-xl">
          <div className="mb-2 flex items-center justify-between border-b border-slate-100 px-2 pb-2">
            <span className="text-xs font-medium text-slate-500">{selectedValues.length} selecionado(s)</span>
            <button
              type="button"
              onClick={clearSelection}
              className="text-xs font-medium text-red-700 transition hover:text-red-800"
            >
              Limpar
            </button>
          </div>
          <div className="max-h-56 space-y-1 overflow-y-auto pr-1">
            {options.map((option) => (
              <label
                key={`${label}-${option}`}
                className="flex cursor-pointer items-center gap-2 rounded-lg px-2 py-2 text-sm text-slate-700 transition hover:bg-slate-50"
              >
                <input
                  type="checkbox"
                  checked={selectedValues.includes(option)}
                  onChange={() => toggleValue(option)}
                  className="h-4 w-4 rounded border-slate-300 text-red-800 focus:ring-red-800/30"
                />
                <span className="truncate">{formatOption ? formatOption(option) : option}</span>
              </label>
            ))}
            {options.length === 0 && <p className="px-2 py-2 text-sm text-slate-400">Sem opções na base</p>}
          </div>
        </div>
      </details>
    </div>
  );
}

function ComboInput({
  value,
  onChange,
  options,
  placeholder,
  className,
  wrapperClassName,
  onOptionSelect,
}: {
  value: string;
  onChange: (value: string) => void;
  options: string[];
  placeholder?: string;
  className?: string;
  wrapperClassName?: string;
  onOptionSelect?: (value: string) => void;
}) {
  const [isOpen, setIsOpen] = useState(false);

  const filteredOptions = useMemo(() => {
    const query = cleanText(value).toLowerCase();
    const uniqueOptions = sortTextValues(options);

    if (!query) return uniqueOptions.slice(0, 12);

    return uniqueOptions
      .filter((option) => option.toLowerCase().includes(query))
      .slice(0, 12);
  }, [options, value]);

  const selectOption = (option: string) => {
    onChange(option);
    onOptionSelect?.(option);
    setIsOpen(false);
  };

  return (
    <div className={cn('relative', wrapperClassName)}>
      <input
        type="text"
        value={value}
        onChange={(e) => {
          onChange(e.target.value);
          setIsOpen(true);
        }}
        onFocus={() => setIsOpen(true)}
        onBlur={() => window.setTimeout(() => setIsOpen(false), 120)}
        placeholder={placeholder}
        className={cn(
          'w-full rounded-lg border border-slate-200 bg-white px-4 py-2 pr-9 text-sm outline-none focus:border-red-800 focus:ring-2 focus:ring-red-800/20',
          className,
        )}
      />
      <ChevronDown className="pointer-events-none absolute right-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />

      {isOpen && filteredOptions.length > 0 && (
        <div className="absolute left-0 right-0 top-full z-50 mt-1 max-h-72 overflow-y-auto rounded-lg border border-slate-300 bg-white py-1 shadow-xl">
          {filteredOptions.map((option) => {
            const isSelected = option === value;
            return (
              <button
                key={option}
                type="button"
                onMouseDown={(e) => {
                  e.preventDefault();
                  selectOption(option);
                }}
                className={cn(
                  'block w-full px-4 py-2 text-left text-sm transition',
                  isSelected
                    ? 'bg-blue-600 text-white'
                    : 'bg-white text-slate-900 hover:bg-blue-50',
                )}
              >
                {option}
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}

// Initial Data
const initialUsers: User[] = [
  { id: 1, username: 'admin', password: 'admin123', role: 'admin', status: 'approved', name: 'Administrador Principal', createdAt: new Date().toISOString() },
  { id: 2, username: 'operador1', password: 'op123', role: 'operator', status: 'approved', name: 'Operador Teste', createdAt: new Date().toISOString() }
];

const initialDieselPrice: DieselPrice = {
  id: 1,
  price: 7.38,
  updatedAt: new Date().toISOString(),
  updatedBy: 'admin'
};

const createDefaultDashboardFilters = (): FilterState => ({
  ccNovo: [],
  diretoria: [],
  gerencia: [],
  areaLotacao: [],
  fornecedor: [],
  equipamento: [],
  dia: [],
  mes: [],
  semana: [],
  ano: [],
});

type DatabaseFilterState = {
  gerencia: string[];
  semana: string[];
  dataInicio: string;
  dataFim: string;
  fornecedor: string[];
  ccNovo: string[];
  equipamento: string[];
};

const createDefaultDatabaseFilters = (): DatabaseFilterState => ({
  gerencia: [],
  semana: [],
  dataInicio: '',
  dataFim: '',
  fornecedor: [],
  ccNovo: [],
  equipamento: [],
});

const COLORS = ['#8B1538', '#1E3A8A', '#059669', '#D97706', '#7C3AED', '#DC2626', '#0891B2', '#7C2D12'];
const RATEIO_COLORS = ['#1E3A8A', '#F97316', '#10B981', '#A855F7', '#EAB308', '#EC4899', '#06B6D4', '#84CC16'];

const BASE_EXPORT_COLUMNS = [
  { key: 'id', label: 'ID' },
  { key: 'ccNovo', label: 'CC NOVO' },
  { key: 'diretoria', label: 'DIRETORIA' },
  { key: 'gerencia', label: 'GERÊNCIA' },
  { key: 'areaLotacao', label: 'ÁREA LOT.' },
  { key: 'fornecedor', label: 'FORNECEDOR' },
  { key: 'equipamento', label: 'EQUIPAMENTO' },
  { key: 'area', label: 'ÁREA' },
  { key: 'semana', label: 'SEMANA' },
  { key: 'data', label: 'DATA' },
  { key: 'litros', label: 'LITROS' },
  { key: 'valor', label: 'VALOR (R$)' },
] as const;

const EXPORT_FORMAT_OPTIONS = [
  {
    id: 'base',
    title: 'Base Completa',
    description: 'Todos os registros com todas as colunas da tabela tb. abastecimento',
    sheetName: 'BASE DADOS',
  },
  {
    id: 'diretoria',
    title: 'Resumo por Diretoria',
    description: 'Totais de litros e valor agrupados por diretoria',
    sheetName: 'RESUMO DIRETORIA',
  },
  {
    id: 'semana',
    title: 'Resumo por Semana',
    description: 'Totais de consumo agrupados por semana do mês',
    sheetName: 'RESUMO SEMANA',
  },
  {
    id: 'equipamento',
    title: 'Resumo por Equipamento',
    description: 'Consumo total por equipamento, ordenado do maior para o menor',
    sheetName: 'RESUMO EQUIPAMENTO',
  },
  {
    id: 'orcado',
    title: 'Orçado vs Realizado',
    description: 'Comparativo entre orçamento e realizado por diretoria com % de execução',
    sheetName: 'ORÇADO VS REALIZADO',
  },
] as const;

type ExportFormatId = (typeof EXPORT_FORMAT_OPTIONS)[number]['id'];

type BaseExportColumnKey = (typeof BASE_EXPORT_COLUMNS)[number]['key'];

export default function App() {
  // State
  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [currentPage, setCurrentPage] = useState<string>('login');
  const [users, setUsers] = useState<User[]>(initialUsers);
  const [equipments, setEquipments] = useState<Equipment[]>([]);
  const [editingEquipment, setEditingEquipment] = useState<Equipment | null>(null);
  const [rateios, setRateios] = useState<Rateio[]>([]);
  const [budgets, setBudgets] = useState<Budget[]>([]);
  const [abastecimentos, setAbastecimentos] = useState<Abastecimento[]>([]);
  const [dieselPrice, setDieselPrice] = useState<DieselPrice>(initialDieselPrice);
  const [notifications, setNotifications] = useState<Notification[]>([]);
  const [showPassword, setShowPassword] = useState(false);
  const [sidebarVisible, setSidebarVisible] = useState(true);
  
  // Login form state
  const [loginUsername, setLoginUsername] = useState('');
  const [loginPassword, setLoginPassword] = useState('');
  
  // Register form state
  const [showRegister, setShowRegister] = useState(false);
  const [regUsername, setRegUsername] = useState('');
  const [regPassword, setRegPassword] = useState('');
  const [regName, setRegName] = useState('');
  const [regEmail, setRegEmail] = useState('');
  const [regRole, setRegRole] = useState<UserRole>('operator');

  // Forgot password state
  const [showForgotPassword, setShowForgotPassword] = useState(false);
  const [forgotStep, setForgotStep] = useState<'email' | 'code' | 'reset'>('email');
  const [forgotEmail, setForgotEmail] = useState('');
  const [forgotUserId, setForgotUserId] = useState<number | null>(null);
  const [generatedCode, setGeneratedCode] = useState('');
  const [enteredCode, setEnteredCode] = useState('');
  const [codeExpiresAt, setCodeExpiresAt] = useState<number | null>(null);
  const [newPassword, setNewPassword] = useState('');
  const [confirmNewPassword, setConfirmNewPassword] = useState('');
  
  // Filters
  const [filters, setFilters] = useState<FilterState>(createDefaultDashboardFilters());
  
  // History filters
  const [historyDateFilter, setHistoryDateFilter] = useState('');
  const [historyUserFilter, setHistoryUserFilter] = useState('');
  const [selectedObservation, setSelectedObservation] = useState<string | null>(null);
  
  // Diesel price form
  const [newDieselPrice, setNewDieselPrice] = useState('');

  // Export state
  const [exportFormat, setExportFormat] = useState<ExportFormatId>('base');
  const [exportFileName, setExportFileName] = useState('controle_abastecimento');
  const [selectedBaseColumns, setSelectedBaseColumns] = useState<BaseExportColumnKey[]>(
    BASE_EXPORT_COLUMNS.map((column) => column.key),
  );

  // Import preview state
  const [importPreview, setImportPreview] = useState<{ records: Abastecimento[]; fileName: string } | null>(null);

  // Rateio state
  const [showRateioForm, setShowRateioForm] = useState(false);
  const [expandedRateioId, setExpandedRateioId] = useState<number | null>(null);
  const [editingRateio, setEditingRateio] = useState<Rateio | null>(null);

  // Budget global period filter
  const [budgetPeriod, setBudgetPeriod] = useState<{ inicio: string; fim: string }>({ inicio: '', fim: '' });
  const [showBudgetPeriodMenu, setShowBudgetPeriodMenu] = useState(false);

  // Base de Dados tab
  const [databaseTab, setDatabaseTab] = useState<'abastecimentos' | 'equipamentos'>('abastecimentos');
  const [equipmentSearch, setEquipmentSearch] = useState('');
  const [databaseFilters, setDatabaseFilters] = useState<DatabaseFilterState>(createDefaultDatabaseFilters());
  const [editingRecord, setEditingRecord] = useState<Abastecimento | null>(null);
  
  // Add notification
  const addNotification = (type: Notification['type'], message: string) => {
    const id = Date.now();
    setNotifications(prev => [...prev, { id, type, message: cleanText(message) }]);
    setTimeout(() => {
      setNotifications(prev => prev.filter(n => n.id !== id));
    }, 5000);
  };

  // Login handler
  const handleLogin = () => {
    const user = users.find(u => u.username === loginUsername && u.password === loginPassword);
    if (!user) {
      addNotification('error', 'Usuário ou senha inválidos');
      return;
    }
    if (user.status === 'pending') {
      addNotification('warning', 'Sua conta ainda está aguardando aprovação');
      return;
    }
    if (user.status === 'rejected') {
      addNotification('error', 'Sua conta foi rejeitada');
      return;
    }
    setCurrentUser(user);
    setCurrentPage(user.role === 'operator' ? 'preenchimento' : 'dashboard');
    // Supervisor e Admin iniciam no Dashboard, Operador no Preenchimento
    addNotification('success', `Bem-vindo, ${user.name}!`);
    setLoginUsername('');
    setLoginPassword('');
  };

  // Register handler
  const handleRegister = () => {
    if (!regUsername || !regPassword || !regName || !regEmail) {
      addNotification('error', 'Preencha todos os campos');
      return;
    }

    const emailNormalizado = cleanText(regEmail).toLowerCase();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailNormalizado)) {
      addNotification('error', 'Digite um e-mail válido');
      return;
    }
    if (users.find(u => u.username === regUsername)) {
      addNotification('error', 'Nome de usuário já existe');
      return;
    }
    if (users.find(u => (u.email || '').toLowerCase() === emailNormalizado)) {
      addNotification('error', 'E-mail já cadastrado');
      return;
    }
    const newUser: User = {
      id: Date.now(),
      username: cleanText(regUsername),
      password: regPassword,
      role: regRole,
      status: 'pending',
      name: cleanText(regName),
      email: emailNormalizado,
      createdAt: new Date().toISOString()
    };
    setUsers(prev => [...prev, newUser]);
    addNotification('success', 'Cadastro realizado! Aguarde a aprovação do administrador. O e-mail poderá ser usado para recuperação de senha.');
    setShowRegister(false);
    setRegUsername('');
    setRegPassword('');
    setRegName('');
    setRegEmail('');
  };

  // Forgot password handlers
  const handleForgotPasswordEmail = () => {
    const emailNormalizado = cleanText(forgotEmail).toLowerCase();
    if (!emailNormalizado) {
      addNotification('error', 'Digite seu e-mail');
      return;
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailNormalizado)) {
      addNotification('error', 'Digite um e-mail válido');
      return;
    }
    const user = users.find(u => (u.email || '').toLowerCase() === emailNormalizado);
    if (!user) {
      addNotification('error', 'Não encontramos uma conta com este e-mail');
      return;
    }
    if (user.status !== 'approved') {
      addNotification('warning', 'Esta conta ainda não foi aprovada pelo administrador');
      return;
    }

    // Gera código de 6 dígitos
    const code = String(Math.floor(100000 + Math.random() * 900000));
    const expiresAt = Date.now() + 10 * 60 * 1000; // expira em 10 minutos
    setGeneratedCode(code);
    setCodeExpiresAt(expiresAt);
    setForgotUserId(user.id);
    setForgotStep('code');
    setEnteredCode('');

    // Em produção, esse código seria enviado por e-mail real.
    // Aqui simulamos exibindo na notificação e no console.
    console.log(`[Recuperação de senha] Código enviado para ${emailNormalizado}: ${code}`);
    addNotification('success', `Código enviado para ${emailNormalizado}! Código: ${code} (válido por 10 min)`);
  };

  const handleVerifyCode = () => {
    if (!enteredCode) {
      addNotification('error', 'Digite o código recebido');
      return;
    }
    if (enteredCode.length !== 6) {
      addNotification('error', 'O código deve ter 6 dígitos');
      return;
    }
    if (!codeExpiresAt || Date.now() > codeExpiresAt) {
      addNotification('error', 'O código expirou. Solicite um novo.');
      setForgotStep('email');
      return;
    }
    if (enteredCode !== generatedCode) {
      addNotification('error', 'Código incorreto. Tente novamente.');
      return;
    }
    setForgotStep('reset');
    addNotification('success', 'Código verificado com sucesso! Defina sua nova senha.');
  };

  const handleResendCode = () => {
    const user = users.find(u => u.id === forgotUserId);
    if (!user) return;
    const code = String(Math.floor(100000 + Math.random() * 900000));
    const expiresAt = Date.now() + 10 * 60 * 1000;
    setGeneratedCode(code);
    setCodeExpiresAt(expiresAt);
    setEnteredCode('');
    console.log(`[Recuperação de senha] Novo código para ${user.email}: ${code}`);
    addNotification('success', `Novo código enviado! Código: ${code} (válido por 10 min)`);
  };

  const handleResetPassword = () => {
    if (!newPassword || !confirmNewPassword) {
      addNotification('error', 'Preencha os dois campos de senha');
      return;
    }
    if (newPassword.length < 4) {
      addNotification('error', 'A senha deve ter no mínimo 4 caracteres');
      return;
    }
    if (newPassword !== confirmNewPassword) {
      addNotification('error', 'As senhas não coincidem');
      return;
    }
    setUsers(prev => prev.map(u => u.id === forgotUserId ? { ...u, password: newPassword } : u));
    addNotification('success', 'Senha redefinida com sucesso! Faça login com a nova senha.');
    closeForgotPassword();
  };

  const closeForgotPassword = () => {
    setShowForgotPassword(false);
    setForgotStep('email');
    setForgotEmail('');
    setForgotUserId(null);
    setGeneratedCode('');
    setEnteredCode('');
    setCodeExpiresAt(null);
    setNewPassword('');
    setConfirmNewPassword('');
  };

  // Approve/Reject user
  const handleUserApproval = (userId: number, status: UserStatus) => {
    setUsers(prev => prev.map(u => u.id === userId ? { ...u, status } : u));
    addNotification('success', `Usuário ${status === 'approved' ? 'aprovado' : 'rejeitado'} com sucesso!`);
  };

  const handleDeleteUser = (userId: number) => {
    const user = users.find(u => u.id === userId);
    if (!user) return;
    if (user.username === 'admin' || user.name === 'Administrador Principal') {
      addNotification('warning', 'O Administrador Principal não pode ser excluído.');
      return;
    }
    if (!window.confirm(`Tem certeza que deseja excluir o usuário ${user.name}?`)) return;

    setUsers(prev => prev.filter(u => u.id !== userId));
    addNotification('success', 'Usuário excluído com sucesso!');
  };

  // Logout
  const handleLogout = () => {
    setCurrentUser(null);
    setCurrentPage('login');
    addNotification('info', 'Você saiu do sistema');
  };

  // Add equipment
  const handleAddEquipment = (equipment: Omit<Equipment, 'id' | 'createdAt'>) => {
    const newEquipment: Equipment = {
      ...equipment,
      equipment: cleanText(equipment.equipment),
      plate: cleanText(equipment.plate),
      ccNovo: equipment.ccNovo.map(cleanText),
      gerencia: cleanText(equipment.gerencia),
      areaLotacao: cleanText(equipment.areaLotacao),
      area: cleanText(equipment.area),
      fornecedor: cleanText(equipment.fornecedor),
      id: Date.now(),
      createdAt: new Date().toISOString()
    };
    setEquipments(prev => [...prev, newEquipment]);
    setDatabaseTab('equipamentos');
    setCurrentPage('database');
    addNotification('success', 'Equipamento cadastrado! Veja na Base de Dados › Equipamentos Cadastrados.');
  };

  const handleUpdateEquipment = (equipment: Omit<Equipment, 'id' | 'createdAt'>) => {
    if (!editingEquipment) return;

    const updatedEquipment: Equipment = {
      ...editingEquipment,
      ...equipment,
      equipment: cleanText(equipment.equipment),
      plate: cleanText(equipment.plate),
      ccNovo: equipment.ccNovo.map(cleanText).filter(Boolean),
      gerencia: cleanText(equipment.gerencia),
      areaLotacao: cleanText(equipment.areaLotacao),
      area: cleanText(equipment.area),
      fornecedor: cleanText(equipment.fornecedor),
    };

    setEquipments(prev => prev.map(item => item.id === editingEquipment.id ? updatedEquipment : item));
    setEditingEquipment(null);
    addNotification('success', 'Equipamento atualizado com sucesso!');
  };

  const handleDeleteEquipment = (equipmentId: number) => {
    if (!window.confirm('Tem certeza que deseja excluir este equipamento?')) return;
    setEquipments(prev => prev.filter(item => item.id !== equipmentId));
    setRateios(prev => prev.filter(rateio => rateio.equipmentId !== equipmentId));
    if (editingEquipment?.id === equipmentId) setEditingEquipment(null);
    addNotification('success', 'Equipamento excluído com sucesso!');
  };

  // Add rateio
  const handleAddRateio = (rateio: Omit<Rateio, 'id' | 'createdAt'>) => {
    const newRateio: Rateio = {
      ...rateio,
      equipment: cleanText(rateio.equipment),
      description: cleanText(rateio.description),
      items: rateio.items.map((item) => ({
        ...item,
        gerencia: cleanText(item.gerencia),
        ccNovo: cleanText(item.ccNovo),
        description: cleanText(item.description),
      })),
      id: Date.now(),
      createdAt: new Date().toISOString()
    };
    setRateios(prev => [...prev, newRateio]);
    addNotification('success', 'Rateio cadastrado com sucesso!');
  };

  // Add budget
  const handleAddBudget = (budget: Omit<Budget, 'id' | 'createdAt' | 'realizado'>) => {
    const newBudget: Budget = {
      ...budget,
      diretoria: cleanText(budget.diretoria),
      periodo: cleanText(budget.periodo),
      id: Date.now(),
      realizado: 0,
      createdAt: new Date().toISOString()
    };
    setBudgets(prev => [...prev, newBudget]);
    addNotification('success', 'Orçamento cadastrado! Dashboard atualizado automaticamente.');
  };

  // Add abastecimento
  const handleAddAbastecimento = (abastecimento: Omit<Abastecimento, 'id' | 'createdAt' | 'valor' | 'createdBy'>) => {
    const valor = abastecimento.litros * dieselPrice.price;
    const newAbastecimento: Abastecimento = {
      ...abastecimento,
      ccNovo: cleanText(abastecimento.ccNovo),
      diretoria: cleanText(abastecimento.diretoria),
      gerencia: cleanText(abastecimento.gerencia),
      areaLotacao: cleanText(abastecimento.areaLotacao),
      fornecedor: cleanText(abastecimento.fornecedor),
      equipamento: cleanText(abastecimento.equipamento),
      area: cleanText(abastecimento.area),
      semana: cleanText(abastecimento.semana) || deriveWeekLabel(abastecimento.data),
      data: normalizeDateValue(abastecimento.data),
      observacoes: cleanText(abastecimento.observacoes || ''),
      rateioInfo: abastecimento.rateioInfo
        ? {
            ...abastecimento.rateioInfo,
            gerenciaRateio: cleanText(abastecimento.rateioInfo.gerenciaRateio || ''),
            ccRateio: cleanText(abastecimento.rateioInfo.ccRateio || ''),
          }
        : undefined,
      id: Date.now(),
      valor,
      createdBy: cleanText(currentUser?.name || 'Unknown'),
      createdAt: new Date().toISOString()
    };
    setAbastecimentos(prev => [...prev, newAbastecimento]);
    
    // Update budget realizado
    if (newAbastecimento.diretoria) {
      setBudgets(prev => prev.map(b => 
        b.diretoria === newAbastecimento.diretoria 
          ? { ...b, realizado: b.realizado + valor }
          : b
      ));
    }
    
    addNotification('success', 'Abastecimento registrado com sucesso! Dashboard atualizado automaticamente.');
  };

  // Update diesel price
  const handleUpdateDieselPrice = () => {
    const price = parseFloat(newDieselPrice);
    if (isNaN(price) || price <= 0) {
      addNotification('error', 'Preço inválido');
      return;
    }

    const newDiesel = {
      id: 1,
      price,
      updatedAt: new Date().toISOString(),
      updatedBy: currentUser?.username || 'unknown'
    };

    setDieselPrice(newDiesel);

    // Recalcula o valor de todos os abastecimentos com o novo preço
    setAbastecimentos(prev => prev.map(a => ({
      ...a,
      valor: a.litros * price
    })));

    // O Orçado vs Realizado e os KPIs do dashboard já usam o dieselPrice.price 
    // ou o a.valor recalculado acima, então a atualização será automática.

    addNotification('success', `Preço atualizado para ${formatCurrency(price)}! Dashboard recalculado.`);
    setNewDieselPrice('');
  };

  // Import Excel
  const handleImportExcel = (data: any[][], fileName = '') => {
    const headerAliases: Record<string, string[]> = {
      ccNovo: ['CC NOVO', 'CC', 'CENTRO DE CUSTO', 'CENTRO CUSTO'],
      diretoria: ['DIRETORIA'],
      gerencia: ['GERENCIA', 'GERÊNCIA'],
      areaLotacao: ['ÁREA LOT.', 'AREA LOT.', 'ÁREA LOTACAO', 'AREA LOTACAO', 'ÁREA LOTAÇÃO', 'AREA LOTAÇÃO'],
      fornecedor: ['FORNECEDOR'],
      equipamento: ['EQUIPAMENTO'],
      area: ['ÁREA', 'AREA'],
      data: ['DATA'],
      litros: ['LITROS'],
      semana: ['SEMANA'],
    };

    const normalizeHeader = (value: any) =>
      cleanText(String(value || ''))
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toUpperCase();

    const firstRow = data[0] || [];
    const normalizedFirstRow = firstRow.map(normalizeHeader);

    const detectedIndexes = Object.entries(headerAliases).reduce<Record<string, number>>((acc, [field, aliases]) => {
      const index = normalizedFirstRow.findIndex((cell) => aliases.includes(cell));
      if (index >= 0) acc[field] = index;
      return acc;
    }, {});

    const hasHeader = ['ccNovo', 'diretoria', 'gerencia', 'equipamento', 'data', 'litros'].some(
      (field) => detectedIndexes[field] !== undefined,
    );

    const startIndex = hasHeader ? 1 : 0;

    const fallbackIndexes = {
      ccNovo: 0,
      diretoria: 1,
      gerencia: 2,
      areaLotacao: 3,
      fornecedor: 4,
      equipamento: 5,
      area: 6,
      data: 7,
      litros: 8,
      semana: 9,
    } as const;

    const getCellValue = (row: any[], field: keyof typeof fallbackIndexes) => {
      const index = hasHeader ? detectedIndexes[field] ?? fallbackIndexes[field] : fallbackIndexes[field];
      return row[index] ?? '';
    };

    const newAbastecimentos: Abastecimento[] = [];

    for (let i = startIndex; i < data.length; i++) {
      const row = data[i];
      if (!row || row.every((cell) => cleanText(String(cell || '')) === '')) continue;

      const rawDate = getCellValue(row, 'data');
      const normalizedDate = normalizeDateValue(rawDate);
      const litros = parseFloat(String(getCellValue(row, 'litros')).replace(',', '.')) || 0;
      const semanaInformada = cleanText(String(getCellValue(row, 'semana')));
      const semana = semanaInformada || deriveWeekLabel(normalizedDate);

      if (!normalizedDate && !litros) continue;

      const newAbastecimento: Abastecimento = {
        id: Date.now() + i,
        ccNovo: cleanText(String(getCellValue(row, 'ccNovo'))),
        diretoria: cleanText(String(getCellValue(row, 'diretoria'))),
        gerencia: cleanText(String(getCellValue(row, 'gerencia'))),
        areaLotacao: cleanText(String(getCellValue(row, 'areaLotacao'))),
        fornecedor: cleanText(String(getCellValue(row, 'fornecedor'))),
        equipamento: cleanText(String(getCellValue(row, 'equipamento'))),
        area: cleanText(String(getCellValue(row, 'area'))),
        semana,
        data: normalizedDate,
        litros,
        valor: litros * dieselPrice.price,
        createdBy: cleanText(currentUser?.name || 'Importação'),
        createdAt: new Date().toISOString(),
      };

      newAbastecimentos.push(newAbastecimento);
    }

    if (newAbastecimentos.length === 0) {
      addNotification('warning', 'Nenhum registro válido foi encontrado para importar.');
      return;
    }

    setImportPreview({ records: newAbastecimentos, fileName });
    addNotification('info', `${newAbastecimentos.length} registros prontos para pré-visualização. Confirme para importar.`);
  };

  const handleConfirmImport = () => {
    if (!importPreview || importPreview.records.length === 0) {
      addNotification('warning', 'Nenhum registro disponível para importar.');
      return;
    }

    const newAbastecimentos = importPreview.records;

    setAbastecimentos((prev) => [...prev, ...newAbastecimentos]);

    const realizadoPorDiretoria = newAbastecimentos.reduce<Record<string, number>>((acc, item) => {
      if (!item.diretoria) return acc;
      acc[item.diretoria] = (acc[item.diretoria] || 0) + item.valor;
      return acc;
    }, {});

    setBudgets((prev) =>
      prev.map((budget) => ({
        ...budget,
        realizado: budget.realizado + (realizadoPorDiretoria[budget.diretoria] || 0),
      })),
    );

    addNotification('success', `${newAbastecimentos.length} registros importados com sucesso! Dashboard atualizado automaticamente.`);
    setImportPreview(null);
  };

  const abastecimentosEnriquecidos = useMemo(
    () =>
      abastecimentos.map((item) => ({
        ...item,
        ...getDateParts(item.data),
      })),
    [abastecimentos],
  );

  const updateDashboardFilter = (key: keyof FilterState, value: string[]) => {
    setFilters((prev) => ({ ...prev, [key]: value }));
  };

  // Filtered abastecimentos
  const filteredAbastecimentos = useMemo(() => {
    return abastecimentosEnriquecidos.filter(a => {
      if (filters.ccNovo.length > 0 && !filters.ccNovo.includes(a.ccNovo)) return false;
      if (filters.diretoria.length > 0 && !filters.diretoria.includes(a.diretoria)) return false;
      if (filters.gerencia.length > 0 && !filters.gerencia.includes(a.gerencia)) return false;
      if (filters.areaLotacao.length > 0 && !filters.areaLotacao.includes(a.areaLotacao)) return false;
      if (filters.fornecedor.length > 0 && !filters.fornecedor.includes(a.fornecedor)) return false;
      if (filters.equipamento.length > 0 && !filters.equipamento.includes(a.equipamento)) return false;
      if (filters.dia.length > 0 && !filters.dia.includes(a.dia)) return false;
      if (filters.mes.length > 0 && !filters.mes.includes(a.mes)) return false;
      if (filters.semana.length > 0 && !filters.semana.includes(a.semana)) return false;
      if (filters.ano.length > 0 && !filters.ano.includes(a.ano)) return false;
      return true;
    });
  }, [abastecimentosEnriquecidos, filters]);

  // History filtered
  const filteredHistory = useMemo(() => {
    return abastecimentos.filter(a => {
      if (historyDateFilter && a.data !== historyDateFilter) return false;
      if (historyUserFilter && !a.createdBy.toLowerCase().includes(historyUserFilter.toLowerCase())) return false;
      return true;
    });
  }, [abastecimentos, historyDateFilter, historyUserFilter]);

  // Dashboard data
  const dashboardData = useMemo(() => {
    const totalLitros = filteredAbastecimentos.reduce((sum, a) => sum + a.litros, 0);
    const totalValor = filteredAbastecimentos.reduce((sum, a) => sum + a.valor, 0);
    const mediaAbastecimento = filteredAbastecimentos.length > 0 ? totalLitros / filteredAbastecimentos.length : 0;
    const maiorAbastecimento = filteredAbastecimentos.reduce((max, a) => a.litros > max ? a.litros : max, 0);
    
    // Consumption over time
    const consumoTempo = filteredAbastecimentos.reduce((acc, a) => {
      const date = a.data;
      const existing = acc.find(i => i.data === date);
      if (existing) {
        existing.litros += a.litros;
        existing.valor += a.valor;
      } else {
        acc.push({ data: date, litros: a.litros, valor: a.valor });
      }
      return acc;
    }, [] as { data: string; litros: number; valor: number }[]).sort((a, b) => a.data.localeCompare(b.data));

    // Consumption by week
    const semanas = ['Sem 1', 'Sem 2', 'Sem 3', 'Sem 4', 'Sem 5'];
    const consumoSemana = semanas.map(semana => ({
      semana,
      litros: filteredAbastecimentos.filter(a => a.semana === semana).reduce((sum, a) => sum + a.litros, 0)
    }));

    // Distribution by area
    const areas = [...new Set(filteredAbastecimentos.map(a => a.areaLotacao))];
    const distribuicaoArea = areas
      .map(area => ({
        area,
        litros: filteredAbastecimentos.filter(a => a.areaLotacao === area).reduce((sum, a) => sum + a.litros, 0)
      }))
      .sort((a, b) => b.litros - a.litros);

    // Top equipment (with rateio info)
    const equipamentos = [...new Set(filteredAbastecimentos.map(a => a.equipamento))];
    const topEquipamentos = equipamentos.map(eq => {
      const litros = filteredAbastecimentos.filter(a => a.equipamento === eq).reduce((sum, a) => sum + a.litros, 0);
      const rateio = rateios.find(r => r.equipment === eq);
      const gerencias = rateio ? [...new Set(rateio.items.map(i => i.gerencia))] : [
        ...new Set(filteredAbastecimentos.filter(a => a.equipamento === eq).map(a => a.gerencia))
      ];
      const rateioItems = rateio
        ? rateio.items.map(item => ({
            ccNovo: item.ccNovo,
            gerencia: item.gerencia,
            percentage: item.percentage,
            litros: (litros * item.percentage) / 100,
          }))
        : null;

      // Cria chaves dinâmicas no formato "rateio_<index>_<cc>" para empilhar
      const stackedData: Record<string, number> = {};
      if (rateioItems) {
        rateioItems.forEach((item, idx) => {
          stackedData[`rateio_${idx}`] = item.litros;
        });
      } else {
        stackedData['litrosTotal'] = litros;
      }

      return {
        equipamento: eq,
        litros,
        gerencias,
        rateioItems,
        ...stackedData,
      };
    }).sort((a, b) => b.litros - a.litros).slice(0, 5);

    // Encontra o número máximo de itens de rateio para criar barras
    const maxRateioSegments = Math.max(0, ...topEquipamentos.map(e => e.rateioItems?.length || 0));

    // Budget vs Realized by gerência (combina orçamentos cadastrados + abastecimentos realizados)
    const gerenciasFromAbastecimento = filteredAbastecimentos.map(a => a.gerencia).filter(Boolean);
    const gerenciasFromBudget = budgets.map(b => b.diretoria).filter(Boolean); // campo "diretoria" do budget guarda a gerência
    const todasGerencias = [...new Set([...gerenciasFromAbastecimento, ...gerenciasFromBudget])];

    const orcamentoRealizado = todasGerencias.map(ger => {
      const budgetsDaGerencia = budgets.filter(b => b.diretoria === ger);
      const orcamentoTotal = budgetsDaGerencia.reduce((sum, b) => sum + b.orcamento, 0);
      const realizado = filteredAbastecimentos
        .filter(a => a.gerencia === ger)
        .reduce((sum, a) => sum + a.valor, 0);
      return {
        gerencia: ger,
        orcamento: orcamentoTotal,
        realizado,
      };
    }).sort((a, b) => (b.orcamento + b.realizado) - (a.orcamento + a.realizado));

    // Consumo total por gerência somando todos os equipamentos vinculados
    const consumoGerencia = filteredAbastecimentos.reduce((acc, a) => {
      const existing = acc.find((item) => item.gerencia === a.gerencia);
      if (existing) {
        existing.litros += a.litros;
        existing.equipamentos.add(a.equipamento);
      } else {
        acc.push({
          gerencia: a.gerencia,
          litros: a.litros,
          equipamentos: new Set([a.equipamento]),
        });
      }
      return acc;
    }, [] as { gerencia: string; litros: number; equipamentos: Set<string> }[])
      .map((item) => ({
        gerencia: item.gerencia,
        litros: item.litros,
        totalEquipamentos: item.equipamentos.size,
      }))
      .sort((a, b) => b.litros - a.litros);

    // Rateio data for charts
    const rateioData = rateios.map(r => {
      const totalLitrosRateio = filteredAbastecimentos
        .filter(a => a.rateioInfo?.rateioId === r.id)
        .reduce((sum, a) => sum + a.litros, 0);
      return {
        rateio: r.description,
        litros: totalLitrosRateio,
        items: r.items.map(item => ({
          gerencia: item.gerencia,
          cc: item.ccNovo,
          percentage: item.percentage,
          litros: (totalLitrosRateio * item.percentage) / 100
        }))
      };
    });

    return {
      totalLitros,
      totalValor,
      mediaAbastecimento,
      maiorAbastecimento,
      consumoTempo,
      consumoSemana,
      distribuicaoArea,
      topEquipamentos,
      maxRateioSegments,
      orcamentoRealizado,
      consumoGerencia,
      rateioData
    };
  }, [filteredAbastecimentos, budgets, rateios]);

  // Unique values for filters
  const uniqueValues = useMemo(
    () => ({
      ccs: sortTextValues(abastecimentosEnriquecidos.map((a) => a.ccNovo)),
      diretorias: sortTextValues(abastecimentosEnriquecidos.map((a) => a.diretoria)),
      gerencias: sortTextValues(abastecimentosEnriquecidos.map((a) => a.gerencia)),
      areasLotacao: sortTextValues(abastecimentosEnriquecidos.map((a) => a.areaLotacao)),
      fornecedores: sortTextValues(abastecimentosEnriquecidos.map((a) => a.fornecedor)),
      equipamentos: sortTextValues(abastecimentosEnriquecidos.map((a) => a.equipamento)),
      dias: sortNumericTextValues(abastecimentosEnriquecidos.map((a) => a.dia)),
      meses: sortNumericTextValues(abastecimentosEnriquecidos.map((a) => a.mes)),
      semanas: sortTextValues(abastecimentosEnriquecidos.map((a) => a.semana)),
      anos: sortNumericTextValues(abastecimentosEnriquecidos.map((a) => a.ano)),
    }),
    [abastecimentosEnriquecidos],
  );

  const dashboardFilterFields: Array<{
    key: keyof FilterState;
    label: string;
    options: string[];
    formatOption?: (value: string) => string;
  }> = [
    { key: 'ccNovo', label: 'Centro de Custo (CC)', options: uniqueValues.ccs },
    { key: 'diretoria', label: 'Diretoria', options: uniqueValues.diretorias },
    { key: 'gerencia', label: 'Gerência', options: uniqueValues.gerencias },
    { key: 'areaLotacao', label: 'Área Lotação', options: uniqueValues.areasLotacao },
    { key: 'fornecedor', label: 'Fornecedor', options: uniqueValues.fornecedores },
    { key: 'equipamento', label: 'Equipamento', options: uniqueValues.equipamentos },
    { key: 'dia', label: 'Dia', options: uniqueValues.dias },
    { key: 'mes', label: 'Mês', options: uniqueValues.meses, formatOption: getMonthLabel },
    { key: 'semana', label: 'Semana', options: uniqueValues.semanas },
    { key: 'ano', label: 'Ano', options: uniqueValues.anos },
  ];

  const activeExportFilters = dashboardFilterFields.flatMap((field) =>
    filters[field.key].map((value) => ({
      label: field.label,
      value: field.formatOption ? field.formatOption(value) : value,
    })),
  );

  const toggleBaseExportColumn = (columnKey: BaseExportColumnKey) => {
    setSelectedBaseColumns((prev) =>
      prev.includes(columnKey)
        ? prev.filter((item) => item !== columnKey)
        : [...prev, columnKey],
    );
  };

  const exportPreview = useMemo(() => {
    const selectedBaseColumnDefs = BASE_EXPORT_COLUMNS.filter((column) => selectedBaseColumns.includes(column.key));
    const currentFormat = EXPORT_FORMAT_OPTIONS.find((option) => option.id === exportFormat)!;

    if (exportFormat === 'base') {
      const rows = filteredAbastecimentos.map((item) =>
        Object.fromEntries(
          selectedBaseColumnDefs.map((column) => {
            const value = item[column.key as keyof Abastecimento] ?? item[column.key as keyof typeof item];
            const formattedValue = column.key === 'valor'
              ? Number(item.valor).toFixed(2)
              : column.key === 'litros'
                ? Number(item.litros).toFixed(2)
                : value;
            return [column.label, formattedValue];
          }),
        ),
      );

      return {
        rows,
        columns: selectedBaseColumnDefs.map((column) => column.label),
        sheetNames: [currentFormat.sheetName, 'INFORMAÇÕES'],
        records: rows.length,
        litros: filteredAbastecimentos.reduce((sum, item) => sum + item.litros, 0),
        valor: filteredAbastecimentos.reduce((sum, item) => sum + item.valor, 0),
      };
    }

    if (exportFormat === 'diretoria') {
      const grouped = filteredAbastecimentos.reduce<Record<string, { diretoria: string; qtd: number; litros: number; valor: number }>>((acc, item) => {
        const key = item.diretoria || 'Sem Diretoria';
        if (!acc[key]) acc[key] = { diretoria: key, qtd: 0, litros: 0, valor: 0 };
        acc[key].qtd += 1;
        acc[key].litros += item.litros;
        acc[key].valor += item.valor;
        return acc;
      }, {});

      const rows = Object.values(grouped)
        .sort((a, b) => b.litros - a.litros)
        .map((item) => ({
          DIRETORIA: item.diretoria,
          'QTD. REGISTROS': item.qtd,
          'TOTAL LITROS': item.litros.toFixed(2),
          'TOTAL VALOR (R$)': item.valor.toFixed(2),
          'MÉDIA LITROS': (item.litros / item.qtd).toFixed(2),
        }));

      return {
        rows,
        columns: ['DIRETORIA', 'QTD. REGISTROS', 'TOTAL LITROS', 'TOTAL VALOR (R$)', 'MÉDIA LITROS'],
        sheetNames: [currentFormat.sheetName, 'INFORMAÇÕES'],
        records: rows.length,
        litros: filteredAbastecimentos.reduce((sum, item) => sum + item.litros, 0),
        valor: filteredAbastecimentos.reduce((sum, item) => sum + item.valor, 0),
      };
    }

    if (exportFormat === 'semana') {
      const grouped = filteredAbastecimentos.reduce<Record<string, { semana: string; qtd: number; litros: number; valor: number }>>((acc, item) => {
        const key = item.semana || 'Sem Semana';
        if (!acc[key]) acc[key] = { semana: key, qtd: 0, litros: 0, valor: 0 };
        acc[key].qtd += 1;
        acc[key].litros += item.litros;
        acc[key].valor += item.valor;
        return acc;
      }, {});

      const rows = Object.values(grouped)
        .sort((a, b) => a.semana.localeCompare(b.semana, 'pt-BR', { numeric: true }))
        .map((item) => ({
          SEMANA: item.semana,
          'QTD. REGISTROS': item.qtd,
          'TOTAL LITROS': item.litros.toFixed(2),
          'TOTAL VALOR (R$)': item.valor.toFixed(2),
          'MÉDIA LITROS': (item.litros / item.qtd).toFixed(2),
        }));

      return {
        rows,
        columns: ['SEMANA', 'QTD. REGISTROS', 'TOTAL LITROS', 'TOTAL VALOR (R$)', 'MÉDIA LITROS'],
        sheetNames: [currentFormat.sheetName, 'INFORMAÇÕES'],
        records: rows.length,
        litros: filteredAbastecimentos.reduce((sum, item) => sum + item.litros, 0),
        valor: filteredAbastecimentos.reduce((sum, item) => sum + item.valor, 0),
      };
    }

    if (exportFormat === 'equipamento') {
      const grouped = filteredAbastecimentos.reduce<Record<string, { equipamento: string; qtd: number; litros: number; valor: number }>>((acc, item) => {
        const key = item.equipamento || 'Sem Equipamento';
        if (!acc[key]) acc[key] = { equipamento: key, qtd: 0, litros: 0, valor: 0 };
        acc[key].qtd += 1;
        acc[key].litros += item.litros;
        acc[key].valor += item.valor;
        return acc;
      }, {});

      const rows = Object.values(grouped)
        .sort((a, b) => b.litros - a.litros)
        .map((item) => ({
          EQUIPAMENTO: item.equipamento,
          'QTD. ABASTECIMENTOS': item.qtd,
          'TOTAL LITROS': item.litros.toFixed(2),
          'TOTAL VALOR (R$)': item.valor.toFixed(2),
          'MÉDIA LITROS': (item.litros / item.qtd).toFixed(2),
        }));

      return {
        rows,
        columns: ['EQUIPAMENTO', 'QTD. ABASTECIMENTOS', 'TOTAL LITROS', 'TOTAL VALOR (R$)', 'MÉDIA LITROS'],
        sheetNames: [currentFormat.sheetName, 'INFORMAÇÕES'],
        records: rows.length,
        litros: filteredAbastecimentos.reduce((sum, item) => sum + item.litros, 0),
        valor: filteredAbastecimentos.reduce((sum, item) => sum + item.valor, 0),
      };
    }

    const filteredBudgets = budgets.filter((budget) =>
      filters.diretoria.length === 0 || filters.diretoria.includes(budget.diretoria),
    );

    const rows = filteredBudgets.map((budget) => {
      const saldo = budget.orcamento - budget.realizado;
      const execucao = budget.orcamento > 0 ? (budget.realizado / budget.orcamento) * 100 : 0;
      return {
        DIRETORIA: budget.diretoria,
        'ORÇAMENTO (R$)': budget.orcamento.toFixed(2),
        'REALIZADO (R$)': budget.realizado.toFixed(2),
        'SALDO (R$)': saldo.toFixed(2),
        '% EXECUÇÃO': `${execucao.toFixed(1)}%`,
        STATUS: saldo >= 0 ? 'Dentro do Orçamento' : 'Acima do Orçamento',
      };
    });

    return {
      rows,
      columns: ['DIRETORIA', 'ORÇAMENTO (R$)', 'REALIZADO (R$)', 'SALDO (R$)', '% EXECUÇÃO', 'STATUS'],
      sheetNames: [currentFormat.sheetName, 'INFORMAÇÕES'],
      records: rows.length,
      litros: 0,
      valor: filteredBudgets.reduce((sum, item) => sum + item.realizado, 0),
    };
  }, [exportFormat, filteredAbastecimentos, budgets, filters.diretoria, selectedBaseColumns]);

  const databaseFilterFields: Array<{
    key: keyof DatabaseFilterState;
    label: string;
    options: string[];
    formatOption?: (value: string) => string;
  }> = useMemo(
    () => [
      {
        key: 'gerencia',
        label: 'Gerência',
        options: sortTextValues(abastecimentos.map((item) => item.gerencia)),
      },
      {
        key: 'semana',
        label: 'Semana',
        options: sortTextValues(abastecimentos.map((item) => item.semana)),
      },
      {
        key: 'fornecedor',
        label: 'Fornecedor',
        options: sortTextValues(abastecimentos.map((item) => item.fornecedor)),
      },
      {
        key: 'ccNovo',
        label: 'CC Novo',
        options: sortTextValues(abastecimentos.map((item) => item.ccNovo)),
      },
      {
        key: 'equipamento',
        label: 'Equipamento',
        options: sortTextValues(abastecimentos.map((item) => item.equipamento)),
      },
    ],
    [abastecimentos],
  );

  const databaseFilteredAbastecimentos = useMemo(
    () =>
      abastecimentos.filter((item) => {
        if (databaseFilters.gerencia.length > 0 && !databaseFilters.gerencia.includes(item.gerencia)) return false;
        if (databaseFilters.semana.length > 0 && !databaseFilters.semana.includes(item.semana)) return false;
        if (databaseFilters.dataInicio && item.data < databaseFilters.dataInicio) return false;
        if (databaseFilters.dataFim && item.data > databaseFilters.dataFim) return false;
        if (databaseFilters.fornecedor.length > 0 && !databaseFilters.fornecedor.includes(item.fornecedor)) return false;
        if (databaseFilters.ccNovo.length > 0 && !databaseFilters.ccNovo.includes(item.ccNovo)) return false;
        if (databaseFilters.equipamento.length > 0 && !databaseFilters.equipamento.includes(item.equipamento)) return false;
        return true;
      }),
    [abastecimentos, databaseFilters],
  );

  const updateDatabaseFilter = (key: keyof DatabaseFilterState, values: string[]) => {
    setDatabaseFilters((prev) => ({ ...prev, [key]: values }));
  };

  const handleClearAllRecords = () => {
    if (window.confirm('Tem certeza que deseja apagar TODOS os registros? Esta ação não pode ser desfeita.')) {
      setAbastecimentos([]);
      setBudgets(prev => prev.map(b => ({ ...b, realizado: 0 })));
      addNotification('success', 'Base de dados limpa com sucesso!');
    }
  };

  const handleDeleteRecord = (id: number) => {
    const item = abastecimentos.find(a => a.id === id);
    setAbastecimentos(prev => prev.filter(a => a.id !== id));
    if (item && item.diretoria) {
      setBudgets(prev => prev.map(b => 
        b.diretoria === item.diretoria 
          ? { ...b, realizado: Math.max(0, b.realizado - item.valor) }
          : b
      ));
    }
    addNotification('success', 'Registro excluído com sucesso!');
  };

  const handleSaveEditRecord = (updated: Abastecimento) => {
    const oldRecord = abastecimentos.find(a => a.id === updated.id);
    const newValor = updated.litros * dieselPrice.price;
    const saved: Abastecimento = {
      ...updated,
      valor: newValor,
      semana: updated.semana || deriveWeekLabel(updated.data),
    };
    setAbastecimentos(prev => prev.map(a => a.id === saved.id ? saved : a));

    // Ajusta orçamento: subtrai valor antigo, soma valor novo
    if (oldRecord) {
      setBudgets(prev => prev.map(b => {
        let realizado = b.realizado;
        if (b.diretoria === oldRecord.diretoria) realizado = Math.max(0, realizado - oldRecord.valor);
        if (b.diretoria === saved.diretoria) realizado += newValor;
        return { ...b, realizado };
      }));
    }

    setEditingRecord(null);
    addNotification('success', 'Registro atualizado com sucesso!');
  };

  const handleExportExcelFile = () => {
    if (exportPreview.rows.length === 0) {
      addNotification('warning', 'Nenhum dado encontrado com os filtros aplicados.');
      return;
    }

    const workbook = XLSX.utils.book_new();
    const mainSheet = XLSX.utils.json_to_sheet(exportPreview.rows);
    XLSX.utils.book_append_sheet(workbook, mainSheet, exportPreview.sheetNames[0].slice(0, 31));

    const infoRows = [
      ['Exportação de Dados', ''],
      ['Formato', EXPORT_FORMAT_OPTIONS.find((option) => option.id === exportFormat)?.title || ''],
      ['Gerado em', new Date().toLocaleString('pt-BR')],
      ['Usuário', currentUser?.name || ''],
      ['Preço Diesel (R$/L)', dieselPrice.price.toFixed(2)],
      ['Registros', String(exportPreview.records)],
      ['Litros', exportPreview.litros.toFixed(2)],
      ['Valor Total (R$)', exportPreview.valor.toFixed(2)],
      ['Filtros Aplicados', activeExportFilters.length > 0 ? activeExportFilters.map((item) => `${item.label}: ${item.value}`).join(' | ') : 'Nenhum'],
    ];

    const infoSheet = XLSX.utils.aoa_to_sheet(infoRows);
    XLSX.utils.book_append_sheet(workbook, infoSheet, 'INFORMAÇÕES');

    const safeName = cleanText(exportFileName || 'controle_abastecimento').replace(/\s+/g, '_');
    const dateSuffix = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(workbook, `${safeName}_${dateSuffix}.xlsx`);
    addNotification('success', 'Arquivo exportado com sucesso!');
  };

  // Render functions
  const renderLogin = () => (
    <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-slate-900 via-slate-800 to-slate-900 p-4">
      <div className="w-full max-w-md">
        <div className="text-center mb-8">
          <img src="/logo.png" 
            alt="Logo" 
            className="h-full w-full object-contain" 
            />
          <div className="mx-auto mb-6 flex h-20 w-full max-w-72 items-center justify-center rounded-2xl bg-white shadow-lg">
          </div>
          <p className="text-lg text-slate-300">Controle de Abastecimento</p>
          <p className="text-sm text-slate-400">Sistema Corporativo de Gestão de Combustível</p>
        </div>

        <div className="bg-white rounded-2xl shadow-2xl p-8">
          <h2 className="text-lg font-semibold text-slate-800 mb-6 text-center">Entrar no Sistema</h2>
          
          <div className="space-y-4">
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1">Usuário</label>
              <div className="relative">
                <UserIcon className="absolute left-3 top-1/2 -translate-y-1/2 w-5 h-5 text-slate-400" />
                <input
                  type="text"
                  value={loginUsername}
                  onChange={(e) => setLoginUsername(e.target.value)}
                  placeholder="Digite seu usuário"
                  className="w-full pl-10 pr-4 py-3 rounded-lg border border-slate-200 focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none transition"
                />
              </div>
            </div>

            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1">Senha</label>
              <div className="relative">
                <svg className="absolute left-3 top-1/2 -translate-y-1/2 w-5 h-5 text-slate-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z" />
                </svg>
                <input
                  type={showPassword ? 'text' : 'password'}
                  value={loginPassword}
                  onChange={(e) => setLoginPassword(e.target.value)}
                  placeholder="Digite sua senha"
                  className="w-full pl-10 pr-12 py-3 rounded-lg border border-slate-200 focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none transition"
                />
                <button
                  type="button"
                  onClick={() => setShowPassword(!showPassword)}
                  className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                >
                  {showPassword ? <EyeOff className="w-5 h-5" /> : <Eye className="w-5 h-5" />}
                </button>
              </div>
            </div>

            <button
              onClick={handleLogin}
              className="w-full py-3 px-4 bg-red-800 hover:bg-red-900 text-white font-medium rounded-lg transition-colors"
            >
              Entrar
            </button>

            <div className="text-center">
              <button
                type="button"
                onClick={() => setShowForgotPassword(true)}
                className="text-sm text-red-800 hover:text-red-900 font-medium transition"
              >
                Esqueci minha senha
              </button>
            </div>

            <div className="relative my-6">
              <div className="absolute inset-0 flex items-center">
                <div className="w-full border-t border-slate-200"></div>
              </div>
              <div className="relative flex justify-center text-sm">
                <span className="px-2 bg-white text-slate-500">Não tem uma conta?</span>
              </div>
            </div>

            <button
              onClick={() => setShowRegister(true)}
              className="w-full py-3 px-4 border-2 border-red-800 text-red-800 hover:bg-red-800 hover:text-white font-medium rounded-lg transition-colors flex items-center justify-center gap-2"
            >
              <UserIcon className="w-5 h-5" />
              Criar conta
            </button>
          </div>
        </div>

        <p className="text-center text-slate-500 text-sm mt-6">
          Controle de Abastecimento v1.0 — Sistema Corporativo
        </p>
      </div>

      {/* Forgot Password Modal */}
      {showForgotPassword && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl p-8 w-full max-w-md">
            <div className="flex items-center justify-between mb-6">
              <h3 className="text-xl font-semibold text-slate-800">
                {forgotStep === 'email' ? 'Recuperar Senha' : forgotStep === 'code' ? 'Verificar Código' : 'Nova Senha'}
              </h3>
              <button
                onClick={closeForgotPassword}
                className="text-slate-400 hover:text-slate-600 transition"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Steps indicator */}
            <div className="flex items-center justify-center gap-2 mb-6">
              {['email', 'code', 'reset'].map((step, idx) => {
                const stepIndex = ['email', 'code', 'reset'].indexOf(forgotStep);
                const isActive = stepIndex >= idx;
                return (
                  <div key={step} className="flex items-center">
                    <div className={cn(
                      "h-8 w-8 rounded-full flex items-center justify-center text-xs font-bold transition",
                      isActive ? "bg-red-800 text-white" : "bg-slate-200 text-slate-400"
                    )}>
                      {idx + 1}
                    </div>
                    {idx < 2 && (
                      <div className={cn(
                        "w-12 h-0.5 transition",
                        stepIndex > idx ? "bg-red-800" : "bg-slate-200"
                      )} />
                    )}
                  </div>
                );
              })}
            </div>

            {forgotStep === 'email' && (
              <>
                <p className="text-sm text-slate-600 mb-5">
                  Digite o e-mail cadastrado na sua conta. Enviaremos um <strong>código de verificação</strong> para confirmar sua identidade.
                </p>
                <div className="space-y-4">
                  <div>
                    <label className="block text-sm font-medium text-slate-700 mb-1">E-mail cadastrado</label>
                    <div className="relative">
                      <Mail className="absolute left-3 top-1/2 -translate-y-1/2 w-5 h-5 text-slate-400" />
                      <input
                        type="email"
                        value={forgotEmail}
                        onChange={(e) => setForgotEmail(e.target.value)}
                        placeholder="seu@email.com"
                        className="w-full pl-10 pr-4 py-3 rounded-lg border border-slate-200 focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
                      />
                    </div>
                  </div>
                  <div className="flex gap-3 pt-2">
                    <button
                      onClick={handleForgotPasswordEmail}
                      className="flex-1 py-3 px-4 bg-red-800 hover:bg-red-900 text-white font-medium rounded-lg transition-colors"
                    >
                      Enviar Código
                    </button>
                    <button
                      onClick={closeForgotPassword}
                      className="flex-1 py-3 px-4 border border-slate-200 hover:bg-slate-50 text-slate-700 font-medium rounded-lg transition-colors"
                    >
                      Cancelar
                    </button>
                  </div>
                </div>
              </>
            )}

            {forgotStep === 'code' && (
              <>
                <div className="bg-blue-50 border border-blue-200 rounded-lg p-3 mb-5 flex items-start gap-2">
                  <Mail className="w-5 h-5 text-blue-600 shrink-0 mt-0.5" />
                  <div className="text-sm text-blue-800">
                    <p className="font-medium">Código enviado para:</p>
                    <p className="break-all">{forgotEmail}</p>
                    <p className="mt-1 text-xs text-blue-600">Válido por 10 minutos. Verifique sua caixa de entrada.</p>
                  </div>
                </div>
                <div className="space-y-4">
                  <div>
                    <label className="block text-sm font-medium text-slate-700 mb-1">Código de Verificação (6 dígitos)</label>
                    <input
                      type="text"
                      inputMode="numeric"
                      maxLength={6}
                      value={enteredCode}
                      onChange={(e) => setEnteredCode(e.target.value.replace(/\D/g, ''))}
                      placeholder="000000"
                      className="w-full px-4 py-3 rounded-lg border border-slate-200 focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none text-center text-2xl font-bold tracking-[0.5em]"
                    />
                  </div>
                  <button
                    type="button"
                    onClick={handleResendCode}
                    className="text-sm text-red-800 hover:text-red-900 font-medium transition"
                  >
                    Não recebi o código. Reenviar
                  </button>
                  <div className="flex gap-3 pt-2">
                    <button
                      onClick={handleVerifyCode}
                      className="flex-1 py-3 px-4 bg-red-800 hover:bg-red-900 text-white font-medium rounded-lg transition-colors"
                    >
                      Verificar Código
                    </button>
                    <button
                      onClick={() => { setForgotStep('email'); setEnteredCode(''); }}
                      className="flex-1 py-3 px-4 border border-slate-200 hover:bg-slate-50 text-slate-700 font-medium rounded-lg transition-colors"
                    >
                      Voltar
                    </button>
                  </div>
                </div>
              </>
            )}

            {forgotStep === 'reset' && (
              <>
                <div className="bg-green-50 border border-green-200 rounded-lg p-3 mb-5 flex items-start gap-2">
                  <CheckCircle className="w-5 h-5 text-green-600 shrink-0 mt-0.5" />
                  <p className="text-sm text-green-800">
                    Código verificado! Defina abaixo sua nova senha.
                  </p>
                </div>
                <div className="space-y-4">
                  <div>
                    <label className="block text-sm font-medium text-slate-700 mb-1">Nova senha</label>
                    <input
                      type="password"
                      value={newPassword}
                      onChange={(e) => setNewPassword(e.target.value)}
                      placeholder="Mínimo 4 caracteres"
                      className="w-full px-4 py-3 rounded-lg border border-slate-200 focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-700 mb-1">Confirmar nova senha</label>
                    <input
                      type="password"
                      value={confirmNewPassword}
                      onChange={(e) => setConfirmNewPassword(e.target.value)}
                      placeholder="Repita a nova senha"
                      className="w-full px-4 py-3 rounded-lg border border-slate-200 focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
                    />
                  </div>
                  <div className="flex gap-3 pt-2">
                    <button
                      onClick={handleResetPassword}
                      className="flex-1 py-3 px-4 bg-red-800 hover:bg-red-900 text-white font-medium rounded-lg transition-colors"
                    >
                      Redefinir Senha
                    </button>
                    <button
                      onClick={closeForgotPassword}
                      className="flex-1 py-3 px-4 border border-slate-200 hover:bg-slate-50 text-slate-700 font-medium rounded-lg transition-colors"
                    >
                      Cancelar
                    </button>
                  </div>
                </div>
              </>
            )}
          </div>
        </div>
      )}

      {/* Register Modal */}
      {showRegister && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl p-8 w-full max-w-md">
            <h3 className="text-xl font-semibold text-slate-800 mb-6">Criar Nova Conta</h3>
            <div className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Nome Completo</label>
                <input
                  type="text"
                  value={regName}
                  onChange={(e) => setRegName(e.target.value)}
                  className="w-full px-4 py-3 rounded-lg border border-slate-200 focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
                  placeholder="Seu nome"
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Usuário</label>
                <input
                  type="text"
                  value={regUsername}
                  onChange={(e) => setRegUsername(e.target.value)}
                  className="w-full px-4 py-3 rounded-lg border border-slate-200 focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
                  placeholder="Nome de usuário"
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">E-mail</label>
                <input
                  type="email"
                  value={regEmail}
                  onChange={(e) => setRegEmail(e.target.value)}
                  className="w-full px-4 py-3 rounded-lg border border-slate-200 focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
                  placeholder="seu@email.com"
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Senha</label>
                <input
                  type="password"
                  value={regPassword}
                  onChange={(e) => setRegPassword(e.target.value)}
                  className="w-full px-4 py-3 rounded-lg border border-slate-200 focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
                  placeholder="Sua senha"
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Tipo de Conta</label>
                <select
                  value={regRole}
                  onChange={(e) => setRegRole(e.target.value as UserRole)}
                  className="w-full px-4 py-3 rounded-lg border border-slate-200 focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
                >
                  <option value="operator">Operador</option>
                  <option value="supervisor">Supervisor</option>
                </select>
                <p className="text-xs text-slate-400 mt-1">
                  * Contas de Administrador são criadas apenas pelo sistema
                </p>
              </div>
              <div className="flex gap-3 pt-4">
                <button
                  onClick={handleRegister}
                  className="flex-1 py-3 px-4 bg-red-800 hover:bg-red-900 text-white font-medium rounded-lg transition-colors"
                >
                  Criar Conta
                </button>
                <button
                  onClick={() => setShowRegister(false)}
                  className="flex-1 py-3 px-4 border border-slate-200 hover:bg-slate-50 text-slate-700 font-medium rounded-lg transition-colors"
                >
                  Cancelar
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );

  const renderDashboard = () => (
    <div className="space-y-6">
      {/* Filters */}
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
        <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
          <h3 className="text-lg font-semibold text-slate-800 flex items-center gap-2">
            <Filter className="w-5 h-5 text-red-800" />
            Filtros Dinâmicos
          </h3>
          <div className="flex items-center gap-3">
            <span className="text-sm text-slate-500">
              {Object.values(filters).filter((values) => values.length > 0).length}/{Object.keys(filters).length} filtros ativos
            </span>
            <button
              onClick={() => setFilters(createDefaultDashboardFilters())}
              disabled={!Object.values(filters).some((values) => values.length > 0)}
              className="inline-flex items-center gap-2 rounded-lg border border-slate-200 px-3 py-2 text-sm text-slate-600 transition hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-50"
            >
              <X className="h-4 w-4" />
              Limpar filtros
            </button>
          </div>
        </div>
        <div className="grid grid-cols-1 gap-4 md:grid-cols-2 xl:grid-cols-5">
          {dashboardFilterFields.map(({ key, label, options, formatOption }) => (
            <DashboardMultiSelect
              key={key}
              label={label}
              options={options}
              selectedValues={filters[key]}
              onChange={(values) => updateDashboardFilter(key, values)}
              formatOption={formatOption}
            />
          ))}
        </div>
      </div>

      {/* KPI Cards */}
      <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-4">
        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm text-slate-600">Total Mensal (Litros)</p>
              <p className="text-2xl font-bold text-slate-800">{formatNumber(dashboardData.totalLitros, 2)}</p>
              <p className="text-xs text-slate-400">{filteredAbastecimentos.length} registros</p>
            </div>
            <div className="w-10 h-10 rounded-lg bg-blue-50 flex items-center justify-center">
              <Droplets className="w-5 h-5 text-blue-600" />
            </div>
          </div>
        </div>

        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm text-slate-600">Valor Total (R$)</p>
              <p className="text-2xl font-bold text-slate-800">{formatCurrency(dashboardData.totalValor)}</p>
              <p className="text-xs text-slate-400">Preço: {formatCurrency(dieselPrice.price)}/L</p>
            </div>
            <div className="w-10 h-10 rounded-lg bg-amber-50 flex items-center justify-center">
              <DollarSign className="w-5 h-5 text-amber-600" />
            </div>
          </div>
        </div>

        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm text-slate-600">Média por Abastecimento</p>
              <p className="text-2xl font-bold text-slate-800">{formatLiters(dashboardData.mediaAbastecimento)}</p>
              <p className="text-xs text-slate-400">{formatCurrency(dashboardData.mediaAbastecimento * dieselPrice.price)} em média</p>
            </div>
            <div className="w-10 h-10 rounded-lg bg-green-50 flex items-center justify-center">
              <TrendingUp className="w-5 h-5 text-green-600" />
            </div>
          </div>
        </div>

        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm text-slate-600">Maior Abastecimento</p>
              <p className="text-2xl font-bold text-slate-800">{formatLiters(dashboardData.maiorAbastecimento)}</p>
              <p className="text-xs text-slate-400">{formatCurrency(dashboardData.maiorAbastecimento * dieselPrice.price)}</p>
            </div>
            <div className="w-10 h-10 rounded-lg bg-red-50 flex items-center justify-center">
              <Fuel className="w-5 h-5 text-red-600" />
            </div>
          </div>
        </div>
      </div>

      {/* Charts Row 1 */}
      <div className="grid grid-cols-1 gap-6 lg:grid-cols-2">
        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
          <h4 className="text-sm font-semibold text-slate-700 mb-4 flex items-center gap-2">
            <Calendar className="w-4 h-4 text-red-800" />
            Consumo ao Longo do Tempo
          </h4>
          <ResponsiveContainer width="100%" height={280}>
            <AreaChart data={dashboardData.consumoTempo}>
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
              <XAxis dataKey="data" stroke="#64748b" fontSize={12} />
              <YAxis stroke="#64748b" fontSize={12} tickFormatter={(v) => formatCompactNumber(Number(v))} />
              <Tooltip
                contentStyle={{ backgroundColor: '#fff', border: '1px solid #e2e8f0', borderRadius: '8px' }}
                labelStyle={{ color: '#1e293b' }}
                formatter={(value: any, name: any) => {
                  if (name === 'Valor (R$)') return [formatCurrency(Number(value || 0)), name];
                  return [formatLiters(Number(value || 0)), name];
                }}
              />
              <Legend />
              <Area type="monotone" dataKey="litros" stroke="#8B1538" fill="#8B153820" name="Litros" />
              <Area type="monotone" dataKey="valor" stroke="#1E3A8A" fill="#1E3A8A20" name="Valor (R$)" />
            </AreaChart>
          </ResponsiveContainer>
        </div>

        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
          <h4 className="text-sm font-semibold text-slate-700 mb-4 flex items-center gap-2">
            <Wallet className="w-4 h-4 text-red-800" />
            Orçado vs Realizado por Gerência
          </h4>
          {dashboardData.orcamentoRealizado.length > 0 ? (
            <ResponsiveContainer width="100%" height={280}>
              <BarChart data={dashboardData.orcamentoRealizado} margin={{ top: 8, right: 8, left: 8, bottom: 8 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                <XAxis dataKey="gerencia" stroke="#64748b" fontSize={11} tickLine={false} axisLine={false} />
                <YAxis
                  stroke="#64748b"
                  fontSize={12}
                  tickLine={false}
                  axisLine={false}
                  tickFormatter={(v) => formatCompactNumber(Number(v))}
                />
                <Tooltip
                  contentStyle={{ backgroundColor: '#fff', border: '1px solid #e2e8f0', borderRadius: '8px' }}
                  labelStyle={{ color: '#1e293b' }}
                  formatter={(value: any, name: any) => [formatCurrency(Number(value || 0)), name]}
                />
                <Legend />
                <Bar dataKey="orcamento" fill="#94A3B8" name="Orçado" radius={[6, 6, 0, 0]} maxBarSize={48} />
                <Bar dataKey="realizado" fill="#8B1538" name="Realizado" radius={[6, 6, 0, 0]} maxBarSize={48} />
              </BarChart>
            </ResponsiveContainer>
          ) : (
            <p className="py-12 text-center text-slate-400 text-sm">
              Nenhum orçamento ou abastecimento cadastrado ainda
            </p>
          )}
        </div>
      </div>

      {/* Charts Row 2 */}
      <div className="grid grid-cols-1 gap-6 lg:grid-cols-2">
        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6 flex flex-col">
          <h4 className="text-sm font-semibold text-slate-700 mb-4 flex items-center gap-2">
            <Calendar className="w-4 h-4 text-red-800" />
            Consumo por Semana
          </h4>
          <div className="flex-1 min-h-[340px]">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={dashboardData.consumoSemana} margin={{ top: 8, right: 8, left: 8, bottom: 8 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                <XAxis dataKey="semana" stroke="#64748b" fontSize={12} tickLine={false} axisLine={false} />
                <YAxis stroke="#64748b" fontSize={12} tickLine={false} axisLine={false} tickFormatter={(v) => formatCompactNumber(Number(v))} />
                <Tooltip 
                  contentStyle={{ backgroundColor: '#fff', border: '1px solid #e2e8f0', borderRadius: '8px' }}
                  labelStyle={{ color: '#1e293b' }}
                  formatter={(value: any) => [formatLiters(Number(value || 0)), 'Consumo']}
                />
                <Bar dataKey="litros" fill="#8B1538" name="Litros" radius={[8, 8, 0, 0]} maxBarSize={72} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
          <h4 className="text-sm font-semibold text-slate-700 mb-4 flex items-center gap-2">
            <Droplets className="w-4 h-4 text-red-800" />
            Distribuição por Área de Lotação
          </h4>
          {dashboardData.distribuicaoArea.length > 0 ? (
            <div className="grid grid-cols-1 gap-6 xl:grid-cols-[320px,1fr] xl:items-center">
              <div className="relative h-[300px]">
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie
                      data={dashboardData.distribuicaoArea}
                      dataKey="litros"
                      nameKey="area"
                      cx="50%"
                      cy="50%"
                      innerRadius={62}
                      outerRadius={102}
                      paddingAngle={2}
                      stroke="#ffffff"
                      strokeWidth={2}
                    >
                      {dashboardData.distribuicaoArea.map((_, index) => (
                        <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                      ))}
                    </Pie>
                    <Tooltip
                      contentStyle={{ backgroundColor: '#fff', border: '1px solid #e2e8f0', borderRadius: '8px' }}
                      labelStyle={{ color: '#1e293b' }}
                      formatter={(value: any) => [formatLiters(Number(value || 0)), 'Consumo']}
                    />
                  </PieChart>
                </ResponsiveContainer>
                <div className="pointer-events-none absolute inset-0 flex flex-col items-center justify-center text-center">
                  <p className="text-xs font-semibold uppercase tracking-[0.18em] text-slate-400">Total</p>
                  <p className="mt-1 text-2xl font-bold text-slate-900">{formatLiters(dashboardData.totalLitros, 0)}</p>
                  <p className="text-xs text-slate-500">Áreas de lotação</p>
                </div>
              </div>

              <div className="space-y-3 xl:max-h-[300px] xl:overflow-y-auto xl:pr-2">
                {dashboardData.distribuicaoArea.map((item, index) => {
                  const percentual = dashboardData.totalLitros > 0 ? (item.litros / dashboardData.totalLitros) * 100 : 0;
                  return (
                    <div
                      key={`${item.area}-${index}`}
                      className="flex items-center justify-between rounded-xl border border-slate-100 bg-slate-50 px-4 py-3"
                    >
                      <div className="flex min-w-0 items-center gap-3">
                        <span
                          className="h-3.5 w-3.5 shrink-0 rounded-full"
                          style={{ backgroundColor: COLORS[index % COLORS.length] }}
                        />
                        <div className="min-w-0">
                          <p className="truncate text-sm font-medium text-slate-800" title={item.area}>
                            {item.area}
                          </p>
                          <p className="text-xs text-slate-500">{formatNumber(percentual, 2)}% do total</p>
                        </div>
                      </div>
                      <p className="shrink-0 text-sm font-semibold text-slate-900">{formatLiters(item.litros)}</p>
                    </div>
                  );
                })}
              </div>
            </div>
          ) : (
            <p className="py-8 text-center text-slate-400">Nenhum dado disponível</p>
          )}
        </div>
      </div>

      {/* Consumo total por gerência */}
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
        <h4 className="text-sm font-semibold text-slate-700 mb-4 flex items-center gap-2">
          <Wrench className="w-4 h-4 text-red-800" />
          Consumo Total de Equipamentos por Gerência
        </h4>
        <ResponsiveContainer width="100%" height={330}>
          <BarChart data={dashboardData.consumoGerencia.slice(0, 12)} margin={{ top: 8, right: 8, left: 8, bottom: 18 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
            <XAxis
              dataKey="gerencia"
              stroke="#64748b"
              fontSize={11}
              angle={0}
              textAnchor="middle"
              height={42}
              interval={0}
              tickLine={false}
              axisLine={false}
            />
            <YAxis stroke="#64748b" fontSize={12} tickFormatter={(v) => formatCompactNumber(Number(v))} />
            <Tooltip
              contentStyle={{ backgroundColor: '#fff', border: '1px solid #e2e8f0', borderRadius: '8px' }}
              labelStyle={{ color: '#1e293b' }}
              formatter={(value: any, _: any, props: any) => [
                formatLiters(Number(value || 0)),
                `${props?.payload?.totalEquipamentos || 0} equipamento(s)`
              ]}
            />
            <Bar dataKey="litros" fill="#8B1538" name="Litros" radius={[8, 8, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      {/* Top Equipment */}
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
        <h4 className="text-sm font-semibold text-slate-700 mb-4 flex items-center gap-2">
          <TrendingUp className="w-4 h-4 text-red-800" />
          Top Equipamentos — Consumo (Litros)
        </h4>
        {dashboardData.topEquipamentos.length > 0 ? (
          <ResponsiveContainer width="100%" height={360}>
            <BarChart
              data={dashboardData.topEquipamentos}
              layout="vertical"
              barCategoryGap={18}
              margin={{ top: 8, right: 28, left: 28, bottom: 8 }}
            >
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" horizontal={false} />
              <XAxis type="number" stroke="#64748b" fontSize={12} tickLine={false} axisLine={false} tickFormatter={(v) => formatCompactNumber(Number(v))} />
              <YAxis
                type="category"
                dataKey="equipamento"
                tick={(props) => <CategoryAxisTick {...props} anchorX={-12} textAnchor="end" />}
                tickLine={false}
                axisLine={false}
                width={320}
              />
              <Tooltip
                cursor={{ fill: 'rgba(148, 163, 184, 0.15)' }}
                content={({ active, payload }) => {
                  if (!active || !payload?.length) return null;
                  const data = payload[0].payload;
                  return (
                    <div className="max-w-[calc(100vw-2rem)] rounded-xl border border-slate-200 bg-white p-4 shadow-xl sm:min-w-[280px]">
                      <div className="flex items-center gap-2 mb-2">
                        <Fuel className="h-4 w-4 text-blue-500" />
                        <p className="font-semibold text-slate-800 text-sm leading-tight">{data.equipamento}</p>
                      </div>
                      {data.gerencias?.length > 0 && (
                        <div className="flex flex-wrap gap-1 mb-3">
                          {data.gerencias.map((ger: string) => (
                            <span key={ger} className="rounded-full bg-blue-100 px-2 py-0.5 text-xs font-medium text-blue-700">{ger}</span>
                          ))}
                        </div>
                      )}
                      {data.rateioItems ? (
                        <>
                          <p className="text-xs font-semibold text-slate-500 mb-2">Rateio aplicado:</p>
                          <div className="space-y-1">
                            {data.rateioItems.map((item: any, idx: number) => (
                              <div key={idx} className="flex items-center justify-between text-xs">
                                <span className="text-slate-600">{item.ccNovo} ({item.percentage}%)</span>
                                <span className="font-semibold text-slate-800">{formatLiters(item.litros)}</span>
                              </div>
                            ))}
                          </div>
                          <div className="mt-2 border-t border-slate-100 pt-2 flex justify-between text-sm font-bold">
                            <span>Total</span>
                            <span>{formatLiters(Number(data.litros))}</span>
                          </div>
                        </>
                      ) : (
                        <p className="text-sm font-bold text-slate-800">{formatLiters(Number(data.litros))}</p>
                      )}
                    </div>
                  );
                }}
              />
              {/* Barras: se houver rateios, divide em segmentos coloridos */}
              {dashboardData.maxRateioSegments > 0 ? (
                <>
                  {/* Barra de fundo cinza para mostrar total quando não há rateio */}
                  <Bar
                    dataKey="litrosTotal"
                    stackId="rateio"
                    fill="#8B1538"
                    radius={[0, 10, 10, 0]}
                    maxBarSize={42}
                  />
                  {Array.from({ length: dashboardData.maxRateioSegments }).map((_, idx) => {
                    const isLast = idx === dashboardData.maxRateioSegments - 1;
                    const isFirst = idx === 0;
                    return (
                      <Bar
                        key={`rateio_${idx}`}
                        dataKey={`rateio_${idx}`}
                        stackId="rateio"
                        fill={RATEIO_COLORS[idx % RATEIO_COLORS.length]}
                        radius={isLast ? [0, 10, 10, 0] : isFirst ? [10, 0, 0, 10] : [0, 0, 0, 0]}
                        maxBarSize={42}
                      />
                    );
                  })}
                </>
              ) : (
                <Bar dataKey="litros" fill="#8B1538" name="Litros" radius={[0, 10, 10, 0]} maxBarSize={42} />
              )}
            </BarChart>
          </ResponsiveContainer>
        ) : (
          <p className="text-slate-400 text-center py-8">Nenhum dado disponível</p>
        )}

        {/* Legenda dos rateios */}
        {dashboardData.maxRateioSegments > 0 && (
          <div className="mt-4 flex flex-wrap gap-3 justify-center border-t border-slate-100 pt-4">
            <div className="flex items-center gap-2 text-xs">
              <span className="h-3 w-3 rounded-sm" style={{ backgroundColor: '#8B1538' }} />
              <span className="text-slate-600">Sem rateio</span>
            </div>
            {Array.from({ length: dashboardData.maxRateioSegments }).map((_, idx) => (
              <div key={`legend-${idx}`} className="flex items-center gap-2 text-xs">
                <span
                  className="h-3 w-3 rounded-sm"
                  style={{ backgroundColor: RATEIO_COLORS[idx % RATEIO_COLORS.length] }}
                />
                <span className="text-slate-600">Rateio item {idx + 1}</span>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Rateio Charts */}
      {dashboardData.rateioData.some(r => r.litros > 0) && (
        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
          <h4 className="text-sm font-semibold text-slate-700 mb-4 flex items-center gap-2">
            <Scale className="w-4 h-4 text-red-800" />
            Detalhamento de Rateios
          </h4>
          <div className="grid grid-cols-2 gap-6">
            {dashboardData.rateioData.filter(r => r.litros > 0).map(rateio => (
              <div key={rateio.rateio} className="border border-slate-200 rounded-lg p-4">
                <h5 className="font-medium text-slate-800 mb-3">{rateio.rateio}</h5>
                <p className="text-sm text-slate-600 mb-3">Total: {formatLiters(rateio.litros)}</p>
                <div className="space-y-2">
                  {rateio.items.map((item, idx) => (
                    <div key={idx} className="flex items-center justify-between text-sm">
                      <span className="text-slate-600">{item.gerencia} ({item.cc})</span>
                      <span className="font-medium text-slate-800">
                        {formatLiters(item.litros)} ({formatNumber(item.percentage, 2)}%)
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );

  const renderDatabase = () => {
    const filteredEquipments = equipments.filter((eq) => {
      if (!equipmentSearch) return true;
      const search = equipmentSearch.toLowerCase();
      return (
        eq.equipment.toLowerCase().includes(search) ||
        eq.plate.toLowerCase().includes(search) ||
        eq.gerencia.toLowerCase().includes(search) ||
        eq.areaLotacao.toLowerCase().includes(search) ||
        eq.area.toLowerCase().includes(search) ||
        eq.fornecedor.toLowerCase().includes(search) ||
        eq.ccNovo.some((cc) => cc.toLowerCase().includes(search))
      );
    });

    return (
      <div className="space-y-6">
        <div className="bg-white rounded-xl shadow-sm border border-slate-200">
          <div className="border-b border-slate-200 p-4">
            <div className="flex items-center justify-between">
              <div className="flex flex-wrap gap-3">
                <button
                  onClick={() => setDatabaseTab('abastecimentos')}
                  className={cn(
                    'flex items-center gap-2 rounded-lg px-4 py-2 text-sm font-medium transition',
                    databaseTab === 'abastecimentos'
                      ? 'bg-red-50 text-red-800'
                      : 'text-slate-600 hover:bg-slate-50',
                  )}
                >
                  <Database className="h-4 w-4" />
                  Abastecimentos
                  <span className="ml-1 rounded-full bg-white px-2 text-xs text-slate-500 ring-1 ring-slate-200">
                    {abastecimentos.length}
                  </span>
                </button>
                <button
                  onClick={() => setDatabaseTab('equipamentos')}
                  className={cn(
                    'flex items-center gap-2 rounded-lg px-4 py-2 text-sm font-medium transition',
                    databaseTab === 'equipamentos'
                      ? 'bg-red-50 text-red-800'
                      : 'text-slate-600 hover:bg-slate-50',
                  )}
                >
                  <Wrench className="h-4 w-4" />
                  Equipamentos Cadastrados
                  <span className="ml-1 rounded-full bg-white px-2 text-xs text-slate-500 ring-1 ring-slate-200">
                    {equipments.length}
                  </span>
                </button>
              </div>
            </div>
          </div>

          {databaseTab === 'abastecimentos' && (
            <div className="p-6">
              <div className="flex flex-col gap-3 mb-4 md:flex-row md:items-center md:justify-between">
                <h3 className="text-lg font-semibold text-slate-800 flex items-center gap-2">
                  <Database className="w-5 h-5 text-red-800" />
                  Base de Dados - Abastecimentos
                  <span className="text-sm font-normal text-slate-500">
                    {databaseFilteredAbastecimentos.length} / {abastecimentos.length} registros
                  </span>
                </h3>
              </div>

              <div className="mb-6 rounded-xl border border-slate-200 bg-slate-50 p-4">
                <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                  <h4 className="flex items-center gap-2 text-sm font-semibold text-slate-700">
                    <Filter className="h-4 w-4 text-red-800" />
                    Filtros da Base de Dados
                  </h4>
                  <div className="flex items-center gap-3">
                    <span className="text-sm text-slate-500">
                      {Object.values(databaseFilters).filter((values) => values.length > 0).length}/{Object.keys(databaseFilters).length} filtros ativos
                    </span>
                    <button
                      onClick={() => setDatabaseFilters(createDefaultDatabaseFilters())}
                      disabled={!Object.values(databaseFilters).some((values) => values.length > 0)}
                      className="inline-flex items-center gap-2 rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm text-slate-600 transition hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-50"
                    >
                      <X className="h-4 w-4" />
                      Limpar filtros
                    </button>
                    <button
                      onClick={handleClearAllRecords}
                      disabled={abastecimentos.length === 0}
                      className="inline-flex items-center gap-2 rounded-lg border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-700 transition hover:bg-red-100 disabled:cursor-not-allowed disabled:opacity-50"
                    >
                      <Trash2 className="h-4 w-4" />
                      Limpar tudo
                    </button>
                  </div>
                </div>

                <div className="grid grid-cols-1 gap-4 md:grid-cols-2 xl:grid-cols-6 items-end">
                  <div className="xl:col-span-1">
                    <label className="mb-1 block text-xs font-medium text-slate-600">Data Início</label>
                    <input
                      type="date"
                      value={databaseFilters.dataInicio}
                      onChange={(e) => setDatabaseFilters(prev => ({ ...prev, dataInicio: e.target.value }))}
                      className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
                    />
                  </div>
                  <div className="xl:col-span-1">
                    <label className="mb-1 block text-xs font-medium text-slate-600">Data Fim</label>
                    <input
                      type="date"
                      value={databaseFilters.dataFim}
                      onChange={(e) => setDatabaseFilters(prev => ({ ...prev, dataFim: e.target.value }))}
                      className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
                    />
                  </div>
                  {databaseFilterFields.map(({ key, label, options, formatOption }) => (
                    <DashboardMultiSelect
                      key={key}
                      label={label}
                      options={options}
                      selectedValues={databaseFilters[key] as string[]}
                      onChange={(values) => updateDatabaseFilter(key, values)}
                      formatOption={formatOption}
                    />
                  ))}
                </div>
              </div>

              <div className="overflow-x-auto">
                <table className="min-w-[900px] w-full">
                  <thead>
                    <tr className="border-b border-slate-200">
                      <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">ID</th>
                      <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">CC Novo</th>
                      <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Diretoria</th>
                      <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Gerência</th>
                      <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Área Lot.</th>
                      <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Fornecedor</th>
                      <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Equipamento</th>
                      <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Área</th>
                      <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Sem</th>
                      <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Data</th>
                      <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Litros</th>
                      <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Valor (R$)</th>
                      <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Ação</th>
                    </tr>
                  </thead>
                  <tbody>
                    {databaseFilteredAbastecimentos.length > 0 ? (
                      databaseFilteredAbastecimentos.map((a) => (
                        <tr key={a.id} className="border-b border-slate-100 hover:bg-slate-50">
                          <td className="py-3 px-4 text-sm text-red-800 font-medium">{a.id}</td>
                          <td className="py-3 px-4 text-sm text-slate-600">{a.ccNovo}</td>
                          <td className="py-3 px-4 text-sm text-slate-600">{a.diretoria}</td>
                          <td className="py-3 px-4 text-sm text-slate-600">{a.gerencia}</td>
                          <td className="py-3 px-4 text-sm text-slate-600">{a.areaLotacao}</td>
                          <td className="py-3 px-4 text-sm text-slate-600">{a.fornecedor}</td>
                          <td className="py-3 px-4 text-sm text-slate-600">{a.equipamento}</td>
                          <td className="py-3 px-4 text-sm text-slate-600">{a.area}</td>
                          <td className="py-3 px-4 text-sm text-slate-600">{a.semana}</td>
                          <td className="py-3 px-4 text-sm text-slate-600">{a.data}</td>
                      <td className="py-3 px-4 text-sm text-slate-800 font-medium">{formatNumber(a.litros, 2)}</td>
                      <td className="py-3 px-4 text-sm text-slate-800 font-medium">{formatCurrency(a.valor)}</td>
                      <td className="py-3 px-4">
                        <div className="flex items-center gap-2">
                          <button
                            onClick={() => setEditingRecord({ ...a })}
                            className="text-slate-400 hover:text-blue-600 transition"
                            title="Editar registro"
                          >
                            <Pencil className="h-4 w-4" />
                          </button>
                          <button
                            onClick={() => handleDeleteRecord(a.id)}
                            className="text-slate-400 hover:text-red-600 transition"
                            title="Excluir registro"
                          >
                            <Trash2 className="h-4 w-4" />
                          </button>
                        </div>
                      </td>
                    </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={13} className="py-12 text-center text-slate-400">
                          Nenhum registro encontrado
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>

              {/* Edit Modal */}
              {editingRecord && (
                <div className="fixed inset-0 bg-black/50 flex items-center justify-center p-4 z-50">
                  <div className="bg-white rounded-2xl shadow-xl w-full max-w-2xl max-h-[90vh] overflow-y-auto">
                    <div className="border-b border-slate-200 px-6 py-4 flex items-center justify-between">
                      <h3 className="text-lg font-semibold text-slate-800 flex items-center gap-2">
                        <Pencil className="h-5 w-5 text-blue-600" />
                        Editar Registro #{editingRecord.id}
                      </h3>
                      <button
                        onClick={() => setEditingRecord(null)}
                        className="text-slate-400 hover:text-slate-600 transition"
                      >
                        <X className="h-5 w-5" />
                      </button>
                    </div>
                    <div className="px-6 py-5 grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div>
                        <label className="block text-sm font-medium text-slate-700 mb-1">CC Novo</label>
                        <input type="text" value={editingRecord.ccNovo}
                          onChange={e => setEditingRecord({ ...editingRecord, ccNovo: e.target.value })}
                          className="w-full px-3 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none" />
                      </div>
                      <div>
                        <label className="block text-sm font-medium text-slate-700 mb-1">Diretoria</label>
                        <input type="text" value={editingRecord.diretoria}
                          onChange={e => setEditingRecord({ ...editingRecord, diretoria: e.target.value })}
                          className="w-full px-3 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none" />
                      </div>
                      <div>
                        <label className="block text-sm font-medium text-slate-700 mb-1">Gerência</label>
                        <input type="text" value={editingRecord.gerencia}
                          onChange={e => setEditingRecord({ ...editingRecord, gerencia: e.target.value })}
                          className="w-full px-3 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none" />
                      </div>
                      <div>
                        <label className="block text-sm font-medium text-slate-700 mb-1">Área Lotação</label>
                        <input type="text" value={editingRecord.areaLotacao}
                          onChange={e => setEditingRecord({ ...editingRecord, areaLotacao: e.target.value })}
                          className="w-full px-3 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none" />
                      </div>
                      <div>
                        <label className="block text-sm font-medium text-slate-700 mb-1">Fornecedor</label>
                        <input type="text" value={editingRecord.fornecedor}
                          onChange={e => setEditingRecord({ ...editingRecord, fornecedor: e.target.value })}
                          className="w-full px-3 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none" />
                      </div>
                      <div>
                        <label className="block text-sm font-medium text-slate-700 mb-1">Equipamento</label>
                        <input type="text" value={editingRecord.equipamento}
                          onChange={e => setEditingRecord({ ...editingRecord, equipamento: e.target.value })}
                          className="w-full px-3 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none" />
                      </div>
                      <div>
                        <label className="block text-sm font-medium text-slate-700 mb-1">Área</label>
                        <input type="text" value={editingRecord.area}
                          onChange={e => setEditingRecord({ ...editingRecord, area: e.target.value })}
                          className="w-full px-3 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none" />
                      </div>
                      <div>
                        <label className="block text-sm font-medium text-slate-700 mb-1">Semana</label>
                        <input type="text" value={editingRecord.semana}
                          onChange={e => setEditingRecord({ ...editingRecord, semana: e.target.value })}
                          className="w-full px-3 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none" />
                      </div>
                      <div>
                        <label className="block text-sm font-medium text-slate-700 mb-1">Data</label>
                        <input type="date" value={editingRecord.data}
                          onChange={e => setEditingRecord({ ...editingRecord, data: e.target.value })}
                          className="w-full px-3 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none" />
                      </div>
                      <div>
                        <label className="block text-sm font-medium text-slate-700 mb-1">Litros</label>
                        <input type="number" step="0.01" value={editingRecord.litros}
                          onChange={e => setEditingRecord({ ...editingRecord, litros: parseFloat(e.target.value) || 0 })}
                          className="w-full px-3 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none" />
                      </div>
                      <div className="col-span-2 bg-slate-50 rounded-lg p-3 text-sm text-slate-600">
                        Valor calculado: <span className="font-semibold text-slate-800">{formatCurrency((editingRecord.litros || 0) * dieselPrice.price)}</span>
                        <span className="text-xs text-slate-400 ml-2">(Litros × {formatCurrency(dieselPrice.price)}/L)</span>
                      </div>
                    </div>
                    <div className="border-t border-slate-200 px-6 py-4 flex justify-end gap-3">
                      <button
                        onClick={() => setEditingRecord(null)}
                        className="px-5 py-2.5 border border-slate-200 text-slate-700 font-medium rounded-lg hover:bg-slate-50 transition text-sm"
                      >
                        Cancelar
                      </button>
                      <button
                        onClick={() => handleSaveEditRecord(editingRecord)}
                        className="px-5 py-2.5 bg-blue-600 hover:bg-blue-700 text-white font-medium rounded-lg transition text-sm flex items-center gap-2"
                      >
                        <Check className="h-4 w-4" />
                        Salvar alterações
                      </button>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}

          {databaseTab === 'equipamentos' && (
            <div className="p-6">
              <div className="flex flex-col gap-3 mb-4 md:flex-row md:items-center md:justify-between">
                <h3 className="text-lg font-semibold text-slate-800 flex items-center gap-2">
                  <Wrench className="w-5 h-5 text-red-800" />
                  Equipamentos Cadastrados
                  <span className="text-sm font-normal text-slate-500">
                    {filteredEquipments.length} / {equipments.length} equipamentos
                  </span>
                </h3>
                <div className="flex gap-2">
                  <div className="relative">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                    <input
                      type="text"
                      value={equipmentSearch}
                      onChange={(e) => setEquipmentSearch(e.target.value)}
                      placeholder="Buscar equipamento, placa, gerência..."
                      className="w-full rounded-lg border border-slate-200 py-2 pl-10 pr-4 text-sm outline-none focus:border-red-800 focus:ring-2 focus:ring-red-800/20 md:w-72"
                    />
                  </div>
                  <button
                    onClick={() => setCurrentPage('equipments')}
                    className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white font-medium rounded-lg text-sm"
                  >
                    <Plus className="w-4 h-4" />
                    Novo Equipamento
                  </button>
                </div>
              </div>

              {filteredEquipments.length > 0 ? (
                <div className="overflow-x-auto">
                  <table className="min-w-[900px] w-full">
                    <thead>
                      <tr className="border-b border-slate-200">
                        <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">ID</th>
                        <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Equipamento</th>
                        <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Placa</th>
                        <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">CC Novo</th>
                        <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Gerência</th>
                        <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Área Lot.</th>
                        <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Área</th>
                        <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Fornecedor</th>
                        <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Cadastrado em</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredEquipments.map((eq) => (
                        <tr key={eq.id} className="border-b border-slate-100 hover:bg-slate-50">
                          <td className="py-3 px-4 text-sm text-red-800 font-medium">{eq.id}</td>
                          <td className="py-3 px-4 text-sm text-slate-800 font-medium">{eq.equipment}</td>
                          <td className="py-3 px-4 text-sm text-slate-600">{eq.plate || '-'}</td>
                          <td className="py-3 px-4 text-sm text-slate-600">
                            <div className="flex flex-wrap gap-1">
                              {eq.ccNovo.length > 0 ? (
                                eq.ccNovo.map((cc) => (
                                  <span key={cc} className="rounded bg-slate-100 px-2 py-0.5 text-xs">
                                    {cc}
                                  </span>
                                ))
                              ) : (
                                <span className="text-slate-400">-</span>
                              )}
                            </div>
                          </td>
                          <td className="py-3 px-4 text-sm text-slate-600">{eq.gerencia}</td>
                          <td className="py-3 px-4 text-sm text-slate-600">{eq.areaLotacao}</td>
                          <td className="py-3 px-4 text-sm text-slate-600">{eq.area}</td>
                          <td className="py-3 px-4 text-sm text-slate-600">{eq.fornecedor}</td>
                          <td className="py-3 px-4 text-sm text-slate-500">
                            {new Date(eq.createdAt).toLocaleDateString('pt-BR')}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ) : (
                <div className="py-12 text-center">
                  <div className="mx-auto mb-4 flex h-16 w-16 items-center justify-center rounded-full bg-slate-100">
                    <Wrench className="h-8 w-8 text-slate-400" />
                  </div>
                  <p className="text-slate-600 font-medium mb-1">
                    {equipments.length === 0
                      ? 'Nenhum equipamento cadastrado'
                      : 'Nenhum equipamento encontrado com a busca'}
                  </p>
                  <p className="text-sm text-slate-400 mb-4">
                    {equipments.length === 0
                      ? 'Cadastre um novo equipamento na aba Equipamentos'
                      : 'Tente ajustar o termo da busca'}
                  </p>
                  {equipments.length === 0 && (
                    <button
                      onClick={() => setCurrentPage('equipments')}
                      className="px-6 py-2 bg-blue-600 hover:bg-blue-700 text-white font-medium rounded-lg inline-flex items-center gap-2"
                    >
                      <Plus className="w-4 h-4" />
                      Novo Equipamento
                    </button>
                  )}
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    );
  };

  const renderBudget = () => {
    const budgetFilteredAbastecimentos = filteredAbastecimentos.filter((a) => {
      if (budgetPeriod.inicio && a.data < budgetPeriod.inicio) return false;
      if (budgetPeriod.fim && a.data > budgetPeriod.fim) return false;
      return true;
    });

    const periodLabel = budgetPeriod.inicio || budgetPeriod.fim
      ? `${budgetPeriod.inicio || '...'} → ${budgetPeriod.fim || '...'}`
      : 'Período Global';

    return (
    <div className="space-y-6">
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
        <div className="flex items-center justify-between mb-6">
          <div>
            <h3 className="text-lg font-semibold text-slate-800 flex items-center gap-2">
              <Wallet className="w-5 h-5 text-red-800" />
              Orçamento por Gerência
            </h3>
            <p className="text-sm text-slate-500">Gerencie orçamentos e períodos de vigência</p>
          </div>
          <div className="relative">
            <button
              onClick={() => setShowBudgetPeriodMenu(v => !v)}
              className="flex items-center gap-2 px-4 py-2 border border-slate-200 rounded-lg text-sm text-slate-600 hover:bg-slate-50"
            >
              <Calendar className="w-4 h-4" />
              {periodLabel}
              <ChevronDown className={cn('w-4 h-4 transition', showBudgetPeriodMenu && 'rotate-180')} />
            </button>
            {showBudgetPeriodMenu && (
              <div className="absolute right-0 z-30 mt-2 w-[calc(100vw-2rem)] max-w-80 rounded-xl border border-slate-200 bg-white p-4 shadow-xl">
                <p className="text-xs font-semibold text-slate-500 uppercase mb-3">Filtrar por Período</p>
                <div className="space-y-3">
                  <div>
                    <label className="text-xs font-medium text-slate-600 mb-1 block">Data Início</label>
                    <input
                      type="date"
                      value={budgetPeriod.inicio}
                      onChange={(e) => setBudgetPeriod(prev => ({ ...prev, inicio: e.target.value }))}
                      className="w-full px-3 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
                    />
                  </div>
                  <div>
                    <label className="text-xs font-medium text-slate-600 mb-1 block">Data Fim</label>
                    <input
                      type="date"
                      value={budgetPeriod.fim}
                      onChange={(e) => setBudgetPeriod(prev => ({ ...prev, fim: e.target.value }))}
                      className="w-full px-3 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
                    />
                  </div>
                </div>
                <div className="flex gap-2 mt-4">
                  <button
                    onClick={() => setBudgetPeriod({ inicio: '', fim: '' })}
                    className="flex-1 px-3 py-2 border border-slate-200 rounded-lg text-sm text-slate-600 hover:bg-slate-50"
                  >
                    Limpar
                  </button>
                  <button
                    onClick={() => setShowBudgetPeriodMenu(false)}
                    className="flex-1 px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-sm font-medium"
                  >
                    Aplicar
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>

        <div className="mb-6 grid grid-cols-1 gap-4 sm:grid-cols-2 xl:grid-cols-3">
          <div className="bg-blue-50 rounded-lg p-4">
            <p className="text-sm text-blue-600 mb-1">Total Orçado</p>
            <p className="text-2xl font-bold text-blue-800">
              {formatCurrency(budgets.reduce((sum, b) => sum + b.orcamento, 0))}
            </p>
          </div>
          <div className="bg-amber-50 rounded-lg p-4">
            <p className="text-sm text-amber-600 mb-1">Total Realizado</p>
            <p className="text-2xl font-bold text-amber-800">
              {formatCurrency(
                budgets.reduce((sum, b) => {
                  const realizadoNaBase = budgetFilteredAbastecimentos
                    .filter(a => a.gerencia === b.diretoria)
                    .reduce((acc, curr) => acc + curr.valor, 0);
                  return sum + realizadoNaBase;
                }, 0)
              )}
            </p>
          </div>
          <div className="bg-green-50 rounded-lg p-4">
            <p className="text-sm text-green-600 mb-1">Situação</p>
            <p className="text-lg font-bold text-green-800">Dentro do Orçamento</p>
            <p className="text-xs text-green-600">
              {budgets.reduce((sum, b) => sum + b.orcamento, 0) > 0 
                ? `${formatNumber((budgets.reduce((sum, b) => sum + b.realizado, 0) / budgets.reduce((sum, b) => sum + b.orcamento, 0)) * 100, 2)}% executado`
                : '0% executado'}
            </p>
          </div>
        </div>

        <div className="overflow-x-auto mb-8">
          <table className="min-w-[760px] w-full">
            <thead>
              <tr className="border-b border-slate-200">
                <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Gerência</th>
                <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Período</th>
                <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Orçamento (R$)</th>
                <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Realizado (R$)</th>
                <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">% Exec.</th>
                <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Status</th>
                <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Ações</th>
              </tr>
            </thead>
            <tbody>
              {budgets.length > 0 ? (
                budgets.map(b => (
                  <tr key={b.id} className="border-b border-slate-100">
                    <td className="py-3 px-4 text-sm text-slate-800">{b.diretoria}</td>
                    <td className="py-3 px-4 text-sm text-slate-600">{b.periodo}</td>
                    <td className="py-3 px-4 text-sm text-slate-800 font-medium">{formatCurrency(b.orcamento)}</td>
                    <td className="py-3 px-4 text-sm text-slate-800 font-medium">
                      {(() => {
                        const realizadoTotal = budgetFilteredAbastecimentos
                          .filter(a => a.gerencia === b.diretoria)
                          .reduce((acc, curr) => acc + curr.valor, 0);
                        return formatCurrency(realizadoTotal);
                      })()}
                    </td>
                    <td className="py-3 px-4 text-sm text-slate-600">
                      {(() => {
                        const realizadoTotal = budgetFilteredAbastecimentos
                          .filter(a => a.gerencia === b.diretoria)
                          .reduce((acc, curr) => acc + curr.valor, 0);
                        return b.orcamento > 0 ? formatNumber((realizadoTotal / b.orcamento) * 100, 2) : '0,00';
                      })()}%
                    </td>
                    <td className="py-3 px-4">
                      {(() => {
                        const realizadoTotal = budgetFilteredAbastecimentos
                          .filter(a => a.gerencia === b.diretoria)
                          .reduce((acc, curr) => acc + curr.valor, 0);
                        const isDentro = realizadoTotal <= b.orcamento;
                        return (
                          <span className={cn(
                            "px-2 py-1 rounded-full text-xs font-medium",
                            isDentro ? "bg-green-100 text-green-800" : "bg-red-100 text-red-800"
                          )}>
                            {isDentro ? 'Dentro' : 'Acima'}
                          </span>
                        );
                      })()}
                    </td>
                    <td className="py-3 px-4">
                      <button className="text-slate-400 hover:text-red-800">
                        <Settings className="w-4 h-4" />
                      </button>
                    </td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={7} className="py-12 text-center text-slate-400">
                    Nenhum orçamento cadastrado
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>

        <div className="border-t border-slate-200 pt-6">
          <h4 className="text-sm font-semibold text-slate-700 mb-4 flex items-center gap-2">
            <Plus className="w-4 h-4 text-red-800" />
            Adicionar Nova Gerência
          </h4>
          <BudgetForm onAdd={handleAddBudget} />
        </div>
      </div>
    </div>
    );
  };

  const BudgetForm = ({ onAdd }: { onAdd: (b: any) => void }) => {
    const [diretoria, setDiretoria] = useState('');
    const [orcamento, setOrcamento] = useState('');
    const [dataInicio, setDataInicio] = useState('');
    const [dataFim, setDataFim] = useState('');

    const handleSubmit = () => {
      if (!diretoria || !orcamento) return;
      onAdd({
        diretoria,
        periodo: `${dataInicio} - ${dataFim}`,
        orcamento: parseFloat(orcamento) || 0,
        dataInicio,
        dataFim
      });
      setDiretoria('');
      setOrcamento('');
      setDataInicio('');
      setDataFim('');
    };

    return (
      <div className="flex flex-col gap-4 md:flex-row md:items-end">
        <div className="w-full flex-1">
          <label className="block text-xs font-medium text-slate-600 mb-1">Gerência *</label>
          <select
            value={diretoria}
            onChange={(e) => setDiretoria(e.target.value)}
            className="w-full px-4 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
          >
            <option value="">Selecione a gerência...</option>
            {sortTextValues([
              ...abastecimentos.map((item) => item.gerencia),
              ...budgets.map((item) => item.diretoria),
            ]).map((gerencia) => (
              <option key={gerencia} value={gerencia}>{gerencia}</option>
            ))}
          </select>
        </div>
        <div className="w-full md:w-36">
          <label className="block text-xs font-medium text-slate-600 mb-1">Orçamento (R$) *</label>
          <input
            type="number"
            value={orcamento}
            onChange={(e) => setOrcamento(e.target.value)}
            placeholder="Ex: 85000"
            className="w-full px-4 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
          />
        </div>
        <div className="w-full md:w-36">
          <label className="block text-xs font-medium text-slate-600 mb-1">Data Início</label>
          <input
            type="date"
            value={dataInicio}
            onChange={(e) => setDataInicio(e.target.value)}
            className="w-full px-4 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
          />
        </div>
        <div className="w-full md:w-36">
          <label className="block text-xs font-medium text-slate-600 mb-1">Data Fim</label>
          <input
            type="date"
            value={dataFim}
            onChange={(e) => setDataFim(e.target.value)}
            className="w-full px-4 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
          />
        </div>
        <button
          onClick={handleSubmit}
          className="flex w-full items-center justify-center gap-2 rounded-lg bg-blue-600 px-6 py-2 font-medium text-white hover:bg-blue-700 md:w-auto"
        >
          <Plus className="w-4 h-4" />
          Adicionar
        </button>
      </div>
    );
  };

  const renderRateio = () => {
    // Total de litros de um equipamento a partir dos abastecimentos
    const litrosPorEquipamento = (equipName: string) =>
      abastecimentos
        .filter(a => a.equipamento === equipName)
        .reduce((sum, a) => sum + a.litros, 0);

    const allGerencias = (rateio: Rateio) => [
      ...new Set(rateio.items.map(i => i.gerencia).filter(Boolean))
    ];

    return (
      <div className="space-y-6">
        {/* Header */}
        <div className="rounded-xl border border-slate-200 bg-white p-6 shadow-sm">
          <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
            <div className="flex min-w-0 items-center gap-4">
              <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-blue-50 text-blue-600">
                <Scale className="h-6 w-6" />
              </div>
              <div>
                <h3 className="text-lg font-semibold text-slate-800 sm:whitespace-nowrap">Rateio de Centro de Custo</h3>
                <p className="text-sm text-slate-500">Divida o custo de um equipamento entre múltiplos CCs por porcentagem</p>
              </div>
            </div>
            <div className="flex w-full flex-col gap-3 sm:w-auto sm:flex-row sm:items-center">
              <span className="flex items-center gap-1.5 rounded-full bg-slate-100 px-3 py-1.5 text-sm text-slate-600">
                <CheckCircle className="h-4 w-4 text-green-600" />
                {rateios.length} ativo{rateios.length !== 1 ? 's' : ''}
              </span>
              <button
                onClick={() => {
                  setEditingRateio(null);
                  setShowRateioForm(f => !f);
                }}
                className="flex w-full items-center justify-center gap-2 rounded-xl bg-blue-600 px-5 py-2.5 text-sm font-semibold text-white shadow transition hover:bg-blue-700 sm:w-auto"
              >
                <Plus className="h-4 w-4" />
                Novo Rateio
              </button>
            </div>
          </div>
        </div>

        {/* Form */}
        {showRateioForm && (
          <div className="rounded-xl border border-blue-200 bg-white p-6 shadow-sm">
            <div className="mb-5 flex items-center justify-between">
              <h4 className="flex items-center gap-2 text-sm font-semibold text-slate-700">
                {editingRateio ? <Pencil className="h-4 w-4 text-blue-600" /> : <Plus className="h-4 w-4 text-blue-600" />}
                {editingRateio ? 'Editar Rateio' : 'Novo Rateio'}
              </h4>
              <button onClick={() => {
                setShowRateioForm(false);
                setEditingRateio(null);
              }} className="text-slate-400 hover:text-slate-600">
                <X className="h-5 w-5" />
              </button>
            </div>
            <RateioForm
              equipments={equipments}
              initialRateio={editingRateio}
              onAdd={(r: any) => {
                if (editingRateio) {
                  setRateios(prev => prev.map(item => item.id === editingRateio.id ? {
                    ...editingRateio,
                    ...r,
                    equipment: cleanText(r.equipment),
                    description: cleanText(r.description),
                    items: r.items.map((rateioItem: RateioItem) => ({
                      ...rateioItem,
                      gerencia: cleanText(rateioItem.gerencia),
                      ccNovo: cleanText(rateioItem.ccNovo),
                      description: cleanText(rateioItem.description),
                    })),
                  } : item));
                  addNotification('success', 'Rateio atualizado com sucesso!');
                } else {
                  handleAddRateio(r);
                }
                setShowRateioForm(false);
                setEditingRateio(null);
              }}
            />
          </div>
        )}

        {/* Rateios configurados */}
        {rateios.length > 0 && (
          <div className="space-y-3">
            <p className="text-xs font-semibold uppercase tracking-[0.18em] text-slate-500 px-1">
              Rateios Configurados ({rateios.length})
            </p>
            {rateios.map(r => {
              const totalLitros = litrosPorEquipamento(r.equipment);
              const isExpanded = expandedRateioId === r.id;

              return (
                <div key={r.id} className="rounded-xl border border-slate-200 bg-white shadow-sm">
                  <div className="flex items-center gap-4 px-5 py-4">
                    <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-blue-50 text-blue-500">
                      <Scale className="h-5 w-5" />
                    </div>
                    <div className="flex-1 min-w-0">
                      <p className="font-semibold text-slate-800 truncate">{r.equipment}</p>
                      <div className="mt-1 flex flex-wrap items-center gap-2">
                        {allGerencias(r).map(ger => (
                          <span key={ger} className="rounded-full bg-blue-100 px-2 py-0.5 text-xs font-medium text-blue-700">
                            {ger}
                          </span>
                        ))}
                        {r.items.map((item, idx) => (
                          <span key={idx} className="text-xs text-slate-500">
                            {item.ccNovo} → <span className="font-semibold text-red-700">{item.percentage}%</span>
                          </span>
                        ))}
                      </div>
                    </div>
                    <div className="flex items-center gap-2 shrink-0">
                      <button
                        onClick={() => setExpandedRateioId(isExpanded ? null : r.id)}
                        className="rounded-lg p-2 text-slate-400 hover:bg-slate-100 hover:text-blue-600 transition"
                        title={isExpanded ? 'Recolher' : 'Expandir'}
                      >
                        <Eye className="h-4 w-4" />
                      </button>
                      <button
                        onClick={() => {
                          setEditingRateio(r);
                          setShowRateioForm(true);
                          setExpandedRateioId(null);
                        }}
                        className="rounded-lg p-2 text-slate-400 hover:bg-blue-50 hover:text-blue-600 transition"
                        title="Editar rateio"
                      >
                        <Pencil className="h-4 w-4" />
                      </button>
                      <button
                        onClick={() => {
                          setRateios(prev => prev.filter(x => x.id !== r.id));
                          addNotification('success', 'Rateio removido com sucesso!');
                        }}
                        className="rounded-lg p-2 text-slate-400 hover:bg-red-50 hover:text-red-600 transition"
                        title="Excluir"
                      >
                        <Trash2 className="h-4 w-4" />
                      </button>
                      <button
                        onClick={() => setExpandedRateioId(isExpanded ? null : r.id)}
                        className="rounded-lg p-2 text-slate-400 hover:bg-slate-100 transition"
                      >
                        <ChevronDown className={cn('h-4 w-4 transition-transform', isExpanded && 'rotate-180')} />
                      </button>
                    </div>
                  </div>

                  {isExpanded && (
                    <div className="border-t border-slate-100 px-5 pb-5 pt-4">
                      <div className="overflow-x-auto">
                        <table className="min-w-[820px] w-full">
                          <thead>
                            <tr className="border-b border-slate-100 text-left text-xs font-semibold uppercase tracking-[0.14em] text-slate-500">
                              <th className="pb-2 pr-4">Equipamento</th>
                              <th className="pb-2 pr-4">Gerência</th>
                              <th className="pb-2 pr-4">CC NOVO</th>
                              <th className="pb-2 pr-4 text-center">%</th>
                              <th className="pb-2 pr-4 text-right">Litros Rateados</th>
                              <th className="pb-2 text-right">Valor Rateado</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-50">
                            {r.items.map((item, idx) => {
                              const litrosRateados = (totalLitros * item.percentage) / 100;
                              const valorRateado = litrosRateados * dieselPrice.price;
                              return (
                                <tr key={idx} className="text-sm">
                                  <td className="py-3 pr-4">
                                    <div className="flex items-center gap-2">
                                      <Fuel className="h-4 w-4 shrink-0 text-blue-500" />
                                      <span className="text-slate-700 font-medium">{r.equipment}</span>
                                    </div>
                                  </td>
                                  <td className="py-3 pr-4">
                                    <span className="rounded-full bg-blue-100 px-2.5 py-1 text-xs font-medium text-blue-700">
                                      {item.gerencia || '-'}
                                    </span>
                                  </td>
                                  <td className="py-3 pr-4 text-slate-600">{item.ccNovo || '-'}</td>
                                  <td className="py-3 pr-4 text-center">
                                    <span className="rounded-full bg-slate-100 px-2.5 py-1 text-xs font-semibold text-slate-700">
                                      {item.percentage}%
                                    </span>
                                  </td>
                                  <td className="py-3 pr-4 text-right font-semibold text-slate-800">
                                    {formatLiters(litrosRateados)}
                                  </td>
                                  <td className="py-3 text-right font-semibold text-slate-800">
                                    {formatCurrency(valorRateado)}
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}


      </div>
    );
  };

  const RateioForm = ({ equipments, onAdd, initialRateio }: { equipments: Equipment[], onAdd: (r: any) => void, initialRateio?: Rateio | null }) => {
    const [equipment, setEquipment] = useState(initialRateio?.equipment || '');
    const [description, setDescription] = useState(initialRateio?.description || '');
    const [items, setItems] = useState<RateioItem[]>(
      initialRateio?.items?.length
        ? initialRateio.items.map(item => ({ ...item }))
        : [
            { gerencia: '', ccNovo: '', description: '', percentage: 50 },
            { gerencia: '', ccNovo: '', description: '', percentage: 50 }
          ]
    );

    // Busca informações do equipamento em ambas as fontes
    const equipmentInfo = useMemo(() => {
      const eq = equipments.find(e => e.equipment === equipment);
      const histItems = abastecimentos.filter(a => a.equipamento === equipment);
      
      return {
        gerencia: eq?.gerencia || histItems[0]?.gerencia || '',
        ccs: [
          ...new Set([
            ...(eq?.ccNovo || []),
            ...histItems.map(h => h.ccNovo)
          ])
        ].filter(Boolean)
      };
    }, [equipment, equipments, abastecimentos]);

    const equipmentOptions = useMemo(
      () => sortTextValues([
        ...equipments.map(e => e.equipment),
        ...abastecimentos.map(a => a.equipamento),
      ]),
      [equipments, abastecimentos],
    );

    const gerenciaOptions = useMemo(
      () => sortTextValues([
        ...abastecimentos.map(a => a.gerencia),
        ...equipments.map(e => e.gerencia),
      ]),
      [abastecimentos, equipments],
    );

    const getCcsByGerencia = (gerencia: string) => {
      const ccsFromAbastecimentos = abastecimentos
        .filter(a => a.gerencia === gerencia)
        .map(a => a.ccNovo);

      const ccsFromEquipments = equipments
        .filter(e => e.gerencia === gerencia)
        .flatMap(e => e.ccNovo);

      return sortTextValues([...ccsFromAbastecimentos, ...ccsFromEquipments]);
    };

    const totalPercentage = items.reduce((sum, item) => sum + item.percentage, 0);

    const addItem = () => {
      setItems(prev => [...prev, { 
        gerencia: equipmentInfo.gerencia, 
        ccNovo: '', 
        description: '', 
        percentage: 0 
      }]);
    };

    const updateItem = (index: number, field: keyof RateioItem, value: any) => {
      setItems(prev => prev.map((item, i) => i === index ? { ...item, [field]: value } : item));
    };

    const removeItem = (index: number) => {
      setItems(prev => prev.filter((_, i) => i !== index));
    };

    const handleSubmit = () => {
      if (!equipment || totalPercentage !== 100) return;
      onAdd({
        equipmentId: equipments.find(e => e.equipment === equipment)?.id || 0,
        equipment,
        description,
        items
      });
      setEquipment('');
      setDescription('');
      setItems([
        { gerencia: '', ccNovo: '', description: '', percentage: 50 },
        { gerencia: '', ccNovo: '', description: '', percentage: 50 }
      ]);
    };

    const submitButtonText = initialRateio ? 'Salvar Alterações' : 'Salvar Rateio';

    return (
      <div className="space-y-6">
        <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Equipamento *</label>
            <select
              value={equipment}
              onChange={(e) => {
                const val = e.target.value;
                setEquipment(val);
                // Procura primeiro nos equipamentos cadastrados, depois na base histórica
                const eq = equipments.find(e => e.equipment === val);
                const hist = abastecimentos.find(a => a.equipamento === val);
                
                const gerenciaEncontrada = eq?.gerencia || hist?.gerencia || '';
                
                if (gerenciaEncontrada) {
                  const ccsDaGerencia = getCcsByGerencia(gerenciaEncontrada);
                  setItems(prev => prev.map(item => ({
                    ...item,
                    gerencia: gerenciaEncontrada,
                    ccNovo: ccsDaGerencia.includes(item.ccNovo) ? item.ccNovo : (ccsDaGerencia[0] || '')
                  })));
                }
              }}
              className="w-full px-4 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
            >
              <option value="">Selecione o equipamento...</option>
              {equipmentOptions.map(eqName => (
                <option key={eqName} value={eqName}>{eqName}</option>
              ))}
            </select>
            {equipmentOptions.length === 0 && (
              <p className="text-xs text-amber-600 mt-1 flex items-center gap-1">
                <AlertCircle className="w-3 h-3" />
                Nenhum equipamento na base ainda
              </p>
            )}
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Descrição do Rateio</label>
            <input
              type="text"
              value={description}
              onChange={(e) => setDescription(e.target.value)}
              placeholder="Ex: Rateio Mina A / Mina B"
              className="w-full px-4 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
            />
          </div>
        </div>

        <div>
          <div className="flex items-center justify-between mb-3">
            <p className="text-sm text-slate-600">Gerências, Centros de Custo e Porcentagens</p>
            <div className="flex gap-2">
              <button onClick={() => {
                const equal = 100 / items.length;
                setItems(prev => prev.map(item => ({ ...item, percentage: Math.round(equal) })));
              }} className="text-xs text-blue-600 hover:text-blue-800">
                Dividir Igualmente
              </button>
              <button onClick={addItem} className="text-xs px-3 py-1 bg-blue-600 text-white rounded">
                + Adicionar Gerência/CC
              </button>
            </div>
          </div>

          <div className="space-y-4">
            {items.map((item, index) => (
              <div key={index} className="bg-slate-50 rounded-lg p-4 border border-slate-200">
                <div className="flex items-center gap-2 mb-3">
                  <div className="w-6 h-6 rounded-full bg-red-100 flex items-center justify-center text-red-800 text-sm font-medium">
                    {index + 1}
                  </div>
                  <button onClick={() => removeItem(index)} className="text-slate-400 hover:text-red-600 ml-auto">
                    <X className="w-4 h-4" />
                  </button>
                </div>
                <div className="grid grid-cols-1 gap-4 md:grid-cols-2 xl:grid-cols-4">
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">Gerência *</label>
                    <select
                      value={item.gerencia}
                      onChange={(e) => {
                        const novaGerencia = e.target.value;
                        const ccsDaGerencia = getCcsByGerencia(novaGerencia);
                        setItems(prev => prev.map((current, i) =>
                          i === index
                            ? {
                                ...current,
                                gerencia: novaGerencia,
                                ccNovo: ccsDaGerencia.includes(current.ccNovo) ? current.ccNovo : (ccsDaGerencia[0] || ''),
                              }
                            : current
                        ));
                      }}
                      className="w-full px-3 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
                    >
                      <option value="">Selecione a gerência...</option>
                      {equipmentInfo.gerencia && (
                        <optgroup label={`Gerência do equipamento (${equipment})`}>
                          <option key={`eq-${equipmentInfo.gerencia}`} value={equipmentInfo.gerencia}>
                            {equipmentInfo.gerencia}
                          </option>
                        </optgroup>
                      )}
                      <optgroup label="Todas as gerências da base">
                        {gerenciaOptions
                          .filter(ger => ger !== equipmentInfo.gerencia)
                          .map(ger => (
                            <option key={`all-${ger}`} value={ger}>{ger}</option>
                          ))}
                      </optgroup>
                    </select>
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">CC NOVO *</label>
                    {(() => {
                      const ccsDaGerencia = getCcsByGerencia(item.gerencia);
                      const todosCcs = sortTextValues([
                        ...abastecimentos.map(a => a.ccNovo),
                        ...equipments.flatMap(e => e.ccNovo),
                      ]);
                      return (
                    <select
                      value={item.ccNovo}
                      onChange={(e) => updateItem(index, 'ccNovo', e.target.value)}
                      className="w-full px-3 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
                    >
                      <option value="">Selecione o CC...</option>
                      {item.gerencia && ccsDaGerencia.length > 0 && (
                        <optgroup label={`CCs da gerência (${item.gerencia})`}>
                          {ccsDaGerencia.map((cc: string) => (
                            <option key={`ger-${cc}`} value={cc}>{cc}</option>
                          ))}
                        </optgroup>
                      )}
                      <optgroup label="Todos os CCs da base">
                        {todosCcs
                          .filter(cc => !ccsDaGerencia.includes(cc))
                          .map(cc => (
                            <option key={`all-${cc}`} value={cc}>{cc}</option>
                          ))}
                      </optgroup>
                    </select>
                      );
                    })()}
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">Descrição</label>
                    <input
                      type="text"
                      value={item.description}
                      onChange={(e) => updateItem(index, 'description', e.target.value)}
                      placeholder="Ex: Mina A - Operações"
                      className="w-full px-3 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">% Rateio</label>
                    <div className="relative">
                      <input
                        type="number"
                        value={item.percentage}
                        onChange={(e) => updateItem(index, 'percentage', parseFloat(e.target.value) || 0)}
                        className="w-full px-3 py-2 pr-8 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
                      />
                      <span className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400">%</span>
                    </div>
                  </div>
                </div>
              </div>
            ))}
          </div>

          <div className="mt-4 p-3 rounded-lg border border-slate-200">
            <div className="flex items-center justify-between">
              <span className="text-sm font-medium text-slate-700">Total das porcentagens</span>
              <span className={cn(
                "text-sm font-bold",
                totalPercentage === 100 ? "text-green-600" : "text-red-600"
              )}>
                {formatNumber(totalPercentage, 2)}%
              </span>
            </div>
            <div className="mt-2 w-full bg-slate-200 rounded-full h-2">
              <div 
                className={cn("h-2 rounded-full transition-all", totalPercentage === 100 ? "bg-green-600" : "bg-red-600")}
                style={{ width: `${Math.min(totalPercentage, 100)}%` }}
              />
            </div>
          </div>

          <div className="flex gap-3 pt-4">
            <button
              onClick={handleSubmit}
              disabled={totalPercentage !== 100 || !equipment}
              className="px-6 py-2 bg-blue-600 hover:bg-blue-700 disabled:bg-slate-300 text-white font-medium rounded-lg flex items-center gap-2"
            >
              <Check className="w-4 h-4" />
              {submitButtonText}
            </button>
            <button className="px-6 py-2 border border-slate-200 hover:bg-slate-50 text-slate-700 font-medium rounded-lg flex items-center gap-2">
              <X className="w-4 h-4" />
              Cancelar
            </button>
          </div>
        </div>
      </div>
    );
  };

  const renderEquipments = () => (
    <div className="space-y-6">
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-4 sm:p-6 overflow-hidden">
        <div className="mb-6 grid grid-cols-1 gap-4 xl:grid-cols-[minmax(0,1fr),320px] xl:items-center">
          <div className="min-w-0 max-w-full">
            <h3 className="flex items-center gap-2 text-lg font-semibold text-slate-800 break-words">
              <Wrench className="h-5 w-5 shrink-0 text-red-800" />
              <span>Cadastro de Equipamentos</span>
            </h3>
            <p className="mt-1 text-sm text-slate-500">{equipments.length} equipamento(s) cadastrado(s)</p>
          </div>
          <div className="relative w-full max-w-full xl:justify-self-end">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
            <input
              type="text"
              placeholder="Buscar equipamento..."
              className="w-full max-w-full rounded-lg border border-slate-200 py-2 pl-10 pr-4 text-sm outline-none focus:border-red-800 focus:ring-2 focus:ring-red-800/20"
            />
          </div>
        </div>

        <EquipmentForm
          key={editingEquipment ? `edit-${editingEquipment.id}` : 'new-equipment'}
          onAdd={editingEquipment ? handleUpdateEquipment : handleAddEquipment}
          initialEquipment={editingEquipment}
          onCancel={() => setEditingEquipment(null)}
        />

        {equipments.length === 0 ? (
          <div className="mt-8 text-center py-12 border border-slate-200 rounded-lg">
            <div className="w-16 h-16 rounded-full bg-slate-100 flex items-center justify-center mx-auto mb-4">
              <Wrench className="w-8 h-8 text-slate-400" />
            </div>
            <p className="text-slate-600 font-medium mb-1">Nenhum equipamento cadastrado</p>
            <p className="text-sm text-slate-400 mb-4">Clique em "Novo Equipamento" para começar</p>
            <button className="px-6 py-2 bg-blue-600 hover:bg-blue-700 text-white font-medium rounded-lg">
              + Novo Equipamento
            </button>
          </div>
        ) : (
          <div className="mt-8 grid grid-cols-1 gap-4 sm:grid-cols-2 xl:grid-cols-3">
            {equipments.map(e => (
              <div key={e.id} className="border border-slate-200 rounded-lg p-4 hover:shadow-md transition">
                <div className="flex items-start justify-between mb-3">
                  <h4 className="font-semibold text-slate-800">{e.equipment}</h4>
                  <div className="flex items-center gap-2">
                    <span className="text-xs bg-slate-100 px-2 py-1 rounded text-slate-600">{e.plate || 'Sem placa'}</span>
                    <button
                      onClick={() => setEditingEquipment(e)}
                      className="rounded-lg p-1.5 text-slate-400 transition hover:bg-blue-50 hover:text-blue-600"
                      title="Editar equipamento"
                    >
                      <Pencil className="h-4 w-4" />
                    </button>
                    <button
                      onClick={() => handleDeleteEquipment(e.id)}
                      className="rounded-lg p-1.5 text-slate-400 transition hover:bg-red-50 hover:text-red-600"
                      title="Excluir equipamento"
                    >
                      <Trash2 className="h-4 w-4" />
                    </button>
                  </div>
                </div>
                <div className="space-y-1 text-sm text-slate-600">
                  <p><span className="text-slate-400">CC:</span> {e.ccNovo.join(', ')}</p>
                  <p><span className="text-slate-400">Gerência:</span> {e.gerencia}</p>
                  <p><span className="text-slate-400">Área:</span> {e.area}</p>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );

  const EquipmentForm = ({
    onAdd,
    initialEquipment,
    onCancel,
  }: {
    onAdd: (e: any) => void;
    initialEquipment?: Equipment | null;
    onCancel?: () => void;
  }) => {
    const [equipment, setEquipment] = useState(initialEquipment?.equipment || '');
    const [plate, setPlate] = useState(initialEquipment?.plate || '');
    const [ccNovo, setCcNovo] = useState<string[]>(initialEquipment?.ccNovo?.length ? [...initialEquipment.ccNovo] : ['']);
    const [gerencia, setGerencia] = useState(initialEquipment?.gerencia || '');
    const [areaLotacao, setAreaLotacao] = useState(initialEquipment?.areaLotacao || '');
    const [area, setArea] = useState(initialEquipment?.area || '');
    const [fornecedor, setFornecedor] = useState(initialEquipment?.fornecedor || '');

    const equipmentFieldOptions = useMemo(() => ({
      ccNovo: sortTextValues([
        ...equipments.flatMap((item) => item.ccNovo),
        ...abastecimentos.map((item) => item.ccNovo),
      ]),
      gerencias: sortTextValues([
        ...equipments.map((item) => item.gerencia),
        ...abastecimentos.map((item) => item.gerencia),
      ]),
      areasLotacao: sortTextValues([
        ...equipments.map((item) => item.areaLotacao),
        ...abastecimentos.map((item) => item.areaLotacao),
      ]),
      areas: sortTextValues([
        ...equipments.map((item) => item.area),
        ...abastecimentos.map((item) => item.area),
      ]),
      fornecedores: sortTextValues([
        ...equipments.map((item) => item.fornecedor),
        ...abastecimentos.map((item) => item.fornecedor),
      ]),
    }), [equipments, abastecimentos]);

    const handleSubmit = () => {
      if (!equipment || !gerencia || !areaLotacao || !area || !fornecedor) return;
      onAdd({
        equipment,
        plate,
        ccNovo: ccNovo.filter(c => c),
        gerencia,
        areaLotacao,
        area,
        fornecedor
      });
      setEquipment('');
      setPlate('');
      setCcNovo(['']);
      setGerencia('');
      setAreaLotacao('');
      setArea('');
      setFornecedor('');
    };

    const isEditing = Boolean(initialEquipment);

    const handleCancel = () => {
      setEquipment('');
      setPlate('');
      setCcNovo(['']);
      setGerencia('');
      setAreaLotacao('');
      setArea('');
      setFornecedor('');
      onCancel?.();
    };

    return (
      <div className="border border-slate-200 rounded-lg p-4 sm:p-6 mb-6">
        <h4 className="text-sm font-semibold text-slate-700 mb-4 flex items-center gap-2">
          <Plus className="w-4 h-4 text-red-800" />
          {isEditing ? 'Editar Equipamento' : 'Novo Equipamento'}
        </h4>
        <div className="grid grid-cols-1 gap-4">
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Equipamento *</label>
            <input
              type="text"
              value={equipment}
              onChange={(e) => setEquipment(e.target.value)}
              placeholder="Ex: Caminhão CA-001..."
              className="w-full px-4 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Placa</label>
            <input
              type="text"
              value={plate}
              onChange={(e) => setPlate(e.target.value)}
              placeholder="Ex: ABC-1234"
              className="w-full px-4 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">CC NOVO <span className="text-slate-400 font-normal">(pode ter mais de um)</span></label>
            <div className="flex flex-col gap-2">
              {ccNovo.map((cc, idx) => (
                <div key={idx} className="flex flex-col sm:flex-row gap-2">
                  <ComboInput
                    value={cc}
                    onChange={(value) => {
                      const newCc = [...ccNovo];
                      newCc[idx] = value;
                      setCcNovo(newCc);
                    }}
                    options={equipmentFieldOptions.ccNovo}
                    placeholder="Ex: 42105500"
                    wrapperClassName="flex-1 w-full"
                  />
                  {idx === ccNovo.length - 1 && (
                    <button
                      onClick={() => setCcNovo([...ccNovo, ''])}
                      className="px-3 py-2 text-blue-600 hover:text-blue-800 text-sm font-medium whitespace-nowrap"
                    >
                      + Adicionar CC
                    </button>
                  )}
                </div>
              ))}
            </div>
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Gerência *</label>
            <ComboInput
              value={gerencia}
              onChange={setGerencia}
              options={equipmentFieldOptions.gerencias}
              placeholder="Ex: Mineração, Mecânica..."
              wrapperClassName="w-full"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Área de Lotação *</label>
            <ComboInput
              value={areaLotacao}
              onChange={setAreaLotacao}
              options={equipmentFieldOptions.areasLotacao}
              placeholder="Ex: Mina A, Usina..."
              wrapperClassName="w-full"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Área *</label>
            <ComboInput
              value={area}
              onChange={setArea}
              options={equipmentFieldOptions.areas}
              placeholder="Ex: Produção, Manutenção..."
              wrapperClassName="w-full"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Fornecedor *</label>
            <ComboInput
              value={fornecedor}
              onChange={setFornecedor}
              options={equipmentFieldOptions.fornecedores}
              placeholder="Ex: Posto Shell..."
              wrapperClassName="w-full"
            />
          </div>
        </div>
        <div className="flex flex-col sm:flex-row gap-3 mt-6">
          <button
            onClick={handleSubmit}
            className="w-full sm:w-auto px-6 py-2.5 bg-blue-600 hover:bg-blue-700 text-white font-medium rounded-lg flex items-center justify-center gap-2"
          >
            <Check className="w-4 h-4" />
            {isEditing ? 'Salvar Alterações' : 'Cadastrar Equipamento'}
          </button>
          <button
            onClick={handleCancel}
            className="w-full sm:w-auto px-6 py-2.5 border border-slate-200 hover:bg-slate-50 text-slate-700 font-medium rounded-lg"
          >
            Cancelar
          </button>
        </div>
      </div>
    );
  };

  const renderHistory = () => (
    <div className="space-y-6">
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
        <div className="mb-6">
          <div className="mb-4">
            <h3 className="text-lg font-semibold text-slate-800 flex items-center gap-2">
              <History className="w-5 h-5 text-red-800" />
              Histórico de Preenchimentos
            </h3>
            <p className="text-sm text-slate-500">Registro completo de todos os abastecimentos realizados</p>
          </div>
          <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
            <div className="relative">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
              <input
                type="text"
                value={historyUserFilter}
                onChange={(e) => setHistoryUserFilter(e.target.value)}
                placeholder="Filtrar por usuário..."
                className="w-full pl-10 pr-4 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none sm:w-64"
              />
            </div>
            <div className="relative">
              <Calendar className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
              <input
                type="date"
                value={historyDateFilter}
                onChange={(e) => setHistoryDateFilter(e.target.value)}
                className="w-full pl-10 pr-4 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none sm:w-auto"
              />
            </div>
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="min-w-[980px] w-full">
            <thead>
              <tr className="border-b border-slate-200">
                <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">ID</th>
                <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Data</th>
                <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Equipamento</th>
                <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Litros</th>
                <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Valor (R$)</th>
                <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Usuário</th>
                <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Horário</th>
                <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Observações</th>
                <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Rateio</th>
              </tr>
            </thead>
            <tbody>
              {filteredHistory.length > 0 ? (
                filteredHistory.map(a => (
                  <tr key={a.id} className="border-b border-slate-100 hover:bg-slate-50">
                    <td className="py-3 px-4 text-sm text-red-800 font-medium">{a.id}</td>
                    <td className="py-3 px-4 text-sm text-slate-600">{a.data}</td>
                    <td className="py-3 px-4 text-sm text-slate-800">{a.equipamento}</td>
                    <td className="py-3 px-4 text-sm text-slate-800 font-medium">{formatLiters(a.litros)}</td>
                    <td className="py-3 px-4 text-sm text-slate-800 font-medium">{formatCurrency(a.valor)}</td>
                    <td className="py-3 px-4 text-sm text-slate-600">{a.createdBy || 'Usuário não identificado'}</td>
                    <td className="py-3 px-4 text-sm text-slate-500">
                      {new Date(a.createdAt).toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}
                    </td>
                    <td className="py-3 px-4 text-sm text-slate-500 max-w-[200px]">
                      {a.observacoes ? (
                        <div className="flex items-center gap-2">
                          <span className="truncate block" title={a.observacoes}>{a.observacoes}</span>
                          <button
                            onClick={() => setSelectedObservation(a.observacoes || '')}
                            className="shrink-0 rounded-md p-1 text-slate-400 transition hover:bg-blue-50 hover:text-blue-600"
                            title="Ver observação completa"
                          >
                            <Eye className="h-4 w-4" />
                          </button>
                        </div>
                      ) : (
                        <span className="text-slate-400 text-xs">-</span>
                      )}
                    </td>
                    <td className="py-3 px-4">
                      {a.rateioInfo?.rateado ? (
                        <span className="px-2 py-1 bg-blue-100 text-blue-800 rounded text-xs">
                          {a.rateioInfo.gerenciaRateio} ({a.rateioInfo.percentage}%)
                        </span>
                      ) : (
                        <span className="text-slate-400 text-xs">-</span>
                      )}
                    </td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={9} className="py-12 text-center">
                    <div className="w-16 h-16 rounded-full bg-slate-100 flex items-center justify-center mx-auto mb-4">
                      <Clock className="w-8 h-8 text-slate-400" />
                    </div>
                    <p className="text-slate-400">Nenhum registro encontrado</p>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
      {selectedObservation && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4">
          <div className="w-full max-w-lg rounded-2xl bg-white shadow-xl">
            <div className="flex items-center justify-between border-b border-slate-200 px-6 py-4">
              <h3 className="text-lg font-semibold text-slate-800">Observação completa</h3>
              <button
                onClick={() => setSelectedObservation(null)}
                className="rounded-lg p-1 text-slate-400 transition hover:bg-slate-100 hover:text-slate-600"
              >
                <X className="h-5 w-5" />
              </button>
            </div>
            <div className="px-6 py-5">
              <p className="whitespace-pre-wrap break-words text-sm leading-6 text-slate-700">
                {selectedObservation}
              </p>
            </div>
            <div className="flex justify-end border-t border-slate-200 px-6 py-4">
              <button
                onClick={() => setSelectedObservation(null)}
                className="rounded-lg bg-blue-600 px-5 py-2 text-sm font-medium text-white transition hover:bg-blue-700"
              >
                Fechar
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );

  const renderPreenchimento = () => (
    <div className="max-w-3xl mx-auto">
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
        <div className="mb-6">
          <h3 className="text-lg font-semibold text-slate-800 flex items-center gap-2">
            <Plus className="w-5 h-5 text-red-800" />
            Preenchimento de Abastecimento
          </h3>
          <p className="text-sm text-slate-500">Próximo ID: {abastecimentos.length + 1}</p>
        </div>

        <PreenchimentoForm 
          equipments={equipments} 
          rateios={rateios}
          onAdd={handleAddAbastecimento} 
        />
      </div>
    </div>
  );

  const PreenchimentoForm = ({ equipments, rateios, onAdd }: { equipments: Equipment[], rateios: Rateio[], onAdd: (a: any) => void }) => {
    const [plate, setPlate] = useState('');
    const [equipment, setEquipment] = useState('');
    const [ccNovo, setCcNovo] = useState('');
    const [diretoria, setDiretoria] = useState('');
    const [gerencia, setGerencia] = useState('');
    const [areaLotacao, setAreaLotacao] = useState('');
    const [fornecedor, setFornecedor] = useState('');
    const [area, setArea] = useState('');
    const [data, setData] = useState(new Date().toISOString().split('T')[0]);
    const [litros, setLitros] = useState('');
    const [observacoes, setObservacoes] = useState('');
    const [selectedRateio, setSelectedRateio] = useState<string>('');

    const preenchimentoOptions = useMemo(() => ({
      placas: sortTextValues(equipments.map((item) => item.plate).filter(Boolean)),
      equipamentos: sortTextValues([
        ...equipments.map((item) => item.equipment),
        ...abastecimentos.map((item) => item.equipamento),
      ]),
      ccs: sortTextValues([
        ...equipments.flatMap((item) => item.ccNovo),
        ...abastecimentos.map((item) => item.ccNovo),
      ]),
      diretorias: sortTextValues(abastecimentos.map((item) => item.diretoria)),
      gerencias: sortTextValues([
        ...equipments.map((item) => item.gerencia),
        ...abastecimentos.map((item) => item.gerencia),
      ]),
      areasLotacao: sortTextValues([
        ...equipments.map((item) => item.areaLotacao),
        ...abastecimentos.map((item) => item.areaLotacao),
      ]),
      fornecedores: sortTextValues([
        ...equipments.map((item) => item.fornecedor),
        ...abastecimentos.map((item) => item.fornecedor),
      ]),
      areas: sortTextValues([
        ...equipments.map((item) => item.area),
        ...abastecimentos.map((item) => item.area),
      ]),
    }), [equipments, abastecimentos]);

    const getEquipmentInfo = (equipmentName: string, plateValue?: string) => {
      const normalizedEquipment = cleanText(equipmentName).toLowerCase();
      const normalizedPlate = cleanText(plateValue || '').toLowerCase();

      const registered = equipments.find((item) =>
        (normalizedEquipment && item.equipment.toLowerCase() === normalizedEquipment) ||
        (normalizedPlate && item.plate.toLowerCase() === normalizedPlate)
      );

      const historical = [...abastecimentos].reverse().find((item) =>
        normalizedEquipment && item.equipamento.toLowerCase() === normalizedEquipment
      );

      const finalEquipment = registered?.equipment || historical?.equipamento || equipmentName;
      const historicalItems = abastecimentos.filter((item) => item.equipamento === finalEquipment);

      return {
        equipment: finalEquipment,
        plate: registered?.plate || plateValue || '',
        ccNovo: registered?.ccNovo?.[0] || historical?.ccNovo || '',
        ccOptions: sortTextValues([
          ...(registered?.ccNovo || []),
          ...historicalItems.map((item) => item.ccNovo),
          ...preenchimentoOptions.ccs,
        ]),
        diretoria: historical?.diretoria || '',
        gerencia: registered?.gerencia || historical?.gerencia || '',
        areaLotacao: registered?.areaLotacao || historical?.areaLotacao || '',
        fornecedor: registered?.fornecedor || historical?.fornecedor || '',
        area: registered?.area || historical?.area || '',
      };
    };

    const fillFromEquipmentInfo = (info: ReturnType<typeof getEquipmentInfo>) => {
      setPlate(info.plate);
      setEquipment(info.equipment);
      setCcNovo(info.ccNovo);
      setDiretoria(info.diretoria);
      setGerencia(info.gerencia);
      setAreaLotacao(info.areaLotacao);
      setFornecedor(info.fornecedor);
      setArea(info.area);
    };

    const currentEquipmentInfo = equipment ? getEquipmentInfo(equipment, plate) : null;

    const handlePlateSearch = () => {
      const info = getEquipmentInfo('', plate);
      if (info.equipment) fillFromEquipmentInfo(info);
    };

    const handleSubmit = () => {
      if (!equipment || !ccNovo || !gerencia || !areaLotacao || !fornecedor || !area || !data || !litros) return;
      
      const rateioInfo = selectedRateio ? (() => {
        const rateio = rateios.find(r => r.id.toString() === selectedRateio);
        if (!rateio || rateio.items.length === 0) return undefined;
        return {
          rateado: true,
          rateioId: rateio.id,
          gerenciaRateio: rateio.items[0].gerencia,
          ccRateio: rateio.items[0].ccNovo,
          percentage: rateio.items[0].percentage
        };
      })() : undefined;

      onAdd({
        ccNovo,
        diretoria,
        gerencia,
        areaLotacao,
        fornecedor,
        equipamento: equipment,
        area,
        semana: `Sem ${Math.ceil(new Date(data).getDate() / 7)}`,
        data,
        litros: parseFloat(litros) || 0,
        observacoes,
        rateioInfo
      });

      setPlate('');
      setEquipment('');
      setCcNovo('');
      setDiretoria('');
      setGerencia('');
      setAreaLotacao('');
      setFornecedor('');
      setArea('');
      setData(new Date().toISOString().split('T')[0]);
      setLitros('');
      setObservacoes('');
      setSelectedRateio('');
    };

    return (
      <div className="space-y-6">
        <div className="bg-blue-50 rounded-lg p-4 border border-blue-100">
          <p className="text-sm font-medium text-blue-800 mb-3 flex items-center gap-2">
            <Droplets className="w-4 h-4" />
            DIGITE A PLACA OU SELECIONE O EQUIPAMENTO PARA PREENCHIMENTO AUTOMÁTICO
          </p>
          <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">Placa</label>
              <div className="flex gap-2">
                <ComboInput
                  value={plate}
                  onChange={(value) => {
                    setPlate(value);
                    const exactPlate = equipments.find((item) => item.plate.toLowerCase() === value.toLowerCase());
                    if (exactPlate) fillFromEquipmentInfo(getEquipmentInfo(exactPlate.equipment, value));
                  }}
                  onOptionSelect={(value) => {
                    const exactPlate = equipments.find((item) => item.plate.toLowerCase() === value.toLowerCase());
                    if (exactPlate) fillFromEquipmentInfo(getEquipmentInfo(exactPlate.equipment, value));
                  }}
                  options={preenchimentoOptions.placas}
                  placeholder="Digite ou selecione a placa..."
                  wrapperClassName="flex-1"
                />
                <button
                  onClick={handlePlateSearch}
                  className="px-4 py-2 bg-blue-600 text-white rounded-lg text-sm hover:bg-blue-700"
                >
                  Buscar
                </button>
              </div>
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">Equipamento *</label>
              <ComboInput
                value={equipment}
                onChange={(value) => {
                  setEquipment(value);
                  const exactEquipment = preenchimentoOptions.equipamentos.find((option) => option.toLowerCase() === value.toLowerCase());
                  if (exactEquipment) fillFromEquipmentInfo(getEquipmentInfo(exactEquipment, plate));
                }}
                onOptionSelect={(value) => fillFromEquipmentInfo(getEquipmentInfo(value, plate))}
                options={preenchimentoOptions.equipamentos}
                placeholder="Digite ou selecione o equipamento..."
              />
            </div>
          </div>
        </div>

        <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">CC Novo *</label>
            <select
              value={ccNovo}
              onChange={(e) => setCcNovo(e.target.value)}
              className="w-full px-4 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
            >
              <option value="">Selecione...</option>
              {(currentEquipmentInfo?.ccOptions || preenchimentoOptions.ccs).map((option) => (
                <option key={option} value={option}>{option}</option>
              ))}
            </select>
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Diretoria *</label>
            <select
              value={diretoria}
              onChange={(e) => setDiretoria(e.target.value)}
              className="w-full px-4 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
            >
              <option value="">Selecione...</option>
              {preenchimentoOptions.diretorias.map((option) => (
                <option key={option} value={option}>{option}</option>
              ))}
            </select>
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Gerência *</label>
            <ComboInput
              value={gerencia}
              onChange={setGerencia}
              options={preenchimentoOptions.gerencias}
              placeholder="Ex: Mineração, Mecânica..."
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Área de Lotação *</label>
            <ComboInput
              value={areaLotacao}
              onChange={setAreaLotacao}
              options={preenchimentoOptions.areasLotacao}
              placeholder="Ex: Mina A, Usina..."
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Fornecedor *</label>
            <ComboInput
              value={fornecedor}
              onChange={setFornecedor}
              options={preenchimentoOptions.fornecedores}
              placeholder="Ex: Posto Shell, Posto BR..."
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Área *</label>
            <ComboInput
              value={area}
              onChange={setArea}
              options={preenchimentoOptions.areas}
              placeholder="Ex: Produção, Manutenção..."
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Data *</label>
            <input
              type="date"
              value={data}
              onChange={(e) => setData(e.target.value)}
              className="w-full px-4 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Litros *</label>
            <input
              type="number"
              value={litros}
              onChange={(e) => setLitros(e.target.value)}
              placeholder="Ex: 250.5"
              step="0.01"
              className="w-full px-4 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
            />
          </div>
          {rateios.length > 0 && (
            <div className="col-span-2">
              <label className="block text-sm font-medium text-slate-700 mb-1">Rateio (Opcional)</label>
              <select
                value={selectedRateio}
                onChange={(e) => setSelectedRateio(e.target.value)}
                className="w-full px-4 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
              >
                <option value="">Sem rateio</option>
                {rateios.map(r => (
                  <option key={r.id} value={r.id}>{r.description} - {r.equipment}</option>
                ))}
              </select>
            </div>
          )}
          <div className="col-span-2">
            <label className="block text-sm font-medium text-slate-700 mb-1">Observações</label>
            <input
              type="text"
              value={observacoes}
              onChange={(e) => setObservacoes(e.target.value)}
              placeholder="Observações opcionais..."
              className="w-full px-4 py-2 rounded-lg border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
            />
          </div>
        </div>

        <button
          onClick={handleSubmit}
          className="w-full py-3 px-4 bg-blue-600 hover:bg-blue-700 text-white font-medium rounded-lg flex items-center justify-center gap-2"
        >
          <Database className="w-5 h-5" />
          Salvar na Base de Dados
        </button>
      </div>
    );
  };

  const renderImport = () => (
    <div className="max-w-4xl mx-auto">
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
        <div className="mb-6">
          <h3 className="text-lg font-semibold text-slate-800 flex items-center gap-2">
            <Upload className="w-5 h-5 text-red-800" />
            Importação de Dados
          </h3>
          <p className="text-sm text-slate-500">Importe direto do Excel ou cole dados tabulados</p>
        </div>

        {!importPreview ? (
        <>
        <div className="border-2 border-dashed border-slate-300 rounded-xl p-12 text-center hover:border-red-800 transition-colors">
          <div className="w-16 h-16 rounded-full bg-slate-100 flex items-center justify-center mx-auto mb-4">
            <Upload className="w-8 h-8 text-slate-400" />
          </div>
          <p className="text-slate-600 mb-2">Arraste o arquivo ou clique para selecionar</p>
          <p className="text-sm text-slate-400 mb-4">Suporta .xlsx, .xls e .csv</p>
          <label className="inline-block px-6 py-2 bg-blue-600 hover:bg-blue-700 text-white font-medium rounded-lg cursor-pointer">
            Selecionar arquivo Excel
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={(e) => {
                const file = e.target.files?.[0];
                if (!file) return;
                const reader = new FileReader();
                reader.onload = (evt) => {
                  const bstr = evt.target?.result;
                  const wb = XLSX.read(bstr, { type: 'binary' });
                  const wsname = wb.SheetNames[0];
                  const ws = wb.Sheets[wsname];
                  const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
                  handleImportExcel(data as any[][], file.name);
                };
                reader.readAsBinaryString(file);
              }}
            />
          </label>
        </div>

        <div className="mt-6 rounded-xl border border-amber-300 bg-amber-50 p-5">
          <h4 className="mb-4 flex items-center gap-2 text-base font-semibold text-amber-800">
            <FileInput className="w-5 h-5" />
            Como preparar seu arquivo Excel
          </h4>
          <p className="mb-4 text-sm text-amber-700">O arquivo deve ter as colunas nessa ordem (com ou sem cabeçalho):</p>

          <div className="rounded-lg border border-amber-300 bg-white px-4 py-3 font-mono text-xs text-slate-700">
            CC NOVO | DIRETORIA | GERÊNCIA | ÁREA LOT. | FORNECEDOR | EQUIPAMENTO | ÁREA | DATA | LITROS
          </div>

          <p className="mt-4 text-xs text-amber-700 flex items-center gap-1">
            <Calendar className="w-3 h-3" />
            A data pode estar no formato <strong>DD/MM/AAAA</strong> ou <strong>AAAA-MM-DD</strong>
          </p>
        </div>
        </>
        ) : (
          <div className="space-y-5">
            <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
              <div className="flex flex-wrap items-center gap-3">
                <h4 className="flex items-center gap-2 text-base font-semibold text-slate-800">
                  <FileInput className="h-5 w-5 text-blue-600" />
                  Pré-visualização
                </h4>
                <span className="rounded-full bg-blue-100 px-2.5 py-1 text-xs font-semibold text-blue-700">
                  {importPreview.records.length} registros
                </span>
                {importPreview.fileName && (
                  <span className="rounded-md bg-slate-100 px-2.5 py-1 text-xs font-medium text-slate-600">
                    {importPreview.fileName}
                  </span>
                )}
              </div>
              <button
                onClick={() => setImportPreview(null)}
                className="text-sm font-medium text-slate-500 hover:text-slate-700"
              >
                Voltar
              </button>
            </div>

            <div className="overflow-hidden rounded-xl border border-slate-200">
              <div className="max-h-[560px] overflow-auto">
                <table className="w-full text-sm">
                  <thead className="sticky top-0 bg-slate-50 text-xs uppercase text-slate-600">
                    <tr>
                      <th className="px-3 py-3 text-left">CC Novo</th>
                      <th className="px-3 py-3 text-left">Diretoria</th>
                      <th className="px-3 py-3 text-left">Gerência</th>
                      <th className="px-3 py-3 text-left">Área Lot.</th>
                      <th className="px-3 py-3 text-left">Fornecedor</th>
                      <th className="px-3 py-3 text-left">Equipamento</th>
                      <th className="px-3 py-3 text-left">Data</th>
                      <th className="px-3 py-3 text-right">Litros</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {importPreview.records.slice(0, 20).map((item) => (
                      <tr key={item.id} className="hover:bg-slate-50">
                        <td className="px-3 py-2 text-slate-700">{item.ccNovo}</td>
                        <td className="px-3 py-2 text-slate-700">{item.diretoria}</td>
                        <td className="px-3 py-2 text-slate-700">{item.gerencia}</td>
                        <td className="px-3 py-2 text-slate-700">{item.areaLotacao}</td>
                        <td className="px-3 py-2 text-slate-700">{item.fornecedor}</td>
                        <td className="px-3 py-2 text-slate-700">{item.equipamento}</td>
                        <td className="px-3 py-2 text-slate-700">{item.data}</td>
                        <td className="px-3 py-2 text-right font-medium text-slate-900">{formatNumber(item.litros, 2)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              {importPreview.records.length > 20 && (
                <p className="border-t border-slate-200 bg-white px-4 py-3 text-center text-xs text-slate-500">
                  ... e mais {importPreview.records.length - 20} registros
                </p>
              )}
            </div>

            <div className="flex flex-col gap-3 sm:flex-row">
              <button
                onClick={handleConfirmImport}
                className="inline-flex items-center justify-center gap-2 rounded-xl bg-green-600 px-5 py-3 text-sm font-semibold text-white transition hover:bg-green-700"
              >
                <CheckCircle className="h-4 w-4" />
                Confirmar Importação ({importPreview.records.length} registros)
              </button>
              <button
                onClick={() => setImportPreview(null)}
                className="inline-flex items-center justify-center gap-2 rounded-xl bg-slate-100 px-5 py-3 text-sm font-semibold text-slate-700 transition hover:bg-slate-200"
              >
                <X className="h-4 w-4" />
                Cancelar
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );

  const ProfilePage = () => {
    const [editName, setEditName] = useState(currentUser?.name || '');
    const [editEmail, setEditEmail] = useState(currentUser?.email || '');
    const [editFuncao, setEditFuncao] = useState(currentUser?.funcao || '');
    const [editPassword, setEditPassword] = useState('');
    const [confirmPassword, setConfirmPassword] = useState('');
    const [showDelete, setShowDelete] = useState(false);
    const [deleteConfirm, setDeleteConfirm] = useState('');
    const [showPassFields, setShowPassFields] = useState(false);

    const handleSaveProfile = () => {
      if (!editName.trim()) {
        addNotification('error', 'Nome não pode estar vazio');
        return;
      }
      if (showPassFields && editPassword) {
        if (editPassword !== confirmPassword) {
          addNotification('error', 'As senhas não coincidem');
          return;
        }
        if (editPassword.length < 4) {
          addNotification('error', 'Senha deve ter no mínimo 4 caracteres');
          return;
        }
      }
      setUsers(prev => prev.map(u => {
        if (u.id !== currentUser?.id) return u;
        return {
          ...u,
          name: cleanText(editName),
          email: cleanText(editEmail),
          funcao: cleanText(editFuncao),
          ...(showPassFields && editPassword ? { password: editPassword } : {}),
        };
      }));
      if (currentUser) {
        setCurrentUser(prev => prev ? {
          ...prev,
          name: cleanText(editName),
          email: cleanText(editEmail),
          funcao: cleanText(editFuncao),
          ...(showPassFields && editPassword ? { password: editPassword } : {}),
        } : null);
      }
      setShowPassFields(false);
      setEditPassword('');
      setConfirmPassword('');
      addNotification('success', 'Perfil atualizado com sucesso!');
    };

    const handleAvatarChange = (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (evt) => {
        const dataUrl = evt.target?.result as string;
        setUsers(prev => prev.map(u => u.id === currentUser?.id ? { ...u, avatar: dataUrl } : u));
        setCurrentUser(prev => prev ? { ...prev, avatar: dataUrl } : null);
        addNotification('success', 'Foto de perfil atualizada!');
      };
      reader.readAsDataURL(file);
    };

    const handleDeleteProfile = () => {
      if (deleteConfirm !== currentUser?.username) {
        addNotification('error', 'Usuário digitado não confere');
        return;
      }
      handleLogout();
      setUsers(prev => prev.filter(u => u.id !== currentUser?.id));
      addNotification('info', 'Perfil excluído com sucesso.');
    };

    const handleRemoveAvatar = () => {
      setUsers(prev => prev.map(u => u.id === currentUser?.id ? { ...u, avatar: undefined } : u));
      setCurrentUser(prev => prev ? { ...prev, avatar: undefined } : null);
      addNotification('success', 'Foto de perfil removida!');
    };

    return (
      <div className="max-w-2xl mx-auto space-y-6">
        {/* Card de foto e identidade */}
        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
          <div className="h-24 bg-gradient-to-r from-red-900 via-red-800 to-slate-800" />
          <div className="px-6 pb-6">
            <div className="relative -mt-12 mb-4 flex items-end justify-between">
              <div className="relative">
                <div className="h-24 w-24 rounded-2xl border-4 border-white bg-slate-200 shadow-lg overflow-hidden">
                  {currentUser?.avatar ? (
                    <img src={currentUser.avatar} alt="Avatar" className="h-full w-full object-cover" />
                  ) : (
                    <div className="h-full w-full flex items-center justify-center bg-gradient-to-br from-red-800 to-slate-700 text-white text-3xl font-bold">
                      {(currentUser?.name || '?')[0].toUpperCase()}
                    </div>
                  )}
                </div>
                <label className="absolute -bottom-1 -right-1 flex h-7 w-7 cursor-pointer items-center justify-center rounded-full bg-blue-600 text-white shadow hover:bg-blue-700 transition">
                  <Camera className="h-3.5 w-3.5" />
                  <input type="file" accept="image/*" className="hidden" onChange={handleAvatarChange} />
                </label>
              </div>
              {currentUser?.avatar && (
                <button
                  onClick={handleRemoveAvatar}
                  className="text-xs text-slate-400 hover:text-red-600 transition"
                >
                  Remover foto
                </button>
              )}
            </div>
            <div>
              <p className="text-xl font-bold text-slate-900">{currentUser?.name}</p>
              <p className="text-sm text-slate-500">@{currentUser?.username}</p>
              <div className="mt-2 flex flex-wrap gap-2">
                <span className={cn(
                  "inline-flex items-center gap-1 rounded-full px-3 py-1 text-xs font-medium",
                  currentUser?.role === 'admin' ? "bg-red-100 text-red-800" : currentUser?.role === 'supervisor' ? "bg-blue-100 text-blue-800" : "bg-slate-100 text-slate-700"
                )}>
                  {getRoleLabel(currentUser?.role || 'operator')}
                </span>
                {currentUser?.funcao && (
                  <span className="inline-flex items-center gap-1 rounded-full bg-blue-50 px-3 py-1 text-xs font-medium text-blue-700">
                    <Briefcase className="h-3 w-3" />
                    {currentUser.funcao}
                  </span>
                )}
                {currentUser?.email && (
                  <span className="inline-flex items-center gap-1 rounded-full bg-slate-50 px-3 py-1 text-xs font-medium text-slate-600">
                    <Mail className="h-3 w-3" />
                    {currentUser.email}
                  </span>
                )}
              </div>
            </div>
          </div>
        </div>

        {/* Editar dados */}
        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
          <h3 className="text-base font-semibold text-slate-800 mb-5 flex items-center gap-2">
            <Pencil className="h-4 w-4 text-red-800" />
            Editar Informações
          </h3>
          <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1">Nome completo *</label>
              <input
                type="text"
                value={editName}
                onChange={e => setEditName(e.target.value)}
                className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1">Usuário</label>
              <input
                type="text"
                value={currentUser?.username || ''}
                disabled
                className="w-full px-4 py-2.5 rounded-xl border border-slate-200 bg-slate-50 text-sm text-slate-400 cursor-not-allowed"
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1">
                <span className="flex items-center gap-1"><Mail className="h-3.5 w-3.5" /> E-mail</span>
              </label>
              <input
                type="email"
                value={editEmail}
                onChange={e => setEditEmail(e.target.value)}
                placeholder="seu@email.com"
                className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1">
                <span className="flex items-center gap-1"><Briefcase className="h-3.5 w-3.5" /> Função / Cargo</span>
              </label>
              <input
                type="text"
                value={editFuncao}
                onChange={e => setEditFuncao(e.target.value)}
                placeholder="Ex: Operador de Campo, Gerente..."
                className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
              />
            </div>
          </div>

          {/* Alterar senha */}
          <div className="mt-5 border-t border-slate-100 pt-5">
            <button
              onClick={() => setShowPassFields(!showPassFields)}
              className="text-sm font-medium text-blue-600 hover:text-blue-800 transition flex items-center gap-1"
            >
              <Eye className="h-4 w-4" />
              {showPassFields ? 'Cancelar alteração de senha' : 'Alterar senha'}
            </button>
            {showPassFields && (
              <div className="mt-4 grid grid-cols-1 gap-4 md:grid-cols-2">
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">Nova senha</label>
                  <input
                    type="password"
                    value={editPassword}
                    onChange={e => setEditPassword(e.target.value)}
                    placeholder="Mínimo 4 caracteres"
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">Confirmar senha</label>
                  <input
                    type="password"
                    value={confirmPassword}
                    onChange={e => setConfirmPassword(e.target.value)}
                    placeholder="Repita a nova senha"
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
                  />
                </div>
              </div>
            )}
          </div>

          <button
            onClick={handleSaveProfile}
            className="mt-6 flex items-center gap-2 rounded-xl bg-blue-600 px-6 py-2.5 text-sm font-semibold text-white transition hover:bg-blue-700"
          >
            <Check className="h-4 w-4" />
            Salvar alterações
          </button>
        </div>

        {/* Zona de perigo */}
        <div className="bg-white rounded-2xl shadow-sm border border-red-200 p-6">
          <h3 className="text-base font-semibold text-red-800 mb-2 flex items-center gap-2">
            <AlertCircle className="h-4 w-4" />
            Zona de Perigo
          </h3>
          <p className="text-sm text-slate-500 mb-4">
            Excluir o perfil é uma ação irreversível. Você será desconectado e seu usuário será removido do sistema.
          </p>
          {!showDelete ? (
            <button
              onClick={() => setShowDelete(true)}
              className="flex items-center gap-2 rounded-xl border border-red-300 bg-red-50 px-5 py-2.5 text-sm font-semibold text-red-700 transition hover:bg-red-100"
            >
              <Trash2 className="h-4 w-4" />
              Excluir meu perfil
            </button>
          ) : (
            <div className="space-y-3">
              <p className="text-sm font-medium text-slate-700">
                Digite seu usuário <strong>({currentUser?.username})</strong> para confirmar:
              </p>
              <input
                type="text"
                value={deleteConfirm}
                onChange={e => setDeleteConfirm(e.target.value)}
                placeholder={currentUser?.username}
                className="w-full px-4 py-2.5 rounded-xl border border-red-300 text-sm focus:border-red-600 focus:ring-2 focus:ring-red-600/20 outline-none"
              />
              <div className="flex gap-3">
                <button
                  onClick={handleDeleteProfile}
                  disabled={deleteConfirm !== currentUser?.username}
                  className="flex items-center gap-2 rounded-xl bg-red-600 px-5 py-2.5 text-sm font-semibold text-white transition hover:bg-red-700 disabled:cursor-not-allowed disabled:opacity-50"
                >
                  <Trash2 className="h-4 w-4" />
                  Confirmar exclusão
                </button>
                <button
                  onClick={() => { setShowDelete(false); setDeleteConfirm(''); }}
                  className="rounded-xl border border-slate-200 px-5 py-2.5 text-sm font-medium text-slate-600 transition hover:bg-slate-50"
                >
                  Cancelar
                </button>
              </div>
            </div>
          )}
        </div>
      </div>
    );
  };

  const renderProfile = () => <ProfilePage />;

  const renderDieselPrice = () => (
    <div className="max-w-2xl mx-auto">
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
        <div className="mb-6">
          <h3 className="text-lg font-semibold text-slate-800 flex items-center gap-2">
            <Fuel className="w-5 h-5 text-red-800" />
            Valor do Diesel
          </h3>
          <p className="text-sm text-slate-500">Atualize o preço do diesel por litro</p>
        </div>

        <div className="bg-gradient-to-r from-blue-50 to-indigo-50 rounded-lg p-6 mb-6">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm text-slate-600 mb-1">Preço Atual</p>
              <p className="text-4xl font-bold text-blue-800">{formatCurrency(dieselPrice.price)}</p>
              <p className="text-xs text-slate-500 mt-2">
                Atualizado em {new Date(dieselPrice.updatedAt).toLocaleDateString('pt-BR')} às{' '}
                {new Date(dieselPrice.updatedAt).toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}
              </p>
              <p className="text-xs text-slate-500">Por: {dieselPrice.updatedBy}</p>
            </div>
            <Fuel className="w-16 h-16 text-blue-300" />
          </div>
        </div>

        <div className="space-y-4">
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Novo Preço (R$ por litro)</label>
            <input
              type="number"
              value={newDieselPrice}
              onChange={(e) => setNewDieselPrice(e.target.value)}
              placeholder="Ex: 7.45"
              step="0.01"
              className="w-full px-4 py-3 rounded-lg border border-slate-200 text-lg focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
            />
          </div>
          <button
            onClick={handleUpdateDieselPrice}
            className="w-full py-3 px-4 bg-blue-600 hover:bg-blue-700 text-white font-medium rounded-lg flex items-center justify-center gap-2"
          >
            <Check className="w-5 h-5" />
            Atualizar Preço
          </button>
        </div>
      </div>
    </div>
  );

  const renderUsers = () => (
    <div className="space-y-6">
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
        <div className="mb-6">
          <h3 className="text-lg font-semibold text-slate-800 flex items-center gap-2">
            <Users className="w-5 h-5 text-red-800" />
            Gestão de Usuários
          </h3>
          <p className="text-sm text-slate-500">Aprove ou rejeite novos usuários</p>
        </div>

        <div className="space-y-4">
          {users.filter(u => u.status === 'pending').length === 0 ? (
            <div className="text-center py-12">
              <div className="w-16 h-16 rounded-full bg-green-100 flex items-center justify-center mx-auto mb-4">
                <CheckCircle className="w-8 h-8 text-green-600" />
              </div>
              <p className="text-slate-600 font-medium">Todos os usuários estão aprovados</p>
              <p className="text-sm text-slate-400">Não há pendências de aprovação</p>
            </div>
          ) : (
            users.filter(u => u.status === 'pending').map(user => (
              <div key={user.id} className="border border-slate-200 rounded-lg p-4 flex items-center justify-between">
                <div className="flex items-center gap-4">
                  <div className="w-12 h-12 rounded-full bg-slate-100 flex items-center justify-center">
                    <UserIcon className="w-6 h-6 text-slate-400" />
                  </div>
                  <div>
                    <p className="font-medium text-slate-800">{user.name}</p>
                    <p className="text-sm text-slate-500">@{user.username}</p>
                    <p className="text-xs text-slate-400">
                      {getRoleLabel(user.role)} • 
                      Criado em {new Date(user.createdAt).toLocaleDateString('pt-BR')}
                    </p>
                  </div>
                </div>
                <div className="flex gap-2">
                  <button
                    onClick={() => handleUserApproval(user.id, 'approved')}
                    className="px-4 py-2 bg-green-600 hover:bg-green-700 text-white font-medium rounded-lg flex items-center gap-2"
                  >
                    <Check className="w-4 h-4" />
                    Aprovar
                  </button>
                  <button
                    onClick={() => handleUserApproval(user.id, 'rejected')}
                    className="px-4 py-2 bg-red-600 hover:bg-red-700 text-white font-medium rounded-lg flex items-center gap-2"
                  >
                    <X className="w-4 h-4" />
                    Rejeitar
                  </button>
                </div>
              </div>
            ))
          )}
        </div>

        <div className="mt-8 border-t border-slate-200 pt-6">
          <h4 className="text-sm font-semibold text-slate-700 mb-4">Todos os Usuários</h4>
          <div className="overflow-x-auto">
            <table className="min-w-[760px] w-full">
              <thead>
                <tr className="border-b border-slate-200">
                  <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Nome</th>
                  <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Usuário</th>
                  <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Função</th>
                  <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Status</th>
                  <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Criado em</th>
                  <th className="text-left py-3 px-4 text-xs font-semibold text-slate-600 uppercase">Ações</th>
                </tr>
              </thead>
              <tbody>
                {users.map(user => (
                  <tr key={user.id} className="border-b border-slate-100">
                    <td className="py-3 px-4 text-sm text-slate-800">{user.name}</td>
                    <td className="py-3 px-4 text-sm text-slate-600">@{user.username}</td>
                    <td className="py-3 px-4">
                      <span className={cn(
                        "px-2 py-1 rounded text-xs font-medium",
                        user.role === 'admin' ? "bg-red-100 text-red-800" : user.role === 'supervisor' ? "bg-blue-100 text-blue-800" : "bg-slate-100 text-slate-600"
                      )}>
                        {getRoleLabel(user.role)}
                      </span>
                    </td>
                    <td className="py-3 px-4">
                      <span className={cn(
                        "px-2 py-1 rounded text-xs font-medium",
                        user.status === 'approved' ? "bg-green-100 text-green-800" :
                        user.status === 'pending' ? "bg-amber-100 text-amber-800" :
                        "bg-red-100 text-red-800"
                      )}>
                        {user.status === 'approved' ? 'Aprovado' :
                         user.status === 'pending' ? 'Pendente' : 'Rejeitado'}
                      </span>
                    </td>
                    <td className="py-3 px-4 text-sm text-slate-500">
                      {new Date(user.createdAt).toLocaleDateString('pt-BR')}
                    </td>
                    <td className="py-3 px-4">
                      {user.username === 'admin' || user.name === 'Administrador Principal' ? (
                        <span className="text-xs text-slate-400">Protegido</span>
                      ) : (
                        <button
                          onClick={() => handleDeleteUser(user.id)}
                          className="rounded-lg p-2 text-slate-400 transition hover:bg-red-50 hover:text-red-600"
                          title="Excluir usuário"
                        >
                          <Trash2 className="h-4 w-4" />
                        </button>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );

  const renderExport = () => {
    const selectedFormat = EXPORT_FORMAT_OPTIONS.find((option) => option.id === exportFormat)!;
    const isBaseFormat = exportFormat === 'base';

    return (
      <div className="mx-auto max-w-6xl space-y-4">
        <div className="rounded-xl border border-slate-200 bg-white p-6 shadow-sm">
          <div className="flex items-start gap-3">
            <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-emerald-50 text-emerald-600">
              <FileOutput className="h-5 w-5" />
            </div>
            <div>
              <h3 className="text-lg font-semibold text-slate-800">Exportação de Dados</h3>
              <p className="text-sm text-slate-500">Gere arquivos .xlsx prontos para Excel com filtros e formatos personalizados</p>
            </div>
          </div>
        </div>

        <div className="grid grid-cols-1 gap-4 xl:grid-cols-[300px,1fr]">
          <div className="space-y-4">
            <div className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm">
              <h4 className="mb-3 flex items-center gap-2 text-sm font-semibold text-slate-700">
                <FileOutput className="h-4 w-4 text-blue-600" />
                Formato de Exportação
              </h4>
              <div className="space-y-2">
                {EXPORT_FORMAT_OPTIONS.map((option) => (
                  <button
                    key={option.id}
                    onClick={() => setExportFormat(option.id)}
                    className={cn(
                      'w-full rounded-lg border px-4 py-3 text-left transition',
                      exportFormat === option.id
                        ? 'border-blue-500 bg-blue-50 shadow-sm'
                        : 'border-slate-200 bg-white hover:bg-slate-50',
                    )}
                  >
                    <p className="text-sm font-medium text-slate-800">{option.title}</p>
                    <p className="mt-1 text-xs text-slate-500">{option.description}</p>
                  </button>
                ))}
              </div>
            </div>

            <div className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm">
              <h4 className="mb-3 text-sm font-semibold text-slate-700">Nome do Arquivo</h4>
              <div className="flex items-center gap-2">
                <input
                  type="text"
                  value={exportFileName}
                  onChange={(e) => setExportFileName(cleanText(e.target.value))}
                  className="flex-1 rounded-lg border border-slate-200 px-3 py-2 text-sm focus:border-red-800 focus:ring-2 focus:ring-red-800/20 outline-none"
                />
                <span className="rounded-lg bg-slate-100 px-3 py-2 text-xs text-slate-500">.xlsx</span>
              </div>
              <p className="mt-3 text-xs text-slate-400">
                Será salvo como: {cleanText(exportFileName || 'controle_abastecimento')}_{new Date().toISOString().slice(0, 10)}.xlsx
              </p>
            </div>
          </div>

          <div className="space-y-4">
            <div className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm">
              <div className="flex items-center justify-between gap-4">
                <h4 className="flex items-center gap-2 text-sm font-semibold text-slate-700">
                  <Filter className="h-4 w-4 text-slate-500" />
                  Filtros de Exportação
                </h4>
                <span className="text-sm font-semibold text-slate-700">{activeExportFilters.length}</span>
              </div>
              {activeExportFilters.length > 0 && (
                <div className="mt-3 flex flex-wrap gap-2">
                  {activeExportFilters.map((item, index) => (
                    <span key={`${item.label}-${item.value}-${index}`} className="rounded-full bg-slate-100 px-2.5 py-1 text-xs text-slate-600">
                      {item.label}: {item.value}
                    </span>
                  ))}
                </div>
              )}
            </div>

            {isBaseFormat && (
              <details className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm" open>
                <summary className="flex cursor-pointer list-none items-center justify-between gap-4 text-sm font-semibold text-slate-700">
                  <span className="flex items-center gap-2">
                    <Database className="h-4 w-4 text-slate-500" />
                    Selecionar Colunas para Exportar
                  </span>
                  <span className="text-xs text-slate-400">{selectedBaseColumns.length}/{BASE_EXPORT_COLUMNS.length} selecionadas</span>
                </summary>
                <div className="mt-4 grid grid-cols-2 gap-2 md:grid-cols-3">
                  {BASE_EXPORT_COLUMNS.map((column) => (
                    <label key={column.key} className="flex items-center gap-2 rounded-lg border border-slate-200 px-3 py-2 text-sm text-slate-700">
                      <input
                        type="checkbox"
                        checked={selectedBaseColumns.includes(column.key)}
                        onChange={() => toggleBaseExportColumn(column.key)}
                        className="h-4 w-4 rounded border-slate-300 text-red-800 focus:ring-red-800/30"
                      />
                      <span>{column.label}</span>
                    </label>
                  ))}
                </div>
              </details>
            )}

            <div className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm">
              <h4 className="mb-4 flex items-center gap-2 text-sm font-semibold text-slate-700">
                <FileOutput className="h-4 w-4 text-emerald-600" />
                Resumo da Exportação
              </h4>

              <div className="grid grid-cols-2 gap-3 md:grid-cols-4">
                <div className="rounded-xl bg-blue-50 p-4 text-center">
                  <p className="text-2xl font-bold text-blue-700">{exportPreview.records}</p>
                  <p className="text-xs text-blue-500">Registros</p>
                </div>
                <div className="rounded-xl bg-emerald-50 p-4 text-center">
                  <p className="text-2xl font-bold text-emerald-700">{formatNumber(exportPreview.litros, 0)}</p>
                  <p className="text-xs text-emerald-500">Litros</p>
                </div>
                <div className="rounded-xl bg-amber-50 p-4 text-center">
                  <p className="text-2xl font-bold text-amber-700">{formatCurrency(exportPreview.valor)}</p>
                  <p className="text-xs text-amber-500">Valor Total</p>
                </div>
                <div className="rounded-xl bg-slate-50 p-4 text-center">
                  <p className="text-2xl font-bold text-slate-700">{exportPreview.columns.length}</p>
                  <p className="text-xs text-slate-500">Colunas</p>
                </div>
              </div>

              <div className="mt-4 rounded-xl bg-slate-50 p-4">
                <p className="mb-2 text-xs font-semibold uppercase tracking-[0.18em] text-slate-500">Colunas na planilha</p>
                <div className="flex flex-wrap gap-2">
                  {exportPreview.columns.map((column) => (
                    <span key={column} className="rounded-md border border-slate-200 bg-white px-2 py-1 text-[11px] text-slate-600">
                      {column}
                    </span>
                  ))}
                </div>
              </div>

              <div className="mt-4 rounded-xl bg-slate-50 p-4">
                <p className="mb-2 text-xs font-semibold uppercase tracking-[0.18em] text-slate-500">Abas que serão criadas no XLSX</p>
                <div className="flex flex-wrap gap-2">
                  {exportPreview.sheetNames.map((sheet) => (
                    <span
                      key={sheet}
                      className={cn(
                        'rounded-md px-2 py-1 text-[11px] font-medium',
                        sheet === selectedFormat.sheetName ? 'bg-emerald-100 text-emerald-700' : 'bg-slate-200 text-slate-600',
                      )}
                    >
                      {sheet}
                    </span>
                  ))}
                </div>
              </div>
            </div>

            <button
              onClick={handleExportExcelFile}
              disabled={exportPreview.rows.length === 0 || exportPreview.columns.length === 0}
              className={cn(
                'flex w-full items-center justify-center gap-2 rounded-xl px-4 py-4 text-sm font-medium transition',
                exportPreview.rows.length === 0 || exportPreview.columns.length === 0
                  ? 'cursor-not-allowed bg-slate-200 text-slate-400'
                  : 'bg-blue-600 text-white hover:bg-blue-700',
              )}
            >
              <Download className="h-5 w-5" />
              Exportar para Excel (.xlsx)
            </button>

            {exportPreview.rows.length === 0 && (
              <div className="flex items-center gap-2 rounded-xl border border-amber-300 bg-amber-50 px-4 py-3 text-sm text-amber-700">
                <AlertCircle className="h-4 w-4" />
                Nenhum dado encontrado com os filtros aplicados.
              </div>
            )}
          </div>
        </div>
      </div>
    );
  };

  // Main render
  if (!currentUser) {
    return (
      <>
        {renderLogin()}
        {/* Notifications */}
        <div className="fixed top-4 right-4 space-y-2 z-50">
          {notifications.map(n => (
            <div
              key={n.id}
              className={cn(
                "px-4 py-3 rounded-lg shadow-lg flex items-center gap-2 animate-pulse",
                n.type === 'success' ? "bg-green-600 text-white" :
                n.type === 'error' ? "bg-red-600 text-white" :
                n.type === 'warning' ? "bg-amber-600 text-white" :
                "bg-blue-600 text-white"
              )}
            >
              {n.type === 'success' && <CheckCircle className="w-5 h-5" />}
              {n.type === 'error' && <AlertCircle className="w-5 h-5" />}
              {n.type === 'warning' && <AlertCircle className="w-5 h-5" />}
              {n.type === 'info' && <Bell className="w-5 h-5" />}
              {n.message}
            </div>
          ))}
        </div>
      </>
    );
  }

  const navItems = [
    { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
    { id: 'database', label: 'Base de Dados', icon: Database },
    { id: 'budget', label: 'Orçamento', icon: Wallet },
    { id: 'rateio', label: 'Rateio CC', icon: Scale },
    { id: 'equipments', label: 'Equipamentos', icon: Wrench },
    { id: 'history', label: 'Histórico', icon: History },
    { id: 'preenchimento', label: 'Preenchimento', icon: FileInput },
    { id: 'import', label: 'Importação', icon: Upload },
    { id: 'export', label: 'Exportação', icon: Download },
  ];

  const adminNavItems = [
    { id: 'diesel', label: 'Valor Diesel', icon: Fuel },
    { id: 'users', label: 'Usuários', icon: Users },
  ];



  return (
    <div className="min-h-screen bg-slate-100 lg:flex">
      <aside
        className={cn(
          "fixed inset-y-0 left-0 z-50 overflow-hidden bg-slate-900 text-white transition-all duration-300 ease-in-out lg:sticky lg:top-0 lg:z-auto lg:min-h-screen",
          sidebarVisible
            ? "w-72 translate-x-0 border-r border-slate-800 shadow-2xl lg:shadow-none"
            : "-translate-x-full lg:w-0 lg:translate-x-0"
        )}
      >
        <div className="flex h-full w-72 flex-col">
          <div className="border-b border-slate-800 px-5 py-5">
            {/* ESPAÇO PARA A LOGO — substitua por: <img src="/sua-logo.png" alt="Logo" className="h-full w-full object-contain" /> */}
            <div className="flex h-14 w-full items-center justify-center rounded-xl bg-white">
            </div>
          </div>

          <div className="flex-1 space-y-6 overflow-y-auto px-4 py-5">
            <div>
              <p className="mb-3 px-3 text-[11px] font-semibold uppercase tracking-[0.22em] text-slate-500">
                Navegação
              </p>
              <nav className="space-y-1.5">
                {navItems
                  .filter(item => {
                    if (currentUser?.role === 'admin') return true;
                    if (currentUser?.role === 'supervisor') return ['dashboard', 'budget', 'history', 'export', 'diesel'].includes(item.id);
                    return item.id === 'preenchimento';
                  })
                  .map(item => {
                    const Icon = item.icon;
                    const isActive = currentPage === item.id;
                    return (
                      <button
                        key={item.id}
                        onClick={() => setCurrentPage(item.id)}
                        className={cn(
                          "flex w-full items-center gap-3 rounded-xl px-3 py-3 text-left text-sm font-medium transition",
                          isActive
                            ? "bg-red-800 text-white shadow-lg shadow-red-950/20"
                            : "text-slate-300 hover:bg-slate-800 hover:text-white"
                        )}
                      >
                        <Icon className="h-4 w-4 shrink-0" />
                        <span>{item.label}</span>
                      </button>
                    );
                  })}
              </nav>
            </div>

            {(currentUser.role === 'admin' || currentUser.role === 'supervisor') && (
              <div>
                <p className="mb-3 px-3 text-[11px] font-semibold uppercase tracking-[0.22em] text-slate-500">
                  Administração
                </p>
                <nav className="space-y-1.5">
                  {adminNavItems
                    .filter(item => currentUser.role === 'admin' || item.id === 'diesel')
                    .map(item => {
                    const Icon = item.icon;
                    const isActive = currentPage === item.id;
                    return (
                      <button
                        key={item.id}
                        onClick={() => setCurrentPage(item.id)}
                        className={cn(
                          "flex w-full items-center gap-3 rounded-xl px-3 py-3 text-left text-sm font-medium transition",
                          isActive
                            ? "bg-red-800 text-white shadow-lg shadow-red-950/20"
                            : "text-slate-300 hover:bg-slate-800 hover:text-white"
                        )}
                      >
                        <Icon className="h-4 w-4 shrink-0" />
                        <span>{item.label}</span>
                      </button>
                    );
                  })}
                </nav>
              </div>
            )}
          </div>

          <div className="border-t border-slate-800 px-4 py-4">
            <button
              onClick={() => setCurrentPage('profile')}
              className={cn(
                "flex w-full items-center gap-3 rounded-xl px-3 py-3 text-left text-sm font-medium transition mb-2",
                currentPage === 'profile'
                  ? "bg-red-800 text-white shadow-lg shadow-red-950/20"
                  : "text-slate-300 hover:bg-slate-800 hover:text-white"
              )}
            >
              {currentUser.avatar ? (
                <img src={currentUser.avatar} alt="Avatar" className="h-7 w-7 rounded-full object-cover shrink-0" />
              ) : (
                <div className="h-7 w-7 rounded-full bg-slate-600 flex items-center justify-center shrink-0 text-xs font-bold text-white">
                  {(currentUser.name || '?')[0].toUpperCase()}
                </div>
              )}
              <div className="min-w-0">
                <p className="truncate font-semibold">{currentUser.name}</p>
                <p className="truncate text-xs text-slate-400">{currentUser.funcao || getRoleLabel(currentUser.role)}</p>
              </div>
            </button>
            <button
              onClick={handleLogout}
              className="flex w-full items-center justify-center gap-2 rounded-xl border border-slate-700 bg-slate-900 px-3 py-2.5 text-sm font-medium text-white transition hover:border-slate-600 hover:bg-slate-700"
            >
              <LogOut className="h-4 w-4" />
              Sair
            </button>
          </div>
        </div>
      </aside>

      {/* Overlay para mobile */}
      {sidebarVisible && (
        <div
          className="fixed inset-0 z-40 bg-black/50 lg:hidden"
          onClick={() => setSidebarVisible(false)}
        />
      )}

      <div className="flex min-w-0 max-w-full flex-1 flex-col overflow-x-hidden">
        <div className="sticky top-0 z-30 border-b border-slate-200 bg-white/95 backdrop-blur">
          <div className="flex flex-col gap-3 px-4 py-4 sm:flex-row sm:items-center sm:justify-between lg:px-6">
            <div className="flex items-start gap-3">
              <button
                onClick={() => setSidebarVisible((prev) => !prev)}
                className="flex h-11 w-11 items-center justify-center rounded-xl border border-slate-200 bg-white text-slate-700 shadow-sm transition hover:border-slate-300 hover:bg-slate-50"
                aria-label="Mostrar ou esconder menu lateral"
                title="Mostrar ou esconder menu lateral"
              >
                <Menu className="h-5 w-5" />
              </button>
              <div className="min-w-0">
                <p className="truncate text-xs font-semibold uppercase tracking-[0.18em] text-slate-500 sm:tracking-[0.22em]">Painel de controle</p>
                <h2 className="mt-1 truncate text-2xl font-semibold text-slate-900">
                  {([...navItems, ...(currentUser.role === 'admin' ? adminNavItems : [])].find(item => item.id === currentPage)?.label) || 'Sistema'}
                </h2>
              </div>
            </div>
            <div className="flex min-w-0 flex-wrap items-center gap-2 text-sm text-slate-600 sm:gap-3">
              <div className="max-w-full rounded-full bg-white px-3 py-1.5 shadow-sm ring-1 ring-slate-200">
                Preço Diesel: <span className="font-semibold text-slate-900">{formatCurrency(dieselPrice.price)}/L</span>
              </div>
              <div className="rounded-full bg-white px-3 py-1.5 shadow-sm ring-1 ring-slate-200">
                {abastecimentos.length} registros
              </div>
            </div>
          </div>
        </div>

        <main className="min-w-0 max-w-full flex-1 overflow-x-hidden p-3 sm:p-4 lg:p-6">
          {currentPage === 'dashboard' && renderDashboard()}
          {currentPage === 'database' && renderDatabase()}
          {currentPage === 'budget' && renderBudget()}
          {currentPage === 'rateio' && renderRateio()}
          {currentPage === 'equipments' && renderEquipments()}
          {currentPage === 'history' && renderHistory()}
          {currentPage === 'preenchimento' && renderPreenchimento()}
          {currentPage === 'import' && renderImport()}
          {currentPage === 'export' && renderExport()}
          {currentPage === 'diesel' && renderDieselPrice()}
          {currentPage === 'users' && renderUsers()}
          {currentPage === 'profile' && renderProfile()}
        </main>

        <footer className="max-w-full overflow-hidden border-t border-slate-200 bg-white px-4 py-3 lg:px-6">
          <div className="flex min-w-0 flex-col gap-2 text-xs text-slate-500 sm:flex-row sm:items-center sm:justify-between">
            <p>Controle de Abastecimento v1.0 — Sistema Corporativo</p>
            <div className="flex flex-wrap items-center gap-4">
              <p>Preço Diesel: {formatCurrency(dieselPrice.price)}/L</p>
              <p>{abastecimentos.length} registros</p>
            </div>
          </div>
        </footer>
      </div>

      <div className="fixed right-4 top-4 z-50 space-y-2">
        {notifications.map(n => (
          <div
            key={n.id}
            className={cn(
              "w-[calc(100vw-2rem)] max-w-sm rounded-lg px-4 py-3 shadow-lg flex items-center gap-2 sm:w-auto sm:min-w-72",
              n.type === 'success' ? "bg-green-600 text-white" :
              n.type === 'error' ? "bg-red-600 text-white" :
              n.type === 'warning' ? "bg-amber-600 text-white" :
              "bg-blue-600 text-white"
            )}
          >
            {n.type === 'success' && <CheckCircle className="w-5 h-5" />}
            {n.type === 'error' && <AlertCircle className="w-5 h-5" />}
            {n.type === 'warning' && <AlertCircle className="w-5 h-5" />}
            {n.type === 'info' && <Bell className="w-5 h-5" />}
            {n.message}
          </div>
        ))}
      </div>
    </div>
  );
}
