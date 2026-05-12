export type UserRole = 'admin' | 'operator' | 'supervisor';

export type UserStatus = 'pending' | 'approved' | 'rejected';

export interface User {
  id: number;
  username: string;
  password: string;
  role: UserRole;
  status: UserStatus;
  name: string;
  email?: string;
  funcao?: string;
  avatar?: string;
  createdAt: string;
}

export interface Equipment {
  id: number;
  equipment: string;
  plate: string;
  ccNovo: string[];
  gerencia: string;
  areaLotacao: string;
  area: string;
  fornecedor: string;
  createdAt: string;
}

export interface Rateio {
  id: number;
  equipmentId: number;
  equipment: string;
  description: string;
  items: RateioItem[];
  createdAt: string;
}

export interface RateioItem {
  gerencia: string;
  ccNovo: string;
  description: string;
  percentage: number;
}

export interface Budget {
  id: number;
  diretoria: string;
  periodo: string;
  orcamento: number;
  realizado: number;
  dataInicio: string;
  dataFim: string;
  createdAt: string;
}

export interface Abastecimento {
  id: number;
  ccNovo: string;
  diretoria: string;
  gerencia: string;
  areaLotacao: string;
  fornecedor: string;
  equipamento: string;
  area: string;
  semana: string;
  data: string;
  litros: number;
  valor: number;
  observacoes?: string;
  createdBy: string;
  createdAt: string;
  rateioInfo?: {
    rateado: boolean;
    rateioId?: number;
    gerenciaRateio?: string;
    ccRateio?: string;
    percentage?: number;
  };
}

export interface DieselPrice {
  id: number;
  price: number;
  updatedAt: string;
  updatedBy: string;
}

export interface FilterState {
  ccNovo: string[];
  diretoria: string[];
  gerencia: string[];
  areaLotacao: string[];
  fornecedor: string[];
  equipamento: string[];
  dia: string[];
  mes: string[];
  semana: string[];
  ano: string[];
}

export interface Notification {
  id: number;
  type: 'success' | 'error' | 'info' | 'warning';
  message: string;
}
