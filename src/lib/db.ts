import Dexie, { type Table } from 'dexie';

export interface Statement {
  id?: number;
  vendor: string;
  amount: number;
  date: string; // YYYY-MM
  issueDate?: string;
  statementNumber?: string;
  status: 'pending' | 'matched' | 'error';
  localFilePath?: string;
  outlookId?: string;
  matchedInvoiceId?: number;
  subject?: string;
  projectId?: string;
  invoiceNumbers?: string[]; // Extracted from PDF content
  items?: { invoiceNumber: string, amount: number }[];
}

export interface Invoice {
  id?: number;
  vendor: string;
  amount: number;
  date: string;
  invoiceNumber: string;
  localFilePath?: string;
  outlookId?: string;
  isProcessed: boolean;
  matchedStatementId?: number;
  projectId?: string;
}

export interface Project {
  id: string; // Project Number/ID
  name: string;
  createdAt: string;
}

export interface Setting {
  key: string;
  value: any;
}

export class AppDatabase extends Dexie {
  statements!: Table<Statement>;
  invoices!: Table<Invoice>;
  projects!: Table<Project>;
  settings!: Table<Setting>;

  constructor() {
    super('FinanceAssistantDB');
    this.version(3).stores({
      statements: '++id, vendor, amount, date, status, outlookId, projectId, *invoiceNumbers',
      invoices: '++id, vendor, amount, date, invoiceNumber, isProcessed, projectId',
      projects: 'id, name',
      settings: 'key'
    });
  }
}

export const db = new AppDatabase();
