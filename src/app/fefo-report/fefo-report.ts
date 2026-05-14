import { Component, Inject, OnInit, PLATFORM_ID, ViewChild, computed, signal } from '@angular/core';
import { CommonModule, isPlatformBrowser } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { Table, TableModule } from 'primeng/table';
import { ButtonModule } from 'primeng/button';
import { InputTextModule } from 'primeng/inputtext';
import { ProgressSpinnerModule } from 'primeng/progressspinner';
import { TooltipModule } from 'primeng/tooltip';
import { BreadcrumbModule } from 'primeng/breadcrumb';
import { ToastModule } from 'primeng/toast';
import { TagModule } from 'primeng/tag';
import { MessageService, MenuItem } from 'primeng/api';
import { PrimeNG } from 'primeng/config';
import * as XLSX from 'xlsx';

type FileType = 'pedidos' | 'inventario' | 'reglas';

interface FileData {
  name: string;
  data: string | ArrayBuffer | null;
  size: number;
  uploadDate: Date;
}

interface PedidosRow {
  PRO_CREATED_ON: string;
  OBO_ORDER_DATE: string;
  PRO_ID: string;
  OBO_ID: string;
  OBO_SOLDTO: string;
  OBO_SHIPTO_CODE: string;
  OBO_SHIPTO_CODE2: string;
  OBO_SHIPTO_NAME: string;
  OBO_TRACKING_NUMBER: string;
  OBO_INVOICE: string;
  OBO_ORDER: string;
  OBO_ORDER2: string;
  OBO_PURCHASE_ORDER: string;
  MAI_SKU: string;
  MAI_DESCRIPTION: string;
  LPI_LOT: string;
  LPI_EXPIRY_DATE: string;
  PKO_EACH_QTY: string;
  PKO_CASE_COUNT: string;
  PKO_MODULE_QTY: string;
  PKO_MODULE_CASE_COUNT: string;
  STRATEGY_NAME: string;
  DIFF_DAYS: string;
  [key: string]: any;
}

interface InventarioRow {
  SKU: string;
  DESCRIPCION: string;
  LOTE: string;
  FECHA_VENCIMIENTO: string;
  CATEGORIA: string;
  DISPONIBLE: string;
  FPC: string;
}

interface ReglasRow {
  ESTRATEGIA: string;
  CRUCE_CONACTENA: string;
  TVU_REGLA: string;
}

interface FefoResultRow {
  proCreatedOn: string;
  oboOrderDate: string;
  proId: string;
  oboId: string;
  oboSoldto: string;
  oboShiptoCode: string;
  oboShiptoCode2: string;
  oboShiptoName: string;
  oboTrackingNumber: string;
  oboInvoice: string;
  oboOrder: string;
  oboOrder2: string;
  oboPurchaseOrder: string;
  sku: string;
  descripcion: string;
  lote: string;
  fechaVencimientoPedido: string;
  fechaMasProximaVencer: string;
  cumplimientoFefo: string;
  justificacionFefo: string;
  pkoEachQty: string;
  pkoCaseCount: string;
  pkoModuleQty: string;
  pkoModuleCaseCount: string;
  strategyName: string;
  diffDays: string;
  cruceConcatena: string;
  tvuRegla: string;
  fechaMinDespacho: string;
  cumpleTvuValor: string;
  comentariosMdlz: string;
  decisionFinal: string;
  observacion: string;
}

const pedidosMapping: Record<string, string[]> = {
  PRO_CREATED_ON: ['PRO_CREATED_ON', 'PRO CREATED ON'],
  OBO_ORDER_DATE: ['OBO_ORDER_DATE', 'OBO ORDER DATE'],
  PRO_ID: ['PRO_ID', 'PRO ID'],
  OBO_ID: ['OBO_ID', 'OBO ID'],
  OBO_SOLDTO: ['OBO_SOLDTO', 'OBO SOLDTO'],
  OBO_SHIPTO_CODE: ['OBO_SHIPTO_CODE', 'OBO SHIPTO CODE'],
  OBO_SHIPTO_CODE2: ['OBO_SHIPTO_CODE2', 'OBO SHIPTO CODE2'],
  OBO_SHIPTO_NAME: ['OBO_SHIPTO_NAME', 'OBO SHIPTO NAME'],
  OBO_TRACKING_NUMBER: ['OBO_TRACKING_NUMBER', 'OBO TRACKING NUMBER'],
  OBO_INVOICE: ['OBO_INVOICE', 'OBO INVOICE'],
  OBO_ORDER: ['OBO_ORDER', 'OBO ORDER'],
  OBO_ORDER2: ['OBO_ORDER2', 'OBO ORDER2'],
  OBO_PURCHASE_ORDER: ['OBO_PURCHASE_ORDER', 'OBO PURCHASE ORDER'],
  MAI_SKU: ['MAI_SKU', 'SKU', 'CODIGO SKU'],
  MAI_DESCRIPTION: ['MAI_DESCRIPTION', 'DESCRIPCION', 'DESCRIPCIÓN'],
  LPI_LOT: ['LPI_LOT', 'LOTE'],
  LPI_EXPIRY_DATE: ['LPI_EXPIRY_DATE', 'FECHA DE VENCIMIENTO', 'EXPIRY DATE'],
  STRATEGY_NAME: ['STRATEGY_NAME', 'ESTRATEGIA'],
  DIFF_DAYS: ['DIFF_DAYS', 'DIAS', 'DIFERENCIA DIAS'],
  PKO_EACH_QTY: ['PKO_EACH_QTY', 'EACH QTY'],
  PKO_CASE_COUNT: ['PKO_CASE_COUNT', 'CASE COUNT'],
  PKO_MODULE_QTY: ['PKO_MODULE_QTY', 'MODULE QTY'],
  PKO_MODULE_CASE_COUNT: ['PKO_MODULE_CASE_COUNT', 'MODULE CASE COUNT'],
};

const inventarioMapping: Record<string, string[]> = {
  SKU: ['SKU', 'CODIGO SKU'],
  DESCRIPCION: ['DESCRIPCION', 'DESCRIPCIÓN', 'MAI_DESCRIPTION'],
  LOTE: ['LOTE', 'LPI_LOT'],
  FECHA_VENCIMIENTO: ['FECHA DE VENCIMIENTO', 'FECHA DE VENCIMIENTO ', 'LPI_EXPIRY_DATE'],
  CATEGORIA: ['CATEGORIA DE MATERIAL', 'CATEGORÍA DE MATERIAL', 'CATEGORIA'],
  DISPONIBLE: ['DISPONIBLE', 'AVAILABLE'],
  FPC: ['FPC'],
};

const reglasMapping: Record<string, string[]> = {
  ESTRATEGIA: ['ESTRATEGIA', 'STRATEGY_NAME'],
  CRUCE_CONACTENA: ['CRUCE CONACTENA', 'CRUCECONACTENA'],
  TVU_REGLA: ['TVU REGLA', 'TVU REGla', 'TVU'],
};

const requiredColumnsByType: Record<FileType, string[]> = {
  pedidos: [
    'PRO_CREATED_ON',
    'MAI_SKU',
    'LPI_LOT',
    'LPI_EXPIRY_DATE',
    'STRATEGY_NAME',
    'DIFF_DAYS',
  ],
  inventario: ['SKU', 'LOTE', 'FECHA_VENCIMIENTO', 'DISPONIBLE'],
  reglas: ['ESTRATEGIA', 'CRUCE_CONACTENA', 'TVU_REGLA'],
};

@Component({
  selector: 'app-fefo-report',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    TableModule,
    ButtonModule,
    InputTextModule,
    ProgressSpinnerModule,
    TooltipModule,
    BreadcrumbModule,
    ToastModule,
    TagModule,
  ],
  providers: [MessageService],
  templateUrl: './fefo-report.html',
  styleUrl: './fefo-report.css',
})
export class FefoReport implements OnInit {
  @ViewChild('dt1') dt1!: Table;

  pedidosFile = signal<FileData | null>(null);
  inventarioFile = signal<FileData | null>(null);
  reglasFile = signal<FileData | null>(null);
  loading = signal(false);
  analysisResult = signal<FefoResultRow[]>([]);

  isDraggingPedidos = signal(false);
  isDraggingInventario = signal(false);
  isDraggingReglas = signal(false);

  searchValue = '';
  isBrowser: boolean;

  readonly breadcrumbItems: MenuItem[] = [{ label: 'Reporte FEFO', routerLink: '/fefo-report' }];
  readonly breadcrumbHome: MenuItem = { icon: 'pi pi-home', label: 'Inicio', routerLink: '/' };

  filteredResultados = computed(() => {
    const rows = this.analysisResult();
    if (!this.searchValue.trim()) return rows;
    const q = this.searchValue.toLowerCase();
    return rows.filter(
      (r) =>
        r.sku.toLowerCase().includes(q) ||
        r.descripcion.toLowerCase().includes(q) ||
        r.strategyName.toLowerCase().includes(q) ||
        r.decisionFinal.toLowerCase().includes(q),
    );
  });

  constructor(
    @Inject(PLATFORM_ID) platformId: Object,
    private messageService: MessageService,
    private primeng: PrimeNG,
  ) {
    this.isBrowser = isPlatformBrowser(platformId);
  }

  ngOnInit(): void {
    if (this.primeng && typeof this.primeng.setTranslation === 'function') {
      this.primeng.setTranslation({
        startsWith: 'Empieza con',
        contains: 'Contiene',
        notContains: 'No contiene',
        endsWith: 'Termina con',
        equals: 'Igual a',
        notEquals: 'Distinto de',
        noFilter: 'Sin filtro',
        lt: 'Menor que',
        lte: 'Menor o igual que',
        gt: 'Mayor que',
        gte: 'Mayor o igual que',
        is: 'Es',
        isNot: 'No es',
        before: 'Antes',
        after: 'Después',
        clear: 'Limpiar',
        apply: 'Aplicar',
        matchAll: 'Cumplir todas',
        matchAny: 'Cumplir alguna',
        addRule: 'Agregar regla',
        removeRule: 'Eliminar regla',
      });
    }
  }

  onSearch() {
    if (this.dt1) {
      this.dt1.filterGlobal(this.searchValue, 'contains');
    }
  }

  clearFilters() {
    if (this.dt1) {
      this.dt1.clear();
      this.dt1.filterGlobal('', 'contains');
    }
    this.searchValue = '';
    this.messageService.add({
      severity: 'info',
      summary: 'Filtros limpiados',
      detail: 'Se han eliminado los filtros de búsqueda',
      life: 3000,
    });
  }

  getTotalRegistros(): number {
    return this.analysisResult().length;
  }

  getFefoCumplen(): number {
    return this.analysisResult().filter(
      (r) => String(r.cumplimientoFefo).trim().toLowerCase() !== 'no cumple',
    ).length;
  }

  getFefoNoCumplen(): number {
    return this.getTotalRegistros() - this.getFefoCumplen();
  }

  getTvuCumplen(): number {
    return this.analysisResult().filter(
      (r) => String(r.comentariosMdlz).trim().toLowerCase() === 'cumple',
    ).length;
  }

  getTvuNoCumplen(): number {
    return this.getTotalRegistros() - this.getTvuCumplen();
  }

  getDecisionSeverity(value: string): 'success' | 'warn' | 'danger' {
    const v = String(value ?? '')
      .trim()
      .toLowerCase();
    if (v === 'cumple' || v === 'cumple la linea') return 'success';
    if (v === 'no cumple' || v === 'revisar la linea') return 'danger';
    return 'warn';
  }

  getTvuSeverity(value: string): 'success' | 'warn' | 'danger' {
    const v = String(value ?? '')
      .trim()
      .toLowerCase();
    if (v === 'cumple') return 'success';
    if (v === 'no cumple') return 'danger';
    return 'warn';
  }

  isNegativeTvu(value: unknown): boolean {
    if (value === null || value === undefined || value === '') return false;
    const n = typeof value === 'number' ? value : Number(String(value).replace(',', '.').trim());
    return !Number.isNaN(n) && n < 0;
  }

  onRowExpand(_event: unknown): void {
    // Hook disponible para futuras acciones al expandir filas
  }

  formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  onDragOver(event: DragEvent, type: FileType) {
    event.preventDefault();
    event.stopPropagation();
    if (type === 'pedidos') this.isDraggingPedidos.set(true);
    if (type === 'inventario') this.isDraggingInventario.set(true);
    if (type === 'reglas') this.isDraggingReglas.set(true);
  }

  onDragLeave(event: DragEvent, type: FileType) {
    event.preventDefault();
    event.stopPropagation();
    if (type === 'pedidos') this.isDraggingPedidos.set(false);
    if (type === 'inventario') this.isDraggingInventario.set(false);
    if (type === 'reglas') this.isDraggingReglas.set(false);
  }

  onDrop(event: DragEvent, type: FileType) {
    event.preventDefault();
    event.stopPropagation();
    if (type === 'pedidos') this.isDraggingPedidos.set(false);
    if (type === 'inventario') this.isDraggingInventario.set(false);
    if (type === 'reglas') this.isDraggingReglas.set(false);

    const files = event.dataTransfer?.files;
    if (files && files.length > 0) {
      this.processFile(files[0], type);
    }
  }

  onFileUpload(type: FileType, event: Event) {
    if (!this.isBrowser) return;
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    this.processFile(file, type);
  }

  removeFile(type: FileType) {
    if (type === 'pedidos') this.pedidosFile.set(null);
    if (type === 'inventario') this.inventarioFile.set(null);
    if (type === 'reglas') this.reglasFile.set(null);
    this.analysisResult.set([]);
  }

  private processFile(file: File, type: FileType) {
    if (!this.isBrowser) return;

    const validExtensions = ['.csv', '.xlsx', '.xls', '.txt'];
    const fileExt = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    if (!validExtensions.includes(fileExt)) {
      this.messageService.add({
        severity: 'error',
        summary: 'Formato no soportado',
        detail: 'Use archivos CSV, Excel o TXT',
        life: 5000,
      });
      return;
    }

    this.analysisResult.set([]);

    const reader = new FileReader();
    reader.onload = () => {
      const data = reader.result;
      const headers = this.extractHeaders(data, file.name);
      const count = this.estimateRecordCount(data, file.name);

      const validation = this.validarArchivoSegunCampo(headers, type);
      if (!validation.ok) {
        this.messageService.add({
          severity: 'error',
          summary: validation.templateError ? 'Plantilla incorrecta' : 'Columnas faltantes',
          detail:
            validation.templateError || `Faltan columnas: ${validation.missingCols.join(', ')}`,
          life: 7000,
        });
        return;
      }

      if (count <= 0) {
        this.messageService.add({
          severity: 'error',
          summary: 'Sin datos',
          detail: 'El archivo no contiene filas válidas',
          life: 5000,
        });
        return;
      }

      const fileData: FileData = { name: file.name, data, size: file.size, uploadDate: new Date() };
      if (type === 'pedidos') this.pedidosFile.set(fileData);
      if (type === 'inventario') this.inventarioFile.set(fileData);
      if (type === 'reglas') this.reglasFile.set(fileData);

      this.messageService.add({
        severity: 'success',
        summary: 'Archivo cargado',
        detail: `${type.toUpperCase()}: ${count} registros detectados`,
        life: 3000,
      });
    };

    reader.onerror = () => {
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: 'No se pudo leer el archivo',
        life: 5000,
      });
    };

    if (fileExt === '.xlsx' || fileExt === '.xls') reader.readAsBinaryString(file);
    else reader.readAsText(file, 'UTF-8');
  }

  private extractHeaders(data: string | ArrayBuffer | null, fileName: string): string[] {
    const ext = fileName.split('.').pop()?.toLowerCase();
    if (ext === 'xlsx' || ext === 'xls') {
      const workbook = XLSX.read(data, { type: 'binary' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: '',
        raw: false,
      }) as any[][];
      return (rows[0] || []).map((h: any) => String(h).trim());
    }
    if (typeof data === 'string') {
      const lines = data.split(/\r?\n/).filter((l) => l.trim().length > 0);
      if (!lines.length) return [];
      const delimiter = this.detectDelimiter(lines[0]);
      return lines[0].split(delimiter).map((h) => h.trim().replace(/^['"]|['"]$/g, ''));
    }
    return [];
  }

  private estimateRecordCount(data: string | ArrayBuffer | null, fileName: string): number {
    const ext = fileName.split('.').pop()?.toLowerCase();
    if (ext === 'xlsx' || ext === 'xls') {
      const workbook = XLSX.read(data, { type: 'binary' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: '',
        raw: false,
      }) as any[][];
      return Math.max(rows.length - 1, 0);
    }
    if (typeof data === 'string') {
      return Math.max(data.split(/\r?\n/).filter((l) => l.trim().length > 0).length - 1, 0);
    }
    return 0;
  }

  private detectDelimiter(headerLine: string): string {
    if (headerLine.includes('\t')) return '\t';
    if (headerLine.includes(';')) return ';';
    return ',';
  }

  private normalizarHeader(value: string): string {
    return String(value ?? '')
      .trim()
      .replace(/^\uFEFF/, '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/\s+/g, '')
      .toLowerCase();
  }

  private evaluarCoberturaColumnas(headers: string[], mapping: Record<string, string[]>) {
    const normalized = headers.map((h) => this.normalizarHeader(h));
    const missing: string[] = [];
    for (const key of Object.keys(mapping)) {
      const aliases = mapping[key].map((a) => this.normalizarHeader(a));
      const found = normalized.some((h) =>
        aliases.some((alias) => h === alias || h.startsWith(alias)),
      );
      if (!found) missing.push(key);
    }
    return {
      isFullMatch: missing.length === 0,
      missing,
    };
  }

  private validarArchivoSegunCampo(
    headers: string[],
    type: FileType,
  ): { ok: boolean; missingCols: string[]; templateError?: string } {
    const expected =
      type === 'pedidos'
        ? pedidosMapping
        : type === 'inventario'
          ? inventarioMapping
          : reglasMapping;

    const requiredKeys = requiredColumnsByType[type];
    const expectedRequiredMapping = requiredKeys.reduce(
      (acc, key) => {
        const aliases = expected[key];
        if (aliases) acc[key] = aliases;
        return acc;
      },
      {} as Record<string, string[]>,
    );

    const oppositeCandidates: Record<FileType, Record<string, string[]>>[] = [];
    if (type !== 'pedidos') oppositeCandidates.push({ pedidos: pedidosMapping } as any);
    if (type !== 'inventario') oppositeCandidates.push({ inventario: inventarioMapping } as any);
    if (type !== 'reglas') oppositeCandidates.push({ reglas: reglasMapping } as any);

    const expectedCoverage = this.evaluarCoberturaColumnas(headers, expectedRequiredMapping);
    const missingCols = expectedCoverage.missing.map((k) => expectedRequiredMapping[k][0]);
    if (missingCols.length > 0) {
      return { ok: false, missingCols };
    }

    for (const candidate of oppositeCandidates) {
      const key = Object.keys(candidate)[0] as FileType;
      const mapping = (candidate as any)[key] as Record<string, string[]>;
      const coverage = this.evaluarCoberturaColumnas(headers, mapping);
      if (coverage.isFullMatch) {
        const nombre =
          key === 'pedidos' ? 'Pedidos' : key === 'inventario' ? 'Inventario' : 'Reglas';
        return {
          ok: false,
          missingCols: [],
          templateError: `Este archivo coincide con la plantilla de ${nombre}. Cárgalo en el campo correcto.`,
        };
      }
    }

    return { ok: true, missingCols: [] };
  }

  private parseFile(file: FileData): any[] {
    try {
      const ext = file.name.split('.').pop()?.toLowerCase();
      if (ext === 'xlsx' || ext === 'xls') {
        const wb =
          typeof file.data === 'string'
            ? XLSX.read(file.data, { type: 'binary', cellDates: true })
            : XLSX.read(file.data, { type: 'array', cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        // raw: true devuelve Date objects para celdas de fecha (con cellDates: true)
        // evitando ambigüedad de formato m/d/yy que produce años incorrectos (1928, 1929...)
        return XLSX.utils.sheet_to_json(ws, { defval: '', raw: true });
      }
      if ((ext === 'csv' || ext === 'txt') && typeof file.data === 'string') {
        const lines = file.data.split(/\r?\n/).filter((l) => l.trim().length > 0);
        if (lines.length < 2) return [];
        const delimiter = this.detectDelimiter(lines[0]);
        const headers = lines[0].split(delimiter).map((h) => h.trim().replace(/^['"]|['"]$/g, ''));
        return lines.slice(1).map((line) => {
          const cols = line.split(delimiter).map((c) => c.trim().replace(/^['"]|['"]$/g, ''));
          const row: any = {};
          headers.forEach((h, i) => {
            row[h] = cols[i] ?? '';
          });
          return row;
        });
      }
      return [];
    } catch {
      return [];
    }
  }

  private mapearColumnas(row: any, mapping: Record<string, string[]>): any {
    const out: any = {};
    const keys = Object.keys(row);
    const normalized = keys.map((k) => this.normalizarHeader(k));

    for (const [targetKey, aliases] of Object.entries(mapping)) {
      const aliasNormalized = aliases.map((a) => this.normalizarHeader(a));
      const idx = normalized.findIndex((k) => aliasNormalized.some((a) => k === a));
      out[targetKey] = idx >= 0 ? row[keys[idx]] : '';
    }
    return out;
  }

  private parseDate(value: any): Date | null {
    if (!value) return null;
    if (value instanceof Date && !isNaN(value.getTime())) return value;

    const raw = String(value).trim();
    if (!raw) return null;

    // ISO: YYYY-MM-DD o YYYY-MM-DD HH:MM:SS (con hora sin zero-pad)
    if (/^\d{4}-\d{2}-\d{2}/.test(raw)) {
      const isoString = raw
        .replace(' ', 'T')
        .replace(/T(\d):/, 'T0$1:') // "T0:00:00" → "T00:00:00"
        .replace(/T(\d{2}):(\d):/, 'T$1:0$2:'); // minutos sin pad
      const d = new Date(isoString);
      return isNaN(d.getTime()) ? null : d;
    }

    // D/M/YY, DD/MM/YYYY, M/D/YYYY — detectar cuál por posición y magnitud
    const parts = raw.split(/[\/\-]/);
    if (parts.length >= 3) {
      let p0 = Number(parts[0]);
      let p1 = Number(parts[1]);
      let p2 = Number(parts[2]);

      if (!isNaN(p0) && !isNaN(p1) && !isNaN(p2)) {
        // Normalizar año de 2 dígitos → siempre 2000+
        if (p2 < 100) p2 = 2000 + p2;

        let dd: number, mm: number, yyyy: number;

        // Si p0 > 12, solo puede ser día (formato DD/MM/YYYY)
        // Si p1 > 12, solo puede ser día en posición 2 (MM/DD/YYYY → no aplica en Colombia)
        // Por defecto asumimos DD/MM/YYYY (formato colombiano)
        if (p0 > 12) {
          dd = p0;
          mm = p1;
          yyyy = p2; // DD/MM/YYYY
        } else if (p1 > 12) {
          mm = p0;
          dd = p1;
          yyyy = p2; // MM/DD/YYYY (US)
        } else {
          dd = p0;
          mm = p1;
          yyyy = p2; // Por defecto DD/MM/YYYY
        }

        const d = new Date(yyyy, mm - 1, dd);
        if (!isNaN(d.getTime())) return d;
      }
    }

    const fallback = new Date(raw);
    return isNaN(fallback.getTime()) ? null : fallback;
  }

  private formatDate(date: Date | null): string {
    if (!date) return '';
    const dd = String(date.getDate()).padStart(2, '0');
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const yyyy = date.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
  }

  private dateDiffInDays(a: Date, b: Date): number {
    const ms = b.getTime() - a.getTime();
    return Math.floor(ms / (1000 * 60 * 60 * 24));
  }

  private addDays(date: Date, days: number): Date {
    const out = new Date(date);
    out.setDate(out.getDate() + days);
    return out;
  }

  private parseNumber(value: any): number | null {
    const s = String(value ?? '').trim();
    if (!s) return null;
    const normalized = s.replace(/\./g, '').replace(',', '.').replace(/\s+/g, '');
    const n = Number(normalized);
    return isNaN(n) ? null : n;
  }

  private normalizeKey(value: string): string {
    return String(value ?? '')
      .trim()
      .toUpperCase()
      .replace(/\s+/g, ' ');
  }

  private buildFechaSkuKey(fecha: Date, sku: string): string {
    return `${this.formatDate(fecha)}${String(sku ?? '').trim()}`;
  }

  generateReport() {
    if (!this.pedidosFile() || !this.inventarioFile() || !this.reglasFile()) {
      this.messageService.add({
        severity: 'warn',
        summary: 'Archivos faltantes',
        detail: 'Debes subir Pedidos, Inventario y Reglas para generar el reporte FEFO',
        life: 5000,
      });
      return;
    }

    this.loading.set(true);

    setTimeout(() => {
      try {
        const pedidosRaw = this.parseFile(this.pedidosFile()!);
        const inventarioRaw = this.parseFile(this.inventarioFile()!);
        const reglasRaw = this.parseFile(this.reglasFile()!);

        if (!pedidosRaw.length || !inventarioRaw.length || !reglasRaw.length) {
          this.messageService.add({
            severity: 'error',
            summary: 'Sin datos',
            detail: 'Uno o más archivos no contienen datos válidos',
            life: 5000,
          });
          this.loading.set(false);
          return;
        }

        const pedidos = pedidosRaw.map((r) => this.mapearColumnas(r, pedidosMapping) as PedidosRow);
        const inventario = inventarioRaw.map(
          (r) => this.mapearColumnas(r, inventarioMapping) as InventarioRow,
        );
        const reglas = reglasRaw.map((r) => this.mapearColumnas(r, reglasMapping) as ReglasRow);

        const minExpiryBySku = new Map<string, Date>();
        const analisisLookup = new Map<string, number>();

        for (const inv of inventario) {
          const sku = this.normalizeKey(inv.SKU);
          const exp = this.parseDate(inv.FECHA_VENCIMIENTO);
          const disponible = this.parseNumber(inv.DISPONIBLE);
          const disponibleLine = this.parseNumber(inv.DISPONIBLE);

          // Replica la macro: solo considera inventario con disponible distinto de cero.
          if (!sku || !exp || disponible === null || disponible === 0) continue;

          const curr = minExpiryBySku.get(sku);
          if (!curr || exp.getTime() < curr.getTime()) {
            minExpiryBySku.set(sku, exp);
          }

          const key = this.buildFechaSkuKey(exp, sku);
          // El valor de referencia para FEFO/justificación se toma del disponible por clave Fecha+SKU.
          if (disponibleLine !== null) {
            const existing = analisisLookup.get(key) ?? 0;
            analisisLookup.set(key, existing + disponibleLine);
          }
        }

        const ruleByCruce = new Map<string, ReglasRow>();
        for (const reg of reglas) {
          const cruce = this.normalizeKey(reg.CRUCE_CONACTENA);
          if (cruce) ruleByCruce.set(cruce, reg);
        }

        const resultados: FefoResultRow[] = pedidos.map((p) => {
          const sku = String(p.MAI_SKU ?? '').trim();
          const strategy = String(p.STRATEGY_NAME ?? '').trim();
          const cruce = `${sku}${strategy}`.trim();

          const expPedido = this.parseDate(p.LPI_EXPIRY_DATE);
          const proCreatedOnDate = this.parseDate(p.PRO_CREATED_ON);
          const minExp = minExpiryBySku.get(this.normalizeKey(sku)) || null;

          const analisisKey = minExp ? this.buildFechaSkuKey(minExp, this.normalizeKey(sku)) : '';
          const lookupDisponible = analisisKey ? analisisLookup.get(analisisKey) : undefined;

          // Cumplimiento FEFO:
          // SI.ERROR(SI(LPI_EXPIRY_DATE <= Fecha mas proxima;"Cumple";SI(lookup<300;"cumple";"No cumple"));"No cumple")
          let cumplimientoFefo = 'No cumple';
          if (expPedido && minExp) {
            if (expPedido.getTime() <= minExp.getTime()) {
              cumplimientoFefo = 'Cumple';
            } else if (lookupDisponible !== undefined && lookupDisponible < 300) {
              cumplimientoFefo = 'cumple';
            } else {
              cumplimientoFefo = 'No cumple';
            }
          }

          // Justificación FEFO:
          // SI(lookup<200;"saldo de lote";"")
          const justificacionFefo =
            lookupDisponible !== undefined && lookupDisponible < 200 ? 'saldo de lote' : '';

          const regla = ruleByCruce.get(this.normalizeKey(cruce));
          const tvuReglaN = this.parseNumber(regla?.TVU_REGLA);

          let fechaMinDespacho = 'No se encuentra la regla';
          let cumpleTvuValor = 'No se encuentra la regla';
          let comentariosMdlz = 'No se encuentra la regla';

          // Fecha Min Despacho:
          // Fecha de Despacho = PRO_CREATED_ON + 1 día
          // Fecha Min Despacho = Fecha de Despacho + TVU
          if (regla && tvuReglaN !== null && proCreatedOnDate) {
            const fechaDespacho = this.addDays(proCreatedOnDate, 1);
            const minDate = this.addDays(fechaDespacho, tvuReglaN);
            fechaMinDespacho = this.formatDate(minDate);

            if (expPedido) {
              const diff = this.dateDiffInDays(minDate, expPedido);
              cumpleTvuValor = String(diff);
              comentariosMdlz = diff >= 0 ? 'Cumple' : 'No cumple';
            } else {
              cumpleTvuValor = 'Sin fecha de vencimiento';
              comentariosMdlz = 'Sin fecha de vencimiento';
            }
          }

          const decisionFinal =
            comentariosMdlz === 'No cumple' ? 'Revisar la linea' : 'Cumple la linea';

          return {
            proCreatedOn:
              this.formatDate(this.parseDate(p.PRO_CREATED_ON)) || String(p.PRO_CREATED_ON ?? ''),
            oboOrderDate:
              this.formatDate(this.parseDate(p.OBO_ORDER_DATE)) || String(p.OBO_ORDER_DATE ?? ''),
            proId: String(p.PRO_ID ?? ''),
            oboId: String(p.OBO_ID ?? ''),
            oboSoldto: String(p.OBO_SOLDTO ?? ''),
            oboShiptoCode: String(p.OBO_SHIPTO_CODE ?? ''),
            oboShiptoCode2: String(p.OBO_SHIPTO_CODE2 ?? ''),
            oboShiptoName: String(p.OBO_SHIPTO_NAME ?? ''),
            oboTrackingNumber: String(p.OBO_TRACKING_NUMBER ?? ''),
            oboInvoice: String(p.OBO_INVOICE ?? ''),
            oboOrder: String(p.OBO_ORDER ?? ''),
            oboOrder2: String(p.OBO_ORDER2 ?? ''),
            oboPurchaseOrder: String(p.OBO_PURCHASE_ORDER ?? ''),
            sku,
            descripcion: String(p.MAI_DESCRIPTION ?? ''),
            lote: String(p.LPI_LOT ?? ''),
            fechaVencimientoPedido: this.formatDate(expPedido),
            fechaMasProximaVencer: this.formatDate(minExp),
            cumplimientoFefo,
            justificacionFefo,
            pkoEachQty: String(p.PKO_EACH_QTY ?? ''),
            pkoCaseCount: String(p.PKO_CASE_COUNT ?? ''),
            pkoModuleQty: String(p.PKO_MODULE_QTY ?? ''),
            pkoModuleCaseCount: String(p.PKO_MODULE_CASE_COUNT ?? ''),
            strategyName: strategy,
            diffDays: String(p.DIFF_DAYS ?? ''),
            cruceConcatena: cruce,
            tvuRegla: regla?.TVU_REGLA
              ? String(regla.TVU_REGLA).trim()
              : 'No se encuentra la regla',
            fechaMinDespacho,
            cumpleTvuValor,
            comentariosMdlz,
            decisionFinal,
            observacion: '',
          };
        });

        this.analysisResult.set(resultados);
        this.messageService.add({
          severity: 'success',
          summary: 'Reporte FEFO generado',
          detail: `${resultados.length} registros procesados correctamente`,
          life: 4000,
        });
      } catch (error) {
        this.messageService.add({
          severity: 'error',
          summary: 'Error de análisis',
          detail: (error as Error).message || 'No fue posible generar el reporte FEFO',
          life: 7000,
        });
      } finally {
        this.loading.set(false);
      }
    }, 100);
  }

  exportToExcel() {
    const data = this.analysisResult();
    if (!data.length) {
      this.messageService.add({
        severity: 'warn',
        summary: 'Sin datos',
        detail: 'No hay resultados para exportar',
        life: 3000,
      });
      return;
    }

    const totalLineas = data.length;
    const tvuCumple = data.filter(
      (r) => String(r.comentariosMdlz).trim().toLowerCase() === 'cumple',
    ).length;
    const tvuNoCumple = totalLineas - tvuCumple;

    const fefoCumple = data.filter(
      (r) => String(r.cumplimientoFefo).trim().toLowerCase() !== 'no cumple',
    ).length;
    const fefoNoCumple = totalLineas - fefoCumple;

    const totalCumple = data.filter(
      (r) => String(r.decisionFinal).trim().toLowerCase() === 'cumple la linea',
    ).length;
    const totalNoCumple = totalLineas - totalCumple;

    const formatPercent = (num: number, den: number): string => {
      if (!den) return '0,00%';
      const value = (num / den) * 100;
      return `${value.toFixed(2).replace('.', ',')}%`;
    };

    const fechaPicking = data[0]?.proCreatedOn || this.formatDate(new Date());
    const fechaDespacho = this.formatDate(
      this.addDays(this.parseDate(fechaPicking) || new Date(), 1),
    );

    const tableHeaders = [
      'PRO_CREATED_ON',
      'OBO_ORDER_DATE',
      'PRO_ID',
      'OBO_ID',
      'OBO_SOLDTO',
      'OBO_SHIPTO_CODE',
      'OBO_SHIPTO_CODE2',
      'OBO_SHIPTO_NAME',
      'OBO_TRACKING_NUMBER',
      'OBO_INVOICE',
      'OBO_ORDER',
      'OBO_ORDER2',
      'OBO_PURCHASE_ORDER',
      'MAI_SKU',
      'MAI_DESCRIPTION',
      'LPI_LOT',
      'LPI_EXPIRY_DATE',
      'Fecha mas proxima a vencer',
      'Cumplimiento de FEFO',
      'Justificación de FEFO',
      'PKO_EACH_QTY',
      'PKO_CASE_COUNT',
      'PKO_MODULE_QTY',
      'PKO_MODULE_CASE_COUNT',
      'STRATEGY_NAME',
      'DIFF_DAYS',
      'CRUCE CONACTENA',
      'TVU Regla',
      'Fecha Min Descpacho',
      'Cumple TVU',
      'Comentarios Mdlz',
      'Desición Final',
      'OBSERVACION',
    ];

    const aoa: any[][] = [];
    aoa.push(['Fecha de PICKING:', '', fechaPicking]);
    aoa.push([]);
    aoa.push([
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      'Reporte parametro TVU',
      '',
      '',
      '',
      'Reporte parametro FEFO',
      '',
      '',
      '',
      'Reporte Total',
    ]);
    aoa.push([
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      'Número de lineas totales',
      'Número y %',
      '',
      'Número de lineas totales',
      'Número y %',
      '',
      'Número de lineas totales',
      'Número y %',
    ]);
    aoa.push([
      'Fecha de Despacho:',
      '',
      fechaDespacho,
      '',
      '',
      '',
      '',
      'Número de lineas con cumplimiento',
      tvuCumple,
      '',
      'Número de lineas con cumplimiento',
      fefoCumple,
      '',
      'Número de lineas con cumplimiento',
      totalCumple,
    ]);
    aoa.push([
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      'Número de lineas sin cumplimiento',
      tvuNoCumple,
      '',
      'Número de lineas sin cumplimiento',
      fefoNoCumple,
      '',
      'Número de lineas sin cumplimiento',
      totalNoCumple,
    ]);
    aoa.push([
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '% de Acierto',
      formatPercent(tvuCumple, totalLineas),
      '',
      '% de Acierto',
      formatPercent(fefoCumple, totalLineas),
      '',
      '% de Acierto',
      formatPercent(totalCumple, totalLineas),
    ]);
    aoa.push([]);
    aoa.push([
      'Vendido',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      'Reglas del cliente',
      '',
      '',
      'Analisis',
    ]);
    aoa.push(tableHeaders);

    for (const r of data) {
      aoa.push([
        r.proCreatedOn,
        r.oboOrderDate,
        r.proId,
        r.oboId,
        r.oboSoldto,
        r.oboShiptoCode,
        r.oboShiptoCode2,
        r.oboShiptoName,
        r.oboTrackingNumber,
        r.oboInvoice,
        r.oboOrder,
        r.oboOrder2,
        r.oboPurchaseOrder,
        r.sku,
        r.descripcion,
        r.lote,
        r.fechaVencimientoPedido,
        r.fechaMasProximaVencer,
        r.cumplimientoFefo,
        r.justificacionFefo,
        r.pkoEachQty,
        r.pkoCaseCount,
        r.pkoModuleQty,
        r.pkoModuleCaseCount,
        r.strategyName,
        r.diffDays,
        r.cruceConcatena,
        r.tvuRegla,
        r.fechaMinDespacho,
        r.cumpleTvuValor,
        r.comentariosMdlz,
        r.decisionFinal,
        r.observacion,
      ]);
    }

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 1 } },
      { s: { r: 2, c: 7 }, e: { r: 2, c: 9 } },
      { s: { r: 2, c: 11 }, e: { r: 2, c: 13 } },
      { s: { r: 2, c: 15 }, e: { r: 2, c: 17 } },
      { s: { r: 4, c: 0 }, e: { r: 4, c: 1 } },
      { s: { r: 8, c: 0 }, e: { r: 8, c: 25 } },
      { s: { r: 8, c: 26 }, e: { r: 8, c: 28 } },
      { s: { r: 8, c: 29 }, e: { r: 8, c: 32 } },
    ];
    ws['!cols'] = tableHeaders.map((h) => ({ wch: Math.max(14, Math.min(40, h.length + 2)) }));
    XLSX.utils.book_append_sheet(wb, ws, 'FEFO Reporte');

    const fecha = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `Reporte_FEFO_${fecha}.xlsx`);

    this.messageService.add({
      severity: 'success',
      summary: 'Exportado',
      detail: 'Archivo FEFO exportado a Excel',
      life: 3000,
    });
  }
}
