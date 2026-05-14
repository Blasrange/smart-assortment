import { Component, Inject, PLATFORM_ID, ViewChild, signal } from '@angular/core';
import { CommonModule, isPlatformBrowser } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { Table, TableModule } from 'primeng/table';
import { ButtonModule } from 'primeng/button';
import { InputTextModule } from 'primeng/inputtext';
import { BreadcrumbModule } from 'primeng/breadcrumb';
import { ProgressSpinnerModule } from 'primeng/progressspinner';
import { ToastModule } from 'primeng/toast';
import { TagModule } from 'primeng/tag';
import { TooltipModule } from 'primeng/tooltip';
import { MessageService, MenuItem } from 'primeng/api';
import { PrimeNG } from 'primeng/config';
import * as XLSX from 'xlsx';

// ============================================================
// MAPEO DE COLUMNAS
// ============================================================

// Mapeo para inventario Cliente (SAP)
const CLIENT_INVENTORY_MAPPING = {
  material: ['Material', 'Material', 'SKU', 'Codigo'],
  description: ['Material Description', 'Descripción', 'Descripcion', 'Description'],
  saldo: ['saldo', 'Saldo', 'Saldo disponible', 'Unrestricted', 'Disponible'],
  estado: ['estado', 'Estado', 'ESTADO'],
};

// Mapeo para inventario WMS
const WMS_INVENTORY_MAPPING = {
  sku: ['SKU', 'CODIGO', 'Material', 'Codigo'],
  description: ['DESCRIPCION', 'Descripcion', 'Description'],
  status: ['ESTADO', 'Estado', 'Situación'],
  boxes: ['CAJAS', 'Cajas', 'Boxes'],
  units: ['UNIDADES', 'Unidades', 'Cantidad'],
};

// Mapeo para Maestro de Materiales (Familia)
const FAMILY_MAPPING = {
  material: ['Material', 'SKU', 'Codigo', 'Material'],
  description: ['Material Description', 'Descripción', 'Descripcion'],
  family: ['Familia', 'Family'],
  subfamily: ['Subfamilia', 'Subfamily'],
  caseConfig: ['cases configuration', 'Case Config', 'Case Configuration'],
  gtinCases: ['GTIN CASES', 'GTIN Cases', 'Cases GTIN'],
  gtinInner: ['GTIN INNER PACK', 'GTIN Inner', 'Inner GTIN'],
  gtinEach: ['GTIN EACH', 'GTIN Each', 'Each GTIN'],
  unitValue: ['Valor por caja', 'Valor Unitario', 'Unit Value'],
};

// ============================================================
// MAPEO DE ESTADOS DE CALIDAD
// ============================================================

// Mapeo de estados SAP (texto a categoría)
const SAP_STATUS_MAPPING: Record<string, string> = {
  disponible: 'disponible',
  dsp: 'disponible',
  'dsp (disponible)': 'disponible',
  recuperable: 'recuperable',
  pr: 'recuperable',
  averiado: 'averiado',
  'av (averiado) /dstr - prod dest': 'averiado',
  av: 'averiado',
  dp: 'averiado',
  vencido: 'vencido',
  oa: 'vencido',
  maquilas: 'maquilas',
  ss: 'maquilas',
  cuarentena: 'recuperable',
  qa: 'recuperable',
};

// Mapeo de estados WMS (código a categoría) - Configurable por el usuario
const DEFAULT_WMS_STATUS_MAPPINGS = [
  { code: 'DSP', category: 'disponible' },
  { code: 'DSP  (Disponible)', category: 'disponible' },
  { code: 'Disponible', category: 'disponible' },
  { code: 'REC', category: 'recuperable' },
  { code: 'Recuperable', category: 'recuperable' },
  { code: 'AVR', category: 'averiado' },
  { code: 'Averiado', category: 'averiado' },
  { code: 'AV', category: 'averiado' },
  { code: 'VNC', category: 'vencido' },
  { code: 'Vencido', category: 'vencido' },
  { code: 'MAQ', category: 'maquilas' },
  { code: 'Maquilas', category: 'maquilas' },
  { code: 'QA', category: 'recuperable' },
  { code: 'PR', category: 'recuperable' },
  { code: 'OA', category: 'vencido' },
  { code: 'DP', category: 'averiado' },
  { code: '1', category: 'disponible' },
  { code: 'SS', category: 'maquilas' },
];

// ============================================================
// INTERFACES
// ============================================================

interface WmsStatusMapping {
  code: string;
  category: string;
}

interface ReconciliationRow {
  material: string;
  materialDescription: string;
  family: string;
  subfamily: string;
  caseConfig: string;
  // Disponible
  disponibleSAP: number;
  disponibleWMS: number;
  disponibleDif: number;
  // Recuperable
  recuperableSAP: number;
  recuperableWMS: number;
  recuperableDif: number;
  // Averiado
  averiadoSAP: number;
  averiadoWMS: number;
  averiadoDif: number;
  // Vencido
  vencidoSAP: number;
  vencidoWMS: number;
  vencidoDif: number;
  // Maquilas
  maquilasSAP: number;
  maquilasWMS: number;
  maquilasDif: number;
  // Totales
  totalSAP: number;
  totalWMS: number;
  diferencia: number;
  faltantesSAP: number;
  sobrantesWMS: number;
  valorUnitario: number;
  valorTotal: number;
  gtinCases: string;
  gtinInner: string;
  gtinEach: string;
}

interface ReconciliationResult {
  summary: string;
  results: ReconciliationRow[];
  stats: {
    totalMateriales: number;
    coincidencias: number;
    conDiferencias: number;
    porcentajeCoincidencia: number;
  };
}

interface FileData {
  name: string;
  data: string | ArrayBuffer | null;
  size: number;
  uploadDate: Date;
}

@Component({
  selector: 'app-inventory-reconciliation',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    TableModule,
    ButtonModule,
    InputTextModule,
    BreadcrumbModule,
    ProgressSpinnerModule,
    ToastModule,
    TagModule,
    TooltipModule,
  ],
  providers: [MessageService],
  templateUrl: './inventory-reconciliation.html',
  styleUrl: './inventory-reconciliation.css',
})
export class InventoryReconciliation {
  @ViewChild('dt') dt!: Table;

  // Archivos
  clientFile = signal<FileData | null>(null);
  wmsFile = signal<FileData | null>(null);
  familyFile = signal<FileData | null>(null);

  // Resultados y estado
  analysisResult = signal<ReconciliationResult | null>(null);
  loading = signal(false);
  searchValue = '';
  isBrowser: boolean;
  showStatusMapping = false;
  editingStatus = false;

  // Mapeo de headers
  clientHeaderMap: Record<string, string> = {};
  wmsHeaderMap: Record<string, string> = {};
  familyHeaderMap: Record<string, string> = {};

  // Drag state
  isDraggingClient = signal(false);
  isDraggingWms = signal(false);
  isDraggingFamily = signal(false);

  // Configuración de mapeo de estados WMS
  wmsStatusMappings: WmsStatusMapping[] = [...DEFAULT_WMS_STATUS_MAPPINGS];
  originalWmsMappings: WmsStatusMapping[] = [...DEFAULT_WMS_STATUS_MAPPINGS];

  breadcrumbItems: MenuItem[] = [
    { label: 'Conciliación de Inventarios', routerLink: '/inventory-reconciliation' },
  ];
  breadcrumbHome: MenuItem = { icon: 'pi pi-home', label: 'Inicio', routerLink: '/' };

  // Filtro global
  get filteredResults(): ReconciliationRow[] {
    const result = this.analysisResult();
    if (!result) return [];
    const q = this.searchValue.toLowerCase().trim();
    if (!q) return result.results;
    return result.results.filter(
      (row) =>
        row.material.toLowerCase().includes(q) || row.materialDescription.toLowerCase().includes(q),
    );
  }

  constructor(
    @Inject(PLATFORM_ID) platformId: Object,
    private messageService: MessageService,
    private primeng: PrimeNG,
  ) {
    this.isBrowser = isPlatformBrowser(platformId);
  }

  ngOnInit() {
    // Traducción de filtros PrimeNG
    if (this.primeng?.setTranslation) {
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
    this.loadWmsStatusMapping();
  }

  // ============================================================
  // GESTIÓN DE MAPEO DE ESTADOS WMS
  // ============================================================

  private loadWmsStatusMapping() {
    const saved = localStorage.getItem('wms_status_mappings');
    if (saved) {
      try {
        this.wmsStatusMappings = JSON.parse(saved);
        this.originalWmsMappings = JSON.parse(saved);
      } catch (e) {
        console.error(e);
      }
    }
  }

  private saveWmsStatusMapping() {
    localStorage.setItem('wms_status_mappings', JSON.stringify(this.wmsStatusMappings));
  }

  toggleStatusMapping() {
    this.showStatusMapping = !this.showStatusMapping;
  }
  toggleStatusEdit() {
    this.editingStatus = true;
  }

  cancelStatusEdit() {
    this.wmsStatusMappings = JSON.parse(JSON.stringify(this.originalWmsMappings));
    this.editingStatus = false;
  }

  saveStatusMapping() {
    this.originalWmsMappings = JSON.parse(JSON.stringify(this.wmsStatusMappings));
    this.saveWmsStatusMapping();
    this.editingStatus = false;
    this.messageService.add({
      severity: 'success',
      summary: 'Mapeo guardado',
      detail: 'Configuración de estados actualizada',
      life: 3000,
    });
  }

  // ============================================================
  // FUNCIONES DE MAPEO DE ESTADOS
  // ============================================================

  getSapStatusCategory(estadoText: string): string {
    const normalized = String(estadoText || '')
      .toLowerCase()
      .trim();
    for (const [key, category] of Object.entries(SAP_STATUS_MAPPING)) {
      if (normalized === key || normalized.includes(key)) {
        return category;
      }
    }
    return 'disponible';
  }

  getWmsStatusCategory(statusCode: string): string {
    const upper = String(statusCode || '')
      .toUpperCase()
      .trim();
    for (const mapping of this.wmsStatusMappings) {
      if (upper === mapping.code.toUpperCase() || upper.includes(mapping.code.toUpperCase())) {
        return mapping.category;
      }
    }
    return 'disponible';
  }

  // ============================================================
  // MANEJO DE ARCHIVOS (DRAG & DROP / UPLOAD)
  // ============================================================

  onDragOver(event: DragEvent, type: 'client' | 'wms' | 'family') {
    event.preventDefault();
    if (type === 'client') this.isDraggingClient.set(true);
    else if (type === 'wms') this.isDraggingWms.set(true);
    else this.isDraggingFamily.set(true);
  }

  onDragLeave(event: DragEvent, type: 'client' | 'wms' | 'family') {
    event.preventDefault();
    if (type === 'client') this.isDraggingClient.set(false);
    else if (type === 'wms') this.isDraggingWms.set(false);
    else this.isDraggingFamily.set(false);
  }

  onDrop(event: DragEvent, type: 'client' | 'wms' | 'family') {
    event.preventDefault();
    if (type === 'client') this.isDraggingClient.set(false);
    else if (type === 'wms') this.isDraggingWms.set(false);
    else this.isDraggingFamily.set(false);
    const files = event.dataTransfer?.files;
    if (files?.length) this.processFile(files[0], type);
  }

  onFileUpload(type: 'client' | 'wms' | 'family', event: Event) {
    const input = event.target as HTMLInputElement;
    if (input.files?.length) this.processFile(input.files[0], type);
  }

  private processFile(file: File, type: 'client' | 'wms' | 'family') {
    const ext = file.name.split('.').pop()?.toLowerCase();
    if (!['csv', 'xlsx', 'xls', 'txt'].includes(ext || '')) {
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: 'Formato no soportado. Use CSV, Excel o TXT',
        life: 5000,
      });
      return;
    }

    this.analysisResult.set(null);
    const reader = new FileReader();

    reader.onload = () => {
      const data = reader.result;
      const rows = this.parseFileData(data, file.name);
      const headers = rows.length > 0 ? Object.keys(rows[0]) : [];
      let mapping, headerMap;

      if (type === 'client') {
        mapping = CLIENT_INVENTORY_MAPPING;
        headerMap = this.mapHeaders(headers, mapping);
        this.clientFile.set({ name: file.name, data, size: file.size, uploadDate: new Date() });
        this.clientHeaderMap = headerMap;
      } else if (type === 'wms') {
        mapping = WMS_INVENTORY_MAPPING;
        headerMap = this.mapHeaders(headers, mapping);
        this.wmsFile.set({ name: file.name, data, size: file.size, uploadDate: new Date() });
        this.wmsHeaderMap = headerMap;
      } else {
        mapping = FAMILY_MAPPING;
        headerMap = this.mapHeaders(headers, mapping);
        this.familyFile.set({ name: file.name, data, size: file.size, uploadDate: new Date() });
        this.familyHeaderMap = headerMap;
      }

      this.messageService.add({
        severity: 'success',
        summary: 'Archivo cargado',
        detail: `${rows.length} registros`,
        life: 3000,
      });
    };

    reader.onerror = () =>
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: 'No se pudo leer el archivo',
        life: 5000,
      });

    if (ext === 'xlsx' || ext === 'xls') reader.readAsBinaryString(file);
    else reader.readAsText(file, 'UTF-8');
  }

  removeFile(type: 'client' | 'wms' | 'family') {
    if (type === 'client') {
      this.clientFile.set(null);
      this.clientHeaderMap = {};
    } else if (type === 'wms') {
      this.wmsFile.set(null);
      this.wmsHeaderMap = {};
    } else {
      this.familyFile.set(null);
      this.familyHeaderMap = {};
    }
    this.analysisResult.set(null);
    this.messageService.add({
      severity: 'info',
      summary: 'Archivo removido',
      detail: `Archivo eliminado`,
      life: 3000,
    });
  }

  // ============================================================
  // CONCILIACIÓN PRINCIPAL
  // ============================================================

  generateReconciliation() {
    if (!this.clientFile() || !this.wmsFile()) {
      this.messageService.add({
        severity: 'warn',
        summary: 'Archivos faltantes',
        detail: 'Debes subir archivos Cliente y WMS',
        life: 5000,
      });
      return;
    }

    this.loading.set(true);
    setTimeout(() => {
      try {
        const result = this.processReconciliation();
        this.analysisResult.set(result);
        this.messageService.add({
          severity: 'success',
          summary: 'Conciliación completada',
          detail: result.summary,
          life: 5000,
        });
      } catch (err) {
        this.messageService.add({
          severity: 'error',
          summary: 'Error',
          detail: (err as Error).message,
          life: 7000,
        });
      } finally {
        this.loading.set(false);
      }
    }, 100);
  }

  private processReconciliation(): ReconciliationResult {
    // Parsear archivos
    const clientRows = this.parseFileData(this.clientFile()!.data, this.clientFile()!.name);
    const wmsRows = this.parseFileData(this.wmsFile()!.data, this.wmsFile()!.name);
    const familyRows = this.familyFile()
      ? this.parseFileData(this.familyFile()!.data, this.familyFile()!.name)
      : [];

    // ============================================================
    // 1. MAPA DE FAMILIAS (Maestro de Materiales)
    // ============================================================
    const familyMap: Record<string, any> = {};
    for (const row of familyRows) {
      const sku = this.getVal(row, this.familyHeaderMap['material'] || 'Material', '');
      if (!sku) continue;
      const key = String(sku).trim().toUpperCase();
      familyMap[key] = {
        family: this.getVal(row, this.familyHeaderMap['family'] || 'Familia', ''),
        subfamily: this.getVal(row, this.familyHeaderMap['subfamily'] || 'Subfamilia', ''),
        caseConfig: this.getVal(
          row,
          this.familyHeaderMap['caseConfig'] || 'cases configuration',
          '',
        ),
        gtinCases: this.getVal(row, this.familyHeaderMap['gtinCases'] || 'GTIN CASES', ''),
        gtinInner: this.getVal(row, this.familyHeaderMap['gtinInner'] || 'GTIN INNER PACK', ''),
        gtinEach: this.getVal(row, this.familyHeaderMap['gtinEach'] || 'GTIN EACH', ''),
        unitValue: this.toNumber(
          this.getVal(row, this.familyHeaderMap['unitValue'] || 'Vr Unit', 0),
        ),
      };
    }

    // ============================================================
    // 2. AGRUPAR INVENTARIO CLIENTE (SAP) por SKU y estado
    //    Usa los campos: material, saldo, estado
    // ============================================================
    const clientAgg: Record<
      string,
      {
        description: string;
        disponible: number;
        recuperable: number;
        averiado: number;
        vencido: number;
        maquilas: number;
      }
    > = {};

    for (const row of clientRows) {
      // Obtener SKU
      const sku = this.getVal(row, this.clientHeaderMap['material'] || 'Material', '');
      if (!sku) continue;
      const key = String(sku).trim().toUpperCase();

      // Obtener estado y cantidad (saldo)
      const estadoText = this.getVal(row, this.clientHeaderMap['estado'] || 'estado', '');
      const saldo = this.toNumber(this.getVal(row, this.clientHeaderMap['saldo'] || 'saldo', 0));

      if (saldo === 0) continue;

      // Clasificar según el estado
      const statusCategory = this.getSapStatusCategory(estadoText);

      if (!clientAgg[key]) {
        clientAgg[key] = {
          description: this.getVal(
            row,
            this.clientHeaderMap['description'] || 'Material Description',
            '',
          ),
          disponible: 0,
          recuperable: 0,
          averiado: 0,
          vencido: 0,
          maquilas: 0,
        };
      }

      // Acumular en la categoría correspondiente
      switch (statusCategory) {
        case 'disponible':
          clientAgg[key].disponible += saldo;
          break;
        case 'recuperable':
          clientAgg[key].recuperable += saldo;
          break;
        case 'averiado':
          clientAgg[key].averiado += saldo;
          break;
        case 'vencido':
          clientAgg[key].vencido += saldo;
          break;
        case 'maquilas':
          clientAgg[key].maquilas += saldo;
          break;
        default:
          clientAgg[key].disponible += saldo;
      }
    }

    // ============================================================
    // 3. AGRUPAR INVENTARIO WMS por SKU y estado
    //    Usa el campo CAJAS como unidad principal
    // ============================================================
    const wmsAgg: Record<
      string,
      {
        disponible: number;
        recuperable: number;
        averiado: number;
        vencido: number;
        maquilas: number;
      }
    > = {};

    for (const row of wmsRows) {
      const sku = this.getVal(row, this.wmsHeaderMap['sku'] || 'SKU', '');
      if (!sku) continue;
      const key = String(sku).trim().toUpperCase();

      const statusRaw = String(this.getVal(row, this.wmsHeaderMap['status'] || 'ESTADO', ''));
      const statusCategory = this.getWmsStatusCategory(statusRaw);

      // Obtener cantidad (priorizar CAJAS, luego UNIDADES)
      let qty = this.toNumber(this.getVal(row, this.wmsHeaderMap['boxes'] || 'CAJAS', 0));
      if (qty === 0) {
        qty = this.toNumber(this.getVal(row, this.wmsHeaderMap['units'] || 'UNIDADES', 0));
      }
      if (qty === 0) continue;

      if (!wmsAgg[key]) {
        wmsAgg[key] = { disponible: 0, recuperable: 0, averiado: 0, vencido: 0, maquilas: 0 };
      }

      switch (statusCategory) {
        case 'disponible':
          wmsAgg[key].disponible += qty;
          break;
        case 'recuperable':
          wmsAgg[key].recuperable += qty;
          break;
        case 'averiado':
          wmsAgg[key].averiado += qty;
          break;
        case 'vencido':
          wmsAgg[key].vencido += qty;
          break;
        case 'maquilas':
          wmsAgg[key].maquilas += qty;
          break;
        default:
          wmsAgg[key].disponible += qty;
      }
    }

    // ============================================================
    // 4. COMBINAR RESULTADOS por SKU
    // ============================================================
    const allKeys = new Set([...Object.keys(clientAgg), ...Object.keys(wmsAgg)]);
    const results: ReconciliationRow[] = [];

    for (const key of allKeys) {
      const client = clientAgg[key] || {
        description: '',
        disponible: 0,
        recuperable: 0,
        averiado: 0,
        vencido: 0,
        maquilas: 0,
      };
      const wms = wmsAgg[key] || {
        disponible: 0,
        recuperable: 0,
        averiado: 0,
        vencido: 0,
        maquilas: 0,
      };
      const family = familyMap[key] || {};

      const totalSAP =
        client.disponible + client.recuperable + client.averiado + client.vencido + client.maquilas;
      const totalWMS = wms.disponible + wms.recuperable + wms.averiado + wms.vencido + wms.maquilas;
      const diferencia = totalWMS - totalSAP;

      const valorUnitario = family.unitValue || 0;
      const valorTotal = Math.abs(diferencia) * valorUnitario;

      results.push({
        material: key,
        materialDescription: client.description,
        family: family.family || '',
        subfamily: family.subfamily || '',
        caseConfig: family.caseConfig || '',
        disponibleSAP: client.disponible,
        disponibleWMS: wms.disponible,
        disponibleDif: wms.disponible - client.disponible,
        recuperableSAP: client.recuperable,
        recuperableWMS: wms.recuperable,
        recuperableDif: wms.recuperable - client.recuperable,
        averiadoSAP: client.averiado,
        averiadoWMS: wms.averiado,
        averiadoDif: wms.averiado - client.averiado,
        vencidoSAP: client.vencido,
        vencidoWMS: wms.vencido,
        vencidoDif: wms.vencido - client.vencido,
        maquilasSAP: client.maquilas,
        maquilasWMS: wms.maquilas,
        maquilasDif: wms.maquilas - client.maquilas,
        totalSAP,
        totalWMS,
        diferencia,
        faltantesSAP: diferencia < 0 ? Math.abs(diferencia) : 0,
        sobrantesWMS: diferencia > 0 ? diferencia : 0,
        valorUnitario,
        valorTotal,
        gtinCases: family.gtinCases || '',
        gtinInner: family.gtinInner || '',
        gtinEach: family.gtinEach || '',
      });
    }

    // Ordenar por SKU
    results.sort((a, b) => a.material.localeCompare(b.material));

    const totalMateriales = results.length;
    const conDiferencias = results.filter((r) => Math.abs(r.diferencia) > 0.001).length;
    const coincidencias = totalMateriales - conDiferencias;
    const porcentajeCoincidencia =
      totalMateriales === 0 ? 100 : Math.round((coincidencias / totalMateriales) * 100);

    return {
      summary: `Conciliación completada: ${totalMateriales} materiales, ${conDiferencias} con diferencias`,
      results,
      stats: { totalMateriales, coincidencias, conDiferencias, porcentajeCoincidencia },
    };
  }

  // ============================================================
  // MÉTRICAS Y UTILIDADES
  // ============================================================

  getTotalMateriales(): number {
    return this.analysisResult()?.stats.totalMateriales || 0;
  }
  getCoincidencias(): number {
    return this.analysisResult()?.stats.coincidencias || 0;
  }
  getConDiferencias(): number {
    return this.analysisResult()?.stats.conDiferencias || 0;
  }
  getPorcentajeCoincidencia(): number {
    return this.analysisResult()?.stats.porcentajeCoincidencia || 0;
  }

  applyGlobalFilter() {
    if (this.dt) this.dt.filterGlobal(this.searchValue, 'contains');
  }
  clearGlobalFilter() {
    this.searchValue = '';
    if (this.dt) this.dt.filterGlobal('', 'contains');
  }

  clearAllFilters() {
    this.searchValue = '';
    if (this.dt) {
      this.dt.clear();
      this.dt.filterGlobal('', 'contains');
    }
    this.messageService.add({
      severity: 'info',
      summary: 'Filtros limpiados',
      detail: 'Se eliminaron todos los filtros',
      life: 3000,
    });
  }

  formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    const k = 1024,
      sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  // ============================================================
  // EXPORTAR A EXCEL
  // ============================================================

  exportToExcel() {
    const result = this.analysisResult();
    if (!result) return;

    const toGtinNumber = (value: any): number | '' => {
      const digits = String(value ?? '').replace(/\D/g, '');
      if (!digits) return '';
      const num = Number(digits);
      return Number.isFinite(num) ? num : '';
    };

    // Fila 1: encabezados agrupados (como en la tabla)
    const headerGroupRow = [
      'Material',
      'Descripcion',
      'Familia',
      'Subfamilia',
      'Case Config',
      'DISPONIBLE',
      '',
      '',
      'RECUPERABLE',
      '',
      '',
      'AVERIADO',
      '',
      '',
      'VENCIDO',
      '',
      '',
      'MAQUILAS',
      '',
      '',
      'Diferencia Total',
      'Valor Unitario',
      'Valor Total',
      'GTIN Cases',
      'GTIN Inner',
      'GTIN Each',
    ];

    // Fila 2: sub-encabezados
    const headerSubRow = [
      '',
      '',
      '',
      '',
      '',
      'SAP',
      'WMS',
      'DIF',
      'SAP',
      'WMS',
      'DIF',
      'SAP',
      'WMS',
      'DIF',
      'SAP',
      'WMS',
      'DIF',
      'SAP',
      'WMS',
      'DIF',
      '',
      '',
      '',
      '',
      '',
      '',
    ];

    const dataRows = result.results.map((r) => [
      r.material,
      r.materialDescription,
      r.family,
      r.subfamily,
      r.caseConfig,
      r.disponibleSAP,
      r.disponibleWMS,
      r.disponibleDif,
      r.recuperableSAP,
      r.recuperableWMS,
      r.recuperableDif,
      r.averiadoSAP,
      r.averiadoWMS,
      r.averiadoDif,
      r.vencidoSAP,
      r.vencidoWMS,
      r.vencidoDif,
      r.maquilasSAP,
      r.maquilasWMS,
      r.maquilasDif,
      r.diferencia,
      r.valorUnitario,
      r.valorTotal,
      toGtinNumber(r.gtinCases),
      toGtinNumber(r.gtinInner),
      toGtinNumber(r.gtinEach),
    ]);

    const ws = XLSX.utils.aoa_to_sheet([headerGroupRow, headerSubRow, ...dataRows]);

    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } },
      { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } },
      { s: { r: 0, c: 2 }, e: { r: 1, c: 2 } },
      { s: { r: 0, c: 3 }, e: { r: 1, c: 3 } },
      { s: { r: 0, c: 4 }, e: { r: 1, c: 4 } },
      { s: { r: 0, c: 5 }, e: { r: 0, c: 7 } },
      { s: { r: 0, c: 8 }, e: { r: 0, c: 10 } },
      { s: { r: 0, c: 11 }, e: { r: 0, c: 13 } },
      { s: { r: 0, c: 14 }, e: { r: 0, c: 16 } },
      { s: { r: 0, c: 17 }, e: { r: 0, c: 19 } },
      { s: { r: 0, c: 20 }, e: { r: 1, c: 20 } },
      { s: { r: 0, c: 21 }, e: { r: 1, c: 21 } },
      { s: { r: 0, c: 22 }, e: { r: 1, c: 22 } },
      { s: { r: 0, c: 23 }, e: { r: 1, c: 23 } },
      { s: { r: 0, c: 24 }, e: { r: 1, c: 24 } },
      { s: { r: 0, c: 25 }, e: { r: 1, c: 25 } },
    ];

    ws['!cols'] = [
      { wch: 18 },
      { wch: 40 },
      { wch: 16 },
      { wch: 16 },
      { wch: 16 },
      { wch: 12 },
      { wch: 12 },
      { wch: 12 },
      { wch: 12 },
      { wch: 12 },
      { wch: 12 },
      { wch: 12 },
      { wch: 12 },
      { wch: 12 },
      { wch: 12 },
      { wch: 12 },
      { wch: 12 },
      { wch: 12 },
      { wch: 12 },
      { wch: 12 },
      { wch: 16 },
      { wch: 16 },
      { wch: 16 },
      { wch: 16 },
      { wch: 16 },
      { wch: 16 },
    ];

    // Formatos: moneda en valores y numero entero en EAN/GTIN
    const firstDataRow = 2;
    const lastDataRow = firstDataRow + dataRows.length - 1;
    for (let r = firstDataRow; r <= lastDataRow; r += 1) {
      const unitCellRef = XLSX.utils.encode_cell({ r, c: 21 });
      const totalCellRef = XLSX.utils.encode_cell({ r, c: 22 });
      if (ws[unitCellRef]) ws[unitCellRef].z = '$#,##0.00';
      if (ws[totalCellRef]) ws[totalCellRef].z = '$#,##0.00';

      for (const c of [23, 24, 25]) {
        const eanRef = XLSX.utils.encode_cell({ r, c });
        if (ws[eanRef]) ws[eanRef].z = '0';
      }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Conciliación');
    XLSX.writeFile(wb, `conciliacion_inventarios_${new Date().toISOString().slice(0, 10)}.xlsx`);
    this.messageService.add({
      severity: 'success',
      summary: 'Exportado',
      detail: 'Archivo Excel generado',
      life: 3000,
    });
  }

  // ============================================================
  // FUNCIONES AUXILIARES
  // ============================================================

  private mapHeaders(headers: string[], mapping: Record<string, string[]>): Record<string, string> {
    const result: Record<string, string> = {};
    const headersLower = headers.map((h) => h.toLowerCase().trim());
    for (const key in mapping) {
      for (const alias of mapping[key]) {
        const idx = headersLower.indexOf(alias.toLowerCase().trim());
        if (idx !== -1) {
          result[key] = headers[idx];
          break;
        }
      }
    }
    return result;
  }

  private parseFileData(data: any, fileName: string): any[] {
    const ext = fileName.split('.').pop()?.toLowerCase();
    if (ext === 'xlsx' || ext === 'xls') {
      const workbook = XLSX.read(data, { type: 'binary' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      return XLSX.utils.sheet_to_json(sheet);
    }
    const text = String(data);
    const lines = text.split(/\r?\n/).filter((l) => l.trim());
    if (lines.length < 2) return [];
    let delimiter = lines[0].includes(';') ? ';' : ',';
    const headers = lines[0].split(delimiter).map((h) => h.trim().replace(/^['"]|['"]$/g, ''));
    return lines.slice(1).map((line) => {
      const vals = line.split(delimiter);
      const obj: any = {};
      headers.forEach((h, i) => (obj[h] = vals[i]?.trim().replace(/^['"]|['"]$/g, '') || ''));
      return obj;
    });
  }

  private getVal(obj: any, path: string, def: any): any {
    if (!obj || !path) return def;
    const val = obj[path];
    return val !== undefined && val !== null && val !== '' ? val : def;
  }

  private toNumber(val: any): number {
    if (typeof val === 'number') return isFinite(val) ? val : 0;
    const cleaned = String(val ?? '')
      .replace(/\./g, '')
      .replace(',', '.')
      .replace(/[^\d.-]/g, '');
    const num = parseFloat(cleaned);
    return isFinite(num) ? num : 0;
  }
}
