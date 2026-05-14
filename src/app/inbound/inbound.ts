import { Component, Inject, PLATFORM_ID, ViewChild, signal } from '@angular/core';
import { CommonModule, isPlatformBrowser } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { Table, TableModule } from 'primeng/table';
import { ButtonModule } from 'primeng/button';
import { InputTextModule } from 'primeng/inputtext';
import { ProgressSpinnerModule } from 'primeng/progressspinner';
import { TooltipModule } from 'primeng/tooltip';
import { BreadcrumbModule } from 'primeng/breadcrumb';
import { ToastModule } from 'primeng/toast';
import { MessageService, MenuItem } from 'primeng/api';
import { PrimeNG } from 'primeng/config';
import * as XLSX from 'xlsx';

interface FileData {
  name: string;
  data: string | ArrayBuffer | null;
  size: number;
  uploadDate: Date;
}

interface InboundEntry {
  N_ORDER: string;
  ORDER2: string;
  PURCHASE_ORDER: string;
  INVOICE: string;
  PROVIDER_UID: string;
  ORDER_DATE: string;
  SERVICE_DATE: string;
  INBOUNDTYPE_CODE: string;
  NOTE: string;
  SKU: string;
  LOTE: string;
  FECHA_DE_VENCIMIENTO: string;
  FECHA_DE_FABRICACION: string;
  SERIAL: string;
  ESTADO_CALIDAD: string;
  QTY: number;
  UOM_CODE: string;
  REFERENCE: string;
  PRICE: number;
  TAXES: number;
  IBL_LPN_CODE: string;
  IBL_WEIGHT: number;
}

type InboundFieldType = 'text' | 'number' | 'date';

interface InboundField {
  key: keyof InboundEntry;
  label: string;
  type: InboundFieldType;
}

interface InboundResult {
  results: InboundEntry[];
  summary: string;
  stats: {
    sourceRows: number;
    processedRows: number;
    groupedRows: number;
  };
}

const INBOUND_FIELDS: InboundField[] = [
  { key: 'N_ORDER', label: 'N_ORDER', type: 'text' },
  { key: 'ORDER2', label: 'ORDER2', type: 'text' },
  { key: 'PURCHASE_ORDER', label: 'PURCHASE_ORDER', type: 'text' },
  { key: 'INVOICE', label: 'INVOICE', type: 'text' },
  { key: 'PROVIDER_UID', label: 'PROVIDER_UID', type: 'text' },
  { key: 'ORDER_DATE', label: 'ORDER_DATE', type: 'date' },
  { key: 'SERVICE_DATE', label: 'SERVICE_DATE', type: 'date' },
  { key: 'INBOUNDTYPE_CODE', label: 'INBOUNDTYPE_CODE', type: 'text' },
  { key: 'NOTE', label: 'NOTE', type: 'text' },
  { key: 'SKU', label: 'SKU', type: 'text' },
  { key: 'LOTE', label: 'LOTE', type: 'text' },
  { key: 'FECHA_DE_VENCIMIENTO', label: 'FECHA DE VENCIMIENTO', type: 'date' },
  { key: 'FECHA_DE_FABRICACION', label: 'FECHA DE FABRICACION', type: 'date' },
  { key: 'SERIAL', label: 'SERIAL', type: 'text' },
  { key: 'ESTADO_CALIDAD', label: 'ESTADO CALIDAD', type: 'text' },
  { key: 'QTY', label: 'QTY', type: 'number' },
  { key: 'UOM_CODE', label: 'UOM_CODE', type: 'text' },
  { key: 'REFERENCE', label: 'REFERENCE', type: 'text' },
  { key: 'PRICE', label: 'PRICE', type: 'number' },
  { key: 'TAXES', label: 'TAXES', type: 'number' },
  { key: 'IBL_LPN_CODE', label: 'IBL_LPN_CODE', type: 'text' },
  { key: 'IBL_WEIGHT', label: 'IBL_WEIGHT', type: 'number' },
];

const REQUIRED_KEYS: Array<keyof InboundEntry> = ['SKU', 'LOTE', 'QTY'];
const EXACT_MATCH_ONLY_FIELDS = new Set<keyof InboundEntry>([
  'UOM_CODE',
  'PROVIDER_UID',
  'ESTADO_CALIDAD',
]);

const BASE_FIELD_ALIASES: Record<string, string[]> = {
  N_ORDER: [
    'N_ORDER',
    'N ORDER',
    'NRO ORDEN',
    'NUMERO ORDEN',
    'ALBARAN',
    'ALBARAN DE ENTRADA',
    'Numero de orden de transporte',
    'Número de orden de transporte',
  ],
  ORDER2: ['ORDER2', 'ORDEN2', 'Numero de orden de transporte', 'Número de orden de transporte'],
  PURCHASE_ORDER: [
    'PURCHASE_ORDER',
    'PURCHASE ORDER',
    'OC',
    'ORDEN DE COMPRA',
    'PEDIDO',
    'Ubicacion de destino',
    'Ubicación de destino',
  ],
  INVOICE: ['INVOICE', 'FACTURA'],
  PROVIDER_UID: ['PROVIDER_UID', 'PROVEEDOR', 'PROVIDER', 'CODIGO PROVEEDOR', 'ID PROVEEDOR'],
  ORDER_DATE: ['ORDER_DATE', 'ORDER DATE', 'FECHA ORDEN', 'FECHA PEDIDO'],
  SERVICE_DATE: ['SERVICE_DATE', 'SERVICE DATE', 'FECHA SERVICIO'],
  INBOUNDTYPE_CODE: ['INBOUNDTYPE_CODE', 'TIPO ENTRADA', 'TIPO_INGRESO'],
  NOTE: ['NOTE', 'NOTA', 'OBSERVACION', 'OBSERVACIONES'],
  SKU: ['SKU', 'MATERIAL', 'SKU 2', 'CODIGO', 'CODIGO ARTICULO'],
  LOTE: ['LOTE', 'CE. LOTE', 'BATCH'],
  FECHA_DE_VENCIMIENTO: [
    'FECHA DE VENCIMIENTO',
    'FECHA_VENCIMIENTO',
    'VENCIMIENTO',
    'SLED/BBD',
    'F. VENC.',
    'FeCaduc/FePreferCons',
  ],
  FECHA_DE_FABRICACION: ['FECHA DE FABRICACION', 'FECHA_FABRICACION', 'FABRICACION'],
  SERIAL: ['SERIAL', 'SERIE'],
  ESTADO_CALIDAD: ['ESTADO CALIDAD', 'ESTADO_CALIDAD', 'ESTADO', 'CALIDAD'],
  QTY: [
    'QTY',
    'CANTIDAD',
    'UNIDADES',
    'Ctd.real dest.',
    'Ctd real dest',
    'Ctd.real dest',
    'Ctd real dest.',
    "Ctd teórica 'desde'",
    "Ctd real 'desde'",
    'Ctd. teór. hacia',
    'Ctd Teórica',
    'Cantidad Teórica',
    'Stock disponible',
  ],
  UOM_CODE: ['UOM_CODE', 'UNIDAD MEDIDA', 'Unidad medida base'],
  REFERENCE: ['REFERENCE', 'REFERENCIA'],
  PRICE: ['PRICE', 'PRECIO'],
  TAXES: ['TAXES', 'IMPUESTOS', 'IVA'],
  IBL_LPN_CODE: ['IBL_LPN_CODE', 'LPN', 'LPN CODE'],
  IBL_WEIGHT: ['IBL_WEIGHT', 'PESO', 'WEIGHT'],
};

// Aqui puedes registrar equivalencias fijas por plantilla de cliente.
// Ejemplo: si el archivo del cliente trae "Pedido", el auto mapeo lo unira con PURCHASE_ORDER.
const CLIENT_TEMPLATE_ALIASES: Record<string, string[]> = {
  PURCHASE_ORDER: ['Pedido Cliente', 'Pedido', 'Pedido OC'],
  N_ORDER: ['Orden', 'Order', 'Numero de Orden', 'Orden Cliente', 'Número de orden de transporte'],
  ORDER2: ['Orden', 'Order', 'Numero de Orden', 'Orden Cliente', 'Número de orden de transporte'],
  PROVIDER_UID: ['Proveedor Cliente', 'Codigo Proveedor', 'Código Proveedor', 'ID Proveedor'],
  SKU: ['Codigo Cliente', 'SKU Cliente', 'Material Cliente'],
  LOTE: ['Lote Cliente', 'Batch Cliente'],
  QTY: ['Cantidad Cliente', 'Cantidad Pedido', 'Cant Pedido', "Ctd real 'desde'"],
  ESTADO_CALIDAD: ['Estado Calidad', 'Calidad'],
  FECHA_DE_VENCIMIENTO: ['FeCaduc/FePreferCons'],
  UOM_CODE: ['Unidad medida base'],
};

function mergeAliasMaps(...maps: Array<Record<string, string[]>>): Record<string, string[]> {
  const merged: Record<string, string[]> = {};

  for (const map of maps) {
    for (const [fieldKey, aliases] of Object.entries(map)) {
      const current = merged[fieldKey] || [];
      merged[fieldKey] = Array.from(new Set([...current, ...aliases]));
    }
  }

  return merged;
}

const FIELD_ALIASES = mergeAliasMaps(BASE_FIELD_ALIASES, CLIENT_TEMPLATE_ALIASES);

@Component({
  selector: 'app-inbound',
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
  ],
  providers: [MessageService],
  templateUrl: './inbound.html',
  styleUrl: './inbound.css',
})
export class Inbound {
  @ViewChild('dtResults') dtResults!: Table;

  readonly targetFields = INBOUND_FIELDS;
  readonly numberKeys = new Set(['QTY', 'PRICE', 'TAXES', 'IBL_WEIGHT']);

  sourceFile = signal<FileData | null>(null);
  sourceRows = signal<any[]>([]);
  availableHeaders = signal<string[]>([]);
  loading = signal(false);
  isDraggingSource = signal(false);
  analysisResult = signal<InboundResult | null>(null);
  searchValue = '';
  showMapping = false;
  useManualMapping = false;

  mapping: Record<string, string> = {};
  fixedValues: Record<string, string> = {};

  breadcrumbItems: MenuItem[] = [{ label: 'Entrada de Mercancía' }];
  breadcrumbHome = { icon: 'pi pi-home', label: 'Inicio', routerLink: '/' };

  isBrowser: boolean;

  constructor(
    private messageService: MessageService,
    @Inject(PLATFORM_ID) platformId: Object,
    private primeng: PrimeNG,
  ) {
    this.isBrowser = isPlatformBrowser(platformId);
    this.initDefaults();
  }

  ngOnInit() {
    // Traducción de filtros PrimeNG al español
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

  get filteredResults(): InboundEntry[] {
    const result = this.analysisResult();
    if (!result) return [];

    const q = this.searchValue.toLowerCase().trim();
    if (!q) return result.results;

    return result.results.filter((row) =>
      Object.values(row).some((value) =>
        String(value ?? '')
          .toLowerCase()
          .includes(q),
      ),
    );
  }

  onDragOver(event: DragEvent): void {
    event.preventDefault();
    this.isDraggingSource.set(true);
  }

  onDragLeave(event: DragEvent): void {
    event.preventDefault();
    this.isDraggingSource.set(false);
  }

  onDrop(event: DragEvent): void {
    event.preventDefault();
    this.isDraggingSource.set(false);
    const files = event.dataTransfer?.files;
    if (files && files.length > 0) this.processFile(files[0]);
  }

  onFileUpload(event: Event): void {
    if (!this.isBrowser) return;
    const input = event.target as HTMLInputElement;
    if (input.files && input.files.length > 0) this.processFile(input.files[0]);
  }

  removeFile(): void {
    this.sourceFile.set(null);
    this.sourceRows.set([]);
    this.availableHeaders.set([]);
    this.analysisResult.set(null);
    this.searchValue = '';
    this.useManualMapping = false;
    this.initDefaults();

    this.messageService.add({
      severity: 'info',
      summary: 'Archivo removido',
      detail: 'Se ha eliminado el archivo de entrada',
      life: 3000,
    });
  }

  clearMappings(): void {
    this.initDefaults();
    this.messageService.add({
      severity: 'info',
      summary: 'Mapeo limpiado',
      detail: 'Se reiniciaron mapeos y valores fijos',
      life: 2500,
    });
  }

  autoMapHeaders(showToast = true): void {
    const headers = this.availableHeaders();
    if (!headers.length) return;

    for (const field of this.targetFields) {
      const match = this.findBestHeader(field.key, headers);
      if (match) this.mapping[field.key] = match;
    }

    if (showToast) {
      this.messageService.add({
        severity: 'info',
        summary: 'Auto mapeo aplicado',
        detail: 'Se asignaron columnas automáticamente',
        life: 2000,
      });
    }
  }

  hasMappedRequiredFields(): boolean {
    return REQUIRED_KEYS.every(
      (key) => !!this.mapping[key] || !!String(this.fixedValues[key] || '').trim(),
    );
  }

  calculateInbound(): void {
    if (!this.sourceFile()) {
      this.messageService.add({
        severity: 'warn',
        summary: 'Sin archivo',
        detail: 'Primero carga un archivo para calcular',
        life: 3000,
      });
      return;
    }

    if (this.useManualMapping) {
      this.processInbound();
      return;
    }

    this.autoMapHeaders(false);
    this.processInbound();
  }

  processInbound(): void {
    const rows = this.sourceRows();
    if (!rows.length) {
      this.messageService.add({
        severity: 'warn',
        summary: 'Sin datos',
        detail: 'Primero carga un archivo con registros',
        life: 3500,
      });
      return;
    }

    const missing = REQUIRED_KEYS.filter(
      (key) => !this.mapping[key] && !(this.fixedValues[key] || '').trim(),
    );
    if (missing.length > 0) {
      this.messageService.add({
        severity: 'error',
        summary: 'Campos requeridos faltantes',
        detail: `Completa mapeo o valor fijo para: ${missing.join(', ')}`,
        life: 6000,
      });
      return;
    }

    this.loading.set(true);

    try {
      const processed = rows.map((row) => this.buildEntry(row));
      const groupedMap: Record<string, InboundEntry> = {};

      for (const item of processed) {
        const key = `${String(item.SKU).trim().toUpperCase()}||${String(item.LOTE).trim().toUpperCase()}`;
        if (!groupedMap[key]) {
          groupedMap[key] = { ...item };
        } else {
          groupedMap[key].QTY += item.QTY;
        }
      }

      const grouped = Object.values(groupedMap);
      const summary = `Procesadas ${processed.length} filas y agrupadas en ${grouped.length} registro(s) por SKU+LOTE`;

      this.analysisResult.set({
        results: grouped,
        summary,
        stats: {
          sourceRows: rows.length,
          processedRows: processed.length,
          groupedRows: grouped.length,
        },
      });

      this.messageService.add({
        severity: 'success',
        summary: 'Proceso completado',
        detail: summary,
        life: 3500,
      });
    } catch {
      this.messageService.add({
        severity: 'error',
        summary: 'Error al procesar',
        detail: 'No fue posible transformar la información',
        life: 5000,
      });
    } finally {
      this.loading.set(false);
    }
  }

  clearFilters(): void {
    this.searchValue = '';
    if (this.dtResults) {
      this.dtResults.clear();
      this.dtResults.filterGlobal('', 'contains');
    }

    this.messageService.add({
      severity: 'info',
      summary: 'Filtros limpiados',
      detail: 'Se han eliminado todos los filtros',
      life: 3000,
    });
  }

  applySearchFilter(): void {
    if (this.dtResults) {
      this.dtResults.filterGlobal(this.searchValue, 'contains');
    }
  }

  setCalculationMode(useManual: boolean): void {
    this.useManualMapping = useManual;
    if (!useManual) {
      this.showMapping = false;
    }
  }

  isNumberField(fieldKey: string): boolean {
    return this.numberKeys.has(fieldKey);
  }

  exportToExcel(): void {
    const result = this.analysisResult();
    if (!result) return;

    const exportRows = result.results.map((r) => ({
      N_ORDER: r.N_ORDER,
      ORDER2: r.ORDER2,
      PURCHASE_ORDER: r.PURCHASE_ORDER,
      INVOICE: r.INVOICE,
      PROVIDER_UID: r.PROVIDER_UID,
      ORDER_DATE: this.formatDateForTextExport(r.ORDER_DATE),
      SERVICE_DATE: this.formatDateForTextExport(r.SERVICE_DATE),
      INBOUNDTYPE_CODE: r.INBOUNDTYPE_CODE,
      NOTE: r.NOTE,
      SKU: r.SKU,
      LOTE: r.LOTE,
      'FECHA DE VENCIMIENTO': this.formatDateForTextExport(r.FECHA_DE_VENCIMIENTO),
      'FECHA DE FABRICACION': this.formatDateForTextExport(r.FECHA_DE_FABRICACION),
      SERIAL: r.SERIAL,
      'ESTADO CALIDAD': r.ESTADO_CALIDAD,
      QTY: r.QTY,
      UOM_CODE: r.UOM_CODE,
      REFERENCE: r.REFERENCE,
      PRICE: r.PRICE,
      TAXES: r.TAXES,
      IBL_LPN_CODE: r.IBL_LPN_CODE,
      IBL_WEIGHT: r.IBL_WEIGHT,
    }));

    const wb = XLSX.utils.book_new();

    const wsResumen = XLSX.utils.json_to_sheet([
      { Metrica: 'Filas fuente', Valor: result.stats.sourceRows },
      { Metrica: 'Filas procesadas', Valor: result.stats.processedRows },
      { Metrica: 'Filas agrupadas SKU+LOTE', Valor: result.stats.groupedRows },
    ]);
    XLSX.utils.book_append_sheet(wb, wsResumen, 'Resumen');

    const wsData = XLSX.utils.json_to_sheet(exportRows);
    this.forceTextFormatInColumns(wsData, [
      'ORDER_DATE',
      'SERVICE_DATE',
      'FECHA DE VENCIMIENTO',
      'FECHA DE FABRICACION',
    ]);

    wsData['!cols'] = [
      14, 12, 18, 14, 16, 14, 14, 16, 24, 14, 14, 20, 20, 14, 16, 10, 12, 14, 10, 10, 14, 12,
    ].map((wch) => ({ wch }));
    XLSX.utils.book_append_sheet(wb, wsData, 'Inbound');

    const date = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `inbound_${date}.xls`, { bookType: 'biff8' });

    this.messageService.add({
      severity: 'success',
      summary: 'Exportado',
      detail: 'Archivo de entrada generado correctamente',
      life: 3000,
    });
  }

  formatFileSize(bytes: number): string {
    if (!bytes) return '0 B';
    const units = ['B', 'KB', 'MB', 'GB'];
    let size = bytes;
    let i = 0;
    while (size >= 1024 && i < units.length - 1) {
      size /= 1024;
      i += 1;
    }
    return `${size.toFixed(size >= 100 || i === 0 ? 0 : 1)} ${units[i]}`;
  }

  private initDefaults(): void {
    const today = this.getTodayISODate();

    this.mapping = {};
    this.fixedValues = {
      ORDER_DATE: today,
      SERVICE_DATE: today,
      INBOUNDTYPE_CODE: 'EXP',
      // Estos campos pueden venir por valor fijo o por mapeo desde el archivo.
      PROVIDER_UID: 'BLAS',
      ESTADO_CALIDAD: 'DSP',
      UOM_CODE: 'UND',
      TAXES: '',
      IBL_WEIGHT: '',
    };
  }

  private getTodayISODate(): string {
    return new Date().toISOString().split('T')[0];
  }

  private processFile(file: File): void {
    if (!this.isBrowser) return;

    const validExtensions = ['.csv', '.xlsx', '.xls', '.txt'];
    const fileExt = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    if (!validExtensions.includes(fileExt)) {
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: 'Use archivos CSV, Excel o TXT',
        life: 5000,
      });
      return;
    }

    this.analysisResult.set(null);

    const reader = new FileReader();
    reader.onload = (e: ProgressEvent<FileReader>) => {
      try {
        const data = e.target?.result || null;
        const rows = this.parseFileData(data, file.name);
        const headers =
          rows.length > 0 ? Object.keys(rows[0]) : this.extractHeaders(data, file.name);

        if (rows.length === 0) {
          this.messageService.add({
            severity: 'error',
            summary: 'Sin datos',
            detail: 'El archivo no contiene filas de datos válidos',
            life: 5000,
          });
          return;
        }

        this.sourceRows.set(rows);
        this.availableHeaders.set(headers);
        this.sourceFile.set({
          name: file.name,
          data,
          size: file.size,
          uploadDate: new Date(),
        });

        this.messageService.add({
          severity: 'success',
          summary: 'Archivo cargado',
          detail: `Entrada: ${rows.length} registros encontrados`,
          life: 3000,
        });

        this.autoMapHeaders(false);

        if (!this.useManualMapping) {
          this.processInbound();
        } else {
          this.messageService.add({
            severity: 'info',
            summary: 'Modo mapeo activo',
            detail: 'Configura el mapeo y luego pulsa Calcular con mapeo',
            life: 3000,
          });
        }
      } catch {
        this.messageService.add({
          severity: 'error',
          summary: 'Error',
          detail: 'No se pudo leer el archivo Excel',
          life: 5000,
        });
      }
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

  private buildEntry(row: any): InboundEntry {
    const getValue = (field: InboundField): any => {
      const fixed = String(this.fixedValues[field.key] || '').trim();
      const mappedHeader = this.mapping[field.key];
      const mappedValue = mappedHeader ? row[mappedHeader] : '';
      const value = fixed !== '' ? fixed : mappedValue;

      if (field.type === 'date') return this.formatExcelDate(value);
      if (field.type === 'number') return this.toNumber(value);
      if (typeof value === 'number') return String(value);
      return value === null || value === undefined ? '' : String(value).trim();
    };

    return {
      N_ORDER: String(getValue(INBOUND_FIELDS[0])),
      ORDER2: String(getValue(INBOUND_FIELDS[1])),
      PURCHASE_ORDER: String(getValue(INBOUND_FIELDS[2])),
      INVOICE: String(getValue(INBOUND_FIELDS[3])),
      PROVIDER_UID: String(getValue(INBOUND_FIELDS[4])),
      ORDER_DATE: String(getValue(INBOUND_FIELDS[5])),
      SERVICE_DATE: String(getValue(INBOUND_FIELDS[6])),
      INBOUNDTYPE_CODE: String(getValue(INBOUND_FIELDS[7])),
      NOTE: String(getValue(INBOUND_FIELDS[8])),
      SKU: String(getValue(INBOUND_FIELDS[9])),
      LOTE: String(getValue(INBOUND_FIELDS[10])),
      FECHA_DE_VENCIMIENTO: String(getValue(INBOUND_FIELDS[11])),
      FECHA_DE_FABRICACION: String(getValue(INBOUND_FIELDS[12])),
      SERIAL: String(getValue(INBOUND_FIELDS[13])),
      ESTADO_CALIDAD: String(getValue(INBOUND_FIELDS[14])),
      QTY: Number(getValue(INBOUND_FIELDS[15])) || 0,
      UOM_CODE: String(getValue(INBOUND_FIELDS[16])),
      REFERENCE: String(getValue(INBOUND_FIELDS[17])),
      PRICE: Number(getValue(INBOUND_FIELDS[18])) || 0,
      TAXES: Number(getValue(INBOUND_FIELDS[19])) || 0,
      IBL_LPN_CODE: String(getValue(INBOUND_FIELDS[20])),
      IBL_WEIGHT: Number(getValue(INBOUND_FIELDS[21])) || 0,
    };
  }

  private findBestHeader(fieldKey: string, headers: string[]): string {
    const aliases = (FIELD_ALIASES[fieldKey] || []).map((a) => this.normalizeText(a));
    const normalizedHeaders = headers.map((h) => ({ raw: h, normalized: this.normalizeText(h) }));

    for (const alias of aliases) {
      const exact = normalizedHeaders.find((h) => h.normalized === alias);
      if (exact) return exact.raw;
    }

    if (EXACT_MATCH_ONLY_FIELDS.has(fieldKey as keyof InboundEntry)) {
      return '';
    }

    for (const alias of aliases) {
      const partial = normalizedHeaders.find(
        (h) => h.normalized.includes(alias) || alias.includes(h.normalized),
      );
      if (partial) return partial.raw;
    }

    return '';
  }

  private parseFileData(data: string | ArrayBuffer | null, fileName: string): any[] {
    if (!data) return [];

    const ext = fileName.split('.').pop()?.toLowerCase();
    if (ext === 'xlsx' || ext === 'xls') {
      const workbook = XLSX.read(data, { type: 'binary' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' }) as any[][];

      const headerIndex = this.findHeaderRowIndex(rows);
      if (headerIndex < 0) return [];

      const headers = (rows[headerIndex] || []).map((h) => String(h || '').trim());
      const result: any[] = [];

      for (let i = headerIndex + 1; i < rows.length; i += 1) {
        const row = rows[i] || [];
        if (!row.some((cell) => String(cell || '').trim() !== '')) continue;

        const obj: Record<string, any> = {};
        headers.forEach((header, idx) => {
          if (header) obj[header] = row[idx] ?? '';
        });
        result.push(obj);
      }

      return result;
    }

    const text = String(data);
    const lines = text.split(/\r?\n/).filter((line) => line.trim().length > 0);
    if (!lines.length) return [];

    const delimiter = this.detectDelimiter(lines[0]);
    const headers = this.splitDelimitedLine(lines[0], delimiter).map((h) => this.cleanHeader(h));
    const result: any[] = [];

    for (let i = 1; i < lines.length; i += 1) {
      const values = this.splitDelimitedLine(lines[i], delimiter);
      const obj: Record<string, any> = {};
      headers.forEach((header, idx) => {
        obj[header] = (values[idx] ?? '').trim();
      });
      result.push(obj);
    }

    return result;
  }

  private extractHeaders(data: string | ArrayBuffer | null, fileName: string): string[] {
    if (!data) return [];
    const rows = this.parseFileData(data, fileName);
    return rows.length ? Object.keys(rows[0]) : [];
  }

  private findHeaderRowIndex(rows: any[][]): number {
    if (!rows.length) return -1;
    const knownTerms = [
      'sku',
      'material',
      'lote',
      'batch',
      'descripcion',
      'descripción',
      'unidades',
      'cantidad',
      'serial',
      'pedido',
      'factura',
      'vencimiento',
    ];

    const maxScan = Math.min(rows.length, 60);
    for (let i = 0; i < maxScan; i += 1) {
      const row = rows[i] || [];
      const normalized = row.map((cell) => this.normalizeText(String(cell || ''))).filter(Boolean);
      if (normalized.length < 2) continue;

      let matches = 0;
      for (const cell of normalized) {
        if (knownTerms.some((term) => cell.includes(term))) matches += 1;
      }

      if (matches >= 2) return i;
    }

    return 0;
  }

  private formatExcelDate(value: any): string {
    if (value === null || value === undefined || value === '') return '';

    if (typeof value === 'number' && Number.isFinite(value)) {
      const utc = Date.UTC(1899, 11, 30) + Math.round(value * 86400 * 1000);
      const d = new Date(utc);
      if (!Number.isNaN(d.getTime())) return d.toISOString().split('T')[0];
    }

    const str = String(value).trim();
    if (!str) return '';

    const parsed = new Date(str);
    if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().split('T')[0];

    return str;
  }

  private formatDateForTextExport(value: any): string {
    if (value === null || value === undefined || value === '') return '';

    const str = String(value).trim();
    if (!str) return '';

    const isoMatch = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (isoMatch) {
      return `${isoMatch[3]}/${isoMatch[2]}/${isoMatch[1]}`;
    }

    const slashMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (slashMatch) {
      const dd = slashMatch[1].padStart(2, '0');
      const mm = slashMatch[2].padStart(2, '0');
      const yyyy = slashMatch[3];
      return `${dd}/${mm}/${yyyy}`;
    }

    const parsed = new Date(str);
    if (!Number.isNaN(parsed.getTime())) {
      const dd = String(parsed.getUTCDate()).padStart(2, '0');
      const mm = String(parsed.getUTCMonth() + 1).padStart(2, '0');
      const yyyy = parsed.getUTCFullYear();
      return `${dd}/${mm}/${yyyy}`;
    }

    return str;
  }

  private forceTextFormatInColumns(sheet: XLSX.WorkSheet, headers: string[]): void {
    const ref = sheet['!ref'];
    if (!ref) return;

    const range = XLSX.utils.decode_range(ref);
    const headerSet = new Set(headers);
    const targetColumns: number[] = [];

    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const headerCellRef = XLSX.utils.encode_cell({ r: range.s.r, c });
      const headerValue = String(sheet[headerCellRef]?.v ?? '').trim();
      if (headerSet.has(headerValue)) {
        targetColumns.push(c);
      }
    }

    for (let r = range.s.r + 1; r <= range.e.r; r += 1) {
      for (const c of targetColumns) {
        const cellRef = XLSX.utils.encode_cell({ r, c });
        const cell = sheet[cellRef];
        if (!cell) continue;

        cell.t = 's';
        cell.z = '@';
        cell.v = String(cell.v ?? '');
      }
    }
  }

  private toNumber(value: any): number {
    if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
    const cleaned = String(value ?? '')
      .replace(/\s/g, '')
      .replace(/\./g, '')
      .replace(',', '.');
    const parsed = Number(cleaned);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  private cleanHeader(header: string): string {
    return header.replace(/^['"]|['"]$/g, '').trim();
  }

  private detectDelimiter(headerLine: string): string {
    const comma = (headerLine.match(/,/g) || []).length;
    const semicolon = (headerLine.match(/;/g) || []).length;
    const tab = (headerLine.match(/\t/g) || []).length;
    if (tab >= comma && tab >= semicolon) return '\t';
    return semicolon > comma ? ';' : ',';
  }

  private splitDelimitedLine(line: string, delimiter: string): string[] {
    const values: string[] = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < line.length; i += 1) {
      const ch = line[i];

      if (ch === '"') {
        if (inQuotes && line[i + 1] === '"') {
          current += '"';
          i += 1;
        } else {
          inQuotes = !inQuotes;
        }
        continue;
      }

      if (!inQuotes && ch === delimiter) {
        values.push(current);
        current = '';
      } else {
        current += ch;
      }
    }

    values.push(current);
    return values;
  }

  private normalizeText(value: string): string {
    return value
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-zA-Z0-9]+/g, ' ')
      .toLowerCase()
      .replace(/\s+/g, ' ')
      .trim();
  }
}
