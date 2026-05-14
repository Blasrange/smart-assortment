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

const IGNORED_LOCATIONS = ['PDIF-INV-1-10', 'DEV-1-10', 'PDIF-RES-1-10', 'DEV-RES-1-10'];
const IGNORED_LOCATIONS_MUELLE = ['MUELLE'];
const MS_PER_DAY = 24 * 60 * 60 * 1000;

export const inventoryColumnMapping = {
  sku: ['SKU', 'Item Code'],
  lpn: ['LPN', 'Pallet ID'],
  descripcion: ['Descripcion', 'Description'],
  localizacion: ['Localizacion', 'Location'],
  disponible: ['Disponible', 'Available'],
  estado: ['Estado', 'Status'],
  fechaEntrada: [
    'Fecha de entrada',
    'Fecha Entrada',
    'Fecha ingreso',
    'Fecha de ingreso',
    'Fecha recepcion',
    'Fecha recepcion',
    'Entry Date',
    'Receipt Date',
  ],
  fechaVencimiento: ['Fecha de vencimiento', 'Expiration', 'fecha caducidad'],
  diasFPC: ['FPC', 'Days to Exp', 'DIAS FPC'],
  lote: ['Lote', 'Batch', 'Ce. Lote'],
};

interface FileData {
  name: string;
  data: string | ArrayBuffer | null;
  size: number;
  uploadDate: Date;
}

type AgeBucket = '0-3 meses' | '3-6 meses' | '6-12 meses' | '> 12 meses' | 'Sin fecha de entrada';

interface AgeInventoryRow {
  sku: string;
  descripcion: string;
  lpn: string;
  localizacion: string;
  lote: string;
  disponible: number;
  estado: string;
  fechaEntrada: string;
  diasEnInventario: number | null;
  rangoEdad: AgeBucket;
  fechaVencimiento: string | null;
}

interface AgeAnalysisResult {
  summary: string;
  results: AgeInventoryRow[];
  stats: {
    totalRegistros: number;
    masDe12Meses: number;
    de6a12Meses: number;
    de3a6Meses: number;
    de0a3Meses: number;
    sinFechaEntrada: number;
  };
}

@Component({
  selector: 'app-age-inventory',
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
  templateUrl: './age-inventory.html',
  styleUrl: './age-inventory.css',
})
export class AgeInventory {
  @ViewChild('dt1') dt1!: Table;

  inventoryFile = signal<FileData | null>(null);
  analysisResult = signal<AgeAnalysisResult | null>(null);
  loading = signal(false);
  error = signal<string | null>(null);
  searchValue = '';
  isBrowser: boolean;

  inventoryHeaderMap: Record<string, string> = {};
  isDraggingInventory = signal(false);

  readonly breadcrumbItems: MenuItem[] = [
    { label: 'Antigüedad de Inventario', routerLink: '/age-inventory' },
  ];
  readonly breadcrumbHome: MenuItem = { icon: 'pi pi-home', label: 'Inicio', routerLink: '/' };

  constructor(
    @Inject(PLATFORM_ID) platformId: Object,
    private messageService: MessageService,
    private primeng: PrimeNG,
  ) {
    this.isBrowser = isPlatformBrowser(platformId);
  }

  ngOnInit() {
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
        after: 'Despues',
        clear: 'Limpiar',
        apply: 'Aplicar',
        matchAll: 'Cumplir todas',
        matchAny: 'Cumplir alguna',
        addRule: 'Agregar regla',
        removeRule: 'Eliminar regla',
      });
    }
  }

  onDragOver(event: DragEvent) {
    event.preventDefault();
    event.stopPropagation();
    this.isDraggingInventory.set(true);
  }

  onDragLeave(event: DragEvent) {
    event.preventDefault();
    event.stopPropagation();
    this.isDraggingInventory.set(false);
  }

  onDrop(event: DragEvent) {
    event.preventDefault();
    event.stopPropagation();
    this.isDraggingInventory.set(false);

    const files = event.dataTransfer?.files;
    if (files && files.length > 0) {
      this.processFile(files[0]);
    }
  }

  onFileUpload(event: Event) {
    if (!this.isBrowser) return;
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    this.processFile(file);
  }

  private processFile(file: File) {
    if (!this.isBrowser) return;

    const validExtensions = ['.csv', '.xlsx', '.xls', '.txt'];
    const fileExt = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    if (!validExtensions.includes(fileExt)) {
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: 'Formato no soportado. Use CSV, Excel o TXT',
        life: 5000,
      });
      return;
    }

    this.analysisResult.set(null);
    this.error.set(null);

    const reader = new FileReader();
    reader.onload = () => {
      const data = reader.result;
      const ext = file.name.split('.').pop()?.toLowerCase();
      let headers: string[] = [];

      if (ext === 'xlsx' || ext === 'xls') {
        try {
          const workbook = XLSX.read(data, { type: 'binary' });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[];
          headers = (rows[0] || []).map((h: any) => String(h).trim());
        } catch {
          this.messageService.add({
            severity: 'error',
            summary: 'Error',
            detail: 'No se pudo leer el archivo Excel',
            life: 5000,
          });
          return;
        }
      } else if (typeof data === 'string') {
        const lines = data.split(/\r?\n/).filter((l) => l.trim().length > 0);
        if (lines.length > 0) {
          let delimiter = ',';
          if (lines[0].includes(';')) delimiter = ';';
          headers = lines[0].split(delimiter).map((h) => h.trim().replace(/^['"]|['"]$/g, ''));
        }
      }

      const requiredKeys: Array<keyof typeof inventoryColumnMapping> = [
        'sku',
        'localizacion',
        'disponible',
      ];
      const headersLower = headers.map((h) => h.toLowerCase().trim());
      const missing: string[] = [];

      for (const key of requiredKeys) {
        const aliases = inventoryColumnMapping[key].map((a) => a.toLowerCase().trim());
        const found = headersLower.some((h) => aliases.includes(h));
        if (!found) {
          missing.push(inventoryColumnMapping[key][0]);
        }
      }

      if (missing.length > 0) {
        this.messageService.add({
          severity: 'error',
          summary: 'Columnas faltantes',
          detail: `El archivo no contiene las columnas requeridas: ${missing.join(', ')}`,
          life: 7000,
        });
        return;
      }

      const rows = this.parseFileData(data as string | ArrayBuffer, file.name);
      const fileHeaders = rows.length > 0 ? Object.keys(rows[0]) : headers;
      this.inventoryHeaderMap = this.mapHeadersToStandard(fileHeaders, inventoryColumnMapping);
      this.inventoryFile.set({
        name: file.name,
        data,
        size: file.size,
        uploadDate: new Date(),
      });

      this.messageService.add({
        severity: 'success',
        summary: 'Archivo cargado',
        detail: `Inventario: ${rows.length} registros encontrados`,
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

    const ext = file.name.split('.').pop()?.toLowerCase();
    if (ext === 'xlsx' || ext === 'xls') {
      reader.readAsBinaryString(file);
      return;
    }
    reader.readAsText(file, 'UTF-8');
  }

  removeFile() {
    this.inventoryFile.set(null);
    this.inventoryHeaderMap = {};
    this.analysisResult.set(null);

    this.messageService.add({
      severity: 'info',
      summary: 'Archivo removido',
      detail: 'Se ha eliminado el archivo de inventario',
      life: 3000,
    });
  }

  generateAgeAnalysis() {
    if (!this.inventoryFile()) {
      this.messageService.add({
        severity: 'warn',
        summary: 'Archivo faltante',
        detail: 'Debes subir un archivo de inventario para analizar',
        life: 5000,
      });
      return;
    }

    this.loading.set(true);
    this.error.set(null);

    setTimeout(() => {
      try {
        const result = this.processAnalysis();
        this.analysisResult.set(result);
        this.loading.set(false);
        this.messageService.add({
          severity: 'success',
          summary: 'Analisis completado',
          detail: result.summary,
          life: 5000,
        });
      } catch (err) {
        const errorMsg = 'Error al procesar el archivo: ' + (err as Error).message;
        this.error.set(errorMsg);
        this.loading.set(false);
        this.messageService.add({
          severity: 'error',
          summary: 'Error',
          detail: errorMsg,
          life: 7000,
        });
      }
    }, 100);
  }

  exportToExcel() {
    const result = this.analysisResult();
    if (!result) return;

    const rows = result.results;
    const formatAvg = (value: number | null): string => {
      if (value === null || !Number.isFinite(value)) return '-';
      return value.toFixed(1).replace('.', ',');
    };

    const buildRangeRow = (
      range: AgeBucket,
      label: string,
    ): [string, number, number, number, string] => {
      const rangeRows = rows.filter((r) => r.rangoEdad === range);
      const skuCount = new Set(
        rangeRows.map((r) => r.sku).filter((sku) => String(sku).trim() !== ''),
      ).size;
      const unidades = rangeRows.reduce((sum, r) => sum + r.disponible, 0);
      const registros = rangeRows.length;
      const days = rangeRows
        .map((r) => r.diasEnInventario)
        .filter((d): d is number => d !== null && Number.isFinite(d));
      const promedio = days.length > 0 ? days.reduce((sum, d) => sum + d, 0) / days.length : null;

      return [label, skuCount, unidades, registros, formatAvg(promedio)];
    };

    const summaryRows: Array<[string, number, number, number, string]> = [
      buildRangeRow('0-3 meses', '0-3 meses'),
      buildRangeRow('3-6 meses', '3-6 meses'),
      buildRangeRow('6-12 meses', '6-12 meses'),
      buildRangeRow('> 12 meses', '>12 meses'),
      buildRangeRow('Sin fecha de entrada', 'Sin fecha'),
    ];

    const totalSkus = new Set(rows.map((r) => r.sku).filter((sku) => String(sku).trim() !== ''))
      .size;
    const totalUnidades = rows.reduce((sum, r) => sum + r.disponible, 0);
    const totalRegistros = rows.length;
    const totalDays = rows
      .map((r) => r.diasEnInventario)
      .filter((d): d is number => d !== null && Number.isFinite(d));
    const totalPromedio =
      totalDays.length > 0 ? totalDays.reduce((sum, d) => sum + d, 0) / totalDays.length : null;

    const wb = XLSX.utils.book_new();
    const resumenAoa: Array<Array<string | number>> = [
      ['Rango de Antiguedad', 'SKUs', 'Unidades', 'Registros', 'Promedio Dias'],
      ...summaryRows,
      ['TOTAL', totalSkus, totalUnidades, totalRegistros, formatAvg(totalPromedio)],
    ];

    const wsResumen = XLSX.utils.aoa_to_sheet(resumenAoa);
    wsResumen['!cols'] = [24, 10, 14, 12, 14].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsResumen, 'Resumen');

    const exportData = result.results.map((r) => ({
      SKU: r.sku,
      Descripcion: r.descripcion,
      LPN: r.lpn,
      Localizacion: r.localizacion,
      Lote: r.lote,
      Disponible: r.disponible,
      Estado: r.estado,
      Fecha_Entrada: r.fechaEntrada,
      Fecha_Vencimiento: r.fechaVencimiento || '',
      Dias_En_Inventario: r.diasEnInventario,
      Rango_Edad: r.rangoEdad,
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    ws['!cols'] = [16, 38, 18, 20, 16, 14, 18, 16, 18, 18, 18].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws, 'Antigüedad de Inventario');

    const fecha = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `Antigüedad_inventario_${fecha}.xlsx`);

    this.messageService.add({
      severity: 'success',
      summary: 'Exportado',
      detail: 'Archivo Excel generado',
      life: 3000,
    });
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
      detail: 'Se han eliminado todos los filtros',
      life: 3000,
    });
  }

  formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  getRangeSeverity(range: AgeBucket): 'success' | 'info' | 'warn' | 'danger' | 'secondary' {
    if (range === '> 12 meses') return 'danger';
    if (range === '6-12 meses') return 'warn';
    if (range === '3-6 meses') return 'info';
    if (range === '0-3 meses') return 'success';
    return 'secondary';
  }

  private excelSerialToDate(serial: number): Date | null {
    if (!Number.isFinite(serial)) return null;

    // Excel stores dates as days since 1899-12-30 in the 1900 date system.
    const excelEpochOffset = 25569;
    const wholeDays = Math.floor(serial);
    const dayFraction = serial - wholeDays;
    const utcMs =
      (wholeDays - excelEpochOffset) * MS_PER_DAY + Math.round(dayFraction * MS_PER_DAY);
    const parsedDate = new Date(utcMs);

    return Number.isNaN(parsedDate.getTime()) ? null : parsedDate;
  }

  private parseInventoryDate(value: unknown): Date | null {
    if (!value) return null;

    if (typeof value === 'number') {
      return this.excelSerialToDate(value);
    }

    if (value instanceof Date) {
      return Number.isNaN(value.getTime()) ? null : value;
    }

    const rawValue = String(value).trim();
    if (!rawValue) return null;

    // Example: 46020.51582175926 (Excel serial date)
    if (/^\d+(\.\d+)?$/.test(rawValue)) {
      return this.excelSerialToDate(Number(rawValue));
    }

    const isoCandidate = new Date(rawValue);
    if (!Number.isNaN(isoCandidate.getTime())) {
      return isoCandidate;
    }

    const match = rawValue.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
    if (!match) return null;

    const [, day, month, year] = match;
    const normalizedYear = year.length === 2 ? `20${year}` : year;
    const parsedDate = new Date(Date.UTC(Number(normalizedYear), Number(month) - 1, Number(day)));

    return Number.isNaN(parsedDate.getTime()) ? null : parsedDate;
  }

  private getAgeBucket(daysInInventory: number | null): AgeBucket {
    if (daysInInventory === null) return 'Sin fecha de entrada';
    if (daysInInventory <= 90) return '0-3 meses';
    if (daysInInventory <= 180) return '3-6 meses';
    if (daysInInventory <= 365) return '6-12 meses';
    return '> 12 meses';
  }

  private processAnalysis(): AgeAnalysisResult {
    const raw = this.inventoryFile()?.data;
    const fileName = this.inventoryFile()?.name || '';
    if (typeof raw !== 'string') {
      throw new Error('Archivo no valido');
    }

    const rows = this.parseFileData(raw, fileName);
    const map = this.inventoryHeaderMap;

    const today = new Date();
    const todayUtc = Date.UTC(today.getUTCFullYear(), today.getUTCMonth(), today.getUTCDate());
    const excludedLocationPrefixes = [...IGNORED_LOCATIONS, ...IGNORED_LOCATIONS_MUELLE].map((l) =>
      l.trim().toUpperCase(),
    );

    const inventoryWithoutMuelle = rows.filter((item) => {
      const location = String(this.getNestedValue(item, map['localizacion'] || 'Localizacion', ''))
        .trim()
        .toUpperCase();
      return !excludedLocationPrefixes.some((prefix) => location.startsWith(prefix));
    });

    const rank: Record<AgeBucket, number> = {
      '> 12 meses': 0,
      '6-12 meses': 1,
      '3-6 meses': 2,
      '0-3 meses': 3,
      'Sin fecha de entrada': 4,
    };

    const results: AgeInventoryRow[] = inventoryWithoutMuelle
      .map((item) => {
        const parsedEntryDate = this.parseInventoryDate(
          this.getNestedValue(item, map['fechaEntrada'] || 'Fecha de entrada', ''),
        );
        const entryUtc = parsedEntryDate
          ? Date.UTC(
              parsedEntryDate.getUTCFullYear(),
              parsedEntryDate.getUTCMonth(),
              parsedEntryDate.getUTCDate(),
            )
          : null;

        const daysInInventory =
          entryUtc === null ? null : Math.max(0, Math.floor((todayUtc - entryUtc) / MS_PER_DAY));
        const ageBucket = this.getAgeBucket(daysInInventory);

        const fechaVencRaw = this.getNestedValue(
          item,
          map['fechaVencimiento'] || 'Fecha de vencimiento',
          null,
        );
        let parsedExpiryDate: string | null = null;
        if (fechaVencRaw) {
          const parsed = this.parseInventoryDate(fechaVencRaw);
          parsedExpiryDate = parsed ? parsed.toISOString().split('T')[0] : String(fechaVencRaw);
        }

        return {
          sku: String(this.getNestedValue(item, map['sku'] || 'SKU', '')).trim(),
          descripcion: String(
            this.getNestedValue(item, map['descripcion'] || 'Descripcion', ''),
          ).trim(),
          lpn: String(this.getNestedValue(item, map['lpn'] || 'LPN', '')).trim(),
          localizacion: String(
            this.getNestedValue(item, map['localizacion'] || 'Localizacion', ''),
          ).trim(),
          lote: String(this.getNestedValue(item, map['lote'] || 'Lote', '')).trim(),
          disponible: this.toNumber(
            this.getNestedValue(item, map['disponible'] || 'Disponible', 0),
          ),
          estado: String(this.getNestedValue(item, map['estado'] || 'Estado', '')).trim(),
          fechaEntrada: parsedEntryDate
            ? parsedEntryDate.toISOString().split('T')[0]
            : String(
                this.getNestedValue(item, map['fechaEntrada'] || 'Fecha de entrada', '') || '',
              ),
          diasEnInventario: daysInInventory,
          rangoEdad: ageBucket,
          fechaVencimiento: parsedExpiryDate,
        };
      })
      .sort((a, b) => {
        if (rank[a.rangoEdad] !== rank[b.rangoEdad]) {
          return rank[a.rangoEdad] - rank[b.rangoEdad];
        }
        return (b.diasEnInventario ?? -1) - (a.diasEnInventario ?? -1);
      });

    const stats = {
      totalRegistros: results.length,
      masDe12Meses: results.filter((r) => r.rangoEdad === '> 12 meses').length,
      de6a12Meses: results.filter((r) => r.rangoEdad === '6-12 meses').length,
      de3a6Meses: results.filter((r) => r.rangoEdad === '3-6 meses').length,
      de0a3Meses: results.filter((r) => r.rangoEdad === '0-3 meses').length,
      sinFechaEntrada: results.filter((r) => r.rangoEdad === 'Sin fecha de entrada').length,
    };

    return {
      summary: `Analisis completado: ${stats.totalRegistros} registros analizados por antigüedad.`,
      results,
      stats,
    };
  }

  private mapHeadersToStandard(
    headers: string[],
    mapping: Record<string, string[]>,
  ): Record<string, string> {
    const result: Record<string, string> = {};
    const headersLower = headers.map((h) => h.toLowerCase().trim());
    for (const key in mapping) {
      const aliases = mapping[key].map((a) => a.toLowerCase().trim());
      for (const alias of aliases) {
        const foundIndex = headersLower.indexOf(alias);
        if (foundIndex !== -1) {
          result[key] = headers[foundIndex];
          break;
        }
      }
    }
    return result;
  }

  private parseCSV(data: string): any[] {
    const lines = data.split(/\r?\n/).filter((l) => l.trim().length > 0);
    if (lines.length < 2) return [];

    let delimiter = ',';
    if (lines[0].includes(';')) delimiter = ';';

    const headers = lines[0].split(delimiter).map((h) => h.trim().replace(/^['"]|['"]$/g, ''));
    return lines.slice(1).map((line) => {
      const values = line.split(delimiter);
      const obj: any = {};
      headers.forEach((h, i) => {
        obj[h] = values[i] ? values[i].trim().replace(/^['"]|['"]$/g, '') : '';
      });
      return obj;
    });
  }

  private parseFileData(data: string | ArrayBuffer, fileName: string): any[] {
    const ext = fileName.split('.').pop()?.toLowerCase();

    if (ext === 'csv' || ext === 'txt') {
      if (typeof data === 'string') {
        return this.parseCSV(data);
      }
      return [];
    }

    if ((ext === 'xlsx' || ext === 'xls') && typeof data === 'string') {
      try {
        const workbook = XLSX.read(data, { type: 'binary' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        return XLSX.utils.sheet_to_json(sheet);
      } catch {
        return [];
      }
    }

    return typeof data === 'string' ? this.parseCSV(data) : [];
  }

  private getNestedValue(obj: any, path: string, defaultValue: any = ''): any {
    if (!obj || !path) return defaultValue;
    const value = obj[path];
    if (value === undefined || value === null || value === '') return defaultValue;
    return value;
  }

  private toNumber(value: any): number {
    if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
    const raw = String(value ?? '').trim();
    if (!raw) return 0;

    let normalized = raw.replace(/\s/g, '').replace(/[^\d,.-]/g, '');
    if (normalized.includes(',') && !normalized.includes('.')) {
      normalized = normalized.replace(',', '.');
    } else if (normalized.includes(',') && normalized.includes('.')) {
      normalized = normalized.replace(/,/g, '');
    }

    const num = Number(normalized);
    return Number.isFinite(num) ? num : 0;
  }
}
