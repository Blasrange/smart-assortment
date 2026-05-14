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

export const masterColumnMapping = {
  sku: [
    'Codigo de articulo',
    'Codigo de artículo',
    'Código de articulo',
    'Código de artículo',
    'Material',
    'SKU',
    'Codigo',
  ],
  diasMinimos: [
    'Vida util del producto',
    'Vida útil del producto',
    'Vida Util',
    'Dias Minimos',
    'Requerido',
    'VIDA UTIL REQUERIDA',
  ],
};

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
    'Fecha recepción',
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

interface ShelfLifeRow {
  sku: string;
  descripcion: string;
  lpn: string;
  localizacion: string;
  lote: string;
  fechaVencimiento: string;
  diasFPC: number;
  diasMinimosMaestra: number;
  cumple: boolean;
  estado: 'OK' | 'ALERTA';
}

interface ShelfLifeAnalysisResult {
  summary: string;
  results: ShelfLifeRow[];
  stats: {
    totalRegistros: number;
    totalAlertas: number;
    totalOK: number;
    porcentajeCumplimiento: number;
  };
}

@Component({
  selector: 'app-shelf-life',
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
  templateUrl: './shelf-life.html',
  styleUrl: './shelf-life.css',
})
export class ShelfLife {
  @ViewChild('dt1') dt1!: Table;

  inventoryFile = signal<FileData | null>(null);
  masterFile = signal<FileData | null>(null);
  analysisResult = signal<ShelfLifeAnalysisResult | null>(null);
  loading = signal(false);
  error = signal<string | null>(null);
  searchValue = '';
  isBrowser: boolean;

  inventoryHeaderMap: Record<string, string> = {};
  masterHeaderMap: Record<string, string> = {};
  isDraggingInventory = signal(false);
  isDraggingMaster = signal(false);

  readonly breadcrumbItems: MenuItem[] = [{ label: 'Vida Util', routerLink: '/shelf-life' }];
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

  onDragOver(event: DragEvent, type: 'inventory' | 'master') {
    event.preventDefault();
    event.stopPropagation();
    if (type === 'inventory') {
      this.isDraggingInventory.set(true);
      return;
    }
    this.isDraggingMaster.set(true);
  }

  onDragLeave(event: DragEvent, type: 'inventory' | 'master') {
    event.preventDefault();
    event.stopPropagation();
    if (type === 'inventory') {
      this.isDraggingInventory.set(false);
      return;
    }
    this.isDraggingMaster.set(false);
  }

  onDrop(event: DragEvent, type: 'inventory' | 'master') {
    event.preventDefault();
    event.stopPropagation();
    if (type === 'inventory') {
      this.isDraggingInventory.set(false);
    } else {
      this.isDraggingMaster.set(false);
    }

    const files = event.dataTransfer?.files;
    if (files && files.length > 0) {
      this.processFile(files[0], type);
    }
  }

  onFileUpload(type: 'inventory' | 'master', event: Event) {
    if (!this.isBrowser) return;
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    this.processFile(file, type);
  }

  private processFile(file: File, type: 'inventory' | 'master') {
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

      const mapping: Record<string, string[]> =
        type === 'inventory'
          ? (inventoryColumnMapping as Record<string, string[]>)
          : (masterColumnMapping as Record<string, string[]>);
      const requiredKeys: string[] =
        type === 'inventory' ? ['sku', 'diasFPC'] : ['sku', 'diasMinimos'];
      const headersLower = headers.map((h) => h.toLowerCase().trim());
      const missing: string[] = [];

      for (const key of requiredKeys) {
        const aliases = mapping[key].map((a: string) => a.toLowerCase().trim());
        const found = headersLower.some((h) => aliases.includes(h));
        if (!found) {
          missing.push(mapping[key][0]);
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
      const headerMap = this.mapHeadersToStandard(fileHeaders, mapping);

      const fileData: FileData = {
        name: file.name,
        data,
        size: file.size,
        uploadDate: new Date(),
      };

      if (type === 'inventory') {
        this.inventoryFile.set(fileData);
        this.inventoryHeaderMap = headerMap;
      } else {
        this.masterFile.set(fileData);
        this.masterHeaderMap = headerMap;
      }

      this.messageService.add({
        severity: 'success',
        summary: 'Archivo cargado',
        detail: `${type === 'inventory' ? 'Inventario' : 'Maestra'}: ${rows.length} registros encontrados`,
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

  removeFile(type: 'inventory' | 'master') {
    if (type === 'inventory') {
      this.inventoryFile.set(null);
      this.inventoryHeaderMap = {};
    } else {
      this.masterFile.set(null);
      this.masterHeaderMap = {};
    }

    this.analysisResult.set(null);
    this.messageService.add({
      severity: 'info',
      summary: 'Archivo removido',
      detail: `Se ha eliminado el archivo de ${type === 'inventory' ? 'inventario' : 'maestra'}`,
      life: 3000,
    });
  }

  generateShelfLifeAnalysis() {
    if (!this.inventoryFile() || !this.masterFile()) {
      this.messageService.add({
        severity: 'warn',
        summary: 'Archivos faltantes',
        detail: 'Debes subir ambos archivos (Inventario y Maestra) para analizar',
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
        const errorMsg = 'Error al procesar los archivos: ' + (err as Error).message;
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

    const wb = XLSX.utils.book_new();

    const resumenData = [
      { Concepto: 'ANALISIS VIDA UTIL', Valor: '', Unidad: '' },
      {
        Concepto: '  - Registros analizados',
        Valor: result.stats.totalRegistros,
        Unidad: 'Registros',
      },
      { Concepto: '  - Alertas', Valor: result.stats.totalAlertas, Unidad: 'Registros' },
      { Concepto: '  - OK', Valor: result.stats.totalOK, Unidad: 'Registros' },
      {
        Concepto: '  - Cumplimiento',
        Valor: `${result.stats.porcentajeCumplimiento}%`,
        Unidad: '',
      },
    ];

    const wsResumen = XLSX.utils.json_to_sheet(resumenData);
    wsResumen['!cols'] = [42, 18, 14].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsResumen, 'Resumen');

    const exportData = result.results.map((r) => ({
      SKU: r.sku,
      Descripcion: r.descripcion,
      LPN: r.lpn,
      Localizacion: r.localizacion,
      Lote: r.lote,
      Fecha_Vencimiento: r.fechaVencimiento || '',
      Dias_FPC: r.diasFPC,
      Dias_Minimos_Maestra: r.diasMinimosMaestra,
      Estado: r.estado,
      Cumple: r.cumple ? 'Si' : 'No',
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    ws['!cols'] = [16, 36, 18, 20, 14, 18, 12, 20, 10, 10].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws, 'Vida Util');

    const fecha = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `vida_util_${fecha}.xlsx`);

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

  getStatusSeverity(estado: 'OK' | 'ALERTA'): 'success' | 'danger' {
    return estado === 'OK' ? 'success' : 'danger';
  }

  private processAnalysis(): ShelfLifeAnalysisResult {
    const inventoryRaw = this.inventoryFile()?.data;
    const masterRaw = this.masterFile()?.data;
    const inventoryName = this.inventoryFile()?.name || '';
    const masterName = this.masterFile()?.name || '';

    if (typeof inventoryRaw !== 'string' || typeof masterRaw !== 'string') {
      throw new Error('Archivos no validos');
    }

    const inventoryRows = this.parseFileData(inventoryRaw, inventoryName);
    const masterRows = this.parseFileData(masterRaw, masterName);
    const invMap = this.inventoryHeaderMap;
    const masterMap = this.masterHeaderMap;

    const norm = (v: any) =>
      String(v || '')
        .trim()
        .toUpperCase();

    const masterLookup = masterRows.reduce(
      (acc, item) => {
        const sku = norm(this.getNestedValue(item, masterMap['sku'] || 'SKU', ''));
        if (!sku) return acc;
        acc[sku] = this.toNumber(
          this.getNestedValue(item, masterMap['diasMinimos'] || 'Vida util del producto', 0),
        );
        return acc;
      },
      {} as Record<string, number>,
    );

    const results: ShelfLifeRow[] = inventoryRows.map((item) => {
      const skuNorm = norm(this.getNestedValue(item, invMap['sku'] || 'SKU', ''));
      const diasLimiteMaestra = masterLookup[skuNorm] || 0;
      const diasFPC = this.toNumber(this.getNestedValue(item, invMap['diasFPC'] || 'FPC', 0));
      const cumple = diasFPC <= diasLimiteMaestra;

      return {
        sku: String(this.getNestedValue(item, invMap['sku'] || 'SKU', '')).trim(),
        descripcion: String(
          this.getNestedValue(item, invMap['descripcion'] || 'Descripcion', ''),
        ).trim(),
        lpn: String(this.getNestedValue(item, invMap['lpn'] || 'LPN', '')).trim(),
        localizacion: String(
          this.getNestedValue(item, invMap['localizacion'] || 'Localizacion', ''),
        ).trim(),
        lote: String(this.getNestedValue(item, invMap['lote'] || 'Lote', '')).trim(),
        fechaVencimiento: String(
          this.getNestedValue(item, invMap['fechaVencimiento'] || 'Fecha de vencimiento', ''),
        ).trim(),
        diasFPC,
        diasMinimosMaestra: diasLimiteMaestra,
        cumple,
        estado: cumple ? 'OK' : 'ALERTA',
      };
    });

    results.sort((a, b) => {
      if (a.cumple !== b.cumple) return a.cumple ? 1 : -1;
      return b.diasFPC - a.diasFPC;
    });

    const totalRegistros = results.length;
    const totalAlertas = results.filter((r) => !r.cumple).length;
    const totalOK = totalRegistros - totalAlertas;
    const porcentajeCumplimiento =
      totalRegistros === 0 ? 100 : Math.round((totalOK / totalRegistros) * 100);

    return {
      summary: `Analisis completado: ${totalRegistros} registros evaluados, ${totalAlertas} alertas de vida util.`,
      results,
      stats: {
        totalRegistros,
        totalAlertas,
        totalOK,
        porcentajeCumplimiento,
      },
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
