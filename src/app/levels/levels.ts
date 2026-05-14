import { Component, signal, Inject, PLATFORM_ID, ViewChild } from '@angular/core';
import { isPlatformBrowser } from '@angular/common';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { Table, TableModule } from 'primeng/table';
import { ButtonModule } from 'primeng/button';
import { InputTextModule } from 'primeng/inputtext';
import { ProgressSpinnerModule } from 'primeng/progressspinner';
import { TooltipModule } from 'primeng/tooltip';
import { TagModule } from 'primeng/tag';
import { BreadcrumbModule } from 'primeng/breadcrumb';
import { ToastModule } from 'primeng/toast';
import { MessageService, MenuItem } from 'primeng/api';
import { PrimeNG } from 'primeng/config';
import * as XLSX from 'xlsx';

const VALID_STATUSES = ['STOCK EN ALMACEN LIBRE', 'DISPONIBLE', '13 - CCL  DISPONIBLE'];
const IGNORED_LOCATIONS = ['PDIF-INV-1-10', 'DEV-1-10', 'PDIF-RES-1-10', 'DEV-RES-1-10'];
const PICKING_LEVELS = ['0', '2', '3', '4', '5', '6', '7', '8', '9', '10', '15'];
const RESERVE_LEVELS = ['20', '30', '40', '50', '60', '70'];
const ADDITIONAL_RESERVE_LOCATIONS = ['MUELLE ENTRADA'];

export const levelColumnMapping = {
  sku: ['SKU', 'Item Code', 'Codigo', 'Código', 'Material', 'CODIGO', 'CODE'],
  lpn: ['LPN', 'Pallet ID', 'Lpn', 'License Plate', 'LOTE'],
  descripcion: ['Descripcion', 'Description', 'Descripción', 'Producto', 'DESCRIPCION'],
  localizacion: ['Localizacion', 'Location', 'Ubicación', 'Posición', 'LOCATION'],
  disponible: ['Disponible', 'Available', 'Cantidad', 'Unidades', 'Stock', 'DISPO', 'STOCK'],
  estado: ['Estado', 'Status', 'Condición', 'Situación', 'ESTADO'],
  fechaVencimiento: [
    'Fecha de vencimiento',
    'Expiration',
    'fecha caducidad',
    'Expiry Date',
    'FECHA_VENC',
  ],
  diasFPC: ['FPC', 'Days to Exp', 'DIAS FPC', 'Días FPC', 'Fech. Porc. Cons.', 'DIAS_RESTANTES'],
  cantidadMinima: ['Cantidad Minima', 'Min Level', 'Minimo', 'CANTIDAD_MINIMA'],
  cantidadMaxima: ['Cantidad Maxima', 'Max Level', 'Maximo', 'CANTIDAD_MAXIMA'],
};

interface LocationData {
  lpn: string;
  localizacion: string;
  disponible: number;
  fechaVencimiento?: string | null;
  diasFPC?: number | null;
}

interface FileData {
  name: string;
  data: string | ArrayBuffer | null;
  size: number;
  uploadDate: Date;
}

interface RestockSuggestion {
  sku: string;
  descripcion: string;
  cantidadActual: number;
  cantidadMinima: number;
  cantidadMaxima: number;
  cantidadARestockear: number;
  accion: string;
  localizacion: string;
  ubicacionesSugeridas: LocationData[];
  lpnDestino?: string | null;
  localizacionDestino?: string | null;
}

interface AnalysisResult {
  summary: string;
  suggestions: RestockSuggestion[];
  stats: {
    totalSkus: number;
    skusARestockear: number;
    porcentajeCobertura: number;
  };
}

@Component({
  selector: 'app-levels',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    TableModule,
    ButtonModule,
    InputTextModule,
    ProgressSpinnerModule,
    TooltipModule,
    TagModule,
    BreadcrumbModule,
    ToastModule,
  ],
  providers: [MessageService],
  templateUrl: './levels.html',
  styleUrl: './levels.css',
})
export class Levels {
  @ViewChild('dt1') dt1!: Table;

  inventoryFile = signal<FileData | null>(null);
  levelsFile = signal<FileData | null>(null);
  analysisResult = signal<AnalysisResult | null>(null);
  loading = signal(false);
  error = signal<string | null>(null);
  searchValue: string = '';
  isBrowser: boolean;
  isExpandedAll: boolean = false;
  readonly breadcrumbItems: MenuItem[] = [{ label: 'Niveles', routerLink: '/levels' }];
  readonly breadcrumbHome: MenuItem = { icon: 'pi pi-home', label: 'Inicio', routerLink: '/' };

  inventoryHeaderMap: Record<string, string> = {};
  levelsHeaderMap: Record<string, string> = {};

  isDraggingInventory = signal(false);
  isDraggingLevels = signal(false);

  visibleColumns: Record<string, boolean> = {
    sku: true,
    localizacion: true,
    cantidadActual: true,
    cantidadMinima: true,
    cantidadMaxima: true,
    cantidadARestockear: true,
    ubicaciones: true,
  };

  get filteredSuggestions() {
    return (this.analysisResult()?.suggestions || []).filter((s) => s.cantidadARestockear > 0);
  }

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

  onDragOver(event: DragEvent, type: 'inventory' | 'levels') {
    event.preventDefault();
    event.stopPropagation();
    if (type === 'inventory') {
      this.isDraggingInventory.set(true);
    } else {
      this.isDraggingLevels.set(true);
    }
  }

  onDragLeave(event: DragEvent, type: 'inventory' | 'levels') {
    event.preventDefault();
    event.stopPropagation();
    if (type === 'inventory') {
      this.isDraggingInventory.set(false);
    } else {
      this.isDraggingLevels.set(false);
    }
  }

  onDrop(event: DragEvent, type: 'inventory' | 'levels') {
    event.preventDefault();
    event.stopPropagation();
    if (type === 'inventory') {
      this.isDraggingInventory.set(false);
    } else {
      this.isDraggingLevels.set(false);
    }
    const files = event.dataTransfer?.files;
    if (files && files.length > 0) {
      this.processFile(files[0], type);
    }
  }

  onFileUpload(type: 'inventory' | 'levels', event: Event) {
    if (!this.isBrowser) return;
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    this.processFile(file, type);
  }

  private processFile(file: File, type: 'inventory' | 'levels') {
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
      let rows: any[] = [];
      let headers: string[] = [];
      if (ext === 'xlsx' || ext === 'xls') {
        try {
          const workbook = XLSX.read(data, { type: 'binary' });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
          headers = (rows[0] || []).map((h: any) => String(h).trim());
        } catch (error) {
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

      let mapping: Record<string, string[]>;
      let requiredKeys: string[];
      if (type === 'inventory') {
        mapping = levelColumnMapping;
        requiredKeys = ['sku', 'localizacion', 'disponible'];
      } else {
        mapping = levelColumnMapping;
        requiredKeys = ['sku', 'localizacion', 'cantidadMinima', 'cantidadMaxima'];
      }

      const headersLower = headers.map((h) => h.toLowerCase().trim());
      const missing: string[] = [];
      for (const key of requiredKeys) {
        const aliases = mapping[key].map((a) => a.toLowerCase().trim());
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

      const fileData: FileData = {
        name: file.name,
        data: data,
        size: file.size,
        uploadDate: new Date(),
      };

      if (type === 'inventory') {
        this.inventoryFile.set(fileData);
        if (typeof data === 'string') {
          const rows = this.parseFileData(data, file.name);
          if (rows.length > 0) {
            const headers = Object.keys(rows[0]);
            this.inventoryHeaderMap = this.mapHeadersToStandard(headers, levelColumnMapping);
            this.messageService.add({
              severity: 'success',
              summary: 'Archivo cargado',
              detail: `Inventario: ${rows.length} registros encontrados`,
              life: 3000,
            });
          }
        }
      } else {
        this.levelsFile.set(fileData);
        if (typeof data === 'string') {
          const rows = this.parseFileData(data, file.name);
          if (rows.length > 0) {
            const headers = Object.keys(rows[0]);
            this.levelsHeaderMap = this.mapHeadersToStandard(headers, levelColumnMapping);
            this.messageService.add({
              severity: 'success',
              summary: 'Archivo cargado',
              detail: `Niveles: ${rows.length} registros encontrados`,
              life: 3000,
            });
          }
        }
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
    const ext = file.name.split('.').pop()?.toLowerCase();
    if (ext === 'xlsx' || ext === 'xls') {
      reader.readAsBinaryString(file);
    } else {
      reader.readAsText(file, 'UTF-8');
    }
  }

  removeFile(type: 'inventory' | 'levels') {
    if (type === 'inventory') {
      this.inventoryFile.set(null);
      this.inventoryHeaderMap = {};
    } else {
      this.levelsFile.set(null);
      this.levelsHeaderMap = {};
    }
    this.analysisResult.set(null);
    this.messageService.add({
      severity: 'info',
      summary: 'Archivo removido',
      detail: `Se ha eliminado el archivo de ${type === 'inventory' ? 'inventario' : 'niveles'}`,
      life: 3000,
    });
  }

  generateLevelsAnalysis() {
    if (!this.inventoryFile() || !this.levelsFile()) {
      this.messageService.add({
        severity: 'warn',
        summary: 'Archivos faltantes',
        detail: 'Debes subir ambos archivos (Inventario y Niveles) para analizar',
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
          summary: 'Análisis completado',
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
    const exportData: any[][] = [
      [
        'SKU',
        'Descripción',
        'Destino',
        'LPN Destino',
        'Cant. en Picking',
        'Cant. a Surtir',
        'Acción',
        'Ubicación Origen',
        'LPN Origen',
        'FPC',
        'Fecha Vencimiento',
        'Cantidad desde Ubicación',
        'Cantidad Mínima',
        'Cantidad Máxima',
      ],
    ];

    for (const s of this.filteredSuggestions) {
      if (!s.ubicacionesSugeridas.length) {
        exportData.push([
          s.sku,
          s.descripcion,
          s.localizacionDestino || s.localizacion,
          s.lpnDestino || '',
          s.cantidadActual,
          s.cantidadARestockear,
          s.accion,
          '',
          '',
          '',
          '',
          0,
          s.cantidadMinima,
          s.cantidadMaxima,
        ]);
        continue;
      }

      s.ubicacionesSugeridas.forEach((u, idx) => {
        exportData.push([
          s.sku,
          s.descripcion,
          s.localizacionDestino || s.localizacion,
          s.lpnDestino || '',
          idx === 0 ? s.cantidadActual : '',
          idx === 0 ? s.cantidadARestockear : '',
          idx === 0 ? s.accion : '',
          u.localizacion,
          u.lpn || '',
          u.diasFPC ?? '',
          u.fechaVencimiento || '',
          u.disponible,
          idx === 0 ? s.cantidadMinima : '',
          idx === 0 ? s.cantidadMaxima : '',
        ]);
      });
    }

    const ws = XLSX.utils.aoa_to_sheet(exportData);
    ws['!cols'] = [14, 42, 20, 16, 15, 14, 14, 18, 16, 8, 14, 20, 14, 14].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws, 'Niveles');
    const fecha = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `analisis_niveles_${fecha}.xlsx`);
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

  getTotalSkus(): number {
    return this.analysisResult()?.stats?.totalSkus || 0;
  }

  getSkusToRestock(): number {
    return this.filteredSuggestions.length;
  }

  getTotalUnitsToRestock(): number {
    return (
      this.analysisResult()?.suggestions?.reduce((sum, s) => sum + s.cantidadARestockear, 0) || 0
    );
  }

  getCoberturaPorcentaje(): number {
    return this.analysisResult()?.stats?.porcentajeCobertura || 0;
  }

  getDestinoTexto(suggestion: RestockSuggestion): string {
    return suggestion.localizacionDestino || suggestion.localizacion || '-';
  }

  formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
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
    if (lines[0].includes(';')) {
      delimiter = ';';
    }
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
        const jsonData = XLSX.utils.sheet_to_json(sheet);
        return jsonData;
      } catch (error) {
        console.error('Error parsing Excel:', error);
        return [];
      }
    }
    return typeof data === 'string' ? this.parseCSV(data) : [];
  }

  private sortByFefo(a: LocationData, b: LocationData): number {
    const aFpc = a.diasFPC;
    const bFpc = b.diasFPC;
    if (aFpc != null && bFpc != null && aFpc !== bFpc) {
      return aFpc - bFpc;
    }
    if (aFpc != null && bFpc == null) return -1;
    if (aFpc == null && bFpc != null) return 1;
    const aDate = a.fechaVencimiento ? new Date(a.fechaVencimiento).getTime() : null;
    const bDate = b.fechaVencimiento ? new Date(b.fechaVencimiento).getTime() : null;
    if (aDate && bDate && aDate !== bDate) return aDate - bDate;
    if (aDate && !bDate) return -1;
    if (!aDate && bDate) return 1;
    return 0;
  }

  private getNestedValue(obj: any, path: string, defaultValue: any = ''): any {
    if (!obj || !path) return defaultValue;
    const value = obj[path];
    if (value === undefined || value === null || value === '') return defaultValue;
    return value;
  }

  private processAnalysis(): AnalysisResult {
    const inventoryRaw = this.inventoryFile()?.data;
    const levelsRaw = this.levelsFile()?.data;
    const inventoryName = this.inventoryFile()?.name || '';
    const levelsName = this.levelsFile()?.name || '';

    if (typeof inventoryRaw !== 'string' || typeof levelsRaw !== 'string') {
      throw new Error('Archivos no válidos');
    }

    const inventoryRows = this.parseFileData(inventoryRaw, inventoryName);
    const levelsRows = this.parseFileData(levelsRaw, levelsName);
    const invMap = this.inventoryHeaderMap;
    const levelsMap = this.levelsHeaderMap;

    // 1. Filtrar inventario válido
    const freeStockInventory = inventoryRows.filter((item) => {
      const estadoRaw = this.getNestedValue(item, invMap['estado'] || 'Estado', null);
      const estado = String(estadoRaw ?? '').toUpperCase();
      const loc = String(
        this.getNestedValue(item, invMap['localizacion'] || 'Localizacion', ''),
      ).toUpperCase();
      const isIgnoredLoc = IGNORED_LOCATIONS.some((iloc) => loc.includes(iloc.toUpperCase()));
      const estadoOk =
        estado === '' || VALID_STATUSES.some((s) => estado.includes(s.toUpperCase()));
      return estadoOk && !isIgnoredLoc;
    });

    // 2. Agrupar inventario por SKU y ubicación
    const inventoryBySku: Record<string, any> = {};
    for (const item of freeStockInventory) {
      const sku = String(this.getNestedValue(item, invMap['sku'] || 'SKU', ''))
        .toUpperCase()
        .trim();
      if (!sku) continue;
      if (!inventoryBySku[sku]) {
        inventoryBySku[sku] = {
          descripcion: this.getNestedValue(item, invMap['descripcion'] || 'Descripcion', ''),
          totalEnReserva: 0,
          reserveLocations: [],
        };
      }
      const loc = String(this.getNestedValue(item, invMap['localizacion'] || 'Localizacion', ''));
      const lastPart = loc.split('-').pop() || '';
      const isReserveByLevel = RESERVE_LEVELS.includes(lastPart);
      const isReserveByAdditional = ADDITIONAL_RESERVE_LOCATIONS.some((prefix) =>
        loc.toUpperCase().startsWith(prefix),
      );
      if (isReserveByLevel || isReserveByAdditional) {
        inventoryBySku[sku].totalEnReserva +=
          parseFloat(this.getNestedValue(item, invMap['disponible'] || 'Disponible', '0')) || 0;
        inventoryBySku[sku].reserveLocations.push({
          lpn: this.getNestedValue(item, invMap['lpn'] || 'LPN', ''),
          localizacion: loc,
          disponible:
            parseFloat(this.getNestedValue(item, invMap['disponible'] || 'Disponible', '0')) || 0,
          fechaVencimiento: this.getNestedValue(
            item,
            invMap['fechaVencimiento'] || 'Fecha de vencimiento',
            null,
          ),
          diasFPC: this.getNestedValue(item, invMap['diasFPC'] || 'FPC', null),
        });
      }
    }

    // 3. Procesar reglas de niveles minmax
    const suggestions: RestockSuggestion[] = [];
    const inventoryMap = new Map<string, number>();
    for (const item of freeStockInventory) {
      const sku = String(this.getNestedValue(item, invMap['sku'] || 'SKU', ''))
        .toUpperCase()
        .trim();
      const loc = String(this.getNestedValue(item, invMap['localizacion'] || 'Localizacion', ''));
      const key = `${sku}__${loc}`;
      if (!inventoryMap.has(key)) inventoryMap.set(key, 0);
      inventoryMap.set(
        key,
        inventoryMap.get(key)! +
          (parseFloat(this.getNestedValue(item, invMap['disponible'] || 'Disponible', '0')) || 0),
      );
    }

    for (const rule of levelsRows) {
      const sku = String(this.getNestedValue(rule, levelsMap['sku'] || 'SKU', ''))
        .toUpperCase()
        .trim();
      const localizacion = String(
        this.getNestedValue(rule, levelsMap['localizacion'] || 'Localizacion', ''),
      );
      const cantidadMinima =
        parseFloat(
          this.getNestedValue(rule, levelsMap['cantidadMinima'] || 'Cantidad Minima', '0'),
        ) || 0;
      const cantidadMaxima =
        parseFloat(
          this.getNestedValue(rule, levelsMap['cantidadMaxima'] || 'Cantidad Maxima', '0'),
        ) || 0;
      const lpn = this.getNestedValue(rule, levelsMap['lpn'] || 'LPN', '');

      const key = `${sku}__${localizacion}`;
      const currentStock = inventoryMap.get(key) || 0;
      const generalSkuInventory = inventoryBySku[sku];

      if (
        currentStock < cantidadMinima &&
        generalSkuInventory &&
        generalSkuInventory.totalEnReserva > 0
      ) {
        const amountToRestock = cantidadMaxima - currentStock;
        const sortedReserveLocations = [...generalSkuInventory.reserveLocations].sort((a, b) =>
          this.sortByFefo(a, b),
        );

        let cantidadARestockear = 0;
        const ubicacionesSugeridas: LocationData[] = [];
        let needed = amountToRestock;

        for (const location of sortedReserveLocations) {
          if (needed <= 0) break;
          const amountToTake = Math.min(location.disponible, needed);
          cantidadARestockear += amountToTake;
          ubicacionesSugeridas.push({
            lpn: location.lpn,
            localizacion: location.localizacion,
            disponible: amountToTake,
            fechaVencimiento: location.fechaVencimiento,
            diasFPC: location.diasFPC,
          });
          needed -= amountToTake;
        }

        suggestions.push({
          sku,
          descripcion: generalSkuInventory.descripcion,
          cantidadActual: currentStock,
          cantidadMinima,
          cantidadMaxima,
          cantidadARestockear,
          accion: 'Reabastecer',
          localizacion,
          ubicacionesSugeridas,
          lpnDestino: lpn,
          localizacionDestino: localizacion,
        });
      }
    }

    suggestions.sort((a, b) => {
      if (a.cantidadARestockear > 0 && b.cantidadARestockear === 0) return -1;
      if (a.cantidadARestockear === 0 && b.cantidadARestockear > 0) return 1;
      return a.sku.localeCompare(b.sku);
    });

    const skusARestockear = suggestions.filter((s) => s.cantidadARestockear > 0).length;
    const summary = `✅ Análisis completado: ${suggestions.length} SKUs analizados, ${skusARestockear} requieren resurtido.`;
    const totalNecesario = suggestions.reduce(
      (sum, s) => sum + (s.cantidadMaxima - s.cantidadActual),
      0,
    );
    const totalCubierto = suggestions.reduce((sum, s) => sum + s.cantidadARestockear, 0);
    const porcentajeCobertura =
      totalNecesario === 0 ? 100 : Math.round((totalCubierto / totalNecesario) * 100);

    return {
      summary,
      suggestions,
      stats: {
        totalSkus: suggestions.length,
        skusARestockear,
        porcentajeCobertura,
      },
    };
  }
}
