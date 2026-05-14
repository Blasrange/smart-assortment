import { Component, Inject, PLATFORM_ID, ViewChild, signal } from '@angular/core';
import { CommonModule, isPlatformBrowser } from '@angular/common';
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
const IGNORED_LOCATION_KEYWORDS = ['PDIF-INV-1-10', 'DEV-1-10', 'PDIF-RES-1-10', 'DEV-RES-1-10'];
const PICKING_PREFIXES = ['PT'];
const RESERVE_LEVELS = ['20', '30', '40', '50', '60', '70'];

export const smartAssortmentSalesMapping: Record<string, string[]> = {
  sku: ['MAI_SKU', 'SKU', 'Material', 'Codigo', 'CODIGO'],
  description: ['MAI_DESCRIPTION', 'Descripcion', 'Descripción', 'Producto'],
  qtyOrder: ['QTY_ORDER', 'cantidad confirmada', 'Cantidad', 'Cant. Facturada', 'Qty'],
};

export const smartAssortmentStockMapping: Record<string, string[]> = {
  Codigo: ['Codigo', 'codigo'],
  LPN: ['LPN', 'lpn'],
  Localizacion: ['Localizacion', 'localizacion'],
  'Area Picking': ['Area Picking', 'Área Picking', 'areaPicking', 'AREA PICKING', 'PICKING'],
  SKU: ['SKU', 'sku', 'Codigo', 'codigo', 'Material'],
  Descripcion: ['Descripcion', 'descripcion', 'Descripción'],
  Disponible: ['Disponible', 'disponible'],
  Estado: ['Estado', 'estado'],
  Lote: ['Lote', 'lote'],
  'Fecha de entrada': ['Fecha de entrada', 'fechaEntrada'],
  'Fecha de vencimiento': ['Fecha de vencimiento', 'fechaVencimiento'],
};

export const smartAssortmentMaterialMasterMapping: Record<string, string[]> = {
  lpn: ['LPN', 'lpn'],
  localizacion: [
    'Localizacion',
    'Ubicacion',
    'Ubicación',
    'Ubicacion destino',
    'Ubicación destino',
    'Localizacion destino',
  ],
  sku: ['SKU', 'sku', 'Codigo', 'Material', 'MAI_SKU', 'Cod Material'],
  descripcion: ['Descripcion', 'Descripción', 'Description'],
  tipoMaterial: ['Tipo de Material', 'TipoMaterial', 'tipoMaterial'],
  embalaje: ['Embalaje', 'embalaje', 'Unidad de Embalaje'],
};

interface FileData {
  name: string;
  data: string | ArrayBuffer | null;
  size: number;
  uploadDate: Date;
}

interface UbicacionSugerida {
  lpn: string;
  localizacion: string;
  lote: string;
  fechaVencimiento: string;
  cantidad: number;
  esEstibaCompleta: boolean;
}

interface RestockSuggestion {
  sku: string;
  descripcion: string;
  cantidadVendida: number;
  cantidadVendidaOriginal: number;
  cantidadDisponiblePicking: number;
  cantidadEnReserva: number;
  cantidadARestockear: number;
  cantidadFaltante: number;
  ubicacionesSugeridas: UbicacionSugerida[];
  lpnDestino: string | null;
  localizacionDestino: string | null;
  proyeccion: boolean;
  proyeccionDias: number;
  cantidadProyectada: number;
  prioridadAlta: boolean;
  existeEnMaestra: boolean;
  estadoSurtido: string;
}

interface MissingProduct {
  sku: string;
  descripcion: string;
  cantidadVendida: number;
  cantidadFaltante: number;
  stockEnPicking: number;
  stockEnReserva: number;
  cantidadCubierta: number;
  tipoFalta: 'SIN_INVENTARIO' | 'SIN_RESERVA' | 'RESERVA_INSUFICIENTE';
}

interface AnalysisStats {
  totalProductos: number;
  productosConSurtido: number;
  productosFaltantes: number;
  totalUnidadesSurtir: number;
  porcentajeCobertura: number;
  productosConProyeccion: number;
}

interface AnalysisResult {
  summary: string;
  suggestions: RestockSuggestion[];
  missingProducts: MissingProduct[];
  stats: AnalysisStats;
}

interface InventoryRow {
  SKU: string;
  LPN: string;
  Localizacion: string;
  AreaPicking: string;
  Descripcion: string;
  Estado: string;
  Disponible: number;
  Lote: string;
  fechaVencimiento: string;
  fechaEntrada: string;
}

@Component({
  selector: 'app-assortment-sales',
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
  templateUrl: './assortment-sales.html',
  styleUrls: ['./assortment-sales.css'],
})
export class AssortmentSales {
  @ViewChild('dt1') dt1!: Table;
  @ViewChild('dt2') dt2!: Table;

  salesFile = signal<FileData | null>(null);
  stockFile = signal<FileData | null>(null);
  masterFile = signal<FileData | null>(null);
  analysisResult = signal<AnalysisResult | null>(null);
  loading = signal(false);
  error = signal<string | null>(null);

  isDraggingSales = signal(false);
  isDraggingStock = signal(false);
  isDraggingMaster = signal(false);

  searchValue = '';
  activeTab: 'missing' | 'suggestions' = 'suggestions';
  isBrowser: boolean;

  salesHeaderMap: Record<string, string> = {};
  stockHeaderMap: Record<string, string> = {};
  masterHeaderMap: Record<string, string> = {};

  readonly breadcrumbItems: MenuItem[] = [
    { label: 'Surtido Inteligente', routerLink: '/assortment-sales' },
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

  get filteredMissingProducts(): MissingProduct[] {
    return (this.analysisResult()?.missingProducts || []).filter((m) => m.cantidadFaltante > 0);
  }

  get filteredRestockSuggestions(): RestockSuggestion[] {
    return (this.analysisResult()?.suggestions || []).filter((s) => s.cantidadARestockear > 0);
  }

  clearFilters() {
    if (this.dt1) {
      this.dt1.clear();
      this.dt1.filterGlobal('', 'contains');
    }
    if (this.dt2) {
      this.dt2.clear();
      this.dt2.filterGlobal('', 'contains');
    }
    this.searchValue = '';
    this.messageService.add({
      severity: 'info',
      summary: 'Filtros limpiados',
      detail: 'Se han eliminado todos los filtros',
      life: 2500,
    });
  }

  onDragOver(event: DragEvent, type: 'sales' | 'stock' | 'master') {
    event.preventDefault();
    event.stopPropagation();
    if (type === 'sales') this.isDraggingSales.set(true);
    if (type === 'stock') this.isDraggingStock.set(true);
    if (type === 'master') this.isDraggingMaster.set(true);
  }

  onDragLeave(event: DragEvent, type: 'sales' | 'stock' | 'master') {
    event.preventDefault();
    event.stopPropagation();
    if (type === 'sales') this.isDraggingSales.set(false);
    if (type === 'stock') this.isDraggingStock.set(false);
    if (type === 'master') this.isDraggingMaster.set(false);
  }

  onDrop(event: DragEvent, type: 'sales' | 'stock' | 'master') {
    event.preventDefault();
    event.stopPropagation();
    if (type === 'sales') this.isDraggingSales.set(false);
    if (type === 'stock') this.isDraggingStock.set(false);
    if (type === 'master') this.isDraggingMaster.set(false);

    const files = event.dataTransfer?.files;
    if (files && files.length > 0) {
      this.processFile(files[0], type);
    }
  }

  onFileUpload(type: 'sales' | 'stock' | 'master', event: Event) {
    if (!this.isBrowser) return;
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    this.processFile(file, type);
  }

  removeFile(type: 'sales' | 'stock' | 'master') {
    if (type === 'sales') {
      this.salesFile.set(null);
      this.salesHeaderMap = {};
    }
    if (type === 'stock') {
      this.stockFile.set(null);
      this.stockHeaderMap = {};
    }
    if (type === 'master') {
      this.masterFile.set(null);
      this.masterHeaderMap = {};
    }

    this.analysisResult.set(null);
    this.messageService.add({
      severity: 'info',
      summary: 'Archivo removido',
      detail: 'Archivo eliminado correctamente',
      life: 2500,
    });
  }

  generateAssortmentAnalysis() {
    if (!this.salesFile() || !this.stockFile()) {
      this.messageService.add({
        severity: 'warn',
        summary: 'Archivos faltantes',
        detail: 'Debes subir Ventas y Stock para iniciar el análisis',
        life: 4500,
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

    const resumenData = [
      { Concepto: 'SURTIDO INTELIGENTE', Valor: '', Unidad: '' },
      { Concepto: '  - SKUs analizados', Valor: result.stats.totalProductos, Unidad: 'SKU' },
      { Concepto: '  - SKUs a surtir', Valor: result.stats.productosConSurtido, Unidad: 'SKU' },
      { Concepto: '  - SKUs faltantes', Valor: result.stats.productosFaltantes, Unidad: 'SKU' },
      {
        Concepto: '  - Unidades a surtir',
        Valor: result.stats.totalUnidadesSurtir,
        Unidad: 'Unidades',
      },
      { Concepto: '  - Cobertura', Valor: `${result.stats.porcentajeCobertura}%`, Unidad: '' },
      {
        Concepto: '  - SKUs con proyección',
        Valor: result.stats.productosConProyeccion,
        Unidad: 'SKU',
      },
    ];

    const wsResumen = XLSX.utils.json_to_sheet(resumenData);
    wsResumen['!cols'] = [42, 20, 14].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsResumen, 'Resumen');

    const surtidoRows = this.filteredRestockSuggestions.flatMap((s) => {
      if (!s.ubicacionesSugeridas.length) {
        return [
          {
            SKU: s.sku,
            Descripcion: s.descripcion,
            Vendidas_Analizadas: s.cantidadVendida,
            Picking: s.cantidadDisponiblePicking,
            Reserva: s.cantidadEnReserva,
            A_Surtir: s.cantidadARestockear,
            Faltante: s.cantidadFaltante,
            Ubicacion_Origen: '',
            LPN_Origen: '',
            Cantidad_Origen: 0,
            Destino: s.localizacionDestino || '',
            Proyeccion: s.proyeccion ? `SI (${s.proyeccionDias}d)` : 'NO',
          },
        ];
      }

      return s.ubicacionesSugeridas.map((u, index) => ({
        SKU: s.sku,
        Descripcion: s.descripcion,
        Vendidas_Analizadas: index === 0 ? s.cantidadVendida : '',
        Picking: index === 0 ? s.cantidadDisponiblePicking : '',
        Reserva: index === 0 ? s.cantidadEnReserva : '',
        A_Surtir: index === 0 ? s.cantidadARestockear : '',
        Faltante: index === 0 ? s.cantidadFaltante : '',
        Ubicacion_Origen: u.localizacion,
        LPN_Origen: u.lpn,
        Cantidad_Origen: u.cantidad,
        Destino: index === 0 ? s.localizacionDestino || '' : '',
        Proyeccion: index === 0 ? (s.proyeccion ? `SI (${s.proyeccionDias}d)` : 'NO') : '',
      }));
    });

    const wsSurtido = XLSX.utils.json_to_sheet(surtidoRows);
    wsSurtido['!cols'] = [14, 34, 18, 10, 10, 10, 10, 20, 16, 14, 20, 16].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsSurtido, 'Surtido');

    const faltantesRows = this.filteredMissingProducts.map((m) => ({
      SKU: m.sku,
      Descripcion: m.descripcion,
      Vendidas: m.cantidadVendida,
      Stock_Picking: m.stockEnPicking,
      Stock_Reserva: m.stockEnReserva,
      Cubierta: m.cantidadCubierta,
      Faltante: m.cantidadFaltante,
      Estado: m.tipoFalta,
    }));

    const wsFaltantes = XLSX.utils.json_to_sheet(faltantesRows);
    wsFaltantes['!cols'] = [14, 34, 12, 14, 14, 12, 12, 22].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsFaltantes, 'Faltantes');

    const fecha = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `surtido_inteligente_${fecha}.xlsx`);

    this.messageService.add({
      severity: 'success',
      summary: 'Exportado',
      detail: 'Archivo Excel generado con Resumen, Surtido y Faltantes',
      life: 3000,
    });
  }

  getTipoFaltaSeverity(
    tipoFalta: string,
  ): 'success' | 'info' | 'warn' | 'danger' | 'secondary' | 'contrast' {
    const severities: Record<
      string,
      'success' | 'info' | 'warn' | 'danger' | 'secondary' | 'contrast'
    > = {
      SIN_INVENTARIO: 'danger',
      SIN_RESERVA: 'danger',
      RESERVA_INSUFICIENTE: 'warn',
    };
    return severities[tipoFalta] || 'info';
  }

  getTipoFaltaLabel(tipoFalta: string): string {
    const labels: Record<string, string> = {
      SIN_INVENTARIO: 'Sin inventario',
      SIN_RESERVA: 'Sin reserva',
      RESERVA_INSUFICIENTE: 'Reserva insuficiente',
    };
    return labels[tipoFalta] || tipoFalta;
  }

  getStatusSeverity(value: number): 'success' | 'warn' {
    return value > 0 ? 'warn' : 'success';
  }

  formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  private processFile(file: File, type: 'sales' | 'stock' | 'master') {
    if (!this.isBrowser) return;

    const validExtensions = ['.csv', '.xlsx', '.xls', '.txt'];
    const fileExt = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    if (!validExtensions.includes(fileExt)) {
      this.messageService.add({
        severity: 'error',
        summary: 'Formato no soportado',
        detail: 'Use CSV, Excel o TXT',
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

      if ((ext === 'xlsx' || ext === 'xls') && typeof data === 'string') {
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
        type === 'sales'
          ? smartAssortmentSalesMapping
          : type === 'stock'
            ? smartAssortmentStockMapping
            : smartAssortmentMaterialMasterMapping;

      const requiredKeys: string[] =
        type === 'sales'
          ? ['sku', 'qtyOrder']
          : type === 'stock'
            ? ['SKU', 'Localizacion', 'Estado', 'Disponible']
            : ['sku', 'localizacion'];

      const headersLower = headers.map((h) => h.toLowerCase().trim());
      const missing: string[] = [];

      for (const key of requiredKeys) {
        const aliases = (mapping[key] || []).map((a) => a.toLowerCase().trim());
        const found = headersLower.some((h) => aliases.includes(h));
        if (!found && mapping[key]?.[0]) {
          missing.push(mapping[key][0]);
        }
      }

      if (missing.length > 0) {
        this.messageService.add({
          severity: 'error',
          summary: 'Columnas faltantes',
          detail: `El archivo no contiene columnas requeridas: ${missing.join(', ')}`,
          life: 7000,
        });
        return;
      }

      const rows = this.parseFileData(data as string | ArrayBuffer, file.name);
      const fileHeaders = rows.length ? Object.keys(rows[0]) : headers;
      const headerMap = this.mapHeadersToStandard(fileHeaders, mapping);

      const fileData: FileData = {
        name: file.name,
        data,
        size: file.size,
        uploadDate: new Date(),
      };

      if (type === 'sales') {
        this.salesFile.set(fileData);
        this.salesHeaderMap = headerMap;
      } else if (type === 'stock') {
        this.stockFile.set(fileData);
        this.stockHeaderMap = headerMap;
      } else {
        this.masterFile.set(fileData);
        this.masterHeaderMap = headerMap;
      }

      this.messageService.add({
        severity: 'success',
        summary: 'Archivo cargado',
        detail: `${type === 'sales' ? 'Ventas' : type === 'stock' ? 'Stock' : 'Maestra'}: ${rows.length} registros`,
        life: 2800,
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

  private processAnalysis(): AnalysisResult {
    const salesRaw = this.salesFile()?.data;
    const stockRaw = this.stockFile()?.data;
    const masterRaw = this.masterFile()?.data;

    if (typeof salesRaw !== 'string' || typeof stockRaw !== 'string') {
      throw new Error('Archivos de ventas o stock inválidos');
    }

    const ventasRows = this.parseFileData(salesRaw, this.salesFile()?.name || 'ventas.csv');
    const stockRows = this.parseFileData(stockRaw, this.stockFile()?.name || 'stock.csv');
    const maestraRows =
      typeof masterRaw === 'string'
        ? this.parseFileData(masterRaw, this.masterFile()?.name || 'maestra.csv')
        : [];

    const salesMap = this.salesHeaderMap;
    const stockMap = this.stockHeaderMap;
    const masterMap = this.masterHeaderMap;

    const norm = (value: any) =>
      String(value || '')
        .trim()
        .toUpperCase();
    const skuKey = (value: any) => this.normalizeSkuKey(value);
    const toNum = (value: any) => this.toNumber(value);

    const ventasNormalizadas = ventasRows
      .map((v: any) => {
        const skuRaw = this.getNestedValue(v, salesMap['sku'] || 'MAI_SKU', '');
        const sku = skuKey(skuRaw);
        const skuDisplay = this.normalizeSkuDisplay(skuRaw);
        const descripcion = norm(
          this.getNestedValue(v, salesMap['description'] || 'MAI_DESCRIPTION', ''),
        );
        const cantidad = toNum(this.getNestedValue(v, salesMap['qtyOrder'] || 'QTY_ORDER', 0));
        return { sku, skuDisplay, descripcion, cantidad };
      })
      .filter((v) => v.sku && v.cantidad > 0);

    if (!ventasNormalizadas.length) {
      throw new Error('No hay ventas válidas después de la normalización');
    }

    const maestraMap = new Map<
      string,
      { lpn: string; localizacion: string; descripcion: string; tipoMaterial: string }
    >();
    for (const m of maestraRows) {
      const sku = skuKey(this.getNestedValue(m, masterMap['sku'] || 'SKU', ''));
      if (!sku || maestraMap.has(sku)) continue;
      maestraMap.set(sku, {
        lpn: String(this.getNestedValue(m, masterMap['lpn'] || 'LPN', '')).trim(),
        localizacion: String(
          this.getNestedValue(m, masterMap['localizacion'] || 'Localizacion', ''),
        ).trim(),
        descripcion: String(
          this.getNestedValue(m, masterMap['descripcion'] || 'Descripcion', ''),
        ).trim(),
        tipoMaterial: String(
          this.getNestedValue(m, masterMap['tipoMaterial'] || 'Tipo de Material', ''),
        ).trim(),
      });
    }

    const validStatuses = VALID_STATUSES.map((s) => s.toUpperCase());
    const debeOmitirsePorKeyword = (ubicacion: string) =>
      IGNORED_LOCATION_KEYWORDS.some((keyword) =>
        ubicacion.toUpperCase().includes(keyword.toUpperCase()),
      );
    const normalizeText = (value: any) =>
      String(value ?? '')
        .trim()
        .toUpperCase();
    const normalizeLocationKey = (value: string) => normalizeText(value).replace(/\s+/g, '');
    const isMarkedAsPicking = (value: string): boolean => {
      const area = normalizeText(value);
      if (!area) return false;
      const exactFlags = ['SI', 'SÍ', 'YES', 'Y', 'TRUE', '1', 'X', 'PICKING', 'PICK'];
      return exactFlags.includes(area) || area.includes('PICK');
    };
    const getLastSegment = (ubicacion: string): string => ubicacion.split('-').pop() || '';
    const isPickingLocation = (ubicacion: string) =>
      PICKING_PREFIXES.some((prefix) => ubicacion.toUpperCase().startsWith(prefix.toUpperCase()));
    const isValidReserveLocation = (ubicacion: string) => {
      const lastSegment = getLastSegment(ubicacion);
      if (lastSegment === '10' || lastSegment.startsWith('10-')) return false;
      return RESERVE_LEVELS.some((level) => lastSegment === level);
    };

    const stockNormalizado: InventoryRow[] = stockRows
      .map((item: any) => ({
        SKU: skuKey(this.getNestedValue(item, stockMap['SKU'] || 'SKU', '')),
        LPN: String(this.getNestedValue(item, stockMap['LPN'] || 'LPN', '')).trim(),
        Localizacion: String(
          this.getNestedValue(item, stockMap['Localizacion'] || 'Localizacion', ''),
        ).trim(),
        AreaPicking: String(
          this.getNestedValue(item, stockMap['Area Picking'] || 'Area Picking', ''),
        ).trim(),
        Descripcion: String(
          this.getNestedValue(item, stockMap['Descripcion'] || 'Descripcion', ''),
        ).trim(),
        Estado: String(this.getNestedValue(item, stockMap['Estado'] || 'Estado', '')).trim(),
        Disponible: toNum(this.getNestedValue(item, stockMap['Disponible'] || 'Disponible', 0)),
        Lote: String(this.getNestedValue(item, stockMap['Lote'] || 'Lote', '')).trim(),
        fechaVencimiento: String(
          this.getNestedValue(item, stockMap['Fecha de vencimiento'] || 'Fecha de vencimiento', ''),
        ).trim(),
        fechaEntrada: String(
          this.getNestedValue(item, stockMap['Fecha de entrada'] || 'Fecha de entrada', ''),
        ).trim(),
      }))
      .filter((item) => {
        const estadoUpper = item.Estado.toUpperCase();
        const esEstadoValido = validStatuses.some((vs) => estadoUpper.includes(vs));
        const esUbicacionIgnorada = debeOmitirsePorKeyword(item.Localizacion || '');
        return item.SKU && esEstadoValido && !esUbicacionIgnorada && item.Disponible > 0;
      });

    const ventasPorSku = new Map<
      string,
      {
        total: number;
        descripcion: string;
        skuDisplay: string;
        tendencia: 'alta' | 'media' | 'baja';
      }
    >();
    for (const venta of ventasNormalizadas) {
      const current = ventasPorSku.get(venta.sku);
      if (current) {
        current.total += venta.cantidad;
        if (!current.skuDisplay || current.skuDisplay === venta.sku) {
          current.skuDisplay = venta.skuDisplay || current.skuDisplay;
        }
      } else {
        ventasPorSku.set(venta.sku, {
          total: venta.cantidad,
          descripcion: venta.descripcion,
          skuDisplay: venta.skuDisplay || venta.sku,
          tendencia: 'media',
        });
      }
    }

    const stockPorSku = new Map<string, InventoryRow[]>();
    for (const item of stockNormalizado) {
      if (!stockPorSku.has(item.SKU)) stockPorSku.set(item.SKU, []);
      stockPorSku.get(item.SKU)!.push(item);
    }

    const ventasArray = Array.from(ventasPorSku.entries())
      .map(([sku, info]) => ({ sku, total: info.total }))
      .sort((a, b) => b.total - a.total);
    const topCount = Math.max(1, Math.ceil(ventasArray.length * 0.2));
    const skusAltaRotacion = new Set(ventasArray.slice(0, topCount).map((i) => i.sku));
    const promedioVentas =
      ventasArray.reduce((sum, i) => sum + i.total, 0) / Math.max(1, ventasArray.length);

    const skusTendenciaAlza = new Set<string>();
    for (const [sku, info] of ventasPorSku) {
      if (info.total > promedioVentas * 1.5) {
        skusTendenciaAlza.add(sku);
        info.tendencia = 'alta';
      } else if (info.total < promedioVentas * 0.5) {
        info.tendencia = 'baja';
      }
    }

    const getReserveLocations = (items: InventoryRow[]) =>
      items.filter((item) => isValidReserveLocation(item.Localizacion) && item.Disponible > 0);

    const getItemKey = (item: InventoryRow) =>
      `${item.SKU}|${item.LPN}|${item.Localizacion}|${item.Lote}|${item.Disponible}`;

    const suggestions: RestockSuggestion[] = [];
    const missingProducts: MissingProduct[] = [];

    for (const [sku, ventaInfo] of ventasPorSku) {
      const stockItems = stockPorSku.get(sku) || [];
      const maestraData = maestraMap.get(sku);
      const existeEnMaestra = maestraMap.has(sku);
      const destinoValido = !!(
        maestraData?.localizacion && maestraData.localizacion.trim().length > 0
      );

      const cantidadVendidaOriginal = ventaInfo.total;
      let cantidadVendida = cantidadVendidaOriginal;
      let esProyeccion = false;
      let proyeccionDias = 0;
      let cantidadProyectada = cantidadVendidaOriginal;

      if (skusAltaRotacion.has(sku)) {
        proyeccionDias = 2;
        cantidadProyectada = Math.ceil(cantidadVendidaOriginal * 3);
        cantidadVendida = cantidadProyectada;
        esProyeccion = true;
      } else if (skusTendenciaAlza.has(sku)) {
        proyeccionDias = 1;
        cantidadProyectada = Math.ceil(cantidadVendidaOriginal * 2);
        cantidadVendida = cantidadProyectada;
        esProyeccion = true;
      }

      const destinoMaestra = maestraData?.localizacion || '';
      const lpnDestinoMaestra = maestraData?.lpn || '';

      const pickingItems = stockItems.filter((i) => {
        const byPrefix = isPickingLocation(i.Localizacion);
        const byAreaFlag = isMarkedAsPicking(i.AreaPicking);
        const byDestinoUbicacion =
          !!destinoMaestra &&
          normalizeLocationKey(i.Localizacion) === normalizeLocationKey(destinoMaestra);
        const byDestinoLpn =
          !!lpnDestinoMaestra && normalizeText(i.LPN) === normalizeText(lpnDestinoMaestra);
        return byPrefix || byAreaFlag || byDestinoUbicacion || byDestinoLpn;
      });

      const cantidadDisponiblePicking = pickingItems.reduce((sum, i) => sum + i.Disponible, 0);
      const pickingKeys = new Set(pickingItems.map((item) => getItemKey(item)));
      const reservaItemsValidos = getReserveLocations(stockItems).filter(
        (item) => !pickingKeys.has(getItemKey(item)),
      );
      const cantidadEnReserva = reservaItemsValidos.reduce((sum, i) => sum + i.Disponible, 0);

      let cantidadNecesaria = Math.max(0, cantidadVendida - cantidadDisponiblePicking);
      const ubicacionesSugeridas: UbicacionSugerida[] = [];

      if (cantidadNecesaria > 0 && reservaItemsValidos.length > 0 && destinoValido) {
        const reservaOrdenada = [...reservaItemsValidos].sort((a, b) => {
          const timeVencA = new Date(a.fechaVencimiento || '9999-12-31').getTime();
          const timeVencB = new Date(b.fechaVencimiento || '9999-12-31').getTime();
          if (timeVencA !== timeVencB) return timeVencA - timeVencB;

          const timeEntA = new Date(a.fechaEntrada || '9999-12-31').getTime();
          const timeEntB = new Date(b.fechaEntrada || '9999-12-31').getTime();
          if (timeEntA !== timeEntB) return timeEntA - timeEntB;

          return (a.Lote || '').localeCompare(b.Lote || '');
        });

        let restante = cantidadNecesaria;
        for (const item of reservaOrdenada) {
          if (restante <= 0) break;
          const tomar = Math.min(item.Disponible, restante);
          ubicacionesSugeridas.push({
            lpn: item.LPN,
            localizacion: item.Localizacion,
            lote: item.Lote,
            fechaVencimiento: item.fechaVencimiento,
            cantidad: tomar,
            esEstibaCompleta: tomar >= 10,
          });
          restante -= tomar;
        }

        cantidadNecesaria = restante;
      }

      const cantidadARestockear = ubicacionesSugeridas.reduce((sum, u) => sum + u.cantidad, 0);
      const cantidadFaltante = Math.max(
        0,
        cantidadVendida - cantidadDisponiblePicking - cantidadARestockear,
      );

      suggestions.push({
        sku: ventaInfo.skuDisplay || sku,
        descripcion: maestraData?.descripcion || ventaInfo.descripcion || '',
        cantidadVendida,
        cantidadVendidaOriginal,
        cantidadDisponiblePicking,
        cantidadEnReserva,
        cantidadARestockear: destinoValido ? cantidadARestockear : 0,
        cantidadFaltante,
        ubicacionesSugeridas,
        lpnDestino: pickingItems[0]?.LPN || maestraData?.lpn || null,
        localizacionDestino: destinoValido ? maestraData?.localizacion || null : null,
        proyeccion: esProyeccion,
        proyeccionDias,
        cantidadProyectada,
        prioridadAlta: skusAltaRotacion.has(sku),
        existeEnMaestra,
        estadoSurtido: cantidadARestockear > 0 ? `Surtir: ${cantidadARestockear}` : 'Cubierto',
      });

      if (cantidadFaltante > 0) {
        let tipoFalta: 'SIN_INVENTARIO' | 'SIN_RESERVA' | 'RESERVA_INSUFICIENTE';
        if (stockItems.length === 0) tipoFalta = 'SIN_INVENTARIO';
        else if (cantidadEnReserva === 0) tipoFalta = 'SIN_RESERVA';
        else tipoFalta = 'RESERVA_INSUFICIENTE';

        missingProducts.push({
          sku: ventaInfo.skuDisplay || sku,
          descripcion: maestraData?.descripcion || ventaInfo.descripcion || '',
          cantidadVendida,
          cantidadFaltante,
          stockEnPicking: cantidadDisponiblePicking,
          stockEnReserva: cantidadEnReserva,
          cantidadCubierta: cantidadARestockear,
          tipoFalta,
        });
      }
    }

    suggestions.sort((a, b) => {
      if (a.cantidadARestockear > 0 && b.cantidadARestockear === 0) return -1;
      if (a.cantidadARestockear === 0 && b.cantidadARestockear > 0) return 1;
      return a.sku.localeCompare(b.sku);
    });

    const totalProductos = suggestions.length;
    const productosConSurtido = suggestions.filter((s) => s.cantidadARestockear > 0).length;
    const productosFaltantes = missingProducts.length;
    const totalUnidadesSurtir = suggestions.reduce((sum, s) => sum + s.cantidadARestockear, 0);
    const totalDemanda = suggestions.reduce((sum, s) => sum + s.cantidadVendida, 0);
    const totalCobertura = suggestions.reduce(
      (sum, s) =>
        sum + Math.min(s.cantidadVendida, s.cantidadDisponiblePicking + s.cantidadARestockear),
      0,
    );
    const porcentajeCobertura = totalDemanda
      ? Math.round((totalCobertura / totalDemanda) * 100)
      : 100;
    const productosConProyeccion = suggestions.filter((s) => s.proyeccion).length;

    return {
      summary: `SKUs analizados: ${totalProductos}. Surtido sugerido para ${productosConSurtido} SKU(s), faltantes: ${productosFaltantes}.`,
      suggestions,
      missingProducts,
      stats: {
        totalProductos,
        productosConSurtido,
        productosFaltantes,
        totalUnidadesSurtir,
        porcentajeCobertura,
        productosConProyeccion,
      },
    };
  }

  private mapHeadersToStandard(
    headers: string[],
    mapping: Record<string, string[]>,
  ): Record<string, string> {
    const result: Record<string, string> = {};
    const normalizedHeaders = headers.map((h) => this.normalizeHeader(h));

    for (const key in mapping) {
      const aliases = mapping[key].map((a) => this.normalizeHeader(a));
      for (const alias of aliases) {
        const idx = normalizedHeaders.findIndex(
          (h) => h === alias || h.includes(alias) || alias.includes(h),
        );
        if (idx !== -1) {
          result[key] = headers[idx];
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

    if ((ext === 'csv' || ext === 'txt') && typeof data === 'string') {
      return this.parseCSV(data);
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

  private normalizeHeader(value: any): string {
    return String(value ?? '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/\s+/g, ' ')
      .trim()
      .toUpperCase();
  }

  private normalizeSkuKey(value: any): string {
    let v = String(value ?? '')
      .trim()
      .toUpperCase();
    if (!v) return '';

    v = v.replace(/\s+/g, '');
    v = v.replace(/\.0+$/, '');

    if (/^\d+$/.test(v)) {
      v = v.replace(/^0+(?=\d)/, '');
    }

    return v;
  }

  private normalizeSkuDisplay(value: any): string {
    let v = String(value ?? '')
      .trim()
      .toUpperCase();
    if (!v) return '';

    v = v.replace(/\s+/g, '');
    v = v.replace(/\.0+$/, '');
    return v;
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
