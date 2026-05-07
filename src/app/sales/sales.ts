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

// ============================================================
// CONFIGURACIÓN DE ANÁLISIS
// ============================================================
const VALID_STATUSES = ["STOCK EN ALMACEN LIBRE", "DISPONIBLE", "13 - CCL  DISPONIBLE"];
const IGNORED_LOCATIONS = ["PDIF-INV-1-10", "DEV-1-10", "PDIF-RES-1-10", "DEV-RES-1-10",];
const PICKING_LEVELS = ["0", "2", "3", "4", "5", "6", "7", "8", "9", "10", "15"];
const RESERVE_LEVELS = ['20', '30', '40', '50', '60', '70'];
const ADDITIONAL_RESERVE_LOCATIONS = ["MUELLE ENTRADA"];
const THRESHOLD_VENTAS_ALTAS = 10;

// ============================================================
// MAPPINGS DE COLUMNAS
// ============================================================
export const salesColumnMapping = {
  material: ["Material", "ID de Producto", "codigo", "Material", "SKU", "Código", "CODIGO", "CODE"],
  descripcion: ["Descripción", "Nombre de Artículo", "Descripcion", "Producto", "DESCRIPCION", "PRODUCTO"],
  cantidadConfirmada: ["cantidad confirmada", "Cant. Facturada", "Cantidad", "Qty", "Unidades", "CANTIDAD", "QTY"],
};

export const inventoryColumnMapping = {
  sku: ["SKU", "Item Code", "Codigo", "Código", "Material", "CODIGO", "CODE"],
  lpn: ["LPN", "Pallet ID", "Lpn", "License Plate", "LOTE"],
  descripcion: ["Descripcion", "Description", "Descripción", "Producto", "DESCRIPCION"],
  localizacion: ["Localizacion", "Location", "Ubicación", "Posición", "LOCATION"],
  disponible: ["Disponible", "Available", "Cantidad", "Unidades", "Stock", "DISPO", "STOCK"],
  estado: ["Estado", "Status", "Condición", "Situación", "ESTADO"],
  fechaEntrada: ["Fecha de entrada", "Fecha Entrada", "Fecha ingreso", "Entry Date", "Receipt Date", "FECHA_ENTRADA"],
  fechaVencimiento: ["Fecha de vencimiento", "Expiration", "fecha caducidad", "Expiry Date", "FECHA_VENC"],
  diasFPC: ["FPC", "Days to Exp", "DIAS FPC", "Días FPC", "Fech. Porc. Cons.", "DIAS_RESTANTES"],
  lote: ["Lote", "Batch", "Ce. Lote", "Lot", "LOTE"],
};

// ============================================================
// INTERFACES
// ============================================================
interface LocationData {
  lpn: string;
  localizacion: string;
  disponible: number;
  fechaVencimiento?: string | null;
  diasFPC?: number | null;
}

interface InventoryAggregate {
  descripcion: string;
  totalEnPicking: number;
  totalPrimaryReserve: number;
  totalAdditionalReserve: number;
  pickingLocations: LocationData[];
  primaryReserveLocations: LocationData[];
  additionalReserveLocations: LocationData[];
}

interface UbicacionSugerida {
  lpn?: string;
  localizacion: string;
  diasFPC?: number | null;
  fechaVencimiento?: string | null;
  cantidad: number;
  esEstibaCompleta?: boolean;
}

interface RestockSuggestion {
  sku: string;
  descripcion: string;
  cantidadVendida: number;
  cantidadDisponible: number;
  cantidadARestockear: number;
  cantidadTotalCubierta?: number;
  cantidadFaltante: number;
  ubicacionesSugeridas: UbicacionSugerida[];
  tipoFalta?: string;
  lpnDestino?: string | null;
  localizacionDestino?: string | null;
}

interface MissingProduct {
  sku: string;
  descripcion: string;
  cantidadVendida: number;
  cantidadFaltante: number;
  stockEnPicking?: number;
  stockEnReserva?: number;
  cantidadCubierta?: number;
  tipoFalta: string;
}

interface AnalysisStats {
  totalProductos: number;
  productosConSurtido: number;
  productosFaltantes: number;
  totalUnidadesSurtir: number;
  porcentajeCobertura: number;
}

interface AnalysisResult {
  summary: string;
  suggestions: RestockSuggestion[];
  missingProducts: MissingProduct[];
  stats: AnalysisStats;
}

interface FileData {
  name: string;
  data: string | ArrayBuffer | null;
  size: number;
  uploadDate: Date;
}

@Component({
  selector: 'app-sales',
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
  templateUrl: './sales.html',
  styleUrls: ['./sales.css']
})
export class Sales {
analizarArchivos() {
throw new Error('Method not implemented.');
}
  @ViewChild('dt1') dt1!: Table;
  @ViewChild('dt2') dt2!: Table;
  
  // Signals
  salesFile = signal<FileData | null>(null);
  inventoryFile = signal<FileData | null>(null);
  analysisResult = signal<AnalysisResult | null>(null);
  loading = signal(false);
  error = signal<string | null>(null);
  
  // Properties
  activeTab: 'missing' | 'suggestions' = 'missing';
  searchValue: string = '';
  isBrowser: boolean;
  isExpandedAll: boolean = false;
  readonly breadcrumbItems: MenuItem[] = [{ label: 'Ventas', routerLink: '/sales' }];
  readonly breadcrumbHome: MenuItem = { icon: 'pi pi-home', label: 'Inicio', routerLink: '/' };
  
  // Header maps
  salesHeaderMap: Record<string, string> = {};
  inventoryHeaderMap: Record<string, string> = {};
  
  // Drag and drop
  isDraggingSales = signal(false);
  isDraggingInventory = signal(false);
  
  // Column visibility
  visibleColumns: Record<string, boolean> = {
    sku: true,
    descripcion: true,
    cantidadVendida: true,
    cantidadDisponible: true,
    cantidadARestockear: true,
    cantidadFaltante: true,
    ubicaciones: true
  };

  // Getters para separar la lógica de filtrado por cada pestaña
  get filteredMissingProducts() {
    return (this.analysisResult()?.missingProducts || []).filter(m => m.cantidadFaltante > 0);
  }

  get filteredRestockSuggestions() {
    return (this.analysisResult()?.suggestions || []).filter(
      s => s.cantidadARestockear > 0 && s.tipoFalta !== 'SIN_INVENTARIO'
    );
  }

  get missingSkusSummary(): string {
    const skus = this.filteredMissingProducts.map(item => String(item.sku).trim());
    return skus.length ? skus.join(', ') : 'Sin SKUs faltantes';
  }

  get restockSkusSummary(): string {
    const skus = this.filteredRestockSuggestions.map(item => String(item.sku).trim());
    return skus.length ? skus.join(', ') : 'Sin SKUs para surtido';
  }

  // Elimina primeng, usaremos primengConfig

  constructor(
    @Inject(PLATFORM_ID) platformId: Object,
    private messageService: MessageService, 
    private primeng: PrimeNG,
  ) {
    this.isBrowser = isPlatformBrowser(platformId);
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

  // ============================================================
  // MÉTODOS DE UI
  // ============================================================
  
  toggleColumn(column: string) {
    this.visibleColumns[column] = !this.visibleColumns[column];
  }
  
  toggleExpandAll() {
    if (!this.dt2) return;

    if (this.isExpandedAll) {
      this.dt2.expandedRowKeys = {};
      this.isExpandedAll = false;
    } else {
      const suggestions = this.analysisResult()?.suggestions || [];
      const allExpanded: { [key: string]: boolean } = {};
      suggestions.forEach(s => {
        if (s && s.sku) {
          allExpanded[s.sku] = true;
        }
      });
      // Asignar un nuevo objeto para forzar el cambio de detección
      this.dt2.expandedRowKeys = { ...allExpanded };
      this.isExpandedAll = true;
    }
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
      life: 3000 
    });
  }
  
  // ============================================================
  // MÉTODOS DE ESTADÍSTICAS
  // ============================================================
  
  getTotalProducts(): number {
    return this.analysisResult()?.suggestions?.length || 0;
  }
  
  getProductsToRestock(): number {
    return this.filteredRestockSuggestions.length;
  }
  
  getMissingProductsCount(): number {
    return this.filteredMissingProducts.length;
  }
  
  getTotalUnitsToRestock(): number {
    return this.analysisResult()?.suggestions?.reduce((sum, s) => sum + s.cantidadARestockear, 0) || 0;
  }
  
  getTotalMissingUnits(): number {
    return this.analysisResult()?.missingProducts?.reduce((sum, m) => sum + m.cantidadFaltante, 0) || 0;
  }
  
  getCoberturaPorcentaje(): number {
    const result = this.analysisResult();
    if (!result) return 0;
    const totalVendido = result.suggestions.reduce((sum, s) => sum + s.cantidadVendida, 0);
    const totalCubierto = result.suggestions.reduce((sum, s) => sum + (s.cantidadDisponible + s.cantidadARestockear), 0);
    if (totalVendido === 0) return 100;
    return Math.round((totalCubierto / totalVendido) * 100);
  }
  
  getTipoFaltaSeverity(tipoFalta: string): 'success' | 'info' | 'warn' | 'danger' | 'secondary' | 'contrast' {
    const severities: Record<string, 'success' | 'info' | 'warn' | 'danger' | 'secondary' | 'contrast'> = {
      'SIN_INVENTARIO': 'danger',
      'SIN_RESERVA': 'danger',
      'RESERVA_INSUFICIENTE': 'warn',
      'OK': 'success'
    };
    return severities[tipoFalta] || 'info';
  }
  
  getTipoFaltaIcon(tipoFalta: string): string {
    const icons: Record<string, string> = {
      'SIN_INVENTARIO': 'pi-times-circle',
      'SIN_RESERVA': 'pi-ban',
      'RESERVA_INSUFICIENTE': 'pi-exclamation-triangle',
      'OK': 'pi-check-circle'
    };
    return icons[tipoFalta] || 'pi-info-circle';
  }

  getTipoFaltaLabel(tipoFalta: string): string {
    if (!tipoFalta) return '';
    const normalized = tipoFalta.toUpperCase();
    const labels: Record<string, string> = {
      'SIN_INVENTARIO': 'Sin inventario',
      'SIN_RESERVA': 'Sin reserva',
      'RESERVA_INSUFICIENTE': 'Reserva insuficiente',
      'OK': 'Ok'
    };

    if (labels[normalized]) {
      return labels[normalized];
    }

    const text = normalized.replace(/_/g, ' ').toLowerCase();
    return text.charAt(0).toUpperCase() + text.slice(1);
  }
  
  formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  // ============================================================
  // MANEJO DE ARCHIVOS
  // ============================================================
  
  onDragOver(event: DragEvent, type: 'sales' | 'inventory') {
    event.preventDefault();
    event.stopPropagation();
    if (type === 'sales') {
      this.isDraggingSales.set(true);
    } else {
      this.isDraggingInventory.set(true);
    }
  }
  
  onDragLeave(event: DragEvent, type: 'sales' | 'inventory') {
    event.preventDefault();
    event.stopPropagation();
    if (type === 'sales') {
      this.isDraggingSales.set(false);
    } else {
      this.isDraggingInventory.set(false);
    }
  }
  
  onDrop(event: DragEvent, type: 'sales' | 'inventory') {
    event.preventDefault();
    event.stopPropagation();
    
    if (type === 'sales') {
      this.isDraggingSales.set(false);
    } else {
      this.isDraggingInventory.set(false);
    }
    
    const files = event.dataTransfer?.files;
    if (files && files.length > 0) {
      this.processFile(files[0], type);
    }
  }
  
  onFileUpload(type: 'sales' | 'inventory', event: Event) {
    if (!this.isBrowser) return;
    
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    
    this.processFile(file, type);
  }
  
  private processFile(file: File, type: 'sales' | 'inventory') {
    if (!this.isBrowser) return;

    const validExtensions = ['.csv', '.xlsx', '.xls', '.txt'];
    const fileExt = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    if (!validExtensions.includes(fileExt)) {
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: 'Formato no soportado. Use CSV, Excel o TXT',
        life: 5000
      });
      return;
    }

    this.analysisResult.set(null);
    this.error.set(null);

    const reader = new FileReader();
    reader.onload = () => {
      const data = reader.result;
      // Validar columnas requeridas antes de guardar el archivo
      let rows: any[] = [];
      let headers: string[] = [];
      const ext = file.name.split('.').pop()?.toLowerCase();
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
            life: 5000
          });
          return;
        }
      } else if (typeof data === 'string') {
        const lines = data.split(/\r?\n/).filter(l => l.trim().length > 0);
        if (lines.length > 0) {
          let delimiter = ',';
          if (lines[0].includes(';')) delimiter = ';';
          headers = lines[0].split(delimiter).map(h => h.trim().replace(/^['"]|['"]$/g, ''));
        }
      }

      // Validar columnas requeridas usando los mapeos
      let mapping: Record<string, string[]>;
      let requiredKeys: string[];
      if (type === 'sales') {
        mapping = salesColumnMapping;
        requiredKeys = ['material', 'descripcion', 'cantidadConfirmada'];
      } else {
        mapping = inventoryColumnMapping;
        requiredKeys = ['sku', 'descripcion', 'localizacion', 'disponible', 'estado'];
      }
      const headersLower = headers.map(h => h.toLowerCase().trim());
      const missing: string[] = [];
      for (const key of requiredKeys) {
        const aliases = mapping[key].map(a => a.toLowerCase().trim());
        const found = headersLower.some(h => aliases.includes(h));
        if (!found) {
          // Mostrar el primer alias como nombre de columna esperada
          missing.push(mapping[key][0]);
        }
      }
      if (missing.length > 0) {
        this.messageService.add({
          severity: 'error',
          summary: 'Columnas faltantes',
          detail: `El archivo no contiene las columnas requeridas: ${missing.join(', ')}`,
          life: 7000
        });
        return;
      }

      // Si pasa la validación, guardar el archivo
      const fileData: FileData = {
        name: file.name,
        data: data,
        size: file.size,
        uploadDate: new Date()
      };

      if (type === 'sales') {
        this.salesFile.set(fileData);
        if (typeof data === 'string') {
          const rows = this.parseFileData(data, file.name);
          if (rows.length > 0) {
            const headers = Object.keys(rows[0]);
            this.salesHeaderMap = this.mapHeadersToStandard(headers, salesColumnMapping);
            this.messageService.add({
              severity: 'success',
              summary: 'Archivo cargado',
              detail: `Ventas: ${rows.length} registros encontrados`,
              life: 3000
            });
          }
        }
      } else {
        this.inventoryFile.set(fileData);
        if (typeof data === 'string') {
          const rows = this.parseFileData(data, file.name);
          if (rows.length > 0) {
            const headers = Object.keys(rows[0]);
            this.inventoryHeaderMap = this.mapHeadersToStandard(headers, inventoryColumnMapping);
            this.messageService.add({
              severity: 'success',
              summary: 'Archivo cargado',
              detail: `Inventario: ${rows.length} registros encontrados`,
              life: 3000
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
        life: 5000
      });
    };

    const ext = file.name.split('.').pop()?.toLowerCase();
    if (ext === 'xlsx' || ext === 'xls') {
      reader.readAsBinaryString(file);
    } else {
      reader.readAsText(file, 'UTF-8');
    }
  }
  
  removeFile(type: 'sales' | 'inventory') {
    if (type === 'sales') {
      this.salesFile.set(null);
      this.salesHeaderMap = {};
    } else {
      this.inventoryFile.set(null);
      this.inventoryHeaderMap = {};
    }
    this.analysisResult.set(null);
    this.messageService.add({ 
      severity: 'info', 
      summary: 'Archivo removido', 
      detail: `Se ha eliminado el archivo de ${type === 'sales' ? 'ventas' : 'inventario'}`, 
      life: 3000 
    });
  }
  
  // ============================================================
  // ANÁLISIS PRINCIPAL
  // ============================================================
  
  generateSalesAnalysis() {
    if (!this.salesFile() || !this.inventoryFile()) {
      this.messageService.add({ 
        severity: 'warn', 
        summary: 'Archivos faltantes', 
        detail: 'Debes subir ambos archivos (Ventas e Inventario) para analizar', 
        life: 5000 
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
          life: 5000 
        });
      } catch (err) {
        const errorMsg = 'Error al procesar los archivos: ' + (err as Error).message;
        this.error.set(errorMsg);
        this.loading.set(false);
        this.messageService.add({ 
          severity: 'error', 
          summary: 'Error', 
          detail: errorMsg, 
          life: 7000 
        });
      }
    }, 100);
  }
  
  // ============================================================
  // EXPORTAR RESULTADOS
  // ============================================================
  
  exportToExcel() {
    const result = this.analysisResult();
    if (!result) return;

    const wb = XLSX.utils.book_new();

    // ── Hoja 1: Surtido (una fila por ubicación origen) ──────
    const surtidoData: any[][] = [
      ['SKU', 'Descripción', 'Vendido', 'Stock Picking', 'Cant. a Surtir',
       'Ubicación Origen', 'LPN Origen', 'Tipo', 'Estiba', 'Cubierto', 'Faltante']
    ];
    for (const s of this.filteredRestockSuggestions) {
      const locs = s.ubicacionesSugeridas;
      if (locs.length === 0) {
        surtidoData.push([
          s.sku, s.descripcion, s.cantidadVendida, s.cantidadDisponible,
          0, '', '', '', '', 0, s.cantidadFaltante
        ]);
        continue;
      }
      locs.forEach((u, idx) => {
        const esPallet = (u.esEstibaCompleta === true) || (u.cantidad > 10);
        const tipo  = esPallet ? '✅ Pallet Completo'   : '🔄 Unidades Parciales';
        const estiba = esPallet ? '📦 SI' : '📦 NO';
        surtidoData.push([
          s.sku,
          s.descripcion,
          idx === 0 ? s.cantidadVendida    : 0,
          idx === 0 ? s.cantidadDisponible : 0,
          u.cantidad,
          u.localizacion,
          u.lpn ?? '',
          tipo,
          estiba,
          idx === 0 ? s.cantidadARestockear : 0,
          idx === 0 ? s.cantidadFaltante    : 0
        ]);
      });
    }
    const wsSurtido = XLSX.utils.aoa_to_sheet(surtidoData);
    wsSurtido['!cols'] = [12, 50, 10, 14, 14, 14, 14, 22, 10, 10, 10].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsSurtido, 'Surtido');

    // ── Hoja 2: Faltantes ────────────────────────────────────
    const faltantesData: any[][] = [
      ['SKU', 'Descripción', 'Vendidas', 'Stock Picking', 'Stock Reserva', 'Faltante', 'Estado']
    ];
    for (const m of this.filteredMissingProducts) {
      faltantesData.push([
        m.sku,
        m.descripcion,
        m.cantidadVendida,
        m.stockEnPicking ?? 0,
        m.stockEnReserva ?? 0,
        m.cantidadFaltante,
        m.tipoFalta
      ]);
    }
    const wsFaltantes = XLSX.utils.aoa_to_sheet(faltantesData);
    wsFaltantes['!cols'] = [10, 45, 10, 14, 14, 10, 20].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsFaltantes, 'Faltantes');

    // ── Descargar ────────────────────────────────────────────
    const fecha = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `analisis_${fecha}.xlsx`);

    this.messageService.add({
      severity: 'success',
      summary: 'Exportado',
      detail: 'Archivo Excel generado con hojas Surtido y Faltantes',
      life: 3000
    });
  }
  
  private escapeCSV(str: any): string {
    if (str === undefined || str === null) return '';
    const stringified = String(str);
    if (stringified.includes(',') || stringified.includes('"') || stringified.includes('\n')) {
      return `"${stringified.replace(/"/g, '""')}"`;
    }
    return stringified;
  }
  
  // ============================================================
  // MÉTODOS PRIVADOS DE PROCESAMIENTO
  // ============================================================
  
  private mapHeadersToStandard(headers: string[], mapping: Record<string, string[]>): Record<string, string> {
    const result: Record<string, string> = {};
    const headersLower = headers.map(h => h.toLowerCase().trim());

    for (const key in mapping) {
      // Iterar aliases en orden de prioridad: el primer alias que coincida con
      // algún header gana. Así "SKU" tiene prioridad sobre "Codigo" aunque
      // "Codigo" aparezca antes en el archivo.
      const aliases = mapping[key].map(a => a.toLowerCase().trim());
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
    const lines = data.split(/\r?\n/).filter(l => l.trim().length > 0);
    if (lines.length < 2) return [];
    
    let delimiter = ',';
    if (lines[0].includes(';')) {
      delimiter = ';';
    }
    
    const headers = lines[0].split(delimiter).map(h => h.trim().replace(/^["']|["']$/g, ''));
    return lines.slice(1).map(line => {
      const values = line.split(delimiter);
      const obj: any = {};
      headers.forEach((h, i) => { 
        obj[h] = values[i] ? values[i].trim().replace(/^["']|["']$/g, '') : ''; 
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
    const salesRaw = this.salesFile()?.data;
    const inventoryRaw = this.inventoryFile()?.data;
    const salesName = this.salesFile()?.name || '';
    const inventoryName = this.inventoryFile()?.name || '';
    
    if (typeof salesRaw !== 'string' || typeof inventoryRaw !== 'string') {
      throw new Error('Archivos no válidos');
    }
    
    const salesRows = this.parseFileData(salesRaw, salesName);
    const inventoryRows = this.parseFileData(inventoryRaw, inventoryName);
    const salesMap = this.salesHeaderMap;
    const invMap = this.inventoryHeaderMap;

    // 1. Filtrar inventario válido
    // Se acepta si el estado contiene algún valor válido O si no hay campo estado reconocible
    // para no descartar inventario por problema de mapeo de columna.
    const freeStockInventory = inventoryRows.filter(item => {
      const estadoRaw = this.getNestedValue(item, invMap['estado'] || 'Estado', null);
      const estado = String(estadoRaw ?? '').toUpperCase();
      const loc = String(this.getNestedValue(item, invMap['localizacion'] || 'Localizacion', '')).toUpperCase();
      const isIgnoredLoc = IGNORED_LOCATIONS.some(iloc => loc.includes(iloc.toUpperCase()));
      // Si no se puede leer el estado, incluir el item (mejor tener de más que de menos)
      const estadoOk = estado === '' || VALID_STATUSES.some(s => estado.includes(s.toUpperCase()));
      return estadoOk && !isIgnoredLoc;
    });
    
    // 2. Agrupar inventario por SKU
    const inventoryBySku: Record<string, InventoryAggregate> = {};
    
    for (const item of freeStockInventory) {
      const sku = String(this.getNestedValue(item, invMap['sku'] || 'SKU', '')).toUpperCase().trim();
      if (!sku) continue;
      
      if (!inventoryBySku[sku]) {
        inventoryBySku[sku] = {
          descripcion: this.getNestedValue(item, invMap['descripcion'] || 'Descripcion', ''),
          totalEnPicking: 0,
          totalPrimaryReserve: 0,
          totalAdditionalReserve: 0,
          pickingLocations: [],
          primaryReserveLocations: [],
          additionalReserveLocations: [],
        };
      }
      
      const loc = String(this.getNestedValue(item, invMap['localizacion'] || 'Localizacion', ''));
      const lastPart = loc.split('-').pop() || '';
      const isPicking = PICKING_LEVELS.includes(lastPart);
      const isReserveByLevel = RESERVE_LEVELS.includes(lastPart);
      const isReserveByAdditional = ADDITIONAL_RESERVE_LOCATIONS.some(prefix => loc.toUpperCase().startsWith(prefix));
      
      const locationData: LocationData = {
        lpn: this.getNestedValue(item, invMap['lpn'] || 'LPN', ''),
        localizacion: loc,
        disponible: parseFloat(this.getNestedValue(item, invMap['disponible'] || 'Disponible', '0')) || 0,
        fechaVencimiento: this.getNestedValue(item, invMap['fechaVencimiento'] || 'Fecha de vencimiento', null),
        diasFPC: this.getNestedValue(item, invMap['diasFPC'] || 'FPC', null),
      };
      
      if (isPicking) {
        inventoryBySku[sku].totalEnPicking += locationData.disponible;
        inventoryBySku[sku].pickingLocations.push(locationData);
      } else if (isReserveByLevel) {
        inventoryBySku[sku].totalPrimaryReserve += locationData.disponible;
        inventoryBySku[sku].primaryReserveLocations.push(locationData);
      } else if (isReserveByAdditional) {
        inventoryBySku[sku].totalAdditionalReserve += locationData.disponible;
        inventoryBySku[sku].additionalReserveLocations.push(locationData);
      } else {
        // Ubicación no clasificada: se trata como reserva adicional para que
        // su stock siempre sea contado en inventarioTotal.
        inventoryBySku[sku].totalAdditionalReserve += locationData.disponible;
        inventoryBySku[sku].additionalReserveLocations.push(locationData);
      }
    }
    
    // 3. Agrupar ventas por SKU
    const salesBySku: Record<string, { descripcion: string; totalVendida: number }> = {};
    
    for (const row of salesRows) {
      const sku = String(this.getNestedValue(row, salesMap['material'] || 'Material', '')).toUpperCase().trim();
      if (!sku) continue;
      
      if (!salesBySku[sku]) {
        salesBySku[sku] = {
          descripcion: this.getNestedValue(row, salesMap['descripcion'] || 'Descripción', ''),
          totalVendida: 0,
        };
      }
      
      const cantidad = parseFloat(this.getNestedValue(row, salesMap['cantidadConfirmada'] || 'Cantidad', '0')) || 0;
      salesBySku[sku].totalVendida += cantidad;
    }
    
    // 4. Generar sugerencias y faltantes (misma lógica del sistema original)
    const suggestions: RestockSuggestion[] = [];
    const missingProducts: MissingProduct[] = [];
    const okProducts: RestockSuggestion[] = [];

    for (const sku in salesBySku) {
      const venta = salesBySku[sku];
      const inventario = inventoryBySku[sku];

      // Caso 1: vendido sin inventario
      if (!inventario) {
        missingProducts.push({
          sku,
          descripcion: venta.descripcion,
          cantidadVendida: venta.totalVendida,
          cantidadFaltante: venta.totalVendida,
          stockEnPicking: 0,
          stockEnReserva: 0,
          cantidadCubierta: 0,
          tipoFalta: 'SIN_INVENTARIO',
        });
        continue;
      }

      // Caso 2: picking insuficiente
      if (inventario.totalEnPicking < venta.totalVendida) {
        const amountToRestock = venta.totalVendida - inventario.totalEnPicking;
        const usePrimaryReserve = inventario.totalPrimaryReserve > 0;
        const reserveLocationsToUse = usePrimaryReserve
          ? [...inventario.primaryReserveLocations]
          : [...inventario.additionalReserveLocations];

        const totalReserveDisponible = reserveLocationsToUse.reduce((sum, loc) => sum + (loc.disponible || 0), 0);

        // Subcaso 2a: hay reserva para surtir
        if (totalReserveDisponible > 0) {
          const sortedReserveLocations = reserveLocationsToUse.sort((a, b) => this.sortByFefo(a, b));
          const isHighTurnover = amountToRestock >= THRESHOLD_VENTAS_ALTAS;
          const ubicacionesSugeridas: UbicacionSugerida[] = [];
          let cantidadARestockear = 0;

          for (const location of sortedReserveLocations) {
            if (cantidadARestockear >= amountToRestock) break;

            const amountToTake = isHighTurnover
              ? location.disponible
              : Math.min(location.disponible, amountToRestock - cantidadARestockear);

            if (amountToTake <= 0) continue;

            cantidadARestockear += amountToTake;
            ubicacionesSugeridas.push({
              lpn: location.lpn,
              localizacion: location.localizacion,
              diasFPC: location.diasFPC,
              fechaVencimiento: location.fechaVencimiento,
              cantidad: amountToTake,
              esEstibaCompleta: amountToTake === location.disponible,
            });
          }

          const cantidadFaltante = Math.max(0, amountToRestock - cantidadARestockear);
          const cantidadCubierta = Math.min(totalReserveDisponible, amountToRestock);

          suggestions.push({
            sku,
            descripcion: inventario.descripcion || venta.descripcion,
            cantidadVendida: venta.totalVendida,
            cantidadDisponible: inventario.totalEnPicking,
            cantidadARestockear,
            cantidadTotalCubierta: cantidadARestockear,
            cantidadFaltante,
            ubicacionesSugeridas,
            tipoFalta: cantidadFaltante > 0 ? 'RESERVA_INSUFICIENTE' : 'OK',
            lpnDestino: null,
            localizacionDestino: null,
          });

          if (cantidadFaltante > 0) {
            missingProducts.push({
              sku,
              descripcion: inventario.descripcion || venta.descripcion,
              cantidadVendida: venta.totalVendida,
              cantidadFaltante,
              stockEnPicking: inventario.totalEnPicking,
              stockEnReserva: totalReserveDisponible,
              cantidadCubierta,
              tipoFalta: 'RESERVA_INSUFICIENTE',
            });
          }
        } else {
          // Subcaso 2b: no hay reserva
          missingProducts.push({
            sku,
            descripcion: inventario.descripcion || venta.descripcion,
            cantidadVendida: venta.totalVendida,
            cantidadFaltante: amountToRestock,
            stockEnPicking: inventario.totalEnPicking,
            stockEnReserva: 0,
            cantidadCubierta: 0,
            tipoFalta: 'SIN_RESERVA',
          });
        }
      } else {
        // Caso 3: picking suficiente
        okProducts.push({
          sku,
          descripcion: inventario.descripcion || venta.descripcion,
          cantidadVendida: venta.totalVendida,
          cantidadDisponible: inventario.totalEnPicking,
          cantidadARestockear: 0,
          cantidadTotalCubierta: 0,
          cantidadFaltante: 0,
          ubicacionesSugeridas: inventario.pickingLocations.map(loc => ({
            lpn: loc.lpn,
            localizacion: loc.localizacion,
            fechaVencimiento: loc.fechaVencimiento,
            diasFPC: loc.diasFPC,
            cantidad: loc.disponible,
            esEstibaCompleta: false,
          })),
          tipoFalta: 'OK',
          lpnDestino: null,
          localizacionDestino: null,
        });
      }
    }

    const allSuggestions = [...suggestions, ...okProducts];
    allSuggestions.sort((a, b) => {
      if (a.cantidadARestockear > 0 && b.cantidadARestockear === 0) return -1;
      if (a.cantidadARestockear === 0 && b.cantidadARestockear > 0) return 1;
      return a.sku.localeCompare(b.sku);
    });

    const summary = `✅ Análisis completado: ${allSuggestions.length} productos analizados, ${missingProducts.length} con faltante.`;

    const totalVendido = allSuggestions.reduce((sum, s) => sum + s.cantidadVendida, 0);
    const totalCubierto = allSuggestions.reduce((sum, s) => sum + (s.cantidadDisponible + s.cantidadARestockear), 0);
    const porcentajeCobertura = totalVendido === 0 ? 100 : Math.round((totalCubierto / totalVendido) * 100);

    return {
      summary,
      suggestions: allSuggestions,
      missingProducts,
      stats: {
        totalProductos: allSuggestions.length,
        productosConSurtido: allSuggestions.filter(s => s.cantidadARestockear > 0).length,
        productosFaltantes: missingProducts.filter(m => m.cantidadFaltante > 0).length,
        totalUnidadesSurtir: allSuggestions.reduce((sum, s) => sum + s.cantidadARestockear, 0),
        porcentajeCobertura: porcentajeCobertura
      }
    };
  }
}