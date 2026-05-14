import { Component, signal, Inject, PLATFORM_ID, ViewChild, OnInit, computed } from '@angular/core';
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
const EXPORT_INVENTORY_TARGET_AISLE = 'P10';
const IGNORED_LOCATION_KEYWORDS = ['PDIF-INV-1-10', 'DEV-1-10', 'PDIF-RES-1-10', 'DEV-RES-1-10'];

// ============================================================
// MAPPINGS DE COLUMNAS
// ============================================================
export const exportInventoryMasterMapping: Record<string, string[]> = {
  Cod: ['Cod', 'Código', 'Material', 'SKU', 'CODIGO', 'CODE', 'Item Code'],
  REFERENCIA: [
    'REFERENCIA',
    'Referencia',
    'Descripción',
    'Descripcion',
    'Nombre',
    'Description',
    'PRODUCTO',
  ],
};

export const exportInventoryMapping: Record<string, string[]> = {
  Codigo: ['SKU', 'Material', 'CODIGO', 'CODE', 'Item Code', 'SKU2'],
  LPN: ['LPN', 'Pallet', 'Pallet ID', 'License Plate', 'LOTE', 'Batch'],
  Localizacion: ['Localizacion', 'Localización', 'Ubicacion', 'Location', 'Ubicación', 'LOCATION'],
};

// ============================================================
// INTERFACES
// ============================================================
interface FileData {
  name: string;
  data: string | ArrayBuffer | null;
  size: number;
  uploadDate: Date;
}

interface MaestraRow {
  Cod: string;
  REFERENCIA: string;
}

interface InventarioRow {
  Codigo: string;
  LPN: string | null;
  Localizacion: string | null;
}

interface ResultadoExportacion {
  codigo: string;
  referencia: string;
  localizacionActual: string | null;
  estado: 'OK' | 'MOVIMIENTO AL PASILLO SUGERIDO';
  localizacionSugerida: string | null;
  sugerencia: string;
  lpn: string | null;
}

interface AnalysisStats {
  totalRegistros: number;
  totalOK: number;
  totalMovimientos: number;
  porCodigo: Map<string, { total: number; ok: number; movimientos: number }>;
}

interface AnalysisResult {
  resultados: ResultadoExportacion[];
  stats: AnalysisStats;
}

@Component({
  selector: 'app-exclusive-skus',
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
  templateUrl: './exclusive-skus.html',
  styleUrls: ['./exclusive-skus.css'],
})
export class ExclusiveSkus implements OnInit {
  @ViewChild('dt1') dt1!: Table;

  // Signals
  maestraFile = signal<FileData | null>(null);
  inventarioFile = signal<FileData | null>(null);
  analysisResult = signal<AnalysisResult | null>(null);
  loading = signal(false);

  // UI State
  activeTab: 'todos' | 'movimientos' = 'todos';
  searchValue: string = '';
  isBrowser: boolean;

  // Drag and drop
  isDraggingMaestra = signal(false);
  isDraggingInventario = signal(false);

  // Computed filtered results
  filteredResultados = computed(() => {
    const resultados = this.analysisResult()?.resultados || [];
    let filtered = [...resultados];

    // Filter by tab
    if (this.activeTab === 'movimientos') {
      filtered = filtered.filter((r) => r.estado !== 'OK');
    }

    // Filter by search
    if (this.searchValue) {
      const search = this.searchValue.toLowerCase();
      filtered = filtered.filter(
        (r) =>
          r.codigo.toLowerCase().includes(search) || r.referencia.toLowerCase().includes(search),
      );
    }

    return filtered;
  });

  readonly breadcrumbItems: MenuItem[] = [
    { label: 'SKUs Exclusivos', routerLink: '/exclusive-skus' },
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
    // PrimeNG translations
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

  clearFilters() {
    this.searchValue = '';
    this.activeTab = 'todos';
    this.messageService.add({
      severity: 'info',
      summary: 'Filtros limpiados',
      detail: 'Se han eliminado todos los filtros',
      life: 3000,
    });
  }

  onTabChange(tab: 'todos' | 'movimientos') {
    this.activeTab = tab;
  }

  onSearch() {
    // El filtro se actualiza automáticamente por el computed signal
  }

  formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  getConteoOK(): number {
    return this.analysisResult()?.stats.totalOK || 0;
  }

  getConteoMovimientos(): number {
    return this.analysisResult()?.stats.totalMovimientos || 0;
  }

  getPorcentajeOK(): number {
    const stats = this.analysisResult()?.stats;
    if (!stats || stats.totalRegistros === 0) return 0;
    return Math.round((stats.totalOK / stats.totalRegistros) * 100);
  }

  esPasilloDestino(localizacion: string | null): boolean {
    if (!localizacion) return false;
    return localizacion.toUpperCase().startsWith(EXPORT_INVENTORY_TARGET_AISLE);
  }

  // ============================================================
  // MANEJO DE ARCHIVOS
  // ============================================================

  onDragOver(event: DragEvent, type: 'maestra' | 'inventario') {
    event.preventDefault();
    event.stopPropagation();
    if (type === 'maestra') {
      this.isDraggingMaestra.set(true);
    } else {
      this.isDraggingInventario.set(true);
    }
  }

  onDragLeave(event: DragEvent, type: 'maestra' | 'inventario') {
    event.preventDefault();
    event.stopPropagation();
    if (type === 'maestra') {
      this.isDraggingMaestra.set(false);
    } else {
      this.isDraggingInventario.set(false);
    }
  }

  onDrop(event: DragEvent, type: 'maestra' | 'inventario') {
    event.preventDefault();
    event.stopPropagation();

    if (type === 'maestra') {
      this.isDraggingMaestra.set(false);
    } else {
      this.isDraggingInventario.set(false);
    }

    const files = event.dataTransfer?.files;
    if (files && files.length > 0) {
      this.processFile(files[0], type);
    }
  }

  onFileUpload(type: 'maestra' | 'inventario', event: Event) {
    if (!this.isBrowser) return;

    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;

    this.processFile(file, type);
  }

  private processFile(file: File, type: 'maestra' | 'inventario') {
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

    const reader = new FileReader();
    reader.onload = () => {
      const data = reader.result;

      let headers: string[] = [];
      let recordCount = 0;

      const ext = file.name.split('.').pop()?.toLowerCase();
      if (ext === 'xlsx' || ext === 'xls') {
        try {
          const workbook = XLSX.read(data, { type: 'binary' });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          const rows: any[][] = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            defval: '',
            raw: false,
          }) as any[][];
          headers = (rows[0] || []).map((h: any) => String(h).trim());
          recordCount = Math.max(rows.length - 1, 0);
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
          const delimiter = lines[0].includes(';') ? ';' : ',';
          headers = lines[0].split(delimiter).map((h) => h.trim().replace(/^['"]|['"]$/g, ''));
          recordCount = Math.max(lines.length - 1, 0);
        }
      }

      const validacionPlantilla = this.validarArchivoSegunCampo(headers, type);

      if (!validacionPlantilla.ok && validacionPlantilla.templateError) {
        this.messageService.add({
          severity: 'error',
          summary: 'Plantilla incorrecta',
          detail: validacionPlantilla.templateError,
          life: 7000,
        });
        return;
      }

      if (!validacionPlantilla.ok && validacionPlantilla.missingCols.length > 0) {
        this.messageService.add({
          severity: 'error',
          summary: 'Columnas faltantes',
          detail: `El archivo no contiene las columnas requeridas: ${validacionPlantilla.missingCols.join(', ')}`,
          life: 7000,
        });
        return;
      }

      if (recordCount === 0) {
        this.messageService.add({
          severity: 'error',
          summary: 'Sin datos',
          detail: 'El archivo no contiene filas de datos válidos',
          life: 5000,
        });
        return;
      }

      const fileData: FileData = {
        name: file.name,
        data: data,
        size: file.size,
        uploadDate: new Date(),
      };

      if (type === 'maestra') {
        this.maestraFile.set(fileData);
        this.messageService.add({
          severity: 'success',
          summary: 'Archivo cargado',
          detail: `Maestra: ${recordCount} registros encontrados`,
          life: 3000,
        });
      } else {
        this.inventarioFile.set(fileData);
        this.messageService.add({
          severity: 'success',
          summary: 'Archivo cargado',
          detail: `Inventario: ${recordCount} registros encontrados`,
          life: 3000,
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

    const ext = file.name.split('.').pop()?.toLowerCase();
    if (ext === 'xlsx' || ext === 'xls') {
      reader.readAsBinaryString(file);
    } else {
      reader.readAsText(file, 'UTF-8');
    }
  }

  removeFile(type: 'maestra' | 'inventario') {
    if (type === 'maestra') {
      this.maestraFile.set(null);
    } else {
      this.inventarioFile.set(null);
    }
    this.analysisResult.set(null);
    this.messageService.add({
      severity: 'info',
      summary: 'Archivo removido',
      detail: `Se ha eliminado el archivo de ${type === 'maestra' ? 'maestra' : 'inventario'}`,
      life: 3000,
    });
  }

  // ============================================================
  // MÉTODOS PRIVADOS DE PROCESAMIENTO
  // ============================================================

  private evaluarCoberturaColumnas(
    headersNormalizados: string[],
    mapping: Record<string, string[]>,
  ): { totalKeys: number; foundKeys: string[]; missingKeys: string[] } {
    const foundKeys: string[] = [];
    const missingKeys: string[] = [];

    for (const key of Object.keys(mapping)) {
      const aliases = mapping[key].map((a) => this.normalizarHeader(a));
      const found = headersNormalizados.some((h) =>
        aliases.some((alias) => h === alias || h.startsWith(alias)),
      );
      if (found) {
        foundKeys.push(key);
      } else {
        missingKeys.push(key);
      }
    }

    return {
      totalKeys: Object.keys(mapping).length,
      foundKeys,
      missingKeys,
    };
  }

  private validarArchivoSegunCampo(
    headers: string[],
    type: 'maestra' | 'inventario',
  ): { ok: boolean; missingCols: string[]; templateError?: string } {
    const expectedMapping =
      type === 'maestra' ? exportInventoryMasterMapping : exportInventoryMapping;
    const oppositeMapping =
      type === 'maestra' ? exportInventoryMapping : exportInventoryMasterMapping;
    const tipoEsperado = type === 'maestra' ? 'Maestra' : 'Inventario';
    const tipoOpuesto = type === 'maestra' ? 'Inventario' : 'Maestra';

    const headersNormalizados = headers.map((h) => this.normalizarHeader(h));
    const expectedCoverage = this.evaluarCoberturaColumnas(headersNormalizados, expectedMapping);
    const oppositeCoverage = this.evaluarCoberturaColumnas(headersNormalizados, oppositeMapping);

    const missingCols = expectedCoverage.missingKeys.map((key) => expectedMapping[key][0]);
    if (missingCols.length > 0) {
      return { ok: false, missingCols };
    }

    // Obtener columnas discriminantes (únicas de cada plantilla)
    const expectedKeys = Object.keys(expectedMapping);
    const oppositeKeys = Object.keys(oppositeMapping);
    const discriminantesExpected = expectedKeys.filter((k) => !oppositeKeys.includes(k));
    const discriminantesOpposite = oppositeKeys.filter((k) => !expectedKeys.includes(k));

    // Contar cuántas aliases de columnas discriminantes están presentes
    const countDiscriminantsExpected = discriminantesExpected.reduce((count, key) => {
      const aliases = expectedMapping[key].map((a) => this.normalizarHeader(a));
      return count + (headersNormalizados.some((h) => aliases.includes(h)) ? 1 : 0);
    }, 0);

    const countDiscriminantsOpposite = discriminantesOpposite.reduce((count, key) => {
      const aliases = oppositeMapping[key].map((a) => this.normalizarHeader(a));
      return count + (headersNormalizados.some((h) => aliases.includes(h)) ? 1 : 0);
    }, 0);

    // Si la plantilla opuesta tiene MÁS columnas discriminantes, es probable que sea del tipo opuesto
    if (countDiscriminantsOpposite > countDiscriminantsExpected) {
      return {
        ok: false,
        missingCols: [],
        templateError: `Este archivo coincide con la plantilla de ${tipoOpuesto}. Cárgalo en el campo de ${tipoOpuesto}, no en ${tipoEsperado}.`,
      };
    }

    return { ok: true, missingCols: [] };
  }

  private parseExcelFile(file: FileData): any[] {
    try {
      const ext = file.name.split('.').pop()?.toLowerCase();

      if (ext === 'xlsx' || ext === 'xls') {
        const wb =
          typeof file.data === 'string'
            ? XLSX.read(file.data, { type: 'binary' })
            : XLSX.read(file.data, { type: 'array' });
        const sheetName = wb.SheetNames[0];
        const worksheet = wb.Sheets[sheetName];
        return XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false });
      }

      if ((ext === 'csv' || ext === 'txt') && typeof file.data === 'string') {
        const lines = file.data.split(/\r?\n/).filter((l) => l.trim().length > 0);
        if (lines.length < 2) return [];
        const delimiter = lines[0].includes(';') ? ';' : ',';
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
    } catch (error) {
      console.error('Error parsing file:', error);
      return [];
    }
  }

  private mapearColumnas(row: any, mapping: Record<string, string[]>): any {
    const resultado: any = {};
    const keysRow = Object.keys(row);
    const keysRowNormalized = keysRow.map((k) => this.normalizarHeader(k));

    for (const [targetKey, aliases] of Object.entries(mapping)) {
      for (const alias of aliases) {
        const aliasNormalized = this.normalizarHeader(alias);
        const matchIndex = keysRowNormalized.findIndex(
          (k) => k === aliasNormalized || k.startsWith(aliasNormalized),
        );
        const matchKey = matchIndex >= 0 ? keysRow[matchIndex] : undefined;
        if (matchKey) {
          resultado[targetKey] = row[matchKey];
          break;
        }
      }
      if (resultado[targetKey] === undefined) {
        resultado[targetKey] = null;
      }
    }
    return resultado;
  }

  private normalizarSku(sku: any): string {
    return String(sku ?? '').trim();
  }

  private normalizarHeader(header: string): string {
    return String(header ?? '')
      .trim()
      .replace(/^\uFEFF/, '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/\s+/g, '')
      .toLowerCase();
  }

  private esPasilloDestinoInternal(localizacion: string | null | undefined): boolean {
    if (!localizacion) return false;
    return localizacion.toUpperCase().startsWith(EXPORT_INVENTORY_TARGET_AISLE);
  }

  private obtenerPasillo(localizacion: string | null | undefined): string | null {
    if (!localizacion) return null;
    const match = localizacion.toUpperCase().match(/^([A-Za-z0-9]+)/);
    return match ? match[1] : null;
  }

  private validarPasilloP10(
    maestra: MaestraRow[],
    inventario: InventarioRow[],
  ): ResultadoExportacion[] {
    const resultados: ResultadoExportacion[] = [];

    // Crear Set con códigos de maestra
    const codigosMaestra = new Set<string>();
    const maestraMap = new Map<string, string>();

    for (const mat of maestra) {
      const codigoStr = this.normalizarSku(mat.Cod);
      codigosMaestra.add(codigoStr);
      maestraMap.set(codigoStr, mat.REFERENCIA || '');
    }

    // Filtrar inventario: SOLO códigos en maestra y NO en ubicaciones ignoradas
    const inventarioFiltrado = inventario.filter((inv) => {
      const codigoStr = this.normalizarSku(inv.Codigo);
      if (inv.Localizacion) {
        const loc = inv.Localizacion.toUpperCase();
        if (IGNORED_LOCATION_KEYWORDS.some((kw) => loc.includes(kw.toUpperCase()))) {
          return false;
        }
      }
      return codigosMaestra.has(codigoStr);
    });

    // Procesar inventario filtrado
    for (const inv of inventarioFiltrado) {
      const codigoStr = this.normalizarSku(inv.Codigo);
      const referencia = maestraMap.get(codigoStr) || '';
      const ubicacionActual = inv.Localizacion || null;
      const pasilloActual = this.obtenerPasillo(ubicacionActual);
      const enDestino = this.esPasilloDestinoInternal(ubicacionActual);

      if (enDestino) {
        resultados.push({
          codigo: codigoStr,
          referencia: referencia,
          localizacionActual: ubicacionActual,
          estado: 'OK',
          localizacionSugerida: null,
          sugerencia: `✅ Correctamente ubicado en ${ubicacionActual}`,
          lpn: inv.LPN || null,
        });
      } else {
        const sugerencia = pasilloActual
          ? `🚚 Mover de ${pasilloActual} a pasillo ${EXPORT_INVENTORY_TARGET_AISLE}`
          : `🚚 Mover a pasillo ${EXPORT_INVENTORY_TARGET_AISLE} para optimizar la ubicación`;

        resultados.push({
          codigo: codigoStr,
          referencia: referencia,
          localizacionActual: ubicacionActual,
          estado: 'MOVIMIENTO AL PASILLO SUGERIDO',
          localizacionSugerida: EXPORT_INVENTORY_TARGET_AISLE,
          sugerencia: sugerencia,
          lpn: inv.LPN || null,
        });
      }
    }

    return resultados;
  }

  private obtenerEstadisticas(resultados: ResultadoExportacion[]): AnalysisStats {
    const porCodigo = new Map<string, { total: number; ok: number; movimientos: number }>();
    let totalOK = 0;
    let totalMovimientos = 0;

    for (const r of resultados) {
      if (!porCodigo.has(r.codigo)) {
        porCodigo.set(r.codigo, { total: 0, ok: 0, movimientos: 0 });
      }
      const stats = porCodigo.get(r.codigo)!;
      stats.total++;

      if (r.estado === 'OK') {
        stats.ok++;
        totalOK++;
      } else {
        stats.movimientos++;
        totalMovimientos++;
      }
    }

    return {
      totalRegistros: resultados.length,
      totalOK,
      totalMovimientos,
      porCodigo,
    };
  }

  // ============================================================
  // ANÁLISIS PRINCIPAL
  // ============================================================

  generateAnalysis() {
    if (!this.maestraFile() || !this.inventarioFile()) {
      this.messageService.add({
        severity: 'warn',
        summary: 'Archivos faltantes',
        detail: 'Debes subir ambos archivos (Maestra e Inventario) para analizar',
        life: 5000,
      });
      return;
    }

    this.loading.set(true);

    setTimeout(() => {
      try {
        const maestraData = this.parseExcelFile(this.maestraFile()!);
        const inventarioData = this.parseExcelFile(this.inventarioFile()!);

        if (maestraData.length === 0 || inventarioData.length === 0) {
          this.messageService.add({
            severity: 'error',
            summary: 'Error',
            detail: 'Los archivos no contienen datos válidos',
            life: 5000,
          });
          this.loading.set(false);
          return;
        }

        // Mapear columnas
        const maestraMapeada: MaestraRow[] = maestraData.map((row) =>
          this.mapearColumnas(row, exportInventoryMasterMapping),
        );
        const inventarioMapeada: InventarioRow[] = inventarioData.map((row) =>
          this.mapearColumnas(row, exportInventoryMapping),
        );

        // Validar pasillo
        const resultados = this.validarPasilloP10(maestraMapeada, inventarioMapeada);
        const stats = this.obtenerEstadisticas(resultados);

        if (resultados.length === 0) {
          this.messageService.add({
            severity: 'warn',
            summary: 'Sin coincidencias',
            detail: `No hubo cruce entre Maestra (${maestraMapeada.length}) e Inventario (${inventarioMapeada.length}). Revisa formato de código y encabezados de SKU.`,
            life: 7000,
          });
        }

        this.analysisResult.set({ resultados, stats });

        this.messageService.add({
          severity: 'success',
          summary: 'Análisis completado',
          detail: `${resultados.length} registros procesados, ${stats.totalMovimientos} requieren movimiento`,
          life: 5000,
        });
      } catch (error) {
        const msg = 'Error al procesar los archivos: ' + (error as Error).message;
        this.messageService.add({
          severity: 'error',
          summary: 'Error',
          detail: msg,
          life: 7000,
        });
      } finally {
        this.loading.set(false);
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

    // Hoja 1: Resumen
    const resumenData = [
      { Métrica: 'Total Registros', Valor: result.stats.totalRegistros },
      { Métrica: 'Correctamente Ubicados (P10)', Valor: result.stats.totalOK },
      { Métrica: 'Requieren Movimiento', Valor: result.stats.totalMovimientos },
      { Métrica: 'Porcentaje Correctos', Valor: `${this.getPorcentajeOK()}%` },
      { Métrica: 'Fecha Análisis', Valor: new Date().toLocaleString() },
      { Métrica: 'Pasillo Destino', Valor: EXPORT_INVENTORY_TARGET_AISLE },
    ];

    const wsResumen = XLSX.utils.json_to_sheet(resumenData);
    XLSX.utils.book_append_sheet(wb, wsResumen, 'Resumen');

    // Hoja 2: Detalle de Movimientos
    const movimientosData = result.resultados.map((r) => ({
      'Código SKU': r.codigo,
      Referencia: r.referencia,
      'Ubicación Actual': r.localizacionActual || 'N/A',
      'Ubicación Sugerida': r.localizacionSugerida || 'OK',
      Estado: r.estado === 'OK' ? '✅ Correcto' : '⚠️ Requiere Movimiento',
      'Acción Sugerida': r.sugerencia,
      LPN: r.lpn || 'N/A',
    }));

    const wsMovimientos = XLSX.utils.json_to_sheet(movimientosData);
    XLSX.utils.book_append_sheet(wb, wsMovimientos, 'Detalle');

    // Hoja 3: Solo movimientos requeridos
    const soloMovimientos = result.resultados
      .filter((r) => r.estado !== 'OK')
      .map((r) => ({
        'Código SKU': r.codigo,
        Referencia: r.referencia,
        'Ubicación Actual': r.localizacionActual || 'N/A',
        'Ubicación Sugerida': r.localizacionSugerida || EXPORT_INVENTORY_TARGET_AISLE,
        'Acción Sugerida': r.sugerencia,
        LPN: r.lpn || 'N/A',
      }));

    const wsSoloMovimientos = XLSX.utils.json_to_sheet(soloMovimientos);
    XLSX.utils.book_append_sheet(wb, wsSoloMovimientos, 'Movimientos Requeridos');

    const fecha = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `Validacion_Pasillo_P10_${fecha}.xlsx`);

    this.messageService.add({
      severity: 'success',
      summary: 'Exportado',
      detail: `Archivo Excel generado con resumen y detalle`,
      life: 3000,
    });
  }
}
