// lot-cross.component.ts
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

type LotCrossFileType = 'client' | 'wms';

// ============================================================
// MAPPINGS DE COLUMNAS
// ============================================================
export const wmsLotCrossMapping = {
  sku: ['Sku', 'SKU', 'Sku 2', 'Sku2', 'Codigo', 'Código'],
  descripcion: ['Descripción', 'Descripcion', 'Desc'],
  lote: ['Lote', 'Batch', 'Ce. Lote'],
  cantidad: ['Unidades', 'Cajas', 'Cantidad'],
  fechaVencimiento: [
    'Vencimiento',
    'Fecha Vencimiento',
    'Fecha Vcto',
    'F. Venc.',
    'FV',
    'Fecha_Venc',
    'Fecha de Vencimiento',
    'Vencto.',
    'Vto.',
    'Expiry',
    'Exp. Date',
    'SLED/BBD',
    'Fecha caducidad',
    'Fecha de caducidad',
  ],
};

export const clientLotCrossMapping = {
  sku: ['Material', 'SKU', 'Codigo', 'Código'],
  descripcion: ['Texto breve de material', 'Descripcion', 'Descripción'],
  lote: ['Lote', 'Ce. Lote'],
  cantidad: [
    'Ctd.real dest.',
    'Ctd.real dest',
    'Ctd real dest.',
    'Ctd real dest',
    "Ctd teórica 'desde'",
    'Ctd Teórica',
    'Cantidad Teórica',
    'Stock disponible',
    'Ctd. teór. hacia',
    "Ctd real 'desde'",
  ],
  fechaVencimiento: [
    'SLED/BBD',
    'Fecha caducidad',
    'Fecha de caducidad',
    'Vencimiento',
    'Fecha Vencimiento',
    'Vencto.',
    'Vto.',
    'FeCaduc/FePreferCons',
  ],
};

const requiredColumnsByType: Record<LotCrossFileType, string[]> = {
  client: ['sku', 'lote'],
  wms: ['sku', 'lote'],
};

// ============================================================
// INTERFACES
// ============================================================
interface Aggregate {
  descripcion: string;
  lotes: Set<string>;
  cantidad: number;
  loteFechas: Map<string, string>; // lote -> fechaVencimiento
}

interface LotCrossRow {
  sku: string;
  descripcion: string;
  lotesCliente: string[];
  lotesWms: string[];
  lotesSoloCliente: string[];
  lotesSoloWms: string[];
  cantidadCliente: number;
  cantidadWms: number;
  estado: 'OK' | 'DIFERENTE';
  fechasVencimientoCliente: string;
  fechasVencimientoWms: string;
  fechasVencimiento: string;
}

interface LotCrossAnalysisResult {
  results: LotCrossRow[];
  summary: string;
  stats: {
    totalSkus: number;
    skusOk: number;
    skusDiferentes: number;
    diferenciasEnLotes: number;
  };
}

interface FileData {
  name: string;
  data: string | ArrayBuffer | null;
  size: number;
  uploadDate: Date;
}

@Component({
  selector: 'app-lot-cross',
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
  templateUrl: './lot-cross.html',
  styleUrls: ['./lot-cross.css'],
})
export class LotCross {
  @ViewChild('dtResults') dtResults!: Table;

  // ============================================================
  // SIGNALS Y ESTADO
  // ============================================================
  clientFile = signal<FileData | null>(null);
  wmsFile = signal<FileData | null>(null);
  searchValue: string = '';
  loading = signal(false);

  isDraggingClient = signal(false);
  isDraggingWms = signal(false);

  analysisResult = signal<LotCrossAnalysisResult | null>(null);
  activeTab: 'diferente' | 'ok' = 'diferente';

  breadcrumbItems: MenuItem[] = [{ label: 'Cruce de Lotes' }];
  breadcrumbHome = { icon: 'pi pi-home', label: 'Inicio', routerLink: '/' };

  isBrowser: boolean;

  get filteredResults(): LotCrossRow[] {
    const result = this.analysisResult();
    if (!result) return [];

    const query = this.searchValue.toLowerCase().trim();
    let filtered = result.results;

    if (this.activeTab === 'diferente') {
      filtered = filtered.filter((r) => r.estado === 'DIFERENTE');
    } else {
      filtered = filtered.filter((r) => r.estado === 'OK');
    }

    if (!query) return filtered;

    return filtered.filter(
      (r) => r.sku.toLowerCase().includes(query) || r.descripcion.toLowerCase().includes(query),
    );
  }

  constructor(
    private messageService: MessageService,
    @Inject(PLATFORM_ID) platformId: Object,
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

  // ============================================================
  // FILE HANDLING
  // ============================================================
  onDragOver(event: DragEvent, fileType: 'client' | 'wms'): void {
    event.preventDefault();
    if (fileType === 'client') this.isDraggingClient.set(true);
    else this.isDraggingWms.set(true);
  }

  onDragLeave(event: DragEvent, fileType: 'client' | 'wms'): void {
    event.preventDefault();
    if (fileType === 'client') this.isDraggingClient.set(false);
    else this.isDraggingWms.set(false);
  }

  onDrop(event: DragEvent, fileType: 'client' | 'wms'): void {
    event.preventDefault();
    if (fileType === 'client') this.isDraggingClient.set(false);
    else this.isDraggingWms.set(false);

    const files = event.dataTransfer?.files;
    if (files && files.length > 0) {
      this.processFile(files[0], fileType);
    }
  }

  onFileUpload(fileType: 'client' | 'wms', event: Event): void {
    if (!this.isBrowser) return;
    const input = event.target as HTMLInputElement;
    if (input.files && input.files.length > 0) {
      this.processFile(input.files[0], fileType);
    }
  }

  private processFile(file: File, fileType: 'client' | 'wms'): void {
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

    this.analysisResult.set(null);

    const reader = new FileReader();
    reader.onload = (e: ProgressEvent<FileReader>) => {
      try {
        const data = e.target?.result || null;
        const headers = this.extractHeaders(data, file.name);
        const count = this.estimateRecordCount(data, file.name);

        // Validar columnas requeridas
        const mapping = fileType === 'client' ? clientLotCrossMapping : wmsLotCrossMapping;
        const requiredKeys = requiredColumnsByType[fileType];
        const missingCols = this.validateRequiredColumns(headers, mapping, requiredKeys);

        if (missingCols.length > 0) {
          this.messageService.add({
            severity: 'error',
            summary: 'Columnas faltantes',
            detail: `Faltan columnas: ${missingCols.join(', ')}`,
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

        const fileData: FileData = {
          name: file.name,
          data: data,
          size: file.size,
          uploadDate: new Date(),
        };

        if (fileType === 'client') {
          this.clientFile.set(fileData);
        } else {
          this.wmsFile.set(fileData);
        }

        this.messageService.add({
          severity: 'success',
          summary: 'Archivo cargado',
          detail: `${fileType.toUpperCase()}: ${count} registros detectados`,
          life: 3000,
        });
      } catch (err) {
        this.messageService.add({
          severity: 'error',
          summary: 'Error',
          detail: 'No se pudo procesar el archivo',
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

  removeFile(fileType: 'client' | 'wms'): void {
    if (fileType === 'client') this.clientFile.set(null);
    else this.wmsFile.set(null);
    this.analysisResult.set(null);

    this.messageService.add({
      severity: 'info',
      summary: 'Archivo removido',
      detail: `Se ha eliminado el archivo de ${fileType === 'client' ? 'Cliente' : 'WMS'}`,
      life: 3000,
    });
  }

  formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + ' ' + sizes[i];
  }

  // ============================================================
  // VALIDACIÓN DE COLUMNAS
  // ============================================================
  private validateRequiredColumns(
    headers: string[],
    mapping: any,
    requiredKeys: string[],
  ): string[] {
    const headersLower = headers.map((h) => h.toLowerCase().trim());
    const missing: string[] = [];

    for (const key of requiredKeys) {
      const aliases = mapping[key].map((a: string) => a.toLowerCase().trim());
      const found = headersLower.some((h) => aliases.includes(h));
      if (!found) {
        missing.push(mapping[key][0]);
      }
    }
    return missing;
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
      const headerIdx = this.findHeaderRowIndex(rows);
      return (rows[headerIdx] || []).map((h: any) => String(h ?? '').trim());
    }
    if (typeof data === 'string') {
      const lines = data.split(/\r?\n/).filter((l) => l.trim().length > 0);
      if (!lines.length) return [];
      const delimiter = this.detectDelimiter(lines[0]);
      const rows = lines.map((l) =>
        l.split(delimiter).map((c) => c.trim().replace(/^['"\s]+|['"\s]+$/g, '')),
      );
      const headerIdx = this.findHeaderRowIndex(rows);
      return rows[headerIdx] || [];
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
      const headerIdx = this.findHeaderRowIndex(rows);
      return Math.max(rows.length - headerIdx - 1, 0);
    }
    if (typeof data === 'string') {
      const lines = data.split(/\r?\n/).filter((l) => l.trim().length > 0);
      if (!lines.length) return 0;
      const delimiter = this.detectDelimiter(lines[0]);
      const rows = lines.map((l) => l.split(delimiter));
      const headerIdx = this.findHeaderRowIndex(rows);
      return Math.max(lines.length - headerIdx - 1, 0);
    }
    return 0;
  }

  private detectDelimiter(headerLine: string): string {
    if (headerLine.includes('\t')) return '\t';
    if (headerLine.includes(';')) return ';';
    return ',';
  }

  // Detecta la fila de cabecera de datos, saltando metadatos de albarán
  private findHeaderRowIndex(rows: any[][]): number {
    const knownHeaders = new Set([
      'sku',
      'sku 2',
      'sku2',
      'descripcion',
      'descripción',
      'lote',
      'batch',
      'material',
      'codigo',
      'código',
      'unidades',
      'cajas',
      'cantidad',
      'ce. lote',
      'estado',
      'serial',
      'precio',
      'uom primaria',
    ]);
    for (let i = 0; i < Math.min(rows.length, 60); i++) {
      const row = rows[i];
      if (!row || row.length < 3) continue;
      const cellsLower = row.map((c: any) =>
        String(c ?? '')
          .toLowerCase()
          .trim(),
      );
      const matchCount = cellsLower.filter((c: string) => knownHeaders.has(c)).length;
      if (matchCount >= 2) return i;
    }
    return 0;
  }

  // ============================================================
  // PARSEO DE ARCHIVOS
  // ============================================================
  private parseExcelFile(fileData: string | ArrayBuffer | null, mapping: any): any[] {
    if (!fileData) return [];

    try {
      let workbook: XLSX.WorkBook;
      if (typeof fileData === 'string') {
        workbook = XLSX.read(fileData, { type: 'binary' });
      } else {
        workbook = XLSX.read(fileData);
      }

      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const allRows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: '',
        raw: false,
      }) as any[][];

      if (allRows.length === 0) return [];

      // Detectar si es albarán (cabecera de metadatos antes de la tabla de datos)
      const headerIdx = this.findHeaderRowIndex(allRows);
      const headerRow = allRows[headerIdx].map((h: any) => String(h ?? '').trim());

      // Construir mapa columna -> índice usando los aliases del mapping
      const colMap: Record<string, number> = {};
      for (const [key, aliases] of Object.entries(mapping)) {
        const idx = headerRow.findIndex((h) =>
          (aliases as string[]).some((a) => h.toLowerCase() === a.toLowerCase()),
        );
        if (idx >= 0) colMap[key] = idx;
      }

      const parsedData: any[] = [];
      for (let r = headerIdx + 1; r < allRows.length; r++) {
        const row = allRows[r];
        if (!row || row.every((c: any) => String(c ?? '').trim() === '')) continue;

        const mappedRow: any = {};
        for (const [key, idx] of Object.entries(colMap)) {
          mappedRow[key] = row[idx as number];
        }
        if (mappedRow.sku && mappedRow.lote) {
          parsedData.push(mappedRow);
        }
      }

      return parsedData;
    } catch (err: any) {
      throw new Error(`Error al parsear archivo: ${err.message}`);
    }
  }

  // ============================================================
  // ANÁLISIS PRINCIPAL
  // ============================================================
  analyzeLotCross(): void {
    if (!this.clientFile() || !this.wmsFile()) {
      this.messageService.add({
        severity: 'warn',
        summary: 'Archivos faltantes',
        detail: 'Debes subir archivos Cliente y WMS para ejecutar el análisis',
        life: 5000,
      });
      return;
    }

    this.loading.set(true);

    setTimeout(() => {
      try {
        const clientData = this.parseExcelFile(this.clientFile()!.data, clientLotCrossMapping);
        const wmsData = this.parseExcelFile(this.wmsFile()!.data, wmsLotCrossMapping);

        if (!clientData.length || !wmsData.length) {
          this.messageService.add({
            severity: 'error',
            summary: 'Sin datos',
            detail: 'Uno o más archivos no contienen datos válidos',
            life: 5000,
          });
          this.loading.set(false);
          return;
        }

        const result = this.runLotCrossAnalysis(clientData, wmsData);
        this.analysisResult.set(result);

        this.messageService.add({
          severity: 'success',
          summary: 'Cruce de lotes completado',
          detail: `${result.stats.totalSkus} SKUs analizados - ${result.stats.skusDiferentes} con diferencias`,
          life: 4000,
        });
      } catch (err: any) {
        this.messageService.add({
          severity: 'error',
          summary: 'Error en análisis',
          detail: err.message || 'No fue posible completar el análisis',
          life: 7000,
        });
      } finally {
        this.loading.set(false);
      }
    }, 100);
  }

  private runLotCrossAnalysis(clientData: any[], wmsData: any[]): LotCrossAnalysisResult {
    const norm = (value: any) =>
      String(value || '')
        .trim()
        .toUpperCase();
    const toNumber = (value: any) => {
      if (typeof value === 'number') return value;
      if (value === null || value === undefined) return 0;
      const cleaned = String(value).replace(/\s/g, '').replace(/\./g, '').replace(',', '.');
      const parsed = Number(cleaned);
      return Number.isFinite(parsed) ? parsed : 0;
    };
    const buildFechasTexto = (loteFechas?: Map<string, string>) => {
      if (!loteFechas || loteFechas.size === 0) return '';

      const uniqueFechas = Array.from(
        new Set(
          Array.from(loteFechas.values())
            .map((f) => String(f || '').trim())
            .filter(Boolean),
        ),
      );

      return uniqueFechas.join(' | ');
    };

    // Agregación Cliente
    const clientBySku = new Map<string, Aggregate>();
    for (const row of clientData) {
      const sku = norm(row.sku);
      const lote = norm(row.lote);
      if (!sku || !lote) continue;

      const current = clientBySku.get(sku) ?? {
        descripcion: norm(row.descripcion || ''),
        lotes: new Set<string>(),
        cantidad: 0,
        loteFechas: new Map<string, string>(),
      };
      current.lotes.add(lote);
      current.cantidad += toNumber(row.cantidad) || 1;
      if (!current.descripcion && row.descripcion) current.descripcion = norm(row.descripcion);
      if (row.fechaVencimiento) current.loteFechas.set(lote, String(row.fechaVencimiento).trim());
      clientBySku.set(sku, current);
    }

    // Agregación WMS
    const wmsBySku = new Map<string, Aggregate>();
    for (const row of wmsData) {
      const sku = norm(row.sku);
      const lote = norm(row.lote);
      if (!sku || !lote) continue;

      const current = wmsBySku.get(sku) ?? {
        descripcion: norm(row.descripcion || ''),
        lotes: new Set<string>(),
        cantidad: 0,
        loteFechas: new Map<string, string>(),
      };
      current.lotes.add(lote);
      current.cantidad += toNumber(row.cantidad) || 1;
      if (!current.descripcion && row.descripcion) current.descripcion = norm(row.descripcion);
      if (row.fechaVencimiento) current.loteFechas.set(lote, String(row.fechaVencimiento).trim());
      wmsBySku.set(sku, current);
    }

    // Comparación
    const allSkus = new Set([...clientBySku.keys(), ...wmsBySku.keys()]);
    const results: LotCrossRow[] = Array.from(allSkus).map((sku) => {
      const client = clientBySku.get(sku);
      const wms = wmsBySku.get(sku);

      const lotesCliente = Array.from(client?.lotes ?? []).sort();
      const lotesWms = Array.from(wms?.lotes ?? []).sort();
      const clientSet = new Set(lotesCliente);
      const wmsSet = new Set(lotesWms);

      const lotesSoloCliente = lotesCliente.filter((l) => !wmsSet.has(l));
      const lotesSoloWms = lotesWms.filter((l) => !clientSet.has(l));
      const estado =
        lotesSoloCliente.length === 0 && lotesSoloWms.length === 0 ? 'OK' : 'DIFERENTE';

      const fechasVencimientoCliente = buildFechasTexto(client?.loteFechas);
      const fechasVencimientoWms = buildFechasTexto(wms?.loteFechas);
      const fechasVencimiento = [fechasVencimientoCliente, fechasVencimientoWms]
        .filter(Boolean)
        .join(' | ');

      return {
        sku,
        descripcion: client?.descripcion || wms?.descripcion || '',
        lotesCliente,
        lotesWms,
        lotesSoloCliente,
        lotesSoloWms,
        cantidadCliente: client?.cantidad ?? 0,
        cantidadWms: wms?.cantidad ?? 0,
        estado,
        fechasVencimientoCliente,
        fechasVencimientoWms,
        fechasVencimiento,
      };
    });

    results.sort((a, b) => {
      if (a.estado !== b.estado) return a.estado === 'DIFERENTE' ? -1 : 1;
      return a.sku.localeCompare(b.sku);
    });

    const stats = {
      totalSkus: results.length,
      skusOk: results.filter((r) => r.estado === 'OK').length,
      skusDiferentes: results.filter((r) => r.estado === 'DIFERENTE').length,
      diferenciasEnLotes: results.filter(
        (r) => r.lotesSoloCliente.length > 0 || r.lotesSoloWms.length > 0,
      ).length,
    };

    const summaryText =
      stats.skusDiferentes > 0
        ? `⚠️ ${stats.skusDiferentes} SKU(s) con discrepancias detectadas`
        : `✅ Todos los ${stats.totalSkus} SKU(s) están sincronizados`;

    return { results, summary: summaryText, stats };
  }

  clearFilters(): void {
    if (this.dtResults) {
      this.dtResults.clear();
      this.dtResults.filterGlobal('', 'contains');
    }
    this.searchValue = '';
    this.messageService.add({
      severity: 'info',
      summary: 'Filtros limpiados',
      detail: 'Se han eliminado todos los filtros',
      life: 3000,
    });
  }

  // ============================================================
  // UTILIDADES VISUALES
  // ============================================================
  getEstadoSeverity(estado: 'OK' | 'DIFERENTE'): 'success' | 'danger' {
    return estado === 'OK' ? 'success' : 'danger';
  }

  getEstadoIcon(estado: 'OK' | 'DIFERENTE'): string {
    return estado === 'OK' ? 'pi pi-check' : 'pi pi-exclamation-circle';
  }

  // ============================================================
  // EXPORTACIÓN A EXCEL
  // ============================================================
  exportToExcel(): void {
    const result = this.analysisResult();
    if (!result) return;

    const wb = XLSX.utils.book_new();

    // Hoja Resumen
    const summaryData = [
      { Métrica: 'Total SKUs', Valor: result.stats.totalSkus },
      { Métrica: 'SKUs OK', Valor: result.stats.skusOk },
      { Métrica: 'SKUs con Diferencias', Valor: result.stats.skusDiferentes },
      { Métrica: 'Discrepancias en Lotes', Valor: result.stats.diferenciasEnLotes },
    ];
    const wsResumen = XLSX.utils.json_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(wb, wsResumen, 'Resumen');

    // Hoja Diferencias
    const difData = result.results
      .filter((r) => r.estado === 'DIFERENTE')
      .map((r) => ({
        SKU: r.sku,
        Descripción: r.descripcion,
        'Lotes Cliente': r.lotesCliente.join(', '),
        'Lotes WMS': r.lotesWms.join(', '),
        'Solo en Cliente': r.lotesSoloCliente.join(', '),
        'Solo en WMS': r.lotesSoloWms.join(', '),
        'Cant. Cliente': r.cantidadCliente,
        'Cant. WMS': r.cantidadWms,
        'Fecha Venc. Cliente': r.fechasVencimientoCliente || '',
        'Fecha Venc. WMS': r.fechasVencimientoWms || '',
      }));
    const wsDif = XLSX.utils.json_to_sheet(difData);
    XLSX.utils.book_append_sheet(wb, wsDif, 'Diferencias');

    // Hoja Completa
    const allData = result.results.map((r) => ({
      SKU: r.sku,
      Descripción: r.descripcion,
      'Lotes Cliente': r.lotesCliente.join(', '),
      'Lotes WMS': r.lotesWms.join(', '),
      'Solo en Cliente': r.lotesSoloCliente.join(', '),
      'Solo en WMS': r.lotesSoloWms.join(', '),
      'Cant. Cliente': r.cantidadCliente,
      'Cant. WMS': r.cantidadWms,
      'Fecha Venc. Cliente': r.fechasVencimientoCliente || '',
      'Fecha Venc. WMS': r.fechasVencimientoWms || '',
      Estado: r.estado,
    }));
    const wsAll = XLSX.utils.json_to_sheet(allData);
    XLSX.utils.book_append_sheet(wb, wsAll, 'Todos');

    XLSX.writeFile(wb, `cruce-lotes-${new Date().getTime()}.xlsx`);
    this.messageService.add({
      severity: 'success',
      summary: 'Exportado',
      detail: 'Archivo Excel generado correctamente',
      life: 3000,
    });
  }
}
