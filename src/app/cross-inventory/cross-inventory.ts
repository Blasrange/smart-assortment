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

// Modulo: Cruce de Inventario (fuente Cliente)
export const clientInventoryMapping = {
  sku: ['Material', 'Material', 'Item'],
  descripcion: ['Texto breve material', 'Descripcion', 'Descripción'],
  lote: ['Ce. Lote', 'Lote'],
  cantidad: ['Stock disponible', 'WM stock disp.', 'Stock total', 'Disponible', 'Existencia'],
  fechaVencimiento: [
    'Cad./FPC',
    'Fecha de vencimiento',
    'Vencimiento',
    'Expiracion',
    'Caducidad',
    'FeCaduc/FePreferCons',
  ],
};

// Modulo: Cruce de Inventario (fuente WMS)
export const wmsInventoryCrossMapping = {
  sku: ['SKU', 'Codigo'],
  descripcion: ['Descripcion', 'Descripción'],
  lote: ['Lote', 'Batch', 'Ce. Lote'],
  cantidad: ['Disponible', 'Unidades'],
  fechaVencimiento: ['Fecha de vencimiento', 'Vencimiento', 'Expiracion', 'Caducidad'],
};

interface FileData {
  name: string;
  data: string | ArrayBuffer | null;
  size: number;
  uploadDate: Date;
}

interface CrossRow {
  sku: string;
  lote: string;
  descripcion: string;
  cantidadCliente: number;
  cantidadWms: number;
  diferencia: number;
  fechaVencimientoCliente: string;
  fechaVencimientoWms: string;
  diferenciaVencimientoDias: number | null;
  validacionVencimiento: 'Coincide' | 'Diferente' | 'Sin fecha';
}

interface CrossAnalysisResult {
  summary: string;
  results: CrossRow[];
  stats: {
    totalItems: number;
    itemsConDiferencia: number;
    coincidencias: number;
    porcentajeCoincidencia: number;
    diferenciaNeta: number;
  };
}

@Component({
  selector: 'app-cross-inventory',
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
  templateUrl: './cross-inventory.html',
  styleUrl: './cross-inventory.css',
})
export class CrossInventory {
  @ViewChild('dt1') dt1!: Table;

  clientFile = signal<FileData | null>(null);
  wmsFile = signal<FileData | null>(null);
  analysisResult = signal<CrossAnalysisResult | null>(null);
  groupByLot = signal(true);
  loading = signal(false);
  error = signal<string | null>(null);
  searchValue = '';
  isBrowser: boolean;

  clientHeaderMap: Record<string, string> = {};
  wmsHeaderMap: Record<string, string> = {};

  isDraggingClient = signal(false);
  isDraggingWms = signal(false);

  readonly breadcrumbItems: MenuItem[] = [
    { label: 'Cruce de Inventarios', routerLink: '/cross-inventory' },
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

  setGroupMode(value: boolean) {
    this.groupByLot.set(value);
    this.analysisResult.set(null);
  }

  onDragOver(event: DragEvent, type: 'client' | 'wms') {
    event.preventDefault();
    event.stopPropagation();
    if (type === 'client') {
      this.isDraggingClient.set(true);
      return;
    }
    this.isDraggingWms.set(true);
  }

  onDragLeave(event: DragEvent, type: 'client' | 'wms') {
    event.preventDefault();
    event.stopPropagation();
    if (type === 'client') {
      this.isDraggingClient.set(false);
      return;
    }
    this.isDraggingWms.set(false);
  }

  onDrop(event: DragEvent, type: 'client' | 'wms') {
    event.preventDefault();
    event.stopPropagation();
    if (type === 'client') {
      this.isDraggingClient.set(false);
    } else {
      this.isDraggingWms.set(false);
    }

    const files = event.dataTransfer?.files;
    if (files && files.length > 0) {
      this.processFile(files[0], type);
    }
  }

  onFileUpload(type: 'client' | 'wms', event: Event) {
    if (!this.isBrowser) return;
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    this.processFile(file, type);
  }

  private processFile(file: File, type: 'client' | 'wms') {
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

      const mapping = type === 'client' ? clientInventoryMapping : wmsInventoryCrossMapping;
      const requiredKeys: Array<keyof typeof mapping> = ['sku', 'cantidad'];

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
        data,
        size: file.size,
        uploadDate: new Date(),
      };

      const rows = this.parseFileData(data as string | ArrayBuffer, file.name);
      const fileHeaders = rows.length > 0 ? Object.keys(rows[0]) : headers;
      const headerMap = this.mapHeadersToStandard(fileHeaders, mapping);

      if (type === 'client') {
        this.clientFile.set(fileData);
        this.clientHeaderMap = headerMap;
      } else {
        this.wmsFile.set(fileData);
        this.wmsHeaderMap = headerMap;
      }

      this.messageService.add({
        severity: 'success',
        summary: 'Archivo cargado',
        detail: `${type === 'client' ? 'Cliente' : 'WMS'}: ${rows.length} registros encontrados`,
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

  removeFile(type: 'client' | 'wms') {
    if (type === 'client') {
      this.clientFile.set(null);
      this.clientHeaderMap = {};
    } else {
      this.wmsFile.set(null);
      this.wmsHeaderMap = {};
    }

    this.analysisResult.set(null);

    this.messageService.add({
      severity: 'info',
      summary: 'Archivo removido',
      detail: `Se ha eliminado el archivo de ${type === 'client' ? 'Cliente' : 'WMS'}`,
      life: 3000,
    });
  }

  generateCrossAnalysis() {
    if (!this.clientFile() || !this.wmsFile()) {
      this.messageService.add({
        severity: 'warn',
        summary: 'Archivos faltantes',
        detail: 'Debes subir ambos archivos (Cliente y WMS) para analizar',
        life: 5000,
      });
      return;
    }

    if (this.groupByLot() && (!this.clientHeaderMap['lote'] || !this.wmsHeaderMap['lote'])) {
      this.messageService.add({
        severity: 'warn',
        summary: 'Campo lote faltante',
        detail: 'Para agrupar por SKU + Lote, ambos archivos deben incluir columna de lote.',
        life: 6000,
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

    const crossResults = result.results;
    const totalCliente = crossResults.reduce((sum, row) => sum + row.cantidadCliente, 0);
    const totalWms = crossResults.reduce((sum, row) => sum + row.cantidadWms, 0);
    const totalSobranteUnidades = crossResults.reduce(
      (sum, row) => sum + (row.diferencia > 0 ? row.diferencia : 0),
      0,
    );
    const totalFaltanteUnidades = crossResults.reduce(
      (sum, row) => sum + (row.diferencia < 0 ? Math.abs(row.diferencia) : 0),
      0,
    );

    const totalSKUsSobrantes = crossResults.filter((r) => r.diferencia > 0).length;
    const totalSKUsFaltantes = crossResults.filter((r) => r.diferencia < 0).length;
    const totalSKUsOK = crossResults.filter((r) => r.diferencia === 0).length;

    const lotesMap = new Map<
      string,
      {
        lote: string;
        hasPos: boolean;
        hasNeg: boolean;
        sobrantes: number;
        faltantes: number;
        diferenciaNeta: number;
      }
    >();
    for (const row of crossResults) {
      const loteKey = row.lote || 'SIN LOTE';
      const bucket = lotesMap.get(loteKey) || {
        lote: loteKey,
        hasPos: false,
        hasNeg: false,
        sobrantes: 0,
        faltantes: 0,
        diferenciaNeta: 0,
      };
      if (row.diferencia > 0) bucket.hasPos = true;
      if (row.diferencia < 0) bucket.hasNeg = true;
      if (row.diferencia > 0) bucket.sobrantes += 1;
      if (row.diferencia < 0) bucket.faltantes += 1;
      bucket.diferenciaNeta += row.diferencia;
      lotesMap.set(loteKey, bucket);
    }

    let lotesSoloSobrantes = 0;
    let lotesSoloFaltantes = 0;
    let lotesMixtos = 0;
    let lotesOK = 0;

    for (const loteState of lotesMap.values()) {
      if (loteState.hasPos && loteState.hasNeg) {
        lotesMixtos += 1;
      } else if (loteState.hasPos) {
        lotesSoloSobrantes += 1;
      } else if (loteState.hasNeg) {
        lotesSoloFaltantes += 1;
      } else {
        lotesOK += 1;
      }
    }

    const lotesConNovedad = lotesSoloSobrantes + lotesSoloFaltantes + lotesMixtos;
    const totalLotes = lotesMap.size;
    const lotesMixtosDetalle = Array.from(lotesMap.values())
      .filter((lote) => lote.hasPos && lote.hasNeg)
      .map((lote) => ({
        Lote: lote.lote,
        SKUs_con_Sobrante: lote.sobrantes,
        SKUs_con_Faltante: lote.faltantes,
        Diferencia_Neta: lote.diferenciaNeta,
        Estado: 'Mixto',
      }));

    const resumenData = [
      { Concepto: 'STOCK TOTAL', Valor: '', Unidad: '' },
      { Concepto: '  - Stock Cliente', Valor: totalCliente, Unidad: 'Unidades' },
      { Concepto: '  - Stock WMS', Valor: totalWms, Unidad: 'Unidades' },
      { Concepto: '  - Diferencia Total', Valor: totalCliente - totalWms, Unidad: 'Unidades' },
      { Concepto: '', Valor: '', Unidad: '' },
      { Concepto: 'DIFERENCIAS EN UNIDADES', Valor: '', Unidad: '' },
      { Concepto: '  - Total Sobrante', Valor: totalSobranteUnidades, Unidad: 'Unidades' },
      { Concepto: '  - Total Faltante', Valor: totalFaltanteUnidades, Unidad: 'Unidades' },
      {
        Concepto: '  - Diferencia Neta',
        Valor: totalSobranteUnidades - totalFaltanteUnidades,
        Unidad: 'Unidades',
      },
      { Concepto: '', Valor: '', Unidad: '' },
      { Concepto: 'ANALISIS POR SKU', Valor: '', Unidad: '' },
      { Concepto: '  - SKUs con Sobrante', Valor: totalSKUsSobrantes, Unidad: 'SKUs' },
      { Concepto: '  - SKUs con Faltante', Valor: totalSKUsFaltantes, Unidad: 'SKUs' },
      { Concepto: '  - SKUs OK', Valor: totalSKUsOK, Unidad: 'SKUs' },
      { Concepto: '  - TOTAL SKUs', Valor: crossResults.length, Unidad: 'SKUs' },
      { Concepto: '', Valor: '', Unidad: '' },
      { Concepto: 'ANALISIS POR LOTE', Valor: '', Unidad: '' },
      { Concepto: '  - Lotes solo con Sobrante', Valor: lotesSoloSobrantes, Unidad: 'Lotes' },
      { Concepto: '  - Lotes solo con Faltante', Valor: lotesSoloFaltantes, Unidad: 'Lotes' },
      { Concepto: '  - Lotes Mixtos (Sobrante + Faltante)', Valor: lotesMixtos, Unidad: 'Lotes' },
      { Concepto: '  - TOTAL LOTES CON NOVEDAD', Valor: lotesConNovedad, Unidad: 'Lotes' },
      { Concepto: '  - Lotes OK', Valor: lotesOK, Unidad: 'Lotes' },
      { Concepto: '  - TOTAL LOTES ANALIZADOS', Valor: totalLotes, Unidad: 'Lotes' },
    ];

    const wb = XLSX.utils.book_new();

    const wsResumen = XLSX.utils.json_to_sheet(resumenData);
    wsResumen['!cols'] = [46, 18, 14].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsResumen, 'Resumen');

    // Orden: sobrantes (mayor a menor) → coincidencias → faltantes (mayor diferencia abs primero)
    const sortedResults = [...result.results].sort((a, b) => {
      const group = (d: number) => (d > 0 ? 0 : d === 0 ? 1 : 2);
      const ga = group(a.diferencia);
      const gb = group(b.diferencia);
      if (ga !== gb) return ga - gb;
      if (a.diferencia > 0) return b.diferencia - a.diferencia;
      if (a.diferencia < 0) return a.diferencia - b.diferencia;
      return 0;
    });

    const exportData = sortedResults.map((r) => ({
      SKU: r.sku,
      Lote: r.lote,
      Descripcion: r.descripcion,
      Cantidad_Cliente: r.cantidadCliente,
      Cantidad_WMS: r.cantidadWms,
      Diferencia: r.diferencia,
      Fecha_Vencimiento_Cliente: r.fechaVencimientoCliente,
      Fecha_Vencimiento_WMS: r.fechaVencimientoWms,
      Dias_Diferencia_Vencimiento: r.diferenciaVencimientoDias ?? '',
      Validacion_Vencimiento: r.validacionVencimiento,
      Estado: r.diferencia === 0 ? 'Coincide' : r.diferencia > 0 ? 'Sobra en WMS' : 'Falta en WMS',
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    ws['!cols'] = [18, 18, 42, 16, 16, 14, 22, 22, 18, 18, 18].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws, 'Cruce Cliente-WMS');

    // Hoja Movimientos: SKUs con diferencia (no están cuadrados)
    const movimientosData = sortedResults
      .filter((r) => r.diferencia !== 0)
      .map((r) => ({
        SKU: r.sku,
        Descripcion: r.descripcion,
        Lote: r.lote,
        'F.V': r.diferencia > 0 ? r.fechaVencimientoWms : r.fechaVencimientoCliente,
        Unidades: Math.abs(r.diferencia),
        'Estado de Calidad': '',
        Tipo: r.diferencia > 0 ? 'Sobrante' : 'Faltante',
      }));

    if (movimientosData.length > 0) {
      const wsMovimientos = XLSX.utils.json_to_sheet(movimientosData);
      wsMovimientos['!cols'] = [18, 42, 18, 14, 12, 18, 12].map((w) => ({ wch: w }));
      XLSX.utils.book_append_sheet(wb, wsMovimientos, 'Movimientos');
    }

    if (lotesMixtosDetalle.length > 0) {
      const wsMixtos = XLSX.utils.json_to_sheet(lotesMixtosDetalle);
      wsMixtos['!cols'] = [18, 18, 18, 16, 12].map((w) => ({ wch: w }));
      XLSX.utils.book_append_sheet(wb, wsMixtos, 'Lotes Mixtos');
    }

    const fecha = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `cruce_inventarios_${fecha}.xlsx`);

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

  getTotalItems(): number {
    return this.analysisResult()?.stats.totalItems || 0;
  }

  getItemsConDiferencia(): number {
    return this.analysisResult()?.stats.itemsConDiferencia || 0;
  }

  getCoincidencias(): number {
    return this.analysisResult()?.stats.coincidencias || 0;
  }

  getPorcentajeCoincidencia(): number {
    return this.analysisResult()?.stats.porcentajeCoincidencia || 0;
  }

  getDifferenceSeverity(diferencia: number): 'success' | 'danger' | 'warn' {
    if (diferencia === 0) return 'success';
    return diferencia > 0 ? 'warn' : 'danger';
  }

  getDifferenceLabel(diferencia: number): string {
    if (diferencia === 0) return 'Coincide';
    return diferencia > 0 ? 'Sobra en WMS' : 'Falta en WMS';
  }

  private processAnalysis(): CrossAnalysisResult {
    const clientRaw = this.clientFile()?.data;
    const wmsRaw = this.wmsFile()?.data;
    const clientName = this.clientFile()?.name || '';
    const wmsName = this.wmsFile()?.name || '';

    if (!clientRaw || !wmsRaw) {
      throw new Error('Archivos no validos');
    }

    const clientRows = this.parseFileData(clientRaw, clientName);
    const wmsRows = this.parseFileData(wmsRaw, wmsName);

    const sapData = clientRows.map((item) => ({
      sku: this.getNestedValue(item, this.clientHeaderMap['sku'] || 'Material', ''),
      lote: this.getNestedValue(item, this.clientHeaderMap['lote'] || 'Lote', ''),
      descripcion: this.getNestedValue(
        item,
        this.clientHeaderMap['descripcion'] || 'Descripcion',
        '',
      ),
      cantidad: this.toNumber(
        this.getNestedValue(item, this.clientHeaderMap['cantidad'] || 'Stock disponible', 0),
      ),
      fechaVencimiento: this.getNestedValue(
        item,
        this.clientHeaderMap['fechaVencimiento'] || 'Cad./FPC',
        '',
      ),
    }));

    const wmsData = wmsRows.map((item) => ({
      sku: this.getNestedValue(item, this.wmsHeaderMap['sku'] || 'SKU', ''),
      lote: this.getNestedValue(item, this.wmsHeaderMap['lote'] || 'Lote', ''),
      descripcion: this.getNestedValue(item, this.wmsHeaderMap['descripcion'] || 'Descripcion', ''),
      cantidad: this.toNumber(
        this.getNestedValue(item, this.wmsHeaderMap['cantidad'] || 'Disponible', 0),
      ),
      fechaVencimiento: this.getNestedValue(
        item,
        this.wmsHeaderMap['fechaVencimiento'] || 'Fecha de vencimiento',
        '',
      ),
    }));

    const crossed = this.runInventoryCross({
      sapData,
      wmsData,
      groupByLot: this.groupByLot(),
    });

    const totalItems = crossed.results.length;
    const itemsConDiferencia = crossed.results.filter((r) => r.diferencia !== 0).length;
    const coincidencias = totalItems - itemsConDiferencia;
    const porcentajeCoincidencia =
      totalItems === 0 ? 100 : Math.round((coincidencias / totalItems) * 100);
    const diferenciaNeta = crossed.results.reduce((sum, row) => sum + row.diferencia, 0);

    return {
      summary: `Cruce completado: ${totalItems} registros comparados, ${itemsConDiferencia} con diferencia.`,
      results: crossed.results,
      stats: {
        totalItems,
        itemsConDiferencia,
        coincidencias,
        porcentajeCoincidencia,
        diferenciaNeta,
      },
    };
  }

  private runInventoryCross(input: { sapData: any[]; wmsData: any[]; groupByLot: boolean }): {
    results: CrossRow[];
  } {
    const { sapData, wmsData, groupByLot } = input;
    const norm = (v: any) =>
      String(v || '')
        .trim()
        .toUpperCase();

    const sapAggregated = sapData.reduce(
      (acc, item) => {
        const sku = norm(item.sku);
        if (!sku) return acc;
        const lote = groupByLot ? norm(item.lote) : 'UNIFICADO';
        const key = `${sku}__${lote}`;
        if (!acc[key]) {
          acc[key] = {
            sku,
            lote,
            cantidad: 0,
            descripcion: item.descripcion,
            fechaVencimiento: item.fechaVencimiento || '',
          };
        }
        acc[key].cantidad += this.toNumber(item.cantidad);
        if (!acc[key].fechaVencimiento && item.fechaVencimiento) {
          acc[key].fechaVencimiento = item.fechaVencimiento;
        }
        return acc;
      },
      {} as Record<string, any>,
    );

    const wmsAggregated = wmsData.reduce(
      (acc, item) => {
        const sku = norm(item.sku);
        if (!sku) return acc;
        const lote = groupByLot ? norm(item.lote) : 'UNIFICADO';
        const key = `${sku}__${lote}`;
        if (!acc[key]) {
          acc[key] = {
            sku,
            lote,
            cantidad: 0,
            descripcion: item.descripcion,
            fechaVencimiento: item.fechaVencimiento || '',
          };
        }
        acc[key].cantidad += this.toNumber(item.cantidad);
        if (!acc[key].fechaVencimiento && item.fechaVencimiento) {
          acc[key].fechaVencimiento = item.fechaVencimiento;
        }
        return acc;
      },
      {} as Record<string, any>,
    );

    const allKeys = new Set([...Object.keys(sapAggregated), ...Object.keys(wmsAggregated)]);

    const results = Array.from(allKeys).map((key) => {
      const sap = sapAggregated[key];
      const wms = wmsAggregated[key];

      const sku = sap?.sku || wms?.sku;
      const lote = sap?.lote || wms?.lote;
      const descripcion = sap?.descripcion || wms?.descripcion || '';
      const cantidadCliente = this.toNumber(sap?.cantidad || 0);
      const cantidadWms = this.toNumber(wms?.cantidad || 0);
      const diferencia = cantidadWms - cantidadCliente;
      const fechaVencimientoCliente = this.formatDateForDisplay(sap?.fechaVencimiento || '');
      const fechaVencimientoWms = this.formatDateForDisplay(wms?.fechaVencimiento || '');
      const diffVenc = this.calculateDateDifferenceInDays(
        sap?.fechaVencimiento,
        wms?.fechaVencimiento,
      );
      const validacionVencimiento: CrossRow['validacionVencimiento'] =
        diffVenc === null ? 'Sin fecha' : diffVenc === 0 ? 'Coincide' : 'Diferente';

      return {
        sku,
        lote,
        descripcion,
        cantidadCliente,
        cantidadWms,
        diferencia,
        fechaVencimientoCliente,
        fechaVencimientoWms,
        diferenciaVencimientoDias: diffVenc,
        validacionVencimiento,
      };
    });

    results.sort(
      (a, b) => Math.abs(b.diferencia) - Math.abs(a.diferencia) || a.sku.localeCompare(b.sku),
    );

    return { results };
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

  private parseDateValue(value: any): Date | null {
    if (value === undefined || value === null || value === '') return null;

    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return new Date(value.getFullYear(), value.getMonth(), value.getDate());
    }

    if (typeof value === 'number' && Number.isFinite(value)) {
      return this.parseExcelSerialDate(value);
    }

    const raw = String(value).trim();
    if (!raw) return null;

    const numeric = Number(raw.replace(',', '.'));
    if (Number.isFinite(numeric) && /^\d+(?:[\.,]\d+)?$/.test(raw)) {
      return this.parseExcelSerialDate(numeric);
    }

    const isoMatch = raw.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
    if (isoMatch) {
      const [, year, month, day] = isoMatch;
      return this.buildDate(Number(year), Number(month), Number(day));
    }

    const latinMatch = raw.match(/^(\d{1,2})[\.\/-](\d{1,2})[\.\/-](\d{2,4})/);
    if (latinMatch) {
      const [, part1, part2, part3] = latinMatch;
      const first = Number(part1);
      const second = Number(part2);
      const year = Number(part3.length === 2 ? `20${part3}` : part3);

      // Si el primer número es > 12 es definitivamente el día (DD/MM/YYYY)
      // Si el segundo número es > 12 es definitivamente el día (MM/DD/YYYY → raro en LATAM)
      // En ambigüedad (ambos ≤ 12) se asume DD/MM/YYYY (convención LATAM)
      if (second > 12) {
        // MM/DD/YYYY → month=first, day=second
        return this.buildDate(year, first, second);
      }
      // DD/MM/YYYY (incluye caso ambiguo: ambos ≤ 12)
      return this.buildDate(year, second, first);
    }

    const parsed = new Date(raw);
    if (!Number.isNaN(parsed.getTime())) {
      return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
    }

    return null;
  }

  private formatDateForDisplay(value: any): string {
    const date = this.parseDateValue(value);
    if (!date) return '';
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  private calculateDateDifferenceInDays(clientValue: any, wmsValue: any): number | null {
    const clientDate = this.parseDateValue(clientValue);
    const wmsDate = this.parseDateValue(wmsValue);
    if (!clientDate || !wmsDate) return null;

    const start = Date.UTC(clientDate.getFullYear(), clientDate.getMonth(), clientDate.getDate());
    const end = Date.UTC(wmsDate.getFullYear(), wmsDate.getMonth(), wmsDate.getDate());
    return Math.round((end - start) / (1000 * 60 * 60 * 24));
  }

  private parseExcelSerialDate(value: number): Date | null {
    if (!Number.isFinite(value)) return null;

    const days = Math.floor(value);
    const excelEpochUtc = Date.UTC(1899, 11, 30);
    const utcTime = excelEpochUtc + days * 24 * 60 * 60 * 1000;
    const utcDate = new Date(utcTime);

    if (Number.isNaN(utcDate.getTime())) return null;

    return new Date(utcDate.getUTCFullYear(), utcDate.getUTCMonth(), utcDate.getUTCDate());
  }

  private buildDate(year: number, month: number, day: number): Date | null {
    if (!year || !month || !day) return null;
    const date = new Date(year, month - 1, day);
    if (
      Number.isNaN(date.getTime()) ||
      date.getFullYear() !== year ||
      date.getMonth() !== month - 1 ||
      date.getDate() !== day
    ) {
      return null;
    }
    return date;
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
