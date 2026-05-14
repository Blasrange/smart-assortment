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
import { PrintModal } from './modal/print-modal';

interface FileData {
  name: string;
  data: string | ArrayBuffer | null;
  size: number;
  uploadDate: Date;
}

interface StoreRecord {
  Codigo: string;
  Tienda: string;
  Direccion: string;
  Ciudad: string;
  Departamento: string;
}

interface ExitoLabelData {
  nc: string;
  ct: string;
  codigoBarra: string;
  tienda: string;
  depto: string;
  ciudad: string;
  orden: string;
  direccion: string;
  numeroCaja: number;
  totalCajas: number;
  cedi: string;
  descripcion: string;
}

interface ExitoLabelsResult {
  labels: ExitoLabelData[];
  warnings: string[];
  summary: string;
  stats: {
    sourceRows: number;
    labelsCount: number;
    uniqueStores: number;
  };
}

const exitoLabelsColumnMapping = {
  ocMarker: ['OC', 'O/C', 'ORDEN DE COMPRA', 'ORDEN COMPRA', 'NRO OC'],
  barcode: [
    'COD. BARRA',
    'COD BARRA',
    'CODIGO BARRA',
    'CODIGO DE BARRAS',
    'CODBARRA',
    'BARCODE',
    'EAN',
    'GTIN',
  ],
  dependencia: [
    'DEPENDENCIAS',
    'DEPENDENCIA',
    'COD DEPENDENCIA',
    'CODIGO DEPENDENCIA',
    'DEP',
    'TIENDA CODIGO',
  ],
  tienda: [
    'DESC. ITEM',
    'DESC ITEM',
    'DESCRIPCION',
    'DESCRIPCION ITEM',
    'TIENDA',
    'NOMBRE TIENDA',
    'DESTINO',
  ],
  cajas: ['CJ/UN', 'CJ UN', 'CAJA/UN', 'CAJAS/UNIDADES', 'CAJAS UNIDADES'],
  cantidadInterna: ['CANT', 'QTY', 'UNIDADES', 'CANTIDAD'],
};

const STORES_DATA: StoreRecord[] = [
  {
    Codigo: '0217',
    Tienda: '217B GUITA S MAX SAN LUIS SUBA',
    Direccion: 'TRANSVERSAL 91 # 128D - 34',
    Ciudad: 'BOGOTA D.C.',
    Departamento: 'BOGOTA D.C.',
  },
  {
    Codigo: '0224',
    Tienda: '224B GUITA S MAX BRITALIA',
    Direccion: 'CALLE 46 # 81 - 102SUR',
    Ciudad: 'BOGOTA D.C.',
    Departamento: 'BOGOTA D.C.',
  },
  {
    Codigo: '0225',
    Tienda: '225B GUITA S MAX SANTA ANA MOSQUERA',
    Direccion: 'CARRERA 10 # 10 - 04',
    Ciudad: 'BOGOTA D.C.',
    Departamento: 'BOGOTA D.C.',
  },
  {
    Codigo: '0333',
    Tienda: '333B GUITA S MAX LA CEJA',
    Direccion: 'CALLE 19 # 19-20',
    Ciudad: 'LA CEJA',
    Departamento: 'ANTIOQUIA',
  },
  {
    Codigo: '0376',
    Tienda: 'ALMACENES EXITO BOSA 376',
    Direccion: 'CL 65 SUR78H 54',
    Ciudad: 'BOGOTA D.C.',
    Departamento: 'BOGOTA D.C.',
  },
  {
    Codigo: '0020',
    Tienda: 'ALMACENES EXITO S.A.',
    Direccion: 'CRA 24A # 163A -28',
    Ciudad: 'ENVIGADO',
    Departamento: 'ANTIOQUIA',
  },
  {
    Codigo: '0085',
    Tienda: 'CEDI FUNZA 085',
    Direccion: 'KM4 VIA FUNZA PARQ INDT SAN C',
    Ciudad: 'BOGOTA D.C.',
    Departamento: 'BOGOTA D.C.',
  },
  {
    Codigo: '0050',
    Tienda: 'CEDI CALI 050',
    Direccion: 'CR 36 A 16 79 ACOPI',
    Ciudad: 'CALI',
    Departamento: 'VALLE DEL CAUCA',
  },
  {
    Codigo: '0146',
    Tienda: 'CEDI CARIBE 146',
    Direccion: 'PARQUE INDUSTR MALAMBO PIMSA KILOMETRO 3 VIA MALAMBO',
    Ciudad: 'BARRANQUILLA',
    Departamento: 'ATLANTICO',
  },
  {
    Codigo: '0091',
    Tienda: 'EXITO OCCIDENTE 091',
    Direccion: 'CR 114A # 78B -85 AUTOPISTA MEDELLIN',
    Ciudad: 'BOGOTA D.C.',
    Departamento: 'BOGOTA D.C.',
  },
];

const normalizeStoreCode = (value: unknown): string => {
  const raw = String(value ?? '').trim();
  const digits = raw.replace(/\D/g, '');
  if (!digits) return raw;
  return digits.padStart(4, '0');
};

const storeByCode = new Map<string, StoreRecord>(
  STORES_DATA.map((store) => [normalizeStoreCode(store.Codigo), store]),
);

const normalizeHeader = (value: unknown): string =>
  String(value ?? '')
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Z0-9]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

const clean = (value: string): string =>
  String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[\^~]/g, ' ')
    .trim();

@Component({
  selector: 'app-exito-labels',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    TableModule,
    ButtonModule,
    InputTextModule,
    PrintModal,
    ProgressSpinnerModule,
    TooltipModule,
    BreadcrumbModule,
    ToastModule,
  ],
  providers: [MessageService],
  templateUrl: './exito-labels.html',
  styleUrl: './exito-labels.css',
})
export class ExitoLabels {
  @ViewChild('dt1') dt1!: Table;

  sourceFile = signal<FileData | null>(null);
  analysisResult = signal<ExitoLabelsResult | null>(null);
  loading = signal(false);
  isDragging = signal(false);
  showPrintModal = false;
  printZplBatch = '';
  searchValue = '';

  // Configuración para conversión de unidades a cajas
  UNIDADES_POR_CAJA = 24; // Cambia este valor según tu necesidad (1 caja = 24 unidades por defecto)

  readonly breadcrumbItems: MenuItem[] = [
    { label: 'Etiquetas Exito', routerLink: '/exito-labels' },
  ];
  readonly breadcrumbHome = { icon: 'pi pi-home', label: 'Inicio', routerLink: '/' };

  readonly isBrowser: boolean;

  constructor(
    @Inject(PLATFORM_ID) platformId: Object,
    private messageService: MessageService,
    private primeng: PrimeNG,
  ) {
    this.isBrowser = isPlatformBrowser(platformId);

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

  get filteredLabels(): ExitoLabelData[] {
    const result = this.analysisResult();
    if (!result) return [];

    const q = this.searchValue.trim().toLowerCase();
    if (!q) return result.labels;

    return result.labels.filter((label) =>
      [label.ct, label.tienda, label.codigoBarra, label.orden, label.descripcion].some((v) =>
        String(v || '')
          .toLowerCase()
          .includes(q),
      ),
    );
  }

  get warningCount(): number {
    return this.analysisResult()?.warnings.length || 0;
  }

  onDragOver(event: DragEvent): void {
    event.preventDefault();
    this.isDragging.set(true);
  }

  onDragLeave(event: DragEvent): void {
    event.preventDefault();
    this.isDragging.set(false);
  }

  onDrop(event: DragEvent): void {
    event.preventDefault();
    this.isDragging.set(false);
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
    this.analysisResult.set(null);
    this.searchValue = '';
    this.showPrintModal = false;
    this.printZplBatch = '';

    this.messageService.add({
      severity: 'info',
      summary: 'Archivo removido',
      detail: 'Se ha eliminado el archivo de etiquetas',
      life: 3000,
    });
  }

  clearFilters(): void {
    this.searchValue = '';
    if (this.dt1) {
      this.dt1.clear();
      this.dt1.filterGlobal('', 'contains');
    }

    this.messageService.add({
      severity: 'info',
      summary: 'Filtros limpiados',
      detail: 'Se han eliminado todos los filtros',
      life: 3000,
    });
  }

  processLabels(): void {
    if (!this.sourceFile()) {
      this.messageService.add({
        severity: 'warn',
        summary: 'Archivo faltante',
        detail: 'Primero carga el archivo de Exito para procesar',
        life: 4000,
      });
      return;
    }

    const data = this.sourceFile()?.data || null;
    const fileName = this.sourceFile()?.name || '';
    this.loading.set(true);

    try {
      const rows = this.parseRows(data, fileName);
      const result = this.runProcess(rows);
      this.analysisResult.set(result);

      this.messageService.add({
        severity: 'success',
        summary: 'Proceso completado',
        detail: result.summary,
        life: 3500,
      });

      if (result.warnings.length > 0) {
        this.messageService.add({
          severity: 'warn',
          summary: 'Advertencias',
          detail: result.warnings[0],
          life: 7000,
        });
      }
    } catch (error) {
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: (error as Error)?.message || 'No fue posible procesar el archivo',
        life: 6000,
      });
    } finally {
      this.loading.set(false);
    }
  }

  exportToExcel(): void {
    const result = this.analysisResult();
    if (!result || !result.labels.length) return;

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(
      result.labels.map((label) => ({
        CODIGO_TIENDA: label.ct,
        TIENDA: label.tienda,
        CODIGO_BARRA: label.codigoBarra,
        ORDEN_COMPRA: label.orden,
        CEDI: label.cedi,
        DESCRIPCION: label.descripcion,
        CAJA: label.numeroCaja,
        TOTAL_CAJAS: label.totalCajas,
        DIRECCION: label.direccion,
        CIUDAD: label.ciudad,
        DEPARTAMENTO: label.depto,
      })),
    );
    XLSX.utils.book_append_sheet(wb, ws, 'Etiquetas');

    const date = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `etiquetas_exito_${date}.xlsx`);

    this.messageService.add({
      severity: 'success',
      summary: 'Exportado',
      detail: 'Archivo Excel generado correctamente',
      life: 3000,
    });
  }

  exportZpl(): void {
    const result = this.analysisResult();
    if (!result || !result.labels.length || !this.isBrowser) return;

    const zpl = this.generateBatchZpl(result.labels);
    const blob = new Blob([zpl], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const date = new Date().toISOString().slice(0, 10);
    a.href = url;
    a.download = `etiquetas_exito_${date}.zpl`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    this.messageService.add({
      severity: 'success',
      summary: 'ZPL generado',
      detail: 'Archivo ZPL descargado correctamente',
      life: 3000,
    });
  }

  openPrintModal(): void {
    const result = this.analysisResult();
    if (!result?.labels.length) return;
    this.printZplBatch = this.generateBatchZpl(result.labels);
    this.showPrintModal = true;
  }

  onPrintDone(event: { success: boolean; message: string }): void {
    this.messageService.add({
      severity: event.success ? 'success' : 'error',
      summary: event.success ? 'Impresión completada' : 'Error de impresión',
      detail: event.message,
      life: event.success ? 4000 : 6000,
    });

    if (event.success) {
      this.showPrintModal = false;
    }
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

  private processFile(file: File): void {
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
    reader.onload = (e: ProgressEvent<FileReader>) => {
      try {
        const data = e.target?.result || null;
        const rows = this.parseRows(data, file.name);
        if (!rows.length) {
          this.messageService.add({
            severity: 'error',
            summary: 'Sin datos',
            detail: 'El archivo no contiene filas válidas',
            life: 5000,
          });
          return;
        }

        this.sourceFile.set({
          name: file.name,
          data,
          size: file.size,
          uploadDate: new Date(),
        });

        this.messageService.add({
          severity: 'success',
          summary: 'Archivo cargado',
          detail: `Etiquetas Exito: ${rows.length} filas encontradas`,
          life: 3000,
        });
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

  private parseRows(data: string | ArrayBuffer | null, fileName: string): any[][] {
    if (!data) return [];

    const ext = fileName.split('.').pop()?.toLowerCase();
    if (ext === 'xlsx' || ext === 'xls') {
      const workbook = XLSX.read(data, { type: 'binary' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' }) as any[][];
    }

    const text = String(data);
    const lines = text.split(/\r?\n/).filter((line) => line.trim().length > 0);
    if (!lines.length) return [];

    const delimiter = lines[0].includes(';') ? ';' : ',';
    return lines.map((line) => line.split(delimiter).map((v) => v.trim()));
  }

  /**
   * Parsea el campo CJ/UN y lo convierte a número de cajas
   * Soporta formatos como:
   * - "2 CJ" -> 2 cajas
   * - "24 UN" -> 1 caja (si 1 caja = 24 unidades)
   * - "48 UN" -> 2 cajas
   * - "3" -> 3 cajas (asume que son cajas si solo es número)
   * - "2.5 CJ" -> 2.5 cajas
   */
  private parseQuantityToBoxes(quantityStr: string): number {
    const normalized = quantityStr.trim().toUpperCase();
    if (!normalized) return 0;

    // Buscar patrón: número seguido de CJ (cajas)
    const cajasMatch = normalized.match(/(\d+(?:[.,]\d+)?)\s*CJ/i);
    if (cajasMatch) {
      // Ya viene expresado en cajas: redondear hacia arriba para asegurar etiquetas completas.
      return Math.ceil(parseFloat(cajasMatch[1].replace(',', '.')));
    }

    // Buscar patrón: número seguido de UN (unidades)
    const unidadesMatch = normalized.match(/(\d+(?:[.,]\d+)?)\s*UN/i);
    if (unidadesMatch) {
      // Convertir unidades a cajas
      const unidades = parseFloat(unidadesMatch[1].replace(',', '.'));
      return Math.ceil(unidades / this.UNIDADES_POR_CAJA);
    }

    // Si solo es número (sin especificar CJ o UN), asumir que son cajas
    const soloNumero = parseFloat(normalized.replace(',', '.'));
    if (!isNaN(soloNumero)) {
      return Math.ceil(soloNumero);
    }

    // Si no se pudo parsear, retornar 0
    return 0;
  }

  private runProcess(rows: any[][]): ExitoLabelsResult {
    const norm = (v: any) => String(v ?? '').trim();
    const ocCandidates = exitoLabelsColumnMapping.ocMarker.map(normalizeHeader);
    const barcodeCandidates = exitoLabelsColumnMapping.barcode.map(normalizeHeader);
    const dependenciaCandidates = exitoLabelsColumnMapping.dependencia.map(normalizeHeader);

    let ordenCompra = '';
    let cedi = '';

    for (const row of rows) {
      const ocIndex = row.findIndex((v: any) => ocCandidates.includes(normalizeHeader(v)));
      if (ocIndex >= 0) {
        ordenCompra = norm(row[ocIndex + 1]);
        cedi = norm(row[ocIndex + 2]);
        break;
      }
    }

    let headerIndex = -1;
    for (let i = 0; i < rows.length; i += 1) {
      const row = rows[i];
      const hasBarcode = row.some((v: any) => barcodeCandidates.includes(normalizeHeader(v)));
      const hasDependencia = row.some((v: any) =>
        dependenciaCandidates.includes(normalizeHeader(v)),
      );
      if (hasBarcode && hasDependencia) {
        headerIndex = i;
        break;
      }
    }

    if (headerIndex < 0) {
      throw new Error('No se encontro la cabecera de detalle en el archivo.');
    }

    const header = rows[headerIndex];
    const findIndex = (candidates: string[]): number => {
      const normalizedCandidates = candidates.map(normalizeHeader);
      for (let i = 0; i < header.length; i += 1) {
        if (normalizedCandidates.includes(normalizeHeader(header[i]))) {
          return i;
        }
      }
      return -1;
    };

    const barcodeIdx = findIndex(exitoLabelsColumnMapping.barcode);
    const dependenciaIdx = findIndex(exitoLabelsColumnMapping.dependencia);
    const tiendaIdx = findIndex(exitoLabelsColumnMapping.tienda);
    const cajasIdx = findIndex(exitoLabelsColumnMapping.cajas);
    const cantidadInternaIdx = findIndex(exitoLabelsColumnMapping.cantidadInterna);

    if ([barcodeIdx, dependenciaIdx, cajasIdx].includes(-1)) {
      throw new Error('No se encontraron columnas requeridas: COD. BARRA, DEPENDENCIAS y CJ/UN.');
    }

    const labels: ExitoLabelData[] = [];
    const warnings: string[] = [];
    let currentBarcode = '';
    let currentDescription = '';

    for (let i = headerIndex + 1; i < rows.length; i += 1) {
      const row = rows[i] || [];
      const barcodeCell = norm(row[barcodeIdx] ?? '');
      const dependenciaCell = norm(row[dependenciaIdx] ?? '');
      const tiendaCell = tiendaIdx >= 0 ? norm(row[tiendaIdx] ?? '') : '';
      const cajasRaw = norm(row[cajasIdx] ?? '');
      const cantidadInternaRaw = cantidadInternaIdx >= 0 ? norm(row[cantidadInternaIdx] ?? '') : '';

      const isCompletelyEmpty =
        !barcodeCell && !dependenciaCell && !tiendaCell && !cajasRaw && !cantidadInternaRaw;
      if (isCompletelyEmpty) continue;

      // Si hay código de barras, actualizar el actual y continuar
      if (barcodeCell) {
        currentBarcode = barcodeCell;
        currentDescription = tiendaCell || currentDescription;
        continue;
      }

      // Si no hay dependencia o no hay cantidad válida, continuar
      if (!dependenciaCell) continue;

      // Parsear la cantidad a cajas usando la nueva función
      const totalCajas = this.parseQuantityToBoxes(cajasRaw);

      // Si no hay cajas válidas, mostrar advertencia y continuar
      if (totalCajas <= 0) {
        warnings.push(
          `Fila ${i + 1}: Valor CJ/UN no válido "${cajasRaw}" para dependencia ${dependenciaCell}`,
        );
        continue;
      }

      const dependenciaCode = normalizeStoreCode(dependenciaCell);
      const store = storeByCode.get(dependenciaCode);
      const tienda = store?.Tienda || tiendaCell || 'Localizacion No Registrada';
      const direccion = store?.Direccion || 'N/A';
      const ciudad = store?.Ciudad || 'N/A';
      const departamento = store?.Departamento || 'N/A';

      // Verificar que el código de tienda existe en la base de datos
      if (!store) {
        warnings.push(
          `Fila ${i + 1}: Código de tienda "${dependenciaCell}" no encontrado en la base de datos`,
        );
      }

      // Generar una etiqueta por cada caja
      for (let caja = 1; caja <= totalCajas; caja += 1) {
        labels.push({
          nc: 'Corporacion Colombiana de Logistica',
          ct: dependenciaCode,
          codigoBarra: currentBarcode,
          tienda,
          depto: departamento,
          ciudad,
          orden: ordenCompra,
          direccion,
          numeroCaja: caja,
          totalCajas,
          cedi,
          descripcion: currentDescription || 'SIN DESCRIPCION',
        });
      }
    }

    if (!labels.length) {
      warnings.push('No se encontraron filas de tiendas con cantidad para generar etiquetas.');
    }

    const uniqueStores = new Set(labels.map((l) => l.ct)).size;
    const summary = `Generadas ${labels.length} etiqueta(s) para ${uniqueStores} tienda(s)`;

    return {
      labels,
      warnings,
      summary,
      stats: {
        sourceRows: rows.length,
        labelsCount: labels.length,
        uniqueStores,
      },
    };
  }

  private generateLabelZpl(label: ExitoLabelData): string {
    const nc = clean(label.nc);
    const ct = clean(label.ct);
    const codigoBarra = clean(label.codigoBarra);
    const tienda = clean(label.tienda);
    const depto = clean(label.depto);
    const ciudad = clean(label.ciudad);
    const orden = clean(label.orden);
    const direccion = clean(label.direccion);
    const numeroCaja = String(label.numeroCaja || 1);
    const totalCajas = String(label.totalCajas || 1);
    const descripcion = clean(label.descripcion);

    const z: string[] = [];
    z.push('^XA');
    z.push('^CF0,27');
    z.push(`^FO200,20^FD${nc}^FS`);
    z.push('^FO50,45^GB700,3,3^FS');
    z.push(`^FO270,60^BY4,2,190^BCN,140,N,N,N^FD>;${ct}^FS`);
    z.push(`^CF0,40^FO340,210^FD${ct}^FS`);
    z.push('^FO50,246^GB700,3,3^FS');
    z.push(`^CF0,25^FO50,255^FDTIENDA:^FS^FO140,255^FD${tienda}^FS`);
    z.push(
      `^FO50,315^FDDEPARTAMENTO:^FS^FO230,315^FD${depto}^FS^FO435,315^FDCIUDAD:^FS^FO530,315^FD${ciudad}^FS`,
    );
    z.push(`^CF0,25^FO210,345^FDORD COMPRA:^FS^FO370,345^FD${orden}^FS`);
    z.push(`^CF0,25^FO50,285^FDDIRECCION:^FS^FO175,285^FD${direccion}^FS`);
    z.push('^FO50,369^GB700,3,3^FS');
    z.push(`^CF0,25^FO50,345^FDCAJAS:^FS^FO130,345^FD${numeroCaja} de ${totalCajas}^FS`);
    z.push('^CF0,25^FO540,345^FDALMACENES EXITO^FS');
    z.push(`^CF0,25^FO50,378^FD${descripcion}^FS`);
    z.push(`^CF0,22^FO50,408^FDCOD BARRA:^FS^FO200,408^FD${codigoBarra}^FS`);
    z.push('^CF0,25^FO540,400^FDCEDI VEGAS^FS');
    z.push('^XZ');

    return z.join('\n');
  }

  private generateBatchZpl(labels: ExitoLabelData[]): string {
    return labels.map((label) => this.generateLabelZpl(label)).join('\n');
  }
}
