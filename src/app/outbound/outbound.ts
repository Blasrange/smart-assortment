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

interface OutboundEntry {
  ORDER: string;
  ORDER2: string;
  PURCHASE_ORDER: string;
  INVOICE: string;
  TRACKING: string;
  ORDER_DATE: string;
  SERVICE_DATE: string;
  SERVICE_DATE_MIN: string;
  SERVICE_DATE_MAX: string;
  OUTBOUNDTYPE_CODE: string;
  PRIORITY: number;
  NOTE: string;
  CARRIER_CODE: string;
  CUSTOMER_CODE: string;
  STORE_CODE: string;
  CITY_STORE_CODE: string;
  STORE_ADDRESS: string;
  STORE_PHONE: string;
  STORE2_CODE: string;
  STORE2_CITY_CODE: string;
  STORE2_ADDRESS: string;
  STORE2_PHONE: string;
  SKU: string;
  LOT: string;
  CREATED_DATE: string;
  EXPIRY_DATE: string;
  SERIAL: string;
  QUALITY_STATE: string;
  QTY: number;
  PRICE: number;
  TAXES: number;
  UOM_CODE: string;
  REFERENCE: string;
}

type OutboundFieldType = 'text' | 'number' | 'date';

interface OutboundField {
  key: keyof OutboundEntry;
  label: string;
  type: OutboundFieldType;
}

interface OutboundResult {
  results: OutboundEntry[];
  summary: string;
  stats: {
    orderRows: number;
    storeRows: number;
    totalOrders: number;
    totalStores: number;
    matchedRows: number;
    totalQty: number;
  };
}

interface OrderSourceField {
  key: string;
  label: string;
  type: OutboundFieldType;
}

interface StoreSourceField {
  key: string;
  label: string;
  type: OutboundFieldType;
}

const TARGET_FIELDS: OutboundField[] = [
  { key: 'ORDER', label: 'ORDER', type: 'text' },
  { key: 'ORDER2', label: 'ORDER2', type: 'text' },
  { key: 'PURCHASE_ORDER', label: 'PURCHASE_ORDER', type: 'text' },
  { key: 'INVOICE', label: 'INVOICE', type: 'text' },
  { key: 'TRACKING', label: 'TRACKING', type: 'text' },
  { key: 'ORDER_DATE', label: 'ORDER_DATE', type: 'date' },
  { key: 'SERVICE_DATE', label: 'SERVICE_DATE', type: 'date' },
  { key: 'SERVICE_DATE_MIN', label: 'SERVICE_DATE_MIN', type: 'date' },
  { key: 'SERVICE_DATE_MAX', label: 'SERVICE_DATE_MAX', type: 'date' },
  { key: 'OUTBOUNDTYPE_CODE', label: 'OUTBOUNDTYPE_CODE', type: 'text' },
  { key: 'PRIORITY', label: 'PRIORITY', type: 'number' },
  { key: 'NOTE', label: 'NOTE', type: 'text' },
  { key: 'CARRIER_CODE', label: 'CARRIER_CODE', type: 'text' },
  { key: 'CUSTOMER_CODE', label: 'CUSTOMER_CODE', type: 'text' },
  { key: 'STORE_CODE', label: 'STORE_CODE', type: 'text' },
  { key: 'CITY_STORE_CODE', label: 'CITY_STORE_CODE', type: 'text' },
  { key: 'STORE_ADDRESS', label: 'STORE_ADDRESS', type: 'text' },
  { key: 'STORE_PHONE', label: 'STORE_PHONE', type: 'text' },
  { key: 'STORE2_CODE', label: 'STORE2_CODE', type: 'text' },
  { key: 'STORE2_CITY_CODE', label: 'STORE2_CITY_CODE', type: 'text' },
  { key: 'STORE2_ADDRESS', label: 'STORE2_ADDRESS', type: 'text' },
  { key: 'STORE2_PHONE', label: 'STORE2_PHONE', type: 'text' },
  { key: 'SKU', label: 'SKU', type: 'text' },
  { key: 'LOT', label: 'LOT', type: 'text' },
  { key: 'CREATED_DATE', label: 'CREATED_DATE', type: 'date' },
  { key: 'EXPIRY_DATE', label: 'EXPIRY_DATE', type: 'date' },
  { key: 'SERIAL', label: 'SERIAL', type: 'text' },
  { key: 'QUALITY_STATE', label: 'QUALITY_STATE', type: 'text' },
  { key: 'QTY', label: 'QTY', type: 'number' },
  { key: 'PRICE', label: 'PRICE', type: 'number' },
  { key: 'TAXES', label: 'TAXES', type: 'number' },
  { key: 'UOM_CODE', label: 'UOM_CODE', type: 'text' },
  { key: 'REFERENCE', label: 'REFERENCE', type: 'text' },
];

const ORDER_SOURCE_FIELDS: OrderSourceField[] = [
  { key: 'EAN_LOC', label: 'EAN / LOC', type: 'text' },
  { key: 'PROVEEDOR', label: 'PROVEEDOR', type: 'text' },
  { key: 'CR', label: 'CR', type: 'text' },
  { key: 'LOC', label: 'LOC', type: 'text' },
  { key: 'STORE_NAME', label: 'NOMBRE TIENDA', type: 'text' },
  { key: 'ITEM', label: 'ITEM', type: 'text' },
  { key: 'DESCRIPTION', label: 'DESCRIPCIÓN', type: 'text' },
  { key: 'UNITS', label: 'UNIDADES', type: 'number' },
  { key: 'TRAYS', label: 'BANDEJAS / CARTONES', type: 'number' },
  { key: 'UNITS_PER_PACKAGE', label: 'UNIDADES POR EMBALAJE', type: 'number' },
  { key: 'SHIPPING_PACKAGE', label: 'EMBALAJE DE DESPACHO', type: 'text' },
  { key: 'DELIVERY_DATE', label: 'FECHA DE ENTREGA', type: 'date' },
  { key: 'ORDER_NUMBER', label: 'NÚMERO DE PEDIDO', type: 'text' },
  { key: 'TUTA', label: 'TUTA', type: 'text' },
  { key: 'ROUTE', label: 'RUTA', type: 'text' },
  { key: 'QUALITY_STATE', label: 'ESTADO DE CALIDAD', type: 'text' },
];

const STORE_SOURCE_FIELDS: StoreSourceField[] = [
  { key: 'NIT', label: 'NIT (CLIENTE)', type: 'text' },
  { key: 'NAME', label: 'NOMBRE', type: 'text' },
  { key: 'CONTACT', label: 'CONTACTO', type: 'text' },
  { key: 'CITY', label: 'CIUDAD', type: 'text' },
  { key: 'ADDRESS', label: 'DIRECCIÓN', type: 'text' },
  { key: 'PHONE_1', label: 'TELÉFONO 1', type: 'text' },
  { key: 'PHONE_2', label: 'TELÉFONO 2', type: 'text' },
  { key: 'EMAIL', label: 'CORREO ELECTRÓNICO', type: 'text' },
  { key: 'CODE', label: 'CÓDIGO (CR)', type: 'text' },
  { key: 'CODE2', label: 'CÓDIGO 2', type: 'text' },
  { key: 'LATITUDE', label: 'LATITUD', type: 'text' },
  { key: 'LONGITUDE', label: 'LONGITUD', type: 'text' },
  { key: 'NEIGHBORHOOD', label: 'BARRIO', type: 'text' },
  { key: 'REGION', label: 'REGIÓN', type: 'text' },
];

const ORDER_FIELD_ALIASES: Record<string, string[]> = {
  SKU: [
    'ITEM',
    'EAN - LOC',
    'EAN',
    'SKU',
    'CODIGO',
    'CODIGO PRODUCTO',
    'COD PRODUCTO',
    'ARTICULO',
    'PRODUCTO',
    'CODIGO EAN',
    'REFERENCIA',
    'BARCODE',
    'CODIGO ARTICULO',
  ],
  LOT: ['LOTE', 'LOT', 'BATCH', 'LOTE NUMERO', 'NRO LOTE', 'NUMERO LOTE', 'PARTIDA', 'CODIGO LOTE'],
  QTY: [
    'UNIDADES',
    'UNITS',
    'CANTIDAD',
    'QTY',
    'CANT',
    'CANTIDAD TOTAL',
    'TOTAL UNIDADES',
    'CANT PEDIDA',
    'UNIDADES PEDIDAS',
  ],
  ORDER_NUMBER: [
    'PEDIDO',
    'ORDER',
    'NUMERO PEDIDO',
    'DOCUMENTO',
    'NRO PEDIDO',
    'ORDER NUMBER',
    'NUMERO ORDEN',
    'ORDEN',
  ],
  DELIVERY_DATE: [
    'FECHA DE ENTREGA',
    'FECHA ENTREGA',
    'FECHA',
    'FECHA DESPACHO',
    'DELIVERY DATE',
    'FECHA ENVIO',
    'FECHA DISTRIBUCION',
  ],
  STORE_NAME: [
    'STORE_NAME',
    'NOMBRE TIENDA',
    'TIENDA',
    'NOMBRE STORE',
    'NOMBRE SUCURSAL',
    'SUCURSAL',
  ],
  CR: [
    'CR',
    'CODIGO TIENDA',
    'DG.',
    'CODIGO',
    'COD TIENDA',
    'CODIGO SUCURSAL',
    'SUCURSAL',
    'CODIGO CLIENTE',
    'CLIENTE',
    'TIENDA CODE',
    'CODIGO TIENDAS',
  ],
  LOC: ['LOC', 'CODIGO LOC', 'LOCACION', 'LOCALIDAD', 'CODIGO LOCACION'],
  PROVEEDOR: ['PROVEEDOR', 'PROVIDER', 'SUPPLIER', 'CODIGO PROVEEDOR', 'PROV', 'PROVEDOR'],
  ITEM: ['ITEM', 'SKU', 'EAN - LOC', 'EAN', 'ITEM CODE', 'CODIGO ITEM', 'ITEM NUMERO'],
  UNITS: [
    'UNIDADES',
    'UNITS',
    'CANTIDAD',
    'CANT',
    'CANTIDAD TOTAL',
    'TOTAL UNIDADES',
    'CANT PEDIDA',
  ],
  ROUTE: ['RUTA', 'ROUTE', 'TUTA', 'CODIGO RUTA', 'NUMERO RUTA', 'RUTA ENTREGA', 'EXPORTACION'],
  QUALITY_STATE: [
    'BODEGA',
    'ESTADO',
    'CALIDAD',
    'CALIDAD ESTADO',
    'ESTADO CALIDAD',
    'QUALITY',
    'ESTADO BODEGA',
  ],
  NOTE: ['NOTA', 'NOTE', 'NOTAS', 'OBSERVACION', 'OBSERVACIONES', 'COMENTARIO', 'REMARKS'],
  SERIAL: ['SERIAL', 'NUMERO SERIE', 'SERIE', 'SERIAL NUMBER', 'NSN', 'CODIGO SERIE'],
  INVOICE: ['FACTURA', 'INVOICE', 'NUMERO FACTURA', 'NRO FACTURA', 'FACTURA NUMERO'],
};

const STORE_FIELD_ALIASES: Record<string, string[]> = {
  CODE: [
    'Codigo',
    'CODE',
    'CÓDIGO',
    'CR',
    'CODIGO',
    'COD TIENDA',
    'CODIGO TIENDA',
    'SUCURSAL',
    'NUMERO SUCURSAL',
    'CODIGO SUCURSAL',
    'CODIGO CLIENTE',
    'CLIENTE',
  ],
  NAME: [
    'NOMBRE',
    'NAME',
    'TIENDA',
    'NOMBRE TIENDA',
    'NOMBRE SUCURSAL',
    'RAZON SOCIAL',
    'DENOMINACION',
  ],
  ADDRESS: ['DIRECCION', 'ADDRESS', 'DIRECCIÓN', 'DOMICILIO', 'CALLE', 'DIRECCION TIENDA'],
  PHONE_1: [
    'TELEFONO_1',
    'TELÉFONO',
    'PHONE',
    'TELEFONO',
    'FONO',
    'TELÉFONO 1',
    'TELEFONO 1',
    'CONTACTO',
  ],
  CITY: ['CIUDAD', 'CITY', 'MUNICIPIO', 'LOCALIDAD', 'POBLACION'],
  REGION: ['Región*', 'REGION', 'REGIÓN', 'DEPARTAMENTO', 'ESTADO', 'PROVINCIA', 'ZONA'],
  NIT: [
    'NIT(CLIENTE)',
    'NIT',
    'NIT CLIENTE',
    'CEDULA',
    'RUT',
    'NUMERO IDENTIFICACION',
    'ID CLIENTE',
  ],
  PHONE_2: [
    'TELEFONO_2',
    'TELÉFONO 2',
    'TELEFONO 2',
    'FONO 2',
    'TELEFONO ALTERNO',
    'SEGUNDO TELEFONO',
  ],
  CONTACT: ['CONTACTO', 'CONTACT', 'PERSONA CONTACTO', 'NOMBRE CONTACTO', 'REPRESENTANTE'],
  EMAIL: ['CORREO', 'EMAIL', 'CORREO ELECTRONICO', 'CORREO ELECTRÓNICO', 'MAIL', 'E-MAIL'],
  CODE2: ['CODIGO 2', 'CODE 2', 'CODIGO ALTERNATIVO', 'CODIGO SECUNDARIO', 'CODIGO ADICIONAL'],
  LATITUDE: ['LATITUD', 'LATITUDE', 'LAT', 'COORDENADA LATITUD'],
  LONGITUDE: ['LONGITUD', 'LONGITUDE', 'LON', 'COORDENADA LONGITUD'],
  NEIGHBORHOOD: ['BARRIO', 'NEIGHBORHOOD', 'VECINDARIO', 'SECTOR', 'ZONA'],
};

@Component({
  selector: 'app-outbound',
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
  templateUrl: './outbound.html',
  styleUrl: './outbound.css',
})
export class Outbound {
  @ViewChild('dtResults') dtResults!: Table;

  readonly targetFields = TARGET_FIELDS;
  readonly orderSourceFields = ORDER_SOURCE_FIELDS;
  readonly storeSourceFields = STORE_SOURCE_FIELDS;
  readonly numberKeys = new Set(['QTY', 'PRIORITY', 'PRICE', 'TAXES']);

  orderFile = signal<FileData | null>(null);
  storeFile = signal<FileData | null>(null);
  orderRows = signal<any[]>([]);
  storeRows = signal<any[]>([]);
  availableOrderHeaders = signal<string[]>([]);
  availableStoreHeaders = signal<string[]>([]);
  loading = signal(false);
  isDraggingOrder = signal(false);
  isDraggingStore = signal(false);
  analysisResult = signal<OutboundResult | null>(null);
  searchValue = '';
  showMapping = false;
  useManualMapping = false;

  orderMapping: Record<string, string> = {};
  orderFixedValues: Record<string, string> = {};

  storeMapping: Record<string, string> = {};
  storeFixedValues: Record<string, string> = {};

  crossRefOrderKey = '';
  crossRefStoreKey = '';

  breadcrumbItems: MenuItem[] = [{ label: 'Salida de Mercancía' }];
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

  get filteredResults(): OutboundEntry[] {
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

  onDragOverOrder(event: DragEvent): void {
    event.preventDefault();
    this.isDraggingOrder.set(true);
  }

  onDragLeaveOrder(event: DragEvent): void {
    event.preventDefault();
    this.isDraggingOrder.set(false);
  }

  onDropOrder(event: DragEvent): void {
    event.preventDefault();
    this.isDraggingOrder.set(false);
    const files = event.dataTransfer?.files;
    if (files && files.length > 0) this.processOrderFile(files[0]);
  }

  onOrderFileUpload(event: Event): void {
    if (!this.isBrowser) return;
    const input = event.target as HTMLInputElement;
    if (input.files && input.files.length > 0) this.processOrderFile(input.files[0]);
  }

  onDragOverStore(event: DragEvent): void {
    event.preventDefault();
    this.isDraggingStore.set(true);
  }

  onDragLeaveStore(event: DragEvent): void {
    event.preventDefault();
    this.isDraggingStore.set(false);
  }

  onDropStore(event: DragEvent): void {
    event.preventDefault();
    this.isDraggingStore.set(false);
    const files = event.dataTransfer?.files;
    if (files && files.length > 0) this.processStoreFile(files[0]);
  }

  onStoreFileUpload(event: Event): void {
    if (!this.isBrowser) return;
    const input = event.target as HTMLInputElement;
    if (input.files && input.files.length > 0) this.processStoreFile(input.files[0]);
  }

  removeOrderFile(): void {
    this.orderFile.set(null);
    this.orderRows.set([]);
    this.availableOrderHeaders.set([]);
    this.clearResultsIfNoFiles();
    this.messageService.add({
      severity: 'info',
      summary: 'Archivo removido',
      detail: 'Se ha eliminado el archivo de pedidos',
      life: 3000,
    });
  }

  removeStoreFile(): void {
    this.storeFile.set(null);
    this.storeRows.set([]);
    this.availableStoreHeaders.set([]);
    this.clearResultsIfNoFiles();
    this.messageService.add({
      severity: 'info',
      summary: 'Archivo removido',
      detail: 'Se ha eliminado el archivo de tiendas',
      life: 3000,
    });
  }

  private clearResultsIfNoFiles(): void {
    if (!this.orderFile() && !this.storeFile()) {
      this.analysisResult.set(null);
      this.searchValue = '';
    }
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
    const orderHeaders = this.availableOrderHeaders();
    const storeHeaders = this.availableStoreHeaders();

    for (const field of this.orderSourceFields) {
      const match = this.findBestHeader(field.key, orderHeaders, ORDER_FIELD_ALIASES);
      if (match) this.orderMapping[field.key] = match;
    }

    for (const field of this.storeSourceFields) {
      const match = this.findBestHeader(field.key, storeHeaders, STORE_FIELD_ALIASES);
      if (match) this.storeMapping[field.key] = match;
    }

    const crHeader = orderHeaders.find((h) => h === 'CR' || h === 'Cr' || h === 'cr');
    if (crHeader) {
      this.crossRefOrderKey = crHeader;
    } else {
      this.crossRefOrderKey = this.detectCrossRefKey(orderHeaders, [
        'CR',
        'CODIGO',
        'CODE',
        'LOC',
        'DG.',
      ]);
    }

    const codigoHeader = storeHeaders.find(
      (h) => h === 'Codigo' || h === 'CODIGO' || h === 'código' || h === 'CODE',
    );
    if (codigoHeader) {
      this.crossRefStoreKey = codigoHeader;
    } else {
      this.crossRefStoreKey = this.detectCrossRefKey(storeHeaders, [
        'Codigo',
        'CODE',
        'CÓDIGO',
        'CR',
        'DG.',
      ]);
    }

    if (showToast) {
      this.messageService.add({
        severity: 'info',
        summary: 'Auto mapeo aplicado',
        detail: `Cruce por: ${this.crossRefOrderKey || '?'} ↔ ${this.crossRefStoreKey || '?'}`,
        life: 2000,
      });
    }
  }

  private detectCrossRefKey(headers: string[], preferredTerms: string[]): string {
    const normalizedPreferred = preferredTerms.map((t) => this.normalizeText(t));
    const normalizedHeaders = headers.map((h) => ({ raw: h, normalized: this.normalizeText(h) }));

    // Paso 1: Búsqueda exacta
    for (const term of normalizedPreferred) {
      const match = normalizedHeaders.find((h) => h.normalized === term);
      if (match) return match.raw;
    }

    // Paso 2: Búsqueda parcial con puntuación
    const matches: Array<{ raw: string; score: number }> = [];
    for (const term of normalizedPreferred) {
      for (const header of normalizedHeaders) {
        if (header.normalized.includes(term)) {
          const score = (term.length / header.normalized.length) * 100;
          matches.push({ raw: header.raw, score });
        } else if (term.includes(header.normalized)) {
          const score = (header.normalized.length / term.length) * 100;
          matches.push({ raw: header.raw, score });
        }
      }
    }

    // Retornar el mejor match
    if (matches.length > 0) {
      return matches.sort((a, b) => b.score - a.score)[0].raw;
    }

    return headers.length > 0 ? headers[0] : '';
  }

  hasMappedRequiredFields(): boolean {
    const hasSku =
      !!this.orderMapping['ITEM'] || !!String(this.orderFixedValues['SKU'] || '').trim();
    const hasQty =
      !!this.orderMapping['UNITS'] || !!String(this.orderFixedValues['QTY'] || '').trim();
    const hasStoreCode = !!this.crossRefOrderKey && !!this.crossRefStoreKey;

    return hasSku && hasQty && hasStoreCode;
  }

  calculateOutbound(): void {
    if (!this.orderFile()) {
      this.messageService.add({
        severity: 'warn',
        summary: 'Sin archivo',
        detail: 'Primero carga un archivo de pedidos',
        life: 3000,
      });
      return;
    }

    if (!this.storeFile()) {
      this.messageService.add({
        severity: 'warn',
        summary: 'Sin archivo',
        detail: 'Primero carga un archivo de tiendas para el cruce',
        life: 3000,
      });
      return;
    }

    if (this.useManualMapping) {
      this.processOutbound();
      return;
    }

    this.autoMapHeaders(false);
    this.processOutbound();
  }

  processOutbound(): void {
    const orders = this.orderRows();
    const stores = this.storeRows();

    if (!orders.length) {
      this.messageService.add({
        severity: 'warn',
        summary: 'Sin datos',
        detail: 'El archivo de pedidos no contiene registros válidos',
        life: 3500,
      });
      return;
    }

    if (!stores.length) {
      this.messageService.add({
        severity: 'warn',
        summary: 'Sin datos',
        detail: 'El archivo de tiendas no contiene registros válidos',
        life: 3500,
      });
      return;
    }

    if (!this.crossRefOrderKey) {
      this.messageService.add({
        severity: 'error',
        summary: 'Error de cruce',
        detail: 'Selecciona la columna de código de tienda en el archivo de PEDIDOS (columna CR)',
        life: 5000,
      });
      return;
    }

    if (!this.crossRefStoreKey) {
      this.messageService.add({
        severity: 'error',
        summary: 'Error de cruce',
        detail:
          'Selecciona la columna de código de tienda en el archivo de TIENDAS (columna Codigo)',
        life: 5000,
      });
      return;
    }

    this.loading.set(true);

    try {
      const storeMap = new Map<string, any>();

      for (const store of stores) {
        let storeCode = store[this.crossRefStoreKey];
        storeCode = String(storeCode || '').trim();

        if (storeCode) {
          storeMap.set(storeCode, store);
        }
      }

      const processed: OutboundEntry[] = [];
      let matchedCount = 0;
      let totalQty = 0;

      for (const order of orders) {
        let orderCR = order[this.crossRefOrderKey];
        orderCR = String(orderCR || '').trim();

        const storeData = storeMap.get(orderCR);

        const entry = this.buildOutboundEntry(order, storeData, orderCR);
        processed.push(entry);
        totalQty += entry.QTY;

        if (entry.STORE_CODE !== '') {
          matchedCount++;
        }
      }

      const uniqueOrdersSet = new Set(
        processed
          .map((item) => String(item.ORDER || '').trim().toUpperCase())
          .filter((value) => value !== ''),
      );
      const uniqueStoresSet = new Set(
        processed
          .map((item) => String(item.STORE_CODE || '').trim().toUpperCase())
          .filter((value) => value !== ''),
      );
      const totalOrders = uniqueOrdersSet.size > 0 ? uniqueOrdersSet.size : orders.length;
      const totalStores = uniqueStoresSet.size;

      const summary = `Procesadas ${orders.length} filas de pedidos, ${matchedCount} coincidencias con tiendas, mostrando ${processed.length} registro(s) en tabla`;

      this.analysisResult.set({
        results: processed,
        summary,
        stats: {
          orderRows: orders.length,
          storeRows: stores.length,
          totalOrders,
          totalStores,
          matchedRows: matchedCount,
          totalQty,
        },
      });

      this.messageService.add({
        severity: 'success',
        summary: 'Proceso completado',
        detail: summary,
        life: 4000,
      });
    } catch (error) {
      console.error(error);
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

  private buildOutboundEntry(order: any, storeData: any, orderCR: string): OutboundEntry {
    const getOrderValue = (fieldKey: string): any => {
      const fixed = String(this.orderFixedValues[fieldKey] || '').trim();
      if (fixed !== '') return fixed;
      const mappedHeader = this.orderMapping[fieldKey];
      return mappedHeader ? order[mappedHeader] : '';
    };

    const getStoreValue = (fieldKey: string): any => {
      const fixed = String(this.storeFixedValues[fieldKey] || '').trim();
      if (fixed !== '') return fixed;
      const mappedHeader = this.storeMapping[fieldKey];
      return mappedHeader && storeData ? storeData[mappedHeader] : '';
    };

    const today = this.getTodayISODate();
    const orderDate = this.formatExcelDate(getOrderValue('DELIVERY_DATE')) || today;

    let sku = getOrderValue('ITEM');
    if (!sku) {
      sku = getOrderValue('EAN_LOC');
    }
    const rawSku = String(sku ?? '').trim().replace(/\.0+$/, '');
    const normalizedSku = this.formatSkuForTemplate(rawSku) || rawSku;

    let qty = this.toNumber(getOrderValue('UNITS'));
    if (qty === 0) {
      qty = this.toNumber(getOrderValue('QTY'));
    }
    qty = Math.abs(qty);

    const route = getOrderValue('ROUTE') || getOrderValue('TUTA') || '';
    const qualityStateMappedHeader = this.orderMapping['QUALITY_STATE'];
    const qualityStateFromFile = qualityStateMappedHeader
      ? String(order?.[qualityStateMappedHeader] ?? '').trim()
      : '';
    const qualityStateFallback = String(this.orderFixedValues['QUALITY_STATE'] || 'DISP').trim();
    const qualityStateValue = qualityStateFromFile || qualityStateFallback;

    return {
      ORDER: String(getOrderValue('ORDER_NUMBER') || ''),
      ORDER2: String(getOrderValue('ORDER_NUMBER') || ''),
      PURCHASE_ORDER: String(getOrderValue('ORDER_NUMBER') || ''),
      INVOICE: String(getOrderValue('INVOICE') || ''),
      TRACKING: route,
      ORDER_DATE: orderDate,
      SERVICE_DATE: orderDate,
      SERVICE_DATE_MIN: orderDate,
      SERVICE_DATE_MAX: orderDate,
      OUTBOUNDTYPE_CODE: 'STDA',
      PRIORITY: 1,
      NOTE: String(getOrderValue('NOTE') || ''),
      CARRIER_CODE: '',
      CUSTOMER_CODE: String(getStoreValue('NIT') || ''),
      STORE_CODE: orderCR,
      CITY_STORE_CODE: String(getStoreValue('CITY') || ''),
      STORE_ADDRESS: String(getStoreValue('ADDRESS') || ''),
      STORE_PHONE: String(getStoreValue('PHONE_1') || ''),
      STORE2_CODE: orderCR,
      STORE2_CITY_CODE: String(getStoreValue('CITY') || ''),
      STORE2_ADDRESS: String(getStoreValue('ADDRESS') || ''),
      STORE2_PHONE: String(getStoreValue('PHONE_1') || ''),
      SKU: normalizedSku,
      LOT: String(getOrderValue('LOT') || ''),
      CREATED_DATE: '',
      EXPIRY_DATE: '',
      SERIAL: String(getOrderValue('SERIAL') || ''),
      QUALITY_STATE: qualityStateValue,
      QTY: qty,
      PRICE: 0,
      TAXES: 0,
      UOM_CODE: 'UND',
      REFERENCE: String(getOrderValue('REFERENCE') || ''),
    };
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
      ORDER: r.ORDER,
      ORDER2: r.ORDER2,
      PURCHASE_ORDER: r.PURCHASE_ORDER,
      INVOICE: r.INVOICE,
      TRACKING: r.TRACKING,
      ORDER_DATE: this.formatDateForTextExport(r.ORDER_DATE),
      SERVICE_DATE: this.formatDateForTextExport(r.SERVICE_DATE),
      SERVICE_DATE_MIN: this.formatDateForTextExport(r.SERVICE_DATE_MIN),
      SERVICE_DATE_MAX: this.formatDateForTextExport(r.SERVICE_DATE_MAX),
      OUTBOUNDTYPE_CODE: r.OUTBOUNDTYPE_CODE,
      PRIORITY: r.PRIORITY,
      NOTE: r.NOTE,
      CARRIER_CODE: r.CARRIER_CODE,
      CUSTOMER_CODE: r.CUSTOMER_CODE,
      STORE_CODE: r.STORE_CODE,
      CITY_STORE_CODE: r.CITY_STORE_CODE,
      STORE_ADDRESS: r.STORE_ADDRESS,
      STORE_PHONE: r.STORE_PHONE,
      STORE2_CODE: r.STORE2_CODE,
      STORE2_CITY_CODE: r.STORE2_CITY_CODE,
      STORE2_ADDRESS: r.STORE2_ADDRESS,
      STORE2_PHONE: r.STORE2_PHONE,
      SKU: r.SKU,
      LOT: r.LOT,
      CREATED_DATE: this.formatDateForTextExport(r.CREATED_DATE),
      EXPIRY_DATE: this.formatDateForTextExport(r.EXPIRY_DATE),
      SERIAL: r.SERIAL,
      QUALITY_STATE: r.QUALITY_STATE,
      QTY: r.QTY,
      PRICE: r.PRICE,
      TAXES: r.TAXES,
      UOM_CODE: r.UOM_CODE,
      REFERENCE: r.REFERENCE,
    }));

    const wb = XLSX.utils.book_new();

    const wsSummary = XLSX.utils.json_to_sheet([
      { Métrica: 'Filas pedidos', Valor: result.stats.orderRows },
      { Métrica: 'Filas tiendas', Valor: result.stats.storeRows },
      { Métrica: 'Pedidos únicos', Valor: result.stats.totalOrders },
      { Métrica: 'Tiendas únicas', Valor: result.stats.totalStores },
      { Métrica: 'Coincidencias', Valor: result.stats.matchedRows },
      { Métrica: 'Cantidad total', Valor: result.stats.totalQty },
      { Métrica: 'Registros en tabla', Valor: result.results.length },
    ]);
    XLSX.utils.book_append_sheet(wb, wsSummary, 'Resumen');

    const wsData = XLSX.utils.json_to_sheet(exportRows);
    const textColumns = [
      'SKU',
      'ORDER_DATE',
      'SERVICE_DATE',
      'SERVICE_DATE_MIN',
      'SERVICE_DATE_MAX',
      'CREATED_DATE',
      'EXPIRY_DATE',
    ];
    this.forceTextFormatInColumns(wsData, textColumns);

    XLSX.utils.book_append_sheet(wb, wsData, 'Outbound');

    const date = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `outbound_${date}.xlsx`);

    this.messageService.add({
      severity: 'success',
      summary: 'Exportado',
      detail: 'Archivo de salida generado correctamente',
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
    this.orderMapping = {};
    this.orderFixedValues = {
      OUTBOUNDTYPE_CODE: 'SXP',
      PRIORITY: '1',
      QUALITY_STATE: 'DSP',
      UOM_CODE: 'UND',
    };

    this.storeMapping = {};
    this.storeFixedValues = {};

    this.crossRefOrderKey = '';
    this.crossRefStoreKey = '';
  }

  private getTodayISODate(): string {
    return new Date().toISOString().split('T')[0];
  }

  private processOrderFile(file: File): void {
    if (!this.isBrowser) return;
    this.processFile(file, 'order');
  }

  private processStoreFile(file: File): void {
    if (!this.isBrowser) return;
    this.processFile(file, 'store');
  }

  private processFile(file: File, type: 'order' | 'store'): void {
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

        const fileData: FileData = {
          name: file.name,
          data,
          size: file.size,
          uploadDate: new Date(),
        };

        if (type === 'order') {
          this.orderRows.set(rows);
          this.availableOrderHeaders.set(headers);
          this.orderFile.set(fileData);

          const crHeader = headers.find((h) => h === 'CR' || h === 'Cr' || h === 'cr');
          if (crHeader) {
            this.crossRefOrderKey = crHeader;
          }
        } else {
          this.storeRows.set(rows);
          this.availableStoreHeaders.set(headers);
          this.storeFile.set(fileData);

          const codigoHeader = headers.find(
            (h) => h === 'Codigo' || h === 'CODIGO' || h === 'código' || h === 'CODE',
          );
          if (codigoHeader) {
            this.crossRefStoreKey = codigoHeader;
          }
        }

        this.messageService.add({
          severity: 'success',
          summary: 'Archivo cargado',
          detail: `${type === 'order' ? 'Pedidos' : 'Tiendas'}: ${rows.length} registros encontrados`,
          life: 3000,
        });

        this.autoMapHeaders(false);

        if (!this.useManualMapping && this.orderFile() && this.storeFile()) {
          if (this.crossRefOrderKey && this.crossRefStoreKey) {
            this.processOutbound();
          } else {
            this.messageService.add({
              severity: 'warn',
              summary: 'Columnas de cruce no detectadas',
              detail:
                'Verifica que el archivo de pedidos tenga columna "CR" y el de tiendas "Codigo"',
              life: 5000,
            });
          }
        } else if (this.orderFile() && this.storeFile()) {
          this.messageService.add({
            severity: 'info',
            summary: 'Modo mapeo activo',
            detail: 'Configura el mapeo y luego pulsa Calcular con mapeo',
            life: 3000,
          });
        }
      } catch (error) {
        this.messageService.add({
          severity: 'error',
          summary: 'Error',
          detail: 'No se pudo leer el archivo',
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

  private findBestHeader(
    fieldKey: string,
    headers: string[],
    aliasesMap: Record<string, string[]>,
  ): string {
    const aliases = (aliasesMap[fieldKey] || []).map((a) => this.normalizeText(a));
    const normalizedHeaders = headers.map((h) => ({ raw: h, normalized: this.normalizeText(h) }));

    // Paso 1: Buscar coincidencias exactas
    for (const alias of aliases) {
      const exact = normalizedHeaders.find((h) => h.normalized === alias);
      if (exact) return exact.raw;
    }

    // Paso 2: Buscar coincidencias parciales (uno contiene el otro)
    const partialMatches: Array<{ raw: string; score: number }> = [];
    for (const alias of aliases) {
      const aliasWords = alias.split(/\s+/);
      for (const header of normalizedHeaders) {
        const headerWords = header.normalized.split(/\s+/);
        let matchingWords = 0;

        for (const word of aliasWords) {
          if (headerWords.some((hw) => hw === word || hw.includes(word) || word.includes(hw))) {
            matchingWords++;
          }
        }

        if (matchingWords > 0) {
          const score = (matchingWords / Math.max(aliasWords.length, headerWords.length)) * 100;
          partialMatches.push({ raw: header.raw, score });
        }
      }
    }

    // Retornar el mejor match por puntuación
    if (partialMatches.length > 0) {
      const best = partialMatches.sort((a, b) => b.score - a.score)[0];
      if (best.score >= 50) return best.raw;
    }

    // Paso 3: Buscar contención simple
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
      'item',
      'pedido',
      'order',
      'nombre',
      'name',
      'direccion',
      'address',
      'codigo',
      'code',
      'cr',
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
    if (isoMatch) return `${isoMatch[3]}/${isoMatch[2]}/${isoMatch[1]}`;

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
      if (headerSet.has(headerValue)) targetColumns.push(c);
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

  private formatSkuForTemplate(value: any): string {
    const raw = String(value ?? '').trim();
    if (!raw) return '';

    // Si Excel trae un entero como "12345.0", se limpia para no romper el largo esperado.
    const normalized = raw.replace(/\.0+$/, '');
    if (normalized.length === 6) return normalized;
    if (normalized.length === 5) return `0${normalized}`;
    if (normalized.length === 4) return `00${normalized}`;
    return '';
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
