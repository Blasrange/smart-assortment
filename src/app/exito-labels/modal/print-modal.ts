import {
  Component,
  Inject,
  Input,
  OnChanges,
  OnDestroy,
  Output,
  EventEmitter,
  PLATFORM_ID,
  SimpleChanges,
} from '@angular/core';
import { CommonModule, isPlatformBrowser } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DialogModule } from 'primeng/dialog';
import { ButtonModule } from 'primeng/button';
import { InputTextModule } from 'primeng/inputtext';
import { ProgressBarModule } from 'primeng/progressbar';
import { TooltipModule } from 'primeng/tooltip';
import { ExitoLabelsService, type SystemPrinter } from '../../service/exito-labels.service';

interface ExitoLabelPreviewData {
  ct?: string;
  tienda?: string;
  codigoBarra?: string;
  numeroCaja?: number;
  totalCajas?: number;
}

interface LabelOption {
  index: number;
  storeCode: string;
  storeName: string;
  barcode: string;
  boxLabel: string;
}

@Component({
  selector: 'app-print-modal',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    DialogModule,
    ButtonModule,
    InputTextModule,
    ProgressBarModule,
    TooltipModule,
  ],
  templateUrl: './print-modal.html',
  styleUrl: './print-modal.css',
})
export class PrintModal implements OnChanges, OnDestroy {
  @Input() visible = false;
  @Input() totalLabels = 0;
  @Input() zplBatch = '';
  @Input() labels: ExitoLabelPreviewData[] = [];
  @Output() visibleChange = new EventEmitter<boolean>();
  @Output() printDone = new EventEmitter<{ success: boolean; message: string }>();

  printers: SystemPrinter[] = [];
  selectedPrinter = '';
  loadingPrinters = false;
  printing = false;
  progress = 0;
  processedCount = 0;
  previewIndex = 0;
  printCompleted = false;
  printSucceeded = false;
  searchQuery = '';
  selectedMatchIndex: number | null = null;
  printOnlySelected = false;

  private currentPrintTotal = 0;
  private deferredLoadTimer: ReturnType<typeof setTimeout> | null = null;
  private deferredScrollTimer: ReturnType<typeof setTimeout> | null = null;

  private progressInterval: ReturnType<typeof setInterval> | null = null;
  private readonly isBrowser: boolean;

  constructor(
    @Inject(PLATFORM_ID) platformId: object,
    private service: ExitoLabelsService,
  ) {
    this.isBrowser = isPlatformBrowser(platformId);
  }

  get labelPreviews(): string[] {
    if (!this.zplBatch) return [];
    return this.zplBatch
      .split(/(?=\^XA)/)
      .map((s) => s.trim())
      .filter((s) => s.length > 0);
  }

  get currentPreview(): string {
    return this.labelPreviews[this.previewIndex] ?? '';
  }

  get labelOptions(): LabelOption[] {
    const previews = this.labelPreviews;
    if (!previews.length) return [];

    return previews.map((preview, index) => {
      const source = this.labels[index] ?? {};
      const parsed = this.parseLabelFromZpl(preview);

      const storeCode = String(source.ct ?? parsed.ct ?? '').trim();
      const storeName =
        String(source.tienda ?? parsed.tienda ?? '').trim() || 'Tienda no detectada';
      const barcode = String(source.codigoBarra ?? parsed.codigoBarra ?? '').trim() || 'Sin código';
      const numeroCaja = Number(source.numeroCaja ?? parsed.numeroCaja ?? 0);
      const totalCajas = Number(source.totalCajas ?? parsed.totalCajas ?? 0);
      const boxLabel =
        numeroCaja > 0 && totalCajas > 0
          ? `Caja ${numeroCaja} de ${totalCajas}`
          : 'Caja no definida';

      return {
        index,
        storeCode,
        storeName,
        barcode,
        boxLabel,
      };
    });
  }

  get filteredLabelOptions(): LabelOption[] {
    const query = this.searchQuery.trim().toLowerCase();
    if (!query) return this.labelOptions;

    return this.labelOptions.filter((item) =>
      `${item.storeCode} ${item.storeName} ${item.barcode}`.toLowerCase().includes(query),
    );
  }

  get currentSearchResult(): LabelOption | null {
    if (!this.filteredLabelOptions.length) return null;

    if (this.selectedMatchIndex !== null) {
      const selected = this.filteredLabelOptions.find(
        (item) => item.index === this.selectedMatchIndex,
      );
      if (selected) return selected;
    }

    return this.filteredLabelOptions[0];
  }

  get progressLabel(): string {
    if (this.printCompleted) {
      return this.printSucceeded
        ? `${this.currentPrintTotal} etiqueta(s) enviadas con éxito`
        : 'Error en el envío — intente de nuevo';
    }
    if (this.printing) {
      return `Procesando etiqueta ${this.processedCount} de ${this.currentPrintTotal}...`;
    }
    return '';
  }

  get zplSizeKb(): string {
    return (this.zplBatch.length / 1024).toFixed(1);
  }

  get progressBarClass(): string {
    if (this.printCompleted && this.printSucceeded) return 'success-bar';
    if (this.printCompleted && !this.printSucceeded) return 'error-bar';
    return '';
  }

  get hasSearch(): boolean {
    return this.searchQuery.trim().length > 0;
  }

  get selectedLabelNumber(): number | null {
    return this.selectedMatchIndex === null ? null : this.selectedMatchIndex + 1;
  }

  get selectedLabelOption(): LabelOption | null {
    if (this.selectedMatchIndex === null) return null;
    return this.labelOptions.find((opt) => opt.index === this.selectedMatchIndex) ?? null;
  }

  get canPrint(): boolean {
    return (
      !!this.selectedPrinter &&
      !!this.printableZplBatch &&
      !this.printing &&
      this.printableTotal > 0
    );
  }

  get printableZplBatch(): string {
    if (this.printOnlySelected && this.selectedMatchIndex !== null) {
      return this.labelPreviews[this.selectedMatchIndex] ?? '';
    }
    return this.zplBatch;
  }

  get printableTotal(): number {
    if (this.printOnlySelected) {
      return this.selectedMatchIndex === null ? 0 : 1;
    }
    return this.labelPreviews.length || this.totalLabels;
  }

  get currentLabelData(): any {
    const preview = this.labelPreviews[this.previewIndex];
    if (!preview) return null;

    const source = this.labels[this.previewIndex] ?? {};
    const parsed = this.parseLabelFromZpl(preview);

    const numeroCaja = Number(source.numeroCaja ?? parsed.numeroCaja ?? 0);
    const totalCajas = Number(source.totalCajas ?? parsed.totalCajas ?? 0);

    return {
      ct: String(source.ct ?? parsed.ct ?? '').trim(),
      tienda: String(source.tienda ?? parsed.tienda ?? '').trim(),
      codigoBarra: String(source.codigoBarra ?? parsed.codigoBarra ?? '').trim(),
      numeroCaja,
      totalCajas,
      cajaTexto: numeroCaja > 0 && totalCajas > 0 ? `${numeroCaja} / ${totalCajas}` : 'N/A',
    };
  }

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['visible']?.currentValue === true) {
      this.resetState();

      this.deferredScrollTimer = setTimeout(() => {
        this.resetModalScroll();
        this.deferredScrollTimer = null;
      }, 0);

      if (this.isBrowser && !this.printers.length) {
        // Avoid NG0100 by deferring async-bound state changes to the next macrotask.
        this.deferredLoadTimer = setTimeout(() => {
          this.loadPrinters();
          this.deferredLoadTimer = null;
        }, 0);
      }
    }
  }

  ngOnDestroy(): void {
    this.clearTimer();
    if (this.deferredLoadTimer !== null) {
      clearTimeout(this.deferredLoadTimer);
      this.deferredLoadTimer = null;
    }
    if (this.deferredScrollTimer !== null) {
      clearTimeout(this.deferredScrollTimer);
      this.deferredScrollTimer = null;
    }
  }

  loadPrinters(): void {
    if (!this.isBrowser) return;
    this.loadingPrinters = true;
    this.service.getPrinters().subscribe({
      next: (res) => {
        setTimeout(() => {
          this.printers = res.printers ?? [];
          if (!this.selectedPrinter && this.printers.length > 0) {
            this.selectedPrinter = this.printers[0].name;
          }
          this.loadingPrinters = false;
        }, 0);
      },
      error: () => {
        setTimeout(() => {
          this.printers = [];
          this.loadingPrinters = false;
        }, 0);
      },
    });
  }

  prevPreview(): void {
    if (this.previewIndex > 0) this.previewIndex--;
  }

  nextPreview(): void {
    if (this.previewIndex < this.labelPreviews.length - 1) this.previewIndex++;
  }

  onSearchInputChange(): void {
    this.printOnlySelected = false;

    const first = this.filteredLabelOptions[0];
    if (!first) {
      this.selectedMatchIndex = null;
      return;
    }

    this.selectedMatchIndex = first.index;
    this.previewIndex = first.index;
  }

  selectMatchedLabel(indexValue: number | string | null): void {
    if (indexValue === null || indexValue === undefined || indexValue === '') {
      this.selectedMatchIndex = null;
      this.printOnlySelected = false;
      return;
    }

    const parsed = typeof indexValue === 'number' ? indexValue : Number(indexValue);
    if (Number.isNaN(parsed)) {
      this.selectedMatchIndex = null;
      this.printOnlySelected = false;
      return;
    }

    this.selectedMatchIndex = parsed;
    this.previewIndex = parsed;
  }

  clearSearch(): void {
    this.searchQuery = '';
    this.selectedMatchIndex = null;
    this.printOnlySelected = false;

    if (this.labelOptions.length > 0) {
      this.previewIndex = this.labelOptions[0].index;
      this.selectedMatchIndex = this.labelOptions[0].index;
    }
  }

  selectStoreLabel(index: number): void {
    this.selectMatchedLabel(index);
  }

  closeModal(): void {
    this.visibleChange.emit(false);
  }

  downloadZpl(): void {
    if (!this.isBrowser) return;

    const zplToDownload = this.printableZplBatch;
    if (!zplToDownload) return;

    const scope = this.printOnlySelected ? `label-${this.selectedLabelNumber}` : 'batch';
    const filename = `zpl-${scope}-${new Date().toISOString().slice(0, 19).replace(/[T:]/g, '-')}.zpl`;
    const blob = new Blob([zplToDownload], { type: 'text/plain;charset=utf-8' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);

    link.href = url;
    link.download = filename;
    link.click();
    URL.revokeObjectURL(url);
  }

  startPrint(): void {
    const zplToPrint = this.printableZplBatch;
    const labelsToPrint = this.printableTotal;

    if (!this.selectedPrinter || !zplToPrint || this.printing || labelsToPrint <= 0) return;

    this.printing = true;
    this.progress = 0;
    this.processedCount = 0;
    this.printCompleted = false;
    this.printSucceeded = false;
    this.currentPrintTotal = labelsToPrint;

    const totalMs = Math.min(15000, Math.max(2000, labelsToPrint * 350));
    const intervalMs = 100;
    const increment = 90 / (totalMs / intervalMs);

    this.progressInterval = setInterval(() => {
      if (this.progress < 90) {
        this.progress = Math.min(90, +(this.progress + increment).toFixed(2));
        this.processedCount = Math.floor((this.progress / 100) * labelsToPrint);
      }
    }, intervalMs);

    this.service.printZpl(this.selectedPrinter, zplToPrint).subscribe({
      next: (res) => {
        this.clearTimer();
        this.progress = 100;
        this.processedCount = labelsToPrint;
        this.printing = false;
        this.printCompleted = true;
        this.printSucceeded = true;
        this.printDone.emit({
          success: true,
          message: res.message || `${labelsToPrint} etiqueta(s) enviadas correctamente`,
        });

        // Cierra automáticamente la modal al completar la impresión en 100%.
        setTimeout(() => this.visibleChange.emit(false), 250);
      },
      error: (err) => {
        this.clearTimer();
        this.progress = 0;
        this.printing = false;
        this.printCompleted = true;
        this.printSucceeded = false;
        this.printDone.emit({
          success: false,
          message: err?.error?.error || 'No fue posible enviar el ZPL a la impresora',
        });
      },
    });
  }

  resetPrint(): void {
    this.printCompleted = false;
    this.printSucceeded = false;
    this.progress = 0;
    this.processedCount = 0;
    this.currentPrintTotal = this.printableTotal;
  }

  resetState(): void {
    this.clearTimer();
    this.progress = 0;
    this.processedCount = 0;
    this.printing = false;
    this.printCompleted = false;
    this.printSucceeded = false;
    this.previewIndex = 0;
    this.currentPrintTotal = this.totalLabels;
    this.clearSearch();
    if (this.labelOptions.length > 0) {
      this.selectedMatchIndex = this.labelOptions[0].index;
      this.previewIndex = this.labelOptions[0].index;
    }
  }

  private clearTimer(): void {
    if (this.progressInterval !== null) {
      clearInterval(this.progressInterval);
      this.progressInterval = null;
    }
  }

  private resetModalScroll(): void {
    if (!this.isBrowser) return;

    const leftColumn = document.querySelector('.print-dialog .modal-left') as HTMLElement | null;
    const previewColumn = document.querySelector(
      '.print-dialog .label-preview-container',
    ) as HTMLElement | null;
    const content = document.querySelector('.print-dialog .p-dialog-content') as HTMLElement | null;

    if (leftColumn) leftColumn.scrollTop = 0;
    if (previewColumn) previewColumn.scrollTop = 0;
    if (content) content.scrollTop = 0;
  }

  private parseLabelFromZpl(zpl: string): {
    ct: string;
    tienda: string;
    codigoBarra: string;
    numeroCaja: number;
    totalCajas: number;
  } {
    const extractFd = (pattern: RegExp): string => {
      const match = zpl.match(pattern);
      return (match?.[1] ?? '').trim();
    };

    const ct = extractFd(/\^BCN,140,N,N,N\^FD>;([^\^]+)\^FS/);
    const tienda = extractFd(/\^FDTIENDA:\^FS\^FO140,255\^FD([^\^]+)\^FS/);
    const codigoBarra = extractFd(/\^FDCOD BARRA:\^FS\^FO200,408\^FD([^\^]+)\^FS/);

    const boxRaw = extractFd(/\^FDCAJAS:\^FS\^FO130,345\^FD([^\^]+)\^FS/);
    const boxMatch = boxRaw.match(/(\d+)\s+de\s+(\d+)/i);
    const numeroCaja = boxMatch ? Number(boxMatch[1]) : 0;
    const totalCajas = boxMatch ? Number(boxMatch[2]) : 0;

    return {
      ct,
      tienda,
      codigoBarra,
      numeroCaja,
      totalCajas,
    };
  }
}
