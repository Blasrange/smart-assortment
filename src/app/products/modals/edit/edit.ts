import { Component, Input, Output, EventEmitter } from '@angular/core';
import { DialogModule } from 'primeng/dialog';
import { FormsModule } from '@angular/forms';

const EMPTY_PRODUCT = {
  title: '',
  sku: '',
  description: '',
  category: '',
  price: 0,
  stock: 0,
  discountPercentage: 0,
  active: true
};

@Component({
  selector: 'app-edit',
  imports: [DialogModule, FormsModule],
  templateUrl: './edit.html',
  styleUrl: './edit.css',
})
export class Edit {
  @Input() visible = false;

  private _product = { ...EMPTY_PRODUCT };

  @Input() set product(value: any) {
    this._product = value
      ? { ...EMPTY_PRODUCT, ...value }
      : { ...EMPTY_PRODUCT };
  }

  get product() {
    return this._product;
  }

  @Output() save = new EventEmitter<any>();
  @Output() cancel = new EventEmitter<void>();

  onSave() {
    this.save.emit(this.product);
  }

  onCancel() {
    this.cancel.emit();
  }
}