import { Component, Input, Output, EventEmitter } from '@angular/core';

import { DialogModule } from 'primeng/dialog';
import { FormsModule } from '@angular/forms';

@Component({
  selector: 'app-create',
  imports: [DialogModule, FormsModule],
  templateUrl: './create.html',
  styleUrl: './create.css',
})
export class Create {
  @Input() visible = false;

  private _product = {
    title: '',
    sku: '',
    description: '',
    category: '',
    price: 0,
    stock: 0,
    discountPercentage: 0,
    active: true
  };

  @Input() set product(val: any) {
    this._product = val
      ? { ...this._product, ...val }
      : { ...this._product };
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