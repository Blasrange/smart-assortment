import { Component } from '@angular/core';
import { PrimeNG } from 'primeng/config';
import { ApiProducts } from '../service/api.products';
import { signal } from '@angular/core';

import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';

// PrimeNG Modules para Angular 21

import { TableModule } from 'primeng/table';
import { ButtonModule } from 'primeng/button';
import { ToastModule } from 'primeng/toast';
import { RippleModule } from 'primeng/ripple';
import { InputTextModule } from 'primeng/inputtext';
import { IconFieldModule } from 'primeng/iconfield';
import { InputIconModule } from 'primeng/inputicon';
import { TagModule } from 'primeng/tag';
import { MultiSelectModule } from 'primeng/multiselect';
import { SelectModule } from 'primeng/select';
import { TooltipModule } from 'primeng/tooltip';
import { DialogModule } from 'primeng/dialog';
import { BreadcrumbModule } from 'primeng/breadcrumb';
import { Create } from './modals/create/create';
import { Edit } from './modals/edit/edit';
import { MessageService, MenuItem } from 'primeng/api';

@Component({
  selector: 'app-products',
  imports: [
    CommonModule,
    FormsModule,
    TableModule,
    ButtonModule,
    ToastModule,
    RippleModule,
    InputTextModule,
    IconFieldModule,
    InputIconModule,
    TagModule,
    MultiSelectModule,
    SelectModule,
    TooltipModule,
    DialogModule,
    BreadcrumbModule,
    Create,
    Edit,
  ],
  templateUrl: './products.html',
  styleUrl: './products.css',
  providers: [MessageService],
})
export class Products {
  products = signal<any[]>([]);
  loading = signal(false);
  searchValue: string = '';
  selectedProducts: any[] = [];
  readonly breadcrumbItems: MenuItem[] = [{ label: 'Productos', routerLink: '/products' }];
  readonly breadcrumbHome: MenuItem = { icon: 'pi pi-home', label: 'Inicio', routerLink: '/' };

  // Señales para modal y producto actual
  showCreateModal = signal(false);
  showEditModal = signal(false);
  editingProduct: any = null;

  constructor(
    private api: ApiProducts,
    private primeng: PrimeNG,
    private messageService: MessageService
  ) {}

  clearFilters(dt: any) {
    dt.clear();
    this.searchValue = '';
  }

  ngOnInit() {
    this.llenarData();
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

  llenarData() {
    this.loading.set(true);
    this.api.getData().subscribe({
      next: products => {
        this.products.set(products.products);
        this.loading.set(false);
      },
      error: err => {
        this.loading.set(false);
        this.messageService.add({
          severity: 'error',
          summary: 'Error al cargar',
          detail: 'No se pudieron cargar los productos.'
        });
      }
    });
  }

  openCreateProduct() {
    this.editingProduct = { title: '', sku: '', description: '', category: '', price: 0, stock: 0, discountPercentage: 0, active: true };
    this.showCreateModal.set(true);
  }

  openEditProduct(product: any) {
    this.editingProduct = { ...product };
    this.showEditModal.set(true);
  }

  saveCreateProduct(product: any) {
    this.api.createProduct(product).subscribe({
      next: () => {
        this.llenarData();
        this.showCreateModal.set(false);
        this.messageService.add({
          severity: 'success',
          summary: 'Producto creado',
          detail: 'El producto se ha creado correctamente.'
        });
      },
      error: err => {
        this.messageService.add({
          severity: 'error',
          summary: 'Error al crear',
          detail: 'No se pudo crear el producto.'
        });
      }
    });
  }

  saveEditProduct(product: any) {
    this.api.updateProduct(product.id, product).subscribe({
      next: () => {
        this.llenarData();
        this.showEditModal.set(false);
        this.messageService.add({
          severity: 'success',
          summary: 'Producto actualizado',
          detail: 'El producto se ha actualizado correctamente.'
        });
      },
      error: err => {
        this.messageService.add({
          severity: 'error',
          summary: 'Error al actualizar',
          detail: 'No se pudo actualizar el producto.'
        });
      }
    });
  }

  cancelCreate() {
    this.showCreateModal.set(false);
  }

  cancelEdit() {
    this.showEditModal.set(false);
  }

  inactivateProduct(product: any) {
    this.api.inactivateProduct(product.id).subscribe({
      next: () => {
        this.llenarData();
        this.messageService.add({
          severity: 'info',
          summary: 'Producto inactivado',
          detail: 'El producto ha sido inactivado.'
        });
      },
      error: err => {
        this.messageService.add({
          severity: 'error',
          summary: 'Error al inactivar',
          detail: 'No se pudo inactivar el producto.'
        });
      }
    });
  }
}
