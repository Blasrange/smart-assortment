import { Routes } from '@angular/router';
import { Products } from './products/products';
import { Sales } from './sales/sales';
import { Home } from './home/home';
import { Levels } from './levels/levels';
import { CrossInventory } from './cross-inventory/cross-inventory';
import { AgeInventory } from './age-inventory/age-inventory';
import { ShelfLife } from './shelf-life/shelf-life';
import { AssortmentSales } from './assortment-sales/assortment-sales';
import { ExclusiveSkus } from './exclusive-skus/exclusive-skus';
import { FefoReport } from './fefo-report/fefo-report';
import { LotCross } from './lot-cross/lot-cross';
import { Inbound } from './inbound/inbound';
import { ExitoLabels } from './exito-labels/exito-labels';
import { Outbound } from './outbound/outbound';
import { InventoryReconciliation } from './inventory-reconciliation/inventory-reconciliation';

export const routes: Routes = [
  {
    path: '',
    redirectTo: 'home',
    pathMatch: 'full',
  },
  {
    path: 'home',
    component: Home,
    title: 'Home',
  },
  {
    path: 'sales',
    component: Sales,
    title: 'Ventas',
  },
  {
    path: 'levels',
    component: Levels,
    title: 'Niveles',
  },
  {
    path: 'cross-inventory',
    component: CrossInventory,
    title: 'Cruce de Inventarios',
  },
  {
    path: 'age-inventory',
    component: AgeInventory,
    title: 'Antigüedad del Inventario',
  },
  {
    path: 'shelf-life',
    component: ShelfLife,
    title: 'Vida Útil',
  },
  {
    path: 'assortment-sales',
    component: AssortmentSales,
    title: 'Ventas por Surtido',
  },
  {
    path: 'exclusive-skus',
    component: ExclusiveSkus,
    title: 'SKUs Exclusivos',
  },
  {
    path: 'fefo-report',
    component: FefoReport,
    title: 'Reporte FEFO',
  },

  {
    path: 'lot-cross',
    component: LotCross,
    title: 'Lote Cruce',
  },
  {
    path: 'inbound',
    component: Inbound,
    title: 'Entrada Inbound',
  },
  {
    path: 'exito-labels',
    component: ExitoLabels,
    title: 'Etiquetas Éxito',
  },
  {
    path: 'outbound',
    component: Outbound,
    title: 'Salida',
  },
  {
    path: 'inventory-reconciliation',
    component: InventoryReconciliation,
    title: 'Reconciliación de Inventarios',
  },
];
