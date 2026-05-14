import { Component, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { Router, RouterModule, RouterOutlet } from '@angular/router';
import { MenuItem } from 'primeng/api';
import { MenubarModule } from 'primeng/menubar';
import { RippleModule } from 'primeng/ripple';

@Component({
  selector: 'app-root',
  imports: [CommonModule, RouterModule, RouterOutlet, MenubarModule, RippleModule],
  templateUrl: './app.html',
  styleUrls: ['./app.css'],
})
export class App {
  protected readonly title = signal('surtidor_logistico');

  constructor(private router: Router) {}

  readonly items: MenuItem[] = [
    {
      label: 'Comercial',
      icon: 'pi pi-chart-line',
      items: [
        {
          label: 'Análisis de Ventas',
          icon: 'pi pi-shopping-cart',
          routerLink: '/sales',
        },
        {
          label: 'Ventas por Surtido',
          icon: 'pi pi-chart-bar',
          routerLink: '/assortment-sales',
        },
      ],
    },
    {
      label: 'Inventario',
      icon: 'pi pi-box',
      items: [
        {
          label: 'Niveles de Inventario',
          icon: 'pi pi-sitemap',
          routerLink: '/levels',
        },
        {
          label: 'Cruce de Inventarios',
          icon: 'pi pi-box',
          routerLink: '/cross-inventory',
        },
        {
          label: 'Conciliación de Inventarios',
          icon: 'pi pi-box',
          routerLink: '/inventory-reconciliation',
        },
        {
          label: 'SKUs Exclusivos',
          icon: 'pi pi-hashtag',
          routerLink: '/exclusive-skus',
        },
      ],
    },
    {
      label: 'Control de Calidad',
      icon: 'pi pi-clock',
      items: [
        {
          label: 'Antigüedad del Inventario',
          icon: 'pi pi-clock',
          routerLink: '/age-inventory',
        },
        {
          label: 'Vida Útil',
          icon: 'pi pi-hourglass',
          routerLink: '/shelf-life',
        },
        {
          label: 'Reporte FEFO',
          icon: 'pi pi-file',
          routerLink: '/fefo-report',
        },
      ],
    },
    {
      label: 'Operación de Entrada',
      icon: 'pi pi-download',
      items: [
        {
          label: 'Entrada Inbound',
          icon: 'pi pi-truck',
          routerLink: '/inbound',
        },
        {
          label: 'Cruce por Lote',
          icon: 'pi pi-box',
          routerLink: '/lot-cross',
        },
        {
          label: 'Etiquetas Éxito',
          icon: 'pi pi-tags',
          routerLink: '/exito-labels',
        },
      ],
    },
    {
      label: 'Operación de Salida',
      icon: 'pi pi-upload',
      items: [
        {
          label: 'Salida Outbound',
          icon: 'pi pi-arrow-circle-right',
          routerLink: '/outbound',
        },
      ],
    },
  ];

  isMenuItemActive(item: MenuItem): boolean {
    const currentUrl = this.router.url.split('?')[0];
    return this.isItemOrChildrenActive(item, currentUrl);
  }

  private isItemOrChildrenActive(item: MenuItem, currentUrl: string): boolean {
    const currentPath = currentUrl.startsWith('/') ? currentUrl : `/${currentUrl}`;

    const ownPath = this.normalizeRouterLink(item.routerLink);
    if (ownPath) {
      if (ownPath === '/home') {
        if (currentPath === '/home' || currentPath === '/') {
          return true;
        }
      } else if (currentPath === ownPath || currentPath.startsWith(`${ownPath}/`)) {
        return true;
      }
    }

    if (item.items?.length) {
      return item.items.some((child) => this.isItemOrChildrenActive(child, currentPath));
    }

    return false;
  }

  private normalizeRouterLink(routerLink: MenuItem['routerLink']): string {
    if (!routerLink) return '';

    if (Array.isArray(routerLink)) {
      const path = routerLink
        .map((segment) => String(segment ?? '').trim())
        .filter((segment) => segment !== '')
        .join('/');
      if (!path) return '';
      return path.startsWith('/') ? path : `/${path}`;
    }

    const path = String(routerLink).trim();
    if (!path) return '';
    return path.startsWith('/') ? path : `/${path}`;
  }
}
