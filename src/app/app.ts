import { Component, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { RouterModule, RouterOutlet } from '@angular/router';
import { MenuItem } from 'primeng/api';
import { MenubarModule } from 'primeng/menubar';
import { RippleModule } from 'primeng/ripple';

@Component({
  selector: 'app-root',
  imports: [CommonModule, RouterModule, RouterOutlet, MenubarModule, RippleModule],
  templateUrl: './app.html',
  styleUrls: ['./app.css']
})
export class App {
  protected readonly title = signal('surtidor_logistico');

  readonly items: MenuItem[] = [
    {
      label: 'Inicio',
      icon: 'pi pi-home',
      routerLink: '/',
    },
    {
      label: 'Productos',
      icon: 'pi pi-box',
      routerLink: '/products',
    },
    {
      label: 'Ventas',
      icon: 'pi pi-shopping-cart',
      routerLink: '/sales',
    },
    {
      label: 'Reportes',
      icon: 'pi pi-chart-line',
      items: [
        {
          label: 'Surtido diario',
          icon: 'pi pi-calendar',
        },
        {
          label: 'Cobertura',
          icon: 'pi pi-chart-bar',
        },
        {
          label: 'Analisis avanzado',
          icon: 'pi pi-chart-line',
        },
      ],
    },
  ];
}