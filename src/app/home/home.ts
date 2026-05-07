import { Component } from '@angular/core';
import { MenuItem } from 'primeng/api';
import { BreadcrumbModule } from 'primeng/breadcrumb';
import { ButtonModule } from 'primeng/button';
import { RouterModule } from '@angular/router';

@Component({
  selector: 'app-home',
  imports: [BreadcrumbModule, ButtonModule, RouterModule],
  templateUrl: './home.html',
  styleUrl: './home.css',
})
export class Home {
  readonly items: MenuItem[] = [];
  readonly home: MenuItem = {
    icon: 'pi pi-home',
    label: 'Inicio',
    routerLink: '/',
  };
}
