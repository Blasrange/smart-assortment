import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';

@Component({
  selector: 'app-home',
  imports: [CommonModule],
  templateUrl: './home.html',
  styleUrl: './home.css',
})
export class Home {
  readonly tableSections = [
    {
      name: 'Productos',
      description: 'Catalogo, stock actual y cobertura por sucursal.',
    },
    {
      name: 'Ventas',
      description: 'Movimientos diarios, ticket promedio y tendencia.',
    },
    {
      name: 'Reportes',
      description: 'Vista consolidada para decisiones de surtido.',
    },
  ];

}