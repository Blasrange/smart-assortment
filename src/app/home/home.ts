import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { Router } from '@angular/router';
import { MessageService } from 'primeng/api';
import { ToastModule } from 'primeng/toast';

@Component({
  selector: 'app-home',
  imports: [CommonModule, ToastModule],
  providers: [MessageService],
  templateUrl: './home.html',
  styleUrl: './home.css',
})
export class Home {
  constructor(
    private router: Router,
    private messageService: MessageService,
  ) {}

  navigateTo(module: string): void {
    switch (module) {
      case 'inbound':
        this.router.navigate(['/inbound']);
        break;
      case 'outbound':
        this.router.navigate(['/outbound']);
        break;
      case 'sales':
        this.router.navigate(['/sales']);
        break;
      case 'cross-inventory':
        this.router.navigate(['/cross-inventory']);
        break;
      default:
        this.messageService.add({
          severity: 'info',
          summary: 'Módulo',
          detail: `Navegando a ${module}`,
          life: 2000,
        });
    }
  }
}
