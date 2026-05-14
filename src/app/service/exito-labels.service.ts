import { HttpClient } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Observable } from 'rxjs';

export interface SystemPrinter {
  name: string;
  status: string;
}

@Injectable({
  providedIn: 'root',
})
export class ExitoLabelsService {
  private readonly baseUrl = '/api/exito-labels';

  constructor(private http: HttpClient) {}

  getPrinters(): Observable<{ printers: SystemPrinter[] }> {
    return this.http.get<{ printers: SystemPrinter[] }>(`${this.baseUrl}/printers`);
  }

  printZpl(printerName: string, zpl: string): Observable<{ success: boolean; message: string }> {
    return this.http.post<{ success: boolean; message: string }>(`${this.baseUrl}/print`, {
      printerName,
      zpl,
    });
  }
}