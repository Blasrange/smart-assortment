import { HttpClient } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Observable } from 'rxjs';

@Injectable({
  providedIn: 'root',
})
export class ApiProducts {
  private apiUrl = 'https://dummyjson.com/products';

  constructor(private http: HttpClient) {}

  public getData(): Observable<any> {
    return this.http.get<any>(this.apiUrl);
  }

  public createProduct(product: any): Observable<any> {
    return this.http.post<any>(this.apiUrl, product);
  }

  public updateProduct(id: number, product: any): Observable<any> {
    return this.http.put<any>(`${this.apiUrl}/${id}`, product);
  }

  // Inactivar: PATCH lógico, ejemplo campo "active: false"
  public inactivateProduct(id: number): Observable<any> {
    return this.http.patch<any>(`${this.apiUrl}/${id}`, { active: false });
  }
}
