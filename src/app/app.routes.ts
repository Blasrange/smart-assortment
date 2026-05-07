import { Routes } from '@angular/router';
import { Products } from './products/products';
import {Sales} from './sales/sales';
import { Home} from './home/home';

export const routes: Routes = [
	{
		path: '',
		component: Home,
		title: 'Home',
	},	
	{
		path: 'products',
		component: Products,
		title: 'Productos',
	},
	{
		path: 'sales',
		component: Sales,
		title: 'Ventas',
	},
];
