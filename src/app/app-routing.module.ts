import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { LayoutComponent } from './layout/layout/layout.component';

const routes: Routes = [
  {
    path: '',
    redirectTo: 'app/home',
    pathMatch: 'full',
  },
  {
    path: '',
    component: LayoutComponent,
    data: {
      title: 'app'
    },
    children: [ 
      {
        path: 'app',
        loadChildren: () => import('./views/pages/page.module').then(m=> m.PageModule)
      },
      {
        path: 'error',
        loadChildren: () => import('./views/common/common.module').then(m=> m.CommonModule)
      }
    ]
  },
  { 
    path: '**', 
    redirectTo: 'error/404'
  }
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule { }
