import { NgModule } from '@angular/core';
import { Routes, RouterModule } from '@angular/router';
import { Error404Component } from './error404/error404.component';
import { Error500Component } from './error500/error500.component';


const routes: Routes = [
    {
        path: '404',
        component: Error404Component,
        data: { title: '404 Error' }
    },
    {
        path: '500',
        component: Error500Component,
        data: { title: '500 Error' }
    }
];

@NgModule({
    imports: [RouterModule.forChild(routes)],
    exports: [RouterModule]
})
export class CommonRoutingModule { }
