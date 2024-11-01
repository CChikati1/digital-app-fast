import { NgModule } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { CommonRoutingModule } from './common-routing.module';
import { Error404Component } from './error404/error404.component';
import { Error500Component } from './error500/error500.component';

@NgModule({
    imports: [
        FormsModule,
        CommonRoutingModule
    ],
    declarations: [
        Error404Component,
        Error500Component
    ]
})
export class CommonModule { }
