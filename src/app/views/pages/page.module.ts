import { NgModule } from '@angular/core';
import { DatePipe } from '@angular/common'
import { HomeComponent } from './home/home.component';
import { PageRoutingModule } from './page-routing.module';
import { DataComponent } from './data/data.component';
import { NotificationComponent } from './notification/notification.component';
import { SettingsComponent } from './settings/settings.component';
import { EditComponent } from './edit/edit.component';
import { IssueComponent } from './Issue/issue.component';
import { UsersComponent } from './users/users.component';
import { MaterialModule } from 'src/app/material-module';
import { CommonModule } from '@angular/common';
import { ReactiveFormsModule } from '@angular/forms';
import { NgbModule } from '@ng-bootstrap/ng-bootstrap';
import { NgSelect2Module } from 'ng-select2';
import { FileDragNDropDirective } from './file-drag-n-drop.directive';
import { HttpClientModule, HttpClientJsonpModule, HttpClient } from '@angular/common/http';
import {DataTablesModule} from 'angular-datatables';
import { RichTextEditorModule } from '@syncfusion/ej2-angular-richtexteditor';
import { FormsModule } from '@angular/forms';

@NgModule({
    imports: [
        PageRoutingModule,
        MaterialModule,
        CommonModule,
        ReactiveFormsModule,
        NgbModule,
        NgSelect2Module,
        HttpClientModule,
        HttpClientJsonpModule,
        DataTablesModule,
        RichTextEditorModule,
        FormsModule
    ],
    declarations: [
        HomeComponent,
        DataComponent,
        NotificationComponent,
        SettingsComponent,
        UsersComponent,
        FileDragNDropDirective,
        EditComponent,
        IssueComponent
    ],
    providers: [
        DatePipe
      ]
})

export class PageModule { }
