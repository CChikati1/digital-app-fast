import { NgModule } from '@angular/core';
import { Routes, RouterModule } from '@angular/router';
import { DataComponent } from './data/data.component';
import { HomeComponent } from './home/home.component';
import { NotificationComponent } from './notification/notification.component';
import { SettingsComponent } from './settings/settings.component';
import { UsersComponent } from './users/users.component';
import { EditComponent } from './edit/edit.component';
import { IssueComponent } from './Issue/issue.component';

const routes: Routes = [
    {
        path: 'home',
        component: HomeComponent,
        data: { title: 'Home', animation: 'isLeft' }
    }, {
        path: 'data',
        component: DataComponent,
        data: { title : 'FS Data' , animation: 'isLeft'}
    }, {
        path: 'notification',
        component: NotificationComponent,
        data : { title : 'Notification', animation: 'isLeft'}
    }, {
        path: 'settings',
        component: SettingsComponent,
        data: { title : 'Settings', animation: 'isRight'}
    }, {
        path : 'users',
        component: UsersComponent,
        data: { title : 'User Management', animation: 'isRight'}
    }, {
        path : 'edit',
        component: EditComponent,
        data: { title : 'Edit Data', animation: 'isRight'}
    }, {
        path : 'issue',
        component: IssueComponent,
        data: { title : 'Issue Tracker', animation: 'isRight'}
    }
];

@NgModule({
    imports: [RouterModule.forChild(routes)],
    exports: [RouterModule]
})
export class PageRoutingModule { }
