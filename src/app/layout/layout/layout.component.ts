import { Component, OnInit } from '@angular/core';
import { NavigationEnd, Router, RouterOutlet } from '@angular/router';
//import {slider } from 'src/app/route-transition-animations';
import { ApiService } from '../../api.service';

@Component({
  selector: 'app-layout',
  templateUrl: './layout.component.html',
  styleUrls: ['./layout.component.scss'],
  //animations: [slider]
})
export class LayoutComponent implements OnInit {
  loginUserName: string = "Test@test.com";
  DisplayName: string = "";
  userRoles:any = [];
  // looged in user role
  isSuperuser:boolean = false;
  constructor(private router: Router, private service: ApiService) {
    
  }
  
  ngOnInit() {
    this.router.events.subscribe((evt) => {
      if (!(evt instanceof NavigationEnd)) {
        return;
      }
      window.scrollTo(0, 0);
    });
    this.getUserName();
  }

  getUserName() {
    this.service.getUserName().subscribe((res) => {
      if (res != null && res !== '') {
        const user = res as any;
        this.loginUserName = user.d.Email;
        this.DisplayName = user.d.Title;
        this.service.getUserRoles().subscribe((res) => {
          if (res != null && res !== '') {
            let resultSet = res as any;
            this.userRoles = resultSet.d.results;            
            let superUser = this.userRoles.filter((items:any) => items.User.EMail.toLowerCase() == this.loginUserName.toLowerCase() && items.Roles == Roles.Superuser);            
            if(superUser.length > 0){
              this.isSuperuser = true;
            }           
          }
        });
        
      }
    });
    
  }
  //prepareRoute(outlet: RouterOutlet) {
    //return outlet && outlet.activatedRouteData && outlet.activatedRouteData['animation'];
  //}
}
enum Roles{
  Superuser = 'Superuser',
  GSteam = 'GS team',
  EY = 'EY',
  KPMG = 'KPMG',
  BU = 'BU',
  Consultant = 'Consultant',
  Auditor = 'Auditor'
}