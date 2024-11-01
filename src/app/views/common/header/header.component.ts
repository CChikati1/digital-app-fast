import { Component, OnInit, ViewEncapsulation, AfterViewInit, Output, EventEmitter, HostListener } from '@angular/core';
import { Router } from '@angular/router';
import { ApiService } from '../../../api.service';
@Component({
  selector: 'app-header',
  templateUrl: './header.component.html',
  styleUrls: ['./header.component.scss']
})
export class HeaderComponent implements OnInit {
  isCollapsed : boolean = false;
  loginUserName: string = "Test@test.com";
  DisplayName: string = ""; 
  constructor(private router: Router, private service: ApiService) { }
  ngOnInit(): void {
    this.getUserName();
  }
  getUserName() {    
    this.service.getUserName().subscribe((res) => {
      if (res != null && res !== '') {
        const user = res as any;
        this.loginUserName = user.d.Email;
        this.DisplayName = user.d.Title;
      }
    });    
  }

}
