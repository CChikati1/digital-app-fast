import { Component, OnInit } from '@angular/core';
import { NgbModalConfig, NgbModal, NgbDateStruct } from '@ng-bootstrap/ng-bootstrap';
import { FormBuilder, FormGroup, FormControl, Validators, NgForm, AbstractControl } from '@angular/forms';
import { ApiService } from '../../../api.service';
import { Select2OptionData } from 'ng-select2';
import * as spns from 'sp-pnp-js';
import { environment } from '../../../../environments/environment';
import Swal from 'sweetalert2'
import { DomSanitizer, SafeResourceUrl } from '@angular/platform-browser';


@Component({
  selector: 'app-users',
  templateUrl: './users.component.html',
  styleUrls: ['./users.component.scss'],
  providers: [NgbModalConfig, NgbModal]
})

export class UsersComponent implements OnInit {
  web = new spns.Web(environment.sp_URL);
  listName:string = "UserRoles";
  tableData : any = [];
  public Users: Array<Select2OptionData> = [];
  submitted = false;
  url:string = "";
  safeSrc:SafeResourceUrl
  angForm = new FormGroup({
    User: new FormControl(''),
    Role: new FormControl('')
  });
  loginUserName: string = "Test@test.com";
  DisplayName: string = "";
  userRoles:any = [];
  isSuperuser:boolean = false;
  constructor(private modalService: NgbModal, private service: ApiService,private sanitizer: DomSanitizer) {
    this.createForm();
    
   }
 
  ngOnInit(): void {
    window.scrollTo(0, 0);
    this.service.getUserName().subscribe((res) => {
      if (res != null && res !== '') {
        const user = res as any;
        this.loginUserName = user.d.Email;
        this.DisplayName = user.d.Title;
        // this.service.getUserRoles().subscribe((res) => {
        //   if (res != null && res !== '') {
        //     let resultSet = res as any;
        //     this.userRoles = resultSet.d.results;
        //     let superUser = this.userRoles.filter((items:any) => items.User.EMail.toLowerCase() == this.loginUserName.toLowerCase() && items.Roles == Roles.Superuser);            
        //     if(superUser.length > 0){
        //       this.isSuperuser = true;
        //     }
        //   }
        // });
        
      }
    });
    //this.safeSrc =  this.sanitizer.bypassSecurityTrustResourceUrl(environment.sp_URL + "/Lists/UserRoles/AllItems.aspx");
    // this.Users = [];
    // this.service.getGroupUsers("FSUsers").subscribe((res) => {
    //   if (res != null) {  
    //     const users = res as any; 
    //     const userArr:any = [];
    //     users.d.results.map((item: { Email: any; Title: any; }) => {
    //       return {
    //         id: item.Email,
    //         text: item.Title
    //       };
    //     }).forEach((item: any) => userArr.push(item));
    //     this.Users = userArr;
    //   }
    // });
    // this.createForm();   
    //this.removeElementsByClass("od-ItemsScopeItemContent-breadcrumb"); 
  }

  removeElementsByClass(className:string){
    var elements = null;
    elements = document.getElementsByClassName(className);
    while(elements.length > 0){
        elements[0].removeChild(elements[0]);
    }
  }
  createForm() {
    this.angForm = new FormGroup({
      User: new FormControl('', Validators.required),
      Role: new FormControl('')
    });
  }
  OpenPopUp(content2:any) { 
    this.createForm();
    this.angForm.controls['Role'].setValue("Superuser");
    this.modalService.open(content2, { size: 'lg', centered: true });
  }
  editAuditRequest(item:any, content:any) {
    this.modalService.open(content, { size: 'lg', centered: true });
  }
  saveUserRequest() {
    
    (window as any).global = window;
    this.submitted = true;
    // stop here if form is invalid
    if (this.angForm.invalid) {
      return;
    } 
    let role:string = this.angForm.value.Role;
    let useremail:string =  this.angForm.value.User;
    let username:string =  "";
    if(this.angForm.value.User != ""){    
      let pEmail = this.Users.filter(data => data.id === this.angForm.value.User);      
      if(pEmail.length > 0){
        username = pEmail[0].text;
      }
    }
    this.service.checkFSUserRoles(role, useremail).subscribe((res) => {
        if (res != null && res !== '') {
          let resultSet = res as any;
          resultSet = resultSet.d.results;
          if(resultSet.length == 0){            
            this.web.lists.getByTitle(this.listName).items.add({       
              'Title': username,
              'Useremail': useremail,
              'Roles': role,
              'Status': 'Active',
              'Joined': new Date().toLocaleDateString()
            }).then((item: any) => {
              //this.itemid = item.data.ID;
              //this.initiateNewItem();
              Swal.fire("Success","","success");
              //this.renderData();
              this.createForm();
              this.modalService.dismissAll();
            }).catch((err: any) => {            
              console.log(err);
              //return item;
              //Swal.fire("Failed to update","","info");
            });
          }else{
            Swal.fire("Already record exists with same role","","info");
          }
      }
    });
  }


  get f(): { [key: string]: AbstractControl } {
    return this.angForm.controls;
  }
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