import { AfterViewInit, Component, OnInit, ChangeDetectorRef } from '@angular/core';
import { NgbModalConfig, NgbModal, NgbDateStruct } from '@ng-bootstrap/ng-bootstrap';
import { FormBuilder, FormGroup, FormControl, Validators, NgForm, AbstractControl } from '@angular/forms';
// import { ToolbarService, LinkService, ImageService, HtmlEditorService } from '@syncfusion/ej2-angular-richtexteditor';
import { ApiService } from '../../../api.service';
import { Router } from '@angular/router';
import * as spns from 'sp-pnp-js';
import { environment } from '../../../../environments/environment';
@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.scss'],
  
})
export class HomeComponent implements OnInit {
  web = new spns.Web(environment.sp_URL);
  loginUserName: string = "";
  DisplayName: string = "";
  isAdmin: boolean = false;
  //Data Table
  tableData:any = [];
  tableOverDueData:any = [];
  tableDataAll:any = [];
  tableDataAllFS:any = [];
  tableDataAllRoles:any = [];
  tableNotificationData:any = [];
  flag:boolean = false;
  docLibPath:string = environment.docLibPath;
  
  //this is a variable that hold number
  total:number = 0;
  stage1:number = 0;
  stage2:number = 0;
  stage3:number = 0;
  stage4:number = 0;
  stage1NumberCount:any;
  stage2NumberCount:any;
  stage3NumberCount:any;
  stage4NumberCount:any;
  messages: any = [];
  // Total
  totalsubmit:number = 0;
  totalUR:number = 0;
  totalapproved:number = 0;
  totalrejected:number = 0;
  totaloverdue:number = 0;
  totalnotdue:number = 0;
  // Stage 1
  stage1submit:number = 0;
  stage1UR:number = 0;
  stage1approved:number = 0;
  stage1rejected:number = 0;
  stage1overdue:number = 0;
  stage1notdue:number = 0;
  // Stage 2
  stage2submit:number = 0;
  stage2UR:number = 0;
  stage2complete:number = 0;
  stage2approved:number = 0;
  stage2rejected:number = 0;
  stage2overdue:number = 0;
  stage2notdue:number = 0;
  // Stage 3
  stage3submit:number = 0;
  stage3UR:number = 0;
  stage3approved:number = 0;
  stage3rejected:number = 0;
  stage3overdue:number = 0;
  stage3notdue:number = 0;
  // Stage 4
  stage4submit:number = 0;
  stage4UR:number = 0;
  stage4approved:number = 0;
  stage4rejected:number = 0;
  stage4overdue:number = 0;
  stage4notdue:number = 0;
  // Other counts
  Critical:number = 0;
  NonCritical:number = 0;
  Standalone:number = 0;
  Consolidated:number = 0;
  Dormant:number = 0;
  Operational:number = 0;

  // BU counts
  Corporate:number = 0;
  SMBU:number = 0;
  CBU:number = 0;
  HBU:number = 0;
  // Table Filter
  categoryType = ['All','Critical','Non-Critical'];
  selectedCategoryTab = "All";
  dTableAllData : any;
  dMyTableAllData : any;
  userRoles:any = [];
  // looged in user role
  isSuperuser:boolean = false;
  isGSTeamMember:boolean = false;  
  isEYTeamMember:boolean = false;
  isKPMGTeamMember:boolean = false;
  isBUTeamMember:boolean = false;
  isReadOnlyMember:boolean = false; 
  // enum values
  stages = stages;
  kpmg = kpmg;
  status = Status;
  ey = EY;
  // On-Track & Delayed
  CriticalOntrack:number = 0;
  CriticalDelayed:number = 0;
  NonCriticalOntrack:number = 0;
  NonCriticalDelayed:number = 0;

  CorporateOntrack:number = 0;
  CorporateDelayed:number = 0;
  SMBUOntrack:number = 0;
  SMBUDelayed:number = 0;
  CBUOntrack:number = 0;
  CBUDelayed:number = 0;
  HBUOntrack:number = 0;
  HBUDelayed:number = 0;
  year:string="2024";
  angEmailForm = new FormGroup({
    EntityCode: new FormControl(''),
    EntityName: new FormControl(''),
    Stage: new FormControl(''),
    EmailType: new FormControl(''),
    EmailTo: new FormControl(''),
    EmailCC: new FormControl(''),
    //EmailContent: new FormControl(''),
    Status: new FormControl('')
  });
  //EmailContent:string = "";
  emailHtmlStr: string = "";

  constructor(private service: ApiService, private router: Router, private chRef: ChangeDetectorRef,private modalService: NgbModal) { }

  ngOnInit(){
    window.scrollTo(0, 0);
    this.getUserName(); 
  }
  
  getUserName() {
    this.service.getUserName().subscribe((res) => {
      console.log(res);
      if (res != null && res !== '') {
        const user = res as any;
        this.loginUserName = user.d.Email;
        this.DisplayName = user.d.Title;
        this.service.getUserRoles().subscribe((res) => {
          if (res != null && res !== '') {
            let resultSet = res as any;
            this.userRoles = resultSet.d.results;
            // this.Category = resultSet.filter((items: { Title: string; }) => items.Title == "Category");
            let gsUser = this.userRoles.filter((items:any) => items.User.EMail.toLowerCase() == this.loginUserName.toLowerCase() && items.Roles == Roles.GSteam);
            if(gsUser.length > 0){
              this.isGSTeamMember = true;
            }
            let superUser = this.userRoles.filter((items:any) => items.User.EMail.toLowerCase() == this.loginUserName.toLowerCase() && items.Roles == Roles.Superuser);            
            if(superUser.length > 0){
              this.isSuperuser = true;
            }
            let eyUser = this.userRoles.filter((items:any) => items.User.EMail.toLowerCase() == this.loginUserName.toLowerCase() && items.Roles == Roles.Consultant);
            if(eyUser.length > 0){
              this.isEYTeamMember = true;
            }
            let kpmgUser = this.userRoles.filter((items:any) => items.User.EMail.toLowerCase() == this.loginUserName.toLowerCase() && items.Roles == Roles.Auditor);            
            if(kpmgUser.length > 0){
              this.isKPMGTeamMember = true;
            }
            let buUser = this.userRoles.filter((items:any) => items.User.EMail.toLowerCase() == this.loginUserName.toLowerCase() && items.Roles == Roles.BU);            
            if(buUser.length > 0){
              this.isBUTeamMember = true;
            }
            
            let readUser = this.userRoles.filter((items:any) => items.User.EMail.toLowerCase() == this.loginUserName.toLowerCase() && items.Roles == Roles.Read);            
            if(readUser.length > 0){
              this.isReadOnlyMember = true;
            }
            this.getDashboardData();
          }
        });
      }
    });
   
    
  }

  async getCurrentUser(){
    this.service.getUserName().subscribe((res) => {
      if (res != null && res !== '') {
        const user = res as any;
        this.loginUserName = user.d.Email;
        this.DisplayName = user.d.Title; 
        this.service.isMemberOfGroup(user.d.Id).subscribe((response) => {
          if (response != null && response != '') {              
              let ruser = response as any;
              if(ruser.d.results.length > 0){
                this.isAdmin = true;
              }
        }
        }), (err: any) => {
          //console.log("Error Occured " + err);
        };       
      }
    });
  }

  selectChangeHandler (event: any) {    
    this.year = event.target.value;
    this.getDashboardData();
  }

  getDashboardData(){    
    this.service.getFsTrackerData().subscribe((res) => {
      if (res != null && res !== '') {
        let resultSet = res as any;
        
        //resultSet = resultSet.d.results;
        resultSet = resultSet.d.results.filter((items:{Year:string}) =>items.Year == this.year);
        this.tableDataAll = resultSet;        
        this.tableDataAllFS = resultSet;

        if(this.isSuperuser){
          //resultSet = resultSet.filter((items: { CurrentStage: string; CurrentStatus: string; }) => items.CurrentStatus != Status.Approved);
        }
        if(this.isGSTeamMember){
          resultSet = resultSet.filter((items: { CurrentStage: string; CurrentStatus: string; }) => ((items.CurrentStage == stages.Stage1) || 
          (items.CurrentStage == stages.Stage2)));
        }
        if(this.isEYTeamMember){
          resultSet = resultSet.filter((items: { CurrentStage: string; CurrentStatus: string; }) => ((items.CurrentStage == stages.Stage1) || 
          (items.CurrentStage == stages.Stage2)));
        }
        if(this.isKPMGTeamMember){
          resultSet = resultSet.filter((items: { CurrentStage: string; CurrentStatus: string; }) => items.CurrentStage == stages.Stage3);
        }
        if(this.isBUTeamMember){
          resultSet = resultSet.filter((items: { CurrentStage: string; CurrentStatus: string; }) => items.CurrentStage == stages.Stage4);
        }
        this.tableDataAllRoles = resultSet;
        this.tableData = resultSet;
        this.refreshDataTable();
        const table3: any = $("#tblPendingData").DataTable();
        table3.destroy();
        setTimeout (() => {                                
          this.chRef.detectChanges();
          this.dTableAllData = $("#tblPendingData").DataTable({
              dom: "Blfrtip",            
              pageLength: 5,
              paging: true,                        
              scrollX: false,
              ordering:true,
              columns: [
                { "width": "12%" },
                { "width": "8%" },
                { "width": "8%" },
                { "width": "25%" },
                { "width": "9%" },
                { "width": "9%" },                
                { "width": "15%" },
                { "width": "9%" },
                { "width": "12%" },
                { "width": "12%" },
                { "width": "8%" }
              ]
          });
        }, 100);
        
        
        this.total = resultSet.length;
        
        this.stage1 = resultSet.filter((items: { Stage1Status:string; }) => items.Stage1Status != null && items.Stage1Status != "" && items.Stage1Status == Status.Approved).length;
        this.stage2 = resultSet.filter((items: { Stage2Status:string;Stage21Status:string; }) => (items.Stage21Status != null && items.Stage21Status != "" && items.Stage21Status == Status.Approved)).length;
        this.stage3 = resultSet.filter((items: { Stage31Status:string; Stage32Status:string; Stage33Status:string; }) => (items.Stage33Status != null && items.Stage33Status != "" && items.Stage33Status == Status.Approved)).length;
        this.stage4 = resultSet.filter((items: { Stage4Status:string; Stage5Status:string; }) => (items.Stage5Status != null && items.Stage5Status != "" && items.Stage5Status == Status.Approved)).length;

        //  // Total
        //  this.totalsubmit = resultSet.filter((items: { CurrentStatus:string; }) => items.CurrentStatus == Status.Submitted).length;
        //  this.totalUR = resultSet.filter((items: { CurrentStatus:string; }) => items.CurrentStatus == Status.Underreview).length;
        //  this.totalapproved = resultSet.filter((items: { CurrentStatus:string; }) => items.CurrentStatus == Status.Approved).length;
        //  this.totalrejected = resultSet.filter((items: { CurrentStatus:string; }) => items.CurrentStatus == Status.Rejected).length;
        // Stage 1
        this.stage1submit = resultSet.filter((items: { Stage1Status:string; }) => items.Stage1Status == Status.Submitted).length;
        this.stage1UR = resultSet.filter((items: { Stage1Status:string; }) => items.Stage1Status == Status.Underreview).length;
        this.stage1approved = resultSet.filter((items: { Stage1Status:string; }) => items.Stage1Status == Status.Approved).length;
        this.stage1rejected = resultSet.filter((items: { Stage1Status:string; }) => items.Stage1Status == Status.Rejected).length;

        // Stage 2
        this.stage2submit = resultSet.filter((items: { Stage2Status:string;Stage21Status:string; }) => (items.Stage2Status == Status.Submitted) || (items.Stage21Status == Status.Submitted)).length;
        this.stage2UR = resultSet.filter((items: { Stage2Status:string;Stage21Status:string; }) => (items.Stage2Status == Status.Underreview) || (items.Stage21Status == Status.Underreview)).length;
        this.stage2approved = resultSet.filter((items: { Stage21Status:string; }) => items.Stage21Status == Status.Approved).length;
        this.stage2rejected = resultSet.filter((items: { Stage2Status:string;Stage21Status:string; }) => (items.Stage2Status == Status.Rejected) || (items.Stage21Status == Status.Rejected)).length;
        
        // Stage 3
        this.stage3submit = resultSet.filter((items: { Stage31Status:string; Stage32Status:string; Stage33Status:string; }) => items.Stage31Status == Status.Submitted || items.Stage32Status == Status.Submitted || items.Stage33Status == Status.Submitted).length;
        this.stage3UR = resultSet.filter((items: { Stage31Status:string; Stage32Status:string; Stage33Status:string; }) => items.Stage31Status == Status.Underreview || items.Stage32Status == Status.Underreview || items.Stage33Status == Status.Underreview).length;
        this.stage3approved = resultSet.filter((items: { Stage31Status:string; Stage32Status:string; Stage33Status:string; }) => items.Stage33Status == Status.Approved).length;
        this.stage3rejected = resultSet.filter((items: { Stage31Status:string; Stage32Status:string; Stage33Status:string; }) => items.Stage31Status == Status.Rejected || items.Stage32Status == Status.Rejected || items.Stage33Status == Status.Rejected).length;

        // Stage 4
        this.stage4submit = resultSet.filter((items: { Stage4Status:string; Stage5Status:string }) => (items.Stage4Status == Status.Submitted || items.Stage5Status == Status.Submitted)).length;
        this.stage4UR = resultSet.filter((items: { Stage4Status:string; Stage5Status:string }) => (items.Stage4Status == Status.Underreview || items.Stage5Status == Status.Underreview)).length;
        this.stage4approved = resultSet.filter((items: { Stage4Status:string; Stage5Status:string }) => items.Stage5Status == Status.Approved).length;
        this.stage4rejected = resultSet.filter((items: { Stage4Status:string; Stage5Status:string }) => (items.Stage4Status == Status.Rejected || items.Stage5Status == Status.Rejected)).length;
        
        // Other Counts
        this.Critical = resultSet.filter((items: { Category: string; }) => items.Category == 'Critical').length;        
        this.NonCritical = resultSet.filter((items: { Category: string; }) => items.Category == 'Non-Critical').length;
        this.Consolidated = resultSet.filter((items: { ConsolidatedStandalone: string; }) => items.ConsolidatedStandalone == 'Consolidated').length;        
        this.Standalone = resultSet.filter((items: { ConsolidatedStandalone: string; }) => items.ConsolidatedStandalone == 'Standalone').length;

        // BU Counts
        this.Corporate = resultSet.filter((items: { BusinessUnit: string; }) => items.BusinessUnit == 'Corporate').length;        
        this.SMBU = resultSet.filter((items: { BusinessUnit: string; }) => items.BusinessUnit == 'SMBU').length;
        this.CBU = resultSet.filter((items: { BusinessUnit: string; }) => items.BusinessUnit == 'CBU').length;        
        this.HBU = resultSet.filter((items: { BusinessUnit: string; }) => items.BusinessUnit == 'HBU').length;
        
        //this.tableOverDueData = resultSet.filter((items: { TargetFSCompletionDate: string;CurrentStatus:string; }) => (items.CurrentStatus != Status.Approved && new Date(items.TargetFSCompletionDate) < new Date()));
        this.tableOverDueData = resultSet.filter((items: { CurrentStage: string;Stage1Deadline:string;Stage2Deadline:string;Stage21Deadline:string;Stage31Deadline:string;Stage32Deadline:string;
          Stage33Deadline:string;Stage4SubmissionDate:string;Stage5SubmissionDate:string;Stage3CurrentState:string;Stage33Status:string;Stage2CurrentState:string;Stage21Status:string; }) => ((items.CurrentStage == stages.Stage1 && new Date(items.Stage1Deadline) < new Date()) 
          || (items.CurrentStage == stages.Stage2 && items.Stage2CurrentState == EY.EYFS && new Date(items.Stage2Deadline) < new Date()) || (items.CurrentStage == stages.Stage2 && items.Stage2CurrentState == EY.FSTOKPMG && new Date(items.Stage21Deadline) < new Date() && items.Stage21Status != Status.Approved) 
          || (items.CurrentStage == stages.Stage3 && items.Stage3CurrentState == kpmg.Manager && new Date(items.Stage31Deadline) < new Date()) || (items.CurrentStage == stages.Stage3 && items.Stage3CurrentState == kpmg.Director && new Date(items.Stage32Deadline) < new Date()) 
          || (items.CurrentStage == stages.Stage3 && items.Stage3CurrentState == kpmg.Final && new Date(items.Stage33Deadline) < new Date() && items.Stage33Status != Status.Approved) 
          || (items.CurrentStage == stages.Stage4 && new Date(items.Stage4SubmissionDate) < new Date()) || (items.CurrentStage == stages.Stage5 && new Date(items.Stage5SubmissionDate) < new Date())));
        
        const table5: any = $("#tblOverdueData").DataTable();
        table5.destroy();
        setTimeout (() => {                                
          this.chRef.detectChanges();
          let table6 = $("#tblOverdueData").DataTable({
              dom: "Blfrtip",            
              pageLength: 5,
              paging: true,                        
              scrollX: false,
              ordering:true,
              columns: [
                { "width": "15%" },
                { "width": "40%" },
                { "width": "20%" },
                { "width": "13%" },
                { "width": "12%" }
              ]
          });
        }, 100);
        this.categoryOntrackDelayed(resultSet);
        this.buOntrackDelayed(resultSet);  

        this.service.getFsEmailData().subscribe((res) => {
          if (res != null && res !== '') {
            let resultSetEmail = res as any;
            resultSetEmail = resultSetEmail.d.results;
            this.tableNotificationData = resultSetEmail;

            const table7: any = $("#tblNotificationData").DataTable();
            table7.destroy();
            setTimeout (() => {                                
              this.chRef.detectChanges();
              let table8 = $("#tblNotificationData").DataTable({
                  dom: "Blfrtip",            
                  pageLength: 5,
                  paging: true,                        
                  scrollX: false,
                  ordering:true,
                  columns: [
                    { "width": "15%" },
                    { "width": "55%" },
                    { "width": "15%" },
                    { "width": "15%" }
                  ]
              });
            }, 100);
          }
        }), (err: string) => {
          console.log("Error Occured " + err);
        };
      }
    }), (err: string) => {
      console.log("Error Occured " + err);
    };


  }

  categoryOntrackDelayed(resultSet:any){
    let criticalontrack = 0, criticaldelayed = 0;
    let noncriticalontrack = 0, noncriticaldelayed = 0;
    this.stage1overdue = 0;
    this.stage1notdue = 0;
    this.stage2overdue = 0;
    this.stage2notdue = 0;
    this.stage3overdue = 0;
    this.stage3notdue = 0;
    this.stage4overdue = 0;
    this.stage4notdue = 0;
    // Stages Overdue
    this.stage1overdue = resultSet.filter((items: { Stage1Status: string; }) => items.Stage1Status == Status.OverDue).length;
    this.stage2overdue = resultSet.filter((items: { Stage2Status: string;Stage21Status: string; }) => (items.Stage2Status == Status.OverDue || items.Stage21Status == Status.OverDue)).length;
    this.stage3overdue = resultSet.filter((items: { Stage31Status: string;Stage32Status: string;Stage33Status: string; }) => (items.Stage31Status == Status.OverDue || items.Stage32Status == Status.OverDue || items.Stage33Status == Status.OverDue)).length;
    this.stage4overdue = resultSet.filter((items: { Stage4Status: string;Stage5Status: string; }) => (items.Stage4Status == Status.OverDue || items.Stage5Status == Status.OverDue)).length;

    // Stages Not yet due
    this.stage1notdue = resultSet.filter((items: { Stage1Status: string; }) => items.Stage1Status == Status.NotYetDue).length;
    this.stage2notdue = resultSet.filter((items: { Stage2Status: string;Stage21Status: string; }) => (items.Stage2Status == Status.NotYetDue || items.Stage21Status == Status.NotYetDue)).length;
    this.stage3notdue = resultSet.filter((items: { Stage31Status: string;Stage32Status: string;Stage33Status: string; }) => (items.Stage31Status == Status.NotYetDue || items.Stage32Status == Status.NotYetDue || items.Stage33Status == Status.NotYetDue)).length;
    this.stage4notdue = resultSet.filter((items: { Stage4Status: string;Stage5Status: string; }) => (items.Stage4Status == Status.NotYetDue || items.Stage5Status == Status.NotYetDue)).length;

    resultSet.filter((items: { Category: string; }) => items.Category == 'Critical').forEach((item:any)=>{      
      if(item.CurrentStage == stages.Stage1 && item.Stage1Deadline !=null && new Date(item.Stage1Deadline) > new Date() && item.Stage1Status != Status.Approved){
        criticalontrack = criticalontrack + 1;
      }else if(item.CurrentStage == stages.Stage1 && item.Stage1Deadline !=null && new Date(item.Stage1Deadline) < new Date() && item.Stage1Status != Status.Approved){
        criticaldelayed = criticaldelayed + 1;
      }      
      if(item.CurrentStage == stages.Stage2){
        if((item.Stage2Deadline !=null && new Date(item.Stage2Deadline) > new Date() && item.Stage2Status != Status.Approved) || 
        (item.Stage21Deadline !=null && new Date(item.Stage21Deadline) > new Date() && item.Stage21Status != Status.Approved)){
          criticalontrack = criticalontrack + 1;
        }else if((item.Stage2Deadline !=null && new Date(item.Stage2Deadline) < new Date() && item.Stage2Status != Status.Approved) || 
        (item.Stage21Deadline !=null && new Date(item.Stage21Deadline) < new Date() && item.Stage21Status != Status.Approved)){
          criticaldelayed = criticaldelayed + 1;
        }
        // if(item.Stage2CurrentState == EY.EYFS  && item.Stage2Status != Status.Approved  && (item.Stage2Deadline !=null && new Date(item.Stage2Deadline) > new Date())){
        //   criticalontrack = criticalontrack + 1;
        // }else if(item.Stage2CurrentState == EY.EYFS  && item.Stage2Status != Status.Approved  && (item.Stage2Deadline !=null && new Date(item.Stage2Deadline) < new Date())){
        //   criticaldelayed = criticaldelayed + 1;
        // }
        // if(item.Stage2CurrentState == EY.FSTOKPMG  && item.Stage21Status != Status.Approved  && (item.Stage21Deadline !=null && new Date(item.Stage21Deadline) > new Date())){
        //   criticalontrack = criticalontrack + 1;
        // }else if(item.Stage2CurrentState == EY.FSTOKPMG  && item.Stage21Status != Status.Approved  && (item.Stage21Deadline !=null && new Date(item.Stage21Deadline) < new Date())){
        //   criticaldelayed = criticaldelayed + 1;
        // }
      }
      if(item.CurrentStage == stages.Stage3){
        // if(item.Stage3CurrentState == kpmg.Manager  && item.Stage31Status != Status.Approved  && (item.Stage31Deadline !=null && new Date(item.Stage31Deadline) > new Date())){
        //   criticalontrack = criticalontrack + 1;
        // }else if(item.Stage3CurrentState == kpmg.Manager  && item.Stage31Status != Status.Approved  && (item.Stage31Deadline !=null && new Date(item.Stage31Deadline) < new Date())){
        //   criticaldelayed = criticaldelayed + 1;
        // }
        // if(item.Stage3CurrentState == kpmg.Director  && item.Stage32Status != Status.Approved  && (item.Stage32Deadline !=null && new Date(item.Stage32Deadline) > new Date())){
        //   criticalontrack = criticalontrack + 1;
        // }else if(item.Stage3CurrentState == kpmg.Director  && item.Stage32Status != Status.Approved  && (item.Stage32Deadline !=null && new Date(item.Stage32Deadline) < new Date())){
        //   criticaldelayed = criticaldelayed + 1;
        // }
        // if(item.Stage3CurrentState == kpmg.Final  && item.Stage33Status != Status.Approved  && item.Stage33Status != Status.Approved && (item.Stage33Deadline !=null && new Date(item.Stage33Deadline) > new Date())){
        //   criticalontrack = criticalontrack + 1;
        // }else if(item.Stage3CurrentState == kpmg.Final  && item.Stage33Status != Status.Approved  && item.Stage33Status != Status.Approved && (item.Stage33Deadline !=null && new Date(item.Stage33Deadline) < new Date())){
        //   criticaldelayed = criticaldelayed + 1;
        // }
        if((item.Stage31Status != Status.Approved  && (item.Stage31Deadline !=null && new Date(item.Stage31Deadline) > new Date())) 
        || (item.Stage32Status != Status.Approved  && (item.Stage32Deadline !=null && new Date(item.Stage32Deadline) > new Date())) 
        || (item.Stage33Status != Status.Approved  && (item.Stage33Deadline !=null && new Date(item.Stage33Deadline) > new Date()))){        
          criticalontrack = criticalontrack + 1;
        }else if((item.Stage31Status != Status.Approved  && (item.Stage31Deadline !=null && new Date(item.Stage31Deadline) < new Date())) 
        || (item.Stage32Status != Status.Approved  && (item.Stage32Deadline !=null && new Date(item.Stage32Deadline) < new Date())) 
        || (item.Stage33Status != Status.Approved  && (item.Stage33Deadline !=null && new Date(item.Stage33Deadline) < new Date()))){        
          criticaldelayed = criticaldelayed + 1;
        }
      }
      // if(item.CurrentStage == stages.Stage4  && item.Stage4Status != Status.Approved  && item.Stage4SubmissionDate !=null && new Date(item.Stage4SubmissionDate) > new Date()){
      //   criticalontrack = criticalontrack + 1;
      // }else if(item.CurrentStage == stages.Stage4  && item.Stage4Status != Status.Approved  && item.Stage4SubmissionDate !=null && new Date(item.Stage4SubmissionDate) < new Date()){
      //   criticaldelayed = criticaldelayed + 1;
      // }
     
      // if(item.CurrentStage == stages.Stage5  && item.Stage5Status != Status.Approved  && item.Stage5SubmissionDate !=null && new Date(item.Stage5SubmissionDate) > new Date()){
      //   criticalontrack = criticalontrack + 1;
      // }else if(item.CurrentStage == stages.Stage5  && item.Stage5Status != Status.Approved  && item.Stage5SubmissionDate !=null && new Date(item.Stage5SubmissionDate) < new Date()){
      //   criticaldelayed = criticaldelayed + 1;
      // }

      if(item.CurrentStage == stages.Stage4 || item.CurrentStage == stages.Stage5)
      {
        if((item.Stage4Status != Status.Approved && (item.Stage4SubmissionDate !=null && new Date(item.Stage4SubmissionDate) > new Date())) 
        || (item.Stage5Status != Status.Approved && (item.Stage5SubmissionDate !=null && new Date(item.Stage5SubmissionDate) > new Date()))){        
          criticalontrack = criticalontrack + 1;
        }else if((item.Stage4Status != Status.Approved && (item.Stage4SubmissionDate !=null && new Date(item.Stage4SubmissionDate) < new Date())) 
        || (item.Stage5Status != Status.Approved && (item.Stage5SubmissionDate !=null && new Date(item.Stage5SubmissionDate) < new Date()))){        
          criticaldelayed = criticaldelayed + 1;
        }
      }

    });

    resultSet.filter((items: { Category: string; }) => items.Category == 'Non-Critical').forEach((item:any)=>{      
      if(item.CurrentStage == stages.Stage1 && item.Stage1Deadline !=null && new Date(item.Stage1Deadline) > new Date() && item.Stage1Status != Status.Approved){
        noncriticalontrack = noncriticalontrack + 1;
      }else if(item.CurrentStage == stages.Stage1 && item.Stage1Deadline !=null && new Date(item.Stage1Deadline) < new Date() && item.Stage1Status != Status.Approved){
        noncriticaldelayed = noncriticaldelayed + 1;
      }
      if(item.CurrentStage == stages.Stage2){
        if((item.Stage2Deadline !=null && new Date(item.Stage2Deadline) > new Date() && item.Stage2Status != Status.Approved) || 
        (item.Stage21Deadline !=null && new Date(item.Stage21Deadline) > new Date() && item.Stage21Status != Status.Approved)){
          noncriticalontrack = noncriticalontrack + 1;
        }else if((item.Stage2Deadline !=null && new Date(item.Stage2Deadline) < new Date() && item.Stage2Status != Status.Approved) || 
        (item.Stage21Deadline !=null && new Date(item.Stage21Deadline) < new Date() && item.Stage21Status != Status.Approved)){
          noncriticaldelayed = noncriticaldelayed + 1;
        }
        // if(item.Stage2CurrentState == EY.EYFS  && item.Stage2Status != Status.Approved  && (item.Stage2Deadline !=null && new Date(item.Stage2Deadline) > new Date())){
        //   noncriticalontrack = noncriticalontrack + 1;
        // }else if(item.Stage2CurrentState == EY.EYFS  && item.Stage2Status != Status.Approved  && (item.Stage2Deadline !=null && new Date(item.Stage2Deadline) < new Date())){
        //   noncriticaldelayed = noncriticaldelayed + 1;
        // }
        // if(item.Stage2CurrentState == EY.FSTOKPMG  && item.Stage21Status != Status.Approved  && (item.Stage21Deadline !=null && new Date(item.Stage21Deadline) > new Date())){
        //   noncriticalontrack = noncriticalontrack + 1;
        // }else if(item.Stage2CurrentState == EY.FSTOKPMG  && item.Stage21Status != Status.Approved  && (item.Stage21Deadline !=null && new Date(item.Stage21Deadline) < new Date())){
        //   noncriticaldelayed = noncriticaldelayed + 1;
        // }
      }
      if(item.CurrentStage == stages.Stage3){        
        // if(item.Stage3CurrentState == kpmg.Manager  && item.Stage31Status != Status.Approved && (item.Stage31Deadline !=null && new Date(item.Stage31Deadline) > new Date())){
        //   noncriticalontrack = noncriticalontrack + 1;
        // }else if(item.Stage3CurrentState == kpmg.Manager  && item.Stage31Status != Status.Approved && (item.Stage31Deadline !=null && new Date(item.Stage31Deadline) < new Date())){
        //   noncriticaldelayed = noncriticaldelayed + 1;
        // }
        // if(item.Stage3CurrentState == kpmg.Director  && item.Stage32Status != Status.Approved && (item.Stage32Deadline !=null && new Date(item.Stage32Deadline) > new Date())){
        //   noncriticalontrack = noncriticalontrack + 1;
        // }else if(item.Stage3CurrentState == kpmg.Director && item.Stage32Status != Status.Approved && (item.Stage32Deadline !=null && new Date(item.Stage32Deadline) < new Date())){
        //   noncriticaldelayed = noncriticaldelayed + 1;
        // }
        // if(item.Stage3CurrentState == kpmg.Final && item.Stage33Status != Status.Approved && (item.Stage33Deadline !=null && new Date(item.Stage33Deadline) > new Date())){
        //   noncriticalontrack = noncriticalontrack + 1;
        // }else if(item.Stage3CurrentState == kpmg.Final && item.Stage33Status != Status.Approved && (item.Stage33Deadline !=null && new Date(item.Stage33Deadline) < new Date())){
        //   noncriticaldelayed = noncriticaldelayed + 1;
        // }
        if((item.Stage31Status != Status.Approved  && (item.Stage31Deadline !=null && new Date(item.Stage31Deadline) > new Date())) 
        || (item.Stage32Status != Status.Approved  && (item.Stage32Deadline !=null && new Date(item.Stage32Deadline) > new Date())) 
        || (item.Stage33Status != Status.Approved  && (item.Stage33Deadline !=null && new Date(item.Stage33Deadline) > new Date()))){        
          noncriticalontrack = noncriticalontrack + 1;
        }else if((item.Stage31Status != Status.Approved  && (item.Stage31Deadline !=null && new Date(item.Stage31Deadline) < new Date())) 
        || (item.Stage32Status != Status.Approved  && (item.Stage32Deadline !=null && new Date(item.Stage32Deadline) < new Date())) 
        || (item.Stage33Status != Status.Approved  && (item.Stage33Deadline !=null && new Date(item.Stage33Deadline) < new Date()))){        
          noncriticaldelayed = noncriticaldelayed + 1;
        }
      }
      // if(item.CurrentStage == stages.Stage4  && item.Stage4Status != Status.Approved && item.Stage4SubmissionDate !=null && new Date(item.Stage4SubmissionDate) > new Date()){
      //   noncriticalontrack = noncriticalontrack + 1;
      // }else if(item.CurrentStage == stages.Stage4  && item.Stage4Status != Status.Approved && item.Stage4SubmissionDate !=null && new Date(item.Stage4SubmissionDate) < new Date()){
      //   noncriticaldelayed = noncriticaldelayed + 1;
      // }
     
      // if(item.CurrentStage == stages.Stage5  && item.Stage5Status != Status.Approved && item.Stage5SubmissionDate !=null && new Date(item.Stage5SubmissionDate) > new Date()){
      //   noncriticalontrack = noncriticalontrack + 1;
      // }else if(item.CurrentStage == stages.Stage5  && item.Stage5Status != Status.Approved && item.Stage5SubmissionDate !=null && new Date(item.Stage5SubmissionDate) < new Date()){
      //   noncriticaldelayed = noncriticaldelayed + 1;
      // }
      if(item.CurrentStage == stages.Stage4 || item.CurrentStage == stages.Stage5)
      {
        if((item.Stage4Status != Status.Approved && (item.Stage4SubmissionDate !=null && new Date(item.Stage4SubmissionDate) > new Date())) 
        || (item.Stage5Status != Status.Approved && (item.Stage5SubmissionDate !=null && new Date(item.Stage5SubmissionDate) > new Date()))){        
          noncriticalontrack = noncriticalontrack + 1;
        }else if((item.Stage4Status != Status.Approved && (item.Stage4SubmissionDate !=null && new Date(item.Stage4SubmissionDate) < new Date())) 
        || (item.Stage5Status != Status.Approved && (item.Stage5SubmissionDate !=null && new Date(item.Stage5SubmissionDate) < new Date()))){        
          noncriticaldelayed = noncriticaldelayed + 1;
        }
      }
    });

    this.CriticalOntrack = criticalontrack;
    this.CriticalDelayed = criticaldelayed;

    this.NonCriticalOntrack = noncriticalontrack;
    this.NonCriticalDelayed = noncriticaldelayed;
    
  }

  buOntrackDelayed(resultSet:any){
    let corpontrack = 0, corpdelayed = 0;
    let smbuontrack = 0, smbudelayed = 0;
    let cbuontrack = 0, cbudelayed = 0;
    let hbuontrack = 0, hbudelayed = 0;
    resultSet.filter((items: { BusinessUnit: string; }) => items.BusinessUnit == 'Corporate').forEach((item:any)=>{      
      if(item.CurrentStage == stages.Stage1 && item.Stage1Deadline !=null && new Date(item.Stage1Deadline) > new Date()){
        corpontrack = corpontrack + 1;
      }else if(item.CurrentStage == stages.Stage1 && item.Stage1Deadline !=null && new Date(item.Stage1Deadline) < new Date()){
        corpdelayed = corpdelayed + 1;
      }

      if(item.CurrentStage == stages.Stage2 && item.Stage2Deadline !=null && new Date(item.Stage1Deadline) > new Date()){
        corpontrack = corpontrack + 1;
      }else if(item.CurrentStage == stages.Stage2 && item.Stage2Deadline !=null && new Date(item.Stage1Deadline) < new Date()){
        corpdelayed = corpdelayed + 1;
      }

      if(item.CurrentStage == stages.Stage3){
        if(item.Stage3CurrentState == kpmg.Manager && (item.Stage31Deadline !=null && new Date(item.Stage31Deadline) > new Date())){
          corpontrack = corpontrack + 1;
        }else if(item.Stage3CurrentState == kpmg.Manager && (item.Stage31Deadline !=null && new Date(item.Stage31Deadline) < new Date())){
          corpdelayed = corpdelayed + 1;
        }
        if(item.Stage3CurrentState == kpmg.Director && (item.Stage32Deadline !=null && new Date(item.Stage32Deadline) > new Date())){
          corpontrack = corpontrack + 1;
        }else if(item.Stage3CurrentState == kpmg.Director && (item.Stage32Deadline !=null && new Date(item.Stage32Deadline) < new Date())){
          corpdelayed = corpdelayed + 1;
        }
        if(item.Stage3CurrentState == kpmg.Final && item.Stage33Status != Status.Approved && (item.Stage33Deadline !=null && new Date(item.Stage33Deadline) > new Date())){
          corpontrack = corpontrack + 1;
        }else if(item.Stage3CurrentState == kpmg.Final && item.Stage33Status != Status.Approved && (item.Stage33Deadline !=null && new Date(item.Stage33Deadline) < new Date())){
          corpdelayed = corpdelayed + 1;
        }
      }
      if(item.CurrentStage == stages.Stage4 && item.Stage4SubmissionDate !=null && new Date(item.Stage4SubmissionDate) > new Date()){
        corpontrack = corpontrack + 1;
      }else if(item.CurrentStage == stages.Stage4 && item.Stage4SubmissionDate !=null && new Date(item.Stage4SubmissionDate) < new Date()){
        corpdelayed = corpdelayed + 1;
      }
     
      if(item.CurrentStage == stages.Stage5 && item.Stage5SubmissionDate !=null && new Date(item.Stage5SubmissionDate) > new Date()){
        corpontrack = corpontrack + 1;
      }else if(item.CurrentStage == stages.Stage5 && item.Stage5SubmissionDate !=null && new Date(item.Stage5SubmissionDate) < new Date()){
        corpdelayed = corpdelayed + 1;
      }
    });

    this.CorporateOntrack = corpontrack;
    this.CorporateDelayed = corpdelayed;

    resultSet.filter((items: { BusinessUnit: string; }) => items.BusinessUnit == 'SMBU').forEach((item:any)=>{      
      if(item.CurrentStage == stages.Stage1 && item.Stage1Deadline !=null && new Date(item.Stage1Deadline) > new Date()){
        smbuontrack = smbuontrack + 1;
      }else if(item.CurrentStage == stages.Stage1 && item.Stage1Deadline !=null && new Date(item.Stage1Deadline) < new Date()){
        smbudelayed = smbudelayed + 1;
      }
      if(item.CurrentStage == stages.Stage2 && item.Stage2Deadline !=null && new Date(item.Stage1Deadline) > new Date()){
        smbuontrack = smbuontrack + 1;
      }else if(item.CurrentStage == stages.Stage2 && item.Stage2Deadline !=null && new Date(item.Stage1Deadline) < new Date()){
        smbudelayed = smbudelayed + 1;
      }
      if(item.CurrentStage == stages.Stage3){        
        if(item.Stage3CurrentState == kpmg.Manager && (item.Stage31Deadline !=null && new Date(item.Stage31Deadline) > new Date())){
          smbuontrack = smbuontrack + 1;
        }else if(item.Stage3CurrentState == kpmg.Manager && (item.Stage31Deadline !=null && new Date(item.Stage31Deadline) < new Date())){
          smbudelayed = smbudelayed + 1;
        }
        if(item.Stage3CurrentState == kpmg.Director && (item.Stage32Deadline !=null && new Date(item.Stage32Deadline) > new Date())){
          smbuontrack = smbuontrack + 1;
        }else if(item.Stage3CurrentState == kpmg.Director && (item.Stage32Deadline !=null && new Date(item.Stage32Deadline) < new Date())){
          smbudelayed = smbudelayed + 1;
        }
        if(item.Stage3CurrentState == kpmg.Final && item.Stage33Status != Status.Approved && (item.Stage33Deadline !=null && new Date(item.Stage33Deadline) > new Date())){
          smbuontrack = smbuontrack + 1;
        }else if(item.Stage3CurrentState == kpmg.Final && item.Stage33Status != Status.Approved && (item.Stage33Deadline !=null && new Date(item.Stage33Deadline) < new Date())){
          smbudelayed = smbudelayed + 1;
        }
      }
      if(item.CurrentStage == stages.Stage4 && item.Stage4SubmissionDate !=null && new Date(item.Stage4SubmissionDate) > new Date()){
        smbuontrack = smbuontrack + 1;
      }else if(item.CurrentStage == stages.Stage4 && item.Stage4SubmissionDate !=null && new Date(item.Stage4SubmissionDate) < new Date()){
        smbudelayed = smbudelayed + 1;
      }
     
      if(item.CurrentStage == stages.Stage5 && item.Stage5SubmissionDate !=null && new Date(item.Stage5SubmissionDate) > new Date()){
        smbuontrack = smbuontrack + 1;
      }else if(item.CurrentStage == stages.Stage5 && item.Stage5SubmissionDate !=null && new Date(item.Stage5SubmissionDate) < new Date()){
        smbudelayed = smbudelayed + 1;
      }
    });

    this.SMBUOntrack = smbuontrack;
    this.SMBUDelayed = smbudelayed;

    resultSet.filter((items: { BusinessUnit: string; }) => items.BusinessUnit == 'CBU').forEach((item:any)=>{      
      if(item.CurrentStage == stages.Stage1 && item.Stage1Deadline !=null && new Date(item.Stage1Deadline) > new Date()){
        cbuontrack = cbuontrack + 1;
      }else if(item.CurrentStage == stages.Stage1 && item.Stage1Deadline !=null && new Date(item.Stage1Deadline) < new Date()){
        cbudelayed = cbudelayed + 1;
      }
      if(item.CurrentStage == stages.Stage2 && item.Stage2Deadline !=null && new Date(item.Stage1Deadline) > new Date()){
        cbuontrack = cbuontrack + 1;
      }else if(item.CurrentStage == stages.Stage2 && item.Stage2Deadline !=null && new Date(item.Stage1Deadline) < new Date()){
        cbudelayed = cbudelayed + 1;
      }
      if(item.CurrentStage == stages.Stage3){        
        if(item.Stage3CurrentState == kpmg.Manager && (item.Stage31Deadline !=null && new Date(item.Stage31Deadline) > new Date())){
          cbuontrack = cbuontrack + 1;
        }else if(item.Stage3CurrentState == kpmg.Manager && (item.Stage31Deadline !=null && new Date(item.Stage31Deadline) < new Date())){
          cbudelayed = cbudelayed + 1;
        }
        if(item.Stage3CurrentState == kpmg.Director && (item.Stage32Deadline !=null && new Date(item.Stage32Deadline) > new Date())){
          cbuontrack = cbuontrack + 1;
        }else if(item.Stage3CurrentState == kpmg.Director && (item.Stage32Deadline !=null && new Date(item.Stage32Deadline) < new Date())){
          cbudelayed = cbudelayed + 1;
        }
        if(item.Stage3CurrentState == kpmg.Final && item.Stage33Status != Status.Approved && (item.Stage33Deadline !=null && new Date(item.Stage33Deadline) > new Date())){
          cbuontrack = cbuontrack + 1;
        }else if(item.Stage3CurrentState == kpmg.Final && item.Stage33Status != Status.Approved && (item.Stage33Deadline !=null && new Date(item.Stage33Deadline) < new Date())){
          cbudelayed = cbudelayed + 1;
        }
      }
      if(item.CurrentStage == stages.Stage4 && item.Stage4SubmissionDate !=null && new Date(item.Stage4SubmissionDate) > new Date()){
        cbuontrack = cbuontrack + 1;
      }else if(item.CurrentStage == stages.Stage4 && item.Stage4SubmissionDate !=null && new Date(item.Stage4SubmissionDate) < new Date()){
        cbudelayed = cbudelayed + 1;
      }
     
      if(item.CurrentStage == stages.Stage5 && item.Stage5SubmissionDate !=null && new Date(item.Stage5SubmissionDate) > new Date()){
        cbuontrack = cbuontrack + 1;
      }else if(item.CurrentStage == stages.Stage5 && item.Stage5SubmissionDate !=null && new Date(item.Stage5SubmissionDate) < new Date()){
        cbudelayed = cbudelayed + 1;
      }
    });

    this.CBUOntrack = cbuontrack;
    this.CBUDelayed = cbudelayed;

    resultSet.filter((items: { BusinessUnit: string; }) => items.BusinessUnit == 'HBU').forEach((item:any)=>{      
      if(item.CurrentStage == stages.Stage1 && item.Stage1Deadline !=null && new Date(item.Stage1Deadline) > new Date()){
        hbuontrack = hbuontrack + 1;
      }else if(item.CurrentStage == stages.Stage1 && item.Stage1Deadline !=null && new Date(item.Stage1Deadline) < new Date()){
        hbudelayed = hbudelayed + 1;
      }
      if(item.CurrentStage == stages.Stage2 && item.Stage2Deadline !=null && new Date(item.Stage1Deadline) > new Date()){
        hbuontrack = hbuontrack + 1;
      }else if(item.CurrentStage == stages.Stage2 && item.Stage2Deadline !=null && new Date(item.Stage1Deadline) < new Date()){
        hbudelayed = hbudelayed + 1;
      }
      if(item.CurrentStage == stages.Stage3){        
        if(item.Stage3CurrentState == kpmg.Manager && (item.Stage31Deadline !=null && new Date(item.Stage31Deadline) > new Date())){
          hbuontrack = hbuontrack + 1;
        }else if(item.Stage3CurrentState == kpmg.Manager && (item.Stage31Deadline !=null && new Date(item.Stage31Deadline) < new Date())){
          hbudelayed = hbudelayed + 1;
        }
        if(item.Stage3CurrentState == kpmg.Director && (item.Stage32Deadline !=null && new Date(item.Stage32Deadline) > new Date())){
          hbuontrack = hbuontrack + 1;
        }else if(item.Stage3CurrentState == kpmg.Director && (item.Stage32Deadline !=null && new Date(item.Stage32Deadline) < new Date())){
          hbudelayed = hbudelayed + 1;
        }
        if(item.Stage3CurrentState == kpmg.Final && item.Stage33Status != Status.Approved && (item.Stage33Deadline !=null && new Date(item.Stage33Deadline) > new Date())){
          hbuontrack = hbuontrack + 1;
        }else if(item.Stage3CurrentState == kpmg.Final && item.Stage33Status != Status.Approved && (item.Stage33Deadline !=null && new Date(item.Stage33Deadline) < new Date())){
          hbudelayed = hbudelayed + 1;
        }
      }
      if(item.CurrentStage == stages.Stage4 && item.Stage4SubmissionDate !=null && new Date(item.Stage4SubmissionDate) > new Date()){
        hbuontrack = hbuontrack + 1;
      }else if(item.CurrentStage == stages.Stage4 && item.Stage4SubmissionDate !=null && new Date(item.Stage4SubmissionDate) < new Date()){
        hbudelayed = hbudelayed + 1;
      }
     
      if(item.CurrentStage == stages.Stage5 && item.Stage5SubmissionDate !=null && new Date(item.Stage5SubmissionDate) > new Date()){
        hbuontrack = hbuontrack + 1;
      }else if(item.CurrentStage == stages.Stage5 && item.Stage5SubmissionDate !=null && new Date(item.Stage5SubmissionDate) < new Date()){
        hbudelayed = hbudelayed + 1;
      }
    });

    this.HBUOntrack = hbuontrack;
    this.HBUDelayed = hbudelayed;
    
  }

  editPageRedirect(item: any) { 
    //this.router.navigate(['/app/edit'],{queryParams:{itemid:item.ID,source:"home"}});
    (window as any).global = window; 
    const component: HomeComponent = this;
    this.service.checkFolderStatus(item.ID).subscribe((res) => {
      if (res != null && res !== '') {
        let resultSet = res as any;
        resultSet = resultSet.d.results;  
        if(resultSet.length == 0){
          let folderpath = environment.docLibPath+item.ID;
          component.web.folders.add(folderpath).then((d) => {             
            component.router.navigate(['/app/edit'],{queryParams:{itemid:item.ID,source:"home"}});
          });
        }else{
          component.router.navigate(['/app/edit'],{queryParams:{itemid:item.ID,source:"home"}});
        }
       
      }
    });
  }
  viewEmailPage(item: any,content2:any) { 
    //this.router.navigate(['/app/edit'],{queryParams:{itemid:item.ID,source:"home"}});
    (window as any).global = window; 
    const component: HomeComponent = this;    
    this.createForm();
    this.angEmailForm.controls['EntityCode'].setValue(item.EntityCode);
    this.angEmailForm.controls['EntityName'].setValue(item.Entity);
    this.angEmailForm.controls['Stage'].setValue(item.Stage);
    this.angEmailForm.controls['EmailType'].setValue(item.EmailType);
    this.angEmailForm.controls['EmailTo'].setValue(item.EmailTo);
    this.angEmailForm.controls['EmailCC'].setValue(item.EmailCC);
    //this.angEmailForm.controls['EmailContent'].setValue(item.EmailContent);
    this.angEmailForm.controls['Status'].setValue(item.Status);
    this.emailHtmlStr = item.EmailContent;
    //ID,Title,EmailTo,EmailCC,Stage,EmailType,Status,Preparer,Comments,Entity,EntityCode,EmailContent
    component.modalService.open(content2, { size: 'lg', centered: true });
  }
  createForm() {
    this.angEmailForm = new FormGroup({
    EntityCode: new FormControl(''),
    EntityName: new FormControl(''),
    Stage: new FormControl(''),
    EmailType: new FormControl(''),
    EmailTo: new FormControl(''),
    EmailCC: new FormControl(''),
    EmailContent: new FormControl(''),
    Status: new FormControl('')
    });
  }
  
  onChangeCategory(item:any){ 
    this.selectedCategoryTab = item; 
    if(item == "All"){
      this.tableDataAllFS = this.tableDataAll;
    }else{
      this.tableDataAllFS = this.tableDataAll.filter((items: { Category: string; }) => items.Category == item);
    }
    const table3: any = $("#tblPendingData").DataTable();
    table3.destroy();
    setTimeout (() => {                                
      this.chRef.detectChanges();
      this.dTableAllData = $("#tblPendingData").DataTable({
          dom: "Blfrtip",            
          pageLength: 5,
          paging: true,                        
          scrollX: false,
          ordering:true,
          columns: [
            { "width": "12%" },
            { "width": "8%" },
            { "width": "8%" },
            { "width": "25%" },
            { "width": "9%" },
            { "width": "9%" },                
            { "width": "15%" },
            { "width": "9%" },
            { "width": "12%" },
            { "width": "12%" },
            { "width": "8%" }
          ]
      });
    }, 100);
    //this.refreshDataTable();   
  }

  onCategoryFilter(item:any){
    if(item == "All"){
      this.tableData = this.tableDataAllRoles;
    }else{
      this.tableData = this.tableDataAllRoles.filter((items: { Category: string; }) => items.Category == item);
    } 
    this.refreshDataTable();
    
        // const table5: any = $("#tblOverdueData").DataTable();
        // table5.destroy();
        // setTimeout (() => {                                
        //   this.chRef.detectChanges();
        //   let table6 = $("#tblOverdueData").DataTable({
        //       dom: "Blfrtip",            
        //       pageLength: 5,
        //       paging: true,                        
        //       scrollX: false,
        //       ordering:false,
        //       columns: [
        //         { "width": "15%" },
        //         { "width": "55%" },
        //         { "width": "15%" },
        //         { "width": "15%" }
        //       ]
        //   });
        // }, 100);
  }

  onBUFilter(item:any){
    // this.dMyTableAllData.search('').columns().search('').draw();
    // if(item!='All')
    // {    
    //   this.dMyTableAllData.columns(3).search(item, true, false, true).draw();
    // }
    if(item == "All"){
      this.tableData = this.tableDataAllRoles;
    }else{
      this.tableData = this.tableDataAllRoles.filter((items: { BusinessUnit: string; }) => items.BusinessUnit == item);  
    }
    this.refreshDataTable();

  }

  onStageFilter(selectedstages:any){ 
    if(selectedstages == "All"){
      this.tableData = this.tableDataAllRoles;//.filter((items: { Stage1Status:string; }) => items.Stage1Status != null && items.Stage1Status != "");
    }
    if(selectedstages == stages.Stage1){
      this.tableData = this.tableDataAllRoles.filter((items: { Stage1Status:string; }) => items.Stage1Status != null && items.Stage1Status != "");
    }
    if(selectedstages == stages.Stage2){
      this.tableData = this.tableDataAllRoles.filter((items: { Stage2Status:string; }) => items.Stage2Status != null && items.Stage2Status != "");
    }
    if(selectedstages == stages.Stage3){
      this.tableData = this.tableDataAllRoles.filter((items: { Stage31Status:string; Stage32Status:string; Stage33Status:string; }) => 
      ((items.Stage31Status != null && items.Stage31Status != "") || (items.Stage32Status != null && items.Stage32Status != "") || (items.Stage33Status != null && items.Stage33Status != "")));
    }
    if(selectedstages == stages.Stage4){
      this.tableData = this.tableDataAllRoles.filter((items: { Stage4Status:string; Stage5Status:string; }) => 
      ((items.Stage4Status != null && items.Stage4Status != "" )|| (items.Stage5Status != null && items.Stage5Status != "")));
    }

    this.refreshDataTable();

  }

  onStageStatusFilter(selectedstages:any, status:any){
    
    if(selectedstages == "All"){
      this.tableData = this.tableDataAllRoles.filter((items: { CurrentStatus:string; }) => items.CurrentStatus == status);
    }
    if(selectedstages == stages.Stage1){
      // Commented on 28/06/2022
      // if(status == Status.NotYetDue){
      //   this.tableData = this.tableDataAllRoles.filter((items: { Stage1Status:string;Stage1Deadline:string;CurrentStage:string; }) => (items.Stage1Status != Status.Approved && new Date(items.Stage1Deadline) > new Date()));
      // }else if(status == Status.OverDue){
      //   this.tableData = this.tableDataAllRoles.filter((items: { Stage1Status:string;Stage1Deadline:string;CurrentStage:string; }) => (items.Stage1Status != Status.Approved && new Date(items.Stage1Deadline) < new Date()));
      // }else{
      //   this.tableData = this.tableDataAllRoles.filter((items: { Stage1Status:string; }) => items.Stage1Status == status);
      // }
      //Newly added on 28/06/2022
      this.tableData = this.tableDataAllRoles.filter((items: { Stage1Status:string; }) => items.Stage1Status == status);
    }
    if(selectedstages == stages.Stage2){
      // Commented on 28/06/2022
      // if(status == Status.NotYetDue){        
      //   this.tableData = this.tableDataAllRoles.filter((items: { Stage2Status:string;Stage2Deadline:string;Stage21Status:string;Stage21Deadline:string;CurrentStage:string; }) => 
      //   ((new Date(items.Stage2Deadline) > new Date() && items.Stage2Status != Status.Approved) || (new Date(items.Stage21Deadline) > new Date() && items.Stage21Status != Status.Approved)));
      // }else if(status == Status.OverDue){
      //   // this.tableData = this.tableDataAllRoles.filter((items: { Stage2Status:string;Stage2Deadline:string;Stage21Status:string;Stage21Deadline:string;CurrentStage:string; }) => 
      //   // ((new Date(items.Stage2Deadline) < new Date() && items.Stage2Status != Status.Approved) || (new Date(items.Stage21Deadline) < new Date() && items.Stage21Status != Status.Approved)));
      //   this.tableData = this.tableDataAllRoles.filter((items: { Stage2Status:string;Stage2Deadline:string;Stage21Status:string;Stage21Deadline:string;CurrentStage:string; }) => 
      //   ((new Date(items.Stage2Deadline) < new Date() || new Date(items.Stage21Deadline) < new Date()) && items.Stage21Status != Status.Approved));
      // }else{
      //   this.tableData = this.tableDataAllRoles.filter((items: { Stage2Status:string;Stage21Status:string; }) => (items.Stage2Status == status || items.Stage21Status == status));
      // }

      //Newly added on 28/06/2022
      this.tableData = this.tableDataAllRoles.filter((items: { Stage2Status:string;Stage21Status:string; }) => (items.Stage2Status == status || items.Stage21Status == status));
    }
    if(selectedstages == stages.Stage3){
      // Commented on 28/06/2022
      // if(status == Status.NotYetDue){        
      //   this.tableData = this.tableDataAllRoles.filter((items: { Stage31Status:string;Stage31Deadline:string;Stage32Status:string;Stage32Deadline:string;Stage33Status:string;Stage33Deadline:string;CurrentStage:string; }) => 
      //   (items.Stage31Status != Status.Approved && new Date(items.Stage31Deadline) > new Date()) || (items.Stage32Status != Status.Approved && new Date(items.Stage32Deadline) > new Date()) 
      //   || (items.Stage33Status != Status.Approved && new Date(items.Stage33Deadline) > new Date()));
      // }else if(status == Status.OverDue){
      //   this.tableData = this.tableDataAllRoles.filter((items: { Stage31Status:string;Stage31Deadline:string;Stage32Status:string;Stage32Deadline:string;Stage33Status:string;Stage33Deadline:string;CurrentStage:string; }) => 
      //   (items.Stage31Status != Status.Approved && new Date(items.Stage31Deadline) < new Date()) || (items.Stage32Status != Status.Approved && new Date(items.Stage32Deadline) < new Date()) 
      //   || (items.Stage33Status != Status.Approved && new Date(items.Stage33Deadline) < new Date()));
      // }else{
      //   this.tableData = this.tableDataAllRoles.filter((items: { Stage31Status:string; Stage32Status:string; Stage33Status:string; }) => 
      //   items.Stage31Status == status || items.Stage32Status == status || items.Stage33Status == status);
      // }

      //Newly added on 28/06/2022
      this.tableData = this.tableDataAllRoles.filter((items: { Stage31Status:string; Stage32Status:string; Stage33Status:string; }) => 
        items.Stage31Status == status || items.Stage32Status == status || items.Stage33Status == status);
    }
    if(selectedstages == stages.Stage4){
      // Commented on 28/06/2022
      // if(status == Status.NotYetDue){        
      //   this.tableData = this.tableDataAllRoles.filter((items: { Stage4Status:string;Stage4SubmissionDate:string;Stage5Status:string;Stage5SubmissionDate:string;CurrentStage:string; }) => 
      //   (items.Stage4Status != Status.Approved && new Date(items.Stage4SubmissionDate) > new Date()) 
      //   || (items.Stage5Status != Status.Approved && new Date(items.Stage5SubmissionDate) > new Date()));
      // }else if(status == Status.OverDue){
      //   this.tableData = this.tableDataAllRoles.filter((items: { Stage4Status:string;Stage4SubmissionDate:string;Stage5Status:string;Stage5SubmissionDate:string;CurrentStage:string; }) => 
      //   (items.Stage4Status != Status.Approved && new Date(items.Stage4SubmissionDate) < new Date()) 
      //   || (items.Stage5Status != Status.Approved && new Date(items.Stage5SubmissionDate) < new Date()));
      // }else{
      //   this.tableData = this.tableDataAllRoles.filter((items: { Stage4Status:string; Stage5Status:string; }) => items.Stage4Status == status || items.Stage5Status == status);
      // }  
      //Newly added on 28/06/2022
      this.tableData = this.tableDataAllRoles.filter((items: { Stage4Status:string; Stage5Status:string; }) => items.Stage4Status == status || items.Stage5Status == status);
    }
    this.refreshDataTable();
  }

  refreshDataTable(){
    const table: any = $("#tblMyPendingData").DataTable();
    table.destroy();
    setTimeout (() => {                                
      this.chRef.detectChanges();
      this.dMyTableAllData = $("#tblMyPendingData").DataTable({
          dom: "Blfrtip",            
          pageLength: 5,
          paging: true,                        
          scrollX: false,
          ordering:true,
          columns: [
            { "width": "12%" },
            { "width": "8%" },
            { "width": "8%" },
            { "width": "20%" },
            { "width": "9%" },
            { "width": "9%" },
            { "width": "15%" },
            { "width": "9%" },
            { "width": "12%" },
            { "width": "18%" },
            { "width": "8%" }
          ]
      });
    }, 100);
  }

  ngAfterViewInit(): void {
   
  }

  myPendingTasks(){
    this.service.getExternalAuditData(this.loginUserName).subscribe((res) => {
      if (res != null && res !== '') {
        let resultSet = res as any;
        resultSet = resultSet.d.results;
        // this.tableDataAll = resultSet;
        // this.tableData = resultSet;
        
       
      }
    }), (err: string) => {
      console.log("Error Occured " + err);
    };
  }
  stage1numberAnimation (digit: number) {
    this.stage1NumberCount = setInterval(()=>{
      this.stage1++;
      if(this.stage1 == digit)
        clearInterval(this.stage1NumberCount);
    },30) 
  }
  stage2numberAnimation (digit: number) {
    this.stage2NumberCount = setInterval(()=>{
      this.stage2++;
      if(this.stage2 == digit)
        clearInterval(this.stage2NumberCount);
    },30) 
  }
  stage3numberAnimation (digit: number) {
    this.stage3NumberCount = setInterval(()=>{
      this.stage3++;
      if(this.stage3 == digit)
        clearInterval(this.stage3NumberCount);
    },30) 
  }
  stage4numberAnimation (digit: number) {
    this.stage4NumberCount = setInterval(()=>{
      this.stage4++;
      if(this.stage4 == digit)
        clearInterval(this.stage4NumberCount);
    },30) 
  }
}

enum stages{
  Stage1 = 'MAFP Data Submission',
  Stage2 = 'Consultant FS Submission',
  Stage3 = 'Audit Reviews and clearance',
  Stage4 = 'FS signoff',
  Stage5 = 'FS Archiving'
}
enum Status{
  Submitted = 'Submitted',
  Underreview = 'Under review',
  Approved = 'Approved',
  Rejected = 'Rejected',
  Complete = 'Complete',
  OverDue = 'Overdue',
  NotYetDue = 'Not yet due'
}
enum Tabs{
  MAFP_Data_Submission = 0,
  EY_FS_Submission = 1,
  Audit_Reviews_and_clearance = 2,
  FS_signoff = 3,
  FS_Archiving = 4
}
enum kpmgTabs{
  Manager = 0,
  Director = 1,
  Final = 2
}
enum kpmg{
  Manager = 'Auditor Manager FS review',
  Director = 'Auditor Director/Partner FS review',
  Final = 'Auditor Final FS review'
}
enum Roles{
  Superuser = 'Superuser',
  GSteam = 'GS team',
  EY = 'EY',
  KPMG = 'KPMG',
  BU = 'BU Finance',
  Read = 'ReadOnly',
  Consultant = 'Consultant',
  Auditor = 'Auditor'
}
enum EY{
  EYFS = 'Consultant FS Submission',
  FSTOKPMG = 'MAFP Review & FS Submission TO Auditor'
}