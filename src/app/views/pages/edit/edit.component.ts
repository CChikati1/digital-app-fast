import { AfterViewInit, Component, OnInit, ChangeDetectorRef, ViewChildren, QueryList } from '@angular/core';
import { FormBuilder, FormGroup, FormControl, Validators, NgForm } from '@angular/forms';
import { DatePipe } from '@angular/common'
import { NgbModalConfig, NgbModal, NgbDateStruct } from '@ng-bootstrap/ng-bootstrap';
import { MatSnackBar } from '@angular/material/snack-bar';
import { MatDialog } from '@angular/material/dialog';
import { DialogConfirmComponent } from './dialog-confirm/dialog-confirm.component';
import * as spns from 'sp-pnp-js';
import { environment } from '../../../../environments/environment';
//import { ApiService } from '/../../../api.service';
import { ApiService } from '../../../api.service';
import { Select2OptionData } from 'ng-select2';
import { ActivatedRoute } from '@angular/router';
import * as moment from 'moment';
import Swal from 'sweetalert2'
import { Router } from '@angular/router';
import { DomSanitizer, SafeResourceUrl } from '@angular/platform-browser';
//import { statSync } from 'fs';

interface Stg2CheckListItem {
  category: string;
  header: string;
  selection: string;
  question: string;
  id: number;
  orderby: string;
  qid: string;
  control: string;
  default: string;
}

interface Stg2Item {
  category: string;
  header: string;
  qid: string;
  control: string;
  id: number;
  default: string;
}

@Component({
  selector: 'app-edit',
  templateUrl: './edit.component.html',
  styleUrls: ['./edit.component.scss']
})
export class EditComponent implements OnInit, AfterViewInit {  
  web = new spns.Web(environment.sp_URL);
  // FS Tracker
  listName:string = "FSTracker";
  docLibPath:string = environment.docLibPath;
  itemid: number = 0;
  pageSource:any;
  loginUserName: string = "Test@test.com";
  DisplayName: string = "";
  Preparer: string = "";
  Reviewer: string = "";
  MasterData:any = [];
  targetDate: string = "";
  tableHistoryData: any = [];  
  public Users: Array<Select2OptionData> = [];
  fileInformationData :any = [];
  public files: any[] = [];
  isLinear = false;
  // Item variables
  selectedItem:any =[];
  turnAroundDays:any =[];
  actionHolder: string = "";
  tabIndex : Tabs = Tabs.MAFP_Data_Submission;
  tabKpmgIndex : kpmgTabs = kpmgTabs.Manager;
  tabEYIndex : EYTabs = EYTabs.EYFS;
  tabSignoffIndex : SignoffTabIndexs = SignoffTabIndexs.GSFS;
  tabS21EYChildIndex : S21TabIndex = S21TabIndex.Tab1;
  stages = stages;
  ey = EY;
  kpmg = kpmg;
  signoff = Signoff;
  status = Status;
  s2Tabs = S21Tabs;
  tabEY = EYTabs;
  userRoles:any = [];
  // ngIf stages
  ifStage1:boolean = false;
  ifStage2:boolean = false;
  ifStage3:boolean = false;
  ifStage4:boolean = false;
  ifStage5:boolean = false;

   // ngIf stages active tab
   ifStage1Active:boolean = false;
   ifStage2Active:boolean = false;
   ifStage3Active:boolean = false;
   ifStage4Active:boolean = false;
   ifStage5Active:boolean = false;
  // ngIf stages active tab
  ifFileUploadAllowed:boolean = false;
  ifFileUploadNotAllowed:boolean = false;

  // looged in user role
  isSuperuser:boolean = false;
  isGSTeamMember:boolean = false;  
  isEYTeamMember:boolean = false;
  isKPMGTeamMember:boolean = false;
  isBUTeamMember:boolean = false;
  safeFileUploadSrc:SafeResourceUrl;
  safeFileView:SafeResourceUrl;

  // Check List
  tableCheckListData: any = [];
  isChecked: boolean = false;

  gsgap:number=0;
  consultantgap:number=0;
  auditorgap:number=0;
  crtgap:number=0;
  

  angForm = new FormGroup({
    //TargetFSCompletionDate: new FormControl(''),
    Stage1Deadline: new FormControl(''),
    Stage1ActualSubmission: new FormControl(''),
    Stage1Status: new FormControl(''),    
    Stage1GAP: new FormControl(''),
    Stg1Comments: new FormControl(''),
    Stage2Deadline: new FormControl(''),
    Stage2ActualSubmission: new FormControl(''),
    Stage2Status: new FormControl(''),
    Stage2GAP: new FormControl(''),
    Stg2Comments: new FormControl(''),
    Stage21Deadline: new FormControl(''),
    Stage21ActualSubmission: new FormControl(''),
    Stage21Status: new FormControl(''),
    Stage21GAP: new FormControl(''),
    Stg21Comments: new FormControl(''),
    Stage31Deadline: new FormControl(''),
    Stage31ActualSubmission: new FormControl(''),
    Stage31Status: new FormControl(''),
    Stage31GAP: new FormControl(''),
    Stg31Comments: new FormControl(''),
    Stage32Deadline: new FormControl(''),
    Stage32ActualSubmission: new FormControl(''),
    Stage32Status: new FormControl(''),
    Stage32GAP: new FormControl(''),
    Stg32Comments: new FormControl(''),
    Stage33Deadline: new FormControl(''),
    Stage33ActualSubmission: new FormControl(''),
    Stage33Status: new FormControl(''),
    Stage33GAP: new FormControl(''),
    Stg33Comments: new FormControl(''),
    Stage4Deadline: new FormControl(''),
    Stage4ActualSubmission: new FormControl(''),
    Stage4Status: new FormControl(''),
    Stage4GAP: new FormControl(''),
    Stg4Comments: new FormControl(''),
    Stage42Deadline: new FormControl(''),
    Stage42ActualSubmission: new FormControl(''),
    Stage42Status: new FormControl(''),
    Stage42GAP: new FormControl(''),
    Stg42Comments: new FormControl(''),
    Stage43Deadline: new FormControl(''),
    Stage43ActualSubmission: new FormControl(''),
    Stage43Status: new FormControl(''),
    Stage43GAP: new FormControl(''),
    Stg43Comments: new FormControl(''),
    Stage5Deadline: new FormControl(''),
    Stage5ActualSubmission: new FormControl(''),
    Stage5Status: new FormControl(''),
    Stage5GAP: new FormControl(''),
    Stg5Comments: new FormControl(''),
    Stage3CurrentState:new FormControl(''),
    Stage4CurrentState:new FormControl(''),

    Stage1ConsultantVariableDeadline:new FormControl(''),
    Stage1GSVariableDeadline:new FormControl(''),
    Stage21ConsultantVariableDeadline:new FormControl(''),
    Stage21GSVariableDeadline:new FormControl(''),
    Stage22GSVariableDeadline:new FormControl(''),
    Stage31AuditorvariableDeadline:new FormControl(''),
    Stage31GSVariableDeadline:new FormControl(''),
    Stage31ConsultantVariableDeadline:new FormControl(''),
    Stage32AuditorvariableDeadline:new FormControl(''),
    Stage32GSVariableDeadline:new FormControl(''),
    Stage32ConsultantVariableDeadline:new FormControl(''),
    Stage33AuditorvariableDeadline:new FormControl(''),
    Stage33GSVariableDeadline:new FormControl(''),
    Stage33ConsultantVariableDeadline:new FormControl(''),
    Stage41GSVariableDeadline:new FormControl(''),
    Stage41CRTVariableDeadline:new FormControl(''),
    Stage42GSVariableDeadline:new FormControl(''),
    Stage42CRTVariableDeadline:new FormControl(''),
    Stage43GSVariableDeadline:new FormControl(''),
    Stage43AuditorVariableDeadline:new FormControl(''),
    Stage5GSVariableDeadline:new FormControl(''),
    Stage5CRTVariableDeadline:new FormControl(''),

  });

  stg1Form: FormGroup;
  // Stage 1 Check List  
  stg1ChecklistData:any = [];

  // Stage 2 Check List
  stg2ChildRowItems: Stg2CheckListItem[] = [];
  stg2HeaderRowItems: Stg2Item[] = []; 
  stg2SelectedValues: { name: string, value: string, id: number }[] = []; 
  stg2CommentValues: { name: string, value: string, id: number }[] = []; 
  stg2LevelCL: string = EY.EYFS;
  stg2IsNew: boolean = true;

   // Stage 3 Check List
   stg3ChildRowItems: Stg2CheckListItem[] = [];
   stg3HeaderRowItems: Stg2Item[] = []; 
   stg3SelectedValues: { name: string, value: string, id: number }[] = []; 
   stg3CommentValues: { name: string, value: string, id: number }[] = []; 
   stg3LevelCL: string = EY.EYFS;
   stg3IsNew: boolean = true;

  // Stage 4 Check List
  stg4ChildRowItems: Stg2CheckListItem[] = [];
  stg4HeaderRowItems: Stg2Item[] = []; 
  stg4SelectedValues: { name: string, value: string, id: number }[] = []; 
  stg4CommentValues: { name: string, value: string, id: number }[] = []; 
  stg4LevelCL: string = Signoff.GSFS;
  stg4IsNew: boolean = true;
  // Stage 4 Check List
  stg5ChildRowItems: Stg2CheckListItem[] = [];
  stg5HeaderRowItems: Stg2Item[] = []; 
  stg5SelectedValues: { name: string, value: string, id: number }[] = []; 
  stg5CommentValues: { name: string, value: string, id: number }[] = []; 
  stg5IsNew: boolean = true;

  constructor(private _snackBar: MatSnackBar, public dialog: MatDialog,private formBuilder: FormBuilder,private service: ApiService,
    private activatedRoute: ActivatedRoute, private router: Router, public datepipe: DatePipe, private sanitizer: DomSanitizer, 
    private modalService: NgbModal, private chRef: ChangeDetectorRef) {

     
    
  }

  ngOnInit() {
    window.scrollTo(0, 0);
    this.activatedRoute.queryParams.subscribe((qs: any) => {
      this.itemid = qs.itemid;
      this.pageSource = qs.source;
    });
    this.getUserName();    
    this.stg1Form = this.formBuilder.group({});
    //this.stg21Form = this.formBuilder.group({});
  }

  handleHeaderSelection(category: string, selection: string) {
    const categoryItems = this.stg2ChildRowItems.filter(item => item.category === category);
    categoryItems.forEach(item => {
      item.selection = selection;
    });
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
            // this.Category = resultSet.filter((items: { Title: string; }) => items.Title == "Category");
            let gsUser = this.userRoles.filter((items:any) => items.User.EMail.toLowerCase() == this.loginUserName.toLowerCase() && items.Roles == Roles.GSteam);
            if(gsUser.length > 0){
              this.isGSTeamMember = true;
            }
            let superUser = this.userRoles.filter((items:any) => {
              
              return items.User.EMail.toLowerCase() == this.loginUserName.toLowerCase() && items.Roles == Roles.Superuser;
            });            
            
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
            (window as any).global = window;
            debugger;
            this.service.checkFolderStatus(String(this.itemid)).subscribe((res) => {
              debugger;
              if (res != null && res !== '') {
                let resultSet = res as any;
                resultSet = resultSet.d.results;  
                if(resultSet.length == 0){
                  let folderpath = environment.docLibPath+String(this.itemid);
                  this.web.folders.add(folderpath).then(() => {             
                    this.initiateEditItem();
                  });
                }else{
                  this.initiateEditItem();
                }
              }
            });
            
          }
        });
        //this.ensureFolder(this.itemid);
      }
    });
    this.Users = [];
    this.service.getSiteUsers().subscribe((res) => {
      if (res != null) {  
        const users = res as any; 
        const userArr:any = [];
        users.d.results.map((item: { Email: any; Title: any; }) => {
          return {
            id: item.Email,
            text: item.Title
          };
        }).forEach((item: any) => userArr.push(item));
        this.Users = userArr;
      }
    });
    
  }
  
  initiateEditItem(){
    debugger;
    this.service.getItemById(this.itemid).subscribe((res) => {
      if (res != null) {  
        const itemresult = res as any; 
        this.selectedItem = itemresult.d;
        if(itemresult.d.ActionHolder != null && itemresult.d.ActionHolder != ""){
          this.actionHolder = itemresult.d.ActionHolder;
        }        
        // set preparer and reviewer profile picture
        this.service.getFSTurnAroundDaysById(this.itemid.toString()).subscribe((resdays) => {
          if(resdays!=null){
            const daysresult = resdays as any; 
            if(daysresult.d.results.length>0){
              this.turnAroundDays = daysresult.d.results[0];
              this.bindFileData();
              this.setCurrentStage()       
              this.setStagesDate();
              this.setStagesStatus();
            }
            else{
              this.service.getFSTurnAroundDaysById('0').subscribe((resdaysnew) => {
                if(resdaysnew!=null){
                  const daysresultnew = resdaysnew as any; 
                  if(daysresultnew.d.results.length>0){
                    this.turnAroundDays = daysresultnew.d.results[0];
                    this.bindFileData();
                    this.setCurrentStage()       
                    this.setStagesDate();
                    this.setStagesStatus();
                  }
                }
              });
            }
            
          }
        });
        
        this.setPRProfilePicture();        
      }
    });
  }

  async ensureFolder(ItemID: any): Promise<any> {
    (window as any).global = window;   
    let folderpath = environment.docLibPath+ItemID;
    const folder = await this.web.getFolderByServerRelativePath(folderpath).select('Exists').get();
    if (!folder.Exists) {
        await this.web.folders.add(folderpath);
    }
  }

  setCurrentStage(){
   
    if(this.selectedItem.CurrentStage == null || this.selectedItem.CurrentStage == "" || this.selectedItem.CurrentStage == stages.Stage1){
      this.setStageFalse();
      this.setActiveTab(Tabs.MAFP_Data_Submission);      
      this.ifStage1 = true;
      this.ifStage1Active = true;
      if(this.isGSTeamMember || this.isSuperuser || this.isEYTeamMember ){
        this.ifFileUploadAllowed = true;
        this.ifFileUploadNotAllowed = false;
      }else{
        this.ifFileUploadAllowed = false;
        this.ifFileUploadNotAllowed = true;
      }

    }else if(this.selectedItem.CurrentStage == stages.Stage2){
      this.setStageFalse();
      this.setActiveTab(Tabs.EY_FS_Submission);
      this.ifStage2 = true;
      this.ifStage2Active = true;

      if(this.selectedItem.Stage2CurrentState == EY.EYFS){  
        this.setEYActiveTab(EYTabs.EYFS);   
        if(this.isEYTeamMember || this.isSuperuser){
          this.ifFileUploadAllowed = true;
          this.ifFileUploadNotAllowed = false;
        }else{
          this.ifFileUploadAllowed = false;
          this.ifFileUploadNotAllowed = true;
        }
      }
      if(this.selectedItem.Stage2CurrentState == EY.FSTOKPMG){   
        this.setEYActiveTab(EYTabs.FSTOKPMG);
        if(this.isGSTeamMember || this.isSuperuser || this.isKPMGTeamMember){
          this.ifFileUploadAllowed = true;
          this.ifFileUploadNotAllowed = false;
        }else{
          this.ifFileUploadAllowed = false;
          this.ifFileUploadNotAllowed = true;
        }
      }
     

      
    }else if(this.selectedItem.CurrentStage == stages.Stage3){    
      this.setStageFalse();
      this.setActiveTab(Tabs.Audit_Reviews_and_clearance);
      this.ifStage3 = true;
      this.ifStage3Active = true;
      if(this.selectedItem.Stage3CurrentState == kpmg.Manager){  
        this.setKpmgActiveTab(kpmgTabs.Manager);     
      }
      if(this.selectedItem.Stage3CurrentState == kpmg.Director){   
        this.setKpmgActiveTab(kpmgTabs.Director);
      }
      if(this.selectedItem.Stage3CurrentState == kpmg.Final){   
        this.setKpmgActiveTab(kpmgTabs.Final);
      }

      if(this.isKPMGTeamMember || this.isSuperuser){
        this.ifFileUploadAllowed = true;
        this.ifFileUploadNotAllowed = false;
      }else{
        this.ifFileUploadAllowed = false;
        this.ifFileUploadNotAllowed = true;
      }      
    }else if(this.selectedItem.CurrentStage == stages.Stage4){
      this.setStageFalse();
      this.setActiveTab(Tabs.FS_signoff);
      this.ifStage4 = true;
      this.ifStage4Active = true;

      if(this.selectedItem.Stage4CurrentState == Signoff.GSFS){  
        this.setStg4ActiveTab(SignoffTabIndexs.GSFS);  
      }
      if(this.selectedItem.Stage4CurrentState == Signoff.MAFP){   
        this.setStg4ActiveTab(SignoffTabIndexs.MAFP);
      }
      if(this.selectedItem.Stage4CurrentState == Signoff.KPMG){   
        this.setStg4ActiveTab(SignoffTabIndexs.Auditor);
      }

      if(this.isGSTeamMember || this.isSuperuser){
        this.ifFileUploadAllowed = true;
        this.ifFileUploadNotAllowed = false;
      }else{
        this.ifFileUploadAllowed = false;
        this.ifFileUploadNotAllowed = true;
      }
    }else if(this.selectedItem.CurrentStage == stages.Stage5){
      this.setStageFalse();
      this.setActiveTab(Tabs.FS_Archiving); 
      this.ifStage5 = true;
      this.ifStage5Active = true;
      
      if((this.isGSTeamMember || this.isSuperuser) && (this.selectedItem.CurrentStatus != Status.Approved)){
        this.ifFileUploadAllowed = true;
        this.ifFileUploadNotAllowed = false;
      }else{
        this.ifFileUploadAllowed = false;
        this.ifFileUploadNotAllowed = true;
      }
    }
  }

  onChangeTabSelect($event:any){   
     
    this.setActiveStageFalse();
    if($event.index == Tabs.MAFP_Data_Submission){
      this.ifStage1Active = true;
      if(this.selectedItem.CurrentStage == stages.Stage1){
        if(this.isGSTeamMember || this.isSuperuser){
          this.ifFileUploadAllowed = true;
          this.ifFileUploadNotAllowed = false;
        }else{
          this.ifFileUploadAllowed = false;
          this.ifFileUploadNotAllowed = true;
        }
      }
    }
    if($event.index == Tabs.EY_FS_Submission){
      this.ifStage2Active = true;
      if(this.selectedItem.CurrentStage == stages.Stage2){

        if(this.selectedItem.Stage2CurrentState == EY.EYFS){  
          this.setEYActiveTab(EYTabs.EYFS);   
          if(this.isGSTeamMember || this.isSuperuser || this.isEYTeamMember){
            this.ifFileUploadAllowed = true;
            this.ifFileUploadNotAllowed = false;
          }else{
            this.ifFileUploadAllowed = false;
            this.ifFileUploadNotAllowed = true;
          }
        }
        if(this.selectedItem.Stage2CurrentState == EY.FSTOKPMG){   
          this.setEYActiveTab(EYTabs.FSTOKPMG);
          if(this.isEYTeamMember || this.isSuperuser || this.isKPMGTeamMember){
            this.ifFileUploadAllowed = true;
            this.ifFileUploadNotAllowed = false;
          }else{
            this.ifFileUploadAllowed = false;
            this.ifFileUploadNotAllowed = true;
          }
        }
    }
    // else{
    //   this.ifFileUploadAllowed = false;
    //   this.ifFileUploadNotAllowed = true;
    // }
  }
    if($event.index == Tabs.Audit_Reviews_and_clearance){
      this.ifStage3Active = true;
      if(this.selectedItem.CurrentStage == stages.Stage3){
        if(this.isKPMGTeamMember || this.isSuperuser){
          this.ifFileUploadAllowed = true;
          this.ifFileUploadNotAllowed = false;
        }else{
          this.ifFileUploadAllowed = false;
          this.ifFileUploadNotAllowed = true;
        }
      }
    }
    if($event.index == Tabs.FS_signoff){
      this.ifStage4Active = true;
      if(this.selectedItem.CurrentStage == stages.Stage4){
        if(this.isBUTeamMember || this.isSuperuser){
          this.ifFileUploadAllowed = true;
          this.ifFileUploadNotAllowed = false;
        }else{
          this.ifFileUploadAllowed = false;
          this.ifFileUploadNotAllowed = true;
        }
      }
    }
    if($event.index == Tabs.FS_Archiving){
      this.ifStage5Active = true;        
      if((this.isGSTeamMember || this.isSuperuser) && (this.selectedItem.CurrentStage == stages.Stage5 && this.selectedItem.CurrentStatus != Status.Approved)){
        this.ifFileUploadAllowed = true;
        this.ifFileUploadNotAllowed = false;
      }else{
        this.ifFileUploadAllowed = false;
        this.ifFileUploadNotAllowed = true;
      }
    }
  }

  onChangeEYTabSelect($event:any){    
    this.tabEYIndex = $event.index;
  }

  setStagesDate(){
    // Date set
    let todaydate = new Date();
    if(this.selectedItem.TargetFSCompletionDate != null){
      let date = new Date(this.selectedItem.TargetFSCompletionDate);
      this.targetDate = date.getDate() + "-" + (date.getMonth() + 1) +"-"+date.getFullYear();
      //this.angForm.controls['TargetFSCompletionDate'].setValue({year: date.getFullYear(),month: date.getMonth() + 1,day: date.getDate()});
    }
    let deadlinedate = new Date();
    let actualsubmissiondate = new Date();
    let diffDays = 0;

    if (this.selectedItem.GSGAP != null){
      this.gsgap=Number(this.selectedItem.GSGAP);
    }
    // else{
    //   this.gsgap=0;
    // }
    if (this.selectedItem.ConsultantGAP != null){
      this.consultantgap=Number(this.selectedItem.ConsultantGAP);
    }
    // else{
    //   this.consultantgap=0;
    // }
    if (this.selectedItem.AuditorGAP != null){
      this.auditorgap=Number(this.selectedItem.AuditorGAP);
    }
    // else{
    //   this.auditorgap=0;
    // }
    if (this.selectedItem.CRTGAP != null){
      this.crtgap=Number(this.selectedItem.CRTGAP);
    }
    // else{
    //   this.crtgap=0;
    // }
    if(this.selectedItem.Stage1Deadline != null){
      this.angForm.controls['Stage1Deadline'].setValue(this.datepipe.transform(this.selectedItem.Stage1Deadline, 'dd-MM-yyyy'));
      deadlinedate = new Date(this.selectedItem.Stage1Deadline);
    }
    if(this.selectedItem.Stage1ActualSubmission != null){
      this.angForm.controls['Stage1ActualSubmission'].setValue(this.datepipe.transform(this.selectedItem.Stage1ActualSubmission, 'dd-MM-yyyy'));
      actualsubmissiondate = new Date(this.selectedItem.Stage1ActualSubmission);
    }else{      
      this.angForm.controls['Stage1ActualSubmission'].setValue(this.datepipe.transform(todaydate, 'dd-MM-yyyy'));
    }
    
    if(actualsubmissiondate > deadlinedate){
      diffDays = this.calculateDiffDays(deadlinedate,actualsubmissiondate);
    }      

    this.angForm.controls['Stage1GAP'].setValue(diffDays);

    let Stage1ConsultantVariableDeadline=new Date();
    if(this.selectedItem.Stage1ConsultantVariableDeadline != null){
      Stage1ConsultantVariableDeadline=new Date(this.selectedItem.Stage1ConsultantVariableDeadline);
    }
    
    let Stage1GSVariableDeadline = new Date();
    if (this.selectedItem.Stage1GSVariableDeadline != null) {
      Stage1GSVariableDeadline = new Date(this.selectedItem.Stage1GSVariableDeadline);
    }

    let Stage21ConsultantVariableDeadline = new Date();
    if (this.selectedItem.Stage21ConsultantVariableDeadlin != null) {
      Stage21ConsultantVariableDeadline = new Date(this.selectedItem.Stage21ConsultantVariableDeadlin);
    }

    let Stage21GSVariableDeadline = new Date();
    if (this.selectedItem.Stage21GSVariableDeadline != null) {
      Stage21GSVariableDeadline = new Date(this.selectedItem.Stage21GSVariableDeadline);
    }

    let Stage22GSVariableDeadline = new Date();
    if (this.selectedItem.Stage22GSVariableDeadline != null) {
      Stage22GSVariableDeadline = new Date(this.selectedItem.Stage22GSVariableDeadline);
    }

    let Stage31AuditorvariableDeadline = new Date();
    if (this.selectedItem.Stage31AuditorvariableDeadline != null) {
      Stage31AuditorvariableDeadline = new Date(this.selectedItem.Stage31AuditorvariableDeadline);
    }

    let Stage31GSVariableDeadline = new Date();
    if (this.selectedItem.Stage31GSVariableDeadline != null) {
      Stage31GSVariableDeadline = new Date(this.selectedItem.Stage31GSVariableDeadline);
    }

    let Stage31ConsultantVariableDeadline = new Date();
    if (this.selectedItem.Stage31ConsultantVariableDeadlin != null) {
      Stage31ConsultantVariableDeadline = new Date(this.selectedItem.Stage31ConsultantVariableDeadlin);
    }

    let Stage32AuditorvariableDeadline = new Date();
    if (this.selectedItem.Stage32AuditorvariableDeadline != null) {
      Stage32AuditorvariableDeadline = new Date(this.selectedItem.Stage32AuditorvariableDeadline);
    }

    let Stage32GSVariableDeadline = new Date();
    if (this.selectedItem.Stage32GSVariableDeadline != null) {
      Stage32GSVariableDeadline = new Date(this.selectedItem.Stage32GSVariableDeadline);
    }

    let Stage32ConsultantVariableDeadline = new Date();
    if (this.selectedItem.Stage32ConsultantVariableDeadlin != null) {
      Stage32ConsultantVariableDeadline = new Date(this.selectedItem.Stage32ConsultantVariableDeadlin);
    }

    let Stage33AuditorvariableDeadline = new Date();
    if (this.selectedItem.Stage33AuditorvariableDeadline != null) {
      Stage33AuditorvariableDeadline = new Date(this.selectedItem.Stage33AuditorvariableDeadline);
    }

    let Stage33GSVariableDeadline = new Date();
    if (this.selectedItem.Stage33GSVariableDeadline != null) {
      Stage33GSVariableDeadline = new Date(this.selectedItem.Stage33GSVariableDeadline);
    }

    let Stage33ConsultantVariableDeadline = new Date();
    if (this.selectedItem.Stage33ConsultantVariableDeadlin != null) {
      Stage33ConsultantVariableDeadline = new Date(this.selectedItem.Stage33ConsultantVariableDeadlin);
    }

    let Stage41GSVariableDeadline = new Date();
    if (this.selectedItem.Stage41GSVariableDeadline != null) {
      Stage41GSVariableDeadline = new Date(this.selectedItem.Stage41GSVariableDeadline);
    }

    let Stage41CRTVariableDeadline = new Date();
    if (this.selectedItem.Stage41CRTVariableDeadline != null) {
      Stage41CRTVariableDeadline = new Date(this.selectedItem.Stage41CRTVariableDeadline);
    }

    let Stage42GSVariableDeadline = new Date();
    if (this.selectedItem.Stage42GSVariableDeadline != null) {
      Stage42GSVariableDeadline = new Date(this.selectedItem.Stage42GSVariableDeadline);
    }

    let Stage42CRTVariableDeadline = new Date();
    if (this.selectedItem.Stage42CRTVariableDeadline != null) {
      Stage42CRTVariableDeadline = new Date(this.selectedItem.Stage42CRTVariableDeadline);
    }

    let Stage43GSVariableDeadline = new Date();
    if (this.selectedItem.Stage43GSVariableDeadline != null) {
      Stage43GSVariableDeadline = new Date(this.selectedItem.Stage43GSVariableDeadline);
    }

    let Stage43AuditorVariableDeadline = new Date();
    if (this.selectedItem.Stage43AuditorVariableDeadline != null) {
      Stage43AuditorVariableDeadline = new Date(this.selectedItem.Stage43AuditorVariableDeadline);
    }

    let Stage5GSVariableDeadline = new Date();
    if (this.selectedItem.Stage5GSVariableDeadline != null) {
      Stage5GSVariableDeadline = new Date(this.selectedItem.Stage5GSVariableDeadline);
    }

    let Stage5CRTVariableDeadline = new Date();
    if (this.selectedItem.Stage5CRTVariableDeadline != null) {
      Stage5CRTVariableDeadline = new Date(this.selectedItem.Stage5CRTVariableDeadline);
    }
    if(this.selectedItem.CurrentStage==stages.Stage1 && this.selectedItem.CurrentStatus=="Draft"){
        
        Stage1GSVariableDeadline = new Date(deadlinedate);
        this.selectedItem.Stage1GSVariableDeadline = new Date(deadlinedate);
        //let Stage1ConsultantVariableDeadline=new Date();
        Stage1ConsultantVariableDeadline = this.addTurnAroundDays(deadlinedate,Number(this.turnAroundDays.Stage1ConsultantApproval));
        this.angForm.controls['Stage1ConsultantVariableDeadline'].setValue(this.datepipe.transform(Stage1ConsultantVariableDeadline, 'dd-MM-yyyy'));
        this.selectedItem.Stage1ConsultantVariableDeadline =Stage1ConsultantVariableDeadline;// new Date(deadlinedate);
        
        Stage21ConsultantVariableDeadline=new Date(this.selectedItem.Stage2Deadline);
        this.angForm.controls['Stage21ConsultantVariableDeadline'].setValue(this.datepipe.transform(Stage21ConsultantVariableDeadline, 'dd-MM-yyyy'));

        //let Stage21GSVariableDeadline=new Date();
        Stage21GSVariableDeadline = this.addTurnAroundDays(Stage21ConsultantVariableDeadline,Number(this.turnAroundDays.Stage21GSReview));
        this.angForm.controls['Stage21GSVariableDeadline'].setValue(this.datepipe.transform(Stage21GSVariableDeadline, 'dd-MM-yyyy'));

        Stage22GSVariableDeadline=new Date(this.selectedItem.Stage21Deadline);
        this.angForm.controls['Stage22GSVariableDeadline'].setValue(this.datepipe.transform(Stage22GSVariableDeadline, 'dd-MM-yyyy'));

        Stage31AuditorvariableDeadline=new Date(this.selectedItem.Stage31Deadline);
        this.angForm.controls['Stage31AuditorvariableDeadline'].setValue(this.datepipe.transform(Stage31AuditorvariableDeadline, 'dd-MM-yyyy'));

        //let Stage31GSVariableDeadline=new Date();
        Stage31GSVariableDeadline = this.addTurnAroundDays(Stage31AuditorvariableDeadline,Number(this.turnAroundDays.Stage31GSReview));
        this.angForm.controls['Stage31GSVariableDeadline'].setValue(this.datepipe.transform(Stage31GSVariableDeadline, 'dd-MM-yyyy'));

        //let Stage31ConsultantVariableDeadline=new Date();
        Stage31ConsultantVariableDeadline = this.addTurnAroundDays(Stage31GSVariableDeadline,Number(this.turnAroundDays.Stage31ConsultantReview));
        this.angForm.controls['Stage31ConsultantVariableDeadline'].setValue(this.datepipe.transform(Stage31ConsultantVariableDeadline, 'dd-MM-yyyy'));

        Stage32AuditorvariableDeadline=new Date(this.selectedItem.Stage32Deadline);
        this.angForm.controls['Stage32AuditorvariableDeadline'].setValue(this.datepipe.transform(Stage32AuditorvariableDeadline, 'dd-MM-yyyy'));

        //let Stage32GSVariableDeadline=new Date();
        Stage32GSVariableDeadline = this.addTurnAroundDays(Stage32AuditorvariableDeadline,Number(this.turnAroundDays.Stage32GSReview));
        this.angForm.controls['Stage32GSVariableDeadline'].setValue(this.datepipe.transform(Stage32GSVariableDeadline, 'dd-MM-yyyy'));

        //let Stage32ConsultantVariableDeadline=new Date();
        Stage32ConsultantVariableDeadline = this.addTurnAroundDays(Stage32GSVariableDeadline,Number(this.turnAroundDays.Stage32ConsultantReview));
        this.angForm.controls['Stage32ConsultantVariableDeadline'].setValue(this.datepipe.transform(Stage32ConsultantVariableDeadline, 'dd-MM-yyyy'));

        //Stage33AuditorvariableDeadline=new Date(this.selectedItem.Stage33Deadline);
        Stage33AuditorvariableDeadline = this.addTurnAroundDays(Stage32ConsultantVariableDeadline,Number(this.turnAroundDays.Stage33AuditorReview));
        this.angForm.controls['Stage33AuditorvariableDeadline'].setValue(this.datepipe.transform(Stage33AuditorvariableDeadline, 'dd-MM-yyyy'));

        //let Stage33GSVariableDeadline=new Date();
        Stage33GSVariableDeadline = this.addTurnAroundDays(Stage33AuditorvariableDeadline,Number(this.turnAroundDays.Stage33GSReview));
        this.angForm.controls['Stage33GSVariableDeadline'].setValue(this.datepipe.transform(Stage33GSVariableDeadline, 'dd-MM-yyyy'));

        //let Stage33ConsultantVariableDeadline=new Date();
        Stage33ConsultantVariableDeadline = this.addTurnAroundDays(Stage33GSVariableDeadline,Number(this.turnAroundDays.Stage33ConsultantReview));
        this.angForm.controls['Stage33ConsultantVariableDeadline'].setValue(this.datepipe.transform(Stage33ConsultantVariableDeadline, 'dd-MM-yyyy'));

        Stage41GSVariableDeadline=new Date(this.selectedItem.Stage4SubmissionDate);
        this.angForm.controls['Stage41GSVariableDeadline'].setValue(this.datepipe.transform(Stage41GSVariableDeadline, 'dd-MM-yyyy'));

        //let Stage41CRTVariableDeadline=new Date();
        Stage41CRTVariableDeadline = this.addTurnAroundDays(Stage41GSVariableDeadline,Number(this.turnAroundDays.Stage41CRTReview));
        this.angForm.controls['Stage41CRTVariableDeadline'].setValue(this.datepipe.transform(Stage41CRTVariableDeadline, 'dd-MM-yyyy'));

        Stage42GSVariableDeadline=new Date(this.selectedItem.Stage42SubmissionDate);
        this.angForm.controls['Stage42GSVariableDeadline'].setValue(this.datepipe.transform(Stage42GSVariableDeadline, 'dd-MM-yyyy'));

        //let Stage42CRTVariableDeadline=new Date();
        Stage42CRTVariableDeadline = this.addTurnAroundDays(Stage42GSVariableDeadline,Number(this.turnAroundDays.Stage42CRTReview));
        this.angForm.controls['Stage42CRTVariableDeadline'].setValue(this.datepipe.transform(Stage42CRTVariableDeadline, 'dd-MM-yyyy'));

        //Stage43GSVariableDeadline=new Date(this.selectedItem.Stage43SubmissionDate);
        Stage43GSVariableDeadline = this.addTurnAroundDays(Stage42CRTVariableDeadline,Number(this.turnAroundDays.Stage43GSReview));
        this.angForm.controls['Stage43GSVariableDeadline'].setValue(this.datepipe.transform(Stage43GSVariableDeadline, 'dd-MM-yyyy'));

        //let Stage43AuditorVariableDeadline=new Date();
        Stage43AuditorVariableDeadline = this.addTurnAroundDays(Stage43GSVariableDeadline,Number(this.turnAroundDays.Stage43AuditorReview));
        this.angForm.controls['Stage43AuditorVariableDeadline'].setValue(this.datepipe.transform(Stage43AuditorVariableDeadline, 'dd-MM-yyyy'));

        Stage5GSVariableDeadline=new Date(this.selectedItem.Stage5SubmissionDate);
        this.angForm.controls['Stage5GSVariableDeadline'].setValue(this.datepipe.transform(Stage5GSVariableDeadline, 'dd-MM-yyyy'));

        //let Stage5CRTVariableDeadline=new Date();
        Stage5CRTVariableDeadline = this.addTurnAroundDays(Stage5GSVariableDeadline,Number(this.turnAroundDays.Stage5CRTReview));
        this.angForm.controls['Stage5CRTVariableDeadline'].setValue(this.datepipe.transform(Stage5CRTVariableDeadline, 'dd-MM-yyyy'));
        
        this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({        
          'Stage1ConsultantVariableDeadline': Stage1ConsultantVariableDeadline,
          'Stage1GSVariableDeadline' : deadlinedate,
          'Stage21ConsultantVariableDeadlin': Stage21ConsultantVariableDeadline,
          'Stage21GSVariableDeadline': Stage21GSVariableDeadline,
          'Stage22GSVariableDeadline': Stage22GSVariableDeadline,
          'Stage31AuditorvariableDeadline': Stage31AuditorvariableDeadline,
          'Stage31GSVariableDeadline': Stage31GSVariableDeadline,
          'Stage31ConsultantVariableDeadlin': Stage31ConsultantVariableDeadline,
          'Stage32AuditorvariableDeadline': Stage32AuditorvariableDeadline,
          'Stage32GSVariableDeadline': Stage32GSVariableDeadline,
          'Stage32ConsultantVariableDeadlin': Stage32ConsultantVariableDeadline,
          'Stage33AuditorvariableDeadline': Stage33AuditorvariableDeadline,
          'Stage33GSVariableDeadline': Stage33GSVariableDeadline,
          'Stage33ConsultantVariableDeadlin': Stage33ConsultantVariableDeadline,
          'Stage41GSVariableDeadline': Stage41GSVariableDeadline,
          'Stage41CRTVariableDeadline': Stage41CRTVariableDeadline,
          'Stage42GSVariableDeadline': Stage42GSVariableDeadline,
          'Stage42CRTVariableDeadline': Stage42CRTVariableDeadline,
          'Stage43GSVariableDeadline': Stage43GSVariableDeadline,
          'Stage43AuditorVariableDeadline': Stage43AuditorVariableDeadline,
          'Stage5GSVariableDeadline': Stage5GSVariableDeadline,
          'Stage5CRTVariableDeadline': Stage5CRTVariableDeadline,
          //  'GSGAP':Number(diffDays)
        }).then(async (item: any) => {
          console.log("Variable deadlines has been updated")
        }).catch((err: any) => { 
          Swal.fire("Failed to update","","info");
        });
      //}
    }
    let actualsubmission = new Date();
    if(this.selectedItem.ActionHolder==null || this.selectedItem.ActionHolder=="" || this.selectedItem.ActionHolder=="GS"){
      
      if(this.selectedItem.CurrentStage==stages.Stage1){
        if(this.calculateDiffDays(Stage1GSVariableDeadline,actualsubmission)>0){
          this.selectedItem.GSGAP=this.gsgap + this.calculateDiffDays(Stage1GSVariableDeadline,actualsubmission);
        }
      }
      if(this.selectedItem.CurrentStage==stages.Stage2 && this.selectedItem.Stage2CurrentState==EY.EYFS){
        if(this.calculateDiffDays(Stage21GSVariableDeadline,actualsubmission)>0){
          this.selectedItem.GSGAP=this.gsgap + this.calculateDiffDays(Stage21GSVariableDeadline,actualsubmission);
        }
      }
      if(this.selectedItem.CurrentStage==stages.Stage2 && this.selectedItem.Stage2CurrentState==EY.FSTOKPMG){
        if(this.calculateDiffDays(Stage22GSVariableDeadline,actualsubmission)>0){
          this.selectedItem.GSGAP=this.gsgap + this.calculateDiffDays(Stage22GSVariableDeadline,actualsubmission);
        }
      }
      if(this.selectedItem.CurrentStage==stages.Stage3 && this.selectedItem.Stage3CurrentState==kpmg.Manager){
        if(this.calculateDiffDays(Stage31GSVariableDeadline,actualsubmission)>0){
          this.selectedItem.GSGAP=this.gsgap + this.calculateDiffDays(Stage31GSVariableDeadline,actualsubmission);
        }
      }
      if(this.selectedItem.CurrentStage==stages.Stage3 && this.selectedItem.Stage3CurrentState==kpmg.Director){
        if(this.calculateDiffDays(Stage32GSVariableDeadline,actualsubmission)>0){
          this.selectedItem.GSGAP=this.gsgap + this.calculateDiffDays(Stage32GSVariableDeadline,actualsubmission);
        }
      }
      if(this.selectedItem.CurrentStage==stages.Stage3 && this.selectedItem.Stage3CurrentState==kpmg.Final){
        if(this.calculateDiffDays(Stage33GSVariableDeadline,actualsubmission)>0){
          this.selectedItem.GSGAP=this.gsgap + this.calculateDiffDays(Stage33GSVariableDeadline,actualsubmission);
        }
      }
      if(this.selectedItem.CurrentStage==stages.Stage4 && this.selectedItem.Stage4CurrentState==Signoff.GSFS){
        if(this.calculateDiffDays(Stage41GSVariableDeadline,actualsubmission)>0){
          this.selectedItem.GSGAP=this.gsgap + this.calculateDiffDays(Stage41GSVariableDeadline,actualsubmission);
        }
      }
      if(this.selectedItem.CurrentStage==stages.Stage4 && this.selectedItem.Stage4CurrentState==Signoff.MAFP){
        if(this.calculateDiffDays(Stage42GSVariableDeadline,actualsubmission)>0){
          this.selectedItem.GSGAP=this.gsgap + this.calculateDiffDays(Stage42GSVariableDeadline,actualsubmission);
        }
      }
      if(this.selectedItem.CurrentStage==stages.Stage4 && this.selectedItem.Stage4CurrentState==Signoff.KPMG){
        if(this.calculateDiffDays(Stage43GSVariableDeadline,actualsubmission)>0){
          this.selectedItem.GSGAP=this.gsgap + this.calculateDiffDays(Stage43GSVariableDeadline,actualsubmission);
        }
      }
      if(this.selectedItem.CurrentStage==stages.Stage5){
        if(this.calculateDiffDays(Stage5GSVariableDeadline,actualsubmission)>0){
          this.selectedItem.GSGAP=this.gsgap + this.calculateDiffDays(Stage5GSVariableDeadline,actualsubmission);
        }
      }
    }
    
    if(this.selectedItem.ActionHolder=="Consultant"){
      if(this.selectedItem.CurrentStage==stages.Stage1){
        if(this.calculateDiffDays(Stage1ConsultantVariableDeadline,actualsubmission)>0){
          this.selectedItem.ConsultantGAP=this.consultantgap + this.calculateDiffDays(Stage1ConsultantVariableDeadline,actualsubmission);
        }
      }
      if(this.selectedItem.CurrentStage==stages.Stage2 && this.selectedItem.Stage2CurrentState==EY.EYFS){
        
        if(this.calculateDiffDays(Stage21ConsultantVariableDeadline,actualsubmission)>0){
          this.selectedItem.ConsultantGAP=this.consultantgap + this.calculateDiffDays(Stage21ConsultantVariableDeadline,actualsubmission);
        }
      }
      if(this.selectedItem.CurrentStage==stages.Stage3 && this.selectedItem.Stage3CurrentState==kpmg.Manager){
        if(this.calculateDiffDays(Stage31ConsultantVariableDeadline,actualsubmission)>0){
          this.selectedItem.ConsultantGAP=this.consultantgap + this.calculateDiffDays(Stage31ConsultantVariableDeadline,actualsubmission);
        }
      }
      if(this.selectedItem.CurrentStage==stages.Stage3 && this.selectedItem.Stage3CurrentState==kpmg.Director){
        if(this.calculateDiffDays(Stage32ConsultantVariableDeadline,actualsubmission)>0){
          this.selectedItem.ConsultantGAP=this.consultantgap + this.calculateDiffDays(Stage32ConsultantVariableDeadline,actualsubmission);
        }
      }
      if(this.selectedItem.CurrentStage==stages.Stage3 && this.selectedItem.Stage3CurrentState==kpmg.Final){
        if(this.calculateDiffDays(Stage33ConsultantVariableDeadline,actualsubmission)>0){
          this.selectedItem.ConsultantGAP=this.consultantgap + this.calculateDiffDays(Stage33ConsultantVariableDeadline,actualsubmission);
        }
      }
    }

    if(this.selectedItem.ActionHolder=="Auditor"){
      if(this.selectedItem.CurrentStage==stages.Stage3 && this.selectedItem.Stage3CurrentState==kpmg.Manager){
        if(this.calculateDiffDays(Stage31AuditorvariableDeadline,actualsubmission)>0){
          this.selectedItem.AuditorGAP=this.auditorgap + this.calculateDiffDays(Stage31AuditorvariableDeadline,actualsubmission);
        }
      }
      if(this.selectedItem.CurrentStage==stages.Stage3 && this.selectedItem.Stage3CurrentState==kpmg.Director){
        if(this.calculateDiffDays(Stage32AuditorvariableDeadline,actualsubmission)>0){
          this.selectedItem.AuditorGAP=this.auditorgap + this.calculateDiffDays(Stage32AuditorvariableDeadline,actualsubmission);
        }
      }
      if(this.selectedItem.CurrentStage==stages.Stage3 && this.selectedItem.Stage3CurrentState==kpmg.Final){
        if(this.calculateDiffDays(Stage33AuditorvariableDeadline,actualsubmission)>0){
          this.selectedItem.AuditorGAP=this.auditorgap + this.calculateDiffDays(Stage33AuditorvariableDeadline,actualsubmission);
        }
      }
      if(this.selectedItem.CurrentStage==stages.Stage4 && this.selectedItem.Stage4CurrentState==Signoff.KPMG){
        if(this.calculateDiffDays(Stage43AuditorVariableDeadline,actualsubmission)>0){
          this.selectedItem.AuditorGAP=this.auditorgap + this.calculateDiffDays(Stage43AuditorVariableDeadline,actualsubmission);
        }
      }
    }

    if(this.selectedItem.ActionHolder=="CRT"){
      if(this.selectedItem.CurrentStage==stages.Stage4 && this.selectedItem.Stage4CurrentState==Signoff.GSFS){
        if(this.calculateDiffDays(Stage41CRTVariableDeadline,actualsubmission)>0){
          this.selectedItem.CRTGAP=this.crtgap + this.calculateDiffDays(Stage41CRTVariableDeadline,actualsubmission);
        }
      }
      if(this.selectedItem.CurrentStage==stages.Stage4 && this.selectedItem.Stage4CurrentState==Signoff.MAFP){
        if(this.calculateDiffDays(Stage42GSVariableDeadline,actualsubmission)>0){
          this.selectedItem.CRTGAP=this.crtgap + this.calculateDiffDays(Stage42CRTVariableDeadline,actualsubmission);
        }
      }
      if(this.selectedItem.CurrentStage==stages.Stage5){
        if(this.calculateDiffDays(Stage5CRTVariableDeadline,actualsubmission)>0){
          this.selectedItem.CRTGAP=this.crtgap + this.calculateDiffDays(Stage5CRTVariableDeadline,actualsubmission);
        }
      }
    }

    

    deadlinedate = new Date();
    actualsubmissiondate = new Date();
    diffDays = 0;

    if(this.selectedItem.Stage2Deadline != null){
      this.angForm.controls['Stage2Deadline'].setValue(this.datepipe.transform(this.selectedItem.Stage2Deadline, 'dd-MM-yyyy'));
      deadlinedate = new Date(this.selectedItem.Stage2Deadline);
    }
    
    if(this.selectedItem.Stage2ActualSubmission != null){
      this.angForm.controls['Stage2ActualSubmission'].setValue(this.datepipe.transform(this.selectedItem.Stage2ActualSubmission, 'dd-MM-yyyy'));
      actualsubmissiondate = new Date(this.selectedItem.Stage2ActualSubmission);
    }else{
      this.angForm.controls['Stage2ActualSubmission'].setValue(this.datepipe.transform(todaydate, 'dd-MM-yyyy'));
    }   
    if(actualsubmissiondate > deadlinedate){
      diffDays = this.calculateDiffDays(deadlinedate,actualsubmissiondate);
    }      
    this.angForm.controls['Stage2GAP'].setValue(diffDays); 
    
    deadlinedate = new Date();
    actualsubmissiondate = new Date();
    diffDays = 0;

    if(this.selectedItem.Stage21Deadline != null){
      this.angForm.controls['Stage21Deadline'].setValue(this.datepipe.transform(this.selectedItem.Stage21Deadline, 'dd-MM-yyyy'));
      deadlinedate = new Date(this.selectedItem.Stage21Deadline);
    }
    
    if(this.selectedItem.Stage21ActualSubmission != null){
      this.angForm.controls['Stage21ActualSubmission'].setValue(this.datepipe.transform(this.selectedItem.Stage21ActualSubmission, 'dd-MM-yyyy'));
      actualsubmissiondate = new Date(this.selectedItem.Stage21ActualSubmission);
    }else{
      this.angForm.controls['Stage21ActualSubmission'].setValue(this.datepipe.transform(todaydate, 'dd-MM-yyyy'));
    }   
    if(actualsubmissiondate > deadlinedate){
      diffDays = this.calculateDiffDays(deadlinedate,actualsubmissiondate);
    }      
    this.angForm.controls['Stage21GAP'].setValue(diffDays); 

    deadlinedate = new Date();
    actualsubmissiondate = new Date();
    diffDays = 0;

    if(this.selectedItem.Stage31Deadline != null){        
      this.angForm.controls['Stage31Deadline'].setValue(this.datepipe.transform(this.selectedItem.Stage31Deadline, 'dd-MM-yyyy'));
      deadlinedate = new Date(this.selectedItem.Stage31Deadline);
    } 
    if(this.selectedItem.Stage31ActualSubmission != null){        
      this.angForm.controls['Stage31ActualSubmission'].setValue(this.datepipe.transform(this.selectedItem.Stage31ActualSubmission, 'dd-MM-yyyy'));
      actualsubmissiondate = new Date(this.selectedItem.Stage31ActualSubmission);
    }else{
      this.angForm.controls['Stage31ActualSubmission'].setValue(this.datepipe.transform(todaydate, 'dd-MM-yyyy'));
    }
    if(actualsubmissiondate > deadlinedate){
      diffDays = this.calculateDiffDays(deadlinedate,actualsubmissiondate);
    }      
    this.angForm.controls['Stage31GAP'].setValue(diffDays); 
    
      
   
      deadlinedate = new Date();
      actualsubmissiondate = new Date();
      diffDays = 0;
      if(this.selectedItem.Stage32Deadline != null){
        this.angForm.controls['Stage32Deadline'].setValue(this.datepipe.transform(this.selectedItem.Stage32Deadline, 'dd-MM-yyyy'));
        deadlinedate = new Date(this.selectedItem.Stage32Deadline);
      } 
      if(this.selectedItem.Stage32ActualSubmission != null){
        this.angForm.controls['Stage32ActualSubmission'].setValue(this.datepipe.transform(this.selectedItem.Stage32ActualSubmission, 'dd-MM-yyyy'));
        actualsubmissiondate = new Date(this.selectedItem.Stage32ActualSubmission);
      }else{
        this.angForm.controls['Stage32ActualSubmission'].setValue(this.datepipe.transform(todaydate, 'dd-MM-yyyy'));
      }
      if(actualsubmissiondate > deadlinedate){
        diffDays = this.calculateDiffDays(deadlinedate,actualsubmissiondate);
      }  
      this.angForm.controls['Stage32GAP'].setValue(diffDays);

      deadlinedate = new Date();
      actualsubmissiondate = new Date();
      diffDays = 0;

      if(this.selectedItem.Stage33Deadline != null){
        this.angForm.controls['Stage33Deadline'].setValue(this.datepipe.transform(this.selectedItem.Stage33Deadline, 'dd-MM-yyyy'));
        deadlinedate = new Date(this.selectedItem.Stage33Deadline);
      } 
      if(this.selectedItem.Stage33ActualSubmission != null){
        this.angForm.controls['Stage33ActualSubmission'].setValue(this.datepipe.transform(this.selectedItem.Stage33ActualSubmission, 'dd-MM-yyyy'));
        actualsubmissiondate = new Date(this.selectedItem.Stage33ActualSubmission);
      }else{
        this.angForm.controls['Stage33ActualSubmission'].setValue(this.datepipe.transform(todaydate, 'dd-MM-yyyy'));
      }
      if(actualsubmissiondate > deadlinedate){
        diffDays = this.calculateDiffDays(deadlinedate,actualsubmissiondate);
      }  
      this.angForm.controls['Stage33GAP'].setValue(diffDays);

      
      
      deadlinedate = new Date();
      actualsubmissiondate = new Date();
      diffDays = 0;
      if(this.selectedItem.Stage4SubmissionDate != null){
        this.angForm.controls['Stage4Deadline'].setValue(this.datepipe.transform(this.selectedItem.Stage4SubmissionDate, 'dd-MM-yyyy'));
        deadlinedate = new Date(this.selectedItem.Stage4SubmissionDate);
      }      
      if(this.selectedItem.Stage4ActualSubmission != null){
        this.angForm.controls['Stage4ActualSubmission'].setValue(this.datepipe.transform(this.selectedItem.Stage4ActualSubmission, 'dd-MM-yyyy'));
        actualsubmissiondate = new Date(this.selectedItem.Stage4ActualSubmission);
      }else{
        this.angForm.controls['Stage4ActualSubmission'].setValue(this.datepipe.transform(todaydate, 'dd-MM-yyyy'));
      }
      if(actualsubmissiondate > deadlinedate){
        diffDays = this.calculateDiffDays(deadlinedate,actualsubmissiondate);
      }  
      this.angForm.controls['Stage4GAP'].setValue(diffDays);
      

      deadlinedate = new Date();
      actualsubmissiondate = new Date();
      diffDays = 0;

      if(this.selectedItem.Stage42SubmissionDate != null){
        this.angForm.controls['Stage42Deadline'].setValue(this.datepipe.transform(this.selectedItem.Stage42SubmissionDate, 'dd-MM-yyyy'));
        deadlinedate = new Date(this.selectedItem.Stage42SubmissionDate);
      }      
      if(this.selectedItem.Stage42ActualSubmission != null){
        this.angForm.controls['Stage42ActualSubmission'].setValue(this.datepipe.transform(this.selectedItem.Stage42ActualSubmission, 'dd-MM-yyyy'));
        actualsubmissiondate = new Date(this.selectedItem.Stage42ActualSubmission);
      }else{
        this.angForm.controls['Stage42ActualSubmission'].setValue(this.datepipe.transform(todaydate, 'dd-MM-yyyy'));
      }
      if(actualsubmissiondate > deadlinedate){
        diffDays = this.calculateDiffDays(deadlinedate,actualsubmissiondate);
      }  
      this.angForm.controls['Stage42GAP'].setValue(diffDays);

      deadlinedate = new Date();
      actualsubmissiondate = new Date();
      diffDays = 0;

      if(this.selectedItem.Stage43SubmissionDate != null){
        this.angForm.controls['Stage43Deadline'].setValue(this.datepipe.transform(this.selectedItem.Stage43SubmissionDate, 'dd-MM-yyyy'));
        deadlinedate = new Date(this.selectedItem.Stage43SubmissionDate);
      }      
      if(this.selectedItem.Stage43ActualSubmission != null){
        this.angForm.controls['Stage43ActualSubmission'].setValue(this.datepipe.transform(this.selectedItem.Stage43ActualSubmission, 'dd-MM-yyyy'));
        actualsubmissiondate = new Date(this.selectedItem.Stage43ActualSubmission);
      }else{
        this.angForm.controls['Stage43ActualSubmission'].setValue(this.datepipe.transform(todaydate, 'dd-MM-yyyy'));
      }
      if(actualsubmissiondate > deadlinedate){
        diffDays = this.calculateDiffDays(deadlinedate,actualsubmissiondate);
      }  
      this.angForm.controls['Stage43GAP'].setValue(diffDays);

      deadlinedate = new Date();
      actualsubmissiondate = new Date();
      diffDays = 0;

    if(this.selectedItem.Stage5SubmissionDate != null){
      this.angForm.controls['Stage5Deadline'].setValue(this.datepipe.transform(this.selectedItem.Stage5SubmissionDate, 'dd-MM-yyyy'));
      deadlinedate = new Date(this.selectedItem.Stage5SubmissionDate);
    }
    if(this.selectedItem.Stage5ActualSubmission != null){
      this.angForm.controls['Stage5ActualSubmission'].setValue(this.datepipe.transform(this.selectedItem.Stage5ActualSubmission, 'dd-MM-yyyy'));
      actualsubmissiondate = new Date(this.selectedItem.Stage5ActualSubmission);
    }else{
      this.angForm.controls['Stage5ActualSubmission'].setValue(this.datepipe.transform(todaydate, 'dd-MM-yyyy'));
    }
    if(actualsubmissiondate > deadlinedate){
      diffDays = this.calculateDiffDays(deadlinedate,actualsubmissiondate);
    }  
    this.angForm.controls['Stage5GAP'].setValue(diffDays); 

    deadlinedate = new Date();
    actualsubmissiondate = new Date();
    diffDays = 0;

  }

getDayDiff(startDate: Date, endDate: Date): number {
  const msInDay = 24 * 60 * 60 * 1000;  
  const days = (Math.round(Math.abs(Number(endDate) - Number(startDate)) / msInDay)) - 1;
  return days;
}



  
calculateDiffDays(startDate: Date, endDate: Date): number {
  // Create copies of the start and end dates to avoid modifying the originals
  let start = new Date(startDate);
  let end = new Date(endDate);

  // Reset the time portion to midnight to compare only the date part
  start.setHours(0, 0, 0, 0);
  end.setHours(0, 0, 0, 0);

  // Return 0 if both dates are the same
  if (start.getTime() === end.getTime()) {
      return 0;
  }

  let count = 0;
  let curDate = new Date(start);

  // Loop through the dates
  while (curDate < end) {
      const dayOfWeek = curDate.getDay();
      if (dayOfWeek !== 0 && dayOfWeek !== 6) { // Exclude Sundays (0) and Saturdays (6)
          count++;
      }
      curDate.setDate(curDate.getDate() + 1);
  }

  return count;
}

calculateDiffDaysWithoutTime(startDate: Date, endDate: Date): number {
  // Normalize the dates to midnight to ignore the time component
  
  const start = new Date(startDate);
  const end = new Date(endDate);

  start.setHours(0, 0, 0, 0);
  end.setHours(0, 0, 0, 0);

  // Handle the case when both dates are the same
  if (start.getTime() === end.getTime()) {
    return 0; // Return 0 if both dates are the same
  }

  let count = 0;
  const increment = start < end ? 1 : -1; // Determine increment direction
  let curDate = new Date(start);

  // Ensure the loop runs until the date before `end` if incrementing forward, or after `end` if decrementing
  while ((increment === 1 && curDate < end) || (increment === -1 && curDate > end)) {
    const dayOfWeek = curDate.getDay();
    if (dayOfWeek !== 0 && dayOfWeek !== 6) { // Exclude Sundays (0) and Saturdays (6)
      count++;
    }
    curDate.setDate(curDate.getDate() + increment); // Move to the next date
  }

  return count;
}

calculateActualDiffDays(startDate: Date, endDate: Date): number {
  
  let count = 0;
  
  // Make copies of the dates to avoid modifying the original values
  let curDate = new Date(startDate);
  let finalDate = new Date(endDate);

  // Normalize the dates to ignore time
  curDate.setHours(0, 0, 0, 0);
  finalDate.setHours(0, 0, 0, 0);

  const increment = curDate <= finalDate ? 1 : -1; // Determine if counting forward or backward

  // Loop through the dates
  while ((increment === 1 && curDate <= finalDate) || (increment === -1 && curDate >= finalDate)) {
      const dayOfWeek = curDate.getDay();
      if (dayOfWeek !== 0 && dayOfWeek !== 6) { // Exclude Sunday (0) and Saturday (6)
        count += increment;
      }
      curDate.setDate(curDate.getDate() + increment);
  }

  return count;
}

  // Method to parse the date and convert to ISO format
  parseDate(dateString:any) {
    if (dateString) {
      let parsedDate = new Date(dateString);
      if (!isNaN(parsedDate.getTime())) {
        return parsedDate.toISOString();
      } else {
        console.error('Invalid date format:', dateString);
        return null; // Or handle the error as needed
      }
    }
    return null;
  }

  addTurnAroundDays(date: Date, days: number): Date {
    let result = new Date(date);
    let addedDays = 0;

    date.setHours(0, 0, 0, 0);

    while (addedDays < days) {
      result.setDate(result.getDate() + 1);
      // Check if the day is not Saturday (6) or Sunday (0)
      if (result.getDay() !== 0 && result.getDay() !== 6) {
        addedDays++;
      }
    }
    return result;
  }

  setStagesStatus() {
    
    // Status set
    this.angForm.controls['Stage1Status'].setValue(this.selectedItem.Stage1Status);

    if(this.selectedItem.Stage2Status!=null && this.selectedItem.Stage2Status!=""){
      this.angForm.controls['Stage2Status'].setValue(this.selectedItem.Stage2Status);
    }
    else{
      this.angForm.controls['Stage2Status'].setValue(this.selectedItem.CurrentStatus);
    }
    if(this.selectedItem.Stage21Status!=null && this.selectedItem.Stage21Status!=""){
      this.angForm.controls['Stage21Status'].setValue(this.selectedItem.Stage21Status); 
    }
    else{
      this.angForm.controls['Stage21Status'].setValue(this.selectedItem.CurrentStatus); 
    }
         
    this.angForm.controls['Stage3CurrentState'].setValue(this.selectedItem.Stage3CurrentState);   
    this.angForm.controls['Stage31Status'].setValue(this.selectedItem.Stage31Status);
    this.angForm.controls['Stage32Status'].setValue(this.selectedItem.Stage32Status);
    this.angForm.controls['Stage33Status'].setValue(this.selectedItem.Stage33Status);
    if(this.selectedItem.Stage4Status!=null && this.selectedItem.Stage4Status!=""){
      this.angForm.controls['Stage4Status'].setValue(this.selectedItem.Stage4Status);
    }
    else{
      this.angForm.controls['Stage4Status'].setValue(this.selectedItem.CurrentStatus);
    }
    if(this.selectedItem.Stage42Status!=null && this.selectedItem.Stage42Status!=""){
      this.angForm.controls['Stage42Status'].setValue(this.selectedItem.Stage42Status);
    }
    else{
      this.angForm.controls['Stage42Status'].setValue(this.selectedItem.CurrentStatus);
    }
    if(this.selectedItem.Stage43Status!=null && this.selectedItem.Stage43Status!=""){
      this.angForm.controls['Stage43Status'].setValue(this.selectedItem.Stage43Status);
    }else{
      this.angForm.controls['Stage43Status'].setValue(this.selectedItem.CurrentStatus);
    }
    
    this.angForm.controls['Stage4CurrentState'].setValue(this.selectedItem.Stage4CurrentState);  
    if(this.selectedItem.Stage5Status!=null && this.selectedItem.Stage5Status!=""){
      this.angForm.controls['Stage5Status'].setValue(this.selectedItem.Stage5Status);  
    } 
    else{
      this.angForm.controls['Stage5Status'].setValue(this.selectedItem.CurrentStatus);  
    }
    

    this.angForm.controls['Stg1Comments'].setValue('');   
    this.angForm.controls['Stg2Comments'].setValue('');     
    this.angForm.controls['Stg21Comments'].setValue(''); 
    this.angForm.controls['Stg31Comments'].setValue('');
    this.angForm.controls['Stg32Comments'].setValue('');
    this.angForm.controls['Stg33Comments'].setValue('');
    this.angForm.controls['Stg4Comments'].setValue('');
    this.angForm.controls['Stg42Comments'].setValue('');
    this.angForm.controls['Stg43Comments'].setValue('');
    this.angForm.controls['Stg5Comments'].setValue('');
  }

  setActiveTab(tab : Tabs){
    this.tabIndex = tab; 
  }

  setKpmgActiveTab(tab : kpmgTabs){
    this.tabKpmgIndex = tab; 
  }

  setEYActiveTab(tab : EYTabs){
    this.tabEYIndex = tab; 
  }

  setStg4ActiveTab(tab : SignoffTabIndexs){
    this.tabSignoffIndex = tab; 
  }

  setS21EYChildActiveTab(tab : S21TabIndex){
    this.tabS21EYChildIndex = tab; 
  }

  setActiveStageFalse(){
    this.ifStage1Active = false;
    this.ifStage2Active = false;
    this.ifStage3Active = false;
    this.ifStage4Active = false;
    this.ifStage5Active = false;
  }

  setStageFalse(){
    this.ifStage1 = false;
    this.ifStage2 = false;
    this.ifStage3 = false;
    this.ifStage4 = false;
    this.ifStage5 = false;
  }

  setPRProfilePicture(){
    //(window as any).global = window;
    this.service.getProfilePictureData().subscribe((res) => {
      if (res != null && res !== '') {
        let resultSet = res as any;
        resultSet = resultSet.d.results;  
        if(this.selectedItem.PreparerEmail !=null && this.selectedItem.PreparerEmail != "") {
          let getItem = resultSet.filter((items: { Title: string; }) => items.Title.toLowerCase() == this.selectedItem.PreparerEmail.toLowerCase());
          if(getItem.length > 0 && getItem[0].Attachments){
            this.Preparer = environment.tenant_URL + getItem[0].AttachmentFiles.results[0].ServerRelativeUrl;
          }
        }  
        if(this.selectedItem.ReviewerEmail !=null && this.selectedItem.ReviewerEmail != "") {          
          let getItem = resultSet.filter((items: { Title: string; }) => items.Title.toLowerCase() == this.selectedItem.ReviewerEmail.toLowerCase());
          if(getItem.length > 0 && getItem[0].Attachments){
            this.Reviewer = environment.tenant_URL + getItem[0].AttachmentFiles.results[0].ServerRelativeUrl;
          }
        }
      }
    });
    
    // spns.setup({
    //   sp: {
    //     baseUrl: environment.sp_URL
    //   }
    // });
    
    // let preparer = "", preparerLM = "";
   
    
    // if(this.selectedItem.PreparerEmail != null && this.selectedItem.PreparerEmail != ""){
    //   let loginName = "i:0#.f|membership|"+this.selectedItem.PreparerEmail+""; 
    //   let pictureURL = await spns.sp.profiles.getUserProfilePropertyFor(loginName,"PictureURL");
    //   this.Preparer = pictureURL;
    // }    

  }

  async saveUpdateRequest(status: string, stage: string) {
    debugger;
    (window as any).global = window;
    if(stage == stages.Stage1){
      if(status == Status.Approved){
        this.processStage1(status, stages.Stage2, Status.Submitted);
      }
      if(status == Status.Rejected){
        if(this.angForm.value.Stg1Comments != null && this.angForm.value.Stg1Comments != ""){
          this.processStage1(status, stages.Stage1, Status.Rejected);
        }else{
          Swal.fire("Please enter your comment","","info");
        }
      }
      if(status == Status.Underreview){
        if(this.angForm.value.Stg1Comments != null && this.angForm.value.Stg1Comments != ""){
          this.processStage1(status, stages.Stage1, Status.Underreview);
        }else{
          Swal.fire("Please enter your comment","","info");
        }
      }
      if(status == Status.Submitted){    
      let getMasterData = await this.web.lists.getByTitle("MasterCheckList").items.select("ID","Title","OrderBy","IsRequired","ControlName")
      .filter("Title eq '"+ stages.Stage1 +"'").orderBy("OrderBy").getAll();

      let getCheckListData = await this.web.lists.getByTitle("FSTrackerCheckList").items.select("ID","Title","Stage1Data")
      .filter("Title eq '"+ String(this.itemid) +"'").get();
      
      if(getCheckListData.length > 0 && getMasterData.length > 0){
          const clItem = getCheckListData[0];
          if(clItem.Stage1Data != null && clItem.Stage1Data != ""){
            const jsonData = clItem.Stage1Data;
            const dataArray2 = JSON.parse(jsonData);
            const dataArray: any[] = [];
            getMasterData.forEach((mitem:any) => {
              dataArray.push({id:mitem.ControlName, checked:mitem.IsRequired});
            });
            const isValid = this.compareCheckedValues(dataArray, dataArray2);
            if (isValid) {
              this.processStage1(status, stages.Stage1, Status.Submitted);
            } else {
              Swal.fire("Please complete checklist items.","","info");
            }
            
          }else{
            Swal.fire("Please complete checklist items.","","info");
            return;
          }
        }else{
          Swal.fire("Please complete checklist items.","","info");
          return;
        }
      }
    }    
    if(stage == stages.Stage4){
      if(status == Status.Approved){
          this.processStage4(status, stages.Stage5, Status.Submitted);
      }
      if(status == Status.Rejected){        
        if(this.angForm.value.Stg4Comments != null && this.angForm.value.Stg4Comments != ""){
          this.processStage4(status, stages.Stage4, Status.Rejected);
        }else{
          Swal.fire("Please enter your comment","","info");
        }
      }
      if(status == Status.Underreview){        
        if(this.angForm.value.Stg4Comments != null && this.angForm.value.Stg4Comments != ""){
          this.processStage4(status, stages.Stage4, Status.Underreview);
        }else{
          Swal.fire("Please enter your comment","","info");
        }
      }
      if(status == Status.Submitted){
        let getCheckListData = await this.web.lists.getByTitle("FSTrackerCheckList").items.select("ID","Title","Stage4Data")
        .filter("Title eq '"+ String(this.itemid) +"'").get();
        if(getCheckListData.length > 0){
            const clItem = getCheckListData[0]; 
            if(clItem.Stage4Data != null && clItem.Stage4Data != ""){
              this.processStage4(status, stages.Stage4, Status.Submitted);
            }else{
              Swal.fire("Please complete checklist items.","","info");
              return;
            }
          }else{
            Swal.fire("Please complete checklist items.","","info");
            return;
          }

      }
    }
    if(stage == stages.Stage5){
      if(status == Status.Approved){
        this.processStage5(status, stages.Stage5, Status.Approved);
      }
      if(status == Status.Rejected){        
        if(this.angForm.value.Stg5Comments != null && this.angForm.value.Stg5Comments != ""){
          this.processStage5(status, stages.Stage5, Status.Rejected);
        }else{
          Swal.fire("Please enter your comment","","info");
        }
      }
      if(status == Status.Submitted){
        let getCheckListData = await this.web.lists.getByTitle("FSTrackerCheckList").items.select("ID","Title","Stage5Data")
        .filter("Title eq '"+ String(this.itemid) +"'").get();
        if(getCheckListData.length > 0){
          const clItem = getCheckListData[0]; 
          if(clItem.Stage5Data != null && clItem.Stage5Data != ""){
            this.processStage5(status, stages.Stage5, Status.Submitted);
          }
          else{
              Swal.fire("Please complete checklist items.","","info");
              return;
          }
        }
        else{
          Swal.fire("Please complete checklist items.","","info");
          return;
        }
      }
    }
    //this.initiateEditItem();
  }
  
  compareCheckedValues(array1: any[], array2: any[]): boolean {
    if (array1.length !== array2.length) {
      return false;
    }  
    for (let i = 0; i < array1.length; i++) {
      const id = array1[i].id;
      const checked1 = array1[i].checked;
      const checked2 = array2.find(obj => obj.id === id)?.checked;  
      if (checked1 !== checked2 && checked1) {
        return false;
      }
    }  
    return true;
  }

  async saveUpdateRequestStage2(status: string, stage: string, level: string) {
    (window as any).global = window;   
     
    if(stage == stages.Stage2){
      if(status == Status.Approved){        
        if(this.selectedItem.Stage2CurrentState == EY.EYFS){
          this.processStage2point1(status, stage, EY.FSTOKPMG);
        }
        if(this.selectedItem.Stage2CurrentState == EY.FSTOKPMG){
          this.processStage2point2(status, stages.Stage3, Status.Submitted);
        }
      }
      if(status == Status.Rejected){
        if(this.selectedItem.Stage2CurrentState == EY.EYFS){          
          if(this.angForm.value.Stg2Comments != null && this.angForm.value.Stg2Comments != ""){
            this.processStage2point1(status, stage, EY.EYFS);
          }else{
            Swal.fire("Please enter your comment","","info");
          }
        }        
        if(this.selectedItem.Stage2CurrentState == EY.FSTOKPMG){
          if(this.angForm.value.Stg21Comments != null && this.angForm.value.Stg21Comments != ""){
            this.processStage2point2(status, stage, Status.Rejected);
          }else{
            Swal.fire("Please enter your comment","","info");
          }
        }
      }
      if(status == Status.Underreview){
        if(this.selectedItem.Stage2CurrentState == EY.EYFS){          
          if(this.angForm.value.Stg2Comments != null && this.angForm.value.Stg2Comments != ""){
            this.processStage2point1(status, stage, EY.EYFS);
          }else{
            Swal.fire("Please enter your comment","","info");
          }
        }        
        if(this.selectedItem.Stage2CurrentState == EY.FSTOKPMG){
          if(this.angForm.value.Stg21Comments != null && this.angForm.value.Stg21Comments != ""){
            this.processStage2point2(status, stage, Status.Underreview);
          }else{
            Swal.fire("Please enter your comment","","info");
          }
        }
      }
      if(status == Status.Submitted){

        if(this.selectedItem.Stage2CurrentState == EY.EYFS){
          let getCheckListData = await this.web.lists.getByTitle("FSTrackerCheckList").items.select("ID","Title","Stage21Data")
          .filter("Title eq '"+ String(this.itemid) +"'").get();
          if(getCheckListData.length > 0){
              const clItem = getCheckListData[0]; 
              if(clItem.Stage21Data != null && clItem.Stage21Data != ""){
                this.processStage2point1(status, stage, EY.EYFS);
              }else{
                Swal.fire("Please complete checklist items.","","info");
                return;
              }
          }else{
            Swal.fire("Please complete checklist items.","","info");
            return;
          }
        }
        if(this.selectedItem.Stage2CurrentState == EY.FSTOKPMG){
          let getCheckListData = await this.web.lists.getByTitle("FSTrackerCheckList").items.select("ID","Title","Stage22Data")
          .filter("Title eq '"+ String(this.itemid) +"'").get();
          if(getCheckListData.length > 0){
              const clItem = getCheckListData[0]; 
              if(clItem.Stage22Data != null && clItem.Stage22Data != ""){
                this.processStage2point2(status, stages.Stage3, Status.Submitted);
              }else{
                Swal.fire("Please complete checklist items.","","info");
                return;
              }
            }else{
              Swal.fire("Please complete checklist items.","","info");
              return;
            }
        }

      }
    }
    //this.initiateEditItem();
  }

  async saveUpdateRequestStage3(status: string, stage: string, level: string) {
    (window as any).global = window;   
    
    if(stage == stages.Stage3){
      if(status == Status.Approved){
        if(this.selectedItem.Stage3CurrentState == kpmg.Manager){
          let getCheckListData = await this.web.lists.getByTitle("FSTrackerCheckList").items.select("ID","Title","Stage31Data")
          .filter("Title eq '"+ String(this.itemid) +"'").get();
          if(getCheckListData.length > 0){
              const clItem = getCheckListData[0]; 
              if(clItem.Stage31Data != null && clItem.Stage31Data != ""){
                this.processStage3point1(status, stage, kpmg.Director);
              }else{
                Swal.fire("Please complete checklist items.","","info");
                return;
              }
            }else{
              Swal.fire("Please complete checklist items.","","info");
              return;
            }
        }
        if(this.selectedItem.Stage3CurrentState == kpmg.Director){
          let getCheckListData = await this.web.lists.getByTitle("FSTrackerCheckList").items.select("ID","Title","Stage32Data")
          .filter("Title eq '"+ String(this.itemid) +"'").get();
          if(getCheckListData.length > 0){
              const clItem = getCheckListData[0]; 
              if(clItem.Stage32Data != null && clItem.Stage32Data != ""){
                this.processStage3point2(status, stage, kpmg.Final);
              }else{
                Swal.fire("Please complete checklist items.","","info");
                return;
              }
            }else{
              Swal.fire("Please complete checklist items.","","info");
              return;
            }
        }
        if(this.selectedItem.Stage3CurrentState == kpmg.Final){
          let getCheckListData = await this.web.lists.getByTitle("FSTrackerCheckList").items.select("ID","Title","Stage33Data")
          .filter("Title eq '"+ String(this.itemid) +"'").get();
          if(getCheckListData.length > 0){
              const clItem = getCheckListData[0]; 
              if(clItem.Stage33Data != null && clItem.Stage33Data != ""){
                this.processStage3point3(status, stages.Stage4, Status.Submitted);
              }else{
                Swal.fire("Please complete checklist items.","","info");
                return;
              }
            }else{
              Swal.fire("Please complete checklist items.","","info");
              return;
            }
        }
      }

      if(status == Status.Rejected){
        if(this.selectedItem.Stage3CurrentState == kpmg.Manager){          
          if(this.angForm.value.Stg31Comments != null && this.angForm.value.Stg31Comments != ""){
            this.processStage3point1(status, stage, kpmg.Manager);
          }else{
            Swal.fire("Please enter your comment","","info");
          }
        }
        if(this.selectedItem.Stage3CurrentState == kpmg.Director){          
          if(this.angForm.value.Stg32Comments != null && this.angForm.value.Stg32Comments != ""){
            this.processStage3point2(status, stage, kpmg.Director);
          }else{
            Swal.fire("Please enter your comment","","info");
          }
        }
        if(this.selectedItem.Stage3CurrentState == kpmg.Final){
          if(this.angForm.value.Stg33Comments != null && this.angForm.value.Stg33Comments != ""){
            this.processStage3point3(status, stage, Status.Rejected);
          }else{
            Swal.fire("Please enter your comment","","info");
          }
        }
      }

      if(status == Status.Underreview){
        if(this.selectedItem.Stage3CurrentState == kpmg.Manager){          
          if(this.angForm.value.Stg31Comments != null && this.angForm.value.Stg31Comments != ""){
            this.processStage3point1(status, stage, kpmg.Manager);
          }else{
            Swal.fire("Please enter your comment","","info");
          }
        }
        if(this.selectedItem.Stage3CurrentState == kpmg.Director){          
          if(this.angForm.value.Stg32Comments != null && this.angForm.value.Stg32Comments != ""){
            this.processStage3point2(status, stage, kpmg.Director);
          }else{
            Swal.fire("Please enter your comment","","info");
          }
        }
        if(this.selectedItem.Stage3CurrentState == kpmg.Final){
          if(this.angForm.value.Stg33Comments != null && this.angForm.value.Stg33Comments != ""){
            this.processStage3point3(status, stage, Status.Underreview);
          }else{
            Swal.fire("Please enter your comment","","info");
          }
        }
      }

      if(status == Status.ReqIteration){
        if(this.selectedItem.Stage3CurrentState == kpmg.Manager){
          this.processStage3point1(status, stage, kpmg.Manager);
        }
        if(this.selectedItem.Stage3CurrentState == kpmg.Director){
          this.processStage3point2(status, stage, kpmg.Director);
        }
        if(this.selectedItem.Stage3CurrentState == kpmg.Final){
          this.processStage3point3(status, stage, kpmg.Final);
        }
      }

      if(status == Status.Submitted){
        if(this.selectedItem.Stage3CurrentState == kpmg.Manager){
          let getCheckListData = await this.web.lists.getByTitle("FSTrackerCheckList").items.select("ID","Title","Stage31Data")
          .filter("Title eq '"+ String(this.itemid) +"'").get();
          if(this.selectedItem.ActionHolder=="Consultant"){
            this.processStage3point1(status, stage, kpmg.Manager);
          }
          else{
            if(getCheckListData.length > 0){
              const clItem = getCheckListData[0]; 
              if(clItem.Stage31Data != null && clItem.Stage31Data != ""){
                this.processStage3point1(status, stage, kpmg.Manager);
              }else{
                Swal.fire("Please complete checklist items.","","info");
                return;
              }
            }else{
              Swal.fire("Please complete checklist items.","","info");
              return;
            }
          }
          
        }
        if(this.selectedItem.Stage3CurrentState == kpmg.Director){          
          let getCheckListData = await this.web.lists.getByTitle("FSTrackerCheckList").items.select("ID","Title","Stage32Data")
          .filter("Title eq '"+ String(this.itemid) +"'").get();
          if(this.selectedItem.ActionHolder=="Consultant"){
            this.processStage3point2(status, stage, kpmg.Director);
          }
          else{
            if(getCheckListData.length > 0){
              const clItem = getCheckListData[0]; 
              if(clItem.Stage32Data != null && clItem.Stage32Data != ""){
                this.processStage3point2(status, stage, kpmg.Director);
              }else{
                Swal.fire("Please complete checklist items.","","info");
                return;
              }
            }else{
              Swal.fire("Please complete checklist items.","","info");
              return;
            }
          }
        }
        if(this.selectedItem.Stage3CurrentState == kpmg.Final){  
          let getCheckListData = await this.web.lists.getByTitle("FSTrackerCheckList").items.select("ID","Title","Stage33Data")
          .filter("Title eq '"+ String(this.itemid) +"'").get();
          if(this.selectedItem.ActionHolder=="Consultant"){
            this.processStage3point3(status, stage, Status.Submitted);
          }
          else{
            if(getCheckListData.length > 0){
              const clItem = getCheckListData[0]; 
              if(clItem.Stage33Data != null && clItem.Stage33Data != ""){
                this.processStage3point3(status, stage, Status.Submitted);
              }else{
                Swal.fire("Please complete checklist items.","","info");
                return;
              }
            }else{
              Swal.fire("Please complete checklist items.","","info");
              return;
            }
          }
        }
      }
    }
    //this.initiateEditItem();
  }

  // Stage wise processing
  processStage1(status: string, updatestage:string, updatedstatus:string){
    debugger;
   console.log(this.selectedItem);
    (window as any).global = window;

    let stage1deadline = new Date();
    let stage1cvdeadline=new Date();
    let stage1gvdeadline=new Date();
    let actualsubmission = new Date();
    let stage21deadline = new Date();

    if(this.selectedItem.Stage1Deadline != null){
      stage1deadline = new Date(this.selectedItem.Stage1Deadline);
    }
    if(this.selectedItem.Stage1ConsultantVariableDeadline != null){
      stage1cvdeadline = new Date(this.selectedItem.Stage1ConsultantVariableDeadline);
    }
    
    if(this.selectedItem.Stage1GSVariableDeadline != null){
      stage1gvdeadline = new Date(this.selectedItem.Stage1GSVariableDeadline);
    }
    if(this.selectedItem.Stage2Deadline != null){
      stage21deadline = new Date(this.selectedItem.Stage2Deadline);
    }
    
    if(updatestage == stages.Stage2){
      
      this.consultantgap=this.consultantgap+this.calculateDiffDays(stage1cvdeadline,actualsubmission);
      let ddays=this.calculateActualDiffDays(actualsubmission,stage21deadline);
      let Stage21ConsultantVariableDeadline=new Date();
      if(ddays-Number(this.turnAroundDays.Stage21ConsultantReview)<0){
        Stage21ConsultantVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage21ConsultantReview));
      }
      else{
        Stage21ConsultantVariableDeadline = stage21deadline;
      }

      this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
        'Stage1ActualSubmission':actualsubmission,
        'Stage1GAP':String(this.angForm.value.Stage1GAP),
        'Stage1Status':status,
        'CurrentStage':updatestage,
        'CurrentStatus':Status.Draft,
        'PendingWith':updatestage,
        'Stg1Comments':this.angForm.value.Stg1Comments,
        //'Stage2Status':updatedstatus,
        'Stage2CurrentState':EY.EYFS,
        'ConsultantGAP':this.consultantgap,
        //'Stage2Deadline':Stage21ConsultantVariableDeadline,
        'Stage21ConsultantVariableDeadlin':Stage21ConsultantVariableDeadline,
        //'ActionHolder':AH.EY
        'ActionHolder':AH.Consultant
      }).then(async (item: any) => {        
        await this.saveEmailHistoryData(stages.Stage1, status, "");
        await this.saveEmailHistoryData(updatestage, Status.Draft, EY.EYFS);
        this.saveFSHistoryData(status,this.angForm.value.Stage1Deadline,String(actualsubmission),this.angForm.value.Stg1Comments,stages.Stage1,"",String(stage1cvdeadline),String(this.calculateDiffDays(stage1cvdeadline,actualsubmission)));
        //this.initiateEditItem();
      }).catch((err: any) => {         
        Swal.fire("Failed to update","","info");
      });

    }else{
      if(status == Status.Submitted){
        debugger;
        let ddays=this.calculateActualDiffDays(actualsubmission,stage1cvdeadline);
        let Stage1ConsultantVariableDeadline=new Date();
        if(ddays-Number(this.turnAroundDays.Stage1ConsultantApproval)<0){
          Stage1ConsultantVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage1ConsultantApproval));
        }
        else{
          Stage1ConsultantVariableDeadline = new Date(this.selectedItem.Stage1ConsultantVariableDeadline);
        }
        this.gsgap=this.gsgap + this.calculateDiffDays(stage1gvdeadline,actualsubmission);
        
        this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
          'Stage1ActualSubmission':actualsubmission,
          'Stage1GAP':String(this.angForm.value.Stage1GAP),
          'Stage1Status':status,
          'CurrentStage':updatestage,
          'CurrentStatus':updatedstatus,
          'PendingWith':updatestage,
          'Stg1Comments':this.angForm.value.Stg1Comments,
          'GSGAP':this.gsgap,
          'Stage1ConsultantVariableDeadline':Stage1ConsultantVariableDeadline,
          //'ActionHolder':AH.EY
          'ActionHolder':AH.Consultant
        }).then(async (item: any) => {
          await this.saveEmailHistoryData(stages.Stage1, status, "");
          this.saveFSHistoryData(status,this.angForm.value.Stage1Deadline,String(actualsubmission),this.angForm.value.Stg1Comments,stages.Stage1,"",String(stage1gvdeadline),String(this.calculateDiffDays(stage1gvdeadline,actualsubmission)));
          
        }).catch((err: any) => { 
          Swal.fire("Failed to update","","info");
        });
      }else if(status == Status.Rejected){

        let ddays=this.calculateActualDiffDays(actualsubmission,stage1gvdeadline);
        let Stage1GSVariableDeadline=new Date();
        if(ddays-Number(this.turnAroundDays.Stage1ConsultantReject)<0){
          Stage1GSVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage1ConsultantReject));
        }
        else{
          Stage1GSVariableDeadline = new Date(this.selectedItem.Stage1GSVariableDeadline);
        }
        this.consultantgap=this.consultantgap + this.calculateDiffDays(stage1cvdeadline,actualsubmission);
        this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
          'Stage1ActualSubmission':actualsubmission,
          'Stage1GAP':String(this.angForm.value.Stage1GAP),
          'Stage1Status':status,
          'CurrentStage':updatestage,
          'CurrentStatus':updatedstatus,
          'PendingWith':updatestage,
          'Stg1Comments':this.angForm.value.Stg1Comments,
          'ActionHolder':AH.GS,
          //'Stage1Deadline':Stage1Deadline,
          'Stage1GSVariableDeadline':Stage1GSVariableDeadline,
          'ConsultantGAP':this.consultantgap,
        }).then(async (item: any) => {
          await this.saveEmailHistoryData(stages.Stage1, status, "");
          this.saveFSHistoryData(status,this.angForm.value.Stage1Deadline,String(actualsubmission),this.angForm.value.Stg1Comments,stages.Stage1,"",String(stage1cvdeadline),String(this.calculateDiffDays(stage1cvdeadline,actualsubmission)));
        }).catch((err: any) => { 
          Swal.fire("Failed to update","","info");
        });
      }else{
        this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
          'Stage1ActualSubmission':actualsubmission,
          'Stage1GAP':String(this.angForm.value.Stage1GAP),
          'Stage1Status':status,
          'CurrentStage':updatestage,
          'CurrentStatus':updatedstatus,
          'PendingWith':updatestage,
          //'ConsultantGAP':this.consultantgap,
          'Stg1Comments':this.angForm.value.Stg1Comments
        }).then(async (item: any) => {
          await this.saveEmailHistoryData(stages.Stage1, status, "");
          this.saveFSHistoryData(status,this.angForm.value.Stage1Deadline,String(actualsubmission),this.angForm.value.Stg1Comments,stages.Stage1,"","","");
        }).catch((err: any) => { 
          Swal.fire("Failed to update","","info");
        });
      }
    }
    
  }

  processStage2point1(status: string, stage: string, nextstage:string){
    (window as any).global = window;
    let actualsubmission = new Date();
    let stage2deadline = new Date();
    let stage21cvdeadline = new Date();
    let stage21gvdeadline = new Date();

    let stage22deadline = new Date();
    let stage22gvdeadline = new Date();

    if(this.selectedItem.Stage2Deadline != null){
      stage2deadline = new Date(this.selectedItem.Stage2Deadline);
    }
    if(this.selectedItem.Stage21Deadline != null){
      stage22deadline = new Date(this.selectedItem.Stage21Deadline);
    }
    if(this.selectedItem.Stage21ConsultantVariableDeadlin != null){
      stage21cvdeadline = new Date(this.selectedItem.Stage21ConsultantVariableDeadlin);
    }
    if(this.selectedItem.Stage21GSVariableDeadline != null){
      stage21gvdeadline = new Date(this.selectedItem.Stage21GSVariableDeadline);
    }

    if(this.selectedItem.Stage22GSVariableDeadline != null){
      stage22gvdeadline = new Date(this.selectedItem.Stage22GSVariableDeadline);
    }
    
    if(status == Status.Approved){
      this.gsgap=this.gsgap+this.calculateDiffDays(stage21gvdeadline,actualsubmission);

      let ddays=this.calculateActualDiffDays(actualsubmission,stage22gvdeadline);
      let Stage22GSVariableDeadline=new Date();
      if(ddays-Number(this.turnAroundDays.Stage22GSReview)<0){
        Stage22GSVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage22GSReview));
      }
      else{
        Stage22GSVariableDeadline = stage22gvdeadline;
      }

      this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({        
        'Stage2ActualSubmission':actualsubmission,
        'Stage2GAP':String(this.angForm.value.Stage2GAP),
        'Stage2Status':status,
        'CurrentStage':stage,
        'CurrentStatus':Status.Draft,
        'PendingWith':stage,
        'Stg2Comments':this.angForm.value.Stg2Comments,        
        'Stage2CurrentState':EY.FSTOKPMG,
        'GSGAP':this.gsgap,
        'ActionHolder':AH.GS,
        //'Stage21Deadline':Stage22GSVariableDeadline,
        'Stage22GSVariableDeadline':Stage22GSVariableDeadline
        //'Stage21Status':Status.Submitted
      }).then(async (item: any) => {
        await this.saveEmailHistoryData(stages.Stage2, status, EY.EYFS); // Done
        await this.saveEmailHistoryData(stage, Status.Draft, EY.FSTOKPMG);
        this.saveFSHistoryData(status,this.angForm.value.Stage2Deadline,String(actualsubmission),this.angForm.value.Stg2Comments,stages.Stage2,EY.EYFS,String(stage21gvdeadline),String(this.calculateDiffDays(stage21gvdeadline,actualsubmission)));
      }).catch((err: any) => { 
        Swal.fire("Failed to update","","info");
      });

    }else{
      if(status == Status.Submitted){
        
        let ddays=this.calculateActualDiffDays(actualsubmission,stage21gvdeadline);
        let Stage21GSVariableDeadline=new Date();
        if(ddays-Number(this.turnAroundDays.Stage21GSReview)<0){
          Stage21GSVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage21GSReview));
        }
        else{
          Stage21GSVariableDeadline = stage21gvdeadline;
        }
        this.consultantgap=this.consultantgap+this.calculateDiffDays(stage21cvdeadline,actualsubmission);

        this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
          'Stage2ActualSubmission':actualsubmission,
          'Stage2GAP':String(this.angForm.value.Stage2GAP),
          'Stage2Status':status,
          'CurrentStage':stage,
          'CurrentStatus':status,
          'PendingWith':stage,
          'Stg2Comments':this.angForm.value.Stg2Comments,
          'ConsultantGAP':this.consultantgap,
          'Stage21GSVariableDeadline':Stage21GSVariableDeadline,
          'ActionHolder':AH.GS
        }).then(async (item: any) => {
          await this.saveEmailHistoryData(stages.Stage2, status, EY.EYFS);
          this.saveFSHistoryData(status,this.angForm.value.Stage2Deadline,String(actualsubmission),this.angForm.value.Stg2Comments,stages.Stage2,EY.EYFS,String(stage21cvdeadline),String(this.calculateDiffDays(stage21cvdeadline,actualsubmission)));
        }).catch((err: any) => { 
          Swal.fire("Failed to update","","info");
        });
      }else if(status == Status.Rejected){
        
        let ddays=this.calculateActualDiffDays(actualsubmission,stage21cvdeadline);
        let Stage21ConsultantVariableDeadline=new Date();
        if(ddays-Number(this.turnAroundDays.Stage21ConsultantReject)<0){
          Stage21ConsultantVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage21ConsultantReject));
        }
        else{
          Stage21ConsultantVariableDeadline = stage21cvdeadline;
        }
        this.gsgap=this.gsgap+this.calculateDiffDays(stage21gvdeadline,actualsubmission);
        this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
          'Stage2ActualSubmission':actualsubmission,
          'Stage2GAP':String(this.angForm.value.Stage2GAP),
          'Stage2Status':status,
          'CurrentStage':stage,
          'CurrentStatus':status,
          'PendingWith':stage,
          'Stg2Comments':this.angForm.value.Stg2Comments,
          'GSGAP':this.gsgap,
          'Stage21ConsultantVariableDeadlin':Stage21ConsultantVariableDeadline,
          //'Stage2Deadline':Stage21ConsultantVariableDeadline,
          //'ActionHolder':AH.EY
          'ActionHolder':AH.Consultant
        }).then(async (item: any) => {
          await this.saveEmailHistoryData(stages.Stage2, status, EY.EYFS);
          this.saveFSHistoryData(status,this.angForm.value.Stage2Deadline,String(actualsubmission),this.angForm.value.Stg2Comments,stages.Stage2,EY.EYFS,String(stage21gvdeadline),String(this.calculateDiffDays(stage21gvdeadline,actualsubmission)));
        }).catch((err: any) => { 
          Swal.fire("Failed to update","","info");
        });
      }else{
        this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
          'Stage2ActualSubmission':actualsubmission,
          'Stage2GAP':String(this.angForm.value.Stage2GAP),
          'Stage2Status':status,
          'CurrentStage':stage,
          'CurrentStatus':status,
          'PendingWith':stage,
          'Stg2Comments':this.angForm.value.Stg2Comments
        }).then(async (item: any) => {
          await this.saveEmailHistoryData(stages.Stage2, status, EY.EYFS);
          this.saveFSHistoryData(status,this.angForm.value.Stage2Deadline,String(actualsubmission),this.angForm.value.Stg2Comments,stages.Stage2,EY.EYFS,"","");
        }).catch((err: any) => { 
          Swal.fire("Failed to update","","info");
        });
      }
    }
  }

  processStage2point2(status: string, updatestage:string, updatedstatus:string){
    (window as any).global = window;    
    let actualsubmission = new Date();
    let stage21deadline = new Date();
    if(this.selectedItem.Stage21Deadline != null){
      stage21deadline = new Date(this.selectedItem.Stage21Deadline);
    }
    let stage31deadline = new Date();
    if(this.selectedItem.Stage31Deadline != null){
      stage31deadline = new Date(this.selectedItem.Stage31Deadline);
    }
    let stage22gvdeadline = new Date();
    if(this.selectedItem.Stage22GSVariableDeadline != null){
      stage22gvdeadline = new Date(this.selectedItem.Stage22GSVariableDeadline);
    }

    let stage31avdeadline = new Date();
    if(this.selectedItem.Stage31AuditorvariableDeadline != null){
      stage31avdeadline = new Date(this.selectedItem.Stage31AuditorvariableDeadline);
    }
    
    if(updatestage == stages.Stage3){
      this.gsgap=this.gsgap+this.calculateDiffDays(stage22gvdeadline,actualsubmission);

      let ddays=this.calculateActualDiffDays(actualsubmission,stage31avdeadline);
      let Stage31AuditorvariableDeadline=new Date();
      if(ddays-Number(this.turnAroundDays.Stage31AuditorReview)<0){
        Stage31AuditorvariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage31AuditorReview));
      }
      else{
        Stage31AuditorvariableDeadline = stage31avdeadline;
      }
      this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
        'Stage21ActualSubmission':actualsubmission,
        'Stage21GAP':String(this.angForm.value.Stage2GAP),
        'Stage21Status':status,
        'CurrentStage':updatestage,
        'CurrentStatus':Status.Submitted,
        //'CurrentStatus':Status.Draft,
        'PendingWith':updatestage,
        'Stg21Comments':this.angForm.value.Stg21Comments,
        'Stage3CurrentState':kpmg.Manager,
        'Stage2CurrentState':EY.FSTOKPMG,
        'Stage31Status':Status.Submitted,
        'GSGAP':this.gsgap,
        //'Stage31Deadline':Stage31AuditorvariableDeadline,
        'Stage31AuditorvariableDeadline':Stage31AuditorvariableDeadline,
        //'ActionHolder':AH.KPMG
        'ActionHolder':AH.Auditor
      }).then(async (item: any) => {
        await this.saveEmailHistoryData(stages.Stage2, status, EY.FSTOKPMG);
        await this.saveEmailHistoryData(updatestage, Status.Submitted, kpmg.Manager);
        this.saveFSHistoryData(status,this.angForm.value.Stage21Deadline,String(actualsubmission),this.angForm.value.Stg21Comments,stages.Stage2,EY.FSTOKPMG,String(stage22gvdeadline),String(this.calculateDiffDays(stage22gvdeadline,actualsubmission)));
      }).catch((err: any) => { 
        Swal.fire("Failed to update","","info");
      });

    }else{
      if(status == Status.Submitted){
        this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
          'Stage21ActualSubmission':actualsubmission,
          'Stage21GAP':String(this.angForm.value.Stage21GAP),
          'Stage21Status':status,
          'CurrentStage':updatestage,
          'CurrentStatus':updatedstatus,
          'PendingWith':updatestage,
          'Stg21Comments':this.angForm.value.Stg21Comments,
          //'ActionHolder':AH.KPMG    
          'ActionHolder':AH.Auditor 
        }).then(async (item: any) => {
          await this.saveEmailHistoryData(stages.Stage2, status, EY.FSTOKPMG);
          this.saveFSHistoryData(status,this.angForm.value.Stage21Deadline,String(actualsubmission),this.angForm.value.Stg21Comments,stages.Stage2,EY.FSTOKPMG,"","");
        }).catch((err: any) => { 
          Swal.fire("Failed to update","","info");
        });
      }else if(status == Status.Rejected){
        this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
          'Stage21ActualSubmission':actualsubmission,
          'Stage21GAP':String(this.angForm.value.Stage21GAP),
          'Stage21Status':status,
          'CurrentStage':updatestage,
          'CurrentStatus':updatedstatus,
          'PendingWith':updatestage,
          'Stg21Comments':this.angForm.value.Stg21Comments,
          'ActionHolder':AH.GS     
        }).then(async (item: any) => {
          await this.saveEmailHistoryData(stages.Stage2, status, EY.FSTOKPMG);
          this.saveFSHistoryData(status,this.angForm.value.Stage21Deadline,String(actualsubmission),this.angForm.value.Stg21Comments,stages.Stage2,EY.FSTOKPMG,"","");
        }).catch((err: any) => { 
          Swal.fire("Failed to update","","info");
        });
      }else{
        this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
          'Stage21ActualSubmission':actualsubmission,
          'Stage21GAP':String(this.angForm.value.Stage21GAP),
          'Stage21Status':status,
          'CurrentStage':updatestage,
          'CurrentStatus':updatedstatus,
          'PendingWith':updatestage,
          'Stg21Comments':this.angForm.value.Stg21Comments
        }).then(async (item: any) => {
          await this.saveEmailHistoryData(stages.Stage2, status, EY.FSTOKPMG);
          this.saveFSHistoryData(status,this.angForm.value.Stage21Deadline,String(actualsubmission),this.angForm.value.Stg21Comments,stages.Stage2,EY.FSTOKPMG,"","");
        }).catch((err: any) => { 
          Swal.fire("Failed to update","","info");
        });
      }
    }
  }

  processStage3point1(status: string, stage: string, nextstage:string){

    (window as any).global = window;    
    let actualsubmission = new Date();

    let stage31deadline = new Date();
    let stage31avdeadline = new Date();
    let stage31gvdeadline = new Date();
    let stage31cvdeadline = new Date();
    let stage32avdeadline = new Date();

    if(this.selectedItem.Stage31Deadline != null){
      stage31deadline = new Date(this.selectedItem.Stage31Deadline);
    }
    if(this.selectedItem.Stage31AuditorvariableDeadline != null){
      stage31avdeadline = new Date(this.selectedItem.Stage31AuditorvariableDeadline);
    }
    if(this.selectedItem.Stage31GSVariableDeadline != null){
      stage31gvdeadline = new Date(this.selectedItem.Stage31GSVariableDeadline);
    }
    if(this.selectedItem.Stage31ConsultantVariableDeadlin != null){
      stage31cvdeadline = new Date(this.selectedItem.Stage31ConsultantVariableDeadlin);
    }
    if(this.selectedItem.Stage32AuditorvariableDeadline != null){
      stage32avdeadline = new Date(this.selectedItem.Stage32AuditorvariableDeadline);
    }

    let Stage31GSVariableDeadline=new Date();
    Stage31GSVariableDeadline = new Date(this.selectedItem.Stage31GSVariableDeadline);
    let Stage31ConsultantVariableDeadline=new Date();
    Stage31ConsultantVariableDeadline = new Date(this.selectedItem.Stage31ConsultantVariableDeadlin);
    let Stage31AuditorvariableDeadline=new Date();
    Stage31AuditorvariableDeadline = new Date(this.selectedItem.Stage31AuditorvariableDeadline);
    
    
    if(status == Status.Approved){
      this.auditorgap=this.auditorgap+this.calculateDiffDays(stage31avdeadline,actualsubmission);
      let Stage32AuditorvariableDeadline=new Date();
      let ddays=this.calculateActualDiffDays(actualsubmission,stage32avdeadline);
      if(ddays-Number(this.turnAroundDays.Stage32AuditorReview)<0){
        Stage32AuditorvariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage32AuditorReview));
      }
      else{
        Stage32AuditorvariableDeadline = stage32avdeadline;
      }
      this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({        
        'Stage31ActualSubmission':actualsubmission,
        'Stage31GAP':String(this.angForm.value.Stage31GAP),
        'Stage31Status':status,
        'CurrentStage':stage,
        'CurrentStatus':Status.Submitted,
        //'CurrentStatus':Status.Draft,
        'PendingWith':stage,
        'Stg31Comments':this.angForm.value.Stg31Comments,
        'Stage3CurrentState':nextstage,
        'Stage32Status': Status.Submitted,
        'AuditorGAP':this.auditorgap,
        //'ActionHolder':AH.KPMG
        'ActionHolder':AH.Auditor,
        'Stage32AuditorvariableDeadline':Stage32AuditorvariableDeadline
      }).then(async (item: any) => {
        await this.saveEmailHistoryData(stages.Stage3, status, kpmg.Manager);
        await this.saveEmailHistoryData(stage, Status.Submitted, kpmg.Director);
        this.saveFSHistoryData(status,this.angForm.value.Stage31Deadline,String(actualsubmission),this.angForm.value.Stg31Comments,stages.Stage3,kpmg.Manager,String(stage31avdeadline),String(this.calculateDiffDays(stage31avdeadline,actualsubmission)));
      }).catch((err: any) => { 
        Swal.fire("Failed to update","","info");
      });
    }else{
      let aHolder = AH.Auditor;
      let vardeadline:string = "";
      let calgap:string="";
      if(status == Status.Rejected){
        aHolder = AH.GS;
        let ddays=this.calculateActualDiffDays(actualsubmission,stage31gvdeadline);
        if(ddays-Number(this.turnAroundDays.Stage31GSReview)<0){
          Stage31GSVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage31GSReview));
        }
        vardeadline = String(stage31avdeadline);
        calgap=String(this.calculateDiffDays(stage31avdeadline,actualsubmission));
        this.auditorgap=this.auditorgap+this.calculateDiffDays(stage31avdeadline,actualsubmission);
      }
      else if(status == Status.Submitted || status == Status.Underreview){
        //aHolder = AH.KPMG;
        if(this.selectedItem.ActionHolder=="Consultant"){
          aHolder = AH.GS;
          this.consultantgap=this.consultantgap+this.calculateDiffDays(stage31cvdeadline,actualsubmission);
          vardeadline = String(stage31cvdeadline);
          calgap=String(this.calculateDiffDays(stage31cvdeadline,actualsubmission));
          let ddays=this.calculateActualDiffDays(actualsubmission,stage31gvdeadline);
          if(ddays-Number(this.turnAroundDays.Stage31GSReview)<0){
            Stage31GSVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage31GSReview));
          }
        }
        // else if(this.selectedItem.ActionHolder=="Auditor"){
        //   aHolder = AH.GS;
        //   this.auditorgap=this.auditorgap+this.calculateDiffDays(stage31avdeadline,actualsubmission);
        //   let ddays=this.calculateActualDiffDays(actualsubmission,stage31gvdeadline);
        //   if(ddays-Number(this.turnAroundDays.Stage31GSReview)<0){
        //     Stage31GSVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage31GSReview));
        //   }
        // }
        else if(this.selectedItem.ActionHolder=="GS"){
          aHolder = AH.Auditor;
          this.gsgap=this.gsgap+this.calculateDiffDays(stage31gvdeadline,actualsubmission);
          vardeadline = String(stage31gvdeadline);
          calgap=String(this.calculateDiffDays(stage31gvdeadline,actualsubmission));
          let ddays=this.calculateActualDiffDays(actualsubmission,stage31avdeadline);
          if(ddays-Number(this.turnAroundDays.Stage31AuditorReview)<0){
            Stage31AuditorvariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage31AuditorReview));
          }
        }
      }
      else if(status == Status.ReqIteration){
        //aHolder = AH.EY;
        aHolder = AH.Consultant;
        this.gsgap=this.gsgap+this.calculateDiffDays(stage31gvdeadline,actualsubmission);
        vardeadline = String(stage31gvdeadline);
        calgap=String(this.calculateDiffDays(stage31gvdeadline,actualsubmission));
        let ddays=this.calculateActualDiffDays(actualsubmission,stage31cvdeadline);
        if(ddays-Number(this.turnAroundDays.Stage31ConsultantReview)<0){
          Stage31ConsultantVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage31ConsultantReview));
        }
      }

      this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({ 
        'Stage31ActualSubmission':actualsubmission,
        'Stage31GAP':String(this.angForm.value.Stage31GAP),
        'Stage31Status':status,
        'CurrentStage':stage,
        'CurrentStatus':status,
        'PendingWith':stage,
        'Stg31Comments':this.angForm.value.Stg31Comments,
        'Stage3CurrentState':nextstage,
        'Stage31AuditorvariableDeadline':Stage31AuditorvariableDeadline,
        'AuditorGAP':this.auditorgap,
        'Stage31GSVariableDeadline':Stage31GSVariableDeadline,
        'GSGAP':this.gsgap,
        'Stage31ConsultantVariableDeadlin':Stage31ConsultantVariableDeadline,
        'ConsultantGAP':this.consultantgap,
        'ActionHolder': aHolder
      }).then(async (item: any) => {
        await this.saveEmailHistoryData(stages.Stage3, status, kpmg.Manager);
        this.saveFSHistoryData(status,this.angForm.value.Stage31Deadline,String(actualsubmission),this.angForm.value.Stg31Comments,stages.Stage3,kpmg.Manager,vardeadline,calgap);
      }).catch((err: any) => { 
        Swal.fire("Failed to update","","info");
      });
    } 
  }

  processStage3point2(status: string, stage: string, nextstage:string){
    (window as any).global = window;
    let actualsubmission = new Date();

    let stage32deadline = new Date();
    let stage32avdeadline = new Date();
    let stage32gvdeadline = new Date();
    let stage32cvdeadline = new Date();
    let stage33avdeadline = new Date();

    if(this.selectedItem.Stage32Deadline != null){
      stage32deadline = new Date(this.selectedItem.Stage32Deadline);
    }
    if(this.selectedItem.Stage32AuditorvariableDeadline != null){
      stage32avdeadline = new Date(this.selectedItem.Stage32AuditorvariableDeadline);
    }
    if(this.selectedItem.Stage32GSVariableDeadline != null){
      stage32gvdeadline = new Date(this.selectedItem.Stage32GSVariableDeadline);
    }
    if(this.selectedItem.Stage32ConsultantVariableDeadlin != null){
      stage32cvdeadline = new Date(this.selectedItem.Stage32ConsultantVariableDeadlin);
    }
    if(this.selectedItem.Stage33AuditorvariableDeadline != null){
      stage33avdeadline = new Date(this.selectedItem.Stage33AuditorvariableDeadline);
    }

    let Stage32GSVariableDeadline=new Date();
    Stage32GSVariableDeadline = new Date(this.selectedItem.Stage32GSVariableDeadline);
    let Stage32ConsultantVariableDeadline=new Date();
    Stage32ConsultantVariableDeadline = new Date(this.selectedItem.Stage32ConsultantVariableDeadlin);
    let Stage32AuditorvariableDeadline=new Date();
    Stage32AuditorvariableDeadline = new Date(this.selectedItem.Stage32AuditorvariableDeadline);
    
    if(status == Status.Approved){

      this.auditorgap=this.auditorgap+this.calculateDiffDays(stage32avdeadline,actualsubmission);
      let Stage33AuditorvariableDeadline=new Date();
      let ddays=this.calculateActualDiffDays(actualsubmission,stage33avdeadline);
      if(ddays-Number(this.turnAroundDays.Stage33AuditorReview)<0){
        Stage33AuditorvariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage33AuditorReview));
      }
      else{
        Stage33AuditorvariableDeadline = stage33avdeadline;
      }

      this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
        'Stage32ActualSubmission':actualsubmission,
        'Stage32GAP':String(this.angForm.value.Stage32GAP),
        'Stage32Status':status,
        'CurrentStage':stage,
        'CurrentStatus':Status.Submitted,
        //'CurrentStatus':Status.Draft,
        'PendingWith':stage,
        'Stg32Comments':this.angForm.value.Stg32Comments,
        'Stage3CurrentState':nextstage,
        'Stage33Status':Status.Submitted,
        //'ActionHolder':AH.KPMG
        'ActionHolder':AH.Auditor,
        'Stage33AuditorvariableDeadline':Stage33AuditorvariableDeadline,
        'AuditorGAP':this.auditorgap
      }).then(async (item: any) => {        
        await this.saveEmailHistoryData(stages.Stage3, status, kpmg.Director);
        await this.saveEmailHistoryData(stage, Status.Submitted, kpmg.Final);
        this.saveFSHistoryData(status,this.angForm.value.Stage32Deadline,String(actualsubmission),this.angForm.value.Stg32Comments,stages.Stage3,kpmg.Director,String(stage32avdeadline),String(this.calculateDiffDays(stage32avdeadline,actualsubmission)));
      }).catch((err: any) => { 
        Swal.fire("Failed to update","","info");
      });
    }else{
      let aHolder = AH.Auditor;
      let vardeadline:string = "";
      let calgap:string="";
      if(status == Status.Rejected){
        aHolder = AH.GS;
        let ddays=this.calculateActualDiffDays(actualsubmission,stage32gvdeadline);
        if(ddays-Number(this.turnAroundDays.Stage32GSReview)<0){
          Stage32GSVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage32GSReview));
        }
        this.auditorgap=this.auditorgap+this.calculateDiffDays(stage32avdeadline,actualsubmission);
        vardeadline = String(stage32avdeadline);
        calgap=String(this.calculateDiffDays(stage32avdeadline,actualsubmission));
      }
      else if(status == Status.Submitted || status == Status.Underreview){
        //aHolder = AH.KPMG;
        if(this.selectedItem.ActionHolder=="Consultant"){
          aHolder = AH.GS;
          this.consultantgap=this.consultantgap+this.calculateDiffDays(stage32cvdeadline,actualsubmission);
          vardeadline = String(stage32cvdeadline);
        calgap=String(this.calculateDiffDays(stage32cvdeadline,actualsubmission));
          let ddays=this.calculateActualDiffDays(actualsubmission,stage32gvdeadline);
          if(ddays-Number(this.turnAroundDays.Stage32GSReview)<0){
            Stage32GSVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage32GSReview));
          }
        }
        else if(this.selectedItem.ActionHolder=="GS"){
          aHolder = AH.Auditor;
          this.gsgap=this.gsgap+this.calculateDiffDays(stage32gvdeadline,actualsubmission);
          vardeadline = String(stage32gvdeadline);
          calgap=String(this.calculateDiffDays(stage32gvdeadline,actualsubmission));
          let ddays=this.calculateActualDiffDays(actualsubmission,stage32avdeadline);
          if(ddays-Number(this.turnAroundDays.Stage32AuditorReview)<0){
            Stage32AuditorvariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage32AuditorReview));
          }
        }
      }
      else if(status == Status.ReqIteration){
        //aHolder = AH.EY;
        aHolder = AH.Consultant;
        this.gsgap=this.gsgap+this.calculateDiffDays(stage32gvdeadline,actualsubmission);
        vardeadline = String(stage32gvdeadline);
          calgap=String(this.calculateDiffDays(stage32gvdeadline,actualsubmission));
        let ddays=this.calculateActualDiffDays(actualsubmission,stage32cvdeadline);
        if(ddays-Number(this.turnAroundDays.Stage32ConsultantReview)<0){
          Stage32ConsultantVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage32ConsultantReview));
        }
      }

      this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
        'Stage32ActualSubmission':actualsubmission,
        'Stage32GAP':String(this.angForm.value.Stage32GAP),
        'Stage32Status':status,
        'CurrentStage':stage,
        'CurrentStatus':status,
        'PendingWith':stage,
        'Stg32Comments':this.angForm.value.Stg32Comments,
        'Stage3CurrentState':nextstage,
        'Stage32AuditorvariableDeadline':Stage32AuditorvariableDeadline,
        'AuditorGAP':this.auditorgap,
        'Stage32GSVariableDeadline':Stage32GSVariableDeadline,
        'GSGAP':this.gsgap,
        'Stage32ConsultantVariableDeadlin':Stage32ConsultantVariableDeadline,
        'ConsultantGAP':this.consultantgap,
        'ActionHolder':aHolder
      }).then(async (item: any) => {
        await this.saveEmailHistoryData(stages.Stage3, status, kpmg.Director);
        this.saveFSHistoryData(status,this.angForm.value.Stage32Deadline,String(actualsubmission),this.angForm.value.Stg32Comments,stages.Stage3,kpmg.Director,vardeadline,calgap);
      }).catch((err: any) => { 
        Swal.fire("Failed to update","","info");
      });
    }
  }

  processStage3point3(status: string, updatestage:string, updatedstatus:string){
    (window as any).global = window;    
    let actualsubmission = new Date();

    let stage33deadline = new Date();
    let stage33avdeadline = new Date();
    let stage33gvdeadline = new Date();
    let stage33cvdeadline = new Date();
    let stage41gvdeadline = new Date();

    if(this.selectedItem.Stage33Deadline != null){
      stage33deadline = new Date(this.selectedItem.Stage33Deadline);
    }
    if(this.selectedItem.Stage33AuditorvariableDeadline != null){
      stage33avdeadline = new Date(this.selectedItem.Stage33AuditorvariableDeadline);
    }
    if(this.selectedItem.Stage33GSVariableDeadline != null){
      stage33gvdeadline = new Date(this.selectedItem.Stage33GSVariableDeadline);
    }
    if(this.selectedItem.Stage33ConsultantVariableDeadlin != null){
      stage33cvdeadline = new Date(this.selectedItem.Stage33ConsultantVariableDeadlin);
    }
    if(this.selectedItem.Stage41GSVariableDeadline != null){
      stage41gvdeadline = new Date(this.selectedItem.Stage41GSVariableDeadline);
    }

    let Stage33GSVariableDeadline=new Date();
    Stage33GSVariableDeadline = new Date(this.selectedItem.Stage33GSVariableDeadline);
    let Stage33ConsultantVariableDeadline=new Date();
    Stage33ConsultantVariableDeadline = new Date(this.selectedItem.Stage33ConsultantVariableDeadlin);
    let Stage33AuditorvariableDeadline=new Date();
    Stage33AuditorvariableDeadline = new Date(this.selectedItem.Stage33AuditorvariableDeadline);
    
    if(updatestage == stages.Stage4){

      this.auditorgap=this.auditorgap+this.calculateDiffDays(stage33avdeadline,actualsubmission);
      let Stage41GSVariableDeadline=new Date();
      let ddays=this.calculateActualDiffDays(actualsubmission,stage41gvdeadline);
      if(ddays-Number(this.turnAroundDays.Stage41GSReview)<0){
        Stage41GSVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage41GSReview));
      }
      else{
        Stage41GSVariableDeadline = stage41gvdeadline;
      }

      this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
        'Stage33ActualSubmission':actualsubmission,
        'Stage33GAP':String(this.angForm.value.Stage33GAP),
        'Stage33Status':status,
        'CurrentStage':updatestage,
        'CurrentStatus':Status.Draft,
        'PendingWith':updatestage,
        'Stg33Comments':this.angForm.value.Stg33Comments,
        'Stage4CurrentState':Signoff.GSFS,
        'ActionHolder':AH.GS,
        'Stage41GSVariableDeadline':Stage41GSVariableDeadline,
        'AuditorGAP':this.auditorgap
        //'Stage4Status':Status.Submitted
      }).then(async (item: any) => {
        await this.saveEmailHistoryData(stages.Stage3, status, kpmg.Final);
        await this.saveEmailHistoryData(updatestage, Status.Draft, Signoff.GSFS);
        this.saveFSHistoryData(status,this.angForm.value.Stage33Deadline,String(actualsubmission),this.angForm.value.Stg33Comments,stages.Stage3,kpmg.Final,String(stage33avdeadline),String(this.calculateDiffDays(stage33avdeadline,actualsubmission)));
      }).catch((err: any) => { 
        Swal.fire("Failed to update","","info");
      });
    }else{
      let aHolder = AH.Auditor;
      let vardeadline:string = "";
      let calgap:string="";
      if(status == Status.Rejected){
        aHolder = AH.GS;
        let ddays=this.calculateActualDiffDays(actualsubmission,stage33gvdeadline);
        if(ddays-Number(this.turnAroundDays.Stage33GSReview)<0){
          Stage33GSVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage33GSReview));
        }
        this.auditorgap=this.auditorgap+this.calculateDiffDays(stage33avdeadline,actualsubmission);
        vardeadline = String(stage33avdeadline);
        calgap=String(this.calculateDiffDays(stage33avdeadline,actualsubmission));
      }
      else if(status == Status.Submitted || status == Status.Underreview){
        //aHolder = AH.KPMG;
        if(this.selectedItem.ActionHolder=="Consultant"){
          aHolder = AH.GS;
          this.consultantgap=this.consultantgap+this.calculateDiffDays(stage33cvdeadline,actualsubmission);
          vardeadline = String(stage33cvdeadline);
          calgap=String(this.calculateDiffDays(stage33cvdeadline,actualsubmission));
          let ddays=this.calculateActualDiffDays(actualsubmission,stage33gvdeadline);
          if(ddays-Number(this.turnAroundDays.Stage33GSReview)<0){
            Stage33GSVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage33GSReview));
          }
        }
        else if(this.selectedItem.ActionHolder=="GS"){
          aHolder = AH.Auditor;
          this.gsgap=this.gsgap+this.calculateDiffDays(stage33gvdeadline,actualsubmission);
          vardeadline = String(stage33gvdeadline);
          calgap=String(this.calculateDiffDays(stage33gvdeadline,actualsubmission));
          let ddays=this.calculateActualDiffDays(actualsubmission,stage33avdeadline);
          if(ddays-Number(this.turnAroundDays.Stage33AuditorReview)<0){
            Stage33AuditorvariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage33AuditorReview));
          }
        }
      }
      else if(status == Status.ReqIteration){
        //aHolder = AH.EY;
        aHolder = AH.Consultant;
        this.gsgap=this.gsgap+this.calculateDiffDays(stage33gvdeadline,actualsubmission);
        vardeadline = String(stage33gvdeadline);
        calgap=String(this.calculateDiffDays(stage33gvdeadline,actualsubmission));
        let ddays=this.calculateActualDiffDays(actualsubmission,stage33cvdeadline);
        if(ddays-Number(this.turnAroundDays.Stage33ConsultantReview)<0){
          Stage33ConsultantVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage33ConsultantReview));
        }
      }

      this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
        'Stage33ActualSubmission':actualsubmission,
        'Stage33GAP':String(this.angForm.value.Stage33GAP),
        'Stage33Status':status,
        'CurrentStage':updatestage,
        'CurrentStatus':status,
        'PendingWith':updatestage,
        'Stg33Comments':this.angForm.value.Stg33Comments,
        'ActionHolder':aHolder,
        'Stage33AuditorvariableDeadline':Stage33AuditorvariableDeadline,
        'AuditorGAP':this.auditorgap,
        'Stage33GSVariableDeadline':Stage33GSVariableDeadline,
        'GSGAP':this.gsgap,
        'Stage33ConsultantVariableDeadlin':Stage33ConsultantVariableDeadline,
        'ConsultantGAP':this.consultantgap,    
      }).then(async (item: any) => {
        await this.saveEmailHistoryData(stages.Stage3, status, kpmg.Final);
        this.saveFSHistoryData(status,this.angForm.value.Stage33Deadline,String(actualsubmission),this.angForm.value.Stg33Comments,stages.Stage3,kpmg.Final,String(vardeadline),String(calgap));
      }).catch((err: any) => { 
        Swal.fire("Failed to update","","info");
      });
    }
  }

  async saveUpdateRequestStage4(status: string){
    (window as any).global = window;    
    
      if(status == Status.Approved){        
        if(this.selectedItem.Stage4CurrentState == Signoff.GSFS){
          this.processStage4point1(status, stages.Stage4, Signoff.MAFP);
        }
        if(this.selectedItem.Stage4CurrentState == Signoff.MAFP){
          this.processStage4point2(status, stages.Stage4, Signoff.KPMG);
        }
        if(this.selectedItem.Stage4CurrentState == Signoff.KPMG){
          this.processStage4point3(status, stages.Stage5, Status.Submitted);
        }
      }
      if(status == Status.Rejected){
        if(this.selectedItem.Stage4CurrentState == Signoff.GSFS){          
          if(this.angForm.value.Stg4Comments != null && this.angForm.value.Stg4Comments != ""){
            this.processStage4point1(status, stages.Stage4, Signoff.GSFS);
          }else{
            Swal.fire("Please enter your comment","","info");
          }
        }        
        if(this.selectedItem.Stage4CurrentState == Signoff.MAFP){
          if(this.angForm.value.Stg42Comments != null && this.angForm.value.Stg42Comments != ""){
            this.processStage4point2(status, stages.Stage4, Signoff.MAFP);
          }else{
            Swal.fire("Please enter your comment","","info");
          }
        }
        if(this.selectedItem.Stage4CurrentState == Signoff.KPMG){
          if(this.angForm.value.Stg43Comments != null && this.angForm.value.Stg43Comments != ""){
            this.processStage4point3(status, stages.Stage4, Signoff.KPMG);
          }else{
            Swal.fire("Please enter your comment","","info");
          }
        }
      }
      if(status == Status.Underreview){
        if(this.selectedItem.Stage2CurrentState == EY.EYFS){          
          if(this.angForm.value.Stg2Comments != null && this.angForm.value.Stg2Comments != ""){
            this.processStage2point1(status, stages.Stage4, EY.EYFS);
          }else{
            Swal.fire("Please enter your comment","","info");
          }
        }        
        if(this.selectedItem.Stage2CurrentState == EY.FSTOKPMG){
          if(this.angForm.value.Stg21Comments != null && this.angForm.value.Stg21Comments != ""){
            this.processStage2point2(status, stages.Stage4, Status.Underreview);
          }else{
            Swal.fire("Please enter your comment","","info");
          }
        }
      }
      if(status == Status.Submitted){
        if(this.selectedItem.Stage4CurrentState == Signoff.GSFS){
          let getCheckListData = await this.web.lists.getByTitle("FSTrackerCheckList").items.select("ID","Title","Stage4Data")
          .filter("Title eq '"+ String(this.itemid) +"'").get();
          if(getCheckListData.length > 0){
              const clItem = getCheckListData[0]; 
              if(clItem.Stage4Data != null && clItem.Stage4Data != ""){
                this.processStage4point1(status, stages.Stage4, Signoff.GSFS);
              }else{
                Swal.fire("Please complete checklist items.","","info");
                return;
              }
          }else{
            Swal.fire("Please complete checklist items.","","info");
            return;
          }
        }
        if(this.selectedItem.Stage4CurrentState == Signoff.MAFP){
          let getCheckListData = await this.web.lists.getByTitle("FSTrackerCheckList").items.select("ID","Title","Stage42Data")
          .filter("Title eq '"+ String(this.itemid) +"'").get();
          if(getCheckListData.length > 0){
              const clItem = getCheckListData[0]; 
              if(clItem.Stage42Data != null && clItem.Stage42Data != ""){
                this.processStage4point2(status, stages.Stage4, Signoff.MAFP);
              }else{
                Swal.fire("Please complete checklist items.","","info");
                return;
              }
            }else{
              Swal.fire("Please complete checklist items.","","info");
              return;
            }
        }
        if(this.selectedItem.Stage4CurrentState == Signoff.KPMG){
          let getCheckListData = await this.web.lists.getByTitle("FSTrackerCheckList").items.select("ID","Title","Stage43Data")
          .filter("Title eq '"+ String(this.itemid) +"'").get();
          if(getCheckListData.length > 0){
              const clItem = getCheckListData[0]; 
              if(clItem.Stage43Data != null && clItem.Stage43Data != ""){
                this.processStage4point3(status, stages.Stage4, Signoff.MAFP);
              }else{
                Swal.fire("Please complete checklist items.","","info");
                return;
              }
            }else{
              Swal.fire("Please complete checklist items.","","info");
              return;
            }
        }
      }    
    //this.initiateEditItem();
  }

  processStage4point1(status: string, stage: string, nextstage:string){
    (window as any).global = window;
    
    let actualsubmission = new Date(); 

    let stage41deadline = new Date();
    let stage41gvdeadline = new Date();
    let stage41crtdeadline = new Date();

    let stage42gvdeadline = new Date();

    if(this.selectedItem.Stage4SubmissionDate != null){
      stage41deadline = new Date(this.selectedItem.Stage4SubmissionDate);
    }
    if(this.selectedItem.Stage41GSVariableDeadline != null){
      stage41gvdeadline = new Date(this.selectedItem.Stage41GSVariableDeadline);
    }
    if(this.selectedItem.Stage41CRTVariableDeadline != null){
      stage41crtdeadline = new Date(this.selectedItem.Stage41CRTVariableDeadline);
    }

    if(this.selectedItem.Stage42SubmissionDate != null){
      stage42gvdeadline = new Date(this.selectedItem.Stage42SubmissionDate);
    }

    let Stage41GSVariableDeadline=new Date();
    Stage41GSVariableDeadline = new Date(this.selectedItem.Stage41GSVariableDeadline);

    let Stage41CRTVariableDeadline=new Date();
    Stage41CRTVariableDeadline = new Date(this.selectedItem.Stage41CRTVariableDeadline);


    if(status == Status.Approved){

      this.crtgap=this.crtgap+this.calculateDiffDays(stage41crtdeadline,actualsubmission);
      let ddays=this.calculateActualDiffDays(actualsubmission,stage42gvdeadline);
      let Stage42GSVariableDeadline=new Date();
      if(ddays-Number(this.turnAroundDays.Stage42GSReview)<0){
        Stage42GSVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage42GSReview));
      }
      else{
        Stage42GSVariableDeadline = stage42gvdeadline;
      }

      this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({        
        'Stage4ActualSubmission':actualsubmission,
        'Stage4GAP':String(this.angForm.value.Stage4GAP),
        'Stage4Status':status,
        'CurrentStage':stage,
        'CurrentStatus':Status.Draft,
        'PendingWith':stage,
        'Stg4Comments':this.angForm.value.Stg4Comments,        
        'Stage4CurrentState':Signoff.MAFP,
        'ActionHolder':AH.GS,
        'Stage42GSVariableDeadline':Stage42GSVariableDeadline,
        'CRTGAP':this.crtgap
        //'Stage21Status':Status.Submitted
      }).then(async (item: any) => {
        await this.saveEmailHistoryData(stages.Stage4, status, Signoff.GSFS); // Done        
        this.saveFSHistoryData(status,this.angForm.value.Stage4Deadline,String(actualsubmission),this.angForm.value.Stg4Comments,stages.Stage4, Signoff.GSFS,String(stage41crtdeadline),String(this.calculateDiffDays(stage41crtdeadline,actualsubmission)));
      }).catch((err: any) => { 
        Swal.fire("Failed to update","","info");
      });
    }
    else if(status == Status.Rejected){

      this.crtgap=this.crtgap+this.calculateDiffDays(stage41crtdeadline,actualsubmission);
      let ddays=this.calculateActualDiffDays(actualsubmission,stage41gvdeadline);
      if(ddays-Number(this.turnAroundDays.Stage41GSReview)<0){
        Stage41GSVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage41GSReview));
      }

      this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({        
        'Stage4ActualSubmission':actualsubmission,
        'Stage4GAP':String(this.angForm.value.Stage4GAP),
        'Stage4Status':status,
        'CurrentStage':stage,
        'CurrentStatus':Status.Rejected,
        'PendingWith':stage,
        'Stg4Comments':this.angForm.value.Stg4Comments,        
        //'Stage4CurrentState':Signoff.MAFP,
        'ActionHolder':AH.GS,
        'Stage41GSVariableDeadline':Stage41GSVariableDeadline,
        'CRTGAP':this.crtgap
        //'Stage21Status':Status.Submitted
      }).then(async (item: any) => {
        await this.saveEmailHistoryData(stages.Stage4, status, Signoff.GSFS); // Done        
        this.saveFSHistoryData(status,this.angForm.value.Stage4Deadline,String(actualsubmission),this.angForm.value.Stg4Comments,stages.Stage4, Signoff.GSFS,String(stage41crtdeadline),String(this.calculateDiffDays(stage41crtdeadline,actualsubmission)));
      }).catch((err: any) => { 
        Swal.fire("Failed to update","","info");
      });
    }
    else{

      this.gsgap=this.gsgap+this.calculateDiffDays(stage41gvdeadline,actualsubmission);
      let ddays=this.calculateActualDiffDays(actualsubmission,stage41crtdeadline);
      if(ddays-Number(this.turnAroundDays.Stage41CRTReview)<0){
        Stage41CRTVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage41CRTReview));
      }

      this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
        'Stage4ActualSubmission':actualsubmission,
        'Stage4GAP':String(this.angForm.value.Stage4GAP),
        'Stage4Status':status,
        'CurrentStage':stage,
        'CurrentStatus':status,
        'PendingWith':stage,
        'Stg4Comments':this.angForm.value.Stg4Comments,
        'ActionHolder':AH.CRT,
        'Stage41CRTVariableDeadline':Stage41CRTVariableDeadline,
        'GSGAP':this.gsgap
      }).then(async (item: any) => {
        await this.saveEmailHistoryData(stages.Stage4, status, Signoff.GSFS); // Done        
        this.saveFSHistoryData(status,this.angForm.value.Stage4Deadline,String(actualsubmission),this.angForm.value.Stg4Comments,stages.Stage4, Signoff.GSFS,String(stage41gvdeadline),String(this.calculateDiffDays(stage41gvdeadline,actualsubmission)));
      }).catch((err: any) => { 
        Swal.fire("Failed to update","","info");
      });      
    }
  }

  processStage4point2(status: string, stage: string, nextstage:string){
    (window as any).global = window;
    let actualsubmission = new Date();

    let stage42deadline = new Date();
    let stage42gvdeadline = new Date();
    let stage42crtdeadline = new Date();

    let stage43gvdeadline = new Date();

    if(this.selectedItem.Stage42SubmissionDate != null){
      stage42deadline = new Date(this.selectedItem.Stage42SubmissionDate);
    }
    if(this.selectedItem.Stage42GSVariableDeadline != null){
      stage42gvdeadline = new Date(this.selectedItem.Stage42GSVariableDeadline);
    }
    if(this.selectedItem.Stage42CRTVariableDeadline != null){
      stage42crtdeadline = new Date(this.selectedItem.Stage42CRTVariableDeadline);
    }

    if(this.selectedItem.Stage43SubmissionDate != null){
      stage43gvdeadline = new Date(this.selectedItem.Stage43SubmissionDate);
    }

    let Stage42GSVariableDeadline=new Date();
    Stage42GSVariableDeadline = new Date(this.selectedItem.Stage42GSVariableDeadline);

    let Stage42CRTVariableDeadline=new Date();
    Stage42CRTVariableDeadline = new Date(this.selectedItem.Stage42CRTVariableDeadline);

    
    if(status == Status.Approved){

      this.crtgap=this.crtgap+this.calculateDiffDays(stage42crtdeadline,actualsubmission);
      let ddays=this.calculateActualDiffDays(actualsubmission,stage43gvdeadline);
      let Stage43GSVariableDeadline=new Date();
      if(ddays-Number(this.turnAroundDays.Stage43GSReview)<0){
        Stage43GSVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage43GSReview));
      }
      else{
        Stage43GSVariableDeadline = stage43gvdeadline;
      }

      this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({        
        'Stage42ActualSubmission':actualsubmission,
        'Stage42GAP':String(this.angForm.value.Stage42GAP),
        'Stage42Status':status,
        'CurrentStage':stage,
        'CurrentStatus':Status.Draft,
        'PendingWith':stage,
        'Stg42Comments':this.angForm.value.Stg42Comments,        
        'Stage4CurrentState':Signoff.KPMG,
        'ActionHolder':AH.GS,
        'Stage43GSVariableDeadline':Stage43GSVariableDeadline,
        'CRTGAP':this.crtgap
        //'Stage21Status':Status.Submitted
      }).then(async (item: any) => {
        await this.saveEmailHistoryData(stages.Stage4, status, Signoff.MAFP); // Done        
        this.saveFSHistoryData(status,this.angForm.value.Stage42Deadline,String(actualsubmission),this.angForm.value.Stg42Comments,stages.Stage4, Signoff.MAFP,String(stage42crtdeadline),String(this.calculateDiffDays(stage42crtdeadline,actualsubmission)));
      }).catch((err: any) => { 
        Swal.fire("Failed to update","","info");
      });
    }
    else if(status == Status.Rejected){

      this.crtgap=this.crtgap+this.calculateDiffDays(stage42crtdeadline,actualsubmission);
      let ddays=this.calculateActualDiffDays(actualsubmission,stage42gvdeadline);
      if(ddays-Number(this.turnAroundDays.Stage42GSReview)<0){
        Stage42GSVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage42GSReview));
      }

      this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({        
        'Stage42ActualSubmission':actualsubmission,
        'Stage42GAP':String(this.angForm.value.Stage42GAP),
        'Stage42Status':status,
        'CurrentStage':stage,
        'CurrentStatus':Status.Rejected,
        'PendingWith':stage,
        'Stg42Comments':this.angForm.value.Stg42Comments,        
        //'Stage4CurrentState':Signoff.KPMG,
        'ActionHolder':AH.GS,
        'Stage42GSVariableDeadline':Stage42GSVariableDeadline,
        'CRTGAP':this.crtgap
        //'Stage21Status':Status.Submitted
      }).then(async (item: any) => {
        await this.saveEmailHistoryData(stages.Stage4, status, Signoff.MAFP); // Done        
        this.saveFSHistoryData(status,this.angForm.value.Stage42Deadline,String(actualsubmission),this.angForm.value.Stg42Comments,stages.Stage4, Signoff.MAFP,String(stage42crtdeadline),String(this.calculateDiffDays(stage42crtdeadline,actualsubmission)));
      }).catch((err: any) => { 
        Swal.fire("Failed to update","","info");
      });
    }
    else{

      this.gsgap=this.gsgap+this.calculateDiffDays(stage42gvdeadline,actualsubmission);
      let ddays=this.calculateActualDiffDays(actualsubmission,stage42crtdeadline);
      if(ddays-Number(this.turnAroundDays.Stage42CRTReview)<0){
        Stage42CRTVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage42CRTReview));
      }

      this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
        'Stage42ActualSubmission':actualsubmission,
        'Stage42GAP':String(this.angForm.value.Stage42GAP),
        'Stage42Status':status,
        'CurrentStage':stage,
        'CurrentStatus':status,
        'PendingWith':stage,
        'Stg42Comments':this.angForm.value.Stg42Comments,
        'ActionHolder':AH.CRT,
        'Stage42CRTVariableDeadline':Stage42CRTVariableDeadline,
        'GSGAP':this.gsgap
      }).then(async (item: any) => {
        await this.saveEmailHistoryData(stages.Stage4, status, Signoff.MAFP); // Done        
        this.saveFSHistoryData(status,this.angForm.value.Stage42Deadline,String(actualsubmission),this.angForm.value.Stg42Comments,stages.Stage4, Signoff.MAFP,String(stage42gvdeadline),String(this.calculateDiffDays(stage42gvdeadline,actualsubmission)));
      }).catch((err: any) => { 
        Swal.fire("Failed to update","","info");
      }); 
    }
  }

  processStage4point3(status: string, updatestage:string, updatedstatus:string){
    (window as any).global = window;    
    let actualsubmission = new Date();

    let stage43deadline = new Date();
    let stage43gvdeadline = new Date();
    let stage43avdeadline = new Date();

    let stage5gvdeadline = new Date();

    if(this.selectedItem.Stage43SubmissionDate != null){
      stage43deadline = new Date(this.selectedItem.Stage43SubmissionDate);
    }
    if(this.selectedItem.Stage43GSVariableDeadline != null){
      stage43gvdeadline = new Date(this.selectedItem.Stage43GSVariableDeadline);
    }
    if(this.selectedItem.Stage43AuditorVariableDeadline != null){
      stage43avdeadline = new Date(this.selectedItem.Stage43AuditorVariableDeadline);
    }

    if(this.selectedItem.Stage5SubmissionDate != null){
      stage5gvdeadline = new Date(this.selectedItem.Stage5SubmissionDate);
    }

    let Stage43GSVariableDeadline=new Date();
    Stage43GSVariableDeadline = new Date(this.selectedItem.Stage43GSVariableDeadline);

    let Stage43AuditorVariableDeadline=new Date();
    Stage43AuditorVariableDeadline = new Date(this.selectedItem.Stage43AuditorVariableDeadline);
    
    if(updatestage == stages.Stage5){

      this.auditorgap=this.auditorgap+this.calculateDiffDays(stage43avdeadline,actualsubmission);
      let ddays=this.calculateActualDiffDays(actualsubmission,stage5gvdeadline);
      let Stage5GSVariableDeadline=new Date();
      if(ddays-Number(this.turnAroundDays.Stage5GSReview)<0){
        Stage5GSVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage5GSReview));
      }
      else{
        Stage5GSVariableDeadline = stage5gvdeadline;
      }

      this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
        'Stage43ActualSubmission':actualsubmission,
        'Stage43GAP':String(this.angForm.value.Stage43GAP),
        'Stage43Status':status,
        'CurrentStage':updatestage,
        'CurrentStatus':Status.Draft,
        'PendingWith':updatestage,
        'Stg43Comments':this.angForm.value.Stg43Comments,
        'ActionHolder':AH.GS,
        'Stage5GSVariableDeadline':Stage5GSVariableDeadline,
        'AuditorGAP':this.auditorgap
        //'Stage5Status':Status.Submitted    
      }).then(async (item: any) => {   
        await this.saveEmailHistoryData(stages.Stage4, status, Signoff.KPMG);     
        await this.saveEmailHistoryData(updatestage, Status.Draft, "");
        this.saveFSHistoryData(status,this.angForm.value.Stage43Deadline,String(actualsubmission),this.angForm.value.Stg43Comments,stages.Stage4,Signoff.KPMG,String(stage43avdeadline),String(this.calculateDiffDays(stage43avdeadline,actualsubmission)));
      }).catch((err: any) => { 
        Swal.fire("Failed to update","","info");
      });
    }else{
      let aHolder = AH.GS;
      if(status == Status.Rejected){
        aHolder = AH.GS;
        this.auditorgap=this.auditorgap+this.calculateDiffDays(stage43avdeadline,actualsubmission);
        let ddays=this.calculateActualDiffDays(actualsubmission,stage43gvdeadline);
        if(ddays-Number(this.turnAroundDays.Stage43GSReview)<0){
          Stage43GSVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage43GSReview));
        }
      }
      else if(status == Status.Submitted){
        //aHolder = AH.KPMG;
        aHolder = AH.Auditor;
        this.gsgap=this.gsgap+this.calculateDiffDays(stage43gvdeadline,actualsubmission);
        let ddays=this.calculateActualDiffDays(actualsubmission,stage43avdeadline);
        if(ddays-Number(this.turnAroundDays.Stage43AuditorReview)<0){
          Stage43AuditorVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage43AuditorReview));
        }
      }
      this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
        'Stage43ActualSubmission':actualsubmission,
        'Stage43GAP':String(this.angForm.value.Stage43GAP),
        'Stage43Status':status,
        'CurrentStage':updatestage,
        'CurrentStatus':status,
        'PendingWith':updatestage,
        'Stg43Comments':this.angForm.value.Stg43Comments,
        'ActionHolder':aHolder,
        'GSGAP':this.gsgap,
        'AuditorGAP':this.auditorgap,
        'Stage43AuditorVariableDeadline':Stage43AuditorVariableDeadline,
        'Stage43GSVariableDeadline':Stage43GSVariableDeadline
        //'Stage5Status':Status.Submitted    
      }).then(async (item: any) => {        
        await this.saveEmailHistoryData(updatestage, status, Signoff.KPMG);
        this.saveFSHistoryData(status,this.angForm.value.Stage43Deadline,String(actualsubmission),this.angForm.value.Stg43Comments,stages.Stage4,Signoff.KPMG,String(stage43avdeadline),String(this.calculateDiffDays(stage43avdeadline,actualsubmission)));
      }).catch((err: any) => { 
        Swal.fire("Failed to update","","info");
      });
    }
  }

  processStage4(status: string, updatestage:string, updatedstatus:string){
    (window as any).global = window;    
    let actualsubmission = new Date();    
    if(updatestage == stages.Stage5){
      this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
        'Stage4ActualSubmission':actualsubmission,
        'Stage4GAP':String(this.angForm.value.Stage4GAP),
        'Stage4Status':status,
        'CurrentStage':updatestage,
        'CurrentStatus':Status.Draft,
        'PendingWith':updatestage,
        'Stg4Comments':this.angForm.value.Stg4Comments,
        'ActionHolder':AH.GS
        //'Stage5Status':Status.Submitted    
      }).then(async (item: any) => {
        await this.saveEmailHistoryData(stages.Stage4, status, "");
        await this.saveEmailHistoryData(updatestage, Status.Draft, "");
        this.saveFSHistoryData(status,this.angForm.value.Stage4Deadline,String(actualsubmission),this.angForm.value.Stg4Comments,stages.Stage4,"","","");
      }).catch((err: any) => { 
        Swal.fire("Failed to update","","info");
      });
    }else{
      if(status == Status.Submitted){
        this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
          'Stage4ActualSubmission':actualsubmission,
          'Stage4GAP':String(this.angForm.value.Stage4GAP),
          'Stage4Status':status,
          'CurrentStage':updatestage,
          'CurrentStatus':updatedstatus,
          'PendingWith':updatestage,
          'Stg4Comments':this.angForm.value.Stg4Comments,
          //'ActionHolder':AH.KPMG   
          'ActionHolder':AH.Auditor   
        }).then(async (item: any) => {
          await this.saveEmailHistoryData(stages.Stage4, status, "");
          this.saveFSHistoryData(status,this.angForm.value.Stage4Deadline,String(actualsubmission),this.angForm.value.Stg4Comments,stages.Stage4,"","","");
        }).catch((err: any) => { 
          Swal.fire("Failed to update","","info");
        });
      }else if(status == Status.Rejected){
        this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
          'Stage4ActualSubmission':actualsubmission,
          'Stage4GAP':String(this.angForm.value.Stage4GAP),
          'Stage4Status':status,
          'CurrentStage':updatestage,
          'CurrentStatus':updatedstatus,
          'PendingWith':updatestage,
          'Stg4Comments':this.angForm.value.Stg4Comments,
          'ActionHolder':AH.GS     
        }).then(async (item: any) => {
          await this.saveEmailHistoryData(stages.Stage4, status, "");
          this.saveFSHistoryData(status,this.angForm.value.Stage4Deadline,String(actualsubmission),this.angForm.value.Stg4Comments,stages.Stage4,"","","");
        }).catch((err: any) => { 
          Swal.fire("Failed to update","","info");
        });
      }else{
        this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
          'Stage4ActualSubmission':actualsubmission,
          'Stage4GAP':String(this.angForm.value.Stage4GAP),
          'Stage4Status':status,
          'CurrentStage':updatestage,
          'CurrentStatus':updatedstatus,
          'PendingWith':updatestage,
          'Stg4Comments':this.angForm.value.Stg4Comments
        }).then(async (item: any) => {
          await this.saveEmailHistoryData(stages.Stage4, status, "");
          this.saveFSHistoryData(status,this.angForm.value.Stage4Deadline,String(actualsubmission),this.angForm.value.Stg4Comments,stages.Stage4,"","","");
        }).catch((err: any) => { 
          Swal.fire("Failed to update","","info");
        });
      }
    }
  }

  processStage5(status: string, updatestage:string, updatedstatus:string){
    (window as any).global = window;    
    let actualsubmission = new Date();

    let stage5deadline = new Date();
    let stage5gvdeadline = new Date();
    let stage5crtdeadline = new Date();

    if(this.selectedItem.Stage5SubmissionDate != null){
      stage5deadline = new Date(this.selectedItem.Stage5SubmissionDate);
    }
    if(this.selectedItem.Stage5GSVariableDeadline != null){
      stage5gvdeadline = new Date(this.selectedItem.Stage5GSVariableDeadline);
    }
    if(this.selectedItem.Stage5CRTVariableDeadline != null){
      stage5crtdeadline = new Date(this.selectedItem.Stage5CRTVariableDeadline);
    }

    let Stage5GSVariableDeadline=new Date();
    Stage5GSVariableDeadline = new Date(this.selectedItem.Stage5GSVariableDeadline);

    let Stage5CRTVariableDeadline=new Date();
    Stage5CRTVariableDeadline = new Date(this.selectedItem.Stage5CRTVariableDeadline);

    if(status == Status.Submitted){

      this.gsgap=this.gsgap+this.calculateDiffDays(stage5gvdeadline,actualsubmission);
      let ddays=this.calculateActualDiffDays(actualsubmission,stage5crtdeadline);
      if(ddays-Number(this.turnAroundDays.Stage5CRTReview)<0){
        Stage5CRTVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage5CRTReview));
      }
      this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({      
        'Stage5ActualSubmission':actualsubmission,
        'Stage5GAP':String(this.angForm.value.Stage5GAP),
        'Stage5Status':status,
        'CurrentStage':updatestage,
        'CurrentStatus':updatedstatus,
        'PendingWith':updatestage,
        'Stg5Comments':this.angForm.value.Stg5Comments,
        'ActionHolder':AH.CRT,
        'Stage5CRTVariableDeadline':Stage5CRTVariableDeadline,
        'GSGAP':this.gsgap   
      }).then(async (item: any) => {
        await this.saveEmailHistoryData(stages.Stage5, status, "");
        this.saveFSHistoryData(status, this.angForm.value.Stage4Deadline, String(actualsubmission), this.angForm.value.Stg5Comments,stages.Stage5,"",String(stage5gvdeadline),String(this.calculateDiffDays(stage5gvdeadline,actualsubmission)));
      }).catch((err: any) => { 
        Swal.fire("Failed to update","","info");
      });
    }else if(status == Status.Rejected){
      this.crtgap=this.crtgap+this.calculateDiffDays(stage5crtdeadline,actualsubmission);
      let ddays=this.calculateActualDiffDays(actualsubmission,stage5gvdeadline);
      if(ddays-Number(this.turnAroundDays.Stage5GSReview)<0){
        Stage5GSVariableDeadline = this.addTurnAroundDays(actualsubmission,Number(this.turnAroundDays.Stage5GSReview));
      }
      this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({      
        'Stage5ActualSubmission':actualsubmission,
        'Stage5GAP':String(this.angForm.value.Stage5GAP),
        'Stage5Status':status,
        'CurrentStage':updatestage,
        'CurrentStatus':updatedstatus,
        'PendingWith':updatestage,
        'Stg5Comments':this.angForm.value.Stg5Comments,
        'ActionHolder':AH.GS,
        'Stage5GSVariableDeadline':Stage5GSVariableDeadline,
        'CRTGAP':this.crtgap     
      }).then(async (item: any) => {
        await this.saveEmailHistoryData(stages.Stage5, status, "");
        this.saveFSHistoryData(status, this.angForm.value.Stage4Deadline, String(actualsubmission), this.angForm.value.Stg5Comments,stages.Stage5,"",String(stage5crtdeadline),String(this.calculateDiffDays(stage5crtdeadline,actualsubmission)));
      }).catch((err: any) => { 
        Swal.fire("Failed to update","","info");
      });
    }else if(status == Status.Approved){
      this.crtgap=this.crtgap+this.calculateDiffDays(stage5crtdeadline,actualsubmission);
      this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({      
        'Stage5ActualSubmission':actualsubmission,
        'Stage5GAP':String(this.angForm.value.Stage5GAP),
        'Stage5Status':status,
        'CurrentStage':updatestage,
        'CurrentStatus':updatedstatus,
        'PendingWith':updatestage,
        'Stg5Comments':this.angForm.value.Stg5Comments,
        'CRTGAP':this.crtgap
      }).then(async (item: any) => {
        await this.saveEmailHistoryData(stages.Stage5, status, "");        
        this.saveFSHistoryData(status, this.angForm.value.Stage4Deadline, String(actualsubmission), this.angForm.value.Stg5Comments,stages.Stage5,"",String(stage5crtdeadline),String(this.calculateDiffDays(stage5crtdeadline,actualsubmission)));
      }).catch((err: any) => { 
        Swal.fire("Failed to update","","info");
      });
    }
  }

  async saveEmailHistoryData(stage:string, status:string, stage2or3level:string) {
    (window as any).global = window;
    //let actualsubmission = this.datepipe.transform(asperplan, 'dd-MM-yyyy');
    //let currentItem = this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid));    
    //const sp = spfi(environment.sp_URL);
    
    spns.setup({
      sp: {
        baseUrl: environment.sp_URL
      }
    });
    
    let preparer = "", preparerLM = "", PreparerBy = "", User = "", submissiondate = null, Comment = "No comments added", nextstage = "", stageduedate = null, nextstageduedate = null, subject = "";
    let users = [];
    let getTemplate = [];
    
    const currentItem: any = await this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).get();
    
    if(currentItem.PreparerEmail != null && currentItem.PreparerEmail != ""){
      let loginName = "i:0#.f|membership|"+currentItem.PreparerEmail+""; 
      preparerLM = await spns.sp.profiles.getUserProfilePropertyFor(loginName,"Manager");
      if(preparerLM != ""){
        preparerLM = preparerLM.split('|')[2];
        if(preparerLM.toLowerCase()== "biju.prakash@maf.ae")
        {
          preparerLM ="";
        }
      }
    }

    let EntityNameCode = currentItem.EntityName + " ("+ currentItem.Title +")";
    let tempstatus = status;
    if(stage == stages.Stage3){
      if(status == Status.Submitted && this.selectedItem.ActionHolder=="Consultant"){
        tempstatus = "Consultant "+status;
      }
    }

    let emailContent = "", emailTo = "", emailCC = "";
     //emailCC = preparerLM;
    if(stage == stages.Stage2 || stage == stages.Stage3 || stage == stages.Stage4){
      getTemplate = await this.web.lists.getByTitle("EmailTemplate")
          .items.select("Title", "EmailType", "EmailContent", "EmailTo", "EmailCC","Status","Stage2Or3Level").top(1)
          .filter("Title eq '" + stage + "' and Stage2Or3Level eq '"+ encodeURIComponent(stage2or3level) +"' and EmailType eq 'StatusChange' and Status eq '"+ tempstatus +"'").getAll();
    }else{
      getTemplate = await this.web.lists.getByTitle("EmailTemplate")
          .items.select("Title", "EmailType", "EmailContent", "EmailTo", "EmailCC","Status","Stage2Or3Level").top(1)
          .filter("Title eq '" + stage + "' and EmailType eq 'StatusChange' and Status eq '"+ tempstatus +"'").getAll();
    }

    let dueDate = null;
    if (getTemplate.length > 0) {
      if (stage == stages.Stage1) {
          stageduedate = this.datepipe.transform(currentItem.Stage1Deadline, 'dd-MM-yyyy');
          submissiondate = this.datepipe.transform(currentItem.Stage1ActualSubmission, 'dd-MM-yyyy');
          if(currentItem.Stg1Comments != null){
            Comment = currentItem.Stg1Comments; 
          }
          nextstage = stages.Stage1;
          nextstageduedate = this.datepipe.transform(currentItem.Stage1Deadline, 'dd-MM-yyyy');
          subject = "Stage 1 - " + String(EntityNameCode) + " has been "+ status +"";
          if(status == Status.Submitted){
            preparer = currentItem.Preparer;
            PreparerBy = currentItem.Preparer;
            User = "Consultant";
            emailTo = getTemplate[0].EmailTo;
            emailCC = currentItem.PreparerEmail + ";" + preparerLM + ";" + getTemplate[0].EmailCC + ";" + currentItem.ReviewerEmail;      
            nextstageduedate = this.datepipe.transform(currentItem.Stage1ConsultantVariableDeadline, 'dd-MM-yyyy');    
            //users = await this.web.lists.getByTitle("UserRoles").items.select("User", "Roles", "Status").filter("Status eq 'Active' and Roles eq 'EY'").getAll();
          }else if(status == Status.Draft){        
            preparer = currentItem.Preparer;
            PreparerBy = currentItem.Preparer;
            User = currentItem.Preparer;
            emailTo = currentItem.PreparerEmail;
            emailCC = preparerLM + ";" + getTemplate[0].EmailCC + ";" + currentItem.ReviewerEmail;
          }else if(status == Status.Underreview){        
            // preparer = currentItem.Preparer;
            // PreparerBy = currentItem.Preparer;
            preparer = "Consultant";
            PreparerBy = "Consultant";
            User = currentItem.Preparer;
            emailTo = currentItem.PreparerEmail;
            emailCC = preparerLM + ";" + getTemplate[0].EmailCC + ";" + currentItem.ReviewerEmail;
            subject = "Stage 1 - " + String(EntityNameCode) + " status changed to under review";
            nextstageduedate = this.datepipe.transform(currentItem.Stage1ConsultantVariableDeadline, 'dd-MM-yyyy');    
          }else if(status == Status.Approved){        
            preparer = "Consultant";
            PreparerBy = "Consultant";
            User = currentItem.Preparer;
            emailTo = currentItem.PreparerEmail;            
            emailCC = preparerLM + ";" + getTemplate[0].EmailCC + ";" + currentItem.ReviewerEmail;
            nextstage = EY.EYFS;
            nextstageduedate = this.datepipe.transform(currentItem.Stage21ConsultantVariableDeadline, 'dd-MM-yyyy');
          }else if(status == Status.Rejected){        
            preparer = "Consultant";
            PreparerBy = "Consultant";
            User = currentItem.Preparer;
            emailTo = currentItem.PreparerEmail;
            emailCC = preparerLM + ";" + getTemplate[0].EmailCC + ";" + currentItem.ReviewerEmail;
            nextstageduedate = this.datepipe.transform(currentItem.Stage1GSVariableDeadline, 'dd-MM-yyyy');
          }
      }
      if (stage == stages.Stage2) {
        if(stage2or3level == EY.EYFS){
          stageduedate = this.datepipe.transform(currentItem.Stage2Deadline, 'dd-MM-yyyy');
          submissiondate = this.datepipe.transform(currentItem.Stage1ActualSubmission, 'dd-MM-yyyy');          
          if(currentItem.Stg2Comments != null){
            Comment = currentItem.Stg2Comments; 
          }
          nextstage = EY.EYFS;
          nextstageduedate = this.datepipe.transform(currentItem.Stage2Deadline, 'dd-MM-yyyy');  
          subject = "Stage 2.1 - " + String(EntityNameCode) + " has been "+ status +"";      
          if(status == Status.Submitted){
            preparer = "Consultant";
            User = currentItem.Preparer;
            PreparerBy = "Consultant";
            emailTo = currentItem.PreparerEmail;
            emailCC = preparerLM + ";" + getTemplate[0].EmailCC + ";" + currentItem.ReviewerEmail;
            nextstageduedate = this.datepipe.transform(currentItem.Stage21GSVariableDeadline, 'dd-MM-yyyy'); 
            //users = await this.web.lists.getByTitle("UserRoles").items.select("User", "Roles", "Status").filter("Status eq 'Active' and Roles eq 'EY'").getAll();
          }else if(status == Status.Draft){        
            preparer = currentItem.Preparer;
            User = "Consultant";
            PreparerBy = currentItem.Preparer;
            emailTo = getTemplate[0].EmailTo;            
            emailCC =currentItem.PreparerEmail + ";" + preparerLM + ";" + getTemplate[0].EmailCC + ";" + currentItem.ReviewerEmail;
            subject = "Stage 2.1 - " + String(EntityNameCode) + " is ready for FS Review & Submission";
          }else if(status == Status.Underreview){        
            preparer = currentItem.Preparer;
            User = "Consultant";
            PreparerBy = currentItem.Preparer;
            emailTo = getTemplate[0].EmailTo;            
            emailCC =currentItem.PreparerEmail + ";" + preparerLM + ";" + getTemplate[0].EmailCC + ";" + currentItem.ReviewerEmail;
            subject = "Stage 2.1 - " + String(EntityNameCode) + " status changed to under review";
          }else if(status == Status.Approved){        
            preparer = currentItem.Preparer;
            User = "Consultant";
            PreparerBy = currentItem.Preparer;
            emailTo = getTemplate[0].EmailTo;
            emailCC = currentItem.PreparerEmail + ";" + preparerLM + ";" + getTemplate[0].EmailCC + ";" + currentItem.ReviewerEmail;
            nextstageduedate = this.datepipe.transform(currentItem.Stage22GSVariableDeadline, 'dd-MM-yyyy');
          }else if(status == Status.Rejected){        
            preparer = currentItem.Preparer;
            User = "Consultant";
            PreparerBy = currentItem.Preparer;
            emailTo = getTemplate[0].EmailTo;
            emailCC = currentItem.PreparerEmail + ";" + preparerLM + ";" + getTemplate[0].EmailCC + ";" + currentItem.ReviewerEmail;
            nextstageduedate = this.datepipe.transform(currentItem.Stage21ConsultantVariableDeadlin, 'dd-MM-yyyy');
          }
        }else if(stage2or3level == EY.FSTOKPMG){
          stageduedate = this.datepipe.transform(currentItem.Stage21Deadline, 'dd-MM-yyyy');
          submissiondate = this.datepipe.transform(currentItem.Stage21ActualSubmission, 'dd-MM-yyyy');          
          if(currentItem.Stg21Comments != null){
            Comment = currentItem.Stg21Comments; 
          }
          nextstage = EY.FSTOKPMG;
          nextstageduedate = this.datepipe.transform(currentItem.Stage21Deadline, 'dd-MM-yyyy'); 
          subject = "Stage 2.2 - " + String(EntityNameCode) + " has been "+ status +"";
          if(status == Status.Submitted){
            preparer = currentItem.Preparer;
            PreparerBy = currentItem.Preparer;
            User = "Auditor";
            emailTo = getTemplate[0].EmailTo;
            emailCC = currentItem.PreparerEmail + ";" +preparerLM + ";" + getTemplate[0].EmailCC + ";" + currentItem.ReviewerEmail; 
            nextstage = kpmg.Manager;
            nextstageduedate = this.datepipe.transform(currentItem.Stage31AuditorvariableDeadline, 'dd-MM-yyyy');
            //users = await this.web.lists.getByTitle("UserRoles").items.select("User", "Roles", "Status").filter("Status eq 'Active' and Roles eq 'EY'").getAll();
          }else if(status == Status.Draft){        
            preparer = "Consultant";
            User = currentItem.Preparer;
            PreparerBy = currentItem.Preparer;
            emailTo = currentItem.PreparerEmail;                  
            emailCC = getTemplate[0].EmailCC + ";" + preparerLM + ";" + currentItem.ReviewerEmail;    
            //subject = "Stage 2.2 - " + String(EntityNameCode) + " is ready for FS preparation";
          }else if(status == Status.Underreview){        
            preparer = "Auditor";
            User = currentItem.Preparer;
            PreparerBy = "Auditor";
            emailTo = currentItem.PreparerEmail;                  
            emailCC = getTemplate[0].EmailCC + ";" + preparerLM + ";" + currentItem.ReviewerEmail;    
            subject = "Stage 2.2 - " + String(EntityNameCode) + " status changed to under review";
          }else if(status == Status.Approved){        
            preparer = "Auditor";
            User = currentItem.Preparer;
            PreparerBy = "Auditor";
            emailTo = currentItem.PreparerEmail;
            emailCC = preparerLM + ";" + getTemplate[0].EmailCC + ";" + currentItem.ReviewerEmail;
          }else if(status == Status.Rejected){        
            preparer = "Auditor";
            PreparerBy = "Auditor";
            User = currentItem.Preparer;
            emailTo = currentItem.PreparerEmail;
            emailCC = preparerLM + ";" + getTemplate[0].EmailCC + ";" + currentItem.ReviewerEmail;
          }
        }
        // stageduedate = this.datepipe.transform(currentItem.Stage2Deadline, 'dd-MM-yyyy');
        // preparer = currentItem.Preparer;
      }
      if (stage == stages.Stage3) {
        if(stage2or3level == kpmg.Manager){
          //stageduedate = this.datepipe.transform(currentItem.Stage31Deadline, 'dd-MM-yyyy');
          stageduedate = this.datepipe.transform(currentItem.Stage31Deadline, 'dd-MM-yyyy');
          submissiondate = this.datepipe.transform(currentItem.Stage21ActualSubmission, 'dd-MM-yyyy');          
          if(currentItem.Stg31Comments != null){
            Comment = currentItem.Stg31Comments; 
          }
          nextstage = kpmg.Manager;
          nextstageduedate = this.datepipe.transform(currentItem.Stage31Deadline, 'dd-MM-yyyy');    
          subject = "Stage 3.1 - " + String(EntityNameCode) + " has been "+ status +"";
          if(status == Status.Submitted && this.selectedItem.ActionHolder=="Consultant"){
            preparer = "Consultant";
            User = currentItem.Preparer;
            PreparerBy = "Consultant";
            emailTo = currentItem.PreparerEmail; 
            emailCC = currentItem.PreparerEmail + ";" +preparerLM + ";" + getTemplate[0].EmailCC + ";" + currentItem.ReviewerEmail; 
            nextstageduedate = this.datepipe.transform(currentItem.Stage31GSVariableDeadline, 'dd-MM-yyyy');    
            //nextstage = kpmg.Manager;
            //nextstageduedate = this.datepipe.transform(currentItem.Stage31Deadline, 'dd-MM-yyyy');
            //users = await this.web.lists.getByTitle("UserRoles").items.select("User", "Roles", "Status").filter("Status eq 'Active' and Roles eq 'EY'").getAll();
          }else if(status == Status.Submitted && (this.selectedItem.ActionHolder=="GS" || this.selectedItem.ActionHolder=="Auditor")){
            preparer = currentItem.Preparer;
            PreparerBy = currentItem.Preparer;
            User = "Auditor";
            emailTo = getTemplate[0].EmailTo;
            emailCC = currentItem.PreparerEmail + ";" +preparerLM + ";" + getTemplate[0].EmailCC + ";" + currentItem.ReviewerEmail; 
            nextstageduedate = this.datepipe.transform(currentItem.Stage31AuditorvariableDeadline, 'dd-MM-yyyy');    
            //nextstage = kpmg.Manager;
            //nextstageduedate = this.datepipe.transform(currentItem.Stage31Deadline, 'dd-MM-yyyy');
            //users = await this.web.lists.getByTitle("UserRoles").items.select("User", "Roles", "Status").filter("Status eq 'Active' and Roles eq 'EY'").getAll();
          }else if(status == Status.Draft){        
            preparer = currentItem.Preparer;
            PreparerBy = currentItem.Preparer;
            User =  "Auditor";
            emailTo = getTemplate[0].EmailTo;                  
            emailCC = getTemplate[0].EmailCC + ";" + preparerLM + ";" + currentItem.ReviewerEmail;   
            //subject = "Stage 3.1 - " + String(EntityNameCode) + " is ready for FS preparation";
          }else if(status == Status.Underreview){        
            preparer = "Auditor";
            User = currentItem.Preparer;
            PreparerBy = "Auditor";
            emailTo = currentItem.PreparerEmail;                  
            emailCC = getTemplate[0].EmailCC + ";" + preparerLM + ";" + currentItem.ReviewerEmail;   
            subject = "Stage 3.1 - " + String(EntityNameCode) + " status changed to under review";
          }else if(status == Status.ReqIteration){        
            preparer = currentItem.Preparer;
            User = "Consultant";
            PreparerBy = currentItem.Preparer;
            emailTo = getTemplate[0].EmailTo;                  
            emailCC = getTemplate[0].PreparerEmail + ";" + preparerLM + ";"  + getTemplate[0].EmailCC+ ";"+ currentItem.ReviewerEmail;   
            subject = "Stage 3.1 - " + String(EntityNameCode) + " status changed to required iteration";
            nextstageduedate = this.datepipe.transform(currentItem.Stage31ConsultantVariableDeadlin, 'dd-MM-yyyy');    
          }else if(status == Status.Approved){        
            preparer = "Auditor";
            User = currentItem.Preparer;
            PreparerBy = "Auditor";
            emailTo = currentItem.PreparerEmail;
            emailCC = getTemplate[0].EmailCC + ";" + preparerLM + ";" + currentItem.ReviewerEmail;
            nextstageduedate = this.datepipe.transform(currentItem.Stage32AuditorvariableDeadline, 'dd-MM-yyyy');    
          }else if(status == Status.Rejected){        
            preparer = "Auditor";
            User = currentItem.Preparer;
            PreparerBy = "Auditor";
            emailTo = currentItem.PreparerEmail;
            emailCC = preparerLM + ";" + getTemplate[0].EmailCC + ";" + currentItem.ReviewerEmail;
            nextstageduedate = this.datepipe.transform(currentItem.Stage31GSVariableDeadline, 'dd-MM-yyyy');    
          }
        }
        if(stage2or3level == kpmg.Director){
          stageduedate = this.datepipe.transform(currentItem.Stage32Deadline, 'dd-MM-yyyy');
          submissiondate = this.datepipe.transform(currentItem.Stage31ActualSubmission, 'dd-MM-yyyy');          
          if(currentItem.Stg32Comments != null){
            Comment = currentItem.Stg32Comments; 
          }
          nextstage = kpmg.Director;
          nextstageduedate = this.datepipe.transform(currentItem.Stage32Deadline, 'dd-MM-yyyy');
          subject = "Stage 3.2 - " + String(EntityNameCode) + " has been "+ status +"";
          if(status == Status.Submitted && this.selectedItem.ActionHolder=="Consultant"){
            preparer = "Consultant";
            User = currentItem.Preparer;
            PreparerBy = "Consultant";
            emailTo = currentItem.PreparerEmail; 
            emailCC = currentItem.PreparerEmail + ";" +preparerLM  + ";" + currentItem.ReviewerEmail; 
            nextstageduedate = this.datepipe.transform(currentItem.Stage32GSVariableDeadline, 'dd-MM-yyyy');
            //nextstage = kpmg.Manager;
            //nextstageduedate = this.datepipe.transform(currentItem.Stage31Deadline, 'dd-MM-yyyy');
            //users = await this.web.lists.getByTitle("UserRoles").items.select("User", "Roles", "Status").filter("Status eq 'Active' and Roles eq 'EY'").getAll();
          }else if(status == Status.Submitted && (this.selectedItem.ActionHolder=="GS"|| this.selectedItem.ActionHolder=="Auditor")){
            preparer = currentItem.Preparer;
            User = "Auditor";
            PreparerBy = currentItem.Preparer;
            emailTo = getTemplate[0].EmailTo;
            emailCC = currentItem.PreparerEmail + ";" +preparerLM  + ";" + currentItem.ReviewerEmail; 
            nextstageduedate = this.datepipe.transform(currentItem.Stage32AuditorvariableDeadline, 'dd-MM-yyyy');
            //nextstage = kpmg.Manager;
            //nextstageduedate = this.datepipe.transform(currentItem.Stage31Deadline, 'dd-MM-yyyy');
            //users = await this.web.lists.getByTitle("UserRoles").items.select("User", "Roles", "Status").filter("Status eq 'Active' and Roles eq 'EY'").getAll();
          }else if(status == Status.Draft){        
            preparer = currentItem.Preparer;
            User = "Auditor";
            PreparerBy = currentItem.Preparer;
            emailTo = getTemplate[0].EmailTo;                    
            emailCC = getTemplate[0].EmailCC + ";" + preparerLM  + ";" + currentItem.ReviewerEmail;    
            //subject = "Stage 3.1 - " + String(EntityNameCode) + " is ready for FS preparation";
          }else if(status == Status.Underreview){        
            preparer = "Auditor";
            User = currentItem.Preparer;
            PreparerBy = "Auditor";
            emailTo = currentItem.PreparerEmail;                  
            emailCC = getTemplate[0].EmailCC + ";" + preparerLM  + ";" + currentItem.ReviewerEmail;    
            subject = "Stage 3.2 - " + String(EntityNameCode) + " status changed to under review";
          }else if(status == Status.ReqIteration){        
            preparer = currentItem.Preparer;
            User = "Consultant";
            PreparerBy = currentItem.Preparer;
            emailTo = getTemplate[0].EmailTo;                  
            emailCC = getTemplate[0].PreparerEmail + ";" + getTemplate[0].EmailCC + ";" + preparerLM  + ";" + currentItem.ReviewerEmail;    
            subject = "Stage 3.2 - " + String(EntityNameCode) + " status changed to requires iteration";
            nextstageduedate = this.datepipe.transform(currentItem.Stage32ConsultantVariableDeadlin, 'dd-MM-yyyy');
          }else if(status == Status.Approved){        
            preparer = "Auditor";
            User = currentItem.Preparer;
            PreparerBy = "Auditor";
            emailTo = currentItem.PreparerEmail;
            emailCC = getTemplate[0].EmailCC + ";" + preparerLM  + ";" + currentItem.ReviewerEmail;  
            nextstageduedate = this.datepipe.transform(currentItem.Stage33AuditorvariableDeadline, 'dd-MM-yyyy');
          }else if(status == Status.Rejected){        
            preparer = "Auditor";
            User = currentItem.Preparer;
            PreparerBy = "Auditor";
            emailTo = currentItem.PreparerEmail;
            emailCC = preparerLM + ";" + getTemplate[0].EmailCC  + ";" + currentItem.ReviewerEmail;
            nextstageduedate = this.datepipe.transform(currentItem.Stage32GSVariableDeadline, 'dd-MM-yyyy');
          }
        }
        if(stage2or3level == kpmg.Final){
          stageduedate = this.datepipe.transform(currentItem.Stage33Deadline, 'dd-MM-yyyy');
          submissiondate = this.datepipe.transform(currentItem.Stage32ActualSubmission, 'dd-MM-yyyy');          
          if(currentItem.Stg33Comments != null){
            Comment = currentItem.Stg33Comments; 
          }
          nextstage = stages.Stage4;
          nextstageduedate = this.datepipe.transform(currentItem.Stage33Deadline, 'dd-MM-yyyy');
          subject = "Stage 3.3 - " + String(EntityNameCode) + " has been "+ status +"";
          if(status == Status.Submitted && this.selectedItem.ActionHolder=="Consultant"){
            preparer = "Consultant";
            User = currentItem.Preparer;
            PreparerBy = "Consultant";
            emailTo = currentItem.PreparerEmail;
            emailCC = currentItem.PreparerEmail + ";" +preparerLM  + ";" + currentItem.ReviewerEmail; 
            nextstageduedate = this.datepipe.transform(currentItem.Stage33GSVariableDeadline, 'dd-MM-yyyy');
            //nextstage = kpmg.Manager;
            //nextstageduedate = this.datepipe.transform(currentItem.Stage31Deadline, 'dd-MM-yyyy');
            //users = await this.web.lists.getByTitle("UserRoles").items.select("User", "Roles", "Status").filter("Status eq 'Active' and Roles eq 'EY'").getAll();
          }else if(status == Status.Submitted && (this.selectedItem.ActionHolder=="GS" || this.selectedItem.ActionHolder=="Auditor")){
            preparer = currentItem.Preparer;
            User = "Auditor";
            PreparerBy = currentItem.Preparer;
            emailTo = getTemplate[0].EmailTo;
            emailCC = currentItem.PreparerEmail + ";" +preparerLM  + ";" + currentItem.ReviewerEmail; 
            nextstageduedate = this.datepipe.transform(currentItem.Stage33AuditorvariableDeadline, 'dd-MM-yyyy');
            //nextstage = kpmg.Manager;
            //nextstageduedate = this.datepipe.transform(currentItem.Stage31Deadline, 'dd-MM-yyyy');
            //users = await this.web.lists.getByTitle("UserRoles").items.select("User", "Roles", "Status").filter("Status eq 'Active' and Roles eq 'EY'").getAll();
          }else if(status == Status.Draft){        
            preparer = currentItem.Preparer;
            User = "Auditor";
            PreparerBy = currentItem.Preparer;
            emailTo = getTemplate[0].EmailTo;                 
            emailCC = getTemplate[0].EmailCC + ";" + preparerLM  + ";" + currentItem.ReviewerEmail;    
            //subject = "Stage 3.1 - " + String(EntityNameCode) + " is ready for FS preparation";
          }else if(status == Status.Underreview){        
            preparer = "Auditor";
            User = currentItem.Preparer;
            PreparerBy = "Auditor";
            emailTo = currentItem.PreparerEmail;                  
            emailCC = getTemplate[0].EmailCC + ";" + preparerLM  + ";" + currentItem.ReviewerEmail;    
            subject = "Stage 3.3 - " + String(EntityNameCode) + " status changed to under review";
            
          }else if(status == Status.ReqIteration){        
            preparer = currentItem.Preparer;
            User = "Consultant";
            PreparerBy = currentItem.Preparer;
            emailTo = getTemplate[0].EmailTo;                   
            emailCC = getTemplate[0].PreparerEmail + ";" + getTemplate[0].EmailCC + ";" + preparerLM  + ";" + currentItem.ReviewerEmail;    
            subject = "Stage 3.3 - " + String(EntityNameCode) + " status changed to requires iteration";
            nextstageduedate = this.datepipe.transform(currentItem.Stage33ConsultantVariableDeadlin, 'dd-MM-yyyy');
          }else if(status == Status.Approved){        
            preparer = "Auditor";
            User = currentItem.Preparer;
            PreparerBy = "Auditor";
            emailTo = currentItem.PreparerEmail;
            emailCC = getTemplate[0].EmailCC + ";" + preparerLM  + ";" + currentItem.ReviewerEmail;
            nextstageduedate = this.datepipe.transform(currentItem.Stage41GSVariableDeadline, 'dd-MM-yyyy');
          }else if(status == Status.Rejected){        
            preparer = "Auditor";
            User = currentItem.Preparer;
            PreparerBy = "Auditor";
            emailTo = currentItem.PreparerEmail;
            emailCC = preparerLM + ";" + getTemplate[0].EmailCC  + ";" + currentItem.ReviewerEmail;
            nextstageduedate = this.datepipe.transform(currentItem.Stage33GSVariableDeadline, 'dd-MM-yyyy');
          }
        }    
        //preparer = "Auditor";
      }
      if (stage == stages.Stage4) {  
        
        if(stage2or3level == Signoff.GSFS){
          
          stageduedate = this.datepipe.transform(currentItem.Stage4SubmissionDate, 'dd-MM-yyyy');
          submissiondate = this.datepipe.transform(currentItem.Stage33ActualSubmission, 'dd-MM-yyyy');          
          if(currentItem.Stg4Comments != null){
            Comment = currentItem.Stg4Comments; 
          }
          nextstage = Signoff.GSFS;
          nextstageduedate = this.datepipe.transform(currentItem.Stage4SubmissionDate, 'dd-MM-yyyy');    
          subject = "Stage 4.1 - " + String(EntityNameCode) + " has been "+ status +"";
          if(status == Status.Submitted){
            preparer = currentItem.Preparer;
            PreparerBy = currentItem.Preparer;
            User = "CRT";
            emailTo = getTemplate[0].EmailTo;
            emailCC = currentItem.PreparerEmail + ";" +preparerLM + ";" + getTemplate[0].EmailCC + ";" + currentItem.ReviewerEmail; 
            //nextstage = kpmg.Manager;
            nextstageduedate = this.datepipe.transform(currentItem.Stage41CRTVariableDeadline, 'dd-MM-yyyy');
            //users = await this.web.lists.getByTitle("UserRoles").items.select("User", "Roles", "Status").filter("Status eq 'Active' and Roles eq 'EY'").getAll();
          }else if(status == Status.Draft){        
            preparer =  "GS";
            User = currentItem.Preparer;
            PreparerBy = "GS";
            emailTo = currentItem.PreparerEmail;                  
            emailCC = getTemplate[0].EmailCC + ";" + preparerLM + ";" + currentItem.ReviewerEmail;   
            //subject = "Stage 3.1 - " + String(EntityNameCode) + " is ready for FS preparation";
          }else if(status == Status.Underreview){        
            preparer = "CRT";
            User = currentItem.Preparer;
            PreparerBy = "CRT";
            emailTo = currentItem.PreparerEmail;                  
            emailCC = getTemplate[0].EmailCC + ";" + preparerLM + ";" + currentItem.ReviewerEmail;   
            subject = "Stage 4.1 - " + String(EntityNameCode) + " status changed to under review";
          }else if(status == Status.Approved){        
            preparer = "CRT";
            User = currentItem.Preparer;
            PreparerBy = "CRT";
            emailTo = currentItem.PreparerEmail;
            emailCC = getTemplate[0].EmailCC + ";" + preparerLM + ";" + currentItem.ReviewerEmail;
            nextstageduedate = this.datepipe.transform(currentItem.Stage42GSVariableDeadline, 'dd-MM-yyyy');    
          }else if(status == Status.Rejected){        
            preparer = "CRT";
            User = currentItem.Preparer;
            PreparerBy = "CRT";
            emailTo = currentItem.PreparerEmail;
            emailCC = preparerLM + ";" + getTemplate[0].EmailCC + ";" + currentItem.ReviewerEmail;
            nextstageduedate = this.datepipe.transform(currentItem.Stage41GSVariableDeadline, 'dd-MM-yyyy');    
          }
        }

        if(stage2or3level == Signoff.MAFP){
          stageduedate = this.datepipe.transform(currentItem.Stage42SubmissionDate, 'dd-MM-yyyy');
          //stageduedate = this.datepipe.transform(currentItem.Stage4SubmissionDate, 'dd-MM-yyyy');
          submissiondate = this.datepipe.transform(currentItem.Stage42ActualSubmission, 'dd-MM-yyyy');          
          if(currentItem.Stg42Comments != null){
            Comment = currentItem.Stg42Comments; 
          }
          nextstage = Signoff.KPMG;
          nextstageduedate = this.datepipe.transform(currentItem.Stage43SubmissionDate, 'dd-MM-yyyy');
          //nextstageduedate = this.datepipe.transform(currentItem.Stage4SubmissionDate, 'dd-MM-yyyy');
          subject = "Stage 4.2 - " + String(EntityNameCode) + " has been "+ status +"";
          if(status == Status.Submitted){
            preparer = currentItem.Preparer;
            User = "CRT";
            PreparerBy = currentItem.Preparer;
            emailTo = getTemplate[0].EmailTo;
            emailCC = currentItem.PreparerEmail + ";" +preparerLM  + ";" + currentItem.ReviewerEmail; 
            nextstageduedate = this.datepipe.transform(currentItem.Stage42CRTVariableDeadline, 'dd-MM-yyyy');
            //nextstage = kpmg.Manager;
            //nextstageduedate = this.datepipe.transform(currentItem.Stage31Deadline, 'dd-MM-yyyy');
            //users = await this.web.lists.getByTitle("UserRoles").items.select("User", "Roles", "Status").filter("Status eq 'Active' and Roles eq 'EY'").getAll();
          }else if(status == Status.Draft){        
            preparer = "GS";
            User = currentItem.Preparer;
            PreparerBy = "GS";
            emailTo = currentItem.PreparerEmail;                  
            emailCC = getTemplate[0].EmailCC + ";" + preparerLM  + ";" + currentItem.ReviewerEmail;    
            //subject = "Stage 3.1 - " + String(EntityNameCode) + " is ready for FS preparation";
          }else if(status == Status.Underreview){        
            preparer = "CRT";
            User = currentItem.Preparer;
            PreparerBy = "CRT";
            emailTo = currentItem.PreparerEmail;                  
            emailCC = getTemplate[0].EmailCC + ";" + preparerLM  + ";" + currentItem.ReviewerEmail;    
            subject = "Stage 4.2 - " + String(EntityNameCode) + " status changed to under review";
          }else if(status == Status.Approved){        
            preparer = "CRT";
            User = currentItem.Preparer;
            PreparerBy = "CRT";
            emailTo = currentItem.PreparerEmail;
            emailCC = getTemplate[0].EmailCC + ";" + preparerLM  + ";" + currentItem.ReviewerEmail;  
            nextstageduedate = this.datepipe.transform(currentItem.Stage43GSVariableDeadline, 'dd-MM-yyyy');
          }else if(status == Status.Rejected){        
            preparer = "CRT";
            User = currentItem.Preparer;
            PreparerBy = "CRT";
            emailTo = currentItem.PreparerEmail;
            emailCC = preparerLM + ";" + getTemplate[0].EmailCC  + ";" + currentItem.ReviewerEmail;
            nextstageduedate = this.datepipe.transform(currentItem.Stage42GSVariableDeadline, 'dd-MM-yyyy');
          }
        }

        if(stage2or3level == Signoff.KPMG){
          stageduedate = this.datepipe.transform(currentItem.Stage43SubmissionDate, 'dd-MM-yyyy');
          //stageduedate = this.datepipe.transform(currentItem.Stage4SubmissionDate, 'dd-MM-yyyy');
          submissiondate = this.datepipe.transform(currentItem.Stage43ActualSubmission, 'dd-MM-yyyy');          
          if(currentItem.Stg43Comments != null){
            Comment = currentItem.Stg43Comments; 
          }
          nextstage = stages.Stage5;
          nextstageduedate = this.datepipe.transform(currentItem.Stage43SubmissionDate, 'dd-MM-yyyy');
          //nextstageduedate = this.datepipe.transform(currentItem.Stage4SubmissionDate, 'dd-MM-yyyy');
          subject = "Stage 4.3 - " + String(EntityNameCode) + " has been "+ status +"";
          if(status == Status.Submitted){
            preparer = currentItem.Preparer;
            User = "Auditor";
            PreparerBy = currentItem.Preparer;
            emailTo = getTemplate[0].EmailTo;
            emailCC = currentItem.PreparerEmail + ";" +preparerLM  + ";" + currentItem.ReviewerEmail; 
            nextstageduedate = this.datepipe.transform(currentItem.Stage43AuditorVariableDeadline, 'dd-MM-yyyy');
            //nextstage = kpmg.Manager;
            //nextstageduedate = this.datepipe.transform(currentItem.Stage31Deadline, 'dd-MM-yyyy');
            //users = await this.web.lists.getByTitle("UserRoles").items.select("User", "Roles", "Status").filter("Status eq 'Active' and Roles eq 'EY'").getAll();
          }else if(status == Status.Draft){        
            preparer = "GS";
            User = currentItem.Preparer;
            PreparerBy = "GS";
            emailTo = currentItem.PreparerEmail;                  
            emailCC = getTemplate[0].EmailCC + ";" + preparerLM  + ";" + currentItem.ReviewerEmail;    
            //subject = "Stage 3.1 - " + String(EntityNameCode) + " is ready for FS preparation";
          }else if(status == Status.Underreview){        
            preparer = "Auditor";
            User = currentItem.Preparer;
            PreparerBy = "Auditor";
            emailTo = currentItem.PreparerEmail;                  
            emailCC = getTemplate[0].EmailCC + ";" + preparerLM  + ";" + currentItem.ReviewerEmail;    
            subject = "Stage 4.3 - " + String(EntityNameCode) + " status changed to under review";
          }else if(status == Status.Approved){        
            preparer = "Auditor";
            User = currentItem.Preparer;
            PreparerBy = "Auditor";
            emailTo = currentItem.PreparerEmail;
            emailCC = getTemplate[0].EmailCC + ";" + preparerLM  + ";" + currentItem.ReviewerEmail;
            nextstageduedate = this.datepipe.transform(currentItem.Stage5GSVariableDeadline, 'dd-MM-yyyy');
          }else if(status == Status.Rejected){        
            preparer = "Auditor";
            User = currentItem.Preparer;
            PreparerBy = "Auditor";
            emailTo = currentItem.PreparerEmail;
            emailCC = preparerLM + ";" + getTemplate[0].EmailCC  + ";" + currentItem.ReviewerEmail;
            nextstageduedate = this.datepipe.transform(currentItem.Stage43GSVariableDeadline, 'dd-MM-yyyy');
          }
        }
      }
      if (stage == stages.Stage5) {        
        stageduedate = this.datepipe.transform(currentItem.Stage5SubmissionDate, 'dd-MM-yyyy');
        submissiondate = this.datepipe.transform(currentItem.Stage4ActualSubmission, 'dd-MM-yyyy');        
        if(currentItem.Stg5Comments != null){
          Comment = currentItem.Stg5Comments; 
        }
        nextstage = stages.Stage5;
        nextstageduedate = this.datepipe.transform(currentItem.Stage5SubmissionDate, 'dd-MM-yyyy');
        subject = "Stage 5 - " + String(EntityNameCode) + " has been "+ status +"";
        let eCC = "";
        users = await this.web.lists.getByTitle("UserRoles").items.select("User/EMail", "Roles", "Status").expand("User")
        .filter("Status eq 'Active' and Roles eq 'Superuser'").getAll();
        users.forEach((item:any)=>{
          eCC = item.User.EMail + ";" + eCC;
        });
        if(status == Status.Submitted){
          preparer = currentItem.Preparer;
          User = "Corporate Reporting Team";
          PreparerBy = currentItem.Preparer;
          emailTo = eCC;
          emailCC = currentItem.PreparerEmail + ";" + preparerLM + ";" + currentItem.ReviewerEmail; 
          nextstageduedate = this.datepipe.transform(currentItem.Stage5CRTVariableDeadline, 'dd-MM-yyyy');
          //nextstage = kpmg.Manager;
          //nextstageduedate = this.datepipe.transform(currentItem.Stage31Deadline, 'dd-MM-yyyy');
          //users = await this.web.lists.getByTitle("UserRoles").items.select("User", "Roles", "Status").filter("Status eq 'Active' and Roles eq 'EY'").getAll();
        }else if(status == Status.Draft){        
          preparer = "Auditor";
          User = currentItem.Preparer;
          PreparerBy = "Auditor";
          emailTo = currentItem.PreparerEmail;                  
          emailCC = getTemplate[0].EmailCC + ";" + preparerLM  + ";" + currentItem.ReviewerEmail;    
          subject = "Stage 5 - " + String(EntityNameCode) + " is ready for FS Archiving";
          if(currentItem.Stg5Comments != null){
            Comment = currentItem.Stg5Comments; 
          }
        }else if(status == Status.Underreview){        
          preparer = "Corporate Reporting Team";
          User = currentItem.Preparer;
          PreparerBy = "Corporate Reporting Team";
          emailTo = currentItem.PreparerEmail;                  
          emailCC = getTemplate[0].EmailCC + ";" + preparerLM  + ";" + currentItem.ReviewerEmail;
          subject = "Stage 5 - " + String(EntityNameCode) + " status changed to under review";
        }else if(status == Status.Approved){        
          preparer = "Corporate Reporting Team";
          User = currentItem.Preparer;
          PreparerBy = "Corporate Reporting Team";
          emailTo = currentItem.PreparerEmail;
          emailCC = getTemplate[0].EmailCC + ";" + preparerLM + ";" + eCC + ";" + currentItem.ReviewerEmail;
        }else if(status == Status.Rejected){        
          preparer = "Corporate Reporting Team";
          PreparerBy = "Corporate Reporting Team";
          User = currentItem.Preparer;
          emailTo = currentItem.PreparerEmail;
          emailCC = preparerLM + ";" + getTemplate[0].EmailCC + ";" + eCC + ";" + currentItem.ReviewerEmail;
          nextstageduedate = this.datepipe.transform(currentItem.Stage5GSVariableDeadline, 'dd-MM-yyyy');
        }
      }
      //let editURL ="<a href='http://majidalfuttaim.sharepoint.com/sites/fast/SiteAssets/Tracker/Index.aspx/app/edit?itemid="+ currentItem.ID +"&source=home'>link</a>";
      let editURL ="<a href='http://majidalfuttaim.sharepoint.com/sites/fasttest/SiteAssets/Tracker/Index.aspx/app/edit?itemid="+ currentItem.ID +"&source=home'>link</a>";
      emailContent = getTemplate[0].EmailContent;
      
      emailContent = String(emailContent).replace('[Preparer]',preparer).replace('[Comment]',Comment).replace('[PreparerBy]',PreparerBy).replace('[clickhere]',editURL)
        .replace('[stageduedate]',String(stageduedate)).replace('[Entity]',String(EntityNameCode)).replace('[User]', User)
        .replace('[submissiondate]',String(submissiondate)).replace('[nextstage]',nextstage).replace('[nextstageduedate]',String(nextstageduedate));
        // emailCC +=";"+ getTemplate[0].EmailCC;
      
    }
    //if(status!="Required Iteration"){
      await this.web.lists.getByTitle("EmailHistory").items.add({
        'Title': (currentItem.ID).toString(),
        'Comments': Comment,
        //'DueDate': dueDate,
        'EntityCode': currentItem.Title,
        'Entity': currentItem.EntityName,
        'Preparer': currentItem.Preparer,
        'Status': status,
        'EmailType': "StatusChange",
        'Stage': currentItem.CurrentStage,
        'EmailCC': emailCC,
        'EmailTo': emailTo,
        'EmailContent': emailContent,
        'Subject': subject
      });
    //}

  }

  async saveFSHistoryData(status: string, actual:string, asperplan:string, comment:string,stage:string,state:string,variabledeadline:string,gap:string){
    (window as any).global = window;
    
    let actualsubmission = this.datepipe.transform(asperplan, 'dd-MM-yyyy');
    let vardeadline = this.datepipe.transform(variabledeadline, 'dd-MM-yyyy');
    this.web.lists.getByTitle("FSTrackerHistory").items.add({
      'Title':this.itemid,
      'ActualFSSubmission':actualsubmission,
      'SubmissionDate':actual,
      'Status':status,
      'Stage':stage,
      'State':state,
      'Comments':comment,
      'ActionHolder': this.actionHolder,
      'GAP':gap,
      'VariableDeadline':vardeadline
    }).then((item: any) => {
      Swal.fire("Successfully Submitted","","success").then(()=>{
        this.initiateEditItem();
      });
    }).catch((err: any) => { 
      Swal.fire("Failed to update","","info");
    });
  }

  scrolltop(){
    window.scrollTo(0, 0);
  }

  ngAfterViewInit(): void {
    window.scrollTo(0, 0);
  }

  onFileChange(pFileList: any){  
    this.files = Object.keys(pFileList).map(key => pFileList[key]);
    (window as any).global = window;    
    const component: EditComponent = this; 
       let currentYear = new Date();
       this.files.forEach(file => { 
        let filename = this.itemid +"_"+ file.name; 
        component.web.getFolderByServerRelativeUrl(component.docLibPath).files.add(filename, file, true).then(async function(f) {
          let year = String(currentYear.getFullYear());
          (await f.file.getItem()).update({
            Title: file.name,
            Stage: component.selectedItem.CurrentStage,
            Status: component.selectedItem.CurrentStatus,
            ReferenceID: component.itemid,
            Year: year,
            IsActive: "Yes"
          }).then((d)=>{
            component.bindFileData();
          });          
        });
        
        
      });
      
      
      
  }

  bindFileData(){  
    // PROD Site  
    //  this.safeFileUploadSrc =  this.sanitizer.bypassSecurityTrustResourceUrl(environment.sp_URL+"FSDocuments/Forms/AllItems.aspx?id=/sites/fast/FSDocuments/"+this.itemid+"&viewid=2e2a461a-7737-4051-8e2a-0cf3162ac2fc");    
    //  this.safeFileView =  this.sanitizer.bypassSecurityTrustResourceUrl(environment.sp_URL+"FSDocuments/Forms/AllItems.aspx?id=/sites/fast/FSDocuments/"+this.itemid+"&viewid=b3bd969e-d517-4825-b2e6-7273e9903008");

    //UAT Site
    this.safeFileUploadSrc =  this.sanitizer.bypassSecurityTrustResourceUrl(environment.sp_URL+"FSDocuments/Forms/AllItems.aspx?id=/sites/fasttest/FSDocuments/"+this.itemid+"&viewid=1d1922cb-99d8-4188-98e2-51c703480e16");    
    this.safeFileView =  this.sanitizer.bypassSecurityTrustResourceUrl(environment.sp_URL+"FSDocuments/Forms/AllItems.aspx?id=/sites/fasttest/FSDocuments/"+this.itemid+"&viewid=c542065b-f078-4e65-b396-678805882bd9");
    
  }

  downloadFile(f:any, stage:string){
    (window as any).global = window;
    if(this.selectedItem.CurrentStage == stage){
      if(stage == stages.Stage1){
        if(this.selectedItem.Stage1Status !=null && this.selectedItem.Stage1Status == Status.Submitted){
          this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({  
            'Stage1Status': Status.Underreview,
            'CurrentStatus': Status.Underreview
          }).then((item: any) => {
            this.initiateEditItem();
          });
        }
      }else if(stage == stages.Stage2){
        if(this.selectedItem.Stage2Status !=null && this.selectedItem.Stage2Status == Status.Submitted){
          this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({  
            'Stage2Status': Status.Underreview,
            'CurrentStatus': Status.Underreview
          }).then((item: any) => {
            this.initiateEditItem();
          });
        }
      }else if(stage == stages.Stage3){
        if(this.selectedItem.Stage3CurrentState == kpmg.Manager){  
          if(this.selectedItem.Stage31Status !=null && this.selectedItem.Stage31Status == Status.Submitted){
            this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
              'Stage31Status': Status.Underreview,
              'CurrentStatus': Status.Underreview
            }).then((item: any) => {
              this.initiateEditItem();
            });  
          }  
        }else if(this.selectedItem.Stage3CurrentState == kpmg.Director){  
          if(this.selectedItem.Stage32Status !=null && this.selectedItem.Stage32Status == Status.Submitted){ 
            this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
              'Stage32Status': Status.Underreview,
              'CurrentStatus': Status.Underreview
            }).then((item: any) => {
              this.initiateEditItem();
            });
          }
        }else if(this.selectedItem.Stage3CurrentState == kpmg.Final){  
          if(this.selectedItem.Stage33Status !=null && this.selectedItem.Stage33Status == Status.Submitted){ 
            this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({
              'Stage33Status': Status.Underreview,
              'CurrentStatus': Status.Underreview
            }).then((item: any) => {
              this.initiateEditItem();
            });
          }
        }
      }else if(stage == stages.Stage4){
        if(this.selectedItem.Stage4Status !=null && this.selectedItem.Stage4Status == Status.Submitted){ 
          this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({  
            'Stage4Status': Status.Underreview,
            'CurrentStatus': Status.Underreview
          }).then((item: any) => {
            this.initiateEditItem();
          });
        }
      }else if(stage == stages.Stage5){ 
        if(this.selectedItem.Stage5Status !=null && this.selectedItem.Stage5Status == Status.Submitted){
          this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({  
            'Stage5Status': Status.Underreview,
            'CurrentStatus': Status.Underreview
          }).then((item: any) => {
            this.initiateEditItem();
        }); 
        }
      }
    }  
  }

  async deleteFile(f: any) {
    (window as any).global = window;
    Swal.fire({
        title: 'Are you sure want to delete the file?',
        icon: 'warning',
        showDenyButton: true,
        confirmButtonText: 'Yes',
        denyButtonText: 'No',
      }).then(async (result) => {
        if (result.isConfirmed) {
          this.web.lists.getByTitle("FSDocuments").items.getById(f.ID).update({       
            IsActive: "No"     
          }).then((item: any) => {
             Swal.fire("Successfully deleted","","success");
             this.bindFileData();
          }).catch((err: any) => { 
             Swal.fire("Failed to delete","","info");
          });
        }
      });    
  }

  openConfirmDialog(pIndex:any): void {
    const dialogRef = this.dialog.open(DialogConfirmComponent, {
      panelClass: 'modal-xs'
    });
    dialogRef.componentInstance.fName = this.files[pIndex].name;
    dialogRef.componentInstance.fIndex = pIndex;


    dialogRef.afterClosed().subscribe(result => {
      if (result !== undefined) {
        this.deleteFromArray(result);
      }
    });
  }

  deleteFromArray(index:any) {
    
    this.files.splice(index, 1);
  }

  goBack() { 
    this.router.navigate(['/app/home']);
  }

  OpenPopUp(content2:any) {
    this.modalService.open(content2, { size: 'lg', centered: true });
  }

  viewHistory(content: any) {
    (window as any).global = window;
    this.tableHistoryData = [];
    this.service.getFSTrackerHistoryData(String(this.itemid)).subscribe((res) => {
      if (res != null && res !== '') {
        let resultSet = res as any;
        resultSet = resultSet.d.results;
        
        this.tableHistoryData = resultSet;
        
        const table: any = $("#tblHistoryData").DataTable();
        table.destroy();
        setTimeout(() => {
          this.chRef.detectChanges();
          const table2: any = $("#tblHistoryData").DataTable({
            dom: "Blfrtip",
            pageLength: 5,
            paging: true,
            scrollX: false,
            ordering: false,
            columns: [
              { "width": "18%" },
              { "width": "10%" },
              { "width": "14%" },
              { "width": "5%" },
              { "width": "28%" },
              { "width": "10%" },
              { "width": "15%" }
            ]
          });
        }, 100);
        //console.table(filterFields);
        this.modalService.open(content, { size: 'lg', centered: true });
      }
    }), (err: string) => {
      console.log("Error Occured " + err);
    };
    
  }

  viewStgHistory(content: any, stage: any, state:any) {
    debugger;
    (window as any).global = window;
    this.tableHistoryData = [];
    this.service.getFSTrackerStgWiseHistoryData(String(this.itemid), stage, state).subscribe((res) => {
      if (res != null && res !== '') {
        let resultSet = res as any;
        resultSet = resultSet.d.results;
        
        this.tableHistoryData = resultSet.filter((item: { Comments: string | null; }) => (item.Comments != null && item.Comments != ""));
        
        const table: any = $("#tblHistoryStgData").DataTable();
        table.destroy();
        setTimeout(() => {
          this.chRef.detectChanges();
          const table2: any = $("#tblHistoryStgData").DataTable({
            dom: "Blfrtip",
            pageLength: 5,
            paging: true,
            scrollX: false,
            ordering: false,
            columns: [
              { "width": "18%" },
              { "width": "10%" },
              { "width": "14%" },
              { "width": "5%" },
              { "width": "28%" },
              { "width": "10%" },              
              { "width": "15%" }
            ]
          });
        }, 100);
        //console.table(filterFields);
        this.modalService.open(content, { size: 'lg', centered: true });
      }
    }), (err: string) => {
      console.log("Error Occured " + err);
    };
    
  }

  async verifyStage1CheckList(content:any, stage: any) {
    (window as any).global = window;
    this.tableHistoryData = [];
    const controlNames: any[] = [];
    this.stg1ChecklistData = [];
    let controls: { isChecked: boolean; controlName:string; controlID: number; Question: string; orderBy:string;isRequired: boolean; }[] = [];
    this.service.getFSCheckListMasterData(stages.Stage1, this.selectedItem.Country).subscribe((masterdata) => {
      if (masterdata != null && masterdata !== '') {
        let masterSet = masterdata as any;
        masterSet = masterSet.d.results; 
        masterSet.forEach((mitem:any) => {
          controlNames.push(mitem.ControlName);
        });
        controlNames.forEach(controlName => {          
          this.stg1Form.addControl(controlName, new FormControl(''));
        });
        this.service.getFSCheckListItemsData(String(this.itemid)).subscribe((cListData) => {
          if (cListData != null && cListData !== '') {
            let cListSet = cListData as any;
            cListSet = cListSet.d.results;        
            if(cListSet.length > 0){
              const clItem = cListSet[0];    
              if(clItem.Stage1Data != null && clItem.Stage1Data != ""){
                const jsonData = clItem.Stage1Data;
                const dataArray = JSON.parse(jsonData);
                // dataArray.forEach((item: { id: string | number; checked: any; }) => {
                //   const control = this.stg1Form.controls[item.id];
                //   if (control) {
                //     control.setValue(item.checked);
                //   }
                // });
                masterSet.forEach((mitem:any) => {
                  let ischeck = dataArray.filter((fitem: { id: any; }) => fitem.id == (mitem.ControlName));                  
                  if(ischeck.length > 0){
                    controls.push({isChecked:ischeck[0].checked, controlName:mitem.ControlName, controlID:mitem.ID, Question:mitem.Question, orderBy:mitem.OrderBy, isRequired:mitem.IsRequired});
                  }
                });
              }else{
                masterSet.forEach((mitem:any) => {
                  controls.push({isChecked:false, controlName:mitem.ControlName, controlID:mitem.ID, Question:mitem.Question, orderBy:mitem.OrderBy, isRequired:mitem.IsRequired});
                });
              }
            }else{
              masterSet.forEach((mitem:any) => {
                controls.push({isChecked:false, controlName:mitem.ControlName, controlID:mitem.ID, Question:mitem.Question, orderBy:mitem.OrderBy, isRequired:mitem.IsRequired});
              });
            }
            this.stg1ChecklistData = controls;
            this.modalService.open(content, { size: 'lg', centered: true });
          }
        });
      }
    });
  }

  async verifyStage2CheckList(content:any, tab:string, stg2Level:string) {
    (window as any).global = window;    
    let controls: Stg2CheckListItem[] = [];
    let stg21LocalCheckList: Stg2Item[] = [];
    //if(content != ''){
      this.stg2SelectedValues = [];
      //this.stg2IsNew = true;
      this.stg2CommentValues = [];
    //}    
    this.stg2LevelCL = stg2Level;    
    this.service.getFSCheckListStg23MasterData(stages.Stage2, this.stg2LevelCL, tab, this.selectedItem.Country).subscribe((masterdata) => {
      if (masterdata != null && masterdata !== '') {
        let masterSet = masterdata as any;
        masterSet = masterSet.d.results; 
        this.service.getFSCheckListItemsData(String(this.itemid)).subscribe(async (cListData) => {
          if (cListData != null && cListData !== '') {
            
            let cListSet = cListData as any;
            cListSet = cListSet.d.results;        
            if(cListSet.length > 0){
              const clItem = cListSet[0];   
              if(this.stg2LevelCL == EY.EYFS){ 
                if(clItem.Stage21Data != null && clItem.Stage21Data != ""){
                  //const jsonData = clItem.Stage21Data;
                  const dataArray = JSON.parse(clItem.Stage21Data);   
                  if(dataArray != null){
                    dataArray.forEach((mitem:any) => {
                      this.stg2SelectedValues.push({ name: mitem.name, value: mitem.value, id: mitem.id });
                    });
                  }
                  //const jsonCommentData = clItem.Stage21Comments;
                  
                  const commentArray = JSON.parse(clItem.Stage21Comments); 
                  if(commentArray != null){
                    commentArray.forEach((citem:any) => {
                      this.stg2CommentValues.push({ name: citem.name, value: citem.value, id: citem.id });
                    });
                  }
                  // this.stg2CommentValues.push({ name: item.control, value: $event.value, id: Number(item.id) });
                  this.stg2IsNew = false;
                }
              } else if(this.stg2LevelCL == EY.FSTOKPMG){
                if(clItem.Stage22Data != null && clItem.Stage22Data != ""){
                  //const jsonData = clItem.Stage22Data;
                  const dataArray = JSON.parse(clItem.Stage22Data);
                  if(dataArray != null){          
                    dataArray.forEach((mitem:any) => {
                      this.stg2SelectedValues.push({ name: mitem.name, value: mitem.value, id: mitem.id });
                    });
                  }
                  const commentArray = JSON.parse(clItem.Stage22Comments); 
                  if(commentArray != null){  
                    commentArray.forEach((citem:any) => {
                      this.stg2CommentValues.push({ name: citem.name, value: citem.value, id: citem.id });
                    });
                  }
                  this.stg2IsNew = false;
                }
              }
            }
            
            
            if(this.stg2IsNew && content != ''){
              let getMasterData = await this.web.lists.getByTitle("MasterCheckList").items.select("ID","Title","OrderBy","State","ControlName")
                  .filter("Title eq '"+ stages.Stage2 +"' and State eq '" + encodeURIComponent(this.stg2LevelCL) + "'").orderBy("OrderBy").getAll();             
              getMasterData.forEach((mitem:any) => {
                this.stg2SelectedValues.push({ name: mitem.ControlName, value: 'Yes', id: Number(mitem.ID) });
              });                      
            }
            const gArr: string[] = Array.from(new Set(masterSet.map((item: { Group: any; Question: string; }) => item.Group)));
            gArr.forEach(element => {
              if(element != null){
                const gArr2 = masterSet.filter((gdItem: { Group: string; }) => gdItem.Group == element);    
                gArr2.forEach((gele: any) => {                   
                    let filterdata = this.stg2SelectedValues.filter((fitem: { name: string;}) => (fitem.name == gele.ControlName));                    
                    if(filterdata.length > 0){
                      stg21LocalCheckList.push({category:gele.Group, header:gele.Question, qid:gele.QuestionID, control:gele.ControlName,id:Number(gele.ID), default: filterdata[0].value})
                    }else{
                      stg21LocalCheckList.push({category:gele.Group, header:gele.Question, qid:gele.QuestionID, control:gele.ControlName,id:Number(gele.ID), default: 'Yes'})
                    }
                });
                const gdArr = masterSet.filter((gdItem: { GroupBy: string; }) => gdItem.GroupBy == element);
                gdArr.forEach((fele: any) => {
                  let filterdata = this.stg2SelectedValues.filter((fitem: { name: string;}) => (fitem.name == fele.ControlName));                    
                  if(filterdata.length > 0){
                    controls.push({ category:fele.GroupBy, header: element, selection: '', question:fele.Question, id:Number(fele.ID), orderby:String(fele.OrderBy), qid:fele.QuestionID, control:fele.ControlName, default: filterdata[0].value });
                  }else{
                    controls.push({ category:fele.GroupBy, header: element, selection: '', question:fele.Question, id:Number(fele.ID), orderby:String(fele.OrderBy), qid:fele.QuestionID, control:fele.ControlName, default: 'Yes' });
                  }
                });
              }
            });
            this.stg2ChildRowItems = controls;
            this.stg2HeaderRowItems = stg21LocalCheckList;
            if(content != ''){
              this.modalService.open(content, { size: 'lg', centered: true });
            }
          }
        });
        
      }
    });
  }

  async verifyStage3CheckList(content:any, tab:string, stg3Level:string) {
    (window as any).global = window;    
    let controls: Stg2CheckListItem[] = [];
    let stg3LocalCheckList: Stg2Item[] = [];
    //if(content != ''){
      this.stg3SelectedValues = [];
      this.stg3CommentValues = [];
      //this.stg3IsNew = true;
    //}
    this.stg3LevelCL = stg3Level;
   
    this.service.getFSCheckListStg23MasterData(stages.Stage3, this.stg3LevelCL, tab, this.selectedItem.Country).subscribe((masterdata) => {
      if (masterdata != null && masterdata !== '') {
        let masterSet = masterdata as any;
        masterSet = masterSet.d.results; 
        this.service.getFSCheckListItemsData(String(this.itemid)).subscribe(async (cListData) => {
          if (cListData != null && cListData !== '') {
            let cListSet = cListData as any;
            cListSet = cListSet.d.results;        
            if(cListSet.length > 0){
              const clItem = cListSet[0];   
              if(this.stg3LevelCL == kpmg.Manager){ 
                if(clItem.Stage31Data != null && clItem.Stage31Data != ""){
                  //const jsonData = clItem.Stage21Data;
                  const dataArray = JSON.parse(clItem.Stage31Data);   
                  if(dataArray != null){
                    dataArray.forEach((mitem:any) => {
                      this.stg3SelectedValues.push({ name: mitem.name, value: mitem.value, id: mitem.id });
                    });
                  }
                  const commentArray = JSON.parse(clItem.Stage31Comments); 
                  if(commentArray != null){
                    commentArray.forEach((citem:any) => {
                      this.stg3CommentValues.push({ name: citem.name, value: citem.value, id: citem.id });
                    });
                  }
                  // this.stg2CommentValues.push({ name: item.control, value: $event.value, id: Number(item.id) });
                  this.stg3IsNew = false;
                }
              } else if(this.stg3LevelCL == kpmg.Director){
                if(clItem.Stage32Data != null && clItem.Stage32Data != ""){
                  //const jsonData = clItem.Stage22Data;
                  const dataArray = JSON.parse(clItem.Stage32Data);
                  if(dataArray != null){          
                    dataArray.forEach((mitem:any) => {
                      this.stg3SelectedValues.push({ name: mitem.name, value: mitem.value, id: mitem.id });
                    });
                  }
                  const commentArray = JSON.parse(clItem.Stage32Comments); 
                  if(commentArray != null){  
                    commentArray.forEach((citem:any) => {
                      this.stg3CommentValues.push({ name: citem.name, value: citem.value, id: citem.id });
                    });
                  }
                  this.stg3IsNew = false;
                }
              } else if(this.stg3LevelCL == kpmg.Final){
                if(clItem.Stage33Data != null && clItem.Stage33Data != ""){
                  //const jsonData = clItem.Stage22Data;
                  const dataArray = JSON.parse(clItem.Stage33Data);
                  if(dataArray != null){          
                    dataArray.forEach((mitem:any) => {
                      this.stg3SelectedValues.push({ name: mitem.name, value: mitem.value, id: mitem.id });
                    });
                  }
                  const commentArray = JSON.parse(clItem.Stage33Comments); 
                  if(commentArray != null){  
                    commentArray.forEach((citem:any) => {
                      this.stg3CommentValues.push({ name: citem.name, value: citem.value, id: citem.id });
                    });
                  }
                  this.stg3IsNew = false;
                }
              }
            }

            if(this.stg3IsNew && content != ''){
              let getMasterData = await this.web.lists.getByTitle("MasterCheckList").items.select("ID","Title","OrderBy","State","ControlName")
                  .filter("Title eq '"+ stages.Stage3 +"' and State eq '" + encodeURIComponent(this.stg3LevelCL) + "'").orderBy("OrderBy").getAll();             
              getMasterData.forEach((mitem:any) => {
                this.stg3SelectedValues.push({ name: mitem.ControlName, value: 'Yes', id: Number(mitem.ID) });
              });              
            }

            const gArr: string[] = Array.from(new Set(masterSet.map((item: { Group: any; Question: string; }) => item.Group)));
            gArr.forEach(element => {
              if(element != null){
                const gArr2 = masterSet.filter((gdItem: { Group: string; }) => gdItem.Group == element);
                gArr2.forEach((gele: any) => {                   
                    let filterdata = this.stg3SelectedValues.filter((fitem: { name: string;}) => (fitem.name == gele.ControlName));                    
                    if(filterdata.length > 0){
                      stg3LocalCheckList.push({category:gele.Group, header:gele.Question, qid:gele.QuestionID, control:gele.ControlName,id:Number(gele.ID), default: filterdata[0].value})
                    }else{
                      stg3LocalCheckList.push({category:gele.Group, header:gele.Question, qid:gele.QuestionID, control:gele.ControlName,id:Number(gele.ID), default: 'Yes'})
                    }
                });
                const gdArr = masterSet.filter((gdItem: { GroupBy: string; }) => gdItem.GroupBy == element);
                gdArr.forEach((fele: any) => {
                  let filterdata = this.stg3SelectedValues.filter((fitem: { name: string;}) => (fitem.name == fele.ControlName));                    
                  if(filterdata.length > 0){
                    controls.push({ category:fele.GroupBy, header: element, selection: '', question:fele.Question, id:Number(fele.ID), orderby:String(fele.OrderBy), qid:fele.QuestionID, control:fele.ControlName, default: filterdata[0].value });
                  }else{
                    controls.push({ category:fele.GroupBy, header: element, selection: '', question:fele.Question, id:Number(fele.ID), orderby:String(fele.OrderBy), qid:fele.QuestionID, control:fele.ControlName, default: 'Yes' });
                  }
                });
              }

            });
            this.stg3ChildRowItems = controls;
            this.stg3HeaderRowItems = stg3LocalCheckList;
            if(content != ''){
              this.modalService.open(content, { size: 'lg', centered: true });
            }
          }
        });
        
      }
    });
  }

  async verifyStage4CheckList(content:any, tab:string, stg4Level:string) {
    (window as any).global = window;    
    
    let controls: Stg2CheckListItem[] = [];
    let stg4LocalCheckList: Stg2Item[] = [];
    //if(content != ''){
      this.stg4SelectedValues = [];
      this.stg4CommentValues = [];   
     // this.stg4IsNew = true; 
    //}
    this.stg4LevelCL = stg4Level;
   
    this.service.getFSCheckListStg23MasterData(stages.Stage4, this.stg4LevelCL, tab, this.selectedItem.Country).subscribe((masterdata) => {
      if (masterdata != null && masterdata !== '') {
        let masterSet = masterdata as any;
        masterSet = masterSet.d.results; 
        this.service.getFSCheckListItemsData(String(this.itemid)).subscribe(async (cListData) => {
          if (cListData != null && cListData !== '') {
            let cListSet = cListData as any;
            cListSet = cListSet.d.results;        
            if(cListSet.length > 0){
              const clItem = cListSet[0];   
              if(this.stg4LevelCL == Signoff.GSFS){ 
                if(clItem.Stage4Data != null && clItem.Stage4Data != ""){
                  //const jsonData = clItem.Stage21Data;
                  const dataArray = JSON.parse(clItem.Stage4Data);   
                  if(dataArray != null){
                    dataArray.forEach((mitem:any) => {
                      this.stg4SelectedValues.push({ name: mitem.name, value: mitem.value, id: mitem.id });
                    });
                  }
                  const commentArray = JSON.parse(clItem.Stage4Comments); 
                  if(commentArray != null){
                    commentArray.forEach((citem:any) => {
                      this.stg4CommentValues.push({ name: citem.name, value: citem.value, id: citem.id });
                    });
                  }
                  // this.stg2CommentValues.push({ name: item.control, value: $event.value, id: Number(item.id) });
                  this.stg4IsNew = false;
                }
              } else if(this.stg4LevelCL == Signoff.MAFP){
                if(clItem.Stage42Data != null && clItem.Stage42Data != ""){
                  //const jsonData = clItem.Stage22Data;
                  const dataArray = JSON.parse(clItem.Stage42Data);
                  if(dataArray != null){          
                    dataArray.forEach((mitem:any) => {
                      this.stg4SelectedValues.push({ name: mitem.name, value: mitem.value, id: mitem.id });
                    });
                  }
                  const commentArray = JSON.parse(clItem.Stage42Comments); 
                  if(commentArray != null){  
                    commentArray.forEach((citem:any) => {
                      this.stg4CommentValues.push({ name: citem.name, value: citem.value, id: citem.id });
                    });
                  }
                  this.stg4IsNew = false;
                }
              } else if(this.stg4LevelCL == Signoff.KPMG){
                if(clItem.Stage43Data != null && clItem.Stage43Data != ""){
                  //const jsonData = clItem.Stage22Data;
                  const dataArray = JSON.parse(clItem.Stage43Data);
                  if(dataArray != null){          
                    dataArray.forEach((mitem:any) => {
                      this.stg4SelectedValues.push({ name: mitem.name, value: mitem.value, id: mitem.id });
                    });
                  }
                  const commentArray = JSON.parse(clItem.Stage43Comments); 
                  if(commentArray != null){  
                    commentArray.forEach((citem:any) => {
                      this.stg4CommentValues.push({ name: citem.name, value: citem.value, id: citem.id });
                    });
                  }
                  this.stg4IsNew = false;
                }
              }
            }

            if(this.stg4IsNew && content != ''){
              let getMasterData = await this.web.lists.getByTitle("MasterCheckList").items.select("ID","Title","OrderBy","State","ControlName")
                  .filter("Title eq '"+ stages.Stage4 +"' and State eq '" + encodeURIComponent(this.stg4LevelCL) + "'").orderBy("OrderBy").getAll();             
              getMasterData.forEach((mitem:any) => {
                this.stg4SelectedValues.push({ name: mitem.ControlName, value: 'Yes', id: Number(mitem.ID) });
              });              
            }

            const gArr: string[] = Array.from(new Set(masterSet.map((item: { Group: any; Question: string; }) => item.Group)));
            gArr.forEach(element => {
              if(element != null){
                const gArr2 = masterSet.filter((gdItem: { Group: string; }) => gdItem.Group == element);
                gArr2.forEach((gele: any) => {                   
                    let filterdata = this.stg4SelectedValues.filter((fitem: { name: string;}) => (fitem.name == gele.ControlName));                    
                    if(filterdata.length > 0){
                      stg4LocalCheckList.push({category:gele.Group, header:gele.Question, qid:gele.QuestionID, control:gele.ControlName,id:Number(gele.ID), default: filterdata[0].value})
                    }else{
                      stg4LocalCheckList.push({category:gele.Group, header:gele.Question, qid:gele.QuestionID, control:gele.ControlName,id:Number(gele.ID), default: 'Yes'})
                    }
                    

                });
                const gdArr = masterSet.filter((gdItem: { GroupBy: string; }) => gdItem.GroupBy == element);
                gdArr.forEach((fele: any) => {
                  let filterdata = this.stg4SelectedValues.filter((fitem: { name: string;}) => (fitem.name == fele.ControlName));                    
                  if(filterdata.length > 0){
                    controls.push({ category:fele.GroupBy, header: element, selection: '', question:fele.Question, id:Number(fele.ID), orderby:String(fele.OrderBy), qid:fele.QuestionID, control:fele.ControlName, default: filterdata[0].value });
                  }else{
                    controls.push({ category:fele.GroupBy, header: element, selection: '', question:fele.Question, id:Number(fele.ID), orderby:String(fele.OrderBy), qid:fele.QuestionID, control:fele.ControlName, default: 'Yes' });
                  }
                });
              }

            });
            this.stg4ChildRowItems = controls;
            this.stg4HeaderRowItems = stg4LocalCheckList;
            if(content != ''){
              this.modalService.open(content, { size: 'lg', centered: true });
            }
          }
        });
        
      }
    });
  }

  async verifyStage5CheckList(content:any, tab:string) {
    (window as any).global = window;    
    let controls5: Stg2CheckListItem[] = [];
    let stg5LocalCheckList: Stg2Item[] = [];
    //if(content != ''){
      this.stg5SelectedValues = [];
      this.stg5CommentValues = [];   
     // this.stg5IsNew = true; 
   // }
    this.service.getFSCheckListStg45MasterData(stages.Stage5, tab, this.selectedItem.Country).subscribe((masterdata) => {
      if (masterdata != null && masterdata !== '') {
        let masterSet = masterdata as any;
        masterSet = masterSet.d.results; 
        this.service.getFSCheckListItemsData(String(this.itemid)).subscribe(async (cListData) => {
          if (cListData != null && cListData !== '') {
            let cListSet = cListData as any;
            cListSet = cListSet.d.results;        
            if(cListSet.length > 0){
              const clItem = cListSet[0];
                if(clItem.Stage5Data != null && clItem.Stage5Data != ""){
                  //const jsonData = clItem.Stage21Data;
                  const dataArray = JSON.parse(clItem.Stage5Data);   
                  if(dataArray != null){
                    dataArray.forEach((mitem:any) => {
                      this.stg5SelectedValues.push({ name: mitem.name, value: mitem.value, id: mitem.id });
                    });
                  }
                  const commentArray = JSON.parse(clItem.Stage5Comments); 
                  if(commentArray != null){
                    commentArray.forEach((citem:any) => {
                      this.stg5CommentValues.push({ name: citem.name, value: citem.value, id: citem.id });
                    });
                  }
                  // this.stg2CommentValues.push({ name: item.control, value: $event.value, id: Number(item.id) });
                  this.stg5IsNew = false;
                }
            }

            if(this.stg5IsNew && content != ''){              
              let getMasterData = await this.web.lists.getByTitle("MasterCheckList").items.select("ID","Title","OrderBy","State","ControlName")
                  .filter("Title eq '"+ stages.Stage5 +"'").orderBy("OrderBy").getAll();             
              getMasterData.forEach((mitem:any) => {
                this.stg5SelectedValues.push({ name: mitem.ControlName, value: 'Yes', id: Number(mitem.ID) });
              });

            }
            const gArr: string[] = Array.from(new Set(masterSet.map((item: { Group: any; Question: string; }) => item.Group)));
            gArr.forEach(element => {
              if(element != null){
                const gArr2 = masterSet.filter((gdItem: { Group: string; }) => gdItem.Group == element);
                gArr2.forEach((gele: any) => {                   
                    let filterdata = this.stg5SelectedValues.filter((fitem: { name: string;}) => (fitem.name == gele.ControlName));                    
                    if(filterdata.length > 0){
                      stg5LocalCheckList.push({category:gele.Group, header:gele.Question, qid:gele.QuestionID, control:gele.ControlName,id:Number(gele.ID), default: filterdata[0].value})
                    }else{
                      stg5LocalCheckList.push({category:gele.Group, header:gele.Question, qid:gele.QuestionID, control:gele.ControlName,id:Number(gele.ID), default: 'Yes'})
                    }
                });
                const gdArr = masterSet.filter((gdItem: { GroupBy: string; }) => gdItem.GroupBy == element);
                gdArr.forEach((fele: any) => {
                  let filterdata = this.stg5SelectedValues.filter((fitem: { name: string;}) => (fitem.name == fele.ControlName));                    
                  if(filterdata.length > 0){
                    controls5.push({ category:fele.GroupBy, header: element, selection: '', question:fele.Question, id:Number(fele.ID), orderby:String(fele.OrderBy), qid:fele.QuestionID, control:fele.ControlName, default: filterdata[0].value });
                  }else{
                    controls5.push({ category:fele.GroupBy, header: element, selection: '', question:fele.Question, id:Number(fele.ID), orderby:String(fele.OrderBy), qid:fele.QuestionID, control:fele.ControlName, default: 'Yes' });
                  }
                });
              }

            });            
            this.stg5ChildRowItems = controls5;
            this.stg5HeaderRowItems = stg5LocalCheckList;
            if(content != ''){
              this.modalService.open(content, { size: 'lg', centered: true });
            }
          }
        });        
      }
    });
  }

  getItemsByCategory(category: string): Stg2CheckListItem[] {
    return this.stg2ChildRowItems.filter(item => item.category === category);
  }
  
  getStg3ItemsByCategory(category: string): Stg2CheckListItem[] {
    return this.stg3ChildRowItems.filter(item => item.category === category);
  }

  getStg4ItemsByCategory(category: string): Stg2CheckListItem[] {
    return this.stg4ChildRowItems.filter(item => item.category === category);
  }

  getStg5ItemsByCategory(category: string): Stg2CheckListItem[] {
    return this.stg5ChildRowItems.filter(item => item.category === category);
  }

  stage1ConvertToJson() {
    (window as any).global = window;
    let stage = 'Stage1';
    const data = Object.keys(this.stg1Form.controls).map(key => {
      return {
        id: key,
        checked: this.stg1Form.controls[key].value || false
      };
    });
    const jsonData = JSON.stringify(data);
    this.service.getFSCheckListItemsData(String(this.itemid)).subscribe((res) => {
      if (res != null && res !== '') {
        let resultSet = res as any;
        resultSet = resultSet.d.results;        
        if(resultSet.length > 0){
          if(stage == 'Stage1'){
            this.web.lists.getByTitle('FSTrackerCheckList').items.getById(Number(resultSet[0].ID)).update({
              'Stage1Data':jsonData
            }).then(() => {
              Swal.fire("Successfully Confirmed","","success").then(()=>{
                this.modalService.dismissAll();
              });
              
            });
          }
        }else{
          if(stage == 'Stage1'){
            this.web.lists.getByTitle('FSTrackerCheckList').items.add({
              'Stage1Data':jsonData,
              'Title':String(this.itemid)
            }).then(() => {
              Swal.fire("Successfully Confirmed","","success").then(()=>{
                this.modalService.dismissAll();
              });
            });
          }
        }
      }
    }), (err: string) => {
      console.log("Error Occured " + err);
    };
  }

  onChangeStg2TabSelect($event:any){
    if(this.stg2LevelCL == EY.EYFS){
      if($event.index == S21TabIndex.Tab1){
        this.verifyStage2CheckList('', S21Tabs.Tab1, EY.EYFS);
      }
      if($event.index == S21TabIndex.Tab2){
        this.verifyStage2CheckList('', S21Tabs.Tab2, EY.EYFS);
      }
      if($event.index == S21TabIndex.Tab3){
        this.verifyStage2CheckList('', S21Tabs.Tab3, EY.EYFS);
      }
      if($event.index == S21TabIndex.Tab4){
        this.verifyStage2CheckList('', S21Tabs.Tab4, EY.EYFS);
      }
    }else if(this.stg2LevelCL == EY.FSTOKPMG){
      if($event.index == S21TabIndex.Tab1){
        this.verifyStage2CheckList('', S21Tabs.Tab1, EY.FSTOKPMG);
      }
      if($event.index == S21TabIndex.Tab2){
        this.verifyStage2CheckList('', S21Tabs.Tab2, EY.FSTOKPMG);
      }
      if($event.index == S21TabIndex.Tab3){
        this.verifyStage2CheckList('', S21Tabs.Tab3, EY.FSTOKPMG);
      }
      if($event.index == S21TabIndex.Tab4){
        this.verifyStage2CheckList('', S21Tabs.Tab4, EY.FSTOKPMG);
      }
    }
  }

  stage2ConvertToJson() {
    (window as any).global = window;  
    
    
    if(this.stg2SelectedValues.length > 0){

      const noItems = this.stg2SelectedValues.filter(item => (item.value === "No" || item.value === "NA"));

      // Check if any 'No' item in Array 1 does not exist in Array 2 or has an empty value
      const invalidItems = noItems.filter(item => {
        const matchingItem = this.stg2CommentValues.find(
          a2Item => a2Item.name === item.name && a2Item.value !== ""
        );
        return !matchingItem;
      });

      if (invalidItems.length > 0) {
        
        Swal.fire("Please enter your comment","","info");
        
      } else {
        const jsonData = JSON.stringify(this.stg2SelectedValues);
        const jsonCommentsData = JSON.stringify(this.stg2CommentValues);
        this.service.getFSCheckListItemsData(String(this.itemid)).subscribe((res) => {
          if (res != null && res !== '') {
            let resultSet = res as any;
            resultSet = resultSet.d.results;        
            if(resultSet.length > 0){
              if(this.stg2LevelCL == EY.EYFS){
                this.web.lists.getByTitle('FSTrackerCheckList').items.getById(Number(resultSet[0].ID)).update({
                  'Stage21Data':jsonData,
                  'Stage21Comments':jsonCommentsData
                }).then(() => {
                  Swal.fire("Successfully Confirmed","","success").then(()=>{
                    this.modalService.dismissAll();
                  });
                });
              }else if(this.stg2LevelCL == EY.FSTOKPMG){
                this.web.lists.getByTitle('FSTrackerCheckList').items.getById(Number(resultSet[0].ID)).update({
                  'Stage22Data':jsonData,
                  'Stage22Comments':jsonCommentsData
                }).then(() => {
                  Swal.fire("Successfully Confirmed","","success").then(()=>{
                    this.modalService.dismissAll();
                  });
                });
              }
            }else{
              if(this.stg2LevelCL == EY.EYFS){
                this.web.lists.getByTitle('FSTrackerCheckList').items.add({
                  'Stage21Data':jsonData,
                  'Title':String(this.itemid)
                }).then(() => {
                  Swal.fire("Successfully Confirmed","","success").then(()=>{
                    this.modalService.dismissAll();
                  });
                });
              }else if(this.stg2LevelCL == EY.FSTOKPMG){
                this.web.lists.getByTitle('FSTrackerCheckList').items.add({
                  'Stage22Data':jsonData,
                  'Title':String(this.itemid)
                }).then(() => {
                  Swal.fire("Successfully Confirmed","","success").then(()=>{
                    this.modalService.dismissAll();
                  });
                });
              }
            }
          }
        }), (err: string) => {
          console.log("Error Occured " + err);
        };
      }

    }
  }

  stage3ConvertToJson() {
    (window as any).global = window;    
    if(this.stg3SelectedValues.length > 0){
      const noItems = this.stg3SelectedValues.filter(item => (item.value === "No" || item.value === "NA"));
      // Check if any 'No' item in Array 1 does not exist in Array 2 or has an empty value
      const invalidItems = noItems.filter(item => {
        const matchingItem = this.stg3CommentValues.find(
          a2Item => a2Item.name === item.name && a2Item.value !== ""
        );
        return !matchingItem;
      });

      if (invalidItems.length > 0) {
        
        Swal.fire("Please enter your comment","","info");
        
      } else {
        const jsonData = JSON.stringify(this.stg3SelectedValues);
        const jsonCommentsData = JSON.stringify(this.stg3CommentValues);
        this.service.getFSCheckListItemsData(String(this.itemid)).subscribe((res) => {
          if (res != null && res !== '') {
            let resultSet = res as any;
            resultSet = resultSet.d.results;        
            if(resultSet.length > 0){
              if(this.stg3LevelCL == kpmg.Manager){
                this.web.lists.getByTitle('FSTrackerCheckList').items.getById(Number(resultSet[0].ID)).update({
                  'Stage31Data':jsonData,
                  'Stage31Comments':jsonCommentsData
                }).then(() => {
                  Swal.fire("Successfully Confirmed","","success").then(()=>{
                    this.modalService.dismissAll();
                  });
                });
              }else if(this.stg3LevelCL == kpmg.Director){
                this.web.lists.getByTitle('FSTrackerCheckList').items.getById(Number(resultSet[0].ID)).update({
                  'Stage32Data':jsonData,
                  'Stage32Comments':jsonCommentsData
                }).then(() => {
                  Swal.fire("Successfully Confirmed","","success").then(()=>{
                    this.modalService.dismissAll();
                  });
                });
              }else if(this.stg3LevelCL == kpmg.Final){
                this.web.lists.getByTitle('FSTrackerCheckList').items.getById(Number(resultSet[0].ID)).update({
                  'Stage33Data':jsonData,
                  'Stage33Comments':jsonCommentsData
                }).then(() => {
                  Swal.fire("Successfully Confirmed","","success").then(()=>{
                    this.modalService.dismissAll();
                  });
                });
              }
            }else{
              if(this.stg3LevelCL == kpmg.Manager){
                this.web.lists.getByTitle('FSTrackerCheckList').items.add({
                  'Stage31Data':jsonData,
                  'Title':String(this.itemid)
                }).then(() => {
                  Swal.fire("Successfully Confirmed","","success").then(()=>{
                    this.modalService.dismissAll();
                  });
                });
              }else if(this.stg3LevelCL == kpmg.Director){
                this.web.lists.getByTitle('FSTrackerCheckList').items.add({
                  'Stage32Data':jsonData,
                  'Title':String(this.itemid)
                }).then(() => {
                  Swal.fire("Successfully Confirmed","","success").then(()=>{
                    this.modalService.dismissAll();
                  });
                });
              }else if(this.stg3LevelCL == kpmg.Final){
                this.web.lists.getByTitle('FSTrackerCheckList').items.add({
                  'Stage33Data':jsonData,
                  'Title':String(this.itemid)
                }).then(() => {
                  Swal.fire("Successfully Confirmed","","success").then(()=>{
                    this.modalService.dismissAll();
                  });
                });
              }
            }
          }
        }), (err: string) => {
          console.log("Error Occured " + err);
        };
      }
    }
  }

  stage4ConvertToJson() {
    (window as any).global = window;
    if(this.stg4SelectedValues.length > 0){
      const noItems = this.stg4SelectedValues.filter(item => (item.value === "No" || item.value === "NA"));
      // Check if any 'No' item in Array 1 does not exist in Array 2 or has an empty value
      const invalidItems = noItems.filter(item => {
        const matchingItem = this.stg4CommentValues.find(
          a2Item => a2Item.name === item.name && a2Item.value !== ""
        );
        return !matchingItem;
      });

      if (invalidItems.length > 0) {
        
        Swal.fire("Please enter your comment","","info");
        
      } else {
        const jsonData = JSON.stringify(this.stg4SelectedValues);
        const jsonCommentsData = JSON.stringify(this.stg4CommentValues);
        this.service.getFSCheckListItemsData(String(this.itemid)).subscribe((res) => {
          if (res != null && res !== '') {
            let resultSet = res as any;
            resultSet = resultSet.d.results;        
            if(resultSet.length > 0){
              if(this.stg4LevelCL == Signoff.GSFS){
                this.web.lists.getByTitle('FSTrackerCheckList').items.getById(Number(resultSet[0].ID)).update({
                  'Stage4Data':jsonData,
                  'Stage4Comments':jsonCommentsData
                }).then(() => {
                  Swal.fire("Successfully Confirmed","","success").then(()=>{
                    this.modalService.dismissAll();
                  });
                });
              }else if(this.stg4LevelCL == Signoff.MAFP){
                this.web.lists.getByTitle('FSTrackerCheckList').items.getById(Number(resultSet[0].ID)).update({
                  'Stage42Data':jsonData,
                  'Stage42Comments':jsonCommentsData
                }).then(() => {
                  Swal.fire("Successfully Confirmed","","success").then(()=>{
                    this.modalService.dismissAll();
                  });
                });
              }else if(this.stg4LevelCL == Signoff.KPMG){
                this.web.lists.getByTitle('FSTrackerCheckList').items.getById(Number(resultSet[0].ID)).update({
                  'Stage43Data':jsonData,
                  'Stage43Comments':jsonCommentsData
                }).then(() => {
                  Swal.fire("Successfully Confirmed","","success").then(()=>{
                    this.modalService.dismissAll();
                  });
                });
              }
            }else{
              if(this.stg4LevelCL == Signoff.GSFS){
                this.web.lists.getByTitle('FSTrackerCheckList').items.add({
                  'Stage4Data':jsonData,
                  'Title':String(this.itemid)
                }).then(() => {
                  Swal.fire("Successfully Confirmed","","success").then(()=>{
                    this.modalService.dismissAll();
                  });
                });
              }else if(this.stg4LevelCL == Signoff.MAFP){
                this.web.lists.getByTitle('FSTrackerCheckList').items.add({
                  'Stage42Data':jsonData,
                  'Title':String(this.itemid)
                }).then(() => {
                  Swal.fire("Successfully Confirmed","","success").then(()=>{
                    this.modalService.dismissAll();
                  });
                });
              }else if(this.stg4LevelCL == Signoff.KPMG){
                this.web.lists.getByTitle('FSTrackerCheckList').items.add({
                  'Stage43Data':jsonData,
                  'Title':String(this.itemid)
                }).then(() => {
                  Swal.fire("Successfully Confirmed","","success").then(()=>{
                    this.modalService.dismissAll();
                  });
                });
              }
            }
          }
        }), (err: string) => {
          console.log("Error Occured " + err);
        };
      }
    }
  }

  stage5ConvertToJson() {
    (window as any).global = window;    
    
    if(this.stg5SelectedValues.length > 0){
      const noItems = this.stg5SelectedValues.filter(item => (item.value === "No" || item.value === "NA"));
      // Check if any 'No' item in Array 1 does not exist in Array 2 or has an empty value
      const invalidItems = noItems.filter(item => {
        const matchingItem = this.stg5CommentValues.find(
          a2Item => a2Item.name === item.name && a2Item.value !== ""
        );
        return !matchingItem;
      });

      if (invalidItems.length > 0) {
        Swal.fire("Please enter your comment","","info");
      } else {
        const jsonData = JSON.stringify(this.stg5SelectedValues);
        const jsonCommentsData = JSON.stringify(this.stg5CommentValues);
        this.service.getFSCheckListItemsData(String(this.itemid)).subscribe((res) => {
          if (res != null && res !== '') {
            let resultSet = res as any;
            resultSet = resultSet.d.results;        
            if(resultSet.length > 0){              
                this.web.lists.getByTitle('FSTrackerCheckList').items.getById(Number(resultSet[0].ID)).update({
                  'Stage5Data':jsonData,
                  'Stage5Comments':jsonCommentsData
                }).then(() => {
                  Swal.fire("Successfully Confirmed","","success").then(()=>{
                    this.modalService.dismissAll();
                  });
                });
            }else{
                this.web.lists.getByTitle('FSTrackerCheckList').items.add({
                  'Stage5Data':jsonData,
                  'Title':String(this.itemid)
                }).then(() => {
                  Swal.fire("Successfully Confirmed","","success").then(()=>{
                    this.modalService.dismissAll();
                  });
                });
            }
          }
        }), (err: string) => {
          console.log("Error Occured " + err);
        };
      }
    }
  }

  getStg2Comments(control: any): string {    
    let comment : string = '';
    let filterdata = this.stg2CommentValues.filter((fitem: { name: string;}) => (fitem.name == control));
    if(filterdata.length > 0){
      comment = filterdata[0].value;
    }
    return comment;
  }

  getStg3Comments(control: any): string {    
    let comment : string = '';
    let filterdata = this.stg3CommentValues.filter((fitem: { name: string;}) => (fitem.name == control));
    if(filterdata.length > 0){
      comment = filterdata[0].value;
    }
    return comment;
  }

  getStg4Comments(control: any): string {    
    let comment : string = '';
    let filterdata = this.stg4CommentValues.filter((fitem: { name: string;}) => (fitem.name == control));
    if(filterdata.length > 0){
      comment = filterdata[0].value;
    }
    return comment;
  }

  getStg5Comments(control: any): string {    
    let comment : string = '';
    let filterdata = this.stg5CommentValues.filter((fitem: { name: string;}) => (fitem.name == control));
    if(filterdata.length > 0){
      comment = filterdata[0].value;
    }
    return comment;
  }


  updateSelectedOptiones(item: any, value: any, level: any){ 
    (window as any).global = window;
    // Check if the item exists in the array
    if(level == 'Stage2'){
      const index = this.stg2SelectedValues.findIndex(arritem => (arritem.name == item.control));
      if (index !== -1) {
        // Item exists, remove it and add the updated value
        this.stg2SelectedValues.splice(index, 1);
        this.stg2SelectedValues.push({ name: item.control, value: value, id: Number(item.id) });
      } else {
        // Item does not exist, add it to the array
        this.stg2SelectedValues.push({ name: item.control, value: value, id: Number(item.id) });
      }
    }

    if(level == 'Stage3'){
      const index = this.stg3SelectedValues.findIndex(arritem => (arritem.name == item.control));
      if (index !== -1) {
        // Item exists, remove it and add the updated value
        this.stg3SelectedValues.splice(index, 1);
        this.stg3SelectedValues.push({ name: item.control, value: value, id: Number(item.id) });
      } else {
        // Item does not exist, add it to the array
        this.stg3SelectedValues.push({ name: item.control, value: value, id: Number(item.id) });
      }
    }

    if(level == 'Stage4'){
      const index = this.stg4SelectedValues.findIndex(arritem => (arritem.name == item.control));
      if (index !== -1) {
        // Item exists, remove it and add the updated value
        this.stg4SelectedValues.splice(index, 1);
        this.stg4SelectedValues.push({ name: item.control, value: value, id: Number(item.id) });
      } else {
        // Item does not exist, add it to the array
        this.stg4SelectedValues.push({ name: item.control, value: value, id: Number(item.id) });
      }
    }

    if(level == 'Stage5'){
      const index = this.stg5SelectedValues.findIndex(arritem => (arritem.name == item.control));
      if (index !== -1) {
        // Item exists, remove it and add the updated value
        this.stg5SelectedValues.splice(index, 1);
        this.stg5SelectedValues.push({ name: item.control, value: value, id: Number(item.id) });
      } else {
        // Item does not exist, add it to the array
        this.stg5SelectedValues.push({ name: item.control, value: value, id: Number(item.id) });
      }
    }

  }

  updateStgComments(item: any, $event: any, level: any){
    (window as any).global = window; 
    if(level == 'Stage2'){
      // Check if the item exists in the array
      const index = this.stg2CommentValues.findIndex(arritem => (arritem.name == item.control));
      if (index !== -1) {
        // Item exists, remove it and add the updated value
        this.stg2CommentValues.splice(index, 1);
        this.stg2CommentValues.push({ name: $event.name, value: $event.value, id: Number(item.id) });
      } else {
        // Item does not exist, add it to the array
        this.stg2CommentValues.push({ name: $event.name, value: $event.value, id: Number(item.id) });
      }
    }

    if(level == 'Stage3'){
      // Check if the item exists in the array
      const index = this.stg3CommentValues.findIndex(arritem => (arritem.name == item.control));
      if (index !== -1) {
        // Item exists, remove it and add the updated value
        this.stg3CommentValues.splice(index, 1);
        this.stg3CommentValues.push({ name: $event.name, value: $event.value, id: Number(item.id) });
      } else {
        // Item does not exist, add it to the array
        this.stg3CommentValues.push({ name: $event.name, value: $event.value, id: Number(item.id) });
      }
    }

    if(level == 'Stage4'){
      // Check if the item exists in the array
      const index = this.stg4CommentValues.findIndex(arritem => (arritem.name == item.control));
      if (index !== -1) {
        // Item exists, remove it and add the updated value
        this.stg4CommentValues.splice(index, 1);
        this.stg4CommentValues.push({ name: $event.name, value: $event.value, id: Number(item.id) });
      } else {
        // Item does not exist, add it to the array
        this.stg4CommentValues.push({ name: $event.name, value: $event.value, id: Number(item.id) });
      }
    }

    if(level == 'Stage5'){
      // Check if the item exists in the array
      const index = this.stg5CommentValues.findIndex(arritem => (arritem.name == item.control));
      if (index !== -1) {
        // Item exists, remove it and add the updated value
        this.stg5CommentValues.splice(index, 1);
        this.stg5CommentValues.push({ name: $event.name, value: $event.value, id: Number(item.id) });
      } else {
        // Item does not exist, add it to the array
        this.stg5CommentValues.push({ name: $event.name, value: $event.value, id: Number(item.id) });
      }
    }

  }

}

enum Tabs{
  MAFP_Data_Submission = 0,
  EY_FS_Submission = 1,
  Audit_Reviews_and_clearance = 2,
  FS_signoff = 3,
  FS_Archiving = 4
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
  Draft = 'Draft',
  ReqIteration = 'Required Iteration'
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

enum EYTabs{
  EYFS = 0,
  FSTOKPMG = 1
}

enum S21TabIndex{
  Tab1 = 0,
  Tab2 = 1,
  Tab3 = 2,
  Tab4 = 3
}

enum S21Tabs{
  Tab1 = 'Tab1',
  Tab2 = 'Tab2',
  Tab3 = 'Tab3',
  Tab4 = 'Tab4'
}

enum EY{
  EYFS = 'Consultant FS Submission',
  FSTOKPMG = 'MAFP Review & FS Submission TO Auditor'

}

enum SignoffTabIndexs{
  GSFS = 0,
  MAFP = 1,
  //KPMG = 2
  Auditor = 2
}

enum Signoff{
  GSFS = 'GS FS Pack Submission',
  MAFP = 'MAFP Sign-off',
  KPMG = 'Auditor Sign-off'
  
}

enum Roles{
  Superuser = 'Superuser',
  GSteam = 'GS team',
  EY = 'EY',
  KPMG = 'KPMG',
  BU = 'BU Finance',
  Consultant = 'Consultant',
  Auditor = 'Auditor'
}

enum AH{
  CRT = 'CRT',
  GS = 'GS',
  EY = 'EY',
  KPMG = 'KPMG',
  BU = 'BU',
  Consultant = 'Consultant',
  Auditor = 'Auditor'
}

