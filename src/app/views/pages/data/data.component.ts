import { AfterViewInit, Component, OnInit } from '@angular/core';
import { FormBuilder, FormGroup, FormControl, Validators, NgForm, AbstractControl } from '@angular/forms';
import { NgbModalConfig, NgbModal, NgbDateStruct } from '@ng-bootstrap/ng-bootstrap';
import { MatSnackBar } from '@angular/material/snack-bar';
import { MatDialog } from '@angular/material/dialog';
import { DialogConfirmComponent } from './dialog-confirm/dialog-confirm.component';
import * as spns from 'sp-pnp-js';
import { environment } from '../../../../environments/environment';
//import { ApiService } from '/../../../api.service';
import { ApiService } from '../../../api.service';
import { Select2OptionData } from 'ng-select2';
import { Options } from 'select2';
import { ActivatedRoute } from '@angular/router';
import * as moment from 'moment';
import Swal from 'sweetalert2'
import { Router } from '@angular/router';

@Component({
  selector: 'app-data',
  templateUrl: './data.component.html',
  styleUrls: ['./data.component.scss'],
  providers: [NgbModalConfig, NgbModal]
})
export class DataComponent implements OnInit, AfterViewInit {  
  web = new spns.Web(environment.sp_URL);
  // FS Tracker
  listName:string = "FSTracker";  
  //docLibPath:string = environment.docLibPath;
  itemid: number = 0;
  pageSource:any;
  loginUserName: string = "Test@test.com";
  DisplayName: string = "";
  MasterData:any = [];
  turnAroundDays:any =[];
  //Users:any = [];
  //Users:any = [];
  public Users: Array<Select2OptionData> = [];
  
  fileInformationData :any = [];
  public files: any[] = [];
  isLinear = false;
  //firstFormGroup: FormGroup;
  //secondFormGroup: FormGroup;
  Category :any = [];
  ConsolidatedStandalone:any = [];
  Entities:any = [];
  DormantOperational:any = [];
  BusinessUnit:any = [];
  Country:any = [];
  Status:any = [];
  submitted = false;
  // Item variables
  selectedItem:any =[];
  angForm = new FormGroup({
    Category: new FormControl(''),
    ConsolidatedStandalone: new FormControl(''),
    EntitityCode: new FormControl(''),
    Entities: new FormControl(''),
    DormantOperational: new FormControl(''),
    BusinessUnit: new FormControl(''),
    Country: new FormControl(''),
    Preparer: new FormControl(''),
    Reviewer: new FormControl(''),
    TargetFSCompletionDate: new FormControl(''),
    Stage1Deadline: new FormControl(''),
    Stg1Comments:new FormControl(''),
    Stage2Deadline: new FormControl(''),
    Stage21Deadline: new FormControl(''),
    Stage31Deadline: new FormControl(''),
    Stage32Deadline: new FormControl(''),
    Stage33Deadline: new FormControl(''),
    Stage4Deadline: new FormControl(''),
    Stage42Deadline: new FormControl(''),
    Stage43Deadline: new FormControl(''),
    Stage5Deadline: new FormControl('')
  });
  
  constructor(private _snackBar: MatSnackBar, public dialog: MatDialog,private _formBuilder: FormBuilder,private service: ApiService,
    private activatedRoute:ActivatedRoute, private router: Router) { 
    //this.createForm();
  }
 

  
// createForm() {
  //   this.angForm = this.fb.group({        
  //     Category:['']
  //     });
  // }
  ngOnInit(): void {
    (window as any).global = window;
    window.scrollTo(0, 0);    
    this.getUserName();    
    this.service.getMasterData().subscribe((res) => {
      if (res != null && res !== '') {
        // this.MasterData = res as any;     
        // this.Category = [];
        let resultSet = res as any;
        resultSet = resultSet.d.results;
        this.Category = resultSet.filter((items: { Title: string; }) => items.Title == "Category");
        this.BusinessUnit = resultSet.filter((items: { Title: string; }) => items.Title == "Business Unit");
        this.Country = resultSet.filter((items: { Title: string; }) => items.Title == "Country");
        this.ConsolidatedStandalone = resultSet.filter((items: { Title: string; }) => items.Title == "Consolidated Standalone");
        this.DormantOperational = resultSet.filter((items: { Title: string; }) => items.Title == "Dormant Operational");
        this.Entities = resultSet.filter((items: { Title: string; }) => items.Title == "Entity Name");
      }
    });
    // this.web.lists.getByTitle(this.listName).items.add({       
    //   'ItemStatus': 'InActive'
    // }).then((item: any) => {
    //   this.itemid = item.data.ID;
    //   this.initiateNewItem();
    // }).catch((err: any) => {            
    //   console.log(err);
    // });

    this.angForm = new FormGroup({
      Category: new FormControl('', Validators.required),
      ConsolidatedStandalone: new FormControl('', Validators.required),
      EntitityCode: new FormControl('', Validators.required),
      Entities: new FormControl('', Validators.required),
      DormantOperational: new FormControl('', Validators.required),
      BusinessUnit: new FormControl('', Validators.required),
      Country: new FormControl('', Validators.required),
      Preparer: new FormControl('', Validators.required),
      Reviewer: new FormControl('', Validators.required),
      TargetFSCompletionDate: new FormControl('', Validators.required),
      Stage1Deadline: new FormControl('', Validators.required),
      Stg1Comments:new FormControl(''),
      Stage2Deadline: new FormControl('', Validators.required),
      Stage21Deadline: new FormControl('', Validators.required),
      Stage31Deadline: new FormControl('', Validators.required),
      Stage32Deadline: new FormControl('', Validators.required),
      Stage33Deadline: new FormControl('', Validators.required),
      Stage4Deadline: new FormControl('', Validators.required),
      Stage42Deadline: new FormControl('', Validators.required),
      Stage43Deadline: new FormControl('', Validators.required),
      Stage5Deadline: new FormControl('', Validators.required)
    });
  }
  getUserName() {    
    this.service.getUserName().subscribe((res) => {
      if (res != null && res !== '') {
        const user = res as any;
        this.loginUserName = user.d.Email;
        this.DisplayName = user.d.Title;
      }
    });
    this.Users = [];
    this.service.getGroupUsers("FSUsers").subscribe((res) => {
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
  initiateNewItem(){
    this.service.getItemById(this.itemid).subscribe((res) => {
      if (res != null) {  
        const itemresult = res as any; 
        this.selectedItem = itemresult.d;        
        // this.setCurrentStage()       
        // this.setStagesDate();
        // this.setStagesStatus();
        // this.bindFileData();
      }
    });
  }
  get f(): { [key: string]: AbstractControl } {
    return this.angForm.controls;
  }
  validateencode(){
    let entitycode:string = this.angForm.value.EntitityCode;
    this.service.checkFSTrackerData(entitycode).subscribe((res) => {
        if (res != null && res !== '') {
          let resultSet = res as any;
          resultSet = resultSet.d.results;
          if(resultSet.length >= 1){
            Swal.fire("Already record exists with same entity code","","info");
          }
      }
    });
  }
  saveUpdateRequest(status: string, stage: string) {
   
    (window as any).global = window;
    //this.initiateEditItem();
    this.submitted = true;
    // stop here if form is invalid
    if (this.angForm.invalid) {
      return;
    }    
    let entitycode:string = this.angForm.value.EntitityCode;
    this.service.checkFSTrackerData(entitycode).subscribe((res) => {
        if (res != null && res !== '') {
          let resultSet = res as any;
          resultSet = resultSet.d.results;
          if(resultSet.length == 0){
            var TargetFSCompletionDate, Stage1Deadline,Stage2Deadline,Stage21Deadline,Stage31Deadline,Stage32Deadline,Stage33Deadline,Stage4Deadline,Stage42Deadline,Stage43Deadline,Stage5Deadline;
            if(this.angForm.value.TargetFSCompletionDate != null && this.angForm.value.TargetFSCompletionDate != ""){      
              let date = this.angForm.value.TargetFSCompletionDate;
              TargetFSCompletionDate = moment([date.year, date.month -1 , date.day]).add(12, 'hours').toDate();
            }
            if(this.angForm.value.Stage1Deadline != null && this.angForm.value.Stage1Deadline != ""){      
              let date = this.angForm.value.Stage1Deadline;
              Stage1Deadline = moment([date.year, date.month -1 , date.day]).add(12, 'hours').toDate();
            }
            if(this.angForm.value.Stage2Deadline != null && this.angForm.value.Stage2Deadline != ""){      
              let date = this.angForm.value.Stage2Deadline;
              Stage2Deadline = moment([date.year, date.month -1 , date.day]).add(12, 'hours').toDate();
            }
            if(this.angForm.value.Stage21Deadline != null && this.angForm.value.Stage21Deadline != ""){      
              let date = this.angForm.value.Stage21Deadline;
              Stage21Deadline = moment([date.year, date.month -1 , date.day]).add(12, 'hours').toDate();
            }
            if(this.angForm.value.Stage31Deadline != null && this.angForm.value.Stage31Deadline != ""){      
              let date = this.angForm.value.Stage31Deadline;
              Stage31Deadline = moment([date.year, date.month -1 , date.day]).add(12, 'hours').toDate();
            }
            if(this.angForm.value.Stage32Deadline != null && this.angForm.value.Stage32Deadline != ""){      
              let date = this.angForm.value.Stage32Deadline;
              Stage32Deadline = moment([date.year, date.month -1 , date.day]).add(12, 'hours').toDate();
            }
            if(this.angForm.value.Stage33Deadline != null && this.angForm.value.Stage33Deadline != ""){      
              let date = this.angForm.value.Stage33Deadline;
              Stage33Deadline = moment([date.year, date.month -1 , date.day]).add(12, 'hours').toDate();
            }
            if(this.angForm.value.Stage4Deadline != null && this.angForm.value.Stage4Deadline != ""){      
              let date = this.angForm.value.Stage4Deadline;
              Stage4Deadline = moment([date.year, date.month -1 , date.day]).add(12, 'hours').toDate();
            }
            if(this.angForm.value.Stage42Deadline != null && this.angForm.value.Stage42Deadline != ""){      
              let date = this.angForm.value.Stage42Deadline;
              Stage42Deadline = moment([date.year, date.month -1 , date.day]).add(12, 'hours').toDate();
            }
            if(this.angForm.value.Stage43Deadline != null && this.angForm.value.Stage43Deadline != ""){      
              let date = this.angForm.value.Stage43Deadline;
              Stage43Deadline = moment([date.year, date.month -1 , date.day]).add(12, 'hours').toDate();
            }
            if(this.angForm.value.Stage5Deadline != null && this.angForm.value.Stage5Deadline != ""){      
              let date = this.angForm.value.Stage5Deadline;
              Stage5Deadline = moment([date.year, date.month -1 , date.day]).add(12, 'hours').toDate();
            }
            let Preparer , Reviewer ;
            if(this.angForm.value.Preparer != ""){
              let rEmail = this.Users.filter(data => data.id === this.angForm.value.Preparer);
              if(rEmail.length > 0){
                Preparer = rEmail[0].text;
              }      
            }
            if(this.angForm.value.Reviewer != ""){    
              let pEmail = this.Users.filter(data => data.id === this.angForm.value.Reviewer);      
              if(pEmail.length > 0){
                Reviewer = pEmail[0].text;
              }
            }
            
            let year=(new Date()).getFullYear();
            this.web.lists.getByTitle(this.listName).items.add({ 
              'Title':this.angForm.value.EntitityCode,
              'Category':this.angForm.value.Category,
              'EntityName':this.angForm.value.Entities,
              'ConsolidatedStandalone':this.angForm.value.ConsolidatedStandalone,
              'DormantOperational':this.angForm.value.DormantOperational,
              'BusinessUnit':this.angForm.value.BusinessUnit,
              'Country':this.angForm.value.Country,
              'TargetFSCompletionDate':TargetFSCompletionDate,      
              'Preparer':Preparer,
              'Reviewer':Reviewer,
              'PreparerEmail':this.angForm.value.Preparer,
              'ReviewerEmail':this.angForm.value.Reviewer,
              'Stage1Deadline': Stage1Deadline,
              'Stage2Deadline': Stage2Deadline,
              'Stage21Deadline': Stage21Deadline,
              'Stage31Deadline': Stage31Deadline,
              'Stage32Deadline': Stage32Deadline,
              'Stage33Deadline': Stage33Deadline,
              'Stage4SubmissionDate': Stage4Deadline,
              'Stage42SubmissionDate': Stage42Deadline,
              'Stage43SubmissionDate': Stage43Deadline,
              'Stage5SubmissionDate': Stage5Deadline,      
              'Stage1Status': "Draft",
              'CurrentStage':stages.Stage1,
              'CurrentStatus': "Draft",
              'PendingWith':stages.Stage1,
              'Stg1Comments':this.angForm.value.Stg1Comments,
              'ItemStatus': 'Active',
              'Year':year.toString()
            }).then((item: any) => {    
              
              this.service.getFSTurnAroundDaysById('0').subscribe((resdays) => {
                if(resdays!=null){
                  const daysresult = resdays as any; 
                  this.turnAroundDays = daysresult.d.results[0];
                  this.web.lists.getByTitle('FSTurnAroundDays').items.add({ 
                    'Stage1ConsultantApproval': this.turnAroundDays.Stage1ConsultantApproval,
                    'Stage1ConsultantReject': this.turnAroundDays.Stage1ConsultantReject,
                    'Stage21ConsultantReview': this.turnAroundDays.Stage21ConsultantReview,
                    'Stage21GSReview': this.turnAroundDays.Stage21GSReview,
                    'Stage21ConsultantReject': this.turnAroundDays.Stage21ConsultantReject,
                    'Stage22GSReview': this.turnAroundDays.Stage22GSReview,
                    'Stage31AuditorReview': this.turnAroundDays.Stage31AuditorReview,
                    'Stage31GSReview': this.turnAroundDays.Stage31GSReview,
                    'Stage31ConsultantReview': this.turnAroundDays.Stage31ConsultantReview,
                    'Stage32AuditorReview': this.turnAroundDays.Stage32AuditorReview,
                    'Stage32GSReview': this.turnAroundDays.Stage32GSReview,
                    'Stage32ConsultantReview': this.turnAroundDays.Stage32ConsultantReview,
                    'Stage33AuditorReview': this.turnAroundDays.Stage33AuditorReview,
                    'Stage33GSReview': this.turnAroundDays.Stage33GSReview,
                    'Stage33ConsultantReview': this.turnAroundDays.Stage33ConsultantReview,
                    'Stage41GSReview': this.turnAroundDays.Stage41GSReview,
                    'Stage41CRTReview': this.turnAroundDays.Stage41CRTReview,
                    'Stage42GSReview': this.turnAroundDays.Stage42GSReview,
                    'Stage42CRTReview': this.turnAroundDays.Stage42CRTReview,
                    'Stage43GSReview': this.turnAroundDays.Stage43GSReview,
                    'Stage43AuditorReview': this.turnAroundDays.Stage43AuditorReview,
                    'Stage5GSReview': this.turnAroundDays.Stage5GSReview,
                    'Stage5CRTReview': this.turnAroundDays.Stage5CRTReview,
                    'FSID': item.data.ID.toString(),
                    'EntityCode':this.angForm.value.EntitityCode,
                    'EntityName':this.angForm.value.Entities
                  }).then((itemTD: any) => {  
                    this.ensureFolder(item.data.ID);    
                  }).catch((err: any) => { 
                    Swal.fire("Failed to update","","info");
                  });
                }
              });
              
            }).catch((err: any) => { 
              Swal.fire("Failed to update","","info");
            });

          }else{
            Swal.fire("Already record exists with same entity code","","info");
          }
      }
    });
  }

  private async ensureFolder(ItemID: string): Promise<any> {
    let folderpath = environment.docLibPath+ItemID;
    const folder = await this.web.getFolderByServerRelativePath(folderpath).select('Exists').get();
    if (!folder.Exists) {
        await this.web.folders.add(folderpath);
    }
    Swal.fire("Successfully Saved","","success").then(() => {
      this.router.navigate(['/app/edit'],{queryParams:{itemid:ItemID,source:"home"}})
    });
  }

  // onFileChange(pFileList: any){
  //   this.files = Object.keys(pFileList).map(key => pFileList[key]);
  //   (window as any).global = window;    
  //   const component: DataComponent = this; 
  //      let currentYear = new Date();
  //      component.files.forEach(file => { 
  //       let filename = component.itemid +"_"+ file.name; 
  //       component.web.getFolderByServerRelativeUrl(component.docLibPath).files.add(filename, file, true).then(async function(f) {
  //         let year = String(currentYear.getFullYear());
  //         (await f.file.getItem()).update({
  //           Title: file.name,
  //           Stage: stages.Stage1,
  //           Status: Status.Submitted,
  //           ReferenceID: String(component.itemid),
  //           Year: year,
  //           IsActive: "Yes"
  //         }).then((d)=>{
  //           component.bindFileData();            
  //         });          
  //       });
  //     });
      
  // }

  bindFileData(){
    this.service.getFileData(String(this.itemid)).subscribe((res) => {
      if (res != null && res !== '') {
        let result= res as any;
        this.fileInformationData = result.d.results;
      }
    });
  }

  downloadFile(f:any, stage:string){
    (window as any).global = window;
        if(this.selectedItem.Stage1Status !=null){
          this.web.lists.getByTitle(this.listName).items.getById(Number(this.itemid)).update({  
            'Stage1Status': 'Submitted'
          }).then((item: any) => {
            //this.initiateEditItem();
          });
    }
  }
  // download(url: string, filename?: string): Observable<Download> {
  //   return this.http.get(url, {
  //     reportProgress: true,
  //     observe: 'events',
  //     responseType: 'blob'
  //   }).pipe(download(blob => saveAs(blob, filename)))
  // }
  async deleteFile(f:any){
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
  scrolltop(){
    window.scrollTo(0, 0);
  }
  ngAfterViewInit(): void {
    window.scrollTo(0, 0);
  }
  goBack() { 
    this.router.navigate(['/app/home']);
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
  Complete = 'Complete'
}