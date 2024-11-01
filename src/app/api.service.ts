import { Injectable } from '@angular/core';
import { HttpClient, HttpHeaders, HttpParams } from '@angular/common/http';
// import * as $ from 'jquery';
import { environment } from '../environments/environment';
import * as spns from 'sp-pnp-js';

@Injectable({
  providedIn: 'root'
})
export class ApiService {  
  sp_URL =environment.sp_URL;
  sproject_URL =environment.sproject_URL;
  digestValue = '';
  name: string = "";  
  timesheet: string = "";
  objTimesheetList: any;

  constructor(private http: HttpClient) { }

  getUserName() {
    return this.http.get(this.sp_URL + '_api/Web/CurrentUser', { headers: { Accept: 'application/json;odata=verbose' } })
  }
  isMemberOfGroup(usedID:number) {     
    const requestURL =  this.sp_URL + "/_api/web/sitegroups/getByName('IsAdmin')/Users?$filter=Id eq "+ usedID + "";
    return this.http.get(requestURL, { headers: { Accept: 'application/json;odata=verbose' } })
  }

  getGroupID(group:string) {     
    const requestURL =  this.sp_URL + "_api/web/sitegroups/?$filter=Title eq '"+group+"'&$top=1";
    return this.http.get(requestURL, { headers: { Accept: 'application/json;odata=verbose' } })
  }

  getGroupUsers(group:string) {  
    return this.http.get(this.sp_URL + "_api/web/sitegroups/getByName('"+group+"')/users?$top=100", { headers: { Accept: 'application/json;odata=verbose' } })
  }
  getSiteUsers() {  
    return this.http.get(this.sp_URL + "_api/web/siteusers?$top=500", { headers: { Accept: 'application/json;odata=verbose' } })
  }

  getFsTrackerData() {
    const url = this.sp_URL +"_api/web/lists/getByTitle('FSTracker')/items?$select=ID,Title,Preparer,Category,EntityName,BusinessUnit,TargetFSCompletionDate,CurrentStage,CurrentStatus,Stage3CurrentState,Stage2CurrentState,Year,ConsolidatedStandalone,DormantOperational,Reviewer,Country,PreparerEmail,ReviewerEmail,"
    +"Stage1Status,Stage2Status,Stage21Status,Stage31Status,Stage32Status,Stage33Status,Stage4Status,Stage5Status,Stage1Deadline,Stage2Deadline,Stage21Deadline,Stage31Deadline,Stage32Deadline,Stage33Deadline,Stage4SubmissionDate,Stage5SubmissionDate,ActionHolder&$top=5000&$orderby=ID desc";
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }
  getFsEmailData() {
    const url = this.sp_URL +"_api/web/lists/getByTitle('EmailHistory')/items?$select=ID,Title,EmailTo,EmailCC,Stage,EmailType,Status,Preparer,Comments,Entity,EntityCode,EmailContent&$top=100&$orderby=ID desc";
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }
  getFSTurnAroundDaysData() {
    const url = this.sp_URL +"_api/web/lists/getByTitle('FSTurnAroundDays')/items?$select=ID,Title,*&$top=100&$orderby=ID desc";
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }
  getFSTurnAroundDaysById(id:string){
    const url = this.sp_URL + "_api/web/lists/getByTitle('FSTurnAroundDays')/items?$select=ID,Title,*,Author/Name,Author/Title&$expand=Author&$top=1000&$filter=(FSID eq '" + id + "')";
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }
  getUserProfile(user:string) {    
    const url = this.sp_URL + "_api/sp.userprofiles.peoplemanager/GetUserProfilePropertyFor(AccountName=@v,propertyName='Manager')?@v='" + encodeURIComponent(user) + "'";
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }

  getItemById(id: Number) {
    return this.http.get(this.sp_URL + "_api/web/lists/getByTitle('FSTracker')/items(" + id + ")??$select=ID,Title,*", { headers: { Accept: 'application/json;odata=verbose' } })
  }

  checkFileStatus(filePath:string) {
    const url = this.sp_URL +"_api/Web/GetFileByServerRelativeUrl('"+ filePath +"')";
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }
  checkFolderStatus(folder:string) {
    const url = this.sp_URL +"_api/Web/GetFolderByServerRelativeUrl('FSDocuments')/Folders?$filter=Name eq '"+folder+"'";
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }
  //https://majidalfuttaim.sharepoint.com/sites/LCM/_api/Web/GetFolderByServerRelativeUrl('LeaseDocuments/DCC')/Folders?$filter=Name eq 'Al Rasasi Trading'
  getFileData(itemid:string) {
    const url =this.sp_URL +"_api/web/lists/getByTitle('FSDocuments')/items?$select=ReferenceID,IsActive,Stage,Status,Year,ID,Title,File/Name,File/ServerRelativeUrl,Created,Modified,Editor/Title&$expand=File,Editor&$filter=((ReferenceID eq '"+ itemid +"') and (IsActive eq 'Yes'))";    
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }
  getFSTrackerHistoryData(itemID: string) {
    const url = this.sp_URL + "_api/web/lists/getByTitle('FSTrackerHistory')/items?$select=ID,Title,Stage,*,Author/Name,Author/Title&$expand=Author&$top=1000&$filter=(Title eq '" + itemID + "')&$orderby=ID desc";
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }
  getFSTrackerStgWiseHistoryData(itemID: string, stage: string, state: string) {
    debugger;
    let url = this.sp_URL + "_api/web/lists/getByTitle('FSTrackerHistory')/items?$select=ID,Title,Stage,State,*,Author/Name,Author/Title&$expand=Author&$top=1000&$filter=((Title eq '" + itemID + "') and (Stage eq '" + stage + "'))&$orderby=ID desc";
    if(stage != 'MAFP Data Submission' && stage != 'FS Archiving'){
      url = this.sp_URL + "_api/web/lists/getByTitle('FSTrackerHistory')/items?$select=ID,Title,Stage,State,*,Author/Name,Author/Title&$expand=Author&$top=1000&$filter=((Title eq '" + itemID + "') and (Stage eq '" + stage + "') and (State eq '" + encodeURIComponent(state) + "'))&$orderby=ID desc";
    }
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }
  getFSCheckListMasterData(stage: string, country: string) {
    const url = this.sp_URL + "_api/web/lists/getByTitle('MasterCheckList')/items?$select=ID,Title,OrderBy,Country,*&$filter=((Title eq '" + stage + "') and (Country eq '" + country + "'))&$orderby=OrderBy";
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }
  getFSCheckListStg23MasterData(stage: string, state: string, tab: string, country: string) {
    const url = this.sp_URL + "_api/web/lists/getByTitle('MasterCheckList')/items?$select=ID,Title,OrderBy,State,Country,*&$filter=((Title eq '" + stage + "') and (State eq '" + encodeURIComponent(state) + "') and (Category eq '"+ tab +"') and (Country eq '" + country + "'))&$orderby=OrderBy";
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }

  getFSCheckListStg23AllMasterData(stage: string, state: string, country: string) {
    const url = this.sp_URL + "_api/web/lists/getByTitle('MasterCheckList')/items?$select=ID,Title,OrderBy,State,Country,*&$filter=((Title eq '" + stage + "') and (State eq '" + encodeURIComponent(state) + "') and (Country eq '" + country + "'))&$orderby=OrderBy";
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }

  getFSCheckListStg45MasterData(stage: string, tab: string, country: string) {
    const url = this.sp_URL + "_api/web/lists/getByTitle('MasterCheckList')/items?$select=ID,Title,OrderBy,State,Country,*&$filter=((Title eq '" + stage + "') and (Category eq '"+ tab +"') and (Country eq '" + country + "'))&$orderby=OrderBy";
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }
  getFSCheckListItemsData(fsid: string) {
    const url = this.sp_URL + "_api/web/lists/getByTitle('FSTrackerCheckList')/items?$select=ID,Title,*&$filter=(Title eq '" + fsid + "')&$top=1";
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }
  checkFSTrackerData(entitycode:string) {
    const url = this.sp_URL +"_api/web/lists/getByTitle('FSTracker')/items?$select=ID,Title&$filter=(Title eq '"+ entitycode +"')";
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }

  
  
  checkFSUserRoles(role:string, email:string) {
    const url = this.sp_URL +"_api/web/lists/getByTitle('UserRoles')/items?$select=ID,Title,Useremail,Roles&$filter=((Useremail eq '"+ email +"') and (Roles eq '"+ role +"'))";
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }
  getUserRoles() {
    const url =this.sp_URL +"_api/web/lists/getByTitle('UserRoles')/items?$select=ID,Title,Roles,User/Title,User/EMail,Status&$expand=User&$top=5000&$filter=(Status eq 'Active')";    
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }
  getProfilePictureData() {
    const url =this.sp_URL +"_api/web/lists/getByTitle('ProfilePicture')/items?$select=AttachmentFiles,Title,ID,Attachments&$expand=AttachmentFiles";    
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }

  // Audit

  getExternalAuditData(assignedtome:string) {
    const url = this.sp_URL +"_api/web/lists/getByTitle('MAFPExternalAuditUAT')/items?$select=ID,*,Status,Level&$top=5000&$filter=((Level ne 'Completed') and (Level ne 'Submitted') and (ResponsibilityEmail eq '"+assignedtome+"'))&$orderby=ID desc";
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }

  getKPMGExternalAuditData(assignedtome:string) {
    const url =this.sp_URL +"_api/web/lists/getByTitle('MAFPExternalAuditUAT')/items?$select=ID,*,Status,Level&$top=5000&$filter=((Level eq 'Submitted') and (ResponsibilityEmail eq '"+assignedtome+"'))";
    //const url =this.sp_URL +"_api/web/lists/getByTitle('MAFPExternalAuditUAT')/items?$select=ID,*,Status,Level&$top=5000&$filter=(ResponsibilityEmail eq '"+assignedtome+"')";
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }
  getMasterData() {
    const url =this.sp_URL +"_api/web/lists/getByTitle('MasterData')/items?$select=ID,Title,Options,Code,Status&$filter=(Status eq 'Active')";    
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }
  getAllExternalAuditData() {
    const url = this.sp_URL +"_api/web/lists/getByTitle('MAFPExternalAuditUAT')/items?$select=ID,*,Status,Level&$top=5000";
    return this.http.get(url, { headers: { Accept: 'application/json;odata=verbose' } });
  }
}
