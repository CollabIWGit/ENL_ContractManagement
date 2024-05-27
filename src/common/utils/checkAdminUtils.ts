// import { MSGraphClient } from '@microsoft/sp-http';
// // import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

// import { loaderUtils } from "../../common/utils/loaderUtils";
// let LoaderUtils = new loaderUtils();

// import { sideMenuUtils } from "../../common/utils/sideMenuUtils";
// let SideMenuUtils = new sideMenuUtils();

// import * as $ from 'jquery';

// // import * as sharepointConfig from '../../common/sharepoint-config.json';

// export class checkAdminUtils {
//     public checkifUserIsAdmin(graphClient: MSGraphClient): Promise<any> {
//         if (!graphClient) {
//             return;
//         }
//         return new Promise((resolve, reject) => {
//             graphClient.api(`/me/transitiveMemberOf/microsoft.graph.group?$count=true&$top=999`).get((errorGroup, groups: any, rawResponseGroup?: any) => {
//                 if (errorGroup) {
//                     console.log(errorGroup);
//                     return reject(errorGroup);
//                 }
//                 //console.log(groups);
//                 var groupList = groups.value;

//                 if (groupList.filter(g => g.displayName == sharepointConfig.Groups.ManageActionsGroup).length < 1)
//                     sessionStorage.setItem("manageActions", "0");
//                 else
//                     sessionStorage.setItem("manageActions","1");

//                 if (groupList.filter(g => g.displayName == sharepointConfig.Groups.AdminGroup).length < 1) {
//                     sessionStorage.setItem("admin", "0");
//                     return resolve(false);
//                 }
//                 else {
//                     sessionStorage.setItem("admin", "1");
//                     return resolve(true);
//                 }
//             });
//         });
//     }

//     public checkifUserIsAdminMain(graphClient: MSGraphClient): Promise<any> {
//         if (!graphClient) {
//             return;
//         }
//         return new Promise((resolve, reject) => {
//             var adminMain = sessionStorage.getItem("adminMain");

//             if (adminMain != null){
//                 if (adminMain == "1")
//                     return resolve(true);
//                 else
//                     return resolve(false);
//             }

//             graphClient.api(`/me/transitiveMemberOf/microsoft.graph.group?$count=true&$top=999`).get((errorGroup, groups: any, rawResponseGroup?: any) => {
//                 if (errorGroup) {
//                     console.log(errorGroup);
//                     return reject(errorGroup);
//                 }
//                 //console.log(groups);
//                 //console.log(groups["@odata.context"]);
//                 //console.log(groups["@odata.nextLink"]); //check if undefined
//                 var groupList = groups.value;

//                 if (groupList.filter(g => g.displayName == sharepointConfig.Groups.AdminMain).length < 1) {
//                     sessionStorage.setItem("adminMain", "0");
//                     return resolve(false);
//                 }
//                 else {
//                     sessionStorage.setItem("adminMain", "1");
//                     return resolve(true);
//                 }
//             });
//         });
//     }

//     public checkifUserIsAdminSub(graphClient: MSGraphClient): Promise<any> {
//         if (!graphClient) {
//             return;
//         }
//         return new Promise((resolve, reject) => {
//             var adminSub = sessionStorage.getItem("adminSub");

//             if (adminSub != null){
//                 if (adminSub == "1")
//                     return resolve(true);
//                 else
//                     return resolve(false);
//             }

//             graphClient.api(`/me/transitiveMemberOf/microsoft.graph.group?$count=true&$top=999`).get((errorGroup, groups: any, rawResponseGroup?: any) => {
//                 if (errorGroup) {
//                     console.log(errorGroup);
//                     return reject(errorGroup);
//                 }
//                 //console.log(groups);
//                 //console.log(groups["@odata.context"]);
//                 //console.log(groups["@odata.nextLink"]); //check if undefined
//                 var groupList = groups.value;

//                 if (groupList.filter(g => g.displayName == sharepointConfig.Groups.AdminSub).length < 1) {
//                     sessionStorage.setItem("adminSub", "0");
//                     return resolve(false);
//                 }
//                 else {
//                     sessionStorage.setItem("adminSub", "1");
//                     return resolve(true);
//                 }
//             });
//         });
//     }

//     public checkIfUserIsAdminAndSetupPage(absoluteUrl: string, graphClient: MSGraphClient) {
//         var admin = sessionStorage.getItem("admin");
//         var manageActions = sessionStorage.getItem("manageActions");
//         if (admin == null || admin == undefined || manageActions == null || manageActions == undefined) {
//             LoaderUtils.toggleLoader(true);
//             this.checkifUserIsAdmin(graphClient).then((response) => {
//                 LoaderUtils.toggleLoader(false);
//                 this.renderPage(response, absoluteUrl);
//             })
//                 .catch((error) => {
//                     LoaderUtils.toggleLoader(false);
//                     console.log(error);
//                 });
//         }
//         else {
//             if (admin == "1") { //user is admin
//                 this.renderPage(true, absoluteUrl);
//             }
//             else {
//                 this.renderPage(false, absoluteUrl);
//             }
//         }
//     }

//     public checkIfUserIsAdminAndSetupHomePage(absoluteUrl: string, graphClient: MSGraphClient) {
//         var admin = sessionStorage.getItem("admin");
//         var manageActions = sessionStorage.getItem("manageActions");
//         if (admin == null || admin == undefined || manageActions == null || manageActions == undefined) {
//             LoaderUtils.toggleLoader(true);
//             this.checkifUserIsAdmin(graphClient).then((response) => {
//                 LoaderUtils.toggleLoader(false);
//                 this.renderPage(response, absoluteUrl);
//                 this.hideShowCard(response);
//             })
//                 .catch((error) => {
//                     LoaderUtils.toggleLoader(false);
//                     console.log(error);
//                 });
//         }
//         else {
//             if (admin == "1") { //user is admin
//                 this.renderPage(true, absoluteUrl);
//                 this.hideShowCard(true);
//             }
//             else {
//                 this.renderPage(false, absoluteUrl);
//                 this.hideShowCard(false);
//             }
//         }
//     }

//     private renderPage(isAdmin: boolean, absoluteUrl: string) {
//         if (!isAdmin) {
//             SideMenuUtils.buildSideMenu(absoluteUrl);
//             $("#content").show();
//         }
//         else {
//             SideMenuUtils.buildSideMenu(absoluteUrl);
//             $("#content").show();
//         }
//     }

//     private hideShowCard(isAdmin: boolean) {
//         if (!isAdmin) {
//             $("#docInArchiveCard").hide();
//             $("#docInDocCard").show();
//             $("#docOutArchiveCard").hide();
//             $("#docOutDocCard").show();
//             $("#action").show();
//         }
//         else {
//             $("#docInArchiveCard").show();
//             $("#docInDocCard").show();
//             $("#docOutArchiveCard").show();
//             $("#docOutDocCard").show();
//             $("#action").show();
//         }
//     }
// }