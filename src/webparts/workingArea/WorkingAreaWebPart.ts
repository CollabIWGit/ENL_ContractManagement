import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';
// import "bootstrap";
// import '@popperjs/core';
import styles from './WorkingAreaWebPart.module.scss';
import * as strings from 'RequestFormWebPartStrings';
import { sp, List, IItemAddResult, UserCustomActionScope, Items, Item, ITerm, ISiteGroup, ISiteGroupInfo } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-groups/web";
import { MSGraphClientV3 } from '@microsoft/sp-http';
// import { SPComponentLoader } from '@microsoft/sp-loader';
import { ISiteUserInfo } from '@pnp/sp/site-users/types';
// import objMyCustomHTML from './Requestor_form';
import navHTML from '../enl_navbar';
import * as $ from 'jquery';
import * as sharepointConfig from '../../common/sharepoint-config.json';
import * as moment from 'moment';
import 'datatables.net';
// import {
//   SPHttpClient,
//   SPHttpClientResponse
// } from '@microsoft/sp-http';
import { sideMenuUtils } from "../../common/utils/sideMenuUtils";
import { Navigation } from 'spfx-navigation';

let SideMenuUtils = new sideMenuUtils();

// SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/4.1.0/css/bootstrap.min.css');
// SPComponentLoader.loadCss('https://cdn.datatables.net/1.10.25/css/jquery.dataTables.min.css');

// require('../../Assets/scripts/styles/mainstyles.css');
require('./../../common/scss/style.scss');
require('./../../common/css/style.css');
require('./../../common/css/common.css');
require('../../../node_modules/@fortawesome/fontawesome-free/css/all.min.css');

let currentUser: string;
var department;

export interface IWorkingAreaWebPartProps {
  description: string;
}

export default class WorkingAreaWebPart extends BaseClientSideWebPart<IWorkingAreaWebPartProps> {

  private graphClient: MSGraphClientV3;

  protected onInit(): Promise<void> {
    currentUser = this.context.pageContext.user.displayName;
    return new Promise<void>(async (resolve: () => void, reject: (error: any) => void): Promise<void> => {
      sp.setup({
        spfxContext: this.context as any
      });
 
      this.context.msGraphClientFactory
      .getClient('3')
      .then((client: MSGraphClientV3): void => {
        this.graphClient = client;
        resolve();
      }, err => reject(err));
    });
  }

  public async render(): Promise<void> {
    this.domElement.innerHTML = `
    <style>
    .main-container {
      display: flex;
      height: fit-content;
      position: relative;
  }
    .left-panel, .right-panel {
        position: fixed;
        transition: width 0.5s ease, margin 0.5s ease;
        height: 100vh;
    }
    .left-panel {
        width: 13%;
        left: 0;
        overflow-x: hidden;
    }
    .right-panel {
        right: 20px;
    }
    .middle-panel {
        flex: 1;
        width: 60%;
        margin-left: 13%;
        transition: width 0.5s ease, margin 0.5s ease;
    }
  
    .form-container {
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        height: 60rem;
    }
    .timeline {
        list-style: none;
        padding-left: 8px;
        height: 70%;
        overflow-y: scroll;
    }
  
    .timeline-item {
        margin-bottom: 20px;
    }
    .comment-box {
        height: 35%;
        padding-bottom: 1rem
    }
    .comment-input {
        height: 70%;
        width: 100%;
        padding: 10px;
        border: 1px solid #ccc;
        border-radius: 5px;
        resize: vertical;
    }
  
    .timeline::-webkit-scrollbar {
      width: 5px; /* Adjust the width to make the scrollbar thinner */
    }
  
    .timeline::-webkit-scrollbar-thumb {
      background-color: #888; /* Color of the scrollbar thumb */
    }
  
    .timelineHeader {
      text-align: center;
      // padding: 0.5rem;
      font-size: 1rem;
      font-weight: 500;
      // box-shadow: 0 7px 6px -6px #222;
      // border: 2px solid black;
      border-bottom: 1px solid black;
      margin-bottom: 0.5rem;
    }

    .fileUploadBtn {
      display: inline-block;
      padding: 6px 12px;
      cursor: pointer;
      border: 1px solid #ccc;
      border-radius: 4px;
      background-color: #f8f8f8;
      color: #333;
  }
  
  .fileUploadBtn {
      background-color: #e2e2e2;
  }

  #addComment {
    cursor: pointer;
    background-color: #062470;
    border: none;
    color: #fff;
    padding: 8px 12px;
    font-size: 16px;
    width: 40%;
    border-radius: 5px;
  }
  
  </style>
  
        <div class="main-container" id="content">
  
          <div id="nav-placeholder" class="left-panel"></div>
  
          <div id="middle-panel" class="middle-panel">
  
            <button id="minimizeButton"></button>
      
            <p id="contractStatus" style="color: green; position: absolute; top: 0; right: 0; margin: 1%;">Status: In Progress</p>
            
            <div id="workingAreaForm" style="width: 100%; height: 100%; padding: 2%">

              <div id="section_review_contract">
                  <div class="col-md-12 table-responsive"  style="border-bottom: 2px solid;">
                    <table class="table" id="table1">
                        <thead class="thead-dark">
                            <tr>
                                <th class="th-lg" scope="col">Contract</th>
                                <th scope="col">View</th>
                            </tr>
                        </thead>
                    </table>
                    <div id="tbl_contract" style="margin-top: 1.5em;">

                    </div>
                  </div>
              </div>

              <br>

              <div class="${styles['col-1-2']}">
                <div class="${styles.controls}" id="uploadFile" class="fileUploadBtn">
                  <label for="uploadContract">Upload Contract to Review</label>
                  <input type="file"  id="uploadContract">
                </div>
              </div>

            </div>

          </div>
    
          <div class="right-panel" id="rightPanel">
  
          </div>
  
        </div>
    `;

    await this.checkCurrentUsersGroupAsync();

    console.log(styles);

    //Generate Side Menu
    SideMenuUtils.buildSideMenu(this.context.pageContext.web.absoluteUrl);

    const urlParams = new URLSearchParams(window.location.search);
    const requestID = urlParams.get('requestid');

    const middlePanelID = document.getElementById('middle-panel');
    middlePanelID.style.marginRight = '27%';
    const rightPanelID = document.getElementById('rightPanel');
    rightPanelID.style.width = '26%';

    //Generate Timeline
    document.getElementById('rightPanel').innerHTML = `
      <div style="width: 100%; height:100%; background: white; padding-bottom: 30%;">
        <div class="timelineHeader">
          <p style="margin-bottom: 0px;">Comments</p>
        </div>
        <ul id="commentTimeline" class="timeline"></ul>
        <div class="comment-box">
          <textarea id="comment" class="comment-input" placeholder="Add your comment..."></textarea>
          <button id="addComment">Add Comment</button>
        </div>
      </div>
    `;

    const companyList = await sp.web.lists.getByTitle("Contract_Request").items.select("Company").filter(`ID eq ${requestID}`).get();
    const companyName = companyList[0].Company;
    console.log(companyList, companyName);

    this.renderRequestDetails(requestID, companyName);
    this.load_comments(requestID);

    //Minimize sidebar
    $('#minimizeButton').on('click', function() {
      const navPlaceholderID = document.getElementById('nav-placeholder');
      const middlePanelID = document.getElementById('middle-panel');
      const minimizeButtonID = document.getElementById('minimizeButton') as HTMLElement;

      if (navPlaceholderID && middlePanelID) {
        if (navPlaceholderID.offsetWidth === 0) {
          navPlaceholderID.style.width = '13%';
          middlePanelID.style.marginLeft = '13%';
          minimizeButtonID.style.left = '13%';
        } 
        else {
          navPlaceholderID.style.width = '0';
          middlePanelID.style.marginLeft = '0%'
          minimizeButtonID.style.left = '0%';
        }
      }
    });

    //Add comment button
    $("#addComment").click(async (e) => {
      console.log("Test New Comment");
      // icon_add_comment.classList.remove('hide');
      // icon_add_comment.classList.add('show');
      // icon_add_comment.classList.add('spinning');

      const currentUser = await sp.web.currentUser();
      let role;

      if(department === "Requestor"){

        role = "Requestor";

      }
      else if(department === "Owner"){
        role = "Owner";
      }
      else{
        role = "Despatcher";
      }

      const data = {

        Title: requestID,
        RequestID: requestID,
        Comment: $("#comment").val(),
        CommentBy: currentUser.UserPrincipalName,
        CommentDate: moment().format("DD/MM/YYYY HH:mm"),
        Role: role
      };

      console.log(data);

      await this.addComment(data);

      // icon_add_comment.classList.remove('spinning', 'show');
      // icon_add_comment.classList.add('hide');

      this.load_comments(requestID);

      $("#comment").val("");

    });

    var filename_add;
    var content_add;

    //Process uploaded file
    $('#uploadContract').on('change', async () => {
      const input = document.getElementById('uploadContract') as HTMLInputElement | null;

      var file = input.files[0];
      var reader = new FileReader();

      reader.onload = ((file1) => {
        return (e) => {
          console.log(file1.name);

          filename_add = file1.name,
            content_add = e.target.result

        };
      })(file);

      reader.readAsArrayBuffer(file);

      const library = "Contracts_ToReview";
      const folderPath = `/sites/ContractMgt/Contracts_ToReview/${companyName}/${requestID}`;

      await this.addFolderToDocumentLibrary(library, requestID)
        .then(async () => {
          try {
            await this.addFileToFolder2(folderPath, filename_add, content_add, requestID.toString());
          }
          catch (e) {
            console.log(e.message);
          }
      });

      this.renderRequestDetails(requestID, companyName);
      
    });
  
  }

  async addFolderToDocumentLibrary(libraryTitle, folderName) {
    try {
      // Initialize the PnP JS Library

      // Replace with the folder name you want to check

      //Check existence of company folder
      const exists = await this.folderExists(libraryTitle, folderName);

      if (exists) {
        console.log(`Folder '${folderName}' exists.`);
      }
      else {
        const library = sp.web.lists.getByTitle(libraryTitle);

        // Create a new folder
        await library.rootFolder.folders.add(folderName);

        console.log(`Folder '${folderName}' created successfully.`);
      }

      // Get the document library by title

    } catch (error) {
      console.error(`Error creating folder: ${error.message}`);
    }
  }

  async addFileToFolder2(folderPath, fileName, fileContent, requestId) {
    try {
      const fileData = await sp.web.getFolderByServerRelativeUrl(folderPath)
        .files.add(fileName, fileContent, false);

      const item = await fileData.file.getItem();
      await item.update({
        Request_Id: requestId
      });

      console.log('File uploaded successfully.');
      alert('File uploaded successfully.');
    } catch (error) {
      console.error('Error uploading file:', error);
      alert('Error uploading file.');
      throw error;
    }
  }

  async folderExists(libraryTitle, folderName) {
    try {
      // Initialize the PnP JS Library
      // Get the document library by title
      const library = sp.web.lists.getByTitle(libraryTitle);

      // Check if the folder exists
      const folder = await library.rootFolder.folders.getByName(folderName).select("Exists").get();

      return folder.Exists;
    }
    catch (error) {
      console.error(`Error checking folder existence: ${error.message}`);
      return false;
    }
  }

  private async renderRequestDetails(id: any, companyName: string) {

    $("#tbl_contract").html('');

    $("#section_review_contract").css("display", "block");

    this.getFileDetailsByFilter('Contracts_ToReview', id, companyName)
    .then((fileDetailsArray) => {
      if (fileDetailsArray && fileDetailsArray.length > 0) {
        console.log("File details:", fileDetailsArray);

        let html: string = '<div class="form-row">';
        html += `
                <div style="width: 100%; max-height: 300px; overflow-y: scroll;">
                    <table class="table">
                        <tbody>
        `;
        
        fileDetailsArray.forEach(fileItem => {
            html += `
                <tr>
                    <td scope="row">${fileItem.Name}</td>
                    <td>
                        <ul class="list-inline m-0">
                            <li class="list-inline-item">
                                <button id="btn_view_${fileItem.UniqueId}" class="btn btn-secondary btn-sm rounded-circle" type="button" data-toggle="tooltip" data-placement="top" title="View" style="display: none;">
                                    <i class="fas fa-eye"></i>
                                </button>
                            </li>
                            <li class="list-inline-item">
                                <button id="modalActivate_${fileItem.UniqueId}" class="btn btn-secondary btn-sm rounded-circle" type="button" data-toggle="modal" data-target="#exampleModalPreview" style="display: block; width: auto;">
                                    <i class="fas fa-eye"></i>
                                </button>
                            </li>
                        </ul>
                    </td>
                </tr>
            `;
        });
        
        html += `
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        `;

        const listContainer: Element = this.domElement.querySelector('#tbl_contract');
        listContainer.innerHTML = html;

        fileDetailsArray.forEach(fileDetails => {
          $(`#modalActivate_${fileDetails.UniqueId}`).click(() => {
            window.open(`ms-word:ofv|u|https://frcidevtest.sharepoint.com/${fileDetails.ServerRelativeUrl}`, '_blank');
          });
        });
      } 
      else {
        console.log("No items found.");
      }
    })
    .catch(error => {
      console.error("Error retrieving file details:", error);
    });

  }

  public async checkCurrentUsersGroupAsync() {
    let groupList = await sp.web.currentUser.groups();

    if (groupList.filter(g => g.Title == sharepointConfig.Groups.Requestor).length == 1) {
      department = "Requestor";
      console.log("You are a requestor", department);
      // $(".legalDept").css("display", "none");
    }
    else if (groupList.filter(g => g.Title == sharepointConfig.Groups.Owner).length == 1) {
      department = "Owner";
      console.log("You are an", department);
      // $(".legalDept").css("display", "none");
      // $('#commentSection').hide();
    }
    else if (groupList.filter(g => g.Title == sharepointConfig.Groups.Despatcher).length == 1) {
      department = "Despatcher";
      console.log("You are a", department);
      // $(".legalDept").css("display", "block");
    }
    else {
      department = "null";
      console.log("You are not in any group");
      // $(".legalDept").css("display", "none");
    }
  }


  //Load Timeline comments
  public async load_comments(updateRequestID) {
    // let userEmail = "";
    const timeline = document.getElementById('commentTimeline');
    timeline.innerHTML = '';
    const CommentList = await sp.web.lists.getByTitle("Comments").items.select("RequestID,Comment,CommentBy,CommentDate").filter(`RequestID eq '${updateRequestID}'`).get();
    console.log('Commentlist',CommentList);
    // userEmail = CommentList[0].CommentBy;
    const users: any[] = await sp.web.siteUsers();
    // let userTitle = '';
    // users.forEach(user => {
      // if (user.Email === userEmail) {
      //   userTitle = user.Title;
      //   return;
      // }
    // });
    // if (userTitle === '') {
    //   console.log('User with email ' + userEmail + ' not found.');
    // }
    CommentList.forEach(item => {
      const comment = item.Comment;
      const commentDate = item.CommentDate;
      let userEmail = item.CommentBy;
      let userTitle = '';
      users.forEach(user => {
        if (user.Email === userEmail) {
          userTitle = user.Title;
          return;
        }
      });
      const timelineItem = document.createElement('li');
      timelineItem.className = 'timeline-item';
      timelineItem.innerHTML = `
        <div style="display: flex">
          <p style="margin-bottom: 0px; font-style: italic; color: #3870ff">#${userTitle} -&nbsp; ${commentDate}</p>
        </div>
        <div>${comment}</div>
      `;
      timeline.appendChild(timelineItem);
    });

    timeline.scrollTop = timeline.scrollHeight;
  }

  async addComment(data) {
    try {
      const iar = await sp.web.lists.getByTitle("Comments").items.add(data);

      alert("Comment added succesfully.");
    }
    catch (e) {
      alert("An error occured." + e.message);
    }
  }

  async getFileDetailsByFilter(libraryName, reqId, companyName) {
    try {
      let folderPath = libraryName + "/" + companyName + "/" + reqId;
      let currentWebUrl = this.context.pageContext.web.absoluteUrl;
      let requestUrl = `${currentWebUrl}/_api/web/GetFolderByServerRelativeUrl('${folderPath}')/Files?$orderby=TimeCreated desc`;

      const response = await fetch(requestUrl, {
        method: 'GET',
        headers: {
            'Accept': 'application/json;odata=verbose'
        }
      });

      if (!response.ok) {
        throw new Error(`Error fetching folder contents: ${response.statusText}`);
      }

      const data = await response.json();
      const files = data.d.results;

      // Log the results
      console.log(files);

      if (files.length > 0) {
        return files;
      }

      return null;
    } 
    catch (error) {
      console.log(error);
      return null;
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
