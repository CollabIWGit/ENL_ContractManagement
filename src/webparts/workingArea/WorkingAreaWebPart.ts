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
import { SPComponentLoader } from '@microsoft/sp-loader';
import { ISiteUserInfo } from '@pnp/sp/site-users/types';
// import objMyCustomHTML from './Requestor_form';
import navHTML from '../enl_navbar';
import * as $ from 'jquery';
import * as sharepointConfig from '../../common/sharepoint-config.json';
import * as moment from 'moment';
import 'datatables.net';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import { sideMenuUtils } from "../../common/utils/sideMenuUtils";
import { Navigation } from 'spfx-navigation';
import { DatePicker } from '@fluentui/react';

// import { SPHttpClient } from '@microsoft/sp-http';

// import DespatcherDashboardObj from './DespatcherDashboard';

let SideMenuUtils = new sideMenuUtils();

SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/4.1.0/css/bootstrap.min.css');
SPComponentLoader.loadCss('https://cdn.datatables.net/1.10.25/css/jquery.dataTables.min.css');

// require('../../Assets/scripts/styles/mainstyles.css');
require('./../../common/scss/style.scss');
require('./../../common/css/style.css');
require('./../../common/css/common.css');
require('../../../node_modules/@fortawesome/fontawesome-free/css/all.min.css');

let currentUser: string;
// var department;
var tableDataLength = '';
let departments = [];
let absoluteUrl = '';
let baseUrl = '';

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

    //Retrieve request id
    const urlParams = new URLSearchParams(window.location.search);
    const requestID = urlParams.get('requestid');
    console.log(requestID);
    const contractDetails = await sp.web.lists.getByTitle("Contract_Request").items.select("NameOfAgreement","Company","NameOfRequestor","Owner","TypeOfContract","Party2_agreement","OwnerEmail","Email","ContractStatus").filter(`ID eq ${requestID}`).get();
    const NameOfAgreement = contractDetails[0].NameOfAgreement;
    const companyName = contractDetails[0].Company;
    const NameOfRequestor = contractDetails[0].NameOfRequestor;
    const RequestorEmail = contractDetails[0].Email;
    const Owner = contractDetails[0].Owner;
    const OwnerEmail = contractDetails[0].OwnerEmail;
    const typeOfAgreement = contractDetails[0].TypeOfContract;
    const party2 = contractDetails[0].Party2_agreement;
    const contractStatus = contractDetails[0].ContractStatus;
    console.log('Here', contractDetails);

    var typeOfContract_Acronym = ''
    const contractType = await sp.web.lists.getByTitle('Type of contracts')
      .items.filter(`Identifier eq '${typeOfAgreement}'`)
      // .select('Identifier')
      .getAll();
    console.log(contractType);

    typeOfContract_Acronym = contractType[0].Identifier;

    // Get the current date in YYYY-MM-DD format
    const currentDate = new Date().toISOString().split('T')[0];
    console.log(currentDate);
    const generalFileName = `${currentDate}_${companyName}_${requestID}_${typeOfContract_Acronym}_${party2}`;

    absoluteUrl = this.context.pageContext.web.absoluteUrl;
    baseUrl = absoluteUrl.split('/sites')[0];

    await this.checkCurrentUsersGroupAsync();
    
    //CSS
    this.domElement.innerHTML = `
    <style>

    .container {
      border: 2px solid #dedede;
      background-color: #f1f1f1;
      border-radius: 5px;
      padding: 5px 10px;
      margin: 10px 0;
      display: flex;
      flex-direction: column;
      position: relative;
    }
  
    .darker {
      border-color: #bbb;
      background-color: #ccc;
    }
  
    .container::after {
      content: "";
      clear: both;
      display: table;
    }
  
    .container .user-title-left {
      font-style: italic;
      color: #3870ff;
      align-self: flex-start;
    }
  
    .container .user-title-right {
      font-style: italic;
      color: #3870ff;
      margin-bottom: 5px;
      align-self: flex-end;
    }
  
    .container .comment-text {
      overflow-wrap: break-word;
    }
  
    .container .time-right {
      align-self: flex-end;
      color: #aaa;
    }
  
    .container .time-left {
      align-self: flex-start;
      color: #999;
    }



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
        width: 26%;
    }
    .middle-panel {
        flex: 1;
        width: 60%;
        margin-left: 13%;
        margin-right: 27%;
        transition: width 0.5s ease, margin 0.5s ease;
        position: relative;
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

  #floatingIframeContainer {
    position: fixed;
    top: 10%;
    left: 10%;
    width: 80%;
    height: 80%;
    background: rgba(0, 0, 0, 0.8);
    display: none;
    justify-content: center;
    align-items: center;
    z-index: 9999;
}

#floatingIframe {
    width: 100%;
    height: 100%;
    border: none;
}

#iframeCloseBtn {
    position: absolute;
    top: 10px;
    right: 10px;
    background: #fff;
    border: none;
    padding: 10px;
    cursor: pointer;
    z-index: 10000;
}

.file-input {
  margin: 0px 2rem;
  cursor: pointer;
  background-color: #062470;
  border: none;
  color: #fff;
  padding: 8px 12px;
  font-size: 16px;
  width: 12rem;
  border-radius: 5px;
}

.column {
  flex: 1;
  display: flex;
  flex-direction: column;
  width: 40%;
}

.detail {
}

.detail label {
  font-weight: bold;
  margin-bottom: 0;
}

.detail span {
  margin-left: 5px;
}

#minimizeCommentSection {
  position: fixed;
  top: 50%;
  right: calc(26% + 22px);
  background: none;
  padding: 0px;
  cursor: pointer;
  border: solid #062470;
  border-width: 3px 3px;
  height: 20px;
  margin-left: 2px;
  border-radius: 4px;
  transition: right 0.5s ease;
}

#minimizeCommentSection:hover{
  border: solid #ef7d17;
}

.contract-name-col {
  width: 30%;
}

.column-width-15 {
  width: 15%;
}

.column-width-12 {
  width: 12.5%;
}

.view-col {
  width: 5%;
}

.column-width-8 {
  width: 8%;
}

.contract-details {
  padding: 0rem 1rem;
}

fieldset {
  border: 1px solid #062470;
  overflow: hidden;
}

legend {
  width: auto;
  margin-bottom: 0;
  font-size: 1.2rem;
  color: #062470;
}

.details-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
}



  </style>
    `;
    //HTML
    this.domElement.innerHTML += `
  
        <div class="main-container" id="content">
  
          <div id="nav-placeholder" class="left-panel"></div>
  
          <div id="middle-panel" class="middle-panel">
  
            <button id="minimizeButton"></button>

            <p id="contractStatus" style="color: green; font-size: x-large; position: absolute; top: 0; right: 0; margin: 0.5% 2%;"></p>

            <h2 style="margin-top: 0.7rem; margin-left: 2rem; color: #888; margin-bottom: 0;">Working Area</h2>

            <section class="contract-details">
              <fieldset class="${styles.contractDetailsFS}">
                <legend>CONTRACT ID: ${requestID}</legend>
                <div class="details-grid">
                  <div><strong>Contract Name:</strong> <span>${NameOfAgreement}</span></div>
                  <div><strong>Requestor:</strong> <span>${NameOfRequestor}</span></div>
                  <div><strong>Company:</strong> <span>${companyName}</span></div>
                  <div><strong>Owner:</strong> <span>${Owner}</span></div>
                </div>
              </fieldset>
            </section>
            
            <div id="workingAreaForm" style="width: 100%; padding: 1rem;">

              <fieldset id="tbl_contract" class="${styles.contractContainer, styles.hideDisplay}">
                <legend class="${styles.datatableLegends}">CONTRACT VERSIONS</legend>
              </fieldset>

              <br>
              
              <div id="workingAreaSubmit" style="width: 100%; margin: auto; display: flex; justify-content: center;""></div>

              <br>

              <div id="sharepointSearch" style="display: flex; justify-content: center;"></div>

              <br>

              <fieldset id="contractsDatatableDiv" class="${styles.contractContainer, styles.hideDisplay}">
                <legend class="${styles.datatableLegends}">LegalLink EXISTING FILES</legend>
              </fieldset>
            
            </div>

            <button id="minimizeCommentSection"></button>

          </div>
    
          <div class="right-panel" id="rightPanel">
            <div style="width: 100%; height:100%; background: white; padding-bottom: 30%;">
              <div class="timelineHeader">
                <p style="margin-bottom: 0px;">Comments</p>
              </div>
              <ul id="commentTimeline" class="timeline"></ul>
              <div class="comment-box">
                <textarea id="comment" class="comment-input" placeholder="Add your comment..."></textarea>
                <div style="display: flex;">
                  <button id="addComment">Add Comment</button>
                  <label for="notifyAll" class="form-check-label" style="font-family: Poppins, Arial, sans-serif; display: flex; align-items: center;">
                    <input type="checkbox" id="notifyAll" name="notifyAll" style="transform: scale(1.9); margin-left: 1rem; margin-right: 0.5rem; accent-color: #f07e12;" value="YES">
                    Notify All
                  </label>
                <div>
              </div>
            </div>
          </div>
  
        </div>
    `;

    if(requestID){
      document.getElementById('contractStatus').innerText = `${contractStatus}`;
    }

    $('#tbl_contract').show();

    //Display buttons for working area
    if(departments.includes('Despatcher') || departments.includes('InternalOwner')){
      document.getElementById('workingAreaSubmit').innerHTML = `
        <button type="button" class="file-input" id="useContractTemplate"><i id="useTemplateLoader" class="fa fa-refresh icon" style="display: none;"></i>Use Existing Files</button>
        <button type="button" class="file-input" id="uploadFile">Upload File</button>
        <input type="file" id="uploadContract" style="display: none">
        <button type="button" class="file-input" id="newContractVersion"><i class="fa fa-refresh icon" style="display: none;"></i>New Version</button>
      `;
      // <button type="button" id="approvedByRequestor" style="border-radius: 5px;">Approve as Final Version</button>
    }
    else if(departments.includes('Requestor')){
      document.getElementById('workingAreaSubmit').innerHTML = `
        <button type="button" class="file-input" id="uploadFile">Upload File</button>
        <input type="file" id="uploadContract" style="display: none">
      `;
    }

    //Disable if cancelled
    if(contractStatus === 'Cancelled'){
      const formElements = this.domElement.querySelectorAll('input, select, textarea, button');
      formElements.forEach(element => {
        if (element instanceof HTMLInputElement || element instanceof HTMLSelectElement || 
          element instanceof HTMLTextAreaElement || element instanceof HTMLButtonElement) {
          element.disabled = true;
        }
      });
    }

    //Generate Side Menu
    SideMenuUtils.buildSideMenu(absoluteUrl, departments);

    this.renderRequestDetails(requestID, companyName);
    this.load_comments(requestID);

    //Minimize sidebar
    $('#minimizeButton').on('click', function () {
      const navPlaceholderID = document.getElementById('nav-placeholder');
      const middlePanelID = document.getElementById('middle-panel');
      const minimizeButtonID = document.getElementById('minimizeButton') as HTMLElement;
      const isMinimized = navPlaceholderID.classList.toggle('minimized');

      if (navPlaceholderID && middlePanelID) {
        if (isMinimized) {
          navPlaceholderID.style.width = '0';
          middlePanelID.style.marginLeft = '0%';
          minimizeButtonID.style.left = '0%';
        }
        else {
          navPlaceholderID.style.width = '13%';
          middlePanelID.style.marginLeft = '13%';
          minimizeButtonID.style.left = '13%';
        }
      }
    });

    //Minimize Comment Section
    $('#minimizeCommentSection').on('click', function () {
      const rightPanel = document.getElementById('rightPanel');
      const minimizeButton = document.getElementById('minimizeCommentSection');
      const middlePanelID = document.getElementById('middle-panel');
      const isMinimized = rightPanel.classList.toggle('minimized');
    
      if (isMinimized) {
        rightPanel.style.width = '0';
        rightPanel.style.right = '0';
        middlePanelID.style.marginRight = '0';
        minimizeButton.style.right = '20px';
      } else {
        rightPanel.style.width = '26%';
        rightPanel.style.right = '20px';
        middlePanelID.style.marginRight = '27%';
        minimizeButton.style.right = 'calc(26% + 22px)';
      }
    });

    //Add comment button
    $("#addComment").click(async (e) => {
      console.log("Test New Comment");
      // icon_add_comment.classList.remove('hide');
      // icon_add_comment.classList.add('show');
      // icon_add_comment.classList.add('spinning');

      const currentUserComment = await sp.web.currentUser();
      console.log('Current user here' + currentUserComment);
      let role;
      let commentToUser = '';
      let CommentByName = '';

      if (departments.includes('Despatcher') || departments.includes('InternalOwner')) {
        role = "Owner";
        console.log('Owner Comment');
        commentToUser = RequestorEmail;
        CommentByName = NameOfRequestor;
      }
      else if (departments.includes('Requestor')){
        role = "Requestor";
        console.log('Requestor Comment');
        commentToUser = OwnerEmail;
        CommentByName = Owner;
      }

      const checkboxNotifyAll = document.getElementById('notifyAll') as HTMLInputElement;
      const notifyAll = checkboxNotifyAll.checked ? 'YES' : 'NO';

      const data = {
        Title: requestID,
        RequestID: requestID,
        Comment: $("#comment").val(),
        CommentBy: currentUserComment.UserPrincipalName, // Use Email
        CommentByName: currentUserComment.Title,
        CommentDate: moment().format("DD/MM/YYYY HH:mm"),
        CommentTo: commentToUser,
        NameOfAgreement: NameOfAgreement,
        NotifyAll: notifyAll,
        Role: role
      };

      console.log(data);

      await this.addComment(data);

      // icon_add_comment.classList.remove('spinning', 'show');
      // icon_add_comment.classList.add('hide');

      this.load_comments(requestID);

      $("#comment").val("");

    });

    //Process uploaded file
    $('#uploadContract').on('change', async () => {
      const input = document.getElementById('uploadContract') as HTMLInputElement | null;

      const libraryTitle = "Contracts";
      const library = sp.web.lists.getByTitle(libraryTitle);

      const folder = await library.rootFolder.folders.getByName(companyName).folders.getByName(requestID);
      const files = await folder.files();

      var content_add;

      var file = input.files[0];
      var reader = new FileReader();

      reader.onload = ((file1) => {
        return (e) => {
          console.log(file1.name);
          content_add = e.target.result;
        };
      })(file);

      reader.readAsArrayBuffer(file);

      var nextVersion;
      var filename = '';

      let notRequestor = true;

      if(departments.includes('Requestor') && departments.length === 1){
        const latestFile = getLatestVersionFileRequestor(files);
        nextVersion = latestFile ? getNextVersionRequestor(latestFile.Name) : 1;
        filename = `${generalFileName}_Requestor_V${nextVersion}.docx`;
      }
      else{
        const filenameInput = await prompt("Enter the filename:");
        if(filenameInput && filenameInput.trim() !== ''){
          const latestFile = getLatestVersionFile(files);
          nextVersion = latestFile ? getNextVersion(latestFile.Name) : 1;
          filename = `${generalFileName}_${filenameInput}_V${nextVersion}.docx`;
        }
        else{
          alert('Filename is required.');
          notRequestor = false;
        }
      }

      if(notRequestor){
        console.log(filename);

        const folderPath = `/sites/LegalLink/${libraryTitle}/${companyName}/${requestID}`;

        await this.addFolderToDocumentLibrary(libraryTitle, companyName, requestID)
          .then(async () => {
          try {
            await this.addFileToFolder2(folderPath, filename, content_add, requestID.toString());
          }
          catch (e) {
            console.log(e.message);
          }
        });

        this.renderRequestDetails(requestID, companyName);
        // location.reload();
      }
      else{
        notRequestor = true;
        if (input) {
          input.value = '';
          content_add = null;
        }
      }

      function getLatestVersionFileRequestor(files) {
        let latestFile = null;
        let maxVersion = 0;

        files.forEach(file => {
            const match = file.Name.match(/Requestor_V(\d+)\.docx$/);
            if (match) {
                const version = parseInt(match[1], 10);
                if (version > maxVersion) {
                    maxVersion = version;
                    latestFile = file;
                }
            }
        });

        return latestFile;
      }

      function getNextVersionRequestor(filename) {
          const match = filename.match(/Requestor_V(\d+)\.docx$/);
          if (match) {
              return parseInt(match[1], 10) + 1;
          }
          return 1;
      }

    });

    // //New document button
    // $("#newContractFile").click(async (e) => {
    //   e.preventDefault(); // Prevent the default form submission
  
    //   const libraryTitle = "Contracts";
    //   const library = sp.web.lists.getByTitle(libraryTitle);
  
    //   // Get the current date in YYYY-MM-DD format
    //   const currentDate = new Date().toISOString().split('T')[0];
    //   console.log(currentDate);
  
    //   // Prompt for filename
    //   const filename = await prompt("Enter the filename:");
    //   console.log(filename);

    //   if (filename) {

    //     try {
    //       const folder = await library.rootFolder.folders.getByName(companyName).folders.getByName(requestID);
    //       const files = await folder.files();
    //       const nextVersion = getNextVersion(files);

    //       const fileNameDocx = `${currentDate}_${companyName}_${requestID}_${typeOfAgreement}_${party2}_${filename}_V${nextVersion}.docx`;
    //       await folder.files.add(fileNameDocx, "", false);
    //       alert('File created successfully.');
    //     }
    //     catch (error) {
    //       console.error("Error creating file: ", error);
    //       alert(`Error creating file: ${error.message}`);
    //     }

    //   } else {
    //     alert("Filename is required.");
    //   }
  
    //   function getNextVersion(files) {
    //     let maxVersion = 0;
    
    //     files.forEach(file => {
    //         const match = file.Name.match(/V(\d+)\.docx$/);
    //         if (match) {
    //             const version = parseInt(match[1], 10);
    //             if (version > maxVersion) {
    //                 maxVersion = version;
    //             }
    //         }
    //     });
    
    //     return maxVersion + 1;
    //   }
    // });

    //New Document Version
    $("#newContractVersion").click(async (e) => {
      e.preventDefault(); // Prevent the default form submission
  
      const libraryTitle = "Contracts";
      const library = sp.web.lists.getByTitle(libraryTitle);
  
      // Prompt for filename
      const filename = await prompt("Enter the filename:");
      console.log(filename);
  
      if (filename) {
        try {
          const folder = await library.rootFolder.folders.getByName(companyName).folders.getByName(requestID);
          const files = await folder.files();
          
          const latestFile = getLatestVersionFile(files);
          const nextVersion = latestFile ? getNextVersion(latestFile.Name) : 1;

          const newFileName = `${generalFileName}_${filename}_V${nextVersion}.docx`;
          
          if (latestFile) {
              console.log('Copied');
              // Retrieve the file content as a blob
              const fileContent = await sp.web.getFileByServerRelativeUrl(latestFile.ServerRelativeUrl).getBlob();
              await folder.files.add(newFileName, fileContent, false);
          } else {
            console.log('New file');
            await folder.files.add(newFileName, "", false);
          }
          
          alert('File created successfully.');
        } catch (error) {
            console.error("Error creating file: ", error);
            alert(`Error creating file: ${error.message}`);
        }
        this.renderRequestDetails(requestID, companyName);
        // location.reload();
      } else {
          alert("Filename is required.");
      }
    });

    function getLatestVersionFile(files) {
      let latestFile = null;
      let maxVersion = 0;

      files.forEach(file => {
        const match = file.Name.match(/_V(\d+)\.docx$/);
        if (match && !file.Name.includes('_Requestor_')) {
            const version = parseInt(match[1], 10);
            if (version > maxVersion) {
                maxVersion = version;
                latestFile = file;
            }
        }
      });

      if(latestFile === null){
        return files[0];
      }

      return latestFile;
    }

    function getNextVersion(filename) {
        const match = filename.match(/V(\d+)\.docx$/);
        if (match) {
            return parseInt(match[1], 10) + 1;
        }
        return 1;
    }
  
    //Use template button
    $("#useContractTemplate").click(async (e) => {
      $('#contractsDatatableDiv').show();

      const searchBarHTML = `
        <input type="text" id="searchQuery" style="width: 20rem;" placeholder="Search Existing Files in LegalLink" autocomplete="off">
          <img id="searchButton" src="${absoluteUrl}/SiteAssets/Images/SearchIcon.png" alt="Search" style="cursor: pointer; height: 30px; width: 30px;" />
        <div id="searchResults"></div>
      `;
    
      // Append table to container
      $('#sharepointSearch').html(searchBarHTML);
      
      // Bind SharePoint search to image button
      this.domElement.querySelector('#searchButton').addEventListener('click', () => this.handleSearch(absoluteUrl, companyName, requestID));

      // const useTemplateLoader = document.getElementById('useTemplateLoader');
      // useTemplateLoader.style.display = 'Block';

      // const libraryTitle = "Contracts";
      // const library = sp.web.lists.getByTitle(libraryTitle);

      // await this.getAllDocuments(library).then(dataTable => {
      //   if (dataTable) {
      //     this.buildDataTable(dataTable);
      //     console.log("Documents retrieved successfully.");
      //   } 
      //   else {
      //     console.log("Failed to retrieve documents.");
      //   }
      // });
      
      const libraryName = 'Contracts';

      try {
        const allDocuments = await this.fetchDocumentsFromLibrary(absoluteUrl, libraryName);
        console.log(allDocuments);
        const tableHtml = `
            <table id="contractsDatatable" class="${styles.existingFilesDatatable}">
              <thead>
                <tr>
                  <th class="column-width-12">Company</th>
                  <th class="column-width-8">ID</th>
                  <th class="contract-name-col">Document Name</th>
                  <th class="column-width-12">Created</th>
                  <th class="column-width-12">Last Modified</th>
                  <th class="view-col"></th>
                  <th class="view-col"></th>
                </tr>
              </thead>
              <tbody>
              </tbody>
            </table>
        `;
  
        // Append table to container
        const existingFilesContainer: Element = this.domElement.querySelector('#contractsDatatableDiv');
        const legend = existingFilesContainer.querySelector('legend');
        existingFilesContainer.innerHTML = '';
        existingFilesContainer.appendChild(legend);
        existingFilesContainer.innerHTML += tableHtml;

        console.log('All documents:', allDocuments);
  
        if (allDocuments) {
          // Initialize DataTable
          //to check unique ID TO REMOVE
          $('#contractsDatatable').DataTable({
            data: allDocuments,
            columns: [
                { data: 'Company', className: 'column-width-12' },
                { data: 'Contract', className: 'column-width-8' },
                { data: 'DocumentName', className: 'contract-name-col' },
                { data: 'Created', className: 'column-width-12' },
                { data: 'Modified', className: 'column-width-12' },
                {
                  data: null, className: 'view-col', render: function (data, type, row) {
                      return `
                        <button id="previewExistingFileBtn" data-url="${row.DocumentUrl}" title="Preview Document" class="${styles.datatableBtn}">
                          <img src="${absoluteUrl}/SiteAssets/Images/PreviewDocumentIcon.png" class="${styles.datatableBtnImg}">
                        </button>
                      `;
                  }
                },
                {
                    data: null, className: 'view-col', render: function (data, type, row) {
                        return `
                          <button id="selectExistingFileBtn" data-url="${row.sourceUrl}" title="Select Document" class="${styles.datatableBtn}">
                            <img src="${absoluteUrl}/SiteAssets/Images/SelectDocumentIcon.png" class="${styles.datatableBtnImg}">
                          </button>
                        `;
                    }
                }
            ],
          });
  
          // Attach click event to the preview buttons
          $('#contractsDatatable').on('click', '#previewExistingFileBtn', function () {
            const url = $(this).data('url');
            console.log('iframe url:', url);
            createFloatingIframe(url);
          });
  
          $('#contractsDatatable').on('click', '#selectExistingFileBtn', async (e) => {
            const sourceUrl = $(e.currentTarget).data('url');

            const filenameInput = await prompt("Enter the filename:");

            if(filenameInput && filenameInput.trim() !== ''){
              const libraryTitle = "Contracts";
              const library = sp.web.lists.getByTitle(libraryTitle);

              const folder = await library.rootFolder.folders.getByName(companyName).folders.getByName(requestID);
              const files = await folder.files();

              const latestFile = getLatestVersionFile(files);
              const nextVersion = latestFile ? getNextVersion(latestFile.Name) : 1;
  
              const newFileName = `${generalFileName}_${filenameInput}_V${nextVersion}.docx`;

              console.log(newFileName);
              // Construct the destination URL dynamically
              const destinationFolder = `/sites/LegalLink/${libraryName}/${companyName}/${requestID}`;
              const destinationFileUrl = `${destinationFolder + '/' + newFileName}`;

              // Proceed with your functionality here, e.g., copying the file
              try {
                const file = await sp.web.getFileByServerRelativeUrl(sourceUrl).getBuffer();
                await sp.web.getFolderByServerRelativeUrl(destinationFolder).files.add(destinationFileUrl, file, true);
                console.log('File copied successfully.');
                alert('File copied successfully.');
              } catch (error) {
                console.error('Error copying file: ', error);
                alert('Error copying file.');
              }
            }
            else{
              alert('Filename is required.');
            }
          });

          // $("#searchButton").click(async (e) => {
          //   await this.useExistingContractsSearch(companyName, requestID);
          // });
  
  
        } else {
          console.log("No documents found.");
        }
      }
      catch (error) {
        console.error("Failed to retrieve documents:", error);
      }
  
      function createFloatingIframe(url) {
        let iframeContainer = document.getElementById('floatingIframeContainer');
        if (!iframeContainer) {
          iframeContainer = document.createElement('div');
          iframeContainer.id = 'floatingIframeContainer';
    
          const closeButton = document.createElement('button');
          closeButton.id = 'iframeCloseBtn';
          closeButton.innerText = 'Close';
          closeButton.onclick = () => {
            iframeContainer.style.display = 'none';
            const iframe = document.getElementById('floatingIframe') as HTMLIFrameElement;
            iframe.src = ''; // clear the iframe content
          };
    
          const iframe = document.createElement('iframe');
          iframe.id = 'floatingIframe';
    
          iframeContainer.appendChild(closeButton);
          iframeContainer.appendChild(iframe);
          document.body.appendChild(iframeContainer);
        }
    
        const iframe = document.getElementById('floatingIframe') as HTMLIFrameElement;
        iframe.src = url;
        iframeContainer.style.display = 'flex';
      }

      // useTemplateLoader.style.display = 'None';
    });

    //Upload File button
    $("#uploadFile").click(async (e) => {
      document.getElementById('uploadContract').click();
    });

    $("#approvedByRequestor").click(async (e) => {
      const confirmation = confirm("Are you sure you want to confirm the final version of the contract to proceed with the next step?");
        if (confirmation) 
        {
          const list = sp.web.lists.getByTitle("Contract_Request");
          await list.items.getById(Number(requestID)).update({
            ContractStatus: 'ApprovedByRequestor'
          });
        }
    });

  }

  private async handleSearch(absoluteUrl, companyName, requestID): Promise<void> {
    const query = (this.domElement.querySelector('#searchQuery') as HTMLInputElement).value;
    const libraryName = 'Contracts';

    if (query) {
      const results = await this.searchLibrary(absoluteUrl, query, libraryName);
      console.log("Results Here:", results);
      this.displayResults(absoluteUrl, results, companyName, requestID);
    }
  }

  private async searchLibrary(siteUrl: string, query: string, libraryName: string): Promise<Array<{ Title: string, CreatedDate: string, ModifiedDate: string, sourceUrl: string, documentUrl: string}>> {

    const libraryPath = `/sites/LegalLink/${libraryName}`;

    //wildcard
    const searchQueryUrl = `${absoluteUrl}/_api/search/query?querytext='${query+"*"}'&selectproperties='Title,Path,FileExtension,CreatedOWSDate,CreatedBy,ModifiedOWSDATE,ModifiedBy'&sourceid='%7B368B4FE5-EB91-4554-9225-3AAABD3FF41E%7D'`;

    const response = await this.context.spHttpClient.get(searchQueryUrl, SPHttpClient.configurations.v1);
    const jsonResponse = await response.json();
    console.log("JSON", jsonResponse);

    if (!response.ok) {
      throw new Error('Error fetching search results');
    }

    const results = jsonResponse.PrimaryQueryResult.RelevantResults.Table.Rows.map(row => {
      const title = row.Cells.find(cell => cell.Key === 'Title').Value;
      const fileExtension = row.Cells.find(cell => cell.Key === 'FileExtension').Value;
      const fileName = `${title}.${fileExtension}`;
      const path = row.Cells.find(cell => cell.Key === 'Path').Value;
      const created = row.Cells.find(cell => cell.Key === 'CreatedOWSDate').Value;
      const createdBy = row.Cells.find(cell => cell.Key === 'CreatedBy').Value;
      const lastModified = row.Cells.find(cell => cell.Key === 'ModifiedOWSDATE').Value;
      const modifiedBy = row.Cells.find(cell => cell.Key === 'ModifiedBy').Value;
      const sourceUrl = this.getRelativeUrl(path);
      let documentUrl = `${siteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${sourceUrl}&action=default`;
      if(fileExtension == 'pdf'){
        documentUrl = path;
      }
      return { Title: fileName, CreatedDate: created, ModifiedDate: lastModified, sourceUrl: sourceUrl, documentUrl: documentUrl};
    });

    const filteredResults = results.filter(result => result.sourceUrl.includes(libraryPath));

    return filteredResults;
  }

  private displayResults(absoluteUrl, results: Array<{ Title: string, CreatedDate: string, ModifiedDate: string, sourceUrl: string, documentUrl: string}>, companyName, requestID): void {
    console.log(results);
    $('#contractsDatatable tbody').empty();

    const formattedResults = results.map(item => ({
      Company: companyName,
      Contract: requestID,
      DocumentName: item.Title,
      Created: this.formatDateToUK(item.CreatedDate),
      // CreatedBy: item.CreatedBy,
      Modified: this.formatDateToUK(item.ModifiedDate),
      // ModifiedBy: item.ModifiedBy,
      DocumentUrl: item.documentUrl,
      sourceUrl: item.sourceUrl
    }));

    console.log('Formatted results:', formattedResults);

    $('#contractsDatatable').DataTable({
      destroy: true,
      data: formattedResults,
      columns: [
        { data: 'Company' },
        { data: 'Contract' },
        { data: 'DocumentName' },
        { data: 'Created' },
        // { data: 'CreatedBy' },
        { data: 'Modified' },
        // { data: 'ModifiedBy' },
        {
          data: null, className: 'view-col', render: function (data, type, row) {
              return `
                <button id="previewExistingFileBtn" data-url="${row.DocumentUrl}" title="Preview Document" class="${styles.datatableBtn}">
                  <img src="${absoluteUrl}/SiteAssets/Images/PreviewDocumentIcon.png" class="${styles.datatableBtnImg}">
                </button>
              `;
          }
        },
        {
            data: null, className: 'view-col', render: function (data, type, row) {
                return `
                  <button id="selectExistingFileBtn" data-url="${row.sourceUrl}" title="Select Document" class="${styles.datatableBtn}">
                    <img src="${absoluteUrl}/SiteAssets/Images/SelectDocumentIcon.png" class="${styles.datatableBtnImg}">
                  </button>
                `;
            }
        }
      ]
    });
  }

  private getRelativeUrl(fullUrl: string): string {
    const baseUrl = this.context.pageContext.web.absoluteUrl.replace(this.context.pageContext.web.serverRelativeUrl, '');
    return fullUrl.replace(baseUrl, '');
  }

  public async addFolderToDocumentLibrary(libraryTitle, companyFolderName, contractFolderName) {
    const library = sp.web.lists.getByTitle(libraryTitle);

    try {
      const exists = await this.folderExists(library, companyFolderName, contractFolderName);

      //None exists
      if (exists === "noneExist") {
        //Create company folder
        await library.rootFolder.folders.add(companyFolderName);
        console.log(`Company Folder '${companyFolderName}' created successfully.`);
        //Create contract folder
        await library.rootFolder.folders.getByName(companyFolderName).folders.add(contractFolderName);
        console.log(`Contract Folder '${contractFolderName}' created successfully.`);
      }
      else if (exists === "companyOnly") {
        //Create contract folder
        await library.rootFolder.folders.getByName(companyFolderName).folders.add(contractFolderName);
        console.log(`Contract Folder '${contractFolderName}' created successfully.`);
      }
      else if (exists === "allExist") {
        console.log(`All folders already exist.`);
      }

    }
    catch (error) {
      console.error(`Error creating folder: ${error.message}`);
    }

    // try {
    //   console.log(1);
    //   //Check existence of company folder
    //   const exists = await this.folderExists(libraryTitle, companyFolderName, contractFolderName);

    //   if(exists == 'allExist'){
    //     console.log(9);
    //     console.log(`All folders exist.`);
    //   }
    //   else {
    //     console.log(10);
    //     if(exists == 'noneExist'){
    //       // Create a new company folder
    //       const library = sp.web.lists.getByTitle(libraryTitle);
    //       await library.rootFolder.folders.add(companyFolderName);
    //       console.log(`Company Folder '${companyFolderName}' created successfully.`);
    //     }
    //  console.log(`Contract Folder '${contractFolderName}'`);
    // const library = sp.web.lists.getByTitle(libraryTitle);
    // await library.rootFolder.folders.add(contractFolderName);
    // console.log(`Contract Folder '${contractFolderName}' created successfully.`);
    // }

    // Get the document library by title

    // } catch (error) {
    //   console.log(11);
    //   console.error(`Error creating folder: ${error.message}`);
    // }
  }

  public async folderExists(library, companyFolderName, contractFolderName) {

    let existResponse = "";

    // Check if company folder exists
    try {
      const companyFolder = await library.rootFolder.folders.getByName(companyFolderName).select("Exists").get();
      console.log("Company folder exists");
      //Company folder exists
      if (companyFolder.Exists) {
        try {
          const contractFolder = await library.rootFolder.folders.getByName(companyFolderName).folders.getByName(contractFolderName).select("Exists").get();
          if (contractFolder.Exists) {
            console.log("Contract folder exists");
            existResponse = "allExist";
            return existResponse;
          }
        }
        catch (error) {
          console.log(error);
          console.log("Contract folder does not exist");
          existResponse = "companyOnly";
          return existResponse;
        }
      }
    }
    catch (error) {
      //Company folder does not exist
      console.log(error);
      console.log("Company folder does not exist");
      existResponse = "noneExist";
      return existResponse;
    }

  }

  public async fetchDocumentsFromLibrary(siteUrl, libraryName) {
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${libraryName}')/items?$expand=File&$select=ID,File/Name,File/ServerRelativeUrl,File/Title,Modified,Created&$top=500`;

    const response = await fetch(endpoint, {
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=verbose',
      },
      credentials: 'include'
    });

    if (!response.ok) {
      throw new Error(`Error fetching documents: ${response.statusText}`);
    }

    const data = await response.json();
    console.log("Here", data);
    const contractFiles = data.d.results.filter(item => item.File);

    return contractFiles.map(item => {
      if (item.File && item.File.Name && item.File.ServerRelativeUrl) {
        const sourceUrl = item.File.ServerRelativeUrl;
        const fileUrlParts = item.File.ServerRelativeUrl.split('/');
        const companyFolder = fileUrlParts[fileUrlParts.length - 3];
        const contractFolder = fileUrlParts[fileUrlParts.length - 2];
        const redirectUrl = `${siteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${item.File.ServerRelativeUrl}&action=default`;

        return {
          Company: companyFolder,
          Contract: contractFolder,
          DocumentName: item.File.Name,
          Created: this.formatDateToUK(item.Created),
          Modified: this.formatDateToUK(item.Modified),
          DocumentUrl: redirectUrl,
          sourceUrl: sourceUrl
        };
      }
      else {
        // console.warn('Skipping folder:', item);
        return null;
      }
    }).filter(Boolean);

  }

  public formatDateToUK(dateString: string): string {
    const date = new Date(dateString);
    return date.toLocaleDateString('en-GB', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
    });
  }

  public async addFileToFolder2(folderPath, fileName, fileContent, requestId) {
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

  // async folderExists(libraryTitle, folderName) {
  //   try {
  //     // Initialize the PnP JS Library
  //     // Get the document library by title
  //     const library = sp.web.lists.getByTitle(libraryTitle);

  //     // Check if the folder exists
  //     const folder = await library.rootFolder.folders.getByName(folderName).select("Exists").get();

  //     return folder.Exists;
  //   }
  //   catch (error) {
  //     console.error(`Error checking folder existence: ${error.message}`);
  //     return false;
  //   }
  // }
  
  private async renderRequestDetails(id: any, companyName: string) {
    const contractVersionText = document.getElementById('newContractVersion');
    if(contractVersionText){
      contractVersionText.innerText = 'New Document';
    }
    this.getFileDetailsByFilter('Contracts', id, companyName)
        .then((fileDetailsArray) => {
            if (fileDetailsArray && fileDetailsArray.length > 0) {
              if(contractVersionText){
                contractVersionText.innerText = 'New Version';
              }
              tableDataLength = fileDetailsArray.length;
                console.log("File details:", fileDetailsArray);

                const tableHtml = `
                    <table id="tableContracts" class="${styles.contractVersionsTable}">
                        <thead>
                            <tr>
                                <th class="th-lg contract-name-col" scope="col">ContractName</th>
                                <th class="column-width-15" scope="col">CreatedAt</th>
                                <th class="column-width-15" scope="col">LastModifiedBy</th>
                                <th class="column-width-15" scope="col">LastModifiedAt</th>
                                <th class="column-width-15" scope="col">UploadedBy</th>
                                <th scope="col"></th>
                            </tr>
                        </thead>
                        <tbody>
                        </tbody>
                    </table>
                `;

                const listContainer: Element = this.domElement.querySelector('#tbl_contract');
                const legend: Element = listContainer.querySelector('legend');
                listContainer.innerHTML = '';
                listContainer.appendChild(legend);
                // const tableHtmlResert: Element = this.domElement.querySelector('#tableContracts');
                // if(tableHtmlResert){
                //   tableHtmlResert.innerHTML = '';
                // }
                listContainer.innerHTML += tableHtml;

                // let requestorFlag = false;

                const tableData = fileDetailsArray.map(fileItem => {
                    const formattedTimeCreated = new Date(fileItem.TimeCreated).toLocaleDateString('en-GB');
                    const unformattedLastModified = new Date(fileItem.TimeLastModified);
                    const day = ('0' + unformattedLastModified.getDate()).slice(-2);
                    const month = ('0' + (unformattedLastModified.getMonth() + 1)).slice(-2);
                    const year = unformattedLastModified.getFullYear();
                    const hours = ('0' + unformattedLastModified.getHours()).slice(-2);
                    const minutes = ('0' + unformattedLastModified.getMinutes()).slice(-2);
                    const seconds = ('0' + unformattedLastModified.getSeconds()).slice(-2);
                    const formattedTimeLastModified = `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;

                    return {
                        Name: fileItem.Name || 'N/A',
                        CreatedAt: formattedTimeCreated || 'N/A',
                        ModifiedBy: fileItem.ModifiedBy?.Title || 'N/A',
                        LastModifiedAt: formattedTimeLastModified || 'N/A',
                        UploadedBy: fileItem.Author?.Title || 'N/A',
                        UniqueId: fileItem.UniqueId,
                        Url: `${baseUrl+fileItem.ServerRelativeUrl}`,
                    };
                });

                $('#tableContracts').DataTable({
                  data: tableData,
                  columns: [
                    { data: 'Name', className: 'contract-name-col' },
                    { data: 'CreatedAt', className: 'column-width-15' },
                    { data: 'ModifiedBy', className: 'column-width-15' },
                    { data: 'LastModifiedAt', className: 'column-width-15' },
                    { data: 'UploadedBy', className: 'column-width-15' },
                    {
                      data: null, className: 'view-col', render: function (data, type, row) {
                        // if (requestorFlag && ((departments.length === 1) && departments.includes('Requestor'))) {
                        //     return '<td class="view-col"></td>';
                        // } 
                        // if (((departments.length === 1) && departments.includes('Requestor'))) {
                        //   return '<td class="view-col"></td>';
                        // } 
                        // else {
                          // requestorFlag = true;
                          return `
                              <ul class="list-inline m-0" style="display: grid; align-items: center;">
                                  <li class="list-inline-item">
                                      <button id="btn_view_${row.UniqueId}" class="btn btn-secondary btn-sm rounded-circle" type="button" data-toggle="tooltip" data-placement="top" title="View" style="display: none;">
                                          <i class="fas fa-eye"></i>
                                      </button>
                                  </li>
                                  <li class="list-inline-item">
                                      <button id="modalActivate_${row.UniqueId}" class="btn btn-secondary btn-sm rounded-circle" type="button" data-toggle="modal" data-target="#exampleModalPreview" style="display: block; width: auto;">
                                          <i class="fas fa-eye"></i>
                                      </button>
                                  </li>
                              </ul>
                          `;
                        // }
                      }
                    }
                  ],
                  order: [],
                  pageLength: -1
                });

                tableData.forEach(fileDetails => {
                  console.log(fileDetails);
                  if((departments.length === 1) && departments.includes('Requestor')){
                    $(`#modalActivate_${fileDetails.UniqueId}`).click(() => {
                      window.open(`${fileDetails.Url}`);
                    });
                  }
                  else{
                    $(`#modalActivate_${fileDetails.UniqueId}`).click(() => {
                      const extension = fileDetails.Name.split('.').pop().toLowerCase();
                      if (extension === 'pdf') {
                          window.open(`${fileDetails.Url}?web=1`, '_blank');
                      } else if (extension === 'docx') {
                        window.open(`ms-word:ofv|u|${fileDetails.Url}`, '_blank');
                      }
                    });
                  }
                });
            } else {
                console.log("No items found.");
            }
        })
        .catch(error => {
            console.error("Error retrieving file details:", error);
        });
  }

  public async checkCurrentUsersGroupAsync() {
    // var currentRole;
    let groupList = await sp.web.currentUser.groups();
    console.log('grouplist: ', groupList);
  
    // const urlParams = new URLSearchParams(window.location.search);
    // const updateRequestID = urlParams.get('requestid');
    
    if (groupList.filter(g => g.Title == sharepointConfig.Groups.Requestor).length == 1) {
      departments.push("Requestor");
    }
    if (groupList.filter(g => g.Title == sharepointConfig.Groups.InternalOwner).length == 1) {
      departments.push("InternalOwner");
    }
    if (groupList.filter(g => g.Title == sharepointConfig.Groups.ExternalOwner).length == 1) {
      departments.push("ExternalOwner");
    }
    if (groupList.filter(g => g.Title == sharepointConfig.Groups.Despatcher).length == 1) {
      departments.push("Despatcher");
    }
    if (groupList.filter(g => g.Title == sharepointConfig.Groups.DirectorsView).length == 1) {
      departments.push("DirectorsView");
    }

    console.log(departments);

    // if (departments.length === 0) {
    //   departments.push("noGroup");
    // }
    // else if(departments.length === 1) {
    //   if (departments.includes('Requestor')) {
    //     if (!updateRequestID){
    //       return currentRole = 'RequestorCreate'; //New Request
    //     }
    //     else{
    //       return currentRole = 'RequestorUpdate'; //Update Request
    //     }
    //   }
    //   else if (departments.includes('ExternalOwner')) {
    //     return currentRole = 'ExternalOwnerOnly' //External Owner Only -> Disable Submit Button
    //   }
    // }
    // else if(departments.length === 2){
    //   if (departments.includes('Requestor') && (departments.includes('InternalOwner') || departments.includes('ExternalOwner') || (departments.includes('DirectorsView')))) {
    //     if (!updateRequestID){
    //       if(departments.includes('DirectorsView')){
    //         return currentRole = 'RequestorCreate'; //New Request by Director's View
    //       }
    //       else{
    //         return currentRole = 'OwnerCreate'; //New Request by Internal Owner or External Owner on behalf of requestor or for themselves
    //       }
    //     }
    //     else {
    //       return currentRole = 'OwnerView'; //Internal Owner or External Owner
    //     }
    //   }
    // }
    // else if(departments.length === 3){
    //   if (departments.includes('Requestor') && departments.includes('InternalOwner') && departments.includes('Despatcher')){
    //     if (!updateRequestID){
    //       return currentRole = 'DespatcherCreate'; //New Request by despatcher on behalf of requestor
    //     }
    //     else{
    //       return currentRole = 'DespatcherAssign'; //Despatcher edit and assign
    //     }
    //   }
    // }
  }

  //Load timeline comments
  public async load_comments(updateRequestID) {
    const timeline = document.getElementById('commentTimeline');
    timeline.innerHTML = '';
  
    const CommentList = await sp.web.lists.getByTitle("Comments").items.select("RequestID,Comment,CommentBy,CommentDate,CommentByName").filter(`RequestID eq '${updateRequestID}'`).get();
    console.log('Commentlist', CommentList);
  
    // const commentUsers: any[] = await sp.web.siteUsers();
  
    // Get current user
    const currentUser = await sp.web.currentUser();
    console.log("Comments  Current:", currentUser);
    const currentUserTitle = currentUser.Title;
    console.log("Comments  Current Name:", currentUserTitle);
  
    CommentList.forEach(item => {
      const comment = item.Comment;
      let formattedCommentDate = '';
  
      if (item.CommentDate) {
        const parts = item.CommentDate.split(/[\/\s:]/);
        if (parts.length >= 5) {
          const day = parseInt(parts[0], 10);
          const month = parseInt(parts[1], 10) - 1; // months are 0-based in JavaScript
          const year = parseInt(parts[2], 10);
          const hours = parseInt(parts[3], 10);
          const minutes = parseInt(parts[4], 10);
          const commentDate = new Date(year, month, day, hours, minutes);
          if (!isNaN(commentDate.getTime())) {
            formattedCommentDate = `${('0' + commentDate.getDate()).slice(-2)}/${('0' + (commentDate.getMonth() + 1)).slice(-2)}/${commentDate.getFullYear()} ${('0' + commentDate.getHours()).slice(-2)}:${('0' + commentDate.getMinutes()).slice(-2)}`;
          }
        }
      }
  
      // let userEmail = item.CommentBy;
      // let userTitle = '';
      // users.forEach(user => {
      //   if (user.Email === userEmail) {
      //     userTitle = user.Title;
      //     return;
      //   }
      // });
  
      const isCurrentUser = item.CommentByName === currentUserTitle ;
      const containerClass = isCurrentUser ? 'container darker' : 'container';
      const timeClass = isCurrentUser ? 'time-left' : 'time-right';
      const userTitleClass = isCurrentUser ? 'user-title-right' : 'user-title-left';
  
      const timelineItem = document.createElement('l');
      timelineItem.className = 'timeline-item';
      timelineItem.innerHTML = `
        <div class="${containerClass}">
          <div class="${userTitleClass}">#${item.CommentByName}</div>
          <div class="comment-text">${comment}</div>
          <span class="${timeClass}">${formattedCommentDate}</span>
        </div>
      `;
      timeline.appendChild(timelineItem);
    });
  
    timeline.scrollTop = timeline.scrollHeight;
  }

  public async addComment(data) {
    try {
      const iar = await sp.web.lists.getByTitle("Comments").items.add(data);

      alert("Comment added succesfully.");
    }
    catch (e) {
      alert("An error occured." + e.message);
    }
  }

  public async getFileDetailsByFilter(libraryName, reqId, companyName) {
    try {
      let folderPath = libraryName + "/" + companyName + "/" + reqId;
      let requestUrl = `${absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderPath}')/Files?$orderby=TimeCreated desc&$expand=Author,ModifiedBy`;
      console.log('RequestURl: ', requestUrl);
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
