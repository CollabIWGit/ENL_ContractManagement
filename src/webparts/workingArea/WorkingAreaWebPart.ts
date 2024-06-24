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

    //Retrieve request id
    const urlParams = new URLSearchParams(window.location.search);
    const requestID = urlParams.get('requestid');
    const contractDetails = await sp.web.lists.getByTitle("Contract_Request").items.select("NameOfAgreement","Company","NameOfRequestor","Owner").filter(`ID eq ${requestID}`).get();
    const NameOfAgreement = contractDetails[0].NameOfAgreement;
    const companyName = contractDetails[0].Company;
    const NameOfRequestor = contractDetails[0].NameOfRequestor;
    const Owner = contractDetails[0].Owner;

    const absoluteUrl = this.context.pageContext.web.absoluteUrl;

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
      border-color: #ccc;
      background-color: #ddd;
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

#tableContracts tbody tr td {
  padding: 3px;
}

.contract-details {
  display: flex;
  justify-content: space-between;
  background-color: #f9f9f9;
  border: 1px solid #ddd;
  border-radius: 5px;
  padding: 5px 20px;
  width: 50rem;
  margin: 0 auto;
  position: relative;
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

.table-responsive {
  max-height: 305px; /* Adjust as needed */
  overflow: hidden;
}

#tableContracts {
  width: 100%;
  border-collapse: collapse; /* Ensure borders do not collapse */
  table-layout: fixed; /* Ensure the table layout is fixed to align columns properly */
}

thead th {
  background-color: #f2f2f2;
  position: sticky;
  top: 0;
  z-index: 1; /* Ensure the header is above the body */
  border-left: none;
  border-right: none;
  border-top: none;
}

th, td {
  border-left: none;
  border-right: none;
  border-top: none;
  border-bottom: 1px solid #ddd;
  padding: 8px;
  word-wrap: break-word;
  overflow-wrap: break-word;
  white-space: normal; /* Allow text to wrap */
}

tbody {
  display: block;
  max-height: 250px; /* Adjust the height as needed */
  overflow-y: auto;
}

thead, tbody tr {
  display: table;
  width: 100%;
  table-layout: fixed;
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
  width: 10%;
}

.column-width-8 {
  width: 8%;
}

#contractsDatatable tbody tr td,
#contractsDatatable thead tr th {
  text-align: center;
}
#contractsDatatable thead {
  padding-right: 12px;
}



  </style>
    `;
    //HTML
    this.domElement.innerHTML += `
  
        <div class="main-container" id="content">
  
          <div id="nav-placeholder" class="left-panel"></div>
  
          <div id="middle-panel" class="middle-panel">

            <h2 style="margin-top: 0.7rem; margin-left: 2rem;">Working Area</h2>
  
            <button id="minimizeButton"></button>
      
            <div class="contract-details">

              <p id="contractStatus" style="color: green; position: absolute; top: 0; right: 0; margin: 0.5% 2%;">Status: In Progress</p>

              <div class="column">
                <div class="detail">
                  <label>Contract Name:</label>
                  <span>${NameOfAgreement}</span>
                </div>
                <div class="detail">
                  <label>Company:</label>
                  <span>${companyName}</span>
                </div>
              </div>
              <div class="column">
                <div class="detail">
                  <label>Requestor:</label>
                  <span>${NameOfRequestor}</span>
                </div>
                <div class="detail">
                  <label>Owner:</label>
                  <span>${Owner}</span>
                </div>
              </div>

            </div>
          
            
            <div id="workingAreaForm" style="width: 100%; padding: 2%">

              <div class="col-md-12 table-responsive"  style="border-bottom: 2px solid;">
                  <div id="tbl_contract"></div>
              </div>

              <br>
              
              <div id="workingAreaSubmit" style="width: 100%; margin: auto; display: flex; justify-content: center;""></div>

              <br>

              <div id="sharepointSearch" style="display: flex; justify-content: center;"></div>

              <br>

              <div id="contractsDatatableDiv"></div>
            
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
                <button id="addComment">Add Comment</button>
              </div>
            </div>
          </div>
  
        </div>
    `;

    //Display buttons for working area
    if(department !== "Requestor"){
      document.getElementById('workingAreaSubmit').innerHTML = `
        <button type="button" class="file-input" id="newContractFile"><i class="fa fa-refresh icon" style="display: none;"></i>New Document</button>
        <button type="button" class="file-input" id="useContractTemplate"><i id="useTemplateLoader" class="fa fa-refresh icon" style="display: none;"></i>Use Existing Files</button>
        <button type="button" class="file-input" id="uploadFile">Upload File</button>
        <input type="file" id="uploadContract" style="display: none">
      `;
    }
    else {
      document.getElementById('workingAreaSubmit').innerHTML = `
        <button type="button" class="file-input" id="uploadFile">Upload File</button>
        <input type="file" id="uploadContract" style="display: none">
      `;
    }

    //Generate Side Menu
    SideMenuUtils.buildSideMenu(this.context.pageContext.web.absoluteUrl);

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

      const currentUser = await sp.web.currentUser();
      let role;

      if (department === "Requestor") {

        role = "Requestor";

      }
      else if (department === "Owner") {
        role = "Owner";
      }
      else {
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

    //Process uploaded file
    $('#uploadContract').on('change', async () => {
      const input = document.getElementById('uploadContract') as HTMLInputElement | null;

      var filename_add;
      var content_add;

      var file = input.files[0];
      var reader = new FileReader();

      reader.onload = ((file1) => {
        return (e) => {
          console.log(file1.name);

          filename_add = file1.name,
            content_add = e.target.result;

        };
      })(file);

      reader.readAsArrayBuffer(file);

      const library = "Contracts_ToReview";
      const folderPath = `/sites/ContractMgt/Contracts_ToReview/${companyName}/${requestID}`;

      await this.addFolderToDocumentLibrary(library, companyName, requestID)
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

    //New document button
    $("#newContractFile").click(async (e) => {
      const libraryTitle = "Contracts_ToReview";
      const library = sp.web.lists.getByTitle(libraryTitle);
      // Prompt for filename
      const filename = await prompt("Enter the filename:");
      console.log(filename);
      const fileNameDocx = filename + '.docx';

      try {
        await library.rootFolder.folders.getByName(companyName).folders.getByName(requestID).files.add(fileNameDocx, false);
        console.log("File created successfully");
      }
      catch (error) {
        console.error(error);
      }
    });

    //Use template button
    $("#useContractTemplate").click(async (e) => {

      const searchBarHTML = `
      <input type="text" id="searchQuery" style="width: 20rem;" placeholder="Search Existing Files" />
      <img id="searchButton" src="${absoluteUrl}/Site%20Assets/SearchIcon.png" alt="Search" style="cursor: pointer; height: 30px; width: 30px;" />
      <div id="searchResults"></div>
    `;
    
    // Append table to container
    $('#sharepointSearch').html(searchBarHTML);
    
    // Bind SharePoint search to image button
    this.domElement.querySelector('#searchButton').addEventListener('click', () => this.handleSearch(siteUrl, companyName, requestID));

      // const useTemplateLoader = document.getElementById('useTemplateLoader');
      // useTemplateLoader.style.display = 'Block';

      // const libraryTitle = "Contracts_ToReview";
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
      
      const siteUrl = 'https://frcidevtest.sharepoint.com/sites/ContractMgt';
      const libraryName = 'Contracts_ToReview';

      try {
        const allDocuments = await this.fetchDocumentsFromLibrary(siteUrl, libraryName);
        console.log(allDocuments);
        const tableHtml = `
            <table id="contractsDatatable" class="${styles.table}">
              <thead>
                <tr>
                  <th class="column-width-12">Company</th>
                  <th class="column-width-12">Contract</th>
                  <th class="contract-name-col">Document Name</th>
                  <th class="column-width-12">Created</th>
                  <th class="column-width-12">Last Modified</th>
                  <th class="view-col">Preview</th>
                  <th class="column-width-8">Select</th>
                </tr>
              </thead>
              <tbody>
              </tbody>
            </table>
        `;
  
        // Append table to container
        $('#contractsDatatableDiv').html(tableHtml);

        console.log('All documents:', allDocuments);
  
        if (allDocuments) {
          // Initialize DataTable
          //to check unique ID TO REMOVE
          $('#contractsDatatable').DataTable({
            data: allDocuments,
            columns: [
                { data: 'Company', className: 'column-width-12' },
                { data: 'Contract', className: 'column-width-12' },
                { data: 'DocumentName', className: 'contract-name-col' },
                { data: 'Created', className: 'column-width-12' },
                { data: 'Modified', className: 'column-width-12' },
                {
                    data: null, className: 'view-col', render: function (data, type, row) {
                        return `<button class="preview-btn" data-url="${row.DocumentUrl}">Preview</button>`;
                    }
                },
                {
                    data: null, className: 'column-width-8', render: function (data, type, row) {
                        return `<button class="select-btn" data-url="${row.sourceUrl}">Select</button>`;
                    }
                }
            ],
          });
  
          // Attach click event to the preview buttons
          $('#contractsDatatable').on('click', '.preview-btn', function () {
            const url = $(this).data('url');
            console.log('iframe url:', url);
            createFloatingIframe(url);
          });
  
          $('#contractsDatatable').on('click', '.select-btn', async (e) => {
            const sourceUrl = $(e.currentTarget).data('url');
  
            const filename = await prompt("Enter the filename:");
            console.log(filename);
            const fileNameDocx = filename + '.docx';
  
            // Construct the destination URL dynamically
            const destinationFolder = `/sites/ContractMgt/${libraryName}/${companyName}/${requestID}`;
            const destinationFileUrl = `${destinationFolder + '/' + fileNameDocx}`;
  
            await this.addFolderToDocumentLibrary(libraryName, companyName, requestID)
              .then(async () => {
                try {
                  //Gives error if file name already exists
                  await sp.web.getFileByServerRelativePath(sourceUrl).copyTo(destinationFileUrl, false);
                  console.log(`File copied successfully to ${destinationFileUrl}`);
                  window.open(`ms-word:ofv|u|https://frcidevtest.sharepoint.com/${destinationFileUrl}`, '_blank');
                }
                catch (e) {
                  console.log(e.message);
                }
              });
  
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

  }

  private async handleSearch(siteUrl, companyName, requestID): Promise<void> {
    const query = (this.domElement.querySelector('#searchQuery') as HTMLInputElement).value;
    const libraryName = 'Contracts_ToReview';

    if (query) {
      const results = await this.searchLibrary(siteUrl, query, libraryName);
      console.log("Results Here:", results);
      this.displayResults(results, companyName, requestID);
    }
  }

  private async searchLibrary(siteUrl: string, query: string, libraryName: string): Promise<Array<{ Title: string, CreatedDate: string, ModifiedDate: string, sourceUrl: string, documentUrl: string}>> {

    const libraryPath = `/sites/ContractMgt/${libraryName}`;

    //wildcard
    const searchQueryUrl = `${this.context.pageContext.web.absoluteUrl}/_api/search/query?querytext='${query+"*"}'&selectproperties='Title,Path,FileExtension,CreatedOWSDate,CreatedBy,ModifiedOWSDATE,ModifiedBy'&sourceid='%7B368B4FE5-EB91-4554-9225-3AAABD3FF41E%7D'`;

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

  private displayResults(results: Array<{ Title: string, CreatedDate: string, ModifiedDate: string, sourceUrl: string, documentUrl: string}>, companyName, requestID): void {
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
            data: 'DocumentUrl', render: function (data, type, row) {
                return `<button class="preview-btn" data-url="${row.DocumentUrl}">Preview</button>`;
            }
        },
        {
            data: 'sourceUrl', render: function (data, type, row) {
                return `<button class="select-btn" data-url="${row.sourceUrl}">Select</button>`;
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

    $("#tbl_contract").html('');

    // $("#section_review_contract").css("display", "block");

    this.getFileDetailsByFilter('Contracts_ToReview', id, companyName)
      .then((fileDetailsArray) => {
        if (fileDetailsArray && fileDetailsArray.length > 0) {
          console.log("File details:", fileDetailsArray);

          let html: string = '';
          
          html += `
                    <table id="tableContracts" class="table">
                      <thead class="thead-dark">
                        <tr>
                          <th class="th-lg contract-name-col" scope="col">Contract Name</th>
                          <th class="column-width-15" scope="col">Created At</th>
                          <th class="column-width-15" scope="col">Last modified By</th>
                          <th class="column-width-15" scope="col">Last modified At</th>
                          <th class="column-width-15" scope="col">Uploaded By</th>

                          <th class="view-col" scope="col">View</th>
                        </tr>
                      </thead>
                      <tbody class="table-body">
        `;

          let requestorFlag = false;

          fileDetailsArray.forEach(fileItem => {
            const formattedTimeCreated = new Date(fileItem.TimeCreated).toLocaleDateString('en-GB');
            // const formattedTimeLastModified = new Date(fileItem.TimeLastModified).toLocaleDateString('en-GB');
            const unformattedLastModified = new Date(fileItem.TimeLastModified);
            const day = ('0' + unformattedLastModified.getDate()).slice(-2);
            const month = ('0' + (unformattedLastModified.getMonth() + 1)).slice(-2);
            const year = unformattedLastModified.getFullYear();
            const hours = ('0' + unformattedLastModified.getHours()).slice(-2);
            const minutes = ('0' + unformattedLastModified.getMinutes()).slice(-2);
            const seconds = ('0' + unformattedLastModified.getSeconds()).slice(-2);

            const formattedTimeLastModified = `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;


            html += `
                <tr>
                    <td class="contract-name-col" scope="row">${fileItem.Name}</td>
                    <td class="column-width-15" scope="row">${formattedTimeCreated}</td>
                    <td class="column-width-15" scope="row">${fileItem.ModifiedBy.Title}</td>
                    <td class="column-width-15" scope="row">${formattedTimeLastModified}</td>
                    <td class="column-width-15" scope="row">${fileItem.Author.Title}</td>
                    `;
                    if (department !== "Requestor" || !requestorFlag) {
                    html+=`
                    <td class="column-width-8">
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
            }
            else{
              html+=`
                    <td style="width: 8%;"></td>
            `;
            }
            requestorFlag = true;
          });

          html += `
                        </tbody>
                    </table>
        `;

          const listContainer: Element = this.domElement.querySelector('#tbl_contract');
          listContainer.innerHTML = html;

          fileDetailsArray.forEach(fileDetails => {
            if(department !== "Requestor"){
              $(`#modalActivate_${fileDetails.UniqueId}`).click(() => {
                window.open(`ms-word:ofv|u|https://frcidevtest.sharepoint.com/${fileDetails.ServerRelativeUrl}`, '_blank');
              });
            }
            else{
              $(`#modalActivate_${fileDetails.UniqueId}`).click(() => {
                window.open(`https://frcidevtest.sharepoint.com/${fileDetails.ServerRelativeUrl}`);
              });
            }
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

  //Load timeline comments
  public async load_comments(updateRequestID) {
    const timeline = document.getElementById('commentTimeline');
    timeline.innerHTML = '';
  
    const CommentList = await sp.web.lists.getByTitle("Comments").items.select("RequestID,Comment,CommentBy,CommentDate").filter(`RequestID eq '${updateRequestID}'`).get();
    console.log('Commentlist', CommentList);
  
    const users: any[] = await sp.web.siteUsers();
  
    // Get current user
    const currentUser = await sp.web.currentUser();
    console.log("Current:", currentUser);
    const currentUserTitle = currentUser.Title;
  
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
  
      let userEmail = item.CommentBy;
      let userTitle = '';
      users.forEach(user => {
        if (user.Email === userEmail) {
          userTitle = user.Title;
          return;
        }
      });
  
      const isCurrentUser = userTitle === currentUserTitle;
      const containerClass = isCurrentUser ? 'container darker' : 'container';
      const timeClass = isCurrentUser ? 'time-left' : 'time-right';
      const userTitleClass = isCurrentUser ? 'user-title-right' : 'user-title-left';
  
      const timelineItem = document.createElement('li');
      timelineItem.className = 'timeline-item';
      timelineItem.innerHTML = `
        <div class="${containerClass}">
          <div class="${userTitleClass}">#${userTitle}</div>
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
      let currentWebUrl = this.context.pageContext.web.absoluteUrl;
      let requestUrl = `${currentWebUrl}/_api/web/GetFolderByServerRelativeUrl('${folderPath}')/Files?$orderby=TimeCreated desc&$expand=Author,ModifiedBy`;

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
