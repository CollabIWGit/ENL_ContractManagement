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
import styles from './RequestFormWebPart.module.scss';
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
import { GraphRequest } from '@microsoft/microsoft-graph-client';

let SideMenuUtils = new sideMenuUtils();

SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/4.1.0/css/bootstrap.min.css');
SPComponentLoader.loadCss('https://cdn.datatables.net/1.10.25/css/jquery.dataTables.min.css');

// require('../../Assets/scripts/styles/mainstyles.css');
require('./../../common/scss/style.scss');
require('./../../common/css/style.css');
require('./../../common/css/common.css');
require('../../../node_modules/@fortawesome/fontawesome-free/css/all.min.css');

let departments = [];
// var currentRole;
let currentUser: string;
let absoluteUrl = '';
let baseUrl = '';

export interface IRequestFormWebPartProps {
  description: string;
}

export default class RequestFormWebPart extends BaseClientSideWebPart<IRequestFormWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
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

  //Render everything
  public async render(): Promise<void> {

    absoluteUrl = this.context.pageContext.web.absoluteUrl;
    baseUrl = absoluteUrl.split('/sites')[0];

    //CSS
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
        width: 87%;
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
      padding: 0.5rem;
      font-size: 1rem;
      font-weight: 500;
      box-shadow: 0 7px 6px -6px #222;
      border: 2px solid black;
      margin-bottom: 0.5rem;
    }
  
    fieldset {
      border: 1px solid #062470;
      padding: 0rem 1rem;
      margin-bottom: 0.5rem;
    }
  
    legend {
      width: auto;
      margin-bottom: 0;
      font-size: 1.2rem;
      color: #062470;
    }
  
    #legalDeptSection{
      display: none;
      background-color: rgb(6, 36, 112, 0.1);
    }
  
    .submitBtnDiv {
      display: flex;
      justify-content: center;
      gap: 3rem;
    }

    .assignBtnDiv {
      display: flex;
      justify-content: right;
      margin-bottom: 10px;
    }
  
    .toggle-container {
      display: flex;
      border-radius: 5px;
      overflow: hidden;
      width: fit-content;
      border: 1px solid #062470;
      font-size: medium;
    }
    
    .toggle-container input[type="radio"] {
      display: none;
    }
    
    .toggle-label {
      padding: 0px 10px;
      cursor: pointer;
      transition: background-color 0.3s ease;
      text-align: center;
      margin-bottom: 0;
    }
    
    #yourDetails:checked + .toggle-label,
    #onBehalf:checked + .toggle-label {
      background-color: rgb(6, 36, 112, 0.1);
    }

    table.dataTable thead th {
      text-align: center!important;
    }

    table.displayContractTable thead th {
      text-align: center!important;
    }

    table.dataTable tbody td {
      text-align: center!important;
    }

    .action-btn {
            border: 2px solid;
            background-color: transparent;
            font-size: 18px;
            padding: 10px 20px;
            cursor: pointer;
            transition: background-color 0.3s, color 0.3s;
            border-radius: 5px;
        }

        /* Accept button styles */
        #acceptBtn {
            color: green;
            border-color: green;
        }

        /* Reject button styles */
        #rejectBtn {
            color: red;
            border-color: red;
        }

        /* Hover styles for buttons */
        #acceptBtn:hover {
            background-color: rgba(0, 128, 0, 0.1); /* Green with slight opacity */
        }

        #rejectBtn:hover {
            background-color: rgba(255, 0, 0, 0.1); /* Red with slight opacity */
        }
            .textarea-container {
  position: relative;
}

.textarea-container textarea {
  width: 100%;
  height: 100px;
  padding: 5px;
  box-sizing: border-box;
  resize: none;
  padding-top: 30px; /* Space for input */
}

.textarea-container input {
  position: absolute;
  top: 5px;
  left: 5px;
  width: calc(100% - 10px);
  padding: 5px;
  box-sizing: border-box;
}
  
  </style>
  
    `;

    //HTML
    this.domElement.innerHTML += `

      <div class="main-container" id="content">

        <div id="nav-placeholder" class="left-panel"></div>

        <div id="middle-panel" class="middle-panel">

          <button id="minimizeButton"></button>

          <div class="${styles.requestForm}" id="form_checklist">

            <form id="requestor_form" style="position: relative; width: 100%;">

              <p id="contractStatus" style="color: green; position: absolute; top: 0; right: 0; font-size: x-large"></p>

              <div class="${styles['form-group']}">
                <h2 style="color: #888;">Request Form</h2>

                <fieldset>
                  <legend id='requestorDetailsLegend'>YOUR DETAILS</legend>

                  <div id="yourDetailsSection" class="${styles.grid}" style="display: flex;align-items: stretch;">

                    <div class="${styles['col-1-3']}">
                      <div class="${styles.controls}">
                        <label for="requestor_name">Name of Requestor*</label>
                        <input type="text" id="requestor_name" required autocomplete="off">
                      </div>
                      <div class="${styles.controls}">
                        <label for="phone_number">Phone Number*</label>
                        <input type="number" id="phone_number" min="0" required autocomplete="off">
                      </div>
                    </div>

                    <div class="${styles['col-1-3']}">
                      <div class="${styles.controls}">
                        <label for="email">Email*</label>
                        <input type="text" id="email" required autocomplete="off">
                      </div>
                      <div class="${styles.controls}">
                        <label for="enl_company">Company*</label>
                        <span id="enl_company_error" class="${styles.errorSpan}">Please select a valid company.</span>
                        <input type="text" placeholder="Please select.." id="enl_company" list='companies_folder' required autocomplete="off">
                        <datalist id="companies_folder"></datalist>
                      </div>
                    </div>

                    <div class="${styles.controls} ${styles['col-1-3']}" style="display: flex; flex-direction: column;">
                      <label for="contributors">Contributors (Email)</label>
                      <div style="display: flex; flex-direction: column; position: relative; height: 100%;">
                        <input type="text" id="contributors_email" style="margin-bottom: 0px; padding-right: 5rem;" autocomplete="off">
                        <div id="contributors" class="${styles.contributorEntryContainer}"></div>
                        <button class="${styles.addPartiesButton}" id="addContributors" type="button">+</button>
                      </div>
                    </div>

                  </div>

                </fieldset>

                <fieldset>
                  <legend title="The more details and info you provide, the better we can assist you. And our team will be reaching out to you shortly to confirm the scope.">
                    HOW CAN WE ASSIST?
                  </legend>

                  <div class="${styles.grid}">
                    <div class="${styles['col-1-4']}">
                      <div class="${styles.controls}">
                        <label for="requestFor"
                        title='You may choose in the list below what you would like us to assist with. If in doubt, no worries, just choose "Other"'
                        >Request For*</label><span id="requestFor_error" class="${styles.errorSpan}">Please select a valid request.</span>
                        <input type="text"  id="requestFor" list='request_List' placeholder="Please select.." required autocomplete="off">
                        <datalist id="request_List"></datalist>
                      </div>
                    </div>

                    <div class="${styles['col-1-4']}">
                      <div class="${styles.controls}" id="uploadFile" style="display: none;">
                        <label for="uploadContract">Upload Contract to Review</label>
                        <input style="background: none; padding: 0px; border: none;" type="file"  id="uploadContract">
                      </div>
                    </div>

                    <div class="${styles['col-1-2']}">
                      <div style="display: flex; flex-direction: column; align-items: flex-start; font-size: large; border: none; height: auto; margin-bottom: 2px;">
                        <p style="font-size: smaller; margin-left: 0; margin-bottom: 4px;">[click if you wish this assignment to be known to Chief Legal Executive only]</p>
                        <label for="checkbox" class="form-check-label" style="font-family: Poppins, Arial, sans-serif; display: flex; align-items: center;">
                          Confidential
                          <input type="checkbox" id="checkbox_confidential" name="checkbox_confidential" style="transform: scale(1.9); margin-left: 1rem; accent-color: #f07e12;" value="YES">
                        </label>
                      </div>
                    </div>

                  </div>

                  <div class="${styles.grid}">
                    <div class="${styles['col-1-1']}">
                      <div class="${styles.controls}">
                        <label for="brief_desc">Tell us more: brief description of the transaction, what do you want to achieve?</label>
                        <textarea type="text"  id="brief_desc" required></textarea>
                      </div>
                    </div>
                  </div>

                </fieldset>

                <fieldset>
                  <legend>PARTIES TO THE AGREEMENT</legend>
                  
                  <div class="${styles.grid}" style="width: 100%; display: flex;">
                    <div style="width: 100%;">
                      <div class="${styles['col-1-2']}">
                        <div class="${styles.controls}">
                          <label for="party1"">Name of Party 1(ENL-Rogers side)*</label><span id="party1_error" class="${styles.errorSpan}">Please select a valid company.</span>
                          <input type="text" placeholder="Please select.." id="party1" list='companies_folder' required autocomplete="off">
                          <datalist id="companies_folder"></datalist>
                        </div>
                      </div>

                      <div class="${styles['col-1-2']}">
                        <div class="${styles.controls}">
                          <label for="party2">Name of Party 2*</label>
                          <input type="text" id="party2" list='companies_folder' required autocomplete="off">
                          <datalist id="companies_folder"></datalist>
                          <select name="party2_type" class="${styles.addPartiesButton} ${styles.dropdownPadding}" id="party2_type">
                            <option value="Internal" selected>Internal</option>
                            <option value="External">External</option>
                          </select>
                        </div>
                      </div>

                      <div class="${styles['col-1-2']}">
                        <div class="${styles.controls}">
                          <div style="position: relative;">
                            <label for="other_parties" 
                             title="If there are more than 2 parties to the agreement, add the remaining parties using the +"
                            >Other Parties</label>
                            <input type="text"  id="other_parties" autocomplete="off">
                            <button class="${styles.addPartiesButton}" id="addOtherParties">+</button>
                          </div>
                        </div>
                      </div>

                      <div class="${styles['col-1-2']}">
                        <div class="${styles.controls}">
                          <div style="position: relative;">
                            <label for="party2_persons">Party 2 Contributors Email (Internal Only)</label>
                            <input type="text" id="party2_persons" autocomplete="off">
                            <button class="${styles.addPartiesButton}" id="addParty2">+</button>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                
                  <div class="${styles.grid}">
                    <div class="${styles['col-1-2']}">
                      <div class="w3-container">
                        <div>
                          <div id="tblOtherParties" class="table-responsive-xl">
                            <div class="form-row">
                              <div class="col-xl-12">
                                <div id="other_parties_tbl">
                                  <table id='tbl_other_Parties' class='table table-striped' style="margin-bottom: 1rem;">
                                    <thead>
                                      <tr>
                                        <th class=" text-left">Other Party</th>
                                        <th class="text-center"></th>
                                      </tr>
                                    </thead>
                                    <tbody id="tb_otherParties"></tbody>
                                  </table>
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>

                    <div class="${styles['col-1-2']}">
                      <div class="w3-container">
                        <div>
                          <div id="tblOtherParties" class="table-responsive-xl">
                            <div class="form-row">
                              <div class="col-xl-12">
                                <div id="party2_tbl">
                                  <table id='tbl_party2' class='table table-striped' style="margin-bottom: 1rem;">
                                    <thead>
                                      <tr>
                                        <th class=" text-left">Party 2</th>
                                        <th class="text-center"></th>
                                      </tr>
                                    </thead>
                                    <tbody id="tb_party2"></tbody>
                                  </table>
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>

                </fieldset>

                <fieldset>
                  <legend>OTHER INFO</legend>

                  <div class="${styles.grid}">
                    <div class="${styles['col-1-3']}">
                      <div class="${styles.controls}">
                        <label for="expectedCommenceDate">Expected Date of Commencement*</label>
                        <input type="date"  id="expectedCommenceDate" required>
                      </div>
                    </div>

                    <div class="${styles['col-1-3']}">
                      <div class="${styles.controls}">
                        <label for="authority_to_approve_contract" title="If you are not the person signing off on this agreement, please let us know who will give clearance in your company">Authority to Approve Contract</label>
                        <div style="display: flex; padding: 8px;">
                          <input type="radio" id="approve_yes" name="authority_to_approve_contract" value="Yes" style="height: 1.5rem; width: 10%;" checked>
                          <label for="approve_yes">Yes</label>
                          <input type="radio" id="approve_no" name="authority_to_approve_contract" value="No" style="height: 1.5rem; width: 10%;">
                          <label for="approve_no">No</label>
                        </div>
                      </div>
                    </div>

                    <div class="${styles.grid}">
                      <div class="${styles['col-1-3']}">
                        <div class="${styles.controls}" id="authorisedApproverDiv" >
                          <label for="authorisedApprover">Name of authorised approver*</label>
                          <input type="text"  id="authorisedApprover" autocomplete="off">
                        </div>
                      </div>
                    </div>

                  </div>

                </fieldset>

              <div id="requestorSubmit" class="submitBtnDiv"></div>

            </form>

            <form id="despatcher_form">
              <fieldset id="legalDeptSection">
                <legend class="${styles.legalLegend}">FOR LEGAL DEPARTMENT ONLY</legend>
              </fieldset>
            </form>

            <div id="section_review_contract">
              <div id="tbl_contract" style="margin-top: 1.5em;"></div>
            </div>

          </div>

        </div>

      </div>
    `;

    //#region SPComponent Loader
    SPComponentLoader.loadScript('https://ajax.googleapis.com/ajax/libs/jquery/3.6.3/jquery.min.js')
      .then(() => {
        // return SPComponentLoader.loadScript('//cdnjs.cloudflare.com/ajax/libs/popper.js/2.9.2/cjs/popper.min.js') 
      })
      .then(() => {
        return SPComponentLoader.loadScript('//cdn.jsdelivr.net/npm/bootstrap@5.0.1/dist/js/bootstrap.min.js');
      })
      .then(() => {
        console.log("Scripts loaded successfully");
      })
      .catch(error => {
        console.error("Error loading scripts: " + error);
      });
    //#endregion

    //Retrieve user roles
    var currentRole = await this.checkCurrentUsersGroupAsync();
    console.group("User's current role:", currentRole);

    //Generate Side Menu
    SideMenuUtils.buildSideMenu(absoluteUrl, departments);
    
    let nameInput = document.getElementById('requestor_name')  as HTMLInputElement;
    let emailInput = document.getElementById('email')  as HTMLInputElement;

    //Retrieve Request ID
    const urlParams = new URLSearchParams(window.location.search);
    const updateRequestID = urlParams.get('requestid');
    const contractDetails = await sp.web.lists.getByTitle("Contract_Request").items.select("NameOfAgreement","Company","NameOfRequestor","Owner","TypeOfContract","Others_parties","Confidential","ContractStatus","Party2_agreement","Party2_Type","Contributors","Party2_Persons").filter(`ID eq ${updateRequestID}`).get();
    console.log(contractDetails);
    // const NameOfAgreement = contractDetails[0].NameOfAgreement;
    let contractStatus = '';
    let companyName = '';
    // let NameOfRequestor = '';
    // let NameOfOwner = '';
    let isConfidential = '';
    let party2_agreement = '';
    let party2_type = '';
    let contributorsArrayInitial = [];
    let party2ContributorsArrayInitial = [];
    // const typeOfAgreement = contractDetails[0].TypeOfContract;
    // const otherParties = contractDetails[0].Others_parties;
    let onBehalf: boolean = false;

    let currentUserNameFromField = '';

    const party2PersonsInput = document.getElementById("party2_persons") as HTMLInputElement;
    const addParty2Button = document.getElementById("addParty2") as HTMLButtonElement;
    const party2TypeInput = document.getElementById("party2") as HTMLInputElement;

    document.getElementById('requestorSubmit').innerHTML += `
      <button type="submit" id="saveToList"><i class="fa fa-refresh icon" style="display: none;"></i>Save</button>
    `;

    //Party 2 datatable
    var party2Table = $('#tbl_party2').DataTable({
      info: false,
      // responsive: true,
      pageLength: 5,
      ordering: false,
      paging: false,
      searching: false,
    });

    //OtherParties datatable
    var otherPartiesTable = $('#tbl_other_Parties').DataTable({
      info: false,
      // responsive: true,
      pageLength: 5,
      ordering: false,
      paging: false,
      searching: false,
    });

    //New Request
    if (!updateRequestID) {
      await this.setRequestorDetails(onBehalf);
      currentUserNameFromField = String($("#requestor_name").val());
      $("#authorisedApprover").val(currentUserNameFromField);

    }
    //Update Request
    else {
      // NameOfAgreement = contractDetails[0].NameOfAgreement;
      companyName = contractDetails[0].Company;
      // NameOfRequestor = contractDetails[0].NameOfRequestor;
      // NameOfOwner = contractDetails[0].Owner;
      isConfidential = contractDetails[0].Confidential;
      party2_agreement = contractDetails[0].Party2_agreement;
      contractStatus = contractDetails[0].ContractStatus;
      party2_type = contractDetails[0].Party2_Type;
      if(contractDetails[0].Contributors){
        contributorsArrayInitial = contractDetails[0].Contributors.split(';').map(email => email.trim());
      }
      if(contractDetails[0].Party2_Persons){
        party2ContributorsArrayInitial = contractDetails[0].Party2_Persons.split(';').map(email => email.trim());
      }
      if(party2_type === "External"){
        party2PersonsInput.disabled = true;
        addParty2Button.disabled = true;
        party2TypeInput.removeAttribute("list");
      }

      //Display Accept or Reject
      if(currentRole === 'OwnerUpdate' && contractStatus === 'ToBeAccepted'){
        console.log('Workings');
        document.getElementById('requestorSubmit').innerHTML = `
          <button id="acceptBtn" class="action-btn">Accept &#10004;</button>
          <button id="rejectBtn" class="action-btn">Reject &#10060;</button>
        `;

        $('#acceptBtn').on('click', async function (event) {
          event.preventDefault();
          const confirmation = confirm("Are you sure you want to accept?");
          if (confirmation) {
            const list = sp.web.lists.getByTitle("Contract_Request");
            await list.items.getById(Number(updateRequestID)).update({
                Owner: currentUser,
                ContractStatus: 'WIP'
            });
            console.log("User accepted:", updateRequestID);
            alert("You accepted the request " + updateRequestID);
            location.reload();
          } 
          else 
          {
            console.log("User cancelled");
          }        
        });

        $('#rejectBtn').on('click', async function (event) {
          event.preventDefault();
          const confirmation = confirm("Are you sure you want to reject?");
          if (confirmation) {
            const list = sp.web.lists.getByTitle("Contract_Request");
            await list.items.getById(Number(updateRequestID)).update({
                AssignedTo: "",
                ContractStatus: 'ToBeAssigned'
            });
            console.log("User rejected:", updateRequestID);
            alert("You rejected the request " + updateRequestID);
            Navigation.navigate(`${absoluteUrl}/SitePages/Dashboard.aspx`, true);
          } 
          else 
          {
            console.log("User deleted cancelled");
          }
        });
      }
      else{
        document.getElementById('saveToList').textContent = 'Update';
      }

      // typeOfAgreement = contractDetails[0].TypeOfContract;
      // otherParties = contractDetails[0].Others_parties;
      this.renderRequestDetails(updateRequestID, otherPartiesTable, party2Table);
      if(contractStatus === 'Cancelled'){
          const formElements = this.domElement.querySelectorAll('input, select, textarea, button');
          formElements.forEach(element => {
          if (element instanceof HTMLInputElement || element instanceof HTMLSelectElement || 
              element instanceof HTMLTextAreaElement || element instanceof HTMLButtonElement) {
              element.disabled = true;
          }
        });
      }
      document.getElementById('contractStatus').innerText = `${contractStatus}`;
    }

    // Root Document library
    const libraryTitle = "Contracts";
    const library = sp.web.lists.getByTitle(libraryTitle);
    var folderId = '';
    if(updateRequestID){
      const consolefolderRetrieval = await library.rootFolder.folders.getByName(companyName).folders.getByName(updateRequestID);
      const consolefolderItem = await consolefolderRetrieval.listItemAllFields.get();
      folderId = consolefolderItem.Id;
      this.consoleFolderUsers(libraryTitle, folderId);
    }

    //Display On Behalf
    if(currentRole === 'OwnerCreate' || currentRole === 'DespatcherCreate'){
      document.getElementById('requestorDetailsLegend').innerHTML = `
        <div class="toggle-container">
          <input type="radio" id="yourDetails" name="toggle" checked>
          <label for="yourDetails" class="toggle-label">YOUR DETAILS</label>
          <input type="radio" id="onBehalf" name="toggle">
          <label for="onBehalf" class="toggle-label">ON BEHALF</label>
        </div>
      `;

      document.getElementById('yourDetails').addEventListener('change', (event: Event) => {
        const target = event.target as HTMLInputElement;
        if (target.checked) {
          onBehalf = false;
          this.setRequestorDetails(onBehalf);
          nameInput.disabled = true;
          emailInput.disabled = true;
        }
      });
  
      document.getElementById('onBehalf').addEventListener('change', (event: Event) => {
        const target = event.target as HTMLInputElement;
        if (target.checked) {
          onBehalf = true;
          this.setRequestorDetails(onBehalf);
          nameInput.disabled = false;
          emailInput.disabled = false;
        }
      });
    }

    //Retrieve requestDigest
    var requestDigest;
    await this.getFormDigest().then(function (data) {
      requestDigest = data.d.GetContextWebInformation.FormDigestValue;
    });

    //Permission levels
    const permissionLevels = {
      FullControl: 1073741829,
      Design: 1073741828,
      Edit: 1073741830,
      Contribute: 1073741827,
      Read: 1073741826,
      LimitedAccess: 1073741825,
      ViewOnly: 1073741924,
      ManageHierarchy: 1073741928
    };

    //Display Legal Department
    if(currentRole === ('DespatcherAssign') || currentRole === 'OwnerUpdate'){
      $('#legalDeptSection').show();

      document.getElementById('legalDeptSection').innerHTML += `
        <div class="legalDept">
          <div class="${styles.grid}" style="display: flex;align-items: stretch;">
            <div class="${styles['col-1-2']}">
              <div id="assignOwners" class="${styles.controls}">
                <label for="assignedTo">Assigned To*</label>
                <input type="text"  placeholder="Please select.." id="assignedTo" list='ownersList' required  autocomplete="off">
                <datalist id="ownersList" style="color: blue"></datalist>
              </div>
              <div class="${styles.controls}">
                <label for="due_date">Due Date*</label>
                <input type="date"  id="due_date" required>
              </div>
              <div class="${styles.controls}">
                <label for="contractType">Type of Contract*</label><span id="contractType_error" class="${styles.errorSpan}">Please select a valid type.</span>
                <input type="text" id="contractType" placeholder="Please select.." list='contractTypeList' required autocomplete="off">
                <datalist id="contractTypeList"></datalist>          
              </div>
            </div>
            <div class="${styles.controls}" style="width: 50%;float: left;display: flex;flex-direction: column;justify-content: flex-end;">
              <div>
                <label for="agreement_name">Name of Agreement</label>
                <input type="text"  id="agreement_name" readonly autocomplete="off">
              </div>
              <label for="DespatcherComments">Comments from Despatcher</label>
              <textarea style="height: 100%;" type="text" placeholder="Your comment here..." id="DespatcherComments"></textarea>
            </div>
          </div>

          <div id='permissionEmails' class="${styles.permissionContainer}"></div>

          <div class="assignBtnDiv" id="assignBtnCont"></div>

        </div>
      `;

      if(currentRole === 'OwnerUpdate'){
        const form = document.getElementById('despatcher_form');
        const inputs = form.querySelectorAll('input');
        inputs.forEach(input => input.readOnly = true);
        $('#DespatcherComments').prop('readonly', true);
      }
      else {
        document.getElementById('assignBtnCont').innerHTML += `
          <button type="submit" id="assignOwner">Assign</button>
        `;
      }

      await this.getSiteUsers();

      this.displayAccessEmails(libraryTitle, folderId);

      //Bind the data-value attribute to the options of the datalist
      // $("#assignedTo").bind('input', () => {
      //   const shownVal = (document.getElementById("assignedTo") as HTMLInputElement).value;
      //   // var shownVal = document.getElementById("name").value;
  
      //   const value2send = (document.querySelector<HTMLSelectElement>(`#ownersList option[value='${shownVal}']`) as HTMLSelectElement).dataset.value;
      //   ownerTitle = value2send;
      //   console.log("LOGG", value2send);
      //   //  $("#created_by").val(value2send);
      // });
    }

    //Disable Name and Email
    if(currentRole === 'RequestorCreate' || currentRole === 'OwnerCreate' || currentRole === 'DespatcherCreate'){
      nameInput.disabled = true;
      emailInput.disabled = true;
    }

    //Display Cancel Button
    if(currentRole === 'RequestorUpdate' || currentRole === 'DespatcherAssign'){
      document.getElementById('requestorSubmit').innerHTML += `
        <button id="cancelRequest" type="button">Cancel</button>
      `;
    }

    $("input[name='authority_to_approve_contract']").change(function() {
      var value = $(this).val();
      currentUserNameFromField = String($("#requestor_name").val());
      if (value === 'Yes') {
        $("#authorisedApprover").val(currentUserNameFromField).prop('disabled', true).removeAttr('required');
      } else {
        $("#authorisedApprover").val('').prop('disabled', false).attr('required', 'required');
      }
    });

    this.load_companies(); //Companies list
    this.load_services(); //Request For list
    await this.load_contractType(); //Companies list
    // this.getAllADUsers2();

    //Valid for dropdown datalists
    document.querySelectorAll('input[list]').forEach(input => {
      input.addEventListener('change', function () {
        const inputElement = input as HTMLInputElement;
        const datalistId = input.getAttribute('list');
        const datalist = document.getElementById(datalistId) as HTMLSelectElement;
        const options = Array.from(datalist.options).map(option => option.value);
        const errorSpan = document.getElementById(input.id + "_error");
        
        if (inputElement.value && !options.includes(inputElement.value)) {
          errorSpan.style.display = "inline";
          inputElement.value = '';
        } else {
          errorSpan.style.display = "none";
        }
      });
    });

    //Display for upload button for file
    $("#requestFor").change(function (e) {
      var $el = $(this);
      var value = $el.val();
      if (value == 'Review of Agreement') {
        $("#uploadFile").css("display", "block");
      } else {
        $("#uploadFile").css("display", "none");
      }
    });

    $("#party2_type").change(function (e) {
      const selectedValue = (this as HTMLSelectElement).value;
    
      if (selectedValue === "External") {
        party2PersonsInput.disabled = true;
        addParty2Button.disabled = true;
        party2TypeInput.removeAttribute("list");
      } else {
        party2PersonsInput.disabled = false;
        addParty2Button.disabled = false;
        party2TypeInput.setAttribute("list", "companies_folder");
      }
    });

    //Process uploaded file
    var filename_add;
    var content_add;
    $('#uploadContract').on('change', () => {
      const input = document.getElementById('uploadContract') as HTMLInputElement | null;

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
    });

    //Add other parties button functionality
    document.querySelector('#addOtherParties').addEventListener('click', (event) => {
      event.preventDefault();
      const otherPartyValue = $("#other_parties").val();
    
      if (otherPartyValue === "") {
        alert("Please enter a value");
      } else {
        this.addNewOtherPartiesRow(otherPartiesTable, otherPartyValue, 'otherParties');
      }
    });

    $('#tbl_other_Parties tbody').on('click', '.delete-row', function (event) {
      event.preventDefault();
      otherPartiesTable.row($(this).closest('tr')).remove().draw(false);
    });

    //Add party 2 contributors button functionality
    document.querySelector('#addParty2').addEventListener('click', (event) => {
      event.preventDefault();
      const otherPartyValue = $("#party2_persons").val();

      if (otherPartyValue === "") {
        alert("Please enter a valid email");
      }
      else {
        this.addNewOtherPartiesRow(party2Table, otherPartyValue, 'party2');
      }
    });

    $('#tbl_party2 tbody').on('click', '.delete-row', function (event) {
      event.preventDefault();
      party2Table.row($(this).closest('tr')).remove().draw(false);
    });

    //Minimize sidebar
    $('#minimizeButton').on('click', function () {
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
          middlePanelID.style.marginLeft = '0%';
          minimizeButtonID.style.left = '0%';
        }
      }
    });

    //ENL CMS GROUP Principal IDs
    var LegalLink_Group_ID = {
      Requestors: 49,
      Despatchers: 46,
      Internal_Owners: 48,
      External_Owners: 47,
      Directors_View: 50
    };

    if(absoluteUrl === 'https://enlmu.sharepoint.com/sites/ContractMgt'){
      LegalLink_Group_ID = {
        Requestors: 13,
        Despatchers: 10,
        Internal_Owners: 12,
        External_Owners: 11,
        Directors_View: 14
      };
    }

    //Retrieve Email of current user
    var currentUserEmail;
    await this.getCurrentUserEmail()
    .then(response => {
        currentUserEmail = response.d.Email;
    })
    .catch(error => {
        console.error('Error retrieving current user email:', error);
    });

    //Retrieve PrincipalId for currrent user
    var principalIdUser;
    await this.getPrincipalIdForUserByEmail(currentUserEmail)
    .then(principalId => {
        principalIdUser = principalId;
    })
    .catch(error => {
        console.error('Error fetching PrincipalId:', error);
    });

    //this.consoleFolderUsers(caseFolderPath);
    
    //Create new request
    var newRequestID;
    document.getElementById("requestor_form").addEventListener("submit", async (event) => {
      event.preventDefault(); // Prevent the default form submission
  
      const form = event.target as HTMLFormElement;
  
      if (form.checkValidity() === false) {
          event.stopPropagation();
          form.classList.add("was-validated");
  
          const firstInvalidElement = form.querySelector(":invalid") as HTMLElement;
          if (firstInvalidElement) {
              firstInvalidElement.focus();
          }
      } 
      else {
        // (document.getElementById('saveToList') as HTMLButtonElement).disabled = true;

        //Other Parties table data
        var dataOtherParties = otherPartiesTable.rows().data();
        var allOtherParties = "";
        var rowCountOtherParties = dataOtherParties.length;

        dataOtherParties.each(function (value, index) {
          var partyName = value[0];
          allOtherParties += partyName;
          if (index < rowCountOtherParties - 1) {
              allOtherParties += ";";
          }
        });

        //Party 2 persons table data
        // var dataParty2 = party2Table.rows().data();
        // var allParty2 = "";
        // var rowCountParty2 = dataParty2.length;

        // dataParty2.each(function (value, index) {
        //   var partyName = value[0];
        //   allParty2 += partyName;
        //   if (index < rowCountParty2 - 1) {
        //     allParty2 += ";";
        //   }
        // });

        const checkbox = document.getElementById('checkbox_confidential') as HTMLInputElement;
        const confidentialValue = checkbox.checked ? 'YES' : 'NO';
        const authorityToApproveContract = $("input[name='authority_to_approve_contract']:checked").val();

        const contributorsValue = getContributorsValue();
        // console.log(contributorsValue);

        const contributorsArrayCurrent = getContributorsArray();
        console.log(contributorsArrayCurrent);

        const party2CntributorsValue = getParty2ContributorsValue();
        // console.log(party2CntributorsValue);

        const party2CntributorsArrayCurrent = getParty2ContributorsArray();
        console.log(party2CntributorsArrayCurrent);

        //Form data
        var formData = {
          NameOfRequestor: $("#requestor_name").val(),
          Email: $("#email").val(),
          Phone_Number: $("#phone_number").val(),
          Company: $("#enl_company").val(),
          Contributors: contributorsValue,
          RequestFor: $("#requestFor").val(),
          Confidential: confidentialValue,
          BriefDescriptionTransaction: $("#brief_desc").val(),
          Party1_agreement: $("#party1").val(),
          Party2_agreement: $("#party2").val(),
          Party2_Type: $("#party2_type").val(),
          Others_parties: allOtherParties,
          Party2_Persons: party2CntributorsValue,
          ExpectedCommencementDate: $("#expectedCommenceDate").val().toString(),
          AuthorityApproveContract: authorityToApproveContract,
          AuthorisedApprover: $("#authorisedApprover").val(),
        };

        console.log(formData);

        //Create Request
        if(currentRole === 'RequestorCreate' || currentRole === 'OwnerCreate' || currentRole === 'DespatcherCreate'){
          try {
            //Add item to Contract Request
            const iar = await sp.web.lists.getByTitle("Contract_Request").items.add(formData)
              .then((iar) => {
                newRequestID = iar.data.ID;
              });
            console.log(newRequestID);

            var contractDetailsLibraryData = {
              Request_ID: newRequestID.toString()
            };
            console.log(contractDetailsLibraryData);

            //Add item to Contract Details
            try {
              await sp.web.lists.getByTitle("Contract_Details").items.add(contractDetailsLibraryData);
            }
            catch (error) {
              console.error('Error adding item in contract_Details:', error);
              throw error;
            }

            const companyFolderName = $("#enl_company").val() as string;
            const contractFolderName = contractDetailsLibraryData.Request_ID;

            //Create contract folder
            const folderCreation = await library.rootFolder.folders.getByName(companyFolderName).folders.add(contractFolderName);
            console.log(`Contract Folder '${contractFolderName}' created successfully.`);
            const folderItem = await folderCreation.folder.listItemAllFields.get();

            console.log('folderItem.Id;', folderItem.Id);
            const folderID = folderItem.Id;
            // const folderID = '508';
            // this.consoleFolderUsers(libraryTitle, folderID);

            //Final path in which document will be stored
            const caseFolderPath = `/sites/ContractMgt/Contracts/${companyFolderName}/${contractFolderName}`;

            //Assign Permissions
            try {
              // Break role inheritance
              await this.breakRoleInheritance(requestDigest, libraryTitle, folderID);
              console.log("Inheritance broken");
            
              // List Folder Permissions before adding new users
              this.consoleFolderUsers(libraryTitle, folderID);
            
              // Add LegalLink_Despatchers group with appropriate permissions
              await this.addRoleAssignment(requestDigest, libraryTitle, folderID, LegalLink_Group_ID.Despatchers, permissionLevels.Edit);
              console.log("LegalLink_Despatchers group added with permissions");
            
              // Loop through each row in party2Table and assign edit permissions
              var dataParty2 = party2Table.rows().data();
              for (var i = 0; i < dataParty2.length; i++) {
                var userEmailParty2 = dataParty2[i][0];
            
                try {
                  // Get the user's principal ID
                  var userPrincipalId = await this.getPrincipalIdForUserByEmail(userEmailParty2);
                  if (userPrincipalId) {
                    await this.addRoleAssignment(requestDigest, libraryTitle, folderID, userPrincipalId, permissionLevels.Edit);
                    console.log(`Permissions assigned to ${userEmailParty2}`);
                  } else {
                    console.error(`Failed to get principal ID for ${userEmailParty2}`);
                  }
                } catch (error) {
                  console.error(`Failed to assign permissions for ${userEmail}:`, error.message);
                }
              }
              
              // Loop through each contributor email and assign edit permissions
              for (var i = 0; i < contributorsArrayCurrent .length; i++) {
                var userEmail = contributorsArrayCurrent [i];
            
                try {
                  // Get the user's principal ID
                  var userPrincipalId = await this.getPrincipalIdForUserByEmail(userEmail);
                  if (userPrincipalId) {
                    await this.addRoleAssignment(requestDigest, libraryTitle, folderID, userPrincipalId, permissionLevels.Edit);
                    console.log(`Permissions assigned to ${userEmail}`);
                  } else {
                    throw new Error(`Principal ID not found for ${userEmail}`);
                  }
                } catch (error) {
                  console.error(`Failed to assign permissions for ${userEmail}:`, error.message);
                }
              }
            
              // List Folder Permissions after adding new users
              this.consoleFolderUsers(libraryTitle, folderID);
            
            } catch (error) {
              console.error("Error updating folder permissions:", error);
            }
            
            //If file has been uploaded
            if ($("#requestFor").val() == 'Review of Agreement') {
              try {
                await this.addFileToContractFolder(caseFolderPath, filename_add, content_add, contractFolderName);
              }
              catch (e) {
                console.log(e.message);
              }
            }

            alert(`Request ${newRequestID} has been submitted successfully.`);

            if(currentRole === 'DespatcherCreate'){
              Navigation.navigate(`${absoluteUrl}/SitePages/Requestor-Form.aspx?requestid=${newRequestID}`, true);
            }
            else{
              Navigation.navigate(`${absoluteUrl}/SitePages/Dashboard.aspx`, true);
            }
          }
          catch (error) {
            console.error('Error adding item:', error);
            throw error;
          }
        }

        //Update Request
        if(currentRole === 'RequestorUpdate' || currentRole === 'OwnerUpdate' || currentRole === 'DespatcherAssign'){

          const removeContributorArray = contributorsArrayInitial.filter(email => !contributorsArrayCurrent.includes(email));
          const addContributorArray = contributorsArrayCurrent.filter(email => !contributorsArrayInitial.includes(email));

          const removeParty2ContributorArray = party2ContributorsArrayInitial.filter(email => !party2CntributorsArrayCurrent.includes(email));
          const addParty2ContributorArray = party2CntributorsArrayCurrent.filter(email => !party2ContributorsArrayInitial.includes(email));

          const duplicateValues = addParty2ContributorArray.filter(email => addContributorArray.includes(email));

          if (duplicateValues.length > 0) {
            return alert(`Duplicate emails found: ${duplicateValues.join(', ')}`);
          }

          const removeCombinedArray = removeParty2ContributorArray.concat(removeContributorArray);
          const addCombinedArray = addParty2ContributorArray.concat(addContributorArray);
          
          console.log('Length', addCombinedArray);
          if (addCombinedArray.length > 0) {
            const validEmails: string[] = [];
            for (const email of addCombinedArray) {
              try {
                const userPrincipalId = await this.getPrincipalIdForUserByEmail(email);
                if (userPrincipalId) {
                  validEmails.push(email);
                } else {
                  return alert(`Invalid email for permissions: ${email}`);
                }
              }
              catch (error) {
                console.error(`Error validating email ${email}:`, error);
                return alert(`Error validating email: ${email}`);
              }
            }
          }

          console.log('Remove', removeCombinedArray);
          console.log('Add', addCombinedArray);

          const folder = await library.rootFolder.folders.getByName(companyName).folders.getByName(updateRequestID).listItemAllFields.select('Id').get();
          console.log(folder);
          const folderId = folder.Id;

          if (removeCombinedArray.length > 0) {
            await removePermissions(removeCombinedArray, folderId);
          }
          
          if (addCombinedArray.length > 0) {
            await addPermissions(addCombinedArray, folderId);
          }

          try {
            const items = await sp.web.lists.getByTitle("Contract_Request").items.filter(`ID eq ${updateRequestID}`).get();
            
            if (items.length > 0) {
              const itemId = items[0].Id;
              console.log('itemId',itemId);
              // Update the item with the new data
              await sp.web.lists.getByTitle("Contract_Request").items.getById(itemId).update(formData);
              console.log("Item updated successfully");
            }

            const caseFolderPath = `/sites/ContractMgt/Contracts/${companyName}/${updateRequestID}`;

            //If file has been uploaded
            if ($("#requestFor").val() == 'Review of Agreement') {
              try {
                await this.addFileToContractFolder(caseFolderPath, filename_add, content_add, updateRequestID);
              }
              catch (e) {
                console.log(e.message);
              }
            }

            alert(`Request has been updated successfully.`);
          }
          catch (error) {
            console.error("Error updating item:", error);
          }

          location.reload();
        }
      }
    });

    function getContributorsValue() {
      const container = document.getElementById('contributors') as HTMLDivElement;
      const spans = container.getElementsByClassName(styles.contributorEmail);
      const contributors = Array.from(spans)
        .map(span => span.textContent?.trim() || '')
        .filter(text => text !== '');
      return contributors.join(';');
    }
    
    function getContributorsArray() {
      const container = document.getElementById('contributors') as HTMLDivElement;
      const spans = container.getElementsByClassName(styles.contributorEmail);
      return Array.from(spans)
        .map(span => span.textContent?.trim() || '')
        .filter(text => text !== '');
    }

    function getParty2ContributorsValue() {
      const table = document.getElementById('tbl_party2') as HTMLTableElement;
      const rows = table.getElementsByTagName('tr');
      
      if (rows.length <= 1) {
        return '';
      }
    
      const party2Values = Array.from(rows)
        .slice(1)
        .map(row => {
          const cells = row.getElementsByTagName('td');
          return cells[0]?.textContent?.trim() || '';
        })
        .filter(text => text !== '' && text !== 'No data available in table');
    
      return party2Values.join(';');
    }

    function getParty2ContributorsArray() {
      const table = document.getElementById('tbl_party2') as HTMLTableElement;
      const rows = table.getElementsByTagName('tr');
      
      const secondRowFirstCell = rows[1].getElementsByTagName('td')[0];
      if (secondRowFirstCell?.textContent?.trim() === 'No data available in table') {
        return [];
      }
    
      return Array.from(rows)
        .slice(1)
        .map(row => {
          const cells = row.getElementsByTagName('td');
          return cells[0]?.textContent?.trim() || '';
        })
        .filter(text => text !== '');
    }    

    const getPrincipalIdForUserByEmail = this.getPrincipalIdForUserByEmail.bind(this);
    const removeRoleAssignments = this.removeRoleAssignments.bind(this);
    const addRoleAssignment = this.addRoleAssignment.bind(this);

    // Function to handle permission removal
    async function removePermissions(removeContributorArray, folderId) {
      for (const email of removeContributorArray) {
          try {
              const userPrincipalId = await getPrincipalIdForUserByEmail(email);
              if (userPrincipalId) {
                  await removeRoleAssignments(requestDigest, libraryTitle, folderId, userPrincipalId, permissionLevels.Edit);
                  console.log(`Permissions removed for ${email}`);
              } else {
                  throw new Error(`Principal ID not found for ${email}`);
              }
          } catch (error) {
              console.error(`Failed to remove permissions for ${email}:`, error.message);
          }
      }
    }

    // Function to handle permission addition
    async function addPermissions(addContributorArray, folderId) {
      for (const email of addContributorArray) {
          try {
              const userPrincipalId = await getPrincipalIdForUserByEmail(email);
              if (userPrincipalId) {
                  await addRoleAssignment(requestDigest, libraryTitle, folderId, userPrincipalId, permissionLevels.Edit);
                  console.log(`Permissions assigned to ${email}`);
              } else {
                  throw new Error(`Principal ID not found for ${email}`);
              }
          } catch (error) {
              console.error(`Failed to assign permissions for ${email}:`, error.message);
          }
      }
    }

    //Assign Owner
    document.getElementById("despatcher_form").addEventListener("submit", async (event) => {
      event.preventDefault();
  
      const form = event.target as HTMLFormElement;
  
      if (form.checkValidity() === false) {
        event.stopPropagation();
        form.classList.add("was-validated");

        const firstInvalidElement = form.querySelector(":invalid") as HTMLElement;
        if (firstInvalidElement) {
          firstInvalidElement.focus();
        }
      } 
      else {
        // (document.getElementById('assignOwner') as HTMLButtonElement).disabled = true;

        // Retrieve the data-value attribute of the selected option in the datalist
        const shownVal = (document.getElementById("assignedTo") as HTMLInputElement).value;
        const selectedOption = document.querySelector(`#ownersList option[value='${shownVal}']`);
        let OwnerEmail = "";
        if (selectedOption) {
          OwnerEmail = selectedOption.getAttribute('data-value') || "";
        }
        const contractTyoe = $("#contractType").val();
        const currentDate = new Date().toISOString().split('T')[0];
        const agreementName = (currentDate + '_' + companyName + '_' + updateRequestID + '_' + contractTyoe + '_' + party2_agreement);

        // Form data
        const assignData = {
          // AssigneeComment: $("#comment").val(),
          AssignedTo: $("#assignedTo").val(),
          OwnerEmail: OwnerEmail,
          DueDate: $("#due_date").val(),
          TypeOfContract: $("#contractType").val(),
          DespatcherComments: contractTyoe,
          NameOfAgreement: agreementName,
          ContractStatus: 'ToBeAccepted'
        };

        console.log(assignData);
      
        try {
          const folderRetrieval = await library.rootFolder.folders.getByName(companyName).folders.getByName(updateRequestID);
          const folderItem = await folderRetrieval.listItemAllFields.get();
          // console.log('folderItem.Id;', folderItem.Id);
          const folderID = folderItem.Id;

          //Assign Permissions
          try {

            //List Folder Permissions
            this.consoleFolderUsers(libraryTitle, folderID);

            if(isConfidential == 'YES'){
              const ownerPrincipalID = await this.getPrincipalIdForUserByEmail(OwnerEmail);
              //Add Owner with appropriate permissions
              await this.addRoleAssignment(requestDigest, libraryTitle, folderID, ownerPrincipalID, permissionLevels.Edit);
              console.log("Owner added with permissions");
            }
            else if(isConfidential === 'NO'){
              //Add LegalLink_Internal_Owners  group with appropriate permissions
              await this.addRoleAssignment(requestDigest, libraryTitle, folderID, LegalLink_Group_ID.Internal_Owners, permissionLevels.Edit);
              console.log("LegalLink_Internal_Owners group added with permissions");
            }

            this.consoleFolderUsers(libraryTitle, folderID);
        
          } catch (error) {
            console.error("Error updating folder permissions:", error);
          }

          // Get the item with the matching request_ID
          const items = await sp.web.lists.getByTitle("Contract_Request").items.filter(`ID eq ${updateRequestID}`).get();
          
          if (items.length > 0) {
            const itemId = items[0].Id;
            console.log('itemId',itemId);
            // Update the item with the new data
            await sp.web.lists.getByTitle("Contract_Request").items.getById(itemId).update(assignData);
            console.log("Item updated successfully");
            alert(`Request has been assigned to ${assignData.AssignedTo} successfully.`);
            Navigation.navigate(`${absoluteUrl}/SitePages/Dashboard.aspx`, true);
          }
          else {
            console.log("Item with the specified request_ID not found");
          }
        } catch (error) {
          console.error("Error updating item:", error);
        }
      }
    });

    //Cancel Request
    $('#cancelRequest').on('click', async function () {
      const userConfirmed = confirm("Are you sure you want to cancel this request?");
      if (!userConfirmed) {
         return;
      }
      const cancelledStatus = {
        ContractStatus: 'Cancelled'
      };

      try {
        await sp.web.lists.getByTitle("Contract_Request").items.getById(Number(updateRequestID)).update(cancelledStatus);
        alert("Request cancelled successfully.");
        Navigation.navigate(`${absoluteUrl}/SitePages/Dashboard.aspx`, true);
      } catch (error) {
        console.error("Error cancelling the request:", error);
      }
    });

    $('#contributors_email').on('keydown', function(event) {
      if (event.key === 'Enter') {
        event.preventDefault();
        $('#addContributors').click();
      }
    });

    document.addEventListener('keydown', function (event) {
      if (event.key === 'Enter') {
        const activeElement = document.activeElement;
        if (activeElement && activeElement.tagName === 'INPUT') {
          event.preventDefault();
        }
      }
    });
    
    $('#addContributors').on('click', function (event) {
      event.preventDefault();
      const input = document.getElementById('contributors_email') as HTMLInputElement;
      const container = document.getElementById('contributors') as HTMLDivElement;
      
      if (input.value.trim() !== '') {
        const entryDiv = document.createElement('div');
        entryDiv.className = `${styles.contributorEntry}`;
        
        const entryText = document.createElement('span');
        entryText.className = `${styles.contributorEmail}`;
        entryText.textContent = input.value;
        
        const removeButton = document.createElement('button');
        removeButton.className = `${styles.removeButton}`;
        removeButton.innerHTML = '&#10060;';
        removeButton.onclick = function () {
          container.removeChild(entryDiv);
        };
        
        entryDiv.appendChild(entryText);
        entryDiv.appendChild(removeButton);
        container.appendChild(entryDiv);
        
        input.value = '';
      }
    
      container.scrollTop = container.scrollHeight;
    });

    document.getElementById('party2_type')!.addEventListener('change', function (event) {
      const party2TypeSelect = event.target as HTMLSelectElement;
      const table = document.getElementById('tbl_party2') as HTMLTableElement;
      const rows = table.getElementsByTagName('tr');
    
      const dataRows = Array.from(rows).filter(row => {
        const cells = row.getElementsByTagName('td');
        return cells.length > 0 && cells[0].textContent?.trim() !== 'No data available in table';
      });
    
      if (party2TypeSelect.value === 'External' && dataRows.length > 0) {
        alert('Cannot change to External while there are entries in the table. Please remove all entries first.');
        party2TypeSelect.value = 'Internal';
      }
    });
    
    document.addEventListener('DOMContentLoaded', function () {
      const datalists = document.querySelectorAll('datalist');
      datalists.forEach(datalist => {
        const options = datalist.querySelectorAll('option');
        options.forEach(option => {
            if (!option.hasAttribute('data-added')) {
                option.remove();
            }
        });
      });
    });
  }

  //RequestDigest
  private getFormDigest() {
    return $.ajax({
        url: absoluteUrl + "/_api/contextinfo",
        method: "POST",
        headers: {
            "Accept": "application/json; odata=verbose"
        }
    });
  }

  //Retrieve Current User Email
  private getCurrentUserEmail() {
    const restUrl = `${absoluteUrl}/_api/web/currentUser?$select=Email`;

    return $.ajax({
        url: restUrl,
        method: "GET",
        headers: {
            "Accept": "application/json; odata=verbose"
        }
    });
  }

  //Retrieve principalId of user using email address
  private getPrincipalIdForUserByEmail(UserEmail) {
    const restUrl = `${absoluteUrl}/_api/web/SiteUserInfoList/items?$filter=EMail eq '${encodeURIComponent(UserEmail)}'`;

    return $.ajax({
        url: restUrl,
        method: "GET",
        headers: {
            "Accept": "application/json; odata=verbose"
        }
    }).then(response => {
        if (response.d && response.d.results && response.d.results.length > 0) {
            return response.d.results[0].ID; // PrincipalId is typically found in the ID field
        } else {
            throw new Error(`User with email ${UserEmail} not found.`);
        }
    });
  }

  //Adding Role Assignment for User
  // private async addRoleAssignmentUser(requestDigest, folderUrl, principalId, roleDefId) {
  //   try {
  //       // Step 1: Break role inheritance (optional if already broken)
  //       // await this.breakRoleInheritance(requestDigest, folderUrl);

  //       // Step 2: Add role assignment for the specified user
  //       const restUrl = `${absoluteUrl}/_api/web/getFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields/RoleAssignments/addroleassignment(principalid=${principalId}, roledefid=${roleDefId})`;
        
  //       await $.ajax({
  //           url: restUrl,
  //           method: "POST",
  //           headers: {
  //               "Accept": "application/json; odata=verbose",
  //               "X-RequestDigest": requestDigest
  //           }
  //       });

  //       console.log(`Role assigned to user with Principal ID ${principalId}`);
  //   } catch (error) {
  //       console.error('Error adding role assignment:', error);
  //       throw error;
  //   }
  // }

  //Display current users on folder on console
  private async consoleFolderUsers(libraryTitle, folderID) {
    const response = await this.getFolderPermissions(libraryTitle, folderID);
      const roleAssignments = response.d.results;
      console.log(roleAssignments);
      //Listing Folder Permissions
      const users = roleAssignments
        .filter(roleAssignment => roleAssignment.Member.PrincipalType === 1 || 8)
        .map(roleAssignment => roleAssignment.Member.Title);
      console.log('Users with access to the folder:', users);
  }

  private async displayAccessEmails(libraryTitle, folderID) {
    const response = await this.getFolderPermissions(libraryTitle, folderID);
      const roleAssignments = response.d.results;
      console.log(roleAssignments);
      //Listing Folder Permissions
      const users = roleAssignments
        .filter(roleAssignment => roleAssignment.Member.PrincipalType === 1 || 8)
        .map(roleAssignment => roleAssignment.Member.Title);

    const container = document.getElementById('permissionEmails') as HTMLDivElement;
  
    users.forEach(email => {
      // Create a div for each email
      const emailDiv = document.createElement('div');
      emailDiv.className = styles.emailPermissionContainer;
  
      // Create a span to display the email
      const emailSpan = document.createElement('span');
      emailSpan.className = styles.emailPermission;
      emailSpan.textContent = email;
  
      emailDiv.appendChild(emailSpan);
      container.appendChild(emailDiv);
    });
  }
  

  //Retrieve folder user and group access
  private getFolderPermissions(libraryTitle, folderID) {
    const restUrl = `${absoluteUrl}/_api/Web/Lists/GetByTitle('${libraryTitle}')/Items(${folderID})/RoleAssignments?$expand=Member,RoleDefinitionBindings`;
    // console.log(restUrl);
  
    return $.ajax({
      url: restUrl,
      method: "GET",
      headers: {
        "Accept": "application/json; odata=verbose"
      }
    });
  }

  //Break Permissions
  private breakRoleInheritance(requestDigest, libraryTitle, folderID) {
    const restUrl = `${absoluteUrl}/_api/web/Lists/GetByTitle('${libraryTitle}')/Items(${folderID})/breakroleinheritance(copyRoleAssignments=false)`;

    return $.ajax({
      url: restUrl,
      method: "POST",
      headers: {
        "Accept": "application/json; odata=verbose",
        "X-RequestDigest": requestDigest
      }
    });
  }

  private removeRoleAssignments(requestDigest, libraryTitle, folderID, roleAssignmentId) {
    const restUrl = `${absoluteUrl}/_api/web/Lists/GetByTitle('${libraryTitle}')/Items(${folderID})/roleassignments/removeroleassignment(principalid=${roleAssignmentId})`;
    return $.ajax({
      url: restUrl,
      method: "POST",
      headers: {
        "Accept": "application/json; odata=verbose",
        "X-RequestDigest": requestDigest
      }
    });
  }
  
  private addRoleAssignment(requestDigest, libraryTitle, folderID, principalId, roleDefId) {
    const restUrl = `${absoluteUrl}/_api/web/Lists/GetByTitle('${libraryTitle}')/Items(${folderID})/roleassignments/addroleassignment(principalid=${principalId}, roledefid=${roleDefId})`;
    // console.log(restUrl);
    return $.ajax({
      url: restUrl,
      method: "POST",
      headers: {
        "Accept": "application/json; odata=verbose",
        "X-RequestDigest": requestDigest
      }
    });
  }

  public async checkCurrentUsersGroupAsync() {
    var currentRole;
    let groupList = await sp.web.currentUser.groups();
    console.log('grouplist: ', groupList);
  
    const urlParams = new URLSearchParams(window.location.search);
    const updateRequestID = urlParams.get('requestid');
    
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

    if (departments.length === 0) {
      departments.push("noGroup");
    }
    else if(departments.length === 1) {
      if (departments.includes('Requestor')) {
        if (!updateRequestID){
          return currentRole = 'RequestorCreate'; //New Request
        }
        else{
          return currentRole = 'RequestorUpdate'; //Update Request
        }
      }
      else if (departments.includes('ExternalOwner')) {
        return currentRole = 'ExternalOwnerOnly' //External Owner Only -> Disable Submit Button
      }
    }
    else if(departments.length === 2){
      if (departments.includes('Requestor') && (departments.includes('InternalOwner') || departments.includes('ExternalOwner') || (departments.includes('DirectorsView')))) {
        if (!updateRequestID){
          if(departments.includes('DirectorsView')){
            return currentRole = 'RequestorCreate'; //New Request by Director's View
          }
          else{
            return currentRole = 'OwnerCreate'; //New Request by Internal Owner or External Owner on behalf of requestor or for themselves
          }
        }
        else {
          if (departments.includes('InternalOwner')){
            return currentRole = 'OwnerUpdate'; //Internal Owner
          }
          else{
            return currentRole = 'ExternalOwnerView';//External Owner
          }
        }
      }
    }
    else if(departments.length === 3){
      if (departments.includes('Requestor') && departments.includes('InternalOwner') && departments.includes('Despatcher')){
        if (!updateRequestID){
          return currentRole = 'DespatcherCreate'; //New Request by despatcher on behalf of requestor
        }
        else{
          return currentRole = 'DespatcherAssign'; //Despatcher edit and assign
        }
      }
    }
  }

  // private async getAllADUsers2(): Promise<void> {
  //   try {
  //     const result: any[] = [];
  //     const datalist = document.getElementById('ADUsers') as HTMLDataListElement;
  //     datalist.innerHTML = ''; // Clear existing options
  
  //     let request: GraphRequest | null = this.graphClient.api('/users').top(999);
  
  //     while (request !== null) {
  //       const response = await request.get();
  //       const users = response.value;
  
  //       result.push(...users);
  
  //       if (response['@odata.nextLink']) {
  //         request = this.graphClient.api(response['@odata.nextLink']);
  //       } else {
  //         request = null;
  //       }
  //     }
  
  //     console.log("USERS", result);
  
  //     // Populate the datalist with users
  //     result.forEach(user => {
  //       const option = document.createElement('option');
  //       option.value = user.mail; // Use the email as the value
  //       option.textContent = user.displayName; // Optionally display the name
  //       datalist.appendChild(option);
  //     });
      
  //   } catch (error) {
  //     console.error("Error fetching users: ", error);
  //     throw error;
  //   }
  // }
  

  //New row for other parties
  addNewOtherPartiesRow(table, party, partyType) {
    table.row.add([
      party,
      '<button class="delete-row" style="background: none; padding: 0px;">&#10060;</button>'
    ]).draw(false);
  
    if (partyType === 'otherParties') {
      $("#other_parties").val("");
    } else if (partyType === 'party2') {
      $("#party2_persons").val("");
    }
  }

  public async load_companies() {
    const drp_companies = document.getElementById("companies_folder") as HTMLSelectElement;
    if (!drp_companies) {
      console.error("Dropdown element not found");
      return;
    }
    const companies = await sp.web.lists.getByTitle('Companies').items.getAll();
    console.log(companies);

    await Promise.all(companies.map(async (result) => {
      const opt = document.createElement('option');
      opt.value = result.Title;
      drp_companies.appendChild(opt);
    }));
  }

  public async load_contractType() {
    const drp_contractType = document.getElementById("contractTypeList") as HTMLSelectElement;
    if (!drp_contractType) {
        console.error("Dropdown element not found");
        return;
    }
    const contractType = await sp.web.lists.getByTitle('Type of contracts').items.get();

    await Promise.all(contractType.map(async (result) => {
        const opt = document.createElement('option');
        opt.value = result.Title;
        drp_contractType.appendChild(opt);
    }));
  } 

  public async load_services() {
    const drp_companies = document.getElementById("request_List") as HTMLSelectElement;
    if (!drp_companies) {
      console.error("Dropdown element not found");
      return;
    }
    const companies = await sp.web.lists.getByTitle('ENL_Services').items.get();
    await Promise.all(companies.map(async (result) => {
      const opt = document.createElement('option');
      opt.value = result.Title;
      drp_companies.appendChild(opt);
    }));
  }

  // Function to set requestor details
  public async setRequestorDetails(onBehalf: boolean) {
    const requestor = await sp.web.currentUser();
    console.log("req:", requestor);

    const fields = [
      { id: "#requestor_name", value: requestor.Title },
      { id: "#email", value: requestor.Email },
      // Add more fields as needed
    ];

    fields.forEach(field => {
      if (!onBehalf) {
        $(field.id).val(field.value);
      } else {
        $(field.id).val('');
      }
    });
  }

  // async addFolderToDocumentLibrary(libraryTitle, companyFolderName, contractFolderName) {
  //   const library = sp.web.lists.getByTitle(libraryTitle);

  //   try {
  //     const exists = await this.folderExists(library, companyFolderName, contractFolderName);

  //     //None exists
  //     if(exists === "noneExist"){
  //       //Create company folder
  //       await library.rootFolder.folders.add(companyFolderName);
  //       console.log(`Company Folder '${companyFolderName}' created successfully.`);
  //       //Create contract folder
  //       await library.rootFolder.folders.getByName(companyFolderName).folders.add(contractFolderName);
  //       console.log(`Contract Folder '${contractFolderName}' created successfully.`);
  //     }
  //     else if(exists === "companyOnly"){
  //       //Create contract folder
  //       await library.rootFolder.folders.getByName(companyFolderName).folders.add(contractFolderName);
  //       console.log(`Contract Folder '${contractFolderName}' created successfully.`);
  //     }
  //     else if(exists === "allExist"){
  //       console.log(`All folders already exist.`);
  //     }

  //   }
  //   catch (error) {
  //     console.error(`Error creating folder: ${error.message}`);
  //   }

  //   // try {
  //   //   console.log(1);
  //   //   //Check existence of company folder
  //   //   const exists = await this.folderExists(libraryTitle, companyFolderName, contractFolderName);

  //   //   if(exists == 'allExist'){
  //   //     console.log(9);
  //   //     console.log(`All folders exist.`);
  //   //   }
  //   //   else {
  //   //     console.log(10);
  //   //     if(exists == 'noneExist'){
  //   //       // Create a new company folder
  //   //       const library = sp.web.lists.getByTitle(libraryTitle);
  //   //       await library.rootFolder.folders.add(companyFolderName);
  //   //       console.log(`Company Folder '${companyFolderName}' created successfully.`);
  //   //     }
  //       //  console.log(`Contract Folder '${contractFolderName}'`);
  //       // const library = sp.web.lists.getByTitle(libraryTitle);
  //       // await library.rootFolder.folders.add(contractFolderName);
  //       // console.log(`Contract Folder '${contractFolderName}' created successfully.`);
  //     // }

  //     // Get the document library by title

  //   // } catch (error) {
  //   //   console.log(11);
  //   //   console.error(`Error creating folder: ${error.message}`);
  //   // }
  // }

  // async folderExists(library, companyFolderName, contractFolderName) {

  //   let existResponse = "";

  //   // Check if company folder exists
  //   try {
  //     const companyFolder = await library.rootFolder.folders.getByName(companyFolderName).select("Exists").get();
  //     console.log("Company folder exists");
  //     //Company folder exists
  //     if(companyFolder.Exists){
  //       try{
  //         const contractFolder = await library.rootFolder.folders.getByName(companyFolderName).folders.getByName(contractFolderName).select("Exists").get();
  //         if(contractFolder.Exists){
  //           console.log("Contract folder exists");
  //           existResponse = "allExist"; 
  //           return existResponse;
  //         }
  //       }
  //       catch(error){
  //         console.log(error);
  //         console.log("Contract folder does not exist");
  //         existResponse = "companyOnly"; 
  //         return existResponse;
  //       }
  //     }
  //   }
  //   catch (error) {
  //     //Company folder does not exist
  //     console.log(error);
  //     console.log("Company folder does not exist");
  //     existResponse = "noneExist"; 
  //     return existResponse;
  //   }

  // }

  // //If file name already exists, file will not be uploaded
  
  async addFileToContractFolder(folderPath, fileName, fileContent, requestId) {
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

  // async addFileToContractFolder(folderPath, fileName, fileContent, requestId, userEmail, role) {
  //   try {
  //     // Add the file to the folder
  //     const fileData = await sp.web.getFolderByServerRelativeUrl(folderPath)
  //       .files.add(fileName, fileContent, false);
  
  //     // Get the list item associated with the file
  //     const item = await fileData.file.getItem();
  //     await item.update({
  //       Request_Id: requestId
  //     });
  
  //     // Get the file's list item to manage permissions
  //     const fileItem: any = sp.web.getFileByServerRelativeUrl(fileData.data.ServerRelativeUrl);
  
  //     // Break permission inheritance (stop inheriting permissions from the parent folder/library)
  //     await fileItem.breakRoleInheritance(true);
  
  //     // Get the user to assign permissions
  //     const user = await sp.web.ensureUser(userEmail);
  
  //     // Define the role (permissions) to assign (e.g., 'read', 'contribute', etc.)
  //     let roleDef;
  //     switch (role.toLowerCase()) {
  //       case 'read':
  //         roleDef = sp.web.roleDefinitions.getByName('Read');
  //         break;
  //       case 'contribute':
  //         roleDef = sp.web.roleDefinitions.getByName('Contribute');
  //         break;
  //       case 'edit':
  //         roleDef = sp.web.roleDefinitions.getByName('Edit');
  //         break;
  //       case 'full control':
  //         roleDef = sp.web.roleDefinitions.getByName('Full Control');
  //         break;
  //       default:
  //         roleDef = sp.web.roleDefinitions.getByName('Read');
  //     }
  
  //     // Assign the permissions to the user
  //     await fileItem.roleAssignments.add(user.data.Id, roleDef.Id);
  
  //     console.log('File uploaded and permissions set successfully.');
  //     alert('File uploaded and permissions set successfully.');
  //   } catch (error) {
  //     console.error('Error uploading file or setting permissions:', error);
  //     alert('Error uploading file or setting permissions.');
  //     throw error;
  //   }
  // }
  
  //Render Update Request Details
  private renderRequestDetails(id: any, otherPartiesTable, party2Table) {
    var checkbox = document.getElementById('checkbox_confidential') as HTMLInputElement;
    try {
      let currentWebUrl = this.context.pageContext.web.absoluteUrl;
      let requestUrl = currentWebUrl.concat(`/_api/web/Lists/GetByTitle('Contract_Request')/items?$select=ID, NameOfRequestor, Email, Phone_Number, Company, Contributors, 
      Party1_agreement, Party2_agreement, Party2_Type, Others_parties, Party2_Persons, BriefDescriptionTransaction, ExpectedCommencementDate, AssignedTo, Owner, AssigneeComment, DueDate,
      Confidential, AuthorityApproveContract, AuthorisedApprover, DueDate, TypeOfContract, NameOfAgreement, DespatcherComments, RequestFor &$filter=(ID eq '${id}') `);
      var doc = null;
      var date = null;
      let html: string = "";

      const response = this.context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            response.json()
              .then((responseJSON) => {
                if (responseJSON != null && responseJSON.value != null) {
                  doc = responseJSON.value;

                  console.log("Items", doc);


                  doc.forEach((result: any) => {
                    const item = {
                      Request_ID: result.ID,
                      NameOfRequestor: result.NameOfRequestor,
                      Email: result.Email,
                      Phone_Number: result.Phone_Number,
                      Company: result.Company,
                      Contributors: result.Contributors,
                      Party1_agreement: result.Party1_agreement,
                      Party2_agreement: result.Party2_agreement,
                      Party2Type: result.Party2_Type,
                      Others_parties: result.Others_parties,
                      Party2_Persons: result.Party2_Persons,
                      BriefDescriptionTransaction: result.BriefDescriptionTransaction,
                      ExpectedCommencementDate: result.ExpectedCommencementDate,
                      AuthorityApproveContract: result.AuthorityApproveContract,
                      AuthorisedApprover: result.AuthorisedApprover,
                      DueDate: result.DueDate,
                      TypeOfContract: result.TypeOfContract,
                      NameOfAgreement: result.NameOfAgreement,
                      RequestFor: result.RequestFor,
                      AssignedTo: result.AssignedTo,
                      Owner: result.Owner,
                      Confidential: result.Confidential,
                      DespatcherComments: result.DespatcherComments

                      // Date_time: result.DateTime,
                      // Attachments: result.AttachmentFiles
                    };

                    // console.log("Comments list:");
                    // console.log(item);

                    if (!Date.parse(item.DueDate)) {
                      date = item.DueDate;
                    }
                    else {
                      date = moment(new Date(item.DueDate)).format("DD/MM/YYYY HH:mm");
                    }

                    // if (item.RequestFor == 'Review of Agreement') {

                    $("#section_review_contract").css("display", "block");

                    this.getFileDetailsByFilter('Contracts', id)
                      .then((fileDetails) => {
                        if (fileDetails) {
                          console.log("File URL:", fileDetails.fileUrl);
                          console.log("File Name:", fileDetails.fileName);

                          let html: string = `
                              <div class="form-row">
                                  <fieldset style="width: 100%;">
                                    <legend>View Contract</legend>
                                      <table id="displayContractTable" style="margin-bottom: 1rem;">
                                          <thead>
                                              <tr>
                                                  <th class="th-lg" scope="col">Contract</th>
                                                  <th scope="col">View</th>
                                              </tr>
                                          </thead>
                                          <tbody>
                                              <tr>
                                                  <td scope="row">${fileDetails.fileName}</td>
                                                  <td>
                                                      <ul class="list-inline m-0">
                                                          <li class="list-inline-item">
                                                              <button id="btn_view" class="btn btn-secondary btn-sm rounded-circle" type="button" data-toggle="tooltip" data-placement="top" title="View" style="display: none;">
                                                                  <i class="fas fa-eye"></i>
                                                              </button>
                                                          </li>
                                                          <li class="list-inline-item">
                                                              <button id="modalActivate" class="btn btn-secondary btn-sm rounded-circle" type="button" data-toggle="modal" data-target="#exampleModalPreview" style="display: block; width: auto;">
                                                                  <i class="fas fa-eye"></i>
                                                              </button>
                                                          </li>
                                                      </ul>
                                                  </td>
                                              </tr>
                                          </tbody>
                                      </table>
                                  </fieldset>
                              </div>
                          `;


                          const listContainer: Element = this.domElement.querySelector('#tbl_contract');
                          listContainer.innerHTML = html;

                          // Initialize DataTable
                          $('#displayContractTable').DataTable({
                            info: false,
                            // responsive: true,
                            // pageLength: 5,
                            ordering: false,
                            paging: false,
                            searching: false,
                          });
                      

                          $("#modalActivate").click(() => {
                            //console.log("calling viewing ...");
                            // this.submit_main();
                            window.open(`ms-word:ofv|u|${baseUrl}/${fileDetails.fileUrl}`, '_blank');
                            // Navigation.navigate(`${urlChoice}`, '_blank');
                          });


                        } else {
                          console.log("Item not found.");
                        }
                      })
                      .catch((error) => {
                        console.log(error);
                      });
                    // }

                    $("#requestor_name").val(item.NameOfRequestor);
                    $("#email").val(item.Email);
                    $("#phone_number").val(item.Phone_Number),
                    $("#enl_company").val(item.Company);
                    
                    if (item.Contributors) {
                      const contributorsArray = item.Contributors.split(';');
                      const container = document.getElementById('contributors') as HTMLDivElement;
                      container.innerHTML = '';
                  
                      contributorsArray.forEach(contributor => {
                        if (contributor.trim() !== '') {
                          const entryDiv = document.createElement('div');
                          entryDiv.className = `${styles.contributorEntry}`;
                          
                          const entryText = document.createElement('span');
                          entryText.className = `${styles.contributorEmail}`;
                          entryText.textContent = contributor;
                          
                          const removeButton = document.createElement('button');
                          removeButton.className = `${styles.removeButton}`;
                          removeButton.innerHTML = '&#10060;';
                          removeButton.onclick = function () {
                            container.removeChild(entryDiv);
                          };
                          
                          entryDiv.appendChild(entryText);
                          entryDiv.appendChild(removeButton);
                          container.appendChild(entryDiv);
                        }
                      });
                    }
                    
                    $("#requestFor").val(item.RequestFor);
                    $("#party1").val(item.Party1_agreement);
                    $("#party2").val(item.Party2_agreement);
                    $("#party2_type").val(item.Party2Type);

                    // Other parties populate table
                    // $("#other_parties").val(item.Others_parties);

                    if (item.Others_parties !== null && item.Others_parties !== "") {
                      var othersPartiesVal = item.Others_parties;
                      othersPartiesVal = othersPartiesVal.replace(/;+$/, '');
                      var otherPartiesArray = othersPartiesVal.split(';');
                      var tbodyOtherParties = document.getElementById('tb_otherParties');
                      tbodyOtherParties.innerHTML = '';
                  
                      otherPartiesArray.forEach(function (value) {
                        otherPartiesTable.row.add([
                          value,
                          '<button class="delete-row" style="background: none; padding: 0px;">&#10060;</button>'
                        ]).draw(false);
                      });
                    }

                    if(item.Party2_Persons !== null && item.Party2_Persons !== "") {
                      var party2_PersonsVal = item.Party2_Persons;
                      party2_PersonsVal = party2_PersonsVal.replace(/;+$/, '');
                      var party2Array = party2_PersonsVal.split(';');
                      var tbodyParty2 = document.getElementById('tb_party2');
                      tbodyParty2.innerHTML = '';

                      party2Array.forEach(function (value) {
                        party2Table.row.add([
                          value,
                          '<button class="delete-row" style="background: none; padding: 0px;">&#10060;</button>'
                        ]).draw(false);
                      });
                    }

                    $("#brief_desc").val(item.BriefDescriptionTransaction);
                    $("#expectedCommenceDate").val(item.ExpectedCommencementDate);
                    var authorityApproveContract = item.AuthorityApproveContract;
                    if (authorityApproveContract === 'Yes') {
                        $("#approve_yes").prop("checked", true);
                    } else {
                        $("#approve_no").prop("checked", true);
                    }

                    if(item.AuthorityApproveContract === 'Yes'){
                      $("#authorisedApproverDiv").show();
                      $("#authorisedApprover").val(item.AuthorisedApprover);
                    }

                    $("#assignedTo").val(item.AssignedTo);
                    $("#contractType").val(item.TypeOfContract);
                    $("#due_date").val(item.DueDate);
                    $("#DespatcherComments").val(item.DespatcherComments);

                    // const currentDate = new Date().toISOString().split('T')[0];
                    // const agreementName = (currentDate + '_' + item.Company + '_' + item.Request_ID + '_' + item.TypeOfContract + '_' + item.Party2_agreement);

                    // if(!item.NameOfAgreement){
                    //   $("#agreement_name").val(agreementName);
                    // }
                    // else {
                    $("#agreement_name").val(item.NameOfAgreement);
                    // }


                    if (item.Confidential == "YES") {
                      checkbox.checked = true;
                    }
                    else {
                      checkbox.checked = false;
                    }



                  });
                }
              });

          }
        });
    }
    catch (err) {
      console.log(err.message);
    }

  }

  async getFileDetailsByFilter(libraryName, reqId) {
    try {
      const items = await sp.web.lists.getByTitle(libraryName).items
        .filter(`Request_Id eq '${reqId}'`)
        .select("File", "File/ServerRelativeUrl", "File/Name")
        .expand("File")
        .get();

      if (items.length > 0) {
        const item = items[0];
        const fileUrl = item.File.ServerRelativeUrl;
        const fileName = item.File.Name;
        return { fileUrl, fileName };
      }

      return null;
    } catch (error) {
      console.log(error);
      return null;
    }
  }

  //Load Timeline comments
  // public async load_comments(updateRequestID) {
  //   // let userEmail = "";
  //   const timeline = document.getElementById('commentTimeline');
  //   timeline.innerHTML = '';
  //   const CommentList = await sp.web.lists.getByTitle("Comments").items.select("RequestID,Comment,CommentBy,CommentDate").filter(`RequestID eq '${updateRequestID}'`).get();
  //   console.log('Commentlist',CommentList);
  //   // userEmail = CommentList[0].CommentBy;
  //   const users: any[] = await sp.web.siteUsers();
  //   // let userTitle = '';
  //   // users.forEach(user => {
  //     // if (user.Email === userEmail) {
  //     //   userTitle = user.Title;
  //     //   return;
  //     // }
  //   // });
  //   // if (userTitle === '') {
  //   //   console.log('User with email ' + userEmail + ' not found.');
  //   // }
  //   CommentList.forEach(item => {
  //     const comment = item.Comment;
  //     const commentDate = item.CommentDate;
  //     let userEmail = item.CommentBy;
  //     let userTitle = '';
  //     users.forEach(user => {
  //       if (user.Email === userEmail) {
  //         userTitle = user.Title;
  //         return;
  //       }
  //     });
  //     const timelineItem = document.createElement('li');
  //     timelineItem.className = 'timeline-item';
  //     timelineItem.innerHTML = `
  //       <div style="display: flex">
  //         <p style="margin-bottom: 0px">@${userTitle} -&nbsp;</p>
  //         ${commentDate}
  //       </div>
  //       <div>${comment}</div>
  //     `;
  //     timeline.appendChild(timelineItem);
  //   });

  //   timeline.scrollTop = timeline.scrollHeight;
  // }

  // async addComment(data) {
  //   try {
  //     const iar = await sp.web.lists.getByTitle("Comments").items.add(data);

  //     alert("Comment added succesfully.");
  //   }
  //   catch (e) {
  //     alert("An error occured." + e.message);
  //   }
  // }

  //Original
  // public async getSiteUsers() {
  //   const MasterList = await sp.web.lists.getByTitle("Contract_Request").items.getAll();
  //   let items: any[] = MasterList.filter(item => item.AssignedTo !== null);
  //   console.log("All contracts", items);

  //   var drp_users = document.getElementById("ownersList") as HTMLDataListElement;
  //   const users: [] = await sp.web.siteUsers();

  //   if (!drp_users) {
  //     console.error("Dropdown element not found");
  //     return;
  //   }

  //   // Clear the options of the datalist
  //   while (drp_users.options.length > 0) {
  //     drp_users.remove();
  //   }


  //   users.forEach(async (result: ISiteUserInfo) => {
  //     if (result.UserPrincipalName != null) {
  //       const groups = await sp.web.siteUsers.getById(result.Id).groups();
  //       groups.forEach((group) => {
  //         if (group.Title == "ENL_CMS_Owners") {
  //           var opt = document.createElement('option');

  //           opt.value = result.Title;

  //           opt.setAttribute('data-value', result.Email);
  //           opt.dataset; // Set the title as the display text
  //           drp_users.appendChild(opt);
  //         }
  //       });
  //     }
  //   });
  // }

  public async getSiteUsers() {
    const drp_users = document.getElementById("ownersList") as HTMLDataListElement;
    const allUsers = await sp.web.siteUsers();
    
    // Fetch all contracts
    const MasterList = await sp.web.lists.getByTitle("Contract_Request").items.getAll();
    const filteredMasterList: any[] = MasterList.filter(item => item.AssignedTo !== null);
    
    // Array to store users who belong to the "Owners" group
    const ownerUsers: ISiteUserInfo[] = [];
    
    // Fetch and filter users who belong to the "Owners" group
    for (const user of allUsers) {
      if (user.UserPrincipalName != null) {
        const groups = await sp.web.siteUsers.getById(user.Id).groups();
        const isOwner = groups.some(group => group.Title === "LegalLink_Internal_Owners");
        if (isOwner) {
          ownerUsers.push(user);
        }
      }
    }
    
    // Create a list that will contain the owner.Title, owner.email, and the number of contracts
    const ownersWithContractCount = ownerUsers.map(owner => {
      const contractCount = filteredMasterList.filter(item => item.AssignedTo === owner.Title).length;
      return {
        Title: owner.Title,
        Email: owner.Email,
        ContractCount: contractCount
      };
    });
    
    console.log("Owners with contract count:", ownersWithContractCount);
    
    // Populate the dropdown list
    if (drp_users) {
      // Clear the options of the datalist
      while (drp_users.options.length > 0) {
        drp_users.remove();
      }
    
      ownersWithContractCount.forEach(owner => {
        const opt = document.createElement('option');
        opt.value = `${owner.Title}`;
        opt.text = `Contract Count: ${owner.ContractCount}`;
        opt.setAttribute('data-value', owner.Email);
        drp_users.appendChild(opt);
      });
    } else {
      console.error("Dropdown element not found");
    }

    const css = `
    #assignOwners #ownersList {
      background-color: #f0f0f0;
      border: 1px solid #ccc;
      padding: 5px;
      width: 200px;
    }
    #assignOwners #ownersList option {
      padding: 5px;
      border-bottom: 1px solid #ccc;
    }
    #assignOwners #ownersList option:hover {
      background-color: #f2f2f2;
    }
  `;
  const style = document.createElement('style');
  style.type = 'text/css';
  style.appendChild(document.createTextNode(css));
  document.head.appendChild(style);

  }

  // public async getSiteUsers() {
  //   // Retrieve the master list of all contracts and filter it
  //   const MasterList = await sp.web.lists.getByTitle("Contract_Request").items.getAll();
  //   let items: any[] = MasterList.filter(item => item.AssignedTo !== null);
  //   console.log("All contracts", items);
  
  //   var drp_users = document.getElementById("ownersList") as HTMLDataListElement;
  //   const users: ISiteUserInfo[] = await sp.web.siteUsers();
  
  //   if (!drp_users) {
  //     console.error("Dropdown element not found");
  //     return;
  //   }
  
  //   // Clear the options of the datalist
  //   while (drp_users.options.length > 0) {
  //     drp_users.remove();
  //   }
  
  //   // Create a map to count the number of contracts assigned to each user
  //   const contractCounts = items.reduce((acc, item) => {
  //     const assignedTo = item.AssignedTo.Title;
  //     if (acc[assignedTo]) {
  //       acc[assignedTo]++;
  //     } else {
  //       acc[assignedTo] = 1;
  //     }
  //     return acc;
  //   }, {} as Record<string, number>);
  
  //   // Populate the dropdown list
  //   for (const result of users) {
  //     if (result.UserPrincipalName != null) {
  //       const groups = await sp.web.siteUsers.getById(result.Id).groups();
  //       const isOwner = groups.some(group => group.Title === "ENL_CMS_Owners");
  //       if (isOwner) {
  //         const contractCount = contractCounts[result.Title] || 0;
  //         var opt = document.createElement('option');
  //         opt.value = `${result.Title} (${contractCount} contracts)`;
  //         opt.setAttribute('data-value', result.Email);
  //         drp_users.appendChild(opt);
  //       }
  //     }
  //   }
  // }

  // private async assignOwners(id: any, item: any) {
  //   const list = sp.web.lists.getByTitle("Contract_Request");
  //   const i = await list.items.getById(id).update(item)
  //     .then(() => {
  //       alert("Task assigned successfully!");
  //     })
  //     .catch(err => {
  //       console.error(err);
  //     });
  // }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
