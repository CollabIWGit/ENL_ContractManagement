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

    const absoluteUrl = this.context.pageContext.web.absoluteUrl;

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

              <p id="contractStatus" style="color: green; position: absolute; top: 0; right: 0; font-size: x-large">In Progress</p>

              <div class="${styles['form-group']}">
                <h2 style="color: #888;">Request Form</h2>

                <fieldset>
                  <legend id='requestorDetailsLegend'>YOUR DETAILS</legend>

                  <div id="yourDetailsSection" class="${styles.grid}">
                    
                    <div class="${styles['col-1-3']}">
                      <div class="${styles.controls}">
                        <label for="requestor_name">Name of Requestor</label>
                        <input type="text" id="requestor_name" >
                      </div>
                    </div>
                    <div class="${styles['col-1-3']}">
                      <div class="${styles.controls}">
                        <label for="status_title">Title</label>
                        <input type="text"  id="status_title">
                      </div>
                    </div>

                    <div class="${styles['col-1-3']}">
                      <div class="${styles.controls}">
                        <label for="email">Email*</label>
                        <input type="text"  id="email" required>
                      </div>
                    </div>
                    <div class="${styles['col-1-3']}">
                      <div class="${styles.controls}">
                        <label for="phone_number">Phone Number*</label>
                        <input type="text"  id="phone_number" required>
                      </div>
                    </div>
            
                    <div class="${styles['col-1-3']}">
                        <div class="${styles.controls}">
                        <label for="enl_company">Company*</label>
                        <input type="text"  placeholder="Please select.." id="enl_company" list='companies_folder' required/>
                        <datalist id="companies_folder"></datalist>
                      </div>
                    </div>
                    <div class="${styles['col-1-3']}">
                      <div class="${styles.controls}">
                        <label for="department">Department</label>
                        <input type="text"  id="department">
                      </div>
                    </div>
                  
                  </div>

                </fieldset>

                <fieldset>
                  <legend>HOW CAN WE ASSIST?</legend>

                  <div class="${styles.grid}">
                    <div class="${styles['col-1-4']}">
                      <div class="${styles.controls}">
                        <label for="requestFor">Request For*</label>
                        <input type="text"  id="requestFor" list='request_List' placeholder="Please select.." />
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
                        <textarea type="text"  id="brief_desc"></textarea>
                      </div>
                    </div>
                  </div>

                </fieldset>

                <fieldset>
                  <legend>PARTIES TO THE AGREEMENT</legend>
                  
                  <div class="${styles.grid}" style="width: 100%; display: flex;">
                    <div style="width: 100%;">
                      <div class="${styles['col-1-3']}">
                        <div class="${styles.controls}">
                          <label for="party1"">Name of Party 1(ENL-Rogers group side)</label>
                          <input type="text"  id="party1">
                        </div>
                      </div>

                      <div class="${styles['col-1-3']}">
                        <div class="${styles.controls}">
                          <label for="party2">Name of Party 2</label>
                          <input type="text"  id="party2">
                        </div>
                      </div>

                      <div class="${styles['col-1-3']}">
                        <div class="${styles.controls}">
                          <div style="position: relative;">
                            <label for="other_parties" class="">Other Parties*</label>
                            <input type="text"  id="other_parties">
                            <button class="${styles.addPartiesButton}" id="addOtherParties">+</button>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                
                  <div class="${styles.grid}">
                    <div id="contentTable">
                      <div class="w3-container" id="table">
                        <div id="content3">
                          <div id="tblOtherParties" class="table-responsive-xl">
                            <div class="form-row">
                              <div class="col-xl-12">
                                <div id="other_parties_tbl">
                                  <table id='tbl_other_Parties' class='table table-striped'>
                                    <thead>
                                      <tr>
                                        <th class=" text-left">Other Party</th>
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
                  </div>

                </fieldset>

                <fieldset>
                  <legend>OTHER INFO</legend>

                  <div class="${styles.grid}">
                    <div class="${styles['col-1-3']}">
                      <div class="${styles.controls}">
                        <label for="expectedCommenceDate">Expected Date of Commencement*</label>
                        <input type="date"  id="expectedCommenceDate">
                      </div>
                    </div>

                    <div class="${styles['col-1-3']}">
                      <div class="${styles.controls}">
                        <label for="authority_to_approve_contract">Authority to Approve Contract*</label>
                        <input type="text"  placeholder="Please select.." id="authority_to_approve_contract" list='authorityApproval'>
                        <datalist id="authorityApproval">
                          <option value="Yes">
                          <option value="No">
                        </datalist>
                      </div>
                    </div>

                    <div class="${styles.grid}">
                      <div class="${styles['col-1-2']}">
                        <div class="${styles.controls}" id="authorisedApproverDiv" >
                          <label for="authorisedApprover">Name of authorised approver</label>
                          <input type="text"  id="authorisedApprover">
                        </div>
                      </div>
                    </div>

                  </div>

                </fieldset>

              <div id="requestorSubmit" class="submitBtnDiv">
                <button type="submit" id="saveToList"><i class="fa fa-refresh icon" style="display: none;"></i>Save</button>
              </div>

              <fieldset id="legalDeptSection">
                <legend class="${styles.legalLegend}">FOR LEGAL DEPARTMENT ONLY</legend>
              </fieldset>

              <div id="section_review_contract">
                <div id="tbl_contract" style="margin-top: 1.5em;"></div>
              </div>

            </form>

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
    SideMenuUtils.buildSideMenu(absoluteUrl);

    let nameInput = document.getElementById('requestor_name')  as HTMLInputElement;
    let emailInput = document.getElementById('email')  as HTMLInputElement;

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

    //Display Legal Department
    if(currentRole === ('DespatcherAssign')){
      $('#legalDeptSection').show();

      document.getElementById('legalDeptSection').innerHTML += `
        <div class="legalDept">
          <div class="${styles.grid}">
            <div class="${styles['col-1-2']}">
              <div id="assignOwners" class="${styles.controls}">
                <label for="assignedTo">Assigned To*</label>
                <input type="text"  placeholder="Please select.." id="assignedTo" list='ownersList' />
                <datalist id="ownersList" style="color: blue"></datalist>
              </div>
            </div>

            <div class="${styles['col-1-2']}">
              <div class="${styles.controls}">
                <label for="due_date">Due Date*</label>
                <input type="date"  id="due_date">
              </div>
            </div>
          </div>

          <div class="${styles.grid}">
            <div class="${styles['col-1-2']}">
              <div class="${styles.controls}">
                <label for="contract_type">Type of Contract*</label>
                <input type="text"  id="contract_type" list='typesOfContracts_list' placeholder="Please select.." />
                <datalist id="typesOfContracts_list"></datalist>
              </div>
            </div>

            <div class="${styles['col-1-2']}">
              <div class="${styles.controls}">
                <label for="agreement_name">Name of Agreement</label>
                <input type="text"  id="agreement_name">
              </div>
            </div>
          </div>

          <div class="assignBtnDiv">
            <button type="submit" id="saveToList">Assign</button>
          </div>
          <br>
        </div>
      `;

      this.getSiteUsers();

      $("#assignedTo").bind('input', () => {
        const shownVal = (document.getElementById("assignedTo") as HTMLInputElement).value;
        // var shownVal = document.getElementById("name").value;
  
        const value2send = (document.querySelector<HTMLSelectElement>(`#ownersList option[value='${shownVal}']`) as HTMLSelectElement).dataset.value;
        ownerTitle = value2send;
        console.log("LOGG", value2send);
        //  $("#created_by").val(value2send);
      });
    }

    //Display Cancel Button
    if(currentRole === 'RequestorUpdate' || currentRole === 'DespatcherAssign'){
      document.getElementById('requestorSubmit').innerHTML += `
        <button id="cancelRequest" type="button">Cancel</button>
      `;
    }

    //Disable Name and Email
    if(currentRole === 'RequestorCreate' || currentRole === 'OwnerCreate' || currentRole === 'DespatcherCreate'){
      nameInput.disabled = true;
      emailInput.disabled = true;
    }

    //Retrieve Request ID
    const urlParams = new URLSearchParams(window.location.search);
    const updateRequestID = urlParams.get('requestid');
    let onBehalf: boolean = false;
    //New Request
    if (!updateRequestID) {
      this.setRequestorDetails(onBehalf);
    }
    //Update Request
    else {
      this.renderRequestDetails(updateRequestID);

      document.getElementById('saveToList').textContent = 'Update';
    }

    var ownerTitle;

    //OtherParties datatable
    var table = $('#tbl_other_Parties').DataTable({
      info: false,
      // responsive: true,
      pageLength: 5
    });

    $("#authorisedApproverDiv").css("display", "none");

    $("#authority_to_approve_contract").change(function (e) {
      var $el = $(this);
      var value = $el.val();
      if (value == 'Yes') {
        $("#authorisedApproverDiv").css("display", "block");
      } else {
        $("#authorisedApproverDiv").css("display", "none");
      }
    });

    this.load_companies(); //Companies list
    this.load_services(); //Request For list

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
      if ($("#other_parties").val() == "") {
        alert("Please enter a value");
      }
      else {
        this.addNewOtherPartiesRow(table, $("#other_parties").val());
      }
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

    //ENL CMS GROUP Principal IDs
    const ENL_CMS_Group_IDs = {
      LegalLink_Requestors: 49,
      LegalLink_Despatchers: 46,
      LegalLink_Internal_Owners: 48,
      LegalLink_External_Owners: 47,
      LegalLink_Directors_View: 50
    };

    //Create new request
    var newRequestID;
    $("#saveToList").click(async (e) => {

      this.checkFormIsValid();

      (document.getElementById('saveToList') as HTMLButtonElement).disabled = true;

      //Other Parties data
      var dataParties = table.rows().data();
      var allOtherParties = "";
      dataParties.each(function (value, index) {
        allOtherParties += `${value};`;
      });

      const checkbox = document.getElementById('checkbox_confidential') as HTMLInputElement;
      const confidentialValue = checkbox.checked ? 'YES' : 'NO';

      //Form data
      var formData = {
        // Title: "Title",
        NameOfRequestor: $("#requestor_name").val(),
        Status_Title: $("#status_title").val(),
        Email: $("#email").val(),
        Phone_Number: $("#phone_number").val(),
        Company: $("#enl_company").val(),
        Department: $("#department").val(),
        RequestFor: $("#requestFor").val(),
        Confidential: confidentialValue,
        BriefDescriptionTransaction: $("#brief_desc").val(),
        Party1_agreement: $("#party1").val(),
        Party2_agreement: $("#party2").val(),
        Others_parties: allOtherParties,
        ExpectedCommencementDate: $("#expectedCommenceDate").val().toString(),
        AuthorityApproveContract: $("#authority_to_approve_contract").val(),
        AuthorisedApprover: $("#authorisedApprover").val()
        // AssigneeComment: $("#comment").val()
        // AssignedTo: $("#requestor_name").val(),
        // DueDate: $("#requestor_name").val(),
        // TypeOfContract: $("#requestor_name").val(),
        // NameOfAgreement: $("#requestor_name").val()
      };

      console.log(formData);

      if(currentRole === 'RequestorCreate' || currentRole === 'OwnerCreate' || currentRole === 'DespatcherCreate'){
        try {
          //Add item to Contract Request
          const iar = await sp.web.lists.getByTitle("Contract_Request").items.add(formData)
            .then((iar) => {
              newRequestID = iar.data.ID;
            });
          console.log(newRequestID);

          var dataC = {
            Request_ID: newRequestID.toString()
          };
          console.log(dataC);

          //Add item to Contract Details
          try {
            await sp.web.lists.getByTitle("Contract_Details").items.add(dataC);
          }
          catch (error) {
            console.error('Error adding item in contract_Details:', error);
            throw error;
          }

          //Root Document library
          // const libraryTitle = "Contracts";
          // const library = sp.web.lists.getByTitle(libraryTitle);
          // const companyFolderName = $("#enl_company").val() as string;
          // const contractFolderName = newRequestID;
          // //Final path in which document will be stored
          // const folderPath = `/sites/ContractMgt/Contracts/${companyFolderName}/${contractFolderName}`;

          // //Create contract folder
          // await library.rootFolder.folders.getByName(companyFolderName).folders.add(newRequestID);
          // console.log(`Contract Folder '${contractFolderName}' created successfully.`);





          // if ($("#requestFor").val() == 'Review of Agreement') {
          //   await this.addFolderToDocumentLibrary(library, $("#enl_company").val(), newRequestID.toString())
          //     .then(async () => {
          //       try {
          //         await this.addFileToFolder2(folderPath, filename_add, content_add, newRequestID.toString());
          //       }
          //       catch (e) {
          //         console.log(e.message);
          //       }
          //     });
          // }

          alert(`Request ${newRequestID} has been submitted successfully.`);

          // if(currentRole === 'DespatcherCreate'){
          //   Navigation.navigate(`${this.context.pageContext.web.absoluteUrl}/SitePages/Requestor-Form.aspx?requestid=${newRequestID}`, true);
          // }
          // else{
          //   Navigation.navigate(`${this.context.pageContext.web.absoluteUrl}/SitePages/Dashboard.aspx`, true);
          // }
        }
        catch (error) {
          console.error('Error adding item:', error);
          throw error;
        }
      }
    });

    //Retrieve requestDigest
    var requestDigest;
    await this.getFormDigest(absoluteUrl).then(function (data) {
      requestDigest = data.d.GetContextWebInformation.FormDigestValue;
    });

    //Retrieve Email of current user
    var currentUserEmail;
    await this.getCurrentUserEmail(absoluteUrl)
    .then(response => {
        currentUserEmail = response.d.Email;
    })
    .catch(error => {
        console.error('Error retrieving current user email:', error);
    });

    currentUserEmail = 'samg@frcidevtest.onmicrosoft.com'

    //Retrieve PrincipalId for currrent user
    var principalIdUser;
    await this.getPrincipalIdForUserByEmail(absoluteUrl, currentUserEmail)
      .then(principalId => {
          principalIdUser = principalId;
      })
      .catch(error => {
          console.error('Error fetching PrincipalId:', error);
      });


    const caseFolderPath = "/sites/ContractMgt/Contracts/FRCI/194";

    this.consoleFolderUsers(absoluteUrl, caseFolderPath);

    //Assign Permissions
    try {

      //Break role inheritance
      await this.breakRoleInheritance(absoluteUrl, requestDigest, caseFolderPath);
      console.log("Inheritance broken");
  
      const response = await this.getFolderPermissions(absoluteUrl, caseFolderPath);
      const roleAssignments = response.d.results;
      console.log(roleAssignments);
      //Listing Folder Permissions
      const users = roleAssignments
        .filter(roleAssignment => roleAssignment.Member.PrincipalType === 1 || 8)
        .map(roleAssignment => roleAssignment.Member.Title);
      console.log('Users with access to the folder:', users);

      //Remove all existing role assignments
      for (let roleAssignment of roleAssignments) {
        await this.removeRoleAssignments(absoluteUrl, requestDigest, caseFolderPath, roleAssignment.PrincipalId);
      }
      console.log("All role assignments removed");

      //Add permission per user
      this.addRoleAssignmentUser(absoluteUrl, requestDigest, caseFolderPath, principalIdUser, permissionLevels.Edit);
  
      //Add ENL_CMS_Despachers group with appropriate permissions
      const ENL_CMS_Despachers_principalId = 14;
      const roleDefId = 1073741827;
      await this.addRoleAssignment(absoluteUrl, requestDigest, caseFolderPath, ENL_CMS_Despachers_principalId, roleDefId);
      console.log("ENL_CMS_Despachers group added with permissions");

      const response2 = await this.getFolderPermissions(absoluteUrl, caseFolderPath);
      const roleAssignments2 = response2.d.results;
      //Listing Folder Permissions
      const users2 = roleAssignments2
        .filter(roleAssignments2 => roleAssignments2.Member.PrincipalType === 1 || 8)
        .map(roleAssignments2 => roleAssignments2.Member.Title);
      console.log('Users with access to the folder2:', users2);
  
    } catch (error) {
      console.error("Error updating folder permissions:", error);
    }

  }

  //RequestDigest
  private getFormDigest(absoluteUrl) {
    return $.ajax({
        url: absoluteUrl + "/_api/contextinfo",
        method: "POST",
        headers: {
            "Accept": "application/json; odata=verbose"
        }
    });
  }

  //Retrieve Current User Email
  private getCurrentUserEmail(absoluteUrl) {
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
  private getPrincipalIdForUserByEmail(absoluteUrl, currentUserEmail) {
    const restUrl = `${absoluteUrl}/_api/web/SiteUserInfoList/items?$filter=EMail eq '${encodeURIComponent(currentUserEmail)}'`;

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
            throw new Error(`User with email ${currentUserEmail} not found.`);
        }
    });
  }

  //Display current users on folder on console
  private async consoleFolderUsers(absoluteUrl, caseFolderPath) {
    const response = await this.getFolderPermissions(absoluteUrl, caseFolderPath);
      const roleAssignments = response.d.results;
      console.log(roleAssignments);
      //Listing Folder Permissions
      const users = roleAssignments
        .filter(roleAssignment => roleAssignment.Member.PrincipalType === 1 || 8)
        .map(roleAssignment => roleAssignment.Member.Title);
      console.log('Users with access to the folder:', users);
  }

  //Adding Role Assignment for User
  private async addRoleAssignmentUser(absoluteUrl, requestDigest, folderUrl, principalId, roleDefId) {
    try {
        // Step 1: Break role inheritance (optional if already broken)
        await this.breakRoleInheritance(absoluteUrl, requestDigest, folderUrl);

        // Step 2: Add role assignment for the specified user
        const restUrl = `${absoluteUrl}/_api/web/getFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields/RoleAssignments/addroleassignment(principalid=${principalId}, roledefid=${roleDefId})`;
        
        await $.ajax({
            url: restUrl,
            method: "POST",
            headers: {
                "Accept": "application/json; odata=verbose",
                "X-RequestDigest": requestDigest
            }
        });

        console.log(`Role assigned to user with Principal ID ${principalId}`);
    } catch (error) {
        console.error('Error adding role assignment:', error);
        throw error;
    }
  }

  //Retrieve folder user and group access
  private getFolderPermissions(absoluteUrl, folderUrl) {
    const restUrl = `${absoluteUrl}/_api/web/getFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields/RoleAssignments?$expand=Member,Member/Users,Member/Owner,RoleDefinitionBindings`;
    console.log(restUrl);
  
    return $.ajax({
      url: restUrl,
      method: "GET",
      headers: {
        "Accept": "application/json; odata=verbose"
      }
    });
  }

  //Break Permissions
  private breakRoleInheritance(absoluteUrl, requestDigest, folderUrl) {
    const restUrl = `${absoluteUrl}/_api/web/getFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields/breakroleinheritance(copyRoleAssignments=false)`;

    return $.ajax({
      url: restUrl,
      method: "POST",
      headers: {
        "Accept": "application/json; odata=verbose",
        "X-RequestDigest": requestDigest
      }
    });
  }

  private removeRoleAssignments(absoluteUrl, requestDigest, folderUrl, roleAssignmentId) {
    const restUrl = `${absoluteUrl}/_api/web/getFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields/roleassignments/removeroleassignment(principalid=${roleAssignmentId})`;
    return $.ajax({
      url: restUrl,
      method: "POST",
      headers: {
        "Accept": "application/json; odata=verbose",
        "X-RequestDigest": requestDigest
      }
    });
  }
  
  private addRoleAssignment(absoluteUrl, requestDigest, folderUrl, principalId, roleDefId) {
    const restUrl = `${absoluteUrl}/_api/web/getFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields/roleassignments/addroleassignment(principalid=${principalId}, roledefid=${roleDefId})`;
    return $.ajax({
      url: restUrl,
      method: "POST",
      headers: {
        "Accept": "application/json; odata=verbose",
        "X-RequestDigest": requestDigest
      }
    });
  }

  //Check Form Validity
  private checkFormIsValid(){
    document.getElementById("requestor_form").addEventListener("submit", function(event) {
      event.preventDefault(); // Prevent the default form submission
    
      const form = event.target as HTMLFormElement;
    
      if (form.checkValidity() === false) {
        event.stopPropagation();
        form.classList.add("was-validated");
      } else {
        console.log('Form is valid');
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
          return currentRole = 'OwnerView'; //Internal Owner or External Owner
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

  //New row for other parties
  addNewOtherPartiesRow(table: any, party) {
    table.row.add([party]).draw(false);
    $("#other_parties").val("");
  }

  public async load_companies() {
    const drp_companies = document.getElementById("companies_folder") as HTMLSelectElement;
    if (!drp_companies) {
      console.error("Dropdown element not found");
      return;
    }
    const companies = await sp.web.lists.getByTitle('Companies').items.get();

    await Promise.all(companies.map(async (result) => {
      const opt = document.createElement('option');
      opt.value = result.Title;
      drp_companies.appendChild(opt);
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


  async addFolderToDocumentLibrary(libraryTitle, companyFolderName, contractFolderName) {
    const library = sp.web.lists.getByTitle(libraryTitle);

    try {
      const exists = await this.folderExists(library, companyFolderName, contractFolderName);

      //None exists
      if(exists === "noneExist"){
        //Create company folder
        await library.rootFolder.folders.add(companyFolderName);
        console.log(`Company Folder '${companyFolderName}' created successfully.`);
        //Create contract folder
        await library.rootFolder.folders.getByName(companyFolderName).folders.add(contractFolderName);
        console.log(`Contract Folder '${contractFolderName}' created successfully.`);
      }
      else if(exists === "companyOnly"){
        //Create contract folder
        await library.rootFolder.folders.getByName(companyFolderName).folders.add(contractFolderName);
        console.log(`Contract Folder '${contractFolderName}' created successfully.`);
      }
      else if(exists === "allExist"){
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

  async folderExists(library, companyFolderName, contractFolderName) {

    let existResponse = "";

    // Check if company folder exists
    try {
      const companyFolder = await library.rootFolder.folders.getByName(companyFolderName).select("Exists").get();
      console.log("Company folder exists");
      //Company folder exists
      if(companyFolder.Exists){
        try{
          const contractFolder = await library.rootFolder.folders.getByName(companyFolderName).folders.getByName(contractFolderName).select("Exists").get();
          if(contractFolder.Exists){
            console.log("Contract folder exists");
            existResponse = "allExist"; 
            return existResponse;
          }
        }
        catch(error){
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

  //If file name already exists, file will not be uploaded
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

  // async addFileToFolder2(folderPath, fileName, fileContent, requestId, userEmail, role) {
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
  private renderRequestDetails(id: any) {
    var checkbox = document.getElementById('checkbox_confidential') as HTMLInputElement;
    try {
      let currentWebUrl = this.context.pageContext.web.absoluteUrl;
      let requestUrl = currentWebUrl.concat(`/_api/web/Lists/GetByTitle('Contract_Request')/items?$select=NameOfRequestor, Status_Title, Email, Phone_Number, Company, 
      Department, Party1_agreement, Party2_agreement, Others_parties, BriefDescriptionTransaction, ExpectedCommencementDate, AssignedTo, Owner, AssigneeComment, DueDate,
      Confidential, AuthorityApproveContract, DueDate, TypeOfContract, NameOfAgreement, RequestFor &$filter=(ID eq '${id}') `);
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

                      NameOfRequestor: result.NameOfRequestor,
                      StatusTitle: result.Status_Title,
                      Email: result.Email,
                      Phone_Number: result.Phone_Number,
                      Company: result.Company,
                      Department: result.Department,
                      Party1_agreement: result.Party1_agreement,
                      Party2_agreement: result.Party2_agreement,
                      Others_parties: result.Others_parties,
                      BriefDescriptionTransaction: result.BriefDescriptionTransaction,
                      ExpectedCommencementDate: result.ExpectedCommencementDate,
                      AuthorityApproveContract: result.AuthorityApproveContract,
                      DueDate: result.DueDate,
                      TypeOfContract: result.TypeOfContract,
                      NameOfAgreement: result.NameOfAgreement,
                      RequestFor: result.RequestFor,
                      AssignedTo: result.AssignedTo,
                      Owner: result.Owner,
                      AssigneeComment: result.AssigneeComment,
                      Confidential: result.Confidential

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

                    if (item.RequestFor == 'Review of Agreement') {

                      $("#section_review_contract").css("display", "block");

                      this.getFileDetailsByFilter('Contracts', id)
                        .then((fileDetails) => {
                          if (fileDetails) {
                            console.log("File URL:", fileDetails.fileUrl);
                            console.log("File Name:", fileDetails.fileName);

                            let html: string = '<div class="form-row">';
                            html += `
                              <div class="col-md-12 table-responsive">
                                <table class="table" id="table1" style="box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%);margin-bottom: 2em;">
                                  <thead class="thead-dark">
                                    <tr>
                                      <th class="th-lg" scope="col">Contract</th>
                                      <th scope="col">View</th>
                                    </tr>
                                  </thead>
                            `;

                            html += `
                              <tbody>
                                <tr>
                                  <td scope="row">${fileDetails.fileName}</td>
                                  <td>
                                    <ul class="list-inline m-0">
                                      <li class="list-inline-item">
                                        <button id="btn_view" class="btn btn-secondary btn-sm rounded-circle" type="button"  data-toggle="tooltip" data-placement="top" title="View" style="display: none;"><i class="fas fa-eye"></i></button>
                                      </li>
                                      <li class="list-inline-item">
                                        <button id="modalActivate" class="btn btn-secondary btn-sm rounded-circle" type="button"  data-toggle="modal" data-target="#exampleModalPreview" style="display: block;width: auto;""><i class="fas fa-eye"></i></button>
                                      </li>
                                    </ul>
                                  </td>
                                </tr>
                              </tbody>
                            </table>
                          </div>
                        </div>
                            `;

                            var baseUrl = '';
                            const listContainer: Element = this.domElement.querySelector('#tbl_contract');
                            listContainer.innerHTML = html;

                            $("#modalActivate").click(() => {
                              //console.log("calling viewing ...");
                              // this.submit_main();
                              window.open(`ms-word:ofv|u|https://frcidevtest.sharepoint.com/${fileDetails.fileUrl}`, '_blank');
                              // Navigation.navigate(`${urlChoice}`, '_blank');
                            });


                          } else {
                            console.log("Item not found.");
                          }
                        })
                        .catch((error) => {
                          console.log(error);
                        });
                    }

                    $("#requestor_name").val(item.NameOfRequestor);
                    $("#status_title").val(item.StatusTitle);
                    $("#email").val(item.Email);
                    $("#phone_number").val(item.Phone_Number),
                      $("#enl_company").val(item.Company);
                    $("#department").val(item.Department);
                    $("#requestFor").val(item.RequestFor);
                    $("#party1").val(item.Party1_agreement);
                    $("#party2").val(item.Party2_agreement);

                    // Other parties populate table
                    // $("#other_parties").val(item.Others_parties);


                    if(item.Others_parties !== null && item.Others_parties !== "") {
                      var othersPartiesVal = item.Others_parties;
                      othersPartiesVal = othersPartiesVal.replace(/;+$/, '');
                      var otherPartiesArray = othersPartiesVal.split(';').reverse();
                      var tbody = document.getElementById('tb_otherParties');
                      tbody.innerHTML = '';
                      otherPartiesArray.forEach(function (value) {
                        var tr = document.createElement('tr');
                        var td = document.createElement('td');
                        td.textContent = value.trim();
                        tr.appendChild(td);
                        tbody.appendChild(tr);
                      });
                    }




                    $("#brief_desc").val(item.BriefDescriptionTransaction);
                    $("#expectedCommenceDate").val(item.ExpectedCommencementDate);
                    $("#authority_to_approve_contract").val(item.AuthorityApproveContract);

                    $("#assignedTo").val(item.AssignedTo);
                    $("#contract_type").val(item.Owner);
                    $("#due_date").val(item.DueDate);
                    $("#agreement_name").val(item.NameOfAgreement);


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
        const isOwner = groups.some(group => group.Title === "ENL_CMS_Owners");
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

  private async assignOwners(id: any, item: any) {
    const list = sp.web.lists.getByTitle("Contract_Request");
    const i = await list.items.getById(id).update(item)
      .then(() => {
        alert("Task assigned successfully!");
      })
      .catch(err => {
        console.error(err);
      });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
