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

var department;
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

  //Render everything
  public async render(): Promise<void> {

    //HTML CSS of form
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

</style>

      <div class="main-container" id="content">

        <div id="nav-placeholder" class="left-panel"></div>

        <div id="middle-panel" class="middle-panel">

          <button id="minimizeButton"></button>

          <div class="${styles.requestForm}" id="form_checklist">

            <form id="requestor_form" style="position: relative; width: 100%;">

              <p id="contractStatus" style="color: green; position: absolute; top: 0; right: 0;">Status: In Progress</p>

              <div class="${styles['form-group']}">
                <h2>Request Form</h2>
                <h5 class="${styles.heading}">Your details</h5>

                <div id="yourDetailsSection" class="${styles.grid}" style="opacity: 1; max-height: 1000px; transition: opacity 0.5s ease, max-height 0.5s ease;">
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
                      <input type="text"  id="email">
                    </div>
                  </div>
                  <div class="${styles['col-1-3']}">
                    <div class="${styles.controls}">
                      <label for="phone_number">Phone Number*</label>
                      <input type="text"  id="phone_number">
                    </div>
                  </div>
          
                  <div class="${styles['col-1-3']}">
                      <div class="${styles.controls}">
                      <label for="enl_company">Company*</label>
                      <input type="text"  placeholder="Please select.." id="enl_company" list='companies_folder'/>
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

                <h5 class="${styles.heading}">How can we assist?</h5>
                <div class="${styles.grid}">
                  <div class="${styles['col-1-2']}">
                    <div class="${styles.controls}">
                      <label for="requestFor">Request For*</label>
                      <input type="text"  id="requestFor" list='request_List' placeholder="Please select.." />
                      <datalist id="request_List"></datalist>
                    </div>
                  </div>

                  <div class="${styles['col-1-2']}">
                    <div class="${styles.controls}" id="uploadFile" style="display: none;">
                      <label for="uploadContract">Upload Contract to Review</label>
                      <input type="file"  id="uploadContract">
                    </div>
                  </div>
                </div>

                <div class="${styles.grid}">
                  <div style="display: flex; align-items: center; font-size: large;border: none;height: 51px; margin-bottom: 2px;">
                    <label for="checkbox" class="form-check-label" style="font-family: Poppins, Arial, sans-serif;"> Confidential</label>
                    <input type="checkbox" id="checkbox_confidential" name="checkbox_confidential" style="transform: scale(1.9); margin-right: 12px; margin-left: 1rem; accent-color: #f07e12;" value="YES">
                    <p style="font-size: smaller; margin-left: 12px; margin-bottom: 4px;">[click if you wish this assignment to be known to Chief Legal Executive only]</p>
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

                <h5 class="${styles.heading}">Parties to the agreement</h5>
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
              </div>
              
              <h5 class="${styles.heading}">Other Parties</h3>
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

              <h5 class="${styles.heading}">Other info</h5>
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

              <div id="legalDeptSection" class="legalDept">

              </div>

              <br>

              <div id="requestorSubmit" class="form-row">
                <button type="button" class="buttoncss" id="saveToList"><i class="fa fa-refresh icon" style="display: none;"></i>Save</button>
                <button type="button" class="buttoncss">Cancel</button>
              </div>

              <br>

              <div id="section_review_contract">
                <div id="tbl_contract" style="margin-top: 1.5em;"></div>
              </div>

            </form>

          </div>

        </div>

      </div>
    `;

    SPComponentLoader.loadScript('https://ajax.googleapis.com/ajax/libs/jquery/3.6.3/jquery.min.js')
      .then(() => {
        // return SPComponentLoader.loadScript('//cdnjs.cloudflare.com/ajax/libs/popper.js/2.9.2/cjs/popper.min.js') 
      })
      .then(() => {
        return SPComponentLoader.loadScript('//cdn.jsdelivr.net/npm/bootstrap@5.0.1/dist/js/bootstrap.min.js')
      })
      .then(() => {
        console.log("Scripts loaded successfully");
      })
      .catch(error => {
        console.error("Error loading scripts: " + error);
      });

    const urlParams = new URLSearchParams(window.location.search);
    const updateRequestID = urlParams.get('requestid');

    if (!urlParams.has('requestid')) {

      this.setRequestorDetails();

    }



    //On Update Request
    if (updateRequestID) {

      // const middlePanelID = document.getElementById('middle-panel');
      // middlePanelID.style.marginRight = '27%';
      // const rightPanelID = document.getElementById('rightPanel');
      // rightPanelID.style.width = '27%';

      //Generate Timeline
      // document.getElementById('rightPanel').innerHTML = `
      // <div style="width: 100%; height:100%; background: white; padding-bottom: 30%;">
      //   <div class="timelineHeader">
      //     <p style="margin-bottom: 0px;">Timeline</p>
      //   </div>
      //   <ul id="commentTimeline" class="timeline"></ul>
      //   <div class="comment-box">
      //     <textarea id="comment" class="comment-input" placeholder="Add your comment..."></textarea>
      //     <button id="addComment">Add Comment</button>
      //   </div>
      // </div>
      // `;

      this.renderRequestDetails(updateRequestID);


      // this.load_comments(updateRequestID);
    }

    //New Request
    // else {
    // $('#rightPanel').hide();
    // const middlePanelID = document.getElementById('middle-panel');
    // middlePanelID.style.marginRight = '0%';
    // middlePanelID.style.width = '83%';
    // $('#contractStatus').hide();
    // }

    await this.checkCurrentUsersGroupAsync();
    var ownerTitle;


    if (department === "Despatcher") {
      document.getElementById('legalDeptSection').innerHTML = `
        <h5 class="${styles.heading}">For Legal Department Only</h5>
        <div class="${styles.grid}">
          <div class="${styles['col-1-2']}">
            <div class="${styles.controls}">
              <label for="assignedTo">Assigned To*</label>
              <input type="text"  placeholder="Please select.." id="assignedTo" list='ownersList' />
              <datalist id="ownersList"></datalist>
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




    //OtherParties datatable
    var table = $('#tbl_other_Parties').DataTable({
      info: false,
      responsive: true,
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

    //CSS for labels
    require('./RequestorForm');

    //Generate Side Menu
    SideMenuUtils.buildSideMenu(this.context.pageContext.web.absoluteUrl);

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

    var filename_add;
    var content_add;

    //Process uploaded file
    $('#uploadContract').on('change', () => {
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
          middlePanelID.style.marginLeft = '0%'
          minimizeButtonID.style.left = '0%';
        }
      }
    });

    var newRequestID;



    //Create new request
    $("#saveToList").click(async (e) => {

      try {
        if (department == "Despatcher") {
          var ifConfidential = "NO";
          // icon_update.classList.remove('hide');
          // icon_update.classList.add('show');
          // icon_update.classList.add('spinning');
          (document.getElementById('saveToList') as HTMLButtonElement).disabled = true;

          if ($('input[name="checkbox_confidential"]').is(':checked')) {
            ifConfidential = "YES";
          }

          const despatcher = await sp.web.currentUser();


          var dataDespatch = {
            AssignedTo: $("#assignedTo").val(),
            DueDate: $("#due_date").val(),
            NameOfAgreement: $("#agreement_name").val(),
            TypeOfContract: $("#contract_type").val(),
            Confidential: ifConfidential,
            DespatcherEmail: despatcher.Email,
            OwnerEmail:  ownerTitle
          };

          await this.assignOwners(parseInt(updateRequestID), dataDespatch);

          // icon_update.classList.remove('spinning', 'show');
          // icon_update.classList.add('hide');


          (document.getElementById('saveToList') as HTMLButtonElement).disabled = false;
          Navigation.navigate(`${this.context.pageContext.web.absoluteUrl}/SitePages/Dashboard.aspx`, true);
        }
        else {


          // icon_add.classList.remove('hide');
          // icon_add.classList.add('show');
          // icon_add.classList.add('spinning');

          (document.getElementById('saveToList') as HTMLButtonElement).disabled = true;

          //Other Parties data
          var dataParties = table.rows().data();
          var allOtherParties = "";
          dataParties.each(function (value, index) {
            allOtherParties += `${value};`;
          });

          //Form data
          var data = {
            // Title: "Title",
            NameOfRequestor: $("#requestor_name").val(),
            Status_Title: $("#status_title").val(),
            Email: $("#email").val(),
            Phone_Number: $("#phone_number").val(),
            Company: $("#enl_company").val(),
            Department: $("#department").val(),
            RequestFor: $("#requestFor").val(),
            BriefDescriptionTransaction: $("#brief_desc").val(),
            Party1_agreement: $("#party1").val(),
            Party2_agreement: $("#party2").val(),
            Others_parties: allOtherParties,
            ExpectedCommencementDate: $("#expectedCommenceDate").val(),
            AuthorityApproveContract: $("#authority_to_approve_contract").val(),
            AuthorisedApprover: $("#authorisedApprover").val(),
            Confidential: $("#checkbox_confidential").val()
            // AssigneeComment: $("#comment").val()
            // AssignedTo: $("#requestor_name").val(),
            // DueDate: $("#requestor_name").val(),
            // TypeOfContract: $("#requestor_name").val(),
            // NameOfAgreement: $("#requestor_name").val()
          };

          try {
            const iar = await sp.web.lists.getByTitle("Contract_Request").items.add(data)
              .then((iar) => {
                newRequestID = iar.data.ID;
              });
            console.log(newRequestID);

            var dataC = {
              Request_ID: newRequestID.toString()
            }

            console.log(dataC);
            try {
              await sp.web.lists.getByTitle("Contract_Details").items.add(dataC);
            }
            catch (error) {
              console.error('Error adding item in contract_Details:', error);
              throw error;
            }
          }
          catch (error) {
            console.error('Error adding item:', error);
            throw error;
          }

          const library = "Contracts_ToReview";
          const folderPath = `/sites/ContractMgt/Contracts_ToReview/${$("#enl_company").val()}/${newRequestID}`;

          if ($("#requestFor").val() == 'Review of Agreement') {
            await this.addFolderToDocumentLibrary(library, $("#enl_company").val())
              .then(async () => {
                try {
                  await this.addFileToFolder2(folderPath, filename_add, content_add, newRequestID.toString());
                }
                catch (e) {
                  console.log(e.message);
                }
              });
          }

          alert("Request has been submitted successfully.");

          // icon_add.classList.remove('spinning', 'show');
          // icon_add.classList.add('hide');

          (document.getElementById('saveToList') as HTMLButtonElement).disabled = false;

          Navigation.navigate(`${this.context.pageContext.web.absoluteUrl}/SitePages/Dashboard.aspx`, true);
        }
      }
      catch (e) {

        console.log("ERROR", e.message);
      }

    });

    //Add comment button
    // $("#addComment").click(async (e) => {
    //   console.log("Test New Comment");
    //   // icon_add_comment.classList.remove('hide');
    //   // icon_add_comment.classList.add('show');
    //   // icon_add_comment.classList.add('spinning');

    //   const currentUser = await sp.web.currentUser();

    //   const data = {

    //     Title: updateRequestID,
    //     RequestID: updateRequestID,
    //     Comment: $("#comment").val(),
    //     CommentBy: currentUser.UserPrincipalName,
    //     CommentDate: moment().format("DD/MM/YYYY HH:mm")
    //   };

    //   console.log(data);

    //   await this.addComment(data);

    //   // icon_add_comment.classList.remove('spinning', 'show');
    //   // icon_add_comment.classList.add('hide');

    //   this.load_comments(updateRequestID);

    //   $("#comment").val("");

    // });
  }

  //New row for other parties
  addNewOtherPartiesRow(table: any, party) {
    table.row.add([party]).draw(false);
    $("#other_parties").val("")
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

  public async setRequestorDetails() {

    const requestor = await sp.web.currentUser();

    $("#requestor_name").val(requestor.Title);
    $("#email").val(requestor.Email);


  }

  // libraryTitle = Contracts_ToReview, foldername = enlCompany.val
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
                      date = moment(new Date(item.DueDate)).format("DD/MM/YYYY HH:mm")
                    }

                    if (item.RequestFor == 'Review of Agreement') {

                      $("#section_review_contract").css("display", "block");

                      this.getFileDetailsByFilter('Contracts_ToReview', id)
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

  public async getSiteUsers() {
    var drp_users = document.getElementById("ownersList") as HTMLDataListElement;
    const users: [] = await sp.web.siteUsers();

    if (!drp_users) {
      console.error("Dropdown element not found");
      return;
    }

    // Clear the options of the datalist
    while (drp_users.options.length > 0) {
      drp_users.remove();
    }


    users.forEach(async (result: ISiteUserInfo) => {
      if (result.UserPrincipalName != null) {
        const groups = await sp.web.siteUsers.getById(result.Id).groups();
        groups.forEach((group) => {
          if (group.Title == "ENL_CMS_Owners") {
            var opt = document.createElement('option');

            opt.value = result.Title;

            opt.setAttribute('data-value', result.Email);
            opt.dataset; // Set the title as the display text
            drp_users.appendChild(opt);
          }
        });
      }
    });
  }

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
