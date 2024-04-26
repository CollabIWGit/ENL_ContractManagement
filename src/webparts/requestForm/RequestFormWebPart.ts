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
import { MSGraphClient } from '@microsoft/sp-http';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { ISiteUserInfo } from '@pnp/sp/site-users/types';
import objMyCustomHTML from './Requestor_form';
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
let SideMenuUtils = new sideMenuUtils();

SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/4.1.0/css/bootstrap.min.css');
SPComponentLoader.loadCss('https://cdn.datatables.net/1.10.25/css/jquery.dataTables.min.css');


// require('../../Assets/scripts/styles/mainstyles.css');
require('./../../common/scss/style.scss');
require('./../../common/css/style.css');
require('./../../common/css/common.css');
require('../../../node_modules/@fortawesome/fontawesome-free/css/all.min.css');



var department;

export interface IRequestFormWebPartProps {
  description: string;
}

export default class RequestFormWebPart extends BaseClientSideWebPart<IRequestFormWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private graphClient: MSGraphClient;


  protected onInit(): Promise<void> {
    return new Promise<void>(async (resolve: () => void, reject: (error: any) => void): Promise<void> => {
      sp.setup({
        spfxContext: this.context
      });

      this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          this.graphClient = client;
          resolve();
        }, err => reject(err));
    });
  }

  public checkifUserIsAdmin(graphClient: MSGraphClient): Promise<void> {
    if (!graphClient) {
      return;
    }
    return new Promise((resolve, reject) => {
      graphClient.api(`/me/transitiveMemberOf/microsoft.graph.group?$count=true&$top=999`).get((errorGroup, groups: any, rawResponseGroup?: any) => {
        if (errorGroup) {
          console.log(errorGroup);
          return reject(errorGroup);
        }
        //console.log(groups);
        var groupList = groups.value;

        //console.log("GROOOUPS", groupList);

        if (groupList.filter(g => g.displayName == sharepointConfig.Groups.Requestor).length == 1) {

          department = "Requestor";
          console.log("You are a requestor");
          $(".legalDept").css("display", "none");

        }


        else if (groupList.filter(g => g.displayName == sharepointConfig.Groups.Owner).length == 1) {

          department = "Owner";
          console.log("You are an Owner");
          $(".legalDept").css("display", "none");

        }


        else if (groupList.filter(g => g.displayName == sharepointConfig.Groups.Despacher).length == 1) {

          department = "Despacher";
          console.log("You are an Despacher");
          $(".legalDept").css("display", "block");


        }
        else {
          department = "null";
          console.log("You are not in any group");
          $(".legalDept").css("display", "none");


        }

        return resolve();
      });
    });
  }

  public async checkCurrentUsersGroupAsync() {

    let groupList = await sp.web.currentUser.groups();

    if (groupList.filter(g => g.Title == sharepointConfig.Groups.Requestor).length == 1) {

      department = "Requestor";
      console.log("You are a requestor");
      $(".legalDept").css("display", "none");

    }


    else if (groupList.filter(g => g.Title == sharepointConfig.Groups.Owner).length == 1) {

      department = "Owner";
      console.log("You are an Owner");
      $(".legalDept").css("display", "none");

    }


    else if (groupList.filter(g => g.Title == sharepointConfig.Groups.Despacher).length == 1) {

      department = "Despacher";
      console.log("You are an Despacher");
      $(".legalDept").css("display", "block");

    }

    else {
      department = "null";
      console.log("You are not in any group");
      $(".legalDept").css("display", "none");


    }

  }

  addNewRow(table: any, party) {
    table.row
      .add([
        party
      ])
      .draw(false);

    party++
  }

  public render(): void {

    this.domElement.innerHTML = `
    <div class="wrapper d-flex align-items-stretch">
        <div id="nav-placeholder"></div>
    
        <div class="p-4 p-md-5 pt-3" id="content">   

            <div class="${styles.requestForm}" id="form_checklist">
    
  
                <form id="requestor_form" style="    border: 1px solid #ccc;
        padding: 20px;
        box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%);
        margin: 2em;
        padding: 2em;
        border-radius: 1rem;">
    
                    <div class="${styles['form-group']}">
    
                        <h2>Request Form</h2>
                        <h5 class="${styles.heading}">Request Details</h5>
    
                        <div class="${styles.grid}">
    
                            <div class="${styles['col-1-3']}">
                                <div class="${styles.controls}">
                                    <input type="text" id="requestor_name" class="floatLabel">
                                    <label for="requestor_name">Name of Requestor</label>
                                </div>
                            </div>
    
    
    
                            <div class="${styles['col-1-3']}">
                                <div class="${styles.controls}">
    
                                    <input type="text" class="floatLabel" id="status_title">
                                    <label for="status_title">Title</label>
                                </div>
                            </div>
    
                            <div class="${styles['col-1-3']}">
                                <div class="${styles.controls}">
    
                                    <input type="text" class="floatLabel" id="email">
                                    <label for="email">Email*</label>
                                </div>
    
                            </div>
    
                            </br>
    
    
    
                            <div class="${styles.grid}">
                                <div class="${styles['col-1-2']}">
                                    <div class="${styles.controls}">
                                        <i class="fa fa-sort"></i>
                                        <input type="text" class="floatLabel2" value="Please select.." id="enl_company" list='companies_folder' />
    
                                        <datalist id="companies_folder">
    
                                        </datalist>
                                        <label for="enl_company">Company*</label>
    
                                    </div>
                                </div>
    
    
                                <div class="${styles['col-1-2']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="department">
                                        <label for="department">Department</label>
                                    </div>
                                </div>
    
    
                            </div>
    
                            </br>
    
                            <div class="${styles.grid}">
    
    
                                <div class="${styles['col-1-2']}">
                                    <div class="${styles.controls}">
                                        <i class="fa fa-sort"></i>
                                        <input type="text" class="floatLabel2" id="requestFor" list='request_List' value="Please select.."/>
    
                                        <datalist id="request_List">
    
    
    
                                        </datalist>
                                        <label for="requestFor">Request For*</label>
    
                                    </div>
                                </div>
    
                                <div class="${styles['col-1-2']}">
                                    <div class="${styles.controls}" id="uploadFile">
                                        <input type="file" class="floatLabel" id="uploadContract">
                                        <label for="uploadContract">Upload Contract to Review</label>
                                    </div>
                                </div>
                            </div>
    
                            </br>
    
                            <div class="${styles.grid}">
    
    
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="party1">
                                        <label for="party1">Party 1 to the Agreement*</label>
    
                                    </div>
                                </div>
    
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="party2">
                                        <label for="party2">Party 2 to the Agreement*</label>
    
                                    </div>
                                </div>
    
                                <div class="${styles['col-1-3']}">
                                <div class="${styles.controls}">
                                <div style="position: relative;">
                                    <input type="text" class="floatLabel" id="other_parties">
                                    <label for="other_parties" class="">Other Parties*</label>
                                    <button class="add-button" id="addOtherParties" style="position: absolute;right: 12%;left: 73%;top: 38%;transform: translateY(-50%);width: 20%;">+</button>
                                </div>
                            </div>
                                </div>
                            </div>
    </br>
                                    <h5 class="${styles.heading}">Other Parties</h3>
                                    <div class="${styles.grid}">
                                        <div id="contentTable">
                                            <div class="w3-container" id="table">
                                                <div id="content3">
                                                    <div id="tblOtherParties" class="table-responsive-xl">
                                                        <div class="form-row">
                                                            <div class="col-xl-12">
                                                                <div id="other_parties_tbl">
                                                                <table id='tbl_other_Parties' class='table table-striped'">
                                                                <thead>
                                                                <tr>
                                                                    <th class="text-left">Other Party</th>              
                                                                </tr>
                                                                </thead>
                                                                <tbody id="tb_otherParties">
                                                               </tbody>
                                                                </table>
    
                                                                </div>
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                            </br>
    
                            <div class="${styles.grid}">
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="textarea" class="floatLabel" id="brief_desc">
                                        <label for="brief_desc">Brief Description of Transaction*</label>
    
                                    </div>
                                </div>
    
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="date" class="floatLabel" id="expectedCommenceDate">
                                        <label for="expectedCommenceDate">Expected Date of Commencement*</label>
    
                                    </div>
                                </div>
    
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                    <i class="fa fa-sort"></i>

                                        <input type="text" class="floatLabel2" value="Please select.." id="authority_to_approve_contract" list='authorityApproval'>
                                        <datalist id="authorityApproval">

                                        <option value="Yes">
                                        <option value="No">
    
    
                                        </datalist>
                                        <label for="authority_to_approve_contract">Authority to Approve Contract*</label>
    
                                    </div>
                                </div>
                            </div>
    
                            </br>
                            <div class="form-row">
    
                                <button type="button" class="buttoncss" id="saveToList"><i class="fa fa-refresh icon"></i>
                                    Save</button>
                                <button type="button" class="buttoncss">Cancel</button>
    
    
                            </div>

                            </br>

                            <div id="section_review_contract" style="display:none">
                              <div id="tbl_contract" style="margin-top: 1.5em;">
                              </div>

                            </div>
    
                            </br>
    
                            <div class="legalDept">

  
                                <h5 class="${styles.heading}">Comments</h3>
                                    <div class="${styles.grid}">
                                        <div id="contentTable">
                                            <div class="w3-container" id="table">
                                                <div id="content3">
                                                    <div id="tblcommentsD" class="table-responsive-xl">
                                                        <div class="form-row">
                                                            <div class="col-xl-12">
                                                                <div id="sp_comments_list_SectD">
    
                                                                </div>
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
    
    
    
    
                                    <h5 class="${styles.heading}">For Legal Department Only</h5>
    
    
    
                                    <div class="${styles.grid}">
                                        <div class="${styles['col-1-2']}">
                                            <div class="${styles.controls}">
                                                <i class="fa fa-sort"></i>
                                                <input type="text" class="floatLabel2" value="Please select.." id="assignedTo" list='ownersList' />
    
                                                <datalist id="ownersList">
    
    
    
                                                </datalist>
                                                <label for="assignedTo">Assigned To*</label>
    
                                            </div>
                                        </div>
    
                                        <div class="${styles['col-1-2']}">
                                            <div class="${styles.controls}">
                                                <input type="date" class="floatLabel" id="due_date">
                                                <label for="due_date">Due Date*</label>
    
                                            </div>
    
    
                                        </div>
    
    
                                    </div>
    
                                    <div class="${styles.grid}">
                                        <div class="${styles['col-1-2']}">
                                            <div class="${styles.controls}">
                                                <i class="fa fa-sort"></i>
                                                <input type="text" class="floatLabel2" id="contract_type"
                                                    list='typesOfContracts_list' value="Please select.." />
    
                                                <datalist id="typesOfContracts_list">
    
    
    
                                                </datalist>
                                                <label for="contract_type">Type of Contract*</label>
    
                                            </div>
                                        </div>
                                        <div class="${styles['col-1-2']}">
                                            <div class="${styles.controls}">
                                                <input type="text" class="floatLabel" id="agreement_name">
                                                <label for="agreement_name">Name of Agreement</label>
    
                                            </div>
    
                                        </div>
                                    </div>
    
                                    <div class="${styles.grid}">
                                        <div class="${styles['col-1-2']}">
                                            <div class="form-check" style="font-size: large;border: solid 0.5px #c6c6c6;height: 51px;">
                                                <input type="checkbox" class="" id="checkbox_confidential" name="checkbox_confidential" style="transform: scale(1.9);margin-top: 18px;margin-left: 151px;margin-right: 12px;accent-color: #f07e12;" value="YES">
    <label for="checkbox" class="form-check-label" style="font-family: Poppins,Arial,sans-serif;">       Confidential</label>
    
    
    
    
    </div>
                                        </div>
                                        <div class="${styles['col-1-2']}">
    
    
                                        </div>
                                    </div>
    
                                    </br>
    
                                    <div class="form-row">
    
                                        <button type="button" class="buttoncss" id="update_request"><i
                                                class="fa fa-refresh icon"></i> Save</button>
                                        <button type="button" class="buttoncss">Cancel</button>
    
    
                                    </div>
    
                                    </br>
    
    
                                    <h5 class="${styles.heading}">Comments</h5>
    
                                    <div class="${styles.grid}">
                                        <div class="${styles['col-1-1']}">
                                            <div class="${styles.controls}">
    
                                                <textarea type="text" class="floatLabel" id="comment"></textarea>
    
    
                                                <label for="comment">Comments</label>
    
                                            </div>
                                        </div>
    
    
    
    
                                    </div>
    
                                    <div class="${styles.grid}">
                                        <div class="${styles['col-1-2']}">
                                            <div class="${styles.controls}">
    
                                                <input type="text" class="floatLabel" id="commentsBy" />
    
    
                                                <label for="commentsBy">Comments by:</label>
    
                                            </div>
                                        </div>
                                        <div class="${styles['col-1-2']}">
                                            <div class="${styles.controls}">
                                                <input type="date" class="floatLabel" id="comment_date">
                                                <label for="comment_date">Date</label>
    
                                            </div>
    
                                        </div>
                                    </div>
    
                                    </br>
    
                                    <div class="form-row">
    
                                        <button type="button" class="buttoncss" id="addComment">Add Comment</button>
                                        <button type="button" class="buttoncss">Cancel</button>
    
    
                                    </div>
    
                            </div>
    
                        </div>
                </form>
            </div>
        </div>
    </div>
`;

    var table = $('#tbl_other_Parties').DataTable({
      info: false,
      responsive: true,
      pageLength: 5
    });


    document.querySelector('#addOtherParties').addEventListener('click', (event) => {
      event.preventDefault();
      this.addNewRow(table, $("#other_parties").val());
    });

    // // Automatically add a first row of data
    // this.addNewRow(table, counter);

    const button_update = document.getElementById('update_request');
    const icon_update = button_update.querySelector('.icon');

    const button_add = document.getElementById('saveToList');
    const icon_add = button_add.querySelector('.icon');

    const button_add_comment = document.getElementById('addComment');
    const icon_add_comment = button_add_comment.querySelector('.icon');

    var filename_add;
    var content_add;

    icon_update.classList.add('hide');
    icon_add.classList.add('hide');
    // icon_add_comment.classList.add('hide');


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

    require('./RequestorForm');

    this.load_companies();
    this.load_services();
    this.load_type_of_contracts();
    var requestID = this.getItemId();




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

    $("#uploadFile").css("display", "none");

    $("#saveToList").click(async (e) => {

      icon_add.classList.remove('hide');
      icon_add.classList.add('show');
      icon_add.classList.add('spinning');
      (document.getElementById('saveToList') as HTMLButtonElement).disabled = true;

      const library = "Contracts_ToReview";
      const folderPath = `/sites/ContractMgt/Contracts_ToReview/${$("#enl_company").val()}`;
      let listUri: string = '/sites/MyGed/Lists/Documents1';

      var dataParties = table.rows().data();

      var allParties = "";

      dataParties.each(function (value, index) {

        allParties += `${value};`;

      });


      var data = {
        Title: "Title",
        NameOfRequestor: $("#requestor_name").val(),
        Status_Title: $("#status_title").val(),
        Email: $("#email").val(),
        Company: $("#enl_company").val(),
        Department: $("#department").val(),
        RequestFor: $("#requestFor").val(),
        Party1_agreement: $("#party1").val(),
        Party2_agreement: $("#party2").val(),
        Others_parties: allParties,
        BriefDescriptionTransaction: $("#brief_desc").val(),
        ExpectedCommencementDate: $("#expectedCommenceDate").val(),
        AuthorityApproveContract: $("#authority_to_approve_contract").val(),
        AssigneeComment: $("#comment").val()
        // AssignedTo: $("#requestor_name").val(),
        // DueDate: $("#requestor_name").val(),
        // TypeOfContract: $("#requestor_name").val(),
        // NameOfAgreement: $("#requestor_name").val()
      };

      var RequestID;

      try {
        const iar = await sp.web.lists.getByTitle("Contract_Request").items.add(data)
          .then((iar) => {

            RequestID = iar.data.ID;

          });

      } catch (error) {
        console.error('Error adding item:', error);
        throw error;
      }

      console.log("LOG ID REVIEW", RequestID);

      if ($("#requestFor").val() == 'Review of Agreement') {

        await this.addFolderToDocumentLibrary(library, $("#enl_company").val())
          .then(async () => {
            try {
              await this.addFileToFolder2(folderPath, filename_add, content_add, RequestID.toString());
            }
            catch (e) {
              console.log(e.message);
            }
          });
      }


      alert("Request has been submitted successfully.");

      icon_add.classList.remove('spinning', 'show');
      icon_add.classList.add('hide');

      (document.getElementById('saveToList') as HTMLButtonElement).disabled = false;

    });

    $("#update_request").click(async (e) => {

      var ifConfidential = "NO";
      icon_update.classList.remove('hide');
      icon_update.classList.add('show');
      icon_update.classList.add('spinning');
      (document.getElementById('update_request') as HTMLButtonElement).disabled = true;

      if ($('input[name="checkbox_confidential"]').is(':checked')) {
        ifConfidential = "YES";
      }

      var data = {
        AssignedTo: $("#assignedTo").val(),
        DueDate: $("#due_date").val(),
        NameOfAgreement: $("#agreement_name").val(),
        TypeOfContract: $("#contract_type").val(),
        Confidential: ifConfidential
      };

      await this.assignOwners(parseInt(requestID), data);

      icon_update.classList.remove('spinning', 'show');
      icon_update.classList.add('hide');


      (document.getElementById('update_request') as HTMLButtonElement).disabled = false;

    });

    $("#requestFor").change(function (e) {

      var $el = $(this);

      var value = $el.val();

      if (value == 'Review of Agreement') {

        $("#uploadFile").css("display", "block");

      } else {

        $("#uploadFile").css("display", "none");

      }

    });

    $("#addComment").click(async (e) => {


      // icon_add_comment.classList.remove('hide');
      // icon_add_comment.classList.add('show');
      // icon_add_comment.classList.add('spinning');

      const currentUser = await sp.web.currentUser();

      const data = {

        Title: requestID,
        RequestID: requestID,
        Comment: $("#comment").val(),
        CommentBy: currentUser.UserPrincipalName,
        CommentDate: moment().format("DD/MM/YYYY HH:mm")
      };

      await this.addComment(data);

      // icon_add_comment.classList.remove('spinning', 'show');
      // icon_add_comment.classList.add('hide');

    });


    //this.require_libraries();

    SideMenuUtils.buildSideMenu(this.context.pageContext.web.absoluteUrl);

    // this.checkUsersGroup();
    this.checkCurrentUsersGroupAsync();
    this.renderRequestDetails(parseInt(requestID));
    this.renderComments(requestID);

    this.getSiteUsers();

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

  async addFolderToDocumentLibrary(libraryTitle, folderName) {
    try {
      // Initialize the PnP JS Library

      // Replace with the folder name you want to check

      const exists = await this.folderExists(libraryTitle, folderName);
      if (exists) {
        console.log(`Folder '${folderName}' exists.`);
      } else {
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

  async addComment(data) {

    try {
      const iar = await sp.web.lists.getByTitle("Comments").items.add(data);

      alert("Comment added succesfully.");

    }
    catch (e) {

      alert("An error occured." + e.message);

    }

  }

  async addFileToFolder2(folderPath, fileName, fileContent, requestId) {
    sp.web.getFolderByServerRelativeUrl(folderPath)
      .files.add(fileName, fileContent, false)
      .then((data) => {
        data.file.getItem()
          .then((item) => {
            item.update({
              Request_Id: requestId
            });
          })
      })
      .catch((error) => {
        alert("Error in uploading");
        console.log(error);
      })

  }

  // Example usage
  // Replace with actual file content

  async folderExists(libraryTitle, folderName) {
    try {
      // Initialize the PnP JS Library
      // Get the document library by title
      const library = sp.web.lists.getByTitle(libraryTitle);

      // Check if the folder exists
      const folder = await library.rootFolder.folders.getByName(folderName).select("Exists").get();

      return folder.Exists;
    } catch (error) {
      console.error(`Error checking folder existence: ${error.message}`);
      return false;
    }
  }

  // Example usage
  // Example usage

  private getItemId() {
    var queryParms = new URLSearchParams(document.location.search.substring(1));
    var myParm = queryParms.get("requestid");
    if (myParm) {
      return myParm.trim();
    }
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

  private renderComments(id: any) {
    try {
      const listContainer3: Element = this.domElement.querySelector('#sp_comments_list_SectD');

      let currentWebUrl = this.context.pageContext.web.absoluteUrl;
      let requestUrl = currentWebUrl.concat(`/_api/web/Lists/GetByTitle('Comments')/items?$select=Comment, CommentBy, CommentDate &$filter=(RequestID eq '${id}') `);
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
                  html += `
                            <table id='tbl_comment_attach_mgt_SectD' class='table table-striped'">
                                <thead>
                                <tr>
                                 <!--   <th class="text-left" style="width: 7%;">ID</th> -->
                                    <th class="text-left">Comment</th>              
                                  <!--  <th class="text-left">Department</th> -->
                                    <th class="text-left">Name of Person</th>
                                    <th class="text-left">Date/Time</th>
                                   <!-- <th class="text-left">Attachment(s)</th> -->
                                </tr>
                                </thead>
                                <tbody id="tb_contract_mgt_SectD">`;

                  doc.forEach((result: any) => {
                    const item = {

                      Comment: result.Comment,
                      CommentBy: result.CommentBy,
                      CommentDate: result.CommentDate,
                      // Date_time: result.DateTime,
                      // Attachments: result.AttachmentFiles
                    };

                    // console.log("Comments list:");
                    // console.log(item);

                    if (!Date.parse(item.CommentDate)) {
                      date = item.CommentDate;
                    }
                    else {
                      date = moment(new Date(item.CommentDate)).format("DD/MM/YYYY HH:mm")
                    }

                    html += `
                                <tr>
                                  
                                    <td class="text-left">${item.Comment}</td>
                                    <td class="text-left">${item.CommentBy}</td>
                                    <td class="text-left">${item.CommentDate}</td>
                                    
                       
                                </tr>
                                `;
                  });

                  html += `</tbody>
                            </table>`;
                  listContainer3.innerHTML += html;
                }
              })
              .then(() => {
                if (doc != null) {
                  $('#tbl_comment_attach_mgt_SectD').DataTable({
                    info: false,
                    responsive: true,
                    pageLength: 5,
                    order: [[0, 'desc']],
                  });
                }
              })
              .catch((error) => {
                console.log(error);
              });
          }
        });
    }
    catch (err) {
      console.log(err.message);
    }
  }

  private renderRequestDetails(id: any) {

    var checkbox = document.getElementById('checkbox_confidential') as HTMLInputElement;

    try {

      let currentWebUrl = this.context.pageContext.web.absoluteUrl;
      let requestUrl = currentWebUrl.concat(`/_api/web/Lists/GetByTitle('Contract_Request')/items?$select=NameOfRequestor, Status_Title, Email, Company, 
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
                              window.open(`https://frcidevtest.sharepoint.com/${fileDetails.fileUrl}`, '_blank');
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
                    $("#enl_company").val(item.Department);
                    $("#department").val(item.Department);
                    $("#requestFor").val(item.RequestFor);
                    $("#party1").val(item.Party1_agreement);
                    $("#party2").val(item.Party2_agreement);
                    $("#other_parties").val(item.Others_parties);
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

  public async checkUsersGroup() {

    await this.checkifUserIsAdmin(this.graphClient);

  }

  public async load_companies() {


    const drp_companies = document.getElementById("companies_folder") as HTMLSelectElement;

    if (!drp_companies) {
      console.error("Dropdown element not found");
      return;
    }

    const companies = await sp.web.lists.getByTitle('Companies').items
      .get();


    await Promise.all(companies.map(async (result) => {

      const opt = document.createElement('option');
      opt.value = result.Title;
      drp_companies.appendChild(opt);

    }));


  }

  public async load_type_of_contracts() {


    const drp_typeOfContracts = document.getElementById("contract_type") as HTMLSelectElement;

    if (!drp_typeOfContracts) {
      console.error("Dropdown element not found");
      return;
    }

    const companies = await sp.web.lists.getByTitle('Type of contracts').items
      .get();


    await Promise.all(companies.map(async (result) => {

      const opt = document.createElement('option');
      opt.value = result.Title;
      drp_typeOfContracts.appendChild(opt);

    }));


  }

  public async InsertRequestData(data) {

    var ID;

    try {
      const iar = await sp.web.lists.getByTitle("Contract_Request").items.add(data)
        .then((iar) => {

          ID = iar.data.ID;
          return ID;
        });

    } catch (error) {
      console.error('Error adding item:', error);
      throw error;
    }

  }

  public async getSiteUsers() {

    var drp_users = document.getElementById("ownersList");

    const users: [] = await sp.web.siteUsers();

    users.forEach(async (result: ISiteUserInfo) => {

      if (result.UserPrincipalName != null) {

        const groups = await sp.web.siteUsers.getById(result.Id).groups();

        groups.forEach((group) => {
          console.log("GROUP", group.Id, group.Title);

          if (group.Title == "ENL_CMS_Owners") {
            console.log("USER", result.Id, result.Email);
            var opt = document.createElement('option');
            opt.appendChild(document.createTextNode(result.Email));
            opt.value = result.Email;
            drp_users.appendChild(opt);
          }
          // Perform further actions with the group information
        });

      }

    });

  }

  public async load_services() {
    const drp_companies = document.getElementById("request_List") as HTMLSelectElement;

    if (!drp_companies) {
      console.error("Dropdown element not found");
      return;
    }

    const companies = await sp.web.lists.getByTitle('ENL_Services').items
      .get();


    await Promise.all(companies.map(async (result) => {

      const opt = document.createElement('option');
      opt.value = result.Title;
      drp_companies.appendChild(opt);

    }));

  }

  public async getItems() {
    const items: any[] = await sp.web.lists.getByTitle("Type Of Contracts").items();
    console.log(items);
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
