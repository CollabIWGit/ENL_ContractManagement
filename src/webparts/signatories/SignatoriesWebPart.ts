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
import styles from './SignatoriesWebPart.module.scss';
import * as strings from 'SignatoriesWebPartStrings';
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
let absoluteUrl = '';

export interface ISignatoriesWebPartProps {
  description: string;
}

export default class SignatoriesWebPart extends BaseClientSideWebPart<ISignatoriesWebPartProps> {

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
  
  </style>
  
    `;

    //HTML
    this.domElement.innerHTML += `

      <div class="main-container" id="content">

        <div id="nav-placeholder" class="left-panel"></div>

        <div id="middle-panel" class="middle-panel">

          <button id="minimizeButton"></button>

          <div class="${styles.signatories}" id="form_checklist">

            <form id="signature_form" style="position: relative; width: 100%;">

              <div class="${styles['form-group']}">
                <h2 style="color: #888;">Signature Form</h2>

                <fieldset class="${styles.fieldsetSig}">
                  <legend class="${styles.legendSig}">Internal Signatory</legend>

                  <div class="${styles.grid}">
                    
                    <div class="${styles['col-1-2']}">
                      <div class="${styles.controls}">
                        <label for="InternalSignatory_1">Signatory 1</label>
                        <input type="text" id="InternalSignatory_1" required>
                      </div>
                    </div>
                    <div class="${styles['col-1-2']}">
                      <div class="${styles.controls}">
                        <label for="InternalSignatory_2">Signatory 2</label>
                        <input type="text"  id="InternalSignatory_2" required>
                      </div>
                    </div>
                  
                  </div>
                </fieldset>

                <fieldset class="${styles.fieldsetSig}">
                  <legend class="${styles.legendSig}">External Signatory</legend>

                  <div class="${styles.grid}">
                    
                    <div class="${styles['col-1-2']}">
                      <div class="${styles.controls}">
                        <label for="ExternalSignatory_1">Signatory 1</label>
                        <input type="text" id="ExternalSignatory_1" required>
                      </div>
                    </div>
                    <div class="${styles['col-1-2']}">
                      <div class="${styles.controls}">
                        <label for="ExternalSignatory_2">Signatory 2</label>
                        <input type="text"  id="ExternalSignatory_2" required>
                      </div>
                    </div>
                  
                  </div>
                </fieldset>

                <div id="signatureSubmit" class="${styles.submitBtnDiv}">
                  <button type="submit" id="sendForSignature"><i class="fa fa-refresh icon" style="display: none;"></i>Send For Signature</button>
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
    SideMenuUtils.buildSideMenu(absoluteUrl, departments);

    //Retrieve Request ID
    const urlParams = new URLSearchParams(window.location.search);
    const updateRequestID = urlParams.get('requestid');

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

    if(updateRequestID){
      document.getElementById("signature_form").addEventListener("submit", async (event) => {
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
          const signatoryData = {
            InternalSignatory1: $("#InternalSignatory_1").val(),
            InternalSignatory2: $("#InternalSignatory_2").val(),
            ExternalSignatory1: $("#ExternalSignatory_1").val(),
            ExternalSignatory2: $("#ExternalSignatory_2").val(),
            Signatory_Date: new Date().toLocaleDateString()
          };

          try {
            const items = await sp.web.lists.getByTitle("Contract_Details").items.filter(`Request_ID eq '${updateRequestID}'`).get();

            if (items.length > 0) {
              for (const item of items) {
                await sp.web.lists.getByTitle("Contract_Details").items.getById(item.Id).update(signatoryData);
              }
            } else {
              console.log("No items found with the specified Request_Id.");
            }

            const contractRequestList = sp.web.lists.getByTitle("Contract_Request");
              await contractRequestList.items.getById(Number(updateRequestID)).update({
                ContractStatus: 'SentForSignature'
            });

            alert(`Request for signature has been sent.`);
          }
          catch (error)
          {
            console.error("Error updating item:", error);
          }
          
        }
      });
    }
    else{
      (document.getElementById('sendForSignature') as HTMLButtonElement).disabled = true;
    }
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

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
