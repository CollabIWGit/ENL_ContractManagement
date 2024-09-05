import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';
import styles from './ContractFormWebPart.module.scss';
import * as strings from 'ContractFormWebPartStrings';
import { sp, List, IItemAddResult, UserCustomActionScope, Items, Item, ITerm, ISiteGroup, ISiteGroupInfo } from "@pnp/sp/presets/all";
import { MSGraphClientV3 } from '@microsoft/sp-http';
import * as moment from 'moment';
import 'datatables.net';
import * as $ from 'jquery';
import {
    SPHttpClient,
    SPHttpClientResponse
} from '@microsoft/sp-http';


import { sideMenuUtils } from "../../common/utils/sideMenuUtils";
let SideMenuUtils = new sideMenuUtils();
let departments = [];

SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/4.1.0/css/bootstrap.min.css');
SPComponentLoader.loadCss('https://cdn.datatables.net/1.10.25/css/jquery.dataTables.min.css');


// require('../../Assets/scripts/styles/mainstyles.css');
require('./../../common/scss/style.scss');
require('./../../common/css/style.css');
require('./../../common/css/common.css');
require('../../../node_modules/@fortawesome/fontawesome-free/css/all.min.css');

let currentUser: string;
let absoluteUrl = '';
let baseUrl = '';
export interface IContractFormWebPartProps {
    description: string;
}

export default class ContractFormWebPart extends BaseClientSideWebPart<IContractFormWebPartProps> {

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

    public async render(): Promise<void> {
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
                    width: 27%;
                    right: 0;
                    padding-right: 20px;
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
                table.dataTable thead th {
                    text-align: center!important;
                }

                table.displayContractTable thead th {
                    text-align: center!important;
                }

                table.dataTable tbody td {
                    text-align: center!important;
                }
            </style>
        `;
        //HTML
        this.domElement.innerHTML += `

            <div class="main-container" id="content">

                <div id="nav-placeholder" class="left-panel"></div>

                <div id="middle-panel" class="middle-panel">
                    <button id="minimizeButton"></button>
                    <div class="${styles.contractForm}" id="form_checklist">
                        <form id="contract_details_form" style="position: relative; width: 100%;">
                                                        
                            <div class="${styles.grid}">

                                <div class="${styles['form-group']}">
                                    <h2 style="color: #888;">Contract Form</h2>
                                    <h5 class="${styles.heading}">Contract Details</h5>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <label for="contract_name"
                                            title="Naming convention applies"
                                        >Name of Contract*
                                        </label>
                                        <input type="text" id="contract_name" required autocomplete="off">
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <label for="internal_ref_num">Internal Reference Number*</label>
                                        <input type="text" id="internal_ref_num" required autocomplete="off">
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <label for="contractStatus">Status*</label><span id="contractStatus_error" class="${styles.errorSpan}">Please select a valid status.</span>
                                        <input type="text" id="contractStatus" placeholder="Please select.." list="statuses" autocomplete="off">
                                        <datalist id="statuses">
                                            <option value="To Be Assigned"></option>
                                            <option value="To Be Accepted"></option>
                                            <option value="WIP"></option>
                                            <option value="Approved by Requestor"></option>
                                            <option value="Sent for Signature"></option>
                                            <option value="Signed"></option>
                                            <option value="Effective"></option>
                                            <option value="Terminated"></option>
                                            <option value="Cancelled"></option>
                                        </datalist>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <p> </p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <p style="margin-bottom: 0px;">Name of Party*</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <p 
                                            title="Name of person within client company responsible for this contract and who will receive alerts related to this contract."
                                        style="margin-bottom: 0px;">Name of person responsible</p>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <p>Party 1 (ENL-Rogers side)*</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" placeholder="Please select.." id="partyENL_Rogers_name" list='companies_folder' required autocomplete="off">
                                        <datalist id="companies_folder"></datalist>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" id="partyENL_Rogers_person" required autocomplete="off">
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <p>Party 2*</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" id="party2_name" required autocomplete="off">
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" id="party2_person" required autocomplete="off">
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <p>Party 3</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" id="party3_name" autocomplete="off">
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" id="party3_person" autocomplete="off">
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <p>Party 4</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" id="party4_name" autocomplete="off">
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" id="party4_person" autocomplete="off">
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <p>Party 5</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" id="party5_name" autocomplete="off">
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" id="party5_person" autocomplete="off">
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <label for="contractType">Type of Contract*</label>
                                        <input type="text" id="contractType" placeholder="Please select.." list='contractTypeList' required autocomplete="off">
                                        <datalist id="contractTypeList"></datalist>                                        
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <div style="position: relative;">
                                            <label for="contract_value">Value of Contract*</label>
                                            <div style="display: flex;">
                                                <select id="contract_currency" class="${styles.addPartiesButton}" required>
                                                    <option value="MUR">MUR</option>
                                                    <option value="USD">USD</option>
                                                    <option value="EUR">EUR</option>
                                                    <option value="GBP">GBP</option>
                                                </select>
                                                <input type="number" id="contract_value" min="0" required autocomplete="off">
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <label for="approvedBy">Approved by</label>
                                        <input type="text" id="approvedBy" autocomplete="off">
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <label for="signatureDate">Date of Signature*</label>
                                        <input type="date" id="signatureDate" required>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <label for="effectDate">Date of Effect*</label>
                                        <input type="date" id="effectDate" required>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <label for="expiryDate">Date of Expiry</label>
                                        <input type="date" id="expiryDate">
                                    </div>
                                </div>
                                
                                <div class="${styles['col-1-3']}" class="${styles.controls}">
                                    <label>Term of Contract*</label>
                                    <div class="${styles.termOfContract}">
                                        <div style="display: flex; width: 50%;">
                                            <label style="width: 50%;" for="termOfContractIndefinite">Indefinite</label>
                                            <input type="radio" name="termOfContractRBG" value="Indefinite" required>
                                        </div>
                                        <div style="display: flex; width: 50%;">
                                            <label style="width: 50%;" for="termOfContractFixedTerm">Fixed Term</label>
                                            <input type="radio" name="termOfContractRBG" value="Fixed Term">
                                        </div>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <div style="position: relative;">
                                        <label for="initial_duration_value" 
                                            title="Please select the duration type and enter the duration value">Duration (Initial)
                                        </label>
                                        <input type="text" id="initial_duration_value" autocomplete="off">
                                        <select id="initial_duration_type" class="${styles.addPartiesButton}" title="Please select...">
                                            <option value="" disabled selected>Please select...</option>
                                            <option value="days">Days</option>
                                            <option value="months">Months</option>
                                            <option value="years">Years</option>
                                            <option value="other">Other</option>
                                        </select>
                                        </div>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}" style="margin-bottom: 5px;">
                                    <div class="${styles.controls}">
                                        <div style="position: relative;">
                                            <label for="renewed_duration_value">Duration (Renewed)</label>
                                            <input type="text" id="renewed_duration_value" autocomplete="off">
                                            <select id="renewed_duration_type" class="${styles.addPartiesButton}" title="Please select...">
                                                <option value="" disabled selected>Please select...</option>
                                                <option value="days">Days</option>
                                                <option value="months">Months</option>
                                                <option value="years">Years</option>
                                                <option value="other">Other</option>
                                            </select>
                                        </div>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <div style="position: relative;">
                                            <label for="noticePeriodTermination_value">Notice period for termination*</label>
                                            <input type="text" id="noticePeriodTermination_value" required autocomplete="off">
                                            <select id="noticePeriodTermination_type" class="${styles.addPartiesButton}" title="Please select..." required>
                                                <option value="" disabled selected>Please select...</option>
                                                <option value="days">Days</option>
                                                <option value="months">Months</option>
                                                <option value="years">Years</option>
                                                <option value="other">Other</option>
                                            </select>
                                        </div>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <div style="position: relative;">
                                        <label for="noticePeriodExtensionRenewal_value">Notice Period for renewal/extension</label>
                                            <input type="text" id="noticePeriodExtensionRenewal_value" autocomplete="off">
                                            <select id="noticePeriodExtensionRenewal_type" class="${styles.addPartiesButton}" title="Please select...">
                                                <option value="" disabled selected>Please select...</option>
                                                <option value="days">Days</option>
                                                <option value="months">Months</option>
                                                <option value="years">Years</option>
                                                <option value="other">Other</option>
                                            </select>
                                        </div>
                                    </div>
                                </div>

                                <div class="${styles['col-2-3']}">
                                    <div class="${styles.controls}">
                                        <label for="salientTerms">Salient Terms</label>
                                        <textarea type="text" id="salientTerms" style="height: 10rem; margin-bottom: 10px;"></textarea>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <label for="renewalTerms">Renewal Terms</label>
                                        <textarea type="text" id="renewalTerms" style="height: 10rem; margin-bottom: 10px;"></textarea>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <label for="jurisdiction">Jurisdiction*</label>
                                        <input type="text" id="jurisdiction" required autocomplete="off">
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <label for="disputeResolution">Dispute Resolution</label>
                                        <input type="text" id="disputeResolution" autocomplete="off">
                                    </div>
                                </div>

                                <div class="${styles['col-1']}">
                                    <div class="${styles.controls}">
                                        <label>For addendum/amendment agreement only: details of initial contract</label>
                                    </div>
                                </div>
                                <div class="${styles['col-2-3']}">
                                    <div class="${styles.controls}">
                                        <label for="addendaName">Name</label>
                                        <input type="text" id="addendaName" autocomplete="off">
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <label for="addendaDate">Date</label>
                                        <input type="date" id="addendaDate">
                                    </div>
                                </div>

                                <table class="${styles['custom-table']}">
                                    <tr>
                                        <td class="${styles['col-1-6']}">
                                        </td>
                                        <td class="${styles['col-03']}">
                                            <div class="${styles.controls}"><p style="text-align: center">Yes</p></div>
                                        </td>
                                        <td class="${styles['col-03']}">
                                            <div class="${styles.controls}"><p style="text-align: center">No</p></div>
                                        </td>
                                        <td></td>
                                    </tr>
                                    <tr>
                                        <td class="${styles['col-1-6']}">
                                            <div class="${styles.controls}"><p>Electronically Signed</p></div>
                                        </td>
                                        <td class="${styles['col-03']}">
                                            <div class="${styles.controls}"><input type="radio" name="eleSignedRBG" value="Yes"></div>
                                        </td>
                                        <td class="${styles['col-03']}">
                                            <div class="${styles.controls}"><input type="radio" name="eleSignedRBG" value="No"></div>
                                        </td>
                                        <td></td>
                                    </tr>
                                    <tr>
                                        <td class="${styles['col-1-6']}">
                                            <div class="${styles.controls}"><p>Exclusivity*</p></div>
                                        </td>
                                        <td class="${styles['col-03']}">
                                            <div class="${styles.controls}"><input type="radio" name="exclusivityRBG" value="Yes" required></div>
                                        </td>
                                        <td class="${styles['col-03']}">
                                            <div class="${styles.controls}"><input type="radio" name="exclusivityRBG" value="No"></div>
                                        </td>
                                        <td></td>
                                    </tr>
                                    <tr>
                                        <td class="${styles['col-1-6']}">
                                            <div class="${styles.controls}"><p>Confidentiality*</p></div>
                                        </td>
                                        <td class="${styles['col-03']}">
                                            <div class="${styles.controls}"><input type="radio" name="confidentialityRBG" value="Yes" required></div>
                                        </td>
                                        <td class="${styles['col-03']}">
                                            <div class="${styles.controls}"><input type="radio" name="confidentialityRBG" value="No"></div>
                                        </td>
                                        <td></td>
                                    </tr>
                                    <tr>
                                        <td class="${styles['col-1-6']}">
                                            <div class="${styles.controls}"><p>Related Party Transaction</p></div>
                                        </td>
                                        <td class="${styles['col-03']}">
                                            <div class="${styles.controls}"><input type="radio" name="relatedPartyTransactionRBG" value="Yes"></div>
                                        </td>
                                        <td class="${styles['col-03']}">
                                            <div class="${styles.controls}"><input type="radio" name="relatedPartyTransactionRBG" value="No"></div>
                                        </td>
                                        <td></td>
                                    </tr>
                                    <tr>
                                        <td class="${styles['col-1-6']}">
                                            <div class="${styles.controls}"><p>Compliance SEM/DEM/FSC</p></div>
                                        </td>
                                        <td class="${styles['col-03']}">
                                            <div class="${styles.controls}"><input type="radio" name="complianceRBG" value="Yes"></div>
                                        </td>
                                        <td class="${styles['col-03']}">
                                            <div class="${styles.controls}"><input type="radio" name="complianceRBG" value="No"></div>
                                        </td>
                                        <td class="${styles['col-1']}">
                                            <div class="${styles['col-1']}" style="position: absolute; transform: translateY(-50%);">
                                                <label for="ComplianceDetails">Details</label>
                                                <div class="${styles.controls} ${styles['col-1-2']}">
                                                    <input type="text" id="ComplianceDetails" style="height: 50%;" autocomplete="off">
                                                </div>
                                            </div>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="${styles['col-1-6']}">
                                            <div class="${styles.controls}"><p>Original Copy Filed*</p></div>
                                        </td>
                                        <td class="${styles['col-03']}">
                                            <div class="${styles.controls}"><input type="radio" name="originalCopyFiledRBG" value="Yes" required></div>
                                        </td>
                                        <td class="${styles['col-03']}">
                                            <div class="${styles.controls}"><input type="radio" name="originalCopyFiledRBG" value="No"></div>
                                        </td>
                                        <td></td>
                                    </tr>
                                    <tr>
                                        <td class="${styles['col-1-6']}">
                                            <div class="${styles.controls}"><p>External Legal advice sought</p></div>
                                        </td>
                                        <td class="${styles['col-03']}">
                                            <div class="${styles.controls}"><input type="radio" name="ELAS_RBG" value="Yes"></div>
                                        </td>
                                        <td class="${styles['col-03']}">
                                            <div class="${styles.controls}"><input type="radio" name="ELAS_RBG" value="No"></div>
                                        </td>
                                        <td class="${styles['col-1']}">
                                            <div class="${styles['col-1']}" style="position: absolute; transform: translateY(-50%);">
                                                <label for="ELAS_Name">Name</label>
                                                <div class="${styles.controls} ${styles['col-1-2']}">
                                                    <input type="text" id="ELAS_Name" style="height: 50%;" autocomplete="off">
                                                </div>
                                            </div>
                                        </td>
                                    </tr>
                                </table>

                                <div class="${styles['col-1-2']}">
                                    <div class="${styles.controls}">
                                        <label for="breachDetails">Details of Breach (if any)</label>
                                        <textarea type="text" id="breachDetails" style="height: 10rem;"></textarea>
                                    </div>
                                </div>
                                <div class="${styles['col-1-2']}">
                                    <div class="${styles.controls}">
                                        <label for="litigationDetails">Details of Litigation (if any)</label>
                                        <textarea type="text" id="litigationDetails" style="height: 10rem;"></textarea>
                                    </div>
                                </div>

                                <div class="${styles['col-1']}">
                                    <div class="${styles.controls}">
                                        <label>Last Updated</label>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <label for="lastUpdatedOn">On</label>
                                        <input type="text" id="lastUpdatedOn" disabled>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <label for="lastUpdatedBy">By</label>
                                        <input type="text" id="lastUpdatedBy" disabled>
                                    </div>
                                </div>

                                <div class="${styles['col-1']}" style="margin-top: 1rem;">
                                    <div class="form-row ${styles.submitBtnDiv}">
                                        <button type="submit" class="buttoncss" id="saveToList"><i class="fa fa-refresh icon" style="display: none"></i>Save</button>
                                        <button type="button" class="buttoncss">Cancel</button>
                                        <button type="button" class="buttoncss" id="sendForSignature">Send For Signature</button>
                                    </div>
                                </div>

                                <div id="section_review_contract">
                                    <div id="tbl_contract" style="margin-top: 1.5em;"></div>
                                </div>
                            
                            </div>

                        </form>
                    </div>
                </div>

            </div>
        `;

        //<input type="file" id="fileUpload" accept=".pdf,.doc,.docx" />    Upload file to be sent for signature

        absoluteUrl = this.context.pageContext.web.absoluteUrl;
        baseUrl = absoluteUrl.split('/sites')[0];

        SideMenuUtils.buildSideMenu(absoluteUrl, departments);

        //#region 
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
        //#endregion

        //Update Contract
        const urlParams = new URLSearchParams(window.location.search);
        const updateRequestID = urlParams.get('requestid');
        const contractDetails = await sp.web.lists.getByTitle("Contract_Request").items.select("NameOfAgreement","Company","NameOfRequestor","Owner","TypeOfContract","Party2_agreement","OwnerEmail","Email","ContractStatus").filter(`ID eq ${updateRequestID}`).get();
        const contractStatus = contractDetails[0].ContractStatus;

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

        this.load_companies(); //Companies list
        this.load_contractType(); //Companies list

        require('./ContractForm');

        const docForSignature = this.renderRequestDetails(parseInt(updateRequestID));

        $('#minimizeButton').on('click', function() {

            const navPlaceholderID = document.getElementById('nav-placeholder');
            const middlePanelID = document.getElementById('middle-panel');
            const minimizeButtonID = document.getElementById('minimizeButton') as HTMLElement;
        
            if (navPlaceholderID && middlePanelID) {
                if (navPlaceholderID.offsetWidth === 0) {
                navPlaceholderID.style.width = '13%';
                middlePanelID.style.width = '60%';
                middlePanelID.style.marginLeft = '13%';
                minimizeButtonID.style.left = '13%';
                } else {
                navPlaceholderID.style.width = '0';
                middlePanelID.style.width = '73%';
                middlePanelID.style.marginLeft = '0%'
                minimizeButtonID.style.left = '0%';
                }
            }
        });

        document.getElementById("contract_details_form").addEventListener("submit", async (event) => {
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

                (document.getElementById('saveToList') as HTMLButtonElement).disabled = true;

                var eleSigned = "";
                var relatedPartyTransaction = "";
                var compliance = "";
                var ELAS = "";

                const termOfContractRB = document.querySelector('input[name="termOfContractRBG"]:checked') as any;  //Req
                const eleSignedRB = document.querySelector('input[name="eleSignedRBG"]:checked') as any;
                const exclusivityRB = document.querySelector('input[name="exclusivityRBG"]:checked') as any; //Req
                const confidentialityRB = document.querySelector('input[name="confidentialityRBG"]:checked') as any; //Req
                const relatedPartyTransactionRB = document.querySelector('input[name="relatedPartyTransactionRBG"]:checked') as any;
                const complianceRB = document.querySelector('input[name="complianceRBG"]:checked') as any;
                const originalCopyFiledRB = document.querySelector('input[name="originalCopyFiledRBG"]:checked') as any; //Req
                const ELAS_RB = document.querySelector('input[name="ELAS_RBG"]:checked') as any;

                if (eleSignedRB) {
                    eleSigned = eleSignedRB.value;
                }
                if (relatedPartyTransactionRB) {
                    relatedPartyTransaction = relatedPartyTransactionRB.value;
                }
                if (complianceRB) {
                    compliance = complianceRB.value;
                }
                if (ELAS_RB) {
                    ELAS = ELAS_RB.value;
                }

                // Retrieve values from the duration fields
                const durationInitialValue = $("#initial_duration_value").val();
                const durationInitialType = $("#initial_duration_type").val();
                var initialDuration = '';
                if(durationInitialValue || durationInitialType){
                    initialDuration = `${durationInitialValue};${durationInitialType}`;
                }

                const durationRenewedValue = $("#renewed_duration_value").val();
                const durationRenewedType = $("#renewed_duration_type").val();
                var renewedDuration = '';
                if(durationRenewedValue || durationRenewedType){
                    renewedDuration = `${durationRenewedValue};${durationRenewedType}`;
                }

                const noticePeriodTermination_value = $("#noticePeriodTermination_value").val();
                const noticePeriodTermination_type = $("#noticePeriodTermination_type").val();
                const noticePeriodTermination = `${noticePeriodTermination_value};${noticePeriodTermination_type}`;

                const NoticePeriodForExtension_value = $("#noticePeriodExtensionRenewal_value").val();
                const NoticePeriodForExtension_type = $("#noticePeriodExtensionRenewal_type").val();
                var NoticePeriodForExtension = '';
                if(NoticePeriodForExtension_value || NoticePeriodForExtension_type){
                    NoticePeriodForExtension = `${NoticePeriodForExtension_value};${NoticePeriodForExtension_type}`;
                }

                const contractValue_Amount = $("#contract_value").val();
                const contractValue_Currency = $("#contract_currency").val();
                const contractValue = `${contractValue_Amount};${contractValue_Currency}`;

                const currentUser = await sp.web.currentUser();
                const currentUserName = currentUser.Title;

                const currentDateUK = this.getCurrentDateUKFormat();

                //Form data
                var data = {
                    NameOfContract: $("#contract_name").val(),
                    InternalReferenceNumber: $("#internal_ref_num").val(),
                    Status: $("#contractStatus").val(),
                    Party_ENL_Rogers_Name: $("#partyENL_Rogers_name").val(),
                    Party_ENL_Rogers_Person: $("#partyENL_Rogers_person").val(),
                    Party2_Name: $("#party2_name").val(),
                    Party2_Person: $("#party2_person").val(),
                    Party3_Name: $("#party3_name").val(),
                    Party3_Person: $("#party3_person").val(),
                    Party4_Name: $("#party4_name").val(),
                    Party4_Person: $("#party4_person").val(),
                    Party5_Name: $("#party5_name").val(),
                    Party5_Person: $("#party5_person").val(),
                    TypeOfContract: $("#contractType").val(),
                    ValueOfContract: contractValue,
                    ApprovedBy: $("#approvedBy").val(),

                    DateOfSignature:$("#signatureDate").val(),
                    DateOfEffect:$("#effectDate").val(),
                    DateOfExpiry:$("#expiryDate").val(),

                    TermOfContract: termOfContractRB.value,
                    Duration_Initial: initialDuration,
                    Duration_Renewed: renewedDuration,
                    NoticePeriodForTermination: noticePeriodTermination,
                    NoticePeriodForExtension: NoticePeriodForExtension,
                    RenewalTerms: $("#renewalTerms").val(),
                    SalientTerms: $("#salientTerms").val(),
                    Jurisdiction: $("#jurisdiction").val(),
                    DisputeResolution: $("#disputeResolution").val(),
                    AddendaName: $("#addendaName").val(),
                    AddendaDate:$("#addendaDate").val(),
                    ElectronicallySigned: eleSigned,
                    Exclusivity: exclusivityRB.value,
                    Confidentiality: confidentialityRB.value,
                    RelatedPartyTransaction: relatedPartyTransaction,
                    Compliance: compliance,
                    OriginalCopyFiled: originalCopyFiledRB.value,
                    ExternalLegalAdvice: ELAS,
                    ComplianceDetails: $("#ComplianceDetails").val(),
                    ExternalLegalAdvicePerson: $("#ELAS_Name").val(),
                    Breach: $("#breachDetails").val(),
                    Litigation: $("#litigationDetails").val(),
                    LastUpdatedOn: currentDateUK,
                    LastUpdatedBy: currentUserName
                };

                console.log(data);

                try {
                    // Get the list item where Request_ID equals updateRequestID
                    const items = await sp.web.lists.getByTitle("Contract_Details").items.filter(`Request_ID eq ${updateRequestID}`).get();
                    console.log(items);
            
                    if (items.length > 0) {
                        const itemId = items[0].Id; // Get the ID of the item to update
                        console.log(itemId);
            
                        // Update the item with the new data
                        await sp.web.lists.getByTitle("Contract_Details").items.getById(itemId).update(data);
                        console.log(`Item with Request_ID ${updateRequestID} updated successfully.`);
                        alert("Contract has been updated successfully.");
                    } else {
                        console.log(`No item found with Request_ID ${updateRequestID}.`);
                    }
                    location.reload();
                    
                } 
                catch (error) {
                    console.error('Error updating item:', error);
                }

                (document.getElementById('saveToList') as HTMLButtonElement).disabled = false;
            }
        });

        $('#sendForSignature').on('click', async () => {
            const fileDetails = await this.getFileDetailsByFilter('Contracts', updateRequestID);
        
            if (!fileDetails) {
                alert('No file found for the specified request ID.');
                return;
            }
            
            const { fileUrl, fileName } = fileDetails;
            
            try {
                const response = await fetch(fileUrl);
                const blob = await response.blob();
            
                const formData = new FormData();
                formData.append("File", blob, fileName);
            
                // Upload the file to Adobe Sign via the local proxy
                const uploadResponse = await fetch('http://localhost:3000/api/proxy/adobeSign', {
                    method: 'POST',
                    body: formData
                });
            
                const result = await uploadResponse.json();
                
                if (result && result.agreementViewList && result.agreementViewList.length > 0) {
                    const url = result.agreementViewList[0].url;
                    if (url) {
                        window.open(url, '_blank');
                    } else {
                        console.error('No URL found in the response');
                    }
                } else {
                    console.error('Invalid response structure:', result);
                }
            } catch (error) {
                console.error('Error uploading file to Adobe Sign:', error);
            }
        });
        
        
        // $('#sendForSignature').on('click', async () => {
        //     // Check if a file is selected by the user
        //     const fileInput = document.getElementById('fileUpload') as HTMLInputElement;
        //     const file = fileInput.files[0]; // Get the first file selected
        
        //     if (!file) {
        //         alert('Please select a file to upload.');
        //         return;
        //     }
        
        //     try {
        //         // Create FormData object
        //         const formData = new FormData();
        //         formData.append("File", file, file.name); // Ensure the field name is correct
        
        //         // Upload the file to Adobe Sign via the local proxy
        //         const uploadResponse = await fetch('http://localhost:3000/api/proxy/adobeSign', {
        //             method: 'POST',
        //             body: formData // Send the FormData object directly
        //             // Do not set headers here; let the browser handle it
        //         });
        
        //         // Handle the response
        //         const result = await uploadResponse.json();
        //         console.log(result);
        //     } catch (error) {
        //         console.error('Error uploading file to Adobe Sign:', error);
        //     }
        // });
        
    }
    
    private async renderRequestDetails(id: any) {

        try {
            // Retrieve the item from the SharePoint list where req_ID matches the provided ID
            const items = await sp.web.lists.getByTitle("Contract_Details").items.filter(`Request_ID eq ${id}`).get();
    
            if (items.length > 0) {
                const item = items[0];
                console.log('Retrieved', item);
    
                $("#contract_name").val(item.NameOfContract);
                $("#internal_ref_num").val(item.InternalReferenceNumber);
                $("#contractStatus").val(item.Status);
                $("#partyENL_Rogers_name").val(item.Party_ENL_Rogers_Name);
                $("#partyENL_Rogers_person").val(item.Party_ENL_Rogers_Person);
                $("#party2_name").val(item.Party2_Name);
                $("#party2_person").val(item.Party2_Person);
                $("#party3_name").val(item.Party3_Name);
                $("#party3_person").val(item.Party3_Person);
                $("#party4_name").val(item.Party4_Name);
                $("#party4_person").val(item.Party4_Person);
                $("#party5_name").val(item.Party5_Name);
                $("#party5_person").val(item.Party5_Person);
                $("#contractType").val(item.TypeOfContract);

                const contractValue = item.ValueOfContract;
                if (contractValue) {
                    const [contract_value, contract_currency] = contractValue.split(';');
                    $("#contract_value").val(contract_value);
                    $("#contract_currency").val(contract_currency);
                }

                
                $("#approvedBy").val(item.ApprovedBy);
                $("#signatureDate").val(item.DateOfSignature);
                $("#effectDate").val(item.DateOfEffect);
                $("#expiryDate").val(item.DateOfExpiry);

                if (item.TermOfContract === "Indefinite") {
                    $('input[name="termOfContractRBG"][value="Indefinite"]').prop('checked', true);
                } else if (item.TermOfContract === "Fixed Term") {
                    $('input[name="termOfContractRBG"][value="Fixed Term"]').prop('checked', true);
                }

                const durationInitial = item.Duration_Initial;
                if (durationInitial) {
                    const [durationInitialValue, durationInitialType] = durationInitial.split(';');
                    $("#initial_duration_value").val(durationInitialValue);
                    $("#initial_duration_type").val(durationInitialType);
                }

                const durationRenewed = item.Duration_Renewed;
                if (durationRenewed) {
                    const [durationRenewedValue, durationRenewedType] = durationRenewed.split(';');
                    $("#renewed_duration_value").val(durationRenewedValue);
                    $("#renewed_duration_type").val(durationRenewedType);
                }

                const noticePeriodForTermination = item.NoticePeriodForTermination;
                if (noticePeriodForTermination) {
                    const [noticePeriodForTerminationValue, noticePeriodForTerminationType] = noticePeriodForTermination.split(';');
                    $("#noticePeriodTermination_value").val(noticePeriodForTerminationValue);
                    $("#noticePeriodTermination_type").val(noticePeriodForTerminationType);
                }

                const NoticePeriodForExtension = item.NoticePeriodForExtension;
                if (NoticePeriodForExtension) {
                    const [NoticePeriodForExtensionValue, NoticePeriodForExtensionType] = NoticePeriodForExtension.split(';');
                    $("#noticePeriodExtensionRenewal_value").val(NoticePeriodForExtensionValue);
                    $("#noticePeriodExtensionRenewal_type").val(NoticePeriodForExtensionType);
                }

                $("#renewalTerms").val(item.RenewalTerms);
                $("#salientTerms").val(item.SalientTerms);
                $("#jurisdiction").val(item.Jurisdiction);
                $("#disputeResolution").val(item.DisputeResolution);
                $("#addendaName").val(item.AddendaName);
                $("#addendaDate").val(item.AddendaDate);

                if (item.ElectronicallySigned === "Yes") {
                    $('input[name="eleSignedRBG"][value="Yes"]').prop('checked', true);
                } if (item.ElectronicallySigned === "No") {
                    $('input[name="eleSignedRBG"][value="No"]').prop('checked', true);
                }
                if (item.Exclusivity === "Yes") {
                    $('input[name="exclusivityRBG"][value="Yes"]').prop('checked', true);
                } if (item.Exclusivity === "No") {
                    $('input[name="exclusivityRBG"][value="No"]').prop('checked', true);
                }
                if (item.Confidentiality === "Yes") {
                    $('input[name="confidentialityRBG"][value="Yes"]').prop('checked', true);
                } if (item.Confidentiality === "No") {
                    $('input[name="confidentialityRBG"][value="No"]').prop('checked', true);
                }
                if (item.RelatedPartyTransaction === "Yes") {
                    $('input[name="relatedPartyTransactionRBG"][value="Yes"]').prop('checked', true);
                } if (item.RelatedPartyTransaction === "No") {
                    $('input[name="relatedPartyTransactionRBG"][value="No"]').prop('checked', true);
                }
                if (item.Compliance === "Yes") {
                    $('input[name="complianceRBG"][value="Yes"]').prop('checked', true);
                } if (item.Compliance === "No") {
                    $('input[name="complianceRBG"][value="No"]').prop('checked', true);
                }
                if (item.OriginalCopyFiled === "Yes") {
                    $('input[name="originalCopyFiledRBG"][value="Yes"]').prop('checked', true);
                } if (item.OriginalCopyFiled === "No") {
                    $('input[name="originalCopyFiledRBG"][value="No"]').prop('checked', true);
                }
                if (item.ExternalLegalAdvice === "Yes") {
                    $('input[name="ELAS_RBG"][value="Yes"]').prop('checked', true);
                } if (item.ExternalLegalAdvice === "No") {
                    $('input[name="ELAS_RBG"][value="No"]').prop('checked', true);
                }

                $("#ComplianceDetails").val(item.ComplianceDetails);
                $("#ELAS_Name").val(item.ExternalLegalAdvicePerson);
                $("#breachDetails").val(item.Breach);
                $("#litigationDetails").val(item.Litigation);
                $("#lastUpdatedOn").val(item.LastUpdatedOn);
                $("#lastUpdatedBy").val(item.LastUpdatedBy);
    
                console.log(`Form populated with data from item with req_ID ${id}.`);
            } else {
                console.log(`No item found with req_ID ${id}.`);
            }
        } catch (error) {
            console.error('Error retrieving item:', error);
        }

        $("#section_review_contract").css("display", "block");

        var fileUrl = '';

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

            fileUrl = `${baseUrl}/${fileDetails.fileUrl}`;

          } else {
            console.log("Item not found.");
          }
        })
        .catch((error) => {
          console.log(error);
        });

        return fileUrl;

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

    private getCurrentDateUKFormat() {
        const date = new Date();
        const day = ('0' + date.getDate()).slice(-2);
        const month = ('0' + (date.getMonth() + 1)).slice(-2);
        const year = date.getFullYear();
        return `${day}/${month}/${year}`;
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

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }
}
