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
import { MSGraphClient } from '@microsoft/sp-http';
import * as moment from 'moment';
import * as $ from 'jquery';
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

let currentUser: string;
export interface IContractFormWebPartProps {
    description: string;
}

export default class ContractFormWebPart extends BaseClientSideWebPart<IContractFormWebPartProps> {

    private _isDarkTheme: boolean = false;
    private _environmentMessage: string = '';
    private graphClient: MSGraphClient;


    protected onInit(): Promise<void> {
        currentUser = this.context.pageContext.user.displayName;
        return new Promise<void>(async (resolve: () => void, reject: (error: any) => void): Promise<void> => {
          sp.setup({
            spfxContext: this.context as any
          });
     
          this.context.msGraphClientFactory
          .getClient()
          .then((client: MSGraphClient): void => {
            this.graphClient = client;
            resolve();
          }, err => reject(err));
        });
      }

    public render(): void {
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
            </style>

            <div class="main-container" id="content">

                <div id="nav-placeholder" class="left-panel"></div>

                <div id="middle-panel" class="middle-panel">
                    <button id="minimizeButton"></button>
                    <div class="${styles.contractForm}" id="form_checklist">
                        <form id="requestor_form" style=" position: relative; border: 1px solid #ccc; padding: 20px; box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%); padding: 2em; border-radius: 1rem; width: 100%;">
                            
                            <p id="contractStatusUp" style="color: green; position: absolute; top: 0; right: 0; margin: 1%;">Status: In progress</p>
                            
                            <div class="${styles.grid}">

                                <div class="${styles['form-group']}">
                                    <h2>Contract Form</h2>
                                    <h5 class="${styles.heading}">Contract Details</h5>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" id="contract_name" class="floatLabel">
                                        <label for="contract_name">Name of Contract*</label>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="internal_ref_num">
                                        <label for="internal_ref_num">Internal Reference Number*</label>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="contractStatus">
                                        <label for="contractStatus">Status*</label>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <p> </p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <p>Name of Party*</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <p>Name of person responsible</p>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <p>Party 1 (ENL-Rogers side)*</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="partyENL_Rogers_name">
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="partyENL_Rogers_person">
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <p>Party 2*</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="party2_name">
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="party2_person">
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <p>Party 3</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="party3_name">
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="party3_person">
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <p>Party 4</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="party4_name">
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="party4_person">
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <p>Party 5</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="party5_name">
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="party5_person">
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="contractType">
                                        <label for="contractType">Type of Contract*</label>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="contractValue">
                                        <label for="contractValue">Value of Contract*</label>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="approvedBy">
                                        <label for="approvedBy">Approved by</label>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="date" class="floatLabel" id="signatureDate">
                                        <label for="signatureDate">Date of Signature*</label>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="date" class="floatLabel" id="effectDate">
                                        <label for="effectDate">Date of Effect*</label>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="date" class="floatLabel" id="expiryDate">
                                        <label for="expiryDate">Date of Expiry</label>
                                    </div>
                                </div>
                                
                                <div class="${styles['col-1-3']}">
                                    <div style="display: flex; flex-direction: column;">
                                        <p style="margin-bottom: 0px;">Term of Contract*</p>
                                        <div style="display: flex;">
                                            <input type="radio" name="termOfContractRBG" value="Indefinite">
                                            <label style="margin-bottom: 0px; margin-left: 5px;" for="termOfContractIndefinite">Indefinite</label>
                                        </div>
                                        <div style="display: flex;">
                                            <input type="radio" name="termOfContractRBG" value="Fixed Term">
                                            <label style="margin-bottom: 0px; margin-left: 5px;" for="termOfContractFixedTerm">Fixed Term</label>
                                        </div>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}" style="margin-bottom: 5px;">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="initial_duration">
                                        <label for="initial_duration">Duration (Initial)</label>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}" style="margin-bottom: 5px;">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="renewed_duration">
                                        <label for="renewed_duration">Duration (Renewed)</label>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="noticePeriodTermination">
                                        <label for="noticePeriodTermination">Notice period for termination*</label>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="noticePeriodExtensionRenewal">
                                        <label for="noticePeriodExtensionRenewal">Notice Period for renewal/extension</label>
                                    </div>
                                </div>

                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="date" class="floatLabel" id="renewalTerms">
                                        <label for="renewalTerms">Renewal Terms</label>
                                    </div>
                                </div>

                                <div class="${styles['col-2-3']}">
                                    <div class="${styles.controls}">
                                        <textarea type="text" class="floatLabel" id="salientTerms" style="height: 8rem; margin-bottom: 10px;"></textarea>
                                        <label for="salientTerms">Salient Terms</label>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="jurisdiction">
                                        <label for="jurisdiction">Jurisdiction*</label>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="disputeResolution">
                                        <label for="disputeResolution">Dispute Resolution</label>
                                    </div>
                                </div>

                                <div class="${styles['col-1']}">
                                    <div class="${styles.controls}">
                                        <p>For addenda/amendment only: details of initial contract</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="addendaName">
                                        <label for="addendaName">Name</label>
                                    </div>
                                </div>
                                <div class="${styles['col-1']}">
                                    <div class="${styles.controls}">
                                        <input type="date" class="floatLabel" id="addendaDate">
                                        <label for="addendaDate">Date</label>
                                    </div>
                                </div>


                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <br><br>
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <p style="text-align: center">Yes</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <p style="text-align: center">No</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <br><br>
                                    </div>
                                </div>


                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <p>Electronically Signed</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <input type="radio" name="eleSignedRBG" value="Yes">
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <input type="radio" name="eleSignedRBG" value="No">
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <br><br>    
                                    </div>
                                </div>


                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <p>Exclusivity*</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <input type="radio" name="exclusivityRBG" value="Yes">
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <input type="radio" name="exclusivityRBG" value="No">
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <br><br>
                                    </div>
                                </div>


                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <p>Confidentiality*</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <input type="radio" name="confidentialityRBG" value="Yes">
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <input type="radio" name="confidentialityRBG" value="No">
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <br><br>
                                    </div>
                                </div>


                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <p>Related Party Transaction</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <input type="radio" name="relatedPartyTransactionRBG" value="Yes">
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <input type="radio" name="relatedPartyTransactionRBG" value="No">
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <br><br>
                                    </div>
                                </div>


                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <p>Compliance SEM/DEM/FSC</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <input type="radio" name="complianceRBG" value="Yes">
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <input type="radio" name="complianceRBG" value="No">
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="ComplianceDetails">
                                        <label for="ComplianceDetails">Details</label>
                                    </div>
                                </div>


                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <p>Original Copy Filed*</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <input type="radio" name="originalCopyFiledRBG" value="Yes">
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <input type="radio" name="originalCopyFiledRBG" value="No">
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <br><br>
                                    </div>
                                </div>


                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <p>External Legal advice sought</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <input type="radio" name="ELAS_RBG" value="Yes">
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <input type="radio" name="ELAS_RBG" value="No">
                                    </div>
                                </div>
                                <div class="${styles['col-1-4']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="ELAS_Name">
                                        <label for="ELAS_Name">Name</label>
                                    </div>
                                </div>

                                <div class="${styles['col-1-2']}">
                                    <div class="${styles.controls}">
                                        <textarea type="text" class="floatLabel" id="breachDetails" style="height: 10rem;"></textarea>
                                        <label for="breachDetails">Details of Breach (if any)</label>
                                    </div>
                                </div>
                                <div class="${styles['col-1-2']}">
                                    <div class="${styles.controls}">
                                        <textarea type="text" class="floatLabel" id="litigationDetails" style="height: 10rem;"></textarea>
                                        <label for="litigationDetails">Details of Litigation (if any)</label>
                                    </div>
                                </div>

                                <div class="${styles['col-1']}">
                                    <div class="${styles.controls}">
                                        <p>Last Updated</p>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="lastUpdatedOn">
                                        <label for="lastUpdatedOn">On</label>
                                    </div>
                                </div>
                                <div class="${styles['col-1-3']}">
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="lastUpdatedBy">
                                        <label for="lastUpdatedBy">By</label>
                                    </div>
                                </div>

                                <div class="${styles['col-1']}" style="margin-bottom: 3rem">
                                    <div class="form-row">
                                        <button type="button" class="buttoncss" id="saveToList"><i class="fa fa-refresh icon" style="display: none"></i>Save</button>
                                        <button type="button" class="buttoncss">Cancel</button>
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

        SideMenuUtils.buildSideMenu(this.context.pageContext.web.absoluteUrl);


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

        //Update Contract
        const urlParams = new URLSearchParams(window.location.search);
        const updateRequestID = urlParams.get('requestid');

        // if (!updateRequestID) {
        // $('#rightPanel').hide();
        // const middlePanelID = document.getElementById('middle-panel');
        // middlePanelID.style.marginRight = '0%';
        // middlePanelID.style.width = '83%';
        // $('#contractStatus').hide();
        // }
        // else {
        // document.getElementById('rightPanel').innerHTML = `
        // <div style="width: 100%; height:100%; background: white; padding-bottom: 30%;">
        //     <div class="timelineHeader">
        //     <p style="margin-bottom: 0px;">Timeline</p>
        //     </div>
        //     <ul id="commentTimeline" class="timeline"></ul>
        //     <div class="comment-box">
        //     <textarea id="comment" class="comment-input" placeholder="Add your comment..."></textarea>
        //     <button id="addComment">Add Comment</button>
        //     </div>
        // </div>
        // `;
        // }

        require('./ContractForm');

        var requestID = this.getItemId();
        console.log('Here :', requestID);

        // $("#addComment").click(async (e) => {
        //     console.log("Test New Comment");
        //     // icon_add_comment.classList.remove('hide');
        //     // icon_add_comment.classList.add('show');
        //     // icon_add_comment.classList.add('spinning');
        
        //     const currentUser = await sp.web.currentUser();
        
        //     const data = {
        
        //         Title: requestID,
        //         RequestID: requestID,
        //         Comment: $("#comment").val(),
        //         CommentBy: currentUser.UserPrincipalName,
        //         CommentDate: moment().format("DD/MM/YYYY HH:mm")
        //     };
        
        //     await this.addComment(data);

        //     this.load_comments(requestID);
        
        //     // icon_add_comment.classList.remove('spinning', 'show');
        //     // icon_add_comment.classList.add('hide');

        //     $("#comment").val("");
        
        //     });

        // this.renderComments(requestID);
        // this.load_comments(requestID);
        this.renderRequestDetails(parseInt(requestID));

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

        $("#saveToList").click(async (e) => {
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

            console.log('test1');

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

                ValueOfContract: $("#contractValue").val(),
                
                ApprovedBy: $("#approvedBy").val(),

                // DateOfSignature: $("#signatureDate").val(),
                // DateOfEffect: $("#effectDate").val(),
                // DateOfExpiry: $("#expiryDate").val(),
                TermOfContract: termOfContractRB.value,

                // NoticePeriodForTermination: $("#noticePeriodTermination").val(),
                // NoticePeriodForExtension: $("#noticePeriodExtensionRenewal").val(),

                // RenewalTerms: $("#renewalTerms").val(),

                SalientTerms: $("#salientTerms").val(),
                Jurisdiction: $("#jurisdiction").val(),
                DisputeResolution: $("#disputeResolution").val(),
                AddendaName: $("#addendaName").val(),

                // AddendaDate: $("#addendaDate").val(),

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
                LastUpdatedOn: $("#lastUpdatedOn").val(),
                LastUpdatedBy: $("#lastUpdatedBy").val()
            };
            //Does not work
            // data["Duration(Initial)"] = $("#initial_duration").val();
            // data["Duration(Renewed)"] = $("#renewed_duration").val();

            console.log(data);

            try {
                // Get the list item where Request_ID equals updateRequestID
                const items = await sp.web.lists.getByTitle("Contract_Details").items.filter(`Request_ID eq ${updateRequestID}`).get();
                console.log(items);
        
                if (items.length > 0) {
                    const itemId = items[0].Id; // Get the ID of the item to update
        
                    // Update the item with the new data
                    await sp.web.lists.getByTitle("Contract_Details").items.getById(itemId).update(data);
                    console.log(`Item with Request_ID ${updateRequestID} updated successfully.`);
                    alert("Contract has been updated successfully.");
                } else {
                    console.log(`No item found with Request_ID ${updateRequestID}.`);
                }
            } 
            catch (error) {
                console.error('Error updating item:', error);
            }

            (document.getElementById('saveToList') as HTMLButtonElement).disabled = false;
                
        });

    }

    private async renderRequestDetails(id: any) {

        // try {
        //     let currentWebUrl = this.context.pageContext.web.absoluteUrl;
        //     let requestUrl = currentWebUrl.concat(`/_api/web/Lists/GetByTitle('Contract_Details')/items?$select=NameOfContract, InternalReferenceNumber, Status, Party_ENL_Rogers_Name, Party_ENL_Rogers_Person, Party2_Name, Party2_Person, Party3_Name, Party3_Person, Party4_Name, Party4_Person, Party5_Name, Party5_Person, TypeOfContract, ValueOfContract, ApprovedBy, DateOfSignature, DateOfEffect, DateOfExpiry, TermOfContract, NoticePeriodForTermination, NoticePeriodForExtension, RenewalTerms, SalientTerms, Jurisdiction, DisputeResolution, AddendaName, AddendaDate, ElectronicallySigned, Exclusivity, Confidentiality, RelatedPartyTransaction, Compliance: compliance, OriginalCopyFiled, ExternalLegalAdvice, ComplianceDetails, ExternalLegalAdvicePerson, Breach,Litigation, LastUpdatedOn, LastUpdatedBy &$filter=(Request_ID eq '${id}') `);
            
            
        //     var retrievedData = null;
        //     const response = this.context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
        //     .then((response: SPHttpClientResponse) => {
        //     if (response.ok) {
        //         response.json()
        //         .then((responseJSON) => {
        //             if (responseJSON != null && responseJSON.value != null) {
        //             retrievedData = responseJSON.value;

        //             console.log("Items", retrievedData);


        //             retrievedData.forEach((result: any) => {
        //                 const item = {

        //                 NameOfRequestor: result.NameOfRequestor,
        //                 StatusTitle: result.Status_Title,
        //                 Email: result.Email,
        //                 Company: result.Company,
        //                 Department: result.Department,
        //                 Party1_agreement: result.Party1_agreement,
        //                 Party2_agreement: result.Party2_agreement,
        //                 Others_parties: result.Others_parties,
        //                 BriefDescriptionTransaction: result.BriefDescriptionTransaction,
        //                 ExpectedCommencementDate: result.ExpectedCommencementDate,
        //                 AuthorityApproveContract: result.AuthorityApproveContract,
        //                 DueDate: result.DueDate,
        //                 TypeOfContract: result.TypeOfContract,
        //                 NameOfAgreement: result.NameOfAgreement,
        //                 RequestFor: result.RequestFor,
        //                 AssignedTo: result.AssignedTo,
        //                 Owner: result.Owner,
        //                 AssigneeComment: result.AssigneeComment,
        //                 Confidential: result.Confidential

        //                 // Date_time: result.DateTime,
        //                 // Attachments: result.AttachmentFiles
        //                 };

        //                 $("#requestor_name").val(item.NameOfRequestor);
        //                 $("#status_title").val(item.StatusTitle);
        //                 // $("#email").val(item.Email);
        //                 // $("#enl_company").val(item.Company);
        //                 // $("#department").val(item.Department);
        //                 // $("#requestFor").val(item.RequestFor);
        //                 // $("#party1").val(item.Party1_agreement);
        //                 // $("#party2").val(item.Party2_agreement);
        //             });
        //         }
        //       });

        //   }
        // });
        // }
        // catch (err) {
        // console.log(err.message);
        // }

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
                $("#contractValue").val(item.ValueOfContract);
                $("#approvedBy").val(item.ApprovedBy);
                $("#signatureDate").val(item.DateOfSignature);
                $("#effectDate").val(item.DateOfEffect);
                $("#expiryDate").val(item.DateOfExpiry);

                if (item.TermOfContract === "Indefinite") {
                    $('input[name="termOfContractRBG"][value="Indefinite"]').prop('checked', true);
                } else if (item.TermOfContract === "Fixed Term") {
                    $('input[name="termOfContractRBG"][value="Fixed Term"]').prop('checked', true);
                }

                $("#NoticePeriodForTermination").val(item.NoticePeriodForTermination);
                $("#NoticePeriodForExtension").val(item.NoticePeriodForExtension);
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
                } if (item.ElectronicallySigned === "No") {
                    $('input[name="exclusivityRBG"][value="No"]').prop('checked', true);
                }
                if (item.Confidentiality === "Yes") {
                    $('input[name="confidentialityRBG"][value="Yes"]').prop('checked', true);
                } if (item.ElectronicallySigned === "No") {
                    $('input[name="confidentialityRBG"][value="No"]').prop('checked', true);
                }
                if (item.RelatedPartyTransaction === "Yes") {
                    $('input[name="relatedPartyTransactionRBG"][value="Yes"]').prop('checked', true);
                } if (item.ElectronicallySigned === "No") {
                    $('input[name="relatedPartyTransactionRBG"][value="No"]').prop('checked', true);
                }
                if (item.Compliance === "Yes") {
                    $('input[name="complianceRBG"][value="Yes"]').prop('checked', true);
                } if (item.ElectronicallySigned === "No") {
                    $('input[name="complianceRBG"][value="No"]').prop('checked', true);
                }
                if (item.OriginalCopyFiled === "Yes") {
                    $('input[name="originalCopyFiledRBG"][value="Yes"]').prop('checked', true);
                } if (item.ElectronicallySigned === "No") {
                    $('input[name="originalCopyFiledRBG"][value="No"]').prop('checked', true);
                }
                if (item.ExternalLegalAdvice === "Yes") {
                    $('input[name="ELAS_RBG"][value="Yes"]').prop('checked', true);
                } if (item.ElectronicallySigned === "No") {
                    $('input[name="ELAS_RBG"][value="No"]').prop('checked', true);
                }

                $("#ComplianceDetails").val(item.ComplianceDetails);
                $("#ELAS_Name").val(item.ExternalLegalAdvicePerson);
                $("#breachDetails").val(item.Breach);
                $("#litigationDetails").val(item.Litigation);
                $("#lastUpdatedOn").val(item.LastUpdatedOn);
                $("#lastUpdatedBy").val(item.LastUpdatedOn);
    
                console.log(`Form populated with data from item with req_ID ${id}.`);
            } else {
                console.log(`No item found with req_ID ${id}.`);
            }
        } catch (error) {
            console.error('Error retrieving item:', error);
        }

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

    // public async load_comments(requestID) {
    //     let userEmail = "";
    //     const timeline = document.getElementById('commentTimeline');
    //     timeline.innerHTML = '';
    //     const CommentList = await sp.web.lists.getByTitle("Comments").items.select("RequestID,Comment,CommentBy,CommentDate").filter(`RequestID eq '${requestID}'`).get();
    //     userEmail = CommentList[0].CommentBy;
    //     const users: any[] = await sp.web.siteUsers();
    //     let userTitle = '';
    //     users.forEach(user => {
    //       if (user.Email === userEmail) {
    //         userTitle = user.Title;
    //         return;
    //       }
    //     });
    //     if (userTitle === '') {
    //       console.log('User with email ' + userEmail + ' not found.');
    //     }
    //     CommentList.forEach(item => {
    //       const comment = item.Comment;
    //       const commentDate = item.CommentDate;
    //       const timelineItem = document.createElement('li');
    //       timelineItem.className = 'timeline-item';
    //       timelineItem.innerHTML = `
    //         <div style="display: flex">
    //           <p style="margin-bottom: 0px">@${userTitle} -&nbsp;</p>
    //           ${commentDate}
    //         </div>
    //         <div>${comment}</div>
    //       `;
    //       timeline.appendChild(timelineItem);
    //     });

    //     timeline.scrollTop = timeline.scrollHeight;
    // }

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

    private getItemId() {
        var queryParms = new URLSearchParams(document.location.search.substring(1));
        var myParm = queryParms.get("requestid");
        if (myParm) {
            return myParm.trim();
        }
    }

    // async addComment(data) {

    //     try {
    //         const iar = await sp.web.lists.getByTitle("Comments").items.add(data);

    //         alert("Comment added succesfully.");

    //     }
    //     catch (e) {

    //         alert("An error occured." + e.message);

    //     }

    // }

    // private async renderComments(id: any) {
    //     try {
    //         const listContainer3: Element = this.domElement.querySelector('#sp_comments_list_SectD');

    //         let currentWebUrl = this.context.pageContext.web.absoluteUrl;
    //         // let requestUrl = currentWebUrl.concat(`/_api/web/Lists/GetByTitle('Comments')/items?$select=Comment, CommentBy, CommentDate &$filter=(ContractID eq '${id}') `);
    //         var doc = null;
    //         var date = null;
    //         let html: string = "";

    //         const commentItemsList = await sp.web.lists.getByTitle("Comments").items.filter(`RequestID eq '${id}'`).select("Comment", "CommentBy", "CommentDate").getAll();
    //         // var commentItemsList: any = sp.web.lists.getByTitle("Comments").items.filter(`RequestID eq '${id}'`).getAll();
    //         console.log(commentItemsList);

    //         html = '';

    //         console.log("Items", doc);
    //         html += `
    //                         <table id='tbl_comment_attach_mgt_SectD' class='table table-striped'>
    //                         <thead>
    //                             <tr>
    //                             <th class="text-left">Comment</th>              
    //                             <th class="text-left">Name of Person</th>
    //                             <th class="text-left">Date/Time</th>
    //                             </tr>
    //                         </thead>
    //                         <tbody id="tb_contract_mgt_SectD">
    //                     `;

    //         console.log('Test 1');
    //         commentItemsList.forEach((result: any) => {
    //             console.log('Test 2');
    //             const item = {
    //                 Comment: result.Comment,
    //                 CommentBy: result.CommentBy,
    //                 CommentDate: result.CommentDate,
    //             };

    //             let date = '';
    //             if (!Date.parse(item.CommentDate)) {
    //                 date = item.CommentDate;
    //             } else {
    //                 date = moment(new Date(item.CommentDate)).format("DD/MM/YYYY HH:mm");
    //             }

    //             html += `
    //                             <tr>
    //                                 <td class="text-left">${item.Comment}</td>
    //                                 <td class="text-left">${item.CommentBy}</td>
    //                                 <td class="text-left">${date}</td>
    //                             </tr>
    //                             `;
    //         });

    //         html += `</tbody></table>`;
    //         listContainer3.innerHTML += html;

    //         if (doc != null) {
    //             $('#tbl_comment_attach_mgt_SectD').DataTable({
    //                 info: false,
    //                 responsive: true,
    //                 pageLength: 5,
    //                 order: [[0, 'desc']],
    //             });
    //         }
    //     }
    //     catch (err) {
    //         console.log(err.message);
    //     }
    // }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

}
