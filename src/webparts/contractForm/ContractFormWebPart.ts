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


export interface IContractFormWebPartProps {
    description: string;
}

export default class ContractFormWebPart extends BaseClientSideWebPart<IContractFormWebPartProps> {

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

    public render(): void {
        this.domElement.innerHTML = `
        <div class="wrapper d-flex align-items-stretch">
            <div id="nav-placeholder"></div>
        
            <div class="p-4 p-md-5 pt-3" id="content">
                <div class="${styles.contractForm}" id="form_checklist">
        
                    <form id="requestor_form" style="    border: 1px solid #ccc;
                padding: 20px;
                box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%);
                margin: 2em;
                padding: 2em;">
        
                        <div class="${styles['form-group']}">
        
                            <h2>Contract Form</h2>
                            <h5 class="${styles.heading}">Contract Details</h5>
        
                            <div class="${styles.grid}">
        
                                <div class="${styles['col-1-2']}">
                                    <div class="${styles.controls}">
                                        <input type="text" id="contract_name" class="floatLabel">
                                        <label for="contract_name">Name of Contract</label>
                                    </div>
                                </div>
        
                                <div class="${styles['col-1-2']}">
                                    <div class="${styles.controls}">
        
                                        <input type="text" class="floatLabel" id="internal_ref_num">
                                        <label for="internal_ref_num">Internal Reference Number</label>
                                    </div>
                                </div>
                                <!-- 
                                    <div class="${styles['col-1-3']}">
                                        <div class="${styles.controls}">
            
                                            <input type="text" class="floatLabel" id="email">
                                            <label for="email">Email*</label>
                                        </div>
            
                                    </div> -->
        
                                </br>
        
                                <div class="${styles.grid}">
                                    <div class="${styles['col-1-2']}">
                                        <div class="${styles.controls}">
                                            <i class="fa fa-sort"></i>
                                            <input type="text" class="floatLabel" id="obligationParty1" />
        
                                            <label for="obligationParty1">Obligation Owner for Party 1</label>
        
                                        </div>
                                    </div>
        
        
                                    <div class="${styles['col-1-2']}">
                                        <div class="${styles.controls}">
                                            <input type="text" class="floatLabel" id="obligationParty2">
                                            <label for="obligationParty2">Obligation owner for Party 2(only where party 2 is an
                                                ENL Co) Same for Parties 3 to...</label>
                                        </div>
                                    </div>
        
        
                                </div>
        
                                </br>
        
                                <div class="${styles.grid}">
                                    <div id="contentTable">
                                        <div class="w3-container" id="table">
                                            <div id="content3">
                                                <div id="tblOtherParties" class="table-responsive-xl">
                                                    <div class="form-row">
                                                        <div class="col-xl-12">
                                                            <div id="otherParties">
        
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
                                            <input type="text" class="floatLabel" id="contractType">
                                            <label for="contractType">Type of Contract</label>
        
                                        </div>
                                    </div>
        
                                    <div class="${styles['col-1-3']}">
                                        <div class="${styles.controls}">
                                            <input type="date" class="floatLabel" id="party2">
                                            <label for="party2">Date of Signature</label>
        
                                        </div>
                                    </div>
        
                                    <div class="${styles['col-1-3']}">
                                        <div class="${styles.controls}">
                                            <input type="date" class="floatLabel" id="effectDate">
                                            <label for="effectDate">Date of Effect</label>
        
                                        </div>
                                    </div>
                                </div>
        
                                </br>
        
                                <div class="${styles.grid}">
                                    <div class="${styles['col-1-3']}">
                                        <div class="${styles.controls}">
                                            <input type="textarea" class="floatLabel" id="termOfContract">
                                            <label for="termOfContract">Term of Contract</label>
        
                                        </div>
                                    </div>
        
                                    <div class="${styles['col-1-3']}">
                                        <div class="${styles.controls}">
                                            <input type="text" class="floatLabel" id="initial_duration">
                                            <label for="initial_duration">Duration(Initial)</label>
        
                                        </div>
                                    </div>
        
                                    <div class="${styles['col-1-3']}">
                                        <div class="${styles.controls}">
                                            <input type="text" class="floatLabel" id="renewed_duration">
                                            <label for="renewed_duration">Duration(Renewed)</label>
        
                                        </div>
                                    </div>
                                </div>
        
                                </br>
        
                                <div class="${styles.grid}">
        
        
                                    <div class="${styles['col-1-3']}">
                                        <div class="${styles.controls}">
                                            <input type="date" class="floatLabel" id="expiry_date">
                                            <label for="expiry_date">Date of Expiry</label>
        
                                        </div>
                                    </div>
        
                                    <div class="${styles['col-1-3']}">
                                        <div class="${styles.controls}">
                                            <input type="date" class="floatLabel" id="renewal_terms">
                                            <label for="renewal_terms">Renewal Terms</label>
        
                                        </div>
                                    </div>
        
                                    <div class="${styles['col-1-3']}">
                                        <div class="${styles.controls}">
                                            <input type="text" class="floatLabel" id="noticePeriod">
                                            <label for="noticePeriod">Notice Period for renewal/extension</label>
        
                                        </div>
                                    </div>
                                </div>
        
                                </br>
        
                                <div class="${styles.grid}">
        
        
                                    <div class="${styles['col-1-3']}">
                                        <div class="${styles.controls}">
                                            <input type="text" class="floatLabel" id="noticePeriodTermination">
                                            <label for="noticePeriodTermination">Notice period for termination</label>
        
                                        </div>
                                    </div>
        
                                    <div class="${styles['col-1-3']}">
                                        <div class="${styles.controls}">
                                            <input type="text" class="floatLabel" id="contractValue">
                                            <label for="contractValue">Value of Contract</label>
        
                                        </div>
                                    </div>
        
                                    <div class="${styles['col-1-3']}">
                                        <div class="${styles.controls}">
                                            <input type="text" class="floatLabel" id="approvedBy">
                                            <label for="approvedBy">Approved by*</label>
        
                                        </div>
                                    </div>
                                </div>
        
                                </br>
        
                                <div class="${styles.grid}">
        
        
                                    <div class="${styles['col-1-3']}">
                                        <div class="${styles.controls}">
                                            <input type="text" class="floatLabel" id="contractStatus">
                                            <label for="contractStatus">Status</label>
        
                                        </div>
                                    </div>
        
                                    <div class="${styles['col-1-3']}">
                                        <div class="${styles.controls}">
                                            <input type="text" class="floatLabel" id="contractValue">
                                            <label for="salientTerms">Salient Terms</label>
        
                                        </div>
                                    </div>
        
                                    <div class="${styles['col-1-3']}">
                                        <div class="${styles.controls}">
                                            <input type="text" class="floatLabel" id="noticePeriod">
                                            <label for="addenda_amendment">For addenda/amendment ony:name,date and reference of
                                                initial agreement.</label>
        
                                        </div>
                                    </div>
                                </div>
        
                                </br>
        
                                <div class="${styles.grid}">
        
        
                                    <div class="${styles['col-1-3']}">
                                        <div class="${styles.controls}">
                                            <input type="text" class="floatLabel" id="electroniacallySigned">
                                            <label for="electroniacallySigned">Electronically signed</label>
        
                                        </div>
                                    </div>
        
                                    <div class="${styles['col-1-3']}">
                                        <div class="${styles.controls}">
                                            <input type="date" class="floatLabel" id="contractValue">
                                            <label for="exclusivity">Excluisivity</label>
        
                                        </div>
                                    </div>
        
                                    <div class="${styles['col-1-3']}">
                                        <div class="${styles.controls}">
                                            <input type="date" class="floatLabel" id="noticePeriod">
                                            <label for="noticePeriod">Condidentiality</label>
        
                                        </div>
                                    </div>
                                </div>
        
                                </br>
        
                                <div class="${styles.grid}">
        
        
                                    <div class="${styles['col-1-4']}">
                                        <div class="${styles.controls}">
                                            <input type="text" class="floatLabel" id="disputeResolution">
                                            <label for="disputeResolution">Dispute Resolution</label>
        
                                        </div>
                                    </div>
        
                                    <div class="${styles['col-1-4']}">
                                        <div class="${styles.controls}">
                                            <input type="text" class="floatLabel" id="jurisdiction">
                                            <label for="jurisdiction">Jurisdiction</label>
        
                                        </div>
                                    </div>
        
                                    <div class="${styles['col-1-4']}">
                                        <div class="${styles.controls}">
                                            <input type="date" class="floatLabel" id="relatedParty">
                                            <label for="relatedParty">Related Party Transaction</label>
        
                                        </div>
                                    </div>
        
                                    <div class="${styles['col-1-4']}">
                                        <div class="${styles.controls}">
                                            <input type="date" class="floatLabel" id="noticePeriod">
                                            <label for="complianceSEMDEM">Compliance SEM/DEM</label>
        
                                        </div>
                                    </div>
                                </div>
        
                                </br>
        
                                <div class="${styles.grid}">
        
        
                                    <div class="${styles['col-1-3']}">
                                        <div class="${styles.controls}">sp_comments_list_SectD
                                            <input type="date" class="floatLabel" id="originalCopyFiled">
                                            <label for="originalCopyFiled">Original Copy Filed</label>
        
                                        </div>
                                    </div>
        
                                    <div class="${styles['col-1-3']}">
                                        <div class="${styles.controls}">
                                            <input type="date" class="floatLabel" id="externalLegalAdvice">
                                            <label for="externalLegalAdvice">External Legal Advice</label>
        
                                        </div>
                                    </div>
        
                                    <div class="${styles['col-1-3']}">
                                        <div class="${styles.controls}">
                                            <input type="text" class="floatLabel" id="breach">
                                            <label for="breach">Breach</label>
        
                                        </div>
                                    </div>
                                </div>
        
                                </br>
        
                                <div class="${styles.grid}">
        
                                    <div class="${styles.controls}">
                                        <input type="text" class="floatLabel" id="ligitigation">
                                        <label for="ligitigation">Litigation</label>
        
                                    </div>
        
        
                                </div>
        
                                </br>
        
        
        
                                <div class="form-row">
        
                                    <button type="button" class="buttoncss" id="saveToList"><i class="fa fa-refresh icon"></i>
                                        Save</button>
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
        
                        
                            </div>
                    </form>
                </div>
            </div>
        </div>`;

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

        require('./ContractForm');

        var requestID = this.getItemId();
        console.log('Here :', requestID);

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

        this.renderComments(requestID);

    }

    private getItemId() {
        var queryParms = new URLSearchParams(document.location.search.substring(1));
        var myParm = queryParms.get("requestid");
        if (myParm) {
            return myParm.trim();
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

    private async renderComments(id: any) {
        try {
            const listContainer3: Element = this.domElement.querySelector('#sp_comments_list_SectD');

            let currentWebUrl = this.context.pageContext.web.absoluteUrl;
            // let requestUrl = currentWebUrl.concat(`/_api/web/Lists/GetByTitle('Comments')/items?$select=Comment, CommentBy, CommentDate &$filter=(ContractID eq '${id}') `);
            var doc = null;
            var date = null;
            let html: string = "";

            const commentItemsList = await sp.web.lists.getByTitle("Comments").items.filter(`RequestID eq '${id}'`).select("Comment", "CommentBy", "CommentDate").getAll();
            // var commentItemsList: any = sp.web.lists.getByTitle("Comments").items.filter(`RequestID eq '${id}'`).getAll();
            console.log(commentItemsList);

            html = '';

            console.log("Items", doc);
            html += `
                            <table id='tbl_comment_attach_mgt_SectD' class='table table-striped'>
                            <thead>
                                <tr>
                                <th class="text-left">Comment</th>              
                                <th class="text-left">Name of Person</th>
                                <th class="text-left">Date/Time</th>
                                </tr>
                            </thead>
                            <tbody id="tb_contract_mgt_SectD">
                        `;

                console.log('Test 1');
                commentItemsList.forEach((result: any) => {
                    console.log('Test 2');
                    const item = {
                        Comment: result.Comment,
                        CommentBy: result.CommentBy,
                        CommentDate: result.CommentDate,
                    };

                let date = '';
                if (!Date.parse(item.CommentDate)) {
                    date = item.CommentDate;
                } else {
                    date = moment(new Date(item.CommentDate)).format("DD/MM/YYYY HH:mm");
                }

                html += `
                                <tr>
                                    <td class="text-left">${item.Comment}</td>
                                    <td class="text-left">${item.CommentBy}</td>
                                    <td class="text-left">${date}</td>
                                </tr>
                                `;
            });

            html += `</tbody></table>`;
            listContainer3.innerHTML += html;

            if (doc != null) {
                $('#tbl_comment_attach_mgt_SectD').DataTable({
                    info: false,
                    responsive: true,
                    pageLength: 5,
                    order: [[0, 'desc']],
                });
            }
        }
        catch (err) {
            console.log(err.message);
        }
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
