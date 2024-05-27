import { escape } from '@microsoft/sp-lodash-subset';
import styles from './RequestFormWebPart.module.scss';
// require('../../Assets/scripts/styles/custom_styles.css');

export default class sample {

    public static templateHTML: string =
        `<div class="tab-pane" id="section_B">
        
        <fieldset class="fieldset_section_B">
            <div class="${styles['form-group']}">
                <h2 class="${styles.heading}">Section B - Department Details</h2>
    
                <div class="${styles.grid}">
                    <div class="${styles['col-1-2']}">
                        <div class="${styles.controls}">
                            <i class="fa fa-sort"></i>
                            <input id="department" list="dept" class="floatLabel2">
                            <datalist id="dept"><select id="select_dept"></select></datalist>
                            <label for="department">Department</label>
    
                        </div>
                    </div>
                    <div class="${styles['col-1-2']}">
                        <div class="${styles.controls} leftmargin">
                            <input id="agreement_signed_by" list="site_users_name" class="floatLabel2">
                            <datalist id="site_users_name"></datalist>
                            <label for="drp_contract_owner">Agreement signed off by: </label>
                        </div>
                    </div>
                </div>
    
                <div class="${styles.grid}">
                    <div class="${styles['col-1-2']}">
                        <div class="${styles.controls}">
                            <input id="title_position" class="floatLabel2">
                            <label for="drp_start_date_year">Title/Position </label>
                        </div>
                    </div>
                    <div class="${styles['col-1-2']}">
                        <div class="${styles.controls} leftmargin">
                            <!-- <i class="fa fa-sort"></i>
                                    <input id="agreement_name" list="owners" class="floatLabel2">
                                    <datalist id="owners"><select id="select_owners"></select></datalist>
                                    <label for="drp_contract_owner">Signature</label> -->
                            <input id="date_section_B" type="date" class="floatLabel2">
                            <label for="date_section_B">Date</label>
                        </div>
                    </div>
                </div>
    
                <!-- <div class="${styles.grid}">
        
                            <div class="">
                                <div class="${styles.controls}">
                                       <i class="fa fa-sort"></i>
                                    <input id="date_section_B" class="floatLabel2">
                                       <datalist id="owners"><select id="select_owners"></select></datalist>
                                    <label for="nature_service">Date </label>
                                </div>
                            </div>
        
                        </div> -->
    
                <div class="form-row">
                    <div class="col-md-8">
                        <h6></h6>
                    </div>
                    <div class="col-md-4 offset-8">
                        <button id="submit" class="btn btn-secondary submit_section_B" type="button">Submit</button>
                        <button id="cancel" class="btn btn-secondary" type="button">Back</button>
                    </div>
                </div>
            </div>
        </fieldset>
    
    </div>`;
}

