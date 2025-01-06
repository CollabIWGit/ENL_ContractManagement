import * as $ from 'jquery';
require('../scss/style.scss');
require('../css/style.css');
require('../css/common.css');
 
import { Navigation } from 'spfx-navigation';
import * as sharepointConfig from '../../../src/common/sharepoint-config.json';
  
// import * as sharepointConfig from '../../common/sharepoint-config.json';
 
export class sideMenuUtils {
 
  public buildSideMenu(absoluteUrl: string, departments) {
    var navbar = `
    <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.11.0/umd/popper.min.js" integrity="sha384-b/U6ypiBEHpOf/4+1nzFpr53nxSS+GLCkfwBdFNTxtclqqenISfwAzpKaMNFNmj4" crossorigin="anonymous"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-beta/js/bootstrap.min.js" integrity="sha384-h0AbiXch4ZDo7tp9hKZ4TsHbi047NrKGLO3SEJAg45jXxnGIfYzk4Si90RDIqNm1" crossorigin="anonymous"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css" integrity="sha512-DTOQO9RWCH3ppGqcWaEA1BIZOC6xxalwEsw9c2QQeAIftl+Vegovlnee1c9QX4TctnWMn13TZye+giMm8e2LwA==" crossorigin="anonymous" referrerpolicy="no-referrer" />
   
    <nav id="sidebar">
      <div style="font-size: 2.5rem; text-align: center; width: 100%; padding: 2rem 0rem; font-weight: 600;">
        <span style="color: white;">Legal</span><span style="color: #ff2738;">Link</span>
      </div>
 
      <ul class="list-unstyled components mb-5">
        <li>
          <a id="contractMgthome"><span class="fas fa-columns mr-2"></span>Dashboard</a>
        </li>
 
        <li id="adminSect">
        <a href="#adminManagment"><span class="fas fa-file-contract mr-2"></span>Admin Management</a>
        <ul id="adminManagment">
        <li>
          <a id="add_contracts"><span class="fas fa-tasks mr-2"></span>Onboard contracts</a>
        </li>
          <li>
          <a id="add_company"><span class="fas fa-tasks mr-2"></span>Add Company</a>
        </li>
        <li>
        <a id="add_service"><span class="fas fa-tasks mr-2"></span>Add Services</a>
      </li>
 
      <li>
      <a id="add_typeOfContract"><span class="fas fa-tasks mr-2"></span>Add Type of Contract</a>
     </li>
 
      <li>
      <a id="manageUsers"><span class="fas fa-tasks mr-2"></span>Manage Users</a>
     </li>

         <li>
          <a id="manageDirectors"><span class="fas fa-tasks mr-2"></span>Manage Directors</a>
        </li>
       
        </ul>
      </li>  
          <li id="siteContentsNB">
            <ul>
              <li>
                <a id="site_contents"><span class="fas fa-tasks mr-2"></span>Contract Libraries</a>
              </li>
            </ul>
          </li> 
      </ul>
      <div>
        <img id="imgLogo" src="${absoluteUrl}/SiteAssets/Images/ENlnRogersLogo.png" alternate="ENL-logo" style="bottom: 20px; width: 100%; position: absolute;">
      </div>
    </nav>`;
 
    $("#nav-placeholder").html(navbar);
    this.sideMenuNavigation(absoluteUrl);

    $('#adminSect').hide();
    $('#siteContentsNB').hide();

    if (departments.includes('Despatcher')){
      $('#adminSect').show();
    }

    if (departments.includes('InternalOwner')){
      $('#siteContentsNB').show();
    }
   
  }
 
  public sideMenuNavigation(absoluteUrl: string) {
    $("#contractMgthome").on("click", () => {
      Navigation.navigate(`${absoluteUrl}/SitePages/Dashboard.aspx`, true);
    });
 
    $("#auditTrail").on("click", () => {
   //   Navigation.navigate(`${absoluteUrl}/SitePages/${page}/?`, true);
    });
 
    $("#add_company").on("click", () => {
      Navigation.navigate(`${absoluteUrl}/SitePages/AddToList.aspx/?list=Company`, true);
    });
 
    $("#add_service").on("click", () => {
      Navigation.navigate(`${absoluteUrl}/SitePages/AddToList.aspx/?list=ENR_Services`, true);
    });
 
    $("#add_typeOfContract").on("click", () => {
      Navigation.navigate(`${absoluteUrl}/SitePages/AddToList.aspx/?list=Type of contracts`, true);
    });
 
    $("#manageUsers").on("click", () => {
      Navigation.navigate(`${absoluteUrl}/SitePages/AddUser.aspx`, true);
    });

    $("#manageDirectors").on("click", () => {
      Navigation.navigate(`${absoluteUrl}/SitePages/ManageDirectors.aspx`, true);
    });
 
    $("#add_contracts").on("click", () => {
      Navigation.navigate(`${absoluteUrl}/SitePages/OnboardActiveContracts.aspx`, true);
    });

    $("#site_contents").on("click", () => {
      Navigation.navigate(`${absoluteUrl}/SitePages/Links-Dashboard.aspx`, true);
    });
 
 
 
 
 
 
 
    // $("#despatcherDashboard").on("click", () => {
    //   Navigation.navigate(`${absoluteUrl}/SitePages/Despatcher-Dashboard.aspx`, true);
    // });
 
    // $("#ownerDashboard").on("click", () => {
    //   Navigation.navigate(`${absoluteUrl}/SitePages/Owner-Dashboard.aspx`, true);
    // });
 
    // $("#requestorDashboard").on("click", () => {
    //   Navigation.navigate(`${absoluteUrl}/SitePages/Requestor-Dashboard.aspx`, true);
    // });
 
    // $("#auditTrailDashboard").on("click", () => {
    //   Navigation.navigate(`${absoluteUrl}/SitePages/Audit-Trail-Dashboard.aspx`, true);
    // });
 
    // $("#addCompany").on("click", () => {
    //   Navigation.navigate(`${absoluteUrl}/SitePages/Requestor-Form.aspx`, true);
    // });
 
  }

}