// import * as $ from 'jquery';

// import { Navigation } from 'spfx-navigation';

// //import * as sharepointConfig from '../../common/sharepoint-config.json';

// export class sideMenuUtils {
//   public buildSideMenu(absoluteUrl: string) {
//     var navbar = `<nav id="sidebar" >
//     <div class="custom-menu">
//       <button type="button" id="sidebarCollapse" class="btn btn-primary">
//         <i class="fa fa-bars"></i>
//         <span class="sr-only">Toggle Menu</span>
//       </button>
//     </div>
    
//   <img id="imgLogo" src="${absoluteUrl}/Site Assets/enl_logo_blue.png" alternate="ENL-logo" style="
//     vertical-align: middle;
//     margin-left: 15%;
//     margin-top: -10%;
//     margin-bottom: 30%;">

//    <!-- <span id="imgLogo" style="
//     vertical-align: middle;
//     margin-left: 26%;
//     margin-top: 20%;
//     margin-bottom: 30%;
//     font-size: xx-large;
//     color:blue;">DEMO</span> -->


//     <ul class="list-unstyled components mb-5">
//     <li>
//       <a id="contractMgthome"><span class="fas fa-columns mr-3"></span>Home</a>
//     </li>

//     <li>

//       <a href="#contractMgtmenu" data-toggle="collapse" aria-expanded="false" class="dropdown-toggle collapsed"><span class="fas fa-file-contract mr-3"></span>Contract Management</a>
//       <ul class="collapse list-unstyled" id="contractMgtmenu">
//         <li>
//           <a id="newContract"><span class="fas fa-tasks mr-3"></span>Request a Contract</a>
//         </li>
//         <li>
//           <a id="draftContact"><span class="fas fa-tasks mr-3"></span>Dashboard</a>
//         </li>

//         <!--      <li>
//           <a id="inProcessContact"><span class="fas fa-tasks mr-3"></span>Amended Contracts</a>
//         </li>
//         <li>
//           <a id="approvedContact"><span class="fas fa-tasks mr-3"></span>Approved by Procurement</a>
//         </li>
//         <li>
//         <a id="approvedContactLegal"><span class="fas fa-tasks mr-3"></span>Approved by Legal</a>
//       </li>
//         <li>
//           <a id="expiringContact"><span class="fas fa-tasks mr-3"></span>Expiring Contracts</a>
//         </li>    -->     
//       </ul>

//     </li>

//     <li>
//         <a href="#adminManagment" data-toggle="collapse" aria-expanded="false" class="dropdown-toggle collapsed"><span class="fas fa-file-contract mr-3"></span>Admin Management</a>
//         <ul class="collapse list-unstyled" id="adminManagment">
//           <li>
//             <a id="auditTrail"><span class="fas fa-tasks mr-3"></span>Add Type of Contract</a>
//           </li>
//           <li>
//           <a id="add_department"><span class="fas fa-tasks mr-3"></span>Add Company</a>
//         </li>
//         <li>
//         <a id="add_supplier"><span class="fas fa-tasks mr-3"></span>Add Services</a>
//       </li>

//       <!--    
//       <li>
//       <a id="add_service_details"><span class="fas fa-tasks mr-3"></span>Add Service Details</a>
//      </li>
   
       
//         </ul>
//       </li>

//     <li>
//     <a id="existing_contract"><span class="fas fa-columns mr-3"></span>Existing Contracts</a>
//   </li> 
//   <li>
//   <a id="signed_contract"><span class="fas fa-columns mr-3"></span>Signed Contracts</a>
// </li> -->

//   </ul>
//   </nav>`;
//     $("#nav-placeholder").html(navbar);
//   //  this.sideMenuNavigation(absoluteUrl);
//   }

//   // public sideMenuNavigation(absoluteUrl: string) {
//   //   $("#contractMgthome").on("click", () => {
//   //     Navigation.navigate(`${absoluteUrl}/SitePages/${sharepointConfig.Page.HomePage}`, true);
//   //   });

//   //   $("#newContract").on("click", () => {
//   //     // Navigation.navigate(`${absoluteUrl}/SitePages/${sharepointConfig.Page.NewContractPage}`, true);
//   //   });

//   //   $("#auditTrail").on("click", () => {
//   //     Navigation.navigate(`${absoluteUrl}/SitePages/${sharepointConfig.Page.AuditTrail}`, true);
//   //   });

//   //   $("#newContact").on("click", () => {
//   //     Navigation.navigate(`${absoluteUrl}/SitePages/${sharepointConfig.Page.UploadCheckList}`, true);
//   //   });
//   //   $("#draftContact").on("click", () => {
//   //     Navigation.navigate(`${absoluteUrl}/SitePages/${sharepointConfig.Page.ListOfAgreements}?status=Draft&process_level=Business unit`, true);
//   //   });
//   //   $("#inProcessContact").on("click", () => {
//   //     Navigation.navigate(`${absoluteUrl}/SitePages/${sharepointConfig.Page.ListOfAgreements}?status=Amend&process_level=Procurement`, true);
//   //   });
//   //   $("#allContact").on("click", () => {
//   //     Navigation.navigate(`${absoluteUrl}/SitePages/${sharepointConfig.Page.ListOfAgreements}`, true);
//   //   });
//   //   $("#expiringContact").on("click", () => {
//   //     // Navigation.navigate(`${absoluteUrl}/SitePages/${sharepointConfig.Page.ExpiringContracts}`, true);
//   //   });

//   //   $("#approvedContact").on("click", () => {
//   //     Navigation.navigate(`${absoluteUrl}/SitePages/${sharepointConfig.Page.ListOfAgreements}?status=Approved&process_level=Procurement`, true);
//   //   });

//   //   $("#approvedContactLegal").on("click", () => {
//   //     Navigation.navigate(`${absoluteUrl}/SitePages/${sharepointConfig.Page.ListOfAgreements}?status=Approved&process_level=Legal`, true);
//   //   });

//   //   $("#existing_contract").on("click", () => {
//   //     Navigation.navigate(`${absoluteUrl}/SitePages/Contracts.aspx?type=Existing`, true);
//   //   });

//   //   $("#signed_contract").on("click", () => {
//   //     Navigation.navigate(`${absoluteUrl}/SitePages/Contracts.aspx?type=Signed`, true);
//   //   });

//   //   //add department/service_providers/supplier
//   //   $("#add_department").on("click", () => {
//   //     Navigation.navigate(`${absoluteUrl}/SitePages/AddListItems.aspx?list=Department`, true);
//   //   });

//   //   $("#add_supplier").on("click", () => {
//   //     Navigation.navigate(`${absoluteUrl}/SitePages/AddListItems.aspx?list=Supplier`, true);
//   //   });

//   //   $("#add_service_details").on("click", () => {
//   //     Navigation.navigate(`${absoluteUrl}/SitePages/AddListItems.aspx?list=Service_details`, true);
//   //   });
    

//   // }
// }

import * as $ from 'jquery';
require('../scss/style.scss');
require('../css/style.css');
require('../css/common.css');

import { Navigation } from 'spfx-navigation';

// import * as sharepointConfig from '../../common/sharepoint-config.json';

export class sideMenuUtils {
  
  public buildSideMenu(absoluteUrl: string) {
    var navbar = `
    <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.11.0/umd/popper.min.js" integrity="sha384-b/U6ypiBEHpOf/4+1nzFpr53nxSS+GLCkfwBdFNTxtclqqenISfwAzpKaMNFNmj4" crossorigin="anonymous"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-beta/js/bootstrap.min.js" integrity="sha384-h0AbiXch4ZDo7tp9hKZ4TsHbi047NrKGLO3SEJAg45jXxnGIfYzk4Si90RDIqNm1" crossorigin="anonymous"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css" integrity="sha512-DTOQO9RWCH3ppGqcWaEA1BIZOC6xxalwEsw9c2QQeAIftl+Vegovlnee1c9QX4TctnWMn13TZye+giMm8e2LwA==" crossorigin="anonymous" referrerpolicy="no-referrer" />
    
    <nav id="sidebar">
      <img id="imgLogo" src="https://frcidevtest.sharepoint.com/sites/ContractMgt/Site%20Assets/enl_logo_blue.png" alternate="ENL-logo" style="
      vertical-align: middle;
      margin-left: 25%;
      width: 50%;
      margin-bottom: 15%;">
      <ul class="list-unstyled components mb-5">
      <li>
        <a id="contractMgthome"><span class="fas fa-columns mr-3"></span>Home</a>
      </li>

      <li>
        <a href="#contractMgtmenu" data-toggle="collapse" aria-expanded="false" class="dropdown-toggle collapsed"><span class="fas fa-file-contract mr-3"></span>Contract Management</a>
        <ul class="collapse list-unstyled" id="contractMgtmenu">
          <li>
            <a id="despatcherDashboard"><span class="fas fa-tasks mr-3"></span>Despatcher Dashboard</a>
          </li>
          <li>
            <a id="ownerDashboard"><span class="fas fa-tasks mr-3"></span>Owner Dashboard</a>
          </li>
          <li>
            <a id="requestorDashboard"><span class="fas fa-tasks mr-3"></span>Requestor Dashboard</a>
          </li>     
        </ul>
      </li>

      <li>
        <a href="#adminManagment" data-toggle="collapse" aria-expanded="false" class="dropdown-toggle collapsed"><span class="fas fa-file-contract mr-3"></span>Admin Management</a>
        <ul class="collapse list-unstyled" id="adminManagment">
          <li>
            <a id="auditTrailDashboard"><span class="fas fa-tasks mr-3"></span>Audit Trail</a>
          </li>
          <li>
            <a id="addCompany"><span class="fas fa-tasks mr-3"></span>Add Company</a>
          </li>
        </ul>
      </li>
    </ul>
  </nav>`;

    $("#nav-placeholder").html(navbar);
    this.sideMenuNavigation(absoluteUrl);

    // $(document).ready(function(){
    //   $('#sidebarCollapse').on('click', function(){
    //     $('#sidebar').toggleClass('#sidebar.active');
    //   });
    // });
    
  }

  public sideMenuNavigation(absoluteUrl: string) {
    $("#contractMgthome").on("click", () => {
      Navigation.navigate(`${absoluteUrl}/SitePages/Home(1).aspx`, true);
    });

    $("#despatcherDashboard").on("click", () => {
      Navigation.navigate(`${absoluteUrl}/SitePages/Despatcher-Dashboard.aspx`, true);
    });

    $("#ownerDashboard").on("click", () => {
      Navigation.navigate(`${absoluteUrl}/SitePages/Owner-Dashboard.aspx`, true);
    });

    $("#requestorDashboard").on("click", () => {
      Navigation.navigate(`${absoluteUrl}/SitePages/Requestor-Dashboard.aspx`, true);
    });

    $("#auditTrailDashboard").on("click", () => {
      Navigation.navigate(`${absoluteUrl}/SitePages/Audit-Trail-Dashboard.aspx`, true);
    });

    // $("#addCompany").on("click", () => {
    //   Navigation.navigate(`${absoluteUrl}/SitePages/Requestor-Form.aspx`, true);
    // });

  }
}