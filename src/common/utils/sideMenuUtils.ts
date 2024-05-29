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
      <div style="font-size: 2.5rem; text-align: center; width: 100%; padding: 4rem 0rem; font-weight: 600;">
        <span style="color: white;">Legal</span><span style="color: #ff2738;">Link</span>
      </div>

      <ul class="list-unstyled components mb-5">
        <li>
          <a id="contractMgthome"><span class="fas fa-columns mr-3"></span>Dashboard</a>
        </li>   
      </ul>
      <div>
        <img id="imgLogo" src="https://frcidevtest.sharepoint.com/sites/ContractMgt/Site%20Assets/ENlnRogersLogo.png" alternate="ENL-logo" style="bottom: 20px; width: 100%; position: absolute;">
      </div>
    </nav>`;

    $("#nav-placeholder").html(navbar);
    this.sideMenuNavigation(absoluteUrl);
    
  }

  public sideMenuNavigation(absoluteUrl: string) {
    $("#contractMgthome").on("click", () => {
      Navigation.navigate(`${absoluteUrl}/SitePages/Dashboard.aspx`, true);
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