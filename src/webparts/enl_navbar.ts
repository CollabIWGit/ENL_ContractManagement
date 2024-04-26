import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';

require('../Assets/scripts/styles/navstyles.css');

SPComponentLoader.loadScript('https://ajax.googleapis.com/ajax/libs/jquery/3.6.3/jquery.min.js');
//SPComponentLoader.loadScript('https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-YvpcrYf0tY3lHB60NNkmXc5s9fDVZLESaAA55NDzOxhy9GkcIdslK1eN7N6jIeHz" crossorigin="anonymous"');


SPComponentLoader.loadScript('https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js', {
  globalExportsName: 'bootstrap'  
}).then(() => {
  // Bootstrap loaded and available globally as window.bootstrap
});

export default class nav {

    public static navHTML: string =
        `
        
        <nav class="side-nav" style="grid-column: 1;">
          <ul class="nav-menu">
            <span class="logo">
                <svg version="1.1" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 80" xml:space="preserve">
                <path class="logo-text" d="M89.4,29.8l6-28.8H83.4l-8,38.8h29.1l2.1-9.9H89.4z M68.3,1l-4.3,21.2h-0.1L56.6,1H44.7l-3.4,16.3
                    c0,2.3-0.2,4.4-0.4,6.5H40l-1,4.6h0.4c-0.3,0.8-0.6,1.6-0.9,2.3l-1.9,9.1H48l4.5-21.7h0.1l7.1,21.7h11.7L79.6,1H68.3z M23.5,0
                    C9.3,0,0.2,10.5,0.2,24.3c0,11.5,7.9,17,18.8,17c8.7,0,16-3,19.6-10.7l0.5-2.3H27c-1,2.6-3.1,4.3-7.5,4.3c-4.2,0-6.5-2.2-6.9-6.2
                    c0-1.3,0.1-1.9,0.1-2.6h28.2c0.3-2.1,0.4-4.2,0.4-6.5v-0.2C41.3,5.8,33.6,0,23.5,0z M29,16.6H13.6c1.4-4.9,4.1-8,9.4-8
                    C27.3,8.7,29.4,12.3,29,16.6z M39.5,28.4c-0.3,0.8-0.6,1.6-0.9,2.3l0.5-2.3H39.5z"></path>
                    <path class="logo-icon" d="M110.7,39.7c0.2,0,0.4,0,0.6,0h27.4c4,0,7.9-3.2,8.7-7.1L154.2,0C151.3,10.5,141.6,29.4,110.7,39.7" style="fill: orange;"></path>                
                </svg>			
            </span>
            <li class="nav-item"><a href="#"><i class="fas fa-tachometer-alt"></i><span class="menu-text">Dashboard</span></a></li>
            <li class="nav-item"><a href="#"><i class="fas fa-user"></i><span class="menu-text">Users</span></a></li>
            <li class="nav-item active"><a href="#"><i class="fas fa-file-alt"></i><span class="menu-text">Posts</span></a></li>
            <li class="nav-item"><a href="#"><i class="fas fa-play "></i><span class="menu-text">Media</span></a></li>
            <li class="nav-item"><a href="#"><i class="fas fa-sign-out-alt"></i><span class="menu-text">exit</span></a></li>
          </ul>
        </nav>
     `;
}
