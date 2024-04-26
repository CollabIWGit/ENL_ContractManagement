import * as $ from 'jquery';
import styles from './RequestFormWebPart.module.scss';

$(function () {


    function floatLabel() {
        $('.floatLabel').each(function() {
          var $this = $(this);
          
          // Check url parameter
          var requestId = getParameterByName('requestid');
          if(requestId){
            $this.next().addClass(`${styles.active}`);
          }
          
          // on focus add class
          $this.focus(function() {
            $this.next().addClass(`${styles.active}`); 
          });
          
          // on blur remove class if empty
          $this.blur(function() {
            if ($this.val() === '' || $this.val() === 'blank') {
              $this.next().removeClass();
            }
          });
        });
      }

    floatLabel();


    function getParameterByName(name, url) {
        if (!url) url = window.location.href;
        name = name.replace(/[\[\]]/g, '\\$&');
        var regex = new RegExp('[?&]' + name + '(=([^&#]*)|&|#|$)');
        var results = regex.exec(url);
        if (!results) return null;
        if (!results[2]) return '';
        return decodeURIComponent(results[2].replace(/\+/g, ' '));
      }

    function floatLabel2() {
        $('.floatLabel2').each(function () {
            var $this = $(this);

            var requestId = getParameterByName('requestid');
            if(requestId){
              $this.next().next().addClass(`${styles.active}`);
            }
            // on focus add cladd active to label
            $this.focus(function () {
                $this.next().next().addClass(`${styles.active}`);
            });
            //on blur check field and remove class if needed
            $this.blur(function () {
                if ($this.val() === '' || $this.val() === 'blank') {
                    $this.next().next().removeClass();
                }
            });
        });
    }
    // just add a class of "floatLabel2 to the input field!"
    floatLabel2();
});