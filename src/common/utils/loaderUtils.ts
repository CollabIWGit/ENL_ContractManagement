import * as $ from 'jquery';

export class loaderUtils {
    public toggleLoader(show) {
        if (show) {
            $('.cd-popup3').addClass('is-visible');
        }
        else {
            $('.cd-popup3').removeClass('is-visible');
        }
    }
}