import * as $ from 'jquery';

export class textUtils {
    public removeHtmlTags(strHtml: string): string {
        var regex = /(<([^>]+)>)/ig
            , body = strHtml
            , result = body.replace(regex, "").replace("&nbsp;", "");
        return result;
    }
}