import * as $ from 'jquery';

export class dataListUtils {
    public isValidDatalistValue(idDataList, inputValue): boolean {
        var val = $("#" + inputValue).val();
        var option = $("#" + idDataList).find("option[value='" + val + "']");

        if (option != null && option.length > 0)
            return true;
        else
            return false;
    }
}

