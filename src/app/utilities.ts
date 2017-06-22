import uiResource from "./resources";
import * as $ from 'jquery';

const ConditionalFormatting = GC.Spread.Sheets.ConditionalFormatting;
const ComparisonOperators = ConditionalFormatting.ComparisonOperators;
const conditionalFormatTexts:any = uiResource.conditionalFormat.texts;

export const resourceMap:any = {};

export const setRadioButtonActive = function(name:any, index:any) {
	let $items = $("div.insp-radio-button-group[data-name='" + name + "'] div>span");

	$items.removeClass("active");
	$($items[index]).addClass("active");
}

export const setFontStyleButtonActive = function(name:any, active:any) {
    var $target = $("div.group-container>span[data-name='" + name + "']");

    if (active) {
        $target.addClass("active");
    } else {
        $target.removeClass("active");
    }
}

export const setColorValue = function(name:any, value:any) {
    $("div.insp-color-picker[data-name='" + name + "'] div.color-view").css("background-color", value || "");
}

export const px2pt = function(pxValue:any) {
    var tempSpan = $("<span></span>");
    tempSpan.css({
        "font-size": "96pt",
        "display": "none"
    });
    tempSpan.appendTo($(document.body));
    var tempPx = tempSpan.css("font-size");
    if (tempPx.indexOf("px") !== -1) {
        var tempPxValue = parseFloat(tempPx);
        return Math.round(pxValue * 96 / tempPxValue);
    }
    else {  // when browser have not convert pt to px, use 96 DPI.
        return Math.round(pxValue * 72 / 96);
    }
}

export const setTextDisabled = function(name:any, isDisabled:any) {
    var $item = $("div.insp-text[data-name='" + name + "']");
    var $input = $item.find("input");
    if (isDisabled) {
        $item.addClass("disabled");
        $input.attr("disabled", "true");
    }
    else {
        $item.removeClass("disabled");
        $input.attr("disabled", "false");
    }
}

export const setRadioItemChecked = function(groupName:any, itemName:any) {
    var $radioGroup = $("div.insp-checkbox[data-name='" + groupName + "']");
    var $radioItems = $("div.radiobutton[data-name='" + itemName + "']");

    $radioGroup.find(".radiobutton").removeClass("checked");
    $radioItems.addClass("checked");
}

export const selectedCurrentSlicerStyle = function(slicer:any) {
    var slicerStyle = slicer.style(),
        styleName = slicerStyle && slicerStyle.name();
    $("#slicerStyles .slicer-format-item").removeClass("slicer-format-item-selected");
    styleName = styleName.split("SlicerStyle")[1];
    if (styleName) {
        $("#slicerStyles .slicer-format-item div[data-name='" + styleName.toLowerCase() + "']").parent().addClass("slicer-format-item-selected");
    }
}

export const getNumberValue = function(name:string):number {
    return +$("div[data-name='" + name + "'] input.editor").val();
}

export const getTextValue = function(name:string):string {
    return <string>$("div.insp-text[data-name='" + name + "'] input.editor").val();
}

export const getBackgroundColor = function(name:string):string {
    return $("div.insp-color-picker[data-name='" + name + "'] div.color-view").css("background-color");
}

export const getCheckValue = function(name:string):boolean {
    let $target = $("div.insp-checkbox[data-name='" + name + "'] div.button");

    return $target.hasClass("checked");
}

export const processBorderLineSetting = function(name:string) {
    let $borderLineType = $('#border-line-type');
    $borderLineType.text("");
    $borderLineType.removeClass();

    switch (name) {
        case "none":
            $('#border-line-type').text(getResource("cellTab.border.noBorder"));
            $('#border-line-type').addClass("no-border");
            return;

        case "hair":
            $('#border-line-type').addClass("line-style-hair");
            break;

        case "dotted":
            $('#border-line-type').addClass("line-style-dotted");
            break;

        case "dash-dot-dot":
            $('#border-line-type').addClass("line-style-dash-dot-dot");
            break;

        case "dash-dot":
            $('#border-line-type').addClass("line-style-dash-dot");
            break;

        case "dashed":
            $('#border-line-type').addClass("line-style-dashed");
            break;

        case "thin":
            $('#border-line-type').addClass("line-style-thin");
            break;

        case "medium-dash-dot-dot":
            $('#border-line-type').addClass("line-style-medium-dash-dot-dot");
            break;

        case "slanted-dash-dot":
            $('#border-line-type').addClass("line-style-slanted-dash-dot");
            break;

        case "medium-dash-dot":
            $('#border-line-type').addClass("line-style-medium-dash-dot");
            break;

        case "medium-dashed":
            $('#border-line-type').addClass("line-style-medium-dashed");
            break;

        case "medium":
            $('#border-line-type').addClass("line-style-medium");
            break;

        case "thick":
            $('#border-line-type').addClass("line-style-thick");
            break;

        case "double":
            $('#border-line-type').addClass("line-style-double");
            break;

        default:
            console.log("processBorderLineSetting not add for ", name);
            break;
    }
}

export const getResource = function(key:string) {
    key = key.replace(/\./g, "_");

    return resourceMap[key];
}

export const setCheckValue = function (name:string, value:boolean, options?:any) {
    let $target = $("div.insp-checkbox[data-name='" + name + "'] div.button");
    if (value) {
        $target.addClass("checked");
    } else {
        $target.removeClass("checked");
    }
    if (options) {
        $target.data(options);
    }
}

export const setNumberValue = function(name:string, value:any) {
    $("div.insp-number[data-name='" + name + "'] input.editor").val(value);
}

export const getCurrentTime = function():string {
    let date = new Date();
    let year = date.getFullYear();
    let month = date.getMonth() + 1;
    let day = date.getDate();

    let strDate = year + "-";
    if (month < 10)
        strDate += "0";
    strDate += month + "-";
    if (day < 10)
        strDate += "0";
    strDate += day;

    return strDate;
}

export const getDropDownValue = function(name:string, host?:HTMLElement|Document):any {
    host = host || document;

    let container:any = $(host).find("div.insp-dropdown-list[data-name='" + name + "']");
    let	refList:string = "#" + $(container).data("list-ref");
    let text:string = $("span.display", container).text();

    let value:any = $("div.text", $(refList)).filter(function () {
        return $(this).text() === text;
    }).data("value");

    return value;
}

export const setTextValue = function(name:string, value:any) {
    $("div.insp-text[data-name='" + name + "'] input.editor").val(value);
}

export const processNumberValidatorComparisonOperatorSetting = function(value:GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators) {
    if (value === ComparisonOperators.between || value === ComparisonOperators.notBetween) {
        $("#numberValue").hide();
        $("#numberBetweenOperator").show();
    }
    else {
        $("#numberBetweenOperator").hide();
        $("#numberValue").show();
    }
}

export const processTextLengthValidatorComparisonOperatorSetting = function(value:GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators) {
    if (value === ComparisonOperators.between || value === ComparisonOperators.notBetween) {
        $("#textLengthValue").hide();
        $("#textLengthBetweenOperator").show();
    }
    else {
        $("#textLengthBetweenOperator").hide();
        $("#textLengthValue").show();
    }
}

export const processDateValidatorComparisonOperatorSetting = function(value:GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators) {
    if (value === ComparisonOperators.between || value === ComparisonOperators.notBetween) {
        $("#dateValue").hide();
        $("#dateBetweenOperator").show();
    }
    else {
        $("#dateBetweenOperator").hide();
        $("#dateValue").show();
    }
}

export const updateIconCriteriaItems = function(iconStyleType:number) {
    let IconSetType = ConditionalFormatting.IconSetType,
        items = $("#iconCriteriaSetting .settinggroup"),
        values:number[];

    if (iconStyleType <= IconSetType.threeSymbolsUncircled) {
        values = [33, 67];
    } else if (iconStyleType <= IconSetType.fourTrafficLights) {
        values = [25, 50, 75];
    } else {
        values = [20, 40, 60, 80];
    }

    items.each(function (index) {
        let value = values[index], $item = $(this), suffix = index + 1;

        if (value) {
            $item.show();
            setDropDownValue("iconSetCriteriaOperator" + suffix, 1, this);
            setDropDownValue("iconSetCriteriaType" + suffix, 4, this);
            $("input.editor", this).val(value);
        } else {
            $item.hide();
        }
    });
}

export const setDropDownValue = function(container:string|any, value:any, host?:HTMLElement|Document) {
    if (typeof container === "string") {
        host = host || document;

        container = $(host).find("div.insp-dropdown-list[data-name='" + container + "']");
    }

    let refList:string = "#" + $(container).data("list-ref");

    $("span.display", container).text($(".menu-item>div.text[data-value='" + value + "']", $(refList)).text());
}

export const processConditionalFormatSetting = function(groupName:string, listRef?:string, rule?:number) {
    $("#conditionalFormatSettingContainer div.details").show();
    setConditionalFormatSettingGroupVisible(groupName);

    var $ruleType = $("#highlightCellsRule"),
        $setButton = $("#setConditionalFormat");
    if (listRef) {
        $ruleType.data("list-ref", listRef);
        $setButton.data("rule-type", rule);
        let item = setDropDownValueByIndex($ruleType, 0);
        updateEnumTypeOfCF(item.value);
    } else {
        $setButton.data("rule-type", groupName);
    }
}

export const setConditionalFormatSettingGroupVisible = function(groupName:string) {
    let $groupItems = $("#conditionalFormatSettingContainer .settingGroup .groupitem");

    $groupItems.hide();
    $groupItems.filter("[data-group='" + groupName + "']").show();
}

export const setDropDownValueByIndex = function(container:any, index:number):{ text:string, value: string } {
    let refList:string = "#" + $(container).data("list-ref");
    let $item = $(".menu-item:eq(" + index + ") div.text", $(refList));

    $("span.display", container).text($item.text());

    return {text: $item.text(), value: $item.data("value")};
}

export const updateEnumTypeOfCF = function(itemType:string|number) {
    let $operator = $("#ComparisonOperator");
    let	$setButton = $("#setConditionalFormat");

    $setButton.data("rule-type", itemType);

    switch ("" + itemType) {
        case "0":
            $("#ruletext").text(conditionalFormatTexts.cells);
            $("#andtext").hide();
            $("#formattext").hide();
            $("#value1").show();
            $("#value1").val("");
            $("#value2").hide();
            $("#colorScale").hide();
            $operator.show();
            $operator.data("list-ref", "cellValueOperatorList");
            setDropDownValueByIndex($operator, 0);
            break;
        case "1":
            $("#ruletext").text(conditionalFormatTexts.cells);
            $("#andtext").hide();
            $("#formattext").hide();
            $("#value1").show();
            $("#value1").val("");
            $("#value2").hide();
            $("#colorScale").hide();
            $operator.show();
            $operator.data("list-ref", "specificTextOperatorList");
            setDropDownValueByIndex($operator, 0);
            break;
        case "2":
            $("#ruletext").text(conditionalFormatTexts.cells);
            $("#andtext").hide();
            $("#formattext").hide();
            $("#value1").hide();
            $("#value2").hide();
            $("#colorScale").hide();
            $operator.show();
            $operator.data("list-ref", "dateOccurringOperatorList");
            setDropDownValueByIndex($operator, 0);
            break;
        case "4":
            $("#ruletext").text(conditionalFormatTexts.rankIn);
            $("#andtext").hide();
            $("#formattext").hide();
            $("#value1").show();
            $("#value1").val("10");
            $("#value2").hide();
            $("#colorScale").hide();
            $operator.show();
            $operator.data("list-ref", "top10OperatorList");
            setDropDownValueByIndex($operator, 0);
            break;
        case "5":
        case "6":
            $("#ruletext").text(conditionalFormatTexts.all);
            $("#andtext").hide();
            $("#formattext").show();
            $("#formattext").text(conditionalFormatTexts.inRange);
            $("#value1").hide();
            $("#value2").hide();
            $("#colorScale").hide();
            $operator.hide();
            break;
        case "7":
            $("#ruletext").text(conditionalFormatTexts.values);
            $("#andtext").hide();
            $("#formattext").show();
            $("#formattext").text(conditionalFormatTexts.average);
            $("#value1").hide();
            $("#value2").hide();
            $("#colorScale").hide();
            $operator.show();
            $operator.data("list-ref", "averageOperatorList");
            setDropDownValueByIndex($operator, 0);
            break;
        case "8":
            $("#ruletext").text(conditionalFormatTexts.allValuesBased);
            $("#andtext").hide();
            $("#formattext").hide();
            $("#value1").hide();
            $("#value2").hide();
            $("#colorScale").show();
            $("#midpoint").hide();
            $("#minType").val("1");
            $("#maxType").val("2");
            $("#minValue").val("");
            $("#maxValue").val("");
            $("#minColor").css("background", "#F8696B");
            $("#maxColor").css("background", "#63BE7B");
            $operator.hide();
            break;
        case "9":
            $("#ruletext").text(conditionalFormatTexts.allValuesBased);
            $("#andtext").hide();
            $("#formattext").hide();
            $("#value1").hide();
            $("#value2").hide();
            $("#colorScale").show();
            $("#midpoint").show();
            $("#minType").val("1");
            $("#midType").val("4");
            $("#maxType").val("2");
            $("#minValue").val("");
            $("#midValue").val("50");
            $("#maxValue").val("");
            $("#minColor").css("background-color", "#F8696B");
            $("#midColor").css("background-color", "#FFEB84");
            $("#maxColor").css("background-color", "#63BE7B");
            $operator.hide();
            break;
        default:
            break;
    }
}

export const getRGBAColor = function(color:any, alpha:any) {
    var result = color,
        prefix = "rgb(";

    // get rgb color use jquery
    if (color.substr(0, 4) !== prefix) {
        var $temp = $("#setfontstyle");
        $temp.css("background-color", color);
        color = $temp.css("background-color");
    }

    // adding alpha to make rgba
    if (color.substr(0, 4) === prefix) {
        var length = color.length;
        result = "rgba(" + color.substring(4, length - 1) + ", " + alpha + ")";
    }

    return result;
}