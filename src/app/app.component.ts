import { Component, OnInit } from '@angular/core';
import uiResource from "./resources";
import * as utilities from "./utilities";

import '../assets/css/styles.css';
import '../assets/css/font-awesome/css/font-awesome.min.css'
import '../assets/css/bootstrap.min.css'
import '../assets/css/bootstrap-theme.min.css'
import '../assets/css/gc.spread.sheets.excel2013white.10.1.0.css';
import '../assets/css/inspector.css';
import '../assets/css/insp-table-format.css';
import '../assets/css/sample.css';
import '../assets/css/colorpicker.css';
import '../assets/css/borderpicker.css';
import '../assets/css/insp-spread.css';

const MARGIN_BOTTOM = 4;
const isIE = navigator.userAgent.toLowerCase().indexOf('compatible') < 0 && /(trident)(?:.*? rv ([\w.]+)|)/.exec(navigator.userAgent.toLowerCase()) !== null;
const ConditionalFormatting = GC.Spread.Sheets.ConditionalFormatting;
const ComparisonOperators = ConditionalFormatting.ComparisonOperators;
const Sparklines = GC.Spread.Sheets.Sparklines;

const Calc = GC.Spread.CalcEngine;
const SheetsCalc = GC.Spread.Sheets.CalcEngine;
const ExpressionType = Calc.ExpressionType;

@Component({
	selector: 'my-app',
	templateUrl: './app.component.html',
	styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit {
	spread:GC.Spread.Sheets.Workbook;
	floatInspector:boolean = false;
	tableIndex:number = 1;
	_needShow:boolean = true;
	_colorpicker:any;
	_dropdownitem:any;
	_documentMousedownHandler:any;
	_activeTable:any;
	_activeComment:any;
	isShiftKey:boolean;

	constructor() {
		this._documentMousedownHandler = this.documentMousedownHandler.bind(this);
	}

	ngOnInit():void {
		var z = this;

		$(document).ready(() => {
			z.getResourceMap(uiResource);
			z.localizeUI();

			z.spread = new GC.Spread.Sheets.Workbook($("#ss")[0], {tabStripRatio: 0.88});

			z.getThemeColor();
			z.initSpread();

			//Change default allowCellOverflow the same with Excel.
			z.spread.sheets.forEach(function (sheet) {
				sheet.options.allowCellOverflow = true;
			});

			//window resize adjust
			$(".insp-container").draggable();
			z.checkMediaSize();
			z.screenAdoption();

			let resizeTimeout:number|null = null;
			$(window).bind("resize", () => {
				if (resizeTimeout === null) {
					resizeTimeout = window.setTimeout(function() {
						z.screenAdoption();
						clearTimeout(resizeTimeout);
						resizeTimeout = null;
					}, 100);
				}
			});

			z.doPrepareWork();
			
			let innerSpread = z.spread;

			$("ul.dropdown-menu>li>a").click(function () {
				let value = $(this).text(),
					$divhost = $(this).parents("div.btn-group"),
					groupName = $divhost.data("name"),
					sheet = innerSpread.getActiveSheet();

				$divhost.find("button:first").text(value);

				switch (groupName) {
					case "fontname":
						z.setStyleFont(sheet, "font-family", false, [value], value);
					break;

					case "fontsize":
						z.setStyleFont(sheet, "font-size", false, [value], value);
					break;
				}
			});

			var toolbarHeight = 0,//$("#toolbar").height(),
				formulaboxDefaultHeight = $("#formulabox").outerHeight(true),
				verticalSplitterOriginalTop = formulaboxDefaultHeight - $("#verticalSplitter").height();

			$("#verticalSplitter").draggable({
				axis: "y",              // vertical only
				containment: "#inner-content-container",  // limit in specified range
				scroll: false,          // not allow container scroll
				zIndex: 100,            // set to move on top
				stop: function (event: Event, ui: JQueryUI.DraggableEventUIParams) {
					let element:HTMLElement = this;
					let $this:any = $(element);
					let top = $this.offset().top;
					let	offset = top - toolbarHeight - verticalSplitterOriginalTop;

					// limit min size
					if (offset < 0) {
						offset = 0;
					}
					// adjust size of related items
					$("#formulabox").css({height: formulaboxDefaultHeight + offset});
					var height = $("div.insp-container").height() - $("#formulabox").outerHeight(true);
					$("#controlPanel").height(height);
					$("#ss").height(height);
					z.spread.refresh();
					// reset
					$(element).css({top: 0});
				}
			});

			z.attachEvents();

			$(document).on("contextmenu", function (e) {
				let evt = window.event || e;
				if (!$(evt.target).data('contextmenu')) {
					evt.preventDefault();
					return false;
				}
			});

			$("#download").on("click", function (e) {
				e.preventDefault();
				return false;
			});

			z.spread.focus();

			z.syncSheetPropertyValues();

			z.onCellSelected();

			z.updatePositionBox(z.spread.getActiveSheet());

			//fix bug 220484
			if (isIE) {
				$("#formulabox").css('padding', 0);
			}
		});
	}

	getResourceMap(src:any) {
		function isObject(item:any) {
			return typeof item === "object";
		}

		function addResourceMap(map:any, obj:any, keys:any) {
			if (isObject(obj)) {
				for (var p in obj) {
					var cur = obj[p];

					addResourceMap(map, cur, keys.concat(p));
				}
			} else {
				var key = keys.join("_");
				map[key] = obj;
			}
		}

		addResourceMap(utilities.resourceMap, src, []);
	}

	localizeUI() {
        function getLocalizeString(text:any) {
            var matchs = text.match(/(?:(@[\w\d\.]*@))/g);

            if (matchs) {
                matchs.forEach(function (item:any) {
                    var s = utilities.getResource(item.replace(/[@]/g, ""));
                    text = text.replace(item, s);
                });
            }

            return text;
        }

        $(".localize").each(function () {
            var text = $(this).text();

            $(this).text(getLocalizeString(text));
        });

        $(".localize-tooltip").each(function () {
            var text = $(this).prop("title");

            $(this).prop("title", getLocalizeString(text));
        });

        $(".localize-value").each(function () {
            var text = $(this).attr("value");

            $(this).attr("value", getLocalizeString(text));
        });
    }

	getThemeColor() {
		var sheet = this.spread.getActiveSheet();
		this.setThemeColorToSheet(sheet);                                            // Set current theme color to sheet

		var $colorUl = $("#default-theme-color");
		var $themeColorLi, cellBackColor;
		for (var col = 3; col < 13; col++) {
			var row = 4;
			cellBackColor = sheet.getActualStyle(row, col).backColor;
			$themeColorLi = $("<li class=\"color-cell seed-color-column\"></li>");
			$themeColorLi.css("background-color", cellBackColor).attr("data-name", sheet.getCell(2, col).text()).appendTo($colorUl);
			for (row = 5; row < 10; row++) {
				cellBackColor = sheet.getActualStyle(row, col).backColor;
				$themeColorLi = $("<li class=\"color-cell\"></li>");
				$themeColorLi.css("background-color", cellBackColor).attr("data-name", this.getColorName(sheet, row, col)).appendTo($colorUl);
			}
		}

		sheet.clear(2, 1, 8, 12, GC.Spread.Sheets.SheetArea.viewport, 255);      // Clear sheet theme color
	}

	setThemeColorToSheet(sheet:any) {
		sheet.suspendPaint();

		sheet.getCell(2, 3).text("Background 1").themeFont("Body");
		sheet.getCell(2, 4).text("Text 1").themeFont("Body");
		sheet.getCell(2, 5).text("Background 2").themeFont("Body");
		sheet.getCell(2, 6).text("Text 2").themeFont("Body");
		sheet.getCell(2, 7).text("Accent 1").themeFont("Body");
		sheet.getCell(2, 8).text("Accent 2").themeFont("Body");
		sheet.getCell(2, 9).text("Accent 3").themeFont("Body");
		sheet.getCell(2, 10).text("Accent 4").themeFont("Body");
		sheet.getCell(2, 11).text("Accent 5").themeFont("Body");
		sheet.getCell(2, 12).text("Accent 6").themeFont("Body");

		sheet.getCell(4, 1).value("100").themeFont("Body");

		sheet.getCell(4, 3).backColor("Background 1");
		sheet.getCell(4, 4).backColor("Text 1");
		sheet.getCell(4, 5).backColor("Background 2");
		sheet.getCell(4, 6).backColor("Text 2");
		sheet.getCell(4, 7).backColor("Accent 1");
		sheet.getCell(4, 8).backColor("Accent 2");
		sheet.getCell(4, 9).backColor("Accent 3");
		sheet.getCell(4, 10).backColor("Accent 4");
		sheet.getCell(4, 11).backColor("Accent 5");
		sheet.getCell(4, 12).backColor("Accent 6");

		sheet.getCell(5, 1).value("80").themeFont("Body");

		sheet.getCell(5, 3).backColor("Background 1 80");
		sheet.getCell(5, 4).backColor("Text 1 80");
		sheet.getCell(5, 5).backColor("Background 2 80");
		sheet.getCell(5, 6).backColor("Text 2 80");
		sheet.getCell(5, 7).backColor("Accent 1 80");
		sheet.getCell(5, 8).backColor("Accent 2 80");
		sheet.getCell(5, 9).backColor("Accent 3 80");
		sheet.getCell(5, 10).backColor("Accent 4 80");
		sheet.getCell(5, 11).backColor("Accent 5 80");
		sheet.getCell(5, 12).backColor("Accent 6 80");

		sheet.getCell(6, 1).value("60").themeFont("Body");

		sheet.getCell(6, 3).backColor("Background 1 60");
		sheet.getCell(6, 4).backColor("Text 1 60");
		sheet.getCell(6, 5).backColor("Background 2 60");
		sheet.getCell(6, 6).backColor("Text 2 60");
		sheet.getCell(6, 7).backColor("Accent 1 60");
		sheet.getCell(6, 8).backColor("Accent 2 60");
		sheet.getCell(6, 9).backColor("Accent 3 60");
		sheet.getCell(6, 10).backColor("Accent 4 60");
		sheet.getCell(6, 11).backColor("Accent 5 60");
		sheet.getCell(6, 12).backColor("Accent 6 60");

		sheet.getCell(7, 1).value("40").themeFont("Body");

		sheet.getCell(7, 3).backColor("Background 1 40");
		sheet.getCell(7, 4).backColor("Text 1 40");
		sheet.getCell(7, 5).backColor("Background 2 40");
		sheet.getCell(7, 6).backColor("Text 2 40");
		sheet.getCell(7, 7).backColor("Accent 1 40");
		sheet.getCell(7, 8).backColor("Accent 2 40");
		sheet.getCell(7, 9).backColor("Accent 3 40");
		sheet.getCell(7, 10).backColor("Accent 4 40");
		sheet.getCell(7, 11).backColor("Accent 5 40");
		sheet.getCell(7, 12).backColor("Accent 6 40");

		sheet.getCell(8, 1).value("-25").themeFont("Body");

		sheet.getCell(8, 3).backColor("Background 1 -25");
		sheet.getCell(8, 4).backColor("Text 1 -25");
		sheet.getCell(8, 5).backColor("Background 2 -25");
		sheet.getCell(8, 6).backColor("Text 2 -25");
		sheet.getCell(8, 7).backColor("Accent 1 -25");
		sheet.getCell(8, 8).backColor("Accent 2 -25");
		sheet.getCell(8, 9).backColor("Accent 3 -25");
		sheet.getCell(8, 10).backColor("Accent 4 -25");
		sheet.getCell(8, 11).backColor("Accent 5 -25");
		sheet.getCell(8, 12).backColor("Accent 6 -25");

		sheet.getCell(9, 1).value("-50").themeFont("Body");

		sheet.getCell(9, 3).backColor("Background 1 -50");
		sheet.getCell(9, 4).backColor("Text 1 -50");
		sheet.getCell(9, 5).backColor("Background 2 -50");
		sheet.getCell(9, 6).backColor("Text 2 -50");
		sheet.getCell(9, 7).backColor("Accent 1 -50");
		sheet.getCell(9, 8).backColor("Accent 2 -50");
		sheet.getCell(9, 9).backColor("Accent 3 -50");
		sheet.getCell(9, 10).backColor("Accent 4 -50");
		sheet.getCell(9, 11).backColor("Accent 5 -50");
		sheet.getCell(9, 12).backColor("Accent 6 -50");
		sheet.resumePaint();
	}


	getColorName(sheet:any, row:any, col:any) {
		var colName = sheet.getCell(2, col).text();
		var rowName = sheet.getCell(row, 1).text();
		return colName + " " + rowName;
	}

	syncSheetPropertyValues() {
		let sheet = this.spread.getActiveSheet();
		let	options = sheet.options;

		this.updateCellStyleState(sheet, sheet.getActiveRowIndex(), sheet.getActiveColumnIndex());
	}

	updateCellStyleState(sheet:any, row:any, column:any) {
		var style = sheet.getActualStyle(row, column);

		if (style) {
			var sfont = style.font;

			// Font
			var font
			if (sfont) {
				font = this.parseFont(sfont);

				utilities.setFontStyleButtonActive("bold", ["bold", "bolder", "700", "800", "900"].indexOf(font.fontWeight) !== -1);
				utilities.setFontStyleButtonActive("italic", font.fontStyle !== 'normal');
				this.setDropDownText($("#cellTab div.insp-dropdown-list[data-name='fontFamily']"), font.fontFamily.replace(/'/g, ""));
				this.setDropDownText($("#cellTab div.insp-dropdown-list[data-name='fontSize']"), parseFloat(font.fontSize));
			}

			var underline = GC.Spread.Sheets.TextDecorationType.underline,
				linethrough = GC.Spread.Sheets.TextDecorationType.lineThrough,
				overline =  GC.Spread.Sheets.TextDecorationType.overline,
				textDecoration = style.textDecoration;
			utilities.setFontStyleButtonActive("underline", textDecoration && ((textDecoration & underline) === underline));
			utilities.setFontStyleButtonActive("strikethrough", textDecoration && ((textDecoration & linethrough) === linethrough));
			utilities.setFontStyleButtonActive("overline", textDecoration && ((textDecoration & overline) === overline));

			utilities.setColorValue("foreColor", style.foreColor || "#000");
			utilities.setColorValue("backColor", style.backColor || "#fff");

			// Alignment
			utilities.setRadioButtonActive("hAlign", style.hAlign);   // general (3, auto detect) without setting button just like Excel
			utilities.setRadioButtonActive("vAlign", style.vAlign);
			utilities.setCheckValue("wrapText", style.wordWrap);

			//cell padding
			var cellPadding = style.cellPadding;
			if (cellPadding) {
				utilities.setTextValue("cellPadding", cellPadding);
			} else {
				utilities.setTextValue("cellPadding", "");
			}
			//watermark
			var watermark = style.watermark;
			if (watermark) {
				utilities.setTextValue("watermark", watermark);
			} else {
				utilities.setTextValue("watermark", "");
			}
			//label options
			var labelOptions = style.labelOptions;
			if (labelOptions) {
				var lFont = labelOptions.font;
				if (lFont) {
					font = this.parseFont(lFont);
					utilities.setFontStyleButtonActive("labelBold", ["bold", "bolder", "700", "800", "900"].indexOf(font.fontWeight) !== -1);
					utilities.setFontStyleButtonActive("labelItalic", font.fontStyle !== 'normal');
					this.setDropDownText($("#cellTab div.insp-dropdown-list[data-name='labelFontFamily']"), font.fontFamily.replace(/'/g, ""));
					this.setDropDownText($("#cellTab div.insp-dropdown-list[data-name='labelFontSize']"), parseFloat(font.fontSize));
				} else {
					utilities.setFontStyleButtonActive("labelBold", ["bold", "bolder", "700", "800", "900"].indexOf(font.fontWeight) !== -1);
					utilities.setFontStyleButtonActive("labelItalic", font.fontStyle !== 'normal');
					this.setDropDownText($("#cellTab div.insp-dropdown-list[data-name='labelFontFamily']"), font.fontFamily.replace(/'/g, ""));
					this.setDropDownText($("#cellTab div.insp-dropdown-list[data-name='labelFontSize']"), parseFloat(font.fontSize));
				}
				utilities.setColorValue("labelForeColor", labelOptions.foreColor || "#000");
				utilities.setTextValue("labelMargin", labelOptions.margin || "");
				utilities.setDropDownValueByIndex($("#cellLabelVisibility"), labelOptions.visibility === undefined ? 2 : labelOptions.visibility);
				utilities.setDropDownValueByIndex($("#cellLabelAlignment"), labelOptions.alignment || 0);
			}
		}
	}

	parseFont(font:any) {
		var fontFamily = null,
			fontSize = null,
			fontStyle = "normal",
			fontWeight = "normal",
			fontVariant = "normal",
			lineHeight = "normal";

		var elements = font.split(/\s+/);
		var element;
		while ((element = elements.shift())) {
			switch (element) {
				case "normal":
					break;

				case "italic":
				case "oblique":
					fontStyle = element;
					break;

				case "small-caps":
					fontVariant = element;
					break;

				case "bold":
				case "bolder":
				case "lighter":
				case "100":
				case "200":
				case "300":
				case "400":
				case "500":
				case "600":
				case "700":
				case "800":
				case "900":
					fontWeight = element;
					break;

				default:
					if (!fontSize) {
						var parts = element.split("/");
						fontSize = parts[0];
						if (fontSize.indexOf("px") !== -1) {
							fontSize = utilities.px2pt(parseFloat(fontSize)) + 'pt';
						}
						if (parts.length > 1) {
							lineHeight = parts[1];
							if (lineHeight.indexOf("px") !== -1) {
								lineHeight = utilities.px2pt(parseFloat(lineHeight)) + 'pt';
							}
						}
						break;
					}

					fontFamily = element;
					if (elements.length)
						fontFamily += " " + elements.join(" ");

					return {
						"fontStyle": fontStyle,
						"fontVariant": fontVariant,
						"fontWeight": fontWeight,
						"fontSize": fontSize,
						"lineHeight": lineHeight,
						"fontFamily": fontFamily
					};
			}
		}

		return {
			"fontStyle": fontStyle,
			"fontVariant": fontVariant,
			"fontWeight": fontWeight,
			"fontSize": fontSize,
			"lineHeight": lineHeight,
			"fontFamily": fontFamily
		};
	}

	updatePositionBox(sheet:any) {
		var selection = sheet.getSelections().slice(-1)[0];
		if (selection) {
			var position;
			if (!this.isShiftKey) {
				position = this.getCellPositionString(sheet,
					sheet.getActiveRowIndex() + 1,
					sheet.getActiveColumnIndex() + 1);
			}
			else {
				position = this.getSelectedRangeString(sheet, selection);
			}

			$("#positionbox").val(position);
		}
	}

	getSelectedRangeString(sheet:any, range:any) {
		var selectionInfo = "",
			rowCount = range.rowCount,
			columnCount = range.colCount,
			startRow = range.row + 1,
			startColumn = range.col + 1;

		if (rowCount == 1 && columnCount == 1) {
			selectionInfo = this.getCellPositionString(sheet, startRow, startColumn);
		}
		else {
			if (rowCount < 0 && columnCount > 0) {
				selectionInfo = columnCount + "C";
			}
			else if (columnCount < 0 && rowCount > 0) {
				selectionInfo = rowCount + "R";
			}
			else if (rowCount < 0 && columnCount < 0) {
				selectionInfo = sheet.getRowCount() + "R x " + sheet.getColumnCount() + "C";
			}
			else {
				selectionInfo = rowCount + "R x " + columnCount + "C";
			}
		}
		return selectionInfo;
	}

	getCellPositionString(sheet:any, row:any, column:any) {
		if (row < 1 || column < 1) {
			return null;
		}
		else {
			var letters = "";
			switch (this.spread.options.referenceStyle) {
				case GC.Spread.Sheets.ReferenceStyle.a1: // 0
					while (column > 0) {
						var num = column % 26;
						if (num === 0) {
							letters = "Z" + letters;
							column--;
						}
						else {
							letters = String.fromCharCode('A'.charCodeAt(0) + num - 1) + letters;
						}
						column = parseInt((column / 26).toString());
					}
					letters += row.toString();
					break;
				case GC.Spread.Sheets.ReferenceStyle.r1c1: // 1
					letters = "R" + row.toString() + "C" + column.toString();
					break;
				default:
					break;
			}
			return letters;
		}
	}

	attachEvents():void {
		this.attachToolbarItemEvents();
		this.attachSpreadEvents();
		this.attachConditionalFormatEvents();
		this.attachCellTypeEvents();
		this.attachBorderTypeClickEvents();
	}

	attachConditionalFormatEvents() {
		let component = this;
		$("#setConditionalFormat").click(function () {
			var ruleType = $(this).data("rule-type");

			switch (ruleType) {
				case "databar":
					component.addDataBarRule();
					break;

				case "iconset":
					component.addIconSetRule();
					break;

				default:
					component.addCondionalFormaterRule("" + ruleType);
					break;
			}
		});
	}

	addIconSetRule() {
		var sheet = this.spread.getActiveSheet();
		sheet.suspendPaint();

		var selections = sheet.getSelections();
		if (selections.length > 0) {
			var ranges:any[] = [];
			$.each(selections, function (i, v) {
				ranges.push(new GC.Spread.Sheets.Range(v.row, v.col, v.rowCount, v.colCount));
			});
			var cfs = sheet.conditionalFormats;
			var iconSetRule = new ConditionalFormatting.IconSetRule(
				+utilities.getDropDownValue("iconSetType"),
				ranges);
			var $divs = $("#iconCriteriaSetting .settinggroup:visible");
			var iconCriteria = iconSetRule.iconCriteria();
			$.each($divs, function (i, v) {
				var suffix = i + 1,
					isGreaterThanOrEqualTo = +utilities.getDropDownValue("iconSetCriteriaOperator" + suffix, this) === 1,
					iconValueType = +utilities.getDropDownValue("iconSetCriteriaType" + suffix, this),
					iconValue = $("input.editor", this).val();
				if (iconValueType !== ConditionalFormatting.IconValueType.formula) {
					iconValue = +iconValue;
				}
				iconCriteria[i] = new ConditionalFormatting.IconCriterion(isGreaterThanOrEqualTo, iconValueType, iconValue);
			});
			iconSetRule.reverseIconOrder(utilities.getCheckValue("reverseIconOrder"));
			iconSetRule.showIconOnly(utilities.getCheckValue("showIconOnly"));
			cfs.addRule(iconSetRule);
		}

		sheet.resumePaint();
	}

	addCondionalFormaterRule(rule:any) {
		var sheet = this.spread.getActiveSheet();
		var sels = sheet.getSelections();
		var style = new GC.Spread.Sheets.Style();

		if (utilities.getCheckValue("useFormatBackColor")) {
			style.backColor = utilities.getBackgroundColor("formatBackColor");
		}
		if (utilities.getCheckValue("useFormatForeColor")) {
			style.foreColor = utilities.getBackgroundColor("formatForeColor");
		}
		if (utilities.getCheckValue("useFormatBorder")) {
			var lineBorder = new GC.Spread.Sheets.LineBorder(utilities.getBackgroundColor("formatBorderColor"), GC.Spread.Sheets.LineStyle.thin);
			style.borderTop = style.borderRight = style.borderBottom = style.borderLeft = lineBorder;
		}
		var value1 = $("#value1").val();
		var value2 = $("#value2").val();
		var cfs = sheet.conditionalFormats;
		var operator = +utilities.getDropDownValue("comparisonOperator");

		var minType = +utilities.getDropDownValue("minType");
		var midType = +utilities.getDropDownValue("midType");
		var maxType = +utilities.getDropDownValue("maxType");
		var midColor = utilities.getBackgroundColor("midColor");
		var minColor = utilities.getBackgroundColor("minColor");
		var maxColor = utilities.getBackgroundColor("maxColor");
		var midValue = utilities.getNumberValue("midValue");
		var maxValue = utilities.getNumberValue("maxValue");
		var minValue = utilities.getNumberValue("minValue");

		switch (rule) {
			case "0":
				var doubleValue1 = parseFloat(value1);
				var doubleValue2 = parseFloat(value2);
				cfs.addCellValueRule(operator, isNaN(doubleValue1) ? value1 : doubleValue1, isNaN(doubleValue2) ? value2 : doubleValue2, style, sels);
				break;
			case "1":
				cfs.addSpecificTextRule(operator, value1, style, sels);
				break;
			case "2":
				cfs.addDateOccurringRule(operator, style, sels);
				break;
			case "4":
				cfs.addTop10Rule(operator, parseInt(value1, 10), style, sels);
				break;
			case "5":
				cfs.addUniqueRule(style, sels);
				break;
			case "6":
				cfs.addDuplicateRule(style, sels);
				break;
			case "7":
				cfs.addAverageRule(operator, style, sels);
				break;
			case "8":
				cfs.add2ScaleRule(minType, minValue, minColor, maxType, maxValue, maxColor, sels);
				break;
			case "9":
				cfs.add3ScaleRule(minType, minValue, minColor, midType, midValue, midColor, maxType, maxValue, maxColor, sels);
				break;
			default:
				var doubleValue1 = parseFloat(value1);
				var doubleValue2 = parseFloat(value2);
				cfs.addCellValueRule(operator, isNaN(doubleValue1) ? value1 : doubleValue1, isNaN(doubleValue2) ? value2 : doubleValue2, style, sels);
				break;
    	}

    	sheet.repaint();
	}

	addDataBarRule() {
		var sheet = this.spread.getActiveSheet();
		sheet.suspendPaint();

		var selections = sheet.getSelections();
		if (selections.length > 0) {
			var ranges:any[] = [];
			$.each(selections, function (i, v) {
				ranges.push(new GC.Spread.Sheets.Range(v.row, v.col, v.rowCount, v.colCount));
			});
			var cfs = sheet.conditionalFormats;
			var dataBarRule = new ConditionalFormatting.DataBarRule(
				+utilities.getDropDownValue("minimumType"), 
				utilities.getNumberValue("minimumValue"), 
				+utilities.getDropDownValue("maximumType"),
				utilities.getNumberValue("maximumValue"), null, ranges);
			
			dataBarRule.gradient(utilities.getCheckValue("gradient"));
			dataBarRule.color(utilities.getBackgroundColor("gradientColor"));
			dataBarRule.showBorder(utilities.getCheckValue("showBorder"));
			dataBarRule.borderColor(utilities.getBackgroundColor("barBorderColor"));
			dataBarRule.dataBarDirection(+utilities.getDropDownValue("dataBarDirection"));
			dataBarRule.negativeFillColor(utilities.getBackgroundColor("negativeFillColor"));
			dataBarRule.useNegativeFillColor(utilities.getCheckValue("useNegativeFillColor"));
			dataBarRule.negativeBorderColor(utilities.getBackgroundColor("negativeBorderColor"));
			dataBarRule.useNegativeBorderColor(utilities.getCheckValue("useNegativeBorderColor"));
			dataBarRule.axisPosition(+utilities.getDropDownValue("axisPosition"));
			dataBarRule.axisColor(utilities.getBackgroundColor("barAxisColor"));
			dataBarRule.showBarOnly(utilities.getCheckValue("showBarOnly"));
			cfs.addRule(dataBarRule);
		}

		sheet.resumePaint();
	}

	attachBorderTypeClickEvents() {
		let component = this;
		let $groupItems = $(".group-item>div");
		$groupItems.bind("mousedown", function () {
			if ($(this).parent().hasClass("disable")) {
				return;
			}
			var name = $(this).data("name").split("Border")[0];
			component.applyBorderSetting(name);
		});
	}

	applyBorderSetting(name:string) {
		var sheet = this.spread.getActiveSheet();
    	var borderLine = this.getBorderLineType($("#border-line-type").attr("class"));
    	var borderColor = utilities.getBackgroundColor("borderColor");
    	this.setBorderlines(sheet, name, borderLine, borderColor);
	}

	setBorderlines(sheet:any, borderType:any, borderStyle:any, borderColor:any) {
		let component = this;
		function setSheetBorder(setting:any) {
			var lineBorder = new GC.Spread.Sheets.LineBorder(borderColor, setting.lineStyle);
			sel.setBorder(lineBorder, setting.options);
			component.setRangeBorder(sheet, sel, setting.options);
		}

		var settings = this.getBorderSettings(borderType, borderStyle);
		var rowCount = sheet.getRowCount(),
			columnCount = sheet.getColumnCount();

		sheet.suspendPaint();
		var sels = sheet.getSelections();

		for (var n = 0; n < sels.length; n++) {
			var sel = this.getActualCellRange(sheet, sels[n], rowCount, columnCount);
			settings.forEach(setSheetBorder);
		}
		sheet.resumePaint();
	}

	getBorderSettings(borderType:any, borderStyle:any) {
		var result = [];

		switch (borderType) {
			case "outside":
				result.push({lineStyle: borderStyle, options: {outline: true}});
				break;

			case "inside":
				result.push({lineStyle: borderStyle, options: {innerHorizontal: true}});
				result.push({lineStyle: borderStyle, options: {innerVertical: true}});
				break;

			case "all":
			case "none":
				result.push({lineStyle: borderStyle, options: {all: true}});
				break;

			case "left":
				result.push({lineStyle: borderStyle, options: {left: true}});
				break;

			case "innerVertical":
				result.push({lineStyle: borderStyle, options: {innerVertical: true}});
				break;

			case "right":
				result.push({lineStyle: borderStyle, options: {right: true}});
				break;

			case "top":
				result.push({lineStyle: borderStyle, options: {top: true}});
				break;

			case "innerHorizontal":
				result.push({lineStyle: borderStyle, options: {innerHorizontal: true}});
				break;

			case "bottom":
				result.push({lineStyle: borderStyle, options: {bottom: true}});
				break;
		}

		return result;
	}

	setRangeBorder(sheet:any, range:any, options:any) {
		var outline = options.all || options.outline,
			rowCount = sheet.getRowCount(),
			columnCount = sheet.getColumnCount(),
			startRow = range.row, endRow = startRow + range.rowCount - 1,
			startCol = range.col, endCol = startCol + range.colCount - 1;

		// update related borders for all cells arround the range

		// left side 
		if ((startCol > 0) && (outline || options.left)) {
			sheet.getRange(startRow, startCol - 1, range.rowCount, 1).borderRight(undefined);
		}
		// top side
		if ((startRow > 0) && (outline || options.top)) {
			sheet.getRange(startRow - 1, startCol, 1, range.colCount).borderBottom(undefined);
		}
		// right side
		if ((endCol < columnCount - 1) && (outline || options.right)) {
			sheet.getRange(startRow, endCol + 1, range.rowCount, 1).borderLeft(undefined);
		}
		// bottom side
		if ((endRow < rowCount - 1) && (outline || options.bottom)) {
			sheet.getRange(endRow + 1, startCol, 1, range.colCount).borderTop(undefined);
		}
	}

	getBorderLineType(className:string) {
		switch (className) {
			case "no-border":
				return GC.Spread.Sheets.LineStyle.empty;

			case "line-style-hair":
				return GC.Spread.Sheets.LineStyle.hair;

			case "line-style-dotted":
				return GC.Spread.Sheets.LineStyle.dotted;

			case "line-style-dash-dot-dot":
				return GC.Spread.Sheets.LineStyle.dashDotDot;

			case "line-style-dash-dot":
				return GC.Spread.Sheets.LineStyle.dashDot;

			case "line-style-dashed":
				return GC.Spread.Sheets.LineStyle.dashed;

			case "line-style-thin":
				return GC.Spread.Sheets.LineStyle.thin;

			case "line-style-medium-dash-dot-dot":
				return GC.Spread.Sheets.LineStyle.mediumDashDotDot;

			case "line-style-slanted-dash-dot":
				return GC.Spread.Sheets.LineStyle.slantedDashDot;

			case "line-style-medium-dash-dot":
				return GC.Spread.Sheets.LineStyle.mediumDashDot;

			case "line-style-medium-dashed":
				return GC.Spread.Sheets.LineStyle.mediumDashed;

			case "line-style-medium":
				return GC.Spread.Sheets.LineStyle.medium;

			case "line-style-thick":
				return GC.Spread.Sheets.LineStyle.thick;

			case "line-style-double":
				return GC.Spread.Sheets.LineStyle.double;
		}
	}

	attachSpreadEvents(rebind?:boolean) {
		this.spread.bind(GC.Spread.Sheets.Events.EnterCell, () => this.onCellSelected());

		this.spread.bind(GC.Spread.Sheets.Events.ValueChanged, (sender:any, args:any) => {
			var row = args.row, col = args.col, sheet = args.sheet;

			if (sheet.getCell(row, col).wordWrap()) {
				sheet.autoFitRow(row);
			}
		});

		function shouldAutofitRow(sheet:any, row:any, col:any, colCount:any) {
			for (var c = 0; c < colCount; c++) {
				if (sheet.getCell(row, col++).wordWrap()) {
					return true;
				}
			}

			return false;
		}

		this.spread.bind(GC.Spread.Sheets.Events.RangeChanged, (sender:any, args:any) => {
			var sheet = args.sheet, row = args.row, rowCount = args.rowCount;

			if (args.action === GC.Spread.Sheets.RangeChangedAction.paste) {
				var col = args.col, colCount = args.colCount;
				for (var i = 0; i < rowCount; i++) {
					if (shouldAutofitRow(sheet, row, col, colCount)) {
						sheet.autoFitRow(row);
					}
					row++;
				}
			}
		});

		this.spread.bind(GC.Spread.Sheets.Events.SelectionChanging, () => {
			var sheet = this.spread.getActiveSheet();
			var selection = sheet.getSelections().slice(-1)[0];
			if (selection) {
				var position = this.getSelectedRangeString(sheet, selection);
				$("#positionbox").val(position);
			}
			this.syncDisabledBorderType();
		});

		this.spread.bind(GC.Spread.Sheets.Events.SelectionChanged, () => {
			this.syncCellRelatedItems();
			this.updatePositionBox(this.spread.getActiveSheet());
		});

		$(document).bind("keydown", (event) => {
			if (event.shiftKey) {
				this.isShiftKey = true;
			}
		});
		$(document).bind("keyup", (event) => {
			if (!event.shiftKey) {
				this.isShiftKey = false;

				var sheet = this.spread.getActiveSheet();
				var	position = this.getCellPositionString(sheet, sheet.getActiveRowIndex() + 1, sheet.getActiveColumnIndex() + 1);
				$("#positionbox").val(position);
			}
		});
	}

	// Cell Type related items
	attachCellTypeEvents() {
		$("#setCellTypeBtn").click(() => {
			let currentCellType:any = utilities.getDropDownValue("cellTypes");
			this.applyCellType(currentCellType);
		});
	}

	toggleInspector() {
		if ($(".insp-container:visible").length > 0) {
			$(".insp-container").hide();
			if (!this.floatInspector) {
				$("#inner-content-container").css({right: 0});
				$("span", this).removeClass("fa-angle-right fa-angle-up fa-angle-down").addClass("fa-angle-left");
			} else {
				$("#inner-content-container").css({right: 0});
				$("span", this).removeClass("fa-angle-right fa-angle-left fa-angle-up").addClass("fa-angle-down");
			}

			$(this).attr("title", uiResource.toolBar.showInspector);
		} else {
			$(".insp-container").show();
			if (!this.floatInspector) {
				$("#inner-content-container").css({right: "301px"});
				$("span", this).removeClass("fa-angle-left fa-angle-up fa-angle-down").addClass("fa-angle-right");
			} else {
				$("#inner-content-container").css({right: 0});
				$("span", this).removeClass("fa-angle-right fa-angle-left fa-angle-down").addClass("fa-angle-up");
			}

			$(this).attr("title", uiResource.toolBar.hideInspector);
		}
		this.spread.refresh();
	}

	attachToolbarItemEvents() {
		$("#addtable").click(() => {
			let sheet = this.spread.getActiveSheet();
			let	row = sheet.getActiveRowIndex();
			let	column = sheet.getActiveColumnIndex();
			let	name = "Table" + this.tableIndex;
			let	rowCount = 1;
			let	colCount = 1;

			this.tableIndex++;

			var selections = sheet.getSelections();

			if (selections.length > 0) {
				var range = selections[0],
					r = range.row,
					c = range.col;

				rowCount = range.rowCount,
					colCount = range.colCount;

				// update row / column for whole column / row was selected
				if (r >= 0) {
					row = r;
				}
				if (c >= 0) {
					column = c;
				}
			}

			sheet.suspendPaint();
			try {
				// handle exception if the specified range intersect with other table etc.
				sheet.tables.add(name, row, column, rowCount, colCount, GC.Spread.Sheets.Tables.TableThemes.light2);
			} catch (e) {
				alert(e.message);
			}
			sheet.resumePaint();

			this.spread.focus();

			this.onCellSelected();
		});

		$("#addcomment").click(() => {
			var sheet = this.spread.getActiveSheet(),
				row = sheet.getActiveRowIndex(),
				column = sheet.getActiveColumnIndex(),
				comment;

			sheet.suspendPaint();
			comment = sheet.comments.add(row, column, new Date().toLocaleString());
			sheet.resumePaint();

			comment.commentState(GC.Spread.Sheets.Comments.CommentState.edit);
		});

		$("#addpicture, #doImport").click(function () {
			$("#fileSelector").data("action", this.id);
			$("#fileSelector").click();
		});

		$("#toggleInspector").click(() => this.toggleInspector());

		$("#doClear").click(() => {
			var $dropdown = $("#clearActionList"),
				$this = $(this),
				offset = $this.offset();

			$dropdown.css({left: offset.left, top: offset.top + $this.outerHeight()});
			$dropdown.show();
			this.processEventListenerHandleClosePopup(true);
		});

		$("#doExport").click(() => {
			var $dropdown = $("#exportActionList"),
				$this = $(this),
				offset = $this.offset();

			$dropdown.css({left: offset.left, top: offset.top + $this.outerHeight()});
			$dropdown.show();
			this.processEventListenerHandleClosePopup(true);
		});

		$("#addslicer").click(() => this.processAddSlicer());
	}

	// slicer related items
	processAddSlicer() {
		this.addTableColumns();                          // get table header data from table, and add them to slicer dialog

		var SLICER_DIALOG_WIDTH = 230;              // slicer dialog width
		this.showModal(uiResource.slicerDialog.insertSlicer, SLICER_DIALOG_WIDTH, $("#insertslicerdialog").children(), () => this.addSlicerEvent());
	}

	addSlicerEvent() {
		var table = this._activeTable;
		if (!table) {
			return;
		}
		var checkedColumnIndexArray:number[] = [];
		$("#slicer-container div.button").each(function (index) {
			if ($(this).hasClass("checked")) {
				checkedColumnIndexArray.push(index);
			}
		});
		var sheet = this.spread.getActiveSheet();
		var posX = 100, posY = 200;
		this.spread.suspendPaint();
		for (var i = 0; i < checkedColumnIndexArray.length; i++) {
			var columnName = table.getColumnName(checkedColumnIndexArray[i]);
			var slicerName = this.getSlicerName(sheet, columnName);
			var slicer:GC.Spread.Sheets.Slicers.Slicer = sheet.slicers.add(slicerName, table.name(), columnName, null);
			(<any>slicer).position(new GC.Spread.Sheets.Point(posX, posY));
			posX = posX + 30;
			posY = posY + 30;
		}
		this.spread.resumePaint();
		slicer.isSelected(true);
		this.initSlicerTab();
	}

	initSlicerTab() {
		var sheet = this.spread.getActiveSheet();
		var selectedSlicers = this.getSelectedSlicers(sheet);
		if (!selectedSlicers || selectedSlicers.length === 0) {
			return;
		}
		if (selectedSlicers.length > 1) {
			this.getMultiSlicerSetting(selectedSlicers);
			utilities.setTextDisabled("slicerName", true);
		}
		else if (selectedSlicers.length === 1) {
			this.getSingleSlicerSetting(selectedSlicers[0]);
			utilities.setTextDisabled("slicerName", false);
		}
	}

	getSingleSlicerSetting(slicer:any) {
		if (!slicer) {
			return;
		}
		utilities.setTextValue("slicerName", slicer.name());
		utilities.setTextValue("slicerCaptionName", slicer.captionName());
		utilities.setDropDownValue("slicerItemSorting", slicer.sortState());
		utilities.setCheckValue("displaySlicerHeader", slicer.showHeader());
		utilities.setNumberValue("slicerColumnNumber", slicer.columnCount());
		utilities.setNumberValue("slicerButtonWidth", this.getSlicerItemWidth(slicer.columnCount(), slicer.width()));
		utilities.setNumberValue("slicerButtonHeight", slicer.itemHeight());
		if (slicer.dynamicMove()) {
			if (slicer.dynamicSize()) {
				utilities.setRadioItemChecked("slicerMoveAndSize", "slicer-move-size");
			}
			else {
				utilities.setRadioItemChecked("slicerMoveAndSize", "slicer-move-nosize");
			}
		}
		else {
			utilities.setRadioItemChecked("slicerMoveAndSize", "slicer-nomove-size");
		}
		utilities.setCheckValue("lockSlicer", slicer.isLocked());
		utilities.selectedCurrentSlicerStyle(slicer);
	}

	getMultiSlicerSetting(selectedSlicers:any) {
		if (!selectedSlicers || selectedSlicers.length === 0) {
			return;
		}
		var slicer = selectedSlicers[0];
		var isDisplayHeader = false,
			isSameSortState = true,
			isSameCaptionName = true,
			isSameColumnCount = true,
			isSameItemHeight = true,
			isSameItemWidth = true,
			isSameLocked = true,
			isSameDynamicMove = true,
			isSameDynamicSize = true;

		var sortState = slicer.sortState(),
			captionName = slicer.captionName(),
			columnCount = slicer.columnCount(),
			itemHeight = slicer.itemHeight(),
			itemWidth = this.getSlicerItemWidth(columnCount, slicer.width()),
			dynamicMove = slicer.dynamicMove(),
			dynamicSize = slicer.dynamicSize();

		for (var item in selectedSlicers) {
			var slicer = selectedSlicers[item];
			isDisplayHeader = isDisplayHeader || slicer.showHeader();
			isSameLocked = isSameLocked && slicer.isLocked();
			if (slicer.sortState() !== sortState) {
				isSameSortState = false;
			}
			if (slicer.captionName() !== captionName) {
				isSameCaptionName = false;
			}
			if (slicer.columnCount() !== columnCount) {
				isSameColumnCount = false;
			}
			if (slicer.itemHeight() !== itemHeight) {
				isSameItemHeight = false;
			}
			if (this.getSlicerItemWidth(slicer.columnCount(), slicer.width()) !== itemWidth) {
				isSameItemWidth = false;
			}
			if (slicer.dynamicMove() !== dynamicMove) {
				isSameDynamicMove = false;
			}
			if (slicer.dynamicSize() !== dynamicSize) {
				isSameDynamicSize = false;
			}
			utilities.selectedCurrentSlicerStyle(slicer);
		}
		utilities.setTextValue("slicerName", "");
		if (isSameCaptionName) {
			utilities.setTextValue("slicerCaptionName", captionName);
		}
		else {
			utilities.setTextValue("slicerCaptionName", "");
		}
		if (isSameSortState) {
			utilities.setDropDownValue("slicerItemSorting", sortState);
		}
		else {
			utilities.setDropDownValue("slicerItemSorting", "");
		}
		utilities.setCheckValue("displaySlicerHeader", isDisplayHeader);
		if (isSameDynamicMove && isSameDynamicSize && dynamicMove) {
			if (dynamicSize) {
				utilities.setRadioItemChecked("slicerMoveAndSize", "slicer-move-size");
			}
			else {
				utilities.setRadioItemChecked("slicerMoveAndSize", "slicer-move-nosize");
			}
		}
		else {
			utilities.setRadioItemChecked("slicerMoveAndSize", "slicer-nomove-size");
		}
		if (isSameColumnCount) {
			utilities.setNumberValue("slicerColumnNumber", columnCount);
		}
		else {
			utilities.setNumberValue("slicerColumnNumber", "");
		}
		if (isSameItemHeight) {
			utilities.setNumberValue("slicerButtonHeight", Math.round(itemHeight));
		}
		else {
			utilities.setNumberValue("slicerButtonHeight", "");
		}
		if (isSameItemWidth) {
			utilities.setNumberValue("slicerButtonWidth", itemWidth);
		}
		else {
			utilities.setNumberValue("slicerButtonWidth", "");
		}
		utilities.setCheckValue("lockSlicer", isSameLocked);
	}

	getSlicerItemWidth(count:any, slicerWidth:any) {
		if (count <= 0) {
			count = 1; //Column count will be converted to 1 if it is set to 0 or negative number.
		}
		var SLICER_PADDING = 6;
		var SLICER_ITEM_SPACE = 2;
		var itemWidth = Math.round((slicerWidth - SLICER_PADDING * 2 - (count - 1) * SLICER_ITEM_SPACE) / count);
		if (itemWidth < 0) {
			return 0;
		}
		else {
			return itemWidth;
		}
	}

	getSelectedSlicers(sheet:any) {
		if (!sheet) {
			return null;
		}
		var slicers = sheet.slicers.all();
		if (!slicers || slicers.length === 0) {
			return null;
		}
		var selectedSlicers = [];
		for (var item in slicers) {
			if (slicers[item].isSelected()) {
				selectedSlicers.push(slicers[item]);
			}
		}
		return selectedSlicers;
	}

	getSlicerName(sheet:any, columnName:any) {
		var autoID = 1;
		var newName = columnName;
		while (sheet.slicers.get(newName)) {
			newName = columnName + '_' + autoID;
			autoID++;
		}
		return newName;
	}

	showModal(title:any, width:any, content:any, callback:any) {
		var $dialog = $("#modalTemplate"),
			$body = $(".modal-body", $dialog);

		$(".modal-title", $dialog).text(title);
		$dialog.data("content-parent", content.parent());
		$body.append(content);

		// remove old and add new event handler since this modal is common used (reused)
		$("#dialogConfirm").off("click");
		$("#dialogConfirm").on("click", function () {
			var result = callback();

			// return an object with  { canceled: true } to tell not close the modal, otherwise close the modal
			if (!(result && result.canceled)) {
				(<any>$("#modalTemplate")).modal("hide");
			}
		});

		if (!$dialog.data("event-attached")) {
			$dialog.on("hidden.bs.modal", function () {
				var $originalParent = $(this).data("content-parent");
				if ($originalParent) {
					$originalParent.append($(".modal-body", this).children());
				}
			});
			$dialog.data("event-attached", true);
		}

		// set width of the dialog
		$(".modal-dialog", $dialog).css({width: width});

		(<any>$dialog).modal("show");
	}

	addTableColumns() {
		var table = this._activeTable;
		if (!table) {
			return;
		}
		var $slicerContainer = $("#slicer-container");
		$slicerContainer.empty();
		for (var col = 0; col < table.range().colCount; col++) {
			var columnName = table.getColumnName(col);
			var $slicerDiv = $(
				"<div>"
				+ "<div class='insp-row'>"
				+ "<div>"
				+ "<div class='insp-checkbox insp-inline-row'>"
				+ "<div class='button insp-inline-row-item'></div>"
				+ "<div class='text insp-inline-row-item localize'>" + columnName + "</div>"
				+ "</div>"
				+ "</div>"
				+ "</div>"
				+ "</div>");
			$slicerDiv.appendTo($slicerContainer);
		}
		let component = this;
		$("#slicer-container .insp-checkbox").click(function() {
			component.checkedChanged(this)
		});
	}

	processEventListenerHandleClosePopup(add:boolean) {
		let _handlePopupCloseEvents = 'mousedown touchstart MSPointerDown pointerdown'.split(' ');

		if (add) {
			_handlePopupCloseEvents.forEach((value) => {
				document.addEventListener(value, this._documentMousedownHandler, true);
			});
		} else {
			_handlePopupCloseEvents.forEach((value) => {
				document.removeEventListener(value, this._documentMousedownHandler, true);
			});
		}
	}

	documentMousedownHandler(event:any) {
		var target = event.target,
			container = this._dropdownitem || this._colorpicker || $("#clearActionList:visible")[0] || $("#exportActionList:visible")[0];

		if (container) {
			if (container === target || $.contains(container, target)) {
				return;
			}

			// click on related item popup the dropdown, close it
			var dropdown = (<any>$(container)).data("dropdown");
			if (dropdown && $.contains(dropdown, target)) {
				this.hidePopups();
				this._needShow = false;
				return false;
			}
		}

		this.hidePopups();
		$("#passwordError").hide();
	}

	hidePopups() {
		this.hideDropdown();
		this.hideColorPicker();
	}

	hideColorPicker() {
		if (this._colorpicker) {
			$(this._colorpicker).removeClass("colorpicker-visible");
			this._colorpicker = null;
		}
		this.processEventListenerHandleClosePopup(false);
	}

	onCellSelected() {
		$("#addslicer").addClass("hidden");
		let sheet = this.spread.getActiveSheet();
		let	row = sheet.getActiveRowIndex();
		let	column = sheet.getActiveColumnIndex();

		let cellInfo = this.getCellInfo(sheet, row, column);
		let	cellType = cellInfo.type;

		this.syncCellRelatedItems();
		this.updatePositionBox(sheet);
		this.updateCellStyleState(sheet, row, column);

		let tabType = "cell";

		this.clearCachedItems();

		// add map from cell type to tab type here
		if (cellType === "table") {
			tabType = "table";
			$("#addslicer").removeClass("hidden");
		} else if (cellType === "comment") {
			tabType = "comment";
		}
	}

	clearCachedItems() {
		this._activeComment = null;
		this._activeTable = null;
	}

	syncCellRelatedItems() {
		this.updateMergeButtonsState();
		this.syncDisabledLockCells();
		this.syncDisabledBorderType();

		// reset conditional format setting
		var item = utilities.setDropDownValueByIndex($("#conditionalFormatType"), -1);
		this.processConditionalFormatDetailSetting(item.value, true);
		// sync cell type related information
		this.syncCellTypeInfo();
	}

	syncDisabledBorderType() {
		var sheet = this.spread.getActiveSheet();
		var selections = sheet.getSelections(), selectionsLength = selections.length;
		var isDisabledInsideBorder = true;
		var isDisabledHorizontalBorder = true;
		var isDisabledVerticalBorder = true;
		for (var i = 0; i < selectionsLength; i++) {
			var selection = selections[i];
			var col = selection.col, row = selection.row,
				rowCount = selection.rowCount, colCount = selection.colCount;
			if (isDisabledHorizontalBorder) {
				isDisabledHorizontalBorder = rowCount === 1;
			}
			if (isDisabledVerticalBorder) {
				isDisabledVerticalBorder = colCount === 1;
			}
			if (isDisabledInsideBorder) {
				isDisabledInsideBorder = rowCount === 1 || colCount === 1;
			}
		}
		[isDisabledInsideBorder, isDisabledVerticalBorder, isDisabledHorizontalBorder].forEach(function (value, index) {
			var $item = $("div.group-item:eq(" + (index * 3 + 1) + ")");
			if (value) {
				$item.addClass("disable");
			} else {
				$item.removeClass("disable");
			}
		});
	}

	syncDisabledLockCells() {
		var cellsLockedState = this.getCellsLockedState();
		utilities.setCheckValue("checkboxLockCell", cellsLockedState);
	}

	getCellsLockedState() {
		var isLocked = false;
		var sheet = this.spread.getActiveSheet();
		var selections = sheet.getSelections(), selectionsLength = selections.length;
		var cell;
		var row, col, rowCount, colCount;
		if (selectionsLength > 0) {
			for (var i = 0; i < selectionsLength; i++) {
				var range = selections[i];
				row = range.row;
				rowCount = range.rowCount;
				colCount = range.colCount;
				if (row < 0) {
					row = 0;
				}
				for (row; row < range.row + rowCount; row++) {
					col = range.col;
					if (col < 0) {
						col = 0;
					}
					for (col; col < range.col + colCount; col++) {
						cell = sheet.getCell(row, col);
						isLocked = isLocked || cell.locked();
						if (isLocked) {
							return isLocked;
						}
					}
				}
			}
			return false;
		} else {
			return sheet.getCell(sheet.getActiveRowIndex(), sheet.getActiveColumnIndex()).locked();
		}
	}

	updateMergeButtonsState() {
		var sheet = this.spread.getActiveSheet();
		var sels = sheet.getSelections(),
			mergable = false,
			unmergable = false;

		sels.forEach(function (range) {
			var ranges = sheet.getSpans(range),
				spanCount = ranges.length;

			if (!mergable) {
				if (spanCount > 1 || (spanCount === 0 && (range.rowCount > 1 || range.colCount > 1))) {
					mergable = true;
				} else if (spanCount === 1) {
					var range2 = ranges[0];
					if (range2.row !== range.row || range2.col !== range.col ||
						range2.rowCount !== range2.rowCount || range2.colCount !== range.colCount) {
                    		mergable = true;
						}
            	}
       		}
       		if (!unmergable) {
       	    	unmergable = spanCount > 0;
       		}
		});

		$("#mergeCells").attr("disabled", mergable ? null : "disabled");
		$("#unmergeCells").attr("disabled", unmergable ? null : "disabled");
	}

	syncCellTypeInfo() {
		function updateButtonCellTypeInfo(cellType:any) {
			utilities.setNumberValue("buttonCellTypeMarginTop", cellType.marginTop());
			utilities.setNumberValue("buttonCellTypeMarginRight", cellType.marginRight());
			utilities.setNumberValue("buttonCellTypeMarginBottom", cellType.marginBottom());
			utilities.setNumberValue("buttonCellTypeMarginLeft", cellType.marginLeft());
			utilities.setTextValue("buttonCellTypeText", cellType.text());
			utilities.setColorValue("buttonCellTypeBackColor", cellType.buttonBackColor());
		}

		function updateCheckBoxCellTypeInfo(cellType:any) {
			utilities.setTextValue("checkboxCellTypeCaption", cellType.caption());
			utilities.setTextValue("checkboxCellTypeTextTrue", cellType.textTrue());
			utilities.setTextValue("checkboxCellTypeTextIndeterminate", cellType.textIndeterminate());
			utilities.setTextValue("checkboxCellTypeTextFalse", cellType.textFalse());
			utilities.setDropDownValue("checkboxCellTypeTextAlign", cellType.textAlign());
			utilities.setCheckValue("checkboxCellTypeIsThreeState", cellType.isThreeState());
		}

		function updateComboBoxCellTypeInfo(cellType:any) {
			utilities.setDropDownValue("comboboxCellTypeEditorValueType", cellType.editorValueType());
			var items = cellType.items(),
				texts = items.map(function (item:any) {
					return item.text || item;
				}).join(","),
				values = items.map(function (item:any) {
					return item.value || item;
				}).join(",");

			utilities.setTextValue("comboboxCellTypeItemsText", texts);
			utilities.setTextValue("comboboxCellTypeItemsValue", values);
		}

		function updateHyperLinkCellTypeInfo(cellType:any) {
			utilities.setColorValue("hyperlinkCellTypeLinkColor", cellType.linkColor());
			utilities.setColorValue("hyperlinkCellTypeVisitedLinkColor", cellType.visitedLinkColor());
			utilities.setTextValue("hyperlinkCellTypeText", cellType.text());
			utilities.setTextValue("hyperlinkCellTypeLinkToolTip", cellType.linkToolTip());
		}

		var sheet = this.spread.getActiveSheet(),
			index,
			cellType = sheet.getCell(sheet.getActiveRowIndex(), sheet.getActiveColumnIndex()).cellType();

		if (cellType instanceof GC.Spread.Sheets.CellTypes.Button) {
			index = 0;
			updateButtonCellTypeInfo(cellType);
		} else if (cellType instanceof GC.Spread.Sheets.CellTypes.CheckBox) {
			index = 1;GC.Spread.Sheets
			updateCheckBoxCellTypeInfo(cellType);
		} else if (cellType instanceof GC.Spread.Sheets.CellTypes.ComboBox) {
			index = 2;
			updateComboBoxCellTypeInfo(cellType);
		} else if (cellType instanceof GC.Spread.Sheets.CellTypes.HyperLink) {
			index = 3;
			updateHyperLinkCellTypeInfo(cellType);
		} else {
			index = -1;
		}
		var cellTypeItem = utilities.setDropDownValueByIndex($("#cellTypes"), index);
		this.processCellTypeSetting(cellTypeItem.value, true);

		if (index >= 0) {
			var $group = $("#groupCellType");
			if ($group.find(".group-state").hasClass("fa-caret-right")) {
				$group.click();
			}
		}
	}

	getCellInfo(sheet:any, row:any, column:any): {type:string, object:any} {
		let result:{type:string, object:any} = {type: "", object: null};
		let object;

		if ((object = sheet.comments.get(row, column))) {
			result.type = "comment";
		} else if ((object = sheet.tables.find(row, column))) {
			result.type = "table";
		}

		result.object = object;

		return result;
	}

	applyCellType(name:string|GC.Spread.Sheets.CellTypes.CheckBoxTextAlign) {
		var sheet = this.spread.getActiveSheet();
		var cellType;
		switch (name) {
			case "button-celltype":
				cellType = new GC.Spread.Sheets.CellTypes.Button();
				cellType.marginTop(utilities.getNumberValue("buttonCellTypeMarginTop"));
				cellType.marginRight(utilities.getNumberValue("buttonCellTypeMarginRight"));
				cellType.marginBottom(utilities.getNumberValue("buttonCellTypeMarginBottom"));
				cellType.marginLeft(utilities.getNumberValue("buttonCellTypeMarginLeft"));
				cellType.text(utilities.getTextValue("buttonCellTypeText"));
				cellType.buttonBackColor(utilities.getBackgroundColor("buttonCellTypeBackColor"));
				break;

			case "checkbox-celltype":
				cellType = new GC.Spread.Sheets.CellTypes.CheckBox();
				cellType.caption(utilities.getTextValue("checkboxCellTypeCaption"));
				cellType.textTrue(utilities.getTextValue("checkboxCellTypeTextTrue"));
				cellType.textIndeterminate(utilities.getTextValue("checkboxCellTypeTextIndeterminate"));
				cellType.textFalse(utilities.getTextValue("checkboxCellTypeTextFalse"));
				cellType.textAlign(<GC.Spread.Sheets.CellTypes.CheckBoxTextAlign>utilities.getDropDownValue("checkboxCellTypeTextAlign"));
				cellType.isThreeState(utilities.getCheckValue("checkboxCellTypeIsThreeState"));
				break;

			case "combobox-celltype":
				cellType = new GC.Spread.Sheets.CellTypes.ComboBox();
				cellType.editorValueType(<GC.Spread.Sheets.CellTypes.EditorValueType>utilities.getDropDownValue("comboboxCellTypeEditorValueType"));
				var comboboxItemsText = utilities.getTextValue("comboboxCellTypeItemsText");
				var comboboxItemsValue = utilities.getTextValue("comboboxCellTypeItemsValue");
				var itemsText = comboboxItemsText.split(",");
				var itemsValue = comboboxItemsValue.split(",");
				var itemsLength = itemsText.length > itemsValue.length ? itemsText.length : itemsValue.length;
				var items = [];
				for (var count = 0; count < itemsLength; count++) {
					var t = itemsText.length > count && itemsText[0] !== "" ? itemsText[count] : undefined;
					var v = itemsValue.length > count && itemsValue[0] !== "" ? itemsValue[count] : undefined;
					if (t !== undefined && v !== undefined) {
						items[count] = {text: t, value: v};
					}
					else if (t !== undefined) {
						items[count] = {text: t};
					} else if (v !== undefined) {
						items[count] = {value: v};
					}
				}
				cellType.items(items);
				break;

			case "hyperlink-celltype":
				cellType = new GC.Spread.Sheets.CellTypes.HyperLink();
				cellType.linkColor(utilities.getBackgroundColor("hyperlinkCellTypeLinkColor"));
				cellType.visitedLinkColor(utilities.getBackgroundColor("hyperlinkCellTypeVisitedLinkColor"));
				cellType.text(utilities.getTextValue("hyperlinkCellTypeText"));
				cellType.linkToolTip(utilities.getTextValue("hyperlinkCellTypeLinkToolTip"));
				break;
		}
		sheet.suspendPaint();
		sheet.suspendEvent();
		var sels = sheet.getSelections();
		var rowCount = sheet.getRowCount(),
			columnCount = sheet.getColumnCount();

		for (var i = 0; i < sels.length; i++) {
			var sel = this.getActualCellRange(sheet, sels[i], rowCount, columnCount);
			for (var r = 0; r < sel.rowCount; r++) {
				for (var c = 0; c < sel.colCount; c++) {
					sheet.setCellType(sel.row + r, sel.col + c, cellType, GC.Spread.Sheets.SheetArea.viewport);
				}
			}
		}
		sheet.resumeEvent();
		sheet.resumePaint();
	}

	initSpread():void {
    	//formulabox
    	let fbx = new GC.Spread.Sheets.FormulaTextBox.FormulaTextBox(document.getElementById('formulabox'), {});
    	fbx.workbook(this.spread);

    	this.setCellContent();
	}
	
	setCellContent():void {
		let sheet = new GC.Spread.Sheets.Worksheet("Cell");
        this.spread.removeSheet(0);
        this.spread.addSheet(this.spread.getSheetCount(), sheet);

        sheet.suspendPaint();
        sheet.setColumnCount(50);

        sheet.setColumnWidth(0, 100);
        sheet.setColumnWidth(1, 20);
        for (let col = 2; col < 11; col++) {
            sheet.setColumnWidth(col, 88);
        }

        let Range = GC.Spread.Sheets.Range;
        let row = 1, col = 0;                               // cell background
        sheet.getCell(row, col).value("Background").font("700 11pt Calibri");
        sheet.getCell(row, col + 2).backColor("#1E90FF");
        sheet.getCell(row, col + 4).backColor("#00ff00");

        row = row + 2;                                      // line border
        let borderColor = "red";
        let lineStyle = GC.Spread.Sheets.LineStyle;
        let lineBorder = GC.Spread.Sheets.LineBorder;
        let option = {all: true};
        sheet.getCell(row, 0).value("Border").font("700 11pt Calibri");
        col = 1;
        sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.empty), option);
        sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.hair), option);
        sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.dotted), option);
        sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.dashDotDot), option);
        sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.dashDot), option);
        sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.dashed), option);
        sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.thin), option);
        row = row + 2, col = 1;
        sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.mediumDashDotDot), option);
        sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.slantedDashDot), option);
        sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.mediumDashDot), option);
        sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.mediumDashed), option);
        sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.medium), option);
        sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.thick), option);
        sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.double), option);
        row = row + 2, col = 1;
        sheet.getRange(row, ++col, 2, 2).setBorder(new lineBorder("blue", lineStyle.dashed), {all: true});
        sheet.getRange(row, col + 3, 2, 2).setBorder(new lineBorder("yellowgreen", lineStyle.double), {outline: true});
        sheet.getRange(row, col + 6, 2, 2).setBorder(new lineBorder("black", lineStyle.mediumDashed), {innerHorizontal: true});
        sheet.getRange(row, col + 6, 2, 2).setBorder(new lineBorder("black", lineStyle.slantedDashDot), {innerVertical: true});
        row = row + 3, col = 2;
        sheet.getRange(row, col, 3, 2).setBorder(new lineBorder("lightgreen", lineStyle.thick), {outline: true});
        sheet.getRange(row, col, 3, 2).setBorder(new lineBorder("lightgreen", lineStyle.thick), {innerHorizontal: true});
        col = col + 3;
        sheet.getRange(row, col, 3, 3).setBorder(new lineBorder("#CDCD00", lineStyle.thick), {outline: true});
        sheet.getRange(row, col, 3, 3).setBorder(new lineBorder("#CDCD00", lineStyle.thick), {innerVertical: true});

        row = row + 3, col = 1;                             // merge cell
        sheet.getCell(row + 1, 0).value("Span").font("700 11pt Calibri");
        sheet.addSpan(row + 1, ++col, 1, 2);
        sheet.addSpan(row, col + 3, 3, 1);
        sheet.addSpan(row, col + 5, 3, 2);

        row = row + 4, col = 1;                             // font
        let TextDecorationType = GC.Spread.Sheets.TextDecorationType;
        let fontText = "SPREADJS";
        sheet.getCell(row, 0).value("Font").font("700 11pt Calibri");
        sheet.getCell(row, ++col).value(fontText);
        sheet.getCell(row, ++col).value(fontText).font("13pt Calibri");
        sheet.getCell(row, ++col).value(fontText).font("11pt Arial");
        sheet.getCell(row, ++col).value(fontText).font("13pt Times New Roman");
        sheet.getCell(row, ++col).value(fontText).backColor("#FFD700");
        sheet.getCell(row, ++col).value(fontText).foreColor("#436EEE");
        row = row + 2, col = 1;
        sheet.getCell(row, ++col).value(fontText).foreColor("#FFD700").backColor("#436EEE");
        sheet.getCell(row, ++col).value(fontText).font("700 11pt Calibri");
        sheet.getCell(row, ++col).value(fontText).font("italic 11pt Calibri");
        sheet.getCell(row, ++col).value(fontText).textDecoration(TextDecorationType.underline);
        sheet.getCell(row, ++col).value(fontText).textDecoration(TextDecorationType.lineThrough);
        sheet.getCell(row, ++col).value(fontText).textDecoration(TextDecorationType.overline);

        row = row + 2, col = 1;                             // format
        let number = 0.25;
        sheet.getCell(row, 0).value("Format").font("700 11pt Calibri");
        sheet.getCell(row, ++col).value(number).formatter("0.00");
        sheet.getCell(row, ++col).value(number).formatter("$#,##0.00");
        sheet.getCell(row, ++col).value(number).formatter("$ #,##0.00;$ (#,##0.00);$ \"-\"??;@");
        sheet.getCell(row, ++col).value(number).formatter("0%");
        sheet.getCell(row, ++col).value(number).formatter("# ?/?");
        row = row + 2, col = 1;
        sheet.getCell(row, ++col).value(number).formatter("0.00E+00");
        sheet.getCell(row, ++col).value(number).formatter("@");
        sheet.getCell(row, ++col).value(number).formatter("h:mm:ss AM/PM");
        sheet.getCell(row, ++col).value(number).formatter("m/d/yyyy");
        sheet.getCell(row, ++col).value(number).formatter("dddd, mmmm dd, yyyy");

        row = row + 2, col = 1;                             // text alignment
        let HorizontalAlign = GC.Spread.Sheets.HorizontalAlign;
        let VerticalAlign = GC.Spread.Sheets.VerticalAlign;
        sheet.setRowHeight(row, 60);
        sheet.getCell(row, 0).value("Alignment").font("700 11pt Calibri");
        sheet.getCell(row, ++col).value("Top Left").vAlign(VerticalAlign.top).hAlign(HorizontalAlign.left);
        sheet.getCell(row, ++col).value("Top Center").vAlign(VerticalAlign.top).hAlign(HorizontalAlign.center);
        sheet.getCell(row, ++col).value("Top Right").vAlign(VerticalAlign.top).hAlign(HorizontalAlign.right);
        sheet.getCell(row, ++col).value("Center Left").vAlign(VerticalAlign.center).hAlign(HorizontalAlign.left);
        sheet.getCell(row, ++col).value("Center Center").vAlign(VerticalAlign.center).hAlign(HorizontalAlign.center);
        sheet.getCell(row, ++col).value("Center Right").vAlign(VerticalAlign.center).hAlign(HorizontalAlign.right);
        sheet.getCell(row, ++col).value("Bottom Left").vAlign(VerticalAlign.bottom).hAlign(HorizontalAlign.left);
        sheet.getCell(row, ++col).value("Bottom Center").vAlign(VerticalAlign.bottom).hAlign(HorizontalAlign.center);
        sheet.getCell(row, ++col).value("Bottom Right").vAlign(VerticalAlign.bottom).hAlign(HorizontalAlign.right);

        row = row + 2, col = 1;                             // lock cell
        sheet.getCell(row, 0).value("Locked").font("700 11pt Calibri");
        sheet.getCell(row, ++col).value("TRUE").locked(true);
        sheet.getCell(row, ++col).value("FALSE").locked(false);

        row = row + 2, col = 1;                             // word wrap
        sheet.setRowHeight(row, 60);
        sheet.getCell(row, 0).value("WordWrap").font("700 11pt Calibri");
        sheet.getCell(row, ++col).value("ABCDEFGHIJKLMNOPQRSTUVWXYZ").wordWrap(true);
        sheet.getCell(row, ++col).value("ABCDEFGHIJKLMNOPQRSTUVWXYZ").wordWrap(false);

        row = row + 2, col = 1;                             // celltype
        sheet.setRowHeight(row, 25);
        let cellType;
        sheet.getCell(row, 0).value("CellType").font("700 11pt Calibri");
        cellType = new GC.Spread.Sheets.CellTypes.Button();
        cellType.buttonBackColor("#FFFF00");
        cellType.text("I'm a button");
        sheet.getCell(row, ++col).cellType(cellType);

        cellType = new GC.Spread.Sheets.CellTypes.CheckBox();
        cellType.caption("caption");
        cellType.textTrue("true");
        cellType.textFalse("false");
        cellType.textIndeterminate("indeterminate");
        cellType.textAlign(GC.Spread.Sheets.CellTypes.CheckBoxTextAlign.right);
        cellType.isThreeState(true);
        sheet.getCell(row, ++col).cellType(cellType);

        cellType = new GC.Spread.Sheets.CellTypes.ComboBox();
        cellType.items(["apple", "banana", "cat", "dog"]);
        sheet.getCell(row, ++col).cellType(cellType);

        cellType = new GC.Spread.Sheets.CellTypes.HyperLink();
        cellType.linkColor("blue");
        cellType.visitedLinkColor("red");
        cellType.text("SpreadJS");
        cellType.linkToolTip("SpreadJS Web Site");
        sheet.getCell(row, ++col).cellType(cellType).value("http://spread.grapecity.com/Products/SpreadJS/");

        row = row + 2, col = 1;                             // celltype
        sheet.setRowHeight(row, 100);
        sheet.setColumnWidth(0, 150);
        sheet.getCell(row, 0).value("CellPadding&Label").font("700 11pt Calibri");
        sheet.getCell(row, ++col, GC.Spread.Sheets.SheetArea.viewport).watermark("User ID").cellPadding('20');
        sheet.getCell(row, col, GC.Spread.Sheets.SheetArea.viewport).labelOptions({
            foreColor: 'red',
            visibility: 2,
            font: 'bold 15px Arial'
        });

        let b = new GC.Spread.Sheets.CellTypes.Button();
        b.text("Click Me!");
        sheet.setColumnWidth(3, 200);
        sheet.setCellType(row, ++col, b, GC.Spread.Sheets.SheetArea.viewport);
        sheet.getCell(row, col, GC.Spread.Sheets.SheetArea.viewport).watermark("Button Cell Type").cellPadding('20 20');
        sheet.getCell(row, col, GC.Spread.Sheets.SheetArea.viewport).labelOptions({
            alignment: 2,
            visibility: 1,
            font: 'bold 15px Arial',
            foreColor: 'grey'
        });

        let c = new GC.Spread.Sheets.CellTypes.CheckBox();
        c.isThreeState(false);
        c.textTrue("Checked!");
        c.textFalse("Check Me!");
        sheet.setColumnWidth(4, 200);
        sheet.setCellType(row, ++col, c, GC.Spread.Sheets.SheetArea.viewport);
        sheet.getCell(row, col, GC.Spread.Sheets.SheetArea.viewport).watermark("CheckBox Cell Type").cellPadding('30');
        sheet.getCell(row, col, GC.Spread.Sheets.SheetArea.viewport).labelOptions({
            alignment: 5,
            visibility: 0,
            foreColor: 'green'
        });
        sheet.resumePaint();
	}

	checkMediaSize():void {
		let mql:MediaQueryList = window.matchMedia("screen and (max-width: 768px)");

		this.processMediaQueryResponse(mql);
		this.adjustInspectorDisplay();

		mql.addListener((v) => this.processMediaQueryResponse(v));
	}

	processMediaQueryResponse(mql:MediaQueryList):void {
		if (mql.matches) {
			if (!this.floatInspector) {
				this.floatInspector = true;
				this.adjustInspectorDisplay();
			}
		} else {
			if (this.floatInspector) {
				this.floatInspector = false;
				this.adjustInspectorDisplay();
			}
		}
	}

	adjustInspectorDisplay():void {
		let $inspectorContainer = $(".insp-container");
		let	$contentContainer = $("#inner-content-container");
		let	toggleInspectorClasses;
			
		if (this.floatInspector) {
			$inspectorContainer.draggable("enable");
			$inspectorContainer.addClass("float-inspector");
			$contentContainer.addClass("float-inspector");
			toggleInspectorClasses = ["fa-angle-down", "fa-angle-up"];
			$("#inner-content-container").css({right: 0});
		} else {
			$inspectorContainer.draggable("disable");
			$inspectorContainer.removeClass("float-inspector");
			$inspectorContainer.css({left: "auto", top: 0});
			$contentContainer.removeClass("float-inspector");
			toggleInspectorClasses = ["fa-angle-left", "fa-angle-right"];
		}
		
		// update toggleInspector
		let classIndex = ($(".insp-container:visible").length > 0) ? 1 : 0;
		$("#toggleInspector > span")
		.removeClass("fa-angle-left fa-angle-right fa-angle-up fa-angle-down")
        .addClass(toggleInspectorClasses[classIndex]);
	}

	screenAdoption():void {
		this.hideSpreadContextMenu();
		this.adjustSpreadSize();
		
		// adjust toolbar items position
		let $toolbar = $("#toolbar");
		let	sectionWidth = Math.floor($toolbar.width() / 3);
			
		$(".toolbar-left-section", $toolbar).width(sectionWidth);
			
		// + 2 to make sure the right section with enough space to show in same line
		if (sectionWidth > 375 + 2) {  // 340 = (380 + 300) / 2, where 380 is min-width of left section, 300 is the width of right section
			$(".toolbar-middle-section", $toolbar).width(sectionWidth);
		} else {
			$(".toolbar-middle-section", $toolbar).width("auto");
		}

    	// explicit set formula box' width instead of 100% because it's contained in table
    	let width = $("#inner-content-container").width() - $("#positionbox").outerWidth() - 1; // 1: border' width of td contains formulabox (left only)
    	$("#formulabox").css({width: width});
	}

	hideSpreadContextMenu():void {
		$("#spreadContextMenu").hide();
		$(document).off("mousedown.contextmenu");
	}

	adjustSpreadSize():void {
		let height = $("#inner-content-container").height() - $("#formulaBar").height() - MARGIN_BOTTOM,
		spreadHeight = $("#ss").height();

		if (spreadHeight !== height) {
			$("#controlPanel").height(height);
        	$("#ss").height(height);
			$("#ss").data("workbook").refresh();
    	}
	}

	doPrepareWork():void {
		/*
		1. expand / collapse .insp-group by checking expanded class
		*/
		function processDisplayGroups() {
			$("div.insp-group").each(function () {
				let $group = $(this),
				expanded = $group.hasClass("expanded"),
                $content = $group.find("div.insp-group-content"),
                $state = $group.find("span.group-state");

				if (expanded) {
					$content.show();
					$state.addClass("fa-caret-down");
				} else {
					$content.hide();
					$state.addClass("fa-caret-right");
				}
			});
		}
		
    	processDisplayGroups();

    	this.addEventHandlers();

    	$("input[type='number']:not('.not-min-zero')").attr("min", 0);

    	// set default values
    	let item = utilities.setDropDownValueByIndex($("#conditionalFormatType"), -1);
    	this.processConditionalFormatDetailSetting(item.value, true);
    	let cellTypeItem = utilities.setDropDownValueByIndex($("#cellTypes"), -1);
    	this.processCellTypeSetting(cellTypeItem.value, true);                     // CellType Setting

    	utilities.setDropDownValue("numberValidatorComparisonOperator", 0);       // NumberValidator Comparison Operator
    	utilities.processNumberValidatorComparisonOperatorSetting(0);
    	utilities.setDropDownValue("dateValidatorComparisonOperator", 0);         // DateValidator Comparison Operator
    	utilities.processDateValidatorComparisonOperatorSetting(0);
    	utilities.setDropDownValue("textLengthValidatorComparisonOperator", 0);   // TextLengthValidator Comparison Operator
    	utilities.processTextLengthValidatorComparisonOperatorSetting(0);
    	utilities.processBorderLineSetting("thin");                               // Border Line Setting

    	utilities.setDropDownValue("minType", 1);                                 // LowestValue
    	utilities.setDropDownValue("midType", 4);                                 // Percentile
    	utilities.setDropDownValue("maxType", 2);                                 // HighestValue
    	utilities.setDropDownValue("minimumType", 5);                             // Automin
    	utilities.setDropDownValue("maximumType", 7);                             // Automax
    	utilities.setDropDownValue("dataBarDirection", 0);                        // Left-to-Right
    	utilities.setDropDownValue("axisPosition", 0);                            // Automatic
    	utilities.setDropDownValue("iconSetType", 0);                             // ThreeArrowsColored
    	utilities.setDropDownValue("checkboxCellTypeTextAlign", 3);               // Right
    	utilities.setDropDownValue("comboboxCellTypeEditorValueType", 2);         // Value
    	utilities.setDropDownValue("errorAlert", 0);                              // Data Validation Error Alert Type
    	utilities.setDropDownValue("zoomSpread", 1);                              // Zoom Value
    	utilities.setDropDownValueByIndex($("#commomFormatType"), 0);             // Format Setting
    	utilities.setDropDownValueByIndex($("#boxplotClassType"), 0);             // BoxPlotSparkline Class
	    utilities.setDropDownValue("boxplotSparklineStyleType", 0);               // BoxPlotSparkline Style
    	utilities.setDropDownValue("dataOrientationType", 0);                     // CompatibleSparkline DataOrientation
    	utilities.setDropDownValue("paretoLabelList", 0);                         // ParetoSparkline Label
    	utilities.setDropDownValue("spreadSparklineStyleType", 4);                // SpreadSparkline Style
    	utilities.setDropDownValue("stackedSparklineTextOrientation", 0);         // StackedSparkline TextOrientation
    	utilities.setDropDownValueByIndex($("#spreadTheme"), 1);                  // Spread Theme
    	utilities.setDropDownValue("resizeZeroIndicator", 1);                     // ResizeZeroIndicator
    	utilities.setDropDownValueByIndex($("#copyPasteHeaderOptions"), 3);       // CopyPasteHeaderOptins
    	utilities.setDropDownValueByIndex($("#cellLabelVisibility"), 0);          // CellLabelVisibility
    	utilities.setDropDownValueByIndex($("#cellLabelAlignment"), 0);           // CellLabelAlignment
	}

	processConditionalFormatDetailSetting(name:string, noAction?:boolean) {
		switch (name) {
        	case "highlight-cells-rules":
            	$("#formatSetting").show();
            	utilities.processConditionalFormatSetting("normal", "highlightCellsRulesList", 0);
            break;

        case "top-bottom-rules":
            	$("#formatSetting").show();
            	utilities.processConditionalFormatSetting("normal", "topBottomRulesList", 4);
            break;

        case "color-scales":
            	$("#formatSetting").hide();
            	utilities.processConditionalFormatSetting("normal", "colorScaleList", 8);
            break;

        case "data-bars":
            	utilities.processConditionalFormatSetting("databar");
            break;

        case "icon-sets":
            	utilities.processConditionalFormatSetting("iconset");
            	utilities.updateIconCriteriaItems(0);
            break;

        case "remove-conditional-formats":
            $("#conditionalFormatSettingContainer div.details").hide();
            if (!noAction) {
                this.removeConditionFormats();
            }
            break;

        default:
            console.log("processConditionalFormatSetting not add for ", name);
            break;
    	}
	}

	processCellTypeSetting(name:string, noAction?:boolean) {
		$("#cellTypeSettingContainer").show();

		switch (name) {
			case "button-celltype":
				$("#celltype-button").show();
				$("#celltype-checkbox").hide();
				$("#celltype-combobox").hide();
				$("#celltype-hyperlink").hide();
				break;

			case "checkbox-celltype":
				$("#celltype-button").hide();
				$("#celltype-checkbox").show();
				$("#celltype-combobox").hide();
				$("#celltype-hyperlink").hide();
				break;

			case "combobox-celltype":
				$("#celltype-button").hide();
				$("#celltype-checkbox").hide();
				$("#celltype-combobox").show();
				$("#celltype-hyperlink").hide();
				break;

			case "hyperlink-celltype":
				$("#celltype-button").hide();
				$("#celltype-checkbox").hide();
				$("#celltype-combobox").hide();
				$("#celltype-hyperlink").show();
				break;

			case "clear-celltype":
				if (!noAction) {
					this.clearCellType();
				}
				$("#cellTypeSettingContainer").hide();
				return;

			default:
				console.log("processCellTypeSetting not process with ", name);
				return;
		}
	}

	clearCellType() {
		var sheet = this.spread.getActiveSheet();
		var sels = sheet.getSelections();
		var rowCount = sheet.getRowCount(),
			columnCount = sheet.getColumnCount();
		sheet.suspendPaint();
		for (var i = 0; i < sels.length; i++) {
			var sel = this.getActualCellRange(sheet, sels[i], rowCount, columnCount);
			sheet.clear(sel.row, sel.col, sel.rowCount, sel.colCount, GC.Spread.Sheets.SheetArea.viewport, GC.Spread.Sheets.StorageType.style);
		}
		sheet.resumePaint();
	}

	removeConditionFormats() {
    	let sheet = this.spread.getActiveSheet();
    	let cfs = sheet.conditionalFormats;
    	let row = sheet.getActiveRowIndex(), col = sheet.getActiveColumnIndex();
    	let rules:GC.Spread.Sheets.ConditionalFormatting.ConditionRuleBase[] = cfs.getRules(row, col);

    	sheet.suspendPaint();

    	$.each(rules, function (i:number, v:GC.Spread.Sheets.ConditionalFormatting.ConditionRuleBase) {
        	cfs.removeRule(v);
    	});

    	sheet.resumePaint();
	}

	addEventHandlers() {
		let component = this;

		$("div.insp-group-title>span").click(this.toggleState);
		$("div.insp-checkbox").click(function() {
			component.checkedChanged(this);
		});
		$("div.insp-number>input.editor").blur(function() {
			component.updateNumberProperty(this)
		});
		$("div.insp-dropdown-list .dropdown").click(function() {
			component.showDropdown(this);
		});
		$("div.insp-menu .menu-item").click(function() {
			component.itemSelected(this);
		});
		$("div.insp-color-picker .picker").click(function() {
			component.showColorPicker(this);
		});
		$("li.color-cell").click(function() {
			component.colorSelected(this);
		});
		$(".insp-button-group span.btn").click(function() {
			component.buttonClicked(this);
		});
		$(".insp-radio-button-group span.btn").click(function() {
			component.buttonClicked(this);
		});
		$(".insp-buttons .btn").click(function() {
			component.divButtonClicked(this);
		});
		$(".insp-text input.editor").blur(function() {
			component.updateStringProperty(this);
		});
	}

	updateNumberProperty(element:HTMLElement) {
		var $element = $(element),
			$parent = $element.parent(),
			name = $parent.data("name"),
			value = parseInt($element.val(), 10);

		if (isNaN(value)) {
			return;
		}

		var sheet = this.spread.getActiveSheet();

		this.spread.suspendPaint();
		switch (name) {
			case "rowCount":
				sheet.setRowCount(value);
				break;

			case "columnCount":
				sheet.setColumnCount(value);
				break;

			case "frozenRowCount":
				sheet.frozenRowCount(value);
				break;

			case "frozenColumnCount":
				sheet.frozenColumnCount(value);
				break;

			case "trailingFrozenRowCount":
				sheet.frozenTrailingRowCount(value);
				break;

			case "trailingFrozenColumnCount":
				sheet.frozenTrailingColumnCount(value);
				break;

			default:
				console.log("updateNumberProperty need add for", name);
				break;
		}
		this.spread.resumePaint();
	}

	divButtonClicked(element:HTMLElement) {
		var sheet = this.spread.getActiveSheet(),
			id = element.id;

		this.spread.suspendPaint();
		switch (id) {
			case "mergeCells":
				this.mergeCells(sheet);
				this.updateMergeButtonsState();
				break;

			case "unmergeCells":
				this.unmergeCells(sheet);
				this.updateMergeButtonsState();
				break;

			default:
				console.log("TODO add code for ", id);
				break;
		}
		this.spread.resumePaint();
	}

	mergeCells(sheet:any) {
		var sels = sheet.getSelections();
		var rowCount = sheet.getRowCount(),
			columnCount = sheet.getColumnCount();

		for (var n = 0; n < sels.length; n++) {
			var sel = this.getActualCellRange(sheet, sels[n], rowCount, columnCount);
			sheet.addSpan(sel.row, sel.col, sel.rowCount, sel.colCount);
		}
	}

	unmergeCells(sheet:any) {
		function removeSpan(range:any) {
			sheet.removeSpan(range.row, range.col);
		}

		var sels = sheet.getSelections();
		var rowCount = sheet.getRowCount(),
			columnCount = sheet.getColumnCount();

		for (var n = 0; n < sels.length; n++) {
			var sel = this.getActualCellRange(sheet, sels[n], rowCount, columnCount);
			sheet.getSpans(sel).forEach(removeSpan);
		}
	}

	checkedChanged(element:HTMLElement) {
		var $element = $(element),
			name = $element.data("name");

		if ($element.hasClass("disabled")) {
			return;
		}

		// radio buttons need special process
		switch (name) {
			case "referenceStyle":
			case "slicerMoveAndSize":
			case "pictureMoveAndSize":
				this.processRadioButtonClicked(name, $(event.target), $element);
				return;
		}


		var $target = $("div.button", $element),
			value = !$target.hasClass("checked");

		var sheet = this.spread.getActiveSheet();

		$target.toggleClass("checked");

		this.spread.suspendPaint();

		var options = this.spread.options;

		switch (name) {
			case "wrapText":
				this.setWordWrap(sheet);
				break;
			default:
				console.log("not added code for", name);
				break;

		}
		this.spread.resumePaint();
	}

	setWordWrap(sheet:any) {
		var sels = sheet.getSelections();
		var rowCount = sheet.getRowCount(),
			columnCount = sheet.getColumnCount();

		sheet.suspendPaint();
		for (var n = 0; n < sels.length; n++) {
			var sel = this.getActualCellRange(sheet, sels[n], rowCount, columnCount),
				wordWrap = !sheet.getCell(sel.row, sel.col).wordWrap(),
				startRow = sel.row,
				endRow = sel.row + sel.rowCount - 1;

			sheet.getRange(startRow, sel.col, sel.rowCount, sel.colCount).wordWrap(wordWrap);

			for (var row = startRow; row <= endRow; row++) {
				sheet.autoFitRow(row);
			}
		}
		sheet.resumePaint();
	}

	colorSelected(element:HTMLElement) {
		var themeColor = $(element).data("name");
		var value = $(element).css("background-color");

		var name = $(this._colorpicker).data("name");
		var sheet = this.spread.getActiveSheet();

		$("div.color-view", $(this._colorpicker).data("dropdown")).css("background-color", value);

		// No Fills need special process
		if ($(this).hasClass("auto-color-cell")) {
			if (name === "backColor") {
				value = undefined;
			}
		}

		var options = this.spread.options;

		this.spread.suspendPaint();
		switch (name) {
			case "spreadBackcolor":
				options.backColor = value;
				break;

			case "grayAreaBackcolor":
				options.grayAreaBackColor = value;
				break;

			case "cutCopyIndicatorBorderColor":
				options.cutCopyIndicatorBorderColor = value;
				break;

			case "sheetTabColor":
				sheet.options.sheetTabColor = value;
				break;

			case "frozenLineColor":
				sheet.options.frozenlineColor = value;
				break;

			case "gridlineColor":
				sheet.options.gridline.color = value;
				break;

			case "foreColor":
			case "backColor":
				this.setColor(sheet, name, themeColor || value);
				break;

			case "labelForeColor":
				this.setLabelOptions(sheet, value, "foreColor");
				break;

			case "selectionBorderColor":
				sheet.options.selectionBorderColor = value;
				break;

			case "selectionBackColor":
				// change to rgba (alpha: 0.2) to make cell content visible
				value = utilities.getRGBAColor(value, 0.2);
				sheet.options.selectionBackColor = value;
				$("div.color-view", $(this._colorpicker).data("dropdown")).css("background-color", value);
				break;

			case "commentBorderColor":
				this._activeComment && this._activeComment.borderColor(value);
				break;

			default:
				console.log("TODO colorSelected", name);
				break;
		}
		this.spread.resumePaint();
	}

	setColor(sheet:any, method:any, value:any) {
		var sels = sheet.getSelections();
		var rowCount = sheet.getRowCount(),
			columnCount = sheet.getColumnCount();

		sheet.suspendPaint();
		for (var n = 0; n < sels.length; n++) {
			var sel = this.getActualCellRange(sheet, sels[n], rowCount, columnCount);
			sheet.getRange(sel.row, sel.col, sel.rowCount, sel.colCount)[method](value);
		}
		sheet.resumePaint();
	}

	updateStringProperty(element:HTMLElement) {
		var $element = $(element),
			$parent = $element.parent(),
			name = $parent.data("name"),
			value = $element.val();

		var sheet = this.spread.getActiveSheet();

		switch (name) {
			case "sheetName":
				if (value && value !== sheet.name()) {
					try {
						sheet.name(value);
					} catch (ex) {
						alert(utilities.getResource("messages.duplicatedSheetName"));
						$element.val(sheet.name());
					}
				}
				break;

			case "tableName":
				if (value && this._activeTable && value !== this._activeTable.name()) {
					if (!sheet.tables.findByName(value)) {
						this._activeTable.name(value);
					} else {
						alert(utilities.getResource("messages.duplicatedTableName"));
						$element.val(this._activeTable.name());
					}
				}
				break;

			case "commentPadding":
				//setCommentPadding(value);
				break;

			case "customFormat":
				this.setFormatter(value);
				break;

			case "slicerName":
				//setSlicerSetting("name", value);
				break;

			case "slicerCaptionName":
				//setSlicerSetting("captionName", value);
				break;

			case "watermark":
				//setWatermark(sheet, value);
				break;

			case "cellPadding":
				this.setCellPadding(sheet, value);
				break;

			case "labelMargin":
				this.setLabelOptions(sheet, value, "margin");
				break;
			default:
				console.log("updateStringProperty w/o process of ", name);
				break;
		}
	}

	setCellPadding(sheet:any, value:any) {
		var selections = sheet.getSelections(),
			rowCount = sheet.getRowCount(),
			columnCount = sheet.getColumnCount();
		sheet.suspendPaint();
		for (var n = 0; n < selections.length; n++) {
			var sel = this.getActualCellRange(sheet, selections[n], rowCount, columnCount);
			for (var r = sel.row; r < sel.row + sel.rowCount; r++) {
				for (var c = sel.col; c < sel.col + sel.colCount; c++) {
					var style = sheet.getStyle(r, c);
					if (!style) {
						style = new GC.Spread.Sheets.Style();
					}
					style.cellPadding = value;
					sheet.setStyle(r, c, style);
				}
			}
		}
		sheet.resumePaint();
	}

	setLabelOptions(sheet:any, value:any, option:any) {
		var selections = sheet.getSelections(),
			rowCount = sheet.getRowCount(),
			columnCount = sheet.getColumnCount();
		sheet.suspendPaint();
		for (var n = 0; n < selections.length; n++) {
			var sel = this.getActualCellRange(sheet, selections[n], rowCount, columnCount);
			for (var r = sel.row; r < sel.row + sel.rowCount; r++) {
				for (var c = sel.col; c < sel.col + sel.colCount; c++) {
					var style = sheet.getStyle(r, c);
					if (!style) {
						style = new GC.Spread.Sheets.Style();
					}
					if (!style.labelOptions) {
						style.labelOptions = {};
					}
					if (option === "foreColor") {
						style.labelOptions.foreColor = value;
					} else if (option === "margin") {
						style.labelOptions.margin = value;
					} else if (option === "visibility") {
						style.labelOptions.visibility = GC.Spread.Sheets.LabelVisibility[value];
					} else if (option === "alignment") {
						style.labelOptions.alignment = GC.Spread.Sheets.LabelAlignment[value];
					}
					sheet.setStyle(r, c, style);
				}
			}
		}
		sheet.resumePaint();
	}

	buttonClicked(element:HTMLElement) {
		let $element = $(element);
		let	name = $element.data("name");
		let	container;

		var sheet = this.spread.getActiveSheet();

		// get group
		if ((container = $element.parents(".insp-radio-button-group")).length > 0) {
			name = container.data("name");
			$element.siblings().removeClass("active");
			$element.addClass("active");
			switch (name) {
				case "vAlign":
				case "hAlign":
					this.setAlignment(sheet, name, $element.data("name"));
					break;
			}
		} else if ($element.parents(".insp-button-group").length > 0) {
			if (!$element.hasClass("no-toggle")) {
				$element.toggleClass("active");
			}

			switch (name) {
				case "bold":
					this.setStyleFont(sheet, "font-weight", false, ["700", "bold"], "normal");
					break;
				case "labelBold":
					this.setStyleFont(sheet, "font-weight", true, ["700", "bold"], "normal");
					break;
				case "italic":
					this.setStyleFont(sheet, "font-style", false, ["italic"], "normal");
					break;
				case "labelItalic":
					this.setStyleFont(sheet, "font-style", true, ["italic"], "normal");
					break;
				case "underline":
					this.setTextDecoration(sheet, GC.Spread.Sheets.TextDecorationType.underline);
					break;
				case "strikethrough":
					this.setTextDecoration(sheet, GC.Spread.Sheets.TextDecorationType.lineThrough);
					break;
				case "overline":
					this.setTextDecoration(sheet, GC.Spread.Sheets.TextDecorationType.overline);
					break;

				case "increaseIndent":
					this.setTextIndent(sheet, 1);
					break;

				case "decreaseIndent":
					this.setTextIndent(sheet, -1);
					break;

				case "percentStyle":
					this.setFormatter(uiResource.cellTab.format.percentValue);
					break;

				case "commaStyle":
					this.setFormatter(uiResource.cellTab.format.commaValue);
					break;

				case "increaseDecimal":
					this.increaseDecimal();
					break;

				case "decreaseDecimal":
					this.decreaseDecimal();
					break;

				default:
					console.log("buttonClicked w/o process code for ", name);
					break;
			}
		}
	}

	increaseDecimal() {
		var sheet = this.spread.getActiveSheet();
		let component = this;
		this.execInSelections(sheet, "formatter", function (sheet:any, row:any, column:any) {
			var style = sheet.getStyle(row, column);
			if (!style) {
				style = new GC.Spread.Sheets.Style();
			}
			var activeCell = sheet.getCell(sheet.getActiveRowIndex(), sheet.getActiveColumnIndex());
			var activeCellValue = activeCell.value();
			var activeCellFormatter = activeCell.formatter();
			var activeCellText = activeCell.text();

			if (activeCellValue) {
				var formatString = null;
				var zero = "0";
				var numberSign = "#";
				var decimalPoint = ".";
				var zeroPointZero = "0" + decimalPoint + "0";

				var scientificNotationCheckingFormatter = component.getScientificNotationCheckingFormattter(activeCellFormatter);
				if (!activeCellFormatter || ((activeCellFormatter == "General" || (scientificNotationCheckingFormatter &&
					(scientificNotationCheckingFormatter.indexOf("E") >= 0 || scientificNotationCheckingFormatter.indexOf('e') >= 0))))) {
					if (!isNaN(activeCellValue)) {
						var result = activeCellText.split('.');
						if (result.length == 1) {
							if (result[0].indexOf('E') >= 0 || result[0].indexOf('e') >= 0)
								formatString = zeroPointZero + "E+00";
							else
								formatString = zeroPointZero;
						}
						else if (result.length == 2) {
							result[0] = "0";
							var isScience = false;
							var sb = "";
							for (var i = 0; i < result[1].length + 1; i++) {
								sb = sb + '0';
								if (i < result[1].length && (result[1].charAt(i) == 'e' || result[1].charAt(i) == 'E')) {
									isScience = true;
									break;
								}
							}
							if (isScience)
								sb = sb + "E+00";
							if (sb) {
								result[1] = sb.toString();
								formatString = result[0] + decimalPoint + result[1];
							}
						}
					}
				}
				else {
					formatString = activeCellFormatter;
					if (formatString) {
						var formatters = formatString.split(';');
						for (var i = 0; i < formatters.length && i < 2; i++) {
							if (formatters[i] && formatters[i].indexOf("/") < 0 && formatters[i].indexOf(":") < 0 && formatters[i].indexOf("?") < 0) {
								var indexOfDecimalPoint = formatters[i].lastIndexOf(decimalPoint);
								if (indexOfDecimalPoint != -1) {
									formatters[i] = formatters[i].slice(0, indexOfDecimalPoint + 1) + zero + formatters[i].slice(indexOfDecimalPoint + 1);
								}
								else {
									var indexOfZero = formatters[i].lastIndexOf(zero);
									var indexOfNumberSign = formatters[i].lastIndexOf(numberSign);
									var insertIndex = indexOfZero > indexOfNumberSign ? indexOfZero : indexOfNumberSign;
									if (insertIndex >= 0)
										formatters[i] = formatters[i].slice(0, insertIndex + 1) + decimalPoint + zero + formatters[i].slice(insertIndex + 1);
								}
							}
						}
						formatString = formatters.join(";");
					}
				}
				style.formatter = formatString;
				sheet.setStyle(row, column, style);
			}
		});
	}

	decreaseDecimal() {
    	var sheet = this.spread.getActiveSheet();
		this.execInSelections(sheet, "formatter", function (sheet:any, row:any, column:any) {
			var style = sheet.getStyle(row, column);
			if (!style) {
				style = new GC.Spread.Sheets.Style();
			}
			var activeCell = sheet.getCell(sheet.getActiveRowIndex(), sheet.getActiveColumnIndex());
			var activeCellValue = activeCell.value();
			var activeCellFormatter = activeCell.formatter();
			var activeCellText = activeCell.text();
			var decimalPoint = ".";
			if (activeCellValue) {
				var formatString = null;
				if (!activeCellFormatter || activeCellFormatter == "General") {
					if (!isNaN(activeCellValue)) {
						var result = activeCellText.split('.');
						if (result.length == 2) {
							result[0] = "0";
							var isScience = false;
							var sb = "";
							for (var i = 0; i < result[1].length - 1; i++) {
								if ((i + 1 < result[1].length) && (result[1].charAt(i + 1) == 'e' || result[1].charAt(i + 1) == 'E')) {
									isScience = true;
									break;
								}
								sb = sb + ('0');
							}

							if (isScience)
								sb = sb + ("E+00");

							if (sb !== null) {
								result[1] = sb.toString();

								formatString = result[0] + (result[1] !== "" ? decimalPoint + result[1] : "");
							}
						}
					}
				}
				else {
					formatString = activeCellFormatter;
					if (formatString) {
						var formatters = formatString.split(';');
						for (var i = 0; i < formatters.length && i < 2; i++) {
							if (formatters[i] && formatters[i].indexOf("/") < 0 && formatters[i].indexOf(":") < 0 && formatters[i].indexOf("?") < 0) {
								var indexOfDecimalPoint = formatters[i].lastIndexOf(decimalPoint);
								if (indexOfDecimalPoint != -1 && indexOfDecimalPoint + 1 < formatters[i].length) {
									formatters[i] = formatters[i].slice(0, indexOfDecimalPoint + 1) + formatters[i].slice(indexOfDecimalPoint + 2);
									var tempString = indexOfDecimalPoint + 1 < formatters[i].length ? formatters[i].substr(indexOfDecimalPoint + 1, 1) : "";
									if (tempString === "" || tempString !== "0")
										formatters[i] = formatters[i].slice(0, indexOfDecimalPoint) + formatters[i].slice(indexOfDecimalPoint + 1);
								}
								else {
									//do nothing.
								}
							}
						}
						formatString = formatters.join(";");
					}
				}
				style.formatter = formatString;
				sheet.setStyle(row, column, style);
			}
		});
	}

	getScientificNotationCheckingFormattter(formatter:any) {
		if (!formatter) {
			return formatter;
		}
		var i;
		var signalQuoteSubStrings = this.getSubStrings(formatter, '\'', '\'');
		for (i = 0; i < signalQuoteSubStrings.length; i++) {
			formatter = formatter.replace(signalQuoteSubStrings[i], '');
		}
		var doubleQuoteSubStrings = this.getSubStrings(formatter, '\"', '\"');
		for (i = 0; i < doubleQuoteSubStrings.length; i++) {
			formatter = formatter.replace(doubleQuoteSubStrings[i], '');
		}
		var colorStrings = this.getSubStrings(formatter, '[', ']');
		for (i = 0; i < colorStrings.length; i++) {
			formatter = formatter.replace(colorStrings[i], '');
		}
		return formatter;
	}

	getSubStrings(source:any, beginChar:any, endChar:any) {
		if (!source) {
			return [];
		}
		var subStrings = [], tempSubString = '', inSubString = false;
		for (var index = 0; index < source.length; index++) {
			if (!inSubString && source[index] === beginChar) {
				inSubString = true;
				tempSubString = source[index];
            continue;
        }
        if (inSubString) {
            tempSubString += source[index];
            if (source[index] === endChar) {
                subStrings.push(tempSubString);
                tempSubString = "";
                inSubString = false;
            }
        }
    }
    return subStrings;
}

	setFormatter(value:any) {
		var sheet = this.spread.getActiveSheet();
		this.execInSelections(sheet, "formatter", function (sheet:any, row:any, column:any) {
			var style = sheet.getStyle(row, column);
			if (!style) {
				style = new GC.Spread.Sheets.Style();
			}
			style.formatter = value;
			sheet.setStyle(row, column, style);
		});
	}

	execInSelections(sheet:any, styleProperty:any, func:any) {
		var selections = sheet.getSelections();
		for (var k = 0; k < selections.length; k++) {
			var selection = selections[k];
			var col = selection.col, row = selection.row,
				rowCount = selection.rowCount, colCount = selection.colCount;
			if ((col === -1 || row === -1) && styleProperty) {
				var style, r, c;
				// whole sheet was selected, need set row / column' style one by one
				if (col === -1 && row === -1) {
					for (r = 0; r < rowCount; r++) {
						if ((style = sheet.getStyle(r, -1)) && style[styleProperty] !== undefined) {
							func(sheet, r, -1);
						}
					}
					for (c = 0; c < colCount; c++) {
						if ((style = sheet.getStyle(-1, c)) && style[styleProperty] !== undefined) {
							func(sheet, -1, c);
						}
					}
				}
				// Get actual range for whole rows / columns / sheet selection
				if (col === -1) {
					col = 0;
				}
				if (row === -1) {
					row = 0;
				}
				// set to each cell with style that in the adjusted selection range
				for (var i = 0; i < rowCount; i++) {
					r = row + i;
					for (var j = 0; j < colCount; j++) {
						c = col + j;
						if ((style = sheet.getStyle(r, c)) && style[styleProperty] !== undefined) {
							func(sheet, r, c);
						}
					}
				}
			}
			if (selection.col == -1 && selection.row == -1) {
				func(sheet, -1, -1);
			}
			else if (selection.row == -1) {
				for (var i = 0; i < selection.colCount; i++) {
					func(sheet, -1, selection.col + i);
				}
			}
			else if (selection.col == -1) {
				for (var i = 0; i < selection.rowCount; i++) {
					func(sheet, selection.row + i, -1);
				}
			}
			else {
				for (var i = 0; i < selection.rowCount; i++) {
					for (var j = 0; j < selection.colCount; j++) {
						func(sheet, selection.row + i, selection.col + j);
					}
				}
			}
		}
	}

	setTextIndent(sheet:any, step:any) {
		var sels = sheet.getSelections();
		var rowCount = sheet.getRowCount(),
			columnCount = sheet.getColumnCount();

		sheet.suspendPaint();
		for (var n = 0; n < sels.length; n++) {
			var sel = this.getActualCellRange(sheet, sels[n], rowCount, columnCount),
				indent = sheet.getCell(sel.row, sel.col).textIndent();

			if (isNaN(indent)) {
				indent = 0;
			}

			var value = indent + step;
			if (value < 0) {
				value = 0;
			}
			sheet.getRange(sel.row, sel.col, sel.rowCount, sel.colCount).textIndent(value);
		}
		sheet.resumePaint();
	}

	setAlignment(sheet:any, type:any, value:any) {
		var sels = sheet.getSelections(),
			rowCount = sheet.getRowCount(),
			columnCount = sheet.getColumnCount(),
			align;

		value = value.toLowerCase();

		if (value === "middle") {
			value = "center";
		}

		if (type === "hAlign") {
			align = GC.Spread.Sheets.HorizontalAlign[value];
		} else {
			align = GC.Spread.Sheets.VerticalAlign[value];
		}

		sheet.suspendPaint();
		for (var n = 0; n < sels.length; n++) {
			var sel = this.getActualCellRange(sheet, sels[n], rowCount, columnCount);
			sheet.getRange(sel.row, sel.col, sel.rowCount, sel.colCount)[type](align);
		}
		sheet.resumePaint();
	}

	setTextDecoration(sheet:any, flag:any) {
		var sels = sheet.getSelections();
		var rowCount = sheet.getRowCount(),
			columnCount = sheet.getColumnCount();

		sheet.suspendPaint();
		for (var n = 0; n < sels.length; n++) {
			var sel = this.getActualCellRange(sheet, sels[n], rowCount, columnCount),
				textDecoration = sheet.getCell(sel.row, sel.col).textDecoration();
			if ((textDecoration & flag) === flag) {
				textDecoration = textDecoration - flag;
			} else {
				textDecoration = textDecoration | flag;
			}
			sheet.getRange(sel.row, sel.col, sel.rowCount, sel.colCount).textDecoration(textDecoration);
		}
		sheet.resumePaint();
	}

	showDropdown(element:HTMLElement) {
		if (!this._needShow) {
			this._needShow = true;
			return;
		}

		var DROPDOWN_OFFSET = 10;
		var $element = $(element),
			$container = $element.parent(),
			name = $container.data("name"),
			targetId = $container.data("list-ref"),
			$target = $("#" + targetId);

		if ($target && !$target.hasClass("show")) {
			$target.data("dropdown", element);
			this._dropdownitem = $target[0];

			var $dropdown = $element,
				offset = $dropdown.offset();

			var height = $element.outerHeight(),
				targetHeight = $target.outerHeight(),
				width = $element.outerWidth(),
				targetWidth = $target.outerWidth(),
				top = offset.top + height;

			// adjust drop down' width to same
			if (targetWidth < width) {
				$target.width(width);
			}

			var $inspContainer = $(".insp-container"),
				maxTop = $inspContainer.height() + $inspContainer.offset().top;

			// adjust top when out of bottom range
			if (top + targetHeight + DROPDOWN_OFFSET > maxTop) {
				top = offset.top - targetHeight;
			}

			$target.css({
				top: top,
				left: offset.left - $target.width() + $dropdown.width() + 16
			});

			// select corresponding item
			if (name === "borderLine") {
				var text = $("#border-line-type").attr("class");
				$("div.image", $target).removeClass("fa-check");
				$("div.text", $target).filter(function () {
					return $(this).find("div").attr("class") === text;
				}).siblings("div.image").addClass("fa fa-check");
				$("div.image.nocheck", $target).removeClass("fa-check");
			}
			else {
				var text = $("span.display", $dropdown).text();
				$("div.image", $target).removeClass("fa-check");
				$("div.text", $target).filter(function () {
					return $(this).text() === text;
				}).siblings("div.image").addClass("fa fa-check");
				// remove check for special items mark with nocheck class
				$("div.image.nocheck", $target).removeClass("fa-check");
			}

			$target.addClass("show");

			this.processEventListenerHandleClosePopup(true);
		}
	}

	itemSelected(element:HTMLElement) {
		// get related dropdown item
		var dropdown = $(this._dropdownitem).data("dropdown");

		this.hideDropdown();

		var sheet = this.spread.getActiveSheet();

		var name = $(dropdown.parentElement).data("name"),
			$text = $("div.text", element),
			dataValue = $text.data("value"),    // data-value includes both number value and string value, should pay attention when use it
			numberValue = +dataValue,
			text = $text.text(),
			value = text,
			nameValue = dataValue || text;

		var options = this.spread.options;

		switch (name) {
			case "scrollTip":
				options.showScrollTip = numberValue;
				break;

			case "resizeTip":
				options.showResizeTip = numberValue;
				break;

			case "fontFamily":
				this.setStyleFont(sheet, "font-family", false, [value], value);
				break;

			case "labelFontFamily":
				this.setStyleFont(sheet, "font-family", true, [value], value);
				break;

			case "fontSize":
				value += "pt";
				this.setStyleFont(sheet, "font-size", false, [value], value);
				break;

			case "labelFontSize":
				value += "pt";
				this.setStyleFont(sheet, "font-size", true, [value], value);
				break;

			case "cellLabelVisibility":
				this.setLabelOptions(sheet, nameValue, "visibility");
				break;

			case "cellLabelAlignment":
				this.setLabelOptions(sheet, nameValue, "alignment");
				break;

			case "selectionPolicy":
				sheet.selectionPolicy(numberValue);
				break;

			case "selectionUnit":
				sheet.selectionUnit(numberValue);
				break;

			case "sheetName":
				var selectedSheet = this.spread.sheets[numberValue];
				utilities.setCheckValue("sheetVisible", selectedSheet.visible(), {
					sheetIndex: numberValue,
					sheetName: selectedSheet.name()
				});
				break;

			case "conditionalFormat":
				this.processConditionalFormatDetailSetting(nameValue);
				break;

			case "ruleType":
				utilities.updateEnumTypeOfCF(numberValue);
				break;

			case "iconSetType":
				utilities.updateIconCriteriaItems(numberValue);
				break;

			case "minType":
				this.processMinItems(numberValue, "minValue");
				break;

			case "midType":
				this.processMidItems(numberValue, "midValue");
				break;

			case "maxType":
				this.processMaxItems(numberValue, "maxValue");
				break;

			case "cellTypes":
				this.processCellTypeSetting(nameValue);
				break;

			case "commomFormat":
				this.processFormatSetting(nameValue, value);
				break;

			case "borderLine":
				utilities.processBorderLineSetting(nameValue);
				break;

			default:
				console.log("TODO add itemSelected for ", name, value);
				break;
		}

		this.setDropDownText(dropdown, text);
	}

	processMinItems(type:any, name:any) {
		var value = "";
		switch (type) {
			case 0: // Number
			case 3: // Percent
				value = "0";
				break;
			case 4: // Percentile
				value = "10";
				break;
			default:
				value = "";
				break;
		}
		utilities.setTextValue(name, value);
	}

	processMidItems(type:number, name:string) {
		var value = "";
		switch (type) {
			case 0: // Number
				value = "0";
				break;
			case 3: // Percent
			case 4: // Percentile
				value = "50";
				break;
			default:
				value = "";
				break;
		}
		utilities.setTextValue(name, value);
	}

	processMaxItems(type:number, name:string) {
		var value = "";
		switch (type) {
			case 0: // Number
				value = "0";
				break;
			case 3: // Percent
				value = "100";
				break;
			case 4: // Percentile
				value = "90";
				break;
			default:
				value = "";
				break;
		}
		utilities.setTextValue(name, value);
	}

	processFormatSetting(name:any, title:any) {
		switch (name) {
			case "nullValue":
				name = null;
			case "0.00":
			case "$#,##0.00":
			case "$ #,##0.00;$ (#,##0.00);$ '-'??;@":
			case "m/d/yyyy":
			case "dddd, mmmm dd, yyyy":
			case "h:mm:ss AM/PM":
			case "0%":
			case "# ?/?":
			case "0.00E+00":
			case "@":
				this.setFormatter(name);
				break;

			default:
				console.log("processFormatSetting not process with ", name, title);
				break;
		}
	}

	setDropDownText(container:any, value:any) {
		var refList = "#" + $(container).data("list-ref"),
			$items = $(".menu-item div.text", refList),
			$item = $items.filter(function () {
				return $(this).data("value") === value;
			});

		var text = $item.text() || value;

		$("span.display", container).text(text);
	}

	hideDropdown() {
		if (this._dropdownitem) {
			$(this._dropdownitem).removeClass("show");
			this._dropdownitem = null;
		}

		this.processEventListenerHandleClosePopup(false);
	}

	toggleState():void {
		let $element = $(this),
        $parent = $element.parent(),
        $content = $parent.siblings(".insp-group-content"),
        $target = $parent.find("span.group-state"),
        collapsed = $target.hasClass("fa-caret-right");

    	if (collapsed) {
        	$target.removeClass("fa-caret-right").addClass("fa-caret-down");
        	$content.slideToggle("fast");
    	} else {
        	$target.addClass("fa-caret-right").removeClass("fa-caret-down");
        	$content.slideToggle("fast");
    	}
	}

	unParseFormula(expr:any, row:any, col:any):string|null {
		if (!expr) {
			return "";
		}
		var sheet = this.spread.getActiveSheet();
		if (!sheet) {
			return null;
		}
		var calcService = (<any>sheet).getCalcService();
		return calcService.unparse(null, expr, row, col);
	}

	parseColorExpression(colorExpression:any, row:any, col:any) {
		if (!colorExpression) {
			return null;
		}
		var sheet = this.spread.getActiveSheet();
		if (colorExpression.type === ExpressionType.string) {
			return colorExpression.value;
		}
		else if (colorExpression.type === ExpressionType.missingArgument) {
			return null;
		}
		else {
			var formula = null;
			try {
				formula = this.unParseFormula(colorExpression, row, col);
			}
			catch (ex) {
			}
			return SheetsCalc.evaluateFormula(sheet, formula, row, col);
		}
	}

	showColorPicker(element:HTMLElement) {
		if (!this._needShow) {
			this._needShow = true;
			return;
		}

		var MIN_TOP = 30, MIN_BOTTOM = 4;
		var $element = $(element),
			$container = $element.parent(),
			name = $container.data("name"),
			$target = $("#colorpicker");

		if ($target && !$target.hasClass("colorpicker-visible")) {
			$target.data("dropdown", element);
			// save related name for later use
			$target.data("name", name);

			var $nofill = $target.find("div.nofill-color");
			if ($container.hasClass("show-nofill-color")) {
				$nofill.show();
			} else {
				$nofill.hide();
			}

			this._colorpicker = $target[0];

			var $dropdown = $element,
				offset = $dropdown.offset();

			var height = $target.height(),
				top = offset.top - (height - $element.height()) / 2 + 3,   // 3 = padding (4) - border-width(1)
				yOffset = 0;

			if (top < MIN_TOP) {
				yOffset = MIN_TOP - top;
				top = MIN_TOP;
			} else {
				var $inspContainer = $(".insp-container"),
					maxTop = $inspContainer.height() + $inspContainer.offset().top;

				// adjust top when out of bottom range
				if (top + height > maxTop - MIN_BOTTOM) {
					var newTop = maxTop - MIN_BOTTOM - height;
					yOffset = newTop - top;
					top = newTop;
				}
			}

			$target.css({
				top: top,
				left: offset.left - $target.width() - 20
			});

			// v-center the pointer
			var $pointer = $target.find(".cp-pointer");
			$pointer.css({top: (height - 24) / 2 - yOffset});   // 24 = pointer height

			$target.addClass("colorpicker-visible");

			this.processEventListenerHandleClosePopup(true);
		}
	}
	
	processRadioButtonClicked(key:any, $item:any, $group:any) {
		var name = $item.data("name");

    	// only need process when click on radio button or relate label like text
    	if ($item.hasClass("radiobutton") || $item.hasClass("text")) {
        	$group.find("div.radiobutton").removeClass("checked");
        	$group.find("div.radiobutton[data-name='" + name + "']").addClass("checked");

        	switch (key) {
			}
    	}
	}
	
	setStyleFont(sheet:GC.Spread.Sheets.Worksheet, prop:string, isLabelStyle:boolean, optionValue1:string[], optionValue2:string) {
		var styleEle = document.getElementById("setfontstyle"),
        	selections = sheet.getSelections(),
        	rowCount = sheet.getRowCount(),
        	columnCount = sheet.getColumnCount(),
        	defaultStyle = sheet.getDefaultStyle();

    	function updateStyleFont(style:any) {
        	if (!style.font) {
            	style.font = defaultStyle.font || "11pt Calibri";
        	}
        	styleEle.style.font = style.font;
        	var styleFont = $(styleEle).css(prop);

        	if (styleFont === optionValue1[0] || styleFont === optionValue1[1]) {
            	if (defaultStyle.font) {
                	styleEle.style.font = defaultStyle.font;
                	var defaultFontProp = $(styleEle).css(prop);
                	styleEle.style.font = style.font;
                	$(styleEle).css(prop, defaultFontProp);
            	}
            	else {
                	$(styleEle).css(prop, optionValue2);
            	}
        	} else {
            	$(styleEle).css(prop, optionValue1[0]);
        	}
        	style.font = styleEle.style.font;
    	}

    	sheet.suspendPaint();

    	for (var n = 0; n < selections.length; n++) {
        	var sel = this.getActualCellRange(sheet, selections[n], rowCount, columnCount);
        	for (var r = sel.row; r < sel.row + sel.rowCount; r++) {
            	for (var c = sel.col; c < sel.col + sel.colCount; c++) {
                	var style = sheet.getStyle(r, c);
                	if (!style) {
                    	style = new GC.Spread.Sheets.Style();
                	}
                	// reset themeFont to make sure font be used
                	style.themeFont = undefined;
                	if (isLabelStyle) {
                    	if (!style.labelOptions) {
                        	style.labelOptions = {};
                    	}
                    	updateStyleFont(style.labelOptions);
                	} else {
                    	updateStyleFont(style)
                	}
                	sheet.setStyle(r, c, style);
            	}
        	}
    	}

    	sheet.resumePaint();
	}

	getActualCellRange(sheet:GC.Spread.Sheets.Worksheet, cellRange:GC.Spread.Sheets.Range, rowCount:number, columnCount:number) {
		if (cellRange.row === -1 && cellRange.col === -1) {
			return new GC.Spread.Sheets.CellRange(sheet, 0, 0, rowCount, columnCount);
		}
		else if (cellRange.row === -1) {
			return new GC.Spread.Sheets.CellRange(sheet, 0, cellRange.col, rowCount, cellRange.colCount);
		}
		else if (cellRange.col === -1) {
			return new GC.Spread.Sheets.CellRange(sheet, cellRange.row, 0, cellRange.rowCount, columnCount);
		}
		return new GC.Spread.Sheets.CellRange(sheet, cellRange.row, cellRange.col, cellRange.rowCount, cellRange.colCount);
	}
}