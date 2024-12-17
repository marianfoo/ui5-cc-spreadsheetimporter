import ManagedObject from "sap/ui/base/ManagedObject";
import Fragment from "sap/ui/core/Fragment";
import SpreadsheetUpload from "../SpreadsheetUpload";
import SpreadsheetDialog, { SpreadsheetDialog$AvailableOptionsChangedEvent, SpreadsheetDialog$DecimalSeparatorChangedEvent, SpreadsheetDialog$FileDropEvent } from "../../control/SpreadsheetDialog";
import ResourceModel from "sap/ui/model/resource/ResourceModel";
import Component from "../../Component";
import Event from "sap/ui/base/Event";
import FileUploader, { FileUploader$ChangeEvent } from "sap/ui/unified/FileUploader";
import MessageToast from "sap/m/MessageToast";
import Preview from "../Preview";
import Util from "../Util";
import ResourceBundle from "sap/base/i18n/ResourceBundle";
import * as XLSX from "xlsx";
import OptionsDialog from "./OptionsDialog";
import MessageHandler from "../MessageHandler";
import Log from "sap/base/Log";
import SheetHandler from "../SheetHandler";
import Parser from "../Parser";
import Button from "sap/m/Button";
import { ArrayData, AvailableOptionsType, DeepDownloadConfig, FireEventReturnType } from "../../types";
import FlexBox from "sap/m/FlexBox";
import JSONModel from "sap/ui/model/json/JSONModel";
import Dialog from "sap/m/Dialog";
import Select from "sap/m/Select";
import Item from "sap/ui/core/Item";
import SpreadsheetDownloadDialog from "../download/SpreadsheetDownloadDialog";
import SpreadsheetGenerator from "../download/SpreadsheetGenerator";
import SpreadsheetDownload from "../download/SpreadsheetDownload";
import OData from "../odata/OData";

type InputType = {
	[key: string]: {
		rawValue: any;
		[additionalKeys: string]: any;
	};
};

/**
 * @namespace cc.spreadsheetimporter.XXXnamespaceXXX
 */
export default class SpreadsheetUploadDialog extends ManagedObject {
	spreadsheetUploadController: SpreadsheetUpload;
	spreadsheetUploadDialog: SpreadsheetDialog;
	spreadsheetDownloadDialog: SpreadsheetDownloadDialog;
	component: Component;
	previewHandler: Preview;
	util: Util;
	componentI18n: ResourceModel;
	optionsHandler: OptionsDialog;
	messageHandler: MessageHandler;
	spreadsheetOptionsModel: JSONModel;
	spreadsheetGenerator: SpreadsheetGenerator;
	spreadsheetDownload: SpreadsheetDownload;

	constructor(spreadsheetUploadController: SpreadsheetUpload, component: Component, componentI18n: ResourceModel, messageHandler: MessageHandler) {
		super();
		this.spreadsheetUploadController = spreadsheetUploadController;
		this.component = component;
		this.componentI18n = componentI18n;
		this.util = new Util(componentI18n.getResourceBundle() as ResourceBundle);
		this.previewHandler = new Preview(this.util);
		this.optionsHandler = new OptionsDialog(spreadsheetUploadController);
		this.messageHandler = messageHandler;
		this.spreadsheetDownloadDialog = new SpreadsheetDownloadDialog(this.spreadsheetUploadController, this);
	}

	async createSpreadsheetUploadDialog() {
		if (!this.spreadsheetUploadDialog) {
			this.spreadsheetOptionsModel = new JSONModel({
				dataRows: 0,
				strict: this.component.getStrict(),
				hidePreview: this.component.getHidePreview(),
				showOptions: this.component.getShowOptions(),
				showDownloadButton: this.component.getShowDownloadButton(),
				hideGenerateTemplateButton: false,
				fileUploadValue: "",
				densityClass: this.component._densityClass
			});
			this.spreadsheetUploadDialog = (await Fragment.load({
				name: "cc.spreadsheetimporter.XXXnamespaceXXX.fragment.SpreadsheetUpload",
				type: "XML",
				controller: this
			})) as SpreadsheetDialog;
			this.spreadsheetUploadDialog.setComponent(this.component);
			this.spreadsheetUploadDialog.setBusyIndicatorDelay(0);
			this.spreadsheetUploadDialog.setModel(this.componentI18n, "i18n");
			this.spreadsheetUploadDialog.setModel(this.spreadsheetOptionsModel, "info");
			this.spreadsheetUploadDialog.setModel(this.component.getModel("device"), "device");
			this.spreadsheetUploadDialog.attachDecimalSeparatorChanged(this.onDecimalSeparatorChanged.bind(this));
			this.spreadsheetUploadDialog.attachAvailableOptionsChanged(this.onAvailableOptionsChanged.bind(this));
			this.spreadsheetUploadDialog.attachFileDrop(this.onFileDrop.bind(this));
		}
		if (this.component.getStandalone() && 
			this.component.getColumns().length === 0 && 
			!this.component.getSpreadsheetTemplateFile()) {
			this.spreadsheetOptionsModel.setProperty("/hideGenerateTemplateButton", true);
		}
	}

	onFileDrop(event: SpreadsheetDialog$FileDropEvent) {
		const files = event.getParameter("files");
		const file = files[0] as File;
		(this.spreadsheetUploadDialog.getModel("info") as JSONModel).setProperty("/fileUploadValue", file.name);
		this.handleFile(file);
	}

	/**
	 * Handles file upload event.
	 * @param {Event} event - The file upload event
	 */
	async onFileUpload(event: FileUploader$ChangeEvent) {
		this.messageHandler.setMessages([]);
		const file = event.getParameter("files")[0] as Blob;
		try {
			const asyncEventPreFileProcessing = await Util.fireEventAsync("preFileProcessing", { file: file }, this.component);
			const isDefaultPrevented = asyncEventPreFileProcessing.bPreventDefault;
			if (isDefaultPrevented) {
				Log.info("preFileProcessing event was prevented", "SpreadsheetUpload: onFileUpload");
				this.resetContent();
				return;
			} else {
				this.handleFile(file);
			}
		} catch (error) {
			Log.error("Error while calling the preFileProcessing event", error as Error, "SpreadsheetUpload: onFileUpload");
		}
	}

	async handleFile(file: Blob) {
		try {
			this.messageHandler.setMessages([]);
			const workbook = (await this._readWorkbook(file)) as XLSX.WorkBook;

			const isStandalone = this.component.getStandalone();
			const readAllSheets = this.component.getReadAllSheets();

			let spreadsheetSheetsData: ArrayData = [];
			let columnNames: string[] = [];

			if (isStandalone && readAllSheets) {
				// Loop over the sheet names in the workbook
				for (const sheetName of Object.keys(workbook.Sheets)) {
					let currSheetData = SheetHandler.sheet_to_json(workbook.Sheets[sheetName]);
					for (const dataVal of currSheetData) {
						Object.keys(dataVal).forEach((key) => {
							dataVal[key].sheetName = sheetName;
						});
					}

					spreadsheetSheetsData = spreadsheetSheetsData.concat(currSheetData);
					columnNames = columnNames.concat(XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 })[0] as string[]);
				}
			} else {
				const sheetOption = this.spreadsheetUploadController.component.getReadSheet();
				let sheetName: string;
				try {
					sheetName = await this._getSheetName(workbook, sheetOption);
				} catch (error) {
					this.resetContent();
					return;
				}
				spreadsheetSheetsData = SheetHandler.sheet_to_json(workbook.Sheets[sheetName]);
				const rawColumns = SheetHandler.sheet_to_json(workbook.Sheets[sheetName], { header: 1 })[0] as Array<{rawValue: string}>;
				columnNames = rawColumns.map(column => column.rawValue || column as unknown as string);
			}

			if (!spreadsheetSheetsData || spreadsheetSheetsData.length === 0) {
				throw new Error(this.util.geti18nText("spreadsheetimporter.emptySheet"));
			}

			//remove empty spaces before and after every value
			for (const object of spreadsheetSheetsData) {
				for (const key in object) {
					object[key].rawValue = typeof object[key].rawValue === "string" ? object[key].rawValue.trim() : object[key].rawValue;
				}
			}

			if (!isStandalone) {
				this.messageHandler.checkFormat(spreadsheetSheetsData);
				this.messageHandler.checkMandatoryColumns(
					spreadsheetSheetsData,
					columnNames,
					this.spreadsheetUploadController.odataKeyList,
					this.component.getMandatoryFields(),
					this.spreadsheetUploadController.typeLabelList
				);
				this.messageHandler.checkDuplicateColumns(columnNames);
				if (!this.component.getSkipMaxLengthCheck()) {
					this.messageHandler.checkMaxLength(spreadsheetSheetsData, this.spreadsheetUploadController.typeLabelList, this.component.getFieldMatchType());
				}
				if (!this.component.getSkipColumnsCheck()) {
					this.messageHandler.checkColumnNames(columnNames, this.component.getFieldMatchType(), this.spreadsheetUploadController.typeLabelList);
				}
			}
			this.spreadsheetUploadController.payload = spreadsheetSheetsData;

			this.spreadsheetUploadController.payloadArray = isStandalone
				? this.spreadsheetUploadController.payload
				: Parser.parseSpreadsheetData(
						this.spreadsheetUploadController.payload,
						this.spreadsheetUploadController.typeLabelList,
						this.component,
						this.messageHandler,
						this.util,
						this.spreadsheetUploadController.isODataV4
					);

			this.getDialog().setBusy(true);
			try {
				await Util.fireEventAsync("checkBeforeRead", { sheetData: spreadsheetSheetsData, parsedData: this.spreadsheetUploadController.payloadArray }, this.component);
			} catch (error) {
				Log.error("Error while calling the checkBeforeRead event", error as Error, "SpreadsheetUpload: onFileUpload");
			}
			this.getDialog().setBusy(false);

			if (this.messageHandler.areMessagesPresent()) {
				// show error dialog
				this.messageHandler.displayMessages();
				return;
			}
			this.setDataRows(this.spreadsheetUploadController.payloadArray.length);
		} catch (error) {
			Util.showError(error, "SpreadsheetUpload.ts", "onFileUpload");
			this.resetContent();
		}
	}

	/**
	 * Sending extracted data to backend
	 * @param {*} event
	 */
	async onUploadSet(event: Event) {
		const source = event.getSource() as Button;
		this.getDialog().setBusy(true);
		let fireEventAsyncReturn: FireEventReturnType;
		let isDefaultPrevented;
		try {
			fireEventAsyncReturn = await Util.fireEventAsync(
				"uploadButtonPress",
				{
					payload: this.spreadsheetUploadController.payloadArray,
					rawData: this._extractRawValues(this.spreadsheetUploadController.payloadArray),
					parsedData: this._extractParsedValues(this.spreadsheetUploadController.payloadArray)
				},
				this.component
			);
			isDefaultPrevented = fireEventAsyncReturn.bPreventDefault;
		} catch (error) {}
		this.getDialog().setBusy(false);
		if (isDefaultPrevented || this.component.getStandalone()) {
			this.onCloseDialog();
			if (this.messageHandler.areMessagesPresent()) {
				this.messageHandler.displayMessages(true);
			}
			return;
		}
		// checking if spreadsheet file contains data or not
		if (!this.spreadsheetUploadController.payloadArray.length) {
			MessageToast.show(this.util.geti18nText("spreadsheetimporter.selectFileUpload"));
			return;
		}

		var that = this;
		const sourceParent = source.getParent() as SpreadsheetDialog;

		sourceParent.setBusyIndicatorDelay(0);
		sourceParent.setBusy(true);
		await Util.sleep(50);

		// creating a promise as the extension api accepts odata call in form of promise only
		var fnAddMessage = function () {
			return new Promise((fnResolve, fnReject) => {
				that.spreadsheetUploadController.getODataHandler().callOdata(fnResolve, fnReject, that.spreadsheetUploadController);
			});
		};

		var mParameters = {
			busy: {
				set: true,
				check: false
			},
			dataloss: {
				popup: false,
				navigation: false
			},
			sActionLabel: this.util.geti18nText("spreadsheetimporter.uploadingFile")
		};

		// calling the OData service using extension api
		if (this.spreadsheetUploadController.isODataV4) {
			try {
				if(this.spreadsheetUploadController.context.editFlow){
					await this.spreadsheetUploadController.context.editFlow.securedExecution(fnAddMessage, mParameters);
				}else{
					await fnAddMessage();
				}
			} catch (error) {
				Log.error("Error while calling the odata service", error as Error, "SpreadsheetUpload: onUploadSet");
				this.spreadsheetUploadController.errorsFound = true;
				this.resetContent();
			}
		} else {
			try {
				if (this.spreadsheetUploadController.context.extensionAPI) {
					await this.spreadsheetUploadController.context.extensionAPI.securedExecution(fnAddMessage, mParameters);
				} else {
					await fnAddMessage();
				}
			} catch (error) {
				Log.error("Error while calling the odata service", error as Error, "SpreadsheetUpload: onUploadSet");
				this.spreadsheetUploadController.errorsFound = true;
				this.resetContent();
			}
		}

		sourceParent.setBusy(false);
		this.getDialog().setBusy(true);
		try {
			await Util.fireEventAsync("requestCompleted", { success: !this.spreadsheetUploadController.errorsFound }, this.component);
		} catch (error) {
			Log.error("Error while calling the requestCompleted event", error as Error, "SpreadsheetUpload: onUploadSet");
		}
		this.getDialog().setBusy(false);
		this.onCloseDialog();
	}

	openSpreadsheetUploadDialog() {
		this.spreadsheetUploadDialog.open();
	}

	/**
	 * Closes the Spreadsheet upload dialog.
	 */
	onCloseDialog() {
		this.component.fireRequestCompleted();
		this.resetContent();
		this.spreadsheetUploadDialog.close();
	}

	onDecimalSeparatorChanged(event: SpreadsheetDialog$DecimalSeparatorChangedEvent) {
		this.component.setDecimalSeparator(event.getParameter("decimalSeparator"));
	}

	onAvailableOptionsChanged(event: SpreadsheetDialog$AvailableOptionsChangedEvent) {
		const availableOptions = event.getParameter("availableOptions") as AvailableOptionsType[];
		if (availableOptions.length > 0) {
			this.component.setShowOptions(true);
			this.spreadsheetOptionsModel.setProperty("/showOptions", true);
		} else {
			this.component.setShowOptions(false);
			this.spreadsheetOptionsModel.setProperty("/showOptions", true);
		}
		this.component.setAvailableOptions(availableOptions);
	}

	resetContent() {
		(this.spreadsheetUploadDialog.getModel("info") as JSONModel).setProperty("/dataRows", 0);
		var fileUploader = (this.spreadsheetUploadDialog.getContent()[0] as FlexBox).getItems()[1] as FileUploader;
		fileUploader.setValue();
	}

	setDataRows(length: number) {
		(this.spreadsheetUploadDialog.getModel("info") as JSONModel).setProperty("/dataRows", length);
	}

	getDialog(): SpreadsheetDialog {
		return this.spreadsheetUploadDialog;
	}

	async showPreview() {
		this.previewHandler.showPreview(this.spreadsheetUploadController.getPayloadArray(), this.spreadsheetUploadController.typeLabelList, this.component.getPreviewColumns());
	}

	async onTempDownload() {
		// check if custom template is provided, otherwise generate it
		if (this.component.getSpreadsheetTemplateFile() !== "") {
			try {
				const templateFile = this.component.getSpreadsheetTemplateFile();
				let arrayBuffer;
				let fileName;

				if (typeof templateFile === "string") {
					// Check if the string is a HTTP/HTTPS address
					if (templateFile.startsWith("http://") || templateFile.startsWith("https://")) {
						const response = await fetch(templateFile);
						if (!response.ok) {
							throw new Error("Network response was not ok " + response.statusText);
						}
						fileName = templateFile.split("/").pop();
						arrayBuffer = await response.arrayBuffer();
					}
					// Assume the string is a local file path
					else {
						const sPath = sap.ui.require.toUrl(templateFile);
						const response = await fetch(sPath);
						if (!response.ok) {
							throw new Error("Network response was not ok " + response.statusText);
						}
						fileName = templateFile.split("/").pop();
						arrayBuffer = await response.arrayBuffer();
					}
				} else if (templateFile instanceof ArrayBuffer) {
					// If the input is already an ArrayBuffer, use it directly
					arrayBuffer = templateFile;
				} else {
					throw new Error("Unsupported type for templateFile");
				}

				// spreadsheet file name wll overwrite the template file name
				if (this.component.getSpreadsheetFileName() !== "Template.xlsx" || fileName === undefined) {
					fileName = this.component.getSpreadsheetFileName();
				}

				Util.downloadSpreadsheetFile(arrayBuffer, fileName);

				// You can now use arrayBuffer to do whatever you need
			} catch (error) {
				console.error("Error loading file", error);
			}
		} else {
			// create spreadsheet column list
			let fieldMatchType = this.component.getFieldMatchType();
			var worksheet = {} as XLSX.WorkSheet;
			let colWidths: { wch: number }[] = []; // array to store column widths
			let sampleData = this.component.getSampleData() as any[];
			let sampleDataDefined = true;
			// if sampledata is empty add one row of empty data
			if (!sampleData || sampleData.length === 0) {
				sampleDataDefined = false;
				sampleData = [{}];
			}
			const colWidthDefault = 15;
			const colWidthDate = 20;
			let col = 0;
			let rows = 1;
			if (this.component.getStandalone()) {
				// loop over this.component.getColumns
				for (let column of this.component.getColumns()) {
					worksheet[XLSX.utils.encode_cell({ c: col, r: 0 })] = { v: column, t: "s" };
					col++;
				}
				col = 0;
				for (let column of this.component.getColumns()) {
					for (const [index, data] of sampleData.entries()) {
						let sampleDataValue;
						rows = index + 1;
						if (data[column]) {
							sampleDataValue = data[column];
							worksheet[XLSX.utils.encode_cell({ c: col, r: 1 })] = { v: sampleDataValue, t: "s" };
						}
					}

					col++;
				}
			} else {
				for (let [key, value] of this.spreadsheetUploadController.typeLabelList.entries()) {
					let cell = { v: "", t: "s" } as XLSX.CellObject;
					let label = "";
					if (fieldMatchType === "label") {
						label = value.label;
					}
					if (fieldMatchType === "labelTypeBrackets") {
						label = `${value.label}[${key}]`;
					}
					worksheet[XLSX.utils.encode_cell({ c: col, r: 0 })] = { v: label, t: "s" };

					for (const [index, data] of sampleData.entries()) {
						let sampleDataValue;
						rows = index + 1;
						if (data[key]) {
							sampleDataValue = data[key];
						}
						if (value.type === "Edm.Boolean") {
							// Set default value for sampleDataValue based on sampleDataDefined flag
							let defaultValue = sampleDataDefined ? "" : "true";
							// Assign sampleDataValue only if sampleDataValue is not undefined
							sampleDataValue = sampleDataValue ? sampleDataValue.toString() : defaultValue;
							cell = {
								v: sampleDataValue,
								t: "b"
							};
							colWidths.push({ wch: colWidthDefault });
						} else if (value.type === "Edm.String") {
							let newStr;
							if (value.maxLength) {
								newStr = sampleDataValue ? sampleDataValue : "test string".substring(0, value.maxLength);
							} else {
								newStr = sampleDataValue ? sampleDataValue : "test string";
							}
							// Set default value for sampleDataValue based on sampleDataDefined flag
							let defaultValue = sampleDataDefined ? "" : newStr;
							// Assign sampleDataValue only if sampleDataValue is not undefined
							sampleDataValue = sampleDataValue ? sampleDataValue : defaultValue;
							cell = { v: sampleDataValue, t: "s" };
							colWidths.push({ wch: colWidthDefault });
						} else if (value.type === "Edm.DateTimeOffset" || value.type === "Edm.DateTime") {
							let format;
							const currentLang = await Util.getLanguage();
							if (currentLang.startsWith("en")) {
								format = "mm/dd/yyyy hh:mm AM/PM";
							} else {
								format = "dd.mm.yyyy hh:mm";
							}

							// Set default value for sampleDataValue based on sampleDataDefined flag
							let defaultValue = sampleDataDefined ? "" : new Date();
							// Assign sampleDataValue only if sampleDataValue is not undefined
							sampleDataValue = sampleDataValue ? sampleDataValue : defaultValue;
							// if sampleDataValue is empty and cellFormat is "d", Excel Generation fails
							let cellFormat: XLSX.ExcelDataType = sampleDataValue ? "d" : "s";
							cell = { v: sampleDataValue, t: cellFormat, z: format };
							colWidths.push({ wch: colWidthDate }); // set column width to 20 for this column
						} else if (value.type === "Edm.Date") {
							// Set default value for sampleDataValue based on sampleDataDefined flag
							let defaultValue = sampleDataDefined ? "" : new Date();
							// Assign sampleDataValue only if sampleDataValue is not undefined
							sampleDataValue = sampleDataValue ? sampleDataValue.toString() : defaultValue;
							// if sampleDataValue is empty and cellFormat is "d", Excel Generation fails
							let cellFormat: XLSX.ExcelDataType = sampleDataValue ? "d" : "s";
							cell = {
								v: sampleDataValue,
								t: cellFormat
							};
							colWidths.push({ wch: colWidthDefault });
						} else if (value.type === "Edm.TimeOfDay" || value.type === "Edm.Time") {
							// Set default value for sampleDataValue based on sampleDataDefined flag
							let defaultValue = sampleDataDefined ? "" : new Date();
							// Assign sampleDataValue only if sampleDataValue is not undefined
							sampleDataValue = sampleDataValue ? sampleDataValue : defaultValue;
							// if sampleDataValue is empty and cellFormat is "d", Excel Generation fails
							let cellFormat: XLSX.ExcelDataType = sampleDataValue ? "d" : "s";
							cell = {
								v: sampleDataValue,
								t: cellFormat,
								z: "hh:mm"
							};
							colWidths.push({ wch: colWidthDefault });
						} else if (
							value.type === "Edm.UInt8" ||
							value.type === "Edm.Int16" ||
							value.type === "Edm.Int32" ||
							value.type === "Edm.Integer" ||
							value.type === "Edm.Int64" ||
							value.type === "Edm.Integer64"
						) {
							// Set default value for sampleDataValue based on sampleDataDefined flag
							let defaultValue = sampleDataDefined ? "" : 85;
							// Assign sampleDataValue only if sampleDataValue is not undefined
							sampleDataValue = sampleDataValue ? sampleDataValue : defaultValue;
							cell = {
								v: sampleDataValue,
								t: "n"
							};
							colWidths.push({ wch: colWidthDefault });
						} else if (value.type === "Edm.Double" || value.type === "Edm.Decimal") {
							const decimalSeparator = this.component.getDecimalSeparator();
							// Set default value for sampleDataValue based on sampleDataDefined flag
							let defaultValue = sampleDataDefined ? "" : `123${decimalSeparator}4`;
							// Assign sampleDataValue only if sampleDataValue is not undefined
							sampleDataValue = sampleDataValue ? sampleDataValue.toString() : defaultValue;
							cell = {
								v: sampleDataValue,
								t: "n"
							};
							colWidths.push({ wch: colWidthDefault });
						}

						if (!this.component.getHideSampleData()) {
							worksheet[XLSX.utils.encode_cell({ c: col, r: rows })] = cell;
						}
					}
					col++;
				}
			}
			worksheet["!ref"] = XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: col, r: sampleData.length } });
			worksheet["!cols"] = colWidths; // assign the column widths to the worksheet

			// creating the new spreadsheet work book
			const wb = XLSX.utils.book_new();
			// set the file value
			XLSX.utils.book_append_sheet(wb, worksheet, "Tabelle1");
			// download the created spreadsheet file
			XLSX.writeFile(wb, this.component.getSpreadsheetFileName());

			MessageToast.show(this.util.geti18nText("spreadsheetimporter.downloadingTemplate"));
		}
	}

	onOpenOptionsDialog() {
		this.optionsHandler.openOptionsDialog();
	}

	/**
	 * Read the uploaded workbook from the file.
	 * @param {File} file - The uploaded file.
	 * @returns {Promise} - Promise object representing the workbook.
	 */
	async _readWorkbook(file: Blob) {
		return new Promise(async (resolve, reject) => {
			try {
				const data = await this.buffer_RS(file.stream());
				let workbook = XLSX.read(data, { cellNF: true, cellDates: true, cellText: true, cellFormula: true });
				resolve(workbook);
			} catch (error) {
				Log.error("Error while reading the uploaded workbook", error as Error, "SpreadsheetUpload: _readWorkbook");
				reject(error);
			}
		});
	}

	async buffer_RS(stream: ReadableStream) {
		/* collect data */
		const buffers = [];
		const reader = stream.getReader();
		for (;;) {
			const res = await reader.read();
			if (res.value) buffers.push(res.value);
			if (res.done) break;
		}

		/* concat */
		const out = new Uint8Array(buffers.reduce((acc, v) => acc + v.length, 0));

		let off = 0;
		for (const u8 of buffers) {
			out.set(u8, off);
			off += u8.length;
		}

		return out;
	}

	_extractRawValues(data: InputType[]): any[] {
		return data.map((item) => {
			const newObj: { [key: string]: any } = {};

			for (const key in item) {
				if (item[key].hasOwnProperty("rawValue")) {
					newObj[key] = item[key].rawValue;
				}
			}

			return newObj;
		});
	}

	_extractParsedValues(data: InputType[]): any[] {
		return data.map((item) => {
			const newObj: { [key: string]: any } = {};

			for (const key in item) {
				if (item[key].hasOwnProperty("formattedValue")) {
					newObj[key] = item[key].formattedValue;
				}
			}

			return newObj;
		});
	}

	async _getSheetName(workbook: XLSX.WorkBook, sheetOption: string): Promise<string> {
		let sheetName: string;
		// Check if sheetOption is a number
		if (typeof sheetOption === "number") {
			if (sheetOption >= 0 && sheetOption < workbook.SheetNames.length) {
				sheetName = workbook.SheetNames[sheetOption];
			} else {
				Log.error("Invalid sheet index, defaulting to first Sheet", "SpreadsheetUpload: _getSheetName");
				sheetName = workbook.SheetNames[0];
			}
		}
		// Check if sheetOption is "XXSelectorXX"
		else if (sheetOption === "XXSelectorXX") {
			if (workbook.SheetNames.length === 1) {
				sheetName = workbook.SheetNames[0];
				Log.debug("Only one sheet in workbook, defaulting to first Sheet", "SpreadsheetUpload: _getSheetName");
			} else {
				// Display a selector dialog and get the selected sheet name
				sheetName = await this._displaySheetSelectorDialog(workbook.SheetNames);
			}
		}
		// Check if sheetOption is a string and exists in workbook.SheetNames
		else if (workbook.SheetNames.includes(sheetOption)) {
			sheetName = sheetOption;
		} else {
			Log.error("Invalid sheet name, defaulting to first Sheet", "SpreadsheetUpload: _getSheetName");
			sheetName = workbook.SheetNames[0];
		}
		return sheetName;
	}

	_displaySheetSelectorDialog(sheetNames: string[]): Promise<string> {
		// Display a selector dialog and get the selected sheet name
		// Assuming you have a function called displaySelectorDialog that returns the selected sheet name
		return new Promise((resolve, reject) => {
			const select = new Select();

			sheetNames.forEach((sheetName) => {
				select.addItem(
					new Item({
						key: sheetName,
						text: sheetName
					})
				);
			});

			const i18n = this.componentI18n.getResourceBundle() as ResourceBundle;

			const dialog = new Dialog({
				title: i18n.getText("spreadsheetimporter.sheetSelectorDialogTitle"),
				type: "Message",
				content: [select],
				beginButton: new Button({
					text: i18n.getText("spreadsheetimporter.ok"),
					press: () => {
						const selectedKey = select.getSelectedKey();
						resolve(selectedKey);
						dialog.close();
					}
				}),
				afterClose: () => dialog.destroy(),
				endButton: new Button({
					text: i18n.getText("close"),
					press: () => {
						reject(new Error(i18n.getText("close")));
						dialog.close();
					}
				})
			}) as Dialog;

			dialog.open();
		});
	}

	setODataHandler(odataHandler: OData) {
		this.spreadsheetGenerator = new SpreadsheetGenerator(this.spreadsheetUploadController, this.component, odataHandler);
		this.spreadsheetDownload = new SpreadsheetDownload(this.spreadsheetUploadController, this.component, odataHandler);
	}

	/**
	 * Initializes the spreadsheet download process.
	 * If showOptions is enabled in the DeepDownloadConfig, opens a dialog allowing users to configure download options.
	 * Otherwise, directly triggers the spreadsheet download.
	 * 
	 * @returns {Promise<void>} A promise that resolves when the download process is initialized
	 */
	async onInitDownloadSpreadsheetProcess(): Promise<void> {
		const showOptionsToUser = (this.component.getDeepDownloadConfig() as DeepDownloadConfig).showOptions;
		if (showOptionsToUser) {
			await this.spreadsheetDownloadDialog.createSpreadsheetDownloadDialog();
			this.spreadsheetDownloadDialog.spreadsheetDownloadDialog.open();
		} else {
			this.onDownloadDataSpreadsheet();
		}
	}

	async onDownloadDataSpreadsheet(): Promise<void> {
		if (!this.spreadsheetUploadController.errorState) {
			try {
				const mainEntitySiblings = await this.spreadsheetDownload.fetchData(this.component.getDeepDownloadConfig() as DeepDownloadConfig);

				let isDefaultPrevented = false;
				try {
					const asyncEventBeforeDownloadFileProcessing = await Util.fireEventAsync("beforeDownloadFileProcessing", { data: mainEntitySiblings }, this.component);
					isDefaultPrevented = asyncEventBeforeDownloadFileProcessing.bPreventDefault;
				} catch (error) {
					Log.error("Error while calling the beforeDownloadFileProcessing event", error as Error, "SpreadsheetUploadDialog.ts");
				}
				if (!isDefaultPrevented) {
					this.spreadsheetGenerator.downloadSpreadsheet(mainEntitySiblings, this.component.getDeepDownloadConfig() as DeepDownloadConfig);
				}
			} catch (error) {
				console.error("Error in onDownloadDataSpreadsheet:", error);
			}
		} else {
			Util.showError(this.spreadsheetUploadController.errorMessage, "SpreadsheetUpload.ts", "initialSetup");
			Log.error("Error opening the dialog", undefined, "SpreadsheetUpload: SpreadsheetUpload");
		}
	}
}
