import UIComponent from "sap/ui/core/UIComponent";
import JSONModel from "sap/ui/model/json/JSONModel";
import Device from "sap/ui/Device";
import ExcelUpload from "./controller/ExcelUpload";
/**
 * @namespace cc.excelUpload.XXXnamespaceXXX
 */
export default class Component extends UIComponent {
	constructor(idOrSettings?: string | $ComponentSettings);
	constructor(id?: string, settings?: $ComponentSettings);
	constructor(id?: string, settings?: $ComponentSettings) {
		super(id, settings);
	}

	public static metadata = {
		manifest: "json",
		properties: {
			excelFileName: { type: "string", defaultValue: "Template.xlsx" },
			context: { type: "object" },
			columns: { type: "string[]" },
			tableId: { type: "string" },
			odataType: { type: "string" },
			mandatoryFields: { type: "string[]" },
			errorResults: { type: "object" },
			fieldMatchType: { type: "string", defaultValue: "label" },
			activateDraft: { type: "boolean", defaultValue: false },
		},
		aggregations: {
			rootControl: {
				type: "sap.ui.core.Control",
				multiple: false,
				visibility: "hidden",
			},
		},
		events: {
			checkBeforeRead: {
				parameters: {
					sheetData: { type: "object" },
					errorResults: { type: "object" },
				},
			},
			changeBeforeCreate: {
				parameters: {
					payload: { type: "object" },
				},
			},
		},
	};

	//=============================================================================
	//LIFECYCLE APIS
	//=============================================================================

	public async init(): Promise<void> {
		var oModel, oCompData;

		oCompData = this.getComponentData();
		// if (typeof oCompData.renderButton === "boolean"){
		// 	this.setRenderButton(oCompData.renderButton);
		// }
		this.setContext(oCompData.context);
		this.setColumns(oCompData.columns);
		this.setTableId(oCompData.tableId);
		this.setOdataType(oCompData.odataType);
		this.setMandatoryFields(oCompData.mandatoryFields);
		this.setFieldMatchType(oCompData.fieldMatchType);
		this.setActivateDraft(oCompData.activateDraft);

		// call the init function of the parent - ATTENTION: this triggers createContent()
		// call the base component's init function
		super.init();

		// this.excelUpload = await sap.ui.core.mvc.Controller.create({ name:"cc.excelUpload.XXXnamespaceXXX.ExcelUpload"})
		// //now this here would work:
		// //var oRoot = this.getRootControl(); → won't work with visibility: "hidden", no getters/setters generated

		// // we could create a device model and use it
		oModel = new JSONModel(Device);
		oModel.setDefaultBindingMode("OneWay");
		this.setModel(oModel, "device");
	}

	// Component.prototype.setContextPublic = function(options) {
	//     this.excelUpload.setContext(options)
	// };

	createContent() {
		this.excelUpload = new ExcelUpload(this, this.getModel("i18n"));
		// this.excelUpload = await Controller.create({ name:"cc.excelUpload.XXXnamespaceXXX.controller.ExcelUpload"})

		// var oBtn, oTSD;

		// oTSD = this._getCustomerSelectDialog();

		// if (this.getRenderButton()) {
		// 	oBtn = this._getOpenButton();
		// 	oBtn.addDependent(oTSD);
		// 	return oBtn;
		// }
		// return oTSD;
	}

	//=============================================================================
	//OVERRIDE SETTERS
	//=============================================================================

	// /**
	//  * Overrides method <code>getErrorResults</code> of the component to set this text in the button.
	//  * @override
	//  */
	// Component.prototype.getErrorResults = function (array) {
	// 	return this.excelUpload.getErrorResults();
	// };
	// /**
	//  * Overrides method <code>setErrorResults</code> of the component to set this text in the button.
	//  * @override
	//  */
	// Component.prototype.setErrorResults = function (array) {
	// 	this.excelUpload.setErrorResults(array)
	// 	return this;
	// };

	//=============================================================================
	//PUBLIC APIS
	//=============================================================================

	/**
	 * Opens the dialog for selecting a customer.
	 * @public
	 */
	openExcelUploadDialog() {
		this.excelUpload.openExcelUploadDialog();
	}

	/**
	 * Set Payload for Event
	 * @public
	 */
	setPayload(payload) {
		this.excelUpload._setPayload(payload);
	}

	/**
	 * add to error array
	 * @public
	 */
	addToErrorsResults(errorArray) {
		this.excelUpload._addToErrorsResults(errorArray);
	}

	//=============================================================================
	//EVENT HANDLERS
	//=============================================================================

	// Component.prototype.onCheckBeforeRead = function (firstSheet) {
	// 		this.fireCheckBeforeRead({sheetData:firstSheet})
	// };

	onChangeBeforeCreate(oEvent) {
		var aContexts, oCustomer;

		aContexts = oEvent.getParameter("selectedContexts");
	}

	//=============================================================================
	//PRIVATE APIS
	//=============================================================================

	/**
	 * This method can be called to determine whether the sapUiSizeCompact or sapUiSizeCozy
	 * design mode class should be set, which influences the size appearance of some controls.
	 * @private
	 * @return {string} css class, either 'sapUiSizeCompact' or 'sapUiSizeCozy' - or an empty string if no css class should be set
	 */
	getContentDensityClass() {
		if (this._sContentDensityClass === undefined) {
			// check whether FLP has already set the content density class; do nothing in this case
			if (document.body.classList.contains("sapUiSizeCozy") || document.body.classList.contains("sapUiSizeCompact")) {
				this._sContentDensityClass = "";
			} else if (!Device.support.touch) {
				// apply "compact" mode if touch is not supported
				this._sContentDensityClass = "sapUiSizeCompact";
			} else {
				// "cozy" in case of touch support; default for most sap.m controls, but needed for desktop-first controls like sap.ui.table.Table
				this._sContentDensityClass = "sapUiSizeCozy";
			}
		}
		return this._sContentDensityClass;
	}
}
