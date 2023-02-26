sap.ui.define(["sap/fe/test/ListReport"], function (ListReport) {
	"use strict";

	var CustomPageDefinitions = {
		actions: {},
		assertions: {}
	};

	return new ListReport(
		{
			appId: "ui.v4.orders",
			componentId: "OrdersList",
			entitySet: "Orders"
		},
		CustomPageDefinitions
	);
});
