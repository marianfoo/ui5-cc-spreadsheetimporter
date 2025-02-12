const { escape } = require("querystring");
const path = require('path');
const downloadDir = path.resolve(__dirname, 'downloads');
const util = require("./../../dev/util");
const { TimelineService } = require('wdio-timeline-reporter/timeline-service');
let scenario = "";
let version = 0;

for (let index = 0; index < process.argv.length; index++) {
	const arg = process.argv[index];
	if (arg.startsWith("orders")) {
		scenario = arg;
		version = process.argv[index + 1];
	}
}

const testappObject = util.getTestappObject(scenario, version);
const specs = testappObject["testMapping"]["specs"];
const port = testappObject.port;
let baseUrl = `http://localhost:${port}/index.html?sap-language=EN`;
global.scenario = scenario;

const isDebugEnabled = process.env.ACTIONS_STEP_DEBUG === 'true';

module.exports.config = {
	wdi5: {
		logLevel: "error",
		waitForUI5Timeout: 180000
	},
	specs: specs,
	exclude: [
		// 'path/to/excluded/files'
	],
	maxInstances: 10,
	//
	capabilities: [
		{
			maxInstances: 5,
			//
			browserName: "chrome",
			browserVersion: 'stable',
			"goog:chromeOptions": {
				args:
					process.argv.indexOf("--headless") > -1
						? ["--headless=new", "--window-size=1920,1080", "--no-sandbox"]
						: process.argv.indexOf("--debug") > -1
						? ["--window-size=1920,1080", "--auto-open-devtools-for-tabs"]
						: ["--window-size=1920,1080"],
            prefs: {
                'download.default_directory': downloadDir,
                'download.prompt_for_download': false,
                'download.directory_upgrade': true
            }
			},
			acceptInsecureCerts: true
		}
	],
	logLevel: "error",
	bail: 0,
	baseUrl: baseUrl,
	waitforTimeout: 60000,
	connectionRetryTimeout: process.argv.indexOf("--debug") > -1 ? 1200000 : 120000,
	connectionRetryCount: 3,
	services: isDebugEnabled 
		? ["ui5", [TimelineService]]
		: ["ui5"],
	framework: "mocha",
	reporters: isDebugEnabled 
		? [
			"spec",
			["timeline", { outputDir: "./reports/timeline", embedImages: true, screenshotStrategy: "before:click" }]
		]
		: ["spec"],
	mochaOpts: {
		ui: "bdd",
		timeout: process.argv.indexOf("--debug") > -1 ? 600000 : 600000
	}
};
