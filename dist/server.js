/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 0);
/******/ })
/************************************************************************/
/******/ ({

/***/ "./src/app/TeamsAppsComponents.ts":
/*!****************************************!*\
  !*** ./src/app/TeamsAppsComponents.ts ***!
  \****************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

function __export(m) {
    for (var p in m) if (!exports.hasOwnProperty(p)) exports[p] = m[p];
}
Object.defineProperty(exports, "__esModule", { value: true });
// Components will be added here
exports.nonce = {}; // Do not remove!
// Automatically added for the powerPillagerBot bot
__export(__webpack_require__(/*! ./powerPillagerBot/PowerPillager */ "./src/app/powerPillagerBot/PowerPillager.ts"));


/***/ }),

/***/ "./src/app/powerPillagerBot/PowerPillager.ts":
/*!***************************************************!*\
  !*** ./src/app/powerPillagerBot/PowerPillager.ts ***!
  \***************************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_msteams_host_1 = __webpack_require__(/*! express-msteams-host */ "express-msteams-host");
const debug = __webpack_require__(/*! debug */ "debug");
const botbuilder_dialogs_1 = __webpack_require__(/*! botbuilder-dialogs */ "botbuilder-dialogs");
const botbuilder_1 = __webpack_require__(/*! botbuilder */ "botbuilder");
const HelpDialog_1 = __webpack_require__(/*! ./dialogs/HelpDialog */ "./src/app/powerPillagerBot/dialogs/HelpDialog.ts");
const WelcomeDialog_1 = __webpack_require__(/*! ./dialogs/WelcomeDialog */ "./src/app/powerPillagerBot/dialogs/WelcomeDialog.ts");
const botbuilder_teams_1 = __webpack_require__(/*! botbuilder-teams */ "botbuilder-teams");
// Initialize debug logging module
const log = debug("msteams");
/**
 * Implementation for Power Pillager
 */
let PowerPillager = class PowerPillager {
    /**
     * The constructor
     * @param conversationState
     */
    constructor(conversationState) {
        this.activityProc = new botbuilder_teams_1.TeamsActivityProcessor();
        this.conversationState = conversationState;
        this.dialogState = conversationState.createProperty("dialogState");
        this.dialogs = new botbuilder_dialogs_1.DialogSet(this.dialogState);
        this.dialogs.add(new HelpDialog_1.default("help"));
        // tslint:disable-next-line: no-console
        console.log("################################ IM RUNNING");
        // Set up the Activity processing
        this.activityProc.messageActivityHandler = {
            // Incoming messages
            onMessage: (context) => __awaiter(this, void 0, void 0, function* () {
                // get the Microsoft Teams context, will be undefined if not in Microsoft Teams
                const teamsContext = botbuilder_teams_1.TeamsContext.from(context);
                // TODO: add your own bot logic in here
                switch (context.activity.type) {
                    case botbuilder_1.ActivityTypes.Message:
                        const text = teamsContext ?
                            teamsContext.getActivityTextWithoutMentions().toLowerCase() :
                            context.activity.text;
                        if (text.startsWith("hello")) {
                            yield context.sendActivity("Oh, hello to you as well!");
                            return;
                        }
                        else if (text.startsWith("help")) {
                            const dc = yield this.dialogs.createContext(context);
                            yield dc.beginDialog("help");
                        }
                        else {
                            yield context.sendActivity(`I\'m terribly sorry, but my master hasn\'t trained me to do anything yet...`);
                        }
                        break;
                    default:
                        break;
                }
                // Save state changes
                return this.conversationState.saveChanges(context);
            })
        };
        this.activityProc.conversationUpdateActivityHandler = {
            onConversationUpdateActivity: (context) => __awaiter(this, void 0, void 0, function* () {
                if (context.activity.membersAdded && context.activity.membersAdded.length !== 0) {
                    for (const idx in context.activity.membersAdded) {
                        if (context.activity.membersAdded[idx].id === context.activity.recipient.id) {
                            const welcomeCard = botbuilder_1.CardFactory.adaptiveCard(WelcomeDialog_1.default);
                            yield context.sendActivity({ attachments: [welcomeCard] });
                        }
                    }
                }
            })
        };
        // Message reactions in Microsoft Teams
        this.activityProc.messageReactionActivityHandler = {
            onMessageReaction: (context) => __awaiter(this, void 0, void 0, function* () {
                const added = context.activity.reactionsAdded;
                if (added && added[0]) {
                    yield context.sendActivity({
                        textFormat: "xml",
                        text: `That was an interesting reaction (<b>${added[0].type}</b>)`
                    });
                }
            })
        };
    }
    /**
     * The Bot Framework `onTurn` handlder.
     * The Microsoft Teams middleware for Bot Framework uses a custom activity processor (`TeamsActivityProcessor`)
     * which is configured in the constructor of this sample
     */
    onTurn(context) {
        return __awaiter(this, void 0, void 0, function* () {
            // transfer the activity to the TeamsActivityProcessor
            yield this.activityProc.processIncomingActivity(context);
        });
    }
};
PowerPillager = __decorate([
    express_msteams_host_1.BotDeclaration("/api/messages", new botbuilder_1.MemoryStorage(), process.env.MICROSOFT_APP_ID, process.env.MICROSOFT_APP_PASSWORD)
], PowerPillager);
exports.PowerPillager = PowerPillager;


/***/ }),

/***/ "./src/app/powerPillagerBot/dialogs/HelpDialog.ts":
/*!********************************************************!*\
  !*** ./src/app/powerPillagerBot/dialogs/HelpDialog.ts ***!
  \********************************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const botbuilder_dialogs_1 = __webpack_require__(/*! botbuilder-dialogs */ "botbuilder-dialogs");
class HelpDialog extends botbuilder_dialogs_1.Dialog {
    constructor(dialogId) {
        super(dialogId);
    }
    beginDialog(context, options) {
        return __awaiter(this, void 0, void 0, function* () {
            context.context.sendActivity(`I'm just a friendly but rather stupid bot, and right now I don't have any valuable help for you!`);
            return yield context.endDialog();
        });
    }
}
exports.default = HelpDialog;


/***/ }),

/***/ "./src/app/powerPillagerBot/dialogs/WelcomeCard.json":
/*!***********************************************************!*\
  !*** ./src/app/powerPillagerBot/dialogs/WelcomeCard.json ***!
  \***********************************************************/
/*! exports provided: $schema, type, version, body, actions, default */
/***/ (function(module) {

module.exports = {"$schema":"http://adaptivecards.io/schemas/adaptive-card.json","type":"AdaptiveCard","version":"1.0","body":[{"type":"Image","url":"https://pillagers-teams-bot.herokuapp.com/assets/icon.png","size":"stretch"},{"type":"TextBlock","spacing":"medium","size":"default","weight":"bolder","text":"Welcome to Pillager bot","wrap":true,"maxLines":0},{"type":"TextBlock","size":"default","isSubtle":true,"text":"Hello, nice to meet you!","wrap":true,"maxLines":0}],"actions":[{"type":"Action.OpenUrl","title":"Learn more about Yo Teams","url":"https://aka.ms/yoteams"},{"type":"Action.OpenUrl","title":"Pillager bot","url":"https://pillagers-teams-bot.herokuapp.com"}]};

/***/ }),

/***/ "./src/app/powerPillagerBot/dialogs/WelcomeDialog.ts":
/*!***********************************************************!*\
  !*** ./src/app/powerPillagerBot/dialogs/WelcomeDialog.ts ***!
  \***********************************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
const WelcomeCard = __webpack_require__(/*! ./WelcomeCard.json */ "./src/app/powerPillagerBot/dialogs/WelcomeCard.json");
exports.default = WelcomeCard;


/***/ }),

/***/ "./src/app/server.ts":
/*!***************************!*\
  !*** ./src/app/server.ts ***!
  \***************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
const Express = __webpack_require__(/*! express */ "express");
const http = __webpack_require__(/*! http */ "http");
const path = __webpack_require__(/*! path */ "path");
const morgan = __webpack_require__(/*! morgan */ "morgan");
const express_msteams_host_1 = __webpack_require__(/*! express-msteams-host */ "express-msteams-host");
const debug = __webpack_require__(/*! debug */ "debug");
// Initialize debug logging module
const log = debug("msteams");
log(`Initializing Microsoft Teams Express hosted App...`);
// Initialize dotenv, to use .env file settings if existing
// tslint:disable-next-line:no-var-requires
__webpack_require__(/*! dotenv */ "dotenv").config();
// The import of components has to be done AFTER the dotenv config
const allComponents = __webpack_require__(/*! ./TeamsAppsComponents */ "./src/app/TeamsAppsComponents.ts");
// Create the Express webserver
const express = Express();
const port = process.env.port || process.env.PORT || 3007;
// Inject the raw request body onto the request object
express.use(Express.json({
    verify: (req, res, buf, encoding) => {
        req.rawBody = buf.toString();
    }
}));
express.use(Express.urlencoded({ extended: true }));
// Express configuration
express.set("views", path.join(__dirname, "/"));
// Add simple logging
express.use(morgan("tiny"));
// Add /scripts and /assets as static folders
express.use("/scripts", Express.static(path.join(__dirname, "web/scripts")));
express.use("/assets", Express.static(path.join(__dirname, "web/assets")));
// routing for bots, connectors and incoming web hooks - based on the decorators
// For more information see: https://www.npmjs.com/package/express-msteams-host
express.use(express_msteams_host_1.MsTeamsApiRouter(allComponents));
// routing for pages for tabs and connector configuration
// For more information see: https://www.npmjs.com/package/express-msteams-host
express.use(express_msteams_host_1.MsTeamsPageRouter({
    root: path.join(__dirname, "web/"),
    components: allComponents
}));
// Set default web page
express.use("/", Express.static(path.join(__dirname, "web/"), {
    index: "index.html"
}));
// Set the port
express.set("port", port);
// Start the webserver
http.createServer(express).listen(port, () => {
    log(`Server running on ${port}`);
});


/***/ }),

/***/ 0:
/*!*********************************!*\
  !*** multi ./src/app/server.ts ***!
  \*********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

module.exports = __webpack_require__(/*! C:\code\acdc\teams bot/src/app/server.ts */"./src/app/server.ts");


/***/ }),

/***/ "botbuilder":
/*!*****************************!*\
  !*** external "botbuilder" ***!
  \*****************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = require("botbuilder");

/***/ }),

/***/ "botbuilder-dialogs":
/*!*************************************!*\
  !*** external "botbuilder-dialogs" ***!
  \*************************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = require("botbuilder-dialogs");

/***/ }),

/***/ "botbuilder-teams":
/*!***********************************!*\
  !*** external "botbuilder-teams" ***!
  \***********************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = require("botbuilder-teams");

/***/ }),

/***/ "debug":
/*!************************!*\
  !*** external "debug" ***!
  \************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = require("debug");

/***/ }),

/***/ "dotenv":
/*!*************************!*\
  !*** external "dotenv" ***!
  \*************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = require("dotenv");

/***/ }),

/***/ "express":
/*!**************************!*\
  !*** external "express" ***!
  \**************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = require("express");

/***/ }),

/***/ "express-msteams-host":
/*!***************************************!*\
  !*** external "express-msteams-host" ***!
  \***************************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = require("express-msteams-host");

/***/ }),

/***/ "http":
/*!***********************!*\
  !*** external "http" ***!
  \***********************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = require("http");

/***/ }),

/***/ "morgan":
/*!*************************!*\
  !*** external "morgan" ***!
  \*************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = require("morgan");

/***/ }),

/***/ "path":
/*!***********************!*\
  !*** external "path" ***!
  \***********************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = require("path");

/***/ })

/******/ });
//# sourceMappingURL=server.js.map