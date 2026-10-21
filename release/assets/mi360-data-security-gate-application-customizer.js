define("f0f108c2-8e86-42f5-993d-cb1863c8f6d4_1.0.0.0", ["@microsoft/sp-application-base","@microsoft/sp-core-library","@microsoft/sp-http","react","react-dom"], (__WEBPACK_EXTERNAL_MODULE__841__, __WEBPACK_EXTERNAL_MODULE__676__, __WEBPACK_EXTERNAL_MODULE__909__, __WEBPACK_EXTERNAL_MODULE__959__, __WEBPACK_EXTERNAL_MODULE__398__) => { return /******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ 479
/*!********************************************************************************************!*\
  !*** ./lib/extensions/mi360DataSecurityGate/Mi360DataSecurityGateApplicationCustomizer.js ***!
  \********************************************************************************************/
(__unused_webpack_module, exports, __webpack_require__) {


var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", ({ value: true }));
var sp_application_base_1 = __webpack_require__(/*! @microsoft/sp-application-base */ 841);
var sp_core_library_1 = __webpack_require__(/*! @microsoft/sp-core-library */ 676);
var sp_http_1 = __webpack_require__(/*! @microsoft/sp-http */ 909);
var React = __webpack_require__(/*! react */ 959);
var ReactDOM = __webpack_require__(/*! react-dom */ 398);
var AcknowledgementPage_1 = __webpack_require__(/*! ./components/AcknowledgementPage */ 870);
var OwnerGroupService_1 = __webpack_require__(/*! ./services/OwnerGroupService */ 700);
var EntraIdentityService_1 = __webpack_require__(/*! ./services/EntraIdentityService */ 984);
var AcknowledgementListService_1 = __webpack_require__(/*! ./services/AcknowledgementListService */ 958);
var SiteScopeService_1 = __webpack_require__(/*! ./services/SiteScopeService */ 861);
var ReturnUrlValidator_1 = __webpack_require__(/*! ./utils/ReturnUrlValidator */ 643);
var LOG_SOURCE = 'Mi360DataSecurityGateApplicationCustomizer';
var DEFAULT_SITE_ROOT = '/sites/MI360';
var DEFAULT_LIST_TITLE = 'MI360 Data Security Acknowledgements';
var ACKNOWLEDGEMENT_SOURCE = 'MI360-SPFx-ApplicationCustomizer';
var CONTAINER_ID = 'mi360DataSecurityGateContainer';
var Mi360DataSecurityGateApplicationCustomizer = /** @class */ (function (_super) {
    __extends(Mi360DataSecurityGateApplicationCustomizer, _super);
    function Mi360DataSecurityGateApplicationCustomizer() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._hostElement = null;
        _this._isEvaluating = false;
        _this._pendingReEvaluation = false;
        /** Non-authoritative, in-memory only. Never persisted to storage; cleared on full reload. */
        _this._acknowledgedObjectIds = new Set();
        _this._onNavigated = function () {
            _this._runGate().catch(function (error) {
                sp_core_library_1.Log.error(LOG_SOURCE, error instanceof Error ? error : new Error(String(error)));
            });
        };
        return _this;
    }
    Mi360DataSecurityGateApplicationCustomizer.prototype.onInit = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.context.application.navigatedEvent.add(this, this._onNavigated);
                        return [4 /*yield*/, this._runGate()];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    Mi360DataSecurityGateApplicationCustomizer.prototype.onDispose = function () {
        this.context.application.navigatedEvent.remove(this, this._onNavigated);
        this._unmountAcknowledgementUi();
    };
    Object.defineProperty(Mi360DataSecurityGateApplicationCustomizer.prototype, "siteRootPath", {
        get: function () {
            return this.properties.siteRelativeRootPath || DEFAULT_SITE_ROOT;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Mi360DataSecurityGateApplicationCustomizer.prototype, "acknowledgementPageRelativeUrl", {
        get: function () {
            return (this.properties.acknowledgementPageRelativeUrl || "".concat(this.siteRootPath, "/SitePages/MI360-Data-Security-Acknowledgement.aspx")).toLowerCase();
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Mi360DataSecurityGateApplicationCustomizer.prototype, "webUrl", {
        get: function () {
            return this.context.pageContext.web.absoluteUrl;
        },
        enumerable: false,
        configurable: true
    });
    Mi360DataSecurityGateApplicationCustomizer.prototype._runGate = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        // Coalesce overlapping navigation events into a single follow-up run rather than
                        // allowing concurrent evaluations to race against each other.
                        if (this._isEvaluating) {
                            this._pendingReEvaluation = true;
                            return [2 /*return*/];
                        }
                        this._isEvaluating = true;
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, , 3, 6]);
                        return [4 /*yield*/, this._evaluateGate()];
                    case 2:
                        _a.sent();
                        return [3 /*break*/, 6];
                    case 3:
                        this._isEvaluating = false;
                        if (!this._pendingReEvaluation) return [3 /*break*/, 5];
                        this._pendingReEvaluation = false;
                        return [4 /*yield*/, this._runGate()];
                    case 4:
                        _a.sent();
                        _a.label = 5;
                    case 5: return [7 /*endfinally*/];
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    Mi360DataSecurityGateApplicationCustomizer.prototype._evaluateGate = function () {
        return __awaiter(this, void 0, void 0, function () {
            var currentUrl, ownerGroupService, isOwner, identityService, objectId, listService, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 4, , 5]);
                        currentUrl = window.location.href;
                        if (!(0, SiteScopeService_1.isWithinSite)(currentUrl, this.siteRootPath)) {
                            return [2 /*return*/];
                        }
                        if ((0, SiteScopeService_1.isTechnicalExclusion)(currentUrl, this.siteRootPath)) {
                            return [2 /*return*/];
                        }
                        if ((0, SiteScopeService_1.isAcknowledgementPage)(currentUrl, this.acknowledgementPageRelativeUrl)) {
                            this._renderAcknowledgementUi();
                            return [2 /*return*/];
                        }
                        this._unmountAcknowledgementUi();
                        ownerGroupService = new OwnerGroupService_1.OwnerGroupService(this.context.spHttpClient, sp_http_1.SPHttpClient.configurations.v1, this.webUrl);
                        return [4 /*yield*/, ownerGroupService.isCurrentUserOwner(this.properties.ownersGroupName || 'MI360 Owners')];
                    case 1:
                        isOwner = _a.sent();
                        if (isOwner) {
                            return [2 /*return*/];
                        }
                        identityService = new EntraIdentityService_1.EntraIdentityService(this.context.spHttpClient, sp_http_1.SPHttpClient.configurations.v1, this.webUrl);
                        return [4 /*yield*/, identityService.getCurrentEntraObjectId()];
                    case 2:
                        objectId = _a.sent();
                        if (!objectId) {
                            this._redirectToAcknowledgementPage(this.siteRootPath);
                            return [2 /*return*/];
                        }
                        if (this._acknowledgedObjectIds.has(objectId)) {
                            return [2 /*return*/];
                        }
                        listService = new AcknowledgementListService_1.AcknowledgementListService(this.context.spHttpClient, sp_http_1.SPHttpClient.configurations.v1, this.webUrl, this.properties.acknowledgementListTitle || DEFAULT_LIST_TITLE);
                        return [4 /*yield*/, listService.hasActiveAcknowledgement(objectId)];
                    case 3:
                        if (_a.sent()) {
                            this._acknowledgedObjectIds.add(objectId);
                            return [2 /*return*/];
                        }
                        this._redirectToAcknowledgementPage(currentUrl);
                        return [3 /*break*/, 5];
                    case 4:
                        error_1 = _a.sent();
                        // Fail closed: any failure while evaluating owner status or acknowledgement status
                        // keeps the user in the acknowledgement experience instead of allowing access.
                        sp_core_library_1.Log.error(LOG_SOURCE, error_1 instanceof Error ? error_1 : new Error('Unknown gate failure'));
                        this._renderAcknowledgementUi();
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    Mi360DataSecurityGateApplicationCustomizer.prototype._redirectToAcknowledgementPage = function (returnUrlCandidate) {
        var validatedReturnUrl = (0, ReturnUrlValidator_1.validateReturnUrl)(returnUrlCandidate, this.siteRootPath, window.location.origin);
        var target = new URL(this.acknowledgementPageRelativeUrl, window.location.origin);
        target.searchParams.set('returnUrl', validatedReturnUrl);
        window.location.assign(target.toString());
    };
    Mi360DataSecurityGateApplicationCustomizer.prototype._unmountAcknowledgementUi = function () {
        if (this._hostElement) {
            ReactDOM.unmountComponentAtNode(this._hostElement);
            this._hostElement.remove();
            this._hostElement = null;
        }
    };
    Mi360DataSecurityGateApplicationCustomizer.prototype._renderAcknowledgementUi = function () {
        var _this = this;
        var currentReturn = (0, ReturnUrlValidator_1.validateReturnUrl)(new URL(window.location.href).searchParams.get('returnUrl') || this.siteRootPath, this.siteRootPath, window.location.origin);
        this._unmountAcknowledgementUi();
        var container = document.createElement('div');
        container.id = CONTAINER_ID;
        container.style.position = 'relative';
        container.style.zIndex = '1000';
        (document.body || document.documentElement).insertBefore(container, (document.body || document.documentElement).firstChild);
        this._hostElement = container;
        ReactDOM.render(React.createElement(AcknowledgementPage_1.AcknowledgementPage, { returnUrl: currentReturn, policyUrl: this.properties.policyUrl || 'https://www.zf.com/corporate/en/our-company/data-protection.html', supportUrl: this.properties.supportUrl || "".concat(this.siteRootPath, "/SitePages/MI360-Support.aspx"), onAcknowledge: function (redirectTarget) { return __awaiter(_this, void 0, void 0, function () {
                var identityService, identity, listService, destination;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            identityService = new EntraIdentityService_1.EntraIdentityService(this.context.spHttpClient, sp_http_1.SPHttpClient.configurations.v1, this.webUrl);
                            return [4 /*yield*/, identityService.getCurrentIdentity()];
                        case 1:
                            identity = _a.sent();
                            listService = new AcknowledgementListService_1.AcknowledgementListService(this.context.spHttpClient, sp_http_1.SPHttpClient.configurations.v1, this.webUrl, this.properties.acknowledgementListTitle || DEFAULT_LIST_TITLE);
                            return [4 /*yield*/, listService.ensureAcknowledgement({
                                    identity: identity,
                                    policyTitle: this.properties.policyTitle || 'ZF Data Security Policy',
                                    source: ACKNOWLEDGEMENT_SOURCE,
                                    requestedPage: redirectTarget || this.siteRootPath
                                })];
                        case 2:
                            _a.sent();
                            this._acknowledgedObjectIds.add(identity.entraObjectId);
                            destination = (0, ReturnUrlValidator_1.validateReturnUrl)(redirectTarget, this.siteRootPath, window.location.origin);
                            window.location.assign(destination);
                            return [2 /*return*/];
                    }
                });
            }); } }), container);
    };
    return Mi360DataSecurityGateApplicationCustomizer;
}(sp_application_base_1.BaseApplicationCustomizer));
exports["default"] = Mi360DataSecurityGateApplicationCustomizer;


/***/ },

/***/ 870
/*!********************************************************************************!*\
  !*** ./lib/extensions/mi360DataSecurityGate/components/AcknowledgementPage.js ***!
  \********************************************************************************/
(__unused_webpack_module, exports, __webpack_require__) {


var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.AcknowledgementPage = void 0;
var React = __webpack_require__(/*! react */ 959);
var AcknowledgementPage = function (_a) {
    var returnUrl = _a.returnUrl, policyUrl = _a.policyUrl, supportUrl = _a.supportUrl, onAcknowledge = _a.onAcknowledge;
    var _b = React.useState(false), isChecked = _b[0], setIsChecked = _b[1];
    var _c = React.useState(false), isSubmitting = _c[0], setIsSubmitting = _c[1];
    var _d = React.useState(null), error = _d[0], setError = _d[1];
    var headingRef = React.useRef(null);
    var errorRef = React.useRef(null);
    React.useEffect(function () {
        var _a;
        (_a = headingRef.current) === null || _a === void 0 ? void 0 : _a.focus();
    }, []);
    React.useEffect(function () {
        var _a;
        if (error) {
            (_a = errorRef.current) === null || _a === void 0 ? void 0 : _a.focus();
        }
    }, [error]);
    var handleSubmit = function () { return __awaiter(void 0, void 0, void 0, function () {
        var e_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!isChecked || isSubmitting) {
                        return [2 /*return*/];
                    }
                    setIsSubmitting(true);
                    setError(null);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, onAcknowledge(returnUrl)];
                case 2:
                    _a.sent();
                    return [3 /*break*/, 4];
                case 3:
                    e_1 = _a.sent();
                    setError('We could not record your acknowledgement. Please try again.');
                    setIsSubmitting(false);
                    return [3 /*break*/, 4];
                case 4: return [2 /*return*/];
            }
        });
    }); };
    var errorMessageId = 'mi360AcknowledgementError';
    return (React.createElement("div", { role: "main", "aria-live": "polite", style: {
            maxWidth: 760,
            margin: '32px auto',
            padding: 24,
            background: '#ffffff',
            borderRadius: 8,
            border: '1px solid #e1e1e1',
            boxShadow: '0 2px 8px rgba(0,0,0,0.08)'
        } },
        React.createElement("h1", { ref: headingRef, tabIndex: -1, style: { marginTop: 0, fontSize: 28, outline: 'none' } }, "MI360 Data Security Acknowledgement"),
        React.createElement("div", { style: { marginTop: 16 } },
            React.createElement("p", null, "To continue using MI360, you must acknowledge the MI360 Data Security requirements and agree to comply with the applicable ZF Data Security policy.")),
        React.createElement("div", { style: { marginTop: 16 } },
            React.createElement("p", null,
                "Policy reference:",
                ' ',
                React.createElement("a", { href: policyUrl, target: "_blank", rel: "noopener noreferrer" }, "ZF Data Security Policy"))),
        React.createElement("div", { style: { marginTop: 16 } },
            React.createElement("p", null, "Selecting the checkbox confirms that you agree to comply with the policy. It does not prove that you read or understood every detail of the full policy.")),
        error && (React.createElement("div", { ref: errorRef, id: errorMessageId, role: "alert", tabIndex: -1, style: { marginTop: 16, padding: 12, color: '#a4262c', background: '#fde7e9', outline: 'none' } }, error)),
        React.createElement("label", { style: { display: 'flex', gap: 8, marginTop: 20, alignItems: 'flex-start' } },
            React.createElement("input", { type: "checkbox", checked: isChecked, disabled: isSubmitting, "aria-describedby": error ? errorMessageId : undefined, onChange: function (event) { return setIsChecked(event.currentTarget.checked); } }),
            React.createElement("span", null, "I acknowledge the MI360 Data Security requirements and agree to comply with the referenced ZF Data Security policy.")),
        React.createElement("div", { style: { marginTop: 24 } },
            React.createElement("button", { type: "button", disabled: !isChecked || isSubmitting, "aria-describedby": error ? errorMessageId : undefined, onClick: handleSubmit, style: { padding: '10px 18px', cursor: isChecked && !isSubmitting ? 'pointer' : 'not-allowed' } }, isSubmitting ? 'Submitting...' : error ? 'Try again' : 'I acknowledge')),
        React.createElement("div", { style: { marginTop: 24 } },
            React.createElement("p", null,
                "If you are having trouble, contact the",
                ' ',
                React.createElement("a", { href: supportUrl }, "MI360 support route"),
                "."))));
};
exports.AcknowledgementPage = AcknowledgementPage;


/***/ },

/***/ 958
/*!*************************************************************************************!*\
  !*** ./lib/extensions/mi360DataSecurityGate/services/AcknowledgementListService.js ***!
  \*************************************************************************************/
(__unused_webpack_module, exports) {


var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.AcknowledgementListService = exports.buildAcknowledgementPayload = void 0;
/**
 * Pure payload builder, separated from the network call so it can be unit tested without a
 * SharePoint context. Every value originates from the authenticated identity or fixed
 * configuration, never from editable form controls.
 */
function buildAcknowledgementPayload(request, entityTypeFullName) {
    return {
        __metadata: { type: entityTypeFullName },
        Title: request.identity.displayName,
        EntraObjectId: request.identity.entraObjectId,
        UserDisplayName: request.identity.displayName,
        UserPrincipalName: request.identity.userPrincipalName,
        AcknowledgedAtUtc: new Date().toISOString(),
        AcknowledgementSource: request.source,
        PolicyTitle: request.policyTitle,
        RequestedPage: request.requestedPage,
        RecordStatus: 'Active'
    };
}
exports.buildAcknowledgementPayload = buildAcknowledgementPayload;
function escapeODataString(value) {
    return value.replace(/'/g, "''");
}
/**
 * Reads and writes the MI360 Data Security Acknowledgements list. Duplicate submissions,
 * repeated selections from multiple tabs, and concurrent requests are treated as a single
 * successful acknowledgement: an existing active record for the same Entra Object ID is never
 * duplicated, and a write conflict is resolved by re-checking for an existing record.
 */
var AcknowledgementListService = /** @class */ (function () {
    function AcknowledgementListService(spHttpClient, httpClientConfiguration, webUrl, listTitle) {
        this.spHttpClient = spHttpClient;
        this.httpClientConfiguration = httpClientConfiguration;
        this.webUrl = webUrl;
        this.listTitle = listTitle;
    }
    AcknowledgementListService.prototype.hasActiveAcknowledgement = function (entraObjectId) {
        return __awaiter(this, void 0, void 0, function () {
            var listTitle, filterObjectId, requestUrl, response, json;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        listTitle = escapeODataString(this.listTitle);
                        filterObjectId = escapeODataString(entraObjectId);
                        requestUrl = "".concat(this.webUrl, "/_api/web/lists/getByTitle('").concat(listTitle, "')/items?$filter=EntraObjectId eq '").concat(filterObjectId, "' and RecordStatus eq 'Active'&$select=Id&$top=1");
                        return [4 /*yield*/, this.spHttpClient.get(requestUrl, this.httpClientConfiguration)];
                    case 1:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("Unable to query acknowledgement list status (status ".concat(response.status, ")."));
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        json = _a.sent();
                        return [2 /*return*/, Array.isArray(json.value) && json.value.length > 0];
                }
            });
        });
    };
    /**
     * Creates the acknowledgement record if one does not already exist. Returns without error
     * whether the record was newly created or already existed, satisfying the "treat an existing
     * record as a successful acknowledgement" requirement.
     */
    AcknowledgementListService.prototype.ensureAcknowledgement = function (request) {
        return __awaiter(this, void 0, void 0, function () {
            var entityTypeFullName, payload, listTitle, createResponse;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.hasActiveAcknowledgement(request.identity.entraObjectId)];
                    case 1:
                        if (_a.sent()) {
                            return [2 /*return*/];
                        }
                        return [4 /*yield*/, this.getEntityTypeFullName()];
                    case 2:
                        entityTypeFullName = _a.sent();
                        payload = buildAcknowledgementPayload(request, entityTypeFullName);
                        listTitle = escapeODataString(this.listTitle);
                        return [4 /*yield*/, this.spHttpClient.post("".concat(this.webUrl, "/_api/web/lists/getByTitle('").concat(listTitle, "')/items"), this.httpClientConfiguration, { body: JSON.stringify(payload) })];
                    case 3:
                        createResponse = _a.sent();
                        if (createResponse.ok) {
                            return [2 /*return*/];
                        }
                        return [4 /*yield*/, this.hasActiveAcknowledgement(request.identity.entraObjectId)];
                    case 4:
                        // A conflict, or a transient failure that still resulted in a write, is only acceptable if
                        // an active record now exists; otherwise this is a genuine failure and must fail closed.
                        if (_a.sent()) {
                            return [2 /*return*/];
                        }
                        throw new Error("Acknowledgement save failed (status ".concat(createResponse.status, ")."));
                }
            });
        });
    };
    AcknowledgementListService.prototype.getEntityTypeFullName = function () {
        return __awaiter(this, void 0, void 0, function () {
            var listTitle, requestUrl, response, json;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (this.entityTypeFullName) {
                            return [2 /*return*/, this.entityTypeFullName];
                        }
                        listTitle = escapeODataString(this.listTitle);
                        requestUrl = "".concat(this.webUrl, "/_api/web/lists/getByTitle('").concat(listTitle, "')?$select=ListItemEntityTypeFullName");
                        return [4 /*yield*/, this.spHttpClient.get(requestUrl, this.httpClientConfiguration)];
                    case 1:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("Unable to resolve acknowledgement list schema (status ".concat(response.status, ")."));
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        json = _a.sent();
                        if (!json.ListItemEntityTypeFullName) {
                            throw new Error('Acknowledgement list did not return an entity type name.');
                        }
                        this.entityTypeFullName = json.ListItemEntityTypeFullName;
                        return [2 /*return*/, this.entityTypeFullName];
                }
            });
        });
    };
    return AcknowledgementListService;
}());
exports.AcknowledgementListService = AcknowledgementListService;


/***/ },

/***/ 984
/*!*******************************************************************************!*\
  !*** ./lib/extensions/mi360DataSecurityGate/services/EntraIdentityService.js ***!
  \*******************************************************************************/
(__unused_webpack_module, exports) {


var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.EntraIdentityService = void 0;
/**
 * Retrieves the current user's permanent Entra Object ID and profile details from the
 * authenticated SharePoint context. Values are never accepted from editable form controls.
 */
var EntraIdentityService = /** @class */ (function () {
    function EntraIdentityService(spHttpClient, httpClientConfiguration, webUrl) {
        this.spHttpClient = spHttpClient;
        this.httpClientConfiguration = httpClientConfiguration;
        this.webUrl = webUrl;
    }
    EntraIdentityService.prototype.getCurrentEntraObjectId = function () {
        return __awaiter(this, void 0, void 0, function () {
            var requestUrl, response, properties, entries, entry;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        requestUrl = "".concat(this.webUrl, "/_api/SP.UserProfiles.PeopleManager/GetMyProperties?$select=UserProfileProperties");
                        return [4 /*yield*/, this.spHttpClient.get(requestUrl, this.httpClientConfiguration)];
                    case 1:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("Unable to read user profile properties (status ".concat(response.status, ")."));
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        properties = _a.sent();
                        entries = Array.isArray(properties.UserProfileProperties) ? properties.UserProfileProperties : [];
                        entry = entries.find(function (item) { return item.Key === 'SPS-ObjectId' || item.Key === 'SPS-ObjectId-Encrypted'; });
                        return [2 /*return*/, (entry === null || entry === void 0 ? void 0 : entry.Value) ? String(entry.Value).trim() : null];
                }
            });
        });
    };
    EntraIdentityService.prototype.getCurrentIdentity = function () {
        return __awaiter(this, void 0, void 0, function () {
            var objectId, requestUrl, response, currentUser, displayName, userPrincipalName;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.getCurrentEntraObjectId()];
                    case 1:
                        objectId = _a.sent();
                        if (!objectId) {
                            throw new Error('Authenticated user is missing an Entra Object ID.');
                        }
                        requestUrl = "".concat(this.webUrl, "/_api/web/currentuser?$select=Title,Email,LoginName");
                        return [4 /*yield*/, this.spHttpClient.get(requestUrl, this.httpClientConfiguration)];
                    case 2:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("Unable to determine current user details (status ".concat(response.status, ")."));
                        }
                        return [4 /*yield*/, response.json()];
                    case 3:
                        currentUser = _a.sent();
                        displayName = (currentUser.Title || currentUser.LoginName || '').trim();
                        userPrincipalName = (currentUser.Email || currentUser.LoginName || '').trim();
                        if (!displayName || !userPrincipalName) {
                            throw new Error('Authenticated user context did not return a display name or principal name.');
                        }
                        return [2 /*return*/, { entraObjectId: objectId, displayName: displayName, userPrincipalName: userPrincipalName }];
                }
            });
        });
    };
    return EntraIdentityService;
}());
exports.EntraIdentityService = EntraIdentityService;


/***/ },

/***/ 700
/*!****************************************************************************!*\
  !*** ./lib/extensions/mi360DataSecurityGate/services/OwnerGroupService.js ***!
  \****************************************************************************/
(__unused_webpack_module, exports) {


var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.OwnerGroupService = exports.isMemberOfGroup = void 0;
/**
 * Pure membership check, separated from the network call so it can be unit tested
 * without a SharePoint context.
 */
function isMemberOfGroup(groups, groupName) {
    var normalizedTarget = groupName.trim().toLowerCase();
    return groups.some(function (group) { return (group.Title || '').trim().toLowerCase() === normalizedTarget; });
}
exports.isMemberOfGroup = isMemberOfGroup;
/**
 * Determines whether the current signed-in user belongs to the configured owners group.
 * Fails closed: any non-OK response or network failure throws, so the caller treats the
 * user as a non-owner and continues through the acknowledgement gate rather than bypassing it.
 */
var OwnerGroupService = /** @class */ (function () {
    function OwnerGroupService(spHttpClient, httpClientConfiguration, webUrl) {
        this.spHttpClient = spHttpClient;
        this.httpClientConfiguration = httpClientConfiguration;
        this.webUrl = webUrl;
    }
    OwnerGroupService.prototype.isCurrentUserOwner = function (groupName) {
        return __awaiter(this, void 0, void 0, function () {
            var requestUrl, response, json, groups;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        requestUrl = "".concat(this.webUrl, "/_api/web/currentuser/groups?$select=Id,Title");
                        return [4 /*yield*/, this.spHttpClient.get(requestUrl, this.httpClientConfiguration)];
                    case 1:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("Unable to resolve current user group membership (status ".concat(response.status, ")."));
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        json = _a.sent();
                        groups = Array.isArray(json.value) ? json.value : [];
                        return [2 /*return*/, isMemberOfGroup(groups, groupName)];
                }
            });
        });
    };
    return OwnerGroupService;
}());
exports.OwnerGroupService = OwnerGroupService;


/***/ },

/***/ 861
/*!***************************************************************************!*\
  !*** ./lib/extensions/mi360DataSecurityGate/services/SiteScopeService.js ***!
  \***************************************************************************/
(__unused_webpack_module, exports) {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.isAcknowledgementPage = exports.isTechnicalExclusion = exports.isWithinSite = void 0;
var TECHNICAL_EXCLUSION_SUFFIXES = [
    '/_layouts/15/validation.aspx',
    '/_layouts/15/error.aspx',
    '/_layouts/15/authenticate.aspx',
    '/_layouts/15/signin.aspx',
    '/_layouts/15/settings.aspx'
];
function getPathname(url) {
    try {
        return new URL(url).pathname.toLowerCase();
    }
    catch (_a) {
        return null;
    }
}
/**
 * Determines whether the given URL falls within the gated site collection.
 */
function isWithinSite(url, siteRootPath) {
    var pathname = getPathname(url);
    if (pathname === null) {
        return false;
    }
    var normalizedRoot = siteRootPath.toLowerCase();
    return pathname === normalizedRoot || pathname.startsWith("".concat(normalizedRoot, "/"));
}
exports.isWithinSite = isWithinSite;
/**
 * Narrow, explicit set of technical pages required to avoid redirect loops during
 * authentication, error handling, and validation. This must never be widened to a broad
 * pattern that could allow normal MI360 content to bypass the gate.
 */
function isTechnicalExclusion(url, siteRootPath) {
    var pathname = getPathname(url);
    if (pathname === null) {
        return false;
    }
    var normalizedRoot = siteRootPath.toLowerCase();
    return TECHNICAL_EXCLUSION_SUFFIXES.some(function (suffix) { return pathname === "".concat(normalizedRoot).concat(suffix); });
}
exports.isTechnicalExclusion = isTechnicalExclusion;
function isAcknowledgementPage(url, acknowledgementPageRelativeUrl) {
    var pathname = getPathname(url);
    if (pathname === null) {
        return false;
    }
    return pathname === acknowledgementPageRelativeUrl.toLowerCase();
}
exports.isAcknowledgementPage = isAcknowledgementPage;


/***/ },

/***/ 643
/*!**************************************************************************!*\
  !*** ./lib/extensions/mi360DataSecurityGate/utils/ReturnUrlValidator.js ***!
  \**************************************************************************/
(__unused_webpack_module, exports) {


/**
 * Validates that a candidate return URL is a safe, site-relative path within the gated site
 * collection. Used to prevent open-redirect attacks (external hosts, protocol-relative URLs,
 * `javascript:`/`data:` schemes, encoded payloads, and path traversal out of the site root).
 */
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.validateReturnUrl = void 0;
var CONTROL_CHARACTER_PATTERN = /[\u0000-\u001f\u007f]/;
var SCHEME_PATTERN = /^[a-zA-Z][a-zA-Z0-9+.-]*:/;
var MAX_DECODE_ATTEMPTS = 3;
function safeDecode(candidate) {
    var current = candidate;
    for (var attempt = 0; attempt < MAX_DECODE_ATTEMPTS; attempt++) {
        var next = void 0;
        try {
            next = decodeURIComponent(current);
        }
        catch (_a) {
            return null;
        }
        if (next === current) {
            return current;
        }
        current = next;
    }
    return current;
}
/**
 * Returns a validated, site-relative return URL, or the site root if the candidate is missing,
 * malformed, or points outside the gated site collection.
 *
 * @param candidate - The raw, untrusted return URL supplied by the browser/query string.
 * @param siteRootPath - The site-relative root, e.g. `/sites/MI360` (no trailing slash).
 * @param origin - The current window origin, used to resolve and re-validate the final path.
 */
function validateReturnUrl(candidate, siteRootPath, origin) {
    var normalizedRoot = siteRootPath.replace(/\/+$/, '').toLowerCase();
    var fallback = siteRootPath;
    if (!candidate || !candidate.trim()) {
        return fallback;
    }
    var decoded = safeDecode(candidate.trim());
    if (decoded === null) {
        return fallback;
    }
    var normalizedSlashes = decoded.replace(/\\/g, '/');
    if (CONTROL_CHARACTER_PATTERN.test(normalizedSlashes)) {
        return fallback;
    }
    if (SCHEME_PATTERN.test(normalizedSlashes)) {
        return fallback;
    }
    if (!normalizedSlashes.startsWith('/') || normalizedSlashes.startsWith('//')) {
        return fallback;
    }
    var lowerPath = normalizedSlashes.toLowerCase();
    if (lowerPath !== normalizedRoot && !lowerPath.startsWith("".concat(normalizedRoot, "/"))) {
        return fallback;
    }
    try {
        var resolved = new URL(normalizedSlashes, origin);
        if (resolved.origin !== origin) {
            return fallback;
        }
        var resolvedLowerPath = resolved.pathname.toLowerCase();
        if (resolvedLowerPath !== normalizedRoot && !resolvedLowerPath.startsWith("".concat(normalizedRoot, "/"))) {
            return fallback;
        }
        return "".concat(resolved.pathname).concat(resolved.search).concat(resolved.hash);
    }
    catch (_a) {
        return fallback;
    }
}
exports.validateReturnUrl = validateReturnUrl;


/***/ },

/***/ 841
/*!*************************************************!*\
  !*** external "@microsoft/sp-application-base" ***!
  \*************************************************/
(module) {

module.exports = __WEBPACK_EXTERNAL_MODULE__841__;

/***/ },

/***/ 676
/*!*********************************************!*\
  !*** external "@microsoft/sp-core-library" ***!
  \*********************************************/
(module) {

module.exports = __WEBPACK_EXTERNAL_MODULE__676__;

/***/ },

/***/ 909
/*!*************************************!*\
  !*** external "@microsoft/sp-http" ***!
  \*************************************/
(module) {

module.exports = __WEBPACK_EXTERNAL_MODULE__909__;

/***/ },

/***/ 959
/*!************************!*\
  !*** external "react" ***!
  \************************/
(module) {

module.exports = __WEBPACK_EXTERNAL_MODULE__959__;

/***/ },

/***/ 398
/*!****************************!*\
  !*** external "react-dom" ***!
  \****************************/
(module) {

module.exports = __WEBPACK_EXTERNAL_MODULE__398__;

/***/ }

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		if (!(moduleId in __webpack_modules__)) {
/******/ 			delete __webpack_module_cache__[moduleId];
/******/ 			var e = new Error("Cannot find module '" + moduleId + "'");
/******/ 			e.code = 'MODULE_NOT_FOUND';
/******/ 			throw e;
/******/ 		}
/******/ 		__webpack_modules__[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/************************************************************************/
/******/ 	
/******/ 	// startup
/******/ 	// Load entry module and return exports
/******/ 	// This entry module is referenced by other modules so it can't be inlined
/******/ 	var __webpack_exports__ = __webpack_require__(479);
/******/ 	
/******/ 	return __webpack_exports__;
/******/ })()
;
});;
//# sourceMappingURL=mi360-data-security-gate-application-customizer.js.map