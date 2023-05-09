"use strict";
/**
 * MIT License
 * Copyright (C) 2023 Masaaki Ota
 * https://opensource.org/license/mit/
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.progressGetData = exports.Progress = void 0;
/**
 * Progress class
 */
class Progress {
    // init
    constructor(args) {
        var _a;
        // singleton
        if (Progress.instance) {
            throw new Error('Progress is singleton. Use Progress.instance');
        }
        Progress.instance = this;
        // property
        this.Prop = PropertiesService.getScriptProperties();
        this.PropKey = 'ProgressData' + new Date();
        // ui
        this.Ui = args.Ui;
        // data
        this.data = {
            total: (_a = args.total) !== null && _a !== void 0 ? _a : 0,
            progress: 0,
            log: '',
            fin: false,
        };
        this._saveData();
    }
    // load from property
    loadData() {
        var _a;
        const dataStr = (_a = this.Prop.getProperty(this.PropKey)) !== null && _a !== void 0 ? _a : '{}';
        this.data = JSON.parse(dataStr);
        return Object.assign({}, this.data);
    }
    // save to property
    _saveData() {
        this.Prop.setProperty(this.PropKey, JSON.stringify(this.data));
    }
    // check progress number
    _checkProgress() {
        if (this.data.total <= this.data.progress) {
            this.fin();
        }
    }
    // finish
    fin() {
        this.data.fin = true;
        this._saveData();
    }
    // add total number
    addTotal(total) {
        this.data.total += total;
        this._saveData();
    }
    // set total number
    setTotal(total) {
        this.data.total = total;
        this._saveData();
    }
    // increment progress number
    incrementProgress() {
        this.data.progress++;
        this._checkProgress();
        this._saveData();
    }
    // set progress number
    setProgress(now) {
        this.data.progress = now;
        this._checkProgress();
        this._saveData();
    }
    // log
    log(str) {
        this.data.log += str + '\n';
        this._saveData();
    }
    // show dialog
    show(title = 'Now processing. Don\'t close this window...') {
        this.Ui.showModalDialog(HtmlService.createHtmlOutput('ProgressDlg.htm'), title);
    }
}
exports.Progress = Progress;
/**
 * global function
 */
function progressGetData() {
    return Progress.instance.loadData();
}
exports.progressGetData = progressGetData;
