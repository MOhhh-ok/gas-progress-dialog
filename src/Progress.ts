/**
 * MIT License
 * Copyright (C) 2023 Masaaki Ota
 * https://opensource.org/license/mit/
 */

/****************************************************************
 * Edit below. SpreadsheetApp, DocumentApp, SlidesApp, or FormApp;
 */
const APP = SpreadsheetApp;
// const APP = DocumentApp;
// const APP = SlidesApp;
// const APP = FormApp;

/**
 * Edit above
 ****************************************************************/


/**
 * Show progress info dialog
 * example:
 *   const progress=new Progress({
 *     total: 5,
 *   });
 * 
 *   progress.show();
 * 
 *   for(let i=0;i<5;i++){
 *     progress.log('log '+i);
 *     Utilities.sleep(1000);
 *   }
 * 
 *   progress.fin();
 */


/**
 * Progress data structure
 */
export interface ProgressData {
    total: number,
    progress: number,
    log: string,
    fin: boolean,
}

/**
 * Progress class
 */
export class Progress {
    // instance
    static instance: Progress;

    // property
    Prop: GoogleAppsScript.Properties.Properties;
    PropKey: string;

    // ui
    Ui: GoogleAppsScript.Base.Ui;

    // data structure 
    data: ProgressData;

    // init
    constructor(args: {
        total?: number,
        fromHtml?: boolean,
    }) {

        // singleton
        if (Progress.instance) {
            throw new Error('Progress is singleton. Use Progress.instance');
        }
        Progress.instance = this;

        // property
        this.Prop = PropertiesService.getUserProperties();
        this.PropKey = 'ProgressData';

        // ui
        this.Ui = APP.getUi();

        // from html
        if (args.fromHtml) {
            this.data = this.loadData();
            return;
        }

        // data
        this.data = {
            total: args.total ?? 0,
            progress: 0,
            log: '',
            fin: false,
        };

        this._saveData();
    }

    // load from property
    loadData(): ProgressData {
        const dataStr = this.Prop.getProperty(this.PropKey) ?? '{}';
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
    addTotal(total: number) {
        this.data.total += total;
        this._saveData();
    }

    // set total number
    setTotal(total: number) {
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
    setProgress(now: number) {
        this.data.progress = now;
        this._checkProgress();
        this._saveData();
    }

    // log
    log(str: string) {
        this.data.log += str + '\n';
        this._saveData();
    }

    // show dialog
    show(title: string = 'Now processing. Don\'t close this window...') {

        // if trigger, not show dialog
        if(!this.Ui){
            return;
        }

        this.Ui.showModalDialog(
            HtmlService.createHtmlOutputFromFile('ProgressDlg'),
            title
        );
    }
}


/**
 * global function. call from html
 */
export function progressGetData() {
    const progress = new Progress({ fromHtml: true });
    return progress.loadData();
}

/**
 * test
 */
function progressTest() {
    const progress = new Progress({
        total: 5,
    });

    progress.show('Now processing. Don\'t close this window...');

    for (let i = 0; i < 5; i++) {
        progress.incrementProgress();
        progress.log('log ' + i);
        Utilities.sleep(1000);
    }

    progress.fin();
}
