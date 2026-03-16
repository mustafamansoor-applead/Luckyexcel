import { LuckyFile } from "./ToLuckySheet/LuckyFile";
// import {SecurityDoor,Car} from './content';

import { HandleZip } from './HandleZip';

import { IuploadfileList } from "./ICommon";

import { WorkBook } from "./UniverToExcel/Workbook";
import exceljs from "@zwight/exceljs";
import { CSV } from "./UniverToCsv/CSV";
import { isObject } from "./common/method";
import { UniverWorkBook } from "./LuckyToUniver/UniverWorkBook";
import { IWorkbookData } from "@univerjs/core";
import { formatSheetData, getDataByFile } from "./common/utils";
import { UniverCsvWorkBook } from "./LuckyToUniver/UniverCsvWorkBook";
import { ILuckyFile } from "./ToLuckySheet/ILuck";
export class LuckyExcel {
    constructor() { }
    static transformExcelToLucky(excelFile: File,
        callback?: (files: ILuckyFile, fs?: string) => void,
        errorHandler?: (err: Error) => void) {
        let handleZip: HandleZip = new HandleZip(excelFile);
        const startedAt = Date.now();
        LuckyExcel.logStep(
            'XLSX -> Lucky import started',
            `${excelFile.name} (${LuckyExcel.formatFileSize(excelFile.size)})`
        );

        // const fileReader = new FileReader();
        // fileReader.onload = async (e) => {
        //     const { result } = e.target as any;
        //     const workbook = new exceljs.Workbook();
        //     const data = await workbook.xlsx.load(result);
        //     // console.log('exceljs', data)
        // }
        // fileReader.readAsArrayBuffer(excelFile)

        handleZip.unzipFile((files: IuploadfileList) => {
            try {
                LuckyExcel.logStep(
                    'ZIP extraction completed',
                    `${Object.keys(files).length} entries in ${Date.now() - startedAt}ms`
                );
                const parseStartedAt = Date.now();
                let luckyFile = new LuckyFile(files, excelFile.name);
                let exportJson = luckyFile.ParseObject();
                LuckyExcel.logStep(
                    'Lucky workbook parsed',
                    `${exportJson?.sheets?.length || 0} sheets in ${Date.now() - parseStartedAt}ms`
                );
                if (callback != undefined) {
                    const luckysheetfile = callback.length > 1 ? JSON.stringify(exportJson) : undefined;
                    callback(exportJson, luckysheetfile);
                }
                LuckyExcel.logStep(
                    'XLSX -> Lucky import finished',
                    `${Date.now() - startedAt}ms total`
                );
            } catch (err) {
                LuckyExcel.logError(
                    `XLSX -> Lucky import failed after ${Date.now() - startedAt}ms`,
                    err
                );
                if (errorHandler) {
                    errorHandler(err as Error);
                } else {
                    console.error(err);
                }
            }
        },
            function (err: Error) {
                LuckyExcel.logError(
                    `ZIP extraction failed after ${Date.now() - startedAt}ms`,
                    err
                );
                if (errorHandler) {
                    errorHandler(err);
                } else {
                    console.error(err);
                }
            });
    }

    static transformExcelToLuckyByUrl(
        url: string,
        name: string,
        callBack?: (files: ILuckyFile, fs?: string) => void,
        errorHandler?: (err: Error) => void) {
        let handleZip: HandleZip = new HandleZip();
        const startedAt = Date.now();
        LuckyExcel.logStep('XLSX URL -> Lucky import started', name || url);
        handleZip.unzipFileByUrl(url, (files: IuploadfileList) => {
            try {
                LuckyExcel.logStep(
                    'ZIP extraction completed',
                    `${Object.keys(files).length} entries in ${Date.now() - startedAt}ms`
                );
                const parseStartedAt = Date.now();
                let luckyFile = new LuckyFile(files, name);
                let exportJson = luckyFile.ParseObject();
                LuckyExcel.logStep(
                    'Lucky workbook parsed',
                    `${exportJson?.sheets?.length || 0} sheets in ${Date.now() - parseStartedAt}ms`
                );
                if (callBack != undefined) {
                    const luckysheetfile = callBack.length > 1 ? JSON.stringify(exportJson) : undefined;
                    callBack(exportJson, luckysheetfile);
                }
                LuckyExcel.logStep(
                    'XLSX URL -> Lucky import finished',
                    `${Date.now() - startedAt}ms total`
                );
            } catch (err) {
                LuckyExcel.logError(
                    `XLSX URL -> Lucky import failed after ${Date.now() - startedAt}ms`,
                    err
                );
                if (errorHandler) {
                    errorHandler(err as Error);
                } else {
                    console.error(err);
                }
            }
        },
            function (err: Error) {
                LuckyExcel.logError(
                    `ZIP extraction failed after ${Date.now() - startedAt}ms`,
                    err
                );
                if (errorHandler) {
                    errorHandler(err);
                } else {
                    console.error(err);
                }
            });
    }


    static transformExcelToUniver(
        excelFile: File,
        callback?: (files: IWorkbookData, fs?: string) => void,
        errorHandler?: (err: Error) => void
    ) {
        let handleZip: HandleZip = new HandleZip(excelFile);
        const startedAt = Date.now();
        LuckyExcel.logStep(
            'XLSX -> Univer import started',
            `${excelFile.name} (${LuckyExcel.formatFileSize(excelFile.size)})`
        );

        handleZip.unzipFile((files: IuploadfileList) => {
            try {
                LuckyExcel.logStep(
                    'ZIP extraction completed',
                    `${Object.keys(files).length} entries in ${Date.now() - startedAt}ms`
                );
                const parseStartedAt = Date.now();
                let luckyFile = new LuckyFile(files, excelFile.name);
                let exportJson = luckyFile.ParseObject();
                LuckyExcel.logStep(
                    'Lucky workbook parsed',
                    `${exportJson?.sheets?.length || 0} sheets in ${Date.now() - parseStartedAt}ms`
                );
                if (callback != undefined) {
                    const convertStartedAt = Date.now();
                    const univerData = new UniverWorkBook(exportJson);
                    LuckyExcel.logStep(
                        'Univer snapshot built',
                        `${univerData.sheetOrder?.length || 0} sheets in ${Date.now() - convertStartedAt}ms`
                    );
                    const luckysheetfile = callback.length > 1 ? JSON.stringify(exportJson) : undefined;
                    callback(univerData.mode, luckysheetfile);
                }
                LuckyExcel.logStep(
                    'XLSX -> Univer import finished',
                    `${Date.now() - startedAt}ms total`
                );
            } catch (err) {
                LuckyExcel.logError(
                    `XLSX -> Univer import failed after ${Date.now() - startedAt}ms`,
                    err
                );
                if (errorHandler) {
                    errorHandler(err as Error);
                } else {
                    console.error(err);
                }
            }
        },
            function (err: Error) {
                LuckyExcel.logError(
                    `ZIP extraction failed after ${Date.now() - startedAt}ms`,
                    err
                );
                if (errorHandler) {
                    errorHandler(err);
                } else {
                    console.error(err);
                }
            });
    }

    static transformCsvToUniver(
        file: File,
        callback?: (files: IWorkbookData, fs?: string[][]) => void,
        errorHandler?: (err: Error) => void
    ) {
        const startedAt = Date.now();
        LuckyExcel.logStep(
            'CSV -> Univer import started',
            `${file.name} (${LuckyExcel.formatFileSize(file.size)})`
        );
        try {
            getDataByFile({ file }).then((source) => {
                LuckyExcel.logStep(
                    'CSV source loaded',
                    `${Date.now() - startedAt}ms`
                );
                const sheetData = formatSheetData(source, file)!;
                const univerData = new UniverCsvWorkBook(sheetData || []);
                LuckyExcel.logStep(
                    'CSV -> Univer import finished',
                    `${sheetData?.length || 0} rows in ${Date.now() - startedAt}ms`
                );
                callback?.(univerData.mode, sheetData);
            }).catch((error) => {
                LuckyExcel.logError(
                    `CSV -> Univer import failed after ${Date.now() - startedAt}ms`,
                    error
                );
                errorHandler?.(error);
            });
        } catch (error) {
            LuckyExcel.logError(
                `CSV -> Univer import failed after ${Date.now() - startedAt}ms`,
                error
            );
            errorHandler?.(error as Error);
        }
    }

    static async transformUniverToExcel(params: {
        snapshot: any,
        fileName?: string,
        getBuffer?: boolean,
        success?: (buffer?: exceljs.Buffer) => void,
        error?: (err: Error) => void
    }) {
        const { snapshot, fileName = `excel_${(new Date).getTime()}.xlsx`, getBuffer = false, success, error } = params;
        const startedAt = Date.now();
        LuckyExcel.logStep(
            'Univer -> XLSX export started',
            `${fileName} (${snapshot?.sheetOrder?.length || 0} sheets)`
        );
        try {
            const workbookStartedAt = Date.now();
            const workbook = new WorkBook(snapshot);
            LuckyExcel.logStep(
                'Excel workbook model created',
                `${Date.now() - workbookStartedAt}ms`
            );
            const bufferStartedAt = Date.now();
            const buffer = await workbook.xlsx.writeBuffer();
            LuckyExcel.logStep(
                'XLSX buffer generated',
                `${Date.now() - bufferStartedAt}ms`
            );
            if (getBuffer) {
                success?.(buffer);
            } else {
                this.downloadFile(fileName, buffer);
                success?.()
            }
            LuckyExcel.logStep(
                'Univer -> XLSX export finished',
                `${Date.now() - startedAt}ms total`
            );

        } catch (err) {
            LuckyExcel.logError(
                `Univer -> XLSX export failed after ${Date.now() - startedAt}ms`,
                err
            );
            error?.(err)
        }
    }

    static async transformUniverToCsv(params: {
        snapshot: any,
        fileName?: string,
        getBuffer?: boolean,
        sheetName?: string,
        success?: (csvContent?: string | { [key: string]: string }) => void,
        error?: (err: Error) => void
    }) {
        const { snapshot, fileName = `csv_${(new Date).getTime()}.csv`, getBuffer = false, success, error, sheetName } = params;
        const startedAt = Date.now();
        LuckyExcel.logStep(
            'Univer -> CSV export started',
            `${fileName} (${sheetName || 'all sheets'})`
        );
        try {
            const csv = new CSV(snapshot);
            LuckyExcel.logStep(
                'CSV content prepared',
                `${Object.keys(csv.csvContent || {}).length} sheet payloads in ${Date.now() - startedAt}ms`
            );

            let contents: string | { [key: string]: string };
            if (sheetName) {
                contents = csv.csvContent[sheetName];
            } else {
                contents = csv.csvContent;
            }
            if (getBuffer) {
                success?.(contents);
            } else {
                if (isObject(contents)) {
                    for (const key in contents) {
                        if (Object.prototype.hasOwnProperty.call(contents, key)) {
                            const element = contents[key];
                            this.downloadFile(`${fileName}_${key}`, element);
                        }
                    }
                } else {
                    this.downloadFile(fileName, contents);
                }
                success?.()
            }
            LuckyExcel.logStep(
                'Univer -> CSV export finished',
                `${Date.now() - startedAt}ms total`
            );
        } catch (err) {
            LuckyExcel.logError(
                `Univer -> CSV export failed after ${Date.now() - startedAt}ms`,
                err
            );
            error?.(err as Error)
        }
    }

    private static logStep(message: string, detail?: string) {
        console.info(`[LuckyExcel] ${message}${detail ? `: ${detail}` : ''}`);
    }

    private static logError(message: string, error: unknown) {
        console.error(`[LuckyExcel] ${message}`, error);
    }

    private static formatFileSize(bytes: number) {
        if (!bytes && bytes !== 0) return 'unknown size';
        if (bytes < 1024) return `${bytes} B`;
        if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
        return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
    }

    private static downloadFile(fileName: string, buffer: exceljs.Buffer | string) {
        const link = document.createElement('a');
        let blob: Blob;
        if (typeof buffer === 'string') {
            blob = new Blob([buffer], { type: "text/csv;charset=utf-8;" });
        } else {
            blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8' });
        }

        const url = URL.createObjectURL(blob);
        link.href = url;
        link.download = fileName;
        document.body.appendChild(link);
        link.click();

        link.addEventListener('click', () => {
            link.remove();
            setTimeout(() => {
                URL.revokeObjectURL(url)
            }, 200);
        })
    }
}
