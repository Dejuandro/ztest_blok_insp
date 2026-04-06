sap.ui.define([
    "sap/m/MessageToast",
    "sap/m/MessageBox",
    "sap/ui/core/Fragment",
    "sap/ui/thirdparty/jquery"
], function (MessageToast, MessageBox, Fragment, jQuery) {
    'use strict';

    var _oPageContext = null;
    var _oDialogController;
    var _pDialog = null;
    var _pXlsxLibrary = null;
    var _sDialogId = "UploadDialog";
    var _sUploaderId = "excelUploader";
    var _sServiceUrl = "/sap/opu/odata4/sap/zui_blok_insp_ov4/srvd/sap/zui_blok_inspc01/0001/";
    var _mColumnMap = {
        estateid: "EstateID",
        estate: "EstateID",
        afdelingid: "AfdelingID",
        afdeling: "AfdelingID",
        blokid: "BlokID",
        blok: "BlokID",
        tglinspeksi: "TglInspeksi",
        tanggalinspeksi: "TglInspeksi",
        tanggal: "TglInspeksi",
        inspectorid: "InspectorID",
        nikinspector: "InspectorID",
        inspector: "InspectorID",
        inspectorname: "InspectorName",
        namainspector: "InspectorName",
        namapemeriksa: "InspectorName",
        statusinsp: "StatusInsp",
        statusinspeksi: "StatusInsp",
        status: "StatusInsp",
        skortotal: "SkorTotal",
        totalskor: "SkorTotal",
        skor: "SkorTotal",
        catatan: "Catatan",
        keterangan: "Catatan"
    };
    var _aRequiredColumns = [
        "EstateID",
        "AfdelingID",
        "BlokID",
        "InspectorID",
        "InspectorName",
        "StatusInsp",
        "SkorTotal"
    ];
    var _mFieldTypes = {
        EstateID: "string",
        AfdelingID: "string",
        BlokID: "string",
        TglInspeksi: "date",
        InspectorID: "string",
        InspectorName: "string",
        StatusInsp: "string",
        SkorTotal: "number",
        Catatan: "string"
    };
    var _aTemplateHeaders = [
        "EstateID",
        "AfdelingID",
        "BlokID",
        "TglInspeksi",
        "InspectorID",
        "InspectorName",
        "StatusInsp",
        "SkorTotal",
        "Catatan"
    ];
    var _aTemplateExampleRow = [
        "EST001",
        "AFD01",
        "BLK001",
        "2026-04-06",
        "INSP001",
        "Budi Santoso",
        "OPEN",
        95,
        "Contoh catatan inspeksi"
    ];

    function _getDialog() {
        return sap.ui.getCore().byId(_sDialogId);
    }

    function _getUploader() {
        return sap.ui.getCore().byId(_sUploaderId);
    }

    function _setDialogBusy(bBusy) {
        var oDialog = _getDialog();

        if (oDialog) {
            oDialog.setBusyIndicatorDelay(0);
            oDialog.setBusy(bBusy);
        }
    }

    function _resetUploader() {
        var oUploader = _getUploader();
        var oFileInput;

        if (!oUploader) {
            return;
        }

        if (oUploader.clear) {
            oUploader.clear();
        } else {
            oUploader.setValue("");
        }

        oFileInput = document.getElementById(oUploader.getId() + "-fu");
        if (oFileInput) {
            oFileInput.value = "";
        }
    }

    function _closeDialog() {
        var oDialog = _getDialog();

        if (oDialog) {
            oDialog.close();
        }

        _resetUploader();
    }

    function _getSelectedFile() {
        var oUploader = _getUploader();
        var oFileInput;

        if (!oUploader) {
            return null;
        }

        oFileInput = document.getElementById(oUploader.getId() + "-fu");
        if (oFileInput && oFileInput.files && oFileInput.files.length) {
            return oFileInput.files[0];
        }

        return null;
    }

    function _getOrLoadDialog() {
        var oDialog = _getDialog();

        if (oDialog) {
            return Promise.resolve(oDialog);
        }

        if (_pDialog) {
            return _pDialog;
        }

        _pDialog = Fragment.load({
            name: "zapp.ztestblokinsp.ext.fragment.ExcelUpload",
            controller: _oDialogController
        }).then(function (oLoadedDialog) {
            return oLoadedDialog;
        }, function (oError) {
            _pDialog = null;
            throw oError;
        });

        return _pDialog;
    }

    function _loadXlsxLibrary() {
        if (window.XLSX) {
            return Promise.resolve(window.XLSX);
        }

        if (_pXlsxLibrary) {
            return _pXlsxLibrary;
        }

        _pXlsxLibrary = new Promise(function (resolve, reject) {
            var sScriptSelector = 'script[data-xlsx-lib="zapp.ztestblokinsp"]';
            var oExistingScript = document.querySelector(sScriptSelector);
            var oScript;

            function fnResolveLibrary() {
                if (window.XLSX) {
                    resolve(window.XLSX);
                } else {
                    reject(new Error("Library XLSX berhasil dimuat tetapi objek global XLSX tidak ditemukan."));
                }
            }

            if (oExistingScript) {
                if (window.XLSX) {
                    resolve(window.XLSX);
                    return;
                }

                oExistingScript.addEventListener("load", fnResolveLibrary);
                oExistingScript.addEventListener("error", function () {
                    reject(new Error("Gagal memuat library XLSX untuk proses upload."));
                });
                return;
            }

            oScript = document.createElement("script");
            oScript.src = sap.ui.require.toUrl("zapp/ztestblokinsp/lib/xlsx.full.min.js");
            oScript.async = true;
            oScript.setAttribute("data-xlsx-lib", "zapp.ztestblokinsp");
            oScript.onload = fnResolveLibrary;
            oScript.onerror = function () {
                reject(new Error("Gagal memuat library XLSX untuk proses upload."));
            };

            document.head.appendChild(oScript);
        });

        return _pXlsxLibrary;
    }

    function _readFileAsArrayBuffer(oFile) {
        return new Promise(function (resolve, reject) {
            var oReader = new FileReader();

            oReader.onload = function (oEvent) {
                resolve(oEvent.target.result);
            };

            oReader.onerror = function () {
                reject(new Error("File Excel tidak bisa dibaca. Pastikan file tidak rusak."));
            };

            oReader.readAsArrayBuffer(oFile);
        });
    }

    function _normalizeHeader(sHeader) {
        return String(sHeader || "")
            .toLowerCase()
            .replace(/[^a-z0-9]/g, "");
    }

    function _zeroPad(iValue) {
        return iValue < 10 ? "0" + iValue : String(iValue);
    }

    function _formatDateParts(iYear, iMonth, iDay) {
        return [
            String(iYear),
            _zeroPad(iMonth),
            _zeroPad(iDay)
        ].join("-");
    }

    function _formatDateValue(vValue, XLSX) {
        var aDateParts;
        var oDate;
        var sValue;

        if (vValue === null || vValue === undefined || vValue === "") {
            return null;
        }

        if (Object.prototype.toString.call(vValue) === "[object Date]" && !isNaN(vValue.getTime())) {
            return _formatDateParts(vValue.getFullYear(), vValue.getMonth() + 1, vValue.getDate());
        }

        if (typeof vValue === "number" && XLSX && XLSX.SSF && XLSX.SSF.parse_date_code) {
            aDateParts = XLSX.SSF.parse_date_code(vValue);
            if (aDateParts) {
                return _formatDateParts(aDateParts.y, aDateParts.m, aDateParts.d);
            }
        }

        sValue = String(vValue).trim();
        if (!sValue) {
            return null;
        }

        if (/^\d{4}-\d{2}-\d{2}$/.test(sValue)) {
            return sValue;
        }

        aDateParts = sValue.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
        if (aDateParts) {
            return _formatDateParts(aDateParts[3], aDateParts[2], aDateParts[1]);
        }

        oDate = new Date(sValue);
        if (!isNaN(oDate.getTime())) {
            return _formatDateParts(oDate.getFullYear(), oDate.getMonth() + 1, oDate.getDate());
        }

        throw new Error("Format tanggal tidak valid: " + sValue);
    }

    function _parseNumber(vValue) {
        var sValue;
        var iLastComma;
        var iLastDot;
        var fValue;

        if (vValue === null || vValue === undefined || vValue === "") {
            return null;
        }

        if (typeof vValue === "number") {
            return vValue;
        }

        sValue = String(vValue).trim();
        if (!sValue) {
            return null;
        }

        iLastComma = sValue.lastIndexOf(",");
        iLastDot = sValue.lastIndexOf(".");

        if (iLastComma > -1 && iLastDot > -1) {
            if (iLastComma > iLastDot) {
                sValue = sValue.replace(/\./g, "").replace(",", ".");
            } else {
                sValue = sValue.replace(/,/g, "");
            }
        } else if (iLastComma > -1) {
            sValue = sValue.replace(",", ".");
        }

        fValue = Number(sValue);
        if (isNaN(fValue)) {
            throw new Error("Nilai numerik tidak valid: " + vValue);
        }

        return fValue;
    }

    function _convertFieldValue(sFieldName, vValue, XLSX) {
        var sFieldType = _mFieldTypes[sFieldName];

        if (sFieldType === "date") {
            return _formatDateValue(vValue, XLSX);
        }

        if (sFieldType === "number") {
            return _parseNumber(vValue);
        }

        if (vValue === null || vValue === undefined) {
            return "";
        }

        return String(vValue).trim();
    }

    function _buildColumnMapping(aRows) {
        var mHeaderToField = {};
        var mMappedFields = {};
        var aHeaders = aRows.length ? Object.keys(aRows[0]) : [];
        var aMissingColumns;

        aHeaders.forEach(function (sHeader) {
            var sMappedField = _mColumnMap[_normalizeHeader(sHeader)];

            if (sMappedField) {
                mHeaderToField[sHeader] = sMappedField;
                mMappedFields[sMappedField] = true;
            }
        });

        aMissingColumns = _aRequiredColumns.filter(function (sRequiredField) {
            return !mMappedFields[sRequiredField];
        });

        if (aMissingColumns.length) {
            throw new Error(
                "Kolom wajib belum lengkap. Minimal harus ada: " +
                _aRequiredColumns.join(", ") +
                ".\nKolom yang belum ditemukan: " +
                aMissingColumns.join(", ")
            );
        }

        return mHeaderToField;
    }

    function _buildPayloads(aRows, XLSX) {
        var mHeaderToField = _buildColumnMapping(aRows);
        var aErrors = [];
        var aPayloads = [];

        aRows.forEach(function (oRow, iIndex) {
            var oPayload = {
                Catatan: ""
            };
            var iExcelRow = iIndex + 2;
            var bHasRowError = false;

            Object.keys(oRow).forEach(function (sHeader) {
                var sFieldName = mHeaderToField[sHeader];

                if (!sFieldName) {
                    return;
                }

                try {
                    oPayload[sFieldName] = _convertFieldValue(sFieldName, oRow[sHeader], XLSX);
                } catch (oError) {
                    bHasRowError = true;
                    aErrors.push("Baris " + iExcelRow + " (" + sHeader + "): " + oError.message);
                }
            });

            _aRequiredColumns.forEach(function (sFieldName) {
                if (oPayload[sFieldName] === null || oPayload[sFieldName] === undefined || oPayload[sFieldName] === "") {
                    bHasRowError = true;
                    aErrors.push("Baris " + iExcelRow + ": kolom " + sFieldName + " wajib diisi.");
                }
            });

            if (!bHasRowError) {
                if (!oPayload.TglInspeksi) {
                    delete oPayload.TglInspeksi;
                }
                aPayloads.push(oPayload);
            }
        });

        if (!aPayloads.length && !aErrors.length) {
            aErrors.push("Sheet pertama tidak berisi data upload yang bisa diproses.");
        }

        return {
            payloads: aPayloads,
            errors: aErrors
        };
    }

    function _ajax(oSettings) {
        return new Promise(function (resolve, reject) {
            jQuery.ajax({
                url: oSettings.url,
                method: oSettings.method || "GET",
                headers: oSettings.headers || {},
                data: oSettings.data,
                contentType: oSettings.contentType,
                dataType: oSettings.dataType,
                processData: oSettings.processData,
                success: function (oData, sTextStatus, jqXHR) {
                    resolve({
                        data: oData,
                        xhr: jqXHR
                    });
                },
                error: function (jqXHR, sTextStatus, sErrorThrown) {
                    reject({
                        xhr: jqXHR,
                        textStatus: sTextStatus,
                        errorThrown: sErrorThrown
                    });
                }
            });
        });
    }

    function _extractAjaxErrorMessage(oAjaxError) {
        var oResponseJSON;
        var sResponseText;

        if (oAjaxError && oAjaxError.xhr && oAjaxError.xhr.responseJSON && oAjaxError.xhr.responseJSON.error) {
            if (typeof oAjaxError.xhr.responseJSON.error.message === "string") {
                return oAjaxError.xhr.responseJSON.error.message;
            }

            if (oAjaxError.xhr.responseJSON.error.message && oAjaxError.xhr.responseJSON.error.message.value) {
                return oAjaxError.xhr.responseJSON.error.message.value;
            }
        }

        sResponseText = oAjaxError && oAjaxError.xhr && oAjaxError.xhr.responseText;
        if (sResponseText) {
            try {
                oResponseJSON = JSON.parse(sResponseText);
                if (oResponseJSON.error) {
                    if (typeof oResponseJSON.error.message === "string") {
                        return oResponseJSON.error.message;
                    }

                    if (oResponseJSON.error.message && oResponseJSON.error.message.value) {
                        return oResponseJSON.error.message.value;
                    }
                }
            } catch (oParseError) {
                return sResponseText;
            }
        }

        if (oAjaxError && oAjaxError.errorThrown) {
            return oAjaxError.errorThrown;
        }

        return "Terjadi kesalahan saat memanggil service upload.";
    }

    function _fetchCsrfToken() {
        return _ajax({
            url: _sServiceUrl,
            method: "GET",
            headers: {
                "X-CSRF-Token": "Fetch",
                "Accept": "application/json"
            }
        }).then(function (oResponse) {
            var sToken = oResponse.xhr.getResponseHeader("X-CSRF-Token");

            if (!sToken) {
                throw new Error("Token CSRF tidak ditemukan dari service OData.");
            }

            return sToken;
        });
    }

    function _createEntry(oPayload, sCsrfToken) {
        return _ajax({
            url: _sServiceUrl + "blok_inspc",
            method: "POST",
            headers: {
                "X-CSRF-Token": sCsrfToken,
                "Accept": "application/json"
            },
            data: JSON.stringify(oPayload),
            contentType: "application/json",
            dataType: "json",
            processData: false
        });
    }

    function _uploadPayloads(aPayloads, sCsrfToken) {
        var iSuccessCount = 0;

        return aPayloads.reduce(function (oPromise, oPayload, iIndex) {
            return oPromise.then(function () {
                return _createEntry(oPayload, sCsrfToken).then(function () {
                    iSuccessCount += 1;
                }).catch(function (oAjaxError) {
                    var oError = new Error(_extractAjaxErrorMessage(oAjaxError));

                    oError.successCount = iSuccessCount;
                    oError.failedRow = iIndex + 2;
                    throw oError;
                });
            });
        }, Promise.resolve()).then(function () {
            return iSuccessCount;
        });
    }

    function _showValidationErrors(aErrors) {
        var aPreviewErrors = aErrors.slice(0, 10);
        var sMessage = aPreviewErrors.join("\n");

        if (aErrors.length > aPreviewErrors.length) {
            sMessage += "\n... dan " + (aErrors.length - aPreviewErrors.length) + " error lainnya.";
        }

        MessageBox.error(sMessage, {
            title: "Upload dibatalkan"
        });
    }

    function _triggerBrowserDownload(oBlob, sFileName) {
        var sObjectUrl = window.URL.createObjectURL(oBlob);
        var oLink = document.createElement("a");

        oLink.href = sObjectUrl;
        oLink.download = sFileName;
        document.body.appendChild(oLink);
        oLink.click();
        document.body.removeChild(oLink);

        window.setTimeout(function () {
            window.URL.revokeObjectURL(sObjectUrl);
        }, 0);
    }

    function _downloadTemplateFile(XLSX) {
        var oWorkbook = XLSX.utils.book_new();
        var oTemplateSheet = XLSX.utils.aoa_to_sheet([
            _aTemplateHeaders,
            _aTemplateExampleRow
        ]);
        var oInstructionSheet = XLSX.utils.aoa_to_sheet([
            ["Petunjuk Upload"],
            ["Isi data pada sheet TemplateUpload mulai baris ke-2."],
            ["Kolom wajib", _aRequiredColumns.join(", ")],
            ["Format tanggal", "YYYY-MM-DD"],
            ["Kolom opsional", "TglInspeksi, Catatan"]
        ]);
        var sFileName = "Template_Upload_Blok_Inspection.xlsx";
        var aWorkbookBuffer;

        XLSX.utils.book_append_sheet(oWorkbook, oTemplateSheet, "TemplateUpload");
        XLSX.utils.book_append_sheet(oWorkbook, oInstructionSheet, "Petunjuk");

        if (XLSX.writeFile) {
            XLSX.writeFile(oWorkbook, sFileName);
            return;
        }

        aWorkbookBuffer = XLSX.write(oWorkbook, {
            bookType: "xlsx",
            type: "array"
        });

        _triggerBrowserDownload(
            new Blob(
                [aWorkbookBuffer],
                {
                    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                }
            ),
            sFileName
        );
    }

    function onCancelUpload() {
        _closeDialog();
    }

    function onDownloadTemplate() {
        _loadXlsxLibrary()
            .then(function (XLSX) {
                _downloadTemplateFile(XLSX);
                MessageToast.show("Template Excel berhasil didownload.");
            })
            .catch(function (oError) {
                MessageBox.error(oError.message || "Gagal menyiapkan template Excel.");
            });
    }

    function onProcessUpload() {
        var oFile = _getSelectedFile();
        var oValidationResult;
        var oProcessPromise;

        if (!oFile) {
            MessageBox.warning("Pilih file Excel (.xlsx) terlebih dahulu sebelum diproses.");
            return;
        }

        _setDialogBusy(true);

        oProcessPromise = _loadXlsxLibrary()
            .then(function (XLSX) {
                return _readFileAsArrayBuffer(oFile).then(function (aBuffer) {
                    var oWorkbook = XLSX.read(aBuffer, {
                        type: "array",
                        cellDates: true
                    });
                    var sFirstSheetName = oWorkbook.SheetNames[0];
                    var oFirstSheet;
                    var aRows;

                    if (!sFirstSheetName) {
                        throw new Error("File Excel tidak memiliki worksheet yang bisa dibaca.");
                    }

                    oFirstSheet = oWorkbook.Sheets[sFirstSheetName];
                    aRows = XLSX.utils.sheet_to_json(oFirstSheet, {
                        raw: true,
                        defval: null
                    });

                    if (!aRows.length) {
                        throw new Error("Sheet pertama kosong. Tidak ada data yang bisa diproses.");
                    }

                    oValidationResult = _buildPayloads(aRows, XLSX);

                    if (oValidationResult.errors.length) {
                        _showValidationErrors(oValidationResult.errors);
                        return null;
                    }

                    return _fetchCsrfToken().then(function (sCsrfToken) {
                        return _uploadPayloads(oValidationResult.payloads, sCsrfToken);
                    });
                });
            })
            .then(function (iSuccessCount) {
                if (iSuccessCount === null) {
                    return;
                }

                MessageToast.show(iSuccessCount + " data inspeksi berhasil diupload.");
                _closeDialog();
            })
            .catch(function (oError) {
                var sMessage = oError.message || "Proses upload gagal dijalankan.";

                if (oError.failedRow) {
                    sMessage = "Upload gagal pada baris Excel " + oError.failedRow + ".\n" + sMessage;
                }

                if (oError.successCount) {
                    sMessage = oError.successCount + " baris sudah berhasil diproses sebelum terjadi error.\n\n" + sMessage;
                }

                MessageBox.error(sMessage, {
                    title: "Upload gagal"
                });
            });

        oProcessPromise.then(function () {
            _setDialogBusy(false);
        }, function () {
            _setDialogBusy(false);
        });
    }

    _oDialogController = {
        onCancelUpload: onCancelUpload,
        onDownloadTemplate: onDownloadTemplate,
        onProcessUpload: onProcessUpload
    };

    return {
        /**
         * Generated event handler.
         *
         * @param oContext the context of the page on which the event was fired. `undefined` for list report page.
         * @param aSelectedContexts the selected contexts of the table rows.
         */
        UploadExcel: function (oContext, aSelectedContexts) {
            // Simpan context halaman agar bisa dipakai saat proses simpan data nanti
            _oPageContext = oContext;

            _getOrLoadDialog().then(function (oDialog) {
                oDialog.open();
            }).catch(function (oError) {
                MessageBox.error(oError.message || "Dialog upload tidak bisa dibuka.");
            });
        },

        onCancelUpload: onCancelUpload,
        onDownloadTemplate: onDownloadTemplate,
        onProcessUpload: onProcessUpload
    };
});
