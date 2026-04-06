sap.ui.define([
    "sap/m/MessageToast",
    "sap/ui/core/Fragment"
], function (MessageToast, Fragment) {
    'use strict';
    var _oPageContext = null;
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

            // Buka Dialog menggunakan Global Core (Tanpa View)
            var oDialog = sap.ui.getCore().byId("UploadDialog");

            if (!oDialog) {
                // Jika belum dirender, load XML Fragment-nya
                Fragment.load({
                    // KITA HAPUS "id: oView.getId()" AGAR MENJADI GLOBAL
                    name: "zapp.ztestblokinsp.ext.fragment.ExcelUpload",
                    controller: this
                }).then(function (oLoadedDialog) {
                    oLoadedDialog.open();
                });
            } else {
                // Jika sudah ada, langsung buka
                oDialog.open();
            }
        }
    };
});
