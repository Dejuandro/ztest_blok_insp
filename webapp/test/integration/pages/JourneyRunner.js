sap.ui.define([
    "sap/fe/test/JourneyRunner",
	"zapp/ztestblokinsp/test/integration/pages/blok_inspcList",
	"zapp/ztestblokinsp/test/integration/pages/blok_inspcObjectPage",
	"zapp/ztestblokinsp/test/integration/pages/blok_inspc_iObjectPage"
], function (JourneyRunner, blok_inspcList, blok_inspcObjectPage, blok_inspc_iObjectPage) {
    'use strict';

    var runner = new JourneyRunner({
        launchUrl: sap.ui.require.toUrl('zapp/ztestblokinsp') + '/test/flp.html#app-preview',
        pages: {
			onTheblok_inspcList: blok_inspcList,
			onTheblok_inspcObjectPage: blok_inspcObjectPage,
			onTheblok_inspc_iObjectPage: blok_inspc_iObjectPage
        },
        async: true
    });

    return runner;
});

