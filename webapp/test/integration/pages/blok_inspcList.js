sap.ui.define(['sap/fe/test/ListReport'], function(ListReport) {
    'use strict';

    var CustomPageDefinitions = {
        actions: {},
        assertions: {}
    };

    return new ListReport(
        {
            appId: 'zapp.ztestblokinsp',
            componentId: 'blok_inspcList',
            contextPath: '/blok_inspc'
        },
        CustomPageDefinitions
    );
});