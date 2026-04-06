sap.ui.define(['sap/fe/test/ObjectPage'], function(ObjectPage) {
    'use strict';

    var CustomPageDefinitions = {
        actions: {},
        assertions: {}
    };

    return new ObjectPage(
        {
            appId: 'zapp.ztestblokinsp',
            componentId: 'blok_inspc_iObjectPage',
            contextPath: '/blok_inspc/_blok_inspc_i'
        },
        CustomPageDefinitions
    );
});