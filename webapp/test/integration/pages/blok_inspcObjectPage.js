sap.ui.define(['sap/fe/test/ObjectPage'], function(ObjectPage) {
    'use strict';

    var CustomPageDefinitions = {
        actions: {},
        assertions: {}
    };

    return new ObjectPage(
        {
            appId: 'zapp.ztestblokinsp',
            componentId: 'blok_inspcObjectPage',
            contextPath: '/blok_inspc'
        },
        CustomPageDefinitions
    );
});