
window['Excel'] = {
    run: (resolve) => {
        resolve(window['context']);
    }
};
// Just to mock the office object
window['Office'] = {
    context: {
        document: {
            url: 'test url'
        }
    }
};
window['context']  = {
    sync: () => '',
    workbook: {
        worksheets: {
            load: () => '',
            getActiveWorksheet: () => {
                return {
                    charts: {
                        onActivated: {
                            add: (cb) => {
                                cb({
                                    chartId: window['chartId'] || 'chartId'
                                });
                            }
                        },
                    }

                };
            }
        },
        getSelectedRange : () => {
            return {
                load: () => '',
                values: window['values'] || [],
                cellCount: window['cellCount'] || 0
            };
        }
    }
};
