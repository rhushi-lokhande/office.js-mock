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

var chartItems = [];





window['context'] = {
	sync: () => '',
	workbook: {
		worksheets: {
			load: () => '',
			getActiveWorksheet: () => {
				return {
					id: 'getActiveWorksheetId',
					load: () => '',
					charts: charts
				};
			},
			getItem: (id) => {
				return workbooks.find(w => w.id === id);
			}
		},
		getSelectedRange: () => {
			return {
				address: 'rangeAddress',
				values: window['values'] || [],
				cellCount: window['cellCount'] || 0,
				load: () => '',
				getImage: () => 'selected range base64 image string'
			};
		}
	}
};

/**
 *  function to set default data
 */
window['officejs-mock-util'] = {
	addChartItems: (data) => {
		chartItems.push({
			id: data.id,
			getImage: () => {
				return {
					value: 'chartbase64imagestring'
				}
			},
			load: () => '',
		});
	},
	addWorkSheet: (w) => {
		workbooks.push({
			id: w.id,
			charts: charts,
			getRange: (id) => {
				return ranges.find(r => r.id === id);
			}
		})
	},
	addRange: (r) => {
		ranges.push({
			id: r.id,
			getImage: () => { return { value: 'range image string' } },
			load: () => '',
			rowCount: 10
		})
	}
}


/**
 * collections
 */

var charts = {
	load: () => '',
	onActivated: {
		add: (cb) => {
			cb({
				chartId: window['chartId'] || 'chartId'
			});
		}
	},
	items: chartItems,
	getItem: (id) => {
		return chartItems.find(c => c.id === id);
	}
}
var workbooks = []
var ranges = [];
