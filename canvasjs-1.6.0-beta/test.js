window.onload = function () 
{
	var chart1 = new CanvasJS.Chart("chartContainer1",
	{
		title:{
			text: "How my time is spent in a week?",
			fontFamily: "arial black"
		},
                animationEnabled: true,
		legend: {
			verticalAlign: "bottom",
			horizontalAlign: "center"
		},
		theme: "theme1",
		data: [
		{        
			type: "pie",
			indexLabelFontFamily: "Garamond",       
			indexLabelFontSize: 20,
			indexLabelFontWeight: "bold",
			startAngle:0,
			indexLabelFontColor: "MistyRose",       
			indexLabelLineColor: "darkgrey", 
			indexLabelPlacement: "inside", 
			toolTipContent: "{name}: {y}hrs",
			showInLegend: true,
			indexLabel: "#percent%", 
			dataPoints: [
				{  y: 52, name: "Time At Work", legendMarkerType: "triangle"},
				{  y: 44, name: "Time At Home", legendMarkerType: "square"},
				{  y: 12, name: "Time Spent Out", legendMarkerType: "circle"}
			]
		}
		]
	});
	chart1.render();
}

function ()
 {
		var chart = new CanvasJS.Chart("chartContainer",
		{
			theme: "theme3",
                        animationEnabled: true,
			title:{
				text: "Income Vs Expense Vs Savings",
				fontSize: 30
			},
			toolTip: {
				shared: true
			},
			axisX:{
				title: "Source:Income Vs Expense Vs Savings"
			},

			axisY: {
				title: "Rupees"
			},
			axisY2: {
				title: "Money"
			},

			legend:{
				verticalAlign: "top",
				horizontalAlign: "center"
			},
			data: [ 
			{
				type: "column",	
				name: "Income",
				legendText: "Income",
				showInLegend: true, 
				dataPoints:[
				{label: "2013", y: 850000},
				{label: "2014", y: 940000},
				{label: "2105", y: 200000}

				]
			},
			{
				type: "column",	
				name: "Expense",
				legendText: "Expense",
				axisYType: "secondary",
				showInLegend: true,
				dataPoints:[
				{label: "2013", y: 750000},
				{label: "2014", y: 640000},
				{label: "2105", y: 10000}

				]
			},
			{
				type: "column",	
				name: "Savings",
				legendText: "Savings",
				axisYType: "secondary",
				showInLegend: true,
				dataPoints:[
				{label: "2013", y: 50000},
				{label: "2014", y: 40000},
				{label: "2105", y: 6000}

				]
			}
			
			],
          legend:{
            cursor:"pointer",
            itemclick: function(e){
              if (typeof(e.dataSeries.visible) === "undefined" || e.dataSeries.visible) {
              	e.dataSeries.visible = false;
              }
              else {
                e.dataSeries.visible = true;
              }
            	chart.render();
            }
          },
        });

chart.render();

var chart2 = new CanvasJS.Chart("chartContainer2",
		{
			theme: "theme3",
                        animationEnabled: true,
			title:{
				text: "Income Vs Expense Vs Savings",
				fontSize: 30
			},
			toolTip: {
				shared: true
			},
			axisX:{
				title: "Source:Income Vs Expense Vs Savings"
			},

			axisY: {
				title: "Pounds"
			},
			axisY2: {
				title: "herio"
			},

			legend:{
				verticalAlign: "top",
				horizontalAlign: "center"
			},
			data: [ 
			{
				type: "column",	
				name: "Income",
				legendText: "Income",
				showInLegend: true, 
				dataPoints:[
				{label: "2016", y: 850000},
				{label: "2017", y: 940000},
				{label: "2108", y: 200000}

				]
			},
			{
				type: "column",	
				name: "Expense",
				legendText: "Expense",
				axisYType: "secondary",
				showInLegend: true,
				dataPoints:[
				{label: "2016", y: 750000},
				{label: "2017", y: 640000},
				{label: "2108", y: 10000}

				]
			},
			{
				type: "column",	
				name: "Savings",
				legendText: "Savings",
				axisYType: "secondary",
				showInLegend: true,
				dataPoints:[
				{label: "2016", y: 50000},
				{label: "2017", y: 40000},
				{label: "2108", y: 6000}

				]
			}
			
			],
          legend:{
            cursor:"pointer",
            itemclick: function(e){
              if (typeof(e.dataSeries.visible) === "undefined" || e.dataSeries.visible) {
              	e.dataSeries.visible = false;
              }
              else {
                e.dataSeries.visible = true;
              }
            	chart2.render();
            }
          },
        });

chart2.render();
}