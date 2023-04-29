var holders = [];
var orgNodes = [
  // {
  //   id: "Shareholders",
  //   color: "blue",
  //   image: null,
  // },
];
let OrgHead = 0;
let OrgName = 0;
let OrgImage = 0;
let DefaultColor = 0;
let ChartHeight = 0;
let ChartWidth = 0;
let NodesWidth = 0;
let TextColor = 0;

document.getElementById("submit").addEventListener("click", () => {
  let upload = document.getElementById("file");
  readXlsxFile(upload.files[0], { sheet: "Options" }).then((data) => {
    orgNodes.push({
      id: data[0][1],
      color: data[2][1],
      image: data[3][1],
    });
    OrgName = data[1][1];
    DefaultColor = data[4][1];
    ChartHeight = data[5][1];
    ChartWidth = data[6][1];
    NodesWidth = data[7][1];
    TextColor = data[8][1];
  });

  //Org Sheet importing
  readXlsxFile(upload.files[0], { sheet: "Database" }).then((data) => {
    // orgNodes.push({
    //   id: data[1][0],
    //   color: "blue",
    //   image: null,
    // });
    data.forEach((r) => {
      if (r[0] != "id") {
        holders.push([r[0], r[1]]);
        preNodes = {
          id: r[1],
          title: r[2],
          name: r[3],
          color: r[6],
          image: r[4],
          level: r[5],
          layout: r[7],
        };
        orgNodes.push(preNodes);
      }
    });
  });

  document.getElementById("submit").setAttribute("disabled", "disabled");
  document.getElementById("start").style.opacity = 1;
});

document.getElementById("start").addEventListener("click", () => {
  document.querySelector(".controller").style.opacity = 0;
  document.querySelector(".controller").style.display = "none";
  orgChart();
});

function orgChart() {
  Highcharts.chart("container", {
    chart: {
      height: ChartHeight,
      inverted: true,
    },
    credits: false,
    title: {
      text: OrgName,
    },

    plotOptions: {
      series: {
        dataLabels: {
          style: {
            fontSize: 16,
          },
        },
      },
    },

    series: [
      {
        type: "organization",
        name: OrgName,
        keys: ["from", "to"],
        data: holders,

        nodes: orgNodes,
        colorByPoint: false,
        color: DefaultColor,
        dataLabels: {
          color: TextColor,
        },
        borderColor: "white",
        nodeWidth: NodesWidth,
      },
    ],
    tooltip: {
      outside: true,
    },
    exporting: {
      allowHTML: true,
      sourceWidth: ChartWidth,
      sourceHeight: ChartHeight,
    },
  });
}
